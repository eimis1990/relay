import {spawn} from 'node:child_process';
import {createServer} from 'node:http';
import {once} from 'node:events';
import {randomBytes,createCipheriv,createHash} from 'node:crypto';
import assert from 'node:assert/strict';
const origin='http://127.0.0.1:3219',appOrigin='https://relay-review.example';
let mode='ok',graphCalls=[],aiCalls=[];
const messages=[{id:'one',subject:'Quote for 20 desks',from:{emailAddress:{name:'Demo Office',address:'sales@office.example'}},receivedDateTime:'2026-11-19T08:00:00Z',bodyPreview:'Please quote',body:{contentType:'text',content:'Demo Office needs 20 desks. Please send a quote before Friday.'},hasAttachments:false,isRead:false,webLink:'https://outlook.live.com/mail/item/one'},
{id:'two',subject:'Long email',from:{emailAddress:{address:'sales@example.com'}},receivedDateTime:'2026-11-19T08:00:00Z',bodyPreview:'short preview',body:{contentType:'text',content:'x'.repeat(9000)},hasAttachments:true,isRead:false,webLink:'https://evil.example/'},
{id:'three',subject:'HTML fallback',from:{emailAddress:{address:'sales@example.com'}},bodyPreview:'Plain preview',body:{contentType:'html',content:'<script>bad()</script>'},webLink:'javascript:alert(1)'}];
const mock=createServer(async(req,res)=>{
  const reply=(value,status=200)=>{res.writeHead(status,{'Content-Type':'application/json'});res.end(JSON.stringify(value))};
  const url=new URL(req.url,'http://localhost');
  if(url.pathname.startsWith('/graph/')){
    graphCalls.push({url:req.url,method:req.method});assert.equal(req.method,'GET');
    assert.equal(req.headers.authorization,'Bearer test-mail-token');
    assert.match(req.headers.prefer,/body-content-type="text"/);
    if(mode==='missing')return reply({},404);
    if(url.pathname==='/graph/me/mailFolders/inbox/messages')return reply({value:messages});
    const id=decodeURIComponent(url.pathname.split('/').at(-1));
    return reply(messages.find(m=>m.id===id)||{},messages.some(m=>m.id===id)?200:404);
  }
  assert.equal(url.pathname,'/openai');assert.equal(req.method,'POST');
  let text='';for await(const chunk of req)text+=chunk;const request=JSON.parse(text);aiCalls.push(request);
  assert.equal(request.store,false);assert.equal(request.tools,undefined);assert.equal(request.text.format.strict,true);assert.ok(!text.includes('test-mail-token'));
  const emails=JSON.parse(request.input).emails;
  assert.ok(emails.length<=5);for(const e of emails){assert.ok(e.text.length<=8000);assert.equal(e.webLink,undefined)}
  if(mode==='rate')return reply({error:{message:'sensitive provider error'}},429);
  if(mode==='incomplete')return reply({status:'incomplete',output:[]});
  if(mode==='refusal')return reply({status:'completed',output:[{type:'message',content:[{type:'refusal',refusal:'no'}]}]});
  const reviews=emails.map(e=>({id:e.id,company:'Demo Office',companyEvidence:mode==='hallucination'?'Invented company':'Demo Office',topic:'Quote request',priority:'Normal',summary:'The sender requests a quote for desks.',priorityReason:'No immediate deadline risk is established.',nextAction:'Prepare a quote and verify the requirements.',deadlineEvidence:mode==='hallucination'?'Imaginary deadline':e.id==='one'?'before Friday':null,needsAttention:false,uncertainty:null}));
  if(mode==='wrong-id')reviews[0].id='not-requested';
  if(mode==='duplicate'&&reviews.length>1)reviews[1].id=reviews[0].id;
  if(mode==='invalid')reviews[0].priority='Critical';
  reply({status:'completed',output:[{type:'message',content:[{type:'output_text',text:JSON.stringify({reviews})}]}]});
});
await new Promise(r=>mock.listen(0,'127.0.0.1',r));
const secret=randomBytes(48).toString('base64url');
const iv=randomBytes(12),cipher=createCipheriv('aes-256-gcm',createHash('sha256').update(secret).digest(),iv);cipher.setAAD(Buffer.from('session'));
const token=Buffer.concat([iv,cipher.update(JSON.stringify({subject:'owner',email:'owner@example.com',token:'test-mail-token',name:'Owner',expires:Date.now()+600000})),cipher.final(),cipher.getAuthTag()]).toString('base64url');
const Cookie='relay_ms_0='+token;
const env={...process.env,NODE_ENV:'production',APP_URL:appOrigin,MICROSOFT_CLIENT_ID:'test',MICROSOFT_CLIENT_SECRET:'test',AUTH_SECRET:secret,RELAY_ALLOWED_MICROSOFT_ACCOUNTS:'owner@example.com',OPENAI_API_KEY:'test-ai-key',SUPABASE_URL:'',SUPABASE_SECRET_KEY:'',REVIEW_MOCK_URL:`http://127.0.0.1:${mock.address().port}`};
let child;
async function start(key='test-ai-key'){
 child=spawn(process.execPath,['--import','./tests/mock-review-fetch.mjs','node_modules/next/dist/bin/next','start','--hostname','127.0.0.1','--port','3219'],{env:{...env,OPENAI_API_KEY:key},stdio:'ignore'});
 for(let i=0;i<60;i++){try{await fetch(origin+'/api/outlook/status');return}catch{await new Promise(r=>setTimeout(r,250))}}throw Error('Server did not start');
}
async function stop(){if(child&&child.exitCode===null){const done=once(child,'exit');child.kill('SIGTERM');await done}}
const payload={source:'outlook',ids:['one'],language:'English',priorityRules:'Prioritize delivery problems.',consent:true};
const post=(body=payload,headers={})=>fetch(origin+'/api/outlook/review',{method:'POST',headers:{Cookie,Origin:appOrigin,'Content-Type':'application/json',...headers},body:JSON.stringify(body)});
try{
 await start();
 assert.equal((await fetch(origin+'/api/outlook/review')).status,401);
 assert.equal((await post(payload,{Cookie:''})).status,401);
 assert.equal((await post(payload,{Cookie:'relay_ms_0=forged'})).status,401);
 assert.equal((await post(payload,{Origin:'https://evil.example'})).status,403);
 assert.equal((await post({...payload,consent:false})).status,400);
 assert.equal((await post({...payload,ids:['one','one']})).status,400);
 assert.equal((await post({...payload,ids:['1','2','3','4','5','6']})).status,400);
 assert.equal((await post({...payload,token:'forged'})).status,400);
 assert.equal((await post({...payload,priorityRules:'x'.repeat(15000)})).status,413);
 assert.equal(graphCalls.length,0);assert.equal(aiCalls.length,0);
 let r=await fetch(origin+'/api/outlook/messages',{headers:{Cookie}});assert.equal(r.status,200);let j=await r.json();assert.equal(j.messages[0].senderEmail,'sales@office.example');assert.equal(j.messages[1].webLink,null);assert.equal(j.messages[2].webLink,null);
 r=await post();assert.equal(r.status,200);assert.match(r.headers.get('cache-control'),/no-store/);j=await r.json();assert.equal(j.items[0].assessment.company,'Demo Office');assert.equal(j.items[0].assessment.deadlineEvidence,'before Friday');assert.equal(j.source,'outlook');
 r=await post({...payload,ids:['two','three']});assert.equal(r.status,200);j=await r.json();assert.equal(j.items[0].email.text.length,8000);assert.equal(j.items[0].email.truncated,true);assert.equal(j.items[0].assessment.needsAttention,true);assert.equal(j.items[1].email.previewOnly,true);assert.equal(j.items[1].email.text,'Plain preview');
 mode='hallucination';r=await post();j=await r.json();assert.equal(j.items[0].assessment.company,null);assert.equal(j.items[0].assessment.deadlineEvidence,null);assert.equal(j.items[0].assessment.needsAttention,true);
 for(const [failure,expected] of [['missing',409],['rate',429],['incomplete',502],['refusal',422],['wrong-id',502],['invalid',502]]){mode=failure;r=await post();assert.equal(r.status,expected,failure);assert.ok(!(await r.text()).includes('sensitive provider error'))}
 mode='duplicate';assert.equal((await post({...payload,ids:['one','two']})).status,502);
 mode='ok';const count=graphCalls.length;r=await post({...payload,source:'demo',ids:[],language:'Lithuanian'});assert.equal(r.status,200);j=await r.json();assert.equal(j.source,'demo');assert.equal(j.items.length,3);assert.equal(graphCalls.length,count);assert.match(aiCalls.at(-1).instructions,/Lithuanian/);
 assert.ok(graphCalls.every(c=>c.method==='GET'&&!c.url.includes('/attachments')));
 await stop();await start('');assert.equal((await post()).status,503);
 console.log('PASS: authenticated same-origin consent, batch limits, read-only Graph, bounded plain text, safe links, evidence checks, ID/schema validation, refusal/incomplete/provider errors, fictional demo isolation, missing-key handling.');
}finally{await stop();mock.close();mock.closeAllConnections()}
