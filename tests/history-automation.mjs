import {spawn} from 'node:child_process';
import {createServer} from 'node:http';
import {once} from 'node:events';
import {randomBytes,createCipheriv,createHash,createHmac} from 'node:crypto';
import assert from 'node:assert/strict';
const secret=randomBytes(48).toString('base64url'),origin='http://127.0.0.1:3220',appOrigin='https://relay-history.example';
const tables={relay_workflows:[],relay_reviews:[],relay_mail_accounts:[],relay_review_queue:[]};let counter=0,aiCalls=0,graphCalls=0,mode='ok',deltaIds=['new-one'],refreshCalls=0;
let categoryMode='ok',categoryCreates=0,categoryPatches=0,etag=1;const master=[{displayName:'Personal',color:'preset8'}],messageCategories=new Map();
function seal(value,purpose){const iv=randomBytes(12),c=createCipheriv('aes-256-gcm',createHash('sha256').update(secret).digest(),iv);c.setAAD(Buffer.from(purpose));return Buffer.concat([iv,c.update(JSON.stringify(value)),c.final(),c.getAuthTag()]).toString('base64url')}
const cookie=subject=>'relay_ms_0='+seal({subject,email:subject+'@example.com',name:subject,token:'test-mail-token',expires:Date.now()+600000},'session');
const cookies={a:cookie('a'),b:cookie('b')};
function matches(row,q){return [...q].filter(([k])=>!['select','limit','order','on_conflict'].includes(k)).every(([k,v])=>{if(v.startsWith('eq.'))return String(row[k])===v.slice(3);if(v.startsWith('lt.'))return row[k]<v.slice(3);if(v.startsWith('lte.'))return row[k]<=v.slice(4);if(v.startsWith('in.('))return v.slice(4,-1).split(',').includes(row[k]);throw Error('Unknown filter '+v)})}
const mock=createServer(async(req,res)=>{
 const url=new URL(req.url,'http://localhost'),reply=(v,status=200)=>{res.writeHead(status,{'Content-Type':'application/json'});res.end(JSON.stringify(v))};
 let raw='';for await(const part of req)raw+=part;
 if(url.pathname==='/token'){refreshCalls++;const grant=new URLSearchParams(raw).get('grant_type');assert.ok(['refresh_token','authorization_code'].includes(grant));if(mode==='revoked')return reply({error:'invalid_grant'},400);return reply({access_token:'test-mail-token',refresh_token:'rotated-secret',expires_in:3600,scope:categoryMode==='missing-scope'?'Mail.Read':new URLSearchParams(raw).get('scope')})}
 if(url.pathname.startsWith('/graph/')){graphCalls++;assert.equal(req.headers.authorization,'Bearer test-mail-token');assert.match(req.headers.prefer,/ImmutableId/);
  if(url.pathname==='/graph/me')return reply({id:'a',mail:'a@example.com',displayName:'Test A'});
  if(url.pathname==='/graph/me/outlook/masterCategories'){
   if(categoryMode==='forbidden')return reply({},403);
   if(req.method==='GET')return reply({value:master});
   assert.equal(req.method,'POST');const created=JSON.parse(raw);assert.deepEqual(Object.keys(created).sort(),['color','displayName']);categoryCreates++;master.push(created);return reply(created,201);
  }
  if(url.pathname.startsWith('/graph/me/messages/')){
   const id=decodeURIComponent(url.pathname.split('/').at(-1)),existing=messageCategories.get(id)||['Personal'];
   if(categoryMode==='deleted')return reply({},404);
   if(req.method==='GET'){assert.equal(url.searchParams.get('$select'),'id,categories');return reply({id,categories:existing,'@odata.etag':'W/"'+etag+'"'})}
   assert.equal(req.method,'PATCH');categoryPatches++;assert.equal(req.headers['if-match'],'W/"'+etag+'"');const body=JSON.parse(raw);assert.deepEqual(Object.keys(body),['categories']);
   if(categoryMode==='throttled'){res.setHeader('Retry-After','600');return reply({},429)}
   if(categoryMode==='conflict'){categoryMode='ok';messageCategories.set(id,[...existing,'Added in Outlook']);etag++;return reply({},412)}
   assert.ok(existing.every(c=>body.categories.includes(c)),'preserves all existing categories');messageCategories.set(id,body.categories);etag++;return reply({id,categories:body.categories});
  }
  assert.equal(req.method,'GET');
  if(url.pathname.endsWith('/delta'))return reply({value:deltaIds.map(id=>({id,receivedDateTime:new Date().toISOString()})),[mode==='next-page'?'@odata.nextLink':'@odata.deltaLink']:mode==='evil-cursor'?'https://evil.example/steal':'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=opaque'});
  const id=decodeURIComponent(url.pathname.split('/').at(-1));return reply({id,subject:'Desk quotation',from:{emailAddress:{name:'Demo Office',address:'sales@office.example'}},receivedDateTime:new Date().toISOString(),bodyPreview:'Quote please',body:{contentType:'text',content:'Demo Office needs a quote for 24 oak desks.'},isRead:false,hasAttachments:false,webLink:'https://outlook.live.com/mail/item/'+id});
 }
 if(url.pathname==='/openai'){aiCalls++;const body=JSON.parse(raw);if(body.text.format.name==='workflow_proposal'){const current=JSON.parse(body.input).current;return reply({status:'completed',output:[{type:'message',content:[{type:'output_text',text:JSON.stringify({reply:'Here is your proposal.',workflow:current})}]}]})}assert.equal(body.store,false);if(mode==='ai-fail')return reply({},429);if(mode==='slow')await new Promise(r=>setTimeout(r,200));const reviews=JSON.parse(body.input).emails.map(e=>({id:e.id,company:'Demo Office',companyEvidence:'Demo Office',topic:'Quote request',priority:'Normal',summary:'A quote for desks is requested.',priorityReason:'No urgent deadline.',nextAction:'Prepare a quotation.',deadlineEvidence:null,needsAttention:false,uncertainty:null}));return reply({status:'completed',output:[{type:'message',content:[{type:'output_text',text:JSON.stringify({reviews})}]}]})}
 assert.equal(req.headers.apikey,'sb_secret_test');assert.equal(req.headers.authorization,undefined);
 if(url.pathname==='/rest/v1/rpc/relay_configure_scheduler'){assert.match(JSON.parse(raw).p_token,/^[a-f0-9]{64}$/);return reply(null)}
 const table=url.pathname.split('/').at(-1),rows=tables[table];assert.ok(rows,'known table');
 if(req.method==='GET'){let result=rows.filter(r=>matches(r,url.searchParams));if(table==='relay_reviews')result.sort((a,b)=>b.id-a.id);return reply(result.slice(0,Number(url.searchParams.get('limit')||1000)))}
 if(req.method==='POST'){const result=[];for(const input of [JSON.parse(raw)].flat()){const existing=rows.find(r=>r.owner_id===input.owner_id&&(table==='relay_mail_accounts'||r.message_id===input.message_id));if(existing){if(req.headers.prefer?.includes('merge-duplicates'))Object.assign(existing,input);else continue;result.push(existing)}else{const row=table==='relay_reviews'?{id:++counter,revision:1,checked:false,priority:null,topic:null,category_status:'none',category_attempts:0,category_revision:'default-version',category_available_at:new Date(0).toISOString(),...input}:table==='relay_review_queue'?{status:'pending',attempts:0,available_at:new Date(0).toISOString(),created_at:new Date().toISOString(),...input}:input;rows.push(row);result.push(row)}}return reply(result)}
 assert.equal(req.method,'PATCH');const changed=rows.filter(r=>matches(r,url.searchParams));for(const row of changed)Object.assign(row,JSON.parse(raw));return reply(changed);
});
await new Promise(r=>mock.listen(0,'127.0.0.1',r));
const env={...process.env,NODE_ENV:'production',APP_URL:appOrigin,MICROSOFT_CLIENT_ID:'test',MICROSOFT_CLIENT_SECRET:'test',AUTH_SECRET:secret,RELAY_ALLOWED_MICROSOFT_ACCOUNTS:'a@example.com,b@example.com',OPENAI_API_KEY:'test',SUPABASE_URL:`http://127.0.0.1:${mock.address().port}`,SUPABASE_SECRET_KEY:'sb_secret_test',REVIEW_MOCK_URL:`http://127.0.0.1:${mock.address().port}`};let child;
async function start(){child=spawn(process.execPath,['--import','./tests/mock-review-fetch.mjs','node_modules/next/dist/bin/next','start','--hostname','127.0.0.1','--port','3220'],{env,stdio:'ignore'});for(let i=0;i<60;i++){try{await fetch(origin+'/api/outlook/status');return}catch{await new Promise(r=>setTimeout(r,250))}}throw Error('No server')}
async function stop(){if(child&&child.exitCode===null){const done=once(child,'exit');child.kill('SIGTERM');await done}}
const call=(path,method='GET',body,user='a',extra={})=>fetch(origin+path,{method,headers:{Cookie:cookies[user]||'',Origin:appOrigin,'Content-Type':'application/json',...extra},...(body?{body:JSON.stringify(body)}:{})});
const review={source:'outlook',ids:['manual-one'],language:'English',priorityRules:'',consent:true};
const auto=action=>call('/api/outlook/automation','POST',{action});
const history=()=>call('/api/outlook/history');
try{
 await start();assert.equal((await call('/api/outlook/history','GET',null,'none')).status,401);assert.equal((await call('/api/cron/review','POST',{})).status,401);
 assert.equal((await call('/api/outlook/history','PATCH',{id:'1',revision:1,checked:true},'a',{Origin:'https://evil.example'})).status,403);
 let r=await call('/api/outlook/review','POST',review);assert.equal(r.status,200);assert.equal((await r.json()).saved,true);assert.equal(tables.relay_reviews.length,1);assert.ok(!tables.relay_reviews[0].snapshot.includes('oak desks'));
 let h=await (await history()).json();assert.equal(h.items[0].email.id,'manual-one');const savedId=h.items[0].savedId;
 assert.equal((await (await call('/api/outlook/history?owner_id=microsoft:a','GET',null,'b')).json()).items.length,0);
 assert.equal((await call('/api/outlook/history','PATCH',{id:savedId,revision:1,checked:true},'b')).status,409);
 assert.equal((await call('/api/outlook/history','PATCH',{id:savedId,revision:1,priority:'High'})).status,200);
 assert.equal((await call('/api/outlook/history','PATCH',{id:savedId,revision:1,checked:true})).status,409);
 assert.equal((await call('/api/outlook/history','PATCH',{id:savedId,revision:2,checked:true})).status,200);
 await stop();await start();h=await (await history()).json();assert.equal(h.items[0].assessment.priority,'High');assert.equal(h.items[0].checked,true);
 await call('/api/outlook/review','POST',review);assert.equal(tables.relay_reviews.length,1);assert.equal((await (await history()).json()).items[0].checked,true);
 const a={owner_id:'microsoft:a',email:'a@example.com',refresh_cipher:seal('refresh-secret','refresh:microsoft:a'),enabled:true,started_at:new Date(Date.now()-60000).toISOString(),cursor_cipher:null,language:'Lithuanian',rules_cipher:null,lock_id:null,lock_until:new Date(0).toISOString(),budget_date:new Date().toISOString().slice(0,10),budget_used:0};tables.relay_mail_accounts.push(a);
 r=await auto('check');assert.equal((await r.json()).processed,1);assert.equal(refreshCalls,1);assert.ok(a.cursor_cipher);const before=aiCalls;await auto('check');assert.equal(aiCalls,before,'duplicate sync does not invoke AI again');
 await auto('pause');deltaIds=['paused-email'];await auto('check');assert.equal(aiCalls,before);await auto('resume');
 mode='slow';deltaIds=['concurrent-email'];const results=await Promise.all([auto('check'),auto('check')]);assert.equal(aiCalls,before+1,'lease allows only one worker');assert.ok((await Promise.all(results.map(r=>r.json()))).some(x=>x.state==='busy-or-paused'));
 mode='evil-cursor';const checkpoint=a.cursor_cipher;deltaIds=['not-queued'];await auto('check');assert.equal(a.cursor_cipher,checkpoint);assert.ok(!tables.relay_review_queue.some(x=>x.message_id==='not-queued'));
 mode='ai-fail';deltaIds=['failed-email'];await auto('check');let job=tables.relay_review_queue.find(x=>x.message_id==='failed-email');assert.equal(job.attempts,1);assert.equal(job.status,'pending');for(let i=0;i<2;i++){job.available_at=new Date(0).toISOString();await auto('check')}assert.equal(job.status,'failed');await auto('retry');assert.equal(job.attempts,0);mode='ok';await auto('check');assert.equal(job.status,'done');
 a.budget_used=300;deltaIds=['budget-email'];const atLimit=aiCalls;assert.equal((await (await auto('check')).json()).state,'daily-limit');assert.equal(aiCalls,atLimit);
 a.budget_used=0;mode='revoked';await auto('check');assert.equal(a.enabled,false);assert.match(a.last_error,/revoked/);
 mode='ok';await auto('disconnect');assert.equal(a.refresh_cipher,null);assert.equal(a.enabled,false);
 const cronAuth='Bearer '+createHmac('sha256',secret).update('relay-review-cron-v1').digest('hex');assert.equal((await call('/api/cron/review','POST',{},'none',{Authorization:cronAuth})).status,200);
 const status=await (await call('/api/outlook/automation')).text();assert.ok(!status.includes('refresh-secret'));assert.ok(!status.includes('refresh_cipher'));
 // New persistent permission only requested with an authenticated, explicit background opt-in.
 const connect=body=>fetch(origin+'/api/outlook/connect',{method:'POST',headers:{Cookie:cookies.a,Origin:appOrigin,'Content-Type':'application/x-www-form-urlencoded'},body,redirect:'manual'});
 assert.equal((await connect('background=true')).status,400);r=await connect('background=true&consent=true&language=Lithuanian');assert.equal(r.status,303);assert.match(new URL(r.headers.get('location')).searchParams.get('scope'),/offline_access/);

 const workflow={language:'English',instructions:'Keep summaries concise.',schedule:{mode:'continuous',time:'07:00',timezone:'Europe/Vilnius'},labels:[{id:'quotes',name:'Sales',color:'green',field:'subject',value:'quotation',priority:'High',enabled:true}]};
 assert.equal((await call('/api/workflow','GET',null,'none')).status,401);
 assert.equal((await call('/api/workflow','PUT',{data:workflow,revision:null},'a',{Origin:'https://evil.example'})).status,403);
 assert.equal((await call('/api/workflow','PUT',{data:workflow,revision:null})).status,200);
 assert.ok(!tables.relay_workflows[0].settings_cipher.includes('concise'));
 assert.deepEqual((await (await call('/api/workflow')).json()).data,workflow);
 assert.equal((await (await call('/api/workflow','GET',null,'b')).json()).revision,null);
 assert.equal((await call('/api/workflow','PUT',{data:workflow,revision:2})).status,409);
 assert.equal((await call('/api/workflow','PUT',{data:{...workflow,labels:[{...workflow.labels[0],value:''}]},revision:1})).status,400);
 assert.equal((await call('/api/workflow/assistant','POST',{message:'Keep these rules',current:workflow})).status,200);
 assert.equal(tables.relay_workflows[0].revision,1,'assistant proposals never save automatically');
 const withRules=await call('/api/outlook/review','POST',{...review,ids:['label-test']});assert.equal(withRules.status,200);
 const labeled=(await (await history()).json()).items.find(x=>x.email.id==='label-test');assert.equal(labeled.labels[0].name,'Sales');assert.equal(labeled.assessment.priority,'High');

 // Daily mode must leave later mail queued without an AI call; continuous mode releases it.
 const daily={...workflow,schedule:{mode:'daily',time:'00:00',timezone:'UTC'}};
 await call('/api/workflow','PUT',{data:daily,revision:1});
 Object.assign(a,{enabled:true,refresh_cipher:seal('refresh-secret','refresh:microsoft:a'),budget_used:0});
 for(const q of tables.relay_review_queue)if(q.status==='pending')q.status='done';
 tables.relay_review_queue.push({owner_id:'microsoft:a',message_id:'scheduled-mail',status:'pending',attempts:0,available_at:new Date(0).toISOString(),created_at:new Date().toISOString(),received_at:new Date(Date.now()+86400000).toISOString()});
 deltaIds=[];const beforeDaily=aiCalls;await auto('check');assert.equal(aiCalls,beforeDaily,'daily cutoff must gate AI dispatch');
 assert.equal(tables.relay_review_queue.find(x=>x.message_id==='scheduled-mail').status,'pending');
 await call('/api/workflow','PUT',{data:workflow,revision:2});await auto('check');assert.equal(aiCalls,beforeDaily+1);
 assert.equal(tables.relay_review_queue.find(x=>x.message_id==='scheduled-mail').status,'done');
 // Categories have separate consent, owner checks and a durable retry path.
 const cat=(action,rest={},user='a',headers={})=>call('/api/outlook/categories','POST',{action,...rest},user,headers);
 assert.equal((await call('/api/outlook/categories','GET',null,'none')).status,401);
 assert.equal((await cat('apply',{id:savedId},'a',{Origin:'https://evil.example'})).status,403);
 assert.equal((await cat('resume')).status,409,'cannot enable category access without OAuth');
 let normal=await connect('');assert.ok(!new URL(normal.headers.get('location')).searchParams.get('scope').includes('ReadWrite'));
 assert.equal((await connect('categories=true')).status,400);assert.equal((await connect('categories=true&background=true')).status,400);
 const authorize=async()=>{const r=await connect('categories=true&background=true&consent=true');assert.equal(r.status,303);const u=new URL(r.headers.get('location'));assert.match(u.searchParams.get('scope'),/Mail.ReadWrite/);assert.match(u.searchParams.get('scope'),/MailboxSettings.ReadWrite/);assert.ok(!u.searchParams.get('scope').includes('Mail.Send'));return fetch(origin+'/api/outlook/callback?state='+u.searchParams.get('state')+'&code=test',{headers:{Cookie:r.headers.get('set-cookie').split(';')[0]},redirect:'manual'})};
 categoryMode='missing-scope';r=await authorize();assert.match(r.headers.get('location'),/category_permission_required/);assert.ok(!a.category_permission);
 categoryMode='ok';const started=a.started_at,language=a.language;r=await authorize();assert.match(r.headers.get('location'),/categories_connected/);assert.equal(a.category_enabled,true);assert.equal(a.started_at,started);assert.equal(a.language,language,'category consent preserves review settings');
 assert.equal((await cat('apply',{id:savedId},'b')).status,409);
 assert.equal((await cat('apply',{id:'999999'})).status,404);
 assert.equal((await cat('apply',{id:savedId,message_id:'arbitrary'})).status,400);
 const beforeCats=aiCalls;
 categoryMode='conflict';r=await cat('apply',{id:savedId});assert.equal((await r.json()).state,'queued');await auto('check');
 assert.equal(aiCalls,beforeCats,'backfill does not rerun AI');assert.equal(categoryCreates,1);assert.equal(categoryPatches,2,'retries etag conflict');
 assert.deepEqual(messageCategories.get('manual-one'),['Personal','Added in Outlook','Sales']);assert.equal(master.find(x=>x.displayName==='Sales').color,'preset4');
 assert.equal(tables.relay_reviews.find(x=>String(x.id)===savedId).category_status,'done');
 await cat('apply',{id:savedId});await auto('check');assert.equal(categoryPatches,2,'idempotent already-applied categories');assert.equal(categoryCreates,1);
 deltaIds=['auto-category'];await auto('check');const automatic=tables.relay_reviews.find(x=>x.message_id==='auto-category');assert.equal(automatic.category_status,'done');assert.deepEqual(messageCategories.get('auto-category'),['Personal','Sales']);assert.ok(!automatic.category_payload.includes('Sales'));
 await cat('pause');deltaIds=['paused-category'];const patches=categoryPatches;await auto('check');assert.equal(categoryPatches,patches);assert.equal(tables.relay_reviews.find(x=>x.message_id==='paused-category').category_status,'none');
 await cat('resume');deltaIds=['retry-category'];categoryMode='throttled';await auto('check');const retry=tables.relay_reviews.find(x=>x.message_id==='retry-category');assert.equal(retry.category_status,'pending');assert.equal(retry.category_attempts,1);assert.ok(Date.parse(retry.category_available_at)>Date.now()+590000);const beforeRetry=aiCalls;
 categoryMode='ok';retry.category_available_at=new Date(0).toISOString();a.budget_used=300;await auto('check');assert.equal(retry.category_status,'done','category retry runs even at AI daily limit');assert.equal(aiCalls,beforeRetry);a.budget_used=0;
 deltaIds=['missing-category'];categoryMode='deleted';await auto('check');assert.equal(tables.relay_reviews.find(x=>x.message_id==='missing-category').category_status,'skipped');
 deltaIds=['permission-category'];categoryMode='forbidden';await auto('check');assert.equal(a.category_enabled,false);assert.equal(a.category_permission,false);assert.equal(a.enabled,true,'category revocation does not disable AI reviews');
 categoryMode='ok';assert.equal((await cat('resume')).status,409);const statusText=await (await call('/api/outlook/categories')).text();assert.ok(!statusText.includes('refresh_cipher'));assert.ok(!statusText.includes('payload'));
 console.log('PASS: category consent and scopes, owner isolation, atomic review/outbox, automatic sync, manual backfill without AI, additive updates, etag conflict retry, deduplication, pause, throttling, daily cap independence, deleted mail and permission revocation.');
 console.log('PASS: encrypted workflow persistence, owner isolation, revision conflicts, invalid rules rejected, AI proposal does not save, saved rules applied to reviews.');
 console.log('PASS: encrypted persistent history, tenant isolation, correction conflicts and restart, duplicate protection, queue leases, pause/resume, refresh rotation, revoked access, bounded retries, daily cap, safe delta URLs, authenticated cron, explicit offline consent.');
}finally{await stop();mock.close();mock.closeAllConnections()}
