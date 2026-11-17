import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import { once } from 'node:events';
import { createCipheriv, createHash, randomBytes } from 'node:crypto';
import assert from 'node:assert/strict';

// HTTP contract tests: a fake PostgREST stores rows across real Next server restarts.
// These do not substitute for the deployed database's RLS and persistence checks.
const rows = new Map();
let databaseFails = false;
const dbKey = 'sb_secret_' + randomBytes(24).toString('hex');
const db = createServer(async (req, res) => {
  assert.equal(req.headers.apikey, dbKey);
  assert.equal(req.headers.authorization, undefined);
  const reply = (value, status=200) => { res.writeHead(status, {'Content-Type':'application/json'}); res.end(JSON.stringify(value)); };
  if (databaseFails) return reply({internal:'must not leak this database error'}, 500);
  const url = new URL(req.url, 'http://localhost');
  assert.equal(url.pathname, '/rest/v1/relay_workspaces');
  if(req.method==='GET') {
    const owner = url.searchParams.get('owner_id');
    assert.ok(owner?.startsWith('eq.microsoft:'));
    const row=rows.get(owner.slice(3)); return reply(row?[row]:[]);
  }
  let body='';for await(const chunk of req)body+=chunk;
  const update=JSON.parse(body);
  if(req.method==='POST') {
    assert.ok(update.owner_id.startsWith('microsoft:'));
    if(rows.has(update.owner_id))return reply({},409);
    rows.set(update.owner_id,update);return reply([update]);
  }
  assert.equal(req.method,'PATCH');
  const owner=url.searchParams.get('owner_id');const revision=url.searchParams.get('revision');
  assert.ok(owner?.startsWith('eq.microsoft:'));assert.ok(revision?.startsWith('eq.'));
  const existing=rows.get(owner.slice(3));
  if(!existing||existing.revision!==Number(revision.slice(3)))return reply([]);
  const row={...existing,...update};rows.set(owner.slice(3),row);return reply([row]);
});
await new Promise(resolve=>db.listen(0,'127.0.0.1',resolve));
const secret=randomBytes(48).toString('base64url');
const origin='http://127.0.0.1:3218', configuredOrigin='https://relay-test.example';
const env={...process.env,NODE_ENV:'production',APP_URL:configuredOrigin,MICROSOFT_CLIENT_ID:'test',MICROSOFT_CLIENT_SECRET:'test',AUTH_SECRET:secret,RELAY_ALLOWED_MICROSOFT_ACCOUNTS:'a@example.com,b@example.com,a-alias@example.com',SUPABASE_URL:`http://127.0.0.1:${db.address().port}`,SUPABASE_SECRET_KEY:dbKey};
let child;
async function start(overrides={}) {
  child=spawn(process.execPath,['node_modules/next/dist/bin/next','start','--hostname','127.0.0.1','--port','3218'],{env:{...env,...overrides},stdio:'ignore'});
  for(let i=0;i<60;i++){try{await fetch(origin+'/api/outlook/status');return}catch{await new Promise(r=>setTimeout(r,250))}}
  throw Error('Server did not start');
}
async function stop(){if(child&&child.exitCode===null){const exit=once(child,'exit');child.kill('SIGTERM');await exit}}
function cookie(subject,email=subject+'@example.com',expires=Date.now()+600000) {
  const iv=randomBytes(12),cipher=createCipheriv('aes-256-gcm',createHash('sha256').update(secret).digest(),iv);
  cipher.setAAD(Buffer.from('session'));
  const value=Buffer.concat([iv,cipher.update(JSON.stringify({subject,email,name:subject,token:'test-token',expires})),cipher.final(),cipher.getAuthTag()]).toString('base64url');
  return 'relay_ms_0='+value;
}
const data={company:'A Company',contact:'Test Owner',email:'a@example.com',industry:'Other',volume:'200–300',mailboxes:'1',provider:'Outlook.com',goals:['Prioritize incoming emails'],pain:'',urgency:'',languages:'English',attachments:'No',access:'One team',notes:'Test only',submitted:false,consent:false,connections:{outlook:{account:'',requested:false},google:{account:'',requested:false},telegram:{account:'',requested:false}}};
const a=cookie('a'),b=cookie('b');
const read=(Cookie,query='')=>fetch(origin+'/api/workspace'+query,{headers:Cookie?{Cookie}:{}});
const contexts=new Map();
const write=(Cookie,payload,Origin=configuredOrigin)=>fetch(origin+'/api/workspace',{method:'POST',headers:{Cookie,Origin,'Content-Type':'application/json'},body:JSON.stringify({workspaceId:contexts.get(Cookie)??null,...payload})});
try {
  await start();
  contexts.set(a,(await (await read(a)).json()).workspaceId);
  contexts.set(b,(await (await read(b)).json()).workspaceId);
  assert.equal((await read()).status,401);
  assert.equal((await read('relay_ms_0=forged')).status,401);
  assert.equal((await read(cookie('a','a@example.com',Date.now()-1000))).status,401);
  assert.equal((await read(cookie('a','outsider@example.com'))).status,401);
  assert.equal((await write('',{data,revision:null})).status,401);
  assert.equal((await write(a,{data,revision:null},'https://evil.example')).status,403);
  assert.equal((await write(a,{data,revision:null,owner_id:'microsoft:b'})).status,400);
  assert.equal((await write(a,{data:{...data,owner_id:'microsoft:b'},revision:null})).status,400);
  assert.equal((await write(a,{data:{...data,submitted:true},revision:null})).status,400);
  assert.equal((await write(a,{data:{...data,notes:'x'.repeat(25000)},revision:null})).status,413);
  let r=await read(a);assert.equal(r.status,200);assert.equal((await r.json()).data,null);
  r=await write(a,{data,revision:null});assert.equal(r.status,200);assert.equal((await r.json()).revision,1);
  assert.equal((await write(a,{data,revision:null})).status,409);
  r=await read(b,'?owner_id=microsoft:a');assert.equal((await r.json()).data,null);
  assert.equal((await write(b,{data,revision:null,workspaceId:contexts.get(a)})).status,409);
  r=await write(b,{data:{...data,company:'B Company'},revision:null});assert.equal(r.status,200);
  r=await write(a,{data:{...data,notes:'Updated'},revision:1});assert.equal(r.status,200);assert.equal((await r.json()).revision,2);
  assert.equal((await write(a,{data,revision:1})).status,409);
  r=await read(cookie('a','a-alias@example.com'));assert.equal((await r.json()).data.notes,'Updated');
  r=await read(b);assert.equal((await r.json()).data.company,'B Company');
  await stop();await start();
  r=await read(a);assert.match(r.headers.get('cache-control'),/no-store/);const saved=await r.json();assert.equal(saved.revision,2);assert.equal(saved.data.notes,'Updated');assert.equal(JSON.stringify(saved).includes('test-token'),false);
  databaseFails=true;r=await read(a);assert.equal(r.status,503);assert.equal((await r.text()).includes('must not leak'),false);
  assert.equal((await write(a,{data,revision:2})).status,503);databaseFails=false;
  assert.equal((await (await read(a)).json()).revision,2);
  await stop();await start({SUPABASE_URL:'',SUPABASE_SECRET_KEY:''});
  assert.equal((await read(a)).status,503); // No local-file fallback in production.
  console.log('PASS: private account isolation, alias-stable ownership, invalid/expired sessions, CSRF, schema and size checks, conflicting saves, restart persistence contract, sanitized storage failures, production fallback denied.');
} finally {await stop();db.close();db.closeAllConnections()}
