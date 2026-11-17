import {spawn} from 'node:child_process';
import {randomBytes} from 'node:crypto';
import assert from 'node:assert/strict';
const origin='http://127.0.0.1:3217';
// Public origin is HTTPS in production; these tests never contact Microsoft.
const configuredOrigin='https://relay-test.example';
const child=spawn(process.execPath,['node_modules/next/dist/bin/next','start','--hostname','127.0.0.1','--port','3217'],{env:{...process.env,NODE_ENV:'production',APP_URL:configuredOrigin,MICROSOFT_CLIENT_ID:'test-client',MICROSOFT_CLIENT_SECRET:'test-secret',AUTH_SECRET:randomBytes(48).toString('base64url'),RELAY_ALLOWED_MICROSOFT_ACCOUNTS:'owner@example.com',SUPABASE_URL:'',SUPABASE_SECRET_KEY:''},stdio:'ignore'});
try{
let ready=false;for(let i=0;i<60;i++){try{await fetch(origin+'/api/outlook/status');ready=true;break}catch{await new Promise(r=>setTimeout(r,250))}}assert.ok(ready,'server starts');
let r=await fetch(origin+'/api/outlook/status');assert.deepEqual(await r.json(),{configured:true,connected:false,account:null});
r=await fetch(origin+'/api/outlook/messages');assert.equal(r.status,401);
r=await fetch(origin+'/api/outlook/connect',{method:'POST',headers:{Origin:'https://evil.example'},redirect:'manual'});assert.equal(r.status,403);
r=await fetch(origin+'/api/outlook/connect',{method:'POST',headers:{Origin:configuredOrigin},redirect:'manual'});assert.equal(r.status,303);const u=new URL(r.headers.get('location'));assert.equal(u.origin,'https://login.microsoftonline.com');assert.equal(u.searchParams.get('code_challenge_method'),'S256');assert.ok(u.searchParams.get('state').length>=32);assert.ok(!u.searchParams.get('scope').includes('Mail.Send'));const cookie=r.headers.get('set-cookie');assert.match(cookie,/HttpOnly/i);assert.match(cookie,/Secure/i);assert.match(cookie,/SameSite=lax/i);
r=await fetch(origin+'/api/outlook/callback?state=forged&code=fake',{headers:{Cookie:cookie.split(';')[0]},redirect:'manual'});assert.equal(r.status,303);assert.match(r.headers.get('location'),/outlook=failed/);assert.match(r.headers.get('set-cookie'),/Max-Age=0/i);
r=await fetch(origin+'/api/outlook/disconnect',{method:'POST',headers:{Origin:'https://evil.example'}});assert.equal(r.status,403);
r=await fetch(origin+'/api/workspace');assert.equal(r.status,401);
console.log('PASS: unauthenticated inbox denied, PKCE + secure cookie, forged callback denied, cross-origin writes denied, production local storage disabled.');
}finally{child.kill('SIGTERM')}
