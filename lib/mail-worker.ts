import 'server-only';
import {readWorkflow} from './workflow-store';
import {eligibleAt,applyWorkflow} from './workflow';
import {createHmac,randomUUID} from 'node:crypto';
import {config,seal,unseal,scope,categoryScope,hasCategoryScopes,graph} from './outlook';
import {account,db,saveReview,type Account} from './review-store';
import {reviewEmails,collectReviewEmails,ReviewError} from './review-service';
import type {ReviewEmail,ReviewRequest} from './inbox-review';
import {syncCategories} from './outlook-categories';
export const DAILY_LIMIT=300;
export function cronToken(){return createHmac('sha256',config().secret).update('relay-review-cron-v1').digest('hex')}
export async function setupScheduler(){await db('rpc/relay_configure_scheduler',{}, {method:'POST',body:JSON.stringify({p_token:cronToken()})})}
export function deltaUrl(value:string){
 const u=new URL(value);
 if(u.origin!=='https://graph.microsoft.com'||u.username||u.password||u.hash||!u.pathname.startsWith('/v1.0/me/mailFolders')&&!u.pathname.startsWith('/v1.0/me/mailfolders')||!u.pathname.endsWith('/messages/delta'))throw new ReviewError('Microsoft returned an invalid sync link.',502);
 return u.href;
}
async function refreshedToken(a:Account,lease:string){
 if(!a.refresh_cipher)throw new ReviewError('Reconnect Outlook for automatic processing.',401);
 const cfg=config();
 const r=await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({client_id:cfg.clientId,client_secret:cfg.clientSecret,grant_type:'refresh_token',refresh_token:unseal<string>(a.refresh_cipher,'refresh:'+a.owner_id),scope:(a.category_permission?categoryScope:scope)+' offline_access'}),cache:'no-store',signal:AbortSignal.timeout(15000)});
 if(!r.ok){const body=await r.json().catch(()=>({}));throw new ReviewError(body.error==='invalid_grant'?'Microsoft authorization expired or was revoked. Reconnect automatic processing.':'Microsoft token refresh is temporarily unavailable.',body.error==='invalid_grant'?401:502)}
 const token=await r.json();if(typeof token.access_token!=='string')throw new ReviewError('Microsoft did not return a token.',502);
 if(token.refresh_token){const saved=await db('relay_mail_accounts',{owner_id:'eq.'+a.owner_id,lock_id:'eq.'+lease,enabled:'eq.true',refresh_cipher:'eq.'+a.refresh_cipher},{method:'PATCH',body:JSON.stringify({refresh_cipher:seal(token.refresh_token,'refresh:'+a.owner_id)})});if(!saved.length)throw new ReviewError('Background connection changed. Check its status before retrying.',409)}
 if(a.category_permission&&!hasCategoryScopes(token.scope)){await db('relay_mail_accounts',{owner_id:'eq.'+a.owner_id,lock_id:'eq.'+lease},{method:'PATCH',body:JSON.stringify({category_enabled:false,category_permission:false,category_error:'Microsoft did not grant category permissions. Reconnect category sync.'})});}
 return token.access_token as string;
}
export async function enqueue(owner:string,ids:string[],received:Record<string,string>={}){if(!ids.length)return;await db('relay_review_queue',{on_conflict:'owner_id,message_id'},{method:'POST',body:JSON.stringify([...new Set(ids)].map(message_id=>({owner_id:owner,message_id,...(received[message_id]?{received_at:received[message_id]}:{})})))},'resolution=ignore-duplicates,return=representation')}
export async function runMailbox(owner:string){
 const lease=randomUUID(),now=new Date(),today=now.toISOString().slice(0,10);
 const claimed=await db('relay_mail_accounts',{owner_id:'eq.'+owner,enabled:'eq.true',lock_until:'lt.'+now.toISOString()},{method:'PATCH',body:JSON.stringify({lock_id:lease,lock_until:new Date(Date.now()+240000).toISOString(),last_polled_at:now.toISOString()})});
 if(!claimed.length)return {state:'busy-or-paused',processed:0};
 let a=claimed[0] as Account,processed=0;
 const patch=async(data:object)=>{const rows=await db('relay_mail_accounts',{owner_id:'eq.'+owner,lock_id:'eq.'+lease},{method:'PATCH',body:JSON.stringify(data)});if(!rows.length)throw new ReviewError('Processing lease expired. Retry later.',409)};
 try{
  if(!config().allowed.includes(a.email.toLowerCase())){await patch({enabled:false,last_error:'Account is no longer allowed.'});return {state:'not-allowed',processed}}
  const workflow=(await readWorkflow(owner,a)).data;
  const token=await refreshedToken(a,lease);
  const initial='https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?'+new URLSearchParams({$select:'id,receivedDateTime',$orderby:'receivedDateTime desc',changeType:'created'});
  const url=deltaUrl(a.cursor_cipher?unseal<string>(a.cursor_cipher,'cursor:'+owner):initial);
  // One durable delta page per run. The nextLink itself is saved, so large arrivals are never skipped.
  const r=await fetch(url,{headers:{Authorization:'Bearer '+token,Prefer:'IdType="ImmutableId", odata.maxpagesize=100'},cache:'no-store',signal:AbortSignal.timeout(15000)});
  if(r.status===410){await patch({cursor_cipher:null,last_error:'Microsoft sync expired; rebuilding from the activation date on the next check.'});return {state:'resync',processed}}
  if(!r.ok)throw new ReviewError(r.status===401||r.status===403?'Microsoft mailbox permission needs reconnecting.':'Microsoft inbox check failed. Will retry.',r.status===401||r.status===403?401:502);
  const data=await r.json();if(!Array.isArray(data.value))throw new ReviewError('Microsoft returned an invalid inbox page.',502);
  const next=data['@odata.nextLink']||data['@odata.deltaLink'];if(typeof next!=='string')throw new ReviewError('Microsoft did not return a sync checkpoint.',502);
  deltaUrl(next);
  await enqueue(owner,data.value.filter((m:any)=>!m['@removed']&&typeof m.id==='string'&&m.id.length<=1500&&typeof m.receivedDateTime==='string'&&Date.parse(m.receivedDateTime)>=Date.parse(a.started_at!)).map((m:any)=>m.id),Object.fromEntries(data.value.filter((m:any)=>typeof m.id==='string'&&Number.isFinite(Date.parse(m.receivedDateTime))).map((m:any)=>[m.id,m.receivedDateTime])));
  await patch({cursor_cipher:seal(next,'cursor:'+owner)});
  const used=a.budget_date===today?a.budget_used:0;
  if(used>=DAILY_LIMIT){await syncCategories(owner,token,lease);await patch({last_error:'Daily limit reached (300 email review attempts). Processing resumes tomorrow UTC.'});return {state:'daily-limit',processed}}
  const candidates=await db('relay_review_queue',{owner_id:'eq.'+owner,status:'eq.pending',available_at:'lte.'+new Date().toISOString(),order:'created_at.asc',limit:'100'});
  const pending=candidates.filter((job:any)=>eligibleAt(job.received_at||job.created_at,workflow)).slice(0,Math.min(5,DAILY_LIMIT-used));
  const emails:ReviewEmail[]=[],jobs:any[]=[];
  for(const job of pending){
   const existing=await db('relay_reviews',{owner_id:'eq.'+owner,message_id:'eq.'+job.message_id,select:'id',limit:'1'});
   if(existing.length){await db('relay_review_queue',{owner_id:'eq.'+owner,message_id:'eq.'+job.message_id},{method:'PATCH',body:JSON.stringify({status:'done',last_error:null})});continue}
   if(job.attempts>=3){await db('relay_review_queue',{owner_id:'eq.'+owner,message_id:'eq.'+job.message_id},{method:'PATCH',body:JSON.stringify({status:'failed',last_error:'Retry limit reached. Retry from Relay.'})});continue}
   try{const request={source:'outlook',ids:[job.message_id],language:a.language,priorityRules:'',consent:true} as ReviewRequest;emails.push(...await collectReviewEmails(token,request));jobs.push(job)}
   catch(e){if(e instanceof ReviewError&&e.status===409){await db('relay_review_queue',{owner_id:'eq.'+owner,message_id:'eq.'+job.message_id},{method:'PATCH',body:JSON.stringify({status:'skipped',last_error:'Email moved or deleted before review.'})})}else throw e}
  }
  if(emails.length){
   // Recheck pause before sharing any email with AI. An already-running AI call can finish after pause.
   const current=await account(owner);if(!current?.enabled)return {state:'paused',processed};
   await patch({budget_date:today,budget_used:used+emails.length});
   for(const job of jobs)await db('relay_review_queue',{owner_id:'eq.'+owner,message_id:'eq.'+job.message_id},{method:'PATCH',body:JSON.stringify({attempts:job.attempts+1,available_at:new Date(Date.now()+300000).toISOString()})});
   try{
    const request={source:'outlook',ids:emails.map(e=>e.id),language:workflow.language,priorityRules:workflow.instructions,consent:true} as ReviewRequest;
    const result=await reviewEmails(emails,request);
    for(const item of result.items){await saveReview(owner,applyWorkflow(item,workflow),workflow.language,result.reviewedAt);await db('relay_review_queue',{owner_id:'eq.'+owner,message_id:'eq.'+item.email.id},{method:'PATCH',body:JSON.stringify({status:'done',last_error:null})});processed++}
   }catch(e){for(const job of jobs)await db('relay_review_queue',{owner_id:'eq.'+owner,message_id:'eq.'+job.message_id},{method:'PATCH',body:JSON.stringify({status:job.attempts+1>=3?'failed':'pending',available_at:new Date(Date.now()+300000*Math.pow(3,job.attempts)).toISOString(),last_error:e instanceof ReviewError?e.message:'Review interrupted. Retrying safely.'})});throw e}
  }
  await syncCategories(owner,token,lease);
  await patch({last_success_at:new Date().toISOString(),last_error:null});return {state:'complete',processed};
 }catch(e){const error=e instanceof ReviewError?e.message:'Processing was interrupted. Relay will retry.';await patch({last_error:error,...(e instanceof ReviewError&&e.status===401?{enabled:false}:{})});return {state:e instanceof ReviewError&&e.status===401?'reconnect':'retry',processed}}
 finally{await patch({lock_id:null,lock_until:new Date(0).toISOString()})}
}
