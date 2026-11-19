import 'server-only';
import {z} from 'zod';
import {randomUUID} from 'node:crypto';
import {seal,unseal} from './outlook';
import {account,db,snapshotPurpose} from './review-store';
import {readWorkflow} from './workflow-store';
import {applyWorkflow,labelColors} from './workflow';
import {ReviewError} from './review-service';
import type {ReviewItem} from './inbox-review';
const labelsSchema=z.array(z.object({name:z.string().trim().min(1).max(40),color:z.enum(labelColors)}).strict()).max(12);
const colors={green:'preset4',teal:'preset5',blue:'preset7',violet:'preset8',amber:'preset3',coral:'preset0'};
const purpose=(owner:string,message:string)=>'categories:'+owner+':'+message;
class CategoryError extends Error {constructor(message:string,public status=502,public delay=300){super(message)}}
async function request(token:string,path:string,method='GET',body?:object,etag?:string){
 const r=await fetch('https://graph.microsoft.com/v1.0/'+path,{method,headers:{Authorization:'Bearer '+token,Prefer:'IdType="ImmutableId"',...(body?{'Content-Type':'application/json'}:{}),...(etag?{'If-Match':etag}:{})},...(body?{body:JSON.stringify(body)}:{}),cache:'no-store',signal:AbortSignal.timeout(10000)});
 if(!r.ok){const status=r.status;throw new CategoryError(status===403||status===401?'Microsoft category permission needs reconnecting.':status===404?'This email is no longer available.':status===412?'The email changed during sync. Retrying with its latest categories.':'Microsoft category sync is temporarily unavailable.',status,Math.max(300,Math.min(86400,Number(r.headers.get('retry-after'))||300)))}
 return r.status===204?{}:r.json();
}
async function masterCategories(token:string){
 const names=new Map<string,string>();let path='me/outlook/masterCategories?$select=id,displayName,color';
 for(let page=0;page<20;page++){
  const data=await request(token,path);if(!Array.isArray(data.value))throw new CategoryError('Microsoft returned an invalid category list.');
  for(const item of data.value)if(typeof item.displayName==='string')names.set(item.displayName.toLowerCase(),item.displayName);
  if(!data['@odata.nextLink'])return names;
  const u=new URL(data['@odata.nextLink']);if(u.origin!=='https://graph.microsoft.com'||u.pathname!=='/v1.0/me/outlook/masterCategories'||u.username||u.password||u.hash)throw new CategoryError('Microsoft returned an invalid category link.');path=u.pathname.slice('/v1.0/'.length)+u.search;
 }
 throw new CategoryError('Too many Outlook categories to sync safely.');
}
export async function queueCurrentLabels(owner:string,id:string){
 const a=await account(owner);if(!a?.enabled||!a.category_enabled||!a.category_permission)throw new ReviewError('Enable Outlook category sync in Assistant & rules first.',409);
 // Never accept message IDs, labels, or another owner from the browser.
 const row=(await db('relay_reviews',{owner_id:'eq.'+owner,id:'eq.'+id,limit:'1'}))[0];if(!row)throw new ReviewError('Saved review not found.',404);
 const item=unseal<ReviewItem>(row.snapshot,snapshotPurpose(owner,row.message_id));
 if(row.topic)item.assessment.topic=row.topic;if(row.priority)item.assessment.priority=row.priority;
 const labels=applyWorkflow(item,(await readWorkflow(owner,a)).data).labels||[];
 if(!labels.length){await db('relay_reviews',{owner_id:'eq.'+owner,id:'eq.'+id},{method:'PATCH',body:JSON.stringify({category_status:'none',category_payload:null,category_revision:randomUUID(),category_error:null})});return {state:'no-matches',labels:[]}}
 await db('relay_reviews',{owner_id:'eq.'+owner,id:'eq.'+id},{method:'PATCH',body:JSON.stringify({category_status:'pending',category_payload:seal(labels,purpose(owner,row.message_id)),category_revision:randomUUID(),category_attempts:0,category_available_at:new Date().toISOString(),category_error:null,category_synced_at:null})});
 return {state:'queued',labels:labels.map(l=>l.name)};
}
export async function syncCategories(owner:string,token:string,lease:string){
 const deadline=Date.now()+40000;
 const active=async()=>{const a=await account(owner);return !!(a?.enabled&&a.category_enabled&&a.category_permission&&a.lock_id===lease&&Date.parse(a.lock_until)>Date.now()+12000)};
 if(Date.now()>deadline||!await active())return;
 const rows=await db('relay_reviews',{owner_id:'eq.'+owner,category_status:'eq.pending',category_available_at:'lte.'+new Date().toISOString(),order:'category_available_at.asc',limit:'5'});
 let names:Map<string,string>|undefined;
 for(const job of rows){
  if(Date.now()>deadline||!await active())break;
  const where={owner_id:'eq.'+owner,id:'eq.'+job.id,category_revision:'eq.'+job.category_revision};
  const update=(v:object)=>db('relay_reviews',where,{method:'PATCH',body:JSON.stringify(v)});
  const attempts=(job.category_attempts||0)+1;
  try{
   if(attempts>3){await update({category_status:'failed',category_error:'Retry limit reached. Retry category sync in Relay.'});continue}
   await update({category_attempts:attempts,category_available_at:new Date(Date.now()+300000).toISOString()});
   const labels=labelsSchema.parse(unseal(job.category_payload,purpose(owner,job.message_id)));
   names??=await masterCategories(token);
   const desired:string[]=[];
   for(const label of labels){
    const key=label.name.toLowerCase();
    if(!names.has(key)){
     if(Date.now()>deadline||!await active())return;
     try{const created=await request(token,'me/outlook/masterCategories','POST',{displayName:label.name,color:colors[label.color]});if(typeof created.displayName!=='string')throw new CategoryError('Microsoft did not confirm category creation.');names.set(key,created.displayName)}
     catch(e){if(e instanceof CategoryError&&e.status===409){names=await masterCategories(token);if(!names.has(key))throw e}else throw e}
    }
    const name=names.get(key)!;if(!desired.includes(name))desired.push(name);
   }
   const path='me/messages/'+encodeURIComponent(job.message_id);
   for(let attempt=0;attempt<3;attempt++){
    const current=await request(token,path+'?$select=id,categories');
    if(!Array.isArray(current.categories)||!current.categories.every((x:unknown)=>typeof x==='string'))throw new CategoryError('Microsoft did not return the email categories.');
    const merged=[...current.categories];for(const name of desired)if(!merged.some(x=>x.toLowerCase()===name.toLowerCase()))merged.push(name);
    if(merged.length===current.categories.length)break;
    const etag=current['@odata.etag'];if(typeof etag!=='string'||!etag)throw new CategoryError('Microsoft did not return a version for this email. No categories were changed.');
    if(Date.now()>deadline||!await active())return;
    try{await request(token,path,'PATCH',{categories:merged},etag);break}
    catch(e){if(e instanceof CategoryError&&e.status===412&&attempt<2)continue;throw e}
   }
   await update({category_status:'done',category_error:null,category_synced_at:new Date().toISOString()});
   await db('relay_mail_accounts',{owner_id:'eq.'+owner,lock_id:'eq.'+lease},{method:'PATCH',body:JSON.stringify({category_last_success:new Date().toISOString(),category_error:null})});
  }catch(error){
   const e=error instanceof CategoryError?error:new CategoryError('Category sync was interrupted. Relay will retry.');
   const revoked=e.status===401||e.status===403;
   await update({category_status:e.status===404?'skipped':revoked||attempts>=3?'failed':'pending',category_error:e.message,category_available_at:new Date(Date.now()+Math.max(e.delay,300*Math.pow(3,attempts-1))*1000).toISOString()});
   await db('relay_mail_accounts',{owner_id:'eq.'+owner,lock_id:'eq.'+lease},{method:'PATCH',body:JSON.stringify({category_error:e.message,...(revoked?{category_enabled:false,category_permission:false}:{})})});
   if(revoked||e.status===429)break;
  }
 }
}
