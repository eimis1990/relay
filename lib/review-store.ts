import 'server-only';
import {randomUUID} from 'node:crypto';
import { seal, unseal, type Session } from './outlook';
import { databaseConfigured } from './workspace-store';
import { ReviewError } from './review-service';
import type { ReviewItem } from './inbox-review';

export function ownerOf(s:Session){if(!s.subject)throw new ReviewError('Reconnect Microsoft to use saved reviews.',401);return 'microsoft:'+s.subject}
export async function db(table:string,query:Record<string,string>={},init:RequestInit={},prefer='return=representation'){
  if(!databaseConfigured())throw new ReviewError('Review storage is not configured.',503);
  const base=new URL(process.env.SUPABASE_URL!);
  if((base.protocol!=='https:'&&!(base.protocol==='http:'&&['localhost','127.0.0.1'].includes(base.hostname)))||base.pathname!=='/'||base.username||base.password)throw new ReviewError('Invalid storage configuration.',503);
  const key=process.env.SUPABASE_SECRET_KEY!;
  const r=await fetch(new URL('/rest/v1/'+table+'?'+new URLSearchParams(query),base),{...init,headers:{apikey:key,...(!key.startsWith('sb_secret_')?{Authorization:'Bearer '+key}:{}),'Content-Type':'application/json',Prefer:prefer},cache:'no-store',signal:AbortSignal.timeout(10000)});
  if(r.status===409)throw new ReviewError('This record changed in another tab. Reload before saving.',409);
  if(!r.ok)throw new ReviewError('Review storage is temporarily unavailable. Please retry.',503);
  return r.status===204?null:await r.json();
}
export type Account={owner_id:string;email:string;refresh_cipher:string|null;enabled:boolean;started_at:string|null;cursor_cipher:string|null;language:'English'|'Lithuanian';rules_cipher:string|null;lock_id:string|null;lock_until:string;last_polled_at:string|null;last_success_at:string|null;last_error:string|null;budget_date:string;budget_used:number;category_enabled?:boolean;category_permission?:boolean;category_error?:string|null;category_last_success?:string|null};
export async function account(owner:string):Promise<Account|null>{return (await db('relay_mail_accounts',{owner_id:'eq.'+owner,limit:'1'}))[0]??null}
export const snapshotPurpose=(owner:string,message:string)=>'review:'+owner+':'+message;
export async function saveReview(owner:string,item:ReviewItem,language:string,reviewedAt:string){
  const a=item.labels?.length?await account(owner):null;
  const categories=!!(a?.enabled&&a.category_enabled&&a.category_permission);
  // The sync request is committed with the review: a crash cannot lose the Outlook update.
  const categoryData=categories?{category_status:'pending',category_payload:seal(item.labels,'categories:'+owner+':'+item.email.id),category_revision:randomUUID(),category_available_at:new Date().toISOString()}:{};
  // Ignore conflicts: re-reading an email must not overwrite a person's checked state or corrections.
  await db('relay_reviews',{on_conflict:'owner_id,message_id'},{method:'POST',body:JSON.stringify({owner_id:owner,message_id:item.email.id,snapshot:seal(item,snapshotPurpose(owner,item.email.id)),language,reviewed_at:reviewedAt,...categoryData})},'resolution=ignore-duplicates,return=representation');
}
export async function history(owner:string,before?:string){
  const rows=await db('relay_reviews',{owner_id:'eq.'+owner,order:'id.desc',limit:'51',...(before?{id:'lt.'+before}:{})});
  const more=rows.length>50;const page=rows.slice(0,50);
  return {items:page.map((r:any)=>{const item=unseal<ReviewItem>(r.snapshot,snapshotPurpose(owner,r.message_id));return {...item,assessment:{...item.assessment,...(r.priority?{priority:r.priority}:{}),...(r.topic?{topic:r.topic}:{})},savedId:String(r.id),revision:r.revision,checked:r.checked,reviewedAt:r.reviewed_at,language:r.language,categoryStatus:r.category_status||'none',categoryError:r.category_error,categorySyncedAt:r.category_synced_at}}),nextCursor:more?String(page.at(-1).id):null};
}
export async function pendingCounts(owner:string){
  const rows=await db('relay_review_queue',{owner_id:'eq.'+owner,status:'in.(pending,failed)',select:'status',limit:'1000'});
  return {pending:rows.filter((r:any)=>r.status==='pending').length,failed:rows.filter((r:any)=>r.status==='failed').length};
}
