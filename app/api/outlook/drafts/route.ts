import {z} from 'zod';
import {config,session,json,seal,unseal} from '../../../../lib/outlook';
import {db,ownerOf,account,snapshotPurpose} from '../../../../lib/review-store';
import {readWorkflow} from '../../../../lib/workflow-store';
import {applyWorkflow} from '../../../../lib/workflow';
import {telegramConnection} from '../../../../lib/telegram';
import {ReviewError} from '../../../../lib/review-service';
import type {ReviewItem} from '../../../../lib/inbox-review';
const schema=z.object({action:z.enum(['prepare','save','outlook']),id:z.string().regex(/^\d+$/),revision:z.number().int().min(0),text:z.string().trim().min(1).max(6000).optional()}).strict();
export async function POST(req:Request){try{
 if(req.headers.get('origin')!==config().origin)return json({error:'Invalid origin.'},403);const s=await session();if(!s)return json({error:'Sign in to manage drafts.'},401);const raw=await req.text();if(raw.length>30000)return json({error:'Request too large.'},413);const p=schema.safeParse(JSON.parse(raw));if(!p.success)return json({error:'Invalid draft request.'},400);const owner=ownerOf(s),v=p.data,a=await account(owner),row=(await db('relay_reviews',{owner_id:'eq.'+owner,id:'eq.'+v.id,limit:'1'}))[0];if(!row)throw new ReviewError('Saved review not found.',404);
 if(row.draft_revision!==v.revision)throw new ReviewError('Draft changed. Refresh before saving.',409);
 const where={owner_id:'eq.'+owner,id:'eq.'+v.id,draft_revision:'eq.'+v.revision,draft_status:'eq.'+row.draft_status,draft_outlook_status:'eq.'+row.draft_outlook_status};let update:object;
 if(v.action==='prepare'){
  if(!a?.enabled)throw new ReviewError('Resume automatic processing first.',409);if(!['none','failed'].includes(row.draft_status))throw new ReviewError('A draft already exists or is being prepared.',409);
  const item=applyWorkflow(unseal<ReviewItem>(row.snapshot,snapshotPurpose(owner,row.message_id)),(await readWorkflow(owner,a)).data);if(!item.draftRequested)throw new ReviewError('Enable Prepare reply draft on a matching rule first.',409);
  const t=await telegramConnection(owner);
  update={...(t?.enabled&&t.chat_cipher?{notify_status:'pending',notify_attempts:0,notify_available_at:new Date().toISOString(),notify_error:null}:{}),snapshot:seal(item,snapshotPurpose(owner,row.message_id)),draft_status:'pending',draft_attempts:0,draft_error:null,draft_available_at:new Date().toISOString()};
 }else{
  if(row.draft_status!=='ready')throw new ReviewError('Wait for the draft to be ready.',409);
  if(!['none','failed'].includes(row.draft_outlook_status))throw new ReviewError('An Outlook copy is queued or already created. Edit that copy in Outlook.',409);
  if(v.action==='save'){if(!v.text)throw new ReviewError('Enter draft text.',400);const draft=unseal<any>(row.draft_cipher,'draft:'+owner+':'+row.message_id);update={draft_cipher:seal({...draft,text:v.text},'draft:'+owner+':'+row.message_id)}}
  else{if(!a?.enabled||!a.category_permission)throw new ReviewError('Enable automatic processing and Outlook write access first.',409);update={draft_outlook_status:'pending',draft_error:null}}
 }
 const changed=await db('relay_reviews',where,{method:'PATCH',body:JSON.stringify({...update,draft_revision:v.revision+1})});if(!changed.length)throw new ReviewError('Draft changed. Refresh and try again.',409);return json({ok:true});
}catch(e){return json({error:e instanceof ReviewError?e.message:'Could not update draft.'},e instanceof ReviewError?e.status:503)}}
