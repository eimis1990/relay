import {z} from 'zod';
import {session,config,json} from '../../../../lib/outlook';
import {ownerOf,account,db} from '../../../../lib/review-store';
import {queueCurrentLabels} from '../../../../lib/outlook-categories';
import {runMailbox} from '../../../../lib/mail-worker';
import {ReviewError} from '../../../../lib/review-service';
export const dynamic='force-dynamic';export const maxDuration=180;
export async function GET(){try{
 const s=await session();if(!s)return json({error:'Sign in to manage Outlook categories.'},401);const owner=ownerOf(s),a=await account(owner);
 const jobs=await db('relay_reviews',{owner_id:'eq.'+owner,category_status:'in.(pending,failed)',select:'category_status',limit:'1000'});
 return json({connected:!!a?.refresh_cipher,processing:!!a?.enabled,permission:!!a?.category_permission,enabled:!!a?.category_enabled,lastSuccess:a?.category_last_success,error:a?.category_error,pending:jobs.filter((x:any)=>x.category_status==='pending').length,failed:jobs.filter((x:any)=>x.category_status==='failed').length});
}catch(e){return json({error:e instanceof ReviewError?e.message:'Category status is unavailable.'},e instanceof ReviewError?e.status:503)}}
const schema=z.discriminatedUnion('action',[z.object({action:z.literal('apply'),id:z.string().regex(/^\d+$/)}).strict(),z.object({action:z.enum(['pause','resume','check','retry'])}).strict()]);
export async function POST(req:Request){try{
 if(req.headers.get('origin')!==config().origin)return json({error:'Invalid origin.'},403);const s=await session();if(!s)return json({error:'Sign in to manage Outlook categories.'},401);
 const text=await req.text();if(text.length>2000)return json({error:'Request too large.'},413);let parsed;try{parsed=schema.safeParse(JSON.parse(text))}catch{return json({error:'Invalid JSON.'},400)}if(!parsed.success)return json({error:'Invalid request.'},400);
 const owner=ownerOf(s),a=await account(owner);if(!a)return json({error:'Connect automatic processing first.'},409);
 const p=parsed.data;
 if(p.action==='pause'){await db('relay_mail_accounts',{owner_id:'eq.'+owner},{method:'PATCH',body:JSON.stringify({category_enabled:false})});return json({ok:true})}
 if(!a.category_permission||!a.refresh_cipher)return json({error:'Connect Outlook category sync first.'},409);
 if(p.action==='resume'){await db('relay_mail_accounts',{owner_id:'eq.'+owner},{method:'PATCH',body:JSON.stringify({category_enabled:true,category_error:null})});return json({ok:true})}
 if(p.action==='apply')return json(await queueCurrentLabels(owner,p.id));
 if(p.action==='retry'){await db('relay_reviews',{owner_id:'eq.'+owner,category_status:'eq.failed'},{method:'PATCH',body:JSON.stringify({category_status:'pending',category_attempts:0,category_available_at:new Date().toISOString(),category_error:null})});return json({ok:true})}
 return json(await runMailbox(owner));
}catch(e){return json({error:e instanceof ReviewError?e.message:'Could not update category sync.'},e instanceof ReviewError?e.status:503)}}
