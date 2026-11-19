import {z} from 'zod';
import {session,config,json,seal} from '../../../../lib/outlook';
import {ownerOf,account,db,pendingCounts} from '../../../../lib/review-store';
import {runMailbox,setupScheduler,DAILY_LIMIT} from '../../../../lib/mail-worker';
import {ReviewError} from '../../../../lib/review-service';
export const dynamic='force-dynamic';export const maxDuration=180;
export async function GET(){try{const s=await session();if(!s)return json({error:'Sign in to view processing status.'},401);const owner=ownerOf(s),a=await account(owner);return json({connected:!!a?.refresh_cipher,enabled:!!a?.enabled,startedAt:a?.started_at,lastChecked:a?.last_polled_at,lastSuccess:a?.last_success_at,error:a?.last_error,language:a?.language,limit:DAILY_LIMIT,...await pendingCounts(owner)})}catch(e){return json({error:e instanceof ReviewError?e.message:'Processing status unavailable.'},503)}}
const schema=z.object({action:z.enum(['pause','resume','check','settings','retry','disconnect']),language:z.enum(['English','Lithuanian']).optional(),priorityRules:z.string().max(1200).optional()}).strict();
export async function POST(req:Request){try{if(req.headers.get('origin')!==config().origin)return json({error:'Invalid origin.'},403);const s=await session();if(!s)return json({error:'Sign in to manage processing.'},401);const text=await req.text();if(text.length>5000)return json({error:'Request too large.'},413);const p=schema.safeParse(JSON.parse(text));if(!p.success)return json({error:'Invalid request.'},400);const owner=ownerOf(s),a=await account(owner);if(!a)return json({error:'Connect automatic processing first.'},409);
 if(p.data.action==='check')return json(await runMailbox(owner));
 if(p.data.action==='retry'){await db('relay_review_queue',{owner_id:'eq.'+owner,status:'eq.failed'},{method:'PATCH',body:JSON.stringify({status:'pending',attempts:0,available_at:new Date().toISOString(),last_error:null})});return json({ok:true})}
 if(p.data.action==='resume'){if(!a.refresh_cipher)return json({error:'Reconnect Outlook first.'},409);await setupScheduler()}
 const update=p.data.action==='disconnect'?{enabled:false,refresh_cipher:null,last_error:null,category_enabled:false,category_permission:false,category_error:null}:p.data.action==='settings'?{...(p.data.language?{language:p.data.language}:{}),...(p.data.priorityRules!==undefined?{rules_cipher:seal(p.data.priorityRules,'rules:'+owner)}:{})}:{enabled:p.data.action==='resume'};
 await db('relay_mail_accounts',{owner_id:'eq.'+owner},{method:'PATCH',body:JSON.stringify(update)});return json({ok:true});
 }catch(e){return json({error:e instanceof ReviewError?e.message:'Could not update processing.'},e instanceof ReviewError?e.status:503)}}
