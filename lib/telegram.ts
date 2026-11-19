import 'server-only';
import {z} from 'zod';
import {randomBytes,randomUUID,createHash} from 'node:crypto';
import {config,seal,unseal} from './outlook';
import {db,snapshotPurpose,account} from './review-store';
import {ReviewError} from './review-service';
import type {ReviewItem} from './inbox-review';
export const notificationSchema=z.object({allReviews:z.boolean(),highPriority:z.boolean(),needsAttention:z.boolean(),draftReady:z.boolean(),includeSummary:z.boolean()}).strict();
export const defaultNotifications={allReviews:true,highPriority:false,needsAttention:false,draftReady:false,includeSummary:false};
export const hash=(value:string)=>createHash('sha256').update(value).digest('hex');
export async function telegramConnection(owner:string){return (await db('relay_telegram',{owner_id:'eq.'+owner,limit:'1'}))[0]||null}
export class TelegramError extends Error{constructor(public code:number,public retry=300){super('Telegram request failed.')}}
export async function telegramCall(token:string,method:'getMe'|'getWebhookInfo'|'setWebhook'|'sendMessage',body:object={}){
 // Never log provider URLs/errors: Telegram places the credential in its URL.
 let r;try{r=await fetch(`https://api.telegram.org/bot${token}/${method}`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body),cache:'no-store',signal:AbortSignal.timeout(10000)})}catch{throw new TelegramError(0)}
 const j=await r.json().catch(()=>({}));if(!r.ok||!j.ok)throw new TelegramError(Number(j.error_code)||r.status,Math.max(30,Math.min(86400,Number(j.parameters?.retry_after)||300)));return j.result;
}
export async function connectTelegram(owner:string,token:string){
 if(!/^\d{5,16}:[A-Za-z0-9_-]{25,100}$/.test(token))throw new ReviewError('Enter the bot token from BotFather.',400);
 const existing=await telegramConnection(owner);if(existing?.chat_cipher)throw new ReviewError('Disconnect the current bot before replacing it.',409);
 const me=await telegramCall(token,'getMe');if(!me.is_bot||!/^\w{5,32}$/.test(me.username||''))throw new ReviewError('Telegram did not return a valid bot.',400);
 const id=existing?.id||randomUUID(),url=config().origin+'/api/telegram/webhook/'+id,info=await telegramCall(token,'getWebhookInfo');
 if(info.url&&info.url!==url)throw new ReviewError('This bot is connected to another service. Create a dedicated Relay bot.',409);
 const secret=randomBytes(32).toString('base64url');
 await db('relay_telegram',{on_conflict:'owner_id'},{method:'POST',body:JSON.stringify({owner_id:owner,id,bot_id:String(me.id),bot_username:me.username,token_cipher:seal(token,'telegram-token:'+owner),webhook_hash:hash(secret),enabled:false,chat_cipher:null,pending_chat_cipher:null,pending_name:null,link_hash:null,last_error:null})},'resolution=merge-duplicates,return=representation');
 try{await telegramCall(token,'setWebhook',{url,secret_token:secret,allowed_updates:['message']})}catch{throw new ReviewError('Could not configure Telegram. Check the token and try connecting again.',502)}
 return createTelegramLink(owner);
}
export async function createTelegramLink(owner:string){const t=await telegramConnection(owner);if(!t?.token_cipher)throw new ReviewError('Connect a bot first.',409);if(t.chat_cipher)throw new ReviewError('Your private chat is already linked.',409);const code=randomBytes(24).toString('base64url');await db('relay_telegram',{owner_id:'eq.'+owner},{method:'PATCH',body:JSON.stringify({link_hash:hash(code),link_expires:new Date(Date.now()+600000).toISOString(),pending_chat_cipher:null,pending_name:null})});return {link:'https://t.me/'+t.bot_username+'?start='+code}}
export async function notifyReviews(owner:string,lease:string,deadline:number){
 let t=await telegramConnection(owner);if(!t?.enabled||!t.chat_cipher||!t.token_cipher)return;
 // A terminated send may have succeeded. Do not automatically duplicate it.
 await db('relay_reviews',{owner_id:'eq.'+owner,notify_status:'eq.sending'},{method:'PATCH',body:JSON.stringify({notify_status:'uncertain',notify_error:'Delivery could not be confirmed. Check Telegram before retrying.'})});
 const rows=await db('relay_reviews',{owner_id:'eq.'+owner,notify_status:'eq.pending',notify_available_at:'lte.'+new Date().toISOString(),order:'id.asc',limit:'10'});
 for(const row of rows){
  if(Date.now()>deadline-12000)break;
  const a=await account(owner);t=await telegramConnection(owner);if(!a?.enabled||a.lock_id!==lease||Date.parse(a.lock_until)<Date.now()+15000||!t?.enabled||!t.chat_cipher)return;
  if(['pending','generating'].includes(row.draft_status))continue;
  const p=notificationSchema.parse(t.preferences),item=unseal<ReviewItem>(row.snapshot,snapshotPurpose(owner,row.message_id));
  const priority=row.priority||item.assessment.priority,ready=row.draft_status==='ready';
  if(!(p.allReviews||p.highPriority&&priority==='High'||p.needsAttention&&item.assessment.needsAttention||p.draftReady&&ready)){await db('relay_reviews',{owner_id:'eq.'+owner,id:'eq.'+row.id},{method:'PATCH',body:JSON.stringify({notify_status:'skipped'})});continue}
  const claimed=await db('relay_reviews',{owner_id:'eq.'+owner,id:'eq.'+row.id,notify_status:'eq.pending'},{method:'PATCH',body:JSON.stringify({notify_status:'sending',notify_attempts:(row.notify_attempts||0)+1})});if(!claimed.length)continue;
  const text=[priority==='High'?'🔴 High-priority review':'✅ Email reviewed',item.email.subject.slice(0,180),`${priority} · ${item.assessment.topic}`,ready?'📝 Reply draft ready':row.draft_status==='failed'?'🛑 Reply draft needs retry':'',item.assessment.needsAttention?'👀 Needs your attention':'',p.includeSummary?item.assessment.summary.slice(0,550):''].filter(Boolean).join('\n');
  try{await telegramCall(unseal(t.token_cipher,'telegram-token:'+owner),'sendMessage',{chat_id:unseal(t.chat_cipher,'telegram-chat:'+owner),text,link_preview_options:{is_disabled:true},reply_markup:{inline_keyboard:[[{text:'Open review in Relay',url:config().origin+'/#review'}]]}});
   await db('relay_reviews',{owner_id:'eq.'+owner,id:'eq.'+row.id},{method:'PATCH',body:JSON.stringify({notify_status:'done',notify_error:null})});await db('relay_telegram',{owner_id:'eq.'+owner},{method:'PATCH',body:JSON.stringify({last_sent_at:new Date().toISOString(),last_error:null})});
  }catch(e){const code=e instanceof TelegramError?e.code:0,retry=e instanceof TelegramError?e.retry:300;await db('relay_reviews',{owner_id:'eq.'+owner,id:'eq.'+row.id},{method:'PATCH',body:JSON.stringify({notify_status:code===429&&(row.notify_attempts||0)<2?'pending':code===0||code>=500?'uncertain':'failed',notify_available_at:new Date(Date.now()+retry*1000).toISOString(),notify_error:code===403?'Bot blocked. Unblock it in Telegram.':'Notification delivery failed or could not be confirmed. Check Telegram.'})});await db('relay_telegram',{owner_id:'eq.'+owner},{method:'PATCH',body:JSON.stringify({last_error:'A notification needs attention. Check delivery status in Review inbox.',...(code===403?{enabled:false}:{})})});break}
 }
}
