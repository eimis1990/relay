import {z} from 'zod';
import {session,config,json,unseal} from '../../../lib/outlook';
import {ownerOf,db} from '../../../lib/review-store';
import {telegramConnection,connectTelegram,createTelegramLink,notificationSchema,telegramCall,defaultNotifications} from '../../../lib/telegram';
import {ReviewError} from '../../../lib/review-service';
export const dynamic='force-dynamic';
export async function GET(){try{const s=await session();if(!s)return json({error:'Sign in to manage Telegram.'},401);const t=await telegramConnection(ownerOf(s));return json({configured:!!t?.token_cipher,bot:t?.bot_username,linked:!!t?.chat_cipher,chat:t?.chat_name,pending:t?.pending_name,enabled:!!t?.enabled,preferences:t?.preferences||defaultNotifications,lastSent:t?.last_sent_at,error:t?.last_error})}catch{return json({error:'Telegram status unavailable.'},503)}}
const schema=z.discriminatedUnion('action',[z.object({action:z.literal('connect'),token:z.string().max(150)}).strict(),z.object({action:z.literal('preferences'),enabled:z.boolean(),preferences:notificationSchema}).strict(),z.object({action:z.enum(['link','confirm','test','disconnect'])}).strict()]);
export async function POST(req:Request){try{
 if(req.headers.get('origin')!==config().origin)return json({error:'Invalid origin.'},403);const s=await session();if(!s)return json({error:'Sign in to manage Telegram.'},401);const raw=await req.text();if(raw.length>3000)return json({error:'Request too large.'},413);const p=schema.safeParse(JSON.parse(raw));if(!p.success)return json({error:'Invalid request.'},400);const owner=ownerOf(s),v=p.data;
 if(v.action==='connect')return json(await connectTelegram(owner,v.token.trim()));
 if(v.action==='link')return json(await createTelegramLink(owner));
 const t=await telegramConnection(owner);if(!t?.token_cipher)throw new ReviewError('Connect your bot first.',409);
 if(v.action==='preferences'){if(v.enabled&&!t.chat_cipher)throw new ReviewError('Confirm your private chat first.',409);await db('relay_telegram',{owner_id:'eq.'+owner},{method:'PATCH',body:JSON.stringify({preferences:v.preferences,enabled:v.enabled})});return json({ok:true})}
 if(v.action==='confirm'){if(!t.pending_chat_cipher||Date.parse(t.link_expires)<Date.now())throw new ReviewError('Open a fresh private link and press Start in Telegram first.',409);const rows=await db('relay_telegram',{owner_id:'eq.'+owner,pending_chat_cipher:'eq.'+t.pending_chat_cipher},{method:'PATCH',body:JSON.stringify({chat_cipher:t.pending_chat_cipher,chat_name:t.pending_name,pending_chat_cipher:null,pending_name:null,link_hash:null,enabled:true})});if(!rows.length)throw new ReviewError('The pending chat changed. Refresh and try again.',409);return json({ok:true})}
 if(v.action==='disconnect'){await db('relay_telegram',{owner_id:'eq.'+owner},{method:'PATCH',body:JSON.stringify({enabled:false,token_cipher:null,webhook_hash:null,chat_cipher:null,chat_name:null,pending_chat_cipher:null,pending_name:null,link_hash:null})});return json({ok:true})}
 if(!t.chat_cipher)throw new ReviewError('Link your private chat first.',409);
 await telegramCall(unseal(t.token_cipher,'telegram-token:'+owner),'sendMessage',{chat_id:unseal(t.chat_cipher,'telegram-chat:'+owner),text:'✅ Relay is connected. Review notifications will follow your saved preferences. No email has been sent.',link_preview_options:{is_disabled:true}});return json({ok:true});
}catch(e){return json({error:e instanceof ReviewError?e.message:'Telegram could not complete this request. Check the bot token and connection.'},e instanceof ReviewError?e.status:502)}}
