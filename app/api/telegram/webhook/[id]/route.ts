import {timingSafeEqual} from 'node:crypto';
import {db} from '../../../../../lib/review-store';
import {hash} from '../../../../../lib/telegram';
import {seal,json} from '../../../../../lib/outlook';
export async function POST(req:Request,{params}:{params:Promise<{id:string}>}){try{
 const {id}=await params;if(!/^[a-f0-9-]{36}$/.test(id))return json({ok:false},404);
 const t=(await db('relay_telegram',{id:'eq.'+id,limit:'1'}))[0],secret=req.headers.get('x-telegram-bot-api-secret-token')||'';
 if(!t?.webhook_hash||!secret||secret.length>256||!timingSafeEqual(Buffer.from(hash(secret)),Buffer.from(t.webhook_hash)))return json({ok:false},403);
 const raw=await req.text();if(raw.length>64000)return json({ok:false},413);const m=JSON.parse(raw).message;
 if(t.chat_cipher||!m||m.chat?.type!=='private'||!Number.isSafeInteger(m.chat.id)||m.chat.id<=0||m.from?.id!==m.chat.id||m.from?.is_bot)return json({ok:true});
 const match=typeof m.text==='string'&&/^\/start ([A-Za-z0-9_-]{32})$/.exec(m.text);if(!match||!t.link_hash||Date.parse(t.link_expires)<Date.now()||hash(match[1])!==t.link_hash)return json({ok:true});
 await db('relay_telegram',{id:'eq.'+id,link_hash:'eq.'+t.link_hash},{method:'PATCH',body:JSON.stringify({pending_chat_cipher:seal(String(m.chat.id),'telegram-chat:'+t.owner_id),pending_name:[m.from.first_name,m.from.last_name,m.from.username?'@'+m.from.username:''].filter(x=>typeof x==='string').join(' ').slice(0,200)||'Private Telegram account',link_hash:null})});return json({ok:true});
}catch{return json({ok:false},503)}}
