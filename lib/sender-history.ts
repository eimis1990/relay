import {emailForReview, safeOutlookLink, type ReviewItem} from './inbox-review';

export type HistoryWindow = 'off' | '6months' | '12months';
export type HistorySource = {id:string; subject:string; date:string; direction:'received'|'sent'; sameThread:boolean; webLink:string|null; excerptLimited:boolean};
export type DraftContext = {window:HistoryWindow; checkedAt:string; since:string|null; sources:HistorySource[]; limited:boolean; candidates:number};
type ContextMessage = HistorySource & {text:string};
export type SenderHistory = {context:DraftContext; messages:ContextMessage[]};
const ROOT='https://graph.microsoft.com/v1.0';
const MAX_CANDIDATES=60, MAX_MESSAGES=12, MAX_TEXT=24000;
const metadata='id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,conversationId,parentFolderId,isDraft,webLink';
const address=(v:any)=>String(v?.emailAddress?.address||'').trim().toLowerCase();
export function monthsBefore(now:Date,months:number){
 const d=new Date(now),day=d.getUTCDate();d.setUTCDate(1);d.setUTCMonth(d.getUTCMonth()-months);
 const last=new Date(Date.UTC(d.getUTCFullYear(),d.getUTCMonth()+1,0)).getUTCDate();d.setUTCDate(Math.min(day,last));return d;
}
export function emptyHistory(window:HistoryWindow='off',now=new Date()):SenderHistory{
 return {context:{window,checkedAt:now.toISOString(),since:window==='off'?null:monthsBefore(now,window==='6months'?6:12).toISOString(),sources:[],limited:false,candidates:0},messages:[]};
}
function terms(subject:string){return new Set(subject.toLowerCase().replace(/\b(re|fw|fwd)\s*:/g,'').match(/[\p{L}\p{N}]{3,}/gu)||[])}
// Exact address checks are repeated locally; Graph search can also match aliases or names.
export function selectHistory(incoming:any[],sent:any[],item:ReviewItem,thread:string,since:string,until:string,excluded:Set<string>){
 const sender=item.email.senderEmail.trim().toLowerCase(),words=terms(item.email.subject),seen=new Set([item.email.id]);
 const candidates:{raw:any;source:HistorySource;score:number}[]=[];
 for(const [direction,rows] of [['received',incoming],['sent',sent]] as const)for(const m of rows){
  const date=direction==='sent'?m.sentDateTime:m.receivedDateTime;
  if(typeof m.id!=='string'||seen.has(m.id)||m.isDraft||excluded.has(m.parentFolderId)||!Number.isFinite(Date.parse(date))||Date.parse(date)<Date.parse(since)||Date.parse(date)>Date.parse(until))continue;
  if(direction==='received'?address(m.from)!==sender:![...(m.toRecipients||[]),...(m.ccRecipients||[])].some(r=>address(r)===sender))continue;
  seen.add(m.id);const sameThread=!!thread&&m.conversationId===thread,subject=String(m.subject||'(No subject)').slice(0,500);
  const overlap=[...terms(subject)].filter(w=>words.has(w)).length;
  candidates.push({raw:m,source:{id:m.id,subject,date,direction,sameThread,webLink:safeOutlookLink(m.webLink),excerptLimited:false},score:(sameThread?100:0)+Math.min(overlap,10)*3});
 }
 const recent=[...candidates].sort((a,b)=>Date.parse(b.source.date)-Date.parse(a.source.date));
 // Keep recent exchanges as well as the strongest thread/subject matches.
 const ranked=[...candidates].sort((a,b)=>b.score-a.score||Date.parse(b.source.date)-Date.parse(a.source.date));
 const chosen=new Map([...ranked.slice(0,8),...recent.slice(0,4),...ranked].map(c=>[c.source.id,c]));
 // A Set preserves the first occurrence; Map replacement above must not change order.
 const selected=[...chosen.values()].slice(0,MAX_MESSAGES).sort((a,b)=>Date.parse(a.source.date)-Date.parse(b.source.date));
 return {selected,count:candidates.length};
}
export async function loadSenderHistory(token:string,item:ReviewItem,window:HistoryWindow,now=new Date()):Promise<SenderHistory>{
 const result=emptyHistory(window,now);if(window==='off')return result;
 const sender=item.email.senderEmail.trim().toLowerCase();
 if(!/^[a-z0-9._%+'-]+@[a-z0-9.-]+\.[a-z]{2,}$/i.test(sender))throw Error('Sender history could not be checked: the sender address is missing or unsupported.');
 const signal=AbortSignal.timeout(20000);
 const get=async(path:string)=>{const r=await fetch(ROOT+path,{headers:{Authorization:'Bearer '+token,Prefer:'IdType="ImmutableId", outlook.body-content-type="text"'},cache:'no-store',signal});if(!r.ok)throw Error('Sender history is temporarily unavailable. Relay will retry before drafting.');return r.json()};
 const incomingQuery=new URLSearchParams({$select:metadata,$top:String(MAX_CANDIDATES),$orderby:'receivedDateTime desc',$filter:`receivedDateTime ge ${result.context.since} and receivedDateTime le ${now.toISOString()} and from/emailAddress/address eq '${sender.replace(/'/g,"''")}'`});
 const sentQuery=new URLSearchParams({$select:metadata,$top:String(MAX_CANDIDATES),$search:'"recipients:'+sender+'"'});
 const [incoming,sent,original,...folders]=await Promise.all([
  get('/me/messages?'+incomingQuery),get('/me/mailFolders/sentitems/messages?'+sentQuery),
  get('/me/messages/'+encodeURIComponent(item.email.id)+'?$select=id,conversationId'),
  ...['deleteditems','junkemail','drafts'].map(name=>get('/me/mailFolders/'+name+'?$select=id')),
 ]);
 if(!Array.isArray(incoming.value)||!Array.isArray(sent.value)||folders.some(f=>typeof f.id!=='string'))throw Error('Sender history returned an invalid response. Relay will retry before drafting.');
 const excluded=new Set<string>(folders.map(f=>f.id));
 const {selected,count}=selectHistory(incoming.value,sent.value,item,original.conversationId||'',result.context.since!,now.toISOString(),excluded);
 result.context.candidates=count;
 result.context.limited=!!incoming['@odata.nextLink']||!!sent['@odata.nextLink']||incoming.value.length>=MAX_CANDIDATES||sent.value.length>=MAX_CANDIDATES||count>MAX_MESSAGES;
 const bodies=await Promise.all(selected.map(c=>get('/me/messages/'+encodeURIComponent(c.source.id)+'?'+new URLSearchParams({$select:metadata+',body,bodyPreview,hasAttachments,isRead'}))));
 let remaining=MAX_TEXT;
 for(let i=0;i<selected.length;i++){
  const candidate=selected[i],m=bodies[i];
  // Recheck identity, date and folder after loading the body, in case a message moved.
  const checked=selectHistory(candidate.source.direction==='received'?[m]:[],candidate.source.direction==='sent'?[m]:[],item,original.conversationId||'',result.context.since!,now.toISOString(),excluded).selected[0];
  if(!checked||m.id!==candidate.source.id)throw Error('Sender history changed while loading. Relay will retry before drafting.');
  const email=emailForReview(m),text=email.text.slice(0,Math.min(2000,remaining));remaining-=text.length;
  if(!text){result.context.limited=true;continue}
  const source={...checked.source,excerptLimited:email.previewOnly||email.truncated||email.text.length>text.length||email.hasAttachments};
  if(source.excerptLimited)result.context.limited=true;
  result.context.sources.push(source);result.messages.push({...source,text});
 }
 return result;
}
