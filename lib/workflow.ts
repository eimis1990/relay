import {z} from 'zod';
import type {ReviewItem} from './inbox-review';
export const labelColors=['green','teal','blue','violet','amber','coral'] as const;
export const ruleFields=['sender','subject','body','company','topic'] as const;
export const workflowSchema=z.object({
 language:z.enum(['English','Lithuanian']),instructions:z.string().trim().max(1200),
 schedule:z.object({mode:z.enum(['continuous','daily']),time:z.string().regex(/^([01]\d|2[0-3]):[0-5]\d$/),timezone:z.string().min(1).max(80).refine(s=>{try{new Intl.DateTimeFormat('en',{timeZone:s});return true}catch{return false}},'Choose a valid IANA time zone.')}).strict(),
 labels:z.array(z.object({id:z.string().regex(/^[a-z0-9_-]{1,40}$/),name:z.string().trim().min(1).max(40),color:z.enum(labelColors),field:z.enum(ruleFields),value:z.string().trim().min(1).max(200),priority:z.enum(['unchanged','High','Normal','Low']),enabled:z.boolean(),draft:z.boolean().optional(),draftInstructions:z.string().trim().max(600).optional()}).strict()).max(12),
}).strict().refine(w=>new Set(w.labels.map(l=>l.id)).size===w.labels.length,'Rule IDs must be unique.');
export type Workflow=z.infer<typeof workflowSchema>;
export const defaultWorkflow:Workflow={language:'English',instructions:'',schedule:{mode:'continuous',time:'07:00',timezone:'Europe/Vilnius'},labels:[]};
function localStamp(date:Date,timezone:string){const parts=new Intl.DateTimeFormat('en-CA',{timeZone:timezone,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).formatToParts(date);const p=Object.fromEntries(parts.map(x=>[x.type,x.value]));return `${p.year}-${p.month}-${p.day}T${p.hour}:${p.minute}`}
// Wall-clock comparison naturally handles DST gaps/repeated hours without a fixed UTC offset.
// Messages arriving after today's scheduled cutoff wait until the next local day's cutoff.
export function eligibleAt(receivedAt:string,w:Workflow,now=new Date()){
 if(w.schedule.mode==='continuous')return true;
 const current=localStamp(now,w.schedule.timezone);if(current.slice(11)<w.schedule.time)return false;
 if(!Number.isFinite(Date.parse(receivedAt)))return false;
 return localStamp(new Date(receivedAt),w.schedule.timezone)<=current.slice(0,10)+'T'+w.schedule.time;
}
export function applyWorkflow(item:ReviewItem,w:Workflow):ReviewItem{
 const matches=w.labels.filter(l=>l.enabled&&({sender:item.email.sender+' '+item.email.senderEmail,subject:item.email.subject,body:item.email.text,company:item.assessment.company||'',topic:item.assessment.topic}[l.field]).toLocaleLowerCase().includes(l.value.toLocaleLowerCase()));
 const priority=matches.find(l=>l.priority!=='unchanged');
 return {...item,draftRequested:matches.some(l=>l.draft),draftInstructions:matches.filter(l=>l.draft).map(l=>l.draftInstructions||'').filter(Boolean).join('\n').slice(0,1200),labels:matches.map(l=>({name:l.name,color:l.color})),assessment:priority?{...item.assessment,priority:priority.priority as 'High'|'Normal'|'Low',priorityReason:`Your “${priority.name}” rule sets this priority. AI assessment: ${item.assessment.priorityReason}`.slice(0,400)}:item.assessment};
}
export const workflowFormat={type:'json_schema',name:'workflow_proposal',strict:true,schema:{type:'object',additionalProperties:false,required:['reply','workflow'],properties:{reply:{type:'string'},workflow:{type:'object',additionalProperties:false,required:['language','instructions','schedule','labels'],properties:{language:{type:'string',enum:['English','Lithuanian']},instructions:{type:'string'},schedule:{type:'object',additionalProperties:false,required:['mode','time','timezone'],properties:{mode:{type:'string',enum:['continuous','daily']},time:{type:'string'},timezone:{type:'string'}}},labels:{type:'array',items:{type:'object',additionalProperties:false,required:['id','name','color','field','value','priority','enabled','draft','draftInstructions'],properties:{id:{type:'string'},name:{type:'string'},color:{type:'string',enum:labelColors},field:{type:'string',enum:ruleFields},value:{type:'string'},priority:{type:'string',enum:['unchanged','High','Normal','Low']},enabled:{type:'boolean'},draft:{type:'boolean'},draftInstructions:{type:'string'}}}}}}}}};
