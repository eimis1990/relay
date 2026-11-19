import { z } from 'zod';

export const MAX_REVIEW_MESSAGES = 5;
export const MAX_EMAIL_TEXT = 8000;
export const topics = ['Quote request','Order','Delivery','Invoice','Complaint','Supplier','General'] as const;
export const priorities = ['High','Normal','Low'] as const;
export const reviewRequestSchema = z.object({
  source: z.enum(['outlook','demo']),
  ids: z.array(z.string().min(1).max(1500)).max(MAX_REVIEW_MESSAGES),
  language: z.enum(['English','Lithuanian']),
  priorityRules: z.string().trim().max(1200),
  consent: z.literal(true),
}).strict().refine(v => v.source === 'demo' ? v.ids.length === 0 : v.ids.length > 0 && new Set(v.ids).size === v.ids.length);
export type ReviewRequest = z.infer<typeof reviewRequestSchema>;
export type InboxMessage = {
  id:string; subject:string; sender:string; senderEmail:string; receivedAt:string;
  preview:string; isRead:boolean; hasAttachments:boolean; webLink:string|null;
};
export type ReviewEmail = InboxMessage & { text:string; truncated:boolean; previewOnly:boolean };
export const assessmentSchema = z.object({
  id:z.string().min(1).max(1500), company:z.string().max(160).nullable(),
  companyEvidence:z.string().max(300).nullable(), topic:z.enum(topics), priority:z.enum(priorities),
  summary:z.string().min(1).max(650), priorityReason:z.string().min(1).max(400),
  nextAction:z.string().min(1).max(450), deadlineEvidence:z.string().max(250).nullable(),
  needsAttention:z.boolean(), uncertainty:z.string().max(400).nullable(),
}).strict();
export type Assessment = z.infer<typeof assessmentSchema>;
export type ReviewItem = { email:ReviewEmail; assessment:Assessment };
const properties = {
  id:{type:'string'},company:{type:['string','null']},companyEvidence:{type:['string','null']},
  topic:{type:'string',enum:topics},priority:{type:'string',enum:priorities},
  summary:{type:'string'},priorityReason:{type:'string'},nextAction:{type:'string'},
  deadlineEvidence:{type:['string','null']},needsAttention:{type:'boolean'},uncertainty:{type:['string','null']},
};
export const reviewOutputFormat = { type:'json_schema',name:'inbox_review',strict:true,schema:{
  type:'object',additionalProperties:false,required:['reviews'],properties:{reviews:{type:'array',items:{
    type:'object',additionalProperties:false,properties,required:Object.keys(properties),
  }}},
}};
export function parseReviewOutput(raw:unknown, emails:ReviewEmail[]):Assessment[] {
  const result=z.object({reviews:z.array(assessmentSchema).max(MAX_REVIEW_MESSAGES)}).strict().parse(raw);
  if(result.reviews.length!==emails.length || new Set(result.reviews.map(r=>r.id)).size!==emails.length || result.reviews.some(r=>!emails.some(e=>e.id===r.id))) throw Error('invalid_review_ids');
  return emails.map(email=>{
    const review=result.reviews.find(r=>r.id===email.id)!;
    // Evidence must actually occur in the supplied email, not be invented by the model.
    const source=[email.sender,email.senderEmail,email.subject,email.text].join(' ').toLowerCase();
    if(review.company && (!review.companyEvidence || !source.includes(review.companyEvidence.toLowerCase()))) {
      review.company=null;review.companyEvidence=null;review.needsAttention=true;
    }
    if(review.deadlineEvidence && !source.includes(review.deadlineEvidence.toLowerCase())) {
      review.deadlineEvidence=null;review.needsAttention=true;
    }
    if(!review.company||email.truncated||email.previewOnly||email.hasAttachments) review.needsAttention=true;
    return review;
  });
}
export function safeOutlookLink(value:unknown):string|null {
  if(typeof value!=='string')return null;
  try{const u=new URL(value);return u.protocol==='https:'&&!u.username&&!u.password&&['outlook.office.com','outlook.office365.com','outlook.live.com'].includes(u.hostname)?u.href:null}catch{return null}
}
export function normalizeMessage(m:any):InboxMessage {
  if(typeof m.id!=='string'||!m.id)throw Error('invalid_message');
  return {id:m.id,subject:String(m.subject||'(No subject)').slice(0,500),sender:String(m.from?.emailAddress?.name||m.from?.emailAddress?.address||'Unknown sender').slice(0,300),senderEmail:String(m.from?.emailAddress?.address||'').slice(0,320),receivedAt:String(m.receivedDateTime||''),preview:String(m.bodyPreview||'').slice(0,1500),isRead:!!m.isRead,hasAttachments:!!m.hasAttachments,webLink:safeOutlookLink(m.webLink)};
}
export function emailForReview(m:any):ReviewEmail {
  const email=normalizeMessage(m);
  const previewOnly=m.body?.contentType?.toLowerCase()!=='text'||typeof m.body.content!=='string';
  const text=previewOnly?email.preview:m.body.content;
  return {...email,text:text.slice(0,MAX_EMAIL_TEXT),truncated:text.length>MAX_EMAIL_TEXT,previewOnly};
}
export function demoEmails():ReviewEmail[] {
  const examples=[
    {id:'demo-delivery',subject:'Order WF-1042: delivery delayed — installation tomorrow',sender:'Rasa · Demo Logistics',senderEmail:'rasa@logistics.example',text:'Hello, this is Rasa from Demo Logistics. The truck for order WF-1042 has broken down. Your furniture delivery will not arrive for the installation tomorrow at 09:00. Please confirm an alternative delivery arrangement today. We need your decision before 16:00. Thank you, Rasa.'},
    {id:'demo-quote',subject:'Quote request: 24 oak desks for our new office',sender:'Jonas · Demo Office',senderEmail:'jonas@office.example',text:'Sveiki, rašo Jonas iš Demo Office. Įrengiame naują biurą ir norėtume gauti pasiūlymą 24 ąžuoliniams darbo stalams, 140 x 70 cm, su metalinėmis kojomis. Prašome nurodyti kainą, pristatymo terminą ir surinkimo kainą. Galutinio užsakymo dar nepatvirtinome. Ačiū!'},
    {id:'demo-catalogue',subject:'New finishes catalogue — for your reference',sender:'Demo Materials',senderEmail:'catalogue@materials.example',text:'Hello from Demo Materials. Our updated range of wood finishes is available this month. This is a general product announcement for your reference. No action or response is required.'},
  ];
  const receivedAt=new Date().toISOString();
  return examples.map(e=>({...e,receivedAt,preview:e.text.slice(0,200),isRead:false,hasAttachments:false,webLink:null,truncated:false,previewOnly:false}));
}
