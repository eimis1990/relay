import {ownerOf,saveReview} from '../../../../lib/review-store';
import {databaseConfigured} from '../../../../lib/workspace-store';
import { config,json,session } from '../../../../lib/outlook';
import { reviewRequestSchema } from '../../../../lib/inbox-review';
import { collectReviewEmails,reviewEmails,ReviewError } from '../../../../lib/review-service';
export const runtime='nodejs';
export const dynamic='force-dynamic';
export const maxDuration=90;
export async function GET(){const s=await session();if(!s)return json({error:'Connect Outlook to review your inbox.'},401);return json({configured:!!process.env.OPENAI_API_KEY,maxMessages:5});}
export async function POST(req:Request){
  try{
    const cfg=config();
    if(req.headers.get('origin')!==cfg.origin)return json({error:'Invalid request origin.'},403);
    const s=await session();if(!s)return json({error:'Reconnect Outlook to review your inbox.'},401);
    if(!req.headers.get('content-type')?.startsWith('application/json'))return json({error:'JSON content is required.'},415);
    const body=await req.text();if(Buffer.byteLength(body)>14000)return json({error:'Request too large.'},413);
    let parsed;try{parsed=reviewRequestSchema.safeParse(JSON.parse(body))}catch{return json({error:'Invalid JSON.'},400)}
    if(!parsed.success)return json({error:'Select 1–5 different emails and confirm the AI review.'},400);
    if(!process.env.OPENAI_API_KEY)return json({error:'AI review is not configured yet. Add the Relay OpenAI API key in Vercel and redeploy.'},503);
    const emails=await collectReviewEmails(s.token,parsed.data);
    const result=await reviewEmails(emails,parsed.data);
    if(parsed.data.source==='outlook'&&databaseConfigured()){for(const item of result.items)await saveReview(ownerOf(s),item,result.language,result.reviewedAt)}
    return json({...result,saved:parsed.data.source==='outlook'&&databaseConfigured()});
  }catch(error){return error instanceof ReviewError?json({error:error.message},error.status):json({error:'The review could not finish. Please retry; no changes were made to Outlook.'},502)}
}
