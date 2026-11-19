import {timingSafeEqual} from 'node:crypto';
import {cronToken,runMailbox} from '../../../../lib/mail-worker';
import {db} from '../../../../lib/review-store';
import {json} from '../../../../lib/outlook';
export const dynamic='force-dynamic';export const maxDuration=180;
export async function POST(req:Request){try{const auth=req.headers.get('authorization')||'',expected='Bearer '+cronToken();if(auth.length!==expected.length||!timingSafeEqual(Buffer.from(auth),Buffer.from(expected)))return json({error:'Unauthorized.'},401);
 // Fair scheduling across the private pilot accounts; each invocation is bounded to one mailbox.
 const rows=await db('relay_mail_accounts',{enabled:'eq.true',lock_until:'lt.'+new Date().toISOString(),order:'last_polled_at.asc.nullsfirst',select:'owner_id',limit:'1'});if(!rows.length)return json({state:'idle'});return json(await runMailbox(rows[0].owner_id));
 }catch{return json({error:'Processing unavailable; the next scheduled run will retry.'},503)}}
