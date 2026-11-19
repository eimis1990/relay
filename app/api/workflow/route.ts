import {z} from 'zod';
import {session,config,json} from '../../../lib/outlook';
import {ownerOf,account} from '../../../lib/review-store';
import {readWorkflow,writeWorkflow} from '../../../lib/workflow-store';
import {workflowSchema} from '../../../lib/workflow';
import {ReviewError} from '../../../lib/review-service';
export const dynamic='force-dynamic';
export async function GET(){try{const s=await session();if(!s)return json({error:'Sign in to configure your assistant.'},401);const owner=ownerOf(s);return json(await readWorkflow(owner,await account(owner)??undefined))}catch(e){return json({error:e instanceof ReviewError?e.message:'Could not load your rules.'},e instanceof ReviewError?e.status:503)}}
export async function PUT(req:Request){try{if(req.headers.get('origin')!==config().origin)return json({error:'Invalid origin.'},403);const s=await session();if(!s)return json({error:'Sign in to save rules.'},401);const raw=await req.text();if(raw.length>24000)return json({error:'Too much configuration text.'},413);const p=z.object({data:workflowSchema,revision:z.number().int().min(1).max(2147483646).nullable()}).strict().safeParse(JSON.parse(raw));if(!p.success)return json({error:'Check label names, matching values and your time zone. Up to 12 rules are supported.'},400);return json(await writeWorkflow(ownerOf(s),p.data.data,p.data.revision))}catch(e){return json({error:e instanceof ReviewError?e.message:'Could not save rules. Reload to check for changes in another tab.'},e instanceof ReviewError?e.status:503)}}
