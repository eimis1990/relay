import { readFile, mkdir, writeFile, rename } from 'node:fs/promises';
import path from 'node:path';
import { z } from 'zod';
export const runtime='nodejs';
export const dynamic='force-dynamic';
const str=z.string().trim().max(2000);
const schema=z.object({company:str,contact:str,email:z.string().email().or(z.literal('')),industry:str,volume:str,mailboxes:str,provider:str,goals:z.array(z.string().max(100)).max(10),pain:str,urgency:str,languages:str,attachments:str,access:str,notes:str,submitted:z.boolean(),consent:z.boolean(),connections:z.object({outlook:z.object({account:str,requested:z.boolean()}),google:z.object({account:str,requested:z.boolean()}),telegram:z.object({account:str,requested:z.boolean()})})});
const file=path.join(process.cwd(),'.local-data','workspace.json');
function json(data:unknown,status=200){return Response.json(data,{status,headers:{'Cache-Control':'private, no-store'}})}
function local(req:Request){return process.env.NODE_ENV==='development'&&['localhost','127.0.0.1','[::1]'].includes(new URL(req.url).hostname)}
export async function GET(req:Request){if(!local(req))return json({error:'Hosted authentication and storage must be configured before this portal can accept client data.'},503);try{return json({data:JSON.parse(await readFile(file,'utf8')),mode:'local-development'})}catch(e:any){if(e.code==='ENOENT')return json({data:null,mode:'local-development'});return json({error:'Local workspace could not be loaded.'},500)}}
export async function POST(req:Request){if(!local(req))return json({error:'Hosted authentication and storage are not configured.'},503);if(req.headers.get('origin')!==`http://${req.headers.get('host')}`)return json({error:'Invalid request origin'},403);try{const text=await req.text();if(text.length>24000)return json({error:'Request too large'},413);const data=schema.parse(JSON.parse(text));if(data.submitted&&(!data.consent||!data.company||!data.contact||!data.email))return json({error:'Complete your contact details and consent before submitting.'},400);await mkdir(path.dirname(file),{recursive:true,mode:0o700});const tmp=file+'.'+crypto.randomUUID()+'.tmp';await writeFile(tmp,JSON.stringify(data),{mode:0o600});await rename(tmp,file);return json({ok:true})}catch{return json({error:'Unable to save. Check your entries and try again.'},400)}}
