import {cookies} from 'next/headers';
import {randomBytes,createHash} from 'node:crypto';
import {config,scope,categoryScope,seal,options,json,session} from '../../../../lib/outlook';
export const runtime='nodejs';
export async function POST(req:Request){try{
 const cfg=config();if(req.headers.get('origin')!==cfg.origin)return json({error:'Invalid origin'},403);
 const text=await req.text();if(text.length>6000)return json({error:'Request too large.'},413);const form=new URLSearchParams(text),background=form.get('background')==='true',categories=form.get('categories')==='true';
 if(categories&&!background)return json({error:'Category sync requires automatic processing and explicit consent.'},400);
 let extra:{}|{background:true;subject:string;language:string;rules:string;categories:boolean}={};
 if(background){const s=await session();if(!s?.subject)return json({error:'Sign in to Relay before enabling automatic processing.'},401);if(form.get('consent')!=='true')return json({error:'Confirm automatic review and storage first.'},400);const language=form.get('language')||'English',rules=form.get('priorityRules')||'';if(!['English','Lithuanian'].includes(language)||rules.length>1200)return json({error:'Invalid review settings.'},400);extra={background:true,subject:s.subject,language,rules,categories}}
 const verifier=randomBytes(32).toString('base64url'),state=randomBytes(32).toString('base64url');
 (await cookies()).set('relay_ms_flow',seal({state,verifier,expires:Date.now()+600000,...extra},'flow'),options(600));
 const u=new URL('https://login.microsoftonline.com/common/oauth2/v2.0/authorize');u.search=new URLSearchParams({client_id:cfg.clientId,response_type:'code',redirect_uri:cfg.redirect,response_mode:'query',scope:(categories?categoryScope:scope)+(background?' offline_access':''),state,code_challenge:createHash('sha256').update(verifier).digest('base64url'),code_challenge_method:'S256',prompt:background?'consent':'select_account'}).toString();return Response.redirect(u,303);
 }catch{return json({error:'Microsoft connection is not configured. See the deployment guide.'},503)}}
