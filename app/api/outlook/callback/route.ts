import {NextResponse} from 'next/server';
import {cookies} from 'next/headers';
import {timingSafeEqual} from 'node:crypto';
import {config,scope,categoryScope,hasCategoryScopes,unseal,seal,options,setSession,graph,json} from '../../../../lib/outlook';
import {account,db} from '../../../../lib/review-store';
import {setupScheduler} from '../../../../lib/mail-worker';
export const runtime='nodejs';
export async function GET(req:Request){
 let cfg;try{cfg=config()}catch{return json({error:'Microsoft setup is incomplete.'},503)}
 let background=false,categories=false;
 const back=(result:string)=>{const response=NextResponse.redirect(cfg.origin+'/?outlook='+result+(categories?'#workflow':background?'#review':'#portal'),303);response.headers.set('Cache-Control','no-store');response.headers.set('Referrer-Policy','no-referrer');return response};
 const c=await cookies(),flowCookie=c.get('relay_ms_flow')?.value;c.set('relay_ms_flow','',options(0));
 try{
  if(!flowCookie)throw Error();const flow=unseal<{state:string;verifier:string;expires:number;background?:boolean;subject?:string;language?:string;rules?:string;categories?:boolean}>(flowCookie,'flow');background=flow.background===true;categories=background&&flow.categories===true;
  const url=new URL(req.url),state=url.searchParams.get('state')||'',code=url.searchParams.get('code');
  if(flow.expires<Date.now()||Buffer.byteLength(state)!==Buffer.byteLength(flow.state)||!timingSafeEqual(Buffer.from(state),Buffer.from(flow.state))||!code||url.searchParams.has('error'))return back('failed');
  const r=await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({client_id:cfg.clientId,client_secret:cfg.clientSecret,grant_type:'authorization_code',code,code_verifier:flow.verifier,redirect_uri:cfg.redirect,scope:(categories?categoryScope:scope)+(background?' offline_access':'')}),cache:'no-store',signal:AbortSignal.timeout(15000)});
  if(!r.ok)return back('failed');const token=await r.json();if(typeof token.access_token!=='string'||!Number.isFinite(token.expires_in))return back('failed');
  const me=await graph(token.access_token,'me?$select=id,displayName,mail,userPrincipalName');if(!me.ok)return back('failed');const profile=await me.json();
  const email=[profile.mail,profile.userPrincipalName].find(v=>typeof v==='string'&&cfg.allowed.includes(v.toLowerCase()));if(!email)return back('not_allowed');
  if(typeof profile.id!=='string'||!profile.id||profile.id.length>200||background&&profile.id!==flow.subject)return back('failed');
  if(categories&&!hasCategoryScopes(token.scope))return back('category_permission_required');
  if(background){
   if(typeof token.refresh_token!=='string'||!token.refresh_token)return back('offline_required');
   const owner='microsoft:'+profile.id,existing=await account(owner);
   if(existing&&Date.parse(existing.lock_until)>Date.now())return back('busy');
   await setupScheduler();
   await db('relay_mail_accounts',{on_conflict:'owner_id'},{method:'POST',body:JSON.stringify({owner_id:owner,email,refresh_cipher:seal(token.refresh_token,'refresh:'+owner),enabled:true,started_at:existing?.started_at||new Date().toISOString(),language:categories?(existing?.language||'English'):(flow.language||'English'),rules_cipher:categories?(existing?.rules_cipher||null):seal(flow.rules||'','rules:'+owner),last_error:null,category_enabled:categories,category_permission:categories,category_error:null})},'resolution=merge-duplicates,return=representation');
  }
  await setSession({subject:profile.id,token:token.access_token,email,name:profile.displayName||email,expires:Date.now()+Math.max(0,Math.min(token.expires_in-60,3600))*1000});
  return back(categories?'categories_connected':background?'automatic_connected':'connected');
 }catch{return back('failed')}
}
