import {config,clearSession,json} from '../../../../lib/outlook';
export async function POST(req:Request){try{if(req.headers.get('origin')!==config().origin)return json({error:'Invalid origin'},403);await clearSession();return json({ok:true})}catch{return json({error:'Connection is not configured.'},503)}}
