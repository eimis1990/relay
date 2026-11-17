import { config, session, json } from '../../../../lib/outlook';
export const dynamic='force-dynamic';
export async function GET(){try{config();const s=await session();return json({configured:true,connected:!!s,account:s?{email:s.email,name:s.name}:null})}catch{return json({configured:false,connected:false,account:null})}}
