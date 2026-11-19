import 'server-only';
import {db} from './review-store';
import {seal,unseal} from './outlook';
import {defaultWorkflow,workflowSchema,type Workflow} from './workflow';
import {ReviewError} from './review-service';
export async function readWorkflow(owner:string,fallback?:{language?:'English'|'Lithuanian';rules_cipher?:string|null}){
 const rows=await db('relay_workflows',{owner_id:'eq.'+owner,limit:'1'});const row=rows[0];
 return row?{data:workflowSchema.parse(unseal(row.settings_cipher,'workflow:'+owner)),revision:row.revision,updatedAt:row.updated_at}:{data:{...defaultWorkflow,language:fallback?.language||'English',instructions:fallback?.rules_cipher?unseal<string>(fallback.rules_cipher,'rules:'+owner):''},revision:null,updatedAt:null};
}
export async function writeWorkflow(owner:string,data:Workflow,revision:number|null){
 const rows=await db('relay_workflows',revision===null?{}:{owner_id:'eq.'+owner,revision:'eq.'+revision},{method:revision===null?'POST':'PATCH',body:JSON.stringify({...(revision===null?{owner_id:owner}:{}),settings_cipher:seal(data,'workflow:'+owner),revision:(revision||0)+1,updated_at:new Date().toISOString()})});
 if(!rows.length)throw new ReviewError('Rules changed in another tab. Reload before saving.',409);
 return {data,revision:rows[0].revision,updatedAt:rows[0].updated_at};
}
