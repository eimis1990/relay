// Test-process preload only. Never imported by application code or enabled in deployments.
const original=globalThis.fetch;
globalThis.fetch=(input,init)=>{
  const url=typeof input==='string'?input:input instanceof URL?input.href:input.url;
  if(url.startsWith('https://api.telegram.org/bot'))return original(process.env.REVIEW_MOCK_URL+'/telegram/'+url.split('/').at(-1),init);
  if(url.startsWith('https://graph.microsoft.com/v1.0/'))return original(process.env.REVIEW_MOCK_URL+'/graph/'+url.slice('https://graph.microsoft.com/v1.0/'.length),init);
  if(url==='https://login.microsoftonline.com/common/oauth2/v2.0/token')return original(process.env.REVIEW_MOCK_URL+'/token',init);
  if(url==='https://api.openai.com/v1/responses')return original(process.env.REVIEW_MOCK_URL+'/openai',init);
  return original(input,init);
};
