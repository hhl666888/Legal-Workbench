const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const { DatabaseSync } = require('node:sqlite');
const { execFileSync } = require('child_process');

const envPath = fs.existsSync(path.join(__dirname, '.env')) ? path.join(__dirname, '.env') : path.join(__dirname, '.env.example');
if (fs.existsSync(envPath)) {
  for (const line of fs.readFileSync(envPath, 'utf8').split(/\r?\n/)) {
    const m = line.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)\s*$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^['"]|['"]$/g, '');
  }
}

const PORT = process.env.PORT || 8789;
const RAW_KEY = process.env.DEEPSEEK_API_KEY || '';
// .env.example 默认是说明文字，不能被当成可用密钥。
const KEY = /^(替换为你的|your[_ -]?deepseek|replace[_ -]?with)/i.test(RAW_KEY.trim()) ? '' : RAW_KEY.trim();
const BASE = (process.env.DEEPSEEK_BASE_URL || 'https://api.deepseek.com/v1').replace(/\/$/, '');
const MODEL = process.env.DEEPSEEK_MODEL || 'deepseek-chat';
const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, 'data');
const GENERATED_DIR = path.join(DATA_DIR, 'generated');
fs.mkdirSync(GENERATED_DIR, {recursive:true});
const db = new DatabaseSync(path.join(DATA_DIR, 'legal-review.db'));
db.exec(`CREATE TABLE IF NOT EXISTS generated_contracts (
  id TEXT PRIMARY KEY, owner TEXT NOT NULL, review_id TEXT, source_name TEXT NOT NULL,
  output_name TEXT NOT NULL, file_path TEXT NOT NULL, adopted INTEGER NOT NULL,
  unadopted INTEGER NOT NULL, applied INTEGER NOT NULL, pending INTEGER NOT NULL,
  created_at TEXT NOT NULL
)`);
const tasks = new Map();
const sessions = new Map();
const SESSION_TTL = 8 * 60 * 60 * 1000;
const users = JSON.parse(process.env.AUTH_USERS || JSON.stringify({admin:{password:'admin123',name:'系统管理员',role:'admin'},legal:{password:'legal123',name:'法务专员',role:'legal'},viewer:{password:'viewer123',name:'只读员工',role:'viewer'}}));
const log = (msg, meta = {}) => console.log(new Date().toISOString(), msg, JSON.stringify(meta));

function json(res, code, data) {
  res.writeHead(code, {'Content-Type':'application/json; charset=utf-8','Access-Control-Allow-Origin':'http://localhost:'+PORT,'Cache-Control':'no-store'});
  res.end(JSON.stringify(data));
}
function auth(req){const h=req.headers.authorization||'';const token=h.startsWith('Bearer ')?h.slice(7):'';const entry=sessions.get(token);if(!entry)return null;if(Date.now()-entry.createdAt>SESSION_TTL){sessions.delete(token);return null}return entry.user}
function canReview(user){return user && (user.role==='admin'||user.role==='legal')}

function extract(name, buffer) {
  const ext = path.extname(name).toLowerCase();
  if (ext === '.txt') return buffer.toString('utf8');
  if (ext === '.docx') {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'contract-docx-'));
    const sourceZip = path.join(dir, 'contract.zip');
    const outputDir = path.join(dir, 'unzipped');
    try {
      fs.writeFileSync(sourceZip, buffer);
      // 使用 UTF-16LE Base64 传递脚本，避免 Windows PowerShell 在中文临时路径上发生 ByteString 编码错误。
      const ps = `Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::ExtractToDirectory('${sourceZip.replace(/'/g,"''")}','${outputDir.replace(/'/g,"''")}')`;
      const encoded = Buffer.from(ps, 'utf16le').toString('base64');
      execFileSync('powershell', ['-NoProfile', '-EncodedCommand', encoded], {maxBuffer:10*1024*1024});
      let xml = fs.readFileSync(path.join(outputDir, 'word', 'document.xml'), 'utf8');
      return xml.replace(/<w:tab[^>]*\/>/g,'\t').replace(/<\/w:p>/g,'\n').replace(/<\/w:tr>/g,'\n').replace(/<[^>]+>/g,'').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&apos;/g,"'");
    } catch (_) {
      throw Error('DOCX 文件解析失败，请确认文件未损坏、未加密，且不是仅修改扩展名的旧版 DOC 文件。');
    } finally { fs.rmSync(dir, {recursive:true, force:true}); }
  }
  if (ext === '.pdf') {
    try { return execFileSync('pdftotext', ['-', '-'], {input:buffer,maxBuffer:50*1024*1024}).toString('utf8'); }
    catch (_) { throw Error('PDF 文件解析失败，请确认 PDF 包含可复制文字且未加密。'); }
  }
  throw Error('当前仅支持 PDF、DOCX 或 TXT 文件。');
}

function chunks(s, n=20000) { const a=[]; for(let i=0;i<s.length;i+=n)a.push(s.slice(i,i+n)); return a; }
function xmlEsc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function xmlText(s){return String(s||'').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&')}
function compactText(s){return String(s||'').replace(/\s+/g,'')}
function replaceParagraph(xml,evidence,draft){const needle=compactText(evidence);if(!needle)return{xml,found:false};let found=false;const next=xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,p=>{if(found)return p;const texts=[...p.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)];const plain=texts.map(x=>xmlText(x[1])).join('');if(!compactText(plain).includes(needle))return p;found=true;let first=true;return p.replace(/<w:t(\s[^>]*)?>[\s\S]*?<\/w:t>/g,(node,attrs='')=>{if(!first)return `<w:t${attrs}></w:t>`;first=false;return `<w:t xml:space="preserve">${xmlEsc(draft)}</w:t>`})});return{xml:next,found}}
function makeDocx(base64,replacements){const dir=fs.mkdtempSync(path.join(os.tmpdir(),'contract-out-'));const src=path.join(dir,'src.zip'),out=path.join(dir,'updated.zip'),unz=path.join(dir,'unz');try{fs.writeFileSync(src,Buffer.from(base64,'base64'));const q=s=>s.replace(/'/g,"''");const unpack=`Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::ExtractToDirectory('${q(src)}','${q(unz)}')`;execFileSync('powershell',['-NoProfile','-EncodedCommand',Buffer.from(unpack,'utf16le').toString('base64')],{maxBuffer:10*1024*1024});const doc=path.join(unz,'word','document.xml');let xml=fs.readFileSync(doc,'utf8');const applied=[],pending=[];for(const x of replacements||[]){if(!x.evidence||!x.draft)continue;const result=replaceParagraph(xml,x.evidence,x.draft);xml=result.xml;if(result.found)applied.push({evidence:x.evidence,draft_clause:x.draft_clause||x.draft});else pending.push({evidence:x.evidence,draft_clause:x.draft_clause||x.draft,reason:'未在原合同 DOCX 中定位到该条款原文，已保留原内容，请人工确认'})}fs.writeFileSync(doc,xml);const pack=`Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${q(unz)}','${q(out)}')`;execFileSync('powershell',['-NoProfile','-EncodedCommand',Buffer.from(pack,'utf16le').toString('base64')],{maxBuffer:10*1024*1024});return{base64:fs.readFileSync(out).toString('base64'),applied,pending}}finally{fs.rmSync(dir,{recursive:true,force:true})}}
async function llm(system,user,attempt=0) {
  if(!KEY) throw Error('未配置有效的 DEEPSEEK_API_KEY，请在 .env.example 或 .env 中填写 DeepSeek API Key');
  const ac=new AbortController(),timer=setTimeout(()=>ac.abort(),90000); let r;
  try { r=await fetch(BASE+'/chat/completions',{method:'POST',headers:{'Content-Type':'application/json','Authorization':'Bearer '+KEY},body:JSON.stringify({model:MODEL,temperature:.1,messages:[{role:'system',content:system},{role:'user',content:user}]}),signal:ac.signal}); }
  finally { clearTimeout(timer); }
  if(!r.ok&&[429,500,502,503].includes(r.status)&&attempt<2){await new Promise(x=>setTimeout(x,700*2**attempt));return llm(system,user,attempt+1)}
  if(!r.ok) throw Error('DeepSeek 请求失败（'+r.status+'）');
  const d=await r.json(); return d.choices?.[0]?.message?.content||'';
}

async function run(id,body) {
  const t=tasks.get(id);
  try {
    t.status='processing'; const text=extract(body.fileName,Buffer.from(body.fileBase64,'base64')); if(!text.trim())throw Error('合同文件为空');
    const parts=chunks(text),issues=[];
    for(let i=0;i<parts.length;i++){t.progress=Math.round((i/parts.length)*45);const seg=await llm('你是专业中国企业法务合同审查助手。严格依据合同原文，按以下固定风险类别审查：付款、违约责任、知识产权、保密、自动续约、争议解决、单方解除、主体与签署、交付验收、数据与合规。只输出 JSON 数组，每条问题格式：{"category":"","level":"高|中|低","clause":"第X条","evidence":"原文证据≤120字","analysis":"风险分析","legal_basis":"法律依据（《民法典》/合同编/司法解释等；无可填未发现明确依据）","suggestion":"修改建议","draft_clause":"建议替换的条款范文，可直接使用"}。未发现问题输出 []。不得编造，不得把常识当成合同事实。只输出 JSON，不要 Markdown，不要解释。',`合同第 ${i+1}/${parts.length} 段：\n${parts[i]}\n\n用户要求：${body.request||''}`);try{issues.push(...JSON.parse(seg.replace(/```json|```/g,'').trim()))}catch(_){log('segment_parse_failed',{taskId:id,segment:i})}}
    t.progress=55; const legal=await llm('你是中国法律法规引用审查助手。只提取并核验合同明确引用的法律法规，判断有效、废止、替代或需确认；没有明确引用时输出“未发现合同中存在明确的法律法规引用”。不得虚构。',text);
    t.progress=70; const metaRaw=await llm('从合同原文提取甲方公司名称和合同标题，只输出 JSON：{"company":"","contract_title":""}，不得推测。',text.slice(0,30000)); let meta={company:'',contract_title:''}; try{meta={...meta,...JSON.parse(metaRaw.replace(/```json|```/g,'').trim())}}catch(_){}
    t.progress=85; const finalRaw=await llm('你是法务报告编辑。把分段审查的风险项整理为最终结构化 JSON 对象，输出格式：{"title":"合同标题","summary":"综合风险概述（含风险等级高/中/低和问题数量）","issues":[...],"legal":"法律引用审查结果","signoff":["签署前重点事项1","..."]}。要求：合并去重分段重复的风险项；为每个 issue 补全 legal_basis 和 draft_clause（若原数据缺失则根据合同类型补上）；draft_clause 必须是可直接替换使用的条款范文；signoff 提炼 3-6 条签署前重点事项。只输出 JSON，不要 Markdown，不要解释。',JSON.stringify({title:(meta.contract_title||'合同审查').trim(),issues,legal}));
    let report={title:(meta.contract_title||'合同审查').trim(),summary:'',issues,legal,signoff:[]};
    try{report=JSON.parse(finalRaw.replace(/```json|```/g,'').trim())}catch(_){log('final_parse_failed',{taskId:id})}
    t.report=report; t.issues=report.issues||issues; t.text=text; t.meta=meta;t.progress=100;t.status='completed';t.audit={segments:parts.length,model:MODEL,completedAt:new Date().toISOString()};log('review_completed',{taskId:id,segments:parts.length,model:MODEL,issues:(report.issues||issues).length});
  } catch(e) { t.status='failed';t.error=e.message;log('review_failed',{taskId:id,error:e.message}); }
}

function receiveBody(req,res,user) {
  let s='',n=0,closed=false;
  req.on('data',c=>{n+=c.length;if(n>28*1024*1024&&!closed){closed=true;res.destroy()}else if(!closed)s+=c});
  req.on('end',()=>{if(!closed)try{const b=JSON.parse(s);if(!b.fileName||!b.fileBase64)throw Error('缺少合同文件');const id=crypto.randomUUID();tasks.set(id,{status:'queued',progress:0,createdAt:Date.now(),owner:user.username});json(res,202,{taskId:id,status:'queued'});run(id,b)}catch(e){json(res,400,{error:e.message})}});
}

const server=http.createServer((req,res)=>{
  if(req.method==='OPTIONS'){res.writeHead(204,{'Access-Control-Allow-Origin':'http://localhost:'+PORT,'Access-Control-Allow-Headers':'Content-Type'});return res.end()}
  if(req.url==='/api/health')return json(res,200,{ok:true,deepseekConfigured:Boolean(KEY),model:MODEL,port:PORT});
  if(req.url==='/api/login'&&req.method==='POST'){let s='';req.on('data',c=>s+=c);req.on('end',()=>{try{const b=JSON.parse(s),u=users[b.username];if(!u||u.password!==b.password)return json(res,401,{error:'账号或密码错误'});const token=crypto.randomBytes(24).toString('hex');const user={username:b.username,name:u.name,role:u.role};sessions.set(token,{user,createdAt:Date.now()});json(res,200,{token,user,expiresIn:SESSION_TTL})}catch(e){json(res,400,{error:'登录请求无效'})}});return}
  if(req.url==='/api/me'&&req.method==='GET'){const u=auth(req);return u?json(res,200,{user:u}):json(res,401,{error:'未登录'})}
  if(req.url==='/api/review'&&req.method==='POST'){const u=auth(req);if(!u)return json(res,401,{error:'请先登录'});if(!canReview(u))return json(res,403,{error:'只读员工没有发起审阅权限'});return receiveBody(req,res,u)}
  if(req.url==='/api/refine'&&req.method==='POST'){const u=auth(req);if(!u)return json(res,401,{error:'请先登录'});if(u.role==='viewer')return json(res,403,{error:'只读账号无权操作'});let s='';req.on('data',c=>s+=c);req.on('end',()=>{try{const b=JSON.parse(s);if(!b.content)return json(res,400,{error:'缺少报告内容'});llm('你是法务报告编辑。基于以下报告内容和法务修改意见，重新整理为结构化 JSON 报告。保留法务已修改的内容，补全每个 issue 的 legal_basis 和 draft_clause（若缺失）。输出格式：{"title":"","summary":"","issues":[{"category":"","level":"高|中|低","clause":"","evidence":"","analysis":"","legal_basis":"","suggestion":"","draft_clause":""}],"legal":"","signoff":[""]}。只输出 JSON，不要 Markdown，不要解释。',`当前报告内容：\n${b.content}\n\n法务修改意见：${b.instruction||'无'}`).then(refined=>{try{const report=JSON.parse(refined.replace(/```json|```/g,'').trim());json(res,200,{report})}catch(_){json(res,200,{report:refined})}}).catch(e=>json(res,500,{error:'AI 调用失败：'+e.message}))}catch(e){json(res,400,{error:'请求解析失败'})}});return}
  if(req.url==='/api/generate'&&req.method==='POST'){const u=auth(req);if(!u)return json(res,401,{error:'请先登录'});if(!canReview(u))return json(res,403,{error:'只读员工没有生成新合同权限'});let s='';req.on('data',c=>s+=c);req.on('end',()=>{try{const b=JSON.parse(s);if(!b.fileName||!b.fileBase64)return json(res,400,{error:'缺少原合同文件'});if(path.extname(b.fileName).toLowerCase()!=='.docx')return json(res,400,{error:'仅 DOCX 合同支持保留排版生成新版本，PDF/TXT 请另存为 DOCX 后重试'});const issues=(b.issues||[]).filter(it=>it._status==='adopted'&&it.evidence&&it.draft_clause);const total=(b.issues||[]).length;if(!issues.length)return json(res,400,{error:'当前没有已采纳的条款，请在报告中点击“采纳”后再生成新合同'});const result=makeDocx(b.fileBase64,issues.map(it=>({evidence:it.evidence,draft:it.draft_clause,draft_clause:it.draft_clause})));const id=crypto.randomUUID(),outName=b.fileName.replace(/\.docx$/i,'')+'_审阅版.docx',filePath=path.join(GENERATED_DIR,id+'.docx'),createdAt=new Date().toISOString();fs.writeFileSync(filePath,Buffer.from(result.base64,'base64'));db.prepare('INSERT INTO generated_contracts (id,owner,review_id,source_name,output_name,file_path,adopted,unadopted,applied,pending,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)').run(id,u.username,b.reviewId||null,b.fileName,outName,filePath,issues.length,total-issues.length,result.applied.length,result.pending.length,createdAt);log('docx_generated',{user:u.username,contractId:id,adopted:issues.length,applied:result.applied.length,pending:result.pending.length});return json(res,200,{generatedId:id,fileName:outName,downloadUrl:'/api/generated/'+id+'/download',adopted:issues.length,unadopted:total-issues.length,applied:result.applied.length,pending:result.pending.length,pendingItems:result.pending})}catch(e){return json(res,500,{error:'生成新合同失败：'+e.message})}});return}
  if(req.url==='/api/generated'&&req.method==='GET'){const u=auth(req);if(!u)return json(res,401,{error:'请先登录'});const rows=(u.role==='admin'?db.prepare('SELECT * FROM generated_contracts ORDER BY created_at DESC'):db.prepare('SELECT * FROM generated_contracts WHERE owner=? ORDER BY created_at DESC')).all(...(u.role==='admin'?[]:[u.username]));return json(res,200,{items:rows.map(x=>({id:x.id,reviewId:x.review_id,sourceName:x.source_name,fileName:x.output_name,adopted:x.adopted,unadopted:x.unadopted,applied:x.applied,pending:x.pending,createdAt:x.created_at,downloadUrl:'/api/generated/'+x.id+'/download'}))})}
  const dm=req.url.match(/^\/api\/generated\/([^/?]+)\/download$/);if(dm&&req.method==='GET'){const u=auth(req);if(!u)return json(res,401,{error:'请先登录'});const row=db.prepare('SELECT * FROM generated_contracts WHERE id=?').get(dm[1]);if(!row||(!u||u.role!=='admin')&&row.owner!==u.username)return json(res,404,{error:'新合同不存在'});if(!fs.existsSync(row.file_path))return json(res,410,{error:'新合同文件已丢失'});const name=encodeURIComponent(row.output_name);res.writeHead(200,{'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','Content-Disposition':`attachment; filename*=UTF-8''${name}`,'Content-Length':fs.statSync(row.file_path).size,'Cache-Control':'no-store'});return fs.createReadStream(row.file_path).pipe(res)}
  const tm=req.url.match(/^\/api\/review\/([^/?]+)\/text$/);if(tm&&req.method==='GET'){const u=auth(req);if(!u)return json(res,401,{error:'请先登录'});const t=tasks.get(tm[1]);if(!t)return json(res,404,{error:'任务不存在或已过期'});if(t.owner!==u.username&&u.role!=='admin')return json(res,403,{error:'无权查看该审阅任务'});return json(res,200,{taskId:tm[1],text:t.text||''})}
  const m=req.url.match(/^\/api\/review\/([^/?]+)$/);if(m&&req.method==='GET'){const u=auth(req);if(!u)return json(res,401,{error:'请先登录'});const t=tasks.get(m[1]);if(!t)return json(res,404,{error:'任务不存在或已过期'});if(t.owner!==u.username&&u.role!=='admin')return json(res,403,{error:'无权查看该审阅任务'});const {text,...rest}=t||{};return json(res,200,{taskId:m[1],...rest})}
  const file=req.url==='/'?'/index.html':req.url.split('?')[0],p=path.resolve(ROOT,'.'+file);if(!p.startsWith(path.resolve(ROOT)+path.sep)||!fs.existsSync(p))return res.writeHead(404).end();res.writeHead(200,{'Content-Type':p.endsWith('.css')?'text/css':p.endsWith('.js')?'text/javascript':'text/html','Cache-Control':'no-store, no-cache, must-revalidate','Pragma':'no-cache','Expires':'0'});fs.createReadStream(p).pipe(res)
});
server.listen(PORT,()=>log('server_started',{port:PORT}));
