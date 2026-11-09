/* Bản thử chỉ lưu trên thiết bị; không dùng để triển khai máy chủ. */
(function(){
let store={sheets:{},files:{},folders:{},properties:{SHEET_ID:'demo',FOLDER_ID:'root',SALT:'demo-only',TEACHER_PASSWORD:'DEMO-STEM'},cache:{}};
const clone=x=>JSON.parse(JSON.stringify(x));
function sheet(name){if(!store.sheets[name])return null;return {
 getDataRange:()=>({getValues:()=>clone(store.sheets[name])}),
 appendRow:r=>store.sheets[name].push(r.map(v=>typeof v==='string'&&v[0]==="'"?v.slice(1):v)),
 setFrozenRows:()=>{},getRange:(r,c,n,m)=>({setValues:values=>{store.sheets[name][r-1]=clone(values[0]);}}),
 deleteRow:r=>store.sheets[name].splice(r-1,1)
};}
const database={getSheetByName:sheet,insertSheet:n=>{store.sheets[n]=[];return sheet(n);},getUrl:()=>''};
const SpreadsheetApp={openById:()=>database,create:()=>({getId:()=> 'demo'})};
const PropertiesService={getScriptProperties:()=>({getProperty:k=>store.properties[k]||'',setProperty:(k,v)=>store.properties[k]=v})};
const CacheService={getScriptCache:()=>({get:k=>store.cache[k]||null,put:(k,v)=>store.cache[k]=v})};
const Utilities={getUuid:()=>(crypto.randomUUID?crypto.randomUUID():Array.from(crypto.getRandomValues(new Uint8Array(16)),v=>v.toString(16).padStart(2,'0')).join('')),DigestAlgorithm:{SHA_256:'demo'},computeDigest:(a,s)=>{let h=2166136261;for(const c of s)h=Math.imul(h^c.charCodeAt(0),16777619);return [h&255,h>>>8&255,h>>>16&255,h>>>24&255];},base64Decode:s=>Array.from(atob(s),c=>c.charCodeAt(0)),base64Encode:b=>btoa(Array.from(b,x=>String.fromCharCode(x&255)).join('')),newBlob:(bytes,mime,name)=>({getBytes:()=>bytes,getContentType:()=>mime,name})};
const LockService={getScriptLock:()=>({waitLock:()=>{},hasLock:()=>true,releaseLock:()=>{}})};
function folder(id){return {getId:()=>id,getUrl:()=>'',createFolder:()=>folder((crypto.randomUUID?crypto.randomUUID():Array.from(crypto.getRandomValues(new Uint8Array(16)),v=>v.toString(16).padStart(2,'0')).join(''))),setTrashed:()=>{},createFile:blob=>{const id=(crypto.randomUUID?crypto.randomUUID():Array.from(crypto.getRandomValues(new Uint8Array(16)),v=>v.toString(16).padStart(2,'0')).join(''));store.files[id]={bytes:blob.getBytes(),mime:blob.getContentType()};return {getId:()=>id};}};}
const DriveApp={createFolder:()=>folder('root'),getFolderById:folder,getFileById:id=>({getBlob:()=>({getBytes:()=>store.files[id].bytes,getContentType:()=>store.files[id].mime})})};
/* STEM Toán 7 — Google Apps Script V8. Không đặt mật khẩu trong GitHub. */
const TABLES = {
  Classes:['id','name','code','created'],
  Students:['id','name','classId','pinHash','groupId','stage','resource','answers','created'],
  Groups:['id','classId','name','code','created'],
  Submissions:['id','groupId','classId','stage','studentId','data','files','created','requestId'],
  Reviews:['id','submissionId','classId','reviewerId','groupId','kind','scores','total','comment','created','requestId'],
  Events:['id','studentId','classId','kind','stage','created'],
  Sessions:['hash','userId','role','expires']
};
function setup() {
  const p=PropertiesService.getScriptProperties();
  if(!p.getProperty('TEACHER_PASSWORD') || p.getProperty('TEACHER_PASSWORD').length<12) throw Error('Đặt TEACHER_PASSWORD ít nhất 12 kí tự trong Script properties trước.');
  if(!p.getProperty('SALT')) p.setProperty('SALT', Utilities.getUuid()+Utilities.getUuid());
  if(!p.getProperty('SHEET_ID')) p.setProperty('SHEET_ID',SpreadsheetApp.create('STEM — Thước song song').getId());
  if(!p.getProperty('FOLDER_ID')) p.setProperty('FOLDER_ID',DriveApp.createFolder('STEM — Ảnh sản phẩm (riêng tư)').getId());
  const db=SpreadsheetApp.openById(p.getProperty('SHEET_ID'));
  Object.keys(TABLES).forEach(n=>{let s=db.getSheetByName(n);if(!s){s=db.insertSheet(n);s.appendRow(TABLES[n]);s.setFrozenRows(1);}});
  console.log('Bảng dữ liệu: '+db.getUrl());
  console.log('Thư mục ảnh: '+DriveApp.getFolderById(p.getProperty('FOLDER_ID')).getUrl());
}
function doGet(e) {
  const t=HtmlService.createTemplateFromFile('Bridge');
  t.channel=String(e.parameter.channel||'').replace(/[^a-zA-Z0-9-]/g,'').slice(0,80);
  t.bootstrap=JSON.stringify({channel:t.channel,allowed:PropertiesService.getScriptProperties().getProperty('ALLOWED_ORIGIN')||'',popup:e.parameter.popup==='1'}).replace(/</g,'\\u003c');
  return t.evaluate().setTitle('Kết nối lớp STEM').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
let DATABASE_;
function db_(){return DATABASE_||(DATABASE_=SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')));}
function rows_(name){const v=db_().getSheetByName(name).getDataRange().getValues();return v.slice(1).map((r,i)=>Object.assign({_row:i+2},Object.fromEntries(TABLES[name].map((k,j)=>[k,r[j]]))));}
function add_(name,o){db_().getSheetByName(name).appendRow(TABLES[name].map(k=>sheetSafe_(o[k]===undefined?'':o[k])));return o;}
function update_(name,o){db_().getSheetByName(name).getRange(o._row,1,1,TABLES[name].length).setValues([TABLES[name].map(k=>sheetSafe_(o[k]===undefined?'':o[k]))]);}
function sheetSafe_(v){return typeof v==='string'&&/^[=+@\-]/.test(v)?"'"+v:v;}
function uid_(){return Utilities.getUuid();}
function now_(){return new Date().toISOString();}
function hash_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s)+PropertiesService.getScriptProperties().getProperty('SALT')).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');}
function parse_(s,f){try{return JSON.parse(s||'');}catch(e){return f;}}
function text_(v,max){if(typeof v!=='string'||!v.trim()||v.trim().length>max)throw Error('Thiếu nội dung hoặc nội dung quá dài.');return v.trim();}
function check_(yes,msg){if(!yes)throw Error(msg);}
function session_(userId,role){const token=uid_()+uid_();add_('Sessions',{hash:hash_(token),userId,role,expires:Date.now()+7*86400000});return token;}
function auth_(token){const s=rows_('Sessions').find(x=>x.hash===hash_(String(token||''))&&Number(x.expires)>Date.now());check_(s,'Phiên đã hết hạn. Em hãy đăng nhập lại.');return s;}
function student_(s){const u=rows_('Students').find(x=>x.id===s.userId);check_(u,'Không tìm thấy học sinh.');return u;}
function unlocked_(u){return 5;}
function allowStage_(u,stage){check_(Number.isInteger(stage)&&stage>=1&&stage<=5,'Hoạt động không hợp lệ.');}
function event_(u,kind,stage){add_('Events',{id:uid_(),studentId:u.id,classId:u.classId,kind,stage:stage||0,created:now_()});}
function limited_(key,max){const c=CacheService.getScriptCache(),k='limit:'+hash_(key),n=Number(c.get(k)||0);check_(n<max,'Thao tác quá nhiều. Vui lòng thử lại sau vài phút.');c.put(k,String(n+1),600);}
function snapshot_(s){
  if(s.role==='teacher')return {role:'teacher',classes:rows_('Classes').map(clean_),students:rows_('Students').map(x=>{let y=clean_(x);delete y.pinHash;y.visits=rows_('Events').filter(e=>e.studentId===x.id&&e.kind==='Lượt tham gia mới').length;return y;}),groups:rows_('Groups').map(clean_),submissions:rows_('Submissions').map(subClean_),reviews:rows_('Reviews').map(reviewClean_),events:rows_('Events').slice(-400).map(clean_)};
  const u=student_(s),group=rows_('Groups').find(x=>x.id===u.groupId),all=rows_('Submissions').filter(x=>x.classId===u.classId),ids=new Set(all.map(x=>x.id));
  return {role:'student',visits:rows_('Events').filter(e=>e.studentId===u.id&&e.kind==='Lượt tham gia mới').length,student:{id:u.id,name:u.name,classId:u.classId,groupId:u.groupId,stage:Number(u.stage),resource:parse_(u.resource,{}),answers:parse_(u.answers,{})},className:(rows_('Classes').find(x=>x.id===u.classId)||{}).name,group:group?clean_(group):null,members:rows_('Students').filter(x=>u.groupId&&x.groupId===u.groupId).map(x=>({id:x.id,name:x.name})),groups:rows_('Groups').filter(x=>x.classId===u.classId).map(x=>({id:x.id,name:x.name})),submissions:all.map(subClean_),reviews:rows_('Reviews').filter(x=>ids.has(x.submissionId)).map(reviewClean_),events:rows_('Events').filter(x=>x.studentId===u.id).slice(-80).map(clean_)};
}
function clean_(x){const y=Object.assign({},x);delete y._row;return y;}
function subClean_(x){const y=clean_(x);y.data=parse_(x.data,{});y.files=parse_(x.files,[]);return y;}
function reviewClean_(x){const y=clean_(x);y.scores=parse_(x.scores,[]);return y;}
function api(payload){
  const lock=LockService.getScriptLock();
  try{check_(payload&&JSON.stringify(payload).length<6000000,'Yêu cầu quá lớn.');lock.waitLock(30000);return {ok:true,data:route_(payload)};}
  catch(e){return {ok:false,error:String(e.message||e)};}finally{if(lock.hasLock())lock.releaseLock();}
}
function route_(p){
  const a=p.action,d=p.data||{};
  if(a==='register'){
    let cl,pinHash='',name=text_(d.name,90);
    if(d.className){
      const className=text_(d.className,60).replace(/\s+/g,' ').toUpperCase();
      limited_('register-name:'+className,100);
      cl=rows_('Classes').find(x=>String(x.name).trim().replace(/\s+/g,' ').toUpperCase()===className);
      if(!cl)cl=add_('Classes',{id:uid_(),name:className,code:uid_().replace(/-/g,'').slice(0,8).toUpperCase(),created:now_()});
    }else{
      const code=text_(d.classCode,30).toUpperCase();cl=rows_('Classes').find(x=>x.code===code);check_(cl,'Không tìm thấy lớp.');limited_('register:'+code,100);
      check_(/^\d{6,12}$/.test(d.pin),'PIN gồm 6–12 chữ số.');
    }
    const id='HS'+uid_().replace(/-/g,'').slice(0,10).toUpperCase();
    if(!d.className)pinHash=hash_(id+':'+d.pin);
    add_('Students',{id,name,classId:cl.id,pinHash,groupId:'',stage:0,resource:'{}',answers:'{}',created:now_()});
    return {token:session_(id,'student'),id};
  }
  if(a==='login'){
    const id=text_(d.id,30).toUpperCase();limited_('login:'+id,15);const u=rows_('Students').find(x=>x.id===id&&x.pinHash===hash_(id+':'+d.pin));check_(u,'Mã học sinh hoặc PIN chưa đúng.');event_(u,'Đăng nhập',0);return {token:session_(u.id,'student')};
  }
  if(a==='teacherLogin'){limited_('teacherLogin',20);check_(hash_(String(d.password))===hash_(PropertiesService.getScriptProperties().getProperty('TEACHER_PASSWORD')),'Mật khẩu giáo viên chưa đúng.');return {token:session_('teacher','teacher')};}
  const s=auth_(p.token);
  if(a==='beginVisit'){
    if(s.role!=='student')return {count:0};const u=student_(s),id='visit:'+u.id+':'+text_(d.visitId,80),events=rows_('Events');
    if(!events.some(x=>x.id===id))add_('Events',{id,studentId:u.id,classId:u.classId,kind:'Lượt tham gia mới',stage:0,created:now_()});
    return {count:rows_('Events').filter(x=>x.studentId===u.id&&x.kind==='Lượt tham gia mới').length};
  }
  if(a==='snapshot'){if(s.role==='student')autoCompleteStage2_(student_(s));return snapshot_(s);}
  if(a==='logout'){const sh=db_().getSheetByName('Sessions');sh.deleteRow(s._row);return true;}
  if(a==='createClass'){
    check_(s.role==='teacher','Chỉ giáo viên được tạo lớp.');return add_('Classes',{id:uid_(),name:text_(d.name,60),code:uid_().replace(/-/g,'').slice(0,8).toUpperCase(),created:now_()});
  }
  if(a==='resetPin'){
    check_(s.role==='teacher','Chỉ giáo viên được đổi PIN.');check_(/^\d{6,12}$/.test(d.pin),'PIN gồm 6–12 chữ số.');const u=rows_('Students').find(x=>x.id===d.id);check_(u,'Không tìm thấy học sinh.');u.pinHash=hash_(u.id+':'+d.pin);update_('Students',u);
    rows_('Sessions').filter(x=>x.userId===u.id).reverse().forEach(x=>db_().getSheetByName('Sessions').deleteRow(x._row));return true;
  }
  if(a==='getImage'){
    const sub=rows_('Submissions').find(x=>x.id===d.submissionId);check_(sub,'Không tìm thấy bài nộp.');if(s.role!=='teacher')check_(student_(s).classId===sub.classId,'Ảnh thuộc lớp khác.');check_(parse_(sub.files,[]).some(x=>x.id===d.fileId),'Ảnh không thuộc bài nộp.');const b=DriveApp.getFileById(d.fileId).getBlob();return {url:'data:'+b.getContentType()+';base64,'+Utilities.base64Encode(b.getBytes())};
  }
  if(a==='review')return review_(s,d);
  check_(s.role==='student','Thao tác dành cho học sinh.');const u=student_(s);
  if(a==='joinGroup'){
    allowStage_(u,3);check_(!u.groupId,'Em đã có nhóm. Nhờ giáo viên kiểm tra nếu cần điều chỉnh.');let g;
    if(d.code)g=rows_('Groups').find(x=>x.classId===u.classId&&x.code===String(d.code).trim().toUpperCase());
    else {g=add_('Groups',{id:uid_(),classId:u.classId,name:text_(d.name,60),code:uid_().replace(/-/g,'').slice(0,8).toUpperCase(),created:now_()});}
    check_(g,'Mã nhóm chưa đúng hoặc nhóm thuộc lớp khác.');u.groupId=g.id;update_('Students',u);event_(u,'Tham gia nhóm',3);return snapshot_(s);
  }
  if(a==='openStage'){const stage=Number(d.stage);allowStage_(u,stage);event_(u,'Mở hoạt động',stage);return true;}
  if(a==='saveResource'){
    allowStage_(u,2);check_(['knowledge','solution'].includes(d.kind),'Học liệu không hợp lệ.');const r=parse_(u.resource,{});check_(JSON.stringify(d.state).length<20000,'Dữ liệu học liệu quá lớn.');r[d.kind]=d.state;u.resource=JSON.stringify(r);update_('Students',u);const completedNow=autoCompleteStage2_(u);return {stage2Completed:Boolean(parse_(u.answers,{})[2]),snapshot:completedNow?snapshot_(s):null};
  }
  if(a==='completeStage'){
    const stage=Number(d.stage);allowStage_(u,stage);
    const answer=validateCompletion_(stage,d.answer||{},parse_(u.resource,{}),rows_('Submissions').filter(x=>x.groupId===u.groupId));const answers=parse_(u.answers,{});answers[stage]=answer;u.answers=JSON.stringify(answers);u.stage=Object.keys(answers).length;update_('Students',u);event_(u,'Hoàn thành hoạt động',stage);return snapshot_(s);
  }
  if(a==='submit'){
    const stage=Number(d.stage);check_(Number.isInteger(stage)&&stage>=3&&stage<=5,'Chỉ nộp ở hoạt động 3–5.');check_(u.groupId,'Em cần tham gia nhóm.');allowStage_(u,stage);
    const requestId=text_(d.requestId,80),old=rows_('Submissions').find(x=>x.requestId===requestId&&x.studentId===u.id);if(old)return subClean_(old);
    const data=validateSubmission_(stage,d.content),photos=d.photos;check_(Array.isArray(photos)&&photos.length>=1&&photos.length<=4,'Nộp từ 1 đến 4 ảnh.');
    const decoded=photos.map((f,i)=>{check_(['image/jpeg','image/png'].includes(f.mime),'Chỉ nhận ảnh JPG, PNG.');check_(typeof f.base64==='string'&&f.base64.length<=1400000,'Ảnh tối đa 1 MB sau nén.');const bytes=Utilities.base64Decode(f.base64);check_(bytes.length>20&&bytes.length<=1048576,'Dung lượng ảnh không hợp lệ.');check_(f.mime==='image/jpeg'?(bytes[0]&255)===255&&(bytes[1]&255)===216:(bytes[0]&255)===137&&(bytes[1]&255)===80,'Nội dung ảnh không hợp lệ.');return {bytes,mime:f.mime,name:'anh-'+(i+1)+(f.mime==='image/jpeg'?'.jpg':'.png')};});
    const id=uid_(),folder=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID')).createFolder(u.classId+'_'+u.groupId+'_HD'+stage+'_'+id),files=[];
    try{decoded.forEach(f=>{const file=folder.createFile(Utilities.newBlob(f.bytes,f.mime,f.name));files.push({id:file.getId(),name:f.name});});const sub=add_('Submissions',{id,groupId:u.groupId,classId:u.classId,stage,studentId:u.id,data:JSON.stringify(data),files:JSON.stringify(files),created:now_(),requestId});event_(u,'Nộp bài nhóm',stage);return subClean_(sub);}catch(e){folder.setTrashed(true);throw e;}
  }
  throw Error('Thao tác không được hỗ trợ.');
}
function validateSubmission_(stage,d){
  check_(d&&typeof d==='object','Thiếu nội dung.');const out={note:text_(d.note,5000)};
  if(stage===3){out.materials=text_(d.materials,2000);out.plan=text_(d.plan,3000);}
  if(stage===4){check_(Array.isArray(d.trials)&&d.trials.length===3,'Cần đủ ba lần thử.');out.trials=d.trials.map(t=>{const length=Number(t.length),a=Number(t.a),b=Number(t.b);check_(t.length!==''&&t.a!==''&&t.b!==''&&Number.isFinite(length)&&length>0&&Number.isFinite(a)&&a>0&&a<180&&Number.isFinite(b)&&b>0&&b<180,'Số đo chưa hợp lệ.');return {length,a,b,delta:Math.abs(a-b),pass:length>=12&&Math.abs(a-b)<2};});out.improvement=text_(d.improvement,3000);}
  if(stage===5)out.change=text_(d.change,3000);return out;
}
function stage2Ready_(r){
  const knowledge=r.knowledge,solution=r.solution;
  return Boolean(knowledge&&Array.isArray(knowledge.done)&&knowledge.done.length===4&&knowledge.done.every(x=>x===true)&&solution&&['yes','no'].includes(solution.guess)&&Array.isArray(solution.records)&&solution.records.length===3&&solution.records.every(x=>Number.isFinite(Number(x.theta))&&Number(x.theta)>=20&&Number(x.theta)<=145)&&new Set(solution.records.map(x=>Number(x.theta))).size===3);
}
function autoCompleteStage2_(u){
  const answers=parse_(u.answers,{});if(answers[2]||!stage2Ready_(parse_(u.resource,{})))return false;
  answers[2]={explanation:'',completedFromResources:true};u.answers=JSON.stringify(answers);u.stage=Object.keys(answers).length;update_('Students',u);event_(u,'Hoàn thành hoạt động',2);return true;
}
function validateCompletion_(stage,a,r,subs){
  if(stage===1)return {problem:text_(a.problem,3000)};
  if(stage===2){check_(r.knowledge&&Array.isArray(r.knowledge.done)&&r.knowledge.done.length===4&&r.knowledge.done.every(x=>x===true),'Hoàn thành đủ bốn thử thách kiến thức nền.');const sol=r.solution;check_(sol&&['yes','no'].includes(sol.guess)&&Array.isArray(sol.records)&&sol.records.length===3&&sol.records.every(x=>Number.isFinite(Number(x.theta))&&Number(x.theta)>=20&&Number(x.theta)<=145)&&new Set(sol.records.map(x=>Number(x.theta))).size===3,'Lưu dự đoán và thử nghiệm ở ba vị trí khác nhau.');const explanation=typeof a.explanation==='string'?a.explanation.trim():'';check_(explanation.length<=3000,'Nội dung giải thích quá dài.');return {explanation};}
  check_(subs.some(x=>Number(x.stage)===stage),'Nhóm cần nộp ảnh và hồ sơ của hoạt động này trước.');
  if(stage===5){check_(a.q1==='parallel'&&a.q2==='first','Xem lại hai câu hỏi cuối bài.');return {q1:a.q1,q2:a.q2,contribution:text_(a.contribution,2000)};}
  return {reflection:text_(a.reflection,2000)};
}
function review_(s,d){
  const sub=rows_('Submissions').find(x=>x.id===d.submissionId);check_(sub,'Không tìm thấy bài nộp.');
  let groupId='',role='teacher';
  if(s.role!=='teacher'){const u=student_(s);check_(u.classId===sub.classId&&u.groupId,'Em cần có nhóm trong cùng lớp.');allowStage_(u,Number(sub.stage));groupId=u.groupId;role=groupId===sub.groupId?'self':'peer';}
  const feedback=[3,4].includes(Number(sub.stage));let scores=[],total='',kind=feedback?'comment_'+role:role;
  if(!feedback){const maxima=[3,2.5,2.5,1,1];check_(Array.isArray(d.scores)&&d.scores.length===5,'Cần đủ năm tiêu chí.');scores=d.scores.map((v,i)=>{const n=Number(v);check_(v!==''&&Number.isFinite(n)&&n>=0&&n<=maxima[i]&&Math.abs(n*4-Math.round(n*4))<1e-8,'Điểm theo bước 0,25 và không vượt điểm tối đa.');return n;});total=scores.reduce((x,y)=>x+y,0);}
  const requestId=text_(d.requestId,80),old=rows_('Reviews').find(x=>x.requestId===requestId&&x.reviewerId===s.userId);if(old)return reviewClean_(old);
  return reviewClean_(add_('Reviews',{id:uid_(),submissionId:sub.id,classId:sub.classId,reviewerId:s.userId,groupId,kind,scores:JSON.stringify(scores),total,comment:text_(d.comment,3000),created:now_(),requestId}));
}

let opened;
function databaseLocal(){return opened||(opened=new Promise((resolve,reject)=>{const r=indexedDB.open('stem-demo-server-v1',1);r.onupgradeneeded=()=>r.result.createObjectStore('state');r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(Error('Bộ nhớ trình duyệt không khả dụng cho bản thử.'));}));}
async function readLocal(){const db=await databaseLocal();return new Promise((resolve,reject)=>{const r=db.transaction('state').objectStore('state').get('all');r.onsuccess=()=>resolve(r.result);r.onerror=reject;});}
async function writeLocal(){const db=await databaseLocal();return new Promise((resolve,reject)=>{const tx=db.transaction('state','readwrite');tx.objectStore('state').put(store,'all');tx.oncomplete=resolve;tx.onerror=()=>reject(Error('Không lưu được dữ liệu thử trên máy.'));});}
let loaded=false,queue=Promise.resolve();
window.STEM_DEMO_API=p=>{const job=queue.then(async()=>{if(!loaded){store=await readLocal()||store;Object.keys(TABLES).forEach(n=>{if(!store.sheets[n])store.sheets[n]=[TABLES[n]];});if(!rows_('Classes').length)add_('Classes',{id:'demo-class',name:'Lớp STEM 7 · thử nghiệm',code:'STEM7',created:now_()});loaded=true;}const before=clone(store);try{const r=api(p);await writeLocal();return r;}catch(e){store=before;return {ok:false,error:e.message};}});queue=job.catch(()=>{});return job;};
if(typeof module!=='undefined')module.exports={api,route_,validateCompletion_,validateSubmission_,review_,rows_,add_,snapshot_,store,TABLES};
})();
