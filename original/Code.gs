/**
 * HISOB — Telegram + Google Sheets + mavjud HTML panel. UMBRELLA V14 PLAN HOLATI + V13 RUXSATLAR.
 * Google Sheets > Extensions > Apps Script. Bot tokeni faqat Script Properties da.
 * UMBRELLA V9.2: Plan/Prixod/Qoldiq/Jami; loyiha tugmalari 3 qator (4/3/3) va Jami Excel. V8 Dashboard saqlangan. Loyiha tanlash xabari faqat bo‘lim nomi.
 * doPost javoblari HtmlService orqali (Telegram 302 muammosini chetlab o‘tish).
 * Deploy: Web app / Execute as Me / Anyone (Telegram webhook uchun); admin API parol bilan himoyalangan.
 */
const HISOB = {
  PRODUCT: 'Mahsulotlar', RECEIPT: 'Kirimlar', META: 'Tizim', AUDIT: 'Amallar_tarixi', IMPORTS: 'Import_tarixi', USERS: 'Foydalanuvchilar', PLAN_STATE: 'Plan_holatlari',
  PH: ['id','project','code','unit','kg','profile','length','plan','opening','material'],
  RH: ['id','productId','qty','date','note','source','userId','createdAt'],
  AH: ['sana_vaqt','amal','kim','loyiha','kod','dona','izoh'],
  MH: ['kalit','qiymat'], IH: ['hash','name','count','date'], PSH: ['loyiha','holat','sana_vaqt','ozgartirgan'],
  UH: ['telegram_id','ism','telefon','rol','holat','royxat_vaqti','tasdiq_vaqti','parol_salt','parol_hash','sessiya_tugashi'],
  CACHE_SECONDS: 21600
};

function props_(){return PropertiesService.getScriptProperties();}
function cfg_(key){return props_().getProperty(key)||'';}
function sheetBook_(){
  const id=cfg_('SPREADSHEET_ID');
  const ss=id?SpreadsheetApp.openById(id):SpreadsheetApp.getActiveSpreadsheet();
  if(!ss)throw Error('SPREADSHEET_ID sozlanmagan yoki Google Sheet ochilmagan');
  return ss;
}
function tab_(name,head){
  const ss=sheetBook_(); let sh=ss.getSheetByName(name);
  if(!sh)sh=ss.insertSheet(name);
  if(sh.getLastRow()===0)sh.getRange(1,1,1,head.length).setValues([head]).setFontWeight('bold');
  return sh;
}
function prepare_(){
  tab_(HISOB.PRODUCT,HISOB.PH);tab_(HISOB.RECEIPT,HISOB.RH);
  const meta=tab_(HISOB.META,HISOB.MH);
  tab_(HISOB.AUDIT,HISOB.AH);tab_(HISOB.IMPORTS,HISOB.IH);tab_(HISOB.USERS,HISOB.UH);tab_(HISOB.PLAN_STATE,HISOB.PSH);
  if(meta.getLastRow()<2)meta.getRange(2,1,1,2).setValues([['revision','0']]);
}
/** Birinchi ishga tushirish: editor ichida shu funksiyani Run qiling. */
function initializeHISOB_(){prepare_();Logger.log('HISOB Sheets tayyor');}
function rows_(name,header){
  const sh=tab_(name,header),end=sh.getLastRow();
  if(end<=1)return [];
  return sh.getRange(2,1,end-1,header.length).getValues().filter(r=>String(r[0])!=='');
}
/** Reja statuslari alohida Sheetda saqlanadi: barcha eski kirim va rejalar saqlanadi. */
function projectStatesV14_(){
  const result={};
  rows_(HISOB.PLAN_STATE,HISOB.PSH).forEach(r=>{
    const project=String(r[0]||'').trim(),state=String(r[1]||'');
    if(project && state==='paused')result[project]='paused';
  });
  return result;
}
function readDb_(){
  prepare_();
  return {
    products:rows_(HISOB.PRODUCT,HISOB.PH).map(r=>Object.fromEntries(HISOB.PH.map((k,i)=>[k,['id','kg','length','plan','opening'].includes(k)?Number(r[i]||0):String(r[i]||'')]))),
    receipts:rows_(HISOB.RECEIPT,HISOB.RH).map(r=>({id:Number(r[0]),productId:Number(r[1]),qty:Number(r[2]),date:asDate_(r[3]),note:String(r[4]||''),source:String(r[5]||''),userId:String(r[6]||''),createdAt:String(r[7]||'')})),
    prixodImports:rows_(HISOB.IMPORTS,HISOB.IH).map(r=>({hash:String(r[0]||''),name:String(r[1]||''),count:Number(r[2]||0),date:asDate_(r[3])})),
    projectStates:projectStatesV14_(),
    version:3
  };
}
function asDate_(v){
  if(v instanceof Date)return Utilities.formatDate(v,hisobTZ_(),'yyyy-MM-dd');
  return String(v||'').slice(0,10);
}
function writeRows_(name,headers,objects){
  const sh=tab_(name,headers);
  const last=sh.getLastRow();
  if(last>1)sh.getRange(2,1,last-1,headers.length).clearContent();
  if(objects.length){
    if(sh.getMaxRows()<objects.length+1)sh.insertRowsAfter(sh.getMaxRows(),objects.length+1-sh.getMaxRows());
    const matrix=objects.map(obj=>headers.map(h=>obj[h]===undefined||obj[h]===null?'':obj[h]));
    sh.getRange(2,1,matrix.length,headers.length).setValues(matrix);
  }
}
function revision_(){
  const m=tab_(HISOB.META,HISOB.MH);
  return Number(m.getRange(2,2).getValue()||0);
}
function bumpRevision_(){const m=tab_(HISOB.META,HISOB.MH);const rev=revision_()+1;m.getRange(2,1,1,2).setValues([['revision',String(rev)]]);return rev;}
function hisobTZ_(){return cfg_('HISOB_TIMEZONE')||'Asia/Tashkent';}
function now_(){return Utilities.formatDate(new Date(),hisobTZ_(),'yyyy-MM-dd HH:mm:ss');}
function audit_(type,who,project,code,qty,description){
  const raw=String(who||'');
  const u=/^\d+$/.test(raw)?umbrellaUserV11_(raw):null;
  const actor=u?(u.name+' [TG:'+u.id+']'):raw;
  tab_(HISOB.AUDIT,HISOB.AH).appendRow([now_(),type,actor,project||'',code||'',qty||0,description||'']);
}
function normalizeKey_(s){return String(s||'').trim().toUpperCase().replace(/\s+/g,'');}
function validateDb_(data){
  if(!data||!Array.isArray(data.products)||!Array.isArray(data.receipts))throw Error('Ma’lumot tuzilmasi noto‘g‘ri');
  if(data.products.length>30000||data.receipts.length>100000)throw Error('Ma’lumotlar limiti oshdi');
  const ids=new Set(),keys=new Set();
  data.products.forEach(p=>{
    if(!Number.isSafeInteger(Number(p.id))||Number(p.id)<=0)throw Error('Mahsulot ID noto‘g‘ri');
    if(ids.has(String(p.id)))throw Error('Takroriy mahsulot ID');ids.add(String(p.id));
    const project=normalizeKey_(p.project),code=normalizeKey_(p.code),key=project+'|'+code;
    if(!project||!code||project.length>100||code.length>100||keys.has(key))throw Error('Loyiha / kod noto‘g‘ri yoki takrorlangan');
    keys.add(key);
    ['kg','length','plan','opening'].forEach(f=>{const x=Number(p[f]||0);if(!Number.isFinite(x)||x<0)throw Error('Noto‘g‘ri '+f)});
    if(!Number.isInteger(Number(p.plan))||!Number.isInteger(Number(p.opening)))throw Error('Reja / boshlang‘ich kirim butun son emas');
  });
  if(data.prixodImports!==undefined && (!Array.isArray(data.prixodImports)||data.prixodImports.length>30000))throw Error('Import tarixi noto‘g‘ri');
  const rids=new Set();data.receipts.forEach(r=>{
    if(!Number.isSafeInteger(Number(r.id))||Number(r.id)<=0||rids.has(String(r.id)))throw Error('Kirim ID takrorlangan / noto‘g‘ri');
    rids.add(String(r.id));
    if(!ids.has(String(r.productId)))throw Error('Kirimda mavjud bo‘lmagan mahsulot ID');
    if(!Number.isInteger(Number(r.qty))||Number(r.qty)===0)throw Error('Kirim miqdori noto‘g‘ri');
    if(!/^\d{4}-\d{2}-\d{2}$/.test(asDate_(r.date)))throw Error('Kirim sanasi noto‘g‘ri');
  });
}
function hash_(s){return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8));}
/** Script Properties: ADMIN_SETUP_PASSWORD >=12 belgi. Bir marta ishga tushiring. */
function setInitialAdminPassword_(){
  const pw=cfg_('ADMIN_SETUP_PASSWORD');
  if(pw.length<12)throw Error('ADMIN_SETUP_PASSWORD kamida 12 belgidan iborat bo‘lsin');
  const salt=Utilities.getUuid();
  props_().setProperty('ADMIN_PASSWORD_HASH',salt+':'+hash_(salt+':'+pw));
  props_().deleteProperty('ADMIN_SETUP_PASSWORD');
  Logger.log('Admin paroli o‘rnatildi (parol chiqarilmaydi)');
}
function requireSession_(token){
  if(!token||!CacheService.getScriptCache().get('auth:'+String(token)))throw Error('Sessiya tugagan. Qayta kiring.');
}
function apiLogin(password){
  const stored=cfg_('ADMIN_PASSWORD_HASH');
  if(!stored)throw Error('Admin paroli hali sozlanmagan');
  const sep=stored.indexOf(':'),salt=stored.slice(0,sep),expected=stored.slice(sep+1);
  if(hash_(salt+':'+String(password||''))!==expected)throw Error('Parol noto‘g‘ri');
  const token=Utilities.getUuid()+Utilities.getUuid();
  CacheService.getScriptCache().put('auth:'+token,'ok',HISOB.CACHE_SECONDS);
  return {token,expiresMinutes:HISOB.CACHE_SECONDS/60};
}
function apiLoad(token){requireSession_(token);return {db:readDb_(),revision:revision_()};}
function apiCommit(token,data,expectedRevision){
  requireSession_(token);validateDb_(data);
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    prepare_();const current=revision_();
    if(Number(expectedRevision)!==current)throw Error('CONFLICT: Bot yoki boshqa administrator hisobni o‘zgartirgan. Eski ma’lumotni ustidan yozmang; avval yangi holatni yuklang.');
    // To‘xtatilgan loyiha bo‘yicha yangi ijobiy prixod yoki mavjud prixodni oshirish mumkin emas.
    const paused=projectStatesV14_(),previousReceipts=rows_(HISOB.RECEIPT,HISOB.RH);
    const previousById=new Map(previousReceipts.map(r=>[String(r[0]),{productId:String(r[1]),qty:Number(r[2]||0)}]));
    const productsById=new Map(data.products.map(p=>[String(p.id),p]));
    for(const r of data.receipts){
      const old=previousById.get(String(r.id)),p=productsById.get(String(r.productId));
      const increased=!old?Number(r.qty)>0:(old.productId!==String(r.productId)?Number(r.qty)>0:Number(r.qty)>old.qty);
      if(increased && p && paused[p.project]==='paused')throw Error(p.project+' — plan to‘xtatilgan. Avval Loyihalar bo‘limidan davom ettiring.');
    }
    writeRows_(HISOB.PRODUCT,HISOB.PH,data.products);
    writeRows_(HISOB.RECEIPT,HISOB.RH,data.receipts.map(r=>({...r,source:r.source||'ADMIN',createdAt:r.createdAt||now_()})));
    writeRows_(HISOB.IMPORTS,HISOB.IH,(data.prixodImports||[]));
    const revision=bumpRevision_();audit_('ADMIN_SAQLASH','Admin','','',0,`${data.products.length} mahsulot, ${data.receipts.length} kirim`);
    SpreadsheetApp.flush();return {revision};
  }finally{lock.releaseLock();}
}
function doGet(e){
  if(e && e.parameter && e.parameter.mini==='1'){
    return HtmlService.createHtmlOutputFromFile('MiniApp')
      .setTitle('UMBRELLA — HISOB')
      .addMetaTag('viewport','width=device-width, initial-scale=1, maximum-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('HISOB — admin va Telegram')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function botToken_(){const token=cfg_('BOT_TOKEN');if(!token)throw Error('BOT_TOKEN Script Properties da kiritilmagan');return token;}
function botCall_(method,payload){
  const res=UrlFetchApp.fetch('https://api.telegram.org/bot'+botToken_()+'/'+method,{method:'post',contentType:'application/json',payload:JSON.stringify(payload),muteHttpExceptions:true});
  const out=JSON.parse(res.getContentText());if(!out.ok)throw Error('Telegram '+method+': '+out.description);
  return out.result;
}
function msg_(chat,text,keyboard){
  const args={chat_id:chat,text:String(text),disable_web_page_preview:true};
  if(keyboard)args.reply_markup=keyboard;
  return botCall_('sendMessage',args);
}
function miniAppUrlV10_(){
  const base=cfg_('WEB_APP_URL').split('?')[0].trim();
  if(!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(base))throw Error('WEB_APP_URL /exec manzili sozlanmagan');
  return base+'?mini=1';
}
function miniLaunchButtonsV10_(){
  return {inline_keyboard:[[{text:'📱 UMBRELLA ilovasini ochish',web_app:{url:miniAppUrlV10_()}}]]};
}
function mainMenu_(id){return umbrellaMenuV11_(id);}
function dashboardMenu_(){return {inline_keyboard:[
  [{text:'📊 Dashboard',callback_data:'dash:main'},{text:'🏗 Loyihalar',callback_data:'dash:projects'}],
  [{text:'📅 Kunlik (30 kun)',callback_data:'dash:daily'}],
  [{text:'📈 Hisobot Dashboard',callback_data:'dash:report'}],
  [{text:'⬅️ Hisobotga qaytish',callback_data:'dash:back'}]
]};}

/** UMBRELLA V11: ro'yxat va operator sessiyalari. Telefon foydalanuvchining o'z kontakti orqali olinadi. */
function umbrellaOwnersV11_(){return cfg_('ALLOWED_TELEGRAM_IDS').split(',').map(x=>x.trim()).filter(Boolean);}
function umbrellaIsOwnerV11_(id){return umbrellaOwnersV11_().includes(String(id));}
function umbrellaUsersSheetV11_(){return tab_(HISOB.USERS,HISOB.UH);}
function umbrellaUserV11_(id){
  const sh=umbrellaUsersSheetV11_(),last=sh.getLastRow(),wanted=String(id);
  if(last<2)return null;
  const values=sh.getRange(2,1,last-1,HISOB.UH.length).getValues();
  for(let i=0;i<values.length;i++)if(String(values[i][0])===wanted){
    const v=values[i];return {row:i+2,id:wanted,name:String(v[1]||''),phone:String(v[2]||''),role:String(v[3]||'viewer'),status:String(v[4]||''),approval:String(v[6]||''),salt:String(v[7]||''),passhash:String(v[8]||''),sessionUntil:Number(v[9]||0)};
  }
  return null;
}
/** YOPIQ REJIM: migratsiya yakunlanmaguncha HECH KIM Telegramdan ma'lumot olmaydi. */
function umbrellaClosedReadyV12_(){return cfg_('UMBRELLA_CLOSED_MODE_V12')==='1';}
function userAllowed_(id){
  if(!umbrellaClosedReadyV12_())return false;
  const u=umbrellaUserV11_(id);
  return !!u&&u.status==='active'&&['viewer','operator'].includes(u.role)&&
    (umbrellaIsOwnerV11_(id)||u.approval.startsWith('V12|'));
}

function umbrellaCanWriteV11_(id) {
  if (!userAllowed_(id)) return false;

  const u = umbrellaUserV11_(id);
  if (!u || u.status !== 'active') return false;

  // Asosiy administratorga to‘liq huquq
  if (umbrellaIsOwnerV11_(id)) return true;

  // Oddiy operatorlarga parol va faol sessiya shart
  return u.role === 'operator' &&
         !!u.passhash &&
         u.sessionUntil > Date.now();
}

function umbrellaRequireWriteV11_(id){
  if(!umbrellaCanWriteV11_(id))throw Error('Faqat faol operator mumkin. Administrator tasdig‘i va Mini App ichida parol bilan kirish talab qilinadi.');
}
function umbrellaNameV11_(text){
  const value=String(text||'').trim().replace(/\s+/g,' ');
  if(value.length<3||value.length>80||!/[\p{L}]/u.test(value)||/^[=+@-]/.test(value))throw Error('Ism va familiyangizni 3–80 belgi bilan yozing.');
  return value;
}
function umbrellaRegIntroV11_(chat,id){
  setMode_(id,'reg_name');
  msg_(chat,'👋 UMBRELLA tizimiga xush kelibsiz.\nRo‘yxatdan o‘tish uchun ism va familiyangizni yuboring.\nKeyin Telegram orqali O‘Z telefon raqamingizni ulashasiz.');
}
function umbrellaContactKeyboardV11_(){return {keyboard:[[{text:'📱 Telefon raqamimni yuborish',request_contact:true}],[{text:'❌ Bekor qilish'}]],resize_keyboard:true,one_time_keyboard:true};}
function umbrellaRegisterV11_(id,name,contact){
  if(!contact||Number(contact.user_id)!==Number(id))throw Error('Faqat Telegramdagi o‘z telefon kontaktingizni yuboring.');
  const phone=String(contact.phone_number||'').replace(/[\s()-]/g,'');
  if(!/^\+?[0-9]{9,15}$/.test(phone))throw Error('Telefon raqami noto‘g‘ri.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const existing=umbrellaUserV11_(id);
    if(existing)throw Error('Bu Telegram ID avval ro‘yxatdan o‘tgan.');
    const owner=umbrellaIsOwnerV11_(id),sh=umbrellaUsersSheetV11_();
    sh.appendRow([String(id),name,phone,owner?'operator':'viewer',owner?'active':'pending',now_(),owner?now_():'','','',0]);
    audit_('RO‘YXAT',''+id,'','',0,'Yangi foydalanuvchi; holat '+(owner?'active':'pending'));
    return {owner,name,phone};
  }finally{lock.releaseLock();}
}
/** Operator, hisobot ko'ruvchisi yoki rad etish qarorini faqat ega chiqaradi. */
function umbrellaAdminDecisionV12_(actorId,targetId,decision){
  if(!umbrellaIsOwnerV11_(actorId)||!userAllowed_(actorId))throw Error('Faqat asosiy administrator tasdiqlashi mumkin.');
  if(!['operator','viewer','deny'].includes(decision))throw Error('Ruxsat turi noto‘g‘ri.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const user=umbrellaUserV11_(targetId);
    if(!user||user.status!=='pending')throw Error('Bu so‘rov endi tasdiq kutmayapti.');
    if(umbrellaIsOwnerV11_(targetId))throw Error('Administrator hisobini bu yerda boshqarib bo‘lmaydi.');
    const sh=umbrellaUsersSheetV11_();
    const status=decision==='deny'?'rejected':'active';
    const role=decision==='deny'?'viewer':decision;
    sh.getRange(user.row,4,1,2).setValues([[role,status]]);
    sh.getRange(user.row,7).setValue(decision==='deny'?'':'V12|'+now_());
    sh.getRange(user.row,10).setValue(0); // eski operator sessiyasi bekor
    audit_(decision==='deny'?'RUXSAT_RAD':'RUXSAT_BERILDI',actorId,'','',0,user.name+' [TG:'+targetId+'] => '+decision);
    return {...user,role,status};
  }finally{lock.releaseLock();}
}
/** Oldingi tugmalar bilan moslik. */
function umbrellaAdminApproveV11_(actorId,targetId,approved){
  return umbrellaAdminDecisionV12_(actorId,targetId,approved?'operator':'viewer');
}
function umbrellaApprovalButtonsV12_(id){return {inline_keyboard:[
  [{text:'✅ Operator',callback_data:'v11:approve:'+id},{text:'📊 Faqat hisobot',callback_data:'v11:viewer:'+id}],
  [{text:'❌ Rad etish',callback_data:'v12:deny:'+id}]
]};}
function umbrellaPendingUsersV12_(max){
  const sh=umbrellaUsersSheetV11_(),last=sh.getLastRow();
  if(last<2)return [];
  return sh.getRange(2,1,last-1,5).getValues().filter(r=>String(r[4])==='pending'&&/^\d+$/.test(String(r[0]))).slice(0,Math.min(15,Math.max(1,Number(max)||10)));
}
function umbrellaPendingMenuV12_(){return {remove_keyboard:true};}
/** Eskidan kelgan ruxsatlar bilan ma'lumot ochilib qolmasligi uchun bir martalik migratsiya.
 * Foydalanuvchilar, mahsulotlar yoki kirimlarni O'CHIRMAYDI. Faqat ro'l/holat/sessiya.
 * Ishga tushirishdan OLDIN Sheets va Code.gs zaxirasini saqlang.
 */
function setupUmbrellaClosedV12(){
  // Eslatma: ommaviy Web App orqali google.script.run bilan chaqirishni rad etamiz.
  // Apps Script muharririda bevosita ishga tushirganda faol foydalanuvchi = egasi.
  const activeEmail=Session.getActiveUser().getEmail();
  const ownerEmail=Session.getEffectiveUser().getEmail();
  if(!activeEmail||activeEmail!==ownerEmail)throw Error('Bu sozlashni faqat Apps Script muharriridagi loyiha egasi bajarishi mumkin.');
  if(umbrellaClosedReadyV12_()){
    Logger.log('UMBRELLA V12 yopiq rejimi allaqachon o‘rnatilgan. Oldingi tasdiqlar saqlandi.');
    return {alreadyInstalled:true};
  }
  const owners=umbrellaOwnersV11_();
  if(!owners.length||owners.some(x=>!/^\d{5,20}$/.test(x)))throw Error('ALLOWED_TELEGRAM_IDS asosiy admin Telegram ID bilan to‘g‘ri kiritilsin.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const sh=umbrellaUsersSheetV11_(),last=sh.getLastRow();
    let pending=0,ownerCount=0,unchanged=0;
    if(last>1){
      const vals=sh.getRange(2,1,last-1,HISOB.UH.length).getValues();
      vals.forEach(function(v,i){
        const id=String(v[0]||''),row=i+2;
        if(!/^\d+$/.test(id))return;
        if(owners.includes(id)){
          if(String(v[3])!=='operator'||String(v[4])!=='active'){
            sh.getRange(row,4,1,2).setValues([['operator','active']]);
            sh.getRange(row,7).setValue('V12|'+now_());
          }
          ownerCount++;
        }else if(String(v[4])==='rejected'||String(v[4])==='pending'){
          unchanged++;
        }else{
          sh.getRange(row,5).setValue('pending');
          // Yangi admin tasdig'igacha eski operator sessiyasini ham tugatamiz.
          sh.getRange(row,10).setValue(0);
          pending++;
        }
      });
    }
    props_().setProperty('UMBRELLA_CLOSED_MODE_V12','1');
    audit_('YOPIQ_TIZIM_V12','SYSTEM','','',0,'Qayta tasdiqlash '+pending+' foydalanuvchi; adminlar '+ownerCount);
    Logger.log('UMBRELLA V12 YOPIQ: yangi va eski tasdiqlanmaganlar BLOK; qayta ko‘rib chiqiladi '+pending+' foydalanuvchi.');
    Logger.log('Asosiy adminlar ro‘yxati: '+owners.length+' ID; bazadagi adminlar: '+ownerCount+'.');
    Logger.log('Sozlama tayyor. /start yuboring; admin yangi va eski so‘rovlarni «⏳ Ruxsat so‘rovlari»da ko‘radi.');
    return {pending,ownerCount,unchanged};
  }finally{lock.releaseLock();}
}

function umbrellaAuditItemsV11_(user,items,sign,rev){
  if(!items||!items.length)return;
  const u=umbrellaUserV11_(user),actor=u?(u.name+' [TG:'+u.id+']'):String(user);
  const sh=tab_(HISOB.AUDIT,HISOB.AH),date=now_(),rows=items.map(item=>[
    date,sign>0?'TELEGRAM_KIRIM_POZITSIYA':'TELEGRAM_QAYTARISH_POZITSIYA',actor,
    item.prod.project,item.prod.code,sign*item.qty,'Rev '+rev
  ]);
  const start=sh.getLastRow()+1;
  if(sh.getMaxRows()<start+rows.length-1)sh.insertRowsAfter(sh.getMaxRows(),start+rows.length-1-sh.getMaxRows());
  sh.getRange(start,1,rows.length,HISOB.AH.length).setValues(rows);
}
function umbrellaRecentActivityV11_(limit){
  const rows=rows_(HISOB.AUDIT,HISOB.AH).slice(-Math.max(1,Math.min(40,Number(limit)||15))).reverse();
  return rows.map(r=>[r[0],r[1],r[2],r[3],r[4],r[5]].filter(v=>String(v||'')!=='').join(' · '));
}
function umbrellaMenuV11_(id){
  if(!userAllowed_(id))return umbrellaPendingMenuV12_();
  const common=[['📊 Hisobot','📱 Ilovani ochish'],['👤 Profil']];
  if(umbrellaIsOwnerV11_(id) || umbrellaCanWriteV11_(id))common.unshift(['📦 Prixod','↩️ Prixodni qaytarish'],['📥 Plan qo‘shish','🚪 Chiqish']);
  else common.push(['🔐 Operator kirishi']);
  if(umbrellaIsOwnerV11_(id))common.push(['⏳ Ruxsat so‘rovlari','📋 Faoliyat jurnali']);
  common.push(['📖 Yordam']);
  return {keyboard:common,resize_keyboard:true};
}
function umbrellaShowProfileV11_(chat,id){
  const u=umbrellaUserV11_(id);
  if(!u){umbrellaRegIntroV11_(chat,id);return;}
  const role=u.role==='operator'?(umbrellaCanWriteV11_(id)?'Operator (kirilgan)':'Operator (parol kerak)'):'Hisobot ko‘ruvchisi';
  msg_(chat,'👤 '+u.name+'\nTelegram ID: '+u.id+'\nTelefon: ****'+u.phone.slice(-4)+'\nHuquq: '+role,umbrellaMenuV11_(id));
}
function umbrellaPasswordDigestV11_(salt,password){
  // Ishchi hisobning alohida tasodifiy tuzlangan xeshi; parol hech qayerga log qilinmaydi.
  let result=salt+':'+String(password);
  for(let i=0;i<7000;i++)result=hash_(result);
  return result;
}
function umbrellaAuthUserV11_(initData){return miniAuthV10_(initData);}
function miniOperatorSetPasswordV11(initData,password){
  const person=umbrellaAuthUserV11_(initData),u=umbrellaUserV11_(person.id);
  if(!u||u.role!=='operator'||u.status!=='active')throw Error('Sizga operator huquqi berilmagan');
  if(u.passhash)throw Error('Parol allaqachon mavjud. Parolni tiklash uchun administratorga murojaat qiling.');
  if(typeof password!=='string'||password.length<12||password.length>128)throw Error('Parol 12–128 belgidan iborat bo‘lsin');
  const salt=Utilities.getUuid(),digest=umbrellaPasswordDigestV11_(salt,password);
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const current=umbrellaUserV11_(person.id);if(!current||current.passhash)throw Error('Parolni o‘rnatish mumkin emas');
    umbrellaUsersSheetV11_().getRange(current.row,8,1,3).setValues([[salt,digest,Date.now()+30*24*3600*1000]]);
    audit_('PAROL_ORNATILDI',person.id,'','',0,'Operator paroli o‘rnatildi');
  }finally{lock.releaseLock();}
  return true;
}
function miniOperatorLoginV11(initData,password){
  const person=umbrellaAuthUserV11_(initData),u=umbrellaUserV11_(person.id);
  if(!u||u.role!=='operator'||u.status!=='active'||!u.passhash)throw Error('Operator huquqi yoki parol hali yo‘q');
  const cache=CacheService.getScriptCache(),key='loginfail:'+person.id,failures=Number(cache.get(key)||0);
  if(failures>=5)throw Error('Ko‘p urinish. 15 daqiqadan so‘ng qaytaring.');
  const given=umbrellaPasswordDigestV11_(u.salt,String(password||''));
  if(given!==u.passhash){cache.put(key,String(failures+1),900);throw Error('Parol noto‘g‘ri');}
  cache.remove(key);
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{const current=umbrellaUserV11_(person.id);umbrellaUsersSheetV11_().getRange(current.row,10).setValue(Date.now()+30*24*3600*1000);}
  finally{lock.releaseLock();}
  audit_('KIRISH',person.id,'','',0,'Mini App operator kirishi');
  return true;
}
function miniOperatorLogoutV11(initData){
  const person=umbrellaAuthUserV11_(initData),u=umbrellaUserV11_(person.id);
  if(u)umbrellaUsersSheetV11_().getRange(u.row,10).setValue(0);
  audit_('CHIQISH',person.id,'','',0,'Operator sessiyasi yakunlandi');return true;
}
function miniOperatorLogoutV11_byId_(id){
  const u=umbrellaUserV11_(id);if(u)umbrellaUsersSheetV11_().getRange(u.row,10).setValue(0);
  audit_('CHIQISH',id,'','',0,'Bot orqali chiqish');
}
function setupUmbrellaUsersV11(){
  umbrellaUsersSheetV11_();
  Logger.log('Foydalanuvchilar varag‘i tayyor. Botga /start yuborib ro‘yxatdan o‘ting. Token yoki parol kiritilmaydi.');
}

function getMode_(id){return CacheService.getScriptCache().get('mode:'+id)||'';}
function setMode_(id,mode){CacheService.getScriptCache().put('mode:'+id,mode,1800);}
function clearMode_(id){CacheService.getScriptCache().remove('mode:'+id);}
function parseCmd_(text){
  const m=String(text||'').trim().match(/^([A-Za-z0-9_-]{1,100})\s+([A-Za-z0-9_.-]{1,100})\s+(\d{1,7})$/);
  if(!m)return null;
  const qty=Number(m[3]);return qty>0&&Number.isSafeInteger(qty)?{project:normalizeKey_(m[1]),code:normalizeKey_(m[2]),qty}:null;
}
function applyReceipt_(user,cmd,sign,confirmed){
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const data=readDb_();
    const prod=data.products.find(p=>normalizeKey_(p.project)===cmd.project&&normalizeKey_(p.code)===cmd.code);
    if(!prod)return {error:'Loyiha yoki mahsulot kodi topilmadi. Admin Excel rejasini tekshiring.'};
    const daily=data.receipts.filter(r=>r.productId===prod.id).reduce((s,r)=>s+Number(r.qty||0),0);
    const current=Number(prod.opening||0)+daily;
    if(sign<0 && daily<cmd.qty)return {error:`Orqaga qaytarish mumkin emas: sanali prixodda faqat ${daily} dona bor. Boshlang‘ich Excel kirimi o‘zgarmaydi.`};
    const after=current+sign*cmd.qty;
    if(sign>0 && after>Number(prod.plan) && !confirmed)return {needConfirm:true,plan:prod.plan,current,after,prod};
    const id=Date.now()*100+Math.floor(Math.random()*100);
    data.receipts.push({id,productId:prod.id,qty:sign*cmd.qty,date:Utilities.formatDate(new Date(),hisobTZ_(),'yyyy-MM-dd'),note:sign<0?'Telegram: prixodni orqaga qaytarish':'Telegram: prixod',source:'TELEGRAM',userId:String(user),createdAt:now_()});
    writeRows_(HISOB.RECEIPT,HISOB.RH,data.receipts);
    const rev=bumpRevision_();
    audit_(sign>0?'TELEGRAM_KIRIM':'TELEGRAM_QAYTARISH',user,prod.project,prod.code,sign*cmd.qty,'Rev '+rev);
    SpreadsheetApp.flush();
    return {prod,qty:cmd.qty,sign,current:after,excess:Math.max(0,after-prod.plan)};
  }finally{lock.releaseLock();}
}
function formatReceipt_(r){
  if(r.error)return '❌ '+r.error;
  const p=r.prod;
  return `${r.sign>0?'✅ Prixod qabul qilindi':'↩️ Prixoddan ayrildi'}\nLoyiha: ${p.project}\nKod: ${p.code}\nProfil: ${p.profile||'—'}\nMiqdor: ${r.sign>0?'+':'−'}${r.qty} dona\nJami vazn: ${(r.qty*Number(p.kg||0)).toFixed(2)} kg\nHozirgi jami: ${r.current} dona\nQoldiq: ${p.plan-r.current} dona${r.excess?'\n⚠️ Rejadan '+r.excess+' dona ortiq!':''}`;
}
function inlineReportMenu_(){return {inline_keyboard:[
  [{text:'📋 Plan',callback_data:'report:plan'},{text:'📦 Prixod',callback_data:'report:receipt'}],
  [{text:'📉 Qoldiq',callback_data:'report:remain'},{text:'📊 Jami',callback_data:'report:total'}],
  [{text:'📊 Dashboard rasmlari',callback_data:'dash:menu'}]
]};}
function reportKindNameV9_(kind){return ({plan:'Plan',receipt:'Prixod',remain:'Qoldiq',total:'Jami'})[kind]||'Hisobot';}
function allProjectsLabelV9_(kind){return '📥 Jami — barcha loyihalar ('+reportKindNameV9_(kind)+')';}
function showReportProjectsV9_(chat,id,kind,page){
  if(!['plan','receipt','remain','total'].includes(kind)){msg_(chat,'Hisobot turi noto‘g‘ri.',inlineReportMenu_());return;}
  setMode_(id,'report:'+kind);
  const names=[...new Set(readDb_().products.map(p=>normalizeKey_(p.project)))].sort();
  if(!names.length){msg_(chat,'Hisobot uchun mahsulotlar bazasi bo‘sh.',inlineReportMenu_());return;}
  // Har bir sahifada 3 qator, har qatorda ko‘pi bilan 4 ta loyiha tugmasi.
  // Masalan, 10 loyiha => 4 / 3 / 3. Telegramda tugmalar juda mayda bo‘lib ketmasligi uchun
  // 12 tadan ortiq loyihalar keyingi sahifaga o‘tadi.
  const pageSize=12,totalPages=Math.ceil(names.length/pageSize);
  const current=Math.min(Math.max(0,Number(page)||0),totalPages-1);
  const start=current*pageSize;
  const visible=names.slice(start,start+pageSize);
  const buttons=[[{text:allProjectsLabelV9_(kind),callback_data:'report:all:'+kind}]];
  const rows=Math.min(3,visible.length);
  let position=0;
  for(let row=0;row<rows;row++){
    const count=Math.ceil((visible.length-position)/(rows-row));
    const keys=visible.slice(position,position+count).map((project,j)=>({
      text:project,
      callback_data:'report:pick:'+kind+':'+(start+position+j)
    }));
    buttons.push(keys);
    position+=count;
  }
  if(totalPages>1){
    const navigation=[];
    if(current>0)navigation.push({text:'⬅️ Oldingi',callback_data:'report:page:'+kind+':'+(current-1)});
    navigation.push({text:(current+1)+' / '+totalPages,callback_data:'report:page:'+kind+':'+current});
    if(current+1<totalPages)navigation.push({text:'Keyingi ➡️',callback_data:'report:page:'+kind+':'+(current+1)});
    buttons.push(navigation);
  }
  buttons.push([{text:'⬅️ Hisobot menyusi',callback_data:'report:menu'}]);
  // Bo‘limga kirganda faqat o‘sha bo‘limning nomi chiqadi; loyihalar tugmalarda qoladi.
  const sectionTitle=({plan:'📋 Plan',receipt:'📦 Prixod',remain:'📉 Qoldiq',total:'📊 Jami'})[kind];
  msg_(chat,sectionTitle,{inline_keyboard:buttons});
}

function doPost(e){
  // Apps Script doPost Telegram secret_token headerini bermaydi; URLdagi sir tekshiriladi.
  const secret=cfg_('WEBHOOK_SECRET');
  if(!secret||String((e&&e.parameter&&e.parameter.webhook)||'')!==secret)return HtmlService.createHtmlOutput('forbidden');
  try{
    const update=JSON.parse(e.postData.contents);
    const key='tg:'+String(update.update_id);
    if(CacheService.getScriptCache().get(key))return HtmlService.createHtmlOutput('ok');
    handleUpdate_(update);
    CacheService.getScriptCache().put(key,'1',21600);
  }catch(err){console.error('Bot error: '+err.message);}
  return HtmlService.createHtmlOutput('ok');
}
function handleUpdate_(update){
  const cb=update.callback_query;
  const message=update.message;
  const sender=cb?cb.from:message?message.from:null;
  const chat=cb?cb.message&&cb.message.chat:message&&message.chat;
  if(!sender||!chat||chat.type!=='private')return;
  const id=sender.id;
  let profile=umbrellaUserV11_(id);
  const incoming=String(message&&message.text||'').trim();
  const modeBefore=getMode_(id);
  if(!profile&&message&&incoming==='❌ Bekor qilish'){clearMode_(id);umbrellaRegIntroV11_(chat.id,id);return;}
  if(!profile){
    if(cb){botCall_('answerCallbackQuery',{callback_query_id:cb.id});umbrellaRegIntroV11_(chat.id,id);return;}
    if(message&&message.contact){
      if(modeBefore!=='reg_contact'){msg_(chat.id,'Avval /start yuborib ismingizni kiriting.');return;}
      const name=CacheService.getScriptCache().get('regname:'+id);
      if(!name){umbrellaRegIntroV11_(chat.id,id);return;}
      try{
        const registered=umbrellaRegisterV11_(id,name,message.contact);
        clearMode_(id);CacheService.getScriptCache().remove('regname:'+id);
        if(registered.owner)msg_(chat.id,'✅ Administrator sifatida ro‘yxatdan o‘tdingiz.',umbrellaMenuV11_(id));
        else msg_(chat.id,'⏳ Ro‘yxatdan o‘tdingiz, '+registered.name+'.\nAdministrator tasdiqlamaguncha hech qanday hisobot, loyiha yoki mahsulot ma’lumoti ochilmaydi.',umbrellaPendingMenuV12_());
        if(registered.owner)msg_(chat.id,'🔐 Siz asosiy adminsiz. Mini App orqali 12+ belgili shaxsiy operator parolini o‘rnating.');
        else umbrellaOwnersV11_().forEach(function(owner){try{msg_(Number(owner),'👤 UMBRELLA — yangi ruxsat so‘rovi\nIsm: '+registered.name+'\nTelefon: '+registered.phone+'\nTelegram ID: '+id+'\n⏳ Holat: TASDIQ KUTILMOQDA',umbrellaApprovalButtonsV12_(id));}catch(ex){console.error('Admin xabari yuborilmadi: '+ex.message);}});
      }catch(e){msg_(chat.id,'❌ '+e.message,umbrellaContactKeyboardV11_());}
      return;
    }
    if(modeBefore==='reg_name'&&incoming&&!incoming.startsWith('/')){
      try{
        const name=umbrellaNameV11_(incoming);
        CacheService.getScriptCache().put('regname:'+id,name,1800);
        setMode_(id,'reg_contact');
        msg_(chat.id,'Rahmat, '+name+'. Endi pastdagi tugma orqali O‘Z telefon raqamingizni yuboring.',umbrellaContactKeyboardV11_());
      }catch(e){msg_(chat.id,'❌ '+e.message);}
      return;
    }
    if(modeBefore==='reg_contact'){msg_(chat.id,'📱 Faqat «Telefon raqamimni yuborish» tugmasini bosing.',umbrellaContactKeyboardV11_());return;}
    umbrellaRegIntroV11_(chat.id,id);return;
  }
  if(!umbrellaClosedReadyV12_()){
    if(cb)botCall_('answerCallbackQuery',{callback_query_id:cb.id});
    msg_(chat.id,'⛔ Yopiq tizim faollashtirilmoqda. Administratorga murojaat qiling.',umbrellaPendingMenuV12_());return;
  }
  // Yopiq ruxsat darvozasi BARCHA callback, eski tugma, buyruq va hujjatdan oldin turadi.
  if(!userAllowed_(id)){
    if(cb)botCall_('answerCallbackQuery',{callback_query_id:cb.id});
    const t=profile.status==='pending'?'⏳ Tasdiq kutilmoqda. Admin ruxsat bermaguncha hech qanday ma’lumot ko‘rinmaydi.':
      profile.status==='rejected'?'⛔ So‘rovingiz rad etilgan. Administratorga murojaat qiling.':'⛔ Sizga kirish taqiqlangan.';
    msg_(chat.id,t,umbrellaPendingMenuV12_());return;
  }
  if(cb&&/^(v11:(approve|viewer)|v12:deny):\d+$/.test(cb.data)){
    botCall_('answerCallbackQuery',{callback_query_id:cb.id});
    try{
      const parts=cb.data.split(':'),decision=parts[0]==='v12'?'deny':(parts[1]==='approve'?'operator':'viewer');
      const u=umbrellaAdminDecisionV12_(id,parts[2],decision);
      msg_(chat.id,'✅ '+u.name+' — '+(decision==='deny'?'ruxsat rad etildi':decision==='operator'?'Operator tasdiqlandi':'Hisobotga ruxsat berildi')+'.');
      const targetText=decision==='deny'?'⛔ Administrator ruxsatni rad etdi.':
        decision==='operator'?'✅ Operator huquqi berildi. Mini Appda parol o‘rnating va kirishingiz mumkin.':
        '✅ Hisobotlarga ruxsat berildi.';
      msg_(Number(parts[2]),targetText,decision==='deny'?umbrellaPendingMenuV12_():umbrellaMenuV11_(parts[2]));
    }catch(e){msg_(chat.id,'❌ '+e.message);}return;
  }
  if(message&&(incoming==='⏳ Ruxsat so‘rovlari'||incoming==='/pending')){
    if(!umbrellaIsOwnerV11_(id)){msg_(chat.id,'⛔ Faqat administratorga ruxsat.');return;}
    const users=umbrellaPendingUsersV12_(10);
    if(!users.length){msg_(chat.id,'✅ Tasdiq kutayotgan foydalanuvchi yo‘q.');return;}
    msg_(chat.id,'⏳ Tasdiq kutayotgan foydalanuvchilar: '+users.length+' ta.');
    users.forEach(function(u){
      msg_(chat.id,'👤 '+String(u[1]||'')+'\n☎️ '+String(u[2]||'')+'\nID: '+u[0],umbrellaApprovalButtonsV12_(String(u[0])));
    });
    return;
  }
  // Xabar tipidan qat’i nazar, eski tugmalar ham qayta tekshiriladi.
  const protectedCallback=cb&&(/^over:/.test(cb.data));
  const protectedText=message&&(['📦 Prixod','↩️ Prixodni qaytarish','📥 Plan qo‘shish'].includes(incoming)||message.document);
  const activeMode=getMode_(id);
  if((protectedCallback||protectedText||(!cb&&['add','undo','plan_excel'].includes(activeMode)&&incoming&&!incoming.startsWith('/')))&&!umbrellaCanWriteV11_(id)){
    if(cb)botCall_('answerCallbackQuery',{callback_query_id:cb.id});
    clearMode_(id);msg_(chat.id,'🔒 Bu amal uchun operator huquqi va Mini App ichida parol bilan kirish kerak.',umbrellaMenuV11_(id));return;
  }
  if(message&&incoming==='🚪 Chiqish'){
    miniOperatorLogoutV11_byId_(id);
    clearMode_(id);msg_(chat.id,'Sessiyadan chiqdingiz.',umbrellaMenuV11_(id));return;
  }
  if(message&&incoming==='🔐 Operator kirishi'){
    if(profile.role!=='operator'){msg_(chat.id,'Hozir siz faqat hisobot ko‘rasiz. Administrator operator huquqini berishi kerak.');return;}
    msg_(chat.id,'🔐 Mini App tugmasi orqali shaxsiy parolingizni o‘rnating yoki kiriting:',miniLaunchButtonsV10_());return;
  }
  if(message&&incoming==='👤 Profil'){umbrellaShowProfileV11_(chat.id,id);return;}
  if(message&&incoming.startsWith('/faoliyat ')){
    if(!umbrellaIsOwnerV11_(id)){msg_(chat.id,'Faqat asosiy administratorga ruxsat.');return;}
    const q=incoming.slice('/faoliyat '.length).trim().toLowerCase();
    if(!q){msg_(chat.id,'Qidiruv: /faoliyat Ism');return;}
    const matches=umbrellaRecentActivityV11_(40).filter(x=>x.toLowerCase().includes(q)).slice(0,15);
    msg_(chat.id,'📋 '+q+' bo‘yicha oxirgi amallar:\n'+(matches.join('\n')||'Mos yozuv topilmadi.'));
    return;
  }
  if(message&&(incoming==='📋 Faoliyat jurnali'||incoming==='/faoliyat')){
    if(!umbrellaIsOwnerV11_(id)){msg_(chat.id,'Faqat asosiy administratorga ruxsat.');return;}
    msg_(chat.id,'📋 Oxirgi amallar:\n'+(umbrellaRecentActivityV11_(15).join('\n')||'Hali amal yo‘q'));return;
  }

  if(cb){
    botCall_('answerCallbackQuery',{callback_query_id:cb.id});
    if(cb.data==='dash:menu'||cb.data==='dash:back'){
      msg_(chat.id,'📊 HISOB Dashboard bo‘limini tanlang:',cb.data==='dash:menu'?dashboardMenu_():inlineReportMenu_());return;
    }
    if(cb.data.indexOf('dash:')===0){
      const section=cb.data.slice(5);
      // Dashboard rasmlari Google Sheets ma'lumotlaridan navbat orqali tayyorlanadi.
      if(['main','projects','daily','report'].includes(section)){
        queueDashboardV8_(chat.id,section,cb.id);
        msg_(chat.id,'🖼 HISOB '+dashboardLabelV8_(section)+' tayyorlanmoqda. Rasm avtomatik keladi (odatda 1–2 daqiqada).');
        return;
      }
      msg_(chat.id,'Noma’lum Dashboard bo‘limi.',dashboardMenu_());return;
    }
    if(cb.data==='report:menu'){clearMode_(id);msg_(chat.id,'📊 Hisobot turini tanlang:',inlineReportMenu_());return;}
    if(cb.data.indexOf('report:page:')===0){const parts=cb.data.split(':');showReportProjectsV9_(chat.id,id,parts[2],Number(parts[3]));return;}
    if(cb.data.indexOf('report:all:')===0){const kind=cb.data.slice('report:all:'.length);sendReport_(chat.id,kind,null);return;}
    if(cb.data.indexOf('report:pick:')===0){
      const parts=cb.data.split(':'),kind=parts[2],ix=Number(parts[3]);
      const names=[...new Set(readDb_().products.map(p=>normalizeKey_(p.project)))].sort();
      if(!Number.isInteger(ix)||ix<0||ix>=names.length){msg_(chat.id,'Loyiha topilmadi. Hisobotni qayta oching.');return;}
      sendReport_(chat.id,kind,[names[ix]]);return;
    }
    if(['report:plan','report:receipt','report:remain','report:total'].includes(cb.data)){
      showReportProjectsV9_(chat.id,id,cb.data.split(':')[1]);return;
    }
    // Eski versiyadan kelgan ochiq inline tugmalar ham ishlaydi.
    if(cb.data.indexOf('file:')===0){const parts=cb.data.split(':');sendReport_(chat.id,parts[1],[parts.slice(2).join(':')]);return;}
    if(cb.data==='over:yes'){umbrellaRequireWriteV11_(id);
      const pending=CacheService.getScriptCache().get('pending:'+id);
      if(!pending){msg_(chat.id,'Tasdiqlash muddati tugadi. Qayta yuboring.');return;}
      CacheService.getScriptCache().remove('pending:'+id);
      const batch=JSON.parse(pending),r=applyBatch_(id,batch.commands,batch.sign,true);
      msg_(chat.id,formatBatch_(r),umbrellaMenuV11_(id));return;
    }
    if(cb.data==='over:no'){CacheService.getScriptCache().remove('pending:'+id);msg_(chat.id,'Bekor qilindi.',umbrellaMenuV11_(id));return;}
    return;
  }
  const text=String(message.text||'').trim();
  const mode=getMode_(id);
  if(message.document){
    if(mode!=='plan_excel'){msg_(chat.id,'Excel reja uchun avval 📥 Plan qo‘shish tugmasini bosing.');return;}
    msg_(chat.id,'📥 Excel qabul qilindi. Tekshirilmoqda, biroz kuting...');
    const response=importPlanDocument_(message.document,id);
    msg_(chat.id,response,umbrellaMenuV11_(id));return;
  }
  if(text==='/start'||text==='/menu'||text==='🏠 Menyu'){
    clearMode_(id);
    msg_(chat.id,'HISOB bot. Salom, '+profile.name+'! Kerakli bo‘limni tanlang.',umbrellaMenuV11_(id));
    msg_(chat.id,'📱 Ilovani Telegram ichida oching:',miniLaunchButtonsV10_());
    return;
  }
  if(text==='📱 Ilovani ochish'){
    msg_(chat.id,'📱 Quyidagi tugmani bosing:',miniLaunchButtonsV10_());return;
  }
  if(text==='📦 Prixod'){setMode_(id,'add');msg_(chat.id,'Prixod: har bir mahsulotni yangi qatordan kiriting.\nMisol:\nCKL-17 P283 1\nCKL-17 P809 1\nCKL-17 P1043 2');return;}
  if(text==='↩️ Prixodni qaytarish'){setMode_(id,'undo');msg_(chat.id,'Har bir qaytarishni yangi qatordan yozing.\nMisol:\nCKL-17 P283 1\nCKL-17 P809 2\nFaqat sanali kirimlardan qaytarish mumkin.');return;}
  if(text==='📥 Plan qo‘shish'){setMode_(id,'plan_excel');msg_(chat.id,'📥 Yangi reja faylini .xlsx shaklida yuboring.\nQo‘llab-quvvatlanadi: HISOB (Loyiha/Kod/Plan/Kg), yoki POZ/PROFİL/ADET/AĞIRLIK jadvali.\nMavjud kodlar ustidan yozilmaydi; konflikt bo‘lsa import rad etiladi.');return;}
  if(text==='📊 Hisobot'){clearMode_(id);msg_(chat.id,'Hisobot turini tanlang:',inlineReportMenu_());return;}
  if(text==='📖 Yordam'){msg_(chat.id,'📦 Ko‘p qatorli prixod: har qatorda LOYIHA KOD DONA.\n↩️ Qaytarish: shu format.\n📥 Plan qo‘shish: .xlsx fayl.\n📊 Hisobot: Plan / Prixod / Qoldiq / Jami Excel va Dashboard.\n📊 Dashboard ichida loyihalar, kunlik grafik, umumiy hisobot.',umbrellaMenuV11_(id));return;}
  if(mode.indexOf('report:')===0){
    const want=text.toLowerCase();
    sendReport_(chat.id,mode.slice(7),['jami','barchasi','all'].includes(want)?null:text.split(/[\s,;]+/).filter(Boolean));return;
  }
  if(mode==='plan_excel'){msg_(chat.id,'📎 Hujjat sifatida .xlsx Excel faylini yuboring yoki /start bosing.');return;}
  if(mode!=='add'&&mode!=='undo'){msg_(chat.id,'Avval menyudan Prixod, Qaytarish, Plan qo‘shish yoki Hisobotni tanlang.',umbrellaMenuV11_(id));return;}
  const parsed=parseBatch_(text);
  if(parsed.error){msg_(chat.id,parsed.error);return;}
  const sign=mode==='add'?1:-1,result=applyBatch_(id,parsed.commands,sign,false);
  if(result.needConfirm){
    CacheService.getScriptCache().put('pending:'+id,JSON.stringify({commands:parsed.commands,sign}),600);
    msg_(chat.id,'⚠️ '+result.excessText+'\nRejadan ortiq. Baribir saqlaysizmi?',{inline_keyboard:[[{text:'✅ Tasdiqlash',callback_data:'over:yes'},{text:'❌ Bekor',callback_data:'over:no'}]]});return;
  }
  msg_(chat.id,formatBatch_(result),umbrellaMenuV11_(id));
}
// Bir xabarda 1–60 ta mahsulotni qabul qilish. Noto‘g‘ri qator bo‘lsa HAMMASI rad etiladi.
function parseBatch_(text){
  const lines=String(text||'').replace(/;/g,'\n').split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
  if(!lines.length||lines.length>60)return {error:'❌ 1 tadan 60 tagacha satr yuboring.'};
  const commands=[],errors=[];
  lines.forEach((line,i)=>{const cmd=parseCmd_(line);if(!cmd)errors.push((i+1)+'-qator: '+line);else commands.push(cmd);});
  if(errors.length)return {error:'❌ Format noto‘g‘ri. Birorta prixod saqlanmadi.\n'+errors.slice(0,12).join('\n')+'\nMisol: CKL-17 P283 1'};
  return {commands};
}
function projectKey_(s){return normalizeKey_(s).replace(/[-_]/g,'');}
function applyBatch_(user,commands,sign,confirmed){
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const db=readDb_(),products=new Map(),totals=new Map(),errors=[];
    db.products.forEach(p=>products.set(projectKey_(p.project)+'|'+normalizeKey_(p.code),p));
    commands.forEach((cmd,i)=>{
      const product=products.get(projectKey_(cmd.project)+'|'+normalizeKey_(cmd.code));
      if(!product){errors.push((i+1)+'-qator: '+cmd.project+' '+cmd.code+' — topilmadi');return;}
      if(sign>0 && db.projectStates[product.project]==='paused'){errors.push((i+1)+'-qator: '+product.project+' — plan to‘xtatilgan');return;}
      const id=String(product.id),entry=totals.get(id)||{product,qty:0};
      entry.qty+=Number(cmd.qty);totals.set(id,entry);
    });
    if(errors.length)return {error:'❌ Bitta ham mahsulot saqlanmadi:\n'+errors.slice(0,15).join('\n')};
    const daily=new Map();db.receipts.forEach(r=>daily.set(String(r.productId),(daily.get(String(r.productId))||0)+Number(r.qty)));
    const excessive=[];
    for(const entry of totals.values()){
      const {product:p,qty}=entry,got=daily.get(String(p.id))||0;
      const current=Number(p.opening||0)+got;
      if(sign<0&&got<qty)errors.push(p.project+' '+p.code+': sanali prixod '+got+' dona, '+qty+' qaytarib bo‘lmaydi');
      if(sign>0&&current+qty>Number(p.plan||0))excessive.push(p.project+' '+p.code+' ('+(current+qty)+' / '+p.plan+')');
    }
    if(errors.length)return {error:'❌ Birorta amal saqlanmadi:\n'+errors.join('\n')};
    if(excessive.length&&!confirmed)return {needConfirm:true,excessText:excessive.slice(0,10).join('\n')};
    const today=Utilities.formatDate(new Date(),hisobTZ_(),'yyyy-MM-dd');
    let nextId=Date.now()*1000; db.receipts.forEach(r=>{nextId=Math.max(nextId,Number(r.id)||0);}); nextId+=1;
    const written=[],rows=[];
    for(const {product:p,qty} of totals.values()){
      const after=Number(p.opening||0)+(daily.get(String(p.id))||0)+sign*qty;
      const rec={id:nextId++,productId:p.id,qty:sign*qty,date:today,note:sign>0?'Telegram: prixod':'Telegram: prixodni qaytarish',source:'TELEGRAM',userId:String(user),createdAt:now_()};
      rows.push(HISOB.RH.map(k=>rec[k]));written.push({prod:p,qty,sign,current:after,excess:Math.max(0,after-p.plan)});
    }
    const sh=tab_(HISOB.RECEIPT,HISOB.RH),start=sh.getLastRow()+1;
    if(sh.getMaxRows()<start+rows.length-1)sh.insertRowsAfter(sh.getMaxRows(),start+rows.length-1-sh.getMaxRows());
    sh.getRange(start,1,rows.length,HISOB.RH.length).setValues(rows);
    const rev=bumpRevision_();
    audit_(sign>0?'TELEGRAM_OMMAVIY_KIRIM':'TELEGRAM_OMMAVIY_QAYTARISH',user,'','',sign*written.reduce((s,x)=>s+x.qty,0),'Satrlar '+commands.length+', mahsulot '+written.length+', rev '+rev);
    umbrellaAuditItemsV11_(user,written,sign,rev);
    SpreadsheetApp.flush();return {items:written,sign,revision:rev};
  }finally{lock.releaseLock();}
}
function formatBatch_(r){
  if(r.error)return r.error;
  const items=r.items||[];
  if(items.length===1)return formatReceipt_(items[0]);
  const qty=items.reduce((s,x)=>s+x.qty,0),kg=items.reduce((s,x)=>s+x.qty*Number(x.prod.kg||0),0);
  const heading=r.sign>0?'✅ Ommaviy prixod qabul qilindi':'↩️ Ommaviy qaytarish qabul qilindi';
  return heading+'\nMahsulotlar: '+items.length+' ta\nJami: '+qty+' dona · '+kg.toFixed(2)+' kg\n\n'+items.slice(0,45).map(x=>x.prod.project+' '+x.prod.code+' '+(r.sign>0?'+':'−')+x.qty+' dona → qoldiq '+(x.prod.plan-x.current)).join('\n')+(items.length>45?'\n...':'');
}
// XLSX → Google Sheets. Asl fayl o'zgarmaydi, vaqtinchalik nusxa o'chiriladi.
function xlsxToSheet_(blob){
  const boundary='----UmbrellaImport'+Date.now();
  const info=JSON.stringify({name:'HISOB_import_'+Date.now(),mimeType:'application/vnd.google-apps.spreadsheet'});
  const prefix=Utilities.newBlob('--'+boundary+'\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n'+info+'\r\n--'+boundary+'\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n').getBytes();
  const suffix=Utilities.newBlob('\r\n--'+boundary+'--').getBytes();
  const bytes=prefix.concat(blob.getBytes(),suffix);
  const response=UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,mimeType',{
    method:'post',contentType:'multipart/related; boundary='+boundary,
    headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},
    payload:bytes,muteHttpExceptions:true
  });
  if(response.getResponseCode()!==200&&response.getResponseCode()!==201)throw Error('Excelni vaqtincha ochib bo‘lmadi (Drive HTTP '+response.getResponseCode()+'). Google Drive API ruxsatini tekshiring.');
  const meta=JSON.parse(response.getContentText());if(!meta.id)throw Error('Excel konvertatsiya ID topilmadi');
  return SpreadsheetApp.openById(meta.id);
}
function findPlanColumns_(headers){
  const cells=headers.map(v=>String(v||'').trim().toUpperCase().normalize('NFKD').replace(/[\u0300-\u036f]/g,''));
  function find(keys){return cells.findIndex(v=>keys.some(k=>v===k||v.replace(/\s+/g,'')===k.replace(/\s+/g,'')));}
  const col={
    project:find(['PROJECT','LOYIHA','ПРОЕКТ','PROJE']),code:find(['CODE','KOD','KODI','КОД','ОБОЗНАЧЕНИЕ','POZ','POSITION']),
    plan:find(['PLAN','REJA','ПЛАН','ADET','QUANTITY','QTY']),kg:find(['KG','КГ','1 ДОНА КГ','AĞIRLIK','AGIRLIK','WEIGHT']),
    profile:find(['PROFILE','PROFIL','ПРОФИЛЬ','PROFİL']),length:find(['LENGTH','ДЛИНА (ММ)','ДЛИНА','BOY','UZUNLIK']),
    material:find(['MATERIAL','MALZEME','MATERİAL','MATERИАЛ','МАТЕРИАЛ']),unit:find(['UNIT','ED.IZ','ЕД.ИЗ','BIRLIK']),
    opening:find(['OPENING','PRIXOD','ПРИХОД','KIRIM']),totalKg:find(['TOPLAM AĞIRLIK','TOPLAM AGIRLIK','TOTAL WEIGHT'])
  };
  return col.code>=0&&col.plan>=0?col:null;
}
function planNumber_(v){
  if(typeof v==='number')return v;
  const s=String(v===null||v===undefined?'':v).replace(/\s/g,'').replace(/,/g,'.');return s===''?NaN:Number(s);
}
function extractPlan_(workbook){
  const results=[],errors=[],seen=new Map();
  workbook.getSheets().forEach(sheet=>{
    if(sheet.isSheetHidden())return;
    const grid=sheet.getDataRange().getValues().slice(0,6000);
    let hi=-1,c=null;
    for(let i=0;i<Math.min(40,grid.length);i++){const found=findPlanColumns_(grid[i]);if(found){hi=i;c=found;break;}}
    if(!c)return;
    for(let i=hi+1;i<grid.length;i++){
      const row=grid[i],code=String(row[c.code]||'').trim();
      if(!code)continue;
      if(/^TOTAL|^JAMI|^ИТОГО|^ВСЕГО/i.test(code))continue;
      const project=c.project>=0?String(row[c.project]||'').trim():sheet.getName().trim();
      if(!project||!/^[-\wА-Яа-я]+$/.test(project)||!/^[-\w.]+$/.test(code)){errors.push(sheet.getName()+': '+(i+1)+'-qator loyiha/kod noto‘g‘ri');continue;}
      const qty=planNumber_(row[c.plan]);
      if(!Number.isInteger(qty)||qty<=0){errors.push(sheet.getName()+': '+(i+1)+'-qator reja miqdori noto‘g‘ri');continue;}
      let kg=c.kg>=0?planNumber_(row[c.kg]):NaN;
      if(!Number.isFinite(kg)&&c.totalKg>=0)kg=planNumber_(row[c.totalKg])/qty;
      if(!Number.isFinite(kg)||kg<0){errors.push(sheet.getName()+': '+(i+1)+'-qator 1 dona vazni yo‘q');continue;}
      let opening=c.opening>=0?planNumber_(row[c.opening]):0;
      if(!Number.isFinite(opening))opening=0;
      if(!Number.isInteger(opening)||opening<0){errors.push(sheet.getName()+': '+(i+1)+'-qator boshlang‘ich kirim noto‘g‘ri');continue;}
      const prod={project,code,plan:qty,kg,opening,unit:c.unit>=0?String(row[c.unit]||'шт'):'шт',
        profile:c.profile>=0?String(row[c.profile]||''):'',length:c.length>=0?Number(planNumber_(row[c.length]))||0:0,
        material:c.material>=0?String(row[c.material]||''):''};
      const key=projectKey_(project)+'|'+normalizeKey_(code);
      if(seen.has(key)){errors.push(sheet.getName()+': '+(i+1)+'-qator takroriy kod '+project+' '+code);continue;}
      seen.set(key,true);results.push(prod);
      if(results.length>3000)throw Error('3000 tadan ko‘p mahsulot: faylni ajrating');
    }
  });
  if(errors.length)throw Error('Excel import rad etildi, ma’lumot saqlanmadi:\n'+errors.slice(0,12).join('\n'));
  if(!results.length)throw Error('Excelda reja ustunlari topilmadi. Kod, Plan/ADET va KG/AĞIRLIK sarlavhalari bo‘lsin.');
  return results;
}
function importPlanDocument_(document,user){
  const filename=String(document.file_name||'');
  if(!/\.xlsx$/i.test(filename))return '❌ Faqat .xlsx Excel faylini yuboring.';
  if(Number(document.file_size||0)>9*1024*1024)return '❌ Excel fayli 9 MB dan kichik bo‘lsin.';
  let tempId='';
  try{
    const f=botCall_('getFile',{file_id:document.file_id});
    const dl=UrlFetchApp.fetch('https://api.telegram.org/file/bot'+botToken_()+'/'+f.file_path,{muteHttpExceptions:true});
    if(dl.getResponseCode()!==200)throw Error('Telegramdan faylni olib bo‘lmadi: HTTP '+dl.getResponseCode());
    const blob=dl.getBlob().setName(filename),digest=Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,blob.getBytes()));
    const hist=tab_('Plan_import_tarixi',['hash','name','count','date','user']);
    if(hist.getLastRow()>1&&hist.getRange(2,1,hist.getLastRow()-1,1).getValues().some(r=>r[0]===digest))return 'ℹ️ Bu Excel ilgari import qilingan. Takroriy yozuv qo‘shilmadi.';
    const converted=xlsxToSheet_(blob);tempId=converted.getId();
    const products=extractPlan_(converted);
    const lock=LockService.getScriptLock();lock.waitLock(30000);
    try{
      const db=readDb_(),existing=new Map(db.products.map(p=>[projectKey_(p.project)+'|'+normalizeKey_(p.code),p]));
      const conflicts=[],newRows=[];
      products.forEach(p=>{
        const old=existing.get(projectKey_(p.project)+'|'+normalizeKey_(p.code));
        if(old){if(Number(old.plan)!==p.plan||Math.abs(Number(old.kg)-p.kg)>0.001)conflicts.push(p.project+' '+p.code+' (avval: '+old.plan+', yangi: '+p.plan+')');}
        else newRows.push(p);
      });
      if(conflicts.length)return '⛔ Faylda oldingi rejadan farq qiladigan kodlar bor. Import rad etildi; mavjud reja o‘zgarmadi:\n'+conflicts.slice(0,12).join('\n');
      const sh=tab_(HISOB.PRODUCT,HISOB.PH);
      const idMax=Math.max(0,...db.products.map(p=>Number(p.id)||0));
      const matrix=newRows.map((p,i)=>HISOB.PH.map(k=>k==='id'?idMax+i+1:(p[k]===undefined?'':p[k])));
      if(matrix.length){
        const begin=sh.getLastRow()+1;
        if(sh.getMaxRows()<begin+matrix.length-1)sh.insertRowsAfter(sh.getMaxRows(),begin+matrix.length-1-sh.getMaxRows());
        sh.getRange(begin,1,matrix.length,HISOB.PH.length).setValues(matrix);
      }
      hist.appendRow([digest,filename,newRows.length,now_(),String(user)]);
      if(newRows.length)bumpRevision_();
      audit_('TELEGRAM_PLAN_IMPORT',user,'','',newRows.length,filename+', jami '+products.length+' pozitsiya');
      SpreadsheetApp.flush();
      return '✅ Plan Excel qabul qilindi.\nYangi mahsulot: '+newRows.length+' ta\nAvvaldan mavjud (o‘zgarmadi): '+(products.length-newRows.length)+' ta\nHISOB admin va Google Sheets rejalari yangilandi.';
    }finally{lock.releaseLock();}
  }catch(err){console.error('Plan import: '+err.message);return '❌ Import bajarilmadi. '+String(err.message).slice(0,950);}
  finally{if(tempId)try{DriveApp.getFileById(tempId).setTrashed(true);}catch(e){console.error('Temp delete: '+e.message);}}
}
function summaryData_(){
  const db=readDb_(),received=new Map(),daily=new Map();
  db.receipts.forEach(r=>{
    received.set(String(r.productId),(received.get(String(r.productId))||0)+Number(r.qty));
    const date=asDate_(r.date),a=daily.get(date)||{positive:0,returned:0,grossKg:0,returnedKg:0};
    const p=db.products.find(x=>String(x.id)===String(r.productId)),kg=Number(p&&p.kg||0);
    if(r.qty>0){a.positive+=r.qty;a.grossKg+=r.qty*kg;}else{a.returned+=-r.qty;a.returnedKg+=-r.qty*kg;}
    daily.set(date,a);
  });
  const summary=new Map();
  let totalPlan=0,totalGot=0,planKg=0,gotKg=0;
  db.products.forEach(p=>{
    const plan=Number(p.plan||0),got=Number(p.opening||0)+(received.get(String(p.id))||0),kg=Number(p.kg||0);
    const name=p.project,st=summary.get(name)||{project:name,positions:0,plan:0,got:0,planKg:0,gotKg:0};
    st.positions++;st.plan+=plan;st.got+=got;st.planKg+=plan*kg;st.gotKg+=got*kg;
    summary.set(name,st);totalPlan+=plan;totalGot+=got;planKg+=plan*kg;gotKg+=got*kg;
  });
  return {projects:[...summary.values()].sort((a,b)=>a.project.localeCompare(b.project)),daily,totals:{totalPlan,totalGot,planKg,gotKg}};
}
function fmtNum_(n){return Number(n||0).toLocaleString('ru-RU',{maximumFractionDigits:1});}
function dashboardText_(section){
  const s=summaryData_(),t=s.totals;
  if(section==='projects')return '🏗 LOYIHALAR — '+s.projects.length+' ta\n\n'+s.projects.map(p=>p.project+': '+p.got+' / '+p.plan+' dona ('+(p.plan?Math.round(p.got/p.plan*100):0)+'%), qoldiq '+(p.plan-p.got)).join('\n');
  return '📊 HISOB DASHBOARD\n\nReja: '+fmtNum_(t.totalPlan)+' dona · '+fmtNum_(t.planKg)+' kg\nKirim: '+fmtNum_(t.totalGot)+' dona · '+fmtNum_(t.gotKg)+' kg\nQoldiq: '+fmtNum_(t.totalPlan-t.totalGot)+' dona · '+fmtNum_(t.planKg-t.gotKg)+' kg\nBajarilish: '+(t.totalPlan?(100*t.totalGot/t.totalPlan).toFixed(1):0)+'% dona\nLoyihalar: '+s.projects.length+' ta';
}
function dailyLast30_(){
  const daily=summaryData_().daily,arr=[];const todayStr=Utilities.formatDate(new Date(),hisobTZ_(),'yyyy-MM-dd'); const today=new Date(todayStr+'T12:00:00Z');
  for(let k=29;k>=0;k--){const date=new Date(today.getTime()-k*86400000),key=Utilities.formatDate(date,'UTC','yyyy-MM-dd');const item=daily.get(key)||{positive:0,returned:0};arr.push({date:key,positive:item.positive,returned:item.returned,net:item.positive-item.returned});}
  return arr;
}
function sendDailyChart_(chat){
  const days=dailyLast30_(),gross=days.reduce((n,x)=>n+x.positive,0),back=days.reduce((n,x)=>n+x.returned,0);
  const dt=Charts.newDataTable().addColumn(Charts.ColumnType.STRING,'Sana').addColumn(Charts.ColumnType.NUMBER,'Prixod').addColumn(Charts.ColumnType.NUMBER,'Qaytarish');
  days.forEach((x,i)=>dt.addRow([i%3===0?x.date.slice(5):'',x.positive,x.returned]));
  const chart=Charts.newColumnChart().setDataTable(dt.build()).setTitle('Oxirgi 30 kunlik prixod va qaytarish').setDimensions(1120,520).setOption('legend',{position:'bottom'}).build();
  const blob=chart.getAs('image/png').setName('HISOB_30_kun.png');
  const result=UrlFetchApp.fetch('https://api.telegram.org/bot'+botToken_()+'/sendPhoto',{method:'post',payload:{chat_id:String(chat),photo:blob,caption:'📅 30 kun: prixod '+gross+' dona, qaytarildi '+back+' dona, sof '+(gross-back)+' dona. Qaytarish alohida ko‘rsatiladi.'},muteHttpExceptions:true});
  const out=JSON.parse(result.getContentText());if(!out.ok)throw Error('Kunlik grafik yuborilmadi: '+out.description);
}
function sendSummaryDashboard_(chat){
  const s=summaryData_(),t=s.totals;
  msg_(chat,dashboardText_('main'),dashboardMenu_());
  const ss=SpreadsheetApp.create('HISOB_Dashboard_'+Date.now());
  try{
    const sh=ss.getSheets()[0];sh.setName('Dashboard');
    sh.getRange('A1:I1').merge().setValue('HISOB — UMUMIY HISOBOT');
    sh.getRange('A2:I2').merge().setValue(now_());
    sh.getRange('A3:B5').setValues([['Reja dona',t.totalPlan],['Kirim dona',t.totalGot],['Qoldiq dona',t.totalPlan-t.totalGot]]);
    sh.getRange('D3:E5').setValues([['Reja kg',t.planKg],['Kirim kg',t.gotKg],['Qoldiq kg',t.planKg-t.gotKg]]);
    const rows=[['Loyiha','Pozitsiya','Reja dona','Reja kg','Kirim dona','Kirim kg','Qoldiq dona','Qoldiq kg','Bajarilish %']];
    s.projects.forEach(p=>rows.push([p.project,p.positions,p.plan,p.planKg,p.got,p.gotKg,p.plan-p.got,p.planKg-p.gotKg,p.plan?100*p.got/p.plan:0]));
    rows.push(['JAMI',s.projects.reduce((n,x)=>n+x.positions,0),t.totalPlan,t.planKg,t.totalGot,t.gotKg,t.totalPlan-t.totalGot,t.planKg-t.gotKg,t.totalPlan?100*t.totalGot/t.totalPlan:0]);
    sh.getRange(7,1,rows.length,9).setValues(rows);
    sh.getRange(1,1,1,9).setBackground('#08745a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(18);
    sh.getRange(7,1,1,9).setBackground('#08745a').setFontColor('#ffffff').setFontWeight('bold');
    sh.getRange(7+rows.length-1,1,1,9).setBackground('#e1f4e8').setFontWeight('bold');
    sh.setFrozenRows(7);sh.autoResizeColumns(1,9);
    SpreadsheetApp.flush();
    const url='https://docs.google.com/spreadsheets/d/'+ss.getId()+'/export?format=pdf&size=A4&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false';
    const resp=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});
    if(resp.getResponseCode()!==200)throw Error('Dashboard PDF HTTP '+resp.getResponseCode());
    const upload=UrlFetchApp.fetch('https://api.telegram.org/bot'+botToken_()+'/sendDocument',{method:'post',payload:{chat_id:String(chat),document:resp.getBlob().setName('HISOB_Dashboard.pdf'),caption:'📈 HISOB — umumiy hisobot Dashboard'},muteHttpExceptions:true});
    const res=JSON.parse(upload.getContentText());if(!res.ok)throw Error('Dashboard Telegram: '+res.description);
  }finally{DriveApp.getFileById(ss.getId()).setTrashed(true);}
}

/** HISOB V9 — hisobotlar faqat bazadan o‘qiladi; mahsulot va kirimni tahrirlamaydi. */
function reportSheetsV9_(db,kind,projects){
  const selected=db.products.filter(p=>projects.includes(normalizeKey_(p.project)));
  const productById=new Map(selected.map(p=>[String(p.id),p]));
  const delta=new Map(),posted=[];
  db.receipts.forEach(r=>{
    const p=productById.get(String(r.productId)); if(!p)return;
    const qty=Number(r.qty||0),kg=Number(p.kg||0);
    delta.set(String(p.id),(delta.get(String(p.id))||0)+qty);
    posted.push({p,r,qty,kg});
  });
  posted.sort((a,b)=>asDate_(a.r.date).localeCompare(asDate_(b.r.date))||Number(a.r.id)-Number(b.r.id));
  const projectTotals=new Map();
  const plans=[['Loyiha','Kod','Profil','Uzunlik (mm)','Birlik','1 dona (kg)','Reja (dona)','Reja (kg)']];
  const incoming=[['Loyiha','Kod','Profil','Boshlang‘ich kirim (dona)','Sanali kirim sof (dona)','Jami kirim (dona)','Jami kirim (kg)','Reja (dona)','Qoldiq (dona)']];
  const remains=[['Loyiha','Kod','Profil','1 dona (kg)','Reja (dona)','Kirim (dona)','Qoldiq (dona)','Qoldiq (kg)','Holat']];
  selected.forEach(p=>{
    const key=normalizeKey_(p.project),pId=String(p.id),plan=Number(p.plan||0),opening=Number(p.opening||0);
    const net=delta.get(pId)||0,got=opening+net,rest=plan-got,kg=Number(p.kg||0);
    plans.push([p.project,p.code,p.profile,p.length,p.unit||'dona',kg,plan,plan*kg]);
    incoming.push([p.project,p.code,p.profile,opening,net,got,got*kg,plan,rest]);
    remains.push([p.project,p.code,p.profile,kg,plan,got,rest,rest*kg,rest<0?'Rejadan ortiq':rest===0?'Bajarilgan':'Qolgan']);
    const s=projectTotals.get(key)||{name:p.project,positions:0,plan:0,planKg:0,opening:0,delta:0,got:0,gotKg:0};
    s.positions++;s.plan+=plan;s.planKg+=plan*kg;s.opening+=opening;s.delta+=net;s.got+=got;s.gotKg+=got*kg;
    projectTotals.set(key,s);
  });
  const totals=[['Loyiha','Pozitsiyalar','Reja (dona)','Reja (kg)','Boshlang‘ich (dona)','Sanali kirim sof (dona)','Jami kirim (dona)','Jami kirim (kg)','Qoldiq (dona)','Qoldiq (kg)','Bajarilish (%)']];
  let sumPos=0,sumPlan=0,sumPlanKg=0,sumOpen=0,sumNet=0,sumGot=0,sumGotKg=0;
  [...projectTotals.values()].sort((a,b)=>a.name.localeCompare(b.name)).forEach(s=>{
    totals.push([s.name,s.positions,s.plan,s.planKg,s.opening,s.delta,s.got,s.gotKg,s.plan-s.got,s.planKg-s.gotKg,s.plan?100*s.got/s.plan:0]);
    sumPos+=s.positions;sumPlan+=s.plan;sumPlanKg+=s.planKg;sumOpen+=s.opening;sumNet+=s.delta;sumGot+=s.got;sumGotKg+=s.gotKg;
  });
  totals.push(['JAMI',sumPos,sumPlan,sumPlanKg,sumOpen,sumNet,sumGot,sumGotKg,sumPlan-sumGot,sumPlanKg-sumGotKg,sumPlan?100*sumGot/sumPlan:0]);
  const history=[['Sana','Loyiha','Kod','Profil','Amal','Dona (+/-)','Og‘irlik kg (+/-)','Manba','Izoh','Kiritgan','Yozuv ID']];
  posted.forEach(({p,r,qty,kg})=>history.push([asDate_(r.date),p.project,p.code,p.profile,qty>0?'Prixod':'Qaytarish',qty,qty*kg,r.source,r.note,r.userId,String(r.id)]));
  // Boshlang‘ich kirim tarix jadvalida alohida ustunda, sanali operatsiya deb ko‘rsatilmaydi.
  const all={Umumiy:totals,Plan:plans,Prixod:incoming,Qoldiq:remains,'Kirim tarixi':history};
  if(kind==='plan')return [{name:'Umumiy',rows:totals},{name:'Plan',rows:plans}];
  if(kind==='receipt')return [{name:'Umumiy',rows:totals},{name:'Prixod',rows:incoming},{name:'Kirim tarixi',rows:history}];
  if(kind==='remain')return [{name:'Umumiy',rows:totals},{name:'Qoldiq',rows:remains}];
  if(kind==='total')return Object.keys(all).map(name=>({name,rows:all[name]}));
  throw Error('Noma’lum hisobot turi: '+kind);
}
function sendReport_(chat,kind,requested){
  if(umbrellaUserV11_(chat))audit_('EXCEL_HISOBOT',chat,'','',0,'Tur: '+kind+'; loyihalar: '+(requested?requested.join(','):'JAMI'));
  if(!['plan','receipt','remain','total'].includes(kind)){msg_(chat,'Hisobot turi noto‘g‘ri.');return;}
  const db=readDb_(),names=[...new Set(db.products.map(p=>normalizeKey_(p.project)))].sort();
  if(!names.length){msg_(chat,'Bazadagi mahsulotlar ro‘yxati bo‘sh.');return;}
  // null = barcha loyihalar. Qo‘lda yozilgan 1–30 ta loyiha ham qabul qilinadi.
  const all=requested===null||requested===undefined;
  const projects=all?names:[...new Set((requested||[]).map(normalizeKey_).filter(Boolean))];
  if(!projects.length||(!all&&projects.length>30)||projects.some(p=>!names.includes(p))){
    msg_(chat,'Loyiha kodi noto‘g‘ri yoki 30 tadan ko‘p. Mavjudlar: '+names.join(' · '));return;
  }
  const sheets=reportSheetsV9_(db,kind,projects);
  const temp=SpreadsheetApp.create('HISOB_V9_'+Date.now());
  try{
    sheets.forEach((info,i)=>{
      const sh=i===0?temp.getSheets()[0]:temp.insertSheet(info.name);sh.setName(info.name);
      const matrix=info.rows,columns=matrix[0].length;
      if(matrix.length>sh.getMaxRows())sh.insertRowsAfter(sh.getMaxRows(),matrix.length-sh.getMaxRows());
      if(columns>sh.getMaxColumns())sh.insertColumnsAfter(sh.getMaxColumns(),columns-sh.getMaxColumns());
      sh.getRange(1,1,matrix.length,columns).setValues(matrix);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,columns).setBackground('#087A5D').setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);
      sh.setRowHeight(1,38);
      sh.setColumnWidths(1,columns,125);
      sh.setColumnWidth(1,150);
      if(columns>1)sh.setColumnWidth(2,140);
      if(matrix.length>1){
        sh.getRange(2,1,matrix.length-1,columns).setVerticalAlignment('middle');
        sh.getRange(1,1,matrix.length,columns).createFilter();
      }
      sh.getDataRange().setFontFamily('Arial').setFontSize(10);
    });
    SpreadsheetApp.flush();
    const exportUrl='https://docs.google.com/spreadsheets/d/'+temp.getId()+'/export?format=xlsx';
    const exportResponse=UrlFetchApp.fetch(exportUrl,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});
    if(exportResponse.getResponseCode()!==200)throw Error('Excel eksport HTTP '+exportResponse.getResponseCode());
    const name='HISOB_'+({plan:'Plan',receipt:'Prixod',remain:'Qoldiq',total:'Jami'})[kind]+'_'+(all?'BARCHA':projects.join('_')).replace(/[^A-Za-z0-9_-]/g,'_').slice(0,95)+'.xlsx';
    const blob=exportResponse.getBlob().setName(name);
    const res=UrlFetchApp.fetch('https://api.telegram.org/bot'+botToken_()+'/sendDocument',{
      method:'post',payload:{chat_id:String(chat),document:blob,caption:'📥 '+reportKindNameV9_(kind)+' — '+(all?'Barcha '+projects.length+' ta loyiha':projects.join(', '))},muteHttpExceptions:true
    });
    const response=JSON.parse(res.getContentText());
    if(!response.ok)throw Error('Telegramga Excel yuborilmadi: '+response.description);
  }catch(err){console.error('Excel '+kind+': '+String(err&&err.message||err));msg_(chat,'❌ Excel tayyorlashda xato: '+String(err&&err.message||err).slice(0,250));}
  finally{DriveApp.getFileById(temp.getId()).setTrashed(true);}
}
/** Deploy qilingan Web App URLni WEB_APP_URL ga kiriting, so‘ng ushbu funksiyani Run qiling. */
function installTelegramWebhook_(){
  const url=cfg_('WEB_APP_URL'),secret=cfg_('WEBHOOK_SECRET');
  if(!url.includes('/exec')||secret.length<20)throw Error('WEB_APP_URL yoki WEBHOOK_SECRET noto‘g‘ri');
  const response=botCall_('setWebhook',{url:url+(url.includes('?')?'&':'?')+'webhook='+encodeURIComponent(secret),allowed_updates:['message','callback_query']});
  Logger.log('Webhook: '+response);
}


// Admin panel: Run menyusida ko'rinadigan ishga tushirish funksiyasi.
function initializeHISOB() {
  initializeHISOB_();
  Logger.log('HISOB bazasi tayyor');
}

// Telegram ID ni tokenni ko'rsatmasdan jurnalga chiqaradi.
function showMyTelegramId() {
  Logger.log("Webhook faol bolsa, getUpdates xato beradi.");
  Logger.log('Ruxsat berilgan IDlar Script Properties dagi ALLOWED_TELEGRAM_IDS da saqlangan.');
  Logger.log("Webhookni ochirmang.");
}


// -------- Apps Script Run menyusida ko'rinadigan funksiyalar --------
function setupAdminPassword() {
  setInitialAdminPassword_();
  Logger.log('Admin paroli faollashtirildi');
}
function installTelegramWebhook() {
  installTelegramWebhook_();
  Logger.log('UMBRELLA webhook sozlandi');
}
function checkHISOBSettings() {
  const p = PropertiesService.getScriptProperties();
  [
    'ADMIN_PASSWORD_HASH', 'BOT_TOKEN', 'SPREADSHEET_ID',
    'ALLOWED_TELEGRAM_IDS', 'WEBHOOK_SECRET', 'WEB_APP_URL'
  ].forEach(key => Logger.log(key + ': ' + (p.getProperty(key) ? 'BOR' : 'YOQ')));
}
function tekshirIndexKodni() {
  const html = HtmlService.createHtmlOutputFromFile('Index').getContent();
  const scripts = [...html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script\s*>/gi)];
  let errors = 0;
  scripts.forEach((part, i) => {
    if (/application\/json/i.test(part[1])) return;
    try { new Function(part[2]); }
    catch (e) { errors++; Logger.log('Skript ' + i + ': ' + e.message); }
  });
  Logger.log('Tekshirildi: ' + scripts.length + ' ta skript. Xatolar: ' + errors);
}

/**
 * UMBRELLA — webhook va Telegram javobini xavfsiz tekshirish.
 * Mavjud Code.gs OXIRIGA QO'SHING. Avvalgi kodlarni o'chirmang.
 * Token, admin paroli, mahsulot va kirimlarni o'zgartirmaydi.
 * Faqat kerak bo'lsa webhook manzilini to'g'rilaydi va admin IDga 1 test xabar yuboradi.
 */
function umbrellaUlanishniTuzatish() {
  const p = PropertiesService.getScriptProperties();
  const token = String(p.getProperty('BOT_TOKEN') || '').trim();
  const webUrl = String(p.getProperty('WEB_APP_URL') || '').trim();
  const secret = String(p.getProperty('WEBHOOK_SECRET') || '').trim();
  const admins = String(p.getProperty('ALLOWED_TELEGRAM_IDS') || '')
    .split(',').map(s => s.trim()).filter(Boolean);

  if (!token || !webUrl.endsWith('/exec') || secret.length < 20 || !admins.length) {
    throw new Error('BOT_TOKEN, WEB_APP_URL, WEBHOOK_SECRET yoki ALLOWED_TELEGRAM_IDS sozlamasini tekshiring');
  }
  function call(method, body) {
    const options = body === undefined
      ? {muteHttpExceptions: true}
      : {method:'post',contentType:'application/json',payload:JSON.stringify(body),muteHttpExceptions:true};
    const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, options);
    const data = JSON.parse(response.getContentText());
    if (!data.ok) throw new Error(method + ': ' + String(data.description || response.getResponseCode()));
    return data.result;
  }

  const me = call('getMe');
  Logger.log('BOT: @' + me.username + ' (token Telegramda amal qiladi)');
  const desired = webUrl + (webUrl.includes('?') ? '&' : '?') + 'webhook=' + encodeURIComponent(secret);
  let hook = call('getWebhookInfo');
  const wasCorrect = String(hook.url || '') === desired;
  Logger.log('WEBHOOK MANZILI: ' + (wasCorrect ? 'MOS' : 'MOS EMAS'));
  Logger.log('NAVBATDAGI XABARLAR: ' + Number(hook.pending_update_count || 0));
  if (hook.last_error_message) Logger.log('TELEGRAM OXIRGI XATO: ' + String(hook.last_error_message).slice(0,300));

  if (!wasCorrect) {
    call('setWebhook', {
      url: desired,
      allowed_updates: ['message','callback_query'],
      drop_pending_updates: false
    });
    hook = call('getWebhookInfo');
    if (hook.url !== desired) throw new Error('Webhook qayta sozlandi, ammo manzil mos kelmadi');
    Logger.log('WEBHOOK: YANGILANDI');
  } else {
    Logger.log('WEBHOOK: OZGARTIRILMADI');
  }

  try {
    call('sendMessage', {
      chat_id: admins[0],
      text: '✅ UMBRELLA diagnostika: bot xabar yubora oladi. Endi /start yuboring.'
    });
    Logger.log('TEST XABAR: YUBORILDI');
  } catch (err) {
    Logger.log('TEST XABAR: XATO — ' + err.message);
  }
  Logger.log('DIAGNOSTIKA YAKUNLANDI. Web App doPost javobini /start bilan tekshiring.');
}


// V7 ishga tushirishdan avval ruxsatni tekshirish. Bazani o'zgartirmaydi.
function umbrellaV7RuxsatTekshirish() {
  const spreadsheetId=cfg_('SPREADSHEET_ID');
  if(!spreadsheetId)throw Error('SPREADSHEET_ID sozlanmagan');
  const name=SpreadsheetApp.openById(spreadsheetId).getName();
  DriveApp.getRootFolder().getId();
  if(!cfg_('BOT_TOKEN')||!cfg_('ALLOWED_TELEGRAM_IDS'))throw Error('BOT_TOKEN yoki ALLOWED_TELEGRAM_IDS topilmadi');
  Logger.log('UMBRELLA V7 uchun Google Drive va HISOB bazasi ruxsatlari mavjud: '+name);
  Logger.log('Bu tekshiruv ma’lumotlarni o‘zgartirmadi. Keyin mavjud Web Appni YANGI VERSIYA orqali yangilang.');
}


/**
 * UMBRELLA V8 — jonli hisob ma’lumotlariga asoslangan hisobot PNG.
 * Texnik izoh: bu brauzerdan olingan aynan bir xil screenshot emas; Google Slides
 * orqali admin dizayniga mos, joriy HISOB_BAZA hisoblaridan yaratilgan tasvirdir.
 * Yangi Script Properties / parol / webhook kiritish talab qilmaydi.
 * Ishni boshlash: editor ichida setupUmbrellaDashboardV8 funksiyasini BIR MARTA Run qiling.
 * Dashboard kliklar 1 daqiqalik trigger orqali rasmlarni tayyorlaydi.
 */
const UMBV8 = {
  SHEET:'Dashboard_navbat',
  HEAD:['id','sana','telegram_chat_id','bolim','holat','urinish','xato'],
  COLOR:{dark:'#073E32',green:'#087A5D',mint:'#E7F5EE',orange:'#EF9756',red:'#B73C3C',ink:'#112A25',grey:'#647C74',bg:'#F3F7F5',white:'#FFFFFF',border:'#DDE8E3'},
  pageW:960,pageH:540
};
function dashboardLabelV8_(s){return ({main:'Dashboard',projects:'Loyihalar',daily:'Kunlik',report:'Hisobot Dashboard'})[s]||s;}
function queueSheetV8_(){
  const ss=sheetBook_();let sh=ss.getSheetByName(UMBV8.SHEET);
  if(!sh){sh=ss.insertSheet(UMBV8.SHEET);sh.getRange(1,1,1,UMBV8.HEAD.length).setValues([UMBV8.HEAD]).setFontWeight('bold');sh.setFrozenRows(1);}
  return sh;
}
function queueDashboardV8_(chat,section,updateId){
  if(!['main','projects','daily','report'].includes(section))throw Error('Dashboard bo‘limi noto‘g‘ri');
  const lock=LockService.getScriptLock();lock.waitLock(12000);
  try{
    const sh=queueSheetV8_(), id='tg_'+String(updateId||Utilities.getUuid()).replace(/[^A-Za-z0-9_-]/g,'').slice(0,80);
    // Telegram takroriy callback yuborganda bir xil tasvir ikkita yaratilmasin.
    const last=sh.getLastRow();
    if(last>1){const rows=sh.getRange(Math.max(2,last-79),1,Math.min(80,last-1),1).getValues();if(rows.some(x=>String(x[0])===id))return;}
    sh.appendRow([id,new Date(),String(chat),section,'KUTILMOQDA',0,'']);
  }finally{lock.releaseLock();}
}
function setupUmbrellaDashboardV8(){
  const old=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='processDashboardJobsV8');
  if(!old.length)ScriptApp.newTrigger('processDashboardJobsV8').timeBased().everyMinutes(1).create();
  // Kerakli Slides/Drive ruxsatlari ham so‘ralishi uchun xizmatlarni chaqiramiz.
  const pres=SlidesApp.create('UMBRELLA_V8_RUXSAT_SINOVI');
  try{
    const id=pres.getId(),slideId=pres.getSlides()[0].getObjectId();
    pres.saveAndClose();
    const png=getSlidePngV8_(id,slideId);
    if(png.getBytes().length<100)throw Error('PNG bo‘sh chiqdi');
    Logger.log('Google Slides PNG test: OK ('+png.getBytes().length+' bytes)');
  }finally{DriveApp.getFileById(pres.getId()).setTrashed(true);}
  const sh=queueSheetV8_();
  Logger.log('UMBRELLA V8 tayyor. Trigger: '+(old.length?'mavjud':'yaratildi')+'. Navbat: '+(sh.getLastRow()-1)+' ta.');
}
function processDashboardJobsV8(){
  const lock=LockService.getScriptLock();if(!lock.tryLock(2000))return;
  let task;
  try{
    const sh=queueSheetV8_();const last=sh.getLastRow();
    if(last<2)return;
    const rows=sh.getRange(2,1,last-1,7).getValues();
    for(let i=0;i<rows.length;i++){
      const r=rows[i],status=String(r[4]),attempt=Number(r[5]||0);
      const age= r[1] instanceof Date ? (Date.now()-r[1].getTime()) : 0;
      if(status==='KUTILMOQDA'||(status==='XATO'&&attempt<2)||(status==='ISHDA'&&age>15*60000&&attempt<2)){
        const n=i+2;sh.getRange(n,5,1,2).setValues([['ISHDA',attempt+1]]);
        task={row:n,chat:String(r[2]),section:String(r[3]),attempt:attempt+1};break;
      }
    }
  }finally{lock.releaseLock();}
  if(!task)return;
  try{
    sendDashboardPicturesV8_(task.chat,task.section);
    queueSheetV8_().getRange(task.row,5,1,3).setValues([['TAYYOR',task.attempt,'']]);
  }catch(err){
    const msg=String(err&&err.message||err).slice(0,450);
    queueSheetV8_().getRange(task.row,5,1,3).setValues([['XATO',task.attempt,msg]]);
    Logger.log('UMBRELLA V8 TASVIR XATO: '+msg);
    if(task.attempt>=2){try{msg_(task.chat,'❌ '+dashboardLabelV8_(task.section)+' tasvirini yaratishda xato: '+msg.slice(0,180));}catch(e){Logger.log('Telegram javob xatosi: '+e.message);}}
  }
}
function sendDashboardPicturesV8_(chat,section){
  const summary=summaryData_(),days=section==='daily'?dailyLast30V8_(summary.daily):null;
  const deck=SlidesApp.create('HISOB_V8_'+section+'_'+Date.now());
  try{
    const original=deck.getSlides()[0];
    const pages=[];
    if(section==='main'){pages.push(makeDashboardSlideV8_(deck,summary));}
    else if(section==='projects'){
      const projects=summary.projects;
      for(let start=0;start<projects.length;start+=6)pages.push(makeProjectsSlideV8_(deck,projects.slice(start,start+6),start,projects.length));
      if(!projects.length)pages.push(makeProjectsSlideV8_(deck,[],0,0));
    }
    else if(section==='daily'){pages.push(makeDailySlideV8_(deck,days,'dona'));pages.push(makeDailySlideV8_(deck,days,'kg'));}
    else if(section==='report'){for(let i=0;i<Math.max(1,summary.projects.length);i+=11)pages.push(makeReportSlideV8_(deck,summary,i));}
    else throw Error('Noma’lum bo‘lim '+section);
    const pageIds=pages.map(p=>p.getObjectId());
    original.remove();
    deck.saveAndClose();
    for(let i=0;i<pageIds.length;i++){
      const blob=getSlidePngV8_(deck.getId(),pageIds[i]);
      blob.setName('HISOB_'+section+'_'+(i+1)+'.png');
      const cap='📊 HISOB · '+dashboardLabelV8_(section)+(pages.length>1?' ('+(i+1)+'/'+pages.length+')':'')+'\n📅 '+now_()+'\nManba: HISOB_BAZA — jonli hisoblar';
      const res=UrlFetchApp.fetch('https://api.telegram.org/bot'+botToken_()+'/sendPhoto',{
        method:'post',payload:{chat_id:String(chat),photo:blob,caption:cap},muteHttpExceptions:true
      });
      const data=JSON.parse(res.getContentText());if(!data.ok)throw Error('Telegram sendPhoto: '+data.description);
    }
  }finally{DriveApp.getFileById(deck.getId()).setTrashed(true);}
}
function getSlidePngV8_(presentationId,objectId){
  // Docs: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations.pages/getThumbnail
  const url='https://slides.googleapis.com/v1/presentations/'+encodeURIComponent(presentationId)+'/pages/'+encodeURIComponent(objectId)+'/thumbnail?thumbnailProperties.mimeType=PNG&thumbnailProperties.thumbnailSize=LARGE';
  let last='';
  for(let tries=0;tries<3;tries++){
    const r=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});
    if(r.getResponseCode()===200){
      const obj=JSON.parse(r.getContentText());
      const img=UrlFetchApp.fetch(obj.contentUrl,{muteHttpExceptions:true});
      if(img.getResponseCode()===200)return img.getBlob().setContentType('image/png');
      last='PNG HTTP '+img.getResponseCode();
    }else last='Slides thumbnail HTTP '+r.getResponseCode()+': '+r.getContentText().slice(0,180);
    Utilities.sleep(1100);
  }
  throw Error(last+' (Slides/Drive ruxsatlarini tekshiring)');
}
function vpV8_(deck){return {sx:deck.getPageWidth()/UMBV8.pageW,sy:deck.getPageHeight()/UMBV8.pageH};}
function slideV8_(deck,title,sub){
  const slide=deck.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(UMBV8.COLOR.bg);
  const c={slide,p:vpV8_(deck)};
  rectV8_(c,0,0,960,58,'#FFFFFF');rectV8_(c,0,0,7,58,UMBV8.COLOR.green);
  txV8_(c,'HISOB.',22,13,110,34,22,UMBV8.COLOR.green,true);
  txV8_(c,title,150,12,680,27,18,UMBV8.COLOR.ink,true);
  txV8_(c,sub+'  ·  '+now_(),150,36,760,14,8,UMBV8.COLOR.grey);
  return c;
}
function rectV8_(c,x,y,w,h,fill){
  const sh=c.slide.insertShape(SlidesApp.ShapeType.RECTANGLE,x*c.p.sx,y*c.p.sy,w*c.p.sx,h*c.p.sy);
  sh.getFill().setSolidFill(fill);sh.getBorder().setTransparent();return sh;
}
function txV8_(c,str,x,y,w,h,size,color,bold,align){
  const sh=c.slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,x*c.p.sx,y*c.p.sy,w*c.p.sx,h*c.p.sy);
  sh.getText().setText(String(str));sh.getText().getTextStyle().setFontFamily('Arial').setFontSize(Math.max(6,size*c.p.sx)).setForegroundColor(color||UMBV8.COLOR.ink).setBold(!!bold);
  sh.getText().getParagraphStyle().setParagraphAlignment(align==='center'?SlidesApp.ParagraphAlignment.CENTER:SlidesApp.ParagraphAlignment.START);
  return sh;
}
function cardV8_(c,x,y,w,h,title,value,small,tone){
  rectV8_(c,x,y,w,h,tone||'#FFFFFF');
  txV8_(c,title,x+12,y+9,w-24,20,12,UMBV8.COLOR.grey,true);
  txV8_(c,value,x+12,y+32,w-24,37,24,UMBV8.COLOR.ink,true);
  if(small)txV8_(c,small,x+12,y+h-21,w-24,16,9,UMBV8.COLOR.grey);
}
function barV8_(c,x,y,w,pct,color){
  rectV8_(c,x,y,w,7,UMBV8.COLOR.border);
  if(pct>0)rectV8_(c,x,y,Math.max(2,w*Math.min(1,Math.max(0,pct))),7,color||UMBV8.COLOR.green);
}
function makeDashboardSlideV8_(deck,s){
  const c=slideV8_(deck,'Dashboard','Reja, kirim va bajarilish nazorati'),t=s.totals,prs=s.projects;
  const plan=t.totalPlan,got=t.totalGot,done=plan?got/plan:0;
  const cards=[['Umumiy reja',fmtNum_(plan)+' dona',fmtNum_(t.planKg)+' kg'],['Jami kirim',fmtNum_(got)+' dona',fmtNum_(t.gotKg)+' kg'],['Rejadan qolgan',fmtNum_(plan-got)+' dona',fmtNum_(t.planKg-t.gotKg)+' kg'],['Reja bajarilishi',(100*done).toFixed(1)+'%',(t.planKg?100*t.gotKg/t.planKg:0).toFixed(1)+'% · kg']];
  cards.forEach((k,i)=>cardV8_(c,18+i*235,72,222,91,k[0],k[1],k[2],i===3?UMBV8.COLOR.mint:'#FFFFFF'));
  const groups=[['Plan',prs.filter(p=>p.got>0&&p.got<p.plan)],['Hali boshlanmagan',prs.filter(p=>p.got<=0)],['To‘liq qabul qilindi',prs.filter(p=>p.got>=p.plan)]];
  groups.forEach((g,i)=>{
    const x=18+i*316;rectV8_(c,x,177,302,215,'#FFFFFF');
    txV8_(c,g[0]+' · '+g[1].length,x+10,184,282,25,15,UMBV8.COLOR.ink,true);
    g[1].slice(0,6).forEach((p,j)=>{const y=217+j*29;txV8_(c,p.project,x+9,y,83,18,11,UMBV8.COLOR.ink,true);barV8_(c,x+94,y+5,151,p.plan?p.got/p.plan:0);txV8_(c,(p.plan?100*p.got/p.plan:0).toFixed(1)+'%',x+249,y,46,19,10,UMBV8.COLOR.green,true);});
  });
  rectV8_(c,18,403,924,121,'#FFFFFF');txV8_(c,'Loyihalar bo‘yicha reja va fakt',30,408,700,23,13,UMBV8.COLOR.ink,true);
  const subset=prs.slice(0,10),half=Math.ceil(subset.length/2);
  subset.forEach((p,i)=>{const col=Math.floor(i/half),row=i%half,x=30+col*457,y=437+row*17;
    txV8_(c,p.project,x,y,75,15,9,UMBV8.COLOR.green,true);txV8_(c,fmtNum_(p.got)+' / '+fmtNum_(p.plan)+' dona',x+79,y,138,15,9,UMBV8.COLOR.ink);barV8_(c,x+221,y+5,150,p.plan?p.got/p.plan:0);txV8_(c,(p.plan?100*p.got/p.plan:0).toFixed(1)+'%',x+376,y,50,15,9,UMBV8.COLOR.ink,true);
  });
  return c.slide;
}
function makeProjectsSlideV8_(deck,projects,start,total){
  const c=slideV8_(deck,'Loyihalar','Reja · kirim · qoldiq · bajarilish  |  jami '+total+' ta');
  projects.forEach((p,i)=>{
    const x=18+(i%2)*468,y=75+Math.floor(i/2)*150,w=450,h=139;
    rectV8_(c,x,y,w,h,'#FFFFFF');
    txV8_(c,p.project,x+12,y+8,200,24,16,UMBV8.COLOR.green,true);
    const d=[['Reja',p.plan,p.planKg],['Kirim',p.got,p.gotKg],['Qoldiq',p.plan-p.got,p.planKg-p.gotKg]];
    d.forEach((a,j)=>{const xx=x+10+j*144;rectV8_(c,xx,y+37,136,57,j===2?'#F8F4F3':'#F2F8F5');
      txV8_(c,a[0],xx+5,y+40,126,13,9,UMBV8.COLOR.grey);txV8_(c,fmtNum_(a[1])+' dona',xx+5,y+52,126,20,13,j===2?UMBV8.COLOR.red:UMBV8.COLOR.ink,true);
      txV8_(c,fmtNum_(a[2])+' kg',xx+5,y+75,126,14,9,UMBV8.COLOR.grey);
    });
    txV8_(c,'Reja bajarilishi',x+11,y+99,175,14,9,UMBV8.COLOR.grey);
    txV8_(c,(p.plan?100*p.got/p.plan:0).toFixed(1)+'%',x+390,y+98,53,17,10,UMBV8.COLOR.green,true);
    barV8_(c,x+12,y+121,426,p.plan?p.got/p.plan:0);
  });
  txV8_(c,'HISOB_BAZA: mahsulot pozitsiyalari barcha loyihalar bo‘yicha hisoblangan',24,524,910,13,9,UMBV8.COLOR.grey);
  return c.slide;
}
function dailyLast30V8_(daily){
  const arr=[],nowDay=Utilities.formatDate(new Date(),hisobTZ_(),'yyyy-MM-dd');
  const today=new Date(nowDay+'T12:00:00Z');
  for(let i=29;i>=0;i--){const key=Utilities.formatDate(new Date(today.getTime()-i*86400000),'UTC','yyyy-MM-dd');
    const x=daily.get(key)||{positive:0,returned:0,grossKg:0,returnedKg:0};arr.push(Object.assign({date:key},x));
  }
  return arr;
}
function makeDailySlideV8_(deck,days,mode){
  const kg=mode==='kg';
  const c=slideV8_(deck,'Kunlik',kg?'So‘nggi 30 kunlik kirim — kilogramm':'So‘nggi 30 kunlik prixod — dona'),gross=days.reduce((s,d)=>s+(kg?d.grossKg:d.positive),0),back=days.reduce((s,d)=>s+(kg?d.returnedKg:d.returned),0);
  const unit=kg?' kg':' dona';
  cardV8_(c,18,74,222,85,'30 kunda qabul',fmtNum_(gross)+unit,'Sana bo‘yicha');
  cardV8_(c,252,74,222,85,'30 kunda qaytarildi',fmtNum_(back)+unit,'Sana bo‘yicha');
  cardV8_(c,486,74,222,85,'Sof kirim',fmtNum_(gross-back)+unit,'Qabul minus qaytarish',UMBV8.COLOR.mint);
  cardV8_(c,720,74,222,85,'Kirim yozuvlari',String(days.filter(d=>d.positive||d.returned).length)+' kun','Faol kunlar');
  rectV8_(c,18,169,924,354,'#FFFFFF');
  txV8_(c,'Qabul',738,181,54,19,10,UMBV8.COLOR.green,true);rectV8_(c,718,185,12,8,UMBV8.COLOR.green);
  txV8_(c,'Qaytarish',834,181,83,19,10,UMBV8.COLOR.orange,true);rectV8_(c,814,185,12,8,UMBV8.COLOR.orange);
  const maxv=Math.max(1,...days.map(d=>Math.max(kg?d.grossKg:d.positive,kg?d.returnedKg:d.returned)));
  for(let j=0;j<=4;j++){const y=470-j*62;rectV8_(c,55,y,842,1,'#E7EEE9');txV8_(c,String(Math.round(maxv*j/4)),28,y-10,29,16,9,UMBV8.COLOR.grey);}
  const bw=13;
  days.forEach((d,i)=>{
    const x=64+i*27.5;
    const pos=kg?d.grossKg:d.positive,neg=kg?d.returnedKg:d.returned;
    const hg=Math.max(pos?2:0,pos/maxv*242),ho=Math.max(neg?2:0,neg/maxv*242);
    if(pos)rectV8_(c,x,469-hg,bw,hg,UMBV8.COLOR.green);
    if(neg)rectV8_(c,x+bw,469-ho,bw,ho,UMBV8.COLOR.orange);
    if(i%3===0||i===29)txV8_(c,d.date.slice(5),x-7,476,45,18,8,UMBV8.COLOR.grey);
  });
  txV8_(c,'Yashil — kelgan mahsulot    /    To‘q sariq — qaytarilgan mahsulot    /    Sof kirim: '+fmtNum_(gross-back)+unit,36,506,860,15,9,UMBV8.COLOR.grey);
  return c.slide;
}
function makeReportSlideV8_(deck,s,offset){
  const c=slideV8_(deck,'Hisobot Dashboard','Loyihalar bo‘yicha umumiy natija'),t=s.totals;
  const vals=[['REJA',fmtNum_(t.totalPlan)+' dona',fmtNum_(t.planKg)+' kg'],['JAMI KIRIM',fmtNum_(t.totalGot)+' dona',fmtNum_(t.gotKg)+' kg'],['QOLDIQ',fmtNum_(t.totalPlan-t.totalGot)+' dona',fmtNum_(t.planKg-t.gotKg)+' kg'],['BAJARILISH',(t.totalPlan?100*t.totalGot/t.totalPlan:0).toFixed(1)+'% · dona',(t.planKg?100*t.gotKg/t.planKg:0).toFixed(1)+'% · kg']];
  vals.forEach((k,i)=>cardV8_(c,17+i*235,70,222,92,k[0],k[1],k[2],i===3?UMBV8.COLOR.mint:'#FFFFFF'));
  rectV8_(c,17,172,926,352,'#FFFFFF');txV8_(c,'1. Loyihalar bo‘yicha umumiy natija',28,178,820,22,14,UMBV8.COLOR.green,true);
  const cols=[['LOYIHA',23,115],['POZITSIYA',142,72],['REJA DONA',223,105],['REJA KG',335,111],['KIRIM DONA',452,112],['KIRIM KG',571,112],['QOLDIQ',688,110],['BAJARILISH',807,112]];
  rectV8_(c,26,207,905,26,UMBV8.COLOR.bg);cols.forEach(x=>txV8_(c,x[0],x[1],213,x[2],15,8,UMBV8.COLOR.grey,true));
  s.projects.slice(offset,offset+11).forEach((p,i)=>{
    const y=236+i*23;rectV8_(c,25,y,907,22,i%2?'#FAFCFB':'#FFFFFF');
    const v=[p.project,String(p.positions),fmtNum_(p.plan),fmtNum_(p.planKg),fmtNum_(p.got),fmtNum_(p.gotKg),fmtNum_(p.plan-p.got),(p.plan?100*p.got/p.plan:0).toFixed(1)+'%'];
    v.forEach((value,j)=>txV8_(c,value,cols[j][1],y+3,cols[j][2],15,9,j===0?UMBV8.COLOR.green:j===6?UMBV8.COLOR.red:UMBV8.COLOR.ink,j===0||j===4));
  });
  const y=236+Math.min(11,Math.max(0,s.projects.length-offset))*23;
  rectV8_(c,25,y,907,25,UMBV8.COLOR.mint);
  const vv=['JAMI',String(s.projects.reduce((n,p)=>n+p.positions,0)),fmtNum_(t.totalPlan),fmtNum_(t.planKg),fmtNum_(t.totalGot),fmtNum_(t.gotKg),fmtNum_(t.totalPlan-t.totalGot),(t.totalPlan?100*t.totalGot/t.totalPlan:0).toFixed(1)+'%'];
  vv.forEach((v,j)=>txV8_(c,v,cols[j][1],y+5,cols[j][2],15,9,UMBV8.COLOR.green,true));
  return c.slide;
}


/** UMBRELLA MINI APP V10 — serverda Telegram initData imzosini tekshirish. */
function miniAuthV10_(initData){
  if(typeof initData!=='string'||!initData||initData.length>8192)throw Error('Telegram identifikatsiyasi topilmadi. Ilovani bot tugmasidan oching.');
  const kv={};
  initData.split('&').forEach(function(part){
    const pos=part.indexOf('=');
    if(pos<0)return;
    const k=decodeURIComponent(part.slice(0,pos).replace(/\+/g,'%20'));
    const v=decodeURIComponent(part.slice(pos+1).replace(/\+/g,'%20'));
    if(Object.prototype.hasOwnProperty.call(kv,k))throw Error('Takroriy Telegram parametri');
    kv[k]=v;
  });
  const provided=(kv.hash||'').toLowerCase();
  if(!/^[0-9a-f]{64}$/.test(provided))throw Error('Telegram imzosi yo‘q');
  const authDate=Number(kv.auth_date);
  const now=Math.floor(Date.now()/1000);
  if(!Number.isSafeInteger(authDate)||authDate>now+120||now-authDate>86400)throw Error('Telegram sessiyasi eskirgan. Ilovani qayta oching.');
  const signed=Object.keys(kv).filter(k=>k!=='hash').sort().map(k=>k+'='+kv[k]).join('\n');
  const b=function(s){return Utilities.newBlob(s).getBytes();};
  const secret=Utilities.computeHmacSha256Signature(b(botToken_()),b('WebAppData'));
  const sig=Utilities.computeHmacSha256Signature(b(signed),secret)
    .map(x=>((x+256)%256).toString(16).padStart(2,'0')).join('');
  let mismatch=0;
  for(let i=0;i<64;i++)mismatch|=provided.charCodeAt(i)^sig.charCodeAt(i);
  if(mismatch)throw Error('Telegram imzosi tasdiqlanmadi');
  let user;
  try{user=JSON.parse(kv.user||'{}');}catch(e){throw Error('Telegram foydalanuvchi ma’lumoti noto‘g‘ri');}
  if(!Number.isSafeInteger(Number(user.id))||!userAllowed_(user.id))throw Error('Bu Telegram ID uchun ruxsat berilmagan');
  return {id:Number(user.id),firstName:String(user.first_name||''),username:String(user.username||'')};
}

/** Hozirgi HISOB_BAZA dan o‘qiladi; hisob-kitobni o‘zgartirmaydi. */
function miniLoadV10(initData) {
  const person = miniAuthV10_(initData); // Serverda Telegram imzosi, auth_date va ruxsat tekshiriladi
  const result = miniDataV10_(person);
  const u = umbrellaUserV11_(person.id);
  if (u && u.role === 'operator' && u.status === 'active' &&
      u.passhash && u.sessionUntil > Date.now() &&
      userAllowed_(person.id)) {
    // Har sahifa ochilganda yangilash o‘rniga, qolgani 20 kundan kam bo‘lsa uzaytiramiz
    const TWENTY_DAYS = 20 * 24 * 3600 * 1000;
    const THIRTY_DAYS = 30 * 24 * 3600 * 1000;
    if (u.sessionUntil - Date.now() < TWENTY_DAYS) {
      const lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        const current = umbrellaUserV11_(person.id);
        if (current && current.role === 'operator' && current.status === 'active' &&
            current.passhash && current.sessionUntil > Date.now() && userAllowed_(person.id)) {
          umbrellaUsersSheetV11_().getRange(current.row, 10).setValue(Date.now() + THIRTY_DAYS);
        }
      } finally {
        lock.releaseLock();
      }
    }
  }
  return result;
}

function miniDataV10_(user){
  const db=readDb_(),receipts=new Map(),products=new Map(),days=new Map();
  db.products.forEach(p=>products.set(String(p.id),p));
  db.receipts.forEach(r=>{
    receipts.set(String(r.productId),(receipts.get(String(r.productId))||0)+Number(r.qty));
    const key=asDate_(r.date),d=days.get(key)||{date:key,positive:0,returned:0,positiveKg:0,returnedKg:0};
    const p=products.get(String(r.productId)),w=Number(p&&p.kg||0);
    if(Number(r.qty)>=0){d.positive+=Number(r.qty);d.positiveKg+=Number(r.qty)*w;}
    else{d.returned+=-Number(r.qty);d.returnedKg+=-Number(r.qty)*w;}
    days.set(key,d);
  });
  const projectMap=new Map();let plan=0,got=0,planKg=0,gotKg=0;
  const list=db.products.map(p=>{
    const planned=Number(p.plan||0),count=Number(p.opening||0)+(receipts.get(String(p.id))||0),w=Number(p.kg||0);
    const name=String(p.project||'');
    const s=projectMap.get(name)||{project:name,plan:0,received:0,planKg:0,receivedKg:0,positions:0};
    s.plan+=planned;s.received+=count;s.planKg+=planned*w;s.receivedKg+=count*w;s.positions++;
    projectMap.set(name,s);plan+=planned;got+=count;planKg+=planned*w;gotKg+=count*w;
    return {project:name,code:String(p.code||''),profile:String(p.profile||''),kg:w,plan:planned,received:count,remaining:planned-count};
  });
  const projects=[...projectMap.values()].sort((a,b)=>a.project.localeCompare(b.project));
  const recent=db.receipts.slice(-25).reverse().map(r=>{
    const p=products.get(String(r.productId));
    return {project:p?String(p.project):'',code:p?String(p.code):'',qty:Number(r.qty),date:asDate_(r.date),kg:Number(r.qty)*Number(p&&p.kg||0)};
  });
  const today=Utilities.formatDate(new Date(),hisobTZ_(),'yyyy-MM-dd');
  const d0=new Date(today+'T12:00:00Z'),daily=[];
  for(let k=29;k>=0;k--){
    const key=Utilities.formatDate(new Date(d0.getTime()-k*86400000),'UTC','yyyy-MM-dd');
    daily.push(days.get(key)||{date:key,positive:0,returned:0,positiveKg:0,returnedKg:0});
  }
  return {user:umbrellaUserPresentationV11_(user),totals:{plan,got,remaining:plan-got,planKg,gotKg,remainingKg:planKg-gotKg,percent:plan?100*got/plan:0},projects,products:list,daily,recent,revision:revision_()};
}

function umbrellaUserPresentationV11_(user){
  if(Number(user.id)===0)return {...user,name:'Administrator',role:'administrator',canWrite:true,needsPassword:false};
  const u=umbrellaUserV11_(user.id);
  if(!u)return {...user,role:'viewer',canWrite:false,needsPassword:false};
  return {...user,firstName:u.name,name:u.name,role:u.role,canWrite:umbrellaCanWriteV11_(user.id),needsPassword:u.role==='operator'&&!u.passhash};
}
/** Matnli ko‘p qatorli prixod, tasdiqli ortiqcha kirim, va qaytarish. */
function miniReceiptV10(initData,text,kind,confirmed){
  const user=miniAuthV10_(initData);
  umbrellaRequireWriteV11_(user.id);
  if(kind!=='add'&&kind!=='undo')throw Error('Noma’lum amal');
  const parsed=parseBatch_(text);
  if(parsed.error)return {ok:false,error:parsed.error};
  const result=applyBatch_(user.id,parsed.commands,kind==='add'?1:-1,confirmed===true);
  if(result.needConfirm)return {ok:false,confirm:true,message:result.excessText};
  if(result.error)return {ok:false,error:result.error};
  return {ok:true,message:formatBatch_(result),revision:result.revision};
}

/** HISOB Excel hisobotini foydalanuvchining Telegram botiga jo‘natadi. */
function miniReportV10(initData,kind,project){
  const user=miniAuthV10_(initData);
  if(!['plan','receipt','remain','total'].includes(kind))throw Error('Hisobot turi noma’lum');
  const projects=(project==='ALL'||!project)?null:[String(project)];
  sendReport_(user.id,kind,projects);
  return 'Excel hisobot yuborish so‘rovi bajarildi. Telegramdagi UMBRELLA chatini tekshiring.';
}

/** Mini App foydalanuvchini Telegram chatida Excel qabul qilish rejimiga o‘tkazadi. */
function miniPlanPromptV10(initData){
  const user=miniAuthV10_(initData);
  umbrellaRequireWriteV11_(user.id);
  setMode_(user.id,'plan_excel');
  msg_(user.id,'📥 Plan qo‘shish rejimi yoqildi. Shu chatga .xlsx Excel reja faylini hujjat qilib yuboring.');
  return true;
}


/** Apps Script ichida Telegram initData bo‘sh qaytsa, mavjud ADMIN parol orqali xavfsiz kirish. */
function miniPasswordLoginV10(password){return apiLogin(password);}
function miniAdminLoadV10(token){
  requireSession_(token);
  return miniDataV10_({id:0,firstName:'Administrator',username:''});
}
function miniAdminReceiptV10(token,text,kind,confirmed){
  requireSession_(token);
  if(kind!=='add'&&kind!=='undo')throw Error('Noma’lum amal');
  const parsed=parseBatch_(text);
  if(parsed.error)return {ok:false,error:parsed.error};
  const result=applyBatch_('MINI_ADMIN',parsed.commands,kind==='add'?1:-1,confirmed===true);
  if(result.needConfirm)return {ok:false,confirm:true,message:result.excessText};
  if(result.error)return {ok:false,error:result.error};
  return {ok:true,message:formatBatch_(result),revision:result.revision};
}
function miniAdminReportV10(token,kind,project){
  requireSession_(token);
  if(!['plan','receipt','remain','total'].includes(kind))throw Error('Hisobot turi noma’lum');
  const adminId=cfg_('ALLOWED_TELEGRAM_IDS').split(',').map(x=>x.trim()).filter(Boolean)[0];
  if(!adminId)throw Error('Hisobot yuborish uchun ALLOWED_TELEGRAM_IDS yo‘q');
  sendReport_(adminId,kind,project==='ALL'||!project?null:[String(project)]);
  return 'Excel hisobotni UMBRELLA bot chatingizdan tekshiring.';
}
function miniAdminPlanPromptV10(token){
  requireSession_(token);
  const adminId=cfg_('ALLOWED_TELEGRAM_IDS').split(',').map(x=>x.trim()).filter(Boolean)[0];
  if(!adminId)throw Error('Telegram ID sozlanmagan');
  setMode_(adminId,'plan_excel');
  msg_(adminId,'📥 Plan qo‘shish rejimi. Endi .xlsx faylni yuboring.');
  return true;
}

/** Bir martalik ixtiyoriy sozlash — Telegram chatning yuqori menyu tugmasini Mini Appga ulaydi. */
function setupUmbrellaMiniAppV10(){
  const url=miniAppUrlV10_();
  const ids=cfg_('ALLOWED_TELEGRAM_IDS').split(',').map(x=>x.trim()).filter(Boolean);
  if(!ids.length)throw Error('ALLOWED_TELEGRAM_IDS sozlanmagan');
  ids.forEach(function(id){
    botCall_('setChatMenuButton',{chat_id:Number(id),menu_button:{type:'web_app',text:'📱 HISOB',web_app:{url:url}}});
  });
  Logger.log('UMBRELLA Mini App tugmasi '+ids.length+' ta foydalanuvchi uchun tayyor.');
}


/** UMBRELLA V13: HISOB administrator panelidan xavfsiz kirish huquqlarini boshqarish.
 * Yopiq rejimdan tashqarida ishlamaydi. Faqat ADMIN_PASSWORD_HASH orqali olingan
 * faol admin sessiyasi bilan chaqiriladi. Parol / salt / sessiya qaytarilmaydi.
 */
function apiUsersV13(token){
  requireSession_(token);
  if(!umbrellaClosedReadyV12_())throw Error('Avval setupUmbrellaClosedV12 ni bir marta ishga tushiring.');
  const ownerIds=umbrellaOwnersV11_(),sh=umbrellaUsersSheetV11_(),last=sh.getLastRow();
  const values=last>1?sh.getRange(2,1,last-1,7).getValues():[];
  const users=values.filter(r=>/^\d{5,20}$/.test(String(r[0]||''))).map(r=>{
    const id=String(r[0]),owner=ownerIds.includes(id),status=String(r[4]||'pending');
    return {id,name:String(r[1]||''),phone:String(r[2]||''),role:owner?'owner':String(r[3]||'viewer'),status,
      joined:String(r[5]||''),approved:String(r[6]||''),owner};
  });
  const hist=rows_(HISOB.AUDIT,HISOB.AH)
    .filter(r=>/^(RUXSAT_|YOPIQ_TIZIM_)/.test(String(r[1]||'')))
    .slice(-30).reverse()
    .map(r=>({time:String(r[0]||''),action:String(r[1]||''),actor:String(r[2]||''),note:String(r[6]||'')}));
  return {users,history:hist};
}
function apiSetUserAccessV13(token,targetId,newRole,newStatus){
  requireSession_(token);
  if(!umbrellaClosedReadyV12_())throw Error('Yopiq rejim yoqilmagan.');
  const id=String(targetId||'').trim(),role=String(newRole||''),status=String(newStatus||'');
  if(!/^\d{5,20}$/.test(id))throw Error('Telegram ID noto‘g‘ri.');
  if(!['viewer','operator'].includes(role))throw Error('Faqat «Operator» yoki «Faqat hisobot» ruxsati mumkin.');
  if(!['active','pending','blocked'].includes(status))throw Error('Holat noto‘g‘ri.');
  if(umbrellaIsOwnerV11_(id))throw Error('Asosiy administratorni paneldan bloklash yoki o‘zgartirish mumkin emas.');
  let previous,person;
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    person=umbrellaUserV11_(id);
    if(!person)throw Error('Foydalanuvchi topilmadi. Avval botda ro‘yxatdan o‘tsin.');
    previous={role:person.role,status:person.status};
    if(previous.role===role&&previous.status===status)return {changed:false,id,role,status};
    const sh=umbrellaUsersSheetV11_();
    // status=active bo‘lmaguncha serverdagi userAllowed_ ma’lumot ochmaydi.
    sh.getRange(person.row,4,1,2).setValues([[role,status]]);
    sh.getRange(person.row,7).setValue(status==='active'?'V12|'+now_():'');
    // Har qanday rol/holat almashinuvi avvalgi operator sessiyasini bekor qiladi.
    sh.getRange(person.row,10).setValue(0);
    audit_(status==='blocked'?'RUXSAT_BEKOR':status==='pending'?'RUXSAT_KUTILMOQDA':'RUXSAT_OZGARDI',
      'HISOB_ADMIN','','',0,person.name+' [TG:'+id+'] '+previous.role+'/'+previous.status+' → '+role+'/'+status);
  }finally{lock.releaseLock();}
  // Telegramda foydalanuvchiga xabar berish yordamchi amal, asosiy o‘zgarishni bekor qilmaydi.
  try{
    const message=status==='blocked'?'⛔ UMBRELLA tizimiga ruxsatingiz bekor qilindi.':
      status==='pending'?'⏳ UMBRELLA ruxsatingiz administrator tasdig‘ini kutmoqda.':
      role==='operator'?'✅ Sizga UMBRELLA Operator ruxsati berildi. Yozish uchun Mini Appda parol bilan kiring.':
      '✅ Sizga UMBRELLA «Faqat hisobot» ruxsati berildi.';
    msg_(Number(id),message);
  }catch(e){Logger.log('Ruxsat o‘zgardi, Telegram bildirishnomasi yetmadi: '+String(e.message||e));}
  return {changed:true,id,role,status};
}


/** HISOB V14: Planni to‘xtatish/davom ettirish. Faqat parolli admin sessiyasi. */
function apiSetProjectPausedV14(token, projectName, shouldPause){
  requireSession_(token);
  if(typeof shouldPause!=='boolean')throw Error('Plan holati noto‘g‘ri.');
  const project=String(projectName||'').trim();
  if(!project || project.length>100)throw Error('Loyiha nomi noto‘g‘ri.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const existing=rows_(HISOB.PRODUCT,HISOB.PH).some(r=>String(r[1]||'').trim()===project);
    if(!existing)throw Error('Loyiha topilmadi. Sahifani yangilang.');
    const current=projectStatesV14_()[project]==='paused';
    if(current===shouldPause)return {project,paused:current,revision:revision_(),changed:false};
    const sh=tab_(HISOB.PLAN_STATE,HISOB.PSH);
    const last=sh.getLastRow();
    const keys=last>1?sh.getRange(2,1,last-1,1).getValues():[];
    const idx=keys.findIndex(r=>String(r[0]||'').trim()===project);
    const values=[project,shouldPause?'paused':'active',now_(),'HISOB_ADMIN'];
    if(idx>=0)sh.getRange(idx+2,1,1,HISOB.PSH.length).setValues([values]);
    else sh.appendRow(values);
    const revision=bumpRevision_();
    audit_(shouldPause?'PLAN_TOXTATILDI':'PLAN_DAVOM_ETDI','HISOB_ADMIN',project,'',0,'Reja va kirim saqlandi; rev '+revision);
    SpreadsheetApp.flush();
    return {project,paused:shouldPause,revision,changed:true};
  }finally{lock.releaseLock();}
}


/** HISOB V15.6: parolni bilib almashtirish — Mini App Telegram operatori. */
function miniOperatorChangePasswordV156(initData,oldPassword,newPassword){
  const person=miniAuthV10_(initData);
  const old=String(oldPassword||''), next=String(newPassword||'');
  if(next.length<12||next.length>128||next===old)throw Error('Yangi parol 12–128 belgi bo‘lsin va eskisidan farq qilsin.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const u=umbrellaUserV11_(person.id);
    if(!u||!userAllowed_(person.id)||u.role!=='operator'||u.status!=='active'||!u.passhash)throw Error('Operator kirish huquqi yoki parol yo‘q.');
    const cache=CacheService.getScriptCache(),key='changefail:'+person.id;
    const fails=Number(cache.get(key)||0);
    if(fails>=5)throw Error('Urinishlar limiti. 15 daqiqadan keyin urinib ko‘ring.');
    if(umbrellaPasswordDigestV11_(u.salt,old)!==u.passhash){cache.put(key,String(fails+1),900);throw Error('Joriy parol noto‘g‘ri.');}
    cache.remove(key);
    const salt=Utilities.getUuid(),digest=umbrellaPasswordDigestV11_(salt,next);
    // Yangi parol ishlaydi, ushbu qurilmadagi sessiya qayta faollashadi.
    umbrellaUsersSheetV11_().getRange(u.row,8,1,3).setValues([[salt,digest,Date.now()+30*24*3600*1000]]);
    audit_('OPERATOR_PAROL_ALMASHTIRILDI',person.id,'','',0,'Operator paroli yangilandi');
    return {ok:true,message:'Operator paroli almashtirildi.'};
  }finally{lock.releaseLock();}
}
/** HISOB V15.6: Admin parolini joriy parol orqali almashtirish. */
function miniAdminChangePasswordV156(token,oldPassword,newPassword){
  requireSession_(token);
  const next=String(newPassword||''),old=String(oldPassword||'');
  if(next.length<12||next.length>128||next===old)throw Error('Yangi parol 12–128 belgi bo‘lsin va eskisidan farq qilsin.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    requireSession_(token);
    const stored=cfg_('ADMIN_PASSWORD_HASH'),i=stored.indexOf(':');
    if(i<=0||hash_(stored.slice(0,i)+':'+old)!==stored.slice(i+1))throw Error('Joriy admin paroli noto‘g‘ri.');
    const salt=Utilities.getUuid();
    props_().setProperty('ADMIN_PASSWORD_HASH',salt+':'+hash_(salt+':'+next));
    CacheService.getScriptCache().remove('auth:'+String(token));
    audit_('ADMIN_PAROL_ALMASHTIRILDI','Admin','','',0,'Admin paroli yangilandi');
    return {ok:true,message:'Admin paroli almashtirildi. Yangi parol bilan qayta kiring.'};
  }finally{lock.releaseLock();}
}
