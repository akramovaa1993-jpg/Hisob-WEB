
/** HISOB V15.1 — GPX jo‘natmalar: faqat yangi GPX_jonatmalar varag‘ida ishlaydi.
 * Admin sessiyasi bilan himoyalangan; mavjud Mahsulotlar/Kirimlar/Loyihalar o‘zgarmaydi.
 * Qabul qilinganda yozuv ko‘chirilmaydi yoki o‘chirilmaydi — holati yangilanadi.
 */
const GPX151_SHEET = 'GPX_jonatmalar';
const GPX151_HEADERS = ['id','sana','akt','loyiha','marka','birlik','miqdor','bir_dona_kg','jami_kg','holat','jonatuvchi','jonatilgan_vaqt','qabul_qiluvchi','qabul_qilingan_vaqt','izoh'];
function gpx151Sheet_(){
  const sh=tab_(GPX151_SHEET,GPX151_HEADERS);
  const head=sh.getRange(1,1,1,GPX151_HEADERS.length).getDisplayValues()[0];
  if(GPX151_HEADERS.some((x,i)=>x!==head[i]))throw Error('GPX varag‘i ustunlari o‘zgargan. Ma’lumotni saqlash to‘xtatildi.');
  return sh;
}
function gpx151Read_(){
  const sh=gpx151Sheet_(),last=sh.getLastRow();
  if(last<2)return [];
  return sh.getRange(2,1,last-1,GPX151_HEADERS.length).getValues().filter(r=>r[0]).map(r=>{
    const obj={};GPX151_HEADERS.forEach((x,i)=>obj[x]=r[i] instanceof Date?Utilities.formatDate(r[i],hisobTZ_(),'yyyy-MM-dd HH:mm:ss'):r[i]);
    obj.miqdor=Number(obj.miqdor)||0;obj.bir_dona_kg=Number(obj.bir_dona_kg)||0;obj.jami_kg=Number(obj.jami_kg)||0;
    return obj;
  });
}
function apiGpxListV151(token){
  requireSession_(token);
  return {items:gpx151Read_()};
}
function gpx151Text_(s,max,label,required){
  const v=String(s===undefined||s===null?'':s).trim().replace(/\s+/g,' ');
  if((required&&!v)||v.length>max)throw Error(label+' noto‘g‘ri (maksimal '+max+' belgi)');
  if(/^[=+@]/.test(v)||/[\u0000-\u001f]/.test(v))throw Error(label+' xavfsiz matn bo‘lishi kerak');
  return v;
}
function gpx151Check_(v){
  if(!v||typeof v!=='object'||Array.isArray(v))throw Error('GPX satri noto‘g‘ri');
  const sana=gpx151Text_(v.sana,10,'Sana',true);
  if(!/^20\d\d-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/.test(sana)||isNaN(new Date(sana+'T12:00:00Z').getTime())||new Date(sana+'T12:00:00Z').toISOString().slice(0,10)!==sana)throw Error('Sana YYYY-MM-DD formatida bo‘lsin');
  const loyiha=gpx151Text_(v.loyiha,80,'Loyiha',true);
  const marka=gpx151Text_(v.marka,100,'Marka',true);
  const akt=gpx151Text_(v.akt,70,'Akt',false);
  const birlik=gpx151Text_(v.birlik||'kg',20,'Birlik',true);
  const izoh=gpx151Text_(v.izoh,350,'Izoh',false);
  const miqdor=Number(v.miqdor),bir_dona_kg=Number(v.bir_dona_kg);
  if(!Number.isInteger(miqdor)||miqdor<1||miqdor>10000000)throw Error('Miqdor musbat butun son bo‘lsin');
  if(!Number.isFinite(bir_dona_kg)||bir_dona_kg<=0||bir_dona_kg>100000000)throw Error('1 dona kg musbat bo‘lishi kerak');
  const jami_kg=Math.round(miqdor*bir_dona_kg*1000)/1000;
  if(!Number.isFinite(jami_kg)||jami_kg<=0||jami_kg>1e12)throw Error('Jami og‘irlik me’yordan oshdi');
  const holat=v.holat==='accepted'?'accepted':'sent';
  return {sana,akt,loyiha,marka,birlik,miqdor,bir_dona_kg,jami_kg,holat,izoh};
}
function gpx151Identity_(o){return [o.sana,o.akt,o.loyiha,o.marka,o.miqdor,o.bir_dona_kg].join('|').toUpperCase();}
function apiGpxAddBatchV151(token,items){
  requireSession_(token);
  if(!Array.isArray(items)||items.length<1||items.length>350)throw Error('Bir marta 1–350 ta GPX yozuvi mumkin');
  const checked=items.map(gpx151Check_);
  const duplicates=new Set();
  checked.forEach(x=>{const key=gpx151Identity_(x);if(duplicates.has(key))throw Error('Fayl ichida takroriy GPX: '+x.marka);duplicates.add(key)});
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    requireSession_(token);
    const sh=gpx151Sheet_(),existing=gpx151Read_(),known=new Set(existing.map(gpx151Identity_));
    for(const x of checked)if(known.has(gpx151Identity_(x)))throw Error(x.loyiha+' / '+x.marka+' — aynan shu GPX avval kiritilgan. Takroriy import bekor qilindi.');
    const stamp=now_(),values=checked.map(x=>{
      const id=Utilities.getUuid(),accepted=x.holat==='accepted';
      return [id,x.sana,x.akt,x.loyiha,x.marka,x.birlik,x.miqdor,x.bir_dona_kg,x.jami_kg,x.holat,'Admin panel',stamp,accepted?'Import: qabul qilingan':'', '',x.izoh];
    });
    // HISOB V15.6: validate available warehouse stock BEFORE writing GPX.
    // Each matching GPX marka represents one complete plan bundle.
    const warehouse=sk156PlanForNewGpx_(values);
    const start=sh.getLastRow()+1;
    if(sh.getMaxRows()<start+values.length-1)sh.insertRowsAfter(sh.getMaxRows(),start+values.length-1-sh.getMaxRows());
    sh.getRange(start,1,values.length,GPX151_HEADERS.length).setValues(values);
    // Write allocation after GPX. If this fails, old-GPX reconciliation repairs the ledger;
    // the GPX record remains and duplicate import is blocked.
    try { sk156Append_(warehouse.rows); }
    catch(e) { throw Error('GPX saqlandi, lekin Sklad rasxod yozilmadi. GPXni qayta yuklamang! Sklad → Eski GPXni tekshirish orqali tuzating. '+String(e.message||e)); }
    audit_('GPX_JONATILDI','Admin panel','','',checked.length,checked.length+' ta GPX yozuvi, '+checked.reduce((s,x)=>s+x.jami_kg,0)+' kg');
    return {count:checked.length,items:values.map((r)=>({id:r[0]}))};
  }finally{lock.releaseLock()}
}
function apiGpxAcceptV151(token,id){
  requireSession_(token);
  if(typeof id!=='string'||!/^[a-f0-9-]{36}$/i.test(id))throw Error('GPX identifikatori noto‘g‘ri');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    requireSession_(token);
    const sh=gpx151Sheet_(),last=sh.getLastRow();
    if(last<2)throw Error('GPX topilmadi');
    const ids=sh.getRange(2,1,last-1,1).getValues(),index=ids.findIndex(r=>String(r[0])===id);
    if(index<0)throw Error('GPX topilmadi yoki o‘chirilgan');
    const row=index+2,record=sh.getRange(row,1,1,GPX151_HEADERS.length).getValues()[0];
    if(record[9]==='accepted')return {changed:false,already:true};
    if(record[9]!=='sent')throw Error('GPX holati noto‘g‘ri');
    const stamp=now_();sh.getRange(row,10).setValue('accepted');sh.getRange(row,13,1,2).setValues([['Admin panel',stamp]]);
    audit_('GPX_QABUL_QILINDI','Admin panel',String(record[3]),String(record[4]),Number(record[6]),'Akt: '+String(record[2])+'; '+Number(record[8])+' kg');
    return {changed:true,acceptedAt:stamp};
  }finally{lock.releaseLock()}
}


/** HISOB V15.2 GPX PNG hisobotni Telegramga yuborish.
 * Google Slides/Drive ishlatilmaydi. Boshqa GPX/Mahsulotlar/Kirimlar ma'lumotlari yozilmaydi.
 * Faqat tasdiqlangan admin sessiyasi, bot tokeni va ALLOWED_TELEGRAM_IDS dagi 1-admin.
 */
function apiGpxSendReportV152(token, pngImages, selection) {
  requireSession_(token);
  if (!Array.isArray(pngImages) || pngImages.length < 1 || pngImages.length > 8) {
    throw Error('Hisobot rasmlari soni 1–8 oralig‘ida bo‘lishi kerak.');
  }
  const sel = selection && typeof selection === 'object' ? selection : {};
  const project = String(sel.project || 'all').trim();
  const from = String(sel.from || '').trim();
  const to = String(sel.to || '').trim();
  const validDate = s => !s || /^20\d\d-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/.test(s);
  if (!validDate(from) || !validDate(to) || project.length > 80) throw Error('Hisobot filtri noto‘g‘ri.');
  const all = gpx151Read_().filter(r => (!from || String(r.sana) >= from) && (!to || String(r.sana) <= to) && (project === 'all' || String(r.loyiha) === project));
  if (!all.length) throw Error('Ushbu filtr bo‘yicha GPX yozuvlari topilmadi.');
  // Raqamlarni foydalanuvchi yuborgan rasmdan emas, bevosita Google Sheetsdan hisoblaymiz.
  const accepted = all.filter(r => r.holat === 'accepted');
  const notAccepted = all.filter(r => r.holat !== 'accepted');
  const sum = rows => rows.reduce((n,r) => n + (Number(r.jami_kg) || 0), 0);
  const mass = n => Number(n).toFixed(1).replace('.', ',');
  const bot = cfg_('BOT_TOKEN');
  const allowed = String(cfg_('ALLOWED_TELEGRAM_IDS') || '').split(',').map(x => x.trim()).filter(x => /^\d{5,20}$/.test(x));
  const chatId = allowed[0];
  if (!bot || !chatId) throw Error('Telegram bot tokeni yoki asosiy administrator ID sozlanmagan.');
  const caption = '📊 HISOB • GPX Hisobot\n' +
    (project === 'all' ? 'Barcha loyihalar' : 'Loyiha: ' + project) + '\n' +
    '📦 Jo‘natilgan, qabul qilinmagan: ' + notAccepted.length + ' ta • ' + mass(sum(notAccepted)) + ' kg\n' +
    '✅ Qabul qilingan: ' + accepted.length + ' ta • ' + mass(sum(accepted)) + ' kg\n' +
    '📋 Jami: ' + all.length + ' ta • ' + mass(sum(all)) + ' kg';
  let sent = 0;
  for (let i=0; i<pngImages.length; i++) {
    requireSession_(token); // har rasm oldidan sessiyani qayta tekshiradi
    const src = pngImages[i];
    if (typeof src !== 'string' || src.length > 13000000 || src.length < 500 || !/^[A-Za-z0-9+/=]+$/.test(src)) {
      throw Error('Hisobot PNG rasmi noto‘g‘ri formatda.');
    }
    const bytes = Utilities.base64Decode(src);
    if (bytes.length > 9500000 || bytes.length < 200 || bytes.slice(0,8).some((v,j) => (v & 255) !== [137,80,78,71,13,10,26,10][j])) {
      throw Error('Hisobot haqiqiy PNG bo‘lishi va 9,5 MB dan kichik bo‘lishi shart.');
    }
    const file = Utilities.newBlob(bytes, 'image/png', 'HISOB_GPX_HISOBOT_'+(i+1)+'.png');
    let resp;
    try {
      resp = UrlFetchApp.fetch('https://api.telegram.org/bot'+bot+'/sendPhoto', {
        method:'post', payload:{chat_id:chatId,photo:file,caption:i===0?caption:'GPX Hisobot davomı '+(i+1)+'/'+pngImages.length},
        muteHttpExceptions:true
      });
    } catch (e) {
      throw Error('Telegramga yuborilmadi ('+sent+'/'+pngImages.length+'). Aloqani tekshiring.');
    }
    let response = {};
    try { response = JSON.parse(resp.getContentText()); } catch(e) {}
    if (resp.getResponseCode() !== 200 || !response.ok) {
      throw Error('Telegram rasmni qabul qilmadi: HTTP '+resp.getResponseCode()+'; yuborildi '+sent+'/'+pngImages.length+'. Bot administratorga yozishi uchun chatda avval /start yuborilgan bo‘lsin.');
    }
    sent++;
  }
  try { audit_('GPX_HISOBOT_TELEGRAM','Admin panel',project==='all'?'':project,'',all.length,'Telegramga '+sent+' ta PNG GPX hisobot, '+mass(sum(all))+' kg'); } catch(e) {}
  return {sent:sent,records:all.length,accepted:accepted.length,totalKg:sum(all)};
}