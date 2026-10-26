'use strict';
const http=require('node:http'),crypto=require('node:crypto');
const port=Number(process.env.PORT||10000), secret=process.env.HISOB_SESSION_SECRET||'';
const username=process.env.HISOB_ADMIN_USER||'',hash=process.env.HISOB_ADMIN_PASSWORD_SCRYPT||'';
const ready=username.length>0&&hash.startsWith('scrypt$')&&secret.length>=32;
const sessions=new Map();
const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const respond=(r,status,body,type='text/html; charset=utf-8',extra={})=>{r.writeHead(status,{'Content-Type':type,'Cache-Control':'no-store','X-Content-Type-Options':'nosniff','Content-Security-Policy':"default-src 'none'; style-src 'unsafe-inline'; form-action 'self'; base-uri 'none'; frame-ancestors 'none'",'Strict-Transport-Security':'max-age=31536000',...extra});r.end(body)};
const page=(inner)=>'<!doctype html><html lang="uz"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>HISOB WEB — Kirish</title><style>*{box-sizing:border-box}body{margin:0;background:#f1f7f4;font:15px Segoe UI,Arial,sans-serif;color:#173a2d;min-height:100vh;display:grid;place-items:center}.card{width:min(420px,92vw);padding:32px;background:white;border:1px solid #dce8df;border-radius:18px;box-shadow:0 15px 50px #00331b15}h1{margin:0 0 6px;color:#087b57}p{color:#688475;line-height:1.5}label{display:block;margin:15px 0 6px;font-weight:700}input{width:100%;padding:13px;border:1px solid #cfdfd5;border-radius:9px;font:inherit}button{background:#087b57;color:white;border:0;padding:13px 18px;border-radius:9px;margin-top:17px;font-weight:bold;cursor:pointer}.err{background:#fff2ee;padding:11px;border-radius:9px;color:#9b3f36}</style><div class="card">'+inner+'</div></html>';
function token(req){const m=String(req.headers.cookie||'').match(/(?:^|;\s*)hisob_session=([a-f0-9]{64})(?:;|$)/);return m?.[1]||''}
function authenticated(req){const t=token(req),v=sessions.get(t);if(!v)return false;if(v<Date.now()){sessions.delete(t);return false}return true}
function validPassword(input){try{const [,salt,hex]=hash.split('$');if(!salt||!/^[a-f0-9]{128}$/i.test(hex))return false;const expected=Buffer.from(hex,'hex'),derived=crypto.scryptSync(input,salt,64);return crypto.timingSafeEqual(expected,derived)}catch{return false}}
function body(req){return new Promise((resolve,reject)=>{let b='';req.on('data',x=>{b+=x;if(b.length>4096){reject(Error('large'));req.destroy()}});req.on('end',()=>resolve(new URLSearchParams(b)));req.on('error',reject)})}
let failures=0,windowStart=Date.now();
const server=http.createServer(async(req,res)=>{
 try{
  const url=new URL(req.url,'http://localhost');
  if(url.pathname==='/health')return respond(res,200,JSON.stringify({ok:true,configured:ready,checks:{loginPresent:username.length>0,passwordFormat: /^scrypt\$[a-f0-9]{32}\$[a-f0-9]{128}$/i.test(hash),sessionSecretLongEnough:secret.length>=32}}), 'application/json; charset=utf-8');
  if(!ready)return respond(res,503,page('<h1>HISOB WEB</h1><p>Login server tayyorlandi. Administrator hisobini xavfsiz sozlash hali yakunlanmagan. Ma’lumotlarga kirish yopiq.</p>'));
  if(req.method==='GET'&&url.pathname==='/login')return respond(res,200,page('<h1>HISOB.</h1><p>Himoyalangan tizimga kirish</p><form method="post" action="/login"><label>Login</label><input name="user" autocomplete="username" required maxlength="80"><label>Parol</label><input name="password" type="password" autocomplete="current-password" required><button>Kirish</button></form>'));
  if(req.method==='POST'&&url.pathname==='/login'){
   const now=Date.now();if(now-windowStart>900000){windowStart=now;failures=0}if(failures>=20)return respond(res,429,page('<div class="err">Kirish urinishlari vaqtincha cheklangan.</div>'));
   const p=await body(req),user=p.get('user')||'',password=p.get('password')||'';
   const match=Buffer.from(crypto.createHash('sha256').update(user).digest()),target=Buffer.from(crypto.createHash('sha256').update(username).digest());
   if(!crypto.timingSafeEqual(match,target)||!validPassword(password)){failures++;return respond(res,401,page('<div class="err">Login yoki parol noto‘g‘ri.</div><p><a href="/login">Qayta kirish</a></p>'))}
   failures=0;const t=crypto.randomBytes(32).toString('hex');sessions.set(t,Date.now()+8*3600_000);
   return respond(res,303,'', 'text/plain; charset=utf-8',{'Location':'/private','Set-Cookie':'hisob_session='+t+'; HttpOnly; Secure; SameSite=Strict; Path=/; Max-Age=28800'});
  }
  if(req.method==='POST'&&url.pathname==='/logout'){sessions.delete(token(req));return respond(res,303,'','text/plain',{'Location':'/login','Set-Cookie':'hisob_session=; HttpOnly; Secure; SameSite=Strict; Path=/; Max-Age=0'})}
  if(!authenticated(req))return respond(res,303,'','text/plain',{'Location':'/login'});
  if(url.pathname==='/private')return respond(res,200,page('<h1>HISOB WEB</h1><p>Kirish tasdiqlandi. To‘liq HISOB_BAZA nusxasi serverdagi maxfiy saqlashga yuklanmaguncha ko‘rsatish bloklangan.</p><form method="post" action="/logout"><button>Chiqish</button></form>'));
  return respond(res,404,page('<h1>404</h1>'));
 }catch{return respond(res,500,'Server xatosi','text/plain')}
});
server.listen(port,'0.0.0.0');
