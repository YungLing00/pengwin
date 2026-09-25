'use strict';
const XLSX = require('xlsx');
const LIMIT = 4 * 1024 * 1024;
const headers = { 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' };
const month = value => {
  if (value instanceof Date && !isNaN(value)) return `${value.getUTCFullYear()}-${String(value.getUTCMonth()+1).padStart(2,'0')}`;
  if (typeof value === 'number' && value > 20000 && value < 90000) {
    const d = XLSX.SSF.parse_date_code(value);
    if (d && d.m >= 1 && d.m <= 12) return `${d.y}-${String(d.m).padStart(2,'0')}`;
  }
  let s = String(value ?? '').trim().replace(/年|[./]/g,'-').replace(/月/g,'').replace(/\s/g,'');
  let m = s.match(/^(\d{4})-?(\d{1,2})(?:-\d{1,2})?$/);
  if (!m && /^\d{3}[-]?\d{1,2}$/.test(s)) {
    m = s.match(/^(\d{3})-?(\d{1,2})$/);
    if (+m[2] < 1 || +m[2] > 12) return null;
    return `${+m[1]+1911}-${String(+m[2]).padStart(2,'0')}`;
  }
  return m && +m[2] >= 1 && +m[2] <= 12 && +m[1] >= 1900 ? `${m[1]}-${String(+m[2]).padStart(2,'0')}` : null;
};
const amount = v => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  let s = String(v ?? '').trim().replace(/[,，\s$＄]/g,'');
  if (/^\(\d+(?:\.\d+)?\)$/.test(s)) s = '-'+s.slice(1,-1);
  return /^-?\d+(?:\.\d+)?$/.test(s) ? Number(s) : null;
};
const unit = v => {
  let s = String(v ?? '').trim().toLowerCase();
  if (/^(元|新臺幣元|新台幣元|ntd|twd)$/.test(s)) return 1;
  if (/^(千元|仟元)$/.test(s)) return 1000;
  if (/^(萬元|萬)$/.test(s)) return 10000;
  if (s === '百萬元') return 1000000;
  return null;
};
const col = n => { let s=''; for(n++;n;n=Math.floor((n-1)/26)) s=String.fromCharCode(65+(n-1)%26)+s;return s; };
function guess(h, kind) {
  const keys={month:[['月份',100],['年月',96],['期間',83],['月別',80],['month',100],['date',78],['日期',78]],amount:[['營業收入',100],['營收',100],['銷售額',91],['銷售收入',90],['收入金額',85],['金額',55],['revenue',100],['sales',90],['amount',55]],unit:[['單位',100],['金額單位',100],['unit',100]]};
  let best={idx:-1,score:0};h.forEach((v,i)=>{let t=String(v??'').trim().toLowerCase().replace(/\s+/g,'');for(const [k,score] of keys[kind]){let n=t===k?score:t.includes(k)?score-18:0;if(n>best.score)best={idx:i,score:n}}});return best;
}
function ingest({name,content}) {
  if(typeof name!=='string'||!/^.{1,200}\.(xlsx|xls|csv)$/i.test(name)||typeof content!=='string'||content.length>LIMIT*1.4||!/^([A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(content)) throw Error('檔案格式或內容不正確');
  const buf=Buffer.from(content,'base64');if(buf.length>3*1024*1024)throw Error('單一檔案不可超過 3 MB');
  const wb=XLSX.read(buf,{type:'buffer',cellDates:false});
  return wb.SheetNames.map(sheetName=>{
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,defval:'',raw:true,blankrows:false});
    if(!rows.length)return null;
    let headerRow=0,best=-1;for(let i=0;i<Math.min(rows.length,15);i++){let score=guess(rows[i]||[],'month').score+guess(rows[i]||[],'amount').score;if(score>best){best=score;headerRow=i}}
    const h=(rows[headerRow]||[]).map((x,i)=>String(x??'').trim()||`未命名欄 ${col(i)}`);
    if(!h.length)return null;
    let mg=guess(h,'month'),ag=guess(h,'amount'),ug=guess(h,'unit');
    return {file:name,name:sheetName,rows,headerRow,headers:h,map:{month:mg.idx,amount:ag.idx,unit:ug.idx},scores:{month:mg.score,amount:ag.score,unit:ug.score},fixedUnit:'',confirmed:false};
  }).filter(Boolean);
}
function analyze(sheets){
  if(!Array.isArray(sheets)||sheets.length>30)throw Error('工作表數量不正確');
  const data=[],issues=[];
  for(const s of sheets){
    if(!s.confirmed)continue;
    if(typeof s.file!=='string'||typeof s.name!=='string'||!Array.isArray(s.rows)||s.rows.length>10000||!s.map||!Number.isInteger(s.headerRow))throw Error('工作表資料不正確');
    const {month:mi,amount:ai,unit:ui}=s.map;
    if(!Number.isInteger(mi)||!Number.isInteger(ai)||!Number.isInteger(ui)||mi<0||ai<0||mi===ai||mi>200||ai>200||ui>200||ui< -1)throw Error('欄位對應不正確');
    let seen=new Map(),months=[],prefix=`${s.file} / ${s.name}`;
    for(let i=s.headerRow+1;i<s.rows.length;i++){
      const row=s.rows[i];if(!Array.isArray(row))continue;
      if(row.every(v=>String(v??'').trim()===''))continue;
      let location=`${prefix} / ${col(mi)}${i+1}、${col(ai)}${i+1}`;
      let m=month(row[mi]),n=amount(row[ai]),factor=ui>=0?unit(row[ui]):null;
      if(!factor&&s.fixedUnit)factor=unit(s.fixedUnit);
      if(!m){issues.push({type:'日期格式待確認',location,detail:`無法辨識月份：${String(row[mi]??'').slice(0,100)}`});continue}
      if(n===null){issues.push({type:'營收金額待確認',location,detail:`無法辨識金額：${String(row[ai]??'').slice(0,100)}`});continue}
      if(!factor)issues.push({type:'單位待確認',location,detail:'未換算金額；請回原表確認單位'});
      const first=seen.get(m);if(first)issues.push({type:'重複月份',location,detail:`${m} 與 ${first} 重複；未自動合併`});else seen.set(m,location);
      months.push(m);data.push({month:m,amount:factor?n*factor:null,unit:factor?'元':'未確認',status:!factor?'單位待確認':first?'重複月份':'已轉換',file:s.file,sheet:s.name,monthCell:`${col(mi)}${i+1}`,amountCell:`${col(ai)}${i+1}`,unitCell:ui<0?'全表指定':`${col(ui)}${i+1}`,source:location});
    }
    if(months.length>1){const nums=[...new Set(months)].map(x=>+x.slice(0,4)*12+(+x.slice(5)-1)).sort((a,b)=>a-b);for(let n=nums[0]+1;n<nums.at(-1)&&n-nums[0]<1200;n++)if(!nums.includes(n))issues.push({type:'缺少月份',location:prefix,detail:`${Math.floor(n/12)}-${String(n%12+1).padStart(2,'0')} 無資料`})}
  }
  return {data,issues};
}
async function handler(req,res){
  for(const [k,v] of Object.entries(headers))res.setHeader(k,v);
  if(req.method==='GET')return res.status(200).json({ok:true,mode:'server',version:1});
  if(req.method!=='POST')return res.status(405).json({error:'Method not allowed'});
  try{
    let body=req.body;if(typeof body==='string')body=JSON.parse(body);
    if(!body||JSON.stringify(body).length>LIMIT*2)return res.status(413).json({error:'請縮小資料量'});
    if(body.action==='ingest')return res.status(200).json({sheets:ingest(body)});
    if(body.action==='analyze')return res.status(200).json(analyze(body.sheets));
    return res.status(400).json({error:'無效操作'});
  }catch(e){return res.status(400).json({error:e.message||'資料解析失敗'})}
}
module.exports=handler;
module.exports.ingest=ingest;
module.exports.analyze=analyze;
