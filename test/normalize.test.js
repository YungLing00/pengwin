const test=require('node:test');
const assert=require('node:assert/strict');
const XLSX=require('xlsx');
const api=require('../api/normalize');

test('ingest and normalize ROC months, units, duplicate months and missing months',()=>{
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
    ['月度報表'],['年月','銷售額','金額單位'],['115年1月',88,'萬元'],['115年2月',93,'萬元'],['115年4月',104,'萬元'],['115年4月',105,'萬元'],['115年5月',98,'待確認']
  ]),'銷售明細');
  const sheets=api.ingest({name:'虛構企業.xlsx',content:XLSX.write(wb,{bookType:'xlsx',type:'buffer'}).toString('base64')});
  assert.deepEqual(sheets[0].map,{month:0,amount:1,unit:2});
  sheets[0].confirmed=true;
  const {data,issues}=api.analyze(sheets);
  assert.equal(data[0].month,'2026-01');
  assert.equal(data[0].amount,880000);
  assert.equal(data[0].amountCell,'B3');
  assert.equal(data[4].amount,null);
  assert.deepEqual(issues.map(x=>x.type).sort(),['單位待確認','缺少月份','重複月份'].sort());
});

test('reject invalid column mapping',()=>{
  assert.throws(()=>api.analyze([{confirmed:true,file:'a',name:'b',headerRow:0,rows:[['x','y'],['1','2']],map:{month:0,amount:0,unit:-1}}]),/欄位對應/);
});
