const ExcelJS=require('exceljs');
const clean=v=>String(v??'').replace(/[\\/*?:\[\]]/g,' ').slice(0,31);
const monthList=(start,end)=>{const a=new Date(start.slice(0,7)+'-01T00:00:00Z'),b=new Date(end.slice(0,7)+'-01T00:00:00Z'),out=[];while(a<=b){out.push(a.toISOString().slice(0,10));a.setUTCMonth(a.getUTCMonth()+1)}return out};
module.exports=async function(report){
 const wb=new ExcelJS.Workbook();wb.creator='Diva Studio';const months=monthList(report.start,report.end);
 const monthName=m=>new Date(m+'T00:00:00Z').toLocaleString('en-IN',{month:'short',year:'numeric',timeZone:'UTC'});
 const header=row=>{row.eachCell({includeEmpty:true},c=>{c.font={bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF245544'}};c.alignment={vertical:'middle',wrapText:true}})};
 const info=wb.addWorksheet('Read me');info.columns=[{width:30},{width:110}];info.addRows([['DIVA DATA','Batch-wise income export'],['Selected dates',report.start+' to '+report.end],['Live records','Manager allocations within exact selected dates.'],['Historical records','Latest legacy workbook, whole months intersecting selected dates; kept separate to prevent double counting.'],['Flagged historical rows','Rows with discrepancies or missing totals are exported separately.'],['Verified income','Independent reconciliation is on a separate sheet; it is never added to manager income.'],['Blank vs zero','Blank = no recorded amount; numeric zero = recorded zero.']]);header(info.getRow(1));
 const sources=[['Live',report.live],['Legacy matched',report.legacy.filter(r=>r.source_status==='matched')],['Legacy flagged',report.legacy.filter(r=>r.source_status!=='matched')]];
 for(const [source,all] of sources){
 for(const studio of [...new Set(all.map(r=>r.studio))]){
 const ws=wb.addWorksheet(clean(source+' '+studio));ws.views=[{state:'frozen',xSplit:2,ySplit:1}];ws.getColumn(1).width=12;ws.getColumn(2).width=34;
 const rows=all.filter(r=>r.studio===studio);
 for(const batch of [...new Set(rows.map(r=>r.batch))].sort()){
 const records=rows.filter(r=>r.batch===batch);
 const keys=[];for(const month of months){const combos=[...new Set(records.filter(r=>r.month===month).map(r=>JSON.stringify([r.vendor||'',r.product||''])))];if(!combos.length)combos.push(JSON.stringify(['','']));for(const combo of combos){const [vendor,product]=JSON.parse(combo);keys.push({month,vendor,product,key:JSON.stringify([month,vendor,product])})}}
 const title=ws.addRow([studio+' | '+batch+' | '+source]);title.font={bold:true,size:14,color:{argb:'FF245544'}};
 const leftTotal=keys.length+3,summaryStart=leftTotal+2,summaryTotal=summaryStart+2+months.length;
 const h1=ws.addRow(['','',...keys.map(k=>monthName(k.month)),'TOTAL','','','MONTHLY INCOME SUMMARY',...months.map(monthName),'TOTAL']);
 let k=0;while(k<keys.length){let j=k+1;while(j<keys.length&&keys[j].month===keys[k].month)j++;if(j-k>1)ws.mergeCells(h1.number,3+k,h1.number,2+j);k=j}
 header(h1);
 header(ws.addRow(['','Vendor',...keys.map(k=>k.vendor),'','','','',...months.map(()=>''),'']));
 header(ws.addRow(['Sr. No.','Beneficiary',...keys.map(k=>k.product),'Total','','Sr. No.','Beneficiary',...months.map(monthName),'Total']));
 const people=new Map();for(const r of records){const key=r.beneficiary_key+'|'+r.beneficiary;let p=people.get(key);if(!p){p={name:r.beneficiary,key:r.beneficiary_key,detail:new Map(),monthly:new Map()};people.set(key,p)}const d=JSON.stringify([r.month,r.vendor||'',r.product||'']);p.detail.set(d,(p.detail.get(d)||0)+Number(r.income));p.monthly.set(r.month,(p.monthly.get(r.month)||0)+Number(r.income))}
 const arr=[...people.values()].sort((a,b)=>a.name.localeCompare(b.name));const first=ws.rowCount+1;for(let i=0;i<arr.length;i++){const p=arr[i],n=ws.rowCount+1;const row=ws.addRow([i+1,p.name+' ('+p.key+')',...keys.map(k=>p.detail.has(k.key)?p.detail.get(k.key):null),{formula:'SUM('+ws.getColumn(3).letter+n+':'+ws.getColumn(leftTotal-1).letter+n+')'},null,i+1,p.name+' ('+p.key+')',...months.map(m=>p.monthly.has(m)?p.monthly.get(m):null),{formula:'SUM('+ws.getColumn(summaryStart+2).letter+n+':'+ws.getColumn(summaryTotal-1).letter+n+')'}]);for(let col=3;col<=summaryTotal;col++){if(col===leftTotal+1||col===summaryStart||col===summaryStart+1)continue;row.getCell(col).numFmt='#,##0.00'}}
 const last=ws.rowCount;const total=ws.addRow(['','TOTAL']);for(let col=3;col<=summaryTotal;col++){if(col===leftTotal+1||col===summaryStart||col===summaryStart+1)continue;const letter=ws.getColumn(col).letter;total.getCell(col).value={formula:'SUM('+letter+first+':'+letter+last+')'};total.getCell(col).numFmt='#,##0.00'}total.font={bold:true};for(let col=3;col<=summaryTotal;col++)ws.getColumn(col).width=Math.max(ws.getColumn(col).width||0,col===summaryStart+1?34:17);ws.addRow([]);
 }
 }}
 const verification=wb.addWorksheet('Income verification');verification.addRow(['Studio','Month','Verified income','Differences','Pending']);header(verification.getRow(1));for(const r of report.verification)verification.addRow([r.studio,r.month,r.verified_income==null?null:Number(r.verified_income),Number(r.differences),Number(r.pending)]);verification.columns.forEach(c=>c.width=22);
 const att=wb.addWorksheet('Attendance');att.addRow(['Studio','Marked','Present','Home-based']);header(att.getRow(1));for(const r of report.attendance)att.addRow([r.studio,Number(r.marked),Number(r.present),Number(r.home_based)]);att.columns.forEach(c=>c.width=22);
 return Buffer.from(await wb.xlsx.writeBuffer()).toString('base64');
};