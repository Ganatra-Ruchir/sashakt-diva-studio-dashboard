// Owner-only export detail. Keep raw production, declarations, and payments separate.
const cents=v=>Math.round(Number(v)*100);
const normal=v=>String(v??'').toLowerCase().replace(/[^a-z0-9]/g,'');
const key=r=>JSON.stringify([r.studio,r.beneficiary_key,r.work_date]);
module.exports=async function(q,args,beneficiary){
const scope=(alias,day)=>"($1::bigint IS NULL OR "+alias+".studio_id=$1) AND "+day+" BETWEEN $2::date AND $3::date AND (b.code NOT LIKE 'DEMO-%' OR lower(coalesce($4::text,'')) LIKE 'demo %') AND ($5::text IS NULL OR strpos(lower(b.name),lower($5))>0 OR strpos(lower(b.code),lower($5))>0)";
const p=[...args,beneficiary||null];
const queries={
production:"SELECT s.name studio,b.code beneficiary_key,b.name beneficiary,coalesce(a.batch_at_work,'Unassigned') batch,p.work_date::text work_date,v.name vendor,pr.name product,a.pieces,a.rate::text rate,a.income::text income,p.id production_id FROM allocations a JOIN production p ON p.id=a.production_id JOIN beneficiaries b ON b.id=a.beneficiary_id JOIN studios s ON s.id=a.studio_id JOIN vendors v ON v.id=p.vendor_id JOIN products pr ON pr.id=p.product_id WHERE "+scope("a","p.work_date")+" ORDER BY p.work_date,a.id LIMIT 20001",
assignments:"SELECT s.name studio,b.code beneficiary_key,b.name beneficiary,coalesce(a.batch_at_assignment,'Unassigned') batch,o.received_date::text received_date,v.name vendor,pr.name product,o.id order_id,o.pieces_received,o.vendor_rate::text vendor_rate,a.id assignment_id,a.pieces,a.rate::text rate,(a.pieces*a.rate)::text expected_income,a.completed_date::text completed_date,CASE WHEN a.completed_date IS NULL THEN 'Pending' ELSE 'Completed' END status FROM studio_work_assignments a JOIN studio_work_orders o ON o.id=a.order_id JOIN beneficiaries b ON b.id=a.beneficiary_id JOIN studios s ON s.id=a.studio_id JOIN vendors v ON v.id=o.vendor_id JOIN products pr ON pr.id=o.product_id WHERE "+scope("a","o.received_date")+" ORDER BY o.received_date,a.id LIMIT 20001",
self:"SELECT s.name studio,b.code beneficiary_key,b.name beneficiary,coalesce(b.batch,'Unassigned') batch,i.work_date::text work_date,i.income::text income,i.outside_income::text outside_income,i.submitted_at::text submitted_at FROM self_income i JOIN beneficiaries b ON b.id=i.beneficiary_id JOIN studios s ON s.id=i.studio_id WHERE "+scope("i","i.work_date")+" ORDER BY i.work_date,i.id LIMIT 20001",
verification:"SELECT s.name studio,b.code beneficiary_key,b.name beneficiary,coalesce(b.batch,'Unassigned') batch,r.work_date::text work_date,r.manager_income::text manager_income,r.beneficiary_income::text beneficiary_income,(r.manager_income-r.beneficiary_income)::text difference,CASE WHEN ir.approved_income IS NOT NULL THEN 'approved' ELSE r.status END status,r.status original_status,COALESCE(ir.approved_income,r.verified_income)::text verified_income,ir.reason,ir.approved_at::text approved_at FROM income_reconciliation r JOIN studios s ON s.id=r.studio_id JOIN beneficiaries b ON b.id=r.beneficiary_id LEFT JOIN income_reviews ir ON ir.beneficiary_id=r.beneficiary_id AND ir.studio_id=r.studio_id AND ir.work_date=r.work_date WHERE "+scope("r","r.work_date")+" ORDER BY r.work_date,r.beneficiary_id LIMIT 20001",
payments:"SELECT s.name studio,b.code beneficiary_key,b.name beneficiary,coalesce(b.batch,'Unassigned') batch,p.payment_date::text payment_date,p.amount::text amount,p.reference FROM payments p JOIN beneficiaries b ON b.id=p.beneficiary_id JOIN studios s ON s.id=p.studio_id WHERE "+scope("p","p.payment_date")+" ORDER BY p.payment_date,p.id LIMIT 20001"};
const all={};
for(const [name,sql] of Object.entries(queries)){all[name]=(await q(sql,p)).rows;if(all[name].length>20000){const e=new Error('Narrow the export date range');e.status=400;throw e}}
const byDay=new Map();for(const r of all.production){const k=key(r);if(!byDay.has(k))byDay.set(k,[]);byDay.get(k).push(r)}
const official=[],unallocated=[];
for(const v of all.verification){
 const rows=byDay.get(key(v))||[];
 const batches=[...new Set(rows.map(r=>r.batch))];
 if(batches.length===1)v.batch=batches[0];
 if(batches.length>1)v.batch='Multiple batches';
 if(v.verified_income==null)continue;
 const total=rows.reduce((n,r)=>n+cents(r.income),0);
 if(rows.length&&total===cents(v.verified_income)){
  for(const r of rows)official.push({...r,month:r.work_date.slice(0,7)+'-01'});
 }else{
  // A day-level approval cannot safely be attributed to any vendor/product.
  const a={...v,month:v.work_date.slice(0,7)+'-01',vendor:'Reconciled adjustment',product:'Not allocated to product',income:v.verified_income};
  official.push(a);unallocated.push(a);
 }
}
for(const r of all.self){const rows=byDay.get(key(r))||[];const batches=[...new Set(rows.map(x=>x.batch))];if(batches.length===1)r.batch=batches[0];else if(batches.length>1)r.batch='Multiple batches'}
const selected=r=>!args[3]||normal(r.batch)===normal(args[3]);
return {...Object.fromEntries(Object.entries(all).map(([k,v])=>[k,v.filter(selected)])),official:official.filter(selected),unallocated:unallocated.filter(selected),ambiguous:all.verification.filter(r=>r.batch==='Multiple batches'&&args[3])};
};
