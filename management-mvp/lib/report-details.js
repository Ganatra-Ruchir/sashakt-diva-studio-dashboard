// Owner-only export detail. Keep raw production, declarations, and payments separate.
const cents=v=>Math.round(Number(v)*100);
const normal=v=>String(v??'').toLowerCase().replace(/[^a-z0-9]/g,'');
const key=r=>JSON.stringify([r.studio,r.beneficiary_key,r.work_date]);
module.exports=async function(q,args,beneficiary){
const scope=(alias,day)=>"($1::bigint IS NULL OR "+alias+".studio_id=$1) AND "+day+" BETWEEN $2::date AND $3::date AND (b.code NOT LIKE 'DEMO-%' OR lower(coalesce($4::text,'')) LIKE 'demo %') AND ($5::text IS NULL OR strpos(lower(b.name),lower($5))>0 OR strpos(lower(b.code),lower($5))>0)";
const p=[...args,beneficiary||null];
const queries={
production:"SELECT s.name studio,b.code beneficiary_key,b.name beneficiary,coalesce(a.batch_at_work,'Unassigned') batch,p.work_date::text work_date,v.name vendor,pr.name product,a.accepted AS pieces,a.rate::text rate,a.earned::text income,p.id production_id,a.rejected,a.accepted+a.rejected AS processed FROM diva_production_fact a JOIN production p ON p.id=a.production_id JOIN beneficiaries b ON b.id=a.beneficiary_id JOIN studios s ON s.id=a.studio_id JOIN vendors v ON v.id=p.vendor_id JOIN products pr ON pr.id=p.product_id WHERE "+scope("a","p.work_date")+" ORDER BY p.work_date,a.allocation_id LIMIT 20001",
assignments:"SELECT s.name studio,b.code beneficiary_key,b.name beneficiary,coalesce(a.batch_at_assignment,'Unassigned') batch,o.received_date::text received_date,v.name vendor,pr.name product,o.id order_id,o.order_reference,o.pieces_received,o.vendor_rate::text vendor_rate,a.id assignment_id,a.pieces,a.rate::text rate,(a.pieces*a.rate)::text expected_income,a.completed_date::text completed_date,COALESCE(g.accepted,CASE WHEN a.completed_date IS NOT NULL THEN a.pieces ELSE 0 END)::int accepted,COALESCE(g.rejected,0)::int rejected,CASE WHEN a.completed_date IS NOT NULL THEN 0 ELSE a.pieces-COALESCE(g.accepted,0)-COALESCE(g.rejected,0) END::int remaining,CASE WHEN a.completed_date IS NOT NULL THEN 'Closed' WHEN g.accepted+g.rejected>0 THEN 'In progress' ELSE 'Pending' END status FROM studio_work_assignments a JOIN studio_work_orders o ON o.id=a.order_id JOIN beneficiaries b ON b.id=a.beneficiary_id JOIN studios s ON s.id=a.studio_id JOIN vendors v ON v.id=o.vendor_id JOIN products pr ON pr.id=o.product_id LEFT JOIN LATERAL (SELECT sum(accepted) accepted,sum(rejected) rejected FROM studio_work_progress WHERE assignment_id=a.id) g ON true WHERE "+scope("a","o.received_date")+" ORDER BY o.received_date,a.id LIMIT 20001",
self:"SELECT s.name studio,b.code beneficiary_key,b.name beneficiary,coalesce(b.batch,'Unassigned') batch,i.work_date::text work_date,i.income::text income,i.outside_income::text outside_income,i.submitted_at::text submitted_at FROM self_income i JOIN beneficiaries b ON b.id=i.beneficiary_id JOIN studios s ON s.id=i.studio_id WHERE "+scope("i","i.work_date")+" ORDER BY i.work_date,i.id LIMIT 20001",
verification:"SELECT s.name studio,b.code beneficiary_key,b.name beneficiary,coalesce(b.batch,'Unassigned') batch,r.work_date::text work_date,r.manager_income::text manager_income,r.beneficiary_income::text beneficiary_income,(r.manager_income-r.beneficiary_income)::text difference,r.status,r.original_status,r.verified_income::text verified_income,r.reason,r.approved_at::text approved_at FROM diva_income_day r JOIN studios s ON s.id=r.studio_id JOIN beneficiaries b ON b.id=r.beneficiary_id WHERE "+scope("r","r.work_date")+" ORDER BY r.work_date,r.beneficiary_id LIMIT 20001",
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
const balances=(await q("SELECT s.name studio,b.code beneficiary_key,b.name beneficiary,COALESCE(b.batch,'Unassigned') batch,COALESCE(e.earned,0)::text earned,COALESCE(e.verified,0)::text verified,COALESCE(p.paid,0)::text paid,(COALESCE(e.verified,0)-COALESCE(p.paid,0))::text net_verified, (COALESCE(e.earned,0)-COALESCE(p.paid,0))::text net_production FROM beneficiaries b JOIN studios s ON s.id=b.studio_id LEFT JOIN LATERAL (SELECT sum(r.manager_income) earned,sum(r.verified_income) verified FROM diva_income_day r WHERE r.beneficiary_id=b.id AND r.studio_id=b.studio_id AND r.work_date <= $3::date) e ON true LEFT JOIN LATERAL (SELECT sum(amount) paid FROM payments WHERE beneficiary_id=b.id AND studio_id=b.studio_id AND payment_date <= $3::date) p ON true WHERE ($1::bigint IS NULL OR b.studio_id=$1) AND $2::date IS NOT NULL AND (b.code NOT LIKE 'DEMO-%' OR lower(coalesce($4::text,'')) LIKE 'demo %') AND ($5::text IS NULL OR strpos(lower(b.name),lower($5))>0 OR strpos(lower(b.code),lower($5))>0) AND (e.earned IS NOT NULL OR e.verified IS NOT NULL OR p.paid IS NOT NULL) ORDER BY s.name,b.name LIMIT 20001",p)).rows;if(balances.length>20000){const e=new Error('Narrow the export scope');e.status=400;throw e}
return {...Object.fromEntries(Object.entries(all).map(([k,v])=>[k,v.filter(selected)])),balances:balances.filter(selected),official:official.filter(selected),unallocated:unallocated.filter(selected),ambiguous:all.verification.filter(r=>r.batch==='Multiple batches'&&args[3])};
};
