// สร้างตารางสรุปรายปีจากไฟล์ดิบ (ไม่แก้ไฟล์ดิบ) → data/labor-market/extracted/
// ใช้: node build_extracted.js <โฟลเดอร์ data/labor-market>
const fs = require('fs'), path = require('path');
const BASE = process.argv[2];
const RAW = path.join(BASE, 'raw'), OUT = path.join(BASE, 'extracted');
fs.mkdirSync(OUT, {recursive: true});

function parseCSV(t){
  t = t.replace(/^﻿+/, '');
  const rows=[]; let r=[], f='', q=false;
  for(let i=0;i<t.length;i++){ const c=t[i];
    if(q){ if(c=='"'){ if(t[i+1]=='"'){f+='"';i++;} else q=false; } else f+=c; }
    else if(c=='"') q=true; else if(c==','){r.push(f);f='';} else if(c=='\n'){r.push(f);rows.push(r);r=[];f='';} else if(c!='\r') f+=c; }
  if(f||r.length){r.push(f);rows.push(r);} return rows;
}
const num = s => { s=(s||'').replace(/[,\s]/g,''); return s===''||s==='-' ? 0 : Number(s); };
const csvCell = v => /[",\n]/.test(String(v)) ? `"${String(v).replace(/"/g,'""')}"` : v;
function writeCSV(name, header, rows, notes){
  const body = [header, ...rows].map(r=>r.map(csvCell).join(',')).join('\n');
  fs.writeFileSync(path.join(OUT, name), '﻿' + body + '\n');   // BOM ให้ Excel อ่านภาษาไทยถูก
  console.log('wrote', name, rows.length, 'rows');
}
const YEARS_CE = [2021,2022,2023,2024,2025];    // = พ.ศ. 2564–2568 (5 ปีเต็มล่าสุด)

/* ---------- A/B. กรมการจัดหางาน: ตำแหน่งงานว่าง vs ผู้สมัคร vs บรรจุงาน รายปี ---------- */
const DOE_OCC = {
  '1':'ผู้บัญญัติกฎหมาย ข้าราชการระดับอาวุโส ผู้จัดการ','2':'ผู้ประกอบวิชาชีพต่าง ๆ',
  '3':'ช่างเทคนิคและผู้ปฏิบัติงานที่เกี่ยวข้อง','4':'เสมียน เจ้าหน้าที่',
  '5':'พนักงานบริการ พนักงานขายในร้านค้าและตลาด','6':'ผู้ปฏิบัติงานฝีมือด้านการเกษตรและประมง',
  '7':'ผู้ปฏิบัติงานโดยใช้ฝีมือในธุรกิจต่าง ๆ','8':'ผู้ปฏิบัติงานในโรงงาน ผู้ควบคุมเครื่องจักรและงานประกอบ',
  '9':'อาชีพงานพื้นฐาน','T':'ผู้ฝึกงาน','X':'ผู้ปฏิบัติงานซึ่งมิอาจจัดประเภท','0':'ทหาร',
};  // ชื่อจาก DataDict ของกรมการจัดหางาน (raw/doe/doe_vacancies_by_occupation_datadict.xls)
const DOE_EDU = {ElementarySchool:'ประถมศึกษา',Secondaryschool:'มัธยมศึกษา',VocationalCertificate:'ปวช.',
  HighVocationalCertificate:'ปวส.',Diploma:'อนุปริญญา',BachelorDegrees:'ปริญญาตรี',MasterDegrees:'ปริญญาโท',DoctorDegrees:'ปริญญาเอก'};

function doeAnnual(file){   // → {year: {col: total}} รวม 12 เดือน
  const rows = parseCSV(fs.readFileSync(path.join(RAW,'doe',file),'utf8'));
  const head = rows[1].map(s=>s.trim()), out = {}, months = {};
  rows.slice(2).forEach(r=>{
    const y = Number(r[0]); if(!YEARS_CE.includes(y)) return;
    months[y] = (months[y]||0) + 1;
    head.forEach((h,i)=>{ if(i<2) return; (out[y]=out[y]||{})[h] = (out[y][h]||0) + num(r[i]); });
  });
  YEARS_CE.forEach(y=>{ if(months[y]!==12) throw new Error(`${file} ปี ${y} มี ${months[y]} เดือน`); });
  return out;
}
for(const [kind, labels, cols] of [['occupation', DOE_OCC, Object.keys(DOE_OCC)], ['education', DOE_EDU, Object.keys(DOE_EDU)]]){
  const vac = doeAnnual(`doe_vacancies_by_${kind}_monthly.csv`);
  const app = doeAnnual(`doe_applicants_by_${kind}_monthly.csv`);
  const pla = doeAnnual(`doe_placements_by_${kind}_monthly.csv`);
  const rows = [];
  YEARS_CE.forEach(y=>{
    [...cols, 'Total'].forEach(c=>{
      const v = vac[y][c]||0, a = app[y][c]||0, p = pla[y][c]||0;
      if(!v && !a && !p) return;
      rows.push([y+543, y, c, c==='Total'?'รวม':labels[c], v, a, p, v ? (a/v).toFixed(2) : '']);
    });
  });
  writeCSV(`doe_annual_by_${kind}_2564-2568.csv`,
    ['ปี_พศ','ปี_คศ','รหัส','กลุ่ม','ตำแหน่งงานว่าง','ผู้สมัครงาน','ผู้บรรจุงาน','ผู้สมัครต่อตำแหน่ง'], rows);
}

/* ---------- C. สสช.: ผู้มีงานทำ จำแนกตามอาชีพ ทั้งประเทศ เฉลี่ยรายปี ---------- */
// ทั้งประเทศ = ผลรวมทุกภาค (ไม่มีแถวรวมทั้งประเทศในไฟล์; ภาคไม่ซ้อนกัน — ตรวจแล้วผลรวม ≈ ยอดผู้มีงานทำทั้งประเทศ)
// ชื่ออาชีพสะกดต่างกันบางปี → ตัดช่องว่างก่อนจับกลุ่ม
{
  const rows = parseCSV(fs.readFileSync(path.join(RAW,'nso','nso_employed_by_occupation.csv'),'utf8'));
  const h = rows[0].map(s=>s.trim());
  const [iy,iq,io,iv] = ['year','quarter','occupation','value'].map(k=>h.indexOf(k));
  const key = s => s.replace(/\s+/g,'');
  const label = {}, sum = {}, quarters = {};
  rows.slice(1).forEach(r=>{
    const y = Number(r[iy]); if(y<2564 || y>2568) return;
    const k = key(r[io]); label[k] = label[k] || r[io].trim();
    (sum[y]=sum[y]||{})[k] = (sum[y][k]||0) + num(r[iv]);
    (quarters[y]=quarters[y]||new Set()).add(r[iq]);
  });
  const out = [];
  Object.keys(sum).sort().forEach(y=>{
    const nq = quarters[y].size; let tot = 0;
    Object.entries(sum[y]).forEach(([k,v])=>{ out.push([y, y-543, label[k], (v/nq).toFixed(1)]); tot += v/nq; });
    out.push([y, y-543, 'รวม', tot.toFixed(1)]);
    if(nq!==4) throw new Error(`สสช. ปี ${y} มี ${nq} ไตรมาส`);
  });
  writeCSV('nso_employed_by_occupation_annual_2564-2568.csv', ['ปี_พศ','ปี_คศ','อาชีพ','ผู้มีงานทำ_พันคน_เฉลี่ย4ไตรมาส'], out);
}

/* ---------- D. สสช.: อัตราการว่างงาน จำแนกตามระดับการศึกษา ทั้งประเทศ เฉลี่ยรายปี ---------- */
{
  const rows = parseCSV(fs.readFileSync(path.join(RAW,'nso','nso_unemployment_rate_by_education_sex.csv'),'utf8'));
  const h = rows[0].map(s=>s.trim());
  const [iy,iq,ia,ie,is,iv] = ['year','quarter','area','level_of_edu','sex','value'].map(k=>h.indexOf(k));
  // ชื่อระดับการศึกษาบางปีมีเลขนำหน้า ("08. อุดมศึกษา สายวิชาการ") หรือมีวงเล็บ → ทำให้เป็นแบบเดียวกัน
  const norm = s => s.trim().replace(/^\d+\.\s*/,'').replace(/[()]/g,'').replace(/\s+/g,' ').replace(/^อื่น ๆ$/,'อื่นๆ');
  const acc = {};
  rows.slice(1).forEach(r=>{
    const y = Number(r[iy]); if(y<2564 || y>2568) return;
    if(r[ia].trim()!=='รวม' || r[is].trim()!=='รวม') return;
    const k = norm(r[ie]); const a = ((acc[y]=acc[y]||{})[k]=acc[y][k]||{s:0,n:0,q:new Set()});
    if(a.q.has(r[iq])) return;     // กันแถวซ้ำในไตรมาสเดียวกัน
    a.q.add(r[iq]); a.s += num(r[iv]); a.n++;
  });
  const out = [];
  Object.keys(acc).sort().forEach(y=> Object.entries(acc[y]).forEach(([k,a])=> out.push([y, y-543, k, (a.s/a.n).toFixed(2), a.n])));
  writeCSV('nso_unemployment_rate_by_education_annual_2564-2568.csv', ['ปี_พศ','ปี_คศ','ระดับการศึกษา','อัตราการว่างงาน_ร้อยละ_เฉลี่ย','จำนวนไตรมาสที่มีข้อมูล'], out);
}
