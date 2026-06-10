// @ts-nocheck
import { useState, useCallback } from "react";
import * as XLSX from "xlsx";

const BASE_RATE = 1850;
const RENDEZ_TYPES = new Set([
  "üzemeltetés","bontás","építés","technikai próba",
  "projekttervezés","szállítás","logisztika","hangpróba","fénypróba",
]);
const C = {
  bg:"#0c0f1a", surf:"#141828", card:"#1a1f30", border:"#242a40",
  accent:"#4f7cff", accentDim:"#1e2d6b", ok:"#22c55e", warn:"#f59e0b",
  rev:"#ef4444", info:"#60a5fa", text:"#e2e8f0", muted:"#5a6585",
  thead:"#10132000",
};

function parseMinutes(val) {
  if (!val && val !== 0) return 0;
  if (typeof val === "number") return Math.round(val * 24 * 60);
  if (typeof val === "string") {
    const m = val.match(/^(\d+):(\d+):(\d+)$/);
    if (m) return +m[1]*60 + +m[2] + Math.round(+m[3]/60);
    const m2 = val.match(/^(\d+):(\d+)$/);
    if (m2) return +m2[1]*60 + +m2[2];
  }
  return 0;
}
const hm = (m) => {
  if (!m) return "0:00";
  const h = Math.floor(Math.abs(m)/60);
  const mm = Math.round(Math.abs(m)%60);
  return (m<0?"-":"") + h + ":" + (mm<10?"0":"") + mm;
};
const fmtHu = (n) => Math.round(n).toLocaleString("hu-HU");
const today = () => new Date().toLocaleDateString("hu-HU");

function readXlsx(file) {
  return new Promise((res,rej)=>{
    const r = new FileReader();
    r.onload = e => {
      const wb = XLSX.read(e.target.result, {type:"array",cellDates:true});
      res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:""}));
    };
    r.onerror = rej;
    r.readAsArrayBuffer(file);
  });
}

function fuzzyScore(a, b) {
  const pa = a.toLowerCase(), pb = b.toLowerCase();
  if (pa===pb) return 1;
  const words = pa.split(/\s+/).filter(w => w.length > 3);
  let hits = 0;
  for (const w of words) if (pb.includes(w)) hits++;
  return words.length > 0 ? hits/words.length : 0;
}

function processFiles(planData, actData) {
  const planProjects = {};
  for (const row of planData) {
    const name = row["Project (Function)"]; if (!name) continue;
    if (!planProjects[name]) planProjects[name] = {plannedMin:0};
    planProjects[name].plannedMin += parseMinutes(row["Planned time"]);
  }
  const actProjects = {};
  const noProject = [];
  for (const row of actData) {
    const proj = row["Feladat, projekt"];
    const altipus = (row["Altípus"]||"").toLowerCase().trim();
    const mins = typeof row["Időtartam"]==="number" ? row["Időtartam"] : 0;
    const isR = RENDEZ_TYPES.has(altipus);
    if (!proj) {
      noProject.push({date:row["Dátum"], altipus:row["Altípus"], mins, note:row["Megjegyzés"]});
      continue;
    }
    if (!actProjects[proj]) actProjects[proj] = {totalMin:0, rendezMin:0, byType:{}};
    actProjects[proj].totalMin += mins;
    if (isR) actProjects[proj].rendezMin += mins;
    actProjects[proj].byType[altipus] = (actProjects[proj].byType[altipus]||0) + mins;
  }
  const matched = {};
  for (const [pn,pd] of Object.entries(planProjects)) {
    let bestKey=null, bestScore=0;
    for (const ak of Object.keys(actProjects)) {
      const s = fuzzyScore(
        pn.replace(/\s+0\d{3}\/\d{4}.*/,""),
        ak.replace(/^\d{4}\s*-\s*/,"").replace(/\s+0\d{3}\/\d{4}.*/,"")
      );
      if (s>bestScore && s>0.3) { bestScore=s; bestKey=ak; }
    }
    matched[pn] = {plan:pd, actKey:bestKey, actData:bestKey?actProjects[bestKey]:null};
  }
  const matchedActKeys = new Set(Object.values(matched).map(v=>v.actKey).filter(k=>k!=null));
  const unmatchedAct = Object.entries(actProjects)
    .filter(([k]) => !matchedActKeys.has(k))
    .map(([k,v]) => ({actKey:k, actData:v}));
  return {matched, unmatchedAct, noProject};
}

function absenceStr(r) {
  if (!r.absenceMin) return "\u2014";
  var s = "-" + hm(r.absenceMin);
  if (r.absenceReason) s += " (" + r.absenceReason + ")";
  return s;
}

function Badge({level, children}) {
  const map = {
    ok:[C.ok,"#052e16"], warn:[C.warn,"#451a03"],
    rev:[C.rev,"#450a0a"], info:[C.info,"#0c1a3e"], muted:[C.muted,C.card]
  };
  const [fg,bg] = map[level]||map.muted;
  return <span style={{color:fg,background:bg,border:"1px solid "+fg+"40",borderRadius:4,padding:"2px 8px",fontSize:10,fontWeight:700,letterSpacing:.3,whiteSpace:"nowrap"}}>{children}</span>;
}
function SectionHead({title, sub}) {
  return (
    <div style={{padding:"13px 18px",borderBottom:"1px solid "+C.border,display:"flex",alignItems:"center"}}>
      <span style={{fontWeight:700,fontSize:13}}>{title}</span>
      {sub && <span style={{color:C.muted,fontSize:11,marginLeft:8}}>{sub}</span>}
    </div>
  );
}
function Num({value, onChange, min, step, style}) {
  const s = step!=null ? step : 1;
  const st = style||{};
  return (
    <input type="number" value={value} min={min!=null?min:undefined} step={s}
      onChange={e=>onChange(Number(e.target.value))}
      style={{width:58,background:C.surf,border:"1px solid "+C.border,color:C.text,borderRadius:5,padding:"3px 6px",fontSize:12,textAlign:"right",...st}}/>
  );
}

export default function PremiumKalkulator() {
  const [planFile,setPlanFile] = useState(null);
  const [actFile,setActFile] = useState(null);
  const [kompScore,setKompScore] = useState(60);
  const [month,setMonth] = useState("");
  const [employee,setEmployee] = useState("Szántó Benedek");
  const [result,setResult] = useState(null);
  const [loading,setLoading] = useState(false);
  const [error,setError] = useState(null);
  const [corrections,setCorrections] = useState({});
  const [manualAllocs,setManualAllocs] = useState([]);
  const [mdText,setMdText] = useState("");
  const [mdVisible,setMdVisible] = useState(false);
  const [copied,setCopied] = useState(false);

  const kompRate = BASE_RATE + kompScore*10;
  const setCorr = (key,field,val) => setCorrections(p => ({...p,[key]:{...(p[key]||{}), [field]:val}}));

  const calculate = useCallback(async () => {
    if (!planFile||!actFile) return;
    setLoading(true); setError(null);
    try {
      const [pd,ad] = await Promise.all([readXlsx(planFile), readXlsx(actFile)]);
      const proc = processFiles(pd, ad);
      setResult(proc);
      const init = {};
      for (const k of Object.keys(proc.matched))
        if (!corrections[k]) init[k] = {overtimeH:0, absenceH:0, absenceReason:"", note:""};
      setCorrections(p => ({...init,...p}));
      setManualAllocs(proc.unmatchedAct.map(u => ({actKey:u.actKey, actData:u.actData, allocH:0, note:""})));
    } catch(e) { setError(e.message); }
    setLoading(false);
  }, [planFile, actFile]);

  const calcRows = () => {
    if (!result) return [];
    return Object.entries(result.matched).map(([planName, m]) => {
      const plan=m.plan, actKey=m.actKey, actData=m.actData;
      const corr = corrections[planName]||{};
      const plannedMin = plan.plannedMin;
      const actRendezMin = actData ? actData.rendezMin : 0;
      const actTotalMin = actData ? actData.totalMin : 0;
      const overtimeMin = Math.round((corr.overtimeH||0)*60);
      const absenceMin = Math.round((corr.absenceH||0)*60);
      let basePayable=0, flag="", flagLevel="ok";
      if (actRendezMin===0 && overtimeMin===0) {
        basePayable=0; flag="Nincs leadott rendezvénytechnikai óra"; flagLevel="warn";
      } else {
        basePayable = plannedMin;
        const diff = plannedMin - (actRendezMin+overtimeMin);
        const diffPct = plannedMin>0 ? (diff/plannedMin)*100 : 0;
        if (diffPct>10) { flag="Proj.vez. jóváhagyás ("+diffPct.toFixed(0)+"% eltérés)"; flagLevel="rev"; }
        else if (actRendezMin+overtimeMin>plannedMin) { flag="Tény > Terv → kifizetve: terv"; flagLevel="info"; }
        else flag="OK";
      }
      const payableMin = Math.max(0, basePayable+overtimeMin-absenceMin);
      return {planName,actKey,plannedMin,actTotalMin,actRendezMin,overtimeMin,absenceMin,payableMin,flag,flagLevel,
        absenceReason:corr.absenceReason||"", note:corr.note||""};
    });
  };

  const rows = calcRows();
  const rentmanPayableMin = rows.reduce((s,r) => s+r.payableMin, 0);
  const manualPayableMin = manualAllocs.reduce((s,r) => s+Math.round((r.allocH||0)*60), 0);
  const totalPayableMin = rentmanPayableMin + manualPayableMin;
  const premium = Math.round((totalPayableMin/60)*kompRate);

  const generatePrintHTML = () => {
    const mLabel = month ? month.replace("-"," / ") : "(hónap nincs megadva)";
    const tbl = (headers, body) =>
      "<table><thead><tr>"+headers.map(h=>"<th>"+h+"</th>").join("")+"</tr></thead><tbody>"+body+"</tbody></table>";

    var rRows = "";
    for (var i=0; i<rows.length; i++) {
      var r = rows[i];
      var ovStr = r.overtimeMin ? "+"+hm(r.overtimeMin) : "—";
      var abStr = absenceStr(r);
      var nameCell = r.planName + (r.actKey&&r.actKey!==r.planName ? "<br><small>↳ "+r.actKey+"</small>" : "");
      var payClass = "num bold " + (r.payableMin>0 ? "ok" : "muted");
      var badgeLabel = r.flagLevel==="ok" ? "✓ OK" : r.flagLevel==="rev" ? "⚑ Jóváhagyás" : r.flagLevel==="info" ? "ℹ Tény>Terv" : "⚠ Hiányzó";
      rRows += "<tr><td>"+nameCell+"</td>"
        +"<td class='num'>"+hm(r.plannedMin)+"</td>"
        +"<td class='num'>"+hm(r.actRendezMin)+"</td>"
        +"<td class='num'>"+ovStr+"</td>"
        +"<td class='num'>"+abStr+"</td>"
        +"<td class='"+payClass+"'>"+hm(r.payableMin)+"</td>"
        +"<td><span class='badge "+r.flagLevel+"'>"+badgeLabel+"</span></td></tr>";
    }
    rRows += "<tr class='total'><td>Összesen</td>"
      +"<td class='num'>"+hm(rows.reduce((s,r)=>s+r.plannedMin,0))+"</td>"
      +"<td class='num'>"+hm(rows.reduce((s,r)=>s+r.actRendezMin,0))+"</td>"
      +"<td></td><td></td>"
      +"<td class='num bold ok'>"+hm(rentmanPayableMin)+"</td><td></td></tr>";

    var manSec = "";
    if (manualAllocs.some(a=>a.allocH>0)) {
      var manRows = "";
      manualAllocs.filter(a=>a.allocH>0).forEach(function(a) {
        manRows += "<tr><td>"+a.actKey+"</td>"
          +"<td class='num'>"+hm(a.actData.totalMin)+"</td>"
          +"<td class='num bold ok'>"+hm(Math.round(a.allocH*60))+"</td>"
          +"<td>"+(a.note||"—")+"</td></tr>";
      });
      manRows += "<tr class='total'><td>Összesen</td><td></td><td class='num bold ok'>"+hm(manualPayableMin)+"</td><td></td></tr>";
      manSec = "<h2>2. Manuálisan elfogadott egyéb prémiumórák</h2>"
        +"<p class='sub'>NORMA-ban leadott, Rentman-ben nem tervezett — egyedi jogcímen elszámolva.</p>"
        +tbl(["Projekt (NORMA)","NORMA összes","Elfogadott prémiumóra","Jogcím"], manRows);
    }

    var unmSec = "";
    if (result && result.unmatchedAct && result.unmatchedAct.some(u=>!manualAllocs.find(a=>a.actKey===u.actKey&&a.allocH>0))) {
      var unmRows = "";
      result.unmatchedAct.filter(u=>!manualAllocs.find(a=>a.actKey===u.actKey&&a.allocH>0)).forEach(function(u) {
        var types = Object.entries(u.actData.byType).map(function(e){return e[0]+": "+hm(e[1]);}).join(", ");
        unmRows += "<tr><td>"+u.actKey+"</td><td class='num'>"+hm(u.actData.totalMin)+"</td><td class='num'>"+hm(u.actData.rendezMin)+"</td><td class='small'>"+types+"</td></tr>";
      });
      unmSec = "<h2>3. Nem elszámolt NORMA projektek</h2><p class='sub'>Nem alapoznak meg prémiumkifizetést.</p>"+tbl(["Projekt","Összes","Rendez.techn.","Altípusok"],unmRows);
    }

    var noProjSec = "";
    if (result && result.noProject && result.noProject.length) {
      var npRows = "";
      result.noProject.forEach(function(r) {
        var d = r.date instanceof Date ? r.date.toLocaleDateString("hu-HU") : String(r.date).slice(0,10);
        npRows += "<tr><td>"+d+"</td><td>"+r.altipus+"</td><td class='num'>"+hm(r.mins)+"</td><td>"+(r.note||"—")+"</td></tr>";
      });
      noProjSec = "<h2>4. Projekt nélküli leadott órák</h2><p class='sub'>Nem kapcsolódnak prémiumkifizetéshez.</p>"+tbl(["Dátum","Típus","Idő","Megjegyzés"],npRows);
    }

    var noteSec = "";
    if (rows.some(r=>r.note)) {
      noteSec = "<h3>Megjegyzések</h3><ul>";
      rows.filter(r=>r.note).forEach(function(r){ noteSec += "<li><strong>"+r.planName+":</strong> "+r.note+"</li>"; });
      noteSec += "</ul>";
    }

    var css = "@page{size:A4;margin:18mm 15mm}*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Segoe UI',Arial,sans-serif;font-size:10pt;color:#1a1a2e;line-height:1.4}.header{border-bottom:3px solid #4f7cff;padding-bottom:10px;margin-bottom:16px;display:flex;justify-content:space-between;align-items:flex-end}.header-left h1{font-size:16pt;color:#4f7cff;margin-bottom:2px}.header-left p{font-size:9pt;color:#555}.header-right{text-align:right;font-size:9pt;color:#555}.summary{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:16px}.summary-card{border:1.5px solid #e0e4f0;border-radius:6px;padding:8px 10px}.summary-card .label{font-size:7.5pt;color:#888;text-transform:uppercase;letter-spacing:.4px;margin-bottom:3px}.summary-card .value{font-size:13pt;font-weight:800;color:#1a1a2e}.summary-card .value.accent{color:#4f7cff}.summary-card .value.ok{color:#16a34a}.summary-card .unit{font-size:8pt;color:#888}.summary-card.highlight{border-color:#4f7cff;background:#f0f4ff}h2{font-size:11pt;color:#4f7cff;margin:14px 0 4px;border-bottom:1px solid #dde3f5;padding-bottom:3px}h3{font-size:10pt;color:#333;margin:10px 0 4px}p.sub{font-size:8.5pt;color:#777;margin-bottom:6px;font-style:italic}table{width:100%;border-collapse:collapse;font-size:8.5pt;margin-bottom:10px;page-break-inside:avoid}th{background:#f0f4ff;color:#4f7cff;text-align:right;padding:5px 7px;font-size:7.5pt;text-transform:uppercase;letter-spacing:.3px;border-bottom:1.5px solid #c7d2f0}th:first-child{text-align:left}td{padding:5px 7px;border-bottom:1px solid #eef0f8;vertical-align:top}td:first-child{text-align:left}td.num{text-align:right;font-family:'Courier New',monospace}td.bold{font-weight:700}td.ok{color:#16a34a}td.muted{color:#aaa}td.small{font-size:7.5pt;color:#888}tr.total td{background:#eef1fc!important;font-weight:700;border-top:1.5px solid #c7d2f0}tr:nth-child(even) td{background:#fafbff}.badge{display:inline-block;border-radius:3px;padding:1px 5px;font-size:7.5pt;font-weight:700}.badge.ok{background:#dcfce7;color:#166534}.badge.warn{background:#fef9c3;color:#854d0e}.badge.rev{background:#fee2e2;color:#991b1b}.badge.info{background:#dbeafe;color:#1e40af}small{font-size:7.5pt;color:#888}ul{margin:0 0 8px 16px}li{font-size:9pt;margin-bottom:2px}.calc-box{background:#f7f9ff;border:1px solid #dde3f5;border-radius:5px;padding:10px 14px;font-family:'Courier New',monospace;font-size:8.5pt;line-height:1.8;margin:8px 0 14px}.approval{margin-top:20px;page-break-inside:avoid}.approval table td{padding:10px 7px;font-size:9pt}.footer{margin-top:16px;padding-top:8px;border-top:1px solid #dde3f5;font-size:7.5pt;color:#aaa;font-style:italic}@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}";

    return "<!DOCTYPE html><html lang='hu'><head><meta charset='UTF-8'/><title>Prémium – "+employee+" – "+mLabel+"</title><style>"+css+"</style></head><body>"
      +"<div class='header'><div class='header-left'><h1>Prémium Elszámolás</h1><p>"+employee+" · "+mLabel+" · Outline Central Europe Kft.</p></div><div class='header-right'>Elkészítve: "+today()+"<br>Szabályzat: 2026.03.01.</div></div>"
      +"<div class='summary'>"
        +"<div class='summary-card'><div class='label'>Rentman elszámolt</div><div class='value ok'>"+hm(rentmanPayableMin)+"</div><div class='unit'>óra</div></div>"
        +"<div class='summary-card'><div class='label'>Manuális prémiumóra</div><div class='value'>"+hm(manualPayableMin)+"</div><div class='unit'>óra</div></div>"
        +"<div class='summary-card'><div class='label'>Összes prémiumóra</div><div class='value ok'>"+hm(totalPayableMin)+"</div><div class='unit'>óra</div></div>"
        +"<div class='summary-card highlight'><div class='label'>Fizetendő prémium</div><div class='value accent'>"+fmtHu(premium)+"</div><div class='unit'>Ft bruttó</div></div>"
      +"</div>"
      +"<h2>1. Rentman projektek részletezése</h2><p class='sub'>Számítási alap: RENTMAN tervezett = kifizethető maximum.</p>"
      +tbl(["Projekt","Terv","NORMA rendez.","Túlóra (+)","Kiesés (−)","Elszámolt","Státusz"], rRows)
      +noteSec+manSec+unmSec+noProjSec
      +"<h2>Kalkuláció</h2><div class='calc-box'>"
        +"Alapdíj: "+fmtHu(BASE_RATE)+" Ft/h<br>"
        +"Kompetencia: "+kompScore+"% x 10 Ft = +"+(kompScore*10)+" Ft/h<br>"
        +"Számított óradíj: "+fmtHu(kompRate)+" Ft/h<br><br>"
        +"Rentman: "+hm(rentmanPayableMin)+" | Egyéb: "+hm(manualPayableMin)+" | Összes: "+hm(totalPayableMin)+" = "+(totalPayableMin/60).toFixed(4)+" h<br><br>"
        +"<strong>Prémium = "+(totalPayableMin/60).toFixed(4)+" h x "+fmtHu(kompRate)+" Ft/h = "+fmtHu(premium)+" Ft</strong>"
      +"</div>"
      +"<div class='approval'><h2>Jóváhagyás</h2><table><thead><tr><th style='text-align:left'>Szerepkör</th><th>Dátum</th><th>Aláírás</th></tr></thead><tbody>"
        +"<tr><td>Munkavállaló</td><td class='num'>_________________</td><td class='num'>_________________</td></tr>"
        +"<tr><td>Projektvezető</td><td class='num'>_________________</td><td class='num'>_________________</td></tr>"
        +"<tr><td>Bérszámfejtő</td><td class='num'>_________________</td><td class='num'>_________________</td></tr>"
      +"</tbody></table></div>"
      +"<div class='footer'>Jelen elszámolás az Outline Central Europe Kft. Prémium Szabályzata (2026.03.01.) alapján készült.</div></body></html>";
  };

  const openPDF = () => {
    const html = generatePrintHTML();
    const w = window.open("","_blank","width=900,height=700");
    if (!w) return;
    w.document.write(html);
    w.document.close();
    setTimeout(() => w.print(), 400);
  };

  const generateMD = () => {
    const mLabel = month ? month.replace("-"," / ") : "(hónap nincs megadva)";
    const lines = [];
    lines.push("# Prémium Elszámolás — "+employee);
    lines.push("**Időszak:** "+mLabel+"  \n**Elkészítve:** "+today()+"  \n**Szabályzat:** Outline Central Europe Kft. 2026.03.01.");
    lines.push("\n---\n## Összefoglaló\n");
    lines.push("| Megnevezés | Érték |\n|---|---|");
    lines.push("| Elszámolt rendezvénytechnikai óra | "+hm(rentmanPayableMin)+" |");
    lines.push("| Elszámolt egyéb prémiumóra | "+hm(manualPayableMin)+" |");
    lines.push("| **Összes prémiumóra** | **"+hm(totalPayableMin)+"** |");
    lines.push("| Prémium óradíj | "+fmtHu(kompRate)+" Ft/h |");
    lines.push("| **Fizetendő prémium (bruttó)** | **"+fmtHu(premium)+" Ft** |");
    lines.push("\n---\n## 1. Rentman projektek\n");
    lines.push("| Projekt | Terv | NORMA rendez. | Túlóra (+) | Kiesés (−) | Elszámolt | Státusz |\n|---|---:|---:|---:|---:|---:|---|");
    for (var i=0; i<rows.length; i++) {
      var r = rows[i];
      var ovMD = r.overtimeMin ? "+"+hm(r.overtimeMin) : "—";
      var abMD = r.absenceMin ? ("-"+hm(r.absenceMin)+(r.absenceReason ? " ("+r.absenceReason+")" : "")) : "—";
      lines.push("| "+r.planName+" | "+hm(r.plannedMin)+" | "+hm(r.actRendezMin)+" | "+ovMD+" | "+abMD+" | **"+hm(r.payableMin)+"** | "+r.flag+" |");
    }
    lines.push("| **ÖSSZESEN** | **"+hm(rows.reduce((s,r)=>s+r.plannedMin,0))+"** | **"+hm(rows.reduce((s,r)=>s+r.actRendezMin,0))+"** | | | **"+hm(rentmanPayableMin)+"** | |");
    if (manualAllocs.some(a=>a.allocH>0)) {
      lines.push("\n---\n## 2. Manuális prémiumórák\n");
      lines.push("| Projekt | NORMA összes | Elfogadott | Jogcím |\n|---|---:|---:|---|");
      manualAllocs.filter(a=>a.allocH>0).forEach(function(a) {
        lines.push("| "+a.actKey+" | "+hm(a.actData.totalMin)+" | **"+hm(Math.round(a.allocH*60))+"** | "+(a.note||"—")+" |");
      });
    }
    lines.push("\n---\n## Kalkuláció\n```");
    lines.push("Alapdíj: "+fmtHu(BASE_RATE)+" Ft/h");
    lines.push("Kompetencia: "+kompScore+"% -> +"+(kompScore*10)+" Ft/h -> "+fmtHu(kompRate)+" Ft/h");
    lines.push("Összes prémiumóra: "+hm(totalPayableMin)+" = "+(totalPayableMin/60).toFixed(4)+" h");
    lines.push("Prémium = "+(totalPayableMin/60).toFixed(4)+" h x "+fmtHu(kompRate)+" Ft/h = "+fmtHu(premium)+" Ft\n```");
    lines.push("\n---\n## Jóváhagyás\n\n| | Dátum | Aláírás |\n|---|---|---|\n| Munkavállaló | _________________ | _________________ |\n| Projektvezető | _________________ | _________________ |\n| Bérszámfejtő | _________________ | _________________ |");
    lines.push("\n*Jelen elszámolás az Outline Central Europe Kft. Prémium Szabályzata (2026.03.01.) alapján készült.*");
    return lines.join("\n");
  };

  const showMD = () => { setMdText(generateMD()); setMdVisible(true); setCopied(false); };
  const copyMD = () => {
    navigator.clipboard.writeText(mdText).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    });
  };

  const FileInput = ({label, file, onChange, hint}) => (
    <label style={{display:"flex",flexDirection:"column",gap:5,cursor:"pointer"}}>
      <span style={{fontSize:10,fontWeight:700,color:C.muted,letterSpacing:1,textTransform:"uppercase"}}>{label}</span>
      <div style={{border:"2px dashed "+(file?C.accent:C.border),borderRadius:8,padding:"11px 14px",background:file?C.accentDim+"44":C.surf,display:"flex",alignItems:"center",gap:10,transition:"all .2s"}}>
        <span style={{fontSize:18}}>{file?"✓":"📂"}</span>
        <div>
          <div style={{color:file?C.accent:C.text,fontSize:13,fontWeight:600}}>{file?file.name:"Kattints a feltöltéshez"}</div>
          <div style={{color:C.muted,fontSize:10}}>{hint}</div>
        </div>
      </div>
      <input type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>onChange(e.target.files&&e.target.files[0]||null)}/>
    </label>
  );

  const brd = (v) => "1px solid " + (v||C.border);

  return (
    <div style={{fontFamily:"'Inter',system-ui,sans-serif",background:C.bg,minHeight:"100vh",color:C.text,padding:"22px 18px"}}>
    <div style={{maxWidth:960,margin:"0 auto",display:"flex",flexDirection:"column",gap:14}}>
      <div>
        <div style={{display:"flex",alignItems:"baseline",gap:10}}>
          <h1 style={{margin:0,fontSize:20,fontWeight:800,letterSpacing:-.5}}>Prémium Elszámoló</h1>
          <span style={{fontSize:11,color:C.muted}}>Outline Central Europe Kft. · Bérszámfejtői eszköz</span>
        </div>
        <p style={{margin:"4px 0 0",color:C.muted,fontSize:12}}>Töltsd fel a RENTMAN és NORMA exportokat, adj meg korrekciókat, majd exportáld az elszámolást.</p>
      </div>

      <div style={{background:C.card,border:brd(C.border),borderRadius:12,overflow:"hidden"}}>
        <SectionHead title="Adatok & Beállítások"/>
        <div style={{padding:16,display:"flex",flexDirection:"column",gap:14}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <FileInput label="RENTMAN — Planning crew export" file={planFile} onChange={setPlanFile} hint="Export_Planning_crew_*.xlsx"/>
            <FileInput label="NORMA — Tevékenységek export" file={actFile} onChange={setActFile} hint="Tevékenységek__*.xlsx"/>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:14,alignItems:"end"}}>
            <div>
              <div style={{fontSize:10,fontWeight:700,color:C.muted,letterSpacing:1,textTransform:"uppercase",marginBottom:5}}>Munkavállaló neve</div>
              <input value={employee} onChange={e=>setEmployee(e.target.value)} style={{width:"100%",boxSizing:"border-box",background:C.surf,border:brd(C.border),color:C.text,borderRadius:6,padding:"8px 11px",fontSize:13}}/>
            </div>
            <div>
              <div style={{fontSize:10,fontWeight:700,color:C.muted,letterSpacing:1,textTransform:"uppercase",marginBottom:5}}>Hónap</div>
              <input type="month" value={month} onChange={e=>setMonth(e.target.value)} style={{width:"100%",boxSizing:"border-box",background:C.surf,border:brd(C.border),color:C.text,borderRadius:6,padding:"8px 11px",fontSize:13}}/>
            </div>
            <div>
              <div style={{fontSize:10,fontWeight:700,color:C.muted,letterSpacing:1,textTransform:"uppercase",marginBottom:5}}>
                Kompetencia: <span style={{color:C.accent}}>{kompScore}%</span>
                <span style={{color:C.muted,fontWeight:400}}> → {fmtHu(kompRate)} Ft/h</span>
              </div>
              <input type="range" min={0} max={100} value={kompScore} onChange={e=>setKompScore(+e.target.value)} style={{width:"100%",accentColor:C.accent}}/>
            </div>
          </div>
          <button onClick={calculate} disabled={!planFile||!actFile||loading}
            style={{padding:"11px",borderRadius:8,border:"none",background:(!planFile||!actFile)?C.border:C.accent,color:"#fff",fontWeight:700,fontSize:14,cursor:(!planFile||!actFile)?"not-allowed":"pointer"}}>
            {loading ? "Feldolgozás…" : "Adatok betöltése és elemzés →"}
          </button>
          {error && <div style={{color:C.rev,fontSize:12}}>Hiba: {error}</div>}
        </div>
      </div>

      {result && <>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10}}>
          {[
            {l:"Rentman elszámolt",v:hm(rentmanPayableMin),u:"óra",hi:false,big:false,dim:false},
            {l:"Manuális prémiumóra",v:hm(manualPayableMin),u:"óra",dim:true,hi:false,big:false},
            {l:"Összes prémiumóra",v:hm(totalPayableMin),u:"óra",hi:true,big:false,dim:false},
            {l:"Fizetendő prémium",v:fmtHu(premium),u:"Ft bruttó",big:true,hi:false,dim:false}
          ].map(function(item) {
            return <div key={item.l} style={{background:item.big?C.accentDim:C.card,border:brd(item.big?C.accent:C.border),borderRadius:10,padding:"13px 15px"}}>
              <div style={{fontSize:9,color:C.muted,fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:5}}>{item.l}</div>
              <div style={{fontSize:item.big?20:17,fontWeight:800,color:item.big?C.accent:item.hi?C.ok:item.dim?C.info:C.text}}>{item.v}</div>
              <div style={{fontSize:10,color:C.muted}}>{item.u}</div>
            </div>;
          })}
        </div>

        <div style={{background:C.card,border:brd(C.border),borderRadius:12,overflow:"hidden"}}>
          <SectionHead title="Rentman projektek" sub="Tervezett vs. leadott — korrekciók hozzáadhatók"/>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:C.thead}}>
              {["Projekt","Terv","NORMA rendez.","Túlóra (+) h","Kiesés (−) h","Kiesés oka","Elszámolt","Státusz"].map(h=>(
                <th key={h} style={{padding:"8px 11px",textAlign:h==="Projekt"||h.includes("oka")?"left":"right",color:C.muted,fontSize:10,fontWeight:700,letterSpacing:.4,textTransform:"uppercase",borderBottom:brd(C.border)}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {rows.map((r,i) => (
                <tr key={i} style={{borderBottom:"1px solid "+C.border+"22",background:i%2?"transparent":C.surf+"33"}}>
                  <td style={{padding:"9px 11px",maxWidth:220}}>
                    <div style={{fontWeight:600}}>{r.planName}</div>
                    {r.actKey && r.actKey!==r.planName && <div style={{color:C.muted,fontSize:10,marginTop:2}}>↳ {r.actKey}</div>}
                    {!r.actKey && <div style={{color:C.warn,fontSize:10}}>Nincs NORMA egyezés</div>}
                  </td>
                  <td style={{padding:"9px 11px",textAlign:"right",fontFamily:"monospace"}}>{hm(r.plannedMin)}</td>
                  <td style={{padding:"9px 11px",textAlign:"right",fontFamily:"monospace",color:C.muted}}>{hm(r.actRendezMin)}</td>
                  <td style={{padding:"9px 11px",textAlign:"right"}}>
                    <Num value={corrections[r.planName]&&corrections[r.planName].overtimeH||0} min={0} step={0.5} onChange={v=>setCorr(r.planName,"overtimeH",v)}/>
                  </td>
                  <td style={{padding:"9px 11px",textAlign:"right"}}>
                    <Num value={corrections[r.planName]&&corrections[r.planName].absenceH||0} min={0} step={0.5} onChange={v=>setCorr(r.planName,"absenceH",v)}/>
                  </td>
                  <td style={{padding:"9px 11px"}}>
                    <input value={corrections[r.planName]&&corrections[r.planName].absenceReason||""}
                      onChange={e=>setCorr(r.planName,"absenceReason",e.target.value)}
                      placeholder="betegség, szabadság…"
                      style={{width:"100%",boxSizing:"border-box",background:C.surf,border:brd(C.border),color:C.text,borderRadius:5,padding:"3px 7px",fontSize:11}}/>
                  </td>
                  <td style={{padding:"9px 11px",textAlign:"right",fontFamily:"monospace",fontWeight:700,color:r.payableMin>0?C.ok:C.muted}}>{hm(r.payableMin)}</td>
                  <td style={{padding:"9px 11px",textAlign:"right"}}>
                    <Badge level={r.flagLevel==="ok"?"ok":r.flagLevel==="rev"?"rev":r.flagLevel==="info"?"info":"warn"}>
                      {r.flagLevel==="ok"?"✓ OK":r.flagLevel==="rev"?"⚑ Jóváhagyás":r.flagLevel==="info"?"ℹ Tény>Terv":"⚠ Hiányzó"}
                    </Badge>
                  </td>
                </tr>
              ))}
              <tr style={{background:C.thead,borderTop:"2px solid "+C.border}}>
                <td style={{padding:"9px 11px",fontWeight:700}}>Összesen</td>
                <td style={{padding:"9px 11px",textAlign:"right",fontFamily:"monospace",fontWeight:700}}>{hm(rows.reduce((s,r)=>s+r.plannedMin,0))}</td>
                <td style={{padding:"9px 11px",textAlign:"right",fontFamily:"monospace",color:C.muted}}>{hm(rows.reduce((s,r)=>s+r.actRendezMin,0))}</td>
                <td/><td/><td/>
                <td style={{padding:"9px 11px",textAlign:"right",fontFamily:"monospace",fontWeight:800,color:C.ok}}>{hm(rentmanPayableMin)}</td>
                <td/>
              </tr>
            </tbody>
          </table>
          <div style={{padding:"10px 14px",borderTop:brd(C.border),display:"flex",flexDirection:"column",gap:6}}>
            <div style={{fontSize:10,color:C.muted,fontWeight:700,letterSpacing:.8,textTransform:"uppercase",marginBottom:2}}>Projekt-szintű megjegyzések</div>
            {rows.map((r,i) => (
              <div key={i} style={{display:"flex",alignItems:"center",gap:8}}>
                <span style={{fontSize:11,color:C.muted,minWidth:200,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{r.planName}:</span>
                <input value={corrections[r.planName]&&corrections[r.planName].note||""}
                  onChange={e=>setCorr(r.planName,"note",e.target.value)}
                  placeholder="megjegyzés az elszámoláshoz…"
                  style={{flex:1,background:C.surf,border:brd(C.border),color:C.text,borderRadius:5,padding:"3px 8px",fontSize:11}}/>
              </div>
            ))}
          </div>
        </div>

        {manualAllocs.length>0 && (
          <div style={{background:C.card,border:brd(C.border),borderRadius:12,overflow:"hidden"}}>
            <SectionHead title="Manuális prémiumóra-hozzárendelés" sub="Nem Rentman-tervezett projektek — egyedi jogcím alapján"/>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr style={{background:C.thead}}>
                {["NORMA projekt","NORMA összes","Rendez.techn.","Altípusok","Elfogadott h","Jogcím / Megjegyzés"].map(h=>(
                  <th key={h} style={{padding:"8px 11px",textAlign:h.includes("projekt")||h.includes("Altípus")||h.includes("Jogcím")?"left":"right",color:C.muted,fontSize:10,fontWeight:700,letterSpacing:.4,borderBottom:brd(C.border)}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {manualAllocs.map((a,i) => (
                  <tr key={i} style={{borderBottom:"1px solid "+C.border+"22",background:i%2?"transparent":C.surf+"33"}}>
                    <td style={{padding:"9px 11px",fontWeight:600,maxWidth:220}}>{a.actKey}</td>
                    <td style={{padding:"9px 11px",textAlign:"right",fontFamily:"monospace",color:C.muted}}>{hm(a.actData.totalMin)}</td>
                    <td style={{padding:"9px 11px",textAlign:"right",fontFamily:"monospace",color:C.muted}}>{hm(a.actData.rendezMin)}</td>
                    <td style={{padding:"9px 11px",fontSize:10,color:C.muted,maxWidth:180}}>
                      {Object.entries(a.actData.byType).map(([t,m])=>t+": "+hm(m)).join(" · ")}
                    </td>
                    <td style={{padding:"9px 11px",textAlign:"right"}}>
                      <Num value={a.allocH||0} min={0} step={0.5}
                        onChange={v=>setManualAllocs(p=>p.map((x,j)=>j===i?{...x,allocH:v}:x))}
                        style={{width:64,color:a.allocH>0?C.ok:C.text}}/>
                    </td>
                    <td style={{padding:"9px 11px"}}>
                      <input value={a.note||""} onChange={e=>setManualAllocs(p=>p.map((x,j)=>j===i?{...x,note:e.target.value}:x))}
                        placeholder="egyedi megbízás, tulajdonosi döntés…"
                        style={{width:"100%",boxSizing:"border-box",background:C.surf,border:brd(C.border),color:C.text,borderRadius:5,padding:"3px 8px",fontSize:11}}/>
                    </td>
                  </tr>
                ))}
                <tr style={{background:C.thead,borderTop:"2px solid "+C.border}}>
                  <td style={{padding:"9px 11px",fontWeight:700}}>Összesen</td>
                  <td colSpan={3}/>
                  <td style={{padding:"9px 11px",textAlign:"right",fontFamily:"monospace",fontWeight:800,color:manualPayableMin>0?C.ok:C.muted}}>{hm(manualPayableMin)}</td>
                  <td/>
                </tr>
              </tbody>
            </table>
          </div>
        )}

        {result.noProject.length>0 && (
          <div style={{background:C.card,border:brd(C.border),borderRadius:12,overflow:"hidden"}}>
            <SectionHead title="Projekt nélküli NORMA sorok" sub="Nem prémiumalapozók"/>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr style={{background:C.thead}}>
                {["Dátum","Típus","Idő","Megjegyzés"].map(h=>(
                  <th key={h} style={{padding:"7px 11px",textAlign:"left",color:C.muted,fontSize:10,fontWeight:700,borderBottom:brd(C.border)}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {result.noProject.map((r,i) => (
                  <tr key={i} style={{borderBottom:"1px solid "+C.border+"22"}}>
                    <td style={{padding:"8px 11px",fontFamily:"monospace",color:C.muted,fontSize:11}}>
                      {r.date instanceof Date ? r.date.toLocaleDateString("hu-HU") : String(r.date).slice(0,10)}
                    </td>
                    <td style={{padding:"8px 11px"}}>{r.altipus}</td>
                    <td style={{padding:"8px 11px",fontFamily:"monospace"}}>{hm(r.mins)}</td>
                    <td style={{padding:"8px 11px",color:C.muted}}>{r.note||"—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        <div style={{background:C.card,border:"1px solid "+C.accent+"44",borderRadius:12,overflow:"hidden"}}>
          <div style={{padding:"14px 18px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:14}}>
            <div>
              <div style={{fontWeight:700,fontSize:14}}>Elszámolás exportálása</div>
              <div style={{fontSize:11,color:C.muted,marginTop:3}}>PDF: új ablakban nyílik, böngészőből Mentés PDF-be (A4). Markdown: vágólapra másolható.</div>
            </div>
            <div style={{display:"flex",gap:8,flexShrink:0}}>
              <button onClick={openPDF} style={{padding:"9px 20px",borderRadius:7,border:"none",background:C.accent,color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>🖨️ Mentés PDF-ként</button>
              <button onClick={showMD} style={{padding:"9px 16px",borderRadius:7,border:brd(C.border),background:C.surf,color:C.muted,fontWeight:600,fontSize:12,cursor:"pointer"}}>
                {mdVisible ? "↻ MD frissítés" : "MD forrás"}
              </button>
              {mdVisible && <button onClick={copyMD} style={{padding:"9px 16px",borderRadius:7,border:brd(copied?C.ok:C.border),background:copied?"#052e16":C.surf,color:copied?C.ok:C.text,fontWeight:700,fontSize:13,cursor:"pointer",transition:"all .2s"}}>
                {copied ? "✓ Másolva!" : "📋 Másolás"}
              </button>}
            </div>
          </div>
          {mdVisible && (
            <div style={{borderTop:brd(C.border),padding:14}}>
              <div style={{fontSize:10,color:C.muted,fontWeight:700,letterSpacing:.8,textTransform:"uppercase",marginBottom:6}}>Markdown — jelöld ki (Ctrl+A) majd másold (Ctrl+C)</div>
              <textarea readOnly value={mdText}
                style={{width:"100%",boxSizing:"border-box",height:280,background:C.bg,border:brd(C.border),color:"#94a3b8",fontFamily:"'Courier New',monospace",fontSize:11,lineHeight:1.6,borderRadius:7,padding:12,resize:"vertical"}}
                onFocus={e=>e.target.select()}/>
            </div>
          )}
        </div>
      </>}
    </div>
    </div>
  );
}
