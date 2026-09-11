applyAppearance();


const demoDoctors=["Dr Daoudi","Dr Daghouane","Dr El Hariri","Dr Ennakhassi","Dr Benarina","Dr Faouzi","Dr Habbouni","Dr Mourabiti","Dr El Mrabet","Dr Belkhadir","Dr El Azher"];
let state=JSON.parse(localStorage.getItem("gardeMed")||"null")||{doctors:[],absences:[],fixed:[],planning:{},planningHistory:{},generationOffset:0,settings:{title:"Tox-Garde",subtitle:"Simple. Équitable. Intelligent",appBg:"#f5f7fb",headerBg:"#3f83c5",cardBg:"#ffffff",bgImage:""}};
if(!state.recoveryLedger) state.recoveryLedger={};
if(!state.confirmedPlanningKeys) state.confirmedPlanningKeys={};
state.settings=Object.assign({title:"Tox-Garde",subtitle:"Simple. Équitable. Intelligent",appBg:"#f5f7fb",headerBg:"#3f83c5",cardBg:"#ffffff",bgImage:""},state.settings||{});
if(!state.planningHistory) state.planningHistory={};
if(!state.settings) state.settings={title:"Tox-Garde",subtitle:"Simple. Équitable. Intelligent",appBg:"#f5f7fb",headerBg:"#3f83c5",cardBg:"#ffffff",bgImage:""};
if(!state.generationOffset) state.generationOffset=0;
if(!state.doctors.length) state.doctors=demoDoctors.map(name=>({name,active:true,shift:"BOTH",weekend:true}));
state.doctors.forEach(d=>{if(!d.shift)d.shift="BOTH";});
const months=["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"];
const monthShort=["Jan","Fév","Mar","Avr","Mai","Jui","Jul","Aoû","Sep","Oct","Nov","Déc"];
months.forEach((m,i)=>document.getElementById("month").innerHTML+=`<option value="${i}">${m}</option>`);
document.getElementById("month").value=8;
if(document.getElementById("weekStart"))document.getElementById("weekStart").value="2026-09-01";
let activePlanningView={type:"current",key:null};
function save(){localStorage.setItem("gardeMed",JSON.stringify(state));renderAll()}
function openTab(id){document.querySelectorAll(".panel").forEach(x=>x.classList.remove("active"));document.getElementById(id).classList.add("active");document.querySelectorAll("nav button").forEach(x=>x.classList.toggle("active",x.dataset.tab===id))}
document.querySelectorAll("nav button").forEach(b=>b.onclick=()=>openTab(b.dataset.tab));
function addDoctor(){let n=document.getElementById("newDoctor").value.trim();if(!n)return;if(state.doctors.some(d=>d.name.toLowerCase()===n.toLowerCase()))return alert("Ce médecin existe déjà.");state.doctors.push({name:n,active:true,shift:"BOTH",weekend:true});document.getElementById("newDoctor").value="";save()}
function loadDemo(){state.doctors=demoDoctors.map(name=>({name,active:true,shift:"BOTH",weekend:true}));save()}
function importTeamFile(ev){
  const f=ev.target.files?.[0]; if(!f)return;
  const r=new FileReader();
  r.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:"array",cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:"",raw:true});
      const norm=x=>String(x??"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim();
      const findKey=(keys,names)=>keys.find(k=>names.some(n=>norm(k).includes(norm(n))));
      const doctors=[]; let skipped=0;
      rows.forEach(row=>{
        const keys=Object.keys(row);
        const nameKey=findKey(keys,["medecin","médecin","nom","name"]);
        if(!nameKey){skipped++;return;}
        const name=String(row[nameKey]??"").trim();
        if(!name){skipped++;return;}
        const activeKey=findKey(keys,["actif","active","disponible"]);
        const shiftKey=findKey(keys,["type de garde","type garde","garde","shift","service"]);
        const weekendKey=findKey(keys,["habilité week-end","habilite week-end","week-end","weekend","garde week-end","garde weekend"]);
        let active=true;
        if(activeKey){const v=norm(row[activeKey]); if(["non","no","0","false","inactif","inactive"].includes(v))active=false;}
        let raw=shiftKey?norm(row[shiftKey]):"both";
        let shift="BOTH";
        if(["j","jour","day"].includes(raw))shift="J";
        else if(["n","nuit","night"].includes(raw))shift="N";
        else if(["j+n","j + n","jn","jour+nuit","jour et nuit","both","les deux","deux"].includes(raw))shift="BOTH";
        else if(raw) {
          if(raw.includes("nuit")&&!raw.includes("jour"))shift="N";
          else if(raw.includes("jour")&&!raw.includes("nuit"))shift="J";
        }
        let weekend=true;
        if(weekendKey){
          const w=norm(row[weekendKey]);
          if(["non","no","0","false","inactif","inactive"].includes(w)) weekend=false;
        }
        if(doctors.some(d=>d.name.toLowerCase()===name.toLowerCase())){skipped++;return;}
        doctors.push({name,active,shift,weekend});
      });
      if(!doctors.length){alert("Aucun médecin valide trouvé dans le fichier.");return;}
      if(!confirm(`Importer ${doctors.length} médecin(s) et remplacer l'équipe actuelle ?`))return;
      state.doctors=doctors; save();
      alert(`${doctors.length} médecin(s) importé(s)${skipped?` ; ${skipped} ligne(s) ignorée(s)`:""}.`);
    }catch(err){alert("Erreur d'import de l'équipe : "+err.message);}
    ev.target.value="";
  };
  r.readAsArrayBuffer(f);
}
function downloadTeamTemplate(){
  const rows=[["Médecin","Actif","Type de garde","Habilité week-end"],["Dr Exemple 1","Oui","J","Oui"],["Dr Exemple 2","Oui","N","Oui"],["Dr Exemple 3","Oui","J+N","Oui"]];
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),"Equipe"); XLSX.writeFile(wb,"modele_equipe_tox-garde.xlsx");
}
function delDoctor(i){if(confirm("Supprimer ce médecin ?")){state.doctors.splice(i,1);save()}}
function addAbsence(){
  let doctor=document.getElementById("absDoctor").value,s=document.getElementById("absStart").value,e=document.getElementById("absEnd").value,type=document.getElementById("absType").value;
  if(!doctor||!s||!e||e<s)return alert("Vérifie les dates.");
  state.absences.push({doctor,start:s,end:e,type});
  save();
  // Une absence ajoutée après validation ne modifie jamais le planning validé.
  // Si elle chevauche une garde déjà validée, signaler le conflit sans masquer la garde.
  const per=getPeriodDates(), h=state.planningHistory?.[per.key];
  if(h?.confirmed && h.planning){
    const conflicts=[];
    for(const x of Object.values(h.planning).flat()){
      if(x?.doctor!==doctor || !["J","N"].includes(x.type)) continue;
      const d=x.date;
      if(d>=s && d<=e) conflicts.push(`${d} ${x.type}`);
    }
    if(conflicts.length){
      const msg=document.getElementById("genMessage");
      if(msg) msg.innerHTML=`<div class="warn"><b>Attention : absence ajoutée après validation.</b><br>Le planning validé reste inchangé et ses gardes restent visibles.<br><small>Conflit détecté : ${escapeHtml(conflicts.join(" • "))}. Régénérez puis validez à nouveau si cette absence doit être prise en compte.</small></div>`;
    }
  }
}
function delAbs(i){state.absences.splice(i,1);save()}
function isWeekend(d){let x=d.getDay();return x===0||x===6}
function iso(d){
  if(typeof d==='string'){
    const m=d.match(/^(\d{4}-\d{2}-\d{2})/);
    if(m)return m[1];
  }
  if(d instanceof Date && !isNaN(d)){
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
  }
  return "";
}
function dateFromISO(s){let [y,m,d]=s.split("-").map(Number);return new Date(y,m-1,d)}
function normalizeDateKey(v){
  if(v instanceof Date && !Number.isNaN(v.getTime())) return iso(v);
  const s=String(v??"").trim();
  const m=s.match(/^(\d{4}-\d{2}-\d{2})/);
  if(m)return m[1];
  return "";
}
function absOn(doc,date){
  const dk=normalizeDateKey(date);
  const doctorName=typeof doc==='string' ? doc.trim() : String(doc?.name??'').trim();
  if(!dk || !doctorName)return false;
  const normName=v=>String(v??'').trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
  const target=normName(doctorName);
  return (state.absences||[]).some(a=>{
    if(!a || normName(a.doctor)!==target)return false;
    const start=normalizeDateKey(a.start), end=normalizeDateKey(a.end);
    if(!start||!end)return false;
    return dk>=start && dk<=end;
  });
}
function fixedOn(doc,date){
  const doctorName=typeof doc==='string' ? doc.trim() : String(doc?.name??'').trim();
  const dk=normalizeDateKey(date);
  if(!doctorName||!dk)return undefined;
  const normName=v=>String(v??'').trim().toLowerCase();
  return (state.fixed||[]).find(x=>normName(x.doctor)===normName(doctorName)&&normalizeDateKey(x.date)===dk);
}
function fixedAny(date){return state.fixed.filter(x=>x.date===iso(date))}
function addFixed(){
  let doctor=document.getElementById("fixedDoctor").value,date=document.getElementById("fixedDate").value,type=document.getElementById("fixedType").value;
  if(!doctor||!date)return alert("Sélectionne le médecin et la date.");
  if(state.fixed.some(x=>x.doctor===doctor&&x.date===date))return alert("Ce médecin a déjà une garde G/F à cette date.");
  const problem=validateFixedGuard(doctor,date);
  if(problem){ alert(problem+"\n\nLa garde G/F n'a pas été enregistrée afin de respecter les contraintes de congé et de repos."); return; }
  state.fixed.push({doctor,date,type,createdAt:Date.now()});save();
}
function parseExcelDate(v){
  if(v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,"0")}-${String(v.getDate()).padStart(2,"0")}`;
  if(typeof v==="number"){
    const d=XLSX.SSF.parse_date_code(v);
    if(d) return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
  }
  let s=String(v??"").trim();
  if(!s) return "";
  // Accepte AAAA-MM-JJ, JJ/MM/AAAA, JJ-MM-AAAA et dates avec heure.
  let m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s.*)?$/);
  if(m) return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
  m=s.match(/^(\d{4})[\-](\d{1,2})[\-](\d{1,2})(?:[T\s].*)?$/);
  if(m) return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  return "";
}
function importExcel(ev){
  const f=ev.target.files?.[0]; if(!f) return;
  const r=new FileReader();
  r.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:"array",cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:"",raw:true});
      let added=0, skipped=0;
      rows.forEach(row=>{
        const keys=Object.keys(row);
        const norm=x=>String(x).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim();
        const get=names=>{const k=keys.find(x=>names.some(n=>norm(x).includes(norm(n))));return k!==undefined?row[k]:""};
        const doctor=String(get(["medecin","nom"])).trim();
        const type=String(get(["type","garde"])).trim().toUpperCase();
        const s=parseExcelDate(get(["date debut","date début","date","jour"]));
        const en=parseExcelDate(get(["date fin","fin"]));
        if(!doctor||!s||!state.doctors.some(d=>d.name.toLowerCase()===doctor.toLowerCase())){skipped++;return;}
        const a=dateFromISO(s), b=dateFromISO(en||s);
        const t=type==="F"?"F":"G";
        for(let d=new Date(a);d<=b;d.setDate(d.getDate()+1)){
          const date=iso(d);
          if(state.fixed.some(x=>x.doctor===doctor&&x.date===date)){skipped++;continue;}
          if(validateFixedGuard(doctor,date)){skipped++;continue;}
          state.fixed.push({doctor,date,type:t,createdAt:Date.now()}); added++;
        }
      });
      save();
      alert(`${added} garde(s) G/F importée(s)${skipped?` ; ${skipped} ligne(s) ignorée(s)`:""}.`);
      ev.target.value="";
    }catch(err){alert("Erreur d'import Excel : "+err.message);}
  };
  r.readAsArrayBuffer(f);
}
function clearFixed(){if(confirm("Effacer toutes les saisies : congés, indisponibilités, récupérations et gardes G/F ?")){state.fixed=[];state.absences=[];save()}}
function downloadTemplate(){let rows=[["Médecin","Date","Type"],["Dr Daoudi","2026-09-05","G"],["Dr Daghouane","2026-09-12","G"],["Dr El Hariri","2026-09-14","F"]];let wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),"Gardes");XLSX.writeFile(wb,"modele_gardes_GF.xlsx")}
function daysInMonth(y,m){return new Date(y,m+1,0).getDate()}
function getPeriodDates(){
  const type=document.getElementById("periodType")?.value||"month";
  if(type==="week"||type==="weeks"){
    let s=document.getElementById("weekStart").value;
    if(!s){s=`${document.getElementById("year").value}-${String(+document.getElementById("month").value+1).padStart(2,"0")}-01`;}
    let start=dateFromISO(s);
    let count=type==="weeks"?Math.max(1,Math.min(52,+document.getElementById("weekCount").value||1)):1;
    let end=new Date(start); end.setDate(end.getDate()+count*7-1);
    return {type,start,end,weeks:count,key:`weeks_${iso(start)}_${count}`};
  }
  const y=+document.getElementById("year").value,m=+document.getElementById("month").value;
  return {type,start:new Date(y,m,1),end:new Date(y,m,daysInMonth(y,m)),key:`month_${y}_${String(m+1).padStart(2,"0")}`};
}
function datesBetween(start,end){let a=[];for(let d=new Date(start);d<=end;d.setDate(d.getDate()+1))a.push(new Date(d));return a}
function togglePeriodInputs(){const t=document.getElementById("periodType").value;const w=t==="week"||t==="weeks";document.getElementById("weekWrap").style.display=w?"block":"none";document.getElementById("weekCountWrap").style.display=t==="weeks"?"block":"none";document.getElementById("monthWrap").style.display=w?"none":"block";renderPlanning()}
function lastGuardDate(doc,date,planning){
  let last=null;
  for(const x of Object.values(planning).flat()){
    if(x.doctor===doc.name&&x.date){const dt=dateFromISO(x.date);if(dt<date&&(!last||dt>last))last=dt;}
  }
  for(const x of state.fixed.filter(z=>z.doctor===doc.name)){
    const dt=dateFromISO(x.date); if(dt<date&&(!last||dt>last))last=dt;
  }
  return last;
}
function specialFixedBlock(doc,date){
  // Règle spécifique : un médecin qui assure tout un week-end (samedi+dimanche)
  // ou deux jours fériés consécutifs ne peut recevoir aucune autre garde
  // dans les 7 jours avant et les 7 jours après ce bloc.
  const fs=state.fixed.filter(x=>x.doctor===doc.name);
  const days=new Set(fs.map(x=>x.date));
  const dateKey=iso(date);
  for(const x of fs){
    if(x.type==='G'){
      const d=dateFromISO(x.date);
      if(d.getDay()===6){
        const sun=new Date(d); sun.setDate(sun.getDate()+1);
        const sunKey=iso(sun);
        if(days.has(sunKey) && fs.some(y=>y.doctor===doc.name&&y.date===sunKey&&y.type==='G')){
          if(dateKey===x.date || dateKey===sunKey) continue;
          const diff=(date-d)/86400000;
          if(diff>=-7 && diff<=8) return true;
        }
      }
    }
    if(x.type==='F'){
      const d=dateFromISO(x.date);
      const next=new Date(d); next.setDate(next.getDate()+1);
      const nextKey=iso(next);
      if(fs.some(y=>y.doctor===doc.name&&y.date===nextKey&&y.type==='F')){
        if(dateKey===x.date || dateKey===nextKey) continue;
        const diff=(date-d)/86400000;
        if(diff>=-7 && diff<=8) return true;
      }
    }
  }
  return false;
}
function canAssign(doc,date,type,planning){
  if(!doc.active)return false;
  if(doc.shift==="J"&&type!=="J")return false;
  if(doc.shift==="N"&&type!=="N")return false;
  // Aucune garde pendant un congé ou une indisponibilité.
  if(absOn(doc,date))return false;
  if(fixedOn(doc,date))return false;
  if(specialFixedBlock(doc,date))return false;
  // Un congé / une indisponibilité interdit uniquement les dates couvertes par
  // cette période. Il ne faut PAS créer artificiellement 48 h supplémentaires
  // après la fin d'un congé ou d'une indisponibilité : cela pouvait éliminer
  // inutilement les seuls candidats capables de compléter un J/N.
  let prev=new Date(date); prev.setDate(prev.getDate()-1);
  if(planning[iso(prev)]?.some(x=>x.doctor===doc.name) || fixedOn(doc,prev))return false;
  const last=lastGuardDate(doc,date,planning);
  if(last && (date-last)/36e5<48)return false;
  return true;
}
function validateFixedGuard(doctor,date){
  const doc=state.doctors.find(d=>d.name===doctor);
  if(doc && doc.weekend===false) return `Le médecin ${doctor} n'est pas habilité aux gardes week-end G/F.`;
  const a=state.absences.find(x=>x.doctor===doctor&&date>=dateFromISO(x.start)&&date<=dateFromISO(x.end));
  if(a) return `Le médecin est en ${a.type.toLowerCase()} le ${date}.`;
  // La fin d'un congé/indisponibilité ne crée pas de blocage automatique
  // supplémentaire : seule la période d'absence elle-même est interdite.
  if(specialFixedBlock(state.doctors.find(d=>d.name===doctor)||{name:doctor},dateFromISO(date))) return `Le médecin est protégé par la règle de repos de 7 jours avant/après un week-end complet ou deux jours fériés consécutifs.`;
  return "";
}
function daysInclusive(start,end){
  const a=dateFromISO(start),b=dateFromISO(end);
  if(!a||!b||b<a)return 0;
  return Math.floor((b-a)/86400000)+1;
}
function equityEntryIsAfterReset(ts){
  const r=Number(state.equityResetAt||0);
  if(!r) return true;
  return Number(ts||0) >= r;
}

function counts(){
  let c={}; state.doctors.forEach(d=>{
    const r=state.recoveryLedger[d.name]||{solde:0,totalAcquise:0,totalPrise:0};
    const solde=Number(r.solde||0);
    const acquise=Number(r.totalAcquise ?? 0);
    const prise=Number(r.totalPrise ?? 0);
    c[d.name]={J:0,N:0,G:0,F:0,total:0,recupAcquise:acquise,recupPrise:prise,recupSolde:solde,last:null};
  });
  // Les gardes sont dédoublonnées par médecin/date/type afin que plusieurs versions
  // d'un même planning ne gonflent pas artificiellement l'équité.
  const all=[];
  // IMPORTANT : seuls les plannings CONFIRMÉS alimentent les compteurs.
  // Une génération/régénération non confirmée est un brouillon et ne compte ni
  // pour J/N/G/F ni pour les récupérations. Le planning courant n'est donc
  // jamais ajouté ici tant qu'il n'est pas confirmé.
  Object.values(state.planningHistory||{}).forEach(h=>{
    if(!h || h.confirmed!==true) return;
    if(h.includeInEquity===false) return;
    if(!equityEntryIsAfterReset(h?.generatedAtISO))return;
    const hp=h?.planning||h;
    all.push(...Object.values(hp||{}).flat());
    if(h?.fixed){
      const hs=h.period?.start?dateFromISO(h.period.start):null;
      const he=h.period?.end?dateFromISO(h.period.end):null;
      all.push(...h.fixed.filter(x=>{
        if(!x?.date)return false;
        const d=dateFromISO(x.date);
        return !hs||!he||(d>=hs&&d<=he);
      }));
    }
  });
  const seen=new Set();
  for(const x of all){
    if(!x || !x.doctor || !x.date || !x.type)continue;
    const key=`${x.doctor}|${x.date}|${x.type}`;
    if(seen.has(key))continue;
    seen.add(key);
    if(!c[x.doctor])continue;
    c[x.doctor][x.type]=(c[x.doctor][x.type]||0)+1;
    // IMPORTANT : les récupérations ne sont créditées qu'après confirmation du planning.
    // Une génération/régénération seule ne modifie donc jamais le score de récupération.
    if(["N","G","F"].includes(x.type) && isConfirmedCurrentPlanningEntry(x)) c[x.doctor].recupAcquise+=2;
    c[x.doctor].total++;
  }
  // Les deux cellules samedi+dimanche d'un même G/F sont UNE garde week-end.
  applyWeekendUnits(c,all);
  // Le solde enregistré est la seule valeur de référence avant toute nouvelle confirmation.
  // Aucun brouillon courant n'est ajouté ici.
  state.doctors.forEach(d=>{
    const r=state.recoveryLedger[d.name]||{};
    c[d.name].recupAcquise=Number(r.totalAcquise ?? 0);
    c[d.name].recupPrise=Number(r.totalPrise ?? 0);
    // Le solde affiché est TOUJOURS dérivé des données confirmées.
    // On ne fait jamais confiance à un ancien champ `solde` éventuellement incohérent.
    c[d.name].recupSolde=c[d.name].recupAcquise-c[d.name].recupPrise;
  });
  return c;
}
function guardPoints(type){
  return ({J:1,N:2,G:3,F:3}[type]||0);
}
function isEligibleForType(doc,type){
  if(!doc || !doc.active) return false;
  const s=String(doc.shift||"BOTH").toUpperCase();
  if(type==="J") return s==="J" || s==="BOTH";
  if(type==="N") return s==="N" || s==="BOTH";
  // G/F are weekend guards. They are normally entered manually/fixed;
  // if a doctor has an explicit weekend=false flag, he is not eligible.
  return doc.weekend !== false;
}
// Charge pondérée : un week-end G/F complet (samedi + dimanche)
// constitue UNE garde week-end = 4 points, même si deux cellules G/F sont
// affichées dans le planning. Les compteurs J/N/G/F restent inchangés pour
// l'affichage ; les points utilisent les unités de garde.
function weekendUnits(c,type){
  if(!c) return 0;
  const meta=type==='G' ? c._weekendUnitsG : c._weekendUnitsF;
  if(meta!==undefined) return Number(meta)||0;
  // Fallback pour les anciens états : deux cellules correspondent à un week-end.
  return Math.ceil(Number(c[type]||0)/2);
}
function computeWeekendUnitsFromEntries(entries){
  const out={};
  const seen=new Set();
  const byKey=new Map();
  for(const x of entries||[]){
    if(!x?.doctor || !x?.date || !['G','F'].includes(x.type)) continue;
    const d=dateFromISO(x.date); if(!d) continue;
    const key=`${x.doctor}|${x.type}|${x.date}`;
    if(seen.has(key)) continue; seen.add(key);
    byKey.set(key,x);
  }
  for(const x of byKey.values()){
    const d=dateFromISO(x.date);
    const y=new Date(d);
    if(d.getDay()===6){
      y.setDate(y.getDate()+1);
      const sunday=`${x.doctor}|${x.type}|${iso(y)}`;
      if(byKey.has(sunday)){
        const u=out[x.doctor]||(out[x.doctor]={G:0,F:0});
        u[x.type]++;
      }
    }else if(d.getDay()!==0){
      // A standalone G/F entry is still one guard unit.
      const u=out[x.doctor]||(out[x.doctor]={G:0,F:0});
      u[x.type]++;
    }
  }
  return out;
}
function applyWeekendUnits(countObj,entries){
  const units=computeWeekendUnitsFromEntries(entries);
  for(const name of Object.keys(countObj||{})){
    const u=units[name]||{};
    countObj[name]._weekendUnitsG=Number(u.G||0);
    countObj[name]._weekendUnitsF=Number(u.F||0);
  }
  return countObj;
}
function workload(c){
  return (c?.J||0)*1 + (c?.N||0)*2 + weekendUnits(c,'G')*3 + weekendUnits(c,'F')*3;
}
function separationPenalty(doc,type,date,planning){
  // Préférence d'équilibrage : lorsqu'un médecin est habilité J+N, on
  // cherche à espacer au maximum ses gardes J et N. Il s'agit d'une
  // préférence souple : les contraintes obligatoires (couverture, repos,
  // congés, indisponibilités et équité de charge) restent prioritaires.
  if(!doc || !date || !planning) return 0;
  const opposite=type==='J' ? 'N' : 'J';
  const target=doc.name;
  const d0=date instanceof Date ? date : dateFromISO(date);
  if(!d0) return 0;
  let minGap=Infinity;
  for(const [key,entries] of Object.entries(planning||{})){
    if(!Array.isArray(entries)) continue;
    if(!entries.some(x=>x?.doctor===target && x.type===opposite)) continue;
    const d=dateFromISO(key);
    if(!d) continue;
    const gap=Math.abs((d0-d)/86400000);
    if(gap>0 && gap<minGap) minGap=gap;
  }
  // Les gardes fixes J/N déjà saisies doivent également compter dans
  // l'espacement recherché.
  for(const x of (state.fixed||[])){
    if(x?.doctor!==target || x.type!==opposite) continue;
    const d=dateFromISO(x.date);
    if(!d) continue;
    const gap=Math.abs((d0-d)/86400000);
    if(gap>0 && gap<minGap) minGap=gap;
  }
  if(!Number.isFinite(minGap)) return 0;
  // Forte pénalisation des alternances rapprochées, décroissante avec la
  // distance. Au-delà de 7 jours, aucune pénalité n'est appliquée.
  if(minGap<=1) return 30000;
  if(minGap<=2) return 15000;
  if(minGap<=3) return 8000;
  if(minGap<=4) return 4000;
  if(minGap<=5) return 2000;
  if(minGap<=6) return 1000;
  if(minGap<=7) return 400;
  return 0;
}

function candidateScore(doc,type,c,monthly,date,planning){
  const x=c[doc.name]||{J:0,N:0,G:0,F:0};
  const m=(monthly&&monthly[doc.name])||{J:0,N:0,G:0,F:0};
  const total=workload(x);
  const month=workload(m);

  // Pour N : un médecin avec la contrainte N uniquement est prioritaire
  // par rapport à un médecin J+N. Les médecins N uniquement n'ont aucune
  // possibilité de recevoir J ; il faut donc utiliser prioritairement leur
  // capacité N, puis départager ces médecins par leur charge.
  const shift=String(doc.shift||'BOTH').toUpperCase();
  const nOnlyPriority=(type==='N' && shift==='N') ? 1 : 0;
  const priorityPenalty=nOnlyPriority ? 0 : 1;

  // Pour J+N, préférence douce pour le poste encore absent, sans jamais
  // passer devant une contrainte obligatoire ni l'équité de charge.
  const shiftBalance=type==='J' ? (x.J||0) : (x.N||0)*2;
  const opposite=type==='J' ? (x.N||0)*2 : (x.J||0);
  const imbalance=Math.max(0,shiftBalance-opposite);

  // En plus de la charge pondérée, on pénalise les J/N trop rapprochés
  // pour favoriser le maximum de séparation possible.
  const separation=separationPenalty(doc,type,date,planning);

  // La priorité N-only est forte mais reste une priorité de sélection :
  // entre plusieurs N-only, la charge cumulée puis mensuelle départage.
  return priorityPenalty*1000000000 + total*100000 + month*1000 + shiftBalance*20 + imbalance*5 + (weekendUnits(x,'G')*2) + (weekendUnits(x,'F')*2) + separation;
}
function recoveryDeltaFor(planning, absences){
  const delta={};
  for(const x of Object.values(planning||{}).flat()){
    if(!x?.doctor) continue;
    if(!delta[x.doctor]) delta[x.doctor]=0;
    if(["N","G","F"].includes(x.type)) delta[x.doctor]+=2;
  }
  for(const a of absences||[]){
    if(a.type!=="RÉCUP" || !delta[a.doctor]) continue;
    delta[a.doctor]-=daysInclusive(a.start,a.end);
  }
  return delta;
}
function undoConfirmedPeriod(h){
  if(!h?.confirmed || !h?.planning) return;
  const pd=h.period||{};
  const start=pd.start?dateFromISO(pd.start):null, end=pd.end?dateFromISO(pd.end):null;
  const delta={};
  for(const x of Object.values(h.planning||{}).flat()){
    if(!x?.doctor) continue;
    if(x.type==="N") delta[x.doctor]=(delta[x.doctor]||0)+2;
  }
  for(const x of (h.fixed||[])){
    if(!x?.doctor || !["G","F"].includes(x.type)) continue;
    const d=dateFromISO(x.date);
    if(start&&end&&(d<start||d>end)) continue;
    delta[x.doctor]=(delta[x.doctor]||0)+2;
  }
  for(const a of (h.absences||[])){
    if(a.type!=="RÉCUP") continue;
    const aStart=dateFromISO(a.start), aEnd=dateFromISO(a.end);
    if(start&&end&&(aEnd<start||aStart>end)) continue;
    const s=start&&aStart<start?start:aStart, e=end&&aEnd>end?end:aEnd;
    delta[a.doctor]=(delta[a.doctor]||0)-(Math.floor((e-s)/86400000)+1);
  }
  for(const d of state.doctors){
    const old=state.recoveryLedger[d.name]||{solde:0,totalAcquise:0,totalPrise:0};
    const earned=Object.values(h.planning||{}).flat().filter(x=>x?.doctor===d.name&&x.type==="N").length*2
      +(h.fixed||[]).filter(x=>x?.doctor===d.name&&["G","F"].includes(x.type)&&(!start||!end||((dateFromISO(x.date)>=start)&&(dateFromISO(x.date)<=end)))).length*2;
    let taken=0;
    for(const a of (h.absences||[])){
      if(a.type!=="RÉCUP"||a.doctor!==d.name)continue;
      const aStart=dateFromISO(a.start),aEnd=dateFromISO(a.end);
      if(start&&end&&(aEnd<start||aStart>end))continue;
      const ss=start&&aStart<start?start:aStart, ee=end&&aEnd>end?end:aEnd;
      taken+=Math.floor((ee-ss)/86400000)+1;
    }
    const newAcquise=Math.max(0,Number(old.totalAcquise||0)-earned);
    const newPrise=Math.max(0,Number(old.totalPrise||0)-taken);
    state.recoveryLedger[d.name]={solde:newAcquise-newPrise,totalAcquise:newAcquise,totalPrise:newPrise,updatedAt:new Date().toISOString()};
  }
}
function generate(){
  const per=getPeriodDates();
  const generationAt=new Date().toLocaleString('fr-FR',{dateStyle:'short',timeStyle:'short'});

  const active=state.doctors.filter(d=>d.active);
  const baseCounts=counts();

  // Fixed guards are hard constraints and contribute to the workload.
  for(const fx of (state.fixed||[])){
    const fd=dateFromISO(fx.date);
    if(fd>=per.start && fd<=per.end && active.some(d=>d.name===fx.doctor) && ["J","N","G","F"].includes(fx.type)){
      baseCounts[fx.doctor][fx.type]=(baseCounts[fx.doctor][fx.type]||0)+1;
      baseCounts[fx.doctor].total=(baseCounts[fx.doctor].total||0)+1;
    }
  }

  // Recalcul précis des unités G/F après ajout des gardes fixes du nouveau planning.
  const currentFixedEntries=(state.fixed||[]).filter(f=>active.some(d=>d.name===f.doctor)&&['G','F'].includes(f.type));
  // counts() contient déjà les gardes confirmées historiques ; on conserve ses unités
  // puis on ajoute les unités des week-ends fixes du planning courant sans doubler.
  const fixedUnits=computeWeekendUnitsFromEntries(currentFixedEntries);
  active.forEach(d=>{
    const u=fixedUnits[d.name]||{};
    baseCounts[d.name]._weekendUnitsG=(baseCounts[d.name]._weekendUnitsG||0)+Number(u.G||0);
    baseCounts[d.name]._weekendUnitsF=(baseCounts[d.name]._weekendUnitsF||0)+Number(u.F||0);
  });

  // Confirmed workload in the calendar month of the period start.
  const targetMonth=per.start.getMonth(), targetYear=per.start.getFullYear();
  const monthly={};
  active.forEach(d=>monthly[d.name]={J:0,N:0,G:0,F:0});
  Object.values(state.planningHistory||{}).forEach(h=>{
    if(!h || h.confirmed!==true || h.includeInEquity===false) return;
    if(!equityEntryIsAfterReset(h.generatedAtISO)) return;
    Object.values(h.planning||{}).flat().forEach(x=>{
      if(!x?.doctor || !monthly[x.doctor] || !x.date) return;
      const dt=dateFromISO(x.date);
      if(dt && dt.getMonth()===targetMonth && dt.getFullYear()===targetYear && monthly[x.doctor][x.type]!==undefined) monthly[x.doctor][x.type]++;
    });
    (h.fixed||[]).forEach(x=>{
      if(!x?.doctor || !monthly[x.doctor] || !x.date || !["J","N","G","F"].includes(x.type)) return;
      const dt=dateFromISO(x.date);
      if(dt && dt.getMonth()===targetMonth && dt.getFullYear()===targetYear) monthly[x.doctor][x.type]++;
    });
  });
  for(const fx of (state.fixed||[])){
    const fd=dateFromISO(fx.date);
    if(fd && fd.getMonth()===targetMonth && fd.getFullYear()===targetYear && monthly[fx.doctor] && ["J","N","G","F"].includes(fx.type)) monthly[fx.doctor][fx.type]++;
  }
  // Pour le mois courant, G/F sont comptés par week-end complet, pas par cellule.
  const monthFixedEntries=(state.fixed||[]).filter(f=>{
    const d=dateFromISO(f.date);
    return d && d.getMonth()===targetMonth && d.getFullYear()===targetYear && ['G','F'].includes(f.type);
  });
  const monthFixedUnits=computeWeekendUnitsFromEntries(monthFixedEntries);
  active.forEach(d=>{
    const u=monthFixedUnits[d.name]||{};
    monthly[d.name]._weekendUnitsG=Number(u.G||0);
    monthly[d.name]._weekendUnitsF=Number(u.F||0);
  });

  const dates=datesBetween(per.start,per.end).filter(d=>!isWeekend(d) && !isHoliday(d));
  const errors=[];
  const initial={planning:{},counts:JSON.parse(JSON.stringify(baseCounts)),monthly:JSON.parse(JSON.stringify(monthly)),score:0,missing:[]};
  // Recherche globale plus large : on conserve davantage d'états candidats
  // afin qu'un choix équilibré au début du mois ne rende pas impossible une
  // couverture plus loin dans le mois (congés/indisponibilités compris).
  const BEAM_WIDTH=Math.min(360,Math.max(120,active.length*18));
  const clone=o=>JSON.parse(JSON.stringify(o));
  const rand=()=>Math.random();

  function addAssignment(st,date,type,doc){
    const k=iso(date);
    if(!st.planning[k]) st.planning[k]=[];
    st.planning[k].push({doctor:doc.name,type,date:k});
    if(!st.counts[doc.name]) st.counts[doc.name]={J:0,N:0,G:0,F:0,total:0};
    st.counts[doc.name][type]=(st.counts[doc.name][type]||0)+1;
    st.counts[doc.name].total=(st.counts[doc.name].total||0)+1;
    if(!st.monthly[doc.name]) st.monthly[doc.name]={J:0,N:0,G:0,F:0};
    st.monthly[doc.name][type]=(st.monthly[doc.name][type]||0)+1;
    st.score += candidateScore(doc,type,st.counts,st.monthly,date,st.planning);
  }

  function eligibleCandidates(st,date,type){
    const list=active.filter(doc=>{
      if(!isEligibleForType(doc,type)) return false;
      if(!canAssign(doc,date,type,st.planning)) return false;
      if((st.planning[iso(date)]||[]).some(x=>x.doctor===doc.name)) return false;
      return true;
    });
    // Randomness is deliberately applied only among candidates that satisfy
    // all hard constraints. The score remains the main ranking criterion.
    return list.sort((a,b)=>{
      const sa=candidateScore(a,type,st.counts,st.monthly,date,st.planning);
      const sb=candidateScore(b,type,st.counts,st.monthly,date,st.planning);
      return (sa-sb)+(rand()-0.5)*Math.max(1000,Math.min(sa+sb,100000)*0.015);
    });
  }

  function dayFixed(date,type){
    return (state.fixed||[]).find(x=>x.date===iso(date)&&x.type===type);
  }

  // Build each day as a J+N pair. This is the key correction:
  // a day is considered covered only when BOTH posts have been found.
  function expandDay(states,date){
    const fixedJ=dayFixed(date,"J"), fixedN=dayFixed(date,"N");
    const next=[];
    for(const st of states){
      const k=iso(date);
      const js=fixedJ ? [active.find(d=>d.name===fixedJ.doctor)].filter(Boolean) : eligibleCandidates(st,date,"J");
      const ns=fixedN ? [active.find(d=>d.name===fixedN.doctor)].filter(Boolean) : eligibleCandidates(st,date,"N");

      // A fixed J/N is authoritative: validate only the basic data/absence
      // constraints here. Do not pass it through canAssign(), because
      // canAssign() intentionally rejects a doctor already fixed on that date.
      if(fixedJ && (!js[0] || !isEligibleForType(js[0],"J") || absOn(js[0],date))) continue;
      if(fixedN && (!ns[0] || !isEligibleForType(ns[0],"N") || absOn(ns[0],date))) continue;
      if(fixedJ && fixedN && fixedJ.doctor===fixedN.doctor) continue;

      // Keep a reasonably broad random shortlist. If either post has no
      // candidate in this state, this state is not allowed to contaminate a
      // state that still has a complete J+N solution.
      if(!js.length || !ns.length){
        const ns0=clone(st);
        ns0.missing=(ns0.missing||[]).concat(`${iso(date)} : ${!js.length?"J":"N"}`);
        ns0.score+=1e15;
        next.push(ns0);
        continue;
      }

      const jList=js.slice(0,Math.min(18,js.length));
      const nList=ns.slice(0,Math.min(18,ns.length));
      const pairs=[];
      for(const j of jList){
        for(const n of nList){
          if(j.name===n.name) continue;
          // Re-check N after the J assignment because same-day and rest rules
          // must be evaluated against the exact partial state.
          const tmp=clone(st);
          if(!fixedJ) addAssignment(tmp,date,"J",j);
          else {
            tmp.planning[k]=tmp.planning[k]||[];
            tmp.planning[k].push({doctor:j.name,type:"J",date:k});
          }
          if(!fixedN && !canAssign(n,date,"N",tmp.planning)) continue;
          if(fixedN && !canAssign(n,date,"N",st.planning)) continue;
          if(fixedN && n.name===j.name) continue;
          if(!fixedJ && fixedN){ /* handled below */ }
          const score=tmp.score + candidateScore(n,"N",tmp.counts,tmp.monthly,date,tmp.planning) + separationPenalty(j,"J",date,st.planning);
          pairs.push({j,n,score,jitter:rand()});
        }
      }
      pairs.sort((a,b)=>(a.score-b.score)+(a.jitter-b.jitter)*Math.max(1,Math.min(5000,a.score*0.002)));
      const keep=pairs.slice(0,Math.min(24,pairs.length));
      for(const p of keep){
        const out=clone(st);
        if(fixedJ){
          out.planning[k]=out.planning[k]||[];
          out.planning[k].push({doctor:p.j.name,type:"J",date:k});
        }else addAssignment(out,date,"J",p.j);
        if(fixedN){
          out.planning[k]=out.planning[k]||[];
          out.planning[k].push({doctor:p.n.name,type:"N",date:k});
        }else addAssignment(out,date,"N",p.n);
        next.push(out);
      }
    }

    // Prefer complete states. Missing coverage is retained only when no
    // complete state exists for this day among the current beam.
    const complete=next.filter(st=>!(st.missing||[]).some(x=>x.startsWith(iso(date))));
    let out=complete.length?complete:next;
    const seen=new Map();
    for(const st of out){
      const sig=Object.entries(st.planning).sort().map(([k,v])=>k+":"+v.map(x=>x.doctor+x.type).sort().join(",")).join("|")+"|m:"+(st.missing||[]).join(",");
      if(!seen.has(sig) || st.score<seen.get(sig).score) seen.set(sig,st);
    }
    return [...seen.values()].sort((a,b)=>a.score-b.score).slice(0,BEAM_WIDTH);
  }

  let beam=[initial];
  for(const date of dates){
    beam=expandDay(beam,date);
    if(!beam.length){
      errors.push(`Aucune solution partielle pour ${iso(date)}`);
      beam=[initial];
    }
  }

  function finalMetric(st){
    const missing=(st.missing||[]).length;
    let metric=missing*1e15;
    const loads=active.map(d=>workload(st.counts[d.name]||{}));
    const monthLoads=active.map(d=>workload(st.monthly[d.name]||{}));
    const mean=loads.length?loads.reduce((a,b)=>a+b,0)/loads.length:0;
    const meanM=monthLoads.length?monthLoads.reduce((a,b)=>a+b,0)/monthLoads.length:0;
    metric += loads.reduce((s,x)=>s+(x-mean)**2,0)*10000;
    metric += monthLoads.reduce((s,x)=>s+(x-meanM)**2,0)*500;

    // Preference for 1 J + 1 N for BOTH/J+N doctors, but never above
    // the hard constraints, coverage, rest, or charge equity.
    for(const d of active){
      const s=String(d.shift||"BOTH").toUpperCase();
      const x=st.counts[d.name]||{};
      // Si un médecin est N uniquement, il doit être utilisé prioritairement
      // pour couvrir les N avant de solliciter les médecins J+N.
      if(s==="N" && (x.N||0)===0) metric-=5000000;
      if(s==="N" && (x.N||0)>0) metric+=0;
      if(s!=="BOTH" && s!=="J+N") continue;
      if((x.J||0)===0) metric+=1e7;
      if((x.N||0)===0) metric+=2e7;
    }
    return metric;
  }

  beam.sort((a,b)=>finalMetric(a)-finalMetric(b));
  const best=beam[0]||initial;
  const planning=best.planning||{};

  // VALIDATION FINALE OBLIGATOIRE : aucune génération partielle ne peut être
  // enregistrée. Pour chaque jour ouvrable, il faut exactement un J et un N.
  // Les gardes fixes J/N comptent comme couverture et les affectations générées
  // complètent le poste manquant. Cette vérification intervient après le moteur
  // aléatoire afin de garantir qu'un résultat incomplet ne soit jamais accepté.
  const coverageErrors=[];
  for(const date of dates){
    const k=iso(date);
    const entries=planning[k]||[];
    const hasJ=!!(state.fixed||[]).some(x=>x.date===k&&x.type==='J') || entries.some(x=>x.type==='J');
    const hasN=!!(state.fixed||[]).some(x=>x.date===k&&x.type==='N') || entries.some(x=>x.type==='N');
    if(!hasJ) coverageErrors.push(`${k} : J`);
    if(!hasN) coverageErrors.push(`${k} : N`);
    const doctors=new Set();
    for(const x of entries){
      if(doctors.has(x.doctor)) coverageErrors.push(`${k} : ${x.doctor} affecté deux fois`);
      doctors.add(x.doctor);
    }
  }
  // Validation dure des affectations générées : aucune garde ne peut être
  // placée pendant une absence, ni deux fois le même jour, et la règle de
  // repos de 48 h entre deux gardes reste appliquée aux gardes réelles.
  const hardErrors=[];
  const byDoctor=new Map();
  for(const d of active) byDoctor.set(d.name,[]);
  for(const fx of (state.fixed||[])){
    if(!byDoctor.has(fx.doctor)) continue;
    byDoctor.get(fx.doctor).push({date:fx.date,type:fx.type,fixed:true});
  }
  for(const [k,arr] of Object.entries(planning)){
    for(const x of (arr||[])){
      if(!byDoctor.has(x.doctor)) continue;
      if(absOn(x.doctor,dateFromISO(x.date))) hardErrors.push(`${x.date} : ${x.doctor} pendant une absence`);
      byDoctor.get(x.doctor).push({date:x.date,type:x.type,fixed:false});
    }
  }
  for(const [doctor,arr] of byDoctor){
    const sorted=arr.map(x=>dateFromISO(x.date)).filter(Boolean).sort((a,b)=>a-b);
    for(let i=1;i<sorted.length;i++){
      if((sorted[i]-sorted[i-1])/86400000<2) hardErrors.push(`${iso(sorted[i])} : repos insuffisant pour ${doctor}`);
    }
  }
  const missing=[...coverageErrors,...hardErrors];

  // Une génération incomplète est refusée : on conserve le planning précédent
  // et on n'archive rien tant qu'une solution complète n'a pas été trouvée.
  if(missing.length){
    const datesText=missing.slice(0,30).join(' • ');
    document.getElementById('genMessage').innerHTML=
      `<div class="warn"><b>Génération refusée : couverture incomplète.</b><br>`+
      `${datesText}${missing.length>30?' • …':''}<br>`+
      `<small>Aucun planning partiel n'est enregistré. Vérifiez les habilitations, congés, indisponibilités, récupérations et gardes fixes.</small></div>`;
    renderAll();
    renderPlanning();
    try{showPage('planning');}catch(e){
      document.querySelectorAll('.page').forEach(x=>x.style.display='none');
      const p=document.getElementById('page-planning'); if(p)p.style.display='block';
    }
    return;
  }

  // La génération est maintenant validée à 100 %. Seulement à ce moment,
  // l'ancien planning de la même période est archivé/remplacé.
  if(state.planningHistory[per.key]){
    const old=historyData(per.key);
    if(old?.confirmed) undoConfirmedPeriod(old);
    const archiveKey=`${per.key}_historique_${Date.now()}`;
    state.planningHistory[archiveKey]={...JSON.parse(JSON.stringify(old)),
      includeInEquity:false,confirmed:false,
      label:`Généré le ${old.generatedAt||'date inconnue'}`,
      generatedAt:old.generatedAt||'',
      period:old.period||{start:iso(per.start),end:iso(per.end),type:per.type,weeks:per.weeks||1}};
    delete state.planningHistory[per.key];
  }

  state.planningHistory[per.key]={
    planning:JSON.parse(JSON.stringify(planning)),
    fixed:JSON.parse(JSON.stringify(state.fixed)),
    absences:JSON.parse(JSON.stringify(state.absences)),
    period:{start:iso(per.start),end:iso(per.end),type:per.type,weeks:per.weeks||1},
    includeInEquity:false,confirmed:false,
    label:`Généré le ${generationAt}`,generatedAt:generationAt,generatedAtISO:Date.now()
  };
  state.planning=planning;
  state.planningGeneratedAt=Date.now();
  localStorage.setItem('gardeMed',JSON.stringify(state));

  const loadsFinal=active.map(d=>workload(best.counts[d.name]||{}));
  const spread=loadsFinal.length?Math.max(...loadsFinal)-Math.min(...loadsFinal):0;
  if(missing.length){
    document.getElementById('genMessage').innerHTML=
      `<div class="warn"><b>Planning généré avec ${missing.length} poste(s) non couvert(s).</b><br>`+
      `${missing.join(" • ")}<br><small>Les contraintes saisies rendent ces postes impossibles à attribuer dans cette génération.</small></div>`;
  }else{
    document.getElementById('genMessage').innerHTML=
      `<div class="ok"><b>Planning généré sans conflit.</b><br>`+
      `Chaque jour ouvrable comporte 1 J + 1 N. Écart de charge pondérée : ${spread} point(s).<br>`+
      `<small>J=1 • N=2 • G=3 • F=3. Le planning reste un brouillon jusqu'à validation.</small></div>`;
  }
  renderAll();
  renderPlanning();
  // Always show the freshly generated planning immediately.
  try{showPage('planning');}catch(e){
    document.querySelectorAll('.page').forEach(x=>x.style.display='none');
    const p=document.getElementById('page-planning'); if(p)p.style.display='block';
  }
}

function historyData(key){
  const h=(state.planningHistory||{})[key];
  if(!h)return null;
  return h.planning? h : {planning:h,fixed:state.fixed,absences:state.absences};
}
function cellFor(doc,date,source){
  const src=source||{planning:state.planning,fixed:state.fixed,absences:state.absences};
  const k=iso(date);
  const fx=(src.fixed||[]).find(x=>x.doctor===doc&&x.date===k);
  if(fx)return fx.type;
  const p=(src.planning?.[k]||[]).find(x=>x.doctor===doc);
  if(p)return p.type;
  const a=(src.absences||[]).find(x=>x.doctor===doc&&date>=dateFromISO(x.start)&&date<=dateFromISO(x.end));
  if(a)return a.type==="CONGÉ"?"CONGÉ":(a.type==="RÉCUP"?"RÉCUP":"INDISP");
  return "";
}
function isHoliday(date){return state.fixed.some(x=>x.date===iso(date)&&x.type==="F")}
function cls(v){return {J:"cellJ",N:"cellN","CONGÉ":"cellC","INDISP":"cellI","RÉCUP":"cellR",G:"cellG",F:"cellF"}[v]||""}
function normalizedHistoryPeriod(key,h){
  const p=h?.period;
  if(p?.start){
    const s=dateFromISO(p.start);
    const e=p.end?dateFromISO(p.end):new Date(s);
    if(s && !Number.isNaN(s.getTime())) return {start:s,end:e,type:p.type||"month",weeks:p.weeks||1};
  }
  const k=String(key||"");
  let m=k.match(/^month_(\d{4})_(\d{1,2})$/);
  if(m){
    const y=Number(m[1]), mo=Number(m[2])-1;
    return {start:new Date(y,mo,1),end:new Date(y,mo+1,0),type:"month",weeks:1};
  }
  m=k.match(/^weeks_(\d{4}-\d{2}-\d{2})_(\d+)$/);
  if(m){
    const s=dateFromISO(m[1]), weeks=Math.max(1,Number(m[2])||1), e=new Date(s);
    e.setDate(e.getDate()+weeks*7-1);
    return {start:s,end:e,type:"weeks",weeks};
  }
  return null;
}

function historyPeriod(key,h){
  return normalizedHistoryPeriod(key,h);
}
function renderHistorySelect(){
  const el=document.getElementById("historySelect");
  const blocks=document.getElementById("historyBlocks");
  if(!el)return;
  const currentKey=getPeriodDates().key;
  const allHistory=state.planningHistory||{};
  const keys=Object.keys(allHistory).filter(k=>k!==currentKey).filter(k=>historyData(k)).sort((a,b)=>{
    const ha=allHistory[a]||{}, hb=allHistory[b]||{};
    const pa=historyPeriod(a,ha), pb=historyPeriod(b,hb);
    const ta=pa?.start?.getTime?.()||0, tb=pb?.start?.getTime?.()||0;
    if(tb!==ta)return tb-ta;
    return (Number(hb.generatedAtISO)||Date.parse(hb.generatedAt)||0)-(Number(ha.generatedAtISO)||Date.parse(ha.generatedAt)||0);
  });
  el.innerHTML=`<option value="">Aucun ancien planning</option>`+keys.map(k=>`<option value="${escapeHtml(k)}">${escapeHtml(historyPeriodLabel(k,allHistory[k]||{}))}</option>`).join("");
  if(!blocks)return;

  // Regroupe les anciens plannings par mois. Le mois actuel est toujours
  // présent et porte simplement le nom du mois courant (ex. « Septembre »).
  const monthMap={};
  keys.forEach(k=>{
    const h=allHistory[k]||{}, hp=historyPeriod(k,h); if(!hp?.start)return;
    const m=hp.start.getMonth();
    if(!monthMap[m])monthMap[m]=[];
    monthMap[m].push({k,h,hp,isCurrent:false});
  });

  // Ajouter explicitement le planning actuel dans le bloc du mois courant.
  const currentPer=getPeriodDates();
  const currentHasPlanning=!!(state.planning && Object.values(state.planning).some(v=>Array.isArray(v)&&v.length));
  const currentHistory=allHistory[currentKey];
  const currentHasFixed=(state.fixed||[]).some(x=>x.date>=iso(currentPer.start)&&x.date<=iso(currentPer.end)&&["G","F"].includes(x.type));
  const currentExists=currentHasPlanning || !!currentHistory || currentHasFixed;
  const cm=currentPer.start.getMonth();
  if(!monthMap[cm])monthMap[cm]=[];
  // Le planning actuel est affiché en tête du mois courant.
  monthMap[cm].unshift({k:currentKey,h:currentHistory||{planning:state.planning,fixed:state.fixed,absences:state.absences,period:{start:iso(currentPer.start),end:iso(currentPer.end),type:currentPer.type,weeks:currentPer.weeks||1},generatedAtISO:state.planningGeneratedAt},hp:currentPer,isCurrent:true,exists:currentExists});

  blocks.innerHTML=Object.keys(monthMap).map(Number).sort((a,b)=>a-b).map(m=>{
    const id=`history-month-${m}`;
    const yearItems=monthMap[m];
    return `<div class="history-month-card"><div class="history-month-head"><button type="button" class="history-month-button" onclick="toggleHistoryMonth('${id}')"><strong title="${months[m]}"><span class="month-full">${months[m]}</span><span class="month-short">${monthShort[m]}</span></strong><span>${yearItems.filter(x=>x.exists!==false).length} planning(s)</span><b>▾</b></button></div><div id="${id}" class="history-year-panel" hidden>${renderHistoryYears(yearItems)}</div></div>`;
  }).join("");
}
function renderHistoryYears(items){
  const years={};
  items.forEach(item=>{const y=item.hp.start.getFullYear();(years[y]||(years[y]=[])).push(item);});
  return Object.keys(years).sort((a,b)=>b-a).map(y=>{
    const id=`history-year-${y}-${Math.random().toString(36).slice(2,8)}`;
    const list=years[y].sort((a,b)=>{
      if(a.isCurrent&&!b.isCurrent)return -1;
      if(!a.isCurrent&&b.isCurrent)return 1;
      return b.hp.start-a.hp.start || (Number(b.h.generatedAtISO)||0)-(Number(a.h.generatedAtISO)||0);
    }).map(({k,h,hp,isCurrent,exists})=>{
      if(isCurrent && !exists){
        return `<div class="history-planning-card history-current-empty"><strong>${escapeHtml(months[hp.start.getMonth()])} ${hp.start.getFullYear()} — Planning actuel</strong><span>Aucun planning actuellement enregistré</span><button type="button" class="history-delete-btn" disabled>Supprimer</button></div>`;
      }
      const age=generatedAge(h);
      const title=isCurrent?`${months[hp.start.getMonth()]} ${hp.start.getFullYear()} — Planning actuel`:historyPeriodLabel(k,h);
      const onclick=isCurrent?`showCurrentPlanning()`: `selectHistoryBlock('${escapeHtml(k)}')`;
      const deleteCall=isCurrent?`deleteDisplayedPlanning()`: `deleteHistoryPlanningByKey('${escapeHtml(k)}')`;
      return `<div class="history-planning-card ${isCurrent?'history-current-card':''}" role="button" tabindex="0" onclick="${onclick}"><div class="history-planning-info"><strong>${escapeHtml(title)}</strong><span>${age?`Généré ${escapeHtml(age)}`:(isCurrent?'Planning actuel':'Date de génération inconnue')}</span></div><button type="button" class="history-delete-btn" onclick="event.stopPropagation();${deleteCall}">🗑 Supprimer</button></div>`;
    }).join("");
    return `<div class="history-year-card"><button type="button" class="history-year-button" onclick="toggleHistoryMonth('${id}')"><strong>${y}</strong><span>${years[y].length} planning(s)</span><b>▾</b></button><div id="${id}" class="history-planning-list" hidden>${list}</div></div>`;
  }).join("");
}
function toggleHistoryMonth(id){const el=document.getElementById(id);if(el)el.hidden=!el.hidden;}
function selectHistoryBlock(k){
  const sel=document.getElementById("historySelect"); if(sel)sel.value=k||"";
  showHistoryPlanningByKey(k);
}
function showHistoryPlanningByKey(k){
  if(!k)return alert("Aucun ancien planning disponible.");
  activePlanningView={type:"history",key:k};
  const h=historyData(k);if(!h)return;
  const per=normalizedHistoryPeriod(k,h);
  renderPlanning(h,per);
}

function showHistoryPlanning(){
  const k=document.getElementById("historySelect")?.value;
  showHistoryPlanningByKey(k);
}
function deleteHistoryPlanningByKey(k){
  if(!isAdmin())return alert("Accès réservé aux administrateurs.");
  if(!k)return alert("Ancien planning introuvable.");
  const h=historyData(k);
  if(!h)return alert("Ancien planning introuvable.");
  const label=historyPeriodLabel(k,h);
  if(!confirm(`Supprimer l'ancien planning ${label} ?`))return;
  delete state.planningHistory[k];
  delete state.confirmedPlanningKeys[k];
  if(activePlanningView?.type==="history" && activePlanningView.key===k)activePlanningView={type:"current",key:null};
  rebuildRecoveryLedger();
  localStorage.setItem("gardeMed",JSON.stringify(state));
  renderAll();
  renderPlanning();
}
function deleteHistoryPlanning(){
  const k=document.getElementById("historySelect")?.value;
  if(!k)return alert("Aucun ancien planning sélectionné.");
  deleteHistoryPlanningByKey(k);
}

function toggleHistoryMonth(id){const el=document.getElementById(id);if(el)el.hidden=!el.hidden;}
function selectHistoryBlock(k){
  const sel=document.getElementById("historySelect"); if(sel)sel.value=k||"";
  showHistoryPlanningByKey(k);
}
function showHistoryPlanningByKey(k){
  if(!k)return alert("Aucun ancien planning disponible.");
  activePlanningView={type:"history",key:k};
  const h=historyData(k);if(!h)return;
  const per=normalizedHistoryPeriod(k,h);
  renderPlanning(h,per);
}

function showHistoryPlanning(){
  const k=document.getElementById("historySelect")?.value;
  showHistoryPlanningByKey(k);
}
function deleteHistoryPlanning(){
  if(!isAdmin())return alert("Accès réservé aux administrateurs.");
  const k=document.getElementById("historySelect")?.value;
  if(!k)return alert("Aucun ancien planning sélectionné.");
  const h=historyData(k);
  if(!h)return alert("Ancien planning introuvable.");
  const label=historyPeriodLabel(k,h);
  if(!confirm(`Supprimer l'ancien planning ${label} ?`))return;
  delete state.planningHistory[k];
  delete state.confirmedPlanningKeys[k];
  if(activePlanningView?.type==="history" && activePlanningView.key===k){
    activePlanningView={type:"current",key:null};
  }
  rebuildRecoveryLedger();
  localStorage.setItem("gardeMed",JSON.stringify(state));
  renderAll();
  renderPlanning();
}
function deleteDisplayedPlanning(){
  if(!isAdmin())return alert("Accès réservé aux administrateurs.");

  // Cas 1 : un ancien planning est actuellement affiché.
  if(activePlanningView?.type==="history" && activePlanningView.key){
    const k=activePlanningView.key;
    const h=state.planningHistory?.[k];
    if(!h)return alert("Le planning affiché n'existe plus.");
    const label=historyPeriodLabel(k,h);
    if(!confirm(`Supprimer le planning affiché : ${label} ?`))return;
    delete state.planningHistory[k];
    delete state.confirmedPlanningKeys[k];
    activePlanningView={type:"current",key:null};
    rebuildRecoveryLedger();
    localStorage.setItem("gardeMed",JSON.stringify(state));
    renderAll();
    renderPlanning();
    return;
  }

  // Cas 2 : le planning actuel est affiché.
  const currentKey=getPeriodDates().key;
  const h=state.planningHistory?.[currentKey];
  const hasCurrent=Object.values(state.planning||{}).some(v=>Array.isArray(v)&&v.length>0);
  const hasCurrentFixed=(state.fixed||[]).some(x=>x.date>=getPeriodDates().start.toISOString().slice(0,10)&&x.date<=getPeriodDates().end.toISOString().slice(0,10)&&["G","F"].includes(x.type));
  if(!h && !hasCurrent && !hasCurrentFixed)return alert("Aucun planning actuel à supprimer.");
  const label=h?historyPeriodLabel(currentKey,h):"planning actuel";
  if(!confirm(`Supprimer le planning actuel : ${label} ?`))return;

  // Le planning généré et son snapshot actuel sont supprimés.
  delete state.planningHistory[currentKey];
  delete state.confirmedPlanningKeys[currentKey];
  state.planning={};
  state.planningGeneratedAt=null;
  activePlanningView={type:"current",key:null};
  rebuildRecoveryLedger();
  localStorage.setItem("gardeMed",JSON.stringify(state));
  renderAll();
  renderPlanning();
}

function showCurrentPlanning(){activePlanningView={type:"current",key:null};renderPlanning();}
function getExportContext(){
  if(activePlanningView?.type==="history" && activePlanningView.key){
    const h=historyData(activePlanningView.key);
    if(h){
      const per=normalizedHistoryPeriod(activePlanningView.key,h);
      return {per,source:h,history:true};
    }
  }
  return {per:getPeriodDates(),source:{planning:state.planning,fixed:state.fixed,absences:state.absences},history:false};
}

function exportPlanning(){
  const ctx=getExportContext(); const per=ctx.per; const src=ctx.source; const dates=datesBetween(per.start,per.end);
  let rows=[["Médecin",...dates.map(d=>`${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`)]];
  state.doctors.filter(d=>d.active).forEach(doc=>rows.push([doc.name,...dates.map(d=>cellFor(doc.name,d,src))]));
  let wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),"Planning"); XLSX.writeFile(wb,`planning_gardes_${ctx.history?"ancien_":""}${per.key}.xlsx`);
}

function exportPlanningWord(){
  const ctx=getExportContext(); const per=ctx.per; const src=ctx.source;
  const dates=datesBetween(per.start,per.end);
  const title=state.settings?.title||"Tox-Garde";
  const subtitle=state.settings?.subtitle||"Planification des gardes médicales";
  const periodTitle=per.type==="month"?`${months[per.start.getMonth()]} ${per.start.getFullYear()}`:per.type==="week"?`Semaine du ${iso(per.start)} au ${iso(per.end)}`:`${per.weeks} semaines : ${iso(per.start)} au ${iso(per.end)}`;
  const esc=escapeHtml;
  const colors={J:"#b7e4c7",N:"#f4a6a6",G:"#d9d9d9",F:"#bfe8ff","CONGÉ":"#ffe36e","RÉCUP":"#ffe36e",INDISP:"#ffe36e"};
  const dayNames=["D","L","M","M","J","V","S"];
  const headers=dates.map(dt=>{
    const h=(src.fixed||[]).some(x=>x.date===iso(dt)&&x.type==="F");
    const bg=h?"#bfe8ff":(isWeekend(dt)?"#d9d9d9":"#eef2f7");
    return `<th style="background:${bg};border:1px solid #777;padding:2px 1px;text-align:center;font-size:${dates.length>31?"5.5pt":dates.length>14?"6.5pt":"8pt"};line-height:1.05;white-space:nowrap"><div>${dayNames[dt.getDay()]}</div><div>${String(dt.getDate()).padStart(2,"0")}/${String(dt.getMonth()+1).padStart(2,"0")}</div></th>`;
  }).join("");
  const body=state.doctors.filter(d=>d.active).map(doc=>{
    const cells=dates.map(dt=>{
      const v=cellFor(doc.name,dt,src); const bg=colors[v]||"#ffffff";
      return `<td style="background:${bg};border:1px solid #777;padding:2px 1px;text-align:center;font-weight:bold;font-size:${dates.length>31?"6.5pt":dates.length>14?"7.5pt":"9pt"};line-height:1.05;white-space:nowrap">${esc(v||"")}</td>`;
    }).join("");
    return `<tr><td style="border:1px solid #777;padding:3px 4px;text-align:left;font-weight:bold;white-space:nowrap;font-size:8pt">${esc(doc.name)}</td>${cells}</tr>`;
  }).join("");

  // Format Word paysage réel + tableau ajusté automatiquement à la largeur de la page.
  // Aucun "Généré le" ni légende dans le document exporté.
  const html=`<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><title>${esc(title)} - ${esc(periodTitle)}</title>
<style>
@page Section1{size:841.89pt 595.28pt; mso-page-orientation:landscape; margin:18pt 18pt 18pt 18pt;}
div.Section1{page:Section1;}
body{font-family:Arial,sans-serif;color:#182230;margin:0;padding:0;}
.header{text-align:center;margin:0 0 7pt 0;}
.center{font-size:12pt;font-weight:700;letter-spacing:.5pt;margin-bottom:3pt;}
h1{font-size:18pt;margin:0 0 2pt 0;}
h2{font-size:10pt;margin:0;color:#4b5563;font-weight:normal;}
.period{text-align:center;font-size:9pt;font-weight:700;margin:5pt 0 7pt 0;}
table.planning{width:100%;border-collapse:collapse;table-layout:fixed;mso-table-layout-alt:fixed;}
table.planning th:first-child,table.planning td:first-child{width:105pt;}
table.planning th,table.planning td{mso-fit-shrink:1;}
</style></head>
<body><div class="Section1">
<div class="header"><div class="center">CENTRE ANTIPOISON</div><h1>${esc(title)}</h1><h2>${esc(subtitle)}</h2></div>
<div class="period">Planning des gardes — ${esc(periodTitle)}</div>
<table class="planning"><thead><tr><th style="background:#eef2f7;border:1px solid #777;padding:3px;text-align:left;font-size:8pt">Médecin</th>${headers}</tr></thead><tbody>${body}</tbody></table>
</div></body></html>`;
  const blob=new Blob(["\ufeff",html],{type:"application/msword;charset=utf-8"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=`${title.replace(/[^a-z0-9àâçéèêëîïôûùüÿñæœ_-]+/gi,"_")}_${ctx.history?"ancien_":""}planning_${per.key}.doc`; document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(a.href),1000);
}

function docByName(name){return state.doctors.find(d=>d.name===name)||null;}
function isConfirmedCurrentPlanningEntry(x){
  const per=getPeriodDates(); const h=state.planningHistory?.[per.key];
  return !!(h?.confirmed && h.planning && Object.values(h.planning).flat().some(y=>y.doctor===x.doctor&&y.date===x.date&&y.type===x.type));
}
function confirmCurrentPlanning(){
  if(!isAdmin())return alert("La validation du planning est réservée à l'administrateur.");
  const per=getPeriodDates();
  if(!state.planning || !Object.values(state.planning).flat().length){alert("Aucun planning à confirmer.");return;}
  const h=state.planningHistory?.[per.key];
  if(h?.confirmed){alert("Ce planning est déjà validé.");return;}
  if(!confirm("Valider ce planning ? Les scores de récupération seront enregistrés et le compteur temporaire sera réinitialisé."))return;
  const periodStart=per.start, periodEnd=per.end;
  const earned={}, taken={};
  for(const x of Object.values(state.planning).flat()){
    if(!x?.doctor)continue;
    if(x.type==="N")earned[x.doctor]=(earned[x.doctor]||0)+2;
  }
  for(const x of (state.fixed||[])){
    const d=dateFromISO(x.date);
    if(d>=periodStart&&d<=periodEnd&&["G","F"].includes(x.type))earned[x.doctor]=(earned[x.doctor]||0)+2;
  }
  for(const a of state.absences||[]){
    if(a.type!=="RÉCUP"||!a.doctor)continue;
    const aStart=dateFromISO(a.start),aEnd=dateFromISO(a.end);
    if(aEnd>=periodStart&&aStart<=periodEnd){
      const start=aStart<periodStart?periodStart:aStart,end=aEnd>periodEnd?periodEnd:aEnd;
      taken[a.doctor]=(taken[a.doctor]||0)+Math.floor((end-start)/86400000)+1;
    }
  }
  for(const d of state.doctors){
    const old=state.recoveryLedger[d.name]||{solde:0,totalAcquise:0,totalPrise:0};
    const e=earned[d.name]||0,t=taken[d.name]||0;
    const newAcquise=Number(old.totalAcquise||0)+e;
    const newPrise=Number(old.totalPrise||0)+t;
    state.recoveryLedger[d.name]={solde:newAcquise-newPrise,totalAcquise:newAcquise,totalPrise:newPrise,updatedAt:new Date().toISOString()};
  }
  if(h){
    h.confirmed=true;
    h.includeInEquity=true;
    h.planning=JSON.parse(JSON.stringify(state.planning));
    h.fixed=JSON.parse(JSON.stringify(state.fixed));
    h.absences=JSON.parse(JSON.stringify(state.absences));
    h.period={start:iso(per.start),end:iso(per.end),type:per.type,weeks:per.weeks||1};
  } else {
    state.planningHistory[per.key]={planning:JSON.parse(JSON.stringify(state.planning)),fixed:JSON.parse(JSON.stringify(state.fixed)),absences:JSON.parse(JSON.stringify(state.absences)),period:{start:iso(per.start),end:iso(per.end),type:per.type,weeks:per.weeks||1},confirmed:true,generatedAt:new Date().toLocaleString('fr-FR',{dateStyle:'short',timeStyle:'short'}),includeInEquity:true};
  }
  state.confirmedPlanningKeys[per.key]=true;
  // Le score temporaire du tableau est réinitialisé : le solde est désormais porté par recoveryLedger.
  localStorage.setItem('gardeMed',JSON.stringify(state));
  document.getElementById('genMessage').innerHTML='<div class="ok"><b>Planning validé.</b><br>Les récupérations ont été enregistrées et le score temporaire a été réinitialisé.</div>';
  renderAll();
}
function resetAllRecoveryScores(){
  if(!confirm("Remettre à 0 tous les scores de récupération ? Cette action ne supprimera pas les plannings ni l’historique."))return;
  state.recoveryLedger={};
  state.doctors.forEach(d=>{state.recoveryLedger[d.name]={solde:0,totalAcquise:0,totalPrise:0,updatedAt:new Date().toISOString()};});
  localStorage.setItem("gardeMed",JSON.stringify(state));
  renderAll();
  alert("Tous les scores de récupération ont été remis à 0.");
}
function resetAllEquityCounters(){
  if(!confirm("Remettre à 0 tous les compteurs J, N, G et F ainsi que les scores de récupération ? Les plannings et l'historique resteront visibles."))return;
  const now=Date.now();
  state.equityResetAt=now;
  state.recoveryLedger={};
  state.doctors.forEach(d=>{state.recoveryLedger[d.name]={solde:0,totalAcquise:0,totalPrise:0,updatedAt:new Date(now).toISOString()};});
  localStorage.setItem("gardeMed",JSON.stringify(state));
  renderAll();
  alert("Tous les compteurs J/N/G/F et les scores de récupération ont été remis à 0. Les plannings restent conservés pour consultation.");
}
function renderStats(){
  const el=document.getElementById("statsTable"); if(!el)return;
  const c=counts();
  el.innerHTML=`<table><thead><tr><th>Médecin</th><th>J</th><th>N</th><th>G</th><th>F</th><th>Total gardes</th><th>Récup acquise</th><th>Récup prise</th><th>Récup restante</th></tr></thead><tbody>`+
    state.doctors.map(d=>{
      const x=c[d.name]||{J:0,N:0,G:0,F:0,total:0,recupAcquise:0,recupPrise:0,recupSolde:0};
      const solde=x.recupSolde||0;
      const clsSolde=solde<0?" style=\"font-weight:700;color:#b42318\"":" style=\"font-weight:700\"";
      return `<tr><td class="name">${d.name}</td><td>${x.J||0}</td><td>${x.N||0}</td><td>${x.G||0}</td><td>${x.F||0}</td><td><b>${x.total||0}</b></td><td>${x.recupAcquise||0}</td><td>${x.recupPrise||0}</td><td${clsSolde}>${solde}</td></tr>`;
    }).join("")+`</tbody></table>`;
}
function renderTodayGuards(){
  const el=document.getElementById("todayGuards");
  if(!el)return;
  const today=iso(new Date());
  const guards=[];
  // Pour le planning actuel, utiliser la dernière version validée si elle existe.
  // Ainsi, le tableau de bord suit toujours les modifications après validation.
  const per=getPeriodDates();
  const h=state.planningHistory?.[per.key];
  const planningSource=(h?.confirmed && h?.planning) ? h.planning : (state.planning||{});
  // Gardes de jour uniquement : J, G et F. La garde N n'est jamais affichée ici.
  (planningSource[today]||[]).filter(x=>x.type==="J").forEach(x=>guards.push({doctor:x.doctor,type:"J"}));
  // G/F fixes du jour.
  (state.fixed||[]).filter(x=>x.date===today && ["G","F"].includes(x.type)).forEach(x=>guards.push({doctor:x.doctor,type:x.type}));
  const seen=new Set();
  const unique=guards.filter(x=>{const k=x.doctor+"|"+x.type;if(seen.has(k))return false;seen.add(k);return true;});
  const dt=dateFromISO(today);
  const dateLabel=`${String(dt.getDate()).padStart(2,"0")}/${String(dt.getMonth()+1).padStart(2,"0")}/${dt.getFullYear()}`;
  el.innerHTML=unique.length
    ? `<div style="width:100%;margin-bottom:4px;font-weight:700">${dateLabel}</div>`+unique.map(x=>`<div class="today-guard"><small>Garde ${x.type}</small>${x.doctor}</div>`).join("")
    : `<div class="today-empty">Aucune garde J, G ou F enregistrée pour aujourd’hui (${dateLabel}).</div>`;
}

function renderAll(){let body=document.getElementById("teamBody");let c=counts();body.innerHTML=state.doctors.map((d,i)=>`<tr><td class='name'><input value="${d.name.replaceAll('"','&quot;')}" onchange="renameDoctor(${i},this.value)"></td><td><input type='checkbox' ${d.active?"checked":""} onchange="state.doctors[${i}].active=this.checked;save()"></td><td><select onchange="state.doctors[${i}].shift=this.value;save()"><option value="BOTH" ${d.shift==="BOTH"?"selected":""}>J + N</option><option value="J" ${d.shift==="J"?"selected":""}>J uniquement</option><option value="N" ${d.shift==="N"?"selected":""}>N uniquement</option></select></td><td><input type="checkbox" ${d.weekend!==false?"checked":""} onchange="state.doctors[${i}].weekend=this.checked;save()"></td><td>${c[d.name]?.J||0}</td><td>${c[d.name]?.N||0}</td><td>${c[d.name]?.G||0}</td><td>${c[d.name]?.F||0}</td><td>${c[d.name]?.total||0}</td><td><button class='btn small danger' onclick='delDoctor(${i})'>Supprimer</button></td></tr>`).join("");
const sesAbs=currentSession();
const doctorSession=sesAbs?.role!=="admin";
const myDoctor=String(sesAbs?.doctor||"").trim();
let opts=doctorSession?state.doctors.filter(d=>String(d.name||"").trim()===myDoctor).map(d=>`<option>${d.name}</option>`).join(""):state.doctors.map(d=>`<option>${d.name}</option>`).join("");
document.getElementById("absDoctor").innerHTML=opts;
if(doctorSession){document.getElementById("absDoctor").value=myDoctor;document.getElementById("absDoctor").disabled=true;}else document.getElementById("absDoctor").disabled=false;
const visibleAbsences=doctorSession?state.absences.map((a,i)=>({a,i})).filter(x=>String(x.a.doctor||"").trim()===myDoctor):state.absences.map((a,i)=>({a,i}));
document.getElementById("absBody").innerHTML=visibleAbsences.map(({a,i})=>`<tr><td class='name'>${a.doctor}</td><td>${a.start}</td><td>${a.end}</td><td>${a.type}</td><td><button class='btn small danger' onclick='delAbs(${i})'>Supprimer</button></td></tr>`).join("");
document.getElementById("fixedBody").innerHTML=state.fixed.map((x,i)=>`<tr><td class='name'>${x.doctor}</td><td>${x.date}</td><td><span class='badge ${x.type==="G"?"bG":"bF"}'>${x.type}</span></td><td><button class='btn small danger' onclick='state.fixed.splice(${i},1);save()'>×</button></td></tr>`).join("");
let total=Object.values(c).reduce((s,x)=>s+x.total,0);document.getElementById("dashStats").innerHTML=`<div class='stat'><span>Médecins actifs</span><b>${state.doctors.filter(d=>d.active).length}</b></div><div class='stat'><span>G importées</span><b>${state.fixed.filter(x=>x.type==="G").length}</b></div><div class='stat'><span>F importées</span><b>${state.fixed.filter(x=>x.type==="F").length}</b></div><div class='stat'><span>Absences</span><b>${state.absences.length}</b></div><div class='stat'><span>Gardes enregistrées</span><b>${total}</b></div>`;renderTodayGuards();renderPlanning();renderHistorySelect()}
function renameDoctor(i,v){let old=state.doctors[i].name;v=v.trim();if(!v)return;state.doctors[i].name=v;state.absences.forEach(a=>{if(a.doctor===old)a.doctor=v});state.fixed.forEach(a=>{if(a.doctor===old)a.doctor=v});save()}
renderAll();
applyAppearance();
