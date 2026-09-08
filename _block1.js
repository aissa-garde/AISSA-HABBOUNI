
const demoDoctors=["Dr Daoudi","Dr Daghouane","Dr El Hariri","Dr Ennakhassi","Dr Benarina","Dr Faouzi","Dr Habbouni","Dr Mourabiti","Dr El Mrabet","Dr Belkhadir","Dr El Azher"];
let state=JSON.parse(localStorage.getItem("gardeMed")||"null")||{doctors:[],absences:[],fixed:[],planning:{},planningHistory:{},generationOffset:0,settings:{title:"Tox-Garde",subtitle:"Simple. Équitable. Intelligent",appBg:"#f5f7fb",headerBg:"#3f83c5",cardBg:"#ffffff",bgImage:"",appLogo:"",centerLogo:"",homeImage:""}};
if(!state.recoveryLedger) state.recoveryLedger={};
if(!state.confirmedPlanningKeys) state.confirmedPlanningKeys={};
state.settings=Object.assign({title:"Tox-Garde",subtitle:"Simple. Équitable. Intelligent",appBg:"#f5f7fb",headerBg:"#3f83c5",cardBg:"#ffffff",bgImage:""},state.settings||{});
if(!state.planningHistory) state.planningHistory={};
if(!state.settings) state.settings={title:"Tox-Garde",subtitle:"Simple. Équitable. Intelligent",appBg:"#f5f7fb",headerBg:"#3f83c5",cardBg:"#ffffff",bgImage:""};
if(!state.generationOffset) state.generationOffset=0;
if(!state.doctors.length) state.doctors=demoDoctors.map(name=>({name,active:true,shift:"BOTH",weekend:true}));
state.doctors.forEach(d=>{if(!d.shift)d.shift="BOTH";});
const months=["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"];
months.forEach((m,i)=>document.getElementById("month").innerHTML+=`<option value="${i}">${m}</option>`);
document.getElementById("month").value=8;
if(document.getElementById("weekStart"))document.getElementById("weekStart").value="2026-09-01";
let activePlanningView={type:"current",key:null};
function save(){save();renderAll()}
function openTab(id){document.querySelectorAll(".panel").forEach(x=>x.classList.remove("active"));document.getElementById(id).classList.add("active");document.querySelectorAll("nav button").forEach(x=>x.classList.toggle("active",x.dataset.tab===id))}
document.querySelectorAll("nav button").forEach(b=>b.onclick=()=>openTab(b.dataset.tab));
function addDoctor(){if(!isAdmin())return alert("Accès réservé aux administrateurs.");let n=document.getElementById("newDoctor").value.trim();if(!n)return;if(state.doctors.some(d=>d.name.toLowerCase()===n.toLowerCase()))return alert("Ce médecin existe déjà.");state.doctors.push({name:n,active:true,shift:"BOTH",weekend:true});document.getElementById("newDoctor").value="";save()}
function loadDemo(){state.doctors=demoDoctors.map(name=>({name,active:true,shift:"BOTH",weekend:true}));save()}
function importTeamFile(ev){if(!isAdmin())return alert("Cette action est réservée aux administrateurs.");
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
function delDoctor(i){if(!isAdmin())return alert("Accès réservé aux administrateurs.");if(confirm("Supprimer ce médecin ?")){state.doctors.splice(i,1);save()}}
function addAbsence(){
  const ses=currentSession(); if(!ses)return;
  const doctor=document.getElementById("absDoctor").value, start=document.getElementById("absStart").value, end=document.getElementById("absEnd").value;
  let type=String(document.getElementById("absType").value||"").trim().toUpperCase(); if(type==="RECUP")type="RÉCUP";
  if(!doctor||!start||!end||end<start)return alert("Vérifie les dates.");
  if(ses.role!=="admin"&&doctor!==ses.doctor)return alert("Un compte Médecin ne peut saisir que ses propres disponibilités/récupérations.");
  if(ses.role!=="admin"&&!['CONGÉ','INDISPONIBLE','INDISP','RÉCUP'].includes(type))return alert("Un compte Médecin peut uniquement saisir un congé, une indisponibilité ou une récupération.");
  if(type==="G"||type==="F"){
    const dates=[];let d=dateFromISO(start),last=dateFromISO(end);while(d<=last){dates.push(iso(d));d.setDate(d.getDate()+1);}
    for(const date of dates){if(state.fixed.some(x=>x.doctor===doctor&&x.date===date))return alert(`Le médecin a déjà une garde G/F le ${date}.`);const problem=validateFixedGuard(doctor,date);if(problem)return alert(problem+"\n\nLa garde G/F n'a pas été enregistrée.");}
    const now=Date.now();dates.forEach(date=>state.fixed.push({doctor,date,type,createdAt:now}));
  }else{
    if(state.absences.some(a=>a.doctor===doctor&&a.start===start&&a.end===end&&a.type===type))return alert("Cette saisie existe déjà.");
    state.absences.push({doctor,start,end,type,createdAt:Date.now()});
  }
  rebuildRecoveryLedger();save();
}
function delAbs(i){const ses=currentSession();if(!ses)return;if(ses.role!=="admin" && (!state.absences[i] || state.absences[i].doctor!==ses.doctor))return alert("Vous ne pouvez supprimer que vos propres congés/indisponibilités.");state.absences.splice(i,1);save()}
function deleteMergedEntry(kind,index){const ses=currentSession();if(!ses)return;if(ses.role!=="admin"){if(kind==="fixed")return alert("Un compte Médecin ne peut pas supprimer une garde G/F.");const a=state.absences[index];if(!a||a.doctor!==ses.doctor)return alert("Vous ne pouvez supprimer que vos propres congés/indisponibilités.");}if(kind==="fixed")state.fixed.splice(index,1);else state.absences.splice(index,1);save()}
function isWeekend(d){let x=d.getDay();return x===0||x===6}
function iso(d){
  if(typeof d==='string'){const m=d.match(/^(\d{4}-\d{2}-\d{2})/);if(m)return m[1];}
  if(d instanceof Date&&!isNaN(d))return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
  return "";
}
function dateFromISO(s){let [y,m,d]=s.split("-").map(Number);return new Date(y,m-1,d)}
function absOn(doc,date){return state.absences.some(a=>a.doctor===doc&&date>=dateFromISO(a.start)&&date<=dateFromISO(a.end))}
function fixedOn(doc,date){return state.fixed.find(x=>x.doctor===doc&&x.date===iso(date))}
function fixedAny(date){return state.fixed.filter(x=>x.date===iso(date))}
function parseExcelDate(v){
  if(v instanceof Date && !isNaN(v)) return iso(v);
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
function importExcel(ev){if(!isAdmin())return alert("Cette action est réservée aux administrateurs.");
  const f=ev.target.files?.[0]; if(!f) return;
  const r=new FileReader();
  r.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:"array",cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:"",raw:true});
      let added=0, skipped=0;
      const allowed={"CONGE":"CONGÉ","CONGÉ":"CONGÉ","INDISPONIBLE":"INDISPONIBLE","INDISP":"INDISPONIBLE","RECUP":"RÉCUP","RÉCUP":"RÉCUP","G":"G","F":"F"};
      rows.forEach(row=>{
        const keys=Object.keys(row);
        const norm=x=>String(x).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim();
        const get=names=>{const k=keys.find(x=>names.some(n=>norm(x).includes(norm(n))));return k!==undefined?row[k]:""};
        const doctor=String(get(["medecin","médecin","nom"])).trim();
        const rawType=String(get(["type","garde"])).trim().toUpperCase();
        const type=allowed[rawType];
        const s=parseExcelDate(get(["date debut","date début","debut","date","jour"]));
        const en=parseExcelDate(get(["date fin","date fin","fin"]));
        if(!doctor||!s||!type||!state.doctors.some(d=>d.name.toLowerCase()===doctor.toLowerCase())){skipped++;return;}
        const a=dateFromISO(s), b=dateFromISO(en||s);
        if(b<a){skipped++;return;}
        if(type==="G"||type==="F"){
          for(let d=new Date(a);d<=b;d.setDate(d.getDate()+1)){
            const date=iso(d);
            if(state.fixed.some(x=>x.doctor===doctor&&x.date===date)){skipped++;continue;}
            const problem=validateFixedGuard(doctor,date);
            if(problem){skipped++;continue;}
            state.fixed.push({doctor,date,type,createdAt:Date.now()}); added++;
          }
        }else{
          const duplicate=state.absences.some(x=>x.doctor===doctor&&x.start===s&&x.end===(en||s)&&x.type===type);
          if(duplicate){skipped++;return;}
          state.absences.push({doctor,start:s,end:en||s,type,createdAt:Date.now()}); added++;
        }
      });
      save();
      alert(`${added} saisie(s) importée(s)${skipped?` ; ${skipped} ligne(s)/jour(s) ignoré(s)`:""}.`);
      ev.target.value="";
    }catch(err){alert("Erreur d'import Excel : "+err.message);}
  };
  r.readAsArrayBuffer(f);
}
function clearFixed(){if(!isAdmin())return alert("Cette action est réservée aux administrateurs.");if(confirm("Effacer toutes les saisies : congés, indisponibilités, récupérations et gardes G/F ?")){state.fixed=[];state.absences=[];save()}}
function downloadTemplate(){let rows=[["Médecin","Date début","Date fin","Type"],["Dr Exemple 1","2026-09-01","2026-09-05","CONGÉ"],["Dr Exemple 2","2026-09-08","2026-09-08","INDISPONIBLE"],["Dr Exemple 3","2026-09-10","2026-09-10","RÉCUP"],["Dr Exemple 1","2026-09-12","2026-09-13","G"],["Dr Exemple 2","2026-09-20","2026-09-20","F"]];let wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),"Disponibilités");XLSX.writeFile(wb,"modele_disponibilites_gardes.xlsx")}
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
  // Après la fin du dernier congé/indisponibilité, imposer 48 h complètes avant toute nouvelle garde.
  const endedAbs=state.absences.filter(a=>a.doctor===doc.name && (a.type==="CONGÉ"||a.type==="INDISPONIBLE"||a.type==="INDISP") && dateFromISO(a.end)<date);
  if(endedAbs.length){
    const end=endedAbs.reduce((latest,a)=>{const d=dateFromISO(a.end);return d>latest?d:latest;},dateFromISO(endedAbs[0].end));
    if((date-end)/36e5<48)return false;
  }
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
  const endAbs=state.absences.filter(x=>x.doctor===doctor&&(x.type==="CONGÉ"||x.type==="INDISPONIBLE"||x.type==="INDISP")&&dateFromISO(x.end)<date).sort((a,b)=>dateFromISO(b.end)-dateFromISO(a.end))[0];
  if(endAbs && (dateFromISO(date)-dateFromISO(endAbs.end))/36e5<48) return `Le médecin reprend après ${endAbs.type.toLowerCase()} le ${endAbs.end} : 48 h de repos sont nécessaires.`;
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

function rebuildRecoveryLedger(){
  const ledger={};
  state.doctors.forEach(d=>ledger[d.name]={solde:0,totalAcquise:0,totalPrise:0,updatedAt:new Date().toISOString()});
  const earnedSeen=new Set();
  Object.values(state.planningHistory||{}).forEach(h=>{
    if(!h||h.confirmed!==true||h.includeInEquity===false)return;
    const hp=h.period||{}; const start=hp.start?dateFromISO(hp.start):null,end=hp.end?dateFromISO(hp.end):null;
    Object.values(h.planning||{}).flat().forEach(x=>{
      if(!x?.doctor||x.type!=="N"||!ledger[x.doctor])return;
      const key=`N|${x.doctor}|${x.date}`; if(!earnedSeen.has(key)){earnedSeen.add(key);ledger[x.doctor].totalAcquise+=2;}
    });
    (h.fixed||[]).forEach(x=>{
      if(!x?.doctor||!["G","F"].includes(x.type)||!ledger[x.doctor])return;
      const d=dateFromISO(x.date); if(start&&end&&(d<start||d>end))return;
      const key=`${x.type}|${x.doctor}|${x.date}`; if(!earnedSeen.has(key)){earnedSeen.add(key);ledger[x.doctor].totalAcquise+=2;}
    });
  });
  const takenSeen=new Set();
  (state.absences||[]).forEach(a=>{
    if(a?.type!=="RÉCUP"||!a.doctor||!ledger[a.doctor])return;
    let d=dateFromISO(a.start),end=dateFromISO(a.end);
    while(d<=end){const key=`${a.doctor}|${iso(d)}`;if(!takenSeen.has(key)){takenSeen.add(key);ledger[a.doctor].totalPrise+=1;}d.setDate(d.getDate()+1);}
  });
  Object.keys(ledger).forEach(name=>{const x=ledger[name];x.solde=x.totalAcquise-x.totalPrise;});
  state.recoveryLedger=ledger;
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
  return ({J:1,N:2,G:4,F:4}[type]||0);
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
function workload(c){
  return (c?.J||0)*1 + (c?.N||0)*2 + (c?.G||0)*4 + (c?.F||0)*4;
}
function candidateScore(doc,type,c,monthly){
  const x=c[doc.name]||{J:0,N:0,G:0,F:0};
  const m=(monthly&&monthly[doc.name])||{J:0,N:0,G:0,F:0};
  // EQUITE PRIORITAIRE : la charge pondérée cumulée est l'objectif
  // principal. La charge du mois courant vient ensuite.
  // Important : les G/F fixes sont déjà inclus avec leur vrai poids
  // (G=4, F=4). Ils doivent donc naturellement réduire les J/N
  // attribués ensuite à un médecin déjà très chargé.
  const total=workload(x);
  const month=workload(m);
  const shiftBalance=type==="J" ? (x.J||0) : (x.N||0)*2;
  const opposite=type==="J" ? (x.N||0)*2 : (x.J||0);
  const imbalance=Math.max(0,shiftBalance-opposite);
  return total*100000 + month*1000 + shiftBalance*20 + imbalance*5;
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

  // Archive/replace an existing draft or confirmed planning for the same period.
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

  const active=state.doctors.filter(d=>d.active);
  const baseCounts=counts();

  // Fixed guards (G/F, and any legacy/manual J/N records) are hard constraints.
  // Their points are included in the workload. A fixed J or N also counts as
  // coverage for that day, so the generator only fills the missing post.
  for(const fx of (state.fixed||[])){
    const fd=dateFromISO(fx.date);
    if(fd>=per.start && fd<=per.end && active.some(d=>d.name===fx.doctor) && ["J","N","G","F"].includes(fx.type)){
      baseCounts[fx.doctor][fx.type]=(baseCounts[fx.doctor][fx.type]||0)+1;
      baseCounts[fx.doctor].total=(baseCounts[fx.doctor].total||0)+1;
    }
  }

  // Confirmed workload in the calendar month of the period start.
  const targetMonth=per.start.getMonth(), targetYear=per.start.getFullYear();
  const monthly={};
  active.forEach(d=>monthly[d.name]={J:0,N:0,G:0,F:0});
  Object.values(state.planningHistory||{}).forEach(h=>{
    if(!h || h.confirmed!==true || h.includeInEquity===false) return;
    if(!equityEntryIsAfterReset(h.generatedAtISO)) return;
    const hp=h.planning||{};
    Object.values(hp).flat().forEach(x=>{
      if(!x?.doctor || !monthly[x.doctor] || !x.date) return;
      const dt=dateFromISO(x.date);
      if(dt && dt.getMonth()===targetMonth && dt.getFullYear()===targetYear && monthly[x.doctor][x.type]!==undefined){
        monthly[x.doctor][x.type]++;
      }
    });
    (h.fixed||[]).forEach(x=>{
      if(!x?.doctor || !monthly[x.doctor] || !x.date || !["J","N","G","F"].includes(x.type)) return;
      const dt=dateFromISO(x.date);
      if(dt && dt.getMonth()===targetMonth && dt.getFullYear()===targetYear) monthly[x.doctor][x.type]++;
    });
  });
  // Include already-fixed G/F of the generated period in the current-month load.
  for(const fx of (state.fixed||[])){
    const fd=dateFromISO(fx.date);
    if(fd && fd.getMonth()===targetMonth && fd.getFullYear()===targetYear && monthly[fx.doctor] && ["J","N","G","F"].includes(fx.type)){
      monthly[fx.doctor][fx.type]++;
    }
  }

  const dates=datesBetween(per.start,per.end).filter(d=>!isWeekend(d));
  const requiredTypes=[];
  // The requested generator always seeks both daily posts.
  requiredTypes.push("J","N");

  // Génération aléatoire robuste : plusieurs tentatives indépendantes sont
  // construites en respectant d'abord toutes les contraintes. Les candidats
  // les moins chargés sont favorisés, mais un choix aléatoire est conservé
  // afin que "Régénérer" puisse réellement produire une autre répartition.
  function clone(o){return JSON.parse(JSON.stringify(o));}
  function addAssignment(st,date,type,doc){
    const k=iso(date);
    if(!st.planning[k]) st.planning[k]=[];
    st.planning[k].push({doctor:doc.name,type,date:k});
    if(!st.counts[doc.name]) st.counts[doc.name]={J:0,N:0,G:0,F:0,total:0};
    st.counts[doc.name][type]=(st.counts[doc.name][type]||0)+1;
    st.counts[doc.name].total=(st.counts[doc.name].total||0)+1;
    if(!st.monthly[doc.name]) st.monthly[doc.name]={J:0,N:0,G:0,F:0};
    st.monthly[doc.name][type]=(st.monthly[doc.name][type]||0)+1;
  }

  function eligibleCandidates(st,date,type){
    return active.filter(doc=>{
      if(!isEligibleForType(doc,type)) return false;
      if(!canAssign(doc,date,type,st.planning)) return false;
      // Un médecin ne peut jamais faire J et N le même jour.
      if((st.planning[iso(date)]||[]).some(x=>x.doctor===doc.name)) return false;
      return true;
    });
  }

  function randomPick(candidates,st,type){
    if(!candidates.length)return null;
    // EQUITE AVANT L'OBJECTIF J/N : un médecin ayant déjà reçu un G/F
    // doit être naturellement moins prioritaire pour J/N, puisque G/F
    // portent déjà 4 points chacun. L'objectif 1 J + 1 N reste un objectif
    // souple, jamais une priorité qui pourrait surcharger un médecin.
    const ranked=candidates.map(doc=>{
      const x=st.counts[doc.name]||{};
      const m=st.monthly[doc.name]||{};
      const both=String(doc.shift||"BOTH").toUpperCase()==="BOTH";
      const missingShift=both && (m[type]||0)===0;
      const score=candidateScore(doc,type,st.counts,st.monthly);
      return {doc,missingShift,score,total:workload(x),month:workload(m)};
    }).sort((a,b)=>a.score-b.score);

    // On ne donne le petit bonus J/N qu'aux candidats proches du meilleur
    // niveau de charge. Ainsi un médecin avec un week-end G+G (8 points)
    // ne reçoit pas automatiquement J+N simplement parce qu'il lui manque
    // ces deux types.
    const best=ranked[0];
    const eligibleBand=ranked.filter(r=>
      r.total<=best.total+2 && r.month<=best.month+2
    );
    const targetPool=eligibleBand.filter(r=>r.missingShift);
    const pool=(targetPool.length?targetPool:eligibleBand.length?eligibleBand:ranked)
      .slice(0,Math.min(6,ranked.length));

    // Tirage pondéré parmi les meilleurs candidats : la génération reste
    // réellement aléatoire sans sacrifier l'équité de charge.
    const weights=pool.map((_,i)=>1/(1+i*0.72));
    const totalWeight=weights.reduce((a,b)=>a+b,0);
    let r=Math.random()*totalWeight;
    for(let i=0;i<pool.length;i++){r-=weights[i];if(r<=0)return pool[i].doc;}
    return pool[0].doc;
  }

  function expandRandom(initialState){
    const st=clone(initialState);
    // IMPORTANT : les jours restent chronologiques.
    // La randomisation des jours était dangereuse : elle pouvait attribuer
    // une garde au jour J+1 avant de connaître la garde du jour J et donc
    // contourner la règle de repos. La randomisation se fait uniquement
    // entre les candidats d'une même journée.
    const dayOrder=dates.map(d=>new Date(d));

    for(const date of dayOrder){
      const k=iso(date);
      const dayFixed=(state.fixed||[]).filter(x=>x.date===k&&["J","N"].includes(x.type));
      st.planning[k]=st.planning[k]||[];
      for(const fx of dayFixed){
        if(!st.planning[k].some(x=>x.type===fx.type)) st.planning[k].push({doctor:fx.doctor,type:fx.type,date:k});
      }

      const missingTypes=requiredTypes.filter(t=>!dayFixed.some(x=>x.type===t));
      // L'ordre J/N est également légèrement randomisé.
      if(Math.random()<0.5) missingTypes.reverse();
      for(const type of missingTypes){
        const candidates=eligibleCandidates(st,date,type);
        if(!candidates.length){
          st.missing=(st.missing||[]).concat(`${k} : ${type}`);
          continue;
        }
        const chosen=randomPick(candidates,st,type);
        addAssignment(st,date,type,chosen);
      }
    }
    return st;
  }

  function finalMetric(st){
    const missing=(st.missing||[]).length;
    // Repos obligatoire : toute garde à moins de 48 h d'une autre garde
    // du même médecin est considérée comme une violation absolue.
    const restSeen={};
    for(const d of active){
      const ds=Object.values(st.planning||{}).flat().filter(x=>x?.doctor===d.name&&x?.date).map(x=>dateFromISO(x.date)).sort((a,b)=>a-b);
      for(let i=1;i<ds.length;i++){
        if((ds[i]-ds[i-1])/36e5<48) restSeen[d.name]=(restSeen[d.name]||0)+1;
      }
      const fds=state.fixed.filter(x=>x?.doctor===d.name&&x?.date).map(x=>dateFromISO(x.date));
      for(const a of ds) for(const b of fds){
        const h=Math.abs(a-b)/36e5;
        if(h>0&&h<48) restSeen[d.name]=(restSeen[d.name]||0)+1;
      }
    }
    const restCount=Object.values(restSeen).reduce((a,b)=>a+b,0);
    let metric=missing*1e15 + restCount*1e14;
    const loads=active.map(d=>workload(st.counts[d.name]||{}));
    const monthLoads=active.map(d=>workload(st.monthly[d.name]||{}));
    const mean=loads.length?loads.reduce((a,b)=>a+b,0)/loads.length:0;
    const meanM=monthLoads.length?monthLoads.reduce((a,b)=>a+b,0)/monthLoads.length:0;
    // Equité globale puis équité du mois courant.
    metric += loads.reduce((s,x)=>s+(x-mean)**2,0)*10000;
    metric += monthLoads.reduce((s,x)=>s+(x-meanM)**2,0)*500;

    // OBJECTIF J+N = SOUPLE et SUBORDONNE A L'EQUITE.
    // On ne pénalise que légèrement l'absence d'un J ou d'un N chez un
    // médecin J+N, et seulement si ce médecin n'est pas déjà nettement
    // plus chargé que les autres. Cela évite le problème observé : les
    // médecins ayant assuré un week-end G+G recevaient encore J+N, alors
    // que leur charge pondérée était déjà supérieure.
    const bothLoads=active.filter(d=>String(d.shift||"BOTH").toUpperCase()==="BOTH")
      .map(d=>({d,load:workload(st.counts[d.name]||{}),month:workload(st.monthly[d.name]||{})}));
    const minBothLoad=bothLoads.length?Math.min(...bothLoads.map(x=>x.load)):0;
    for(const item of bothLoads){
      const d=item.d, x=st.counts[d.name]||{};
      const alreadyHeavy=item.load>minBothLoad+2;
      if(!alreadyHeavy){
        if((x.J||0)===0){
          const possible=dates.some(dt=>isEligibleForType(d,"J") && canAssign(d,dt,"J",st.planning));
          if(possible) metric+=2500;
        }
        if((x.N||0)===0){
          const possible=dates.some(dt=>isEligibleForType(d,"N") && canAssign(d,dt,"N",st.planning));
          if(possible) metric+=2500;
        }
      }
      metric += workload(x)*0.01;
    }
    // Une petite composante aléatoire permet de choisir entre deux plannings
    // pratiquement équivalents et rend les régénérations réellement variées.
    metric += Math.random()*0.5;
    return metric;
  }

  const initial={
    planning:{}, counts:JSON.parse(JSON.stringify(baseCounts)),
    monthly:JSON.parse(JSON.stringify(monthly)), score:0
  };

  // Plusieurs essais indépendants. On retient le meilleur planning faisable
  // selon les objectifs ci-dessus. Ainsi la génération reste aléatoire tout en
  // étant beaucoup plus robuste qu'un simple tirage aléatoire.
  const attempts=Math.min(700,Math.max(220,active.length*35));
  let best=null, bestMetric=Infinity;
  for(let i=0;i<attempts;i++){
    const candidate=expandRandom(initial);
    const metric=finalMetric(candidate);
    if(metric<bestMetric){bestMetric=metric;best=candidate;}
    // Si la couverture est parfaite et que la charge est déjà très équilibrée,
    // quelques centaines d'essais suffisent généralement.
    if((candidate.missing||[]).length===0 && i>120 && bestMetric<1000 && Math.random()<0.03) break;
  }

  // Sécurité : si aucun essai n'a été construit, revenir à l'état initial.
  const chosen=best||initial;
  const planning=chosen.planning||{};
  const missing=chosen.missing||[];

  // Contrôle final de sécurité : aucune garde générée ne doit être
  // consécutive à une autre garde du même médecin. Cette vérification est
  // indépendante du tirage aléatoire et protège contre toute régression.
  const restViolations=[];
  for(const d of active){
    const datesDoc=Object.values(planning).flat().filter(x=>x?.doctor===d.name&&x?.date).map(x=>dateFromISO(x.date)).sort((a,b)=>a-b);
    for(let i=1;i<datesDoc.length;i++){
      const hours=(datesDoc[i]-datesDoc[i-1])/36e5;
      if(hours<48) restViolations.push(`${d.name} : ${iso(datesDoc[i-1])} → ${iso(datesDoc[i])}`);
    }
    const fixedDoc=state.fixed.filter(x=>x.doctor===d.name&&x.date).map(x=>dateFromISO(x.date));
    for(const a of datesDoc){
      for(const b of fixedDoc){
        const hours=Math.abs(a-b)/36e5;
        if(hours>0 && hours<48) restViolations.push(`${d.name} : ${iso(a)} → garde fixe ${iso(b)}`);
      }
    }
  }
  if(restViolations.length){
    missing.push(`Repos insuffisant : ${restViolations.slice(0,8).join(" • ")}`);
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

  const loadsFinal=active.map(d=>workload(chosen.counts[d.name]||{}));
  const spread=loadsFinal.length?Math.max(...loadsFinal)-Math.min(...loadsFinal):0;
  if(missing.length){
    document.getElementById('genMessage').innerHTML=
      `<div class="warn"><b>Planning généré avec ${missing.length} poste(s) non couvert(s).</b><br>`+
      `${missing.join(" • ")}<br><small>Les contraintes saisies rendent ces postes impossibles à attribuer.</small></div>`;
  }else{
    document.getElementById('genMessage').innerHTML=
      `<div class="ok"><b>Planning généré sans conflit.</b><br>`+
      `Chaque jour ouvrable comporte 1 J + 1 N. Écart de charge pondérée : ${spread} point(s).<br>`+
      `<small>J=1 • N=2 • G=4 • F=4. Le planning reste un brouillon jusqu'à validation.</small></div>`;
  }
  // Affichage automatique du planning nouvellement généré.
  // On fixe d'abord la vue courante, puis on rafraîchit les composants.
  activePlanningView={type:"current",key:null};
  renderAll();
  renderPlanning();
  openTab("planning");
}

function cellFor(doc,date){
  let k=iso(date),fx=fixedOn(doc,date);
  if(fx)return fx.type;
  let a=state.absences.find(x=>x.doctor===doc&&date>=dateFromISO(x.start)&&date<=dateFromISO(x.end));
  if(a)return a.type==="CONGÉ"?"CONGÉ":(a.type==="RÉCUP"?"RÉCUP":"INDISP");
  let p=(state.planning[k]||[]).find(x=>x.doctor===doc);
  if(p)return p.type;
  return "";
}
function isHoliday(date){return fixedAny(date).some(x=>x.type==="F")}
function cls(v){return {J:"cellJ",N:"cellN","CONGÉ":"cellC","INDISP":"cellI","RÉCUP":"cellR",G:"cellG",F:"cellF"}[v]||""}
function generatedAgeText(value){
  if(value==null || value==='') return '';
  let t=null;
  if(typeof value==='number') t=value;
  else if(/^\d{1,2}\/\d{1,2}\/\d{4}(?: \d{1,2}:\d{2}(?::\d{2})?)?$/.test(String(value))){
    const m=String(value).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?: (\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if(m){ t=new Date(Number(m[3]),Number(m[2])-1,Number(m[1]),Number(m[4]||0),Number(m[5]||0),Number(m[6]||0)).getTime(); }
  }
  if(!t || Number.isNaN(t)) return '';
  const diff=Math.max(0,Date.now()-t);
  const min=Math.floor(diff/60000);
  if(min<1) return 'à l’instant';
  if(min<60) return `${min} min`;
  const h=Math.floor(min/60);
  if(h<24) return `${h} h`;
  const d=Math.floor(h/24);
  if(d<30) return `${d} j`;
  const mo=Math.floor(d/30);
  if(mo<12) return `${mo} mois`;
  return `${Math.floor(mo/12)} an${Math.floor(mo/12)>1?'s':''}`;
}

function generatedAge(h){
  if(!h) return '';
  return generatedAgeText(h.generatedAtISO || h.generatedAt);
}

function isHistoryConfirmed(key,h){
  // Compatibilité avec les anciens plannings validés : certaines anciennes
  // versions enregistraient la validation dans confirmedPlanningKeys sans
  // ajouter h.confirmed=true dans la fiche historique.
  return !!(h?.confirmed===true || state.confirmedPlanningKeys?.[key]===true);
}
function historyData(key){
  const h=(state.planningHistory||{})[key];
  if(!h)return null;
  return h.planning? h : {planning:h,fixed:state.fixed,absences:state.absences};
}
function cellFor(doc,date,source){
  const src=source||{planning:state.planning,fixed:state.fixed,absences:state.absences};
  let k=iso(date),fx=(src.fixed||[]).find(x=>x.doctor===doc&&x.date===k);
  if(fx)return fx.type;
  let a=(src.absences||[]).find(x=>x.doctor===doc&&date>=dateFromISO(x.start)&&date<=dateFromISO(x.end));
  if(a)return a.type==="CONGÉ"?"CONGÉ":(a.type==="RÉCUP"?"RÉCUP":"INDISP");
  let p=(src.planning?.[k]||[]).find(x=>x.doctor===doc);
  return p?p.type:"";
}
function isHoliday(date){return state.fixed.some(x=>x.date===iso(date)&&x.type==="F")}
function cls(v){return {J:"cellJ",N:"cellN","CONGÉ":"cellC","INDISP":"cellI","RÉCUP":"cellR",G:"cellG",F:"cellF"}[v]||""}
function renderHistorySelect(){
  const el=document.getElementById("historySelect");
  const blocks=document.getElementById("historyBlocks");
  if(!el)return;
  const currentKey=getPeriodDates().key;
  const hasData=k=>{
    const h=state.planningHistory?.[k];
    if(!h)return false;
    const p=Object.values(h.planning||{}).flat();
    const fixed=Array.isArray(h.fixed)?h.fixed:[];
    return p.length>0 || fixed.length>0;
  };
  const keys=Object.keys(state.planningHistory||{})
    .filter(k=>k!==currentKey && hasData(k))
    .sort((a,b)=>{
      const pa=historyPeriod(a,state.planningHistory[a]||{}), pb=historyPeriod(b,state.planningHistory[b]||{});
      return pb.start-pa.start;
    });

  // La liste native reste cachée : l'historique est présenté sous forme de blocs.
  el.innerHTML=keys.length?keys.map(k=>{
    const h=historyData(k), age=generatedAge(h), label=historyPeriodLabel(k,h);
    return `<option value="${escapeHtml(k)}">${escapeHtml(label)} — Généré${age?` il y a ${escapeHtml(age)}`:""}</option>`;
  }).join(""):``;

  if(!blocks)return;

  // On affiche toujours les 12 mois. Le mois courant est marqué comme tel.
  // Les anciens plannings sont ensuite regroupés : MOIS -> ANNÉE -> PLANNING.
  const monthMap={};
  for(let m=0;m<12;m++)monthMap[m]=[];
  keys.forEach(k=>{
    const h=historyData(k), hp=historyPeriod(k,h);
    if(!hp?.start)return;
    monthMap[hp.start.getMonth()].push({key:k,year:hp.start.getFullYear(),h,current:false});
  });

  // Ajouter le planning actuel dans le bloc du mois actuel, s'il existe.
  const currentPer=getPeriodDates();
  const hasCurrent=Object.values(state.planning||{}).flat().length>0 ||
    (state.fixed||[]).some(x=>x.date>=iso(currentPer.start)&&x.date<=iso(currentPer.end)&&["G","F"].includes(x.type));
  if(hasCurrent){
    monthMap[currentPer.start.getMonth()].push({key:currentKey,year:currentPer.start.getFullYear(),h:state.planningHistory?.[currentKey]||null,current:true});
  }

  const currentMonth=new Date().getMonth();
  const monthHtml=Object.keys(monthMap).map(m=>{
    const items=monthMap[m];
    const isCurrentMonth=Number(m)===currentMonth;
    const id=`history-month-${m}`;
    const years=[...new Set(items.map(x=>x.year))].sort((a,b)=>b-a);
    const yearHtml=years.map(y=>{
      const yi=items.filter(x=>x.year===y).sort((a,b)=>generatedTimestamp(b.h)-generatedTimestamp(a.h));
      const yid=`${id}-year-${y}`;
      const list=yi.map(x=>{
        const age=x.current
          ? generatedAgeText(x.h?.generatedAtISO || x.h?.generatedAt)
          : generatedAge(x.h);
        const label=x.current ? `${months[currentPer.start.getMonth()]} ${currentPer.start.getFullYear()}` : historyPeriodLabel(x.key,x.h);
        const viewBtn=x.current
          ? `<button type="button" class="history-view-btn" title="Voir le tableau" aria-label="Voir le tableau" onclick="event.stopPropagation();selectHistoryBlock('',true)">👁</button>`
          : `<button type="button" class="history-view-btn" title="Voir le tableau" aria-label="Voir le tableau" onclick="event.stopPropagation();selectHistoryBlock('${escapeHtml(x.key)}',false)">👁</button>`;
        const deleteBtn=x.current
          ? `<button type="button" class="history-delete-btn" title="Supprimer ce planning" aria-label="Supprimer ce planning" onclick="event.stopPropagation();deleteDisplayedPlanning()">🗑</button>`
          : `<button type="button" class="history-delete-btn" title="Supprimer ce planning" aria-label="Supprimer ce planning" onclick="event.stopPropagation();deleteHistoryBlock('${escapeHtml(x.key)}')">🗑</button>`;
        return `<div class="history-planning-card ${x.current?'history-current-card':''}" role="button" tabindex="0" onclick="selectHistoryBlock('${escapeHtml(x.key)}',${x.current})" onkeydown="if(event.key==='Enter'||event.key===' ')selectHistoryBlock('${escapeHtml(x.key)}',${x.current})">
          <div class="history-planning-info"><strong>${escapeHtml(label)}${x.current?' — Actuel':''}</strong><span>${age?`Généré ${x.current?'':'il y a '}${escapeHtml(age)}`:(x.current?'Planning actuel':'Date de génération inconnue')}</span></div>
          <div class="history-actions">${viewBtn}${deleteBtn}</div>
        </div>`;
      }).join("");
      return `<div class="history-year-card"><button type="button" class="history-year-button" onclick="toggleHistoryPanel('${yid}')"><strong>${y}</strong><span>${yi.length} planning${yi.length>1?'s':''}</span><b>▾</b></button><div id="${yid}" class="history-planning-list" hidden>${list}</div></div>`;
    }).join("");

    // Le bloc du mois reste visible même s'il n'a aucun planning.
    const count=items.length;
    return `<div class="history-month-card ${isCurrentMonth?'history-current-month':''}">
      <button type="button" class="history-month-button" onclick="toggleHistoryPanel('${id}')">
        <strong>${months[Number(m)]}</strong>
        <span>${count} planning${count>1?'s':''}${isCurrentMonth?' · actuel':''}</span><b>▾</b>
      </button>
      <div id="${id}" class="history-year-panel" hidden>${yearHtml || `<div class="history-empty">Aucun planning pour ce mois.</div>`}</div>
    </div>`;
  }).join("");

  blocks.innerHTML=monthHtml;
}
function generatedTimestamp(h){
  if(!h)return 0;
  const v=h.generatedAtISO||h.generatedAt;
  if(typeof v==='number')return v;
  const t=Date.parse(v||'');
  return Number.isFinite(t)?t:0;
}
function historyPeriodLabel(k,h){
  const hp=historyPeriod(k,h);
  if(!hp?.start)return "Planning";
  if(hp.type==="month")return `${months[hp.start.getMonth()]} ${hp.start.getFullYear()}`;
  if(hp.type==="week")return `Semaine du ${iso(hp.start)} au ${iso(hp.end)}`;
  return `${hp.weeks||1} semaines : ${iso(hp.start)} au ${iso(hp.end)}`;
}
function toggleHistoryPanel(id){
  const el=document.getElementById(id);
  if(el)el.hidden=!el.hidden;
}
function selectHistoryBlock(k,isCurrent=false){
  // Pour un compte Médecin, le bloc « Actuel » doit utiliser le planning
  // historique validé du mois courant, et jamais state.planning (qui peut être
  // un brouillon). C'est notamment le cas de Septembre 2026.
  if(!isAdmin() && (isCurrent || k===getPeriodDates().key)){
    const validatedKey=doctorCurrentMonthHistoryKey();
    if(validatedKey){
      const sel=document.getElementById("doctorHistorySelect");
      if(sel)sel.value=validatedKey;
      showDoctorValidatedPlanning(validatedKey);
    }else{
      const notice=document.getElementById("doctorPlanningNotice");
      if(notice)notice.textContent="Aucun planning validé n'est disponible pour le mois actuel.";
      document.getElementById("planningTable").innerHTML="";
      document.getElementById("planningTitle").textContent="";
    }
    return;
  }
  if(isCurrent || k===getPeriodDates().key){
    showCurrentPlanning();
    return;
  }
  const sel=document.getElementById("historySelect");
  if(sel)sel.value=k||"";
  showHistoryPlanningByKey(k);
}
function deleteHistoryBlock(k){
  if(!isAdmin())return alert("Accès réservé aux administrateurs.");
  if(!k)return alert("Planning introuvable.");
  const h=state.planningHistory?.[k];
  if(!h)return alert("Planning introuvable.");
  const label=historyPeriodLabel(k,h);
  if(!confirm(`Supprimer l'ancien planning ${label} ?`))return;
  delete state.planningHistory[k];
  delete state.confirmedPlanningKeys[k];
  if(activePlanningView?.type==="history" && activePlanningView.key===k)activePlanningView={type:"current",key:null};
  rebuildRecoveryLedger();
  save();
  renderAll();
  renderPlanning();
}
function showHistoryPlanningByKey(k){
  if(!k)return;
  const h=historyData(k);
  if(!h)return alert("Ancien planning introuvable.");
  activePlanningView={type:"history",key:k};
  renderPlanning(h,historyPeriod(k,h));
}

function doctorCurrentMonthHistoryKey(){
  const per=getPeriodDates();
  const currentKey=per.key;
  // 1) Cas normal : le planning validé est enregistré sous la clé exacte du mois.
  if(isHistoryConfirmed(currentKey,state.planningHistory?.[currentKey])) return currentKey;

  // 2) Cas des anciennes versions : le planning de septembre peut avoir été
  // archivé avec un suffixe (ex. _historique_...) ou avec une autre clé, tout
  // en conservant la période septembre 2026 dans h.period. On recherche donc
  // par PERIODE, pas uniquement par nom de clé.
  const candidates=Object.keys(state.planningHistory||{}).filter(k=>{
    const h=state.planningHistory[k];
    if(!isHistoryConfirmed(k,h))return false;
    const hp=historyPeriod(k,h);
    return hp?.start && hp?.end &&
      hp.start.getFullYear()===per.start.getFullYear() &&
      hp.start.getMonth()===per.start.getMonth() &&
      hp.end.getFullYear()===per.end.getFullYear() &&
      hp.end.getMonth()===per.end.getMonth();
  });
  return candidates.sort((a,b)=>generatedTimestamp(state.planningHistory[b])-generatedTimestamp(state.planningHistory[a]))[0]||null;
}

function confirmedHistoryKeyForCurrentMonth(){
  // Même mécanisme que les autres historiques : on privilégie l'enregistrement
  // historique validé qui contient réellement les données du mois.
  const k=doctorCurrentMonthHistoryKey();
  if(k && isHistoryConfirmed(k,state.planningHistory?.[k])) return k;
  return null;
}
function renderDoctorHistorySelect(){
  const el=document.getElementById("doctorHistorySelect");if(!el)return;
  const previous=activePlanningView?.type==="history"?activePlanningView.key:el.value;
  // IMPORTANT : utiliser historyData() ici. Certaines anciennes sauvegardes
  // (notamment des plannings mensuels plus anciens) stockent directement le
  // planning dans planningHistory[key] au lieu de l'envelopper dans {planning:...}.
  // Si on lit h.planning directement, le planning peut être validé mais ne pas
  // apparaître dans le compte Médecin.
  const hasData=k=>{
    const h=historyData(k);
    if(!h)return false;
    return Object.values(h.planning||{}).flat().length>0 ||
      (Array.isArray(h.fixed)&&h.fixed.length>0);
  };
  const keys=Object.keys(state.planningHistory||{})
    .filter(k=>isHistoryConfirmed(k,state.planningHistory[k])&&hasData(k))
    .sort((a,b)=>historyPeriod(b,state.planningHistory[b]).start-historyPeriod(a,state.planningHistory[a]).start);
  el.innerHTML=keys.length?keys.map(k=>{const h=historyData(k),hp=historyPeriod(k,h);const label=hp?.type==="month"?`${months[hp.start.getMonth()]} ${hp.start.getFullYear()}`:hp?.type==="week"?`Semaine du ${iso(hp.start)} au ${iso(hp.end)}`:`${hp?.weeks||1} semaines : ${iso(hp.start)} au ${iso(hp.end)}`;return `<option value="${escapeHtml(k)}">${label} — Validé</option>`;}).join(""):``;
  const currentDoctorKey=doctorCurrentMonthHistoryKey();
  if(currentDoctorKey && keys.includes(currentDoctorKey) && (!previous || !keys.includes(previous))) el.value=currentDoctorKey;
  else if(previous && keys.includes(previous)) el.value=previous;
}
function showDoctorValidatedPlanning(keyOverride){
  if(isAdmin())return;
  // Pour le médecin, le mois actuel doit lui aussi être consultable dès qu'il
  // a été validé. Certaines anciennes versions ne remettaient pas toujours
  // la sélection sur la clé exacte du mois courant.
  let k=keyOverride || document.getElementById("doctorHistorySelect")?.value || "";
  if(!k) k=confirmedHistoryKeyForCurrentMonth()||"";
  const notice=document.getElementById("doctorPlanningNotice");
  // Pour septembre (mois courant), utiliser exactement la même source historique
  // que pour les autres mois : l'enregistrement validé dans planningHistory.
  // Cela évite de retomber sur state.planning, qui peut être un brouillon courant.
  const h=k?historyData(k):null;
  if(!k || !h || !isHistoryConfirmed(k,state.planningHistory?.[k])){
    activePlanningView={type:"doctorValidated",key:null};
    if(notice)notice.textContent="Aucun planning validé n'est disponible.";
    document.getElementById("planningTable").innerHTML="";
    document.getElementById("planningTitle").textContent="";
    return;
  }
  activePlanningView={type:"history",key:k};
  const sel=document.getElementById("doctorHistorySelect");
  if(sel && [...sel.options].some(o=>o.value===k)) sel.value=k;
  if(notice)notice.textContent="Consultation seule : ce planning est validé et ne peut pas être modifié.";
  renderPlanning(h,historyPeriod(k,h));
}
function addDoctorRecovery(){
  const ses=currentSession();
  if(!ses || ses.role==="admin")return alert("Cette saisie est réservée aux comptes Médecin.");
  const date=document.getElementById("doctorRecoveryDate")?.value;
  if(!date)return alert("Sélectionnez la date de récupération.");
  if(!ses.doctor || !state.doctors.some(d=>d.name===ses.doctor))return alert("Médecin associé au compte introuvable.");
  if(state.absences.some(a=>a.doctor===ses.doctor&&a.start===date&&a.end===date&&["RÉCUP","RECUP"].includes(String(a.type).toUpperCase())))
    return alert("Une récupération est déjà enregistrée à cette date.");
  state.absences.push({doctor:ses.doctor,start:date,end:date,type:"RÉCUP",createdAt:Date.now(),createdBy:ses.username});
  save();
  const m=document.getElementById("doctorRecoveryMessage");
  if(m){m.textContent=`Récupération du ${date} enregistrée pour ${ses.doctor}.`;m.style.display="block";}
  document.getElementById("doctorRecoveryDate").value="";
  renderAll();
  if(!isAdmin())openTab("planning");
}
function escapeHtml(v){return String(v??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;");}

function renderPlanning(source,periodOverride){
  // SECURITE : un compte Médecin ne peut jamais afficher le planning courant/brouillon.
  // Il ne peut afficher qu'un historique explicitement validé par l'administrateur.
  const ses=currentSession();
  if(ses?.role!=="admin") {
    // Pour un médecin, si aucun source n'est fourni, recharger systématiquement
    // le planning historique actuellement sélectionné (s'il est validé).
    if(!source && activePlanningView?.type==="history" && activePlanningView.key){
      const selected=state.planningHistory?.[activePlanningView.key];
      if(isHistoryConfirmed(activePlanningView.key,selected)){
        source=selected;
        periodOverride=periodOverride||historyPeriod(activePlanningView.key,selected);
      }
    }
    const selectedKey = activePlanningView?.type==="history" ? activePlanningView.key : null;
    // Vérifier la validation sur l'enregistrement original de l'historique.
    // Certaines anciennes versions (notamment pour septembre) stockent
    // directement le planning dans planningHistory[key] : historyData()
    // fabrique alors un wrapper qui ne contient pas le champ confirmed.
    const validationRecord = selectedKey ? state.planningHistory?.[selectedKey] : null;
    if(!source || !isHistoryConfirmed(selectedKey,validationRecord)){
      activePlanningView={type:"doctorValidated",key:null};
      const notice=document.getElementById("doctorPlanningNotice");
      if(notice)notice.textContent="Aucun planning validé à afficher. Les plannings ne sont accessibles qu'après validation par l'administrateur.";
      const table=document.getElementById("planningTable");
      if(table)table.innerHTML="";
      const title=document.getElementById("planningTitle");
      if(title)title.textContent="";
      return;
    }
  }
  const per=periodOverride||getPeriodDates();
  // IMPORTANT : si le planning courant est déjà validé, son affichage repose sur
  // le snapshot validé. Les nouveaux congés/indisponibilités saisis ensuite ne
  // doivent jamais masquer les gardes de ce planning. Ils sont pris en compte
  // uniquement lors d'une nouvelle génération.
  const currentValidated=(!source && state.planningHistory?.[per.key]?.confirmed && state.planningHistory?.[per.key]?.planning)
    ? state.planningHistory[per.key] : null;
  let src=source||currentValidated||{planning:state.planning,fixed:state.fixed,absences:state.absences};
  if(currentValidated && !source){ src={...currentValidated,__validatedSnapshot:true}; }
  const editable=!source && isAdmin();
  const hasPlanning=Object.values(src.planning||{}).flat().length>0 || (src.fixed||[]).some(x=>x.date>=iso(per.start)&&x.date<=iso(per.end)&&["G","F"].includes(x.type));
  if(!hasPlanning){
    document.getElementById("planningTable").innerHTML="<div class='hint' style='padding:15px'>Aucun planning à afficher.</div>";
    document.getElementById("planningTitle").textContent="";
    renderHistorySelect(); renderDoctorHistorySelect(); renderStats();
    return;
  }
  let html="<table class='schedule'><thead><tr><th class='name'>Médecin</th>";
  for(const dt of datesBetween(per.start,per.end)){let h=(src.fixed||[]).some(x=>x.date===iso(dt)&&x.type==="F");let hc=h?"holiday":(isWeekend(dt)?"weekend":"");html+=`<th class='${hc}'><div>${["D","L","M","M","J","V","S"][dt.getDay()]}</div><div class='dayhead'>${String(dt.getDate()).padStart(2,"0")}/${String(dt.getMonth()+1).padStart(2,"0")}</div></th>`}
  html+="</tr></thead><tbody>";
  state.doctors.filter(d=>d.active).forEach(doc=>{html+=`<tr><td class='name'><b>${escapeHtml(doc.name)}</b></td>`;for(const dt of datesBetween(per.start,per.end)){let v=cellFor(doc.name,dt,src);let attrs=editable?`data-edit-name="${encodeURIComponent(doc.name)}" data-edit-date="${iso(dt)}"`:'';html+=`<td class='${cls(v)}' title='Cliquer pour modifier' ${attrs}>${editable?`<button type="button" class="cellEditBtn" onclick="editCell(this)">${escapeHtml(v||"")}</button>`:escapeHtml(v||"")}</td>`}html+="</tr>"});
  html+="</tbody></table>";document.getElementById("planningTable").innerHTML=html;
  const currentHist=state.planningHistory?.[per.key]; const generatedAge=generatedAgeText(currentHist?.generatedAtISO || currentHist?.generatedAt); document.getElementById("planningTitle").textContent=(per.type==="month"?`${months[per.start.getMonth()]} ${per.start.getFullYear()}`:per.type==="week"?`Semaine du ${iso(per.start)} au ${iso(per.end)}`:`${per.weeks} semaines : ${iso(per.start)} au ${iso(per.end)}`)+(generatedAge?` — Généré ${generatedAge}`:"");
  // Pour un compte Médecin, conserver impérativement le planning historique sélectionné.
  // Ne jamais remplacer activePlanningView par "current" après l'affichage d'un planning validé.
  if(ses?.role!=="admin" && activePlanningView?.type==="history" && activePlanningView.key){
    activePlanningView={type:"history",key:activePlanningView.key};
  } else if(ses?.role==="admin" && !source){
    activePlanningView={type:"current",key:null};
  }
  renderHistorySelect(); renderDoctorHistorySelect(); renderStats();
}
function editCell(btn){
  if(!isAdmin())return alert("Un compte Médecin peut uniquement consulter les plannings validés. Modification interdite.");
  const td=btn.closest("td[data-edit-name][data-edit-date]");
  if(!td)return;
  const name=decodeURIComponent(td.dataset.editName);
  const date=td.dataset.editDate;
  const current=cellFor(name,date,{planning:state.planning,fixed:state.fixed,absences:state.absences})||"";
  const value=window.prompt("Modifier la garde de "+name+" le "+date+"\n\nValeurs : J, N, G, F, CONGÉ, RÉCUP, INDISP\nLaissez vide pour effacer.",current);
  if(value===null)return;
  applyInlineEdit(name,date,value);
}
function startInlineEdit(td,name,date){ editCell(td.querySelector(".cellEditBtn")); }
function applyInlineEdit(name,date,raw){
  if(!isAdmin())return alert("Modification du planning réservée à l'administrateur.");
  let value=String(raw||"").trim().toUpperCase();
  const allowed=["","J","N","G","F","CONGÉ","CONGE","RECUP","RÉCUP","INDISP","INDISPONIBLE"];
  if(!allowed.includes(value)){alert("Valeur non reconnue. Utilisez J, N, G, F, CONGÉ, RÉCUP, INDISP ou vide.");return;}
  if(value==="CONGE")value="CONGÉ";
  if(value==="RECUP")value="RÉCUP";
  if(value==="INDISPONIBLE")value="INDISP";

  const dt=dateFromISO(date), doc=docByName(name);
  if(!doc){alert("Médecin introuvable.");return;}

  const oldFixedIndex=state.fixed.findIndex(x=>x.doctor===name&&x.date===date);
  const oldPlanningEntries=(state.planning[date]||[]).filter(x=>x.doctor===name);
  const otherPlanningEntries=(state.planning[date]||[]).filter(x=>x.doctor!==name);
  const sameDayFixed=(state.fixed||[]).filter(x=>x.date===date&&x.doctor!==name);
  const sameDoctorOtherFixed=(state.fixed||[]).find(x=>x.date===date&&x.doctor===name);

  // IMPORTANT : effectuer tous les contrôles AVANT de modifier le planning.
  // Ainsi une tentative refusée ne dévalide ni ne détruit un planning confirmé.
  if(["J","N"].includes(value)){
    // Un seul J et un seul N sont autorisés par date, quel que soit le type
    // de saisie (génération, garde fixe ou modification manuelle).
    const duplicateSameType=[...otherPlanningEntries,...sameDayFixed]
      .filter(x=>x && x.type===value);
    if(duplicateSameType.length){
      alert(`CONFLIT : une garde ${value} est déjà attribuée le ${date} à ${duplicateSameType[0].doctor}.\\n\\nUn seul ${value} est autorisé par jour.`);
      renderPlanning();return;
    }
    // Un même médecin ne peut jamais cumuler deux postes le même jour.
    if(sameDoctorOtherFixed){
      alert(`CONFLIT : ${name} a déjà une garde ${sameDoctorOtherFixed.type} le ${date}.\\n\\nUn médecin ne peut pas avoir deux gardes le même jour.`);
      renderPlanning();return;
    }

    // Remplacer la garde J par N (ou inversement) ne doit pas être considéré
    // comme une seconde garde du même médecin : on teste sur le planning
    // temporairement débarrassé de son ancienne cellule.
    const planningForCheck={...state.planning,[date]:otherPlanningEntries};
    if(!canAssign(doc,dt,value,planningForCheck)){
      alert(`CONFLIT : la garde ${value} de ${name} le ${date} ne respecte pas les contraintes de garde, congé, indisponibilité, récupération ou repos.`);
      renderPlanning();return;
    }
  }else if(["G","F"].includes(value)){
    const problem=validateFixedGuard(name,date);
    if(problem){alert(problem);renderPlanning();return;}
    if(sameDoctorOtherFixed){
      alert(`CONFLIT : ${name} a déjà une garde ${sameDoctorOtherFixed.type} le ${date}.`);
      renderPlanning();return;
    }
  }

  // Toute modification manuelle réellement acceptée d'un planning déjà validé
  // annule sa validation. Une tentative en conflit ci-dessus ne le fait jamais.
  const per=getPeriodDates();
  const currentHistory=state.planningHistory?.[per.key];
  if(currentHistory?.confirmed){
    undoConfirmedPeriod(currentHistory);
    const archiveKey=`${per.key}_historique_${Date.now()}`;
    state.planningHistory[archiveKey]={...JSON.parse(JSON.stringify(currentHistory)),includeInEquity:false,confirmed:false,label:`Généré le ${currentHistory.generatedAt||'date inconnue'}`};
    currentHistory.confirmed=false;
    currentHistory.includeInEquity=false;
    delete state.confirmedPlanningKeys[per.key];
  }

  // Maintenant seulement, appliquer la modification.
  state.absences=state.absences.filter(a=>!(a.doctor===name&&a.start===date&&a.end===date));
  state.planning[date]=(state.planning[date]||[]).filter(x=>x.doctor!==name);
  if(oldFixedIndex>=0 && !["G","F"].includes(value)) state.fixed.splice(oldFixedIndex,1);

  if(["G","F"].includes(value)){
    if(oldFixedIndex>=0)state.fixed[oldFixedIndex].type=value;
    else state.fixed.push({doctor:name,date,type:value,createdAt:Date.now()});
  }else if(["J","N"].includes(value)){
    state.planning[date].push({doctor:name,type:value,date});
  }else if(["CONGÉ","RÉCUP","INDISP"].includes(value)){
    state.absences.push({doctor:name,start:date,end:date,type:value});
  }
  save(); renderAll();
}
function manualEditPlanning(){alert("Pour modifier le planning, cliquez directement sur la cellule souhaitée.");}
function imageToDataURL(file,maxW,maxH,quality){return new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>{const img=new Image();img.onload=()=>{const scale=Math.min(1,maxW/img.width,maxH/img.height);const c=document.createElement("canvas");c.width=Math.max(1,Math.round(img.width*scale));c.height=Math.max(1,Math.round(img.height*scale));c.getContext("2d").drawImage(img,0,0,c.width,c.height);resolve(c.toDataURL("image/jpeg",quality));};img.onerror=reject;img.src=reader.result;};reader.onerror=reject;reader.readAsDataURL(file);});}
async function saveAppearance(){
  state.settings=Object.assign({title:"Tox-Garde",subtitle:"Simple. Équitable. Intelligent",appBg:"#f5f7fb",headerBg:"#3f83c5",cardBg:"#ffffff",bgImage:"",appLogo:"",centerLogo:"",homeImage:""},state.settings||{});
  state.settings.title=(document.getElementById("appTitle").value||"Tox-Garde").trim(); state.settings.subtitle=(document.getElementById("appSubtitle").value||"Simple. Équitable. Intelligent").trim();
  state.settings.appBg=document.getElementById("appBg").value; state.settings.headerBg=document.getElementById("headerBg").value; state.settings.cardBg=document.getElementById("cardBg").value;
  try{
    const jobs=[['appLogoFile','appLogo',900,900,.82],['centerLogoFile','centerLogo',900,900,.82],['homeImageFile','homeImage',1600,900,.78],['bgImageFile','bgImage',1600,1000,.78]];
    for(const [id,key,mw,mh,q] of jobs){const f=document.getElementById(id)?.files?.[0];if(f)state.settings[key]=await imageToDataURL(f,mw,mh,q);}
    persistAppearance();
  }catch(e){alert("Impossible d’enregistrer une image. Choisissez une image plus légère.");}
}
function persistAppearance(){try{save();applyAppearance();}catch(e){alert("Le stockage du navigateur est insuffisant. Choisissez des images plus légères.");}}
function removeOneImage(key,fileId){state.settings=Object.assign({appLogo:"",centerLogo:"",homeImage:"",bgImage:""},state.settings||{});state.settings[key]="";const f=document.getElementById(fileId);if(f)f.value="";persistAppearance();}
function removeImages(){state.settings.appLogo="";state.settings.centerLogo="";state.settings.homeImage="";state.settings.bgImage="";["appLogoFile","centerLogoFile","homeImageFile","bgImageFile"].forEach(id=>{const f=document.getElementById(id);if(f)f.value="";});persistAppearance();}
function resetAppearance(){state.settings={title:"Tox-Garde",subtitle:"Simple. Équitable. Intelligent",appBg:"#f5f7fb",headerBg:"#3f83c5",cardBg:"#ffffff",bgImage:"",appLogo:"",centerLogo:"",homeImage:""};save();applyAppearance();}
function applyAppearance(){
  state.settings=Object.assign({title:"Tox-Garde",subtitle:"Simple. Équitable. Intelligent",appBg:"#f5f7fb",headerBg:"#3f83c5",cardBg:"#ffffff",bgImage:"",appLogo:"",centerLogo:"",homeImage:""},state.settings||{});
  document.getElementById("appTitleDisplay").textContent=state.settings.title||"Tox-Garde"; document.getElementById("appSubtitleDisplay").textContent=state.settings.subtitle||"Simple. Équitable. Intelligent";
  document.getElementById("appTitle").value=state.settings.title||"Tox-Garde"; document.getElementById("appSubtitle").value=state.settings.subtitle||"Simple. Équitable. Intelligent"; document.getElementById("appBg").value=state.settings.appBg||"#f5f7fb"; document.getElementById("headerBg").value=state.settings.headerBg||"#3f83c5"; document.getElementById("cardBg").value=state.settings.cardBg||"#ffffff";
  const setImg=(id,src,statusId,yes,no)=>{const el=document.getElementById(id);if(el){el.src=src||"";el.classList.toggle("show",!!src);}const st=document.getElementById(statusId);if(st)st.textContent=src?yes:no;};
  setImg("appLogoDisplay",state.settings.appLogo,"appLogoStatus","Logo enregistré","Aucun logo"); setImg("centerLogoDisplay",state.settings.centerLogo,"centerLogoStatus","Logo enregistré","Aucun logo"); setImg("homeImageDisplay",state.settings.homeImage,"homeImageStatus","Image enregistrée","Aucune image");
  document.body.style.backgroundColor=state.settings.appBg||"#f5f7fb"; document.body.style.backgroundImage=state.settings.bgImage?`linear-gradient(#ffffffaa,#ffffffaa),url(${state.settings.bgImage})`:"none"; document.body.style.backgroundSize="cover"; document.body.style.backgroundAttachment="fixed"; document.body.style.backgroundPosition="center";
  document.querySelector("header").style.background=state.settings.headerBg||"#3f83c5"; document.querySelectorAll(".card").forEach(x=>x.style.backgroundColor=state.settings.cardBg||"#ffffff"); const st=document.getElementById("bgImageStatus");if(st)st.textContent=state.settings.bgImage?"Image enregistrée":"Aucune image";
}
function historyPeriod(key,h){
  // Les clés mensuelles peuvent être archivées avec un suffixe (_historique_...).
  const m=String(key||'').match(/^month_(\d{4})_(\d{2})(?:_|$)/);
  if(m){
    const y=+m[1], mo=+m[2]-1;
    return {type:"month",start:new Date(y,mo,1),end:new Date(y,mo,daysInMonth(y,mo)),weeks:undefined,key};
  }
  const pd=h?.period||{};
  const start=pd.start?dateFromISO(pd.start):getPeriodDates().start;
  const end=pd.end?dateFromISO(pd.end):getPeriodDates().end;
  return {type:pd.type||"week",start,end,weeks:pd.weeks||Math.max(1,Math.round((end-start)/86400000+1)/7),key};
}
function showHistoryPlanning(){
  const k=document.getElementById("historySelect")?.value;
  if(!k)return alert("Aucun ancien planning disponible.");
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
  save();
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
    save();
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
  save();
  renderAll();
  renderPlanning();
}

function showCurrentPlanning(){activePlanningView={type:"current",key:null};renderPlanning();}
function getExportContext(){
  if(activePlanningView?.type==="history" && activePlanningView.key){
    const h=historyData(activePlanningView.key);
    if(h){
      const per=historyPeriod(activePlanningView.key,h);
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
</style><style>
.history-groups{display:flex;flex-direction:column;gap:10px;margin-top:10px}
.history-month-group{border:1px solid #dbe3ec;border-radius:10px;background:#f8fafc;padding:10px}
.history-month-title{font-weight:700;font-size:15px;margin-bottom:8px}
.history-month-items{display:flex;flex-wrap:wrap;gap:8px}
.history-block{display:flex;flex-direction:column;align-items:flex-start;min-width:170px;padding:10px 12px;border:1px solid #cfd9e5;border-radius:9px;background:#fff;cursor:pointer;text-align:left}
.history-block:hover{box-shadow:0 2px 8px rgba(0,0,0,.08);transform:translateY(-1px)}
.history-block strong{font-size:14px}.history-block span{font-size:12px;color:#667085;margin-top:4px}
.history-empty{padding:12px;border:1px dashed #cbd5e1;border-radius:9px;color:#667085;background:#f8fafc}
</style>

<style>
.history-groups{display:grid;grid-template-columns:repeat(4,minmax(130px,1fr));gap:10px;margin-top:10px}
.history-month-card{border:1px solid #dbe3ec;border-radius:12px;background:#f8fafc;overflow:hidden}
.history-month-button,.history-year-button{width:100%;border:0;background:#fff;cursor:pointer;text-align:left;padding:12px;display:grid;grid-template-columns:1fr auto;gap:3px 8px}
.history-month-button strong{font-size:15px}.history-month-button span,.history-year-button span{font-size:12px;color:#667085}.history-month-button b,.history-year-button b{grid-row:1/3;grid-column:2;align-self:center}
.history-year-panel{padding:7px;background:#f8fafc;border-top:1px solid #e5e7eb}.history-year-card{border:1px solid #dbe3ec;border-radius:9px;background:#fff;margin-bottom:7px;overflow:hidden}.history-year-card:last-child{margin-bottom:0}
.history-year-button{padding:9px 10px;background:#f9fafb}.history-planning-list{padding:7px;display:flex;flex-direction:column;gap:6px}.history-planning-card{display:flex;flex-direction:column;align-items:flex-start;border:1px solid #cfd9e5;border-radius:8px;background:#fff;cursor:pointer;text-align:left;padding:9px 10px}.history-planning-card:hover{box-shadow:0 2px 8px rgba(0,0,0,.08)}.history-planning-card strong{font-size:13px}.history-planning-card span{font-size:11px;color:#667085;margin-top:3px}
@media(max-width:800px){.history-groups{grid-template-columns:repeat(3,minmax(110px,1fr))}}
@media(max-width:560px){.history-groups{grid-template-columns:repeat(2,minmax(120px,1fr))}}
</style>
<style>
.history-month-head{display:flex;align-items:stretch}.history-month-button{flex:1}
.history-planning-card{position:relative;display:flex!important;flex-direction:row!important;align-items:center!important;justify-content:space-between;gap:10px;width:100%;box-sizing:border-box}
.history-planning-info{display:flex;flex-direction:column;align-items:flex-start;min-width:0}.history-planning-info strong{font-size:13px}.history-planning-info span{font-size:11px;color:#667085;margin-top:3px}
.history-actions{display:flex;align-items:center;gap:5px;flex:0 0 auto}.history-view-btn,.history-delete-btn{flex:0 0 auto;border:1px solid #f1b5b5;border-radius:7px;background:#fff5f5;color:#b42318;padding:6px 8px;font-size:11px;font-weight:700;cursor:pointer}.history-view-btn{border:1px solid #b8d2ea;border-radius:7px;background:#f4f9ff;color:#24608f;padding:6px 8px;font-size:11px;font-weight:700;cursor:pointer}.history-view-btn:hover{background:#e8f3ff}.history-delete-btn:hover{background:#fee4e2}.history-delete-btn:disabled{opacity:.55;cursor:not-allowed}
.history-current-card{border-color:#8bb8df;background:#f7fbff}.history-current-empty{cursor:default}
@media(max-width:560px){.history-planning-card{flex-direction:row!important;align-items:center!important}.history-actions{align-self:center}.history-view-btn,.history-delete-btn{width:22px!important}}
</style>

<style>
/* Tableau de bord : un seul bloc et cartes uniformes */
.dashboard-card{width:100%;}
.dashboard-card #dashStats{grid-template-columns:repeat(4,minmax(0,1fr));margin-bottom:14px;}
.dashboard-card .stat{min-height:82px;display:flex;flex-direction:column;justify-content:center;}
.today-dashboard{border-top:1px solid var(--line);padding-top:12px;}
.today-dashboard-title{font-size:14px;font-weight:700;margin-bottom:8px;}
.today-guards-list{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px;width:100%;}
.today-guard,.today-status-box{min-height:68px;display:flex;flex-direction:column;justify-content:center;box-sizing:border-box;}
.today-guard{text-align:center;}
.today-status-grid{grid-template-columns:repeat(3,minmax(0,1fr));}
@media(max-width:900px){.dashboard-card #dashStats{grid-template-columns:repeat(2,minmax(0,1fr));}.today-guards-list{grid-template-columns:repeat(2,minmax(0,1fr));}}
@media(max-width:600px){.dashboard-card #dashStats{grid-template-columns:1fr 1fr;}.today-guards-list,.today-status-grid{grid-template-columns:repeat(2,minmax(0,1fr));}.today-guard,.today-status-box{min-height:60px;padding:6px;}}
</style>
</head>
<body><div class="Section1">
<div class="header"><div class="center">CENTRE ANTIPOISON</div><h1>${esc(title)}</h1><h2>${esc(subtitle)}</h2></div>
<div class="period">Planning des gardes — ${esc(periodTitle)}</div>
<table class="planning"><thead><tr><th style="background:#eef2f7;border:1px solid #777;padding:3px;text-align:left;font-size:8pt">Médecin</th>${headers}</tr></thead><tbody>${body}</tbody></table>
</div>

</body></html>`;
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
  save();
  document.getElementById('genMessage').innerHTML='<div class="ok"><b>Planning validé.</b><br>Les récupérations ont été enregistrées et le score temporaire a été réinitialisé.</div>';
  renderAll();
}
function resetAllRecoveryScores(){
  if(!confirm("Remettre à 0 tous les scores de récupération ? Cette action ne supprimera pas les plannings ni l’historique."))return;
  state.recoveryLedger={};
  state.doctors.forEach(d=>{state.recoveryLedger[d.name]={solde:0,totalAcquise:0,totalPrise:0,updatedAt:new Date().toISOString()};});
  save();
  renderAll();
  alert("Tous les scores de récupération ont été remis à 0.");
}
function resetAllEquityCounters(){
  if(!confirm("Remettre à 0 tous les compteurs J, N, G et F ainsi que les scores de récupération ? Les plannings et l'historique resteront visibles."))return;
  const now=Date.now();
  state.equityResetAt=now;
  state.recoveryLedger={};
  state.doctors.forEach(d=>{state.recoveryLedger[d.name]={solde:0,totalAcquise:0,totalPrise:0,updatedAt:new Date(now).toISOString()};});
  save();
  renderAll();
  alert("Tous les compteurs J/N/G/F et les scores de récupération ont été remis à 0. Les plannings restent conservés pour consultation.");
}
function renderStats(){
  const el=document.getElementById("statsTable"); if(!el)return;
  const c=counts();
  // Un médecin ne voit que ses propres compteurs de récupération.
  // L'administrateur conserve la vue complète de l'équipe.
  const ses=currentSession();
  const doctorsToShow=isAdmin()?state.doctors:state.doctors.filter(d=>d.name===ses?.doctor);
  if(!isAdmin() && !doctorsToShow.length){
    el.innerHTML='<div class="hint">Aucune donnée de récupération disponible pour votre compte.</div>';
    return;
  }
  el.innerHTML=`<table><thead><tr><th>Médecin</th><th>J</th><th>N</th><th>G</th><th>F</th><th>Total gardes</th><th>Récup acquise</th><th>Récup prise</th><th>Récup restante</th></tr></thead><tbody>`+
    doctorsToShow.map(d=>{
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
  const dt=dateFromISO(today);
  const dateLabel=`${String(dt.getDate()).padStart(2,"0")}/${String(dt.getMonth()+1).padStart(2,"0")}/${dt.getFullYear()}`;

  // IMPORTANT : le tableau de bord regarde TOUJOURS la date réelle d'aujourd'hui,
  // indépendamment du mois actuellement régénéré ou affiché dans Planning.
  // On cherche le planning validé qui couvre aujourd'hui parmi tout l'historique.
  const histories=Object.entries(state.planningHistory||{})
    .map(([key,h])=>({key,h}))
    .filter(({key,h})=>isHistoryConfirmed(key,h))
    .map(({key,h})=>{
      const hp=historyPeriod(key,h);
      const startISO=h?.period?.start || (hp?.start?iso(hp.start):"");
      const endISO=h?.period?.end || (hp?.end?iso(hp.end):"");
      // Sécurité supplémentaire : si les dates de période sont absentes ou erronées,
      // on détermine la couverture d'aujourd'hui directement à partir des gardes enregistrées.
      const hasTodayEntry=Object.values(h?.planning||{}).flat().some(x=>x?.date===today) || (h?.fixed||[]).some(x=>x?.date===today);
      const coversByPeriod=!!(startISO&&endISO&&today>=startISO&&today<=endISO);
      return {key,h,startISO,endISO,covers: coversByPeriod || hasTodayEntry};
    })
    .filter(x=>x.covers)
    .sort((a,b)=>generatedTimestamp(b.h)-generatedTimestamp(a.h));

  // Le planning validé couvrant aujourd'hui est prioritaire.
  // Si aucun historique validé ne couvre aujourd'hui, utiliser le planning courant
  // seulement s'il couvre réellement aujourd'hui.
  const h=histories[0]?.h||null;
  const currentPer=getPeriodDates();
  const currentCoversToday=currentPer?.start&&currentPer?.end&&dt>=currentPer.start&&dt<=currentPer.end;
  const planningSource=h?.planning || (currentCoversToday ? (state.planning||{}) : {});
  const fixedSource=h ? (Array.isArray(h.fixed)?h.fixed:[]) : (currentCoversToday ? (state.fixed||[]) : []);

  const guards=[];
  (planningSource[today]||[]).filter(x=>["J","N"].includes(x.type)).forEach(x=>guards.push({doctor:x.doctor,type:x.type}));
  fixedSource.filter(x=>x.date===today&&["G","F"].includes(x.type)).forEach(x=>guards.push({doctor:x.doctor,type:x.type}));
  const seen=new Set();
  const unique=guards.filter(x=>{const k=x.doctor+"|"+x.type;if(seen.has(k))return false;seen.add(k);return true;});

  const activeToday=(state.absences||[]).filter(a=>{
    const st=dateFromISO(a.start),en=dateFromISO(a.end);
    return st&&en&&dt>=st&&dt<=en;
  });
  const byType=types=>{
    const names=new Set();
    return activeToday.filter(a=>types.includes(String(a.type||"").toUpperCase()))
      .map(a=>a.doctor).filter(Boolean).filter(n=>{if(names.has(n))return false;names.add(n);return true;});
  };
  const conges=byType(["CONGÉ","CONGE"]);
  const recups=byType(["RÉCUP","RECUP"]);
  const indisponibles=byType(["INDISPONIBLE","INDISP"]);

  const guardHtml=unique.length
    ? unique.map(x=>`<div class="today-guard"><small>Garde ${x.type}</small><b>${escapeHtml(x.doctor)}</b></div>`).join("")
    : `<div class="today-empty" style="grid-column:1/-1">Aucune garde J, N, G ou F enregistrée pour aujourd’hui (${dateLabel}).</div>`;

  const listHtml=(title,names,emptyLabel)=>`
    <div class="today-status-box">
      <div class="today-status-title">${title} <span>${names.length}</span></div>
      ${names.length?names.map(n=>`<div class="today-status-person">${escapeHtml(n)}</div>`).join(""):`<div class="today-status-empty">${emptyLabel}</div>`}
    </div>`;

  el.innerHTML=`
    <div class="today-dashboard-title">Aujourd’hui — ${dateLabel}</div>
    <div class="today-guards-list">${guardHtml}</div>
    <div class="today-status-grid">
      ${listHtml("Congé",conges,"Aucun médecin en congé")}
      ${listHtml("Récupération",recups,"Aucune récupération aujourd’hui")}
      ${listHtml("Indisponible",indisponibles,"Aucun médecin indisponible")}
    </div>`;
}
function formatEntryCreatedAt(ts){
  if(!ts)return "Non renseignée";
  const d=typeof ts==="number"?new Date(ts):new Date(ts);
  if(isNaN(d))return "Non renseignée";
  return d.toLocaleDateString("fr-FR",{day:"2-digit",month:"2-digit",year:"numeric"})+" "+d.toLocaleTimeString("fr-FR",{hour:"2-digit",minute:"2-digit"});
}
function renderAll(){let body=document.getElementById("teamBody");let c=counts();body.innerHTML=state.doctors.map((d,i)=>`<tr><td class='name'><input value="${d.name.replaceAll('"','&quot;')}" onchange="renameDoctor(${i},this.value)"></td><td><input type='checkbox' ${d.active?"checked":""} onchange="state.doctors[${i}].active=this.checked;save()"></td><td><select onchange="state.doctors[${i}].shift=this.value;save()"><option value="BOTH" ${d.shift==="BOTH"?"selected":""}>J + N</option><option value="J" ${d.shift==="J"?"selected":""}>J uniquement</option><option value="N" ${d.shift==="N"?"selected":""}>N uniquement</option></select></td><td><input type="checkbox" ${d.weekend!==false?"checked":""} onchange="state.doctors[${i}].weekend=this.checked;save()"></td><td>${c[d.name]?.J||0}</td><td>${c[d.name]?.N||0}</td><td>${c[d.name]?.G||0}</td><td>${c[d.name]?.F||0}</td><td>${c[d.name]?.total||0}</td><td><button class='btn small danger' onclick='delDoctor(${i})'>Supprimer</button></td></tr>`).join("");
let opts=state.doctors.map(d=>`<option>${d.name}</option>`).join("");document.getElementById("absDoctor").innerHTML=opts;
const mergedEntries=[...state.absences.map((a,i)=>({kind:"abs",index:i,doctor:a.doctor,start:a.start,end:a.end,type:a.type,createdAt:a.createdAt})),...state.fixed.map((x,i)=>({kind:"fixed",index:i,doctor:x.doctor,start:x.date,end:x.date,type:x.type,createdAt:x.createdAt}))].sort((a,b)=>(a.start||"").localeCompare(b.start||"")||a.doctor.localeCompare(b.doctor));
document.getElementById("absBody").innerHTML=mergedEntries.map(a=>{const ses=currentSession();const canDelete=ses?.role==="admin" || (a.kind==="abs" && a.doctor===ses?.doctor && !["G","F"].includes(a.type));const action=canDelete?`<button class="btn small danger" onclick='deleteMergedEntry("${a.kind}",${a.index})'>Supprimer</button>`:`<span class="hint">Consultation</span>`;return `<tr><td class="name">${a.doctor}</td><td>${a.start}</td><td>${a.end}</td><td>${["G","F"].includes(a.type)?`<span class="badge ${a.type==="G"?"bG":"bF"}">${a.type}</span>`:a.type}</td><td>${formatEntryCreatedAt(a.createdAt)}</td><td>${action}</td></tr>`}).join("");
document.getElementById("dashStats").innerHTML=`<div class='stat'><span>Médecins actifs</span><b>${state.doctors.filter(d=>d.active).length}</b></div>`;renderTodayGuards(); if(isAdmin())renderPlanning(); else {renderDoctorHistorySelect(); const k=activePlanningView?.type==="history"?activePlanningView.key:document.getElementById("doctorHistorySelect")?.value; if(k)showDoctorValidatedPlanning(k);}}
function renameDoctor(i,v){if(!isAdmin())return alert("Accès réservé aux administrateurs.");let old=state.doctors[i].name;v=v.trim();if(!v)return;state.doctors[i].name=v;state.absences.forEach(a=>{if(a.doctor===old)a.doctor=v});state.fixed.forEach(a=>{if(a.doctor===old)a.doctor=v});if(state.accounts){state.accounts.forEach(a=>{if(a.doctor===old)a.doctor=v})};save()}

/* ===== Authentification et rôles : couche ajoutée sans modifier les fonctions métier ===== */
const INITIAL_ADMIN={username:"Aissahabbouni",password:"202020",role:"admin",doctor:"Administrateur initial"};
if(!Array.isArray(state.accounts)) state.accounts=[];
const SUPABASE_URL="https://hoxaqlefmdctjyauwymj.supabase.co";
const SUPABASE_KEY="sb_publishable_pjXtMOX-VOFMhp8PkIQ7hQ_Fb011ct6";
function supaHeaders(extra={}){return Object.assign({"apikey":SUPABASE_KEY,"Content-Type":"application/json","Accept":"application/json"},extra)}
function cloudToken(){return localStorage.getItem("toxGardeCloudToken")||sessionStorage.getItem("toxGardeCloudToken")||""}
async function supaRpc(name,body){
  let r;
  try{
    r=await fetch(`${SUPABASE_URL}/rest/v1/rpc/${name}`,{
      method:"POST",
      headers:supaHeaders({"Prefer":"return=representation"}),
      body:JSON.stringify(body)
    });
  }catch(e){
    throw new Error(`Supabase inaccessible (${e?.message||"erreur réseau"}). Vérifiez l’URL du projet et votre connexion Internet.`);
  }
  const text=await r.text();
  let payload=null;
  try{payload=text?JSON.parse(text):null}catch(_){payload=text}
  if(!r.ok){
    const msg=typeof payload==="string"?payload:(payload?.message||payload?.hint||payload?.details||payload?.error||"Erreur Supabase");
    throw new Error(`Supabase ${r.status} — ${msg}`);
  }
  return payload;
}
async function cloudLoad(token){const rows=await supaRpc("toxgarde_get_data",{p_token:token});return rows}
function cleanStateForCloud(){const copy=JSON.parse(JSON.stringify(state));delete copy.accounts;return copy}
let cloudSaveQueue=Promise.resolve();
async function cloudSave(){
  const token=cloudToken();
  if(!token)return;
  const snapshot=cleanStateForCloud();
  cloudSaveQueue=cloudSaveQueue.catch(()=>{}).then(()=>supaRpc("toxgarde_save_data",{p_token:token,p_data:snapshot}));
  await cloudSaveQueue;
}
function authStorage(){return JSON.parse(localStorage.getItem("toxGardeSession")||"null")}
function persistSession(session,keep){(keep?localStorage:sessionStorage).setItem("toxGardeSession",JSON.stringify(session)); if(keep) sessionStorage.removeItem("toxGardeSession"); else localStorage.removeItem("toxGardeSession")}
function currentSession(){return JSON.parse(sessionStorage.getItem("toxGardeSession")||localStorage.getItem("toxGardeSession")||"null")}
function isAdmin(){return currentSession()?.role==="admin"}
function normalizeUser(v){return String(v||"").trim().toLowerCase()}
async function login(ev){
  ev.preventDefault();
  const u=document.getElementById("loginUsername").value.trim();
  const pw=document.getElementById("loginPassword").value;
  const err=document.getElementById("authError");
  err.style.display="none";
  try{
    const result=await supaRpc("toxgarde_login",{p_username:u,p_password:pw});
    const a=Array.isArray(result)?result[0]:result;
    if(!a?.token)throw new Error("Identifiant ou mot de passe incorrect.");

    const keep=document.getElementById("keepConnected").checked;

    // Vérifier la lecture des données AVANT d'enregistrer la session.
    // Cela évite de laisser une session partiellement connectée si un RPC suivant échoue.
    let remote;
    try{
      remote=await cloudLoad(a.token);
    }catch(e){
      throw new Error(`Connexion au compte réussie, mais récupération des données Tox-Garde impossible. ${e.message}`);
    }

    // La session n'est mémorisée qu'après authentification + lecture cloud réussies.
    persistSession({username:a.username,role:a.role,doctor:a.doctor||""},keep);
    (keep?localStorage:sessionStorage).setItem("toxGardeCloudToken",a.token);
    if(keep)sessionStorage.removeItem("toxGardeCloudToken");
    else localStorage.removeItem("toxGardeCloudToken");
    document.getElementById("loginPassword").value="";
    localStorage.setItem("toxGardeRememberedUsername",u);

    if(remote&&typeof remote==="object"&&Object.keys(remote).length){
      state=Object.assign(state,remote);
      state.accounts=[];
    }else if(a.role==="admin"){
      const localAccounts=Array.isArray(state.accounts)?JSON.parse(JSON.stringify(state.accounts)):[];
      for(const acc of localAccounts){
        if(acc?.username&&acc?.password&&normalizeUser(acc.username)!==normalizeUser(INITIAL_ADMIN.username)){
          await supaRpc("toxgarde_upsert_account",{
            p_token:a.token,
            p_doctor:acc.doctor||"",
            p_username:acc.username,
            p_password:acc.password,
            p_role:acc.role||"doctor"
          });
        }
      }
      await cloudSave();
    }

    state.accounts=[];
    applyAuthUI();
  }catch(e){
    console.error("Tox-Garde / connexion Supabase :",e);
    err.textContent=e?.message||"Connexion impossible.";
    err.style.display="block";
  }
}
async function logout(){const token=cloudToken();if(token){try{await supaRpc("toxgarde_logout",{p_token:token})}catch(e){}}localStorage.removeItem("toxGardeSession");sessionStorage.removeItem("toxGardeSession");localStorage.removeItem("toxGardeCloudToken");sessionStorage.removeItem("toxGardeCloudToken");applyAuthUI()}
async function save(){
  localStorage.setItem("gardeMed",JSON.stringify(state));
  cloudSyncBusy=true;
  try{await cloudSave();lastCloudSnapshot=JSON.stringify(cleanStateForCloud());}
  catch(e){console.error("Synchronisation Supabase:",e);alert("Données enregistrées localement, mais la synchronisation en ligne a échoué. Vérifiez Internet.");}
  finally{cloudSyncBusy=false;}
  renderAll();
}
let cloudSyncBusy=false;
let cloudPollTimer=null;
let lastCloudSnapshot="";

async function refreshCloudOnStartup(){
  const ses=currentSession(), token=cloudToken();
  if(!ses||!token)return;
  try{
    const remote=await cloudLoad(token);
    if(remote&&typeof remote==="object"&&Object.keys(remote).length){
      const incoming=JSON.stringify(remote);
      lastCloudSnapshot=incoming;
      state=Object.assign(state,remote);
      state.accounts=[];
      localStorage.setItem("gardeMed",JSON.stringify(state));
      renderAll();
      if(ses.role!=="admin")showDoctorValidatedPlanning();
    }
  }catch(e){
    console.error("Tox-Garde / actualisation Supabase :",e);
  }
}

async function pollCloudChanges(){
  const ses=currentSession(), token=cloudToken();
  if(!ses||!token||cloudSyncBusy)return;
  try{
    const remote=await cloudLoad(token);
    if(!remote||typeof remote!=="object"||!Object.keys(remote).length)return;
    const incoming=JSON.stringify(remote);
    if(incoming===lastCloudSnapshot)return;
    lastCloudSnapshot=incoming;
    state=Object.assign(state,remote);
    state.accounts=[];
    localStorage.setItem("gardeMed",JSON.stringify(state));
    renderAll();
    if(ses.role!=="admin")showDoctorValidatedPlanning();
    console.log("Tox-Garde : données Supabase actualisées automatiquement.");
  }catch(e){
    console.error("Tox-Garde / synchronisation automatique :",e);
  }
}

function startCloudPolling(){
  if(cloudPollTimer)clearInterval(cloudPollTimer);
  cloudPollTimer=setInterval(pollCloudChanges,5000);
}
function renderMyGuards(){
  const el=document.getElementById("myGuardsContent");
  if(!el)return;
  const ses=currentSession();
  if(!ses || ses.role==="admin"){el.innerHTML='<div class="hint">Cette rubrique est réservée aux comptes Médecin.</div>';return;}
  const doctor=ses.doctor, today=iso(new Date()), byKey=new Map();
  Object.entries(state.planningHistory||{}).forEach(([key,raw])=>{
    const h=historyData(key);
    if(!h || !isHistoryConfirmed(key,raw))return;
    Object.entries(h.planning||{}).forEach(([date,items])=>{
      if(date<today)return;
      const item=(items||[]).find(x=>x.doctor===doctor && ["J","N"].includes(x.type));
      if(item)byKey.set(date,{date,type:item.type});
    });
    (h.fixed||[]).forEach(x=>{
      if(x.doctor===doctor && x.date>=today && ["G","F"].includes(x.type))byKey.set(x.date,{date:x.date,type:x.type});
    });
  });
  const guards=[...byKey.values()].sort((a,b)=>a.date.localeCompare(b.date)).slice(0,10);
  const fmt=d=>{const dt=dateFromISO(d);return {date:dt.toLocaleDateString("fr-FR",{day:"2-digit",month:"2-digit",year:"numeric"}),day:dt.toLocaleDateString("fr-FR",{weekday:"long"})};};
  const typeLabel=t=>({J:"☀️ J",N:"🌙 N",G:"🟢 G",F:"🟣 F"}[t]||t);
  if(!guards.length){el.innerHTML='<div class="hint">Aucune prochaine garde dans les plannings validés.</div>';return;}
  const first=guards[0], f=fmt(first.date), days=Math.max(0,Math.round((dateFromISO(first.date)-dateFromISO(today))/86400000));
  const when=days===0?"Aujourd’hui":days===1?"Demain":`Dans ${days} jours`;
  el.innerHTML=`<div style="border:1px solid var(--line);border-radius:10px;padding:14px;margin-bottom:14px;background:#f8fafc"><div class="hint">⏰ Prochaine garde</div><div style="font-size:20px;font-weight:800;margin-top:4px">${f.day} ${f.date}</div><div style="margin-top:5px;font-weight:700">${typeLabel(first.type)} — ${when}</div></div><table><thead><tr><th>Date</th><th>Jour</th><th>Type</th></tr></thead><tbody>${guards.map(g=>{const x=fmt(g.date);return `<tr><td>${x.date}</td><td>${x.day}</td><td><b>${typeLabel(g.type)}</b></td></tr>`}).join("")}</tbody></table>`;
}
function applyAuthUI(){const ses=currentSession(), overlay=document.getElementById("authOverlay");document.body.classList.toggle("auth-locked",!ses);overlay.style.display=ses?"none":"flex";if(!ses){const remembered=localStorage.getItem("toxGardeRememberedUsername")||"";document.getElementById("loginUsername").value=remembered;document.getElementById("keepConnected").checked=!!localStorage.getItem("toxGardeSession");document.getElementById("authUserInfo").innerHTML="";return;}
const admin=ses.role==="admin";document.querySelectorAll(".admin-only").forEach(el=>el.classList.toggle("doctor-hidden",!admin));const myGuardsTab=document.getElementById("myGuardsTab");if(myGuardsTab)myGuardsTab.style.display=admin?"none":"block";document.getElementById("authUserInfo").innerHTML=`${ses.username} — ${admin?"Administrateur":"Médecin"}<button class="btn small" onclick="logout()">Déconnexion</button>`;
if(!admin){
  document.getElementById("doctorAbsenceHint").style.display="block";
  const absDoctor=document.getElementById("absDoctor");
  if(absDoctor){
    absDoctor.innerHTML=`<option value="${String(ses.doctor||"").replaceAll('"','&quot;')}">${ses.doctor||"Médecin"}</option>`;
    absDoctor.value=ses.doctor||""; absDoctor.disabled=true;
  }
  renderDoctorHistorySelect(); openTab("planning"); showDoctorValidatedPlanning();
}else{
  document.getElementById("doctorAbsenceHint").style.display="none";
  const absDoctor=document.getElementById("absDoctor"); if(absDoctor)absDoctor.disabled=false;
}
renderAccountOptions();renderAccounts();renderAll();}
function renderAccountOptions(){const el=document.getElementById("accountDoctor");if(!el)return;el.innerHTML='<option value="">-- Choisir un médecin --</option>'+state.doctors.map(d=>`<option value="${String(d.name).replaceAll('"','&quot;')}">${d.name}</option>`).join("")}
async function loadCloudAccounts(){if(!isAdmin())return [];const token=cloudToken();if(!token)return [];try{return await supaRpc("toxgarde_list_accounts",{p_token:token})}catch(e){console.error(e);return []}}
async function renderAccounts(){const el=document.getElementById("accountsBody");if(!el)return;if(!isAdmin()){el.innerHTML='<tr><td colspan="4" class="hint">Accès administrateur.</td></tr>';return;}const accounts=await loadCloudAccounts();el.innerHTML=accounts.map((a,i)=>`<tr><td>${a.doctor||"—"}</td><td>${a.username}</td><td>${a.role==="admin"?"Administrateur":"Médecin"}</td><td>${a.username!==INITIAL_ADMIN.username?`<button class="btn small danger" onclick="deleteDoctorAccount(${i},'${String(a.username).replaceAll("'","\\'")}')">Supprimer</button>`:""}</td></tr>`).join("")||'<tr><td colspan="4" class="hint">Aucun compte médecin créé.</td></tr>'}
async function saveDoctorAccount(){if(!isAdmin())return alert("Accès réservé aux administrateurs.");const doctor=document.getElementById("accountDoctor").value,username=document.getElementById("accountUsername").value.trim(),password=document.getElementById("accountPassword").value,role=document.getElementById("accountRole").value;if(!doctor||!username||!password)return alert("Veuillez renseigner le médecin, l’identifiant et le mot de passe.");if(normalizeUser(username)===normalizeUser(INITIAL_ADMIN.username))return alert("Cet identifiant est réservé à l’administrateur initial.");try{await supaRpc("toxgarde_upsert_account",{p_token:cloudToken(),p_doctor:doctor,p_username:username,p_password:password,p_role:role});document.getElementById("accountUsername").value="";document.getElementById("accountPassword").value="";await renderAccounts();const m=document.getElementById("accountFormMessage");if(m){m.textContent="Compte médecin créé/mis à jour avec succès.";m.style.display="block";}}catch(e){alert("Impossible d'enregistrer le compte : "+e.message)}}
async function deleteDoctorAccount(i,username){if(!isAdmin())return alert("Accès réservé aux administrateurs.");if(!confirm("Supprimer ce compte ?"))return;try{await supaRpc("toxgarde_delete_account",{p_token:cloudToken(),p_username:username});await renderAccounts()}catch(e){alert("Impossible de supprimer le compte : "+e.message)}}
function clearDoctorAccountForm(){
  const d=document.getElementById("accountDoctor"),u=document.getElementById("accountUsername"),p=document.getElementById("accountPassword"),r=document.getElementById("accountRole"),m=document.getElementById("accountFormMessage");
  if(d)d.value="";
  if(u)u.value="";
  if(p){p.value="";p.type="password";}
  if(r)r.value="doctor";
  if(m){m.style.display="none";m.textContent="";}
}
function toggleAccountPassword(){
  const p=document.getElementById("accountPassword"); if(!p)return;
  p.type=p.type==="password"?"text":"password";
}

const _openTab=openTab;openTab=function(id){if(!isAdmin()&&!['dashboard','planning','absences','stats','myGuards'].includes(id))id="planning";if(isAdmin()&&id==="myGuards")id="dashboard";_openTab(id);if(id==="myGuards")renderMyGuards();};
function addDoctorRecovery(){
  const ses=currentSession(); if(!ses||ses.role==="admin")return alert("Cette saisie est destinée aux comptes Médecin.");
  const date=document.getElementById("doctorRecoveryDate")?.value; if(!date)return alert("Sélectionnez une date de récupération.");
  if(state.absences.some(a=>a.doctor===ses.doctor&&a.start===date&&a.end===date&&a.type==="RÉCUP"))return alert("Cette récupération existe déjà.");
  state.absences.push({doctor:ses.doctor,start:date,end:date,type:"RÉCUP",createdAt:Date.now()}); rebuildRecoveryLedger();
  save();
  const msg=document.getElementById("doctorRecoveryMessage"); if(msg){msg.textContent="Récupération enregistrée.";msg.style.display="block";}
  renderAll();
}
const _addAbsence=addAbsence;addAbsence=function(){const ses=currentSession();if(ses?.role!=="admin"){const sel=document.getElementById("absDoctor");sel.value=ses.doctor||"";const type=document.getElementById("absType").value.toUpperCase();if(!["CONGÉ","INDISPONIBLE","INDISP","RÉCUP","RECUP"].includes(type))return alert("Un médecin peut saisir uniquement un congé, une indisponibilité ou une récupération.");}return _addAbsence()};
const _renderAll=renderAll;renderAll=function(){_renderAll();renderHistorySelect();renderMyGuards();const ses=currentSession();if(ses&&ses.role!=="admin"){const sel=document.getElementById("absDoctor");if(sel){sel.innerHTML=`<option>${ses.doctor||""}</option>`;sel.value=ses.doctor||"";sel.disabled=true;}const t=document.getElementById("absType");if(t){[...t.options].forEach(o=>o.disabled=!( ["CONGÉ","INDISPONIBLE","INDISP","RÉCUP","RECUP"].includes(o.value.toUpperCase())));if(!["CONGÉ","INDISPONIBLE","INDISP","RÉCUP","RECUP"].includes(t.value.toUpperCase()))t.value="CONGÉ";}renderAccounts();}else renderAccountOptions();};
const remembered=localStorage.getItem("toxGardeRememberedUsername");if(remembered){const lu=document.getElementById("loginUsername");if(lu)lu.value=remembered;}
rebuildRecoveryLedger();
applyAuthUI();
renderAll();
applyAppearance();
refreshCloudOnStartup().finally(startCloudPolling);
