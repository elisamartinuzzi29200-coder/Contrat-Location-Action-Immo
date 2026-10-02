const sections=["Type de contrat","Bailleur(s)","Locataire(s)","Logement","Durée","Finances","Honoraires de location","Travaux / garanties","Annexes","Avenant au contrat","Récapitulatif"];
let step=0,viewMode="dashboard",currentId=null,cautionView=false;
const baseModels={nu_gestion:"Location nue — gestion",nu_hors:"Location nue — hors gestion",meuble_gestion:"Location meublée — gestion",meuble_hors:"Location meublée — hors gestion"};
const agencies={brest:"Brest",quimper:"Quimper"};
const models=Object.fromEntries(Object.entries(agencies).flatMap(([agency,agencyLabel])=>Object.entries(baseModels).map(([key,label])=>[agency+"_"+key,{label:agencyLabel+" — "+label}])));
const modelKey=d=>(d.agence||"brest")+"_"+d.type+"_"+d.gestion;
const extraModels={
  caution_brest_location:{label:"Brest — Acte de caution — location"},
  caution_brest_colocation:{label:"Brest — Acte de caution — colocation"},
  caution_quimper_location:{label:"Quimper — Acte de caution — location"},
  caution_quimper_colocation:{label:"Quimper — Acte de caution — colocation"}
};
const cautionModelKey=d=>"caution_"+(d.agence||"brest")+"_"+((d.cautionActe||{}).mode||"location");
const blankBailleur=()=>({type:"physique",nom:"",prenoms:"",denomination:"",adresse:"",email:"",tel:""});
const blankLoc=()=>({nom:"",prenoms:"",naissance:"",lieuNaissance:"",email:"",tel:""});
const initial=()=>({agence:"brest",type:"nu",gestion:"gestion",bailleurs:[blankBailleur()],locataires:[blankLoc()],localisation:"",habitat:"collectif",identifiantFiscal:"",regime:"copropriete",periode:"depuis2005",surface:"",pieces:"",caracteristiques:"",autresParties:"",autresPartiesChoix:{},autresPartiesAutre:"",equipements:"",equipementsChoix:{},equipementsAutre:"",chauffageMode:"individuel",chauffageEnergie:"electricite",chauffageAutre:"",eauMode:"individuel",eauEnergie:"electricite",eauAutre:"",destination:"habitation",professionMixte:"",accessoiresPrivatifs:"",accessoiresPrivatifsChoix:{},accessoiresPrivatifsNumeros:{},accessoiresPrivatifsAutre:"",partiesCommunes:"",partiesCommunesChoix:{},partiesCommunesNumeros:{},partiesCommunesAutre:"",technologies:"",depensesEnergie:"",anneeEnergie:"",dateEffet:"",duree:"",raisonDureeReduite:"",loyer:"",decretRelocation:"non",encadrement:"non",loyerReference:"",loyerReferenceMajore:"",loyerBase:"",complementLoyer:"",dernierLoyer:"",dateVersementDernier:"",dateDerniereRevision:"",dateRevision:"",irl:"",irlTrimestre:"",irlAnnee:"",irlValeur:"",chargesMode:"provision",chargesMontant:"",contribution:"",justifContribution:"",assuranceColocAnnuelle:"",assuranceColocMensuelle:"",depotGarantie:"",honorairesVisiteBailleur:"",honorairesVisiteLocataire:"",honorairesEdlBailleur:"",honorairesEdlLocataire:"",travauxRecents:"",majorationTravaux:"",diminutionTravaux:"",sinistre:"non",congeLocataire:"",conditionsLocataire:"",conditionsBailleur:"",caution:"",cautionActe:{mode:"location",civilite:"Monsieur",nom:"",prenoms:"",naissance:"",lieuNaissance:"",domicile:"",email:"",tel:"",situation:"Célibataire",profession:"",employeur:"",contratType:"indeterminee",contratFin:"",remuneration:"",autresRevenus:"",dateBail:"",irlTrimestre:"",irlAnnee:"",irlValeur:"",reconductions:"",dureeMax:"",montantMax:"",anneesLoyers:""},annexes:{},avenantCharges:{},avenantInfos:{}});
let data=initial(),autosaveTimer=null,lastGeneratedBlob=null,lastGeneratedFilename="";
function loadDossiers(){try{return JSON.parse(localStorage.getItem("ai-location-dossiers")||"[]")}catch{return []}}
function storeDossiers(list){localStorage.setItem("ai-location-dossiers",JSON.stringify(list))}
function dossierTitle(d){
  const b=(d.bailleurs||[]).map(p=>p.type==="morale"?(p.denomination||"").trim():((p.nom||"")+" "+(p.prenoms||"")).trim()).filter(Boolean).join(" / ");
  const l=(d.locataires||[]).map(p=>((p.nom||"")+" "+(p.prenoms||"")).trim()).filter(Boolean).join(" / ");
  return [b&&"Bailleur : "+b,l&&"Locataire : "+l,d.localisation].filter(Boolean).join(" — ")||"Dossier sans nom";
}
function newDossier(){
  $("newDossierModal").classList.remove("hidden");
}
function startBlankDossier(){
  currentId="loc_"+Date.now()+"_"+Math.random().toString(36).slice(2,8);
  data=initial();step=0;cautionView=false;viewMode="editor";$("newDossierModal").classList.add("hidden");render();
}
function sanitizeImportedData(imported){
  const clean={...initial(),...imported};
  clean.locataires=[blankLoc()];
  clean.dateEffet="";
  clean.irl="";
  clean.dateRevision="";
  clean.loyer="";
  clean.loyerReference="";
  clean.loyerReferenceMajore="";
  clean.loyerBase="";
  clean.complementLoyer="";
  clean.dernierLoyer="";
  clean.dateVersementDernier="";
  clean.dateDerniereRevision="";
  clean.depotGarantie="";
  clean.honorairesVisiteBailleur="";
  clean.honorairesVisiteLocataire="";
  clean.honorairesEdlBailleur="";
  clean.honorairesEdlLocataire="";
  return clean;
}
function importSummaryHtml(d){
  const b=(d.bailleurs||[]).map(p=>p.type==="morale"?(p.denomination||""):[p.nom,p.prenoms].filter(Boolean).join(" ")).filter(Boolean).join(", ")||"Non détecté";
  const annexCount=Object.values(d.annexes||{}).filter(Boolean).length;
  const avenantCount=Object.values(d.avenantCharges||{}).filter(Boolean).length+Object.values(d.avenantInfos||{}).filter(Boolean).length;
  const reusable=[
    ["Type de contrat",models[modelKey(d)]?.label||""],
    ["Bailleur(s)",b],
    ["Adresse du logement",d.localisation||"Non détectée"],
    ["Habitat",d.habitat==="individuel"?"Individuel":"Collectif"],
    ["Régime",d.regime==="monopropriete"?"Monopropriété":"Copropriété"],
    ["Période de construction",({"avant1949":"Avant 1949","1949-1974":"1949 à 1974","1975-1989":"1975 à 1989","1989-2005":"1989 à 2005","depuis2005":"Depuis 2005"})[d.periode]||""],
    ["Surface",d.surface?d.surface+" m²":"Non détectée"],
    ["Pièces principales",d.pieces||"Non détecté"],
    ["Caractéristiques",d.caracteristiques||"Non détectées"],
    ["Équipements",d.equipements||"Non détectés"],
    ["Chauffage",[d.chauffageMode==="collectif"?"Collectif":"Individuel",d.chauffageAutre].filter(Boolean).join(" — ")],
    ["Eau chaude",[d.eauMode==="collectif"?"Collective":"Individuelle",d.eauAutre].filter(Boolean).join(" — ")],
    ["Technologies",d.technologies||"Non détectées"],
    ["Dépenses énergétiques",d.depensesEnergie?d.depensesEnergie+" €"+(d.anneeEnergie?" — réf. "+d.anneeEnergie:""):"Non détectées"],
    ["Durée",d.duree||"Non détectée"],
    ["Charges",d.chargesMontant?d.chargesMontant+" €":"Non détectées"],
    ["Honoraires visite/rédaction",d.honorairesVisiteBailleur?d.honorairesVisiteBailleur+" € bailleur / "+(d.honorairesVisiteLocataire||"")+" € locataire":"Non détectés"],
    ["Honoraires état des lieux",d.honorairesEdlBailleur?d.honorairesEdlBailleur+" € bailleur / "+(d.honorairesEdlLocataire||"")+" € locataire":"Non détectés"],
    ["Annexes cochées",String(annexCount)],
    ["Éléments d’avenant cochés",String(avenantCount)]
  ];
  return `<div class="notice">Les informations ci-dessous ont été récupérées du bail. Les champs incertains restent vides plutôt que de reprendre du texte standard du contrat.</div>
  <div class="importGrid">${reusable.map(([k,v])=>`<div class="importItem"><strong>${esc(k)}</strong><span>${esc(v)}</span></div>`).join("")}</div>
  <div class="section">Informations volontairement non reprises</div>
  <div class="notice">Locataire(s), date de prise d’effet du bail, IRL, loyer et éléments liés au loyer, dépôt de garantie.</div>`;
}
function saveDossier(statusText="Dossier sauvegardé"){
  if(!currentId)currentId="loc_"+Date.now()+"_"+Math.random().toString(36).slice(2,8);
  const list=loadDossiers(),idx=list.findIndex(x=>x.id===currentId);
  const item={id:currentId,title:dossierTitle(data),updatedAt:new Date().toISOString(),data:JSON.parse(JSON.stringify(data))};
  if(idx>=0)list[idx]=item;else list.unshift(item);
  storeDossiers(list);
  if(statusText!==false)$("status").textContent=statusText;
}
function scheduleAutosave(){
  if(viewMode==="dashboard"||!currentId)return;
  clearTimeout(autosaveTimer);
  $("status").textContent="Enregistrement…";
  autosaveTimer=setTimeout(()=>saveDossier("✓ Enregistré automatiquement"),500);
}
function openDossier(id){
  const item=loadDossiers().find(x=>x.id===id);if(!item)return;
  currentId=id;data={...initial(),...item.data};step=0;cautionView=false;viewMode="editor";render();
}
function duplicateDossier(id){
  const item=loadDossiers().find(x=>x.id===id);if(!item)return;
  currentId="loc_"+Date.now()+"_"+Math.random().toString(36).slice(2,8);
  data={...initial(),...JSON.parse(JSON.stringify(item.data))};step=0;viewMode="editor";saveDossier();render();
}
function deleteDossier(id){
  if(!confirm("Supprimer définitivement ce dossier de ce navigateur ?"))return;
  storeDossiers(loadDossiers().filter(x=>x.id!==id));renderDashboard();
}
try{
  const legacy=localStorage.getItem("ai-location-form");
  if(legacy&&loadDossiers().length===0){
    const old={...initial(),...JSON.parse(legacy)},id="loc_migration_"+Date.now();
    storeDossiers([{id,title:dossierTitle(old),updatedAt:new Date().toISOString(),data:old}]);
    localStorage.removeItem("ai-location-form");
  }
}catch{}
const $=id=>document.getElementById(id),esc=s=>String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m]));
const F=(l,k,t="text",full=false)=>`<div class="field ${full?"full":""}"><label>${l}</label><input type="${t}" data-key="${k}" value="${esc(data[k])}"></div>`;
const cautionDefaults=()=>({civilite:"Monsieur",nom:"",prenoms:"",naissance:"",lieuNaissance:"",domicile:"",email:"",tel:"",situation:"Célibataire",profession:"",employeur:"",contratType:"indeterminee",contratFin:"",remuneration:"",autresRevenus:"",dateBail:"",irlTrimestre:"",irlAnnee:"",irlValeur:"",reconductions:"",dureeMax:"",montantMax:"",anneesLoyers:""});
function ensureCaution(){data.cautionActe={...cautionDefaults(),...(data.cautionActe||{})};return data.cautionActe}
function cautionField(label,key,type="text",full=false){const s=ensureCaution();return `<div class="field ${full?"full":""}"><label>${label}</label><input type="${type}" data-caution="${key}" value="${esc(s[key])}"></div>`}
function cautionSelect(label,key,opts){const s=ensureCaution();return `<div class="field"><label>${label}</label><select data-caution="${key}">${opts.map(([v,n])=>`<option value="${v}" ${s[key]===v?"selected":""}>${n}</option>`).join("")}</select></div>`}

const IRL_VALUES={
"2022-1":"133.93","2022-2":"135.84","2022-3":"136.27","2022-4":"137.26",
"2023-1":"138.61","2023-2":"140.59","2023-3":"141.03","2023-4":"142.06",
"2024-1":"143.46","2024-2":"145.17","2024-3":"144.51","2024-4":"144.64",
"2025-1":"145.47","2025-2":"146.68","2025-3":"145.77","2025-4":"145.78",
"2026-1":"146.60","2026-2":"148.37"};
function syncIrlValue(){
 const s=ensureCaution();
 if(!data.irlTrimestre&&s.irlTrimestre)data.irlTrimestre=s.irlTrimestre;
 if(!data.irlAnnee&&s.irlAnnee)data.irlAnnee=s.irlAnnee;
 const key=String(data.irlAnnee||"")+"-"+String(data.irlTrimestre||"");
 const v=IRL_VALUES[key]||"";
 data.irlValeur=v;
 s.irlTrimestre=String(data.irlTrimestre||"");s.irlAnnee=String(data.irlAnnee||"");s.irlValeur=v;
 if(v){
   const q=Number(data.irlTrimestre)||0;
   const qLabel=q===1?"1er trimestre":q+"ème trimestre";
   data.irl=qLabel+" "+data.irlAnnee+" d'une valeur de "+v;
 }
 return v;
}

function parseIrlForCaution(){
  const s=ensureCaution(),raw=String(data.irl||"");
  if(data.irlTrimestre)s.irlTrimestre=String(data.irlTrimestre);
  if(data.irlAnnee)s.irlAnnee=String(data.irlAnnee);
  if(data.irlValeur)s.irlValeur=String(data.irlValeur);
  if(!s.irlTrimestre){const m=raw.match(/\b([1-4])(?:er|e|ème)?\s*(?:trimestre|trim\.?)/i);if(m)s.irlTrimestre=m[1]}
  if(!s.irlAnnee){const m=raw.match(/\b(20\d{2})\b/);if(m)s.irlAnnee=m[1]}
  if(!s.irlValeur){const nums=[...raw.matchAll(/\b(\d{2,3}[.,]\d{1,3})\b/g)].map(x=>x[1]);if(nums.length)s.irlValeur=nums[nums.length-1]}syncIrlValue();
  return s;
}

const T=(l,k)=>`<div class="field full"><label>${l}</label><textarea data-key="${k}">${esc(data[k])}</textarea></div>`;
const Sel=(l,k,opts)=>`<div class="field"><label>${l}</label><select data-key="${k}">${opts.map(([v,n])=>`<option value="${v}" ${data[k]===v?"selected":""}>${n}</option>`).join("")}</select></div>`;
function bailleursHtml(){return data.bailleurs.map((p,i)=>`<div class="box"><div class="boxhead"><strong>Bailleur ${i+1}</strong>${data.bailleurs.length>1?`<button data-rmb="${i}">Retirer</button>`:""}</div>
<label class="check"><input type="radio" name="bt-${i}" data-btype="${i}" value="physique" ${p.type==="physique"?"checked":""}>Personne physique</label>
<label class="check"><input type="radio" name="bt-${i}" data-btype="${i}" value="morale" ${p.type==="morale"?"checked":""}>Personne morale</label>
<div class="grid">${p.type==="morale"?`<div class="field"><label>Dénomination</label><input data-b="${i}:denomination" value="${esc(p.denomination)}"></div>`:`<div class="field"><label>Nom</label><input data-b="${i}:nom" value="${esc(p.nom)}"></div><div class="field"><label>Prénom(s)</label><input data-b="${i}:prenoms" value="${esc(p.prenoms)}"></div>`}
<div class="field full"><label>${p.type==="morale"?"Siège social":"Adresse"}</label><input data-b="${i}:adresse" value="${esc(p.adresse)}"></div><div class="field"><label>Email</label><input data-b="${i}:email" value="${esc(p.email)}"></div><div class="field"><label>Téléphone</label><input data-b="${i}:tel" value="${esc(p.tel)}"></div></div></div>`).join("")+`<button data-addb>Ajouter un bailleur</button>`}
function locatairesHtml(){return data.locataires.map((p,i)=>`<div class="box"><div class="boxhead"><strong>Locataire ${i+1}</strong>${data.locataires.length>1?`<button data-rml="${i}">Retirer</button>`:""}</div><div class="grid"><div class="field"><label>Nom</label><input data-l="${i}:nom" value="${esc(p.nom)}"></div><div class="field"><label>Prénom(s)</label><input data-l="${i}:prenoms" value="${esc(p.prenoms)}"></div><div class="field"><label>Date de naissance</label><input type="date" data-l="${i}:naissance" value="${esc(p.naissance)}"></div><div class="field"><label>Lieu de naissance</label><input data-l="${i}:lieuNaissance" value="${esc(p.lieuNaissance)}"></div><div class="field"><label>Email</label><input data-l="${i}:email" value="${esc(p.email)}"></div><div class="field"><label>Téléphone</label><input data-l="${i}:tel" value="${esc(p.tel)}"></div></div></div>`).join("")+`<button data-addl>Ajouter un locataire</button>`}
const annexes=["Un extrait du règlement concernant la destination de l’immeuble","Le règlement intérieur de l’immeuble","Un document informatif sur les risques de nuisances sonores aériennes","Un diagnostic de performance énergétique","Un constat de risque d‘exposition au plomb","Une copie d’un état mentionnant l’absence ou la présence de matériaux","Un état de l’installation intérieure d’électricité et de gaz","Un état des risques naturels et technologiques","Une notice d’information relative aux droits et obligations","Un état des lieux","Une autorisation préalable de mise en location","Les références aux loyers habituellement constatés","Une grille de vétusté"];

const autresPartiesOptions=["Grenier","Comble aménagé","Comble non aménagé","Terrasse","Balcon","Loggia","Jardin"];
const equipementOptions=["Cuisine équipée","Salle de bain","Salle de douche","WC séparé"];
const privatifOptions=["Cave","Parking","Garage"];
const communOptions=["Garage à vélo","Ascenseur","Espaces verts","Aires et équipements de jeux","Laverie","Local poubelle","Gardiennage"];
function checkGroupHtml(title,key,options,otherKey){
  data[key]=data[key]||{};
  return `<div class="field full"><label>${title}</label><div class="grid">${options.map(o=>`<label class="check"><input type="checkbox" data-group="${key}" data-option="${esc(o)}" ${data[key][o]?"checked":""}>${esc(o)}</label>`).join("")}</div>${otherKey?`<input data-key="${otherKey}" placeholder="Autre (préciser)" value="${esc(data[otherKey]||"")}">`:""}</div>`;
}
function numberedGroupHtml(title,key,numKey,options,otherKey){
 data[key]=data[key]||{};data[numKey]=data[numKey]||{};
 return `<div class="field full"><label>${title}</label><div class="grid">${options.map(o=>`<div class="box"><label class="check"><input type="checkbox" data-group="${key}" data-option="${esc(o)}" ${data[key][o]?"checked":""}>${esc(o)}</label><input data-number-group="${numKey}" data-option="${esc(o)}" placeholder="N° (si applicable)" value="${esc(data[numKey][o]||"")}"></div>`).join("")}</div>${otherKey?`<input data-key="${otherKey}" placeholder="Autre (préciser)" value="${esc(data[otherKey]||"")}">`:""}</div>`;
}

const avenantCharges=[
"L’eau (elle sera réajustée en plus ou en moins selon les consommations réelles)",
"L’électricité","Le gaz","Internet","La minuterie","Le ménage des parties communes","L’ascenseur","Le Contrat d’entretien de la chaudière"
];
const avenantInfos=[
"La taxe d’ordure ménagère sera à payer par le locataire.",
"Le compteur électrique et/ou gaz devra être ouvert au nom du locataire.",
"Le compteur d’eau devra être ouvert au nom du locataire.",
"La taxe d'habitation sera due par le locataire au 1er janvier.",
"Le locataire devra contracter une police d'assurance incendie et dégâts des eaux avant la remise des clés et s'engage à fournir au bailleur un justificatif annuel.",
"Le locataire s’engage à prendre un contrat d’entretien pour la chaudière et à fournir au bailleur un justificatif annuel.",
"Le locataire s’engage à faire un ramonage annuel de la cheminée ou du poêle à bois et à fournir au bailleur un justificatif annuel.",
"Le locataire s’engage à entretenir le jardin et les abords de la maison (pelouses, haies, plantations, terrasse)",
"Le locataire s’engage à fournir une pile pour le détecteur de fumée, lors de l’état des lieux de sortie si celle-ci ne fonctionne plus.",
"Le preneur ou locataire s’engage à prendre un contrat d’entretien pour la VMC chaque année.",
"Le preneur ou locataire s’engage à prendre un contrat d’entretien pour la pompe à chaleur chaque année"
];
function renderDashboard(){
  viewMode="dashboard";currentId=null;$("nav").innerHTML="";
  const list=loadDossiers().sort((a,b)=>String(b.updatedAt).localeCompare(String(a.updatedAt)));
  $("content").innerHTML=`<div class="dashboardHead"><div><h2>MES DOSSIERS LOCATION</h2><p class="hint">Les dossiers sont enregistrés uniquement dans ce navigateur.</p></div><button class="primary" id="dashNew">+ Nouveau dossier</button></div>
  ${list.length?`<div class="dossierList">${list.map(x=>`<article class="dossierCard"><div><strong>${esc(x.title)}</strong><small>Dernière modification : ${new Date(x.updatedAt).toLocaleString("fr-FR")}</small></div><div class="dossierActions"><button data-open="${x.id}">Ouvrir</button><button data-duplicate="${x.id}">Dupliquer</button><button class="dangerBtn" data-delete="${x.id}">Supprimer</button></div></article>`).join("")}</div>`:`<div class="emptyState"><strong>Aucun dossier enregistré</strong><p>Crée un dossier de location, sauvegarde-le, puis retrouve-le ici.</p></div>`}`;
  $("prev").style.display="none";$("next").style.display="none";
  $("dashNew").onclick=newDossier;
  document.querySelectorAll("[data-open]").forEach(x=>x.onclick=()=>openDossier(x.dataset.open));
  document.querySelectorAll("[data-duplicate]").forEach(x=>x.onclick=()=>duplicateDossier(x.dataset.duplicate));
  document.querySelectorAll("[data-delete]").forEach(x=>x.onclick=()=>deleteDossier(x.dataset.delete));
}
function dossierChecks(){
  const issues=[],add=(level,step,label,detail)=>issues.push({level,step,label,detail});
  const bailleurOk=p=>p.type==="morale"?!!(p.denomination&&p.adresse):!!(p.nom&&p.prenoms&&p.adresse);
  const locOk=p=>!!(p.nom&&p.prenoms&&p.naissance&&p.lieuNaissance);
  if(!(data.bailleurs||[]).length||(data.bailleurs||[]).some(p=>!bailleurOk(p)))add("error",1,"Bailleur(s) incomplet(s)","Nom, prénom et adresse, ou dénomination et siège social.");
  if(!(data.locataires||[]).length||(data.locataires||[]).some(p=>!locOk(p)))add("error",2,"Locataire(s) incomplet(s)","Nom, prénom, date et lieu de naissance.");
  if(!data.localisation)add("error",3,"Adresse du logement manquante","Renseigne la localisation complète du logement.");
  if(!data.surface)add("error",3,"Surface habitable manquante","Renseigne la surface habitable.");
  if(!data.pieces)add("warn",3,"Nombre de pièces non renseigné","Vérifie le nombre de pièces principales.");
  if(!data.dateEffet)add("error",4,"Date de prise d'effet manquante","Renseigne la date de début du bail.");
  if(!data.duree)add("error",4,"Durée du contrat manquante","Sélectionne la durée du bail.");
  if(!data.loyer)add("error",5,"Loyer mensuel manquant","Renseigne le montant du loyer.");
  if(!data.irlTrimestre||!data.irlAnnee)add("warn",5,"Référence IRL incomplète","Renseigne le trimestre et l'année de référence si le bail prévoit une révision.");
  else if(!data.irlValeur)add("warn",5,"Valeur IRL indisponible","La valeur IRL n'est pas disponible dans l'application pour cette période.");
  if(!data.chargesMontant)add("warn",5,"Charges non renseignées","Vérifie le montant des charges, y compris s'il est nul.");
  if(!data.depotGarantie)add("warn",7,"Dépôt de garantie non renseigné","Vérifie le dépôt de garantie applicable.");
  if(!(data.annexes||{})["Un diagnostic de performance énergétique"])add("warn",8,"DPE non coché dans les annexes","Vérifie que le DPE est bien joint au bail.");
  if(!(data.annexes||{})["Une notice d’information relative aux droits et obligations"])add("warn",8,"Notice d'information non cochée","Vérifie que la notice d'information est bien jointe.");
  if(!(data.annexes||{})["Un état des lieux"])add("warn",8,"État des lieux non coché","Vérifie s'il doit être joint au dossier.");
  return issues;
}
function dossierControllerHtml(){
  const issues=dossierChecks(),errors=issues.filter(x=>x.level==="error").length,warns=issues.filter(x=>x.level==="warn").length;
  if(!issues.length)return '<div class="controlSummary ok"><strong>✓ Dossier prêt à générer</strong><span>Aucun point bloquant ou point de vigilance détecté.</span></div>';
  return '<div class="controlSummary"><strong>Contrôle du dossier — '+errors+' bloquant(s), '+warns+' vigilance(s)</strong><div class="controlList">'+issues.map(x=>'<button type="button" class="controlItem '+x.level+'" data-goto-step="'+x.step+'"><span><b>'+(x.level==="error"?"Bloquant":"À vérifier")+' — '+esc(x.label)+'</b></span><small>'+esc(x.detail)+'</small><em>Corriger →</em></button>').join("")+'</div></div>';
}
function render(){if(viewMode==="dashboard"){renderDashboard();return}if(cautionView){renderCautionView();return}$("prev").style.display="";$("next").style.display="";$("nav").innerHTML=sections.map((s,i)=>`<button data-step="${i}" class="${i===step?"active":""}">${i+1}. ${s}</button>`).join("");let h=`<h2>${sections[step]}</h2><p class="hint">Les données saisies servent à remplir automatiquement le bon modèle Word sans modifier ses clauses fixes.</p>`;
if(step===0)h+=`<div class="choiceGrid"><div class="box"><strong>Agence</strong><label class="check"><input type="radio" name="agence" value="brest" ${data.agence==="brest"?"checked":""}>Brest</label><label class="check"><input type="radio" name="agence" value="quimper" ${data.agence==="quimper"?"checked":""}>Quimper</label></div><div class="box"><strong>Type de location</strong><label class="check"><input type="radio" name="type" value="nu" ${data.type==="nu"?"checked":""}>Logement nu</label><label class="check"><input type="radio" name="type" value="meuble" ${data.type==="meuble"?"checked":""}>Logement meublé</label></div><div class="box"><strong>Gestion</strong><label class="check"><input type="radio" name="gestion" value="gestion" ${data.gestion==="gestion"?"checked":""}>Gestion Action Immobilière</label><label class="check"><input type="radio" name="gestion" value="hors" ${data.gestion==="hors"?"checked":""}>Hors gestion</label></div></div><div class="notice">Modèle sélectionné : <strong>${models[modelKey(data)].label}</strong></div>`;
if(step===1)h+=bailleursHtml();
if(step===2)h+=locatairesHtml();
if(step===3)h+=`<div class="grid">${T("Localisation du logement","localisation")}${Sel("Type d’habitat","habitat",[["collectif","Collectif"],["individuel","Individuel"]])}${F("Identifiant fiscal du logement","identifiantFiscal")}${Sel("Régime juridique","regime",[["copropriete","Copropriété"],["monopropriete","Monopropriété"]])}${Sel("Période de construction","periode",[["avant1949","Avant 1949"],["1949-1974","1949 à 1974"],["1975-1989","1975 à 1989"],["1989-2005","1989 à 2005"],["depuis2005","Depuis 2005"]])}${F("Surface habitable (m²)","surface","number")}${F("Nombre de pièces principales","pieces","number")}${T("Caractéristiques du logement","caracteristiques")}${checkGroupHtml("Autres parties du logement","autresPartiesChoix",autresPartiesOptions,"autresPartiesAutre")}${checkGroupHtml("Éléments d’équipement du logement","equipementsChoix",equipementOptions,"equipementsAutre")}${Sel("Répartition du chauffage","chauffageMode",[["individuel","Individuel"],["collectif","Collectif"]])}${Sel("Énergie du chauffage","chauffageEnergie",[["electricite","Électricité"],["gaz","Gaz"],["autre","Autre"]])}${data.chauffageEnergie==="autre"?F("Précision chauffage / production","chauffageAutre"):""}${Sel("Eau chaude sanitaire","eauMode",[["individuel","Individuel"],["collectif","Collectif"]])}${Sel("Énergie de l’eau chaude","eauEnergie",[["electricite","Électricité"],["gaz","Gaz"],["autre","Autre"]])}${data.eauEnergie==="autre"?F("Précision eau chaude / production","eauAutre"):""}${Sel("Destination","destination",[["habitation","Habitation principale"],["mixte","Habitation + professionnelle"]])}${data.destination==="mixte"?F("Profession exercée","professionMixte"):""}${numberedGroupHtml("Locaux / équipements accessoires à usage privatif","accessoiresPrivatifsChoix","accessoiresPrivatifsNumeros",privatifOptions,"accessoiresPrivatifsAutre")}${checkGroupHtml("Locaux / parties / équipements à usage commun","partiesCommunesChoix",communOptions,"partiesCommunesAutre")}${T("Accès technologies / communication","technologies")}${F("Dépenses énergétiques annuelles estimées","depensesEnergie")}${F("Année de référence des prix énergétiques","anneeEnergie","number")}</div>`;
if(step===4)h+=`<div class="grid">${F("Date de prise d’effet","dateEffet","date")}<div class="field"><label>Durée du contrat</label><select data-key="duree"><option value=""></option>${Array.from({length:9},(_,i)=>{const n=i+1,v=n+" an"+(n>1?"s":"");return `<option value="${v}" ${data.duree===v?"selected":""}>${v}</option>`}).join("")}</select></div>${T("Événement / raison justifiant une durée réduite","raisonDureeReduite")}</div>`;
if(step===5)h+=`<div class="grid">${F("Loyer mensuel (€)","loyer","number")}${Sel("Décret évolution des loyers à la relocation","decretRelocation",[["non","Non"],["oui","Oui"]])}${Sel("Loyer de référence majoré applicable","encadrement",[["non","Non"],["oui","Oui"]])}${F("Loyer de référence","loyerReference")}${F("Loyer de référence majoré","loyerReferenceMajore")}${F("Loyer de base","loyerBase")}${F("Complément de loyer","complementLoyer")}${F("Dernier loyer","dernierLoyer")}${F("Date de versement du dernier loyer","dateVersementDernier","date")}${F("Date de dernière révision","dateDerniereRevision","date")}${F("Date de révision","dateRevision")}<div class="field"><label>Trimestre de référence IRL</label><select data-key="irlTrimestre"><option value=""></option>${[1,2,3,4].map(n=>`<option value="${n}" ${String(data.irlTrimestre)===String(n)?"selected":""}>${n}${n===1?"er":"ème"} trimestre</option>`).join("")}</select></div><div class="field"><label>Année IRL</label><select data-key="irlAnnee"><option value=""></option>${[2026,2027,2028].map(y=>`<option value="${y}" ${String(data.irlAnnee)===String(y)?"selected":""}>${y}</option>`).join("")}</select></div><div class="field"><label>Valeur IRL</label><input value="${esc(data.irlValeur||"")}" readonly><small>${data.irlTrimestre&&data.irlAnnee&&!data.irlValeur?"Valeur non encore disponible dans l’application pour cette période.":"Calculée automatiquement selon le trimestre et l'année."}</small></div>${Sel("Charges récupérables","chargesMode",[["provision","Provision mensuelle"],["forfait","Forfait"],["justificatif","Remboursement sur justificatif"]])}${F("Montant des charges (€)","chargesMontant","number")}${data.type==="nu"?T("Contribution partage des économies de charges","contribution"):""}${data.type==="nu"?T("Justification de cette contribution","justifContribution"):""}${F("Assurance colocation annuelle (€)","assuranceColocAnnuelle","number")}${F("Assurance colocation mensuelle (€)","assuranceColocMensuelle","number")}</div>`;
if(step===6){
  const totalB=(Number(data.honorairesVisiteBailleur)||0)+(Number(data.honorairesEdlBailleur)||0);
  const totalL=(Number(data.honorairesVisiteLocataire)||0)+(Number(data.honorairesEdlLocataire)||0);
  h+=`<div class="notice">Montants repris automatiquement dans le tableau « Honoraires de location » du Word.</div><div class="grid">${F("Visite, constitution du dossier et rédaction — Bailleur (€)","honorairesVisiteBailleur","number")}${F("Visite, constitution du dossier et rédaction — Locataire (€)","honorairesVisiteLocataire","number")}${F("État des lieux d’entrée — Bailleur (€)","honorairesEdlBailleur","number")}${F("État des lieux d’entrée — Locataire (€)","honorairesEdlLocataire","number")}<div class="field"><label>Total bailleur</label><input value="${totalB.toFixed(2)} €" readonly></div><div class="field"><label>Total locataire</label><input value="${totalL.toFixed(2)} €" readonly></div></div>`
}
if(step===7)h+=`<div class="grid">${F("Dépôt de garantie (€)","depotGarantie","number")}${T("Travaux récents / mise en conformité","travauxRecents")}${T("Majoration de loyer liée à des travaux","majorationTravaux")}${T("Diminution de loyer liée à des travaux","diminutionTravaux")}${Sel("Sinistre indemnisé catastrophe naturelle / technologique","sinistre",[["non","Non"],["oui","Oui"]])}${F("Date de congé du locataire en place","congeLocataire","date")}${T("Conditions particulières / obligations locataire","conditionsLocataire")}${T("Conditions particulières / obligations bailleur","conditionsBailleur")}${T("Caution solidaire","caution")}</div>`;
if(step===8)h+=`<div class="notice">Coche uniquement les annexes effectivement jointes au bail.</div><div class="grid">${annexes.map(a=>`<label class="check"><input type="checkbox" data-annexe="${esc(a)}" ${data.annexes[a]?"checked":""}>${a}</label>`).join("")}</div>`;
if(step===9)h+=`<div class="notice">Ces cases correspondent à la page « AVENANT AU CONTRAT DE LOCATION » du modèle Word.</div><div class="section">La provision sur charges comprend</div><div class="grid">${avenantCharges.map(a=>`<label class="check"><input type="checkbox" data-avcharge="${esc(a)}" ${data.avenantCharges[a]?"checked":""}>${a}</label>`).join("")}</div><div class="section">Informations diverses</div><div class="grid">${avenantInfos.map(a=>`<label class="check"><input type="checkbox" data-avinfo="${esc(a)}" ${data.avenantInfos[a]?"checked":""}>${a}</label>`).join("")}</div>`;
function cautionMaxAmount(){return (Number(data.loyer)||0)+(Number(data.chargesMontant)||0)?((Number(data.loyer)||0)+(Number(data.chargesMontant)||0))*(data.type==="meuble"?12:36):0}
function cautionMaxDisplay(){const n=cautionMaxAmount();if(!n)return "";return n.toLocaleString("fr-FR")+" € ("+numberToFrench(Math.round(n))+" euros)"}

function renderCautionView(){
  const s=parseIrlForCaution();
  $("nav").innerHTML="";$("prev").style.display="none";$("next").style.display="none";let h=`<h2>Acte de caution</h2><div class="box"><strong>Type d'acte</strong><label class="check"><input type="radio" name="cautionMode" value="location" ${((data.cautionActe||{}).mode||"location")==="location"?"checked":""}>Location</label><label class="check"><input type="radio" name="cautionMode" value="colocation" ${(data.cautionActe||{}).mode==="colocation"?"checked":""}>Colocation</label><small>Le modèle ${data.agence==="quimper"?"Quimper":"Brest"} correspondant sera utilisé automatiquement.</small></div>`;

  const bailleurs=data.bailleurs.map(p=>p.type==="morale"?(p.denomination||""):[p.nom,p.prenoms].filter(Boolean).join(" ")).filter(Boolean).join(" / ");
  const locataires=data.locataires.map(p=>[p.nom,p.prenoms].filter(Boolean).join(" ")).filter(Boolean).join(" / ");
  h+=`<div class="notice">Les informations du bail sont reprises automatiquement. Renseigne uniquement les informations propres à la caution et les paramètres de durée qui ne figurent pas dans le bail.</div>
  <div class="section">Caution signataire</div><div class="grid">
  ${cautionSelect("Civilité","civilite",[["Monsieur","Monsieur"],["Madame","Madame"]])}
  ${cautionField("Nom","nom")}${cautionField("Prénom(s)","prenoms")}
  ${cautionField("Date de naissance","naissance","date")}${cautionField("Lieu de naissance","lieuNaissance")}
  ${cautionField("Domicile","domicile","text",true)}${cautionField("Email","email","email")}${cautionField("Téléphone","tel")}
  ${cautionSelect("Situation familiale","situation",[["Célibataire","Célibataire"],["Marié(e)","Marié(e)"],["Pacsé(e)","Pacsé(e)"],["Divorcé(e)","Divorcé(e)"],["Veuf(ve)","Veuf(ve)"]])}
  </div><div class="section">Situation professionnelle</div><div class="grid">
  ${cautionField("Profession","profession")}${cautionField("Employeur","employeur")}
  ${cautionSelect("Contrat de travail","contratType",[["indeterminee","Durée indéterminée"],["determinee","Durée déterminée"]])}
  ${s.contratType==="determinee"?cautionField("Date de fin du contrat","contratFin","date"):""}
  ${cautionField("Rémunération mensuelle nette (€)","remuneration","number")}${cautionField("Autre(s) revenu(s) mensuel(s) (€)","autresRevenus","number")}
  </div><div class="section">Informations de l'acte</div><div class="grid">
  ${cautionField("Date du bail","dateBail","date")}${cautionField("Trimestre IRL (1 à 4)","irlTrimestre","number")}
  ${cautionField("Année IRL","irlAnnee","number")}${cautionField("Valeur IRL","irlValeur")}
  ${cautionField("Nombre de reconductions / renouvellements couverts","reconductions","number")}
  ${cautionField("Durée maximale du cautionnement","dureeMax")}
  <div class="field"><label>Montant maximum garanti</label><input value="${esc(cautionMaxDisplay())}" readonly><small>Calcul automatique : ${data.type==="meuble"?"12":"36"} × (loyer + charges).</small></div>
  <div class="field"><label>Équivalent en années de loyers charges comprises</label><input value="${data.type==="meuble"?"1":"3"}" readonly></div>
  </div>
  <div class="box"><strong>Données reprises du bail</strong><p>Bailleur(s) : ${esc(bailleurs||"Non renseigné")}</p><p>Locataire(s) : ${esc(locataires||"Non renseigné")}</p><p>Bien : ${esc(data.localisation||"Non renseigné")}</p><p>Loyer : ${esc(data.loyer||"")} € — Charges : ${esc(data.chargesMontant||"")} € — Dépôt : ${esc(data.depotGarantie||"")} €</p><p>Prise d'effet : ${esc(data.dateEffet||"")} — Durée : ${esc(data.duree||"")}</p></div>
  <button type="button" class="primary bigAction" id="generateCautionBtn">Générer l'acte de caution</button>`;
$("content").innerHTML=h;$("generateBtn").textContent="Générer l'acte de caution";bind();const gc=$("generateCautionBtn");if(gc)gc.onclick=generateCautionDocument;
}
if(step===10){const b=data.bailleurs.map(p=>p.type==="morale"?p.denomination:[p.nom,p.prenoms].filter(Boolean).join(" ")).filter(Boolean).join(", "),l=data.locataires.map(p=>[p.nom,p.prenoms].filter(Boolean).join(" ")).filter(Boolean).join(", ");h+=dossierControllerHtml()+`<div class="notice">Le bouton « Générer le Word » utilisera automatiquement <strong>${models[modelKey(data)].label}</strong>.</div><div class="box"><strong>Bailleur(s)</strong><p>${esc(b||"Non renseigné")}</p><strong>Locataire(s)</strong><p>${esc(l||"Non renseigné")}</p><strong>Logement</strong><p>${esc(data.localisation||"Non renseigné")}</p><strong>Loyer</strong><p>${esc(data.loyer||"Non renseigné")} €</p></div>`}
$("content").innerHTML=h;$("generateBtn").textContent="Générer le Word";bind();const gc=$("generateCautionBtn");if(gc)gc.onclick=generateCautionDocument}
function bind(){document.querySelectorAll("[data-goto-step]").forEach(x=>x.onclick=()=>{step=+x.dataset.gotoStep;cautionView=false;render()});document.querySelectorAll("[data-step]").forEach(x=>x.onclick=()=>{step=+x.dataset.step;render()});document.querySelectorAll("[data-key]").forEach(x=>x.oninput=x.onchange=()=>{data[x.dataset.key]=x.value;if(["irlTrimestre","irlAnnee"].includes(x.dataset.key)){syncIrlValue();render()}else if(["destination","chauffageEnergie","eauEnergie"].includes(x.dataset.key))render()});document.querySelectorAll('input[name="agence"]').forEach(x=>x.onchange=()=>{data.agence=x.value;render()});document.querySelectorAll('input[name="type"]').forEach(x=>x.onchange=()=>{data.type=x.value;render()});document.querySelectorAll('input[name="gestion"]').forEach(x=>x.onchange=()=>{data.gestion=x.value;render()});document.querySelectorAll("[data-b]").forEach(x=>x.oninput=()=>{const[i,k]=x.dataset.b.split(":");data.bailleurs[+i][k]=x.value});document.querySelectorAll("[data-btype]").forEach(x=>x.onchange=()=>{data.bailleurs[+x.dataset.btype].type=x.value;render()});document.querySelectorAll("[data-l]").forEach(x=>x.oninput=()=>{const[i,k]=x.dataset.l.split(":");data.locataires[+i][k]=x.value});document.querySelectorAll("[data-addb]").forEach(x=>x.onclick=()=>{data.bailleurs.push(blankBailleur());render()});document.querySelectorAll("[data-addl]").forEach(x=>x.onclick=()=>{data.locataires.push(blankLoc());render()});document.querySelectorAll("[data-rmb]").forEach(x=>x.onclick=()=>{data.bailleurs.splice(+x.dataset.rmb,1);render()});document.querySelectorAll("[data-rml]").forEach(x=>x.onclick=()=>{data.locataires.splice(+x.dataset.rml,1);render()});document.querySelectorAll("[data-annexe]").forEach(x=>x.onchange=()=>data.annexes[x.dataset.annexe]=x.checked);document.querySelectorAll("[data-avcharge]").forEach(x=>x.onchange=()=>data.avenantCharges[x.dataset.avcharge]=x.checked);document.querySelectorAll("[data-avinfo]").forEach(x=>x.onchange=()=>data.avenantInfos[x.dataset.avinfo]=x.checked);document.querySelectorAll("[data-group]").forEach(x=>x.onchange=()=>{data[x.dataset.group]=data[x.dataset.group]||{};data[x.dataset.group][x.dataset.option]=x.checked});document.querySelectorAll("[data-number-group]").forEach(x=>x.oninput=()=>{data[x.dataset.numberGroup]=data[x.dataset.numberGroup]||{};data[x.dataset.numberGroup][x.dataset.option]=x.value});document.querySelectorAll('input[name="cautionMode"]').forEach(x=>x.onchange=()=>{ensureCaution().mode=x.value;render()});document.querySelectorAll("[data-caution]").forEach(x=>x.oninput=x.onchange=()=>{ensureCaution()[x.dataset.caution]=x.value;if(["irlTrimestre","irlAnnee"].includes(x.dataset.caution)){syncIrlValue();render()}else if(x.dataset.caution==="contratType")render()});document.querySelectorAll("#content input,#content textarea,#content select").forEach(el=>{el.addEventListener("input",scheduleAutosave);el.addEventListener("change",scheduleAutosave)})}
function dbOpen(){return new Promise((res,rej)=>{const q=indexedDB.open("action-immo-location-models",1);q.onupgradeneeded=()=>q.result.createObjectStore("files");q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error)})}
async function putFile(key,file){const db=await dbOpen(),bytes=new Uint8Array(await file.arrayBuffer());return new Promise((res,rej)=>{const tx=db.transaction("files","readwrite");tx.objectStore("files").put({name:file.name,bytes},"model_"+key);tx.oncomplete=res;tx.onerror=()=>rej(tx.error)})}
async function getFile(key){try{const db=await dbOpen();return await new Promise((res,rej)=>{const tx=db.transaction("files","readonly"),q=tx.objectStore("files").get("model_"+key);q.onsuccess=()=>res(q.result||null);q.onerror=()=>rej(q.error)})}catch{return null}}
async function renderTemplates(){let h="";for(const [k,v] of Object.entries({...models,...extraModels})){const f=await getFile(k);h+=`<div class="templateRow"><div><strong>${v.label}</strong><small>${f?'<span class="ready">Modèle chargé : '+f.name+'</span>':'<span class="missing">Modèle à charger</span>'}</small></div><div><button data-load="${k}">${f?"Remplacer":"Charger"}</button><input type="file" accept=".docx" id="file-${k}"></div></div>`} $("templates").innerHTML=h;document.querySelectorAll("[data-load]").forEach(b=>b.onclick=()=>document.getElementById("file-"+b.dataset.load).click());for(const k of Object.keys({...models,...extraModels})){const inp=document.getElementById("file-"+k);inp.onchange=async()=>{if(!inp.files[0])return;await putFile(k,inp.files[0]);$("status").textContent="Modèle enregistré";await renderTemplates()}}}
$("modelsBtn").onclick=async()=>{$("modelsModal").classList.remove("hidden");await renderTemplates()};$("closeModels").onclick=()=>$("modelsModal").classList.add("hidden");
$("closeNewDossier").onclick=()=>$("newDossierModal").classList.add("hidden");
$("blankDossierBtn").onclick=startBlankDossier;
$("importDossierBtn").onclick=()=>$("existingLeaseInput").click();
$("closeImportResult").onclick=()=>{$("importResultModal").classList.add("hidden");render()};
$("existingLeaseInput").onchange=async()=>{
  const file=$("existingLeaseInput").files[0];if(!file)return;
  try{
    $("status").textContent="Lecture du bail existant…";
    const bytes=new Uint8Array(await file.arrayBuffer());
    const isPdf=file.type==="application/pdf"||/\.pdf$/i.test(file.name);
    $("status").textContent=isPdf?"Lecture du PDF existant…":"Lecture du bail Word existant…";
    const imported=isPdf?await extractExistingLeasePdf(bytes):await extractExistingLease(bytes);
    data=sanitizeImportedData(imported);
    currentId="loc_"+Date.now()+"_"+Math.random().toString(36).slice(2,8);
    step=0;viewMode="editor";
    $("newDossierModal").classList.add("hidden");
    $("importSummary").innerHTML=importSummaryHtml(data);
    $("importResultModal").classList.remove("hidden");
    $("status").textContent="Informations du bail récupérées";
    $("existingLeaseInput").value="";
  }catch(e){
    $("status").textContent="";
    $("existingLeaseInput").value="";
    alert("Impossible de récupérer les informations de ce bail : "+e.message+"\n\nTu peux utiliser un bail Word .docx ou un PDF. Pour un PDF scanné, la lecture peut prendre un peu plus de temps.");
  }
};
function downloadGeneratedDoc(blob,filename){
  lastGeneratedBlob=blob;lastGeneratedFilename=filename;
  $("status").textContent="✓ Document généré — prêt à enregistrer";
  let box=document.getElementById("generatedDownload");
  if(!box){box=document.createElement("div");box.id="generatedDownload";box.className="notice";box.style.marginTop="16px";$("content").prepend(box)}
  box.innerHTML='<strong>✓ Word généré.</strong><p>Clique sur le bouton ci-dessous pour choisir où enregistrer le fichier sur ton ordinateur.</p><button type="button" class="primary bigAction" id="saveGeneratedWordBtn">Enregistrer le Word</button><button type="button" id="openGeneratedWordBtn" style="margin-left:8px">Ouvrir / télécharger autrement</button><p class="hint" id="generatedSaveHint"></p>';

  const saveBtn=document.getElementById("saveGeneratedWordBtn");
  const openBtn=document.getElementById("openGeneratedWordBtn");
  const hint=document.getElementById("generatedSaveHint");

  if(saveBtn)saveBtn.onclick=async()=>{
    try{
      if(window.showSaveFilePicker){
        const handle=await window.showSaveFilePicker({
          suggestedName:lastGeneratedFilename,
          types:[{description:"Document Word",accept:{"application/vnd.openxmlformats-officedocument.wordprocessingml.document":[".docx"]}}]
        });
        const writable=await handle.createWritable();
        await writable.write(lastGeneratedBlob);
        await writable.close();
        $("status").textContent="✓ Word enregistré";
        if(hint)hint.textContent="Le fichier a bien été enregistré.";
        return;
      }
      const url=URL.createObjectURL(lastGeneratedBlob);
      window.location.href=url;
      setTimeout(()=>URL.revokeObjectURL(url),60000);
    }catch(e){
      if(e&&e.name==="AbortError"){if(hint)hint.textContent="Enregistrement annulé.";return}
      if(hint)hint.textContent="Impossible d'ouvrir la fenêtre d'enregistrement : "+e.message;
    }
  };

  if(openBtn)openBtn.onclick=()=>{
    try{
      const url=URL.createObjectURL(lastGeneratedBlob);
      const w=window.open(url,"_blank");
      if(!w)window.location.href=url;
      if(hint)hint.textContent="Si un nouvel onglet s'ouvre, utilise Ctrl+S pour enregistrer le fichier.";
      setTimeout(()=>URL.revokeObjectURL(url),10*60*1000);
    }catch(e){
      if(hint)hint.textContent="Impossible d'ouvrir le fichier : "+e.message;
    }
  };
}

async function generateCautionDocument(){
  try{
    if(viewMode==="dashboard"){alert("Ouvre d’abord un dossier.");return}
    ensureCaution();saveDossier();
    const key=cautionModelKey(data),f=await getFile(key);
    if(!f){$("modelsModal").classList.remove("hidden");await renderTemplates();alert("Charge d’abord le modèle Word : "+extraModels[key].label);return}
    $("status").textContent="Génération de l'acte…";
    const bytes=f.bytes instanceof Uint8Array?f.bytes:new Uint8Array(f.bytes),blob=await buildCaution(bytes,data);
    const n=(data.cautionActe.nom||"CAUTION").replace(/[^A-Za-zÀ-ÿ0-9]+/g,"_").replace(/^_+|_+$/g,"");
    downloadGeneratedDoc(blob,"ACTE_DE_CAUTION_"+(data.agence||"brest").toUpperCase()+"_"+(((data.cautionActe||{}).mode||"location").toUpperCase())+"_"+n+".docx");
  }catch(e){alert("Impossible de générer l'acte de caution : "+e.message)}
}
$("saveBtn").onclick=()=>{if(viewMode==="dashboard")return;clearTimeout(autosaveTimer);saveDossier("✓ Dossier sauvegardé")};$("dashboardBtn").onclick=()=>{cautionView=false;viewMode="dashboard";renderDashboard()};$("cautionTopBtn").onclick=()=>{if(viewMode==="dashboard"){alert("Ouvre d’abord un dossier.");return}cautionView=true;render()};$("newBtn").onclick=newDossier;
$("generateBtn").onclick=async()=>{try{if(viewMode==="dashboard"){alert("Ouvre d’abord un dossier.");return}if(cautionView){await generateCautionDocument();return}const blocking=dossierChecks().filter(x=>x.level==="error");if(blocking.length&&!confirm("Le contrôle du dossier signale "+blocking.length+" point(s) bloquant(s).\n\nGénérer quand même le contrat ?"))return;const key=modelKey(data),f=await getFile(key);if(!f){$("modelsModal").classList.remove("hidden");await renderTemplates();alert("Charge d’abord le modèle Word : "+models[key].label);return}saveDossier();$("status").textContent="Génération du contrat…";const bytes=f.bytes instanceof Uint8Array?f.bytes:new Uint8Array(f.bytes),blob=await buildLocation(bytes,data);downloadGeneratedDoc(blob,"CONTRAT_LOCATION_"+(data.agence||"brest").toUpperCase()+"_"+data.type.toUpperCase()+"_"+(data.gestion==="gestion"?"GESTION":"HORS_GESTION")+"_REMPLI.docx")}catch(e){alert("Impossible de générer le Word : "+e.message)}};
$("prev").onclick=()=>{if(step>0){step--;render()}};$("next").onclick=()=>{if(step<sections.length-1){step++;render()}};
renderDashboard();