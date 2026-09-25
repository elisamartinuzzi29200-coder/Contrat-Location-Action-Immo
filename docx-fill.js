const te=new TextEncoder(),td=new TextDecoder();
const u16=(b,o)=>b[o]|b[o+1]<<8,u32=(b,o)=>(b[o]|b[o+1]<<8|b[o+2]<<16|b[o+3]<<24)>>>0,p16=v=>[v&255,v>>>8&255],p32=v=>[v&255,v>>>8&255,v>>>16&255,v>>>24&255];
function crc32(d){let c=0xffffffff;for(const x of d){c^=x;for(let k=0;k<8;k++)c=c>>>1^(0xedb88320&-(c&1))}return(c^0xffffffff)>>>0}
async function inflate(d){const ds=new DecompressionStream("deflate-raw");return new Uint8Array(await new Response(new Blob([d]).stream().pipeThrough(ds)).arrayBuffer())}
async function unzip(b){let e=-1;for(let i=b.length-22;i>=Math.max(0,b.length-65557);i--)if(u32(b,i)===0x06054b50){e=i;break}if(e<0)throw Error("DOCX invalide");const n=u16(b,e+10);let p=u32(b,e+16),f=[];for(let i=0;i<n;i++){const m=u16(b,p+10),cs=u32(b,p+20),nl=u16(b,p+28),el=u16(b,p+30),cl=u16(b,p+32),lo=u32(b,p+42),name=td.decode(b.slice(p+46,p+46+nl)),ln=u16(b,lo+26),le=u16(b,lo+28),st=lo+30+ln+le,comp=b.slice(st,st+cs);f.push({name,data:m===0?comp:m===8?await inflate(comp):(()=>{throw Error("Compression non prise en charge")})()});p+=46+nl+el+cl}return f}
function zip(files){let locals=[],centrals=[],off=0;for(const f of files){const n=te.encode(f.name),crc=crc32(f.data),lh=new Uint8Array([...p32(0x04034b50),...p16(20),...p16(0),...p16(0),...p16(0),...p16(0),...p32(crc),...p32(f.data.length),...p32(f.data.length),...p16(n.length),...p16(0)]),l=new Uint8Array(lh.length+n.length+f.data.length);l.set(lh);l.set(n,lh.length);l.set(f.data,lh.length+n.length);locals.push(l);const ch=new Uint8Array([...p32(0x02014b50),...p16(20),...p16(20),...p16(0),...p16(0),...p16(0),...p16(0),...p32(crc),...p32(f.data.length),...p32(f.data.length),...p16(n.length),...p16(0),...p16(0),...p16(0),...p16(0),...p32(0),...p32(off)]),c=new Uint8Array(ch.length+n.length);c.set(ch);c.set(n,ch.length);centrals.push(c);off+=l.length}const cs=centrals.reduce((s,x)=>s+x.length,0),end=new Uint8Array([...p32(0x06054b50),...p16(0),...p16(0),...p16(files.length),...p16(files.length),...p32(cs),...p32(off),...p16(0)]),out=new Uint8Array(off+cs+end.length);let q=0;for(const x of locals){out.set(x,q);q+=x.length}for(const x of centrals){out.set(x,q);q+=x.length}out.set(end,q);return out}
const W="http://schemas.openxmlformats.org/wordprocessingml/2006/main",WP="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
function setBox(doc,name,occ,value,prepend=false){const d=[...doc.getElementsByTagNameNS(WP,"docPr")].filter(x=>x.getAttribute("name")===name)[occ];if(!d)return;let a=d.parentElement;while(a&&a.localName!=="anchor")a=a.parentElement;if(!a)return;const b=a.getElementsByTagNameNS(W,"txbxContent")[0];if(!b)return;let p;if(prepend){p=doc.createElementNS(W,"w:p");b.insertBefore(p,b.firstChild)}else{p=b.getElementsByTagNameNS(W,"p")[0]||doc.createElementNS(W,"w:p");if(!p.parentNode)b.appendChild(p);[...p.childNodes].forEach(x=>{if(x.localName!=="pPr")p.removeChild(x)})}const r=doc.createElementNS(W,"w:r");String(value||"").split("\n").forEach((s,i)=>{if(i)r.appendChild(doc.createElementNS(W,"w:br"));const t=doc.createElementNS(W,"w:t");t.setAttributeNS("http://www.w3.org/XML/1998/namespace","xml:space","preserve");t.textContent=s;r.appendChild(t)});p.appendChild(r)}
function setChoice(doc,needle,selected,count=2){const ps=[...doc.getElementsByTagNameNS(W,"p")].filter(p=>(p.textContent||"").includes(needle));const p=ps[0];if(!p)return;let idx=0;for(const t of p.getElementsByTagNameNS(W,"t")){if(/[☐☒]/.test(t.textContent||"")){t.textContent=(t.textContent||"").replace(/[☐☒]/,idx===selected?"☒":"☐");idx++;if(idx>=count)break}}}
function setCheck(doc,needle,on,which=0){const ps=[...doc.getElementsByTagNameNS(W,"p")].filter(p=>(p.textContent||"").includes(needle));const p=ps[0];if(!p)return;let idx=0;for(const t of p.getElementsByTagNameNS(W,"t")){if(/[☐☒]/.test(t.textContent||"")){if(idx===which){t.textContent=(t.textContent||"").replace(/[☐☒]/,on?"☒":"☐");return}idx++}}}
function setTableCell(doc,ti,ri,ci,value){const tbl=doc.getElementsByTagNameNS(W,"tbl")[ti];if(!tbl)return;const rows=[...tbl.children].filter(x=>x.localName==="tr"),r=rows[ri];if(!r)return;const cells=[...r.children].filter(x=>x.localName==="tc"),c=cells[ci];if(!c)return;let p=c.getElementsByTagNameNS(W,"p")[0]||doc.createElementNS(W,"w:p");if(!p.parentNode)c.appendChild(p);[...p.childNodes].forEach(x=>{if(x.localName!=="pPr")p.removeChild(x)});const rr=doc.createElementNS(W,"w:r"),tt=doc.createElementNS(W,"w:t");tt.textContent=value||"";rr.appendChild(tt);p.appendChild(rr)}
function partyName(p){return p.type==="morale"?(p.denomination||"").trim():((p.nom||"")+" "+(p.prenoms||"")).trim()}
function bailleurText(p){return p.type==="morale"?`${p.denomination||"___"}, personne morale, siège social : ${p.adresse||"___"}`:`${(p.nom||"").toUpperCase()} ${p.prenoms||""}, demeurant ${p.adresse||"___"}`}
function money(v){return String(v||"").trim()}
function num(v){return Number(String(v||"").replace(/\s/g,"").replace(/€/g,"").replace(",",".").replace(/[^0-9.-]/g,""))||0}
function numberToFrench(n){
  n=Math.round(n);
  if(n===0)return "zéro";
  const u=["","un","deux","trois","quatre","cinq","six","sept","huit","neuf","dix","onze","douze","treize","quatorze","quinze","seize"];
  const under100=x=>{
    if(x<17)return u[x];
    if(x<20)return "dix-"+u[x-10];
    const t=Math.floor(x/10),r=x%10;
    if(t===7)return "soixante-"+under100(10+r);
    if(t===9)return "quatre-vingt-"+under100(10+r);
    const names={2:"vingt",3:"trente",4:"quarante",5:"cinquante",6:"soixante",8:"quatre-vingt"};
    let s=names[t]||"";
    if(r===1&&t!==8)s+=" et un"; else if(r)s+="-"+u[r];
    if(t===8&&r===0)s+="s";
    return s;
  };
  const under1000=x=>{
    if(x<100)return under100(x);
    const h=Math.floor(x/100),r=x%100;
    let s=h===1?"cent":u[h]+" cent";
    if(r===0&&h>1)s+="s";
    return r?s+" "+under100(r):s;
  };
  const parts=[];
  const milliards=Math.floor(n/1e9);n%=1e9;
  const millions=Math.floor(n/1e6);n%=1e6;
  const milliers=Math.floor(n/1000);n%=1000;
  if(milliards)parts.push((milliards===1?"un":under1000(milliards))+" milliard"+(milliards>1?"s":""));
  if(millions)parts.push((millions===1?"un":under1000(millions))+" million"+(millions>1?"s":""));
  if(milliers)parts.push((milliers===1?"":under1000(milliers)+" ")+"mille");
  if(n)parts.push(under1000(n));
  return parts.join(" ");
}
function moneyWords(v){
  const n=num(v);
  if(!n&&String(v||"").trim()==="")return "";
  return Math.round(n).toLocaleString("fr-FR")+" € ("+numberToFrench(n)+" euros)";
}
function setBailleurTypes(doc,bailleurs){
  const p=[...doc.getElementsByTagNameNS(W,"p")].find(p=>(p.textContent||"").includes("Personne physique")&&(p.textContent||"").includes("Personne morale"));
  if(!p)return;
  const physical=(bailleurs||[]).some(x=>(x.type||"physique")==="physique");
  const moral=(bailleurs||[]).some(x=>x.type==="morale");
  let i=0;
  for(const t of p.getElementsByTagNameNS(W,"t")){
    if(/[☐☒]/.test(t.textContent||"")){
      const on=i===0?physical:moral;
      t.textContent=(t.textContent||"").replace(/[☐☒]/,on?"☒":"☐");
      if(++i>=2)break;
    }
  }
}
function getBoxText(doc,name,occ=0){
  const d=[...doc.getElementsByTagNameNS(WP,"docPr")].filter(x=>x.getAttribute("name")===name)[occ];
  if(!d)return "";
  let a=d.parentElement;while(a&&a.localName!=="anchor")a=a.parentElement;
  if(!a)return "";
  const b=a.getElementsByTagNameNS(W,"txbxContent")[0];
  return b?(b.textContent||"").replace(/\s+/g," ").trim():"";
}
function readChoice(doc,needle,count=2){
  const p=[...doc.getElementsByTagNameNS(W,"p")].find(p=>(p.textContent||"").includes(needle));
  if(!p)return -1;
  let idx=0;
  for(const t of p.getElementsByTagNameNS(W,"t")){
    const s=t.textContent||"";
    if(/[☐☒]/.test(s)){
      if(s.includes("☒"))return idx;
      idx++;if(idx>=count)break;
    }
  }
  return -1;
}
function readCheck(doc,needle){
  const p=[...doc.getElementsByTagNameNS(W,"p")].find(p=>(p.textContent||"").includes(needle));
  if(!p)return false;
  return (p.textContent||"").includes("☒");
}
function extractBailleurs(doc){
  const txt=getBoxText(doc,"Zone de texte 2",0);
  const names=getBoxText(doc,"Zone de texte 219",0);
  if(!txt&&!names)return [ {type:"physique",nom:"",prenoms:"",denomination:"",adresse:"",email:"",tel:""} ];
  const chunks=(txt||names).split(/\n|D’une part,|D'UNE PART/i).map(x=>x.trim()).filter(Boolean);
  const out=[];
  for(const raw of chunks){
    if(/personne morale|SOCI[ÉE]T[ÉE]/i.test(raw)){
      const dm=raw.match(/(?:SOCI[ÉE]T[ÉE]\s+)?([^,]+).*?(?:si[eè]ge social\s*:\s*|si[eè]ge social\s+)(.+?)(?:\.|$)/i);
      out.push({type:"morale",nom:"",prenoms:"",denomination:(dm&&dm[1]?dm[1].trim():raw.split(",")[0].trim()),adresse:(dm&&dm[2]?dm[2].trim():""),email:"",tel:""});
    }else{
      const m=raw.match(/^([^,]+?)(?:,\s*demeurant\s+(.+?))(?:\.|$)/i);
      const full=(m&&m[1]?m[1]:(names||raw)).trim(),parts=full.split(/\s+/),nom=parts.shift()||"",prenoms=parts.join(" ");
      out.push({type:"physique",nom,prenoms,denomination:"",adresse:(m&&m[2]?m[2].trim():""),email:"",tel:""});
    }
  }
  return out.length?out:[{type:"physique",nom:names,prenoms:"",denomination:"",adresse:"",email:"",tel:""}];
}
function readTableCell(doc,ti,ri,ci){
  const tbl=doc.getElementsByTagNameNS(W,"tbl")[ti];if(!tbl)return "";
  const rows=[...tbl.children].filter(x=>x.localName==="tr"),r=rows[ri];if(!r)return "";
  const cells=[...r.children].filter(x=>x.localName==="tc"),c=cells[ci];
  return c?(c.textContent||"").replace(/\s+/g," ").trim():"";
}
async function extractExistingLease(bytes){
  const files=await unzip(bytes),xf=files.find(f=>f.name==="word/document.xml");
  if(!xf)throw Error("Ce fichier Word ne ressemble pas à un bail compatible.");
  const doc=new DOMParser().parseFromString(td.decode(xf.data),"application/xml"),all=(doc.documentElement.textContent||"");
  const type=/LOGEMENT\s+MEUBL[ÉE]/i.test(all)?"meuble":"nu";
  const habitatChoice=readChoice(doc,"Type d’habitat",2),regimeChoice=readChoice(doc,"Régime juridique de l’immeuble",2),periodChoice=readChoice(doc,"Avant 1949",5),destChoice=readChoice(doc,"À usage exclusif d’habitation principale",2),chauffChoice=readChoice(doc,"Modalité de répartition du chauffage",2),eauChoice=readChoice(doc,"Modalité de répartition de l’eau chaude",2);
  const periods=["avant1949","1949-1974","1975-1989","1989-2005","depuis2005"];
  const chargesMode=readCheck(doc,"Provision mensuelle")?"provision":readCheck(doc,"Forfait d’un montant")?"forfait":readCheck(doc,"Remboursement sur justificatif")?"justificatif":"provision";
  const gestion=/Lieu de paiement[\s\S]{0,120}Agence Action Immobili[eè]re/i.test(all)&&/☒\s*Agence Action Immobili[eè]re/i.test(all)?"gestion":"hors";
  const out={
    type,gestion,bailleurs:extractBailleurs(doc),locataires:[{nom:"",prenoms:"",naissance:"",lieuNaissance:"",email:"",tel:""}],
    localisation:getBoxText(doc,"Zone de texte 2",1),habitat:habitatChoice===1?"individuel":"collectif",
    identifiantFiscal:getBoxText(doc,"Zone de texte 126861955",0),regime:regimeChoice===1?"monopropriete":"copropriete",
    periode:periodChoice>=0?periods[periodChoice]:"depuis2005",surface:getBoxText(doc,"Zone de texte 660593113",0),
    pieces:"",caracteristiques:getBoxText(doc,"Zone de texte 507952355",0),autresParties:getBoxText(doc,"Zone de texte 1328110326",0),
    equipements:getBoxText(doc,"Zone de texte 868199202",0),chauffageMode:chauffChoice===1?"collectif":"individuel",chauffageAutre:getBoxText(doc,"Zone de texte 1728029714",0),
    eauMode:eauChoice===1?"collectif":"individuel",eauAutre:getBoxText(doc,"Zone de texte 85956290",0),destination:destChoice===1?"mixte":"habitation",
    professionMixte:getBoxText(doc,"Zone de texte 1029441271",0),accessoiresPrivatifs:getBoxText(doc,"Zone de texte 1625481562",0),partiesCommunes:getBoxText(doc,"Zone de texte 6783576",0),technologies:getBoxText(doc,"Zone de texte 630456886",0),
    depensesEnergie:getBoxText(doc,"Zone de texte 561389577",0),anneeEnergie:getBoxText(doc,"Zone de texte 374328480",0),
    dateEffet:"",duree:getBoxText(doc,"Zone de texte 998965608",0),raisonDureeReduite:getBoxText(doc,"Zone de texte 760777164",0),
    loyer:"",decretRelocation:readChoice(doc,"décret fixant annuellement",2)===0?"oui":"non",encadrement:readChoice(doc,"loyer de référence majoré",2)===0?"oui":"non",
    loyerReference:"",loyerReferenceMajore:"",loyerBase:"",complementLoyer:"",dernierLoyer:"",dateVersementDernier:"",dateDerniereRevision:"",
    dateRevision:"",irl:"",chargesMode,chargesMontant:getBoxText(doc,"Zone de texte 2048181660",0)||getBoxText(doc,"Zone de texte 792361571",0),
    contribution:getBoxText(doc,"Zone de texte 1214938885",0),justifContribution:getBoxText(doc,"Zone de texte 904827124",0),
    assuranceColocAnnuelle:getBoxText(doc,"Zone de texte 1352093729",0),assuranceColocMensuelle:getBoxText(doc,"Zone de texte 1942955975",0),
    depotGarantie:"",honorairesVisiteBailleur:readTableCell(doc,2,1,1),honorairesVisiteLocataire:readTableCell(doc,2,1,2),honorairesEdlBailleur:readTableCell(doc,2,2,1),honorairesEdlLocataire:readTableCell(doc,2,2,2),
    travauxRecents:getBoxText(doc,"Zone de texte 2135196039",0),majorationTravaux:getBoxText(doc,"Zone de texte 1974909847",0),diminutionTravaux:getBoxText(doc,"Zone de texte 1136262404",0),
    sinistre:readChoice(doc,"a-t-il subi un sinistre",2)===0?"oui":"non",congeLocataire:getBoxText(doc,"Zone de texte 520617941",0),
    conditionsLocataire:getBoxText(doc,"Zone de texte 1418093311",0),conditionsBailleur:getBoxText(doc,"Zone de texte 1121375395",0),caution:getBoxText(doc,"Zone de texte 676414816",0),
    annexes:{},avenantCharges:{},avenantInfos:{}
  };
  const annexLabels=["Un extrait du règlement concernant la destination de l’immeuble","Le règlement intérieur de l’immeuble","Un document informatif sur les risques de nuisances sonores aériennes","Un diagnostic de performance énergétique","Un constat de risque d‘exposition au plomb","Une copie d’un état mentionnant l’absence ou la présence de matériaux","Un état de l’installation intérieure d’électricité et de gaz","Un état des risques naturels et technologiques","Une notice d’information relative aux droits et obligations","Un état des lieux","Une autorisation préalable de mise en location","Les références aux loyers habituellement constatés","Une grille de vétusté"];
  for(const x of annexLabels)out.annexes[x]=readCheck(doc,x);
  return out;
}

async function buildLocation(templateBytes,data){const files=await unzip(templateBytes),xf=files.find(f=>f.name==="word/document.xml");if(!xf)throw Error("Modèle Word incomplet");const doc=new DOMParser().parseFromString(td.decode(xf.data),"application/xml");
const bailNames=data.bailleurs.map(partyName).filter(Boolean).join(" / "),locNames=data.locataires.map(partyName).filter(Boolean).join(" / ");
setBox(doc,"Zone de texte 219",0,bailNames);setBox(doc,"Zone de texte 218",0,locNames);
setBox(doc,"Zone de texte 2",0,data.bailleurs.map(bailleurText).join("\n"),true);setBailleurTypes(doc,data.bailleurs);
if(data.gestion==="hors"){setBox(doc,"Zone de texte 478161165",0,data.bailleurs.map(x=>x.email).filter(Boolean).join(" / "));setBox(doc,"Zone de texte 1135107845",0,data.bailleurs.map(x=>x.tel).filter(Boolean).join(" / "))}
data.locataires.slice(0,4).forEach((p,i)=>{setTableCell(doc,0,i+1,0,partyName(p));setTableCell(doc,0,i+1,1,[p.naissance,p.lieuNaissance].filter(Boolean).join(" à "));setTableCell(doc,0,i+1,2,p.email);setTableCell(doc,0,i+1,3,p.tel)});
setBox(doc,"Zone de texte 2",1,data.localisation);setChoice(doc,"Type d’habitat",data.habitat==="individuel"?1:0);setBox(doc,"Zone de texte 126861955",0,data.identifiantFiscal);setChoice(doc,"Régime juridique de l’immeuble",data.regime==="monopropriete"?1:0);
const periods=["avant1949","1949-1974","1975-1989","1989-2005","depuis2005"];setChoice(doc,"Avant 1949",Math.max(0,periods.indexOf(data.periode)),5);
setBox(doc,"Zone de texte 660593113",0,data.surface);setBox(doc,"Zone de texte 507952355",0,data.caracteristiques);setBox(doc,"Zone de texte 1328110326",0,data.autresParties);setBox(doc,"Zone de texte 868199202",0,data.equipements);setBox(doc,"Zone de texte 1728029714",0,data.chauffageAutre);setBox(doc,"Zone de texte 85956290",0,data.eauAutre);
setChoice(doc,"Modalité de répartition du chauffage",data.chauffageMode==="collectif"?1:0);setChoice(doc,"Modalité de répartition de l’eau chaude",data.eauMode==="collectif"?1:0);
setChoice(doc,"À usage exclusif d’habitation principale",data.destination==="mixte"?1:0);setBox(doc,"Zone de texte 1029441271",0,data.professionMixte);setBox(doc,"Zone de texte 1625481562",0,data.accessoiresPrivatifs);setBox(doc,"Zone de texte 6783576",0,data.partiesCommunes);setBox(doc,"Zone de texte 630456886",0,data.technologies);
setBox(doc,"Zone de texte 2077289405",0,data.dateEffet);setBox(doc,"Zone de texte 998965608",0,data.duree);setBox(doc,"Zone de texte 760777164",0,data.raisonDureeReduite);
setBox(doc,"Zone de texte 628403819",0,moneyWords(data.loyer));setChoice(doc,"décret fixant annuellement",data.decretRelocation==="oui"?0:1);setChoice(doc,"loyer de référence majoré",data.encadrement==="oui"?0:1);setBox(doc,"Zone de texte 268514384",0,money(data.loyerBase));setBox(doc,"Zone de texte 475039207",0,money(data.complementLoyer));setBox(doc,"Zone de texte 1040080526",0,money(data.dernierLoyer));setBox(doc,"Zone de texte 31722978",0,data.dateVersementDernier);setBox(doc,"Zone de texte 1259670285",0,data.dateDerniereRevision);setBox(doc,"Zone de texte 1359800534",0,data.dateRevision);setBox(doc,"Zone de texte 85742337",0,data.irl);
if(data.chargesMode==="provision"){setCheck(doc,"Provision mensuelle",true,0);setCheck(doc,"Forfait d’un montant",false,0);setCheck(doc,"Remboursement sur justificatif",false,0);setBox(doc,"Zone de texte 2048181660",0,money(data.chargesMontant))}
if(data.chargesMode==="forfait"){setCheck(doc,"Provision mensuelle",false,0);setCheck(doc,"Forfait d’un montant",true,0);setCheck(doc,"Remboursement sur justificatif",false,0);setBox(doc,"Zone de texte 2048181660",0,"");setBox(doc,"Zone de texte 792361571",0,money(data.chargesMontant))}
if(data.chargesMode==="justificatif"){setCheck(doc,"Provision mensuelle",false,0);setCheck(doc,"Forfait d’un montant",false,0);setCheck(doc,"Remboursement sur justificatif",true,0)}
setBox(doc,"Zone de texte 1214938885",0,data.contribution);setBox(doc,"Zone de texte 904827124",0,data.justifContribution);setBox(doc,"Zone de texte 1352093729",0,money(data.assuranceColocAnnuelle));setBox(doc,"Zone de texte 1942955975",0,money(data.assuranceColocMensuelle));
setTableCell(doc,1,0,1,num(data.loyer)?num(data.loyer).toLocaleString("fr-FR")+" €":"");setTableCell(doc,1,1,1,num(data.chargesMontant)?num(data.chargesMontant).toLocaleString("fr-FR")+" €":"");let row=2;if(data.type==="nu"){setTableCell(doc,1,row++,1,num(data.contribution)?num(data.contribution).toLocaleString("fr-FR")+" €":"");}setTableCell(doc,1,row++,1,num(data.assuranceColocMensuelle)?num(data.assuranceColocMensuelle).toLocaleString("fr-FR")+" €":"");const total=num(data.loyer)+num(data.chargesMontant)+num(data.assuranceColocMensuelle)+(data.type==="nu"?num(data.contribution):0);setTableCell(doc,1,row,1,total?total.toLocaleString("fr-FR")+" €":"");
setBox(doc,"Zone de texte 561389577",0,data.depensesEnergie);setBox(doc,"Zone de texte 374328480",0,data.anneeEnergie);setBox(doc,"Zone de texte 2135196039",0,data.travauxRecents);setBox(doc,"Zone de texte 1974909847",0,data.majorationTravaux);setBox(doc,"Zone de texte 1136262404",0,data.diminutionTravaux);setBox(doc,"Zone de texte 2071132409",0,moneyWords(data.depotGarantie));setBox(doc,"Zone de texte 520617941",0,data.congeLocataire);setBox(doc,"Zone de texte 1418093311",0,data.conditionsLocataire);setBox(doc,"Zone de texte 1121375395",0,data.conditionsBailleur);setBox(doc,"Zone de texte 676414816",0,data.caution);
const hvb=num(data.honorairesVisiteBailleur),hvl=num(data.honorairesVisiteLocataire),heb=num(data.honorairesEdlBailleur),hel=num(data.honorairesEdlLocataire);
setTableCell(doc,2,1,1,hvb?hvb.toLocaleString("fr-FR")+" €":"");setTableCell(doc,2,1,2,hvl?hvl.toLocaleString("fr-FR")+" €":"");
setTableCell(doc,2,2,1,heb?heb.toLocaleString("fr-FR")+" €":"");setTableCell(doc,2,2,2,hel?hel.toLocaleString("fr-FR")+" €":"");
const totalHB=hvb+heb,totalHL=hvl+hel;
setTableCell(doc,2,3,1,totalHB?totalHB.toLocaleString("fr-FR")+" €":"");setTableCell(doc,2,3,2,totalHL?totalHL.toLocaleString("fr-FR")+" €":"");
setChoice(doc,"a-t-il subi un sinistre",data.sinistre==="oui"?0:1);
for(const [label,key] of Object.entries(data.annexes||{}))setCheck(doc,label,!!key,0);
for(const [label,key] of Object.entries(data.avenantCharges||{}))setCheck(doc,label,!!key,0);
for(const [label,key] of Object.entries(data.avenantInfos||{}))setCheck(doc,label,!!key,0);
xf.data=te.encode(new XMLSerializer().serializeToString(doc));return new Blob([zip(files)],{type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"})}