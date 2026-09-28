const te=new TextEncoder(),td=new TextDecoder();
const u16=(b,o)=>b[o]|b[o+1]<<8,u32=(b,o)=>(b[o]|b[o+1]<<8|b[o+2]<<16|b[o+3]<<24)>>>0,p16=v=>[v&255,v>>>8&255],p32=v=>[v&255,v>>>8&255,v>>>16&255,v>>>24&255];
function crc32(d){let c=0xffffffff;for(const x of d){c^=x;for(let k=0;k<8;k++)c=c>>>1^(0xedb88320&-(c&1))}return(c^0xffffffff)>>>0}
async function inflate(d){const ds=new DecompressionStream("deflate-raw");return new Uint8Array(await new Response(new Blob([d]).stream().pipeThrough(ds)).arrayBuffer())}
async function unzip(b){let e=-1;for(let i=b.length-22;i>=Math.max(0,b.length-65557);i--)if(u32(b,i)===0x06054b50){e=i;break}if(e<0)throw Error("DOCX invalide");const n=u16(b,e+10);let p=u32(b,e+16),f=[];for(let i=0;i<n;i++){const m=u16(b,p+10),cs=u32(b,p+20),nl=u16(b,p+28),el=u16(b,p+30),cl=u16(b,p+32),lo=u32(b,p+42),name=td.decode(b.slice(p+46,p+46+nl)),ln=u16(b,lo+26),le=u16(b,lo+28),st=lo+30+ln+le,comp=b.slice(st,st+cs);f.push({name,data:m===0?comp:m===8?await inflate(comp):(()=>{throw Error("Compression non prise en charge")})()});p+=46+nl+el+cl}return f}
function zip(files){let locals=[],centrals=[],off=0;for(const f of files){const n=te.encode(f.name),crc=crc32(f.data),lh=new Uint8Array([...p32(0x04034b50),...p16(20),...p16(0),...p16(0),...p16(0),...p16(0),...p32(crc),...p32(f.data.length),...p32(f.data.length),...p16(n.length),...p16(0)]),l=new Uint8Array(lh.length+n.length+f.data.length);l.set(lh);l.set(n,lh.length);l.set(f.data,lh.length+n.length);locals.push(l);const ch=new Uint8Array([...p32(0x02014b50),...p16(20),...p16(20),...p16(0),...p16(0),...p16(0),...p16(0),...p32(crc),...p32(f.data.length),...p32(f.data.length),...p16(n.length),...p16(0),...p16(0),...p16(0),...p16(0),...p32(0),...p32(off)]),c=new Uint8Array(ch.length+n.length);c.set(ch);c.set(n,ch.length);centrals.push(c);off+=l.length}const cs=centrals.reduce((s,x)=>s+x.length,0),end=new Uint8Array([...p32(0x06054b50),...p16(0),...p16(0),...p16(files.length),...p16(files.length),...p32(cs),...p32(off),...p16(0)]),out=new Uint8Array(off+cs+end.length);let q=0;for(const x of locals){out.set(x,q);q+=x.length}for(const x of centrals){out.set(x,q);q+=x.length}out.set(end,q);return out}
const W="http://schemas.openxmlformats.org/wordprocessingml/2006/main",WP="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
function setBox(doc,name,occ,value,prepend=false){const d=[...doc.getElementsByTagNameNS(WP,"docPr")].filter(x=>x.getAttribute("name")===name)[occ];if(!d)return;let a=d.parentElement;while(a&&a.localName!=="anchor")a=a.parentElement;if(!a)return;const b=a.getElementsByTagNameNS(W,"txbxContent")[0];if(!b)return;let p;if(prepend){p=doc.createElementNS(W,"w:p");b.insertBefore(p,b.firstChild)}else{p=b.getElementsByTagNameNS(W,"p")[0]||doc.createElementNS(W,"w:p");if(!p.parentNode)b.appendChild(p);[...p.childNodes].forEach(x=>{if(x.localName!=="pPr")p.removeChild(x)})}const r=doc.createElementNS(W,"w:r");String(value||"").split("\n").forEach((s,i)=>{if(i)r.appendChild(doc.createElementNS(W,"w:br"));const t=doc.createElementNS(W,"w:t");t.setAttributeNS("http://www.w3.org/XML/1998/namespace","xml:space","preserve");t.textContent=s;r.appendChild(t)});p.appendChild(r)}

function setBoxNearLabel(doc,label,value,after=true){
  const ps=[...doc.getElementsByTagNameNS(W,"p")];
  const p=ps.find(x=>String(x.textContent||"").toLowerCase().includes(String(label).toLowerCase()));
  if(!p)return false;
  const boxes=[...p.getElementsByTagNameNS(WP,"docPr")];
  if(!boxes.length)return false;
  const d=after?boxes[boxes.length-1]:boxes[0];
  const name=d.getAttribute("name");
  const all=[...doc.getElementsByTagNameNS(WP,"docPr")].filter(x=>x.getAttribute("name")===name);
  const occ=all.indexOf(d);
  if(occ<0)return false;
  setBox(doc,name,occ,value);return true;
}

function setChoice(doc,needle,selected,count=2){const ps=[...doc.getElementsByTagNameNS(W,"p")].filter(p=>(p.textContent||"").includes(needle));const p=ps[0];if(!p)return;let idx=0;for(const t of p.getElementsByTagNameNS(W,"t")){if(/[☐☒]/.test(t.textContent||"")){t.textContent=(t.textContent||"").replace(/[☐☒]/,idx===selected?"☒":"☐");idx++;if(idx>=count)break}}}
function setCheck(doc,needle,on,which=0){const ps=[...doc.getElementsByTagNameNS(W,"p")].filter(p=>(p.textContent||"").includes(needle));const p=ps[0];if(!p)return;let idx=0;for(const t of p.getElementsByTagNameNS(W,"t")){if(/[☐☒]/.test(t.textContent||"")){if(idx===which){t.textContent=(t.textContent||"").replace(/[☐☒]/,on?"☒":"☐");return}idx++}}}

function findParagraphCI(doc,needle){
  const n=String(needle||"").toLowerCase();
  return [...doc.getElementsByTagNameNS(W,"p")].find(p=>String(p.textContent||"").toLowerCase().includes(n));
}



function setSdtCheckboxState(sdt,on){
  const W14="http://schemas.microsoft.com/office/word/2010/wordml";
  const cb=[...sdt.getElementsByTagNameNS(W14,"checkbox")][0];
  if(cb){
    let checked=[...cb.getElementsByTagNameNS(W14,"checked")][0];
    if(!checked){checked=sdt.ownerDocument.createElementNS(W14,"w14:checked");cb.insertBefore(checked,cb.firstChild)}
    checked.setAttributeNS(W14,"w14:val",on?"1":"0");
  }
  for(const t of sdt.getElementsByTagNameNS(W,"t")){
    if(/[☐☒]/.test(t.textContent||""))t.textContent=(t.textContent||"").replace(/[☐☒]/,on?"☒":"☐");
  }
}
function paragraphCheckboxSdts(p){
  const W14="http://schemas.microsoft.com/office/word/2010/wordml";
  return [...p.getElementsByTagNameNS(W,"sdt")].filter(s=>s.getElementsByTagNameNS(W14,"checkbox").length);
}
function nextCheckboxParagraph(doc,labelNeedle){
  const ps=[...doc.getElementsByTagNameNS(W,"p")],n=String(labelNeedle||"").toLowerCase();
  const i=ps.findIndex(p=>String(p.textContent||"").toLowerCase().includes(n));
  if(i<0)return null;
  for(let j=i;j<Math.min(ps.length,i+6);j++)if(paragraphCheckboxSdts(ps[j]).length)return ps[j];
  return null;
}
function setEnergyExact(doc,labelNeedle,mode,energy,customText,boxTitle){
  const p=nextCheckboxParagraph(doc,labelNeedle);if(!p)return false;
  const boxes=paragraphCheckboxSdts(p);
  const states=[mode==="individuel",mode==="collectif",energy==="gaz",energy==="electricite",energy==="autre"];
  boxes.slice(0,5).forEach((s,i)=>setSdtCheckboxState(s,!!states[i]));
  setBox(doc,boxTitle,0,energy==="autre"?(customText||""):"");
  return true;
}
function setLabeledCheckboxExact(doc,labelNeedle,on){
  const ps=[...doc.getElementsByTagNameNS(W,"p")],n=String(labelNeedle||"").toLowerCase();
  for(const p of ps){
    const text=String(p.textContent||"").toLowerCase();
    if(!text.includes(n))continue;
    const sdts=paragraphCheckboxSdts(p),runs=[...p.children];
    for(const sdt of sdts){
      const idx=runs.indexOf(sdt);
      let after="";
      for(let j=idx+1;j<runs.length;j++){
        if(runs[j].localName==="sdt")break;
        after+=runs[j].textContent||"";
      }
      if(after.toLowerCase().includes(n)){setSdtCheckboxState(sdt,on);return true}
    }
  }
  return false;
}

function setParagraphCheckboxes(doc,paragraphNeedle,states){
  const p=findParagraphCI(doc,paragraphNeedle);if(!p)return false;
  let changed=false,idx=0;

  // 1) Cases représentées par des glyphes ☐ / ☒.
  for(const t of p.getElementsByTagNameNS(W,"t")){
    let s=t.textContent||"";
    if(/[☐☒]/.test(s)&&idx<states.length){
      t.textContent=s.replace(/[☐☒]/,states[idx++]?"☒":"☐");
      changed=true;
    }
  }
  if(idx>=states.length)return changed;

  // 2) Anciennes cases formulaire Word (w:checkBox / w:default).
  const checkBoxes=[...p.getElementsByTagNameNS(W,"checkBox")];
  for(const cb of checkBoxes){
    if(idx>=states.length)break;
    let def=cb.getElementsByTagNameNS(W,"default")[0];
    if(!def){def=doc.createElementNS(W,"w:default");cb.appendChild(def)}
    def.setAttributeNS(W,"w:val",states[idx++]?"1":"0");
    changed=true;
  }
  if(idx>=states.length)return changed;

  // 3) Cases modernes Word (w14:checkbox / w14:checked).
  const all=[...p.getElementsByTagName("*")];
  for(const el of all){
    if(idx>=states.length)break;
    if(el.localName==="checkbox"){
      const checked=[...el.getElementsByTagName("*")].find(x=>x.localName==="checked");
      if(checked){
        checked.setAttribute("w14:val",states[idx++]?"1":"0");
        changed=true;
      }
    }
  }
  return changed;
}

function setEnergyLine(doc,paragraphNeedle,mode,energy,customText,boxTitle){
  // Ordre réel des cases dans les modèles : Individuel, Collectif, Gaz, Electrique, Autre.
  const states=[mode==="individuel",mode==="collectif",energy==="gaz",energy==="electricite",energy==="autre"];
  setParagraphCheckboxes(doc,paragraphNeedle,states);
  setCheckboxBeforeBox(doc,boxTitle,energy==="autre");
  setBox(doc,boxTitle,0,energy==="autre"?(customText||""):"");
}

function setCheckboxBeforeBox(doc,boxTitle,on){
  const controls=[...doc.getElementsByTagNameNS(W,"sdt")];
  const box=controls.find(s=>[...s.getElementsByTagNameNS(W,"alias")].some(a=>a.getAttributeNS(W,"val")===boxTitle||a.getAttribute("w:val")===boxTitle));
  if(!box)return false;
  const p=box.closest?box.closest("w\\:p"):null;
  const para=p||(()=>{let n=box.parentNode;while(n&&n.localName!=="p")n=n.parentNode;return n})();
  if(!para)return false;
  const ts=[...para.getElementsByTagNameNS(W,"t")];
  const boxTexts=[...box.getElementsByTagNameNS(W,"t")];
  const firstBoxText=boxTexts[0];
  const idx=firstBoxText?ts.indexOf(firstBoxText):ts.length;
  for(let i=idx-1;i>=0;i--){
    if(/[☐☒]/.test(ts[i].textContent||"")){ts[i].textContent=(ts[i].textContent||"").replace(/[☐☒]/,on?"☒":"☐");return true}
  }
  return false;
}
function setCheckByLabel(doc,paragraphNeedle,labelNeedle,on){
  const p=findParagraphCI(doc,paragraphNeedle)||findParagraphCI(doc,labelNeedle);if(!p)return false;
  const ts=[...p.getElementsByTagNameNS(W,"t")],label=String(labelNeedle||"").toLowerCase();
  let li=ts.findIndex(t=>String(t.textContent||"").toLowerCase().includes(label));if(li<0)return false;
  for(let i=li;i>=0;i--){
    if(/[☐☒]/.test(ts[i].textContent||"")){ts[i].textContent=(ts[i].textContent||"").replace(/[☐☒]/,on?"☒":"☐");return true}
  }
  for(let i=li+1;i<ts.length;i++){
    if(/[☐☒]/.test(ts[i].textContent||"")){ts[i].textContent=(ts[i].textContent||"").replace(/[☐☒]/,on?"☒":"☐");return true}
  }
  return false;
}
function setCheckAndNumber(doc,needle,on,numberValue){
  const p=findParagraphCI(doc,needle);if(!p)return false;
  const ts=[...p.getElementsByTagNameNS(W,"t")],label=String(needle||"").toLowerCase();
  let li=ts.findIndex(t=>String(t.textContent||"").toLowerCase().includes(label));
  if(li<0)return false;
  let changed=false;
  for(let i=li;i>=0;i--){
    if(/[☐☒]/.test(ts[i].textContent||"")){ts[i].textContent=(ts[i].textContent||"").replace(/[☐☒]/,on?"☒":"☐");changed=true;break}
  }
  if(!changed)for(let i=li+1;i<ts.length;i++){
    if(/[☐☒]/.test(ts[i].textContent||"")){ts[i].textContent=(ts[i].textContent||"").replace(/[☐☒]/,on?"☒":"☐");changed=true;break}
  }
  if(on&&String(numberValue||"").trim()){
    const suffix=" n° "+String(numberValue).trim();
    if(!String(p.textContent||"").includes(suffix)){
      ts[li].textContent=(ts[li].textContent||"")+suffix;
    }
  }
  return changed;
}
function appendToParagraphIfMissing(doc,needle,value){
  if(!value||String(doc.documentElement.textContent||"").includes(value))return;
  const p=findParagraphCI(doc,needle);if(!p)return;
  const r=doc.createElementNS(W,"w:r"),t=doc.createElementNS(W,"w:t");
  t.setAttributeNS("http://www.w3.org/XML/1998/namespace","xml:space","preserve");
  t.textContent=" "+value;r.appendChild(t);p.appendChild(r);
}

function setTableCell(doc,ti,ri,ci,value){const tbl=doc.getElementsByTagNameNS(W,"tbl")[ti];if(!tbl)return;const rows=[...tbl.children].filter(x=>x.localName==="tr"),r=rows[ri];if(!r)return;const cells=[...r.children].filter(x=>x.localName==="tc"),c=cells[ci];if(!c)return;let p=c.getElementsByTagNameNS(W,"p")[0]||doc.createElementNS(W,"w:p");if(!p.parentNode)c.appendChild(p);[...p.childNodes].forEach(x=>{if(x.localName!=="pPr")p.removeChild(x)});const rr=doc.createElementNS(W,"w:r"),tt=doc.createElementNS(W,"w:t");tt.textContent=value||"";rr.appendChild(tt);p.appendChild(rr)}
function partyName(p){return p.type==="morale"?(p.denomination||"").trim():((p.nom||"")+" "+(p.prenoms||"")).trim()}
function bailleurText(p){return p.type==="morale"?`${p.denomination||"___"}, personne morale, siège social : ${p.adresse||"___"}`:`${(p.nom||"").toUpperCase()} ${p.prenoms||""}, demeurant ${p.adresse||"___"}`}
function money(v){return String(v||"").trim()}

function dateFr(v){
  const s=String(v||"").trim();
  const m=s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m?m[3]+"/"+m[2]+"/"+m[1]:s;
}

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
function bodyBlocks(doc){
  const body=doc.getElementsByTagNameNS(W,"body")[0];
  return body?[...body.children]:[];
}
function blockText(n){return (n?.textContent||"").replace(/\s+/g," ").trim()}
function between(doc,startNeedle,endNeedle){
  const xs=bodyBlocks(doc);
  const s=xs.findIndex(n=>blockText(n).includes(startNeedle));
  if(s<0)return "";
  let e=endNeedle?xs.findIndex((n,i)=>i>s&&blockText(n).includes(endNeedle)):-1;
  if(e<0)e=Math.min(xs.length,s+8);
  return xs.slice(s+1,e).map(blockText).filter(Boolean).join(" ").replace(/\s+/g," ").trim();
}
function readChoice(doc,needle,count=2){
  const p=[...doc.getElementsByTagNameNS(W,"p")].find(p=>(p.textContent||"").includes(needle));
  if(!p)return -1;
  let idx=0;
  for(const t of p.getElementsByTagNameNS(W,"t")){
    const s=t.textContent||"";
    if(/[☐☒]/.test(s)){if(s.includes("☒"))return idx;idx++;if(idx>=count)break}
  }
  return -1;
}
function readCheck(doc,needle){
  const p=[...doc.getElementsByTagNameNS(W,"p")].find(p=>(p.textContent||"").includes(needle));
  return !!p&&(p.textContent||"").includes("☒");
}
function readTableCell(doc,ti,ri,ci){
  const tbl=doc.getElementsByTagNameNS(W,"tbl")[ti];if(!tbl)return "";
  const rows=[...tbl.children].filter(x=>x.localName==="tr"),r=rows[ri];if(!r)return "";
  const cells=[...r.children].filter(x=>x.localName==="tc"),cell=cells[ci];
  return cell?blockText(cell):"";
}
function parseBailleurs(doc){
  const box=getBoxText(doc,"Zone de texte 2",0);
  const cover=getBoxText(doc,"Zone de texte 219",0);
  const raw=(box||cover||"").replace(/\s+/g," ").trim();
  const typeIdx=readChoice(doc,"Personne physique",2);
  const blank=()=>({type:typeIdx===1?"morale":"physique",nom:"",prenoms:"",denomination:"",adresse:"",email:"",tel:""});
  if(!raw)return [blank()];

  const isMorale=typeIdx===1||/personne morale|soci[ée]t[ée]|sarl|sci|sas|eurl/i.test(raw);
  if(isMorale){
    let denomination=raw,adresse="";
    const m=raw.match(/^(.+?)(?:,\s*(?:personne morale,?\s*)?(?:si[eè]ge social\s*:?|dont le si[eè]ge social est situ[ée]\s*:?))\s*(.+)$/i);
    if(m){denomination=m[1].trim();adresse=m[2].trim()}
    denomination=denomination.replace(/,?\s*personne morale.*$/i,"").trim();
    return [{type:"morale",nom:"",prenoms:"",denomination,adresse,email:"",tel:""}];
  }

  let ident=raw,adresse="";
  const m=raw.match(/^(.+?)(?:,\s*demeurant\s+|\s+demeurant\s+)(.+)$/i);
  if(m){ident=m[1].trim();adresse=m[2].trim()}
  ident=ident.replace(/^(M\.?|Mme|Madame|Monsieur)\s+/i,"").trim();
  const parts=ident.split(/\s+/).filter(Boolean);
  const nom=(parts.shift()||"").trim(),prenoms=parts.join(" ");
  return [{type:"physique",nom,prenoms,denomination:"",adresse,email:"",tel:""}];
}
function boxOrBlank(doc,name,occ=0){return getBoxText(doc,name,occ)||""}
async function extractExistingLease(bytes){
  const files=await unzip(bytes),xf=files.find(f=>f.name==="word/document.xml");
  if(!xf)throw Error("Ce fichier Word ne ressemble pas à un bail compatible.");
  const doc=new DOMParser().parseFromString(td.decode(xf.data),"application/xml"),all=(doc.documentElement.textContent||"").replace(/\s+/g," ");
  const type=/LOGEMENT\s+MEUBL[ÉE]/i.test(all)?"meuble":"nu";
  const habitatChoice=readChoice(doc,"Type d’habitat",2),regimeChoice=readChoice(doc,"Régime juridique de l’immeuble",2),periodChoice=readChoice(doc,"Avant 1949",5),destChoice=readChoice(doc,"À usage exclusif d’habitation principale",2),chauffChoice=readChoice(doc,"Modalité de répartition du chauffage",2),eauChoice=readChoice(doc,"Modalité de répartition de l’eau chaude",2);
  const periods=["avant1949","1949-1974","1975-1989","1989-2005","depuis2005"];
  const paiement=[...doc.getElementsByTagNameNS(W,"p")].find(p=>(p.textContent||"").includes("Lieu de paiement"));
  const ptxt=paiement?(paiement.textContent||""):"";
  const gestion=/☒\s*Agence Action Immobili[eè]re/i.test(ptxt)?"gestion":"hors";
  const chargesMode=readCheck(doc,"Provision mensuelle")?"provision":readCheck(doc,"Forfait d’un montant")?"forfait":readCheck(doc,"Remboursement sur justificatif")?"justificatif":"provision";
  const out={
    type,gestion,bailleurs:parseBailleurs(doc),locataires:[{nom:"",prenoms:"",naissance:"",lieuNaissance:"",email:"",tel:""}],
    localisation:boxOrBlank(doc,"Zone de texte 2",1),
    habitat:habitatChoice===1?"individuel":"collectif",
    identifiantFiscal:boxOrBlank(doc,"Zone de texte 126861955",0),
    regime:regimeChoice===1?"monopropriete":"copropriete",
    periode:periodChoice>=0?periods[periodChoice]:"depuis2005",
    surface:boxOrBlank(doc,"Zone de texte 660593113",0),
    pieces:"",
    caracteristiques:boxOrBlank(doc,"Zone de texte 507952355",0),
    autresParties:boxOrBlank(doc,"Zone de texte 1328110326",0),
    equipements:boxOrBlank(doc,"Zone de texte 868199202",0),
    chauffageMode:chauffChoice===1?"collectif":"individuel",
    chauffageAutre:boxOrBlank(doc,"Zone de texte 1728029714",0),
    eauMode:eauChoice===1?"collectif":"individuel",
    eauAutre:boxOrBlank(doc,"Zone de texte 85956290",0),
    destination:destChoice===1?"mixte":"habitation",
    professionMixte:boxOrBlank(doc,"Zone de texte 1029441271",0),
    accessoiresPrivatifs:boxOrBlank(doc,"Zone de texte 1625481562",0),
    partiesCommunes:boxOrBlank(doc,"Zone de texte 6783576",0),
    technologies:boxOrBlank(doc,"Zone de texte 630456886",0),
    depensesEnergie:boxOrBlank(doc,"Zone de texte 561389577",0),
    anneeEnergie:boxOrBlank(doc,"Zone de texte 374328480",0),
    dateEffet:"",
    duree:boxOrBlank(doc,"Zone de texte 998965608",0),
    raisonDureeReduite:boxOrBlank(doc,"Zone de texte 760777164",0),
    loyer:"",
    decretRelocation:readChoice(doc,"décret fixant annuellement",2)===0?"oui":"non",
    encadrement:readChoice(doc,"loyer de référence majoré",2)===0?"oui":"non",
    loyerReference:"",loyerReferenceMajore:"",loyerBase:"",complementLoyer:"",dernierLoyer:"",dateVersementDernier:"",dateDerniereRevision:"",
    dateRevision:"",irl:"",
    chargesMode,
    chargesMontant:boxOrBlank(doc,"Zone de texte 2048181660",0)||boxOrBlank(doc,"Zone de texte 792361571",0),
    contribution:type==="nu"?boxOrBlank(doc,"Zone de texte 1214938885",0):"",
    justifContribution:type==="nu"?boxOrBlank(doc,"Zone de texte 904827124",0):"",
    assuranceColocAnnuelle:boxOrBlank(doc,"Zone de texte 1352093729",0),
    assuranceColocMensuelle:boxOrBlank(doc,"Zone de texte 1942955975",0),
    depotGarantie:"",
    honorairesVisiteBailleur:readTableCell(doc,2,1,1),
    honorairesVisiteLocataire:readTableCell(doc,2,1,2),
    honorairesEdlBailleur:readTableCell(doc,2,2,1),
    honorairesEdlLocataire:readTableCell(doc,2,2,2),
    travauxRecents:boxOrBlank(doc,"Zone de texte 2135196039",0),
    majorationTravaux:boxOrBlank(doc,"Zone de texte 1974909847",0),
    diminutionTravaux:boxOrBlank(doc,"Zone de texte 1136262404",0),
    sinistre:readChoice(doc,"a-t-il subi un sinistre",2)===0?"oui":"non",
    congeLocataire:boxOrBlank(doc,"Zone de texte 520617941",0),
    conditionsLocataire:boxOrBlank(doc,"Zone de texte 1418093311",0),
    conditionsBailleur:boxOrBlank(doc,"Zone de texte 1121375395",0),
    caution:boxOrBlank(doc,"Zone de texte 676414816",0),
    annexes:{},avenantCharges:{},avenantInfos:{}
  };
  const annexLabels=["Un extrait du règlement concernant la destination de l’immeuble","Le règlement intérieur de l’immeuble","Un document informatif sur les risques de nuisances sonores aériennes","Un diagnostic de performance énergétique","Un constat de risque d‘exposition au plomb","Une copie d’un état mentionnant l’absence ou la présence de matériaux","Un état de l’installation intérieure d’électricité et de gaz","Un état des risques naturels et technologiques","Une notice d’information relative aux droits et obligations","Un état des lieux","Une autorisation préalable de mise en location","Les références aux loyers habituellement constatés","Une grille de vétusté"];
  for(const x of annexLabels)out.annexes[x]=readCheck(doc,x);
  return out;
}

function leaseBlankParty(){return {type:"physique",nom:"",prenoms:"",denomination:"",adresse:"",email:"",tel:""}}
function pdfClean(s){return String(s||"").replace(/[\u0000-\u001f]+/g," ").replace(/[ \t]+/g," ").replace(/\n[ \t]+/g,"\n").trim()}
function pdfSection(text,start,end,max=1200){
  const m=text.match(start);if(!m)return "";
  const pos=(m.index||0)+m[0].length,tail=text.slice(pos,pos+max);
  if(!end)return pdfClean(tail);
  const e=tail.search(end);return pdfClean(e>=0?tail.slice(0,e):tail);
}
function pdfSimpleValue(text,start,end){return pdfSection(text,start,end,700).replace(/^[\s:;.-]+/,"").replace(/[☐☒□■]+/g," ").replace(/\s+/g," ").trim()}
function pdfUsefulValue(v,maxLen=180){
  v=String(v||"").replace(/\s+/g," ").trim();
  if(!v||v.length>maxLen)return "";
  const boiler=/(?:le locataire|le bailleur|le présent contrat|les parties conviennent|il est expressément|en l’absence|rappel\s*:|article\s+\d|loi du|décret|modalités? de|montant du|le cas échéant|dénommés ci-après|contrat de location|s’oblige à|s'engage à|s’engage à)/i;
  if(boiler.test(v))return "";
  const optionWords=(v.match(/\b(?:grenier|terrasse|balcon|loggia|jardin|cuisine équipée|baignoire|douche|ascenseur|garage|parking|piscine|local poubelle|espaces verts)\b/gi)||[]);
  if(optionWords.length>=4)return "";
  return v;
}
function pdfShort(text,start,end,maxLen=180){return pdfUsefulValue(pdfSimpleValue(text,start,end),maxLen)}
function pdfMoney(v){const m=String(v||"").match(/(?:^|\s)(\d[\d\s.,]*)(?:\s*€|\s*$)/);return m?m[1].replace(/\s/g,"").replace(",","."):""}
function parsePdfBailleur(text){
  let raw=pdfSimpleValue(text,/Nom et pr[ée]nom,? ou d[ée]nomination du bailleur[^:]*:/i,/Personne physique/i);
  if(!raw){raw=pdfSimpleValue(text,/LE PR[ÉE]SENT CONTRAT EST CONCLU ENTRE LES SOUSSIGN[ÉE]S\s*:/i,/Personne physique/i).replace(/Nom et pr[ée]nom,? ou d[ée]nomination du bailleur[^:]*:/i,"").trim()}
  raw=raw.replace(/Adresse [ée]lectronique.*$/i,"").trim();
  const typeBlock=pdfSection(text,/Personne physique/i,/D[ée]nomm[ée]s ci-apr[èe]s/i,700);
  const morale=/[☒■✓✔Xx]\s*Personne morale/i.test(typeBlock)||/\b(SCI|SARL|SAS|EURL|SASU|SOCI[ÉE]T[ÉE])\b/i.test(raw);
  const emailBlock=pdfSimpleValue(text,/Adresse [ée]lectronique\s*\(2\)\s*:/i,/T[ée]l[ée]phone\s*\(2\)\s*:/i);
  const email=(emailBlock.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i)||[""])[0];
  const telBlock=pdfSimpleValue(text,/T[ée]l[ée]phone\s*\(2\)\s*:/i,/D[ée]nomm[ée]s ci-apr[èe]s/i);
  const tel=(telBlock.match(/(?:\+33|0)[1-9](?:[ .-]?\d{2}){4}/)||[""])[0];
  if(morale){
    let denomination=raw,adresse="";const m=raw.match(/^(.+?)(?:\s+si[eè]ge social\s*:?|,\s*(?:dont le )?si[eè]ge social\s*:?)(.+)$/i);
    if(m){denomination=m[1].trim();adresse=m[2].trim()}
    return [{type:"morale",nom:"",prenoms:"",denomination,adresse,email,tel}];
  }
  let ident=raw,adresse="";const m=raw.match(/^(.+?)(?:,?\s+demeurant\s+|\s+-\s+)(.+)$/i);
  if(m){ident=m[1].trim();adresse=m[2].trim()}
  ident=ident.replace(/^(Monsieur|Madame|M\.?|Mme)\s+/i,"").trim();
  const parts=ident.split(/\s+/).filter(Boolean);if(!parts.length)return [leaseBlankParty()];
  return [{type:"physique",nom:parts.shift()||"",prenoms:parts.join(" "),denomination:"",adresse,email,tel}];
}

function pdfEscRe(s){return String(s||"").replace(/[.*+?^$()|[\]\\]/g,"\\$&")}
function pdfLines(page){return String(page||"").split(/\n+/).map(x=>x.trim()).filter(Boolean)}
function pdfLineAfter(page,re){
  const lines=pdfLines(page);
  for(let i=0;i<lines.length;i++){
    const m=lines[i].match(re);
    if(m){
      const same=lines[i].slice((m.index||0)+m[0].length).replace(/^[\s:;-]+/,"").trim();
      if(same)return same;
      return lines[i+1]||"";
    }
  }
  return "";
}
function pdfBetweenLines(page,startRe,endRe,maxLines=6){
  const lines=pdfLines(page);let s=-1,e=lines.length;
  for(let i=0;i<lines.length;i++){
    if(s<0&&startRe.test(lines[i]))s=i;
    else if(s>=0&&endRe&&endRe.test(lines[i])){e=i;break}
  }
  if(s<0)return "";
  const first=lines[s].replace(startRe,"").replace(/^[\s:;-]+/,"").trim();
  return (first?[first]:[]).concat(lines.slice(s+1,Math.min(e,s+1+maxLines))).join(" ").replace(/\s+/g," ").trim();
}
function parseAgencyPdfPages(pages){
  const p=n=>pages[n-1]||"";
  const type=/LOGEMENT\s+MEUBL[ÉE]/i.test(p(1)+p(2))?"meuble":"nu";
  const gestion=/Repr[ée]sent[ée]?\(s\)? par Action Immobili[èe]re/i.test(p(2))||/Lieu de paiement[\s\S]{0,250}[☒■✓✔Xx]\s*Agence Action Immobili[èe]re/i.test(p(7))?"gestion":"hors";

  let bailleurName="";
  let m=p(1).match(/Nom\(s\)\s*bailleur\(s\)\s*:\s*([^\n]+)/i);
  if(m)bailleurName=m[1].trim();
  if(!bailleurName){m=p(2).match(/Nom et pr[ée]nom,?\s*ou d[ée]nomination du bailleur[^:\n]*:\s*([^\n]+)/i);if(m)bailleurName=m[1].trim()}
  const morale=/[☒■✓✔Xx]\s*Personne morale/i.test(p(2))||/\b(SCI|SARL|SAS|EURL|SASU)\b/i.test(bailleurName);
  const bailleur=leaseBlankParty();
  if(morale){bailleur.type="morale";bailleur.denomination=bailleurName}
  else{const parts=bailleurName.replace(/^(M\.?|Mme|Monsieur|Madame)\s+/i,"").split(/\s+/).filter(Boolean);bailleur.nom=parts.shift()||"";bailleur.prenoms=parts.join(" ")}

  let localisation=pdfBetweenLines(p(3),/Localisation du logement[^:]*:/i,/Type d[’']habitat/i,4);
  localisation=pdfUsefulValue(localisation,220);

  let identifiantFiscal=pdfBetweenLines(p(4),/Identifiant fiscal du logement[^:]*:/i,/R[ée]gime juridique/i,2);
  identifiantFiscal=pdfUsefulValue(identifiantFiscal,80);

  const sm=p(4).match(/(\d+(?:[.,]\d+)?)\s*M[²2]\s+(\d+)/i);
  const surface=sm?sm[1].replace(",","."):"",pieces=sm?sm[2]:"";
  let caracteristiques=pdfBetweenLines(p(4),/Caract[ée]ristiques du logement\s*:/i,/Autres parties du logement/i,4);
  caracteristiques=pdfUsefulValue(caracteristiques.replace(/[☐☒□■].*$/,"").trim(),240);
  let equipements=pdfBetweenLines(p(4),/El[ée]ments d.?[ée]quipements du logement\s*:/i,/Modalit[ée] de r[ée]partition du chauffage/i,5);
  equipements=equipements.split(/\s+(?=[☐☒□■])/)[0].trim();
  equipements=pdfUsefulValue(equipements,180);

  const chauffageBlock=pdfBetweenLines(p(4),/Modalit[ée] de r[ée]partition du chauffage/i,/Modalit[ée] de r[ée]partition de l.?eau chaude/i,4);
  const eauBlock=pdfBetweenLines(p(4),/Modalit[ée] de r[ée]partition de l.?eau chaude sanitaire/i,/Rappel\s*:/i,4);
  const chauffageMode=/[☒■✓✔Xx]\s*Collectif/i.test(chauffageBlock)?"collectif":"individuel";
  const eauMode=/[☒■✓✔Xx]\s*Collectif/i.test(eauBlock)?"collectif":"individuel";
  const chauffageAutre=/[☒■✓✔Xx]\s*Gaz/i.test(chauffageBlock)?"Gaz":/[☒■✓✔Xx]\s*[ÉE]lectrique/i.test(chauffageBlock)?"Électrique":"";
  const eauAutre=/[☒■✓✔Xx]\s*Gaz/i.test(eauBlock)?"Gaz":/[☒■✓✔Xx]\s*[ÉE]lectrique/i.test(eauBlock)?"Électrique":"";

  let technologies=pdfBetweenLines(p(5),/Equipement d.?acc[èe]s aux technologies[\s\S]*?:/i,/F\.\s*D[ée]penses [ée]nerg[ée]tiques/i,4);
  technologies=pdfUsefulValue(technologies,120);
  const energy=p(5).match(/(\d[\d\s.,]*\s*-\s*\d[\d\s.,]*)\s*€/i);
  const years=p(5).match(/\b20\d{2}\b/g)||[];

  let duree=pdfLineAfter(p(6),/B\.\s*Dur[ée]e du contrat\s*:/i);
  if(!/^\d+\s*(?:an|ans|mois)\b/i.test(duree)){m=p(6).match(/\b(\d+\s*(?:an|ans|mois))\b/i);duree=m?m[1]:""}

  let chargesMontant="";
  m=p(8).match(/Charges r[ée]cup[ée]rables\s+(\d+(?:[.,]\d+)?)\s*€/i);
  if(m)chargesMontant=m[1].replace(",",".");
  const fees=p(10);
  const visite=fees.match(/Visite, constitution du dossier et r[ée]daction\s+(\d+(?:[.,]\d+)?)\s*€\s+(\d+(?:[.,]\d+)?)\s*€/i)||[];
  const edl=fees.match(/R[ée]alisation de l.?[ée]tat des lieux d.?entr[ée]e\s+(\d+(?:[.,]\d+)?)\s*€\s+(\d+(?:[.,]\d+)?)\s*€/i)||[];

  const out={
    type,gestion,bailleurs:[bailleur],locataires:[leaseBlankParty()],localisation,
    habitat:/[☒■✓✔Xx]\s*individuel/i.test(p(3))?"individuel":"collectif",identifiantFiscal,
    regime:/[☒■✓✔Xx]\s*monopropri[ée]t[ée]/i.test(p(4))?"monopropriete":"copropriete",
    periode:/[☒■✓✔Xx]\s*Avant 1949/i.test(p(4))?"avant1949":/[☒■✓✔Xx]\s*de 1949/i.test(p(4))?"1949-1974":/[☒■✓✔Xx]\s*de 1975/i.test(p(4))?"1975-1989":/[☒■✓✔Xx]\s*de 1989/i.test(p(4))?"1989-2005":"depuis2005",
    surface,pieces,caracteristiques,autresParties:"",equipements,chauffageMode,chauffageAutre,eauMode,eauAutre,
    destination:/[☒■✓✔Xx]\s*[ÀA] usage mixte/i.test(p(5))?"mixte":"habitation",professionMixte:"",accessoiresPrivatifs:"",partiesCommunes:"",technologies,
    depensesEnergie:energy?energy[1].replace(/\s+/g," ").trim():"",anneeEnergie:years.length?years[years.length-1]:"",
    dateEffet:"",duree,raisonDureeReduite:"",loyer:"",decretRelocation:"non",encadrement:"non",loyerReference:"",loyerReferenceMajore:"",loyerBase:"",complementLoyer:"",dernierLoyer:"",dateVersementDernier:"",dateDerniereRevision:"",dateRevision:"",irl:"",
    chargesMode:/[☒■✓✔Xx]\s*Forfait/i.test(p(7))?"forfait":/[☒■✓✔Xx]\s*Remboursement sur justificatif/i.test(p(7))?"justificatif":"provision",chargesMontant,
    contribution:"",justifContribution:"",assuranceColocAnnuelle:"",assuranceColocMensuelle:"",depotGarantie:"",
    honorairesVisiteBailleur:visite[1]?visite[1].replace(",","."):"",honorairesVisiteLocataire:visite[2]?visite[2].replace(",","."):"",
    honorairesEdlBailleur:edl[1]?edl[1].replace(",","."):"",honorairesEdlLocataire:edl[2]?edl[2].replace(",","."):"",
    travauxRecents:"",majorationTravaux:"",diminutionTravaux:"",sinistre:/a-t-il subi un sinistre[\s\S]{0,300}[☒■✓✔Xx]\s*Oui/i.test(p(10))?"oui":"non",
    congeLocataire:"",conditionsLocataire:"",conditionsBailleur:"",caution:"",annexes:{},avenantCharges:{},avenantInfos:{}
  };

  const annexMap=[
    ["Un extrait du règlement concernant la destination de l’immeuble","extrait du règlement concernant la destination"],["Le règlement intérieur de l’immeuble","règlement intérieur de l’immeuble"],
    ["Un document informatif sur les risques de nuisances sonores aériennes","risques de nuisances sonores aériennes"],["Un diagnostic de performance énergétique","diagnostic de performance énergétique"],
    ["Un constat de risque d‘exposition au plomb","constat de risque"],["Une copie d’un état mentionnant l’absence ou la présence de matériaux","absence ou la présence de matériaux"],
    ["Un état de l’installation intérieure d’électricité et de gaz","installation intérieure d’électricité et de gaz"],["Un état des risques naturels et technologiques","risques naturels et technologiques"],
    ["Une notice d’information relative aux droits et obligations","notice d’information relative aux droits"],["Un état des lieux","état des lieux"],
    ["Une autorisation préalable de mise en location","autorisation préalable de mise en location"],["Les références aux loyers habituellement constatés","références aux loyers habituellement constatés"],["Une grille de vétusté","grille de vétusté"]
  ];
  for(const [label,snippet] of annexMap)out.annexes[label]=new RegExp("[☒■✓✔Xx]\\s*[^\\n]{0,20}"+pdfEscRe(snippet),"i").test(p(12));

  const av=p(15);
  const ac=[["L’eau (elle sera réajustée en plus ou en moins selon les consommations réelles)","L’eau"],["L’électricité","L’électricité"],["Le gaz","Le gaz"],["Internet","Internet"],["La minuterie","La minuterie"],["Le ménage des parties communes","Le ménage des parties communes"],["L’ascenseur","L’ascenseur"],["Le Contrat d’entretien de la chaudière","Contrat d’entretien de la chaudière"]];
  for(const [label,snippet] of ac)out.avenantCharges[label]=new RegExp("[☒■✓✔Xx]\\s*"+pdfEscRe(snippet),"i").test(av);
  const ai=[["La taxe d’ordure ménagère sera à payer par le locataire.","taxe d’ordure ménagère"],["Le compteur électrique et/ou gaz devra être ouvert au nom du locataire.","compteur électrique et/ou gaz"],["Le compteur d’eau devra être ouvert au nom du locataire.","compteur d’eau"],["La taxe d'habitation sera due par le locataire au 1er janvier.","taxe d'habitation"],["Le locataire devra contracter une police d'assurance incendie et dégâts des eaux avant la remise des clés et s'engage à fournir au bailleur un justificatif annuel.","police d'assurance incendie"],["Le locataire s’engage à prendre un contrat d’entretien pour la chaudière et à fournir au bailleur un justificatif annuel.","contrat d’entretien pour la chaudière"],["Le locataire s’engage à faire un ramonage annuel de la cheminée ou du poêle à bois et à fournir au bailleur un justificatif annuel.","ramonage annuel"],["Le locataire s’engage à entretenir le jardin et les abords de la maison (pelouses, haies, plantations, terrasse)","entretenir le jardin"],["Le locataire s’engage à fournir une pile pour le détecteur de fumée, lors de l’état des lieux de sortie si celle-ci ne fonctionne plus.","pile pour le détecteur de fumée"],["Le preneur ou locataire s’engage à prendre un contrat d’entretien pour la VMC chaque année.","contrat d’entretien pour la VMC"],["Le preneur ou locataire s’engage à prendre un contrat d’entretien pour la pompe à chaleur chaque année","contrat d’entretien pour la pompe à chaleur"]];
  for(const [label,snippet] of ai)out.avenantInfos[label]=new RegExp("[☒■✓✔Xx]\\s*"+pdfEscRe(snippet),"i").test(av);
  return out;
}
function parsePdfLeaseText(raw){
  const text=pdfClean(raw),o={type:/LOGEMENT\s+MEUBL[ÉE]/i.test(text)?"meuble":"nu",gestion:/Action Immobili[èe]re/i.test(text)?"gestion":"hors"};
  return Object.assign({bailleurs:[leaseBlankParty()],locataires:[leaseBlankParty()],localisation:"",habitat:"collectif",identifiantFiscal:"",regime:"copropriete",periode:"depuis2005",surface:"",pieces:"",caracteristiques:"",autresParties:"",equipements:"",chauffageMode:"individuel",chauffageAutre:"",eauMode:"individuel",eauAutre:"",destination:"habitation",professionMixte:"",accessoiresPrivatifs:"",partiesCommunes:"",technologies:"",depensesEnergie:"",anneeEnergie:"",dateEffet:"",duree:"",raisonDureeReduite:"",loyer:"",decretRelocation:"non",encadrement:"non",loyerReference:"",loyerReferenceMajore:"",loyerBase:"",complementLoyer:"",dernierLoyer:"",dateVersementDernier:"",dateDerniereRevision:"",dateRevision:"",irl:"",chargesMode:"provision",chargesMontant:"",contribution:"",justifContribution:"",assuranceColocAnnuelle:"",assuranceColocMensuelle:"",depotGarantie:"",honorairesVisiteBailleur:"",honorairesVisiteLocataire:"",honorairesEdlBailleur:"",honorairesEdlLocataire:"",travauxRecents:"",majorationTravaux:"",diminutionTravaux:"",sinistre:"non",congeLocataire:"",conditionsLocataire:"",conditionsBailleur:"",caution:"",annexes:{},avenantCharges:{},avenantInfos:{}},o);
}
async function extractPdfPageText(page){
  const tc=await page.getTextContent();
  const items=tc.items.filter(x=>x.str&&x.str.trim()).map(x=>({s:x.str.trim(),x:x.transform[4],y:x.transform[5]}));
  items.sort((a,b)=>Math.abs(b.y-a.y)>2?b.y-a.y:a.x-b.x);
  const lines=[];let cur=[],lastY=null;
  for(const it of items){
    if(lastY===null||Math.abs(it.y-lastY)<=2){cur.push(it);lastY=lastY===null?it.y:(lastY+it.y)/2}
    else{lines.push(cur.sort((a,b)=>a.x-b.x).map(z=>z.s).join(" "));cur=[it];lastY=it.y}
  }
  if(cur.length)lines.push(cur.sort((a,b)=>a.x-b.x).map(z=>z.s).join(" "));
  return lines.join("\n");
}
async function ocrPdfPages(pdf,maxPages){
  if(!window.Tesseract)throw Error("Le PDF semble scanné et le module OCR n’a pas pu être chargé.");
  const parts=[];
  for(let n=1;n<=maxPages;n++){
    const page=await pdf.getPage(n),viewport=page.getViewport({scale:1.45}),canvas=document.createElement("canvas"),ctx=canvas.getContext("2d",{willReadFrequently:true});
    canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);
    await page.render({canvasContext:ctx,viewport}).promise;
    const status=document.getElementById("status");if(status)status.textContent="Lecture du PDF scanné — page "+n+"/"+maxPages+"…";
    const res=await Tesseract.recognize(canvas,"fra",{logger:()=>{}});parts.push(res.data.text||"");canvas.width=1;canvas.height=1;
  }
  return parts.join("\n");
}
async function extractExistingLeasePdf(bytes){
  if(!window.pdfjsLib)throw Error("Le lecteur PDF n’a pas pu être chargé. Recharge la page puis réessaie.");
  const pdf=await pdfjsLib.getDocument({data:bytes}).promise,pageCount=Math.min(pdf.numPages,15),parts=[];
  for(let n=1;n<=pageCount;n++){const page=await pdf.getPage(n);parts.push(await extractPdfPageText(page))}
  const joined=parts.join("\n");
  if(joined.replace(/\s/g,"").length<500)return parsePdfLeaseText(await ocrPdfPages(pdf,Math.min(pdf.numPages,15)));
  return parseAgencyPdfPages(parts);
}

async function buildLocation(templateBytes,data){const files=await unzip(templateBytes),xf=files.find(f=>f.name==="word/document.xml");if(!xf)throw Error("Modèle Word incomplet");const doc=new DOMParser().parseFromString(td.decode(xf.data),"application/xml");
const bailNames=data.bailleurs.map(partyName).filter(Boolean).join(" / "),locNames=data.locataires.map(partyName).filter(Boolean).join(" / ");
setBox(doc,"Zone de texte 219",0,bailNames);setBox(doc,"Zone de texte 218",0,locNames);
setBox(doc,"Zone de texte 2",0,data.bailleurs.map(bailleurText).join("\n"),true);setBailleurTypes(doc,data.bailleurs);
if(data.gestion==="hors"){setBox(doc,"Zone de texte 478161165",0,data.bailleurs.map(x=>x.email).filter(Boolean).join(" / "));setBox(doc,"Zone de texte 1135107845",0,data.bailleurs.map(x=>x.tel).filter(Boolean).join(" / "))}
data.locataires.slice(0,4).forEach((p,i)=>{setTableCell(doc,0,i+1,0,partyName(p));setTableCell(doc,0,i+1,1,[dateFr(p.naissance),p.lieuNaissance].filter(Boolean).join(" à "));setTableCell(doc,0,i+1,2,p.email);setTableCell(doc,0,i+1,3,p.tel)});
setBox(doc,"Zone de texte 2",1,data.localisation);setChoice(doc,"Type d’habitat",data.habitat==="individuel"?1:0);setBox(doc,"Zone de texte 126861955",0,data.identifiantFiscal);setChoice(doc,"Régime juridique de l’immeuble",data.regime==="monopropriete"?1:0);
const periods=["avant1949","1949-1974","1975-1989","1989-2005","depuis2005"];setChoice(doc,"Avant 1949",Math.max(0,periods.indexOf(data.periode)),5);
setBox(doc,"Zone de texte 1115947929",0,data.surface?String(data.surface)+" m²":"");setBox(doc,"Zone de texte 660593113",0,data.pieces);setBox(doc,"Zone de texte 507952355",0,data.caracteristiques);setBox(doc,"Zone de texte 1328110326",0,data.autresPartiesAutre||data.autresParties);setBox(doc,"Zone de texte 868199202",0,data.equipementsAutre||data.equipements);setBox(doc,"Zone de texte 1728029714",0,"");
const otherLabels={"Grenier":"grenier","Comble aménagé":"comble aménagé","Comble non aménagé":"comble non aménagé","Terrasse":"terrasse","Balcon":"balcon","Loggia":"loggia","Jardin":"jardin"};
for(const [label,needle] of Object.entries(otherLabels))setCheckByLabel(doc,"Autres parties du logement",needle,!!(data.autresPartiesChoix||{})[label]);
const eqLabels={"Cuisine équipée":"Cuisine équipée","Salle de bain":"Salle de bain","Salle de douche":"Salle de douche","WC séparé":"WC séparé"};
for(const [label,needle] of Object.entries(eqLabels))setCheckByLabel(doc,"Eléments d’équipements du logement",needle,!!(data.equipementsChoix||{})[label]);setBox(doc,"Zone de texte 85956290",0,"");
setEnergyExact(doc,"Modalité de répartition du chauffage",data.chauffageMode,data.chauffageEnergie,data.chauffageAutre,"Zone de texte 1728029714");setEnergyExact(doc,"Modalité de répartition de l’eau chaude sanitaire",data.eauMode,data.eauEnergie,data.eauAutre,"Zone de texte 85956290");setChoice(doc,"À usage exclusif d’habitation principale",data.destination==="mixte"?1:0);setBox(doc,"Zone de texte 1029441271",0,data.professionMixte);setBox(doc,"Zone de texte 1625481562",0,data.accessoiresPrivatifsAutre||data.accessoiresPrivatifs);setBox(doc,"Zone de texte 6783576",0,data.partiesCommunesAutre||data.partiesCommunes);
for(const label of ["Cave","Parking","Garage"]){const on=!!(data.accessoiresPrivatifsChoix||{})[label];setLabeledCheckboxExact(doc,label,on);setCheckAndNumber(doc,label,on,(data.accessoiresPrivatifsNumeros||{})[label]);}
for(const label of ["Garage à vélo","Ascenseur","Espaces verts","Aires et équipements de jeux","Laverie","Local poubelle","Gardiennage"])setLabeledCheckboxExact(doc,label,!!(data.partiesCommunesChoix||{})[label]);setBox(doc,"Zone de texte 630456886",0,data.technologies);
setBox(doc,"Zone de texte 2077289405",0,dateFr(data.dateEffet));setBox(doc,"Zone de texte 998965608",0,data.duree);setBox(doc,"Zone de texte 760777164",0,data.raisonDureeReduite);
setBox(doc,"Zone de texte 628403819",0,moneyWords(data.loyer));setChoice(doc,"décret fixant annuellement",data.decretRelocation==="oui"?0:1);setChoice(doc,"loyer de référence majoré",data.encadrement==="oui"?0:1);setBox(doc,"Zone de texte 268514384",0,money(data.loyerBase));setBox(doc,"Zone de texte 475039207",0,money(data.complementLoyer));

// Correspondance exacte du modèle "CONTRAT DE LOCATION NU GESTION".
setBox(doc,"Zone de texte 31722978",0,moneyWords(data.dernierLoyer));
setBox(doc,"Zone de texte 1040080526",0,dateFr(data.dateVersementDernier));
setBox(doc,"Zone de texte 1259670285",0,dateFr(data.dateDerniereRevision));
setBox(doc,"Zone de texte 1542835909",0,dateFr(data.dateRevision));

const q=Number(data.irlTrimestre)||0,qLabel=q?(q===1?"1er trimestre":q+"ème trimestre"):"";
const irlText=qLabel&&data.irlAnnee?(qLabel+" "+data.irlAnnee+(data.irlValeur?" d'une valeur de "+data.irlValeur:"")):(data.irl||"");
setBox(doc,"Zone de texte 1359800534",0,irlText);

setLabeledCheckboxExact(doc,"Provision mensuelle",data.chargesMode==="provision");
setLabeledCheckboxExact(doc,"Forfait d’un montant",data.chargesMode==="forfait");
setLabeledCheckboxExact(doc,"Remboursement sur justificatif",data.chargesMode==="justificatif");
setBox(doc,"Zone de texte 85742337",0,data.chargesMode==="provision"?moneyWords(data.chargesMontant):"");
setBox(doc,"Zone de texte 2048181660",0,data.chargesMode==="forfait"?moneyWords(data.chargesMontant):"");
setBox(doc,"Zone de texte 792361571",0,data.chargesMode==="justificatif"?moneyWords(data.chargesMontant):"");
setBox(doc,"Zone de texte 1214938885",0,data.contribution);setBox(doc,"Zone de texte 904827124",0,data.justifContribution);setBox(doc,"Zone de texte 1352093729",0,money(data.assuranceColocAnnuelle));setBox(doc,"Zone de texte 1942955975",0,money(data.assuranceColocMensuelle));
setTableCell(doc,1,0,1,num(data.loyer)?num(data.loyer).toLocaleString("fr-FR")+" €":"");setTableCell(doc,1,1,1,num(data.chargesMontant)?num(data.chargesMontant).toLocaleString("fr-FR")+" €":"");let row=2;if(data.type==="nu"){setTableCell(doc,1,row++,1,num(data.contribution)?num(data.contribution).toLocaleString("fr-FR")+" €":"");}setTableCell(doc,1,row++,1,num(data.assuranceColocMensuelle)?num(data.assuranceColocMensuelle).toLocaleString("fr-FR")+" €":"");const total=num(data.loyer)+num(data.chargesMontant)+num(data.assuranceColocMensuelle)+(data.type==="nu"?num(data.contribution):0);setTableCell(doc,1,row,1,total?total.toLocaleString("fr-FR")+" €":"");
setBox(doc,"Zone de texte 561389577",0,data.depensesEnergie);setBox(doc,"Zone de texte 374328480",0,data.anneeEnergie);setBox(doc,"Zone de texte 2135196039",0,data.travauxRecents);setBox(doc,"Zone de texte 1974909847",0,data.majorationTravaux);setBox(doc,"Zone de texte 1136262404",0,data.diminutionTravaux);setBox(doc,"Zone de texte 2071132409",0,moneyWords(data.depotGarantie));setBox(doc,"Zone de texte 520617941",0,dateFr(data.congeLocataire));setBox(doc,"Zone de texte 1418093311",0,data.conditionsLocataire);setBox(doc,"Zone de texte 1121375395",0,data.conditionsBailleur);setBox(doc,"Zone de texte 676414816",0,data.caution);
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

function cautionDateFr(v){
  const m=String(v||"").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m?m[3]+"/"+m[2]+"/"+m[1]:String(v||"");
}
function cautionMoneyNumber(v){
  const n=num(v);
  return n||String(v||"").trim()?n.toLocaleString("fr-FR",{minimumFractionDigits:Number.isInteger(n)?0:2,maximumFractionDigits:2}):"";
}
function cautionMoneyLetters(v){
  const n=num(v);
  if(!n&&String(v||"").trim()==="")return "";
  const euros=Math.floor(n),cents=Math.round((n-euros)*100);
  let s=numberToFrench(euros)+" euro"+(euros>1?"s":"");
  if(cents)s+=" et "+numberToFrench(cents)+" centime"+(cents>1?"s":"");
  return s;
}
async function buildCaution(templateBytes,data){
  const files=await unzip(templateBytes),xf=files.find(f=>f.name==="word/document.xml");
  if(!xf)throw Error("Modèle d'acte de caution incomplet");
  const doc=new DOMParser().parseFromString(td.decode(xf.data),"application/xml");
  const s={...(data.cautionActe||{})};
  const bailNames=(data.bailleurs||[]).map(partyName).filter(Boolean).join(" / ");
  const bailAddresses=(data.bailleurs||[]).map(p=>p.adresse||"").filter(Boolean).join(" / ");
  const locNames=(data.locataires||[]).map(partyName).filter(Boolean).join(" / ");
  const totalRevenus=num(s.remuneration)+num(s.autresRevenus);
  const civilite=s.civilite||"";
  const identite=[civilite,s.prenoms,s.nom].filter(Boolean).join(" ");
  const birth=[cautionDateFr(s.naissance),s.lieuNaissance].filter(Boolean).join(" à ");
  const loyerN=cautionMoneyNumber(data.loyer),chargesN=cautionMoneyNumber(data.chargesMontant),depotN=cautionMoneyNumber(data.depotGarantie);

  setBox(doc,"Zone de texte 2",0,s.nom);
  setBox(doc,"Zone de texte 2",1,s.prenoms);
  setBox(doc,"Zone de texte 2",2,birth);
  setBox(doc,"Zone de texte 2",3,s.domicile);
  setBox(doc,"Zone de texte 2",4,s.email);
  setBox(doc,"Zone de texte 2",5,s.tel);
  setBox(doc,"Zone de texte 2",6,s.profession);
  setBox(doc,"Zone de texte 2",7,s.employeur);
  setBox(doc,"Zone de texte 2",8,s.contratType==="determinee"?cautionDateFr(s.contratFin):"");
  setBox(doc,"Zone de texte 2",9,cautionMoneyNumber(s.remuneration));
  setBox(doc,"Zone de texte 2",10,cautionMoneyNumber(s.autresRevenus));
  setBox(doc,"Zone de texte 2",11,totalRevenus?cautionMoneyNumber(totalRevenus):"");

  const family=["Célibataire","Marié(e)","Pacsé(e)","Divorcé(e)","Veuf(ve)"];
  setChoice(doc,"Célibataire",Math.max(0,family.indexOf(s.situation)),5);
  setChoice(doc,"Contrat de travail à durée",s.contratType==="determinee"?1:0,2);

  setBox(doc,"Zone de texte 2",12,bailNames);
  setBox(doc,"Zone de texte 2",13,bailAddresses);
  setBox(doc,"Zone de texte 2",14,locNames);
  setBox(doc,"Zone de texte 2",15,data.localisation);
  setBox(doc,"Zone de texte 2",16,cautionDateFr(s.dateBail));
  setBox(doc,"Zone de texte 2",17,cautionDateFr(data.dateEffet));
  setBox(doc,"Zone de texte 2",18,data.duree);
  setBox(doc,"Zone de texte 2",19,data.irlValeur||s.irlValeur);
  setBox(doc,"Zone de texte 2",20,data.irlAnnee||s.irlAnnee);
  setBox(doc,"Zone de texte 2",21,identite);
  setBox(doc,"Zone de texte 2",22,locNames);

  setTableCell(doc,0,1,1,loyerN?loyerN+" €":"");
  setTableCell(doc,0,1,2,cautionMoneyLetters(data.loyer));
  setTableCell(doc,0,2,1,chargesN?chargesN+" €":"");
  setTableCell(doc,0,2,2,cautionMoneyLetters(data.chargesMontant));
  setTableCell(doc,0,3,1,depotN?depotN+" €":"");
  setTableCell(doc,0,3,2,cautionMoneyLetters(data.depotGarantie));

  const q=Math.max(1,Math.min(4,Number(data.irlTrimestre||s.irlTrimestre)||1));
  setChoice(doc,"trimestre de l’année",q-1,4);

  setBox(doc,"Zone de texte 2",23,cautionMoneyLetters(data.loyer));
  setBox(doc,"Zone de texte 2",24,loyerN?loyerN+" €":"");
  setBox(doc,"Zone de texte 2",25,data.irlValeur||s.irlValeur);
  setBox(doc,"Zone de texte 2",26,data.irlAnnee||s.irlAnnee);
  setBox(doc,"Zone de texte 2",27,String(data.irlTrimestre||s.irlTrimestre||""));
  setBox(doc,"Zone de texte 2",28,cautionMoneyLetters(data.chargesMontant));
  setBox(doc,"Zone de texte 2",29,chargesN?chargesN+" €":"");
  setBox(doc,"Zone de texte 2",30,cautionMoneyLetters(data.depotGarantie));
  setBox(doc,"Zone de texte 2",31,depotN?depotN+" €":"");
  setBox(doc,"Zone de texte 2",32,data.duree);
  setBox(doc,"Zone de texte 2",33,cautionDateFr(data.dateEffet));
  setBox(doc,"Zone de texte 2",34,locNames);
  setBox(doc,"Zone de texte 2",35,s.reconductions);
  setBox(doc,"Zone de texte 2",36,s.dureeMax);

  const monthlyTotal=num(data.loyer)+num(data.chargesMontant),months=data.type==="meuble"?12:36,maxAmount=monthlyTotal*months,maxN=cautionMoneyNumber(maxAmount);
  const maxText=maxAmount?[cautionMoneyLetters(maxAmount),maxN?maxN+" €":""].filter(Boolean).join(" — "):"";
  setBox(doc,"Zone de texte 2",40,maxText);
  setBox(doc,"Zone de texte 2",41,data.type==="meuble"?"1":"3");

  xf.data=te.encode(new XMLSerializer().serializeToString(doc));
  return new Blob([zip(files)],{type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
}
