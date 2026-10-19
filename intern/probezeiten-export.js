/* Probezeiten-Export (Excel + PDF) für HSO Intern
 * Eingebunden in intern/index.html per <script src="probezeiten-export.js"></script>
 */
(function(){
'use strict';

var XLSX_URL='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

function slug(s){return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Za-z0-9]+/g,'_').replace(/^_+|_+$/g,'')||'x';}
function heute(){var d=new Date();return d.getFullYear()+('0'+(d.getMonth()+1)).slice(-2)+('0'+d.getDate()).slice(-2);}
function heuteDe(){var d=new Date();return ('0'+d.getDate()).slice(-2)+'.'+('0'+(d.getMonth()+1)).slice(-2)+'.'+d.getFullYear();}
function h(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function istHaupt(d){return /dorsch|song/i.test(d);}
function wtag(d){return d?WT[d.getDay()]:'';}

/* ── Daten ── */
function zeitItems(){
  var jetzt=new Date();
  return (_untTermineAlle||[]).map(function(t){
    var slots=t.slots||[];
    var letzter=slots.length?slots[slots.length-1].uhrzeit:t.uhrzeit;
    var ende=untSlotEnde(t.datum,letzter);
    var tn=[].concat.apply([],(t.teilnehmer||[]).map(function(x){return x;})).filter(function(x){return x&&!/^\d+$/.test(String(x));});
    var buch=(_untBuchDaten||[]).filter(function(b){return b.terminId===t.id;});
    var namen=buch.map(function(b){return b.name;});
    tn.forEach(function(x){if(namen.indexOf(x)<0)namen.push(x);});
    var von=slots.length?slots[0].uhrzeit:t.uhrzeit;
    var bis=untUhrPlus(letzter,UNT_SLOT_MIN);
    var n=Math.max(1,slots.length);
    var typ=t.marker||'';
    var d=parseDatum(t.datum);
    return {
      doz:(t.dozent||'').trim()||'(ohne Dozent)', typ:typ, d:d, datum:t.datum, wt:wtag(d),
      von:von, bis:bis||'', ort:t.ort||'', projekt:(t.projekt&&t.projekt!==(typ+' '+t.dozent))?t.projekt:'',
      min:n*UNT_SLOT_MIN, n:n, belegt:buch.length, namen:namen,
      vorbei:!!(ende&&ende<=jetzt)
    };
  });
}
function zeitSichtbar(nurDoz){
  var f=_untZF;
  return zeitItems().filter(function(i){
    if(f.typ!=='alle'&&i.typ!==f.typ)return false;
    if(f.doz==='__gast'){if(istHaupt(i.doz))return false;}
    else if(f.doz!=='alle'&&i.doz!==f.doz)return false;
    if(f.zeit==='kommend'&&i.vorbei)return false;
    if(f.zeit==='vorbei'&&!i.vorbei)return false;
    if(nurDoz&&i.doz!==nurDoz)return false;
    return true;
  }).sort(function(a,b){return (a.d?a.d.getTime():0)-(b.d?b.d.getTime():0)||String(a.von).localeCompare(String(b.von));});
}
function zeitGruppen(nurDoz){
  var gr={};
  zeitSichtbar(nurDoz).forEach(function(i){(gr[i.doz]=gr[i.doz]||[]).push(i);});
  return Object.keys(gr).sort(function(a,b){return (istHaupt(b)-istHaupt(a))||a.localeCompare(b);})
    .map(function(k){return {name:k,haupt:istHaupt(k),items:gr[k]};});
}
function studGruppen(nurName){
  var jetzt=new Date(),instr={};
  (_untStudierende||[]).forEach(function(s){instr[s.name]=s.instrument||'';});
  var pers={};
  (_untBuchDaten||[]).forEach(function(b){
    var key=b.name||b.mtknr||'?';
    if(nurName&&key!==nurName)return;
    var ende=untSlotEnde(b.datum,b.slotStart),d=parseDatum(b.datum);
    var p=pers[key]||(pers[key]={name:key,instrument:instr[key]||'',items:[]});
    p.items.push({
      d:d,datum:b.datum,wt:wtag(d),von:b.slotStart,bis:untUhrPlus(b.slotStart,UNT_SLOT_MIN)||'',
      typ:b.marker||'',doz:b.dozent||'',ort:b.ort||'',werk:b.werk||'',min:UNT_SLOT_MIN,
      vorbei:!!(ende&&ende<=jetzt),ende:ende
    });
  });
  return Object.keys(pers).map(function(k){
    var p=pers[k];
    p.items.sort(function(a,b){return (a.ende?a.ende.getTime():0)-(b.ende?b.ende.getTime():0);});
    return p;
  }).sort(function(a,b){return a.name.localeCompare(b.name);});
}

/* ── Zeilen für Excel ── */
var ZK=['Dozent:in','Art','Datum','Wochentag','Von','Bis','Dauer (min)','Ort','Projekt','Belegung','Studierende','Status'];
function zRow(i){
  return [i.doz,i.typ,i.d||i.datum,i.wt,i.von,i.bis,i.min,i.ort,i.projekt,
    i.typ==='Einzelunterricht'?(i.belegt+'/'+i.n):'',i.namen.join(', '),i.vorbei?'stattgefunden':'kommend'];
}
var SK=['Studierende:r','Instrument','Datum','Wochentag','Von','Bis','Dauer (min)','Art','Dozent:in','Ort','Werk','Status'];
function sRow(p,i){
  return [p.name,p.instrument,i.d||i.datum,i.wt,i.von,i.bis,i.min,i.typ,i.doz,i.ort,i.werk,i.vorbei?'stattgefunden':'kommend'];
}

/* ── Excel ── */
var _xlsxP=null;
function ladeXlsx(){
  if(window.XLSX)return Promise.resolve();
  if(_xlsxP)return _xlsxP;
  _xlsxP=new Promise(function(res,rej){
    var s=document.createElement('script');s.src=XLSX_URL;
    s.onload=res;s.onerror=function(){_xlsxP=null;rej(new Error('XLSX-Bibliothek nicht ladbar'));};
    document.head.appendChild(s);
  });
  return _xlsxP;
}
function blatt(kopf,zeilen,datumCol,breiten){
  var ws=XLSX.utils.aoa_to_sheet([kopf].concat(zeilen),{cellDates:true});
  for(var r=1;r<=zeilen.length;r++){
    var c=ws[XLSX.utils.encode_cell({r:r,c:datumCol})];
    if(c&&c.t==='d')c.z='dd.mm.yyyy';
  }
  ws['!cols']=breiten.map(function(w){return {wch:w};});
  ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:zeilen.length,c:kopf.length-1}})};
  ws['!freeze']={xSplit:0,ySplit:1};
  return ws;
}
function blattName(n,benutzt){
  var s=String(n).replace(/[\[\]:*?\/\\]/g,' ').trim().slice(0,31)||'Blatt';
  var b=s,k=2;
  while(benutzt.indexOf(b.toLowerCase())>=0){b=s.slice(0,28)+' '+k++;}
  benutzt.push(b.toLowerCase());return b;
}
var ZB=[22,18,12,11,7,7,11,16,22,10,40,14],SB=[26,16,12,11,7,7,11,18,22,16,30,14];

function exportZeitenExcel(nurDoz){
  var gr=zeitGruppen(nurDoz);
  if(!gr.length){zeigToast('Nichts zu exportieren');return;}
  return ladeXlsx().then(function(){
    var wb=XLSX.utils.book_new(),benutzt=[];
    if(gr.length>1){
      var alle=[];gr.forEach(function(g){g.items.forEach(function(i){alle.push(zRow(i));});});
      XLSX.utils.book_append_sheet(wb,blatt(ZK,alle,2,ZB),blattName('Alle',benutzt));
    }
    gr.forEach(function(g){
      XLSX.utils.book_append_sheet(wb,blatt(ZK,g.items.map(zRow),2,ZB),blattName(g.name,benutzt));
    });
    var fn='Probezeiten_'+(gr.length===1?slug(gr[0].name):'Dozenten')+'_'+heute()+'.xlsx';
    XLSX.writeFile(wb,fn,{cellDates:true});
    zeigToast('Excel erstellt: '+fn);
  }).catch(function(e){zeigToast('Export fehlgeschlagen: '+e.message);});
}
function exportStudExcel(nurName){
  var gr=studGruppen(nurName);
  if(!gr.length){zeigToast('Nichts zu exportieren');return;}
  return ladeXlsx().then(function(){
    var wb=XLSX.utils.book_new(),benutzt=[];
    if(gr.length>1){
      var alle=[];gr.forEach(function(p){p.items.forEach(function(i){alle.push(sRow(p,i));});});
      XLSX.utils.book_append_sheet(wb,blatt(SK,alle,2,SB),blattName('Alle',benutzt));
    }
    gr.forEach(function(p){
      XLSX.utils.book_append_sheet(wb,blatt(SK,p.items.map(function(i){return sRow(p,i);}),2,SB),blattName(p.name,benutzt));
    });
    var fn='Probezeiten_'+(gr.length===1?slug(gr[0].name):'Studierende')+'_'+heute()+'.xlsx';
    XLSX.writeFile(wb,fn,{cellDates:true});
    zeigToast('Excel erstellt: '+fn);
  }).catch(function(e){zeigToast('Export fehlgeschlagen: '+e.message);});
}

/* ── PDF (Druckdialog → „Als PDF speichern“) ── */
var CSS='@page{size:A4;margin:16mm 14mm}*{box-sizing:border-box}body{font:11px/1.4 Arial,Helvetica,sans-serif;color:#111;margin:0}'+
'h1{font-size:15px;margin:0 0 2px}.sub{color:#555;font-size:10px;margin-bottom:10px}'+
'.grp{page-break-before:always}.grp:first-of-type{page-break-before:auto}'+
'table{border-collapse:collapse;width:100%;margin-top:6px}th,td{border-bottom:1px solid #ddd;padding:4px 5px;text-align:left;vertical-align:top}'+
'th{font-size:9px;text-transform:uppercase;color:#444;border-bottom:1.5px solid #333}tr{page-break-inside:avoid}'+
'.r{text-align:right;white-space:nowrap}.nw{white-space:nowrap}.ok{color:#1e6b3a}.off{color:#8a6e1f}.sum{margin:6px 0;font-weight:bold}';
function fmtMin(m){return untFmt(m);}
function drucke(titel,body){
  var f=document.createElement('iframe');
  f.style.cssText='position:fixed;right:0;bottom:0;width:0;height:0;border:0';
  document.body.appendChild(f);
  var doc=f.contentWindow.document;
  doc.open();
  doc.write('<!DOCTYPE html><html lang="de"><head><meta charset="utf-8"><title>'+h(titel)+'</title><style>'+CSS+'</style></head><body>'+body+'</body></html>');
  doc.close();
  setTimeout(function(){
    try{f.contentWindow.focus();f.contentWindow.print();}catch(e){zeigToast('Drucken nicht möglich');}
    setTimeout(function(){if(f.parentNode)f.parentNode.removeChild(f);},60000);
  },400);
}
function exportZeitenPdf(nurDoz){
  var gr=zeitGruppen(nurDoz);
  if(!gr.length){zeigToast('Nichts zu exportieren');return;}
  var body=gr.map(function(g){
    var min=g.items.reduce(function(a,i){return a+i.min;},0);
    var ein=g.items.filter(function(i){return i.typ==='Einzelunterricht';}).length;
    return '<div class="grp"><h1>'+h(g.name)+'</h1><div class="sub">'+(g.haupt?'Hauptdozent:in':'Gastdozent:in')+' · Probezeiten · Stand '+heuteDe()+'</div>'+
      '<div class="sum">'+g.items.length+' Termin'+(g.items.length!==1?'e':'')+' ('+ein+' Einzel, '+(g.items.length-ein)+' Gruppe) · '+fmtMin(min)+'</div>'+
      '<table><thead><tr><th>Datum</th><th>Zeit</th><th>Art</th><th>Ort / Projekt</th><th>Studierende</th><th class="r">Dauer</th><th class="r">Belegt</th><th></th></tr></thead><tbody>'+
      g.items.map(function(i){
        return '<tr><td class="nw">'+h(i.wt)+', '+h(i.datum)+'</td><td class="nw">'+h(i.von)+(i.bis?'–'+h(i.bis):'')+'</td><td>'+h(i.typ)+'</td><td>'+h(i.ort)+(i.projekt?(i.ort?' · ':'')+h(i.projekt):'')+'</td>'+
          '<td>'+h(i.namen.join(', '))+'</td><td class="r">'+fmtMin(i.min)+'</td><td class="r">'+(i.typ==='Einzelunterricht'?i.belegt+'/'+i.n:'')+'</td>'+
          '<td class="'+(i.vorbei?'ok':'off')+'">'+(i.vorbei?'✓':'○')+'</td></tr>';
      }).join('')+'</tbody></table></div>';
  }).join('');
  drucke('Probezeiten '+(gr.length===1?gr[0].name:'Dozenten')+' '+heuteDe(),body);
}
function exportStudPdf(nurName){
  var gr=studGruppen(nurName);
  if(!gr.length){zeigToast('Nichts zu exportieren');return;}
  var body=gr.map(function(p){
    var min=p.items.reduce(function(a,i){return a+i.min;},0);
    var vor=p.items.filter(function(i){return i.vorbei;}).reduce(function(a,i){return a+i.min;},0);
    return '<div class="grp"><h1>'+h(p.name)+'</h1><div class="sub">'+h(p.instrument)+(p.instrument?' · ':'')+'Probezeiten · Stand '+heuteDe()+'</div>'+
      '<div class="sum">'+p.items.length+' Stunde'+(p.items.length!==1?'n':'')+' · '+fmtMin(min)+' gebucht · davon stattgefunden: '+fmtMin(vor)+'</div>'+
      '<table><thead><tr><th>Datum</th><th>Zeit</th><th>Art</th><th>Dozent:in</th><th>Ort</th><th>Werk</th><th class="r">Dauer</th><th></th></tr></thead><tbody>'+
      p.items.map(function(i){
        return '<tr><td class="nw">'+h(i.wt)+', '+h(i.datum)+'</td><td class="nw">'+h(i.von)+(i.bis?'–'+h(i.bis):'')+'</td><td>'+h(i.typ)+'</td><td>'+h(i.doz)+'</td><td>'+h(i.ort)+'</td><td>'+h(i.werk)+'</td>'+
          '<td class="r">'+i.min+' min</td><td class="'+(i.vorbei?'ok':'off')+'">'+(i.vorbei?'✓':'○')+'</td></tr>';
      }).join('')+'</tbody></table></div>';
  }).join('');
  drucke('Probezeiten '+(gr.length===1?gr[0].name:'Studierende')+' '+heuteDe(),body);
}

/* ── UI ── */
var BTN='padding:5px 10px;border-radius:7px;border:1px solid var(--border);background:#fff;font-size:11px;cursor:pointer;font-family:var(--sans)';
var MINI='padding:2px 7px;border-radius:6px;border:1px solid var(--border);background:#fff;font-size:10px;cursor:pointer;font-family:var(--sans);margin-left:4px';

window.untProbeExport=function(art,fmt,name){
  if(art==='zeiten')return fmt==='xlsx'?exportZeitenExcel(name):exportZeitenPdf(name);
  return fmt==='xlsx'?exportStudExcel(name):exportStudPdf(name);
};

var _bar=window.untProbeBar;
window.untProbeBar=function(){
  var html=_bar.apply(this,arguments);
  var art=(_untProbeAnsicht==='zeiten')?'zeiten':'stud';
  var tip=art==='zeiten'?'Exportiert die aktuell gefilterte Ansicht (alle Dozent:innen oder die gewählte)':'Exportiert alle Studierenden';
  var grp='<span style="margin-left:auto;display:flex;gap:6px" title="'+tip+'">'+
    '<button style="'+BTN+'" onclick="untProbeExport(\''+art+'\',\'xlsx\')">⬇ Excel</button>'+
    '<button style="'+BTN+'" onclick="untProbeExport(\''+art+'\',\'pdf\')">⬇ PDF</button></span>';
  return html.replace('<button onclick="ladeUntProbezeiten(true)" style="margin-left:auto;','<button onclick="ladeUntProbezeiten(true)" style="')
             .replace('<button onclick="ladeUntProbezeiten(true)"',grp+'<button onclick="ladeUntProbezeiten(true)"');
};

var _render=window.renderUntProbe;
window.renderUntProbe=function(){
  var r=_render.apply(this,arguments);
  try{miniKnoepfe();}catch(e){}
  return r;
};
function miniKnoepfe(){
  var el=document.getElementById('unt-probe-inhalt');if(!el)return;
  var art=(_untProbeAnsicht==='zeiten')?'zeiten':'stud';
  var gueltig={};
  (art==='zeiten'?zeitItems().map(function(i){return i.doz;}):studGruppen().map(function(p){return p.name;}))
    .forEach(function(n){gueltig[n]=1;});
  el.querySelectorAll('.card.card-pad').forEach(function(card){
    var hd=card.firstElementChild&&card.firstElementChild.firstElementChild;
    if(!hd||!hd.firstChild||hd.firstChild.nodeType!==3)return;
    var name=hd.firstChild.textContent;
    if(!name||!gueltig[name]||card.querySelector('.pz-exp'))return;
    var box=document.createElement('span');box.className='pz-exp';
    ['xlsx','pdf'].forEach(function(fmt){
      var b=document.createElement('button');b.type='button';b.style.cssText=MINI;
      b.textContent=fmt==='xlsx'?'⬇ Excel':'⬇ PDF';
      b.title='Nur '+name+' als '+(fmt==='xlsx'?'Excel':'PDF')+' exportieren';
      b.addEventListener('click',function(ev){ev.stopPropagation();untProbeExport(art,fmt,name);});
      box.appendChild(b);
    });
    hd.appendChild(box);
  });
}

// falls der Reiter schon gezeichnet ist: neu zeichnen
try{if(document.getElementById('unt-probe-inhalt')&&(_untTermineAlle||_untBuchDaten))renderUntProbe();}catch(e){}
})();
