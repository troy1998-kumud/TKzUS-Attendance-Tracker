(function(){
  const sheets = {};
  function mkSheet(name){ const data=[]; const sh={
    _d:data, getName:()=>name,
    getMaxRows:()=>1000, getLastRow:()=>{ let n=0; data.forEach((r,i)=>{ if(r&&r.some(v=>v!==''&&v!=null)) n=i+1; }); return n; },
    setFrozenRows(){},
    getRange(r,c,nr=1,nc=1){ return {
      setNumberFormat(){return this}, setFontWeight(){return this}, setDataValidation(){return this},
      getValue(){ return (data[r-1]||[])[c-1]||''; },
      setValues(v){ v.forEach((row,i)=>{ data[r-1+i]=data[r-1+i]||[]; row.forEach((x,j)=>data[r-1+i][c-1+j]=String(x)); }); return this; },
      getDisplayValues(){ const out=[]; for(let i=0;i<nr;i++){ const row=[]; for(let j=0;j<nc;j++) row.push(((data[r-1+i]||[])[c-1+j])??''); out.push(row);} return out; }
    }; } }; return sh; }
  const ss = { getId:()=>'SHEETID', getUrl:()=>'https://docs.google.com/spreadsheets/d/SHEETID', setSpreadsheetTimeZone(){},
    getSheetByName:n=>sheets[n]||null, insertSheet:n=>(sheets[n]=mkSheet(n)), getSheets:()=>Object.values(sheets), deleteSheet(){} };
  window.__sheets = sheets;
  window.SpreadsheetApp = { getActiveSpreadsheet:()=>ss, openById:()=>ss,
    newDataValidation(){ const b={requireValueInList(){return b}, build(){return {}}}; return b; } };
  const props={}; window.PropertiesService={ getScriptProperties:()=>({getProperty:k=>props[k]||null,setProperty:(k,v)=>props[k]=v}) };
  window.LockService={ getScriptLock:()=>({waitLock(){},releaseLock(){}}) };
  window.Utilities={ formatDate(d,tz,f){ const p=Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone:tz,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).formatToParts(d).map(x=>[x.type,x.value])); return `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}`; } };
  window.HtmlService={};
  window.google={ script:{ get run(){ let ok=()=>{}, bad=()=>{}; const r=new Proxy({},{get(_,k){
    if(k==='withSuccessHandler') return f=>{ok=f;return r;}; if(k==='withFailureHandler') return f=>{bad=f;return r;};
    return (...a)=>setTimeout(()=>{ try{ ok(JSON.parse(JSON.stringify(window[k](...a)))); }catch(e){ bad(e); } },30); }}); return r; } } };
})();
