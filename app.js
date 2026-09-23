const $=id=>document.getElementById(id);
let audit={file:"",sheets:[],findings:[],stats:{}};

$("fileInput").addEventListener("change",e=>{const f=e.target.files[0];if(f)runAudit(f)});
$("themeBtn").onclick=()=>document.body.classList.toggle("light");
$("severityFilter").onchange=renderFindings;
$("search").oninput=renderFindings;
$("exportBtn").onclick=exportCSV;

async function runAudit(file){
  $("empty").textContent="Reading and auditing "+file.name+"…";
  $("dashboardContent").hidden=false;
  $("fileName").textContent=file.name;
  const buf=await file.arrayBuffer();
  let wb;
  try{wb=XLSX.read(buf,{type:"array",cellFormula:true,cellNF:true,cellDates:true});}
  catch(err){$("empty").textContent="Could not read this workbook. It may be corrupted or unsupported.";return}
  const findings=[]; let formulaCount=0, cellCount=0, numeric=0, text=0, blank=0;
  const sheets=[];
  for(const name of wb.SheetNames){
    const ws=wb.Sheets[name], ref=ws["!ref"]||"A1";
    const range=XLSX.utils.decode_range(ref);
    const rows=range.e.r-range.s.r+1, cols=range.e.c-range.s.c+1;
    let formulas=0, nonblank=0;
    const rowValues=[];
    for(let r=range.s.r;r<=range.e.r;r++){
      const vals=[];
      for(let c=range.s.c;c<=range.e.c;c++){
        const addr=XLSX.utils.encode_cell({r,c}), cell=ws[addr];
        if(!cell){blank++; vals.push(""); continue}
        cellCount++; if(cell.f){formulaCount++;formulas++}
        if(cell.v!==undefined&&cell.v!==null&&cell.v!==""){nonblank++;vals.push(String(cell.v))}
        else {blank++;vals.push("")}
        if(typeof cell.v==="number") numeric++; else if(typeof cell.v==="string"&&!cell.f) text++;
        const raw=String(cell.v??"");
        if(["#REF!","#DIV/0!","#VALUE!","#N/A","#NAME?","#NUM!","#NULL!"].includes(raw)){
          findings.push({severity:raw==="#REF!"?"Critical":"High",type:"Excel error",sheet:name,cell:addr,description:`Cell contains ${raw}.`});
        }
        if(cell.f){
          const formula=String(cell.f);
          if(/#REF!/.test(formula)) findings.push({severity:"Critical",type:"Broken reference",sheet:name,cell:addr,description:`Formula contains a broken #REF! reference: =${formula}`});
          if(/\/0(?![0-9])/.test(formula)) findings.push({severity:"High",type:"Potential division by zero",sheet:name,cell:addr,description:`Formula may divide by zero: =${formula}`});
        }
      }
      rowValues.push(vals);
    }
    // Detect duplicate non-empty rows (useful for transaction-style sheets).
    const seen=new Map();
    rowValues.forEach((vals,i)=>{
      const key=vals.join("¦").trim();
      if(key && vals.filter(Boolean).length>=2){
        if(seen.has(key)) findings.push({severity:"Medium",type:"Duplicate row",sheet:name,cell:`row ${i+1}`,description:`Row duplicates row ${seen.get(key)} based on populated cell values.`});
        else seen.set(key,i+1);
      }
    });
    // Formula pattern breaks by comparing formulas after replacing row numbers with #.
    const formulaByCol={};
    for(let r=range.s.r;r<=range.e.r;r++){
      for(let c=range.s.c;c<=range.e.c;c++){
        const addr=XLSX.utils.encode_cell({r,c}), cell=ws[addr];
        if(cell?.f){
          const norm=String(cell.f).toUpperCase().replace(/\$?[A-Z]{1,3}\$?\d+/g,m=>m.replace(/\d+/g,"#"));
          (formulaByCol[c]??=[]).push({r,norm,addr});
        }
      }
    }
    for(const arr of Object.values(formulaByCol)){
      for(let i=1;i<arr.length-1;i++){
        if(arr[i-1].norm===arr[i+1].norm && arr[i].norm!==arr[i-1].norm){
          findings.push({severity:"High",type:"Formula pattern break",sheet:name,cell:arr[i].addr,description:`Formula pattern differs from adjacent formulas in this column.`});
        }
      }
    }
    // Detect suspicious constants inside formula-heavy columns.
    for(let c=range.s.c;c<=range.e.c;c++){
      let f=0, constants=[];
      for(let r=range.s.r;r<=range.e.r;r++){
        const cell=ws[XLSX.utils.encode_cell({r,c})];
        if(cell?.f) f++;
        else if(cell?.v!==undefined&&cell?.v!==null&&cell?.v!=="") constants.push({r,cell});
      }
      if(f>=3 && constants.length && constants.length<Math.max(8,rows*.35)){
        constants.forEach(x=>{
          const addr=XLSX.utils.encode_cell({r:x.r,c});
          findings.push({severity:"Low",type:"Hard-coded value",sheet:name,cell:addr,description:`A constant appears in a column containing ${f} formulas. Review whether a formula is missing.`});
        });
      }
    }
    sheets.push({name,rows,cols,formulas,nonblank});
  }
  // Remove duplicate identical findings.
  const unique=[...new Map(findings.map(x=>[`${x.severity}|${x.type}|${x.sheet}|${x.cell}|${x.description}`,x])).values()];
  let score=100;
  unique.forEach(x=>score-=({Critical:15,High:10,Medium:5,Low:2}[x.severity]||0));
  score=Math.max(0,score);
  const critical=unique.filter(x=>x.severity==="Critical").length, high=unique.filter(x=>x.severity==="High").length;
  audit={file:file.name,sheets,findings:unique,stats:{formulaCount,cellCount,numeric,text,blank,critical,high,score}};
  render();
}

function render(){
  const s=audit.stats;
  $("score").textContent=s.score;
  $("riskLabel").textContent=s.score>=90?"Low risk":s.score>=70?"Moderate risk":s.score>=40?"High risk":"Critical risk";
  $("sheets").textContent=audit.sheets.length;
  $("formulas").textContent=s.formulaCount;
  $("findingsCount").textContent=audit.findings.length;
  $("criticalHigh").textContent=`${s.critical} / ${s.high}`;
  $("heroSheets").textContent=audit.sheets.length;
  $("heroFormulas").textContent=s.formulaCount;
  $("heroFindings").textContent=audit.findings.length;
  $("heroScore").textContent=s.score;
  $("exportBtn").disabled=false;
  $("empty").textContent=`Audit complete: ${audit.file}`;
  const counts={Critical:0,High:0,Medium:0,Low:0};audit.findings.forEach(f=>counts[f.severity]++);
  $("bars").innerHTML=Object.entries(counts).map(([k,v])=>`<div class="bar"><div class="barTop"><span>${k}</span><b>${v}</b></div><div class="track"><div class="fill" style="width:${audit.findings.length?Math.min(100,v/audit.findings.length*100):0}%"></div></div></div>`).join("");
  $("overview").innerHTML=[
    ["Workbook",audit.file],["Sheets",audit.sheets.length],["Cells inspected",s.cellCount],["Formula cells",s.formulaCount],["Numeric cells",s.numeric],["Text cells",s.text],["Blank cells",s.blank]
  ].map(([a,b])=>`<div class="overviewRow"><span>${a}</span><b>${b}</b></div>`).join("");
  renderFindings();
}

function renderFindings(){
  const sev=$("severityFilter").value, q=$("search").value.toLowerCase();
  const list=audit.findings.filter(f=>(sev==="All"||f.severity===sev)&&JSON.stringify(f).toLowerCase().includes(q));
  $("findingsBody").innerHTML=list.length?list.map(f=>`<tr><td class="sev ${f.severity}">${f.severity}</td><td>${esc(f.type)}</td><td>${esc(f.sheet)}</td><td>${esc(f.cell)}</td><td>${esc(f.description)}</td></tr>`).join(""):`<tr><td colspan="5" class="muted">No findings match the filter.</td></tr>`;
}
function esc(v){return String(v).replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m]))}
function exportCSV(){
  const rows=[["Severity","Type","Sheet","Cell","Description"],...audit.findings.map(f=>[f.severity,f.type,f.sheet,f.cell,f.description])];
  const csv=rows.map(r=>r.map(v=>`"${String(v).replaceAll('"','""')}"`).join(",")).join("\n");
  const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv"}));a.download="sheetguard-findings.csv";a.click();URL.revokeObjectURL(a.href);
}
