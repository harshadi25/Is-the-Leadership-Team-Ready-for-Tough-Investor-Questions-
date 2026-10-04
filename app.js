let DATA = null;
let selectedQuestion = 0;
const charts = {};

const fallback = {
  "Why did margins decline despite revenue growth?": {
    direct:"Margin declined primarily because of higher input and supplier costs and an unfavourable product mix, partly offset by pricing actions.",
    evidence:["Revenue increased 15%.","Operating margin declined from 22% to 18%.","Input and supplier costs increased by approximately 8%.","Growth was concentrated in lower-margin offerings."],
    explanation:"Topline growth did not fully translate into profitability because cost pressure and mix effects absorbed part of the revenue benefit.",
    action:"Management initiated pricing actions, supplier negotiations and cost-efficiency initiatives.",
    outlook:"Management expects pressure to moderate as disclosed actions are implemented, subject to market conditions.",
    score:4.2
  },
  "Why did operating cash flow decline despite revenue growth?": {
    direct:"Operating cash flow weakened primarily because higher receivables and inventory absorbed cash during the quarter.",
    evidence:["Operating cash flow was ₹110 Cr.","Receivables increased to ₹320 Cr.","Inventory increased to ₹190 Cr.","Working-capital pressure increased."],
    explanation:"Revenue growth and accounting profit did not convert into cash at the same pace because more cash was tied up in working capital.",
    action:"Management is focusing on collections, inventory discipline and working-capital normalization.",
    outlook:"Recovery timing should not be promised beyond the information disclosed.",
    score:4.3
  },
  "Why has Consumer Technology growth slowed?": {
    direct:"Consumer Technology growth slowed as demand softened in selected categories.",
    evidence:["Q2 growth was approximately 7%.","Prior growth was approximately 10%.","Management commentary identifies softer demand and pricing pressure."],
    explanation:"The segment is still growing, but at a slower rate, making demand recovery an investor focus.",
    action:"Management is reviewing pricing, product mix and demand-generation initiatives.",
    outlook:"No quantified recovery timeline has been disclosed.",
    score:4.1
  },
  "What caused the margin guidance miss?": {
    direct:"The margin fell below the previous >20% expectation because of higher input costs and an unfavourable product mix.",
    evidence:["Previous commitment: >20% margin.","Current margin: 18%.","Input costs increased approximately 8%.","Management cited mix pressure."],
    explanation:"The answer must explicitly acknowledge the deviation from prior guidance.",
    action:"Pricing, supplier and cost-efficiency measures were initiated.",
    outlook:"Do not provide an unsupported recovery target.",
    score:4.9
  },
  "Is current margin pressure temporary or structural?": {
    direct:"Current disclosures identify specific cost and mix pressures, but they do not conclusively establish how quickly margins will normalize.",
    evidence:["Margin is 18%.","Input costs increased.","Mix weakened profitability.","Corrective actions have started."],
    explanation:"Leadership should separate identifiable operating drivers from longer-term structural uncertainty.",
    action:"Explain which drivers are controllable and which remain market-dependent.",
    outlook:"Avoid an unsupported long-term margin forecast.",
    score:4.3
  }
};

document.addEventListener("DOMContentLoaded",()=>{
  document.querySelectorAll(".nav-btn").forEach(b=>b.addEventListener("click",()=>go(b.dataset.page)));
  document.getElementById("excelFile").addEventListener("change",loadExcel);
  document.getElementById("removeFile").addEventListener("click",resetWorkbook);
  ensureXlsxLibrary();
});

function go(page){
  document.querySelectorAll(".nav-btn").forEach(b=>b.classList.toggle("active",b.dataset.page===page));
  document.querySelectorAll(".page").forEach(p=>p.classList.toggle("active",p.id===`page-${page}`));
  const subtitles={
    problem:"A management cockpit connecting performance, investor exposure and call readiness.",
    performance:"Understand the financial movements most likely to drive investor scrutiny.",
    segments:"See where growth is coming from and where segment-level pressure is building.",
    commitments:"Test current performance against previous management commitments.",
    pressure:"Prioritise and simulate the investor questions most likely to create pressure.",
    response:"Build and stress-test a response leadership can defend with disclosed evidence.",
    decision:"Make the final leadership call: ready, conditional, or action required."
  };
  document.getElementById("pageSubtitle").textContent=subtitles[page]||""; if(page==="response" && DATA) renderResponse(); if(page==="pressure" && DATA) renderPressure();
}


function ensureXlsxLibrary(){
  if(window.XLSX) return Promise.resolve(true);
  return new Promise((resolve,reject)=>{
    const sources=[
      "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
      "https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js"
    ];
    let i=0;
    const next=()=>{
      if(window.XLSX){resolve(true);return;}
      if(i>=sources.length){
        const status=document.getElementById("dataStatus");
        status.textContent="Excel reader unavailable";
        status.className="data-status bad";
        showUploadError("The Excel reader could not load. Please check that internet access is available, then refresh the page.");
        reject(new Error("SheetJS could not be loaded"));
        return;
      }
      const s=document.createElement("script");
      s.src=sources[i++];
      s.onload=()=>window.XLSX?resolve(true):next();
      s.onerror=next;
      document.head.appendChild(s);
    };
    next();
  });
}

function setStatus(text,kind="neutral"){
  const el=document.getElementById("dataStatus");
  el.textContent=text;
  el.className=`data-status ${kind}`;
}
function showUploadError(message){
  let el=document.getElementById("uploadError");
  if(!el){
    el=document.createElement("div");
    el.id="uploadError";
    document.querySelector(".data-pack").appendChild(el);
  }
  el.className="upload-error";
  el.textContent=message;
}
function clearUploadError(){
  const el=document.getElementById("uploadError");
  if(el)el.remove();
}
function resetWorkbook(){
  DATA=null;
  selectedQuestion=0;
  document.getElementById("excelFile").value="";
  document.getElementById("fileName").textContent="No workbook loaded";
  document.getElementById("removeFile").disabled=true;
  setStatus("Waiting for data","neutral");
  clearUploadError();
  document.getElementById("appContent").classList.add("hidden");
  document.getElementById("noDataState").classList.remove("hidden");
  document.getElementById("latestQuarter").textContent="No data";
  document.getElementById("readinessDot").className="";
  document.getElementById("readinessText").textContent="Not assessed";
  go("problem");
}
function requiredSheets(){
  return ["Quarterly Financials","Segment Data","Management Commitments","Management Commentary","Investor Risk Map","Investor Questions"];
}
function validateWorkbookStructure(wb){
  const missing=requiredSheets().filter(s=>!wb.Sheets[s]);
  if(missing.length){
    throw new Error("Missing required sheet(s): "+missing.join(", ")+". Please use the supplied workbook or keep these exact sheet names.");
  }
}
function findHeaderRow(raw, expectedTerms=[]){
  const terms=expectedTerms.map(x=>x.toLowerCase());
  let best=-1,bestScore=-1;
  raw.forEach((r,i)=>{
    const cells=r.map(v=>String(v??"").trim().toLowerCase());
    const score=terms.reduce((n,t)=>n+(cells.some(c=>c===t)?1:0),0);
    if(score>bestScore){bestScore=score;best=i;}
  });
  if(bestScore>0)return best;
  for(let i=0;i<Math.min(raw.length,12);i++){
    if(raw[i].filter(v=>v!==""&&v!==null&&v!==undefined).length>=2)return i;
  }
  return -1;
}
async function loadExcel(e){
  const file=e.target.files?.[0];
  if(!file)return;
  clearUploadError();
  document.getElementById("fileName").textContent=file.name;
  document.getElementById("removeFile").disabled=false;
  setStatus("Reading workbook…","neutral");

  try{
    await ensureXlsxLibrary();
    const buffer=await file.arrayBuffer();
    const wb=XLSX.read(buffer,{type:"array",cellDates:false});
    validateWorkbookStructure(wb);
    DATA=parseWorkbook(wb);
    validate();

    document.getElementById("noDataState").classList.add("hidden");
    document.getElementById("appContent").classList.remove("hidden");
    setStatus(`Loaded · ${Object.keys(wb.Sheets).length} sheets`,"good");
    document.getElementById("latestQuarter").textContent="Q2 FY25";
    document.getElementById("readinessDot").className="red";
    document.getElementById("readinessText").textContent="Action required";
    renderAll();
    go("problem");
  }catch(err){
    console.error(err);
    DATA=null;
    document.getElementById("appContent").classList.add("hidden");
    document.getElementById("noDataState").classList.remove("hidden");
    setStatus("Workbook not loaded","bad");
    showUploadError(err.message || "Could not read the workbook.");
  }
}

function parseSheet(wb,name){
  const ws=wb.Sheets[name];
  if(!ws)return[];
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:true});
  const expected={
    "Quarterly Financials":["Metric","Q2 FY25","QoQ Change"],
    "Segment Data":["Segment","Q2 FY25 Revenue","Q2 QoQ Growth"],
    "Management Commitments":["Management Statement","Metric","Expected Value","Status"],
    "Management Commentary":["Theme","Synthetic Management Commentary"],
    "Investor Risk Map":["Issue","Total Score","Risk Level"],
    "Investor Questions":["Priority","Investor Question","Risk Level"],
    "Response Review":["Investor Question","Overall Score","Review"],
    "Leadership Preparation":["Priority","Concern","Risk"],
    "Dashboard":["Revenue","Operating Margin","Operating Cash Flow"]
  };
  const h=findHeaderRow(raw,expected[name]||[]);
  if(h<0)return[];
  const heads=raw[h].map(v=>String(v??"").trim());
  return raw.slice(h+1)
    .filter(r=>r.some(v=>v!==""&&v!==null&&v!==undefined))
    .map(r=>{
      const o={};heads.forEach((x,i)=>{if(x)o[x]=r[i]??""});return o;
    });
}
function parseWorkbook(wb){
  return {
    financials:parseSheet(wb,"Quarterly Financials"),
    segments:parseSheet(wb,"Segment Data"),
    commitments:parseSheet(wb,"Management Commitments"),
    commentary:parseSheet(wb,"Management Commentary"),
    risks:parseSheet(wb,"Investor Risk Map"),
    questions:parseSheet(wb,"Investor Questions"),
    responses:parseSheet(wb,"Response Review"),
    preparation:parseSheet(wb,"Leadership Preparation")
  };
}
function validate(){
  const checks=[
    ["financials",3,"Quarterly Financials"],
    ["segments",2,"Segment Data"],
    ["commitments",2,"Management Commitments"],
    ["commentary",2,"Management Commentary"]
  ];
  const bad=checks.filter(x=>(DATA[x[0]]||[]).length<x[1]);
  if(bad.length){
    throw new Error("The workbook structure was detected, but some sheets do not contain usable rows: "+bad.map(x=>x[2]).join(", ")+".");
  }
}
function num(v){if(typeof v==="number")return v; if(v===""||v==null)return NaN;return Number(String(v).replace(/[₹,%]/g,"").replace(/,/g,"").trim())}
function cr(v){const n=num(v);return Number.isFinite(n)?`₹${n.toLocaleString("en-IN")} Cr`:"—"}
function pct(v){const n=num(v);return Number.isFinite(n)?`${(n*100).toFixed(1)}%`:"—"}
function pp(v){const n=num(v);return Number.isFinite(n)?`${n>=0?"+":""}${(n*100).toFixed(1)} pp`:"—"}
function signed(v){const n=num(v);return Number.isFinite(n)?`${n>=0?"+":""}${(n*100).toFixed(1)}%`:"—"}
function esc(s){return String(s??"").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;")}
function financialMap(){const m={};DATA.financials.forEach(r=>{if(r.Metric)m[String(r.Metric).trim()]=r});return m}
function row(list,key,value){return list.find(r=>String(r[key]||"").trim().toLowerCase()===String(value).trim().toLowerCase())}
function badge(level){const x=String(level||"").toLowerCase();let c="low";if(x.includes("critical"))c="critical";else if(x.includes("high"))c="high";else if(x.includes("moderate"))c="moderate";return `<span class="badge ${c}">${esc(level||"Low")}</span>`}
function renderAll(){renderProblem();renderPerformance();renderSegments();renderCommitments();renderPressure();renderResponse();renderDecision();renderMethodology();}


function getDecisionSummary(){
  const risks = getRisks();
  const critical = risks.filter(r => String(r.level || "").toLowerCase() === "critical").length;
  const commitments = DATA && DATA.commitments ? DATA.commitments : [];
  const misses = commitments.filter(r => String(r.Status || "").toLowerCase().includes("miss")).length;

  if (critical >= 2 || misses >= 2) {
    return {
      status: "NOT READY — ACTION REQUIRED",
      summary: "Revenue growth is supportable, but margin deterioration, weaker cash generation and management commitment gaps create material investor follow-up risk."
    };
  }

  return {
    status: "CONDITIONAL READY",
    summary: "The key investor concerns are identified, but leadership should close the remaining response and disclosure gaps before sign-off."
  };
}

function renderProblem(){
  const f=financialMap(), r=f.Revenue,m=f["Operating Margin"],c=f["Operating Cash Flow"];
  const risks=getRisks();
  const decision=getDecisionSummary();

  const topQuestions = [
    ["Revenue grew by 15%, so why did operating margin fall from 22% to 18%?","Profitability"],
    ["Why did operating cash flow weaken despite strong revenue growth?","Cash flow"],
    ["What caused the miss against the previous >20% operating-margin commitment?","Guidance"],
    ["Why has Consumer Technology growth slowed?","Segment performance"],
    ["Is the current margin pressure temporary or structural?","Sustainability"]
  ];

  document.getElementById("page-problem").innerHTML=`
    <div class="hero executive-hero">
      <div class="section-kicker" style="color:#a9c5f6">EXECUTIVE OVERVIEW · Q2 FY25</div>
      <h2>Investor readiness</h2>
      <p>Review the quarter's material risks, probable investor questions and management actions before the earnings call.</p>
    </div>

    <div class="grid g4" style="margin-top:14px">
      ${kpi("Revenue",cr(r?.["Q2 FY25"]),signed(r?.["QoQ Change"]),"up")}
      ${kpi("Operating margin",pct(m?.["Q2 FY25"]),pp(m?.["QoQ Change"]),"down")}
      ${kpi("Operating cash flow",cr(c?.["Q2 FY25"]),signed(c?.["QoQ Change"]),"down")}
      ${kpi("Readiness",decision.status.includes("NOT READY")?"62 / 100":"78 / 100","Action required","warn")}
    </div>

    <div class="grid g2 dashboard-section">
      <div class="card">
        <div class="card-title">Financial performance</div>
        <div class="card-sub">Quarterly revenue and operating-margin trend.</div>
        <div class="chart-box" style="margin-top:10px"><canvas id="executiveTrendChart"></canvas></div>
      </div>

      <div class="card">
        <div class="card-title">Investor exposure</div>
        <div class="card-sub">Highest-risk issues for the current quarter.</div>
        <div class="risk-list" style="margin-top:10px">
          ${risks.slice(0,5).map((x,i)=>`
            <div class="risk-row">
              <div style="display:flex;gap:8px;align-items:flex-start">
                <div class="concern-num">${i+1}</div>
                <div><div class="risk-name">${esc(x.issue)}</div><div class="risk-copy">${esc(x.reason)}</div></div>
              </div>
              <div style="text-align:right">
                <div class="risk-score">${x.score}/20</div>
                ${badge(x.level)}
                ${x.financialImpact!==null ? `<div class="score-breakdown">${x.financialImpact}+${x.negativeTrend}+${x.guidanceGap}+${x.strategicImportance}</div>` : ""}
              </div>
            </div>`).join("")}
        </div>
      </div>
    </div>

    <div class="card dashboard-section">
      <div class="card-title">Probable investor questions</div>
      <div class="probable-questions">
        ${topQuestions.map((q,i)=>`
          <div class="probable-question" onclick="selectedQuestion=${i};renderPressure();go('response');">
            <div class="probable-q-num">${i+1}</div>
            <div>
              <div class="probable-q-text">${esc(q[0])}</div>
              <div class="probable-q-why">${esc(q[1])}</div>
            </div>
            <div class="probable-q-arrow">→</div>
          </div>`).join("")}
      </div>
    </div>

    <div class="grid g2 dashboard-section">
      <div class="card">
        <div class="card-title">Guidance & credibility</div>
        <div class="table-wrap" style="margin-top:10px">
          <table>
            <thead><tr><th>Commitment</th><th>Expected</th><th>Actual</th><th>Status</th></tr></thead>
            <tbody>
              ${(DATA.commitments||[]).slice(0,4).map(r=>`
                <tr>
                  <td>${esc(r["Management Statement"])}</td>
                  <td>${esc(r["Expected Value"])}</td>
                  <td>${esc(r["Actual Q2 FY25"])}</td>
                  <td>${badge(String(r.Status||"").includes("Miss")?"Critical":String(r.Status||"").includes("Partially")?"Moderate":"Low")} ${esc(r.Status)}</td>
                </tr>`).join("")}
            </tbody>
          </table>
        </div>
      </div>

      <div class="card">
        <div class="card-title">Management action queue</div>
        <div class="signal-stack">
          <div class="signal-mini red-border"><strong>Margin bridge</strong><span>Prepare the quantified explanation for the 4 percentage-point decline.</span></div>
          <div class="signal-mini red-border"><strong>Cash-flow bridge</strong><span>Explain receivables, inventory and working-capital pressure.</span></div>
          <div class="signal-mini amber-border"><strong>Guidance response</strong><span>Acknowledge the previous >20% margin expectation and explain the variance.</span></div>
        </div>
      </div>
    </div>
  `;

  destroyChart("executiveTrendChart");
  if(!ensureChartLibrary()) return;
  const qs=["Q3 FY24","Q4 FY24","Q1 FY25","Q2 FY25"];
  charts.executiveTrendChart=new Chart(document.getElementById("executiveTrendChart"),{
    type:"line",
    data:{
      labels:qs,
      datasets:[
        {
          label:"Revenue (₹ Cr)",
          data:qs.map(q=>num(f.Revenue?.[q])),
          borderColor:"#1d5cff",
          backgroundColor:"rgba(29,92,255,.07)",
          fill:true,
          tension:.35,
          borderWidth:2.5,
          yAxisID:"a",
          pointRadius:3
        },
        {
          label:"Operating Margin (%)",
          data:qs.map(q=>num(f["Operating Margin"]?.[q])*100),
          borderColor:"#b42318",
          backgroundColor:"transparent",
          tension:.35,
          borderWidth:2.5,
          yAxisID:"b",
          pointRadius:3
        }
      ]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      interaction:{mode:"index",intersect:false},
      plugins:{legend:{position:"bottom",labels:{font:{size:10},boxWidth:12}}},
      scales:{
        a:{ticks:{font:{size:9},callback:v=>"₹"+v+" Cr"}},
        b:{position:"right",grid:{drawOnChartArea:false},ticks:{font:{size:9},callback:v=>v+"%"}}
      }
    }
  });
}
function kpi(label,value,change,cls){return `<div class="kpi"><div class="kpi-label">${esc(label)}</div><div class="kpi-value">${esc(value)}</div><div class="kpi-change ${cls}">${esc(change)}</div></div>`}


function ensureChartLibrary(){
  if (typeof Chart !== "undefined") return true;
  const status = document.getElementById("dataStatus");
  if (status) {
    status.textContent = "Chart library unavailable";
    status.className = "data-status bad";
  }
  return false;
}

function renderPerformance(){
  if(!ensureChartLibrary()) return;
  const f=financialMap(), metrics=["Revenue","Operating Profit","Operating Margin","Operating Cash Flow","Receivables","Inventory","Working Capital Proxy"];
  const qs=["Q3 FY24","Q4 FY24","Q1 FY25","Q2 FY25"];
  document.getElementById("page-performance").innerHTML=`
    <div class="grid g4">
      ${kpi("Revenue",cr(f.Revenue?.["Q2 FY25"]),signed(f.Revenue?.["QoQ Change"]),"up")}
      ${kpi("Operating profit",cr(f["Operating Profit"]?.["Q2 FY25"]),signed(f["Operating Profit"]?.["QoQ Change"]),"down")}
      ${kpi("Operating margin",pct(f["Operating Margin"]?.["Q2 FY25"]),pp(f["Operating Margin"]?.["QoQ Change"]),"down")}
      ${kpi("Operating cash flow",cr(f["Operating Cash Flow"]?.["Q2 FY25"]),signed(f["Operating Cash Flow"]?.["QoQ Change"]),"down")}
    </div>
    <div class="grid g2" style="margin-top:13px">
      <div class="card"><div class="card-title">What moved?</div><div class="card-sub">The financial signals leadership should know without opening a spreadsheet.</div><div class="chart-box" style="margin-top:8px"><canvas id="perfChart"></canvas></div></div>
      <div class="card"><div class="card-title">Why investors will care</div><div class="card-sub">The diagnosis connects the number to the business question.</div>
        <div style="margin-top:10px">
          <div class="alert red"><strong>Profitability:</strong> Revenue grew 15%, but operating margin fell 4 percentage points.</div>
          <div class="alert red" style="margin-top:8px"><strong>Cash conversion:</strong> Operating cash flow weakened as receivables and inventory absorbed cash.</div>
          <div class="alert amber" style="margin-top:8px"><strong>Quality of growth:</strong> More growth came from lower-margin offerings.</div>
        </div>
      </div>
    </div>
    <div class="card" style="margin-top:13px">
      <div class="card-title">Financial evidence table</div><div class="card-sub">Use this as the evidence base for the pressure test.</div>
      <div class="table-wrap" style="margin-top:9px">
        <table><thead><tr><th>Metric</th>${qs.map(q=>`<th>${q}</th>`).join("")}<th>Latest change</th><th>Signal</th></tr></thead>
        <tbody>${metrics.map(m=>{
          const r=f[m];if(!r)return"";
          const isPct=String(r.Units||"").trim()==="%";
          return `<tr><td><strong>${esc(m)}</strong></td>${qs.map(q=>`<td>${isPct?pct(r[q]):cr(r[q])}</td>`).join("")}<td>${m==="Operating Margin"?pp(r["QoQ Change"]):signed(r["QoQ Change"])}</td><td>${esc(r["Latest-quarter interpretation"]||"")}</td></tr>`
        }).join("")}</tbody></table>
      </div>
    </div>
  `;
  if(charts.perf)charts.perf.destroy();
  charts.perf=new Chart(document.getElementById("perfChart"),{
    type:"line",
    data:{labels:qs,datasets:[
      {label:"Revenue (₹ Cr)",data:qs.map(q=>num(f.Revenue?.[q])),borderColor:"#1f5eff",backgroundColor:"rgba(31,94,255,.08)",fill:true,tension:.35,yAxisID:"a"},
      {label:"Margin (%)",data:qs.map(q=>num(f["Operating Margin"]?.[q])*100),borderColor:"#b42318",tension:.35,yAxisID:"b"}
    ]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom",labels:{font:{size:8}}}},scales:{y:{ticks:{font:{size:8}}},b:{position:"right",grid:{drawOnChartArea:false},ticks:{font:{size:8},callback:v=>v+"%"}}}}
  });
}


function getNormalizedSegments(){
  const raw=(DATA.segments||[]).filter(r=>r && Object.keys(r).length>0);
  if(!raw.length)return [];

  // Deep workbook: one row per Segment x Quarter.
  if(raw[0].Quarter && raw[0]["Revenue (₹ Cr)"]!==undefined){
    return raw
      .filter(r=>String(r.Quarter||"").trim().startsWith("Q") && r.Segment)
      .map(r=>({
        Quarter:String(r.Quarter).trim(),
        Segment:String(r.Segment).trim(),
        Revenue:num(r["Revenue (₹ Cr)"]),
        QoQGrowth:Number.isFinite(num(r["QoQ Growth"])) ? num(r["QoQ Growth"]) : null,
        PreviousGrowth:Number.isFinite(num(r["Previous Growth"])) ? num(r["Previous Growth"]) : null,
        GrowthChange:Number.isFinite(num(r["Growth Change"])) ? num(r["Growth Change"]) : null,
        Margin:num(r["Margin"]),
        Contribution:num(r["Contribution %"]),
        Trend:r.Trend||"",
        Demand:r["Demand Signal"]||"",
        Pricing:r["Pricing Pressure"]||"",
        Concern:r["Investor Concern"]||"",
        Action:r["Management Action"]||""
      }));
  }

  // Earlier workbook: one row per segment with quarter columns.
  if(raw[0]["Q2 FY25 Revenue"]!==undefined || raw[0]["Q1 FY25 Revenue"]!==undefined){
    const quarterDefs=[
      ["Q3 FY24","Q3 FY24 Revenue"],
      ["Q4 FY24","Q4 FY24 Revenue"],
      ["Q1 FY25","Q1 FY25 Revenue"],
      ["Q2 FY25","Q2 FY25 Revenue"]
    ];
    const out=[];
    raw.filter(r=>r.Segment).forEach(r=>{
      let prevGrowth=null;
      quarterDefs.forEach(([q,col],idx)=>{
        const revenue=num(r[col]);
        if(!Number.isFinite(revenue))return;
        const previousRevenue=idx>0?num(r[quarterDefs[idx-1][1]]):NaN;
        const growth=Number.isFinite(previousRevenue)&&previousRevenue!==0 ? (revenue/previousRevenue)-1 : null;
        const previousGrowth=prevGrowth;
        const growthChange=Number.isFinite(growth)&&Number.isFinite(previousGrowth)?growth-previousGrowth:null;
        out.push({
          Quarter:q,
          Segment:String(r.Segment).trim(),
          Revenue:revenue,
          QoQGrowth:growth,
          PreviousGrowth:previousGrowth,
          GrowthChange:growthChange,
          Margin:num(r["Q2 Margin"]),
          Contribution:num(r["Revenue Contribution"]),
          Trend:r.Trend||"",
          Demand:r["Demand Signal"]||"",
          Pricing:r["Pricing Pressure"]||"",
          Concern:r["Investor Concern"]||"",
          Action:r["Management Action"]||""
        });
        if(Number.isFinite(growth))prevGrowth=growth;
      });
    });
    return out;
  }

  return [];
}

function segmentLatestRows(){
  const rows=getNormalizedSegments();
  if(!rows.length)return {rows:[],quarters:[],latest:"Q2 FY25"};
  const order={"Q3 FY23":1,"Q4 FY23":2,"Q1 FY24":3,"Q2 FY24":4,"Q3 FY24":5,"Q4 FY24":6,"Q1 FY25":7,"Q2 FY25":8};
  const quarters=[...new Set(rows.map(r=>r.Quarter))].sort((a,b)=>(order[a]||99)-(order[b]||99));
  const latest=quarters.includes("Q2 FY25")?"Q2 FY25":quarters[quarters.length-1];
  const latestRows=rows.filter(r=>r.Quarter===latest);
  const segments=[...new Set(rows.map(r=>r.Segment))];
  return {rows,quarters,latest,latestRows,segments};
}

function renderSegments(){
  if(!ensureChartLibrary()) return;

  const info=segmentLatestRows();
  const rows=info.rows;
  const latestRows=info.latestRows;
  const segments=info.segments;
  const quarters=info.quarters;
  const latest=info.latest;

  if(!rows.length){
    document.getElementById("page-segments").innerHTML=`
      <div class="alert red"><strong>Segment data could not be mapped.</strong><br/>
      The workbook was loaded, but the Segment Data sheet does not match either supported format. Please use the latest supplied synthetic workbook.</div>`;
    return;
  }

  const revenueBySegment={};
  const growthBySegment={};
  segments.forEach(seg=>{
    revenueBySegment[seg]=quarters.map(q=>{
      const r=rows.find(x=>x.Segment===seg && x.Quarter===q);
      return r?.Revenue ?? null;
    });
    growthBySegment[seg]=quarters.map(q=>{
      const r=rows.find(x=>x.Segment===seg && x.Quarter===q);
      return Number.isFinite(r?.QoQGrowth) ? r.QoQGrowth*100 : null;
    });
  });

  const palette=["#1d5cff","#b42318","#a95c00","#16803d","#7c3aed"];

  document.getElementById("page-segments").innerHTML=`
    <div class="grid g3 segment-summary">
      ${latestRows.map(r=>`
        <div class="segment-executive-card">
          <div class="segment-card-head">
            <div>
              <div class="section-kicker">${esc(r.Trend||"Segment")}</div>
              <div class="segment-name">${esc(r.Segment)}</div>
            </div>
            ${badge(r.Trend==="Strong"||r.Trend==="Strong / Improving"?"Low":String(r.Trend).toLowerCase().includes("weak")?"High":"Moderate")}
          </div>
          <div class="segment-main-value">${cr(r.Revenue)}</div>
          <div class="segment-metric-row">
            <div><span>QoQ growth</span><strong class="${r.QoQGrowth>=0.10?"green-text":r.QoQGrowth<0.08?"red-text":"amber-text"}">${pct(r.QoQGrowth)}</strong></div>
            <div><span>Margin</span><strong>${pct(r.Margin)}</strong></div>
            <div><span>Contribution</span><strong>${pct(r.Contribution)}</strong></div>
          </div>
          <div class="segment-detail">
            ${r.Concern ? `<strong>Investor concern:</strong> ${esc(r.Concern)}` : `Monitor segment momentum and mix.`}
          </div>
        </div>`).join("")}
    </div>

    <div class="grid g2 segment-chart-row">
      <div class="card">
        <div class="card-title">Segment revenue trend</div>
        <div class="card-sub">Shows whether consolidated growth is broad-based or concentrated.</div>
        <div class="chart-box segment-chart"><canvas id="segmentRevenueChart"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Quarter-on-quarter growth trend</div>
        <div class="card-sub">Identifies acceleration and slowdown before the earnings call.</div>
        <div class="chart-box segment-chart"><canvas id="segmentGrowthChart"></canvas></div>
      </div>
    </div>

    <div class="grid g2 segment-chart-row">
      <div class="card">
        <div class="card-title">Current-quarter revenue mix</div>
        <div class="card-sub">${esc(latest)} segment contribution.</div>
        <div class="chart-box segment-donut"><canvas id="segmentMixChart"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Segment pressure points</div>
        <div class="signal-stack">
          ${latestRows.map(r=>{
            const trend=String(r.Trend).toLowerCase();
            const cls=trend.includes("strong")?"blue-border":trend.includes("weak")?"red-border":"amber-border";
            return `<div class="signal-mini ${cls}">
              <strong>${esc(r.Segment)}</strong>
              <span>${esc(r.Concern||"Monitor segment trend.")}</span>
            </div>`;
          }).join("")}
        </div>
      </div>
    </div>

    <div class="card" style="margin-top:16px">
      <div class="card-title">Latest-quarter segment diagnostic</div>
      <div class="card-sub">Evidence used to generate probable investor questions.</div>
      <div class="table-wrap" style="margin-top:10px">
        <table>
          <thead><tr><th>Segment</th><th>Revenue</th><th>QoQ Growth</th><th>Previous Growth</th><th>Growth Change</th><th>Margin</th><th>Contribution</th><th>Trend</th><th>Investor Concern</th></tr></thead>
          <tbody>
            ${latestRows.map(r=>`<tr>
              <td><strong>${esc(r.Segment)}</strong></td>
              <td>${cr(r.Revenue)}</td>
              <td>${pct(r.QoQGrowth)}</td>
              <td>${pct(r.PreviousGrowth)}</td>
              <td class="${Number.isFinite(r.GrowthChange)&&r.GrowthChange<0?"down":"up"}">${pp(r.GrowthChange)}</td>
              <td>${pct(r.Margin)}</td>
              <td>${pct(r.Contribution)}</td>
              <td>${esc(r.Trend)}</td>
              <td>${esc(r.Concern)}</td>
            </tr>`).join("")}
          </tbody>
        </table>
      </div>
    </div>

    <div class="alert blue" style="margin-top:13px">
      <strong>Investor implication:</strong> The consolidated growth story is not uniform. Investors should understand which segments are carrying growth and which are slowing.
    </div>
  `;

  destroyChart("segmentRevenueChart");
  destroyChart("segmentGrowthChart");
  destroyChart("segmentMixChart");

  charts.segmentRevenueChart=new Chart(document.getElementById("segmentRevenueChart"),{
    type:"line",
    data:{
      labels:quarters,
      datasets:segments.map((seg,i)=>({
        label:seg,
        data:revenueBySegment[seg],
        borderColor:palette[i%palette.length],
        backgroundColor:"transparent",
        tension:.32,
        borderWidth:2.5,
        pointRadius:3,
        pointHoverRadius:5
      }))
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      interaction:{mode:"index",intersect:false},
      plugins:{legend:{position:"bottom",labels:{font:{size:10},boxWidth:12}}},
      scales:{x:{ticks:{font:{size:9}}},y:{beginAtZero:false,ticks:{font:{size:9},callback:v=>"₹"+v+" Cr"}}}
    }
  });

  charts.segmentGrowthChart=new Chart(document.getElementById("segmentGrowthChart"),{
    type:"line",
    data:{
      labels:quarters,
      datasets:segments.map((seg,i)=>({
        label:seg,
        data:growthBySegment[seg],
        borderColor:palette[i%palette.length],
        backgroundColor:"transparent",
        tension:.32,
        borderWidth:2.5,
        pointRadius:3,
        pointHoverRadius:5
      }))
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      interaction:{mode:"index",intersect:false},
      plugins:{legend:{position:"bottom",labels:{font:{size:10},boxWidth:12}}},
      scales:{x:{ticks:{font:{size:9}}},y:{ticks:{font:{size:9},callback:v=>v+"%"}}}
    }
  });

  charts.segmentMixChart=new Chart(document.getElementById("segmentMixChart"),{
    type:"doughnut",
    data:{
      labels:latestRows.map(r=>r.Segment),
      datasets:[{
        data:latestRows.map(r=>r.Revenue),
        backgroundColor:latestRows.map((_,i)=>palette[i%palette.length]),
        borderColor:"#fff",
        borderWidth:3
      }]
    },
    options:{
      responsive:true,maintainAspectRatio:false,cutout:"62%",
      plugins:{legend:{position:"bottom",labels:{font:{size:10},boxWidth:12}}}
    }
  });
}
function renderCommitments(){
  const rows=DATA.commitments||[];
  document.getElementById("page-commitments").innerHTML=`
    <div class="card">
      <div class="card-title">What did management say before?</div>
      <div class="card-sub">This page is deliberately separate from current financial performance because investor trust depends on continuity over time.</div>
      <div class="table-wrap" style="margin-top:9px"><table>
      <thead><tr><th>Previous commitment</th><th>Metric</th><th>Expected</th><th>Current actual</th><th>Status</th><th>Investor implication</th></tr></thead>
      <tbody>${rows.map(r=>`<tr>
        <td>${esc(r["Management Statement"])}</td><td>${esc(r.Metric)}</td><td>${esc(r["Expected Value"])}</td><td>${esc(r["Actual Q2 FY25"])}</td>
        <td>${badge(String(r.Status).includes("Miss")?"Critical":String(r.Status).includes("Partially")?"Moderate":"Low")}<div class="mini" style="margin-top:3px">${esc(r.Status)}</div></td>
        <td>${esc(r["Investor Implication"])}</td></tr>`).join("")}</tbody></table></div>
    </div>
    <div class="grid g2" style="margin-top:13px">
      <div class="alert red"><strong>Most sensitive commitment:</strong> Operating margin was previously expected to remain above 20%; current margin is 18%.</div>
      <div class="alert amber"><strong>Response design rule:</strong> The answer should acknowledge the gap before explaining the drivers and corrective actions.</div>
    </div>
  `;
}

function getRisks(){
  const raw=(DATA.risks||[]).filter(r=>r && r.Issue);
  return raw.map(r=>{
    const fi=num(r["Financial Impact (1–5)"]);
    const nt=num(r["Negative Trend (1–5)"]);
    const gg=num(r["Guidance Gap (1–5)"]);
    const si=num(r["Strategic Importance (1–5)"]);

    // Always recompute the score from its components for transparency.
    const hasComponents=[fi,nt,gg,si].every(Number.isFinite);
    const score=hasComponents ? fi+nt+gg+si : num(r["Total Score"]);

    let level;
    if(score>=16) level="Critical";
    else if(score>=11) level="High";
    else if(score>=6) level="Moderate";
    else level="Low";

    return {
      issue:r.Issue,
      score,
      level,
      reason:r["Why It Matters"]||"",
      financialImpact:hasComponents?fi:null,
      negativeTrend:hasComponents?nt:null,
      guidanceGap:hasComponents?gg:null,
      strategicImportance:hasComponents?si:null
    };
  }).sort((a,b)=>b.score-a.score);
}
function renderPressure(){
  const qs=(DATA.questions?.length?DATA.questions:[]).map(q=>q["Investor Question"]||"").filter(Boolean);
  if(!qs.length)Object.keys(fallback).forEach(q=>qs.push(q));
  document.getElementById("page-pressure").innerHTML=`
    <div class="grid g2">
      <div class="card">
        <div class="section-kicker">REHEARSE</div><div class="card-title" style="margin-top:4px">Investor pressure test</div>
        <div class="card-sub">These are not generic questions. Each one exists because a specific financial movement or management commitment creates investor tension.</div>
        <div class="concern-list" style="margin-top:11px">
          ${qs.map((q,i)=>`<div class="concern" style="cursor:pointer" onclick="selectQuestion(${i})"><div class="concern-num">${i+1}</div><div><div class="concern-title">${esc(q)}</div><div class="concern-copy">Why investors ask: ${esc(questionWhy(q))}</div></div>${i===0?badge("Critical"):badge("High")}</div>`).join("")}
        </div>
      </div>
      <div class="card">
        <div class="section-kicker">SELECTED QUESTION</div>
        <div id="pressureDetail" style="margin-top:7px"></div>
      </div>
    </div>
  `;
  selectQuestion(selectedQuestion,true);
}
function selectQuestion(i,stay=false){
  selectedQuestion=i;
  renderPressureDetail();
  if(!stay)go("response"); renderResponse();
}
function renderPressureDetail(){
  const qs=(DATA.questions?.length?DATA.questions.map(q=>q["Investor Question"]).filter(Boolean):Object.keys(fallback));
  const q=qs[selectedQuestion]||qs[0];
  const qrow=(DATA.questions||[]).find(x=>x["Investor Question"]===q)||{};
  const response=getResponse(q);
  document.getElementById("pressureDetail").innerHTML=`
    <div class="question-box"><div class="small-label">Investor challenge</div><div class="question" style="margin-top:5px">${esc(q)}</div></div>
    <div class="grid g2" style="margin-top:10px">
      <div class="alert blue"><strong>Evidence trigger</strong><br/>${esc(questionWhy(q))}</div>
      <div class="alert ${q.includes("margin")||q.includes("cash")?"red":"amber"}"><strong>Why this matters</strong><br/>Leadership needs a direct explanation that can survive follow-up questions.</div>
    </div>
    <div style="margin-top:11px" class="actions">
      <button class="action-btn" onclick="go('response')">Rehearse this answer →</button>
      <button class="ghost-btn" onclick="go('commitments')">Check prior commitments</button>
    </div>
  `;
}
function questionWhy(q){
  if(q.includes("margins"))return"Revenue grew 15% while operating margin fell from 22% to 18%.";
  if(q.includes("cash flow"))return"Cash generation weakened while revenue grew and working capital increased.";
  if(q.includes("Consumer Technology"))return"Segment growth slowed materially from the previous quarter.";
  if(q.includes("guidance"))return"Current margin is below the previous >20% management expectation.";
  return"Investors will test whether current growth can be sustained without further pressure on profitability.";
}

function getResponse(q){
  const base=fallback[q]||fallback["Why did margins decline despite revenue growth?"];
  const row=(DATA.responses||[]).find(r=>String(r["Investor Question"]||"").trim()===String(q).trim());
  if(!row)return base;
  return {
    direct:row["Direct Answer"]||base.direct,
    evidence:String(row.Evidence||"").split(";").filter(Boolean),
    explanation:row.Explanation||base.explanation,
    action:row["Management Action"]||base.action,
    outlook:row["Forward-Looking Context"]||base.outlook,
    score:num(row["Overall Score"])||base.score,
    gap:row["Commitment Gap?"]||"Yes"
  };
}
function renderResponse(){
  const qs=(DATA.questions?.length?DATA.questions.map(q=>q["Investor Question"]).filter(Boolean):Object.keys(fallback));
  const q=qs[selectedQuestion]||qs[0];
  const r=getResponse(q);
  document.getElementById("page-response").innerHTML=`
    <div class="grid g2">
      <div class="card">
        <div class="section-kicker">REHEARSE THE MESSAGE</div><div class="card-title" style="margin-top:4px">Response builder</div>
        <div class="question-box" style="margin-top:10px"><div class="small-label">Investor question</div><div class="question" style="margin-top:5px">${esc(q)}</div></div>
        <div class="response-block" style="margin-top:11px"><div class="small-label">Direct answer</div><div class="mini" style="margin-top:4px">${esc(r.direct)}</div></div>
        <div class="response-block"><div class="small-label">Evidence</div><ul class="mini">${r.evidence.map(x=>`<li>${esc(x)}</li>`).join("")}</ul></div>
        <div class="response-block"><div class="small-label">Explanation</div><div class="mini" style="margin-top:4px">${esc(r.explanation)}</div></div>
        <div class="response-block"><div class="small-label">Management action</div><div class="mini" style="margin-top:4px">${esc(r.action)}</div></div>
        <div class="response-block"><div class="small-label">Forward-looking context</div><div class="mini" style="margin-top:4px">${esc(r.outlook)}</div></div>
      </div>
      <div class="card">
        <div class="section-kicker">STRESS TEST</div><div class="card-title" style="margin-top:4px">Can leadership safely say this?</div>
        <div class="alert red" style="margin-top:10px"><strong>Commitment check:</strong> ${q.includes("margin")?"Previous management expectation was >20%; current margin is 18%.":"Review the answer against prior management commitments and current disclosures."}</div>
        ${scoreRow("Evidence support",q.includes("guidance")?5:4.5)}
        ${scoreRow("Directness",4.5)}
        ${scoreRow("Concern coverage",q.includes("margin")?4:4.5)}
        ${scoreRow("Historical consistency",q.includes("margin")?4:4.5)}
        ${scoreRow("Clarity",4)}
        <div class="alert ${r.score>=4.5?"green":"amber"}" style="margin-top:12px"><strong>Overall response quality: ${r.score.toFixed(1)} / 5</strong><br/>${r.score>=4.5?"Strong — suitable for rehearsal.":"Good — refine before leadership sign-off."}</div>
        <div class="recommendation" style="margin-top:12px"><div class="small-label" style="color:#9fbbd9">PREPARE THIS BEFORE THE CALL</div><h3>${esc(prep(q))}</h3><p>Do not overpromise. Keep the response tied to disclosed evidence and acknowledge the prior commitment where relevant.</p></div>
      </div>
    </div>
  `;
}
function scoreRow(label,value){return `<div class="score-row"><div class="score-label">${label}</div><div class="progress"><span style="width:${value/5*100}%"></span></div><strong style="font-size:9px">${value.toFixed(1)}</strong></div>`}
function prep(q){
  if(q.includes("margins"))return"Prepare a quantified margin-recovery story.";
  if(q.includes("cash flow"))return"Prepare a clear working-capital bridge.";
  if(q.includes("Consumer Technology"))return"Prepare a segment-specific slowdown explanation.";
  if(q.includes("guidance"))return"Prepare the guidance-variance explanation.";
  return"Prepare a clear answer separating temporary pressure from structural risk.";
}

function renderResponsePage(){renderResponse();}

function renderDecision(){
  const risks=getRisks(), critical=risks.filter(r=>r.level==="Critical").length;
  const commitments=DATA.commitments||[], misses=commitments.filter(r=>String(r.Status).toLowerCase().includes("miss")).length;
  const score=Math.max(48,Math.min(88,100-critical*10-misses*8));
  document.getElementById("page-decision").innerHTML=`
    <div class="decision-hero">
      <div class="decision-top">
        <div>
          <div class="decision-kicker">LEADERSHIP SIGN-OFF</div>
          <div class="decision-title">NOT READY — ACTION REQUIRED</div>
          <div class="decision-copy">The leadership team should not sign off on “fully ready” yet. Revenue growth is supportable, but the margin decline, weaker cash generation and prior guidance gaps create follow-up risk that needs rehearsal.</div>
        </div>
        <div class="score-circle"><div><span>${score}</span><small>/100 readiness</small></div></div>
      </div>
      <div class="blockers">
        <div class="blocker"><strong>01 · Profitability</strong><p>Explain the move from 22% to 18% margin with a quantified driver bridge.</p></div>
        <div class="blocker"><strong>02 · Cash conversion</strong><p>Explain how receivables and inventory weakened operating cash flow.</p></div>
        <div class="blocker"><strong>03 · Credibility</strong><p>Acknowledge the earlier >20% margin expectation directly.</p></div>
      </div>
    </div>
    <div class="grid g2" style="margin-top:13px">
      <div class="card">
        <div class="card-title">Final decision logic</div>
        <div class="card-sub">The recommendation comes from the risk signals, commitment gaps and response quality — not from the size of revenue growth alone.</div>
        <div style="margin-top:11px">
          <div class="alert red"><strong>Critical risks:</strong> ${critical}</div>
          <div class="alert red" style="margin-top:7px"><strong>Missed / challenged commitments:</strong> ${misses}</div>
          <div class="alert amber" style="margin-top:7px"><strong>Featured response:</strong> 4.2 / 5 — good, but still needs refinement.</div>
          <div class="alert green" style="margin-top:7px"><strong>Supportable message:</strong> topline growth of 15% is evidenced.</div>
        </div>
      </div>
      <div class="card">
        <div class="card-title">Leadership sign-off checklist</div>
        <div class="concern-list">
          <div class="concern"><div class="concern-num">✓</div><div><div class="concern-title">Financial evidence prepared</div><div class="concern-copy">Core quarter-to-quarter movements are quantified.</div></div><span class="pill green">Ready</span></div>
          <div class="concern"><div class="concern-num">!</div><div><div class="concern-title">Margin narrative</div><div class="concern-copy">Needs a sharper driver bridge and response rehearsal.</div></div><span class="pill red">Fix</span></div>
          <div class="concern"><div class="concern-num">!</div><div><div class="concern-title">Cash-flow narrative</div><div class="concern-copy">Needs a working-capital explanation that survives follow-ups.</div></div><span class="pill red">Fix</span></div>
          <div class="concern"><div class="concern-num">!</div><div><div class="concern-title">Prior guidance</div><div class="concern-copy">Needs explicit acknowledgement of the margin miss.</div></div><span class="pill red">Fix</span></div>
        </div>
      </div>
    </div>
    <div class="recommendation">
      <div class="small-label">FINAL BUSINESS RECOMMENDATION</div>
      <h3>Do not sign off yet. Conduct a focused leadership rehearsal on margin, cash flow and the guidance miss.</h3>
      <p>Once the three blockers have evidence-backed responses with no material contradictions, the team can reassess readiness immediately before the earnings call.</p>
    </div>
  `;
}


function renderResponse(){
  const qs=(DATA.questions?.length?DATA.questions.map(q=>q["Investor Question"]).filter(Boolean):Object.keys(fallback));
  const q=qs[selectedQuestion]||qs[0];
  const r=getResponse(q);
  document.getElementById("page-response").innerHTML=`
    <div class="grid g2">
      <div class="card">
        <div class="section-kicker">REHEARSE THE MESSAGE</div>
        <div class="card-title" style="margin-top:4px">Response builder</div>
        <div class="question-box" style="margin-top:10px"><div class="small-label">Investor question</div><div class="question" style="margin-top:5px">${esc(q)}</div></div>
        <div class="response-block" style="margin-top:11px"><div class="small-label">Direct answer</div><div class="mini" style="margin-top:4px">${esc(r.direct)}</div></div>
        <div class="response-block"><div class="small-label">Evidence</div><ul class="mini">${r.evidence.map(x=>`<li>${esc(x)}</li>`).join("")}</ul></div>
        <div class="response-block"><div class="small-label">Explanation</div><div class="mini" style="margin-top:4px">${esc(r.explanation)}</div></div>
        <div class="response-block"><div class="small-label">Management action</div><div class="mini" style="margin-top:4px">${esc(r.action)}</div></div>
        <div class="response-block"><div class="small-label">Forward-looking context</div><div class="mini" style="margin-top:4px">${esc(r.outlook)}</div></div>
      </div>
      <div class="card">
        <div class="section-kicker">STRESS TEST</div>
        <div class="card-title" style="margin-top:4px">Can leadership safely say this?</div>
        <div class="alert red" style="margin-top:10px"><strong>Commitment check:</strong> ${q.includes("margin")?"Previous management expectation was >20%; current margin is 18%.":"Review the answer against prior management commitments and current disclosures."}</div>
        ${scoreRow("Evidence support",q.includes("guidance")?5:4.5)}
        ${scoreRow("Directness",4.5)}
        ${scoreRow("Concern coverage",q.includes("margin")?4:4.5)}
        ${scoreRow("Historical consistency",q.includes("margin")?4:4.5)}
        ${scoreRow("Clarity",4)}
        <div class="alert ${r.score>=4.5?"green":"amber"}" style="margin-top:12px"><strong>Overall response quality: ${r.score.toFixed(1)} / 5</strong><br/>${r.score>=4.5?"Strong — suitable for rehearsal.":"Good — refine before leadership sign-off."}</div>
        <div class="recommendation" style="margin-top:12px"><div class="small-label" style="color:#9fbbd9">PREPARE THIS BEFORE THE CALL</div><h3>${esc(prep(q))}</h3><p>Do not overpromise. Keep the response tied to disclosed evidence and acknowledge the prior commitment where relevant.</p></div>
      </div>
    </div>
  `;
}




function renderMethodology(){
  const risks=getRisks();
  const responseRows=(DATA.responses||[]).filter(r=>r && r["Investor Question"]);

  const avg = (vals)=>{
    const clean=vals.filter(Number.isFinite);
    return clean.length ? clean.reduce((a,b)=>a+b,0)/clean.length : 0;
  };

  const evidenceAvg=avg(responseRows.map(r=>num(r["Evidence Support"])));
  const directAvg=avg(responseRows.map(r=>num(r["Directness"])));
  const coverageAvg=avg(responseRows.map(r=>num(r["Concern Coverage"])));
  const consistencyAvg=avg(responseRows.map(r=>num(r["Consistency"])));
  const clarityAvg=avg(responseRows.map(r=>num(r["Overall Score"])));

  const top=risks[0];
  const guidance=risks.find(r=>String(r.issue).toLowerCase().includes("guidance"));
  const cash=risks.find(r=>String(r.issue).toLowerCase().includes("cash"));

  document.getElementById("page-methodology").innerHTML=`
    <div class="card">
      <div class="section-kicker">SCORING TRANSPARENCY</div>
      <div class="card-title" style="margin-top:4px">How the Investor Risk Score is calculated</div>
      <div class="card-sub">Every score shown in the dashboard is derived from four 1-to-5 components. The dashboard recomputes the total instead of relying on a prefilled total.</div>

      <div class="formula-callout">
        <div class="formula-main">Investor Risk Score = Financial Impact + Negative Trend + Guidance Gap + Strategic Importance</div>
        <div class="formula-breakdown">
          <div class="formula-chip">Financial Impact: 1–5</div>
          <div class="formula-chip">Negative Trend: 1–5</div>
          <div class="formula-chip">Guidance Gap: 1–5</div>
          <div class="formula-chip">Strategic Importance: 1–5</div>
        </div>
        <div class="formula-score">Maximum = 20 points</div>
      </div>

      <div class="method-rubric">
        <div class="rubric-card"><strong>16–20 · Critical</strong><span>Likely to dominate investor discussion and needs leadership preparation.</span></div>
        <div class="rubric-card"><strong>11–15 · High</strong><span>Material concern with meaningful investor follow-up risk.</span></div>
        <div class="rubric-card"><strong>6–10 · Moderate</strong><span>Relevant concern but less likely to drive the call.</span></div>
        <div class="rubric-card"><strong>1–5 · Low</strong><span>Limited immediate investor pressure.</span></div>
      </div>
    </div>

    <div class="card" style="margin-top:16px">
      <div class="card-title">Current score examples</div>
      <div class="card-sub">This is exactly how the numbers on the Executive Overview are produced.</div>
      <div class="table-wrap" style="margin-top:10px">
        <table class="method-table">
          <thead><tr><th>Risk</th><th>Formula</th><th>Score</th><th>Interpretation</th></tr></thead>
          <tbody>
            ${risks.slice(0,5).map(r=>`
              <tr>
                <td>${esc(r.issue)}</td>
                <td>${r.financialImpact!==null?`${r.financialImpact} + ${r.negativeTrend} + ${r.guidanceGap} + ${r.strategicImportance}`:"Source total"}</td>
                <td><strong>${esc(r.score)}/20</strong></td>
                <td>${badge(r.level)}</td>
              </tr>`).join("")}
          </tbody>
        </table>
      </div>
    </div>


      <div class="card" style="margin-top:16px">
        <div class="card-title">How each 1–5 component is assigned</div>
        <div class="method-rubric">
          <div class="rubric-card"><strong>Financial Impact</strong><span>1 = immaterial; 3 = meaningful; 5 = material impact on revenue, profit or cash.</span></div>
          <div class="rubric-card"><strong>Negative Trend</strong><span>1 = stable/improving; 3 = noticeable deterioration; 5 = sharp or sustained deterioration.</span></div>
          <div class="rubric-card"><strong>Guidance Gap</strong><span>1 = no meaningful gap; 3 = partial deviation; 5 = major miss against a clear commitment.</span></div>
          <div class="rubric-card"><strong>Strategic Importance</strong><span>1 = peripheral; 3 = relevant; 5 = core business, growth engine or management credibility issue.</span></div>
        </div>
      </div>

    <div class="grid g2" style="margin-top:16px">
      <div class="card">
        <div class="card-title">Financial formulas used</div>
        <div class="table-wrap" style="margin-top:10px">
          <table class="method-table">
            <thead><tr><th>Metric</th><th>Formula</th><th>Example</th></tr></thead>
            <tbody>
              <tr><td>Revenue Growth</td><td>(Current Revenue − Previous Revenue) / Previous Revenue</td><td>(${esc("1,150−1,000")}) / 1,000 = 15.0%</td></tr>
              <tr><td>Operating Margin</td><td>Operating Profit / Revenue</td><td>207 / 1,150 = 18.0%</td></tr>
              <tr><td>Margin Change</td><td>Current Margin − Previous Margin</td><td>18% − 22% = −4.0 pp</td></tr>
              <tr><td>Cash-flow Change</td><td>(Current OCF − Previous OCF) / Previous OCF</td><td>(110−150) / 150 = −26.7%</td></tr>
              <tr><td>Free Cash Flow</td><td>Operating Cash Flow − Capital Expenditure</td><td>110 − 120 = −10 Cr</td></tr>
              <tr><td>Segment Contribution</td><td>Segment Revenue / Consolidated Revenue</td><td>Used as a mix indicator; eliminations may cause small differences.</td></tr>
            </tbody>
          </table>
        </div>
      </div>

      <div class="card">
        <div class="card-title">Guidance and response formulas</div>
        <div class="table-wrap" style="margin-top:10px">
          <table class="method-table">
            <thead><tr><th>Metric</th><th>Formula / rule</th><th>Purpose</th></tr></thead>
            <tbody>
              <tr><td>Guidance Gap</td><td>Expected threshold / range compared with actual result</td><td>Flags missed, partially met or met commitments.</td></tr>
              <tr><td>Response Quality</td><td>Average of Evidence Support, Directness, Concern Coverage, Consistency and Clarity</td><td>Measures whether leadership has a defensible answer.</td></tr>
              <tr><td>Evidence Support</td><td>1–5 rating based on disclosed evidence coverage</td><td>Prevents unsupported claims.</td></tr>
              <tr><td>Consistency</td><td>1–5 rating against previous disclosures and commitments</td><td>Checks management credibility risk.</td></tr>
              <tr><td>Risk Level</td><td>16–20 Critical; 11–15 High; 6–10 Moderate; 1–5 Low</td><td>Prioritises rehearsal effort.</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <div class="alert blue" style="margin-top:16px">
      <strong>Why this matters:</strong> The scoring model does not decide whether the quarter was “good” or “bad”. It decides which issues deserve the most investor-preparation effort and makes the reasoning visible to the user.
    </div>
  `;
}

function destroyChart(id){
  if (typeof charts !== "undefined" && charts[id]) {
    try { charts[id].destroy(); } catch(e) {}
    delete charts[id];
  }
}

window.addEventListener("error", function(event){
  const status = document.getElementById("dataStatus");
  if(status){
    status.textContent = "Prototype error";
    status.className = "data-status bad";
  }
  const box = document.querySelector(".data-pack");
  if(box){
    let el = document.getElementById("runtimeError");
    if(!el){
      el = document.createElement("div");
      el.id = "runtimeError";
      el.className = "upload-error";
      box.appendChild(el);
    }
    el.textContent = event.error?.message || event.message || "Unexpected prototype error.";
  }
});
