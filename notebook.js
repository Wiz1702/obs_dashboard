
//  AI Adoption vs Employment Rates — Interactive Observable Notebook
//  One unified deck: tab switcher + per-tab variable navigation

// ── 1. Title ──────────────────────────────────────────────────────────────────
function _1(md) {
  return md`# AI Adoption vs Employment Rates
*Use the **tab bar** below to switch between views, and the **variable selector** inside each tab to explore different cuts of the data.*`
}

// ── 2. Data load ──────────────────────────────────────────────────────────────
function _data(FileAttachment) {
  return FileAttachment("ai_jobs_merged_dataset.xlsx").xlsx().then(wb => {
    const sheet = wb.sheet(0, { headers: true });
    return sheet.map(d => ({
      ...d,
      posting_year:          +d["posting_year"],
      salary_usd:            +d["salary_usd"],
      ai_intensity_score:    +d["ai_intensity_score"],
      automation_risk_score: +d["automation_risk_score"],
      salary_change:         +d["salary_change_vs_prev_year_percent"],
      ai_index_score:        +d["ai_index_total_score"] || null,
      ai_mentioned:
        d["ai_mentioned"] === true ||
        d["ai_mentioned"] === "TRUE" ||
        d["ai_mentioned"] === "true",
      reskilling_required:
        d["reskilling_required"] === true ||
        d["reskilling_required"] === "TRUE" ||
        d["reskilling_required"] === "true",
    }));
  });
}

// ── 3. Global filter controls ─────────────────────────────────────────────────
function _filterPanel(Inputs, data, htl) {
  const container = htl.html`<div style="
    background:#f8f9ff;
    border:1px solid #dde3f5;
    border-radius:12px;
    padding:16px 20px;
    margin:12px 0;
    display:grid;
    grid-template-columns:repeat(auto-fit,minmax(170px,1fr));
    gap:10px 16px;
  ">
    <div style="grid-column:1/-1;font-size:12px;font-weight:700;color:#6b7280;text-transform:uppercase;letter-spacing:.6px;margin-bottom:4px">
      🔽 Global Filters — applied to every chart
    </div>
  </div>`;
  return container;
}

function _industryFilter(Inputs, data) {
  return Inputs.select(
    ["All", ...new Set(data.map(d => d.industry))].sort(),
    { label: "Industry", value: "All" }
  );
}

function _seniorityFilter(Inputs, data) {
  return Inputs.select(
    ["All", ...new Set(data.map(d => d.seniority_level))].sort(),
    { label: "Seniority", value: "All" }
  );
}

function _sizeFilter(Inputs, data) {
  return Inputs.select(
    ["All", ...new Set(data.map(d => d.company_size))].sort(),
    { label: "Company size", value: "All" }
  );
}

function _riskFilter(Inputs) {
  return Inputs.select(
    ["All", "Low", "Medium", "High"],
    { label: "Displacement risk", value: "All" }
  );
}

function _adoptionFilter(Inputs, data) {
  return Inputs.select(
    ["All", ...new Set(data.map(d => d.industry_ai_adoption_stage))].sort(),
    { label: "AI adoption stage", value: "All" }
  );
}

function _yearMax(Inputs, d3, data) {
  return Inputs.range(
    d3.extent(data, d => d.posting_year),
    { label: "Show years up to", step: 1, value: d3.max(data, d => d.posting_year) }
  );
}

function _aiOnly(Inputs) {
  return Inputs.toggle({ label: "AI-mentioned jobs only", value: false });
}

// ── 4. Filtered dataset ───────────────────────────────────────────────────────
function _filtered(
  data, industryFilter, seniorityFilter, sizeFilter,
  riskFilter, adoptionFilter, yearMax, aiOnly
) {
  return data.filter(d =>
    (industryFilter === "All" || d.industry === industryFilter) &&
    (seniorityFilter === "All" || d.seniority_level === seniorityFilter) &&
    (sizeFilter === "All" || d.company_size === sizeFilter) &&
    (riskFilter === "All" || d.ai_job_displacement_risk === riskFilter) &&
    (adoptionFilter === "All" || d.industry_ai_adoption_stage === adoptionFilter) &&
    d.posting_year <= yearMax &&
    (!aiOnly || d.ai_mentioned === true)
  );
}

// ── 5. KPI summary strip ──────────────────────────────────────────────────────
function _kpis(filtered, d3, htl) {
  const n      = filtered.length;
  const pctAI  = (d3.mean(filtered, d => d.ai_mentioned ? 1 : 0) * 100).toFixed(1);
  const avgSal = d3.mean(filtered, d => d.salary_usd);
  const avgRisk= (d3.mean(filtered, d => d.automation_risk_score) * 100).toFixed(1);
  const pctResk= (d3.mean(filtered, d => d.reskilling_required ? 1 : 0) * 100).toFixed(0);

  const cards = [
    { label: "Total jobs",          value: n.toLocaleString(),           color: "#3266ad" },
    { label: "AI-mentioned",        value: pctAI + "%",                   color: "#059669" },
    { label: "Avg salary",          value: "$" + Math.round(avgSal/1000) + "K", color: "#d97706" },
    { label: "Avg automation risk", value: avgRisk + "%",                 color: "#dc2626" },
    { label: "Reskilling required", value: pctResk + "%",                 color: "#7c3aed" },
  ];

  return htl.html`<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin:14px 0">
    ${cards.map(c => htl.html`
      <div style="background:#fff;border:1px solid #e5e7f0;border-top:3px solid ${c.color};border-radius:10px;padding:14px 16px;text-align:center">
        <div style="font-size:10px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.6px;margin-bottom:6px">${c.label}</div>
        <div style="font-size:22px;font-weight:800;color:${c.color}">${c.value}</div>
      </div>`
    )}
  </div>`;
}

// ── 6. Tab switcher ───────────────────────────────────────────────────────────
function _activeTab(Inputs) {
  return Inputs.radio(
    ["📊 Overview", "🤖 AI Adoption", "💰 Salary", "⚠️ Risk", "🎓 Reskilling", "🔗 Correlations"],
    {
      label: "View",
      value: "📊 Overview",
    }
  );
}

// ── 7. Per-tab variable selector ──────────────────────────────────────────────
function _tabVariable(Inputs, activeTab) {
  const options = {
    "📊 Overview": [
      "AI mention rate over time",
      "Year-over-year salary change by industry",
      "Automation risk heatmap",
    ],
    "🤖 AI Adoption": [
      "AI intensity score by industry",
      "AI intensity by seniority (box plot)",
      "Country AI index vs. salary",
    ],
    "💰 Salary": [
      "Salary vs. automation risk (scatter)",
      "Avg salary — seniority × company size",
    ],
    "⚠️ Risk": [
      "Displacement risk breakdown by industry",
    ],
    "🎓 Reskilling": [
      "% requiring reskilling by adoption stage",
    ],
    "🔗 Correlations": [
      "Correlation matrix",
    ],
  };

  const choices = options[activeTab] || [];
  return Inputs.select(choices, { label: "Chart", value: choices[0] });
}

// ── 8. Main chart dispatcher ───────────────────────────────────────────────────
function _chart(activeTab, tabVariable, filtered, d3, Plot, htl) {
  if (!filtered.length) {
    return htl.html`<div style="padding:60px;text-align:center;color:#9ca3af;font-size:14px">
      ⚠️ No data matches the current filters. Try relaxing some selections.
    </div>`;
  }

  // ── Overview charts ────────────────────────────────────────────────────────
  if (activeTab === "📊 Overview") {

    if (tabVariable === "AI mention rate over time") {
      const byYear = d3.rollups(
        filtered,
        v => d3.mean(v, d => d.ai_mentioned ? 1 : 0) * 100,
        d => d.posting_year
      ).map(([year, pct]) => ({ year, pct })).sort((a, b) => a.year - b.year);

      return Plot.plot({
        title: "AI mention rate in job postings over time",
        width: 900, height: 280,
        x: { label: "Year", tickFormat: d3.format("d") },
        y: { label: "% of postings mentioning AI", domain: [0, 100] },
        marks: [
          Plot.areaY(byYear, { x: "year", y: "pct", fill: "#3266ad", fillOpacity: 0.1, curve: "monotone-x" }),
          Plot.lineY(byYear, { x: "year", y: "pct", stroke: "#3266ad", strokeWidth: 2.5, curve: "monotone-x" }),
          Plot.dot(byYear, { x: "year", y: "pct", fill: "#3266ad", r: 5, title: d => `${d.year}: ${d.pct.toFixed(1)}%` }),
          Plot.text(byYear, { x: "year", y: "pct", text: d => d.pct.toFixed(1) + "%", dy: -14, fontSize: 11, fill: "#3266ad" }),
          Plot.ruleY([0], { stroke: "#e5e7eb" }),
        ],
      });
    }

    if (tabVariable === "Year-over-year salary change by industry") {
      const byYearIndustry = d3.rollups(
        filtered,
        v => d3.mean(v, d => d.salary_change),
        d => d.posting_year,
        d => d.industry
      ).flatMap(([year, inds]) => inds.map(([industry, avg]) => ({ year, industry, avg })));

      return Plot.plot({
        title: "Year-over-year salary change % by industry",
        width: 900, height: 360,
        x: { label: "Year", tickFormat: d3.format("d") },
        y: { label: "Avg YoY salary change (%)" },
        color: { legend: true, label: "Industry" },
        marks: [
          Plot.lineY(byYearIndustry, {
            x: "year", y: "avg", stroke: "industry",
            strokeWidth: 2, curve: "monotone-x",
            title: d => `${d.industry} ${d.year}: ${d.avg.toFixed(1)}%`,
          }),
          Plot.ruleY([0], { stroke: "#ccc", strokeDasharray: "4,3" }),
        ],
      });
    }

    if (tabVariable === "Automation risk heatmap") {
      const heatData = d3.rollups(
        filtered,
        v => d3.mean(v, d => d.automation_risk_score),
        d => d.industry,
        d => d.seniority_level
      ).flatMap(([industry, seniorities]) =>
        seniorities.map(([seniority, avg]) => ({ industry, seniority, avg }))
      );

      return Plot.plot({
        title: "Automation risk heatmap — industry × seniority",
        width: 720, height: 320,
        marginLeft: 120,
        color: { scheme: "RdYlGn", reverse: true, legend: true, label: "Avg automation risk" },
        marks: [
          Plot.cell(heatData, {
            x: "seniority", y: "industry", fill: "avg", inset: 0.5,
            title: d => `${d.industry} / ${d.seniority}: ${d.avg.toFixed(2)}`,
          }),
          Plot.text(heatData, {
            x: "seniority", y: "industry",
            text: d => d.avg.toFixed(2), fontSize: 11,
            fill: d => d.avg > 0.6 ? "white" : "black",
          }),
        ],
      });
    }
  }

  // ── AI Adoption charts ─────────────────────────────────────────────────────
  if (activeTab === "🤖 AI Adoption") {

    if (tabVariable === "AI intensity score by industry") {
      const byIndustry = d3.rollups(
        filtered,
        v => d3.mean(v, d => d.ai_intensity_score),
        d => d.industry
      ).map(([industry, avg]) => ({ industry, avg }));

      return Plot.plot({
        title: "Average AI intensity score by industry",
        width: 720, height: 320,
        marginLeft: 120,
        x: { label: "Avg AI intensity score", domain: [0, 1] },
        y: { label: null },
        color: { scheme: "blues" },
        marks: [
          Plot.barX(byIndustry, {
            x: "avg", y: "industry", fill: "avg",
            sort: { y: "-x" },
            title: d => `${d.industry}: ${d.avg.toFixed(3)}`,
          }),
          Plot.text(byIndustry, {
            x: "avg", y: "industry",
            text: d => d.avg.toFixed(3),
            dx: 6, textAnchor: "start", fontSize: 11, fill: "#555",
          }),
          Plot.ruleX([0]),
        ],
      });
    }

    if (tabVariable === "AI intensity by seniority (box plot)") {
      const order = ["Intern", "Junior", "Mid", "Lead", "Senior", "Executive"];
      return Plot.plot({
        title: "AI intensity score by seniority level",
        width: 720, height: 340,
        x: { label: "Seniority level", domain: order },
        y: { label: "AI intensity score", domain: [0, 1] },
        color: { scheme: "blues" },
        marks: [
          Plot.boxY(filtered, {
            x: "seniority_level", y: "ai_intensity_score",
            fill: "seniority_level",
          }),
        ],
      });
    }

    if (tabVariable === "Country AI index vs. salary") {
      const withIndex = filtered.filter(d => d.ai_index_score && d.ai_index_score > 0);
      return Plot.plot({
        title: "Country AI index score vs. job salary",
        width: 820, height: 400,
        x: { label: "Country AI index total score" },
        y: { label: "Salary (USD)" },
        color: { legend: true, label: "Income group" },
        marks: [
          Plot.dot(withIndex, {
            x: "ai_index_score", y: "salary_usd",
            fill: "ai_index_income_group", r: 3, fillOpacity: 0.55,
            title: d => `${d.country}\nAI Index: ${d.ai_index_score}\n$${d.salary_usd.toLocaleString()}`,
          }),
          Plot.linearRegressionY(withIndex, {
            x: "ai_index_score", y: "salary_usd",
            stroke: "#c00", strokeWidth: 1.5, strokeDasharray: "5,3",
          }),
        ],
      });
    }
  }

  // ── Salary charts ──────────────────────────────────────────────────────────
  if (activeTab === "💰 Salary") {

    if (tabVariable === "Salary vs. automation risk (scatter)") {
      const sample = filtered.filter((_, i) => i % 3 === 0);
      return Plot.plot({
        title: "Salary vs. automation risk score",
        width: 820, height: 400,
        x: { label: "Automation risk score (0–1)" },
        y: { label: "Salary (USD)" },
        color: { legend: true, label: "Displacement risk" },
        marks: [
          Plot.dot(sample, {
            x: "automation_risk_score", y: "salary_usd",
            fill: "ai_job_displacement_risk", r: 3, fillOpacity: 0.55,
            title: d => `${d.job_title}\n$${d.salary_usd.toLocaleString()}\nRisk: ${d.ai_job_displacement_risk}`,
          }),
          Plot.linearRegressionY(sample, {
            x: "automation_risk_score", y: "salary_usd",
            stroke: "#c00", strokeWidth: 1.5, strokeDasharray: "5,3",
          }),
        ],
      });
    }

    if (tabVariable === "Avg salary — seniority × company size") {
      const sizeOrder = ["Startup", "Small", "Medium", "Large", "Enterprise"];
      const senOrder  = ["Intern", "Junior", "Mid", "Lead", "Senior", "Executive"];

      const bySenioritySize = d3.rollups(
        filtered,
        v => d3.mean(v, d => d.salary_usd),
        d => d.seniority_level,
        d => d.company_size
      ).flatMap(([seniority, sizes]) =>
        sizes.map(([company_size, avg_salary]) => ({ seniority, company_size, avg_salary }))
      );

      return Plot.plot({
        title: "Average salary — seniority × company size",
        width: 820, height: 420,
        marginLeft: 100,
        x: { label: "Avg salary (USD)" },
        y: { label: null, domain: sizeOrder },
        color: { legend: true, label: "Company size" },
        facet: { data: bySenioritySize, y: "seniority", label: null },
        fy: { domain: senOrder },
        marks: [
          Plot.barX(bySenioritySize, {
            x: "avg_salary", y: "company_size", fill: "company_size",
            sort: { y: "-x" },
            title: d => `${d.seniority} / ${d.company_size}: $${Math.round(d.avg_salary).toLocaleString()}`,
          }),
          Plot.ruleX([0]),
        ],
      });
    }
  }

  // ── Risk chart ─────────────────────────────────────────────────────────────
  if (activeTab === "⚠️ Risk") {
    if (tabVariable === "Displacement risk breakdown by industry") {
      const grouped = d3.rollups(
        filtered,
        v => {
          const total = v.length;
          return ["Low", "Medium", "High"].map(risk => ({
            industry: v[0].industry,
            risk,
            count: v.filter(d => d.ai_job_displacement_risk === risk).length,
            pct: (v.filter(d => d.ai_job_displacement_risk === risk).length / total) * 100
          }));
        },
        d => d.industry
      ).flatMap(([, rows]) => rows);

      return Plot.plot({
        title: "Displacement risk breakdown by industry (absolute count)",
        width: 800, height: 300,
        marginLeft: 110,
        x: { label: "Number of jobs" },
        y: { label: null },
        color: {
          domain: ["Low", "Medium", "High"],
          range: ["#4caf50", "#ff9800", "#e53935"],
          legend: true
        },
        marks: [
          Plot.barX(grouped, {
            x: "count", y: "industry", fill: "risk",
            sort: { y: "-x" },
            title: d => `${d.industry} — ${d.risk}: ${d.count} jobs (${d.pct.toFixed(1)}%)`
          })
        ]
      });
    }
  }

  // ── Reskilling chart ───────────────────────────────────────────────────────
  if (activeTab === "🎓 Reskilling") {
    if (tabVariable === "% requiring reskilling by adoption stage") {
      const byStage = d3.rollups(
        filtered,
        v => d3.mean(v, d => d.reskilling_required ? 1 : 0) * 100,
        d => d.industry_ai_adoption_stage
      ).map(([stage, pct]) => ({ stage, pct }));

      return Plot.plot({
        title: "% of jobs requiring reskilling by AI adoption stage",
        width: 540, height: 300,
        x: { label: "AI adoption stage" },
        y: { label: "% requiring reskilling", domain: [0, 100] },
        color: { scheme: "purples" },
        marks: [
          Plot.barY(byStage, {
            x: "stage", y: "pct", fill: "stage",
            title: d => `${d.stage}: ${d.pct.toFixed(1)}%`,
          }),
          Plot.text(byStage, {
            x: "stage", y: "pct",
            text: d => d.pct.toFixed(1) + "%",
            dy: -10, fontSize: 12, fontWeight: "bold", fill: "#333",
          }),
          Plot.ruleY([0]),
        ],
      });
    }
  }

  // ── Correlations chart ────────────────────────────────────────────────────
  if (activeTab === "🔗 Correlations") {
    if (tabVariable === "Correlation matrix") {
      const numVars = [
        { key: "ai_intensity_score",    label: "AI intensity"    },
        { key: "automation_risk_score", label: "Automation risk" },
        { key: "salary_usd",            label: "Salary"          },
        { key: "salary_change",         label: "Salary YoY Δ"   },
        { key: "ai_index_score",        label: "AI index score"  },
      ];

      function pearson(arr, ka, kb) {
        const rows = arr.filter(d => d[ka] && d[kb]);
        const ma = d3.mean(rows, d => d[ka]);
        const mb = d3.mean(rows, d => d[kb]);
        const num = d3.sum(rows, d => (d[ka] - ma) * (d[kb] - mb));
        const den = Math.sqrt(
          d3.sum(rows, d => (d[ka] - ma) ** 2) *
          d3.sum(rows, d => (d[kb] - mb) ** 2)
        );
        return den === 0 ? 0 : num / den;
      }

      const corrData = numVars.flatMap(va =>
        numVars.map(vb => ({
          x: va.label, y: vb.label,
          r: pearson(filtered, va.key, vb.key),
        }))
      );

      return Plot.plot({
        title: "Correlation matrix — key numeric variables",
        width: 540, height: 540,
        marginLeft: 110, marginBottom: 90,
        x: { tickRotate: -35 },
        color: {
          domain: [-1, 0, 1],
          range: ["#e53935", "#fff", "#3266ad"],
          legend: true,
          label: "Pearson r",
        },
        marks: [
          Plot.cell(corrData, { x: "x", y: "y", fill: "r", inset: 0.5 }),
          Plot.text(corrData, {
            x: "x", y: "y",
            text: d => d.r.toFixed(2),
            fill: d => Math.abs(d.r) > 0.4 ? "white" : "black",
            fontSize: 12,
          }),
        ],
      });
    }
  }

  return htl.html`<div style="padding:40px;text-align:center;color:#9ca3af">Select a view above.</div>`;
}

// ── 9. Navigation hint ────────────────────────────────────────────────────────
function _hint(activeTab, tabVariable, htl) {
  const tabMap = {
    "📊 Overview":    3,
    "🤖 AI Adoption": 3,
    "💰 Salary":      2,
    "⚠️ Risk":        1,
    "🎓 Reskilling":  1,
    "🔗 Correlations":1,
  };
  const total   = tabMap[activeTab] || 1;
  const options = {
    "📊 Overview":    ["AI mention rate over time","Year-over-year salary change by industry","Automation risk heatmap"],
    "🤖 AI Adoption": ["AI intensity score by industry","AI intensity by seniority (box plot)","Country AI index vs. salary"],
    "💰 Salary":      ["Salary vs. automation risk (scatter)","Avg salary — seniority × company size"],
    "⚠️ Risk":        ["Displacement risk breakdown by industry"],
    "🎓 Reskilling":  ["% requiring reskilling by adoption stage"],
    "🔗 Correlations":["Correlation matrix"],
  };
  const idx = (options[activeTab] || []).indexOf(tabVariable) + 1;
  return htl.html`<div style="font-size:11px;color:#9ca3af;margin-top:8px;text-align:right">
    Chart ${idx} of ${total} in <em>${activeTab}</em> — use the <strong>Chart</strong> selector above to navigate
  </div>`;
}

// ── Observable module export ───────────────────────────────────────────────────
export default function define(runtime, observer) {
  const main = runtime.module();

  function toString() { return this.url; }
  const fileAttachments = new Map([
    ["ai_jobs_merged_dataset.xlsx", {
      url: new URL(
        "./files/1fb7d2c64de6035b83bb2ccbf5c62c8c86d60e8e43c09e22cbaec180b529edbde89c6f0983280bb370560c662c05cfb9ac87df52fc7c9fc2cf94429eed366cee.xlsx",
        import.meta.url
      ),
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      toString,
    }],
  ]);

  main.builtin("FileAttachment", runtime.fileAttachments(name => fileAttachments.get(name)));

  // Title
  main.variable(observer()).define(["md"], _1);

  // Data
  main.variable(observer("data")).define("data", ["FileAttachment"], _data);

  // Global filters
  main.variable(observer("viewof industryFilter")).define("viewof industryFilter", ["Inputs", "data"], _industryFilter);
  main.variable(observer("industryFilter")).define("industryFilter", ["Generators", "viewof industryFilter"], (G, _) => G.input(_));

  main.variable(observer("viewof seniorityFilter")).define("viewof seniorityFilter", ["Inputs", "data"], _seniorityFilter);
  main.variable(observer("seniorityFilter")).define("seniorityFilter", ["Generators", "viewof seniorityFilter"], (G, _) => G.input(_));

  main.variable(observer("viewof sizeFilter")).define("viewof sizeFilter", ["Inputs", "data"], _sizeFilter);
  main.variable(observer("sizeFilter")).define("sizeFilter", ["Generators", "viewof sizeFilter"], (G, _) => G.input(_));

  main.variable(observer("viewof riskFilter")).define("viewof riskFilter", ["Inputs"], _riskFilter);
  main.variable(observer("riskFilter")).define("riskFilter", ["Generators", "viewof riskFilter"], (G, _) => G.input(_));

  main.variable(observer("viewof adoptionFilter")).define("viewof adoptionFilter", ["Inputs", "data"], _adoptionFilter);
  main.variable(observer("adoptionFilter")).define("adoptionFilter", ["Generators", "viewof adoptionFilter"], (G, _) => G.input(_));

  main.variable(observer("viewof yearMax")).define("viewof yearMax", ["Inputs", "d3", "data"], _yearMax);
  main.variable(observer("yearMax")).define("yearMax", ["Generators", "viewof yearMax"], (G, _) => G.input(_));

  main.variable(observer("viewof aiOnly")).define("viewof aiOnly", ["Inputs"], _aiOnly);
  main.variable(observer("aiOnly")).define("aiOnly", ["Generators", "viewof aiOnly"], (G, _) => G.input(_));

  // Filtered data
  main.variable(observer("filtered")).define(
    "filtered",
    ["data", "industryFilter", "seniorityFilter", "sizeFilter", "riskFilter", "adoptionFilter", "yearMax", "aiOnly"],
    _filtered
  );

  // KPIs
  main.variable(observer()).define(["filtered", "d3", "htl"], _kpis);

  // Tab switcher (radio)
  main.variable(observer("viewof activeTab")).define("viewof activeTab", ["Inputs"], _activeTab);
  main.variable(observer("activeTab")).define("activeTab", ["Generators", "viewof activeTab"], (G, _) => G.input(_));

  // Per-tab variable selector (reacts to activeTab)
  main.variable(observer("viewof tabVariable")).define("viewof tabVariable", ["Inputs", "activeTab"], _tabVariable);
  main.variable(observer("tabVariable")).define("tabVariable", ["Generators", "viewof tabVariable"], (G, _) => G.input(_));

  // Chart output (reacts to both selectors + filtered)
  main.variable(observer()).define(
    ["activeTab", "tabVariable", "filtered", "d3", "Plot", "htl"],
    _chart
  );

  // Navigation hint
  main.variable(observer()).define(["activeTab", "tabVariable", "htl"], _hint);

  return main;
}
