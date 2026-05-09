
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
const SENIORITY_ORDER = ["Intern", "Junior", "Mid", "Lead", "Senior", "Executive"];

function _seniorityRank(level) {
  const index = SENIORITY_ORDER.indexOf(level);
  return index === -1 ? Number.POSITIVE_INFINITY : index;
}

function _orderedSeniorityLevels(rows, key = "seniority_level") {
  const available = new Set(rows.map(d => d[key]).filter(Boolean));
  const ordered = SENIORITY_ORDER.filter(level => available.has(level));
  const extra = Array.from(available)
    .filter(level => !SENIORITY_ORDER.includes(level))
    .sort((a, b) => String(a).localeCompare(String(b)));
  return [...ordered, ...extra];
}

function _filterPanel(industryFilter, seniorityFilter, sizeFilter, riskFilter, adoptionFilter, htl) {
  const row = htl.html`<div class="filter-row"></div>`;
  row.classList.add("filter-row--global");

  const filters = [industryFilter, seniorityFilter, sizeFilter, riskFilter, adoptionFilter];
  filters.forEach(filter => {
    const originalCell = filter.closest(".observablehq");
    filter.classList.add("filter-panel__item");
    row.append(filter);
    if (originalCell && !originalCell.contains(row)) {
      originalCell.style.display = "none";
    }
  });

  requestAnimationFrame(() => {
    row.closest(".observablehq")?.classList.add("filter-panel-cell");
  });

  return row;
}

function _multiSelectFilter(label, values, htl, compare = (a, b) => String(a).localeCompare(String(b))) {
  const options = ["All", ...Array.from(new Set(values.filter(Boolean))).sort(compare)];
  const control = htl.html`<div class="multi-select-filter">
    <div class="multi-select-filter__label">${label}</div>
    <details class="multi-select-filter__dropdown">
      <summary>
        <span class="multi-select-filter__summary">All</span>
        <span class="multi-select-filter__chevron" aria-hidden="true"></span>
      </summary>
      <div class="multi-select-filter__menu">
        ${options.map(option => htl.html`<label class="multi-select-filter__option">
          <input type="checkbox" value=${option} checked=${option === "All"}>
          <span>${option}</span>
        </label>`)}
      </div>
    </details>
  </div>`;

  const details = control.querySelector("details");
  const summary = control.querySelector(".multi-select-filter__summary");
  const inputs = Array.from(control.querySelectorAll("input[type='checkbox']"));
  const allInput = inputs.find(input => input.value === "All");

  function selectedValues() {
    const checked = inputs.filter(input => input.checked).map(input => input.value);
    return checked.includes("All") || !checked.length ? ["All"] : checked;
  }

  function updateSummary() {
    const selected = selectedValues();
    if (selected.includes("All")) {
      summary.textContent = "All";
    } else if (selected.length === 1) {
      summary.textContent = selected[0];
    } else {
      summary.textContent = `${selected.length} selected`;
    }
  }

  function setValue(nextValue) {
    const selected = Array.isArray(nextValue) ? nextValue : [nextValue || "All"];
    const useAll = selected.includes("All") || !selected.length;
    inputs.forEach(input => {
      input.checked = useAll ? input.value === "All" : selected.includes(input.value);
    });
    updateSummary();
  }

  inputs.forEach(input => {
    input.addEventListener("change", () => {
      if (input === allInput && input.checked) {
        inputs.forEach(other => {
          if (other !== allInput) other.checked = false;
        });
      } else if (input !== allInput && input.checked) {
        allInput.checked = false;
      }

      if (!inputs.some(other => other.checked)) allInput.checked = true;
      updateSummary();
      control.dispatchEvent(new Event("input", { bubbles: true }));
    });
  });

  document.addEventListener("click", event => {
    if (!control.contains(event.target)) details.open = false;
  });

  Object.defineProperty(control, "value", {
    get: selectedValues,
    set: setValue,
  });

  updateSummary();
  return control;
}

function _industryFilter(data, htl) {
  return _multiSelectFilter("Industry", data.map(d => d.industry), htl);
}

function _seniorityFilter(data, htl) {
  return _multiSelectFilter(
    "Seniority",
    data.map(d => d.seniority_level),
    htl,
    (a, b) => _seniorityRank(a) - _seniorityRank(b) || String(a).localeCompare(String(b))
  );
}

function _sizeFilter(data, htl) {
  return _multiSelectFilter("Company size", data.map(d => d.company_size), htl);
}

function _riskFilter(htl) {
  return _multiSelectFilter("Displacement risk", ["Low", "Medium", "High"], htl);
}

function _adoptionFilter(data, htl) {
  return _multiSelectFilter("AI adoption stage", data.map(d => d.industry_ai_adoption_stage), htl);
}

function _yearMax(d3, data, htl) {
  const [minYear, maxDataYear] = d3.extent(data, d => d.posting_year);
  const endYear = Math.min(2025, maxDataYear);
  const speedOptions = [
    { label: "Slow", delay: 1800 },
    { label: "Normal", delay: 1000 },
    { label: "Fast", delay: 500 },
  ];

  let animationId = null;
  let lastDispatch = 0;

  const control = htl.html`<div class="year-player">
    <div class="year-player__header">
      <label for="year-player-range">Show years up to</label>
      <output>${endYear}</output>
    </div>
    <div class="year-player__controls">
      <button type="button" aria-label="Play year animation">▶</button>
      <input id="year-player-range" type="range" min=${minYear} max=${endYear} step="0.01" value=${endYear}>
      <select aria-label="Animation speed">
        ${speedOptions.map((speed, index) => htl.html`<option value=${speed.delay} selected=${index === 1}>${speed.label}</option>`)}
      </select>
    </div>
  </div>`;

  const button = control.querySelector("button");
  const range = control.querySelector("input");
  const speed = control.querySelector("select");
  const output = control.querySelector("output");
  speed.value = String(speedOptions[1].delay);

  function formatYear(value) {
    return Number.isInteger(value) ? String(value) : value.toFixed(1);
  }

  function setYear(value, shouldDispatch = true) {
    const nextValue = Math.min(endYear, Math.max(minYear, value));
    range.value = nextValue.toFixed(2);
    output.value = formatYear(nextValue);
    if (shouldDispatch) control.dispatchEvent(new Event("input", { bubbles: true }));
  }

  function stop() {
    if (animationId) cancelAnimationFrame(animationId);
    animationId = null;
    button.textContent = "▶";
    button.setAttribute("aria-label", "Play year animation");
  }

  function play(reset = true) {
    if (reset) setYear(minYear);

    const startYear = Number(range.value);
    const startedAt = performance.now();
    const duration = Math.max(1, endYear - startYear) * Number(speed.value);

    button.textContent = "Ⅱ";
    button.setAttribute("aria-label", "Pause year animation");

    function animate(now) {
      const progress = Math.min(1, (now - startedAt) / duration);
      const finalFrame = progress >= 1;
      const nextYear = startYear + (endYear - startYear) * progress;
      const shouldDispatch = finalFrame || now - lastDispatch > 80;

      setYear(nextYear, shouldDispatch);
      if (shouldDispatch) lastDispatch = now;

      if (finalFrame) {
        stop();
        return;
      }

      animationId = requestAnimationFrame(animate);
    }

    lastDispatch = 0;
    animationId = requestAnimationFrame(animate);
  }

  button.addEventListener("click", () => {
    if (animationId) {
      stop();
    } else {
      play(true);
    }
  });

  range.addEventListener("input", () => {
    setYear(Number(range.value));
    if (Number(range.value) >= endYear) stop();
  });

  speed.addEventListener("change", () => {
    if (!animationId) return;
    stop();
    play(false);
  });

  Object.defineProperty(control, "value", {
    get: () => Number(range.value),
    set: value => {
      setYear(value);
    },
  });

  return control;
}

// ── 4. Filtered dataset ───────────────────────────────────────────────────────
function _filtered(
  data, industryFilter, seniorityFilter, sizeFilter,
  riskFilter, adoptionFilter, yearMax
) {
  const wholeYear = Math.floor(yearMax);
  const nextYearShare = yearMax - wholeYear;

  function yearIsIncluded(d, index) {
    if (d.posting_year <= wholeYear) return true;
    if (d.posting_year !== wholeYear + 1 || nextYearShare <= 0) return false;
    return (((index + 1) * 2654435761) >>> 0) / 4294967296 < nextYearShare;
  }

  function matches(selection, value) {
    const selected = Array.isArray(selection) ? selection : [selection];
    return selected.includes("All") || selected.includes(value);
  }

  return data.filter((d, index) =>
    matches(industryFilter, d.industry) &&
    matches(seniorityFilter, d.seniority_level) &&
    matches(sizeFilter, d.company_size) &&
    matches(riskFilter, d.ai_job_displacement_risk) &&
    matches(adoptionFilter, d.industry_ai_adoption_stage) &&
    yearIsIncluded(d, index)
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
      <div style="background:#fff;border:1px solid #8ddd6f;border-top:3px solid ${c.color};border-radius:10px;padding:14px 16px;text-align:center">
        <div style="font-size:10px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.6px;margin-bottom:6px">${c.label}</div>
        <div style="font-size:22px;font-weight:800;color:${c.color}">${c.value}</div>
      </div>`
    )}
  </div>`;
}

// ── 6. Guided story navigation ────────────────────────────────────────────────
const STORY_STEPS = [
  {
    tab: "📊 Overview",
    chart: "AI mention rate over time",
    title: "Start with the adoption signal",
    summary: "Track whether AI language is becoming more common in job postings for the selected filters.",
  },
  {
    tab: "🤖 AI Adoption",
    chart: "AI intensity score by industry",
    title: "Find where AI work is concentrated",
    summary: "Compare industries by the average AI intensity score to see where AI skills and tasks are more central.",
  },
  {
    tab: "🤖 AI Adoption",
    chart: "AI intensity by seniority (box plot)",
    title: "Move through the career ladder",
    summary: "Compare AI intensity from Intern through Executive, showing only seniority levels present in the filtered data.",
  },
  {
    tab: "💰 Salary",
    chart: "Year-over-year salary change by industry",
    title: "Then check compensation movement",
    summary: "See whether salaries are rising or falling by industry as the selected market changes over time.",
  },
  {
    tab: "💰 Salary",
    chart: "Country AI index vs. salary",
    title: "Connect national AI readiness to pay",
    summary: "Compare country AI index scores against advertised salaries to inspect the broader market context.",
  },
  {
    tab: "💰 Salary",
    chart: "Salary vs. automation risk (scatter)",
    title: "Look for the salary-risk relationship",
    summary: "Inspect whether higher automation risk appears alongside higher or lower pay in sampled postings.",
  },
  {
    tab: "💰 Salary",
    chart: "Avg salary — seniority × company size",
    title: "Break salary down by role level and company scale",
    summary: "Compare compensation across seniority levels and company sizes for the selected filters.",
  },
  {
    tab: "⚠️ Risk",
    chart: "Displacement risk breakdown by industry",
    title: "Shift from counts to risk shares",
    summary: "Use percentages to compare the low, medium, and high displacement-risk mix across industries.",
  },
  {
    tab: "⚠️ Risk",
    chart: "Automation risk heatmap",
    title: "Locate risk across industry and seniority",
    summary: "Read the heatmap from Intern to Executive to see where average automation risk concentrates.",
  },
  {
    tab: "🎓 Reskilling",
    chart: "% requiring reskilling by adoption stage",
    title: "Translate adoption into reskilling demand",
    summary: "Compare how often postings require reskilling at each AI adoption stage.",
  },
  {
    tab: "🔗 Correlations",
    chart: "Correlation matrix",
    title: "Close with the relationships between metrics",
    summary: "Scan the major numeric relationships to see which measures tend to move together.",
  },
];

function _storyStep(htl) {
  let index = 0;
  const control = htl.html`<section class="story-step">
    <div class="story-step__meta">
      <span>Step <output class="story-step__current">1</output> of ${STORY_STEPS.length}</span>
      <span class="story-step__section">${STORY_STEPS[0].tab}</span>
    </div>
    <h2>${STORY_STEPS[0].title}</h2>
    <p>${STORY_STEPS[0].summary}</p>
    <div class="story-step__controls">
      <button type="button" class="story-step__prev">Previous</button>
      <div class="story-step__dots">
        ${STORY_STEPS.map((step, stepIndex) => htl.html`<button type="button" class=${stepIndex === 0 ? "is-active" : ""} aria-label=${`Go to step ${stepIndex + 1}: ${step.title}`}></button>`)}
      </div>
      <button type="button" class="story-step__next">Next</button>
    </div>
  </section>`;

  const current = control.querySelector(".story-step__current");
  const section = control.querySelector(".story-step__section");
  const title = control.querySelector("h2");
  const summary = control.querySelector("p");
  const prev = control.querySelector(".story-step__prev");
  const next = control.querySelector(".story-step__next");
  const dots = Array.from(control.querySelectorAll(".story-step__dots button"));

  function setStep(nextIndex, shouldDispatch = true) {
    index = Math.min(STORY_STEPS.length - 1, Math.max(0, nextIndex));
    const step = STORY_STEPS[index];
    current.textContent = String(index + 1);
    section.textContent = step.tab;
    title.textContent = step.title;
    summary.textContent = step.summary;
    prev.disabled = index === 0;
    next.disabled = index === STORY_STEPS.length - 1;
    dots.forEach((dot, dotIndex) => {
      dot.classList.toggle("is-active", dotIndex === index);
      dot.setAttribute("aria-current", dotIndex === index ? "step" : "false");
    });
    if (shouldDispatch) control.dispatchEvent(new Event("input", { bubbles: true }));
  }

  prev.addEventListener("click", () => setStep(index - 1));
  next.addEventListener("click", () => setStep(index + 1));
  dots.forEach((dot, dotIndex) => {
    dot.addEventListener("click", () => setStep(dotIndex));
  });

  Object.defineProperty(control, "value", {
    get: () => index,
    set: value => setStep(value, false),
  });

  setStep(0, false);
  return control;
}

function _activeTab(storyStep) {
  return STORY_STEPS[storyStep]?.tab || STORY_STEPS[0].tab;
}

function _tabVariable(storyStep) {
  return STORY_STEPS[storyStep]?.chart || STORY_STEPS[0].chart;
}

// ── 7. Chart explanation ──────────────────────────────────────────────────────
function _chartExplanation(tabVariable, htl) {
  const aiIntensityCharts = new Set([
    "AI intensity score by industry",
    "AI intensity by seniority (box plot)",
  ]);
  const explanations = {
    "AI mention rate over time":
      "Shows the share of job postings that mention AI each year. Use it to see whether AI language is becoming more common in the filtered job market.",
    "Year-over-year salary change by industry":
      "Compares average annual salary growth across industries. Lines above zero indicate rising salaries; lines below zero indicate declines for the selected filters.",
    "Automation risk heatmap":
      "Highlights where average automation risk is concentrated across industries and seniority levels. Darker/high-risk cells point to groups more exposed to automation pressure.",
    "AI intensity score by industry":
      "Ranks industries by their average AI intensity score. Higher bars indicate roles where AI-related skills, tasks, or exposure are more prominent.",
    "AI intensity by seniority (box plot)":
      "Compares the distribution of AI intensity across seniority levels. Wider boxes and longer whiskers mean more variation among jobs at that level.",
    "Country AI index vs. salary":
      "Plots salaries against each country's AI index score. The trend line helps show whether higher national AI readiness is associated with higher advertised pay.",
    "Salary vs. automation risk (scatter)":
      "Each point is a sampled job posting, positioned by salary and automation risk. The trend line summarizes whether risk tends to rise or fall with pay.",
    "Avg salary — seniority × company size":
      "Compares average salaries within each seniority level across company sizes. Use it to spot where company scale appears to change compensation most.",
    "Displacement risk breakdown by industry":
      "Shows the percentage of jobs in each industry that fall into low, medium, or high displacement-risk groups. Longer red segments indicate a larger high-risk share.",
    "% requiring reskilling by adoption stage":
      "Shows the share of jobs requiring reskilling at each AI adoption stage. Higher bars suggest that adoption is creating more explicit retraining needs.",
    "Correlation matrix":
      "Summarizes pairwise relationships between the key numeric variables. Blue positive values move together, red negative values move in opposite directions, and values near zero are weak relationships.",
  };
  const metricExplanation = aiIntensityCharts.has(tabVariable)
    ? htl.html`<div class="chart-explanation__metric">
        <strong>AI Intensity Risk Score</strong>
        <p>
          This is a normalized 0-1 score from the dataset's <code>ai_intensity_score</code> field. The dashboard does not recalculate the raw score; it aggregates the precomputed job-level values by averaging them for the selected industry, seniority level, year range, and filters. A low score means AI appears to be a smaller part of the role's skills or tasks. A high score means AI is more central to the role, which can signal stronger demand for AI fluency and greater exposure to AI-driven workflow change.
        </p>
      </div>`
    : "";

  return htl.html`<div class="chart-explanation">
    <strong>How to read this chart</strong>
    <p>${explanations[tabVariable] || "Select a chart to see a short reading guide."}</p>
    ${metricExplanation}
  </div>`;
}

// ── 9. Main chart dispatcher ───────────────────────────────────────────────────
function _chart(activeTab, tabVariable, filtered, d3, Plot, htl, data) {
  if (!filtered.length) {
    return htl.html`<div style="padding:60px;text-align:center;color:#9ca3af;font-size:14px">
      ⚠️ No data matches the current filters. Try relaxing some selections.
    </div>`;
  }

  const [minYear, maxYear] = d3.extent(data, d => d.posting_year);
  const yearDomain = [minYear, Math.min(2025, maxYear)];
  const salaryDomain = [0, Math.ceil((d3.max(data, d => d.salary_usd) || 0) * 1.05)];
  const salaryChangeExtent = d3.extent(data, d => d.salary_change).map(v => Number.isFinite(v) ? v : 0);
  const salaryChangeDomain = (() => {
    const [minChange = 0, maxChange = 0] = salaryChangeExtent;
    const padding = Math.max(2, (maxChange - minChange) * 0.1);
    return [Math.floor(minChange - padding), Math.ceil(maxChange + padding)];
  })();
  const aiIndexDomain = [0, Math.ceil(d3.max(data, d => d.ai_index_score) || 100)];

  const plotLayout = {
    marginTop: 44,
    marginRight: 96,
    marginBottom: 56,
  };

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
        width: 900,
        height: 280,
        x: { label: "Year", tickFormat: d3.format("d"), domain: yearDomain },
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
  }

  // ── AI Adoption charts ─────────────────────────────────────────────────────
  if (activeTab === "🤖 AI Adoption") {

    if (tabVariable === "AI intensity score by industry") {
      const byIndustry = d3.rollups(
        filtered,
        v => d3.mean(v, d => d.ai_intensity_score),
        d => d.industry
      ).map(([industry, avg]) => ({ industry, avg }));
      const industryDomain = Array.from(new Set(filtered.map(d => d.industry))).sort();

      return Plot.plot({
        title: "Average AI intensity score by industry",
        width: 720, height: 320,
        marginLeft: 120,
        x: { label: "Avg AI intensity score", domain: [0, 0.5] },
        y: {label: null, domain: industryDomain},
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
      const seniorityOrder = _orderedSeniorityLevels(filtered);
      return Plot.plot({
        title: "AI intensity score by seniority level",
        width: 720, height: 340,
        x: { label: "Seniority level", domain: seniorityOrder },
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

  }

  // ── Salary charts ──────────────────────────────────────────────────────────
  if (activeTab === "💰 Salary") {

    if (tabVariable === "Year-over-year salary change by industry") {
      const byYearIndustry = d3.rollups(
        filtered,
        v => d3.mean(v, d => d.salary_change),
        d => d.industry,
        d => d.posting_year
      ).flatMap(([industry, years]) =>
        years
          .sort(([a], [b]) => d3.ascending(a, b))
          .map(([year, avg]) => ({ year, industry, avg }))
      );

      return Plot.plot({
        title: "Average year-over-year salary change by industry",
        ...plotLayout,
        width: 900, height: 420,
        marginRight: 120,
        x: { label: "Year", tickFormat: d3.format("d"), domain: yearDomain },
        y: { label: "Avg YoY salary change (%)", grid: true, domain: salaryChangeDomain },
        color: { legend: true, label: "Industry" },
        marks: [
          Plot.lineY(byYearIndustry, {
            x: "year",
            y: "avg",
            stroke: "industry",
            strokeWidth: 2,
            strokeOpacity: 0.75,
            z: "industry",
            title: d => `${d.industry} ${d.year}: ${d.avg.toFixed(1)}%`,
          }),
          Plot.dot(byYearIndustry, {
            x: "year",
            y: "avg",
            fill: "industry",
            r: 2,
            opacity: 0.7,
            title: d => `${d.industry} ${d.year}: ${d.avg.toFixed(1)}%`,
          }),
          Plot.ruleY([0]),
        ],
      });
    }

    if (tabVariable === "Country AI index vs. salary") {
      const withIndex = filtered.filter(d => d.ai_index_score && d.ai_index_score > 0);
      return Plot.plot({
        title: "Country AI index score vs. job salary",
        width: 820, height: 400,
        x: { label: "Country AI index total score", domain: aiIndexDomain },
        y: { label: "Salary (USD)", domain: salaryDomain },
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

    if (tabVariable === "Salary vs. automation risk (scatter)") {
      const sample = filtered.filter((_, i) => i % 3 === 0);
      return Plot.plot({
        title: "Salary vs. automation risk score",
        width: 820, height: 400,
        x: { label: "Automation risk score (0–1)", domain: [0, 1] },
        y: { label: "Salary (USD)", domain: salaryDomain },
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

      const bySenioritySize = d3.rollups(
        filtered,
        v => d3.mean(v, d => d.salary_usd),
        d => d.seniority_level,
        d => d.company_size
      ).flatMap(([seniority, sizes]) =>
        sizes.map(([company_size, avg_salary]) => ({ seniority, company_size, avg_salary }))
      );
      const seniorityOrder = _orderedSeniorityLevels(bySenioritySize, "seniority");

      return Plot.plot({
        title: "Average salary — seniority × company size",
        width: 820, height: 420,
        marginLeft: 100,
        x: { label: "Avg salary (USD)", domain: salaryDomain },
        y: { label: null, domain: sizeOrder },
        color: { legend: true, label: "Company size" },
        facet: { data: bySenioritySize, y: "seniority", label: null },
        fy: { domain: seniorityOrder },
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
        title: "Displacement risk breakdown by industry",
        width: 800, height: 300,
        marginLeft: 110,
        x: { label: "% of jobs", domain: [0, 100], grid: true },
        y: { label: null },
        color: {
          domain: ["Low", "Medium", "High"],
          range: ["#4caf50", "#ff9800", "#e53935"],
          legend: true
        },
        marks: [
          Plot.barX(grouped, {
            x: "pct", y: "industry", fill: "risk",
            title: d => `${d.industry} — ${d.risk}: ${d.count} jobs (${d.pct.toFixed(1)}%)`
          })
        ]
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
      const seniorityOrder = _orderedSeniorityLevels(heatData, "seniority");
      const orderedHeatData = heatData.sort(
        (a, b) => _seniorityRank(a.seniority) - _seniorityRank(b.seniority) ||
          String(a.industry).localeCompare(String(b.industry))
      );

      return Plot.plot({
        title: "Automation risk heatmap — industry × seniority",
        width: 720, height: 320,
        marginLeft: 120,
        x: { type: "band", label: "Seniority level", domain: seniorityOrder },
        color: {scheme: "RdYlGn", reverse: true,domain: [-1, 1],  legend: true, label: "Avg automation risk"},
        marks: [
          Plot.cell(orderedHeatData, {
            x: "seniority", y: "industry", fill: "avg", inset: 0.5,
            title: d => `${d.industry} / ${d.seniority}: ${d.avg.toFixed(2)}`,
          }),
          Plot.text(orderedHeatData, {
            x: "seniority", y: "industry",
            text: d => d.avg.toFixed(2), fontSize: 11,
            fill: d => d.avg > 0.6 ? "white" : "black",
          }),
        ],
      });
    }
  }

  // ── Reskilling chart ───────────────────────────────────────────────────────
  if (activeTab === "🎓 Reskilling") {
    if (tabVariable === "% requiring reskilling by adoption stage") {

      const stageOrder = ["Emerging", "Growing", "Mature"];
      const raw = d3.rollups( filtered,v => d3.mean(v, d => d.reskilling_required ? 1 : 0) * 100,
      d => d.industry_ai_adoption_stage);
      const map = new Map(raw.map(([stage, pct]) => [stage, pct]));
      const byStage = stageOrder.map(stage => ({ stage, pct: map.get(stage) ?? 0 }));


      return Plot.plot({
        title: "% of jobs requiring reskilling by AI adoption stage",
        width: 540, height: 300,
        x: { label: "AI adoption stage", domain: stageOrder },
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
        const rows = arr.filter(d => Number.isFinite(d[ka]) && Number.isFinite(d[kb]));
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
          type: "linear",
          domain: [-1, 0, 1],
          range: ["#dc2626", "#f8fafc", "#2563eb"],
          clamp: true,
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

// ── 10. Story progress hint ───────────────────────────────────────────────────
function _hint(storyStep, htl) {
  const step = STORY_STEPS[storyStep] || STORY_STEPS[0];
  return htl.html`<div class="story-progress">
    Showing <strong>${step.chart}</strong> in <em>${step.tab}</em>
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
  main.variable(observer("viewof industryFilter")).define("viewof industryFilter", ["data", "htl"], _industryFilter);
  main.variable(observer("industryFilter")).define("industryFilter", ["Generators", "viewof industryFilter"], (G, _) => G.input(_));

  main.variable(observer("viewof seniorityFilter")).define("viewof seniorityFilter", ["data", "htl"], _seniorityFilter);
  main.variable(observer("seniorityFilter")).define("seniorityFilter", ["Generators", "viewof seniorityFilter"], (G, _) => G.input(_));

  main.variable(observer("viewof sizeFilter")).define("viewof sizeFilter", ["data", "htl"], _sizeFilter);
  main.variable(observer("sizeFilter")).define("sizeFilter", ["Generators", "viewof sizeFilter"], (G, _) => G.input(_));

  main.variable(observer("viewof riskFilter")).define("viewof riskFilter", ["htl"], _riskFilter);
  main.variable(observer("riskFilter")).define("riskFilter", ["Generators", "viewof riskFilter"], (G, _) => G.input(_));

  main.variable(observer("viewof adoptionFilter")).define("viewof adoptionFilter", ["data", "htl"], _adoptionFilter);
  main.variable(observer("adoptionFilter")).define("adoptionFilter", ["Generators", "viewof adoptionFilter"], (G, _) => G.input(_));

  main.variable(observer("viewof filters")).define(
    "viewof filters",
    ["viewof industryFilter", "viewof seniorityFilter", "viewof sizeFilter", "viewof riskFilter", "viewof adoptionFilter", "htl"],
    _filterPanel
  );

  // Filtered data
  main.variable(observer("filtered")).define(
    "filtered",
    ["data", "industryFilter", "seniorityFilter", "sizeFilter", "riskFilter", "adoptionFilter", "yearMax"],
    _filtered
  );

  // KPIs
  main.variable(observer()).define(["filtered", "d3", "htl"], _kpis);

  // Guided story stepper
  main.variable(observer("viewof storyStep")).define("viewof storyStep", ["htl"], _storyStep);
  main.variable(observer("storyStep")).define("storyStep", ["Generators", "viewof storyStep"], (G, _) => G.input(_));

  // Current story chart
  main.variable(observer("activeTab")).define("activeTab", ["storyStep"], _activeTab);
  main.variable(observer("tabVariable")).define("tabVariable", ["storyStep"], _tabVariable);

  // Chart explanation
  main.variable(observer()).define(["tabVariable", "htl"], _chartExplanation);

  // Chart output (reacts to both selectors + filtered)
  main.variable(observer()).define(
    ["activeTab", "tabVariable", "filtered", "d3", "Plot", "htl", "data"],
    _chart
  );

  main.variable(observer("viewof yearMax")).define("viewof yearMax", ["d3", "data", "htl"], _yearMax);
  main.variable(observer("yearMax")).define("yearMax", ["Generators", "viewof yearMax"], (G, _) => G.input(_));

  // Navigation hint
  main.variable(observer()).define(["storyStep", "htl"], _hint);

  return main;
}
