const state = {
  data: null,
  scope: "Municípios do ERJ",
  indicatorId: "",
  year: 2025,
  location: "",
  query: "",
  showZeros: false,
  sortByValue: false,
};

const el = {
  scopeSelect: document.querySelector("#scopeSelect"),
  indicatorSelect: document.querySelector("#indicatorSelect"),
  yearSelect: document.querySelector("#yearSelect"),
  locationSelect: document.querySelector("#locationSelect"),
  searchInput: document.querySelector("#searchInput"),
  showZerosInput: document.querySelector("#showZerosInput"),
  summaryGrid: document.querySelector("#summaryGrid"),
  scopeLabel: document.querySelector("#scopeLabel"),
  indicatorTitle: document.querySelector("#indicatorTitle"),
  indicatorSubtitle: document.querySelector("#indicatorSubtitle"),
  rareBadge: document.querySelector("#rareBadge"),
  rankingTitle: document.querySelector("#rankingTitle"),
  rankingChart: document.querySelector("#rankingChart"),
  trendTitle: document.querySelector("#trendTitle"),
  trendChart: document.querySelector("#trendChart"),
  capitalsComparisonPanel: document.querySelector("#capitalsComparisonPanel"),
  capitalsComparison: document.querySelector("#capitalsComparison"),
  tableTitle: document.querySelector("#tableTitle"),
  tableHead: document.querySelector("#tableHead"),
  tableBody: document.querySelector("#tableBody"),
  sortButton: document.querySelector("#sortButton"),
  notesPanel: document.querySelector("#notesPanel"),
};

const DEFAULT_LOCATION_BY_SCOPE = {
  "Municípios do ERJ": "Rio de Janeiro",
  "Capitais do Brasil": "Rio de Janeiro (RJ)",
  SCCS: "SCCS",
};

const formatter = new Intl.NumberFormat("pt-BR", {
  maximumFractionDigits: 2,
  minimumFractionDigits: 0,
});

const currencyFormatter = new Intl.NumberFormat("pt-BR", {
  style: "currency",
  currency: "BRL",
  maximumFractionDigits: 0,
});

function formatValue(value, unit = "") {
  if (value === null || value === undefined || Number.isNaN(value)) return "-";
  if (value === 0) return "-";
  if (unit.includes("R$")) return currencyFormatter.format(value);
  return formatter.format(value);
}

function currentIndicator() {
  return state.data.indicators.find((item) => item.id === state.indicatorId);
}

function uniqueScopes() {
  return Object.keys(state.data.summary.scopeCounts);
}

function indicatorsForScope(scope) {
  return state.data.indicators.filter((item) => item.scope === scope);
}

function yearsForIndicator(indicator) {
  return indicator?.years?.length ? indicator.years : state.data.years;
}

function updateYearSelect() {
  const indicator = currentIndicator();
  const years = yearsForIndicator(indicator);
  if (!years.includes(state.year)) {
    state.year = years.at(-1) || state.data.years.at(-1);
  }
  updateSelect(
    el.yearSelect,
    years.map((year) => ({ value: year, label: year })),
    state.year,
  );
}

function displayIndicatorName(indicator) {
  if (indicator.sheet === "Capitais - Produção Amblatoriai") {
    return "Capitais - Produção Ambulatorial";
  }
  return indicator.sheet;
}

function selectedYearRows(indicator) {
  const yearKey = String(state.year);
  return indicator.rows
    .map((row) => ({
      location: row.location,
      value: row.values[yearKey] ?? null,
      values: row.values,
    }))
    .filter((row) => row.value !== null);
}

function hasAnyPositive(row) {
  return Object.values(row.values).some((value) => Number(value) > 0);
}

function preferredLocation(indicator, locations) {
  const defaultLocation = DEFAULT_LOCATION_BY_SCOPE[indicator?.scope];
  if (locations.includes(defaultLocation)) return defaultLocation;
  return locations.find((location) => location.toLowerCase().includes("rio de janeiro")) || "";
}

function scopeDetail(scope) {
  if (scope === "Capitais do Brasil") return "capitais comparadas";
  if (scope === "SCCS") return "unidades comparadas";
  return "municípios comparados";
}

function updateSelect(select, options, selectedValue) {
  select.innerHTML = options
    .map((option) => {
      const value = String(option.value);
      const selected = value === String(selectedValue) ? "selected" : "";
      return `<option value="${escapeHtml(value)}" ${selected}>${escapeHtml(option.label)}</option>`;
    })
    .join("");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function setupFilters() {
  updateSelect(
    el.scopeSelect,
    uniqueScopes().map((scope) => ({ value: scope, label: scope })),
    state.scope,
  );
  updateIndicatorSelect();
  updateYearSelect();
  updateLocationSelect();
}

function updateIndicatorSelect() {
  const indicators = indicatorsForScope(state.scope);
  if (!indicators.some((item) => item.id === state.indicatorId)) {
    state.indicatorId = indicators[0]?.id || "";
  }
  updateSelect(
    el.indicatorSelect,
    indicators.map((indicator) => ({ value: indicator.id, label: displayIndicatorName(indicator) })),
    state.indicatorId,
  );
}

function updateLocationSelect() {
  const indicator = currentIndicator();
  const locations = indicator?.rows.map((row) => row.location) || [];
  if (!locations.includes(state.location)) {
    const preferred = preferredLocation(indicator, locations);
    const firstPositive = indicator?.rows.find(hasAnyPositive);
    state.location = preferred || firstPositive?.location || locations[0] || "";
  }
  updateSelect(
    el.locationSelect,
    locations.map((location) => ({ value: location, label: location })),
    state.location,
  );
}

function renderSummary(indicator) {
  const scopeLocationCount = state.data.summary.locationCounts[indicator.scope] || indicator.stats.locations;
  const years = yearsForIndicator(indicator);
  const yearRows = selectedYearRows(indicator);
  const positiveThisYear = yearRows.filter((row) => row.value > 0).length;
  const zeroShare = Math.round(indicator.stats.zeroShare * 100);
  const sccsRow = indicator.rows.find((row) => row.location === "SCCS");
  const sccsValue = sccsRow?.values[String(state.year)];
  const cards = [
    {
      label: "Indicadores no painel",
      value: state.data.summary.indicatorCount,
      detail: years.length > 1 ? `${years[0]}-${years.at(-1)}` : `${years[0]}`,
    },
    {
      label: indicator.scope,
      value: scopeLocationCount,
      detail: scopeDetail(indicator.scope),
    },
    ...(indicator.scope === "SCCS"
      ? [
          {
            label: `SCCS em ${state.year}`,
            value: formatValue(sccsValue, indicator.unit),
            detail: "unidade em destaque",
          },
        ]
      : []),
    {
      label: `Com registro em ${state.year}`,
      value: positiveThisYear,
      detail: `${yearRows.length - positiveThisYear} sem registro no ano`,
    },
    ...(indicator.scope === "SCCS"
      ? []
      : [
          {
            label: "Zeros no indicador",
            value: `${zeroShare}%`,
            detail: indicator.rare ? "usar leitura por presença e ranking" : "ranking contínuo disponível",
          },
        ]),
  ];
  el.summaryGrid.innerHTML = cards
    .map(
      (card) => `
        <article class="metric-card">
          <span>${escapeHtml(card.label)}</span>
          <strong>${escapeHtml(card.value)}</strong>
          <span>${escapeHtml(card.detail)}</span>
        </article>
      `,
    )
    .join("");
}

function renderIndicatorHead(indicator) {
  el.scopeLabel.textContent = indicator.scope;
  el.indicatorTitle.textContent = indicator.title;
  el.indicatorSubtitle.textContent = `${indicator.label} · unidade: ${indicator.unit}`;
  el.rareBadge.textContent = "muitos zeros";
  el.rareBadge.classList.toggle("is-visible", indicator.rare);
}

function renderRanking(indicator) {
  const yearRows = selectedYearRows(indicator);
  const rows = yearRows
    .filter((row) => row.value > 0)
    .sort((a, b) => b.value - a.value)
    .slice(0, 10);
  const max = rows[0]?.value || 0;
  const focusRow = yearRows.find((row) => row.location === "SCCS" && row.value > 0);
  const showFocusRow = indicator.scope === "SCCS" && focusRow && !rows.some((row) => row.location === "SCCS");
  el.rankingTitle.textContent = `Top 10 em ${state.year}`;
  if (!rows.length) {
    el.rankingChart.innerHTML = `<div class="empty-state">Sem registros positivos para este ano.</div>`;
    return;
  }
  const topRows = rows
    .map((row) => {
      const width = Math.max(3, (row.value / max) * 100);
      return `
        <div class="bar-row ${row.location === "SCCS" ? "focus-row" : ""}" title="${escapeHtml(row.location)}">
          <span class="bar-label">${escapeHtml(row.location)}</span>
          <span class="bar-track"><span class="bar-fill" style="width:${width}%"></span></span>
          <span class="bar-value">${formatValue(row.value, indicator.unit)}</span>
        </div>
      `;
    })
    .join("");
  const focusMarkup = showFocusRow
    ? `
        <div class="bar-row focus-row" title="SCCS">
          <span class="bar-label">SCCS</span>
          <span class="bar-track"><span class="bar-fill" style="width:${Math.max(3, (focusRow.value / max) * 100)}%"></span></span>
          <span class="bar-value">${formatValue(focusRow.value, indicator.unit)}</span>
        </div>
      `
    : "";
  el.rankingChart.innerHTML = topRows + focusMarkup;
}

function renderTrend(indicator) {
  const row = indicator.rows.find((item) => item.location === state.location);
  el.trendTitle.textContent = state.location || "Selecione uma localidade";
  if (!row) {
    el.trendChart.innerHTML = `<div class="empty-state">Sem localidade selecionada.</div>`;
    return;
  }
  const years = yearsForIndicator(indicator);
  const values = years.map((year) => Number(row.values[String(year)] || 0));
  const max = Math.max(...values, 1);
  const min = Math.min(...values, 0);
  const width = 520;
  const height = 260;
  const pad = 34;
  const points = values.map((value, index) => {
    const x = values.length === 1 ? width / 2 : pad + (index * (width - pad * 2)) / (values.length - 1);
    const y = height - pad - ((value - min) / (max - min || 1)) * (height - pad * 2);
    return { x, y, value, year: years[index] };
  });
  const path = points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ");
  const area = points.length > 1 ? `${path} L ${points.at(-1).x} ${height - pad} L ${points[0].x} ${height - pad} Z` : "";
  el.trendChart.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Evolução de ${escapeHtml(state.location)}">
      <line x1="${pad}" y1="${height - pad}" x2="${width - pad}" y2="${height - pad}" stroke="#d8e0e2" />
      <line x1="${pad}" y1="${pad}" x2="${pad}" y2="${height - pad}" stroke="#d8e0e2" />
      ${area ? `<path d="${area}" fill="rgba(31, 122, 109, 0.12)"></path>` : ""}
      <path d="${path}" fill="none" stroke="#1f7a6d" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"></path>
      ${points
        .map(
          (point) => `
          <circle cx="${point.x}" cy="${point.y}" r="5" fill="#173452"></circle>
          <text x="${point.x}" y="${height - 10}" text-anchor="middle" class="axis-label">${point.year}</text>
          <text x="${point.x}" y="${Math.max(16, point.y - 12)}" text-anchor="middle" class="axis-label">${formatValue(point.value, indicator.unit)}</text>
        `,
        )
        .join("")}
    </svg>
  `;
}

function heatClass(value, max) {
  if (!value) return "zero";
  const ratio = value / max;
  if (ratio > 0.75) return "heat-4";
  if (ratio > 0.5) return "heat-3";
  if (ratio > 0.25) return "heat-2";
  return "heat-1";
}

function filteredRows(indicator) {
  const query = state.query.trim().toLowerCase();
  let rows = indicator.rows.filter((row) => {
    if (!state.showZeros && !hasAnyPositive(row)) return false;
    if (!query) return true;
    return row.location.toLowerCase().includes(query);
  });
  if (state.sortByValue) {
    const yearKey = String(state.year);
    rows = [...rows].sort((a, b) => (b.values[yearKey] || 0) - (a.values[yearKey] || 0));
  }
  return rows;
}

function renderTable(indicator) {
  const rows = filteredRows(indicator);
  const years = yearsForIndicator(indicator);
  const allValues = indicator.rows.flatMap((row) => Object.values(row.values).map(Number));
  const max = Math.max(...allValues, 1);
  el.tableTitle.textContent = `${rows.length} linhas exibidas`;
  el.tableHead.innerHTML = `
    <tr>
      <th>${indicator.scope === "Capitais do Brasil" ? "Capital" : indicator.scope === "SCCS" ? "Unidade" : "Município"}</th>
      ${years.map((year) => `<th>${year}</th>`).join("")}
    </tr>
  `;
  el.tableBody.innerHTML = rows
    .map(
      (row) => `
      <tr class="${row.location === "SCCS" ? "is-focus" : ""}">
        <td>${escapeHtml(row.location)}</td>
        ${years
          .map((year) => {
            const value = Number(row.values[String(year)] || 0);
            return `<td class="${heatClass(value, max)}">${formatValue(value, indicator.unit)}</td>`;
          })
          .join("")}
      </tr>
    `,
    )
    .join("");
}

function renderNotes(indicator) {
  const notes = indicator.notes.length ? indicator.notes : ["Sem notas registradas na aba."];
  el.notesPanel.innerHTML = `
    <strong>Fonte e metodologia</strong>
    ${notes.map((note) => `<p>${escapeHtml(note)}</p>`).join("")}
  `;
}

function renderCapitalsComparison() {
  const amb = state.data.indicators.find((item) => item.sheet.startsWith("Capitais - Produção"));
  const aih = state.data.indicators.find((item) => item.sheet === "Capitais - AIH");
  const show = state.scope === "Capitais do Brasil" && amb && aih;
  el.capitalsComparisonPanel.hidden = !show;
  if (!show) return;
  const locations = amb.rows.map((row) => row.location);
  el.capitalsComparison.innerHTML = locations
    .slice(0, 9)
    .map((location) => {
      const ambRow = amb.rows.find((row) => row.location === location);
      const aihRow = aih.rows.find((row) => row.location === location);
      return `
        <div class="comparison-item">
          <strong>${escapeHtml(location)}</strong>
          <div class="comparison-values">
            <span>Ambulatorial 2025: ${formatValue(ambRow?.values["2025"], amb.unit)}</span>
            <span>AIH 2025: ${formatValue(aihRow?.values["2025"], aih.unit)}</span>
          </div>
        </div>
      `;
    })
    .join("");
}

function render() {
  const indicator = currentIndicator();
  if (!indicator) return;
  updateIndicatorSelect();
  updateYearSelect();
  updateLocationSelect();
  renderSummary(indicator);
  renderIndicatorHead(indicator);
  renderRanking(indicator);
  renderTrend(indicator);
  renderCapitalsComparison();
  renderTable(indicator);
  renderNotes(indicator);
}

function bindEvents() {
  el.scopeSelect.addEventListener("change", (event) => {
    state.scope = event.target.value;
    state.location = "";
    state.sortByValue = false;
    updateIndicatorSelect();
    updateYearSelect();
    updateLocationSelect();
    render();
  });
  el.indicatorSelect.addEventListener("change", (event) => {
    state.indicatorId = event.target.value;
    state.location = "";
    state.sortByValue = false;
    updateYearSelect();
    updateLocationSelect();
    render();
  });
  el.yearSelect.addEventListener("change", (event) => {
    state.year = Number(event.target.value);
    render();
  });
  el.locationSelect.addEventListener("change", (event) => {
    state.location = event.target.value;
    render();
  });
  el.searchInput.addEventListener("input", (event) => {
    state.query = event.target.value;
    renderTable(currentIndicator());
  });
  el.showZerosInput.addEventListener("change", (event) => {
    state.showZeros = event.target.checked;
    renderTable(currentIndicator());
  });
  el.sortButton.addEventListener("click", () => {
    state.sortByValue = !state.sortByValue;
    el.sortButton.textContent = state.sortByValue ? "Voltar ordem original" : "Ordenar por valor";
    renderTable(currentIndicator());
  });
}

async function boot() {
  const response = await fetch("./data.json");
  state.data = await response.json();
  state.scope = uniqueScopes().includes("Municípios do ERJ") ? "Municípios do ERJ" : uniqueScopes()[0];
  state.year = state.data.years.at(-1);
  state.indicatorId = indicatorsForScope(state.scope)[0]?.id || "";
  setupFilters();
  bindEvents();
  render();
}

boot().catch((error) => {
  document.body.innerHTML = `<main class="empty-state">Erro ao carregar o painel: ${escapeHtml(error.message)}</main>`;
});
