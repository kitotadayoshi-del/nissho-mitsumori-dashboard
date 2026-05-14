/* global Chart */

const DEFAULT_USAGE_PATHS = ["./data/見積使用状況.csv", "./見積使用状況.csv"];
const DEFAULT_MASTER_PATHS = ["./data/ニッショー73支店_整形.csv", "./ニッショー73支店_整形.csv"];
const BRANCH_CHART_LIMIT = 20;
const STAFF_CHART_LIMIT = 12;

const COLORS = {
  blue: "#0057b8",
  blueSoft: "rgba(0, 87, 184, 0.22)",
  blueGrid: "rgba(0, 87, 184, 0.09)",
  orange: "#f28c28",
  orangeSoft: "rgba(242, 140, 40, 0.30)",
  ink: "#122033",
  muted: "#66748a",
};

function setStatus(text, tone = "info") {
  const el = document.getElementById("status");
  el.textContent = text;
  el.style.color = tone === "error" ? "#b42318" : tone === "ok" ? COLORS.blue : COLORS.muted;
}

function normalizeText(value) {
  return (value ?? "")
    .toString()
    .trim()
    .normalize("NFKC")
    .replace(/[　\s]+/g, "")
    .replace(/[.．･・]/g, "・")
    .replace(/^ニッショー/, "")
    .replace(/^日照/, "")
    .replace(/支店$/g, "");
}

function compactKey(value) {
  return normalizeText(value).replace(/・/g, "");
}

function splitCsvLine(line) {
  const out = [];
  let cur = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"') {
        if (line[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        cur += ch;
      }
    } else if (ch === ",") {
      out.push(cur);
      cur = "";
    } else if (ch === '"') {
      inQuotes = true;
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out.map((v) => v.trim());
}

function parseCsvRows(text) {
  const lines = text.replace(/^\uFEFF/, "").split(/\r?\n/).filter((line) => line.trim() !== "");
  if (lines.length === 0) return { header: [], rows: [] };
  const header = splitCsvLine(lines[0]);
  const rows = lines.slice(1).map((line) => {
    const cells = splitCsvLine(line);
    return header.reduce((acc, key, idx) => {
      acc[key] = cells[idx] ?? "";
      return acc;
    }, {});
  });
  return { header, rows };
}

function parseUsageRows(text) {
  const lines = text.replace(/^\uFEFF/, "").split(/\r?\n/).filter((line) => line.trim() !== "");
  if (lines.length === 0) return [];

  const headerLine = lines[0];
  const isTsv = headerLine.includes("\t");
  const header = isTsv ? headerLine.split("\t") : splitCsvLine(headerLine);
  const hasHeader = header.includes("保存日") || header.includes("支店") || header.includes("担当者");

  const indexFromHeader = () => {
    const dateIndex = header.indexOf("保存日");
    const staffIndex = header.indexOf("担当者");
    let branchIndex = header.indexOf("支店");

    // 旧形式（TSV）の「支店」列は見出しが空欄で5列目固定
    if (branchIndex === -1 && isTsv && header.length >= 5) branchIndex = 4;

    // 新形式（公開用）は「保存日・支店・担当者」の3列想定
    if (branchIndex === -1 && header.length === 3) branchIndex = 1;

    if (dateIndex === -1 || staffIndex === -1 || branchIndex === -1) {
      throw new Error("見積使用状況.csv の列が期待と異なります（保存日/支店/担当者を検出できません）。");
    }

    return { dateIndex, branchIndex, staffIndex, startRow: 1 };
  };

  const indexWithoutHeader = () => {
    const firstCells = isTsv ? headerLine.split("\t") : splitCsvLine(headerLine);

    // 公開用（保存日・支店・担当者）の3列貼り付け想定
    if (firstCells.length === 3) return { dateIndex: 0, branchIndex: 1, staffIndex: 2, startRow: 0 };

    // 旧形式っぽい行（保存日 + 3列 + 支店 + 担当者）
    if (firstCells.length >= 6) return { dateIndex: 0, branchIndex: 4, staffIndex: 5, startRow: 0 };

    throw new Error("見積使用状況.csv の形式を判定できません（ヘッダー行が必要かもしれません）。");
  };

  const { dateIndex, branchIndex, staffIndex, startRow } = hasHeader ? indexFromHeader() : indexWithoutHeader();

  return lines.slice(startRow).flatMap((line) => {
    const cells = isTsv ? line.split("\t") : splitCsvLine(line);
    const savedAt = (cells[dateIndex] ?? "").trim();
    const branchRaw = (cells[branchIndex] ?? "").trim();
    const staffName = (cells[staffIndex] ?? "").trim();
    const date = parseJpDate(savedAt);
    if (!date || !branchRaw || !staffName) return [];
    return [{ savedAt, date, branchRaw, staffName }];
  });
}

function parseJpDate(text) {
  const m = text.trim().match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
  if (!m) return null;
  const dt = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  return Number.isNaN(dt.getTime()) ? null : dt;
}

function formatYmd(date) {
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, "0");
  const d = String(date.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function buildMaster(masterCsvText) {
  const { rows } = parseCsvRows(masterCsvText);
  const masterRows = rows
    .map((row) => ({
      block: (row["課/ブロック"] ?? "").trim(),
      code: (row["支店コード"] ?? "").trim(),
      name: (row["支店名"] ?? "").trim(),
    }))
    .filter((row) => row.name !== "");

  const byKey = new Map();
  for (const row of masterRows) {
    byKey.set(normalizeText(row.name), row);
    byKey.set(compactKey(row.name), row);
  }

  return { rows: masterRows, byKey };
}

function canonicalizeBranch(branchRaw, master) {
  const normalized = normalizeText(branchRaw);
  const compact = compactKey(branchRaw);
  const masterRow = master.byKey.get(normalized) ?? master.byKey.get(compact);
  if (masterRow) {
    return { branch: masterRow.name, block: masterRow.block, isMaster: true };
  }

  return {
    branch: branchRaw.trim().normalize("NFKC").replace(/[.．･]/g, "・"),
    block: "その他",
    isMaster: false,
  };
}

function fiscalMonthKey(date, startDay = 21) {
  const y = date.getUTCFullYear();
  const m = date.getUTCMonth();
  const d = date.getUTCDate();
  const endDate = new Date(Date.UTC(y, d >= startDay ? m + 1 : m, 1));
  return `${endDate.getUTCFullYear()}-${String(endDate.getUTCMonth() + 1).padStart(2, "0")}`;
}

function formatFiscalPeriodLabel(periodKey) {
  const m = periodKey.match(/^(\d{4})-(\d{2})$/);
  if (!m) return periodKey;
  const endY = Number(m[1]);
  const endM = Number(m[2]);
  const start = new Date(Date.UTC(endY, endM - 2, 21));
  const end = new Date(Date.UTC(endY, endM - 1, 20));
  return `${endY}年${endM}月度（${start.getUTCMonth() + 1}/${start.getUTCDate()}〜${end.getUTCMonth() + 1}/${end.getUTCDate()}）`;
}

function countBy(items, keyFn) {
  const map = new Map();
  for (const item of items) {
    const key = keyFn(item);
    map.set(key, (map.get(key) ?? 0) + 1);
  }
  return map;
}

function sortRanking(a, b) {
  return b.count - a.count || a.branch.localeCompare(b.branch, "ja") || (a.staff ?? "").localeCompare(b.staff ?? "", "ja");
}

function filteredRows(usageRows, periodKey, blockFilter) {
  const periodRows = periodKey === "__ALL__"
    ? usageRows
    : usageRows.filter((row) => fiscalMonthKey(row.date) === periodKey);
  if (blockFilter === "__ALL__") return periodRows;
  return periodRows.filter((row) => row.block === blockFilter);
}

function cleanStaffName(value) {
  return (value ?? "").toString().trim().normalize("NFKC").replace(/[　\s]+/g, "");
}

function staffMatchKey(value) {
  return cleanStaffName(value)
    .replace(/髙/g, "高")
    .replace(/[斉齊齋]/g, "斎");
}

function firstNamePart(value) {
  return (value ?? "").toString().trim().normalize("NFKC").split(/[　\s]+/)[0] ?? "";
}

function staffSurnameCandidates(clean) {
  const key = staffMatchKey(clean);
  const candidates = new Set();
  if (!key) return [];
  if (key.length <= 3) candidates.add(key);
  return Array.from(candidates);
}

function buildStaffCanonicalMap(rows) {
  const byBranch = new Map();
  for (const row of rows) {
    if (!byBranch.has(row.branch)) byBranch.set(row.branch, new Map());
    const names = byBranch.get(row.branch);
    const clean = cleanStaffName(row.staffName);
    if (!clean) continue;
    const current = names.get(clean) ?? { raw: row.staffName, count: 0 };
    current.count += 1;
    names.set(clean, current);
  }

  const canonicalByBranch = new Map();
  for (const [branch, names] of byBranch) {
    const surnameSeeds = new Set();
    for (const [clean, meta] of names) {
      const first = firstNamePart(meta.raw);
      if (first && cleanStaffName(first) !== clean) surnameSeeds.add(staffMatchKey(first));
      for (const candidate of staffSurnameCandidates(clean)) surnameSeeds.add(candidate);
    }

    const groups = new Map();
    for (const [clean, meta] of names) {
      const matchKey = staffMatchKey(clean);
      const surname = Array.from(surnameSeeds)
        .filter((seed) => matchKey.startsWith(seed))
        .sort((a, b) => a.length - b.length)[0] ?? matchKey;
      if (!groups.has(surname)) groups.set(surname, []);
      groups.get(surname).push({ clean, count: meta.count });
    }

    const branchMap = new Map();
    for (const members of groups.values()) {
      const canonical = members
        .slice()
        .sort((a, b) => b.clean.length - a.clean.length || b.count - a.count || a.clean.localeCompare(b.clean, "ja"))[0].clean;
      for (const member of members) branchMap.set(member.clean, canonical);
    }
    canonicalByBranch.set(branch, branchMap);
  }

  return canonicalByBranch;
}

function buildRankings(usageRows, master, periodKey, blockFilter, branchFilter) {
  const canonicalRows = usageRows.map((row) => ({
    ...row,
    ...canonicalizeBranch(row.branchRaw, master),
  }));
  const rows = filteredRows(canonicalRows, periodKey, blockFilter);

  const branchCounts = countBy(rows, (row) => row.branch);
  const branchMeta = new Map(rows.map((row) => [row.branch, { block: row.block, isMaster: row.isMaster }]));

  const masterRows = blockFilter === "__ALL__"
    ? master.rows
    : master.rows.filter((row) => row.block === blockFilter);
  for (const row of masterRows) {
    if (!branchCounts.has(row.name)) branchCounts.set(row.name, 0);
    if (!branchMeta.has(row.name)) branchMeta.set(row.name, { block: row.block, isMaster: true });
  }

  const branchRanking = Array.from(branchCounts.entries())
    .map(([branch, count]) => ({ branch, count, ...(branchMeta.get(branch) ?? { block: "その他", isMaster: false }) }))
    .sort(sortRanking);

  const staffSource = branchFilter === "__ALL__" ? rows : rows.filter((row) => row.branch === branchFilter);
  const staffCanonicalMap = buildStaffCanonicalMap(staffSource);
  const staffCounts = countBy(staffSource, (row) => {
    const clean = cleanStaffName(row.staffName);
    const canonical = staffCanonicalMap.get(row.branch)?.get(clean) ?? clean;
    return `${row.branch}\t${canonical}`;
  });
  const staffRanking = Array.from(staffCounts.entries())
    .map(([key, count]) => {
      const [branch, staff] = key.split("\t");
      return { branch, staff, count };
    })
    .sort(sortRanking);

  return { rows, staffSource, branchRanking, staffRanking };
}

function setMetric(id, value) {
  document.getElementById(id).textContent = typeof value === "number" ? value.toLocaleString("ja-JP") : value;
}

function formatCount(value, unit) {
  return `${value.toLocaleString("ja-JP")}${unit}`;
}

function masterCountForBlock(master, blockFilter) {
  if (blockFilter === "__ALL__") return master.rows.length;
  if (blockFilter === "その他") return null;
  return master.rows.filter((row) => row.block === blockFilter).length;
}

function renderTable(tbody, rows) {
  tbody.innerHTML = "";
  const fragment = document.createDocumentFragment();
  for (const row of rows) fragment.appendChild(row);
  tbody.appendChild(fragment);
}

function branchTableRow(row, idx) {
  const tr = document.createElement("tr");
  const countClass = row.count === 0 ? "num zero" : "num";
  const division = row.isMaster ? row.block || "73支店" : "その他";
  const divisionClass = row.isMaster ? "division" : "division division--other";
  tr.innerHTML = `
    <td class="rank">${idx + 1}</td>
    <td class="branch-name">${escapeHtml(row.branch)}</td>
    <td><span class="${divisionClass}">${escapeHtml(division)}</span></td>
    <td class="${countClass}">${row.count.toLocaleString("ja-JP")}</td>
  `;
  return tr;
}

function staffTableRow(row, idx) {
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td class="rank">${idx + 1}</td>
    <td class="branch-name">${escapeHtml(row.branch)}</td>
    <td class="staff-name">${escapeHtml(row.staff)}</td>
    <td class="num">${row.count.toLocaleString("ja-JP")}</td>
  `;
  return tr;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;",
  })[ch]);
}

let branchChart = null;
let staffChart = null;

function chartOptions(color, rotateLabels = true) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: { mode: "index", intersect: false },
    plugins: {
      legend: { display: false },
      tooltip: {
        backgroundColor: color === COLORS.orange ? "#fff4e8" : "#eaf3ff",
        titleColor: COLORS.ink,
        bodyColor: COLORS.ink,
        borderColor: color,
        borderWidth: 1,
        titleFont: { size: 15, weight: "bold" },
        bodyFont: { size: 14, weight: "bold" },
        padding: 12,
        displayColors: false,
      },
    },
    scales: {
      x: {
        ticks: {
          color: COLORS.ink,
          font: { size: 12, weight: "bold" },
          maxRotation: rotateLabels ? 45 : 0,
          minRotation: rotateLabels ? 45 : 0,
          autoSkip: false,
        },
        grid: { display: false },
      },
      y: {
        beginAtZero: true,
        ticks: { color: COLORS.muted, precision: 0, font: { size: 13, weight: "bold" } },
        grid: { color: COLORS.blueGrid },
      },
    },
  };
}

function renderBarChart(canvasId, existingChart, ranking, labelFn, color, limit, rotateLabels = true) {
  const ctx = document.getElementById(canvasId);
  if (existingChart) existingChart.destroy();
  const topRows = ranking.slice(0, limit);
  return new Chart(ctx, {
    type: "bar",
    data: {
      labels: topRows.map(labelFn),
      datasets: [{
        label: "件数",
        data: topRows.map((row) => row.count),
        backgroundColor: color === COLORS.orange ? COLORS.orangeSoft : COLORS.blueSoft,
        borderColor: color,
        borderWidth: 2,
        borderRadius: 5,
        maxBarThickness: 34,
      }],
    },
    options: chartOptions(color, rotateLabels),
  });
}

function populateBranchFilter(branchRanking, selectedValue) {
  const select = document.getElementById("branchFilterSelect");
  const current = selectedValue || select.value || "__ALL__";
  select.innerHTML = "";

  const all = document.createElement("option");
  all.value = "__ALL__";
  all.textContent = "全支店・部署";
  select.appendChild(all);

  for (const row of branchRanking.filter((item) => item.count > 0)) {
    const opt = document.createElement("option");
    opt.value = row.branch;
    opt.textContent = row.branch;
    select.appendChild(opt);
  }

  select.value = Array.from(select.options).some((opt) => opt.value === current) ? current : "__ALL__";
  return select.value;
}

function populateBlockFilter(master, selectedValue) {
  const select = document.getElementById("blockFilterSelect");
  const current = selectedValue || select.value || "__ALL__";
  select.innerHTML = "";

  const all = document.createElement("option");
  all.value = "__ALL__";
  all.textContent = "全ブロック";
  select.appendChild(all);

  const blocks = [];
  for (const row of master.rows) {
    if (row.block && !blocks.includes(row.block)) blocks.push(row.block);
  }
  if (!blocks.includes("その他")) blocks.push("その他");

  for (const block of blocks) {
    const opt = document.createElement("option");
    opt.value = block;
    opt.textContent = block;
    select.appendChild(opt);
  }

  select.value = Array.from(select.options).some((opt) => opt.value === current) ? current : "__ALL__";
  return select.value;
}

function renderDashboard(state) {
  const periodKey = document.getElementById("periodSelect").value || "__ALL__";
  const blockFilter = document.getElementById("blockFilterSelect").value || "__ALL__";
  const branchFilter = document.getElementById("branchFilterSelect").value || "__ALL__";
  const activeBlockFilter = populateBlockFilter(state.master, blockFilter);
  const baseRanking = buildRankings(state.usageRows, state.master, periodKey, activeBlockFilter, "__ALL__");
  const activeBranchFilter = populateBranchFilter(baseRanking.branchRanking, branchFilter);
  const { rows, staffSource, branchRanking, staffRanking } = activeBranchFilter === "__ALL__"
    ? baseRanking
    : buildRankings(state.usageRows, state.master, periodKey, activeBlockFilter, activeBranchFilter);

  const unusedMaster = branchRanking.filter((row) => row.isMaster && row.count === 0).length;
  const usedBranches = branchRanking.filter((row) => row.count > 0).length;
  const blockMasterCount = masterCountForBlock(state.master, activeBlockFilter);
  setMetric("metricUsage", formatCount(activeBranchFilter === "__ALL__" ? rows.length : staffSource.length, "件"));
  setMetric("metricBranches", blockMasterCount ? `${usedBranches}/${blockMasterCount}` : usedBranches);
  setMetric("metricUnused", formatCount(unusedMaster, "支店"));
  setMetric("metricStaff", formatCount(staffRanking.length, "名"));

  branchChart = renderBarChart(
    "branchChart",
    branchChart,
    branchRanking,
    (row) => row.branch,
    COLORS.blue,
    BRANCH_CHART_LIMIT,
  );
  staffChart = renderBarChart(
    "staffChart",
    staffChart,
    staffRanking,
    (row) => [row.staff, row.branch],
    COLORS.orange,
    STAFF_CHART_LIMIT,
    false,
  );

  renderTable(
    document.querySelector("#branchTable tbody"),
    branchRanking.map((row, idx) => branchTableRow(row, idx)),
  );
  renderTable(
    document.querySelector("#staffTable tbody"),
    staffRanking.map((row, idx) => staffTableRow(row, idx)),
  );

  document.getElementById("branchTableNote").textContent = `${branchRanking.length.toLocaleString("ja-JP")}件`;
  document.getElementById("staffTableNote").textContent = `${staffRanking.length.toLocaleString("ja-JP")}件`;

  if (rows.length > 0) {
    const minDate = rows.reduce((a, row) => (a < row.date ? a : row.date), rows[0].date);
    const maxDate = rows.reduce((a, row) => (a > row.date ? a : row.date), rows[0].date);
    const blockText = activeBlockFilter === "__ALL__" ? "全ブロック" : activeBlockFilter;
    setStatus(`表示期間: ${formatYmd(minDate)}〜${formatYmd(maxDate)} / ${blockText} / 73支店マスタ基準`, "ok");
  } else {
    setStatus("この条件に該当するデータはありません。", "error");
  }
}

function showFileFallback(show) {
  document.getElementById("fileFallback").classList.toggle("is-hidden", !show);
}

async function fetchTextOrNull(path) {
  try {
    const res = await fetch(path, { cache: "no-store" });
    if (!res.ok) return null;
    return await res.text();
  } catch {
    return null;
  }
}

function readFileText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result?.toString() ?? "");
    reader.onerror = () => reject(new Error("ファイル読み込みに失敗しました。"));
    reader.readAsText(file, "utf-8");
  });
}

async function fetchFirstText(paths) {
  for (const path of paths) {
    const text = await fetchTextOrNull(path);
    if (text) return text;
  }
  return null;
}

function setupPeriodSelect(usageRows) {
  const periods = Array.from(new Set(usageRows.map((row) => fiscalMonthKey(row.date)))).sort();
  const select = document.getElementById("periodSelect");
  select.innerHTML = "";

  const all = document.createElement("option");
  all.value = "__ALL__";
  all.textContent = "全期間";
  select.appendChild(all);

  for (const period of periods) {
    const opt = document.createElement("option");
    opt.value = period;
    opt.textContent = formatFiscalPeriodLabel(period);
    select.appendChild(opt);
  }

  select.value = "__ALL__";
}

async function bootstrap({ usageText, masterText }) {
  if (!window.Chart) {
    throw new Error("グラフライブラリを読み込めませんでした。インターネット接続を確認してください。");
  }

  const state = {
    usageRows: parseUsageRows(usageText),
    master: buildMaster(masterText),
  };

  setupPeriodSelect(state.usageRows);
  populateBlockFilter(state.master, "__ALL__");
  populateBranchFilter(buildRankings(state.usageRows, state.master, "__ALL__", "__ALL__", "__ALL__").branchRanking, "__ALL__");

  document.getElementById("periodSelect").addEventListener("change", () => renderDashboard(state));
  document.getElementById("blockFilterSelect").addEventListener("change", () => {
    document.getElementById("branchFilterSelect").value = "__ALL__";
    renderDashboard(state);
  });
  document.getElementById("branchFilterSelect").addEventListener("change", () => renderDashboard(state));
  document.getElementById("reloadBtn").addEventListener("click", () => {
    document.getElementById("periodSelect").value = "__ALL__";
    document.getElementById("blockFilterSelect").value = "__ALL__";
    document.getElementById("branchFilterSelect").value = "__ALL__";
    renderDashboard(state);
  });

  renderDashboard(state);
}

async function loadFromDefaultPaths() {
  const usageText = await fetchFirstText(DEFAULT_USAGE_PATHS);
  const masterText = await fetchFirstText(DEFAULT_MASTER_PATHS);
  return usageText && masterText ? { usageText, masterText } : null;
}

async function main() {
  setStatus("データ読み込み中...");
  showFileFallback(false);

  const defaultLoaded = await loadFromDefaultPaths();
  if (defaultLoaded) {
    await bootstrap(defaultLoaded);
    return;
  }

  setStatus("自動読み込みに失敗しました。ファイル選択で読み込めます。", "error");
  showFileFallback(true);
  document.getElementById("loadFilesBtn").addEventListener("click", async () => {
    const usageFile = document.getElementById("usageFile").files?.[0];
    const masterFile = document.getElementById("masterFile").files?.[0];
    if (!usageFile || !masterFile) {
      setStatus("2ファイルを選択してください。", "error");
      return;
    }

    try {
      await bootstrap({
        usageText: await readFileText(usageFile),
        masterText: await readFileText(masterFile),
      });
      showFileFallback(false);
    } catch (e) {
      setStatus(e?.message ?? "読み込みに失敗しました。", "error");
    }
  });
}

main().catch((e) => setStatus(e?.message ?? "エラーが発生しました。", "error"));
