// 自社版 定借ナビ（講師用の完成見本）
// データはブラウザの中だけで処理します。外部への送信は一切ありません。

const TSUBO = 3.30578;
const NEW_COLS = ["物件名", "区", "最寄駅", "駅徒歩分", "竣工年", "借地期限年", "総戸数", "価格", "専有面積"];
const OLD_COLS = ["物件名", "成約年月", "価格", "専有面積", "階"];
const WALKS = [
  { key: "all", label: "全て", test: () => true },
  { key: "le5", label: "5分以内", test: (m) => m != null && m <= 5 },
  { key: "6to9", label: "6〜9分", test: (m) => m != null && m >= 6 && m <= 9 },
  { key: "ge10", label: "10分以上", test: (m) => m != null && m >= 10 },
];

const state = { newRows: null, oldRows: null, props: [], ward: "all", walk: "all", charts: {} };
const $ = (id) => document.getElementById(id);

// ---------- 読み込み ----------

// UTF-8 で読めなければ Shift_JIS（Excel の普通の「CSV」保存）として読み直す
async function readText(file) {
  const buf = await file.arrayBuffer();
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(buf).replace(/^﻿/, "");
  } catch (e) {
    return new TextDecoder("shift_jis").decode(buf);
  }
}

function parseCsv(text, cols, label) {
  const res = Papa.parse(text.trim(), { header: true, skipEmptyLines: true });
  const header = (res.meta.fields || []).map((s) => s.trim());
  const missing = cols.filter((c) => !header.includes(c));
  if (missing.length) {
    throw new Error(`${label}の列名が足りません：${missing.join("、")}\n1行目は「${cols.join(",")}」にしてください。`);
  }
  return res.data.map((r) => {
    const o = {};
    for (const k in r) o[k.trim()] = typeof r[k] === "string" ? r[k].trim() : r[k];
    return o;
  });
}

const num = (v) => {
  if (v == null || v === "") return null;
  const n = Number(String(v).replace(/[,，円万㎡分年\s]/g, ""));
  return Number.isFinite(n) ? n : null;
};
const yearOf = (v) => {
  const m = String(v || "").match(/(\d{4})/);
  return m ? Number(m[1]) : null;
};

async function loadFile(input, kind) {
  const f = input.files[0];
  if (!f) return;
  try {
    const text = await readText(f);
    if (kind === "new") state.newRows = parseCsv(text, NEW_COLS, "新築データ");
    else state.oldRows = parseCsv(text, OLD_COLS, "中古データ");
    $("err").textContent = "";
    rebuild();
  } catch (e) {
    $("err").textContent = e.message;
  }
  input.value = "";
}

function loadSample() {
  state.newRows = parseCsv(window.SAMPLE_SHINCHIKU_CSV, NEW_COLS, "新築データ");
  state.oldRows = parseCsv(window.SAMPLE_CHUKO_CSV, OLD_COLS, "中古データ");
  $("err").textContent = "";
  rebuild();
}

function downloadTemplate(cols, name, example) {
  const csv = "﻿" + cols.join(",") + "\n" + example + "\n";
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
  a.download = name;
  a.click();
}

// ---------- 集計 ----------

function rebuild() {
  const newRows = state.newRows || [];
  const oldRows = state.oldRows || [];
  const byName = new Map();

  for (const r of newRows) {
    const price = num(r["価格"]), area = num(r["専有面積"]);
    if (!r["物件名"]) continue;
    if (!byName.has(r["物件名"])) {
      byName.set(r["物件名"], {
        name: r["物件名"], ward: r["区"] || "不明", station: r["最寄駅"] || "",
        walk: num(r["駅徒歩分"]), built: num(r["竣工年"]), limit: num(r["借地期限年"]),
        units: num(r["総戸数"]), newTsubos: [], deals: [],
      });
    }
    if (price && area) byName.get(r["物件名"]).newTsubos.push((price / area) * TSUBO);
  }

  let unmatched = 0;
  for (const r of oldRows) {
    const p = byName.get(r["物件名"]);
    const price = num(r["価格"]), area = num(r["専有面積"]), y = yearOf(r["成約年月"]);
    if (!p) { unmatched++; continue; }
    if (!price || !area || !y) continue;
    p.deals.push({ year: y, tsubo: (price / area) * TSUBO });
  }

  const nowYear = new Date().getFullYear();
  state.props = [...byName.values()].map((p) => {
    p.newTsubo = p.newTsubos.length ? avg(p.newTsubos) : null;
    p.remainNow = p.limit ? p.limit - nowYear : null;
    p.deals.forEach((d) => {
      d.remain = p.limit ? p.limit - d.year : null;
      d.ratio = p.newTsubo ? d.tsubo / p.newTsubo : null;
    });
    p.deals.sort((a, b) => a.year - b.year);
    const last = p.deals[p.deals.length - 1];
    p.latestTsubo = last ? last.tsubo : null;
    p.ratio = p.deals.length && p.newTsubo ? median(p.deals.map((d) => d.ratio)) : null;
    return p;
  });

  const wards = [...new Set(state.props.map((p) => p.ward))].sort();
  if (state.ward !== "all" && !wards.includes(state.ward)) state.ward = "all";
  $("ward").innerHTML = `<option value="all">全体</option>` +
    wards.map((w) => `<option value="${esc(w)}">${esc(w)}</option>`).join("");
  $("ward").value = state.ward;

  const msg = [];
  msg.push(state.newRows ? `新築データ ${newRows.length}行（${state.props.length}物件）` : "新築データ 未読み込み");
  msg.push(state.oldRows ? `中古データ ${oldRows.length}行` : "中古データ 未読み込み");
  if (unmatched) msg.push(`※物件名が新築データと一致しない中古データ ${unmatched}行は除外`);
  $("status").textContent = msg.join(" ／ ");
  render();
}

function avg(a) { return a.reduce((s, x) => s + x, 0) / a.length; }
function median(a) {
  const s = [...a].sort((x, y) => x - y), m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}
function esc(s) { return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }

// ---------- 表示 ----------

function filtered(walkKey = state.walk) {
  const w = WALKS.find((x) => x.key === walkKey);
  return state.props.filter((p) => (state.ward === "all" || p.ward === state.ward) && w.test(p.walk));
}

function render() {
  // 駅徒歩ボタン（件数つき）
  $("walk").innerHTML = WALKS.map((w) =>
    `<button data-k="${w.key}" class="${w.key === state.walk ? "on" : ""}">${w.label}<small>${filtered(w.key).length}</small></button>`
  ).join("");

  const props = filtered();
  const deals = props.flatMap((p) => p.deals.filter((d) => d.remain != null));

  // 5年刻みのバケツ（例: 55〜59年）
  const buckets = new Map();
  for (const d of deals) {
    const b = Math.floor(d.remain / 5) * 5;
    if (!buckets.has(b)) buckets.set(b, []);
    buckets.get(b).push(d);
  }
  // 新築時（例: 残存70年）の点も描けるよう、新築時のバケツも横軸に含める
  for (const p of props) {
    if (p.newTsubo && p.built && p.limit) {
      const b = Math.floor((p.limit - p.built) / 5) * 5;
      if (!buckets.has(b)) buckets.set(b, []);
    }
  }
  const keys = [...buckets.keys()].sort((a, b) => b - a);
  const labels = keys.map((k) => `${k}〜${k + 4}年`);
  const cnt = keys.map((k) => buckets.get(k).length);

  const ratios = (k) => buckets.get(k).filter((d) => d.ratio != null).map((d) => d.ratio);
  const tsubos = (k) => buckets.get(k).map((d) => d.tsubo);
  const ratioMed = keys.map((k) => (ratios(k).length ? pct(median(ratios(k))) : null));
  const ratioAvg = keys.map((k) => (ratios(k).length ? pct(avg(ratios(k))) : null));
  const tsuboMed = keys.map((k) => (tsubos(k).length ? round(median(tsubos(k))) : null));
  const tsuboAvg = keys.map((k) => (tsubos(k).length ? round(avg(tsubos(k))) : null));
  const newPts = props.filter((p) => p.newTsubo && p.built && p.limit).map((p) => {
    const b = Math.floor((p.limit - p.built) / 5) * 5;
    return { x: `${b}〜${b + 4}年`, y: round(p.newTsubo), name: p.name };
  }).filter((pt) => labels.includes(pt.x));

  const tip = { callbacks: { afterBody: (items) => `データ件数：${cnt[items[0].dataIndex]}件` } };

  draw("chartRatio", {
    data: {
      labels,
      datasets: [
        { type: "line", label: "中央値", data: ratioMed, borderColor: "#059669", backgroundColor: "#059669", tension: 0.3, spanGaps: true },
        { type: "bar", label: "平均", data: ratioAvg, backgroundColor: "rgba(5,150,105,.18)" },
        { type: "line", label: "新築時（100%）", data: labels.map(() => 100), borderColor: "#94a3b8", borderDash: [6, 6], pointRadius: 0 },
      ],
    },
    options: { plugins: { tooltip: tip }, scales: { y: { ticks: { callback: (v) => v + "%" } } } },
  });

  draw("chartTsubo", {
    data: {
      labels,
      datasets: [
        { type: "line", label: "中央値", data: tsuboMed, borderColor: "#2563eb", backgroundColor: "#2563eb", tension: 0.3, spanGaps: true },
        { type: "bar", label: "平均", data: tsuboAvg, backgroundColor: "rgba(37,99,235,.15)" },
        { type: "scatter", label: "新築時坪単価", data: newPts, backgroundColor: "#f59e0b", pointRadius: 6 },
      ],
    },
    options: {
      plugins: { tooltip: { callbacks: { label: (c) => c.raw && c.raw.name ? `${c.raw.name}：${c.raw.y}万円/坪` : `${c.dataset.label}：${c.formattedValue}万円/坪` } } },
      scales: { y: { ticks: { callback: (v) => v + "万" } } },
    },
  });

  // 一覧
  $("listTitle").textContent = `対象物件一覧（${props.length}件）`;
  $("tbody").innerHTML = props.length ? props.map((p) => `
    <tr>
      <td><a data-name="${esc(p.name)}">${esc(p.name)}</a></td>
      <td>${esc(p.ward)}</td>
      <td>${p.built ?? "-"}年</td>
      <td>${p.limit ?? "-"}年</td>
      <td>${p.remainNow ?? "-"}年</td>
      <td>${p.units ?? "-"}戸</td>
      <td>${esc(p.station)}${p.walk != null ? `（徒歩${p.walk}分）` : ""}</td>
      <td>${badge(p.ratio)}</td>
      <td>${p.deals.length}件</td>
    </tr>`).join("")
    : `<tr><td colspan="9" style="color:#94a3b8">条件に合う物件がありません</td></tr>`;
}

function badge(r) {
  if (r == null) return "-";
  const v = Math.round(r * 100), c = v >= 100 ? "b-g" : v >= 80 ? "b-y" : "b-r";
  return `<span class="badge ${c}">${v}%</span>`;
}
const pct = (v) => (Number.isFinite(v) ? Math.round(v * 1000) / 10 : null);
const round = (v) => (Number.isFinite(v) ? Math.round(v) : null);

function draw(id, cfg) {
  if (state.charts[id]) state.charts[id].destroy();
  state.charts[id] = new Chart($(id), {
    type: "bar",
    ...cfg,
    options: { responsive: true, maintainAspectRatio: false, animation: false, ...cfg.options },
  });
}

// ---------- 物件詳細 ----------

function showDetail(name) {
  const p = state.props.find((x) => x.name === name);
  if (!p) return;
  $("listView").classList.add("hidden");
  $("detailView").classList.remove("hidden");
  $("dTitle").textContent = p.name;
  $("dSub").textContent = `${p.ward}／${p.station}${p.walk != null ? ` 徒歩${p.walk}分` : ""}／竣工${p.built}年／借地期限${p.limit}年`;
  $("dStats").innerHTML = [
    ["データ件数", `${p.deals.length}件`],
    ["新築時坪単価", p.newTsubo ? `${round(p.newTsubo)}万円` : "-"],
    ["最新成約坪単価", p.latestTsubo ? `${round(p.latestTsubo)}万円` : "-"],
    ["新築比（中央値）", p.ratio ? `${Math.round(p.ratio * 100)}%` : "-"],
  ].map(([k, v]) => `<div class="stat">${k}<b>${v}</b></div>`).join("");

  const by = new Map();
  p.deals.forEach((d) => { if (!by.has(d.remain)) by.set(d.remain, []); by.get(d.remain).push(d.tsubo); });
  const ks = [...by.keys()].sort((a, b) => b - a);
  draw("chartDetail", {
    data: {
      labels: ks.map((k) => `${k}年`),
      datasets: [
        { type: "line", label: "成約坪単価（中央値）", data: ks.map((k) => round(median(by.get(k)))), borderColor: "#2563eb", backgroundColor: "#2563eb", tension: 0.3 },
        { type: "line", label: "新築時坪単価", data: ks.map(() => round(p.newTsubo)), borderColor: "#f59e0b", borderDash: [6, 6], pointRadius: 0 },
      ],
    },
    options: { scales: { y: { ticks: { callback: (v) => v + "万" } } } },
  });
  window.scrollTo(0, 0);
}

function showList() {
  $("detailView").classList.add("hidden");
  $("listView").classList.remove("hidden");
}

// ---------- イベント ----------

$("btnNew").onclick = () => $("fileNew").click();
$("btnOld").onclick = () => $("fileOld").click();
$("fileNew").onchange = (e) => loadFile(e.target, "new");
$("fileOld").onchange = (e) => loadFile(e.target, "old");
$("btnSample").onclick = loadSample;
$("tplNew").onclick = () => downloadTemplate(NEW_COLS, "新築データ_テンプレート.csv", "〇〇レジデンス,文京区,湯島,3,2012,2082,120,5800,65.2");
$("tplOld").onclick = () => downloadTemplate(OLD_COLS, "中古データ_テンプレート.csv", "〇〇レジデンス,2024/05,6200,65.2,7");
$("ward").onchange = (e) => { state.ward = e.target.value; render(); };
$("walk").onclick = (e) => {
  const b = e.target.closest("button");
  if (b) { state.walk = b.dataset.k; render(); }
};
$("tbody").onclick = (e) => { if (e.target.dataset.name) showDetail(e.target.dataset.name); };
$("btnBack").onclick = showList;

// 研修資料のスクリーンショット用：index.html?demo=1 でサンプルを自動表示
const q = new URLSearchParams(location.search);
if (q.get("demo")) {
  loadSample();
  if (q.get("walk")) { state.walk = q.get("walk"); render(); }
  if (q.get("detail")) showDetail(q.get("detail"));
}
