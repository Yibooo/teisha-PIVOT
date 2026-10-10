// 勉強会スライド「AIで 自社版 定借ナビ をつくる」
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const P = require("./prompts");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.title = "AIで 自社版 定借ナビ をつくる 勉強会";

const W = 13.33, H = 7.5, MX = 0.6;
const C = {
  DARK: "1F2E2B", GOLD: "B08D47", GOLDL: "CBB07A", CREAM: "F5F2EA", INK: "1F2A28", BODY: "3E3E3B",
  MUTE: "7A756B", GREEN: "3C5A4F", LINE: "E4DFD4", CARD: "FBF9F4", RED: "E5484D", CODE: "1E293B",
  OK: "2F7D5B", OKBG: "EAF4EE", WARN: "9A4A3A", WARNBG: "FBEFEC", BLUE: "2563EB", GH: "1F883D",
};
const F = "Yu Gothic";
const SHOT = "../images/";
const BREAK = process.env.BREAK === "1"; // チェッカー検証用にわざと崩す

// ---- 配置の記録（重なりチェック用） ----
const reg = [];
let cur = 0;
function rec(kind, x, y, w, h, extra = {}) { reg.push({ slide: cur, kind, x, y, w, h, ...extra }); }

function T(s, text, o) {
  s.addText(text, { fontFace: F, color: C.BODY, margin: 0, valign: "top", ...o });
  rec(o.overlay ? "overlay" : "text", o.x, o.y, o.w, o.h, { str: typeof text === "string" ? text : text.map((t) => t.text).join("") });
}
function R(s, o, shape = pptx.ShapeType.rect) {
  s.addShape(shape, o);
  rec(o.overlay ? "overlay" : "shape", o.x, o.y, o.w, o.h);
}
function I(s, path, o) {
  s.addImage({ path, ...o });
  rec("image", o.x, o.y, o.w, o.h);
}

// ---- 共通パーツ ----
let pageNo = 0;
function slide(part, title, sub) {
  const s = pptx.addSlide();
  cur = ++pageNo;
  s.background = { color: "FFFFFF" };
  if (part) {
    T(s, part, { x: MX, y: 0.32, w: 6, h: 0.3, fontSize: 11, bold: true, color: C.GOLD, charSpacing: 1 });
  }
  if (title) T(s, title, { x: MX, y: 0.62, w: W - MX * 2, h: 0.6, fontSize: 24, bold: true, color: C.INK });
  if (sub) T(s, sub, { x: MX, y: 1.2, w: W - MX * 2, h: 0.36, fontSize: 13, color: C.MUTE });
  s.addShape(pptx.ShapeType.line, { x: MX, y: 7.0, w: W - MX * 2, h: 0, line: { color: C.LINE, width: 0.75 } });
  T(s, "AIで 自社版 定借ナビ をつくる 勉強会", { x: MX, y: 7.08, w: 6, h: 0.25, fontSize: 9, color: C.MUTE });
  T(s, String(pageNo), { x: W - MX - 1, y: 7.08, w: 1, h: 0.25, fontSize: 9, color: C.MUTE, align: "right" });
  return s;
}

// ブラウザの枠（イメージ図）。中身を描く領域を返す
function browser(s, x, y, w, h, url) {
  R(s, { x, y, w, h, fill: { color: "FFFFFF" }, line: { color: "C9CED6", width: 1 }, rectRadius: 0.08 }, pptx.ShapeType.roundRect);
  R(s, { x: x + 0.02, y: y + 0.02, w: w - 0.04, h: 0.36, fill: { color: "EEF0F3" }, line: { type: "none" } });
  ["FF5F57", "FEBC2E", "28C840"].forEach((c, i) =>
    R(s, { x: x + 0.14 + i * 0.2, y: y + 0.13, w: 0.12, h: 0.12, fill: { color: c }, line: { type: "none" } }, pptx.ShapeType.ellipse));
  R(s, { x: x + 0.85, y: y + 0.07, w: w - 1.0, h: 0.24, fill: { color: "FFFFFF" }, line: { color: "D8DCE2", width: 0.5 }, rectRadius: 0.1 }, pptx.ShapeType.roundRect);
  T(s, url, { x: x + 1.0, y: y + 0.08, w: w - 1.3, h: 0.22, fontSize: 9.5, color: "5B6470", valign: "middle" });
  T(s, "イメージ図（実際の画面と細部が異なることがあります）", { x, y: y + h + 0.04, w, h: 0.22, fontSize: 8.5, color: C.MUTE, align: "right" });
  return { x: x + 0.2, y: y + 0.55, w: w - 0.4, h: h - 0.7 };
}

// 画面の部品（イメージ図の中）
function ui(s, x, y, w, h, text, st = {}) {
  const fill = st.fill || "FFFFFF";
  R(s, { x, y, w, h, fill: { color: fill }, line: st.noline ? { type: "none" } : { color: st.line || "C9CED6", width: 0.75 }, rectRadius: st.r ?? 0.06 }, pptx.ShapeType.roundRect);
  if (text) T(s, text, { x: x + 0.06, y, w: w - 0.12, h, fontSize: st.fs || 10, color: st.color || "24292F", bold: !!st.bold, align: st.align || "center", valign: "middle" });
}

// 赤枠＋番号（押す場所）
function hl(s, x, y, w, h, n, pos = "tl") {
  R(s, { x: x - 0.05, y: y - 0.05, w: w + 0.1, h: h + 0.1, fill: { type: "none" }, line: { color: C.RED, width: 2.5 }, rectRadius: 0.06, overlay: true }, pptx.ShapeType.roundRect);
  if (n != null) {
    let cx = x - 0.05 - 0.17, cy = y - 0.05 - 0.17;
    if (pos === "l") { cx = x - 0.05 - 0.42; cy = y + h / 2 - 0.17; }
    if (pos === "r") { cx = x + w + 0.08; cy = y + h / 2 - 0.17; }
    R(s, { x: cx, y: cy, w: 0.34, h: 0.34, fill: { color: C.RED }, line: { color: "FFFFFF", width: 1 }, overlay: true }, pptx.ShapeType.ellipse);
    T(s, String(n), { x: cx, y: cy, w: 0.34, h: 0.34, fontSize: 12, bold: true, color: "FFFFFF", align: "center", valign: "middle", overlay: true });
  }
}

// テキストの高さの目安（全角=1em、半角=0.55em）
function estH(text, fs, wIn, ls = 1.15) {
  let lines = 0;
  for (const para of String(text).split("\n")) {
    let wpt = 0;
    for (const ch of para) wpt += (ch.charCodeAt(0) < 0x2000 ? 0.56 : 1.0) * fs;
    lines += Math.max(1, Math.ceil(wpt / (wIn * 72 * 0.97)));
  }
  return (lines * fs * 1.25 * ls) / 72;
}

// 右側の手順パネル
function steps(s, x, y, w, items, o = {}) {
  const gap = o.gap ?? 0.16;
  const fs = o.fs || 14;
  let yy = y;
  items.forEach((it, i) => {
    const h = Math.max(0.42, estH(it.t, fs, w - 0.5) + (it.d ? estH(it.d, fs - 2.5, w - 0.5) : 0) + 0.04);
    R(s, { x: x, y: yy + 0.02, w: 0.36, h: 0.36, fill: { color: it.plain ? C.GOLD : C.RED }, line: { type: "none" } }, pptx.ShapeType.ellipse);
    T(s, it.n ?? String(i + 1), { x: x, y: yy + 0.02, w: 0.36, h: 0.36, fontSize: 13, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    const runs = [{ text: it.t, options: { bold: true, color: C.INK, fontSize: o.fs || 14, breakLine: !!it.d } }];
    if (it.d) runs.push({ text: it.d, options: { color: C.BODY, fontSize: (o.fs || 14) - 2.5 } });
    T(s, runs, { x: x + 0.5, y: yy, w: w - 0.5, h, lineSpacingMultiple: 1.15 });
    yy += h + gap;
  });
  return yy;
}

// コピペ用プロンプト
function promptBox(s, x, y, w, h, p, fs = 10.5, ls = 1.12) {
  R(s, { x, y, w, h, fill: { color: C.CODE }, line: { type: "none" }, rectRadius: 0.08 }, pptx.ShapeType.roundRect);
  T(s, "📋 コピーして Claude の入力欄に貼り付け", { x: x + 0.25, y: y + 0.12, w: w - 0.5, h: 0.28, fontSize: 10.5, bold: true, color: C.GOLDL });
  T(s, p.text, { x: x + 0.25, y: y + 0.48, w: w - 0.5, h: h - 0.6, fontSize: fs, color: "F1F5F9", lineSpacingMultiple: ls });
}

function bar(s, y, kind, text, o = {}) {
  const st = { ok: [C.OKBG, C.OK, "✓ 確認ポイント"], warn: [C.WARNBG, C.WARN, "⚠ 注意"], tip: [C.CREAM, C.GOLD, "💡 ヒント"] }[kind];
  const x = o.x ?? MX, w = o.w ?? W - MX * 2, h = o.h ?? 0.5;
  R(s, { x, y, w, h, fill: { color: st[0] }, line: { type: "none" }, rectRadius: 0.06 }, pptx.ShapeType.roundRect);
  R(s, { x, y, w: 0.08, h, fill: { color: st[1] }, line: { type: "none" } });
  T(s, [{ text: (o.label || st[2]) + "　", options: { bold: true, color: st[1] } }, { text, options: { color: C.INK } }],
    { x: x + 0.25, y, w: w - 0.4, h, fontSize: o.fs || 12.5, valign: "middle", lineSpacingMultiple: 1.1 });
}

function card(s, x, y, w, h, head, body, o = {}) {
  R(s, { x, y, w, h, fill: { color: o.fill || C.CARD }, line: { color: o.line || C.LINE, width: 1 }, rectRadius: 0.08 }, pptx.ShapeType.roundRect);
  T(s, head, { x: x + 0.22, y: y + 0.16, w: w - 0.44, h: 0.4, fontSize: o.hfs || 15, bold: true, color: o.hcolor || C.INK });
  if (body) T(s, body, { x: x + 0.22, y: y + 0.6, w: w - 0.44, h: h - 0.72, fontSize: o.fs || 12, color: C.BODY, lineSpacingMultiple: 1.2 });
}

function bullets(arr) {
  return arr.map((t, i) => ({ text: t, options: { bullet: { code: "25CF" }, breakLine: i < arr.length - 1, paraSpaceAfter: 4 } }));
}

// 第◯部の扉
function divider(num, title, desc) {
  const s = pptx.addSlide();
  cur = ++pageNo;
  s.background = { color: C.DARK };
  T(s, `第${num}部`, { x: 1.0, y: 2.4, w: 6, h: 0.5, fontSize: 18, bold: true, color: C.GOLDL });
  T(s, title, { x: 1.0, y: 2.95, w: 11, h: 0.9, fontSize: 36, bold: true, color: "FFFFFF" });
  s.addShape(pptx.ShapeType.line, { x: 1.0, y: 4.0, w: 1.2, h: 0, line: { color: C.GOLD, width: 2 } });
  T(s, desc, { x: 1.0, y: 4.25, w: 11, h: 0.9, fontSize: 15, color: "D9D4C7", lineSpacingMultiple: 1.3 });
  T(s, String(pageNo), { x: W - MX - 1, y: 7.08, w: 1, h: 0.25, fontSize: 9, color: "8C877B", align: "right" });
}

const REPO = "teisha-tool";
const COPY_URL = "github.com/Yibooo/teisha-PIVOT → study-session → prompts.md";

// =====================================================================
// 1. 表紙
{
  const s = pptx.addSlide(); cur = ++pageNo;
  s.background = { color: C.DARK };
  T(s, "社内勉強会", { x: 1.0, y: 1.5, w: 6, h: 0.4, fontSize: 16, bold: true, color: C.GOLDL, charSpacing: 2 });
  T(s, "AIで「自社版 定借ナビ」をつくる", { x: 1.0, y: 2.05, w: 11.3, h: 1.0, fontSize: 40, bold: true, color: "FFFFFF" });
  T(s, "Claude × GitHub 入門 ─ アカウント作成から、自社の新築・中古データで\n資産価値グラフを表示するまで", { x: 1.0, y: 3.2, w: 11.3, h: 0.9, fontSize: 18, color: "D9D4C7", lineSpacingMultiple: 1.3 });
  s.addShape(pptx.ShapeType.line, { x: 1.0, y: 4.45, w: 1.2, h: 0, line: { color: C.GOLD, width: 2 } });
  T(s, "対象：AI・Claude・GitHub をさわったことがない方\n所要時間：約3時間（休憩含む）／ 持ち物：ノートPC（Chrome推奨）", { x: 1.0, y: 4.7, w: 11, h: 0.8, fontSize: 14, color: "D9D4C7", lineSpacingMultiple: 1.4 });
  T(s, "定借ナビ", { x: 1.0, y: 6.4, w: 4, h: 0.4, fontSize: 14, bold: true, color: C.GOLDL });
}

// 2. 今日のゴール
{
  const s = slide("はじめに", "今日のゴール：このツールを、自分の手で作れるようになる", "お手本は「定借ナビ」（teisha-navi.vercel.app）。同じ考え方のツールを、AIに作ってもらいます");
  I(s, SHOT + "c_top.png", { x: MX, y: 1.75, w: 7.4, h: 7.4 * 420 / 1100 });
  R(s, { x: MX, y: 1.75, w: 7.4, h: 7.4 * 420 / 1100, fill: { type: "none" }, line: { color: C.LINE, width: 1 } });
  I(s, SHOT + "c_table_top.png", { x: MX, y: 4.75, w: 7.4, h: 7.4 * 305 / 1100 });
  R(s, { x: MX, y: 4.75, w: 7.4, h: 7.4 * 305 / 1100, fill: { type: "none" }, line: { color: C.LINE, width: 1 } });
  const x = 8.4, w = W - MX - x;
  card(s, x, 1.75, w, 1.55, "① ツールを作れる", "プログラミングの知識はいりません。\n日本語で「お願い」を書くと、AIが作ります", { fs: 12 });
  card(s, x, 3.45, w, 1.55, "② 自社データを入れられる", "自社の新築・中古データ（CSV）を\nボタン1つで読み込み、グラフにします", { fs: 12 });
  card(s, x, 5.15, w, 1.65, "③ グラフを説明できる", "「新築比」「残存年数」など、\nグラフの意味を理解し、お客様に\n説明できるようになります", { fs: 12 });
}

// 3. 今日の流れ
{
  const s = slide("はじめに", "今日の流れ", "スライドのとおりに進めれば、最後に自社データのグラフが表示されます");
  const rows = [
    ["第1部", "準備する", "GitHub と Claude のアカウントを作る", "STEP 1〜4", "40分"],
    ["第2部", "つなぐ", "Claude と GitHub を連携し、つながったか確かめる", "STEP 5〜8", "30分"],
    ["第3部", "作る", "プロンプト①〜③を貼り付けて、ツールを完成させる", "STEP 9〜15", "60分"],
    ["第4部", "自社データを入れる", "Excel を CSV で保存し、ツールに読み込む", "STEP 16〜18", "30分"],
    ["第5部", "困ったとき・使いこなす", "うまくいかないときのプロンプト、お客様への説明", "─", "20分"],
  ];
  let y = 1.8;
  rows.forEach((r, i) => {
    R(s, { x: MX, y, w: W - MX * 2, h: 0.82, fill: { color: i % 2 ? "FFFFFF" : C.CARD }, line: { color: C.LINE, width: 0.75 } });
    T(s, r[0], { x: MX + 0.25, y, w: 1.0, h: 0.82, fontSize: 14, bold: true, color: C.GOLD, valign: "middle" });
    T(s, r[1], { x: MX + 1.3, y, w: 3.2, h: 0.82, fontSize: 17, bold: true, color: C.INK, valign: "middle" });
    T(s, r[2], { x: MX + 4.5, y, w: 5.3, h: 0.82, fontSize: 13, color: C.BODY, valign: "middle" });
    T(s, r[3], { x: MX + 9.9, y, w: 1.4, h: 0.82, fontSize: 12, color: C.MUTE, valign: "middle" });
    T(s, r[4], { x: MX + 11.3, y, w: 0.75, h: 0.82, fontSize: 13, bold: true, color: C.INK, valign: "middle", align: "right" });
    y += 0.82;
  });
  bar(s, 6.25, "tip", "各スライドの左上の「STEP 番号」で、いまどこにいるかが分かります。迷ったら番号を講師に伝えてください");
}

// 4. 用語
{
  const s = slide("はじめに", "最初に覚える 8つの言葉", "今日はこれだけ分かれば大丈夫です。たとえで覚えましょう");
  const terms = [
    ["Claude（クロード）", "Anthropic社のAI。日本語で話しかけると、文章やプログラムを作ってくれる"],
    ["Claude Code", "Claude が「プログラムを書く作業」をしてくれる機能。今日は Web版（claude.ai/code）を使う"],
    ["GitHub（ギットハブ）", "プログラムを保管する場所。いわば「設計図の保管庫」"],
    ["リポジトリ", "GitHub の中の1つのプロジェクト用フォルダ。今日は「" + REPO + "」を作る"],
    ["プロンプト", "AIへの「お願い文」。このスライドのものをコピーして貼るだけでOK"],
    ["プルリクエスト（PR）", "「この変更を取り込んでいいですか？」という確認依頼"],
    ["マージ", "確認した変更を、正式に保管庫へ取り込むこと"],
    ["CSV", "カンマ区切りの表データ。Excel から「CSV で保存」で作れる"],
  ];
  const cw = (W - MX * 2 - 0.3) / 2;
  terms.forEach((t, i) => {
    const x = MX + (i % 2) * (cw + 0.3), y = 1.8 + Math.floor(i / 2) * 1.25;
    R(s, { x, y, w: cw, h: 1.1, fill: { color: C.CARD }, line: { color: C.LINE, width: 0.75 }, rectRadius: 0.06 }, pptx.ShapeType.roundRect);
    T(s, t[0], { x: x + 0.22, y: y + 0.12, w: cw - 0.44, h: 0.36, fontSize: 15, bold: true, color: C.GREEN });
    T(s, t[1], { x: x + 0.22, y: y + 0.5, w: cw - 0.44, h: 0.52, fontSize: 12, color: C.BODY });
  });
}

// 5. 全体像
{
  const s = slide("はじめに", "全体像：だれが・どこで・何をするか", "自社データは、最後まで自分のPCの中から出しません");
  const boxes = [
    ["あなた", "ブラウザで\nプロンプトを貼る", C.GOLD],
    ["Claude Code", "AIがプログラムを\n書いて保存する", C.GREEN],
    ["GitHub", "プログラムを保管\n（自社データは置かない）", "24292F"],
    ["あなたのPC", "ZIPでダウンロードし\nindex.html を開く", C.GOLD],
  ];
  const bw = 2.55, gap = 0.55, x0 = MX + 0.1, y = 2.0;
  boxes.forEach((b, i) => {
    const x = x0 + i * (bw + gap);
    R(s, { x, y, w: bw, h: 2.0, fill: { color: "FFFFFF" }, line: { color: b[2], width: 2 }, rectRadius: 0.1 }, pptx.ShapeType.roundRect);
    R(s, { x, y, w: bw, h: 0.6, fill: { color: b[2] }, line: { type: "none" }, rectRadius: 0.1 }, pptx.ShapeType.roundRect);
    T(s, b[0], { x, y, w: bw, h: 0.6, fontSize: 16, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    T(s, b[1], { x: x + 0.15, y: y + 0.75, w: bw - 0.3, h: 1.1, fontSize: 13, color: C.INK, align: "center", valign: "middle", lineSpacingMultiple: 1.2 });
    if (i < 3) T(s, "▶", { x: x + bw + 0.05, y: y + 0.75, w: gap - 0.1, h: 0.5, fontSize: 20, color: C.MUTE, align: "center", valign: "middle" });
  });
  R(s, { x: MX + 0.1, y: 4.45, w: W - MX * 2 - 0.2, h: 1.6, fill: { color: C.OKBG }, line: { color: C.OK, width: 1.5 }, rectRadius: 0.1 }, pptx.ShapeType.roundRect);
  T(s, "🔒 自社の新築・中古データはここだけ", { x: MX + 0.4, y: 4.6, w: 8, h: 0.4, fontSize: 16, bold: true, color: C.OK });
  T(s, "できあがったツールを自分のPCで開き、画面の「新築データを読み込む」「中古データを読み込む」ボタンから CSV を選びます。\nデータはPCのブラウザの中だけで計算され、Claude にも GitHub にも送られません。",
    { x: MX + 0.4, y: 5.05, w: W - MX * 2 - 0.8, h: 0.9, fontSize: 13, color: C.INK, lineSpacingMultiple: 1.3 });
  bar(s, 6.25, "tip", "Claude に渡すのは「どんなツールにしたいか」というお願い文だけです");
}

// 6. 事前準備
{
  const s = slide("はじめに", "始める前に用意するもの", "当日までにそろえておくと、スムーズに進みます");
  const items = [
    ["ノートPC と Chrome", "Windows / Mac どちらでもOK。ブラウザは Google Chrome を推奨"],
    ["メールアドレス", "GitHub と Claude の登録に使います。会社のルールに従ってください"],
    ["Claude の有料プラン", "Claude Code を使うには Pro 以上のプランが必要です。会社で Team / Enterprise を契約済みなら、管理者に利用できるか確認"],
    ["新築データ・中古データ", "Excel などで管理している物件データ。第4部で使います（事前に送る必要はありません）"],
    ["社内ルールの確認", "AIツールの利用ルール、データの持ち出しルールを情報システム部門などに確認"],
  ];
  let y = 1.8;
  items.forEach((it) => {
    R(s, { x: MX, y, w: W - MX * 2, h: 0.85, fill: { color: C.CARD }, line: { color: C.LINE, width: 0.75 }, rectRadius: 0.06 }, pptx.ShapeType.roundRect);
    T(s, "☐", { x: MX + 0.2, y, w: 0.5, h: 0.85, fontSize: 22, color: C.GOLD, valign: "middle" });
    T(s, it[0], { x: MX + 0.8, y, w: 3.4, h: 0.85, fontSize: 15, bold: true, color: C.INK, valign: "middle" });
    T(s, it[1], { x: MX + 4.2, y, w: W - MX * 2 - 4.4, h: 0.85, fontSize: 12.5, color: C.BODY, valign: "middle" });
    y += 0.97;
  });
}

// 7. データの取り扱い
{
  const s = slide("はじめに", "いちばん大切なルール：データの取り扱い", "ツール作りの前に、全員で確認します");
  const rules = [
    ["1", "自社の物件データは、Claude にも GitHub にも渡さない", "プロンプトに物件名・価格などの中身を貼らない。CSV を GitHub にアップロードしない"],
    ["2", "相談するときは「列名だけ」", "列名が違うなどで相談するときは、CSV の1行目（列名）だけを貼る。中身の行は貼らない"],
    ["3", "リポジトリは必ず「Private（非公開）」", "STEP 2 で作るリポジトリは Private にする。ツールのプログラムも社外に見せない"],
    ["4", "社内規程とデータ提供元のルールに従う", "データの入手元ごとの利用条件、社内の情報管理ルールを守る。迷ったら上長・管理部門に確認"],
  ];
  let y = 1.8;
  rules.forEach((r) => {
    R(s, { x: MX, y, w: W - MX * 2, h: 1.05, fill: { color: "FFFFFF" }, line: { color: C.LINE, width: 1 }, rectRadius: 0.08 }, pptx.ShapeType.roundRect);
    R(s, { x: MX + 0.25, y: y + 0.27, w: 0.5, h: 0.5, fill: { color: C.WARN }, line: { type: "none" } }, pptx.ShapeType.ellipse);
    T(s, r[0], { x: MX + 0.25, y: y + 0.27, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    T(s, r[1], { x: MX + 1.0, y: y + 0.14, w: W - MX * 2 - 1.3, h: 0.42, fontSize: 16, bold: true, color: C.INK });
    T(s, r[2], { x: MX + 1.0, y: y + 0.56, w: W - MX * 2 - 1.3, h: 0.4, fontSize: 12.5, color: C.BODY });
    y += 1.17;
  });
}

// =====================================================================
divider(1, "準備する ─ GitHub と Claude", "STEP 1　GitHub のアカウントを作る\nSTEP 2　ツールを保管するリポジトリを作る\nSTEP 3〜4　Claude のアカウントを作り、プランを選ぶ");

// STEP1 GitHub アカウント
{
  const s = slide("第1部　STEP 1", "GitHub のアカウントを作る", "GitHub の画面は英語です。押す場所は赤枠のとおりです");
  const a = browser(s, MX, 1.7, 7.6, 4.95, "https://github.com");
  R(s, { x: a.x - 0.18, y: a.y - 0.15, w: a.w + 0.36, h: 0.5, fill: { color: "0D1117" }, line: { type: "none" } });
  T(s, "GitHub", { x: a.x, y: a.y - 0.15, w: 1.5, h: 0.5, fontSize: 13, bold: true, color: "FFFFFF", valign: "middle" });
  ui(s, a.x + a.w - 2.1, a.y - 0.07, 0.9, 0.34, "Sign in", { fill: "0D1117", line: "0D1117", color: "FFFFFF", fs: 10 });
  ui(s, a.x + a.w - 1.05, a.y - 0.07, 1.0, 0.34, "Sign up", { fill: "0D1117", line: "8B949E", color: "FFFFFF", fs: 10 });
  hl(s, a.x + a.w - 1.05, a.y - 0.07, 1.0, 0.34, 1);
  T(s, "Create your free account", { x: a.x + 1.2, y: a.y + 0.6, w: 4.8, h: 0.4, fontSize: 15, bold: true, color: "24292F", align: "center" });
  const fx = a.x + 1.6, fw = 4.0;
  [["Email", 1.2], ["Password", 1.9], ["Username", 2.6]].forEach(([l, dy]) => {
    T(s, l, { x: fx, y: a.y + dy, w: fw, h: 0.22, fontSize: 9.5, color: "57606A" });
    ui(s, fx, a.y + dy + 0.23, fw, 0.34, "", {});
  });
  hl(s, fx, a.y + 1.2, fw, 1.97, 2, "l");
  ui(s, fx, a.y + 3.4, fw, 0.4, "Continue", { fill: C.GH, line: C.GH, color: "FFFFFF", bold: true });
  hl(s, fx, a.y + 3.4, fw, 0.4, 3, "l");
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "github.com を開き、右上の「Sign up」", d: "Chrome のアドレス欄に github.com と入力して Enter", h: 0.8 },
    { t: "メール・パスワード・ユーザー名を入力", d: "ユーザー名は英数字。例：tanaka-realestate\n（あとで画面に表示されます）", h: 1.1 },
    { t: "「Continue」→ 画面の指示に従う", d: "パズル認証が出たら自分で解きます。メールに届く数字（認証コード）を入力", h: 1.1 },
    { t: "無料プラン（Free）でOK", d: "「Continue for free」などを選びます", h: 0.8, plain: true, n: "✓" },
  ]);
}

// STEP2 リポジトリ作成 (1)
{
  const s = slide("第1部　STEP 2", "リポジトリ（保管用フォルダ）を作る ①", "ログインした状態で、画面右上の「＋」から作ります");
  const a = browser(s, MX, 1.7, 7.6, 4.95, "https://github.com");
  R(s, { x: a.x - 0.18, y: a.y - 0.15, w: a.w + 0.36, h: 0.5, fill: { color: "F6F8FA" }, line: { color: "D0D7DE", width: 0.5 } });
  T(s, "≡  Dashboard", { x: a.x, y: a.y - 0.15, w: 2, h: 0.5, fontSize: 11, bold: true, color: "24292F", valign: "middle" });
  ui(s, a.x + a.w - 1.3, a.y - 0.07, 0.6, 0.34, "＋ ▾", { fs: 10 });
  hl(s, a.x + a.w - 1.3, a.y - 0.07, 0.6, 0.34, 1);
  R(s, { x: a.x + a.w - 0.55, y: a.y - 0.07, w: 0.34, h: 0.34, fill: { color: "D0D7DE" }, line: { type: "none" } }, pptx.ShapeType.ellipse);
  // メニュー
  const mx = a.x + a.w - 2.9, my = a.y + 0.4;
  R(s, { x: mx, y: my, w: 2.6, h: 1.85, fill: { color: "FFFFFF" }, line: { color: "D0D7DE", width: 0.75 }, rectRadius: 0.05, shadow: { type: "outer", blur: 4, offset: 2, opacity: 0.2, color: "000000" } }, pptx.ShapeType.roundRect);
  ["New repository", "Import repository", "New codespace", "New gist", "New organization"].forEach((t, i) =>
    T(s, t, { x: mx + 0.2, y: my + 0.1 + i * 0.33, w: 2.2, h: 0.3, fontSize: 10.5, color: "24292F", valign: "middle" }));
  hl(s, mx + 0.12, my + 0.1, 2.36, 0.3, 2);
  T(s, "Top repositories", { x: a.x, y: a.y + 0.6, w: 2.4, h: 0.3, fontSize: 10, bold: true, color: "24292F" });
  ui(s, a.x, a.y + 1.0, 2.4, 0.3, "", { fill: "F6F8FA", line: "D0D7DE" });
  T(s, "Home", { x: a.x + 2.9, y: a.y + 2.6, w: 3, h: 0.3, fontSize: 12, bold: true, color: "24292F" });
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "右上の「＋」を押す", d: "自分のアイコン（丸）の左どなりにあります", h: 0.85 },
    { t: "「New repository」を押す", d: "新しいリポジトリを作る画面が開きます", h: 0.85 },
  ]);
  bar(s, 4.1, "tip", "画面の左上にある緑の「New」ボタンからでも同じ画面が開きます", { x, w, h: 0.95, fs: 12 });
}

// STEP2 リポジトリ作成 (2)
{
  const s = slide("第1部　STEP 2", "リポジトリ（保管用フォルダ）を作る ②", "名前を入れて、必ず Private（非公開）を選びます");
  const a = browser(s, MX, 1.7, 7.6, 4.95, "https://github.com/new");
  T(s, "Create a new repository", { x: a.x, y: a.y, w: 5, h: 0.4, fontSize: 15, bold: true, color: "24292F" });
  T(s, "Repository name *", { x: a.x, y: a.y + 0.55, w: 3, h: 0.25, fontSize: 9.5, bold: true, color: "24292F" });
  ui(s, a.x, a.y + 0.82, 3.6, 0.36, REPO, { align: "left", fs: 11 });
  hl(s, a.x, a.y + 0.82, 3.6, 0.36, 1, "r");
  T(s, "Choose visibility", { x: a.x, y: a.y + 1.45, w: 3, h: 0.25, fontSize: 9.5, bold: true, color: "24292F" });
  ui(s, a.x, a.y + 1.72, 2.2, 0.36, "🔒 Private ▾", { fs: 11, bold: true });
  hl(s, a.x, a.y + 1.72, 2.2, 0.36, 2, "r");
  T(s, "Add README", { x: a.x, y: a.y + 2.35, w: 3, h: 0.25, fontSize: 9.5, bold: true, color: "24292F" });
  ui(s, a.x, a.y + 2.62, 1.1, 0.34, "On", { fill: "DDF4FF", line: "54AEFF", fs: 10.5, bold: true });
  hl(s, a.x, a.y + 2.62, 1.1, 0.34, 3, "r");
  ui(s, a.x + a.w - 2.2, a.y + 3.45, 2.2, 0.4, "Create repository", { fill: C.GH, line: C.GH, color: "FFFFFF", bold: true });
  hl(s, a.x + a.w - 2.2, a.y + 3.45, 2.2, 0.4, 4);
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: `Repository name に「${REPO}」`, d: "半角英数字。全員同じ名前にすると、スライドどおりに進めやすい", h: 0.95 },
    { t: "「Private」を選ぶ（重要）", d: "Public だと誰でも見られます。必ず Private", h: 0.8 },
    { t: "「Add README」を On にする", d: "チェックボックスの場合はチェックを入れる。中身が空だと Claude が作業できません", h: 1.05 },
    { t: "緑の「Create repository」", d: "", h: 0.45 },
  ]);
}

// STEP2 確認
{
  const s = slide("第1部　STEP 2", "リポジトリができたことを確かめる", "この画面になれば成功です");
  const a = browser(s, MX, 1.7, 7.6, 4.3, "https://github.com/（あなたのユーザー名）/" + REPO);
  T(s, `（ユーザー名） / ${REPO}`, { x: a.x, y: a.y, w: 4.5, h: 0.36, fontSize: 13, bold: true, color: "0969DA" });
  ui(s, a.x + 4.6, a.y + 0.03, 0.8, 0.3, "Private", { fs: 9, r: 0.15 });
  hl(s, a.x + 4.6, a.y + 0.03, 0.8, 0.3, 1);
  ["Code", "Issues", "Pull requests", "Actions", "Settings"].forEach((t, i) =>
    T(s, t, { x: a.x + i * 1.35, y: a.y + 0.5, w: 1.3, h: 0.3, fontSize: 10, color: "24292F" }));
  ui(s, a.x + a.w - 1.2, a.y + 1.0, 1.2, 0.34, "<> Code ▾", { fill: C.GH, line: C.GH, color: "FFFFFF", fs: 10, bold: true });
  R(s, { x: a.x, y: a.y + 1.5, w: a.w, h: 0.42, fill: { color: "F6F8FA" }, line: { color: "D0D7DE", width: 0.5 } });
  T(s, "📄  README.md", { x: a.x + 0.15, y: a.y + 1.5, w: 3, h: 0.42, fontSize: 10.5, color: "24292F", valign: "middle" });
  hl(s, a.x + 0.05, a.y + 1.55, 2.0, 0.32, 2);
  T(s, REPO, { x: a.x + 0.15, y: a.y + 2.2, w: 4, h: 0.45, fontSize: 18, bold: true, color: "24292F" });
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "名前の横に「Private」", d: "Public と出ていたら Settings から変更（講師に声をかけてください）", h: 1.05 },
    { t: "README.md が1つある", d: "Add README を On にしたので、ファイルが1つできています", h: 0.95 },
  ]);
  bar(s, 6.35, "ok", `このページのURLは、あとで何度も開きます。Chrome のブックマークに登録しておきましょう`, { fs: 12 });
}

// STEP3 Claude アカウント
{
  const s = slide("第1部　STEP 3", "Claude のアカウントを作る", "Claude の画面は、日本語表示のこともあります（英語／日本語を併記しています）");
  const a = browser(s, MX, 1.7, 7.6, 4.95, "https://claude.ai");
  T(s, "Claude", { x: a.x, y: a.y + 0.2, w: a.w, h: 0.6, fontSize: 26, bold: true, color: "C15F3C", align: "center" });
  const fx = a.x + 1.7, fw = 3.8;
  ui(s, fx, a.y + 1.2, fw, 0.42, "G  Continue with Google（Googleで続ける）", { fs: 10.5 });
  T(s, "または", { x: fx, y: a.y + 1.75, w: fw, h: 0.25, fontSize: 9, color: "8C8C8C", align: "center" });
  ui(s, fx, a.y + 2.1, fw, 0.42, "メールアドレスを入力", { fs: 10.5, color: "8C8C8C", align: "left" });
  ui(s, fx, a.y + 2.65, fw, 0.42, "Continue with email（メールで続ける）", { fill: "1F1F1F", line: "1F1F1F", color: "FFFFFF", fs: 10.5 });
  hl(s, fx, a.y + 1.2, fw, 0.42, 1);
  hl(s, fx, a.y + 2.1, fw, 0.97, 2);
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "claude.ai を開く", d: "Google アカウントで続けるか（①）、メールアドレスで続けるか（②）を選ぶ", h: 1.05 },
    { t: "メールの場合：届いたリンク／コードで確認", d: "受信メールのボタンを押すか、6桁ほどのコードを入力", h: 1.05 },
    { t: "名前などを入力して完了", d: "利用規約は、自分で内容を読んでから同意してください", h: 0.95, plain: true, n: "✓" },
  ]);
}

// STEP4 プラン
{
  const s = slide("第1部　STEP 4", "Claude Code が使えるプランにする", "Claude Code（AIにプログラムを書いてもらう機能）は有料プランで使えます");
  const plans = [
    ["個人で始める", "Pro（または Max）", "claude.ai の画面で「Upgrade（アップグレード）」を押し、Pro を選んで支払い手続き。\n会社の経費ルールに従ってください。", C.GOLD],
    ["会社で契約している", "Team / Enterprise", "管理者に「Claude Code を使えるか」を確認。\nGitHub の連携を管理者が有効にする必要があります（→ STEP 7 の注意）。", C.GREEN],
  ];
  plans.forEach((p, i) => {
    const x = MX + i * 6.15, w = 5.95, y = 1.8;
    R(s, { x, y, w, h: 3.2, fill: { color: "FFFFFF" }, line: { color: p[3], width: 2 }, rectRadius: 0.1 }, pptx.ShapeType.roundRect);
    T(s, p[0], { x: x + 0.3, y: y + 0.25, w: w - 0.6, h: 0.35, fontSize: 13, bold: true, color: p[3] });
    T(s, p[1], { x: x + 0.3, y: y + 0.65, w: w - 0.6, h: 0.55, fontSize: 22, bold: true, color: C.INK });
    T(s, p[2], { x: x + 0.3, y: y + 1.35, w: w - 0.6, h: 1.7, fontSize: 13, color: C.BODY, lineSpacingMultiple: 1.35 });
  });
  bar(s, 5.25, "ok", "左下（または右上）のアカウント名の近くに「Pro」などのプラン名が出ていれば準備完了です", { h: 0.6 });
  bar(s, 6.0, "warn", "無料プランでは Claude Code が使えません。プランや料金は変わることがあるので、申込時に画面で確認してください", { h: 0.6 });
}

// =====================================================================
divider(2, "つなぐ ─ Claude と GitHub の連携", "STEP 5〜7　Claude Code の画面で GitHub と連携する\nSTEP 8　つながったかを、Claude 側と GitHub 側の両方で確かめる");

// STEP5 claude.ai/code
{
  const s = slide("第2部　STEP 5", "Claude Code を開き、GitHub でログインする", "Claude にログインした状態で、新しいタブで開きます");
  const a = browser(s, MX, 1.7, 7.6, 4.95, "https://claude.ai/code");
  T(s, "Claude Code", { x: a.x, y: a.y + 0.3, w: a.w, h: 0.5, fontSize: 22, bold: true, color: "1F1F1F", align: "center" });
  T(s, "Connect your GitHub account to get started", { x: a.x, y: a.y + 0.95, w: a.w, h: 0.3, fontSize: 11, color: "6B6B6B", align: "center" });
  ui(s, a.x + 1.9, a.y + 1.6, 3.4, 0.46, "⬤  Sign in with GitHub", { fill: "1F1F1F", line: "1F1F1F", color: "FFFFFF", fs: 11.5, bold: true });
  hl(s, a.x + 1.9, a.y + 1.6, 3.4, 0.46, 2);
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "アドレス欄に claude.ai/code と入力", d: "Claude にログインしていない場合は、先にログイン", h: 0.95 },
    { t: "「Sign in with GitHub」を押す", d: "「GitHub に接続」「Connect GitHub」と表示されることもあります", h: 0.95 },
    { t: "GitHub のログイン画面が出たら", d: "STEP 1 で作ったユーザー名とパスワードを、自分で入力", h: 0.95 },
  ]);
}

// STEP6 GitHub 認可
{
  const s = slide("第2部　STEP 6", "GitHub で「Claude に許可する」を押す", "GitHub の確認画面に切り替わります");
  const a = browser(s, MX, 1.7, 7.6, 4.95, "https://github.com/login/oauth/authorize?…");
  R(s, { x: a.x + 1.5, y: a.y + 0.2, w: 4.2, h: 3.6, fill: { color: "FFFFFF" }, line: { color: "D0D7DE", width: 0.75 }, rectRadius: 0.06 }, pptx.ShapeType.roundRect);
  T(s, "Authorize Claude", { x: a.x + 1.5, y: a.y + 0.4, w: 4.2, h: 0.4, fontSize: 15, bold: true, color: "24292F", align: "center" });
  T(s, "Claude by Anthropic wants to access your account", { x: a.x + 1.7, y: a.y + 0.9, w: 3.8, h: 0.5, fontSize: 9.5, color: "57606A", align: "center" });
  ["✓ Verify your GitHub identity", "✓ Know which resources you can access", "✓ Act on your behalf"].forEach((t, i) =>
    T(s, t, { x: a.x + 1.8, y: a.y + 1.5 + i * 0.32, w: 3.6, h: 0.3, fontSize: 9.5, color: "24292F" }));
  ui(s, a.x + 1.8, a.y + 2.75, 3.6, 0.42, "Authorize Claude", { fill: C.GH, line: C.GH, color: "FFFFFF", bold: true, fs: 11 });
  hl(s, a.x + 1.8, a.y + 2.75, 3.6, 0.42, 1);
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "緑の「Authorize Claude」を押す", d: "「Authorize anthropics」など、表記が少し違うことがあります", h: 0.95 },
    { t: "自動で Claude の画面に戻る", d: "戻らない場合は、claude.ai/code を開き直す", h: 0.85, plain: true, n: "→" },
  ]);
  bar(s, 4.25, "tip", "ここは「あなたが GitHub の持ち主であることを Claude に伝える」手続きです。リポジトリの中身はまだ見せていません", { x, w, h: 1.35, fs: 12 });
}

// STEP7 GitHub App
{
  const s = slide("第2部　STEP 7", "Claude にリポジトリを使う許可を出す（GitHub App）", "Private リポジトリを使うには、この設定が必要です");
  const a = browser(s, MX, 1.7, 7.6, 4.95, "https://github.com/apps/claude/installations/new");
  T(s, "Install Claude", { x: a.x, y: a.y, w: 5, h: 0.4, fontSize: 15, bold: true, color: "24292F" });
  T(s, "Install on your personal account（ユーザー名）", { x: a.x, y: a.y + 0.45, w: 6, h: 0.3, fontSize: 10, color: "57606A" });
  T(s, "◯  All repositories", { x: a.x + 0.1, y: a.y + 0.95, w: 4, h: 0.3, fontSize: 11, color: "24292F" });
  T(s, "◉  Only select repositories", { x: a.x + 0.1, y: a.y + 1.35, w: 4, h: 0.3, fontSize: 11, bold: true, color: "24292F" });
  hl(s, a.x + 0.05, a.y + 1.35, 3.0, 0.3, 1, "r");
  ui(s, a.x + 0.4, a.y + 1.85, 3.2, 0.36, "Select repositories ▾", { align: "left", fs: 10.5 });
  ui(s, a.x + 0.4, a.y + 2.3, 3.2, 0.34, `🔒 ${REPO}  ✕`, { fill: "DDF4FF", line: "54AEFF", fs: 10.5, align: "left" });
  hl(s, a.x + 0.4, a.y + 1.85, 3.2, 0.79, 2, "r");
  ui(s, a.x, a.y + 3.35, 1.6, 0.42, "Install", { fill: C.GH, line: C.GH, color: "FFFFFF", bold: true, fs: 11 });
  hl(s, a.x, a.y + 3.35, 1.6, 0.42, 3);
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "「Only select repositories」を選ぶ", d: "Claude に見せるリポジトリを限定します", h: 0.85 },
    { t: `一覧から「${REPO}」を選ぶ`, d: "名前の一部を入力すると候補が出ます", h: 0.8 },
    { t: "緑の「Install」を押す", d: "パスワードを聞かれたら自分で入力", h: 0.8 },
  ]);
  bar(s, 4.7, "tip", "この画面が自動で出ない場合は、上のURLを直接開きます。「Skip」は押さずに設定してください", { x, w, h: 0.95, fs: 11.5 });
  bar(s, 5.8, "warn", "会社の Team / Enterprise プランでは、管理者が「GitHub コネクタ」を有効にしていないと先に進めません", { x, w, h: 0.95, fs: 11.5, label: "⚠" });
}

// 環境の作成
{
  const s = slide("第2部　STEP 7", "（表示されたら）クラウド環境を作る", "Claude がプログラムを書くための「作業部屋」です。表示されなければ飛ばしてOK");
  const a = browser(s, MX, 1.7, 7.6, 4.6, "https://claude.ai/code");
  T(s, "Create your first cloud environment", { x: a.x, y: a.y + 0.1, w: a.w, h: 0.4, fontSize: 14, bold: true, color: "1F1F1F", align: "center" });
  T(s, "Name", { x: a.x + 1.6, y: a.y + 0.75, w: 3, h: 0.25, fontSize: 9.5, color: "6B6B6B" });
  ui(s, a.x + 1.6, a.y + 1.0, 4.0, 0.38, "Default", { align: "left", fs: 10.5 });
  T(s, "Network access", { x: a.x + 1.6, y: a.y + 1.6, w: 3, h: 0.25, fontSize: 9.5, color: "6B6B6B" });
  ui(s, a.x + 1.6, a.y + 1.85, 4.0, 0.38, "Trusted ▾", { align: "left", fs: 10.5 });
  ui(s, a.x + 1.6, a.y + 2.75, 4.0, 0.44, "Create & finish", { fill: "1F1F1F", line: "1F1F1F", color: "FFFFFF", bold: true, fs: 11 });
  hl(s, a.x + 1.6, a.y + 2.75, 4.0, 0.44, 1);
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "設定は変えずに「Create & finish」", d: "名前・ネットワーク設定は初期値のままでOK", h: 0.95 },
  ]);
  card(s, x, 3.0, w, 2.4, "表示されるかどうか", "Pro / Max プラン：\n自動で「Default」が作られるので、この画面は出ないことが多い\n\nTeam / Enterprise：\nこの画面が出たら上の手順で作成", { fs: 11.5 });
}

// STEP8 接続確認（Claude 側）
{
  const s = slide("第2部　STEP 8", "つながったか確かめる ① Claude 側", "リポジトリを選んで、最初のお願いをしてみます");
  const a = browser(s, MX, 1.7, 7.0, 3.1, "https://claude.ai/code");
  T(s, "What should we work on?", { x: a.x, y: a.y, w: a.w, h: 0.4, fontSize: 14, bold: true, color: "1F1F1F", align: "center" });
  ui(s, a.x + 0.3, a.y + 0.55, a.w - 0.6, 0.85, "ここにプロンプトを貼り付け", { align: "left", color: "8C8C8C", fs: 10.5 });
  hl(s, a.x + 0.3, a.y + 0.55, a.w - 0.6, 0.85, 2);
  ui(s, a.x + 0.3, a.y + 1.6, 2.4, 0.32, `📁 ${REPO} ▾`, { fs: 10, r: 0.15 });
  hl(s, a.x + 0.3, a.y + 1.6, 2.4, 0.32, 1);
  ui(s, a.x + 2.85, a.y + 1.6, 1.1, 0.32, "main ▾", { fs: 10, r: 0.15 });
  ui(s, a.x + a.w - 0.75, a.y + 1.6, 0.45, 0.32, "↑", { fill: "C15F3C", line: "C15F3C", color: "FFFFFF", bold: true, r: 0.15 });
  hl(s, a.x + a.w - 0.75, a.y + 1.6, 0.45, 0.32, 3);
  const x = 8.0, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: `入力欄の下で「${REPO}」を選ぶ`, d: "リポジトリ名の部分を押すと一覧が出ます", h: 0.8 },
    { t: "下のプロンプトを貼り付け", d: "", h: 0.42 },
    { t: "送信（↑）または Enter", d: "", h: 0.42 },
  ], { fs: 13 });
  promptBox(s, MX, 5.15, 7.0, 1.7, P.check, 11.5);
  card(s, x, 4.05, w, 2.8, "こう返ってくれば成功", "「README.md が1つあります」\n「中身は # " + REPO + " だけです」\nといった返事が来る\n\n→ Claude があなたのリポジトリを読めています", { fs: 12, line: C.OK, hcolor: C.OK });
}

// STEP8 接続確認（GitHub 側）
{
  const s = slide("第2部　STEP 8", "つながったか確かめる ② GitHub 側", "GitHub の設定画面でも、Claude が登録されているか確かめます");
  const a = browser(s, MX, 1.7, 7.6, 4.95, "https://github.com/settings/installations");
  // 左メニュー
  R(s, { x: a.x, y: a.y, w: 2.0, h: 3.85, fill: { color: "F6F8FA" }, line: { type: "none" } });
  ["Public profile", "Account", "Appearance", "Notifications", "Billing", "Emails", "Password", "SSH keys", "Applications"].forEach((t, i) =>
    T(s, t, { x: a.x + 0.15, y: a.y + 0.08 + i * 0.36, w: 1.8, h: 0.3, fontSize: 9.5, color: "24292F", bold: t === "Applications", valign: "middle" }));
  hl(s, a.x + 0.1, a.y + 0.08 + 8 * 0.36, 1.6, 0.3, 2, "r");
  const cx = a.x + 2.3;
  ["Installed GitHub Apps", "Authorized GitHub Apps", "Authorized OAuth Apps"].forEach((t, i) =>
    T(s, t, { x: cx + i * 1.7, y: a.y + 0.05, w: 1.65, h: 0.42, fontSize: 9, color: "24292F", bold: i === 0, valign: "middle" }));
  hl(s, cx, a.y + 0.1, 1.55, 0.32, 3);
  R(s, { x: cx, y: a.y + 0.75, w: a.w - 2.3, h: 0.6, fill: { color: "FFFFFF" }, line: { color: "D0D7DE", width: 0.75 } });
  T(s, "Claude", { x: cx + 0.15, y: a.y + 0.75, w: 2, h: 0.6, fontSize: 11, bold: true, color: "24292F", valign: "middle" });
  ui(s, cx + a.w - 3.4, a.y + 0.88, 0.95, 0.34, "Configure", { fs: 9.5 });
  hl(s, cx + a.w - 3.4, a.y + 0.88, 0.95, 0.34, 4);
  T(s, "Repository access", { x: cx, y: a.y + 1.7, w: 3, h: 0.3, fontSize: 10, bold: true, color: "24292F" });
  ui(s, cx, a.y + 2.05, 3.2, 0.34, `🔒 （ユーザー名）/${REPO}`, { fs: 9.5, align: "left" });
  hl(s, cx, a.y + 2.05, 3.2, 0.34, 5, "r");
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "右上の自分のアイコン →「Settings」", d: "", h: 0.42 },
    { t: "左メニューの「Applications」", d: "下のほうにあります", h: 0.62 },
    { t: "「Installed GitHub Apps」タブ", d: "", h: 0.42 },
    { t: "Claude の「Configure」", d: "", h: 0.42 },
    { t: `「${REPO}」が出ていればOK`, d: "出ていなければ、ここで追加して「Save」", h: 0.62 },
  ], { fs: 13, gap: 0.14 });
  bar(s, 5.4, "ok", "① と ② の両方OKなら、連携は完了です", { x, w, h: 0.8, fs: 12 });
}

// 連携トラブル
{
  const s = slide("第2部", "連携がうまくいかないとき", "よくある症状と、直し方です");
  const rows = [
    ["リポジトリの一覧に「" + REPO + "」が出てこない", "Claude の GitHub App に、そのリポジトリが許可されていない", "STEP 7 の URL（github.com/apps/claude/installations/new）を開き、" + REPO + " を追加して Install / Save"],
    ["「GitHub access is required…」と表示される", "会社の Team / Enterprise プランで、GitHub コネクタが無効", "管理者（Owner）に「組織の設定 → Connectors で GitHub を有効にしてほしい」と依頼"],
    ["Claude Code の画面が開かない／使えない", "無料プラン、または会社のプランで権限がない", "STEP 4 に戻ってプランを確認。会社契約なら管理者に確認"],
    ["接続確認で「ファイルが見つからない」", "リポジトリが空（README を作らなかった）", "GitHub のリポジトリ画面で「Add file → Create new file」から README.md を作る"],
  ];
  const cols = [3.6, 3.7, 4.83];
  let y = 1.75;
  const hx = [MX, MX + cols[0], MX + cols[0] + cols[1]];
  ["症状", "原因", "直し方"].forEach((h, i) => {
    R(s, { x: hx[i], y, w: cols[i], h: 0.45, fill: { color: C.DARK }, line: { color: "FFFFFF", width: 1 } });
    T(s, h, { x: hx[i] + 0.15, y, w: cols[i] - 0.3, h: 0.45, fontSize: 13, bold: true, color: "FFFFFF", valign: "middle" });
  });
  y += 0.45;
  rows.forEach((r, ri) => {
    r.forEach((t, i) => {
      R(s, { x: hx[i], y, w: cols[i], h: 1.12, fill: { color: ri % 2 ? "FFFFFF" : C.CARD }, line: { color: C.LINE, width: 0.75 } });
      T(s, t, { x: hx[i] + 0.15, y, w: cols[i] - 0.3, h: 1.12, fontSize: 12, bold: i === 0, color: i === 0 ? C.INK : C.BODY, valign: "middle", lineSpacingMultiple: 1.15 });
    });
    y += 1.12;
  });
}

// =====================================================================
divider(3, "作る ─ プロンプトでツールを完成させる", "STEP 9　Claude Code の画面の見方\nSTEP 10〜12　プロンプト① → 確認 → 取り込み → ダウンロード\nSTEP 13〜15　プロンプト②③で機能を追加する");

// STEP9 画面の見方
{
  const s = slide("第3部　STEP 9", "Claude Code の画面の見方", "お願いを送る前に、この3か所を確認します");
  const a = browser(s, MX, 1.7, 7.6, 3.4, "https://claude.ai/code");
  T(s, "What should we work on?", { x: a.x, y: a.y, w: a.w, h: 0.4, fontSize: 14, bold: true, color: "1F1F1F", align: "center" });
  ui(s, a.x + 0.3, a.y + 0.55, a.w - 0.6, 1.0, "ここにプロンプトを貼り付け", { align: "left", color: "8C8C8C", fs: 10.5 });
  hl(s, a.x + 0.3, a.y + 0.55, a.w - 0.6, 1.0, 1);
  ui(s, a.x + 0.3, a.y + 1.8, 2.3, 0.32, `📁 ${REPO} ▾`, { fs: 10, r: 0.15 });
  hl(s, a.x + 0.3, a.y + 1.8, 2.3, 0.32, 2);
  ui(s, a.x + 2.75, a.y + 1.8, 1.1, 0.32, "main ▾", { fs: 10, r: 0.15 });
  ui(s, a.x + 4.0, a.y + 1.8, 1.55, 0.32, "Accept edits ▾", { fs: 10, r: 0.15 });
  hl(s, a.x + 4.0, a.y + 1.8, 1.55, 0.32, 3);
  ui(s, a.x + a.w - 0.75, a.y + 1.8, 0.45, 0.32, "↑", { fill: "C15F3C", line: "C15F3C", color: "FFFFFF", bold: true, r: 0.15 });
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "入力欄", d: "プロンプトを貼る場所。Enter で送信、改行は Shift + Enter", h: 0.85 },
    { t: `リポジトリ：${REPO}`, d: "毎回ここが正しいか確認。違うと別の場所に作られます", h: 0.85 },
    { t: "モード：Accept edits", d: "「ファイルの変更を任せる」設定。今日はこれを選びます", h: 0.85 },
  ]);
  card(s, MX, 5.5, 7.6, 1.35, "左側の一覧（セッション）", "送ったお願いは1件ずつ左側に残ります。あとから開いて、続きのお願いもできます", { fs: 12 });
  bar(s, 5.5, "tip", "Plan を選ぶと「計画だけ」を作ります。慣れてきたら、先に Plan で計画を確かめる使い方もおすすめ", { x, w, h: 1.35, fs: 12 });
}

// STEP10 プロンプト①
{
  const s = slide("第3部　STEP 10", "プロンプト①：ツールの土台と2つのグラフを作る", `リポジトリ「${REPO}」とモードを確認して送信。コピーは prompts.md（最後のスライド参照）の 📋 ボタンが確実`);
  promptBox(s, MX, 1.65, W - MX * 2, 5.25, P.p1, 8.8, 1.05);
}

// STEP10 待つ
{
  const s = slide("第3部　STEP 10", "送信したら：Claude の作業を見守る", "5〜15分ほどかかります。途中で質問されたら答えます");
  const a = browser(s, MX, 1.7, 7.6, 4.95, "https://claude.ai/code/session_…");
  R(s, { x: a.x, y: a.y, w: a.w, h: 0.55, fill: { color: "F4F1EC" }, line: { type: "none" }, rectRadius: 0.06 }, pptx.ShapeType.roundRect);
  T(s, "あなた：あなたは、不動産会社の営業担当が使う社内ツールを…", { x: a.x + 0.15, y: a.y, w: a.w - 0.3, h: 0.55, fontSize: 10, color: "1F1F1F", valign: "middle" });
  const lines = ["● リポジトリの中身を確認しています", "● npm install chart.js papaparse を実行", "● lib/ にライブラリをコピー", "● index.html を作成", "● app.js を作成", "● sample-data.js を作成"];
  lines.forEach((t, i) => T(s, t, { x: a.x + 0.15, y: a.y + 0.75 + i * 0.32, w: a.w - 0.3, h: 0.3, fontSize: 10, color: "4B5563" }));
  R(s, { x: a.x, y: a.y + 2.8, w: a.w, h: 0.9, fill: { color: "FFFFFF" }, line: { color: "E5E7EB", width: 0.75 }, rectRadius: 0.06 }, pptx.ShapeType.roundRect);
  T(s, "Claude：ツールを作成しました。作ったファイルは次のとおりです…\n確かめ方：index.html をダブルクリックし「サンプルデータで試す」を…", { x: a.x + 0.15, y: a.y + 2.85, w: a.w - 0.3, h: 0.8, fontSize: 10, color: "1F1F1F", lineSpacingMultiple: 1.2 });
  hl(s, a.x, a.y + 2.8, a.w, 0.9, 1);
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "作業の様子が流れる", d: "ファイルを作ったり、コマンドを実行したりします。見ているだけでOK", h: 1.0, plain: true, n: "…" },
    { t: "最後の説明を読む", d: "作ったファイルと確かめ方が日本語で書かれます", h: 0.85 },
  ]);
  bar(s, 4.2, "tip", "途中で「◯◯してよいですか？」と聞かれたら、内容を読んで「はい」と答えます。分からなければ講師へ", { x, w, h: 1.2, fs: 12 });
  bar(s, 5.55, "warn", "画面を閉じても作業は続きます。左側の一覧から、あとで開き直せます", { x, w, h: 0.95, fs: 12, label: "⚠" });
}

// STEP11 Create PR
{
  const s = slide("第3部　STEP 11", "作業が終わったら：変更を確認して「Create PR」", "Claude が作ったファイルを、GitHub に取り込む準備をします");
  const a = browser(s, MX, 1.7, 7.6, 4.95, "https://claude.ai/code/session_…");
  T(s, "Claude：ツールを作成しました。", { x: a.x, y: a.y, w: 4, h: 0.3, fontSize: 10.5, color: "1F1F1F" });
  ui(s, a.x + a.w - 1.3, a.y - 0.02, 1.3, 0.34, "+412  -0", { fs: 10.5, bold: true, color: C.GH });
  hl(s, a.x + a.w - 1.3, a.y - 0.02, 1.3, 0.34, 1);
  R(s, { x: a.x, y: a.y + 0.6, w: a.w, h: 3.2, fill: { color: "FFFFFF" }, line: { color: "D0D7DE", width: 0.75 } });
  R(s, { x: a.x, y: a.y + 0.6, w: a.w, h: 0.48, fill: { color: "F6F8FA" }, line: { type: "none" } });
  T(s, "変更されたファイル（4）", { x: a.x + 0.15, y: a.y + 0.6, w: 3, h: 0.48, fontSize: 10, bold: true, color: "24292F", valign: "middle" });
  ui(s, a.x + a.w - 1.5, a.y + 0.67, 1.35, 0.34, "Create PR", { fill: C.GH, line: C.GH, color: "FFFFFF", bold: true, fs: 10.5 });
  hl(s, a.x + a.w - 1.5, a.y + 0.67, 1.35, 0.34, 2);
  ["index.html", "app.js", "sample-data.js", "lib/chart.umd.min.js …"].forEach((t, i) =>
    T(s, "＋  " + t, { x: a.x + 0.2, y: a.y + 1.25 + i * 0.36, w: 4, h: 0.32, fontSize: 10, color: C.GH }));
  R(s, { x: a.x + 0.2, y: a.y + 2.75, w: a.w - 0.4, h: 0.85, fill: { color: "E6FFEC" }, line: { type: "none" } });
  T(s, "+ <h1>残存年数別 資産価値グラフ</h1>\n+ <button id=\"btnNew\">新築データを読み込む</button>", { x: a.x + 0.35, y: a.y + 2.8, w: a.w - 0.7, h: 0.75, fontSize: 9, color: "1A7F37" });
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "「+412 -0」のような数字を押す", d: "変更の量です。押すと変更内容（差分）が開きます", h: 0.95 },
    { t: "上部の「Create PR」を押す", d: "プルリクエスト（取り込みの確認依頼）が GitHub に作られます", h: 0.95 },
    { t: "「View PR」から GitHub へ", d: "STEP 12 に進みます", h: 0.8, plain: true, n: "→" },
  ]);
  bar(s, 5.3, "tip", "緑の行＝追加された部分。中身を全部理解する必要はありません", { x, w, h: 0.95, fs: 12 });
}

// STEP12 Merge
{
  const s = slide("第3部　STEP 12", "GitHub で取り込む（マージする）", "プルリクエストの画面で、緑のボタンを2回押します");
  const a = browser(s, MX, 1.7, 7.6, 4.95, `https://github.com/（ユーザー名）/${REPO}/pull/1`);
  T(s, "残存年数別 資産価値グラフのツールを作成  #1", { x: a.x, y: a.y, w: a.w, h: 0.4, fontSize: 13, bold: true, color: "24292F" });
  ui(s, a.x, a.y + 0.48, 0.75, 0.3, "Open", { fill: C.GH, line: C.GH, color: "FFFFFF", fs: 9.5, r: 0.15 });
  ["Conversation", "Commits", "Files changed"].forEach((t, i) => T(s, t, { x: a.x + i * 1.6, y: a.y + 0.95, w: 1.5, h: 0.3, fontSize: 10, color: "24292F" }));
  R(s, { x: a.x, y: a.y + 1.5, w: a.w, h: 1.7, fill: { color: "FFFFFF" }, line: { color: "D0D7DE", width: 0.75 }, rectRadius: 0.05 }, pptx.ShapeType.roundRect);
  T(s, "✓  No conflicts with base branch", { x: a.x + 0.2, y: a.y + 1.6, w: 5, h: 0.35, fontSize: 10.5, bold: true, color: "24292F" });
  ui(s, a.x + 0.2, a.y + 2.15, 2.3, 0.42, "Merge pull request ▾", { fill: C.GH, line: C.GH, color: "FFFFFF", bold: true, fs: 10.5 });
  hl(s, a.x + 0.2, a.y + 2.15, 2.3, 0.42, 1);
  ui(s, a.x + 0.2, a.y + 3.45, 2.3, 0.42, "Confirm merge", { fill: C.GH, line: C.GH, color: "FFFFFF", bold: true, fs: 10.5 });
  ui(s, a.x + 2.65, a.y + 3.45, 0.95, 0.42, "Cancel", { fs: 10.5 });
  hl(s, a.x + 0.2, a.y + 3.45, 2.3, 0.42, 2);
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "緑の「Merge pull request」", d: "画面を下にスクロールすると出てきます", h: 0.8 },
    { t: "緑の「Confirm merge」", d: "これで正式に取り込まれます", h: 0.8 },
    { t: "紫の「Merged」になれば完了", d: "", h: 0.42, plain: true, n: "✓" },
  ]);
  bar(s, 4.6, "tip", "GitHub のリポジトリを直接開いて「Pull requests」タブ → 一覧から選んでも、同じ画面に行けます", { x, w, h: 1.2, fs: 12 });
}

// STEP12 Download
{
  const s = slide("第3部　STEP 12", "ツールをPCにダウンロードして開く", "GitHub に保管されたツールを、自分のPCに持ってきます");
  const a = browser(s, MX, 1.7, 7.6, 4.95, `https://github.com/（ユーザー名）/${REPO}`);
  T(s, `（ユーザー名） / ${REPO}`, { x: a.x, y: a.y, w: 4.5, h: 0.36, fontSize: 13, bold: true, color: "0969DA" });
  ui(s, a.x + a.w - 1.3, a.y + 0.5, 1.3, 0.36, "<> Code ▾", { fill: C.GH, line: C.GH, color: "FFFFFF", fs: 10.5, bold: true });
  hl(s, a.x + a.w - 1.3, a.y + 0.5, 1.3, 0.36, 1);
  R(s, { x: a.x + a.w - 3.2, y: a.y + 1.0, w: 3.2, h: 1.7, fill: { color: "FFFFFF" }, line: { color: "D0D7DE", width: 0.75 }, rectRadius: 0.05, shadow: { type: "outer", blur: 4, offset: 2, opacity: 0.2, color: "000000" } }, pptx.ShapeType.roundRect);
  T(s, "Local　|　Codespaces", { x: a.x + a.w - 3.05, y: a.y + 1.08, w: 2.9, h: 0.3, fontSize: 9.5, color: "24292F" });
  T(s, "Clone  HTTPS  SSH …", { x: a.x + a.w - 3.05, y: a.y + 1.5, w: 2.9, h: 0.3, fontSize: 9, color: "57606A" });
  T(s, "📦  Download ZIP", { x: a.x + a.w - 3.05, y: a.y + 2.2, w: 2.9, h: 0.34, fontSize: 10.5, bold: true, color: "24292F", valign: "middle" });
  hl(s, a.x + a.w - 3.05, a.y + 2.2, 2.9, 0.34, 2);
  ["📁 lib", "📁 sample-data", "📄 README.md", "📄 app.js", "📄 index.html", "📄 sample-data.js"].forEach((t, i) =>
    T(s, t, { x: a.x + 0.1, y: a.y + 1.05 + i * 0.4, w: 3, h: 0.34, fontSize: 10.5, color: "24292F" }));
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "緑の「<> Code」を押す", d: "", h: 0.42 },
    { t: "「Download ZIP」を押す", d: "ダウンロードフォルダに ZIP が保存されます", h: 0.8 },
    { t: "ZIP を展開する", d: "Windows：右クリック →「すべて展開」\nMac：ダブルクリック", h: 1.05 },
    { t: "フォルダの中の index.html をダブルクリック", d: "Chrome でツールが開きます", h: 0.95 },
  ]);
}

// STEP12 動作確認
{
  const s = slide("第3部　STEP 12", "サンプルデータで動作を確かめる", "まだ自社データは使いません。架空のサンプルで動くかを確かめます");
  I(s, SHOT + "c_charts.png", { x: MX, y: 1.75, w: 6.0, h: 6.0 * 890 / 1100 });
  R(s, { x: MX, y: 1.75, w: 6.0, h: 6.0 * 890 / 1100, fill: { type: "none" }, line: { color: C.LINE, width: 1 } });
  const x = 7.0, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "「サンプルデータで試す」を押す", d: "", h: 0.42 },
    { t: "2つのグラフが出る", d: "新築比のグラフ（緑）と坪単価のグラフ（青）", h: 0.8 },
    { t: "件数が表示される", d: "「新築データ 60行（10物件）／中古データ 188行」など", h: 0.8 },
  ]);
  bar(s, 4.4, "ok", "見た目は左と多少違ってもOK。ボタン・グラフ・件数がそろっていれば成功です", { x, w, h: 0.95, fs: 12 });
  bar(s, 5.5, "warn", "何も出ない・エラーになるときは → 第5部「うまく動かないとき」のプロンプトで Claude に直してもらう", { x, w, h: 1.3, fs: 12, label: "⚠" });
}

// STEP13 プロンプト②
{
  const s = slide("第3部　STEP 13", "プロンプト②：物件一覧と絞り込みを追加する", "同じセッションの入力欄に続けて貼るか、新しく始める場合もリポジトリを確認");
  promptBox(s, MX, 1.7, 7.4, 4.95, P.p2, 10.2);
  I(s, SHOT + "c_le5.png", { x: 8.3, y: 1.7, w: W - MX - 8.3, h: (W - MX - 8.3) * 500 / 1100 });
  R(s, { x: 8.3, y: 1.7, w: W - MX - 8.3, h: (W - MX - 8.3) * 500 / 1100, fill: { type: "none" }, line: { color: C.LINE, width: 1 } });
  T(s, "↑ 完成イメージ：駅徒歩「5分以内」で絞り込んだところ", { x: 8.3, y: 3.75, w: W - MX - 8.3, h: 0.3, fontSize: 10, color: C.MUTE });
  card(s, 8.3, 4.25, W - MX - 8.3, 2.4, "送信したあとは", "STEP 11〜12 と同じです\n① 変更の数字 →「Create PR」\n② GitHub で Merge → Confirm\n③ Download ZIP → 展開 → index.html", { fs: 12 });
}

// STEP14 プロンプト③
{
  const s = slide("第3部　STEP 14", "プロンプト③：物件ごとの詳細画面を追加する", "一覧のマンション名をクリックすると、その物件の推移が見られるようにします");
  promptBox(s, MX, 1.7, 7.4, 3.9, P.p3, 10.5);
  I(s, SHOT + "c_detail.png", { x: 8.3, y: 1.7, w: W - MX - 8.3, h: (W - MX - 8.3) * 650 / 1100 });
  R(s, { x: 8.3, y: 1.7, w: W - MX - 8.3, h: (W - MX - 8.3) * 650 / 1100, fill: { type: "none" }, line: { color: C.LINE, width: 1 } });
  T(s, "↑ 完成イメージ：物件の詳細画面", { x: 8.3, y: 4.45, w: W - MX - 8.3, h: 0.3, fontSize: 10, color: C.MUTE });
  bar(s, 5.85, "ok", "一覧でマンション名を押す → 詳細が出る →「← 一覧に戻る」で戻れる。これで完成です！", { h: 0.6 });
}

// STEP15 更新サイクル
{
  const s = slide("第3部　STEP 15", "2回目以降の直し方は、いつも同じ5ステップ", "機能を足すときも、見た目を直すときも、この順番です");
  const st = [["1", "プロンプトを送る", "Claude Code"], ["2", "Create PR", "Claude Code"], ["3", "Merge → Confirm", "GitHub"], ["4", "Download ZIP", "GitHub"], ["5", "index.html を開く", "あなたのPC"]];
  const bw = 2.15, gap = 0.3, x0 = MX + 0.05;
  st.forEach((t, i) => {
    const x = x0 + i * (bw + gap), y = 2.0;
    R(s, { x, y, w: bw, h: 2.0, fill: { color: C.CARD }, line: { color: C.GOLD, width: 1.5 }, rectRadius: 0.1 }, pptx.ShapeType.roundRect);
    R(s, { x: x + bw / 2 - 0.3, y: y + 0.2, w: 0.6, h: 0.6, fill: { color: C.GOLD }, line: { type: "none" } }, pptx.ShapeType.ellipse);
    T(s, t[0], { x: x + bw / 2 - 0.3, y: y + 0.2, w: 0.6, h: 0.6, fontSize: 18, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    T(s, t[1], { x: x + 0.1, y: y + 0.95, w: bw - 0.2, h: 0.5, fontSize: 15, bold: true, color: C.INK, align: "center", valign: "middle" });
    T(s, t[2], { x: x + 0.1, y: y + 1.45, w: bw - 0.2, h: 0.35, fontSize: 11, color: C.MUTE, align: "center" });
    if (i < 4) T(s, "▶", { x: x + bw, y: y + 0.75, w: gap, h: 0.5, fontSize: 14, color: C.MUTE, align: "center", valign: "middle" });
  });
  bar(s, 4.4, "warn", "ダウンロードし直したら、古いフォルダは削除するか、別の名前にしておくと混ざりません", { h: 0.6, label: "⚠" });
  bar(s, 5.2, "tip", "前の作業の続きなら、左側の一覧から同じセッションを開いてお願いすると、話が通じやすくなります", { h: 0.6 });
  bar(s, 6.0, "ok", "1回のお願いは「1つのテーマ」に絞ると、うまくいきやすい（例：一覧を足す → 次に色を変える）", { h: 0.6 });
}

// =====================================================================
divider(4, "自社データを入れる", "STEP 16　新築データ・中古データを CSV の形にそろえる\nSTEP 17　Excel から CSV で保存する\nSTEP 18　ツールの画面で読み込む（データはPCの外に出ません）");

// STEP16 新築CSV
function csvTable(s, x, y, w, cols, rows, widths) {
  const tot = widths.reduce((a, b) => a + b, 0);
  let cx = x;
  cols.forEach((c, i) => {
    const cw = (widths[i] / tot) * w;
    R(s, { x: cx, y, w: cw, h: 0.42, fill: { color: C.GREEN }, line: { color: "FFFFFF", width: 0.75 } });
    T(s, c, { x: cx, y, w: cw, h: 0.42, fontSize: 11, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    rows.forEach((r, ri) => {
      R(s, { x: cx, y: y + 0.42 + ri * 0.38, w: cw, h: 0.38, fill: { color: ri % 2 ? "FFFFFF" : C.CARD }, line: { color: C.LINE, width: 0.75 } });
      T(s, r[i], { x: cx, y: y + 0.42 + ri * 0.38, w: cw, h: 0.38, fontSize: 10.5, color: C.BODY, align: "center", valign: "middle" });
    });
    cx += cw;
  });
  return y + 0.42 + rows.length * 0.38;
}
{
  const s = slide("第4部　STEP 16", "新築データの形（1住戸＝1行）", "1行目（列名）は、この9つを左からこの順番で。2行目からデータです");
  const yy = csvTable(s, MX, 1.75, W - MX * 2,
    ["物件名", "区", "最寄駅", "駅徒歩分", "竣工年", "借地期限年", "総戸数", "価格", "専有面積"],
    [["〇〇レジデンス", "文京区", "湯島", "3", "2012", "2082", "120", "5800", "65.2"],
     ["〇〇レジデンス", "文京区", "湯島", "3", "2012", "2082", "120", "7450", "78.4"],
     ["△△タワー", "港区", "田町", "4", "2008", "2078", "600", "6900", "70.1"]],
    [2.2, 1.1, 1.2, 1.1, 1.0, 1.25, 1.0, 1.0, 1.1]);
  const rules = [
    ["価格", "万円の数字だけ（5,800万円 → 5800）"],
    ["専有面積", "㎡の数字だけ（65.2㎡ → 65.2）"],
    ["駅徒歩分", "分の数字だけ（徒歩3分 → 3）。分からなければ空欄"],
    ["竣工年・借地期限年", "西暦4桁（2012、2082）"],
    ["物件名", "中古データと1文字も違わない名前にそろえる（全角・半角、スペースも）"],
  ];
  let y = yy + 0.35;
  rules.forEach((r) => {
    T(s, [{ text: r[0] + "　", options: { bold: true, color: C.GREEN } }, { text: r[1], options: { color: C.INK } }], { x: MX + 0.1, y, w: W - MX * 2 - 0.2, h: 0.4, fontSize: 13.5, valign: "middle" });
    y += 0.45;
  });
  bar(s, 6.3, "tip", "物件の「区・駅・年・戸数」は全住戸で同じ値を繰り返します。新築坪単価は、その物件の全住戸の平均で計算されます", { h: 0.55, fs: 12 });
}

// STEP16 中古CSV
{
  const s = slide("第4部　STEP 16", "中古データの形（1成約＝1行）", "1行目（列名）は、この5つを左からこの順番で");
  const yy = csvTable(s, MX, 1.75, W - MX * 2,
    ["物件名", "成約年月", "価格", "専有面積", "階"],
    [["〇〇レジデンス", "2024/05", "6200", "65.2", "7"],
     ["〇〇レジデンス", "2025/02", "8100", "78.4", "12"],
     ["△△タワー", "2023/11", "7300", "70.1", "21"]],
    [2.6, 1.6, 1.4, 1.4, 1.0]);
  const rules = [
    ["物件名", "新築データの物件名と完全に同じ名前（ここで2つのデータをつなぎます）"],
    ["成約年月", "2024/05 の形（年だけが使われます）"],
    ["価格・専有面積", "新築データと同じく、万円・㎡の数字だけ"],
    ["階", "分からなければ空欄でOK（今回の計算には使いません）"],
  ];
  let y = yy + 0.35;
  rules.forEach((r) => {
    T(s, [{ text: r[0] + "　", options: { bold: true, color: C.GREEN } }, { text: r[1], options: { color: C.INK } }], { x: MX + 0.1, y, w: W - MX * 2 - 0.2, h: 0.4, fontSize: 13.5, valign: "middle" });
    y += 0.45;
  });
  bar(s, 6.05, "warn", "自社データの列名や単位が違う場合は、無理に作り直さず「列名が違うとき」のプロンプト（第4部の最後）でツールを合わせられます", { h: 0.75, fs: 12, label: "⚠" });
}

// STEP17 Excel から CSV
{
  const s = slide("第4部　STEP 17", "Excel から CSV で保存する", "いちばん簡単なのは、ツールの「テンプレートCSV」ボタンでひな形をもらう方法です");
  const a = browser(s, MX, 1.7, 7.6, 4.95, "Excel ─ 名前を付けて保存");
  T(s, "名前を付けて保存", { x: a.x, y: a.y, w: 4, h: 0.4, fontSize: 14, bold: true, color: "1F1F1F" });
  T(s, "ファイル名", { x: a.x, y: a.y + 0.6, w: 2, h: 0.25, fontSize: 9.5, color: "57606A" });
  ui(s, a.x, a.y + 0.87, 4.6, 0.38, "shinchiku.csv", { align: "left", fs: 11 });
  hl(s, a.x, a.y + 0.87, 4.6, 0.38, 2, "r");
  T(s, "ファイルの種類", { x: a.x, y: a.y + 1.5, w: 2, h: 0.25, fontSize: 9.5, color: "57606A" });
  ui(s, a.x, a.y + 1.77, 4.6, 0.38, "CSV UTF-8 (コンマ区切り) (*.csv) ▾", { align: "left", fs: 11, bold: true });
  hl(s, a.x, a.y + 1.77, 4.6, 0.38, 3, "r");
  ui(s, a.x + a.w - 1.4, a.y + 3.4, 1.4, 0.4, "保存", { fill: "107C41", line: "107C41", color: "FFFFFF", bold: true });
  hl(s, a.x + a.w - 1.4, a.y + 3.4, 1.4, 0.4, 4);
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "Excel で「ファイル」→「名前を付けて保存」", d: "Mac は「ファイル」→「名前を付けて保存…」", h: 0.95 },
    { t: "ファイル名を半角英数字で", d: "新築：shinchiku.csv\n中古：chuko.csv", h: 0.95 },
    { t: "種類「CSV UTF-8（コンマ区切り）」", d: "ふつうの「CSV（コンマ区切り）」でも読めます", h: 0.95 },
    { t: "「保存」を押す", d: "警告が出たら「はい」「このまま保存」", h: 0.8 },
  ]);
}

// STEP18 読み込み
{
  const s = slide("第4部　STEP 18", "ツールの画面で、自社データを読み込む", "ここで初めて自社データを使います。データはこのPCの中だけで処理されます");
  I(s, SHOT + "c_top.png", { x: MX, y: 1.7, w: 7.6, h: 7.6 * 420 / 1100 });
  R(s, { x: MX, y: 1.7, w: 7.6, h: 7.6 * 420 / 1100, fill: { type: "none" }, line: { color: C.LINE, width: 1 } });
  const k = 7.6 / 1100, ox = MX, oy = 1.7;
  hl(s, ox + 28 * k, oy + 223 * k, 167 * k, 38 * k, 1);
  hl(s, ox + 205 * k, oy + 223 * k, 166 * k, 38 * k, 2);
  hl(s, ox + 28 * k, oy + 271 * k, 280 * k, 18 * k, 3, "r");
  T(s, "↑ ツールの画面（index.html）", { x: MX, y: 4.65, w: 7.6, h: 0.3, fontSize: 10, color: C.MUTE });
  const x = 8.55, w = W - MX - x;
  steps(s, x, 1.75, w, [
    { t: "「新築データを読み込む」", d: "shinchiku.csv を選んで「開く」", h: 0.8 },
    { t: "「中古データを読み込む」", d: "chuko.csv を選んで「開く」", h: 0.8 },
    { t: "件数の表示を確認", d: "行数・物件数が、Excel の行数と合っているか", h: 0.8 },
  ]);
  bar(s, 5.1, "ok", "グラフ・一覧が自社データに切り替われば完成です。エリアや駅徒歩で絞り込んでみましょう", { h: 0.6 });
  bar(s, 5.85, "tip", "ツールを閉じるとデータは消えます（保存されません）。次に開いたときは、もう一度読み込みます", { h: 0.6 });
}

// うまく読めないとき
{
  const s = slide("第4部　STEP 18", "うまく読み込めないときのチェック", "画面に出るメッセージで、原因が分かります");
  const rows = [
    ["「新築データの列名が足りません：駅徒歩分」", "1行目の列名がちがう／足りない", "列名を前のスライドどおりに直す。または次のスライドのプロンプトでツールを合わせる"],
    ["「物件名が新築データと一致しない中古データ ◯行は除外」", "中古データの物件名が、新築データと少し違う", "「タワー」と「ﾀﾜｰ」、スペースの有無などをそろえる"],
    ["文字化けする（◆や?が並ぶ）", "CSV の文字コードが想定外", "Excel で「CSV UTF-8」を選んで保存し直す"],
    ["グラフが一部しか出ない／件数が少ない", "価格や面積に「万円」「㎡」以外の文字、空欄がある", "数字だけにする。カンマ区切りの「5,800」もOKだが、全角数字は半角に"],
  ];
  const cols = [4.1, 3.4, 4.63];
  const hx = [MX, MX + cols[0], MX + cols[0] + cols[1]];
  let y = 1.75;
  ["画面の表示・症状", "原因", "直し方"].forEach((h, i) => {
    R(s, { x: hx[i], y, w: cols[i], h: 0.45, fill: { color: C.DARK }, line: { color: "FFFFFF", width: 1 } });
    T(s, h, { x: hx[i] + 0.15, y, w: cols[i] - 0.3, h: 0.45, fontSize: 13, bold: true, color: "FFFFFF", valign: "middle" });
  });
  y += 0.45;
  rows.forEach((r, ri) => {
    r.forEach((t, i) => {
      R(s, { x: hx[i], y, w: cols[i], h: 1.08, fill: { color: ri % 2 ? "FFFFFF" : C.CARD }, line: { color: C.LINE, width: 0.75 } });
      T(s, t, { x: hx[i] + 0.15, y, w: cols[i] - 0.3, h: 1.08, fontSize: 12, bold: i === 0, color: i === 0 ? C.INK : C.BODY, valign: "middle", lineSpacingMultiple: 1.15 });
    });
    y += 1.08;
  });
}

// 列名が違うとき
{
  const s = slide("第4部", "自社データの列名が違うときは、ツールの方を合わせる", "CSV の「1行目だけ」を貼ります。2行目以降（物件の中身）は絶対に貼りません");
  promptBox(s, MX, 1.7, 7.6, 3.6, P.cols, 10.8);
  const x = 8.55, w = W - MX - x;
  card(s, x, 1.7, w, 2.3, "1行目だけ、の例", "物件名称,所在区,駅,徒歩,竣工,定借満了,戸数,販売価格(円),面積(坪)\n\n↑ このような列名の行だけをコピーします", { fs: 11.5 });
  card(s, x, 4.15, w, 1.15, "そのあとは", "STEP 15 の5ステップ（Create PR → Merge → Download）", { fs: 11.5 });
  bar(s, 5.6, "warn", "物件名・価格などのデータの中身を Claude に貼るのは禁止です（はじめにのルール1・2）。迷ったら貼らないでください", { h: 0.65, label: "⚠" });
}

// =====================================================================
divider(5, "困ったとき・使いこなす", "うまく動かないとき、見た目を変えたいとき、お客様への説明に使うときの\nプロンプトと、グラフの読み方をまとめます");

// 困ったときのプロンプト
{
  const s = slide("第5部", "困ったときのプロンプト", "＜　＞の部分を自分の状況に書きかえて送ります");
  const bw = (W - MX * 2 - 0.3) / 2;
  T(s, "うまく動かないとき", { x: MX, y: 1.7, w: bw, h: 0.35, fontSize: 15, bold: true, color: C.INK });
  promptBox(s, MX, 2.1, bw, 3.9, P.error, 10.5);
  T(s, "見た目を変えたいとき", { x: MX + bw + 0.3, y: 1.7, w: bw, h: 0.35, fontSize: 15, bold: true, color: C.INK });
  promptBox(s, MX + bw + 0.3, 2.1, bw, 3.9, P.look, 10.5);
  bar(s, 6.2, "tip", "エラーの文字は、画面のスクリーンショットではなく「文字のまま」コピーして貼ると、Claude が理解しやすくなります", { h: 0.6, fs: 12 });
}

// グラフの読み方
{
  const s = slide("第5部", "グラフの読み方 ─ お客様に説明できるように", "3つの言葉が分かれば、グラフが読めます");
  const defs = [
    ["坪単価", "価格 ÷ 専有面積 × 3.30578", "広さのちがう住戸を、同じものさしで比べるための数字（万円/坪）"],
    ["残存年数", "借地期限年 − 成約年", "その住戸が売買された時点で、借地の期限まであと何年あったか"],
    ["新築比", "成約坪単価 ÷ 新築時の坪単価", "100%を超えると「新築時より高く売れた」。残存年数が減るとどう動くかを見る"],
  ];
  let y = 1.75;
  defs.forEach((d) => {
    R(s, { x: MX, y, w: W - MX * 2, h: 1.1, fill: { color: C.CARD }, line: { color: C.LINE, width: 0.75 }, rectRadius: 0.08 }, pptx.ShapeType.roundRect);
    T(s, d[0], { x: MX + 0.3, y, w: 1.9, h: 1.1, fontSize: 20, bold: true, color: C.GREEN, valign: "middle" });
    T(s, d[1], { x: MX + 2.3, y: y + 0.15, w: 9.6, h: 0.4, fontSize: 14, bold: true, color: C.INK });
    T(s, d[2], { x: MX + 2.3, y: y + 0.58, w: 9.6, h: 0.4, fontSize: 12.5, color: C.BODY });
    y += 1.25;
  });
  bar(s, 5.6, "warn", "グラフは過去の成約の傾向です。将来の価格を約束するものではありません。個別の住戸の価格を言い切らないようにしましょう", { h: 0.6, fs: 12, label: "⚠" });
  bar(s, 6.3, "tip", "説明文づくりも Claude に頼めます → 次のスライド", { h: 0.5, fs: 12 });
}

// 説明文プロンプト + 応用
{
  const s = slide("第5部", "お客様向けの説明文を作る／次の一歩", "ツール作り以外にも、Claude は営業の準備に使えます");
  const bw = 7.0;
  promptBox(s, MX, 1.7, bw, 2.6, P.explain, 11);
  const x = MX + bw + 0.35, w = W - MX - x;
  card(s, x, 1.7, w, 2.6, "このほかの使い方", "・ツールに「CSV に書き出す」機能を足す\n・グラフを画像で保存するボタンを足す\n・新人向けに、画面の使い方マニュアルを書いてもらう", { fs: 12 });
  card(s, MX, 4.55, W - MX * 2, 1.45, "社内で共有したいとき", "index.html のフォルダを社内の共有フォルダに置けば、他の人も使えます（データは各自のPCで読み込み）。Web で公開する方法もありますが、公開範囲は情報システム部門と相談して決めてください。", { fs: 12 });
  bar(s, 6.2, "tip", "新しい機能は「1回に1テーマ」。うまくいったら次のお願いへ", { h: 0.55, fs: 12 });
}

// 参考リンク・まとめ
{
  const s = slide("おわりに", "今日のまとめと、参考リンク", "スライドどおりに進めれば、何度でも作り直せます");
  const check = [
    "GitHub に Private リポジトリ「" + REPO + "」がある",
    "Claude Code と GitHub がつながっている（STEP 8 の ①② がOK）",
    "プロンプト①〜③で、グラフ・一覧・詳細画面ができた",
    "自社の新築・中古データを CSV にして、画面で読み込めた",
    "データは Claude にも GitHub にも渡していない",
  ];
  let y = 1.75;
  check.forEach((c) => {
    T(s, [{ text: "☑  ", options: { color: C.OK, bold: true } }, { text: c, options: { color: C.INK } }], { x: MX, y, w: 7.2, h: 0.5, fontSize: 14, valign: "middle" });
    y += 0.58;
  });
  const x = 8.2, w = W - MX - x;
  card(s, x, 1.75, w, 1.5, "お手本（本家）", "定借ナビ\nteisha-navi.vercel.app", { fs: 13 });
  card(s, x, 3.4, w, 1.95, "プロンプト集・完成見本", "github.com/Yibooo/teisha-PIVOT\n→ study-session フォルダ\n（prompts.md／reference）", { fs: 13 });
  bar(s, 5.55, "tip", "困ったら「STEP 番号」と「画面に出た文字」を添えて、講師に連絡してください", { h: 0.6 });
}

// =====================================================================
if (BREAK) {
  // 検証用：3枚目に、わざと他の要素と重なるテキストを置く／はみ出す長文を置く
  cur = 3;
  const s = pptx._slides[2];
  T(s, "【崩し】重なっているテキスト", { x: MX + 1.0, y: 2.0, w: 5, h: 0.5, fontSize: 18, bold: true, color: "FF0000" });
  T(s, "【崩し】はみ出す長文 ".repeat(30), { x: MX, y: 6.6, w: 4, h: 0.3, fontSize: 14, color: "FF0000" });
}

fs.writeFileSync((process.env.OUT || "study.pptx").replace(".pptx", ".layout.json"), JSON.stringify(reg));
const out = process.env.OUT || "study.pptx";
pptx.writeFile({ fileName: out }).then(() => console.log("wrote", out, pageNo, "slides"));
