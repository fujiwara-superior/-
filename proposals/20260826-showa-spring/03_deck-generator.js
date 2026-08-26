const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";           // 13.3 x 7.5
const W = 13.3, H = 7.5, M = 0.62;
const CW = W - M * 2;

// ---- palette : steel + copper (metal / spring manufacturing) ----
const STEEL   = "2B3440";
const STEEL_D = "1B222B";
const INK     = "232A33";
const BODY    = "3D4854";
const MUTED   = "6E7A87";
const LINE    = "D6DBE0";
const PANEL   = "F1F3F5";
const COPPER  = "C0642A";
const COPPER_T= "F7E9DF";
const GREEN   = "2F6B4F";
const GREEN_T = "E3EDE7";
const AMBER   = "9A6B14";
const AMBER_T = "F7EDD8";
const WHITE   = "FFFFFF";

const FJ = "Meiryo";

const sh = () => ({ type: "outer", color: "9AA4AE", blur: 8, offset: 1, angle: 90, opacity: 0.28 });

function slide(dark) {
  const s = pres.addSlide();
  s.background = { color: dark ? STEEL : WHITE };
  return s;
}

function title(s, t, sub) {
  s.addText(t, { x: M, y: 0.42, w: CW, h: 0.62, fontSize: 30, bold: true,
    color: INK, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  if (sub) s.addText(sub, { x: M, y: 1.06, w: CW, h: 0.36, fontSize: 13.5,
    color: MUTED, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
}

function chip(s, n, x, y, size, bg, fg) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w: size, h: size, fill: { color: bg || COPPER },
    rectRadius: 0.06, line: { color: bg || COPPER, width: 0 } });
  s.addText(String(n), { x, y, w: size, h: size, fontSize: size * 34, bold: true,
    color: fg || WHITE, align: "center", valign: "middle", fontFace: FJ, isTextBox: true, margin: 0 });
}

function card(s, x, y, w, h, fill, border) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: fill || PANEL },
    rectRadius: 0.04, line: border ? { color: border, width: 1 } : { color: fill || PANEL, width: 0 },
    shadow: sh() });
}

function foot(s, txt) {
  s.addText(txt, { x: M, y: H - 0.62, w: CW, h: 0.3, fontSize: 10, color: MUTED,
    fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
}

function pageNo(s, n) {
  s.addText(String(n), { x: W - M - 0.5, y: H - 0.62, w: 0.5, h: 0.3, fontSize: 10,
    color: MUTED, align: "right", fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
}

/* ============ 1. 表紙 ============ */
{
  const s = slide(true);
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 1.55, w: 0.42, h: 0.42,
    fill: { color: COPPER }, rectRadius: 0.08, line: { color: COPPER, width: 0 } });
  s.addText("株式会社昌和発條製作所 御中", { x: M + 0.62, y: 1.55, w: 8, h: 0.42,
    fontSize: 15, color: "C9D2DA", fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("受注入力業務の自動化に関するご提案", { x: M, y: 2.35, w: CW, h: 1.0,
    fontSize: 42, bold: true, color: WHITE, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("メール・FAXで届く注文を、基幹システムへ自動で起票する", { x: M, y: 3.45, w: CW, h: 0.5,
    fontSize: 17, color: "A8B4C0", fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addShape(pres.ShapeType.rect, { x: M, y: 4.35, w: 2.2, h: 0.02, fill: { color: COPPER }, line: { width: 0 } });
  s.addText("株式会社スペリオル", { x: M, y: 4.72, w: 6, h: 0.36, fontSize: 15, bold: true,
    color: WHITE, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("2026年8月27日　初回お打ち合わせ資料", { x: M, y: 5.12, w: 6, h: 0.32, fontSize: 12,
    color: "8C99A6", fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("本資料はヒアリング前の仮説に基づくドラフトです。本日確認した内容を反映し、正式提案として確定させていただきます。",
    { x: M, y: 6.35, w: CW, h: 0.4, fontSize: 10.5, color: "7C8996", fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addNotes("初回商談。冒頭は提案せず、ヒアリングから入る。");
}

/* ============ 2. 本日の進行 ============ */
{
  const s = slide();
  title(s, "本日の進行", "60分／前半は、伺うことに時間を使わせてください");
  const rows = [
    ["1", "現状のヒアリング", "業務の流れ、件数の内訳、基幹システムについて", "約20分"],
    ["2", "想定される改善方針のご説明", "何を自動化し、何を自動化しないか", "約20分"],
    ["3", "実現方式・費用・補助金のご相談", "連携方式の選択肢と概算費用", "約15分"],
    ["4", "次のステップの確認", "無償の読み取り精度検証について", "約5分"],
  ];
  let y = 1.72;
  rows.forEach((r, i) => {
    const hgt = 1.02;
    card(s, M, y, CW, hgt, i === 0 ? COPPER_T : PANEL);
    chip(s, r[0], M + 0.34, y + 0.26, 0.5);
    s.addText(r[1], { x: M + 1.05, y: y + 0.16, w: 6.6, h: 0.38, fontSize: 16, bold: true,
      color: INK, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(r[2], { x: M + 1.05, y: y + 0.55, w: 8.2, h: 0.34, fontSize: 12,
      color: BODY, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(r[3], { x: W - M - 1.5, y: y + 0.3, w: 1.15, h: 0.42, fontSize: 14, bold: true,
      color: i === 0 ? COPPER : MUTED, align: "right", fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    y += hgt + 0.16;
  });
  pageNo(s, 2);
}

/* ============ 3. 私たちの理解 ============ */
{
  const s = slide();
  title(s, "私たちの理解", "お申し込みの際にご記入いただいた内容から");
  const items = [
    ["手入力", "メール・FAXで届く見積や注文を、\n基幹システムへ人が打ち込んでいる"],
    ["月400〜500件", "1日あたり約20〜25件。\n担当者の時間が転記に消費されている"],
    ["属人化", "処理・管理の方法が担当者任せで、\n共有できていない"],
  ];
  const cw = (CW - 0.4) / 3;
  items.forEach((it, i) => {
    const x = M + i * (cw + 0.2);
    card(s, x, 1.78, cw, 2.05, i === 2 ? AMBER_T : PANEL);
    s.addText(it[0], { x: x + 0.36, y: 2.14, w: cw - 0.72, h: 0.5, fontSize: 22, bold: true,
      color: i === 2 ? AMBER : INK, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(it[1], { x: x + 0.36, y: 2.76, w: cw - 0.72, h: 0.8, fontSize: 12.5,
      color: BODY, fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 19, valign: "top" });
  });
  card(s, M, 4.12, CW, 1.42, WHITE, LINE);
  s.addText("ご記入いただいた内容（原文）", { x: M + 0.4, y: 4.3, w: 6, h: 0.3, fontSize: 10.5,
    bold: true, color: COPPER, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("「メールやFAXで届く見積や注文内容を基幹システムの受注データに手入力している。毎月400〜500件ほど。\n処理・管理方法が担当者任せで属人化しており、共有化できていない。」",
    { x: M + 0.4, y: 4.62, w: CW - 0.8, h: 0.76, fontSize: 13, italic: true, color: INK,
      fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 20 });
  foot(s, "この3点を、本日の議論の起点とさせてください。");
  pageNo(s, 3);
}

/* ============ 4. 現状の業務フロー ============ */
{
  const s = slide();
  title(s, "現状の業務フロー（弊社の推定）", "認識に相違があればご指摘ください");
  const steps = [
    ["受信", "FAX（紙）\nメール添付・本文", ""],
    ["内容の判読", "得意先・品番・数量\n納期・注文区分", "1〜3分"],
    ["特定・照合", "得意先コード\n品番の特定", "2〜6分"],
    ["基幹へ手入力", "受注ヘッダ＋明細行\nを1件ずつ", "2〜5分"],
    ["確認・返信", "在庫／納期回答\n注文請書の返送", "1〜2分"],
  ];
  const bw = 2.26, gap = 0.26;
  const startX = M + (CW - (bw * 5 + gap * 4)) / 2;
  steps.forEach((st, i) => {
    const x = startX + i * (bw + gap);
    const hi = (i === 2);
    card(s, x, 1.9, bw, 2.32, hi ? COPPER_T : PANEL);
    chip(s, i + 1, x + 0.22, 2.12, 0.4, hi ? COPPER : STEEL);
    s.addText(st[0], { x: x + 0.72, y: 2.12, w: bw - 0.9, h: 0.4, fontSize: 14.5, bold: true,
      color: INK, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(st[1], { x: x + 0.22, y: 2.66, w: bw - 0.44, h: 0.86, fontSize: 11,
      color: BODY, fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 15 });
    if (st[2]) s.addText("1件 " + st[2], { x: x + 0.22, y: 3.62, w: bw - 0.44, h: 0.36, fontSize: 12,
      bold: true, color: hi ? COPPER : MUTED, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    if (i < 4) s.addShape(pres.ShapeType.triangle, { x: x + bw + 0.04, y: 2.92, w: 0.18, h: 0.26,
      fill: { color: "AEB8C2" }, line: { width: 0 }, rotate: 90 });
  });
  card(s, M, 4.46, CW, 1.12, AMBER_T);
  s.addText("全工程が、特定のご担当者の頭の中のルールで運用されている", { x: M + 0.42, y: 4.62, w: CW - 0.84, h: 0.38,
    fontSize: 16, bold: true, color: INK, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("手順書がないため、他の方が代われない。改善すべき点も見えない。", { x: M + 0.42, y: 5.02, w: CW - 0.84, h: 0.34,
    fontSize: 12, color: BODY, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  foot(s, "所要時間は弊社の推定値です。実際の感覚をお聞かせください。");
  pageNo(s, 4);
}

/* ============ 5. 課題の構造 ============ */
{
  const s = slide();
  title(s, "課題は4つに整理できます", "");
  const q = [
    ["A", "入力工数", "受注が増えるほど人を増やさざるを得ない。\n増員できなければ受注機会そのものを制限する", false],
    ["B", "入力ミス", "誤品番・誤数量・誤納期が生産指示まで流れる。\n1個から受注されるため、少量でも段取り替えが重い", false],
    ["C", "属人化", "ご担当者の不在で受注業務が止まる。\n引き継ぎに数ヶ月を要し、改善の議論もできない", true],
    ["D", "見えない", "どの注文がどこまで進んだか分からない。\n納期回答が遅れ、問い合わせが担当者に集中する", false],
  ];
  const cw2 = (CW - 0.24) / 2, ch = 1.52;
  q.forEach((it, i) => {
    const x = M + (i % 2) * (cw2 + 0.24);
    const y = 1.72 + Math.floor(i / 2) * (ch + 0.22);
    card(s, x, y, cw2, ch, it[3] ? COPPER_T : PANEL);
    chip(s, it[0], x + 0.3, y + 0.3, 0.46, it[3] ? COPPER : STEEL);
    s.addText(it[1], { x: x + 0.94, y: y + 0.26, w: cw2 - 1.2, h: 0.4, fontSize: 17, bold: true,
      color: INK, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(it[2], { x: x + 0.94, y: y + 0.7, w: cw2 - 1.24, h: 0.66, fontSize: 11.5,
      color: BODY, fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 16 });
  });
  card(s, M, 5.28, CW, 1.12, STEEL);
  s.addText("最も重いのは C です。", { x: M + 0.42, y: 5.44, w: 3.2, h: 0.38, fontSize: 17, bold: true,
    color: WHITE, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("A（工数）は費用の問題ですが、C は事業が止まるかどうかの問題です。ここを解かない提案に意味はないと考えています。",
    { x: M + 0.42, y: 5.84, w: CW - 0.84, h: 0.36, fontSize: 12.5, color: "C6D0DA",
      fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  pageNo(s, 5);
}

/* ============ 6. 現状コスト ============ */
{
  const s = slide();
  title(s, "現状のコスト試算", "月450件（400〜500件の中央値）を基準とした概算");
  card(s, M, 1.78, 4.5, 2.62, STEEL);
  s.addText("年間の入力工数", { x: M + 0.4, y: 1.98, w: 3.7, h: 0.32, fontSize: 12,
    color: "A8B4C0", fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("900", { x: M + 0.4, y: 2.3, w: 2.4, h: 1.0, fontSize: 66, bold: true,
    color: WHITE, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("時間", { x: M + 2.55, y: 2.72, w: 1.0, h: 0.5, fontSize: 18, bold: true,
    color: COPPER, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("正社員1名の年間労働時間の、およそ半分に相当します", { x: M + 0.4, y: 3.42, w: 3.8, h: 0.62,
    fontSize: 12, color: "C6D0DA", fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 17 });

  const tx = M + 4.78, tw = CW - 4.78;
  s.addTable([
    [{ text: "1件あたり", options: { bold: true } }, { text: "月間工数", options: { bold: true } },
     { text: "年間工数", options: { bold: true } }, { text: "年間人件費", options: { bold: true } }],
    ["6分（保守的）", "45.0 時間", "540 時間", "135〜162 万円"],
    ["10分（標準）", "75.0 時間", "900 時間", "225〜270 万円"],
    ["15分（重い）", "112.5 時間", "1,350 時間", "338〜405 万円"],
  ], {
    x: tx, y: 1.78, w: tw, colW: [tw * 0.26, tw * 0.22, tw * 0.22, tw * 0.30],
    rowH: 0.5, fontSize: 12, fontFace: FJ, color: BODY, valign: "middle",
    border: { type: "solid", color: LINE, pt: 1 }, align: "center",
    fill: { color: WHITE },
  });
  s.addText("計算式：450件 × 1件あたり分数 ÷ 60 ＝ 月間工数 → ×12 ＝ 年間工数 → × 人件費単価（2,500〜3,000円/時）\n人件費単価は、法定福利費等を含む企業負担ベースの想定値です。",
    { x: tx, y: 3.86, w: tw, h: 0.6, fontSize: 10.5, color: MUTED, fontFace: FJ,
      isTextBox: true, margin: 0, lineSpacing: 15 });

  card(s, M, 4.62, CW, 1.22, AMBER_T);
  s.addText("入力ミスに伴うコスト（参考）　約30万円／年", { x: M + 0.42, y: 4.78, w: 7.5, h: 0.38,
    fontSize: 15, bold: true, color: INK, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("誤入力率1%（月4.5件）× 1件あたり手戻り2時間 と仮定。特急便の実費、生産計画の乱れ、信用低下は含みません。",
    { x: M + 0.42, y: 5.18, w: CW - 0.84, h: 0.36, fontSize: 11.5, color: BODY,
      fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  foot(s, "実際の誤入力の頻度と、直近1年で実害の出た件数をお聞かせください。");
  pageNo(s, 6);
}

/* ============ 7. 全体像 ============ */
{
  const s = slide();
  title(s, "ご提案する仕組みの全体像", "人が「確認するだけ」の状態を目指します");
  const boxes = [
    ["受信", "FAX・メール\n添付形式を自動判定"],
    ["読み取り", "AI-OCR ＋ 生成AI\nExcel/PDFは直接読む"],
    ["マスタ突合", "品番・得意先マスタと\n照合して自動補正"],
    ["仕分け", "確信度により\n3レーンに振り分け"],
    ["確認画面", "原本と結果を並べて\n人が最終判断"],
    ["基幹へ登録", "CSV / API / RPA\nいずれかで連携"],
  ];
  const bw = (CW - 0.2 * 5) / 6;
  boxes.forEach((b, i) => {
    const x = M + i * (bw + 0.2);
    const hi = (i === 2 || i === 4);
    card(s, x, 1.86, bw, 2.0, hi ? COPPER_T : PANEL);
    chip(s, i + 1, x + 0.18, 2.04, 0.36, hi ? COPPER : STEEL);
    s.addText(b[0], { x: x + 0.16, y: 2.5, w: bw - 0.32, h: 0.36, fontSize: 13.5, bold: true,
      color: INK, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(b[1], { x: x + 0.16, y: 2.9, w: bw - 0.32, h: 0.84, fontSize: 10.5,
      color: BODY, fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 14 });
    if (i < 5) s.addShape(pres.ShapeType.triangle, { x: x + bw + 0.02, y: 2.74, w: 0.16, h: 0.24,
      fill: { color: "AEB8C2" }, line: { width: 0 }, rotate: 90 });
  });
  card(s, M, 4.06, CW, 0.86, WHITE, LINE);
  s.addText("全処理の履歴を記録　／　進捗ダッシュボードで状況を可視化", { x: M + 0.42, y: 4.06, w: CW - 0.84, h: 0.86,
    fontSize: 13.5, bold: true, color: STEEL, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });

  const notes = [
    ["★ 精度の要は「マスタ突合」", "文字認識の誤りを、御社の品番マスタとの照合で機械的に補正します"],
    ["★ 人の確認は必ず残します", "完全自動化はお約束しません。確認工程を残すことが前提の設計です"],
  ];
  const nw = (CW - 0.24) / 2;
  notes.forEach((n, i) => {
    const x = M + i * (nw + 0.24);
    card(s, x, 5.12, nw, 1.06, GREEN_T);
    s.addText(n[0], { x: x + 0.34, y: 5.24, w: nw - 0.68, h: 0.34, fontSize: 13.5, bold: true,
      color: GREEN, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(n[1], { x: x + 0.34, y: 5.58, w: nw - 0.68, h: 0.44, fontSize: 11,
      color: BODY, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  });
  pageNo(s, 7);
}

/* ============ 8. マスタ突合 ============ */
{
  const s = slide();
  title(s, "精度の要は、OCRではなく「品番マスタ突合」です", "本日いちばんお伝えしたい点です");
  card(s, M, 1.8, CW, 2.5, PANEL);
  const stepX = [M + 0.5, M + 4.35, M + 8.2];
  const RED = "B4271A";
  const labels = [
    ["① OCRの読み取り結果", [
       { text: "AP-1", options: { color: MUTED } },
       { text: "O", options: { color: RED, bold: true, underline: { style: "sng" } } },
       { text: "-25", options: { color: MUTED } }],
     "英字の「O（オー）」と読んでしまった", COPPER],
    ["② 品番マスタと照合", [{ text: "該当なし", options: { color: AMBER, bold: true } }],
     "この品番はマスタに存在しない", AMBER],
    ["③ 自動補正", [
       { text: "AP-1", options: { color: MUTED } },
       { text: "0", options: { color: GREEN, bold: true, underline: { style: "sng" } } },
       { text: "-25", options: { color: MUTED } }],
     "数字の「0（ゼロ）」へ補正　確信度98%", GREEN],
  ];
  labels.forEach((l, i) => {
    s.addText(l[0], { x: stepX[i], y: 2.02, w: 3.5, h: 0.34, fontSize: 12, bold: true,
      color: MUTED, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addShape(pres.ShapeType.roundRect, { x: stepX[i], y: 2.44, w: 3.2, h: 0.82,
      fill: { color: WHITE }, rectRadius: 0.04, line: { color: l[3], width: 1.5 } });
    s.addText(l[1], { x: stepX[i], y: 2.44, w: 3.2, h: 0.82, fontSize: 26, bold: true,
      align: "center", valign: "middle", fontFace: "Courier New", isTextBox: true, margin: 0 });
    s.addText(l[2], { x: stepX[i] - 0.15, y: 3.34, w: 3.5, h: 0.34, fontSize: 11.5,
      color: BODY, align: "center", fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    if (i < 2) s.addShape(pres.ShapeType.triangle, { x: stepX[i] + 3.38, y: 2.72, w: 0.2, h: 0.28,
      fill: { color: "8E99A4" }, line: { width: 0 }, rotate: 90 });
  });
  card(s, M, 4.5, 6.2, 1.72, STEEL);
  s.addText("御社には「正解リスト」があります", { x: M + 0.4, y: 4.68, w: 5.4, h: 0.38, fontSize: 16,
    bold: true, color: WHITE, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("サンエス標準ばね 約1,717種\n圧縮919 ／ 引張424 ／ フリーサイズ104 ／ ねじり270",
    { x: M + 0.4, y: 5.1, w: 5.4, h: 0.72, fontSize: 12, color: "C6D0DA", fontFace: FJ,
      isTextBox: true, margin: 0, lineSpacing: 17 });
  card(s, M + 6.44, 4.5, CW - 6.44, 1.72, COPPER_T);
  s.addText("読み取り精度を上げるより、存在しない品番を機械的に弾くほうが確実です。",
    { x: M + 6.84, y: 4.68, w: CW - 7.24, h: 0.62, fontSize: 14.5, bold: true, color: INK,
      fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 20 });
  s.addText("得意先ごとの呼称（先方品番）も、読替表を整備すれば同じ仕組みで吸収できます。",
    { x: M + 6.84, y: 5.34, w: CW - 7.24, h: 0.62, fontSize: 11.5, color: BODY,
      fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 16 });
  pageNo(s, 8);
}

/* ============ 9. 標準品と特注品 ============ */
{
  const s = slide();
  title(s, "標準品と特注品は、入口で分けます", "すべてを自動化するご提案はいたしません");
  const cols = [
    ["標準品（サンエス）", "型番指定・リピート・在庫品", ["判断の要素が少ない", "自動化の適性が高い", "自動起票の対象とする"], GREEN, GREEN_T],
    ["特注品（図面・異形線ばね等）", "図面ベース・仕様確認が必要", ["技術的な判断が必須", "自動化すべきでない", "見積依頼として技術ご担当へ振り分け"], AMBER, AMBER_T],
  ];
  const cw3 = (CW - 0.3) / 2;
  cols.forEach((c, i) => {
    const x = M + i * (cw3 + 0.3);
    card(s, x, 1.8, cw3, 2.96, c[4]);
    s.addText(c[0], { x: x + 0.42, y: 2.0, w: cw3 - 0.84, h: 0.44, fontSize: 18, bold: true,
      color: INK, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(c[1], { x: x + 0.42, y: 2.46, w: cw3 - 0.84, h: 0.34, fontSize: 12,
      color: MUTED, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    c[2].forEach((li, j) => {
      const yy = 2.94 + j * 0.52;
      s.addShape(pres.ShapeType.roundRect, { x: x + 0.42, y: yy + 0.09, w: 0.16, h: 0.16,
        fill: { color: c[3] }, rectRadius: 0.05, line: { width: 0 } });
      s.addText(li, { x: x + 0.72, y: yy, w: cw3 - 1.14, h: 0.36, fontSize: 13,
        color: BODY, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    });
  });
  card(s, M, 4.96, CW, 1.3, STEEL);
  s.addText("特注品を無理に自動化しません。", { x: M + 0.42, y: 5.12, w: 5.0, h: 0.4, fontSize: 16,
    bold: true, color: WHITE, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("むしろ「特注品を技術ご担当者へ確実に、早く渡す」ことに価値があります。振り分けの自動化だけでも、見積回答のリードタイムは短縮できます。",
    { x: M + 0.42, y: 5.54, w: CW - 0.84, h: 0.4, fontSize: 12.5, color: "C6D0DA",
      fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  pageNo(s, 9);
}

/* ============ 10. 属人化 ============ */
{
  const s = slide();
  title(s, "属人化は、工数削減とは別に必ず解きます", "課題C への打ち手");
  const items = [
    ["判断ルールを外に出す", "「この得意先のこの表記はこの品番」という頭の中のルールを、\n誰でも見て直せるマスタとしてシステム上に置きます"],
    ["処理履歴をすべて記録", "誰が・いつ・何を修正したかを全件記録。\n「なぜこう入力したか」を後から追えるようにします"],
    ["進捗を一覧で見える化", "未処理／確認待ち／起票済を一覧化。\n担当者に聞かなくても状況が分かる状態にします"],
    ["修正パターンを蓄積", "人がAIの結果を直した内容を蓄積し、\n繰り返す修正は自動補正のルールへ昇格させます"],
  ];
  const cw4 = (CW - 0.24) / 2;
  items.forEach((it, i) => {
    const x = M + (i % 2) * (cw4 + 0.24);
    const y = 1.76 + Math.floor(i / 2) * 1.44;
    card(s, x, y, cw4, 1.26, PANEL);
    chip(s, i + 1, x + 0.3, y + 0.24, 0.42, STEEL);
    s.addText(it[0], { x: x + 0.88, y: y + 0.2, w: cw4 - 1.16, h: 0.36, fontSize: 15, bold: true,
      color: INK, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(it[1], { x: x + 0.88, y: y + 0.58, w: cw4 - 1.2, h: 0.58, fontSize: 11,
      color: BODY, fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 15 });
  });
  card(s, M, 4.72, CW, 1.28, GREEN_T);
  s.addText("結果として得られるもの", { x: M + 0.42, y: 4.86, w: 5.0, h: 0.36, fontSize: 13,
    bold: true, color: GREEN, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("ご担当者が不在でも、他の方が画面を見れば処理を継続できます。引き継ぎ資料ではなく、システムそのものが手順書になります。",
    { x: M + 0.42, y: 5.24, w: CW - 0.84, h: 0.5, fontSize: 15, bold: true, color: INK,
      fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  pageNo(s, 10);
}

/* ============ 11. 導入計画 ============ */
{
  const s = slide();
  title(s, "段階的に導入します", "一度に全部は作りません。効果が出やすい範囲から始めます");
  const ph = [
    ["Phase 1", "2.5〜3.5ヶ月", "メール添付＋標準品の自動起票",
      "対象：全体の約35%\n完了判定：対象内の自動確定率60%以上", COPPER, COPPER_T],
    ["Phase 2", "2.5〜3.5ヶ月", "FAX・非定型帳票・手書きへ拡張",
      "対象：全体の約80%へ\n完了判定：全体の自動確定率70%以上", STEEL, PANEL],
    ["Phase 3", "別途ご相談", "見積フロー・可視化・他業務へ展開",
      "特注品の見積フロー整備\n納期回答の半自動化、仕入・請求への展開", MUTED, PANEL],
  ];
  const cw5 = (CW - 0.36) / 3;
  ph.forEach((p, i) => {
    const x = M + i * (cw5 + 0.18);
    card(s, x, 1.76, cw5, 2.62, p[5]);
    s.addText(p[0], { x: x + 0.34, y: 1.96, w: cw5 - 0.68, h: 0.42, fontSize: 19, bold: true,
      color: p[4], fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(p[1], { x: x + 0.34, y: 2.38, w: cw5 - 0.68, h: 0.32, fontSize: 12,
      color: MUTED, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(p[2], { x: x + 0.34, y: 2.8, w: cw5 - 0.68, h: 0.66, fontSize: 13.5, bold: true,
      color: INK, fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 19 });
    s.addText(p[3], { x: x + 0.34, y: 3.5, w: cw5 - 0.68, h: 0.72, fontSize: 11,
      color: BODY, fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 15 });
  });
  card(s, M, 4.58, CW, 1.6, AMBER_T);
  s.addText("正直に申し上げます　—　Phase 1 で最も大変なのは、御社側の作業です",
    { x: M + 0.42, y: 4.74, w: CW - 0.84, h: 0.4, fontSize: 15.5, bold: true, color: INK,
      fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("過去の注文書サンプル30〜50件のご提供／品目・得意先マスタのCSV出力／基幹システムの仕様確認／週1回1時間の打合せ／そして運用テスト期間2週間の並行稼働。\nここを軽く見積もると導入は失敗します。繁忙期を避けた日程をご相談させてください。",
    { x: M + 0.42, y: 5.16, w: CW - 0.84, h: 0.86, fontSize: 11.5, color: BODY,
      fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 16 });
  pageNo(s, 11);
}

/* ============ 12. 効果と投資回収 ============ */
{
  const s = slide();
  title(s, "効果と投資回収　—　都合の悪い数字も含めて", "");
  const tw2 = CW;
  s.addTable([
    [{ text: "", options: { fill: { color: WHITE } } },
     { text: "Phase1 のみ", options: { bold: true } },
     { text: "Phase1＋2（保守的）", options: { bold: true } },
     { text: "Phase1＋2（期待値）", options: { bold: true } }],
    ["累計初期費用（中央値）", "125 万円", "300 万円", "300 万円"],
    ["年間ランニング", "30 万円", "42 万円", "42 万円"],
    ["年間削減額", "59 万円", "83 万円", "151 万円"],
    [{ text: "年間の純削減額", options: { bold: true } },
     { text: "29 万円", options: { bold: true } },
     { text: "41 万円", options: { bold: true } },
     { text: "109 万円", options: { bold: true } }],
    [{ text: "回収期間（補助金なし）", options: { bold: true } },
     { text: "4.3 年", options: { bold: true, color: AMBER } },
     { text: "7.3 年", options: { bold: true, color: "A33018" } },
     { text: "2.8 年", options: { bold: true, color: GREEN } }],
    [{ text: "回収期間（補助金1/2適用時）", options: { bold: true } },
     { text: "2.2 年", options: { bold: true, color: GREEN } },
     { text: "3.7 年", options: { bold: true, color: AMBER } },
     { text: "1.4 年", options: { bold: true, color: GREEN } }],
  ], {
    x: M, y: 1.66, w: tw2, colW: [tw2 * 0.31, tw2 * 0.23, tw2 * 0.23, tw2 * 0.23],
    rowH: 0.43, fontSize: 12, fontFace: FJ, color: BODY, valign: "middle",
    border: { type: "solid", color: LINE, pt: 1 }, align: "center", fill: { color: WHITE },
  });
  card(s, M, 4.86, CW, 1.5, STEEL);
  s.addText("工数削減だけを見れば、補助金なしの投資回収は決して速くありません。",
    { x: M + 0.42, y: 5.0, w: CW - 0.84, h: 0.4, fontSize: 15.5, bold: true, color: WHITE,
      fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("「3年で回収できます」と申し上げることはできません。そのうえで、① 補助金の活用で回収期間はおよそ半分になること、② 属人化の解消は工数削減とは別の価値であること、③ 受注が増えたとき人を増やさずに対応できること　を合わせてご判断ください。",
    { x: M + 0.42, y: 5.42, w: CW - 0.84, h: 0.8, fontSize: 11.5, color: "C6D0DA",
      fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 16 });
  foot(s, "前提：月450件、現状1件10分、人件費単価2,750円/時。数値はヒアリング結果により変動します。");
  pageNo(s, 12);
}

/* ============ 13. 費用と補助金 ============ */
{
  const s = slide();
  title(s, "概算費用と、活用できる可能性のある補助金", "確認事項の確定により変動します");
  s.addText("概算費用", { x: M, y: 1.66, w: 5.6, h: 0.34, fontSize: 14, bold: true,
    color: COPPER, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  const t1w = 5.9;
  s.addTable([
    [{ text: "", options: { fill: { color: WHITE } } }, { text: "初期費用", options: { bold: true } }, { text: "月額費用", options: { bold: true } }],
    [{ text: "Phase 1", options: { bold: true } }, "90〜160 万円", "2.0〜3.0 万円"],
    [{ text: "Phase 2", options: { bold: true } }, "130〜220 万円", "＋1.0〜1.5 万円"],
    [{ text: "Phase 3", options: { bold: true } }, "別途", "別途"],
  ], {
    x: M, y: 2.06, w: t1w, colW: [t1w * 0.28, t1w * 0.38, t1w * 0.34], rowH: 0.44,
    fontSize: 12, fontFace: FJ, color: BODY, valign: "middle", align: "center",
    border: { type: "solid", color: LINE, pt: 1 }, fill: { color: WHITE },
  });
  s.addText("変動要因：基幹システムの連携方式（CSV ≦ API < RPA の順に工数増）／取引先別フォーマットの種類数／マスタ整備の要否",
    { x: M, y: 3.94, w: t1w, h: 0.6, fontSize: 10.5, color: MUTED, fontFace: FJ,
      isTextBox: true, margin: 0, lineSpacing: 15, valign: "top" });

  const t2x = M + 6.24, t2w = CW - 6.24;
  s.addText("活用できる可能性のある補助金", { x: t2x, y: 1.66, w: t2w, h: 0.34,
    fontSize: 14, bold: true, color: COPPER, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addTable([
    [{ text: "制度", options: { bold: true } }, { text: "該当性", options: { bold: true } }, { text: "補助率の目安", options: { bold: true } }],
    [{ text: "IT導入補助金", options: { bold: true } }, "最有力", "1/2〜2/3"],
    ["省力化投資補助金", "可能性あり", "1/2 程度"],
    ["ものづくり補助金", "枠により可能", "1/2〜2/3"],
    ["大阪府・松原市の制度", "要確認", "制度による"],
  ], {
    x: t2x, y: 2.06, w: t2w, colW: [t2w * 0.42, t2w * 0.28, t2w * 0.30], rowH: 0.44,
    fontSize: 11.5, fontFace: FJ, color: BODY, valign: "middle", align: "center",
    border: { type: "solid", color: LINE, pt: 1 }, fill: { color: WHITE },
  });
  s.addText("上記は対象となる可能性を示すもので、採択を保証するものではありません。制度は年度ごとに要件が変わるため、申請前に最新の公募要領での確認が必須です。",
    { x: t2x, y: 4.34, w: t2w, h: 0.6, fontSize: 10.5, color: MUTED, fontFace: FJ,
      isTextBox: true, margin: 0, lineSpacing: 15, valign: "top" });

  card(s, M, 4.78, CW, 1.24, GREEN_T);
  s.addText("弊社の支援範囲", { x: M + 0.42, y: 4.92, w: 4.0, h: 0.34, fontSize: 13, bold: true,
    color: GREEN, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("申請書に添付する見積書・システム構成図・効果算定資料の作成をお手伝いします。申請書本体の作成代行は、制度上の制約により行いません。",
    { x: M + 0.42, y: 5.28, w: CW - 0.84, h: 0.5, fontSize: 12.5, color: INK,
      fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  pageNo(s, 13);
}

/* ============ 14. リスク ============ */
{
  const s = slide();
  title(s, "想定されるリスク　—　隠さず申し上げます", "");
  const risks = [
    ["想定した精度が出ない", "FAX画質が極端に悪い／手書きが多い／フォーマットが想定以上に多様", "契約前にサンプルで検証し、実測値を確認してから判断します"],
    ["基幹システムに連携口がない", "CSV取込・APIともに非対応の場合", "RPAで代替。ただし保守性が下がるため事前確認が最重要です"],
    ["並行稼働期間の負荷", "繁忙期と重なった場合", "繁忙期を避けた日程設定。閑散期をお聞かせください"],
    ["現場で使われない", "確認画面が使いにくく、手入力のほうが速い状態", "実際に入力される方に、設計段階から画面を見ていただきます"],
    ["マスタが整備されていない", "品目マスタに表記ゆれ・重複がある場合", "初期にマスタ状態を調査。整備工数が別途必要になる可能性があります"],
  ];
  let y = 1.6;
  risks.forEach((r) => {
    card(s, M, y, CW, 0.76, PANEL);
    s.addText(r[0], { x: M + 0.34, y: y, w: 3.3, h: 0.76, fontSize: 12.5, bold: true,
      color: INK, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(r[1], { x: M + 3.78, y: y, w: 4.3, h: 0.76, fontSize: 11, color: BODY,
      fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(r[2], { x: M + 8.24, y: y, w: CW - 8.58, h: 0.76, fontSize: 11, color: GREEN,
      fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    y += 0.86;
  });
  s.addText("発生条件", { x: M + 3.78, y: 1.26, w: 4.3, h: 0.28, fontSize: 10, bold: true,
    color: MUTED, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("弊社の対策", { x: M + 8.24, y: 1.26, w: CW - 8.58, h: 0.28, fontSize: 10, bold: true,
    color: MUTED, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  card(s, M, 6.06, CW, 0.72, STEEL);
  s.addText("サンプル検証で精度が出なければ、その旨を正直にご報告し、ご提案を取り下げます。",
    { x: M + 0.42, y: 6.06, w: CW - 0.84, h: 0.72, fontSize: 14, bold: true, color: WHITE,
      fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  pageNo(s, 14);
}

/* ============ 15. 次のステップ ============ */
{
  const s = slide(true);
  s.addText("次のステップのご提案", { x: M, y: 0.72, w: CW, h: 0.7, fontSize: 32, bold: true,
    color: WHITE, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("実際の注文書で、何％読めるかを先にお見せします", { x: M, y: 1.42, w: CW, h: 0.4,
    fontSize: 15, color: "A8B4C0", fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });

  const st = [
    ["サンプルのご提供", "注文書 1〜2枚（10枚程度あれば理想）\n得意先名・金額は黒塗りで構いません"],
    ["読み取り精度の検証", "項目別の精度／マスタ突合による補正効果\n自動化可能率の実測見込み／技術的な課題"],
    ["結果のご報告", "3営業日以内　・　費用は無償\n検証後にデータは削除。AIの学習には使用しません"],
  ];
  const cw6 = (CW - 0.44) / 3;
  st.forEach((x0, i) => {
    const x = M + i * (cw6 + 0.22);
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.1, w: cw6, h: 2.2, fill: { color: "3A4552" },
      rectRadius: 0.04, line: { color: "4A5665", width: 1 } });
    chip(s, i + 1, x + 0.34, 2.34, 0.46);
    s.addText(x0[0], { x: x + 0.34, y: 2.92, w: cw6 - 0.68, h: 0.42, fontSize: 17, bold: true,
      color: WHITE, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(x0[1], { x: x + 0.34, y: 3.38, w: cw6 - 0.68, h: 0.76, fontSize: 11.5,
      color: "B9C4CE", fontFace: FJ, isTextBox: true, margin: 0, lineSpacing: 16 });
  });

  s.addShape(pres.ShapeType.roundRect, { x: M, y: 4.62, w: CW, h: 1.16, fill: { color: COPPER },
    rectRadius: 0.04, line: { width: 0 } });
  s.addText("読み取り精度は、帳票の実物を拝見しないと分かりません。", { x: M + 0.42, y: 4.76, w: CW - 0.84, h: 0.4,
    fontSize: 16, bold: true, color: WHITE, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("「高精度です」と申し上げるより、御社の実際の注文書で何％読めたかをお見せするほうが、判断材料として確実だと考えています。",
    { x: M + 0.42, y: 5.16, w: CW - 0.84, h: 0.44, fontSize: 12, color: "FCEBE0",
      fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("株式会社スペリオル　／　受付番号 20260826-091859-64F2027B", { x: M, y: 6.5, w: CW, h: 0.34,
    fontSize: 10.5, color: "7C8996", fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
}

/* ============ 16. ヒアリング項目 ============ */
{
  const s = slide();
  title(s, "本日お伺いしたいこと", "分かる範囲で結構です。お答えいただける範囲でお聞かせください");
  const groups = [
    ["A. 基幹システム（最重要）", ["製品名・ベンダー名・導入時期", "受注データのCSV取込機能の有無と仕様書", "API連携／DB直接アクセスの可否", "保守ベンダーとの契約上の制約"]],
    ["B. 業務の実態", ["標準品と特注品の比率", "FAX と メール の比率", "1件あたりの実際の処理時間", "1件の注文に入る明細行数", "繁忙期・閑散期"]],
    ["C. 帳票", ["上位10社の注文書は何パターンか", "サンプル現物のご提供可否", "手書きが含まれる割合"]],
    ["D. 体制と実害", ["受注入力の担当者数・専任/兼務", "担当者不在時の現在の対応", "直近1年の入力ミスによる実害"]],
    ["E. マスタ・環境", ["得意先／品目マスタのCSV出力可否", "得意先別の品番読替表の有無", "クラウド利用に関する社内規程"]],
    ["F. 意思決定", ["ご予算の目安と決裁プロセス", "着手のご希望時期", "補助金申請のご意向"]],
  ];
  const cw7 = (CW - 0.4) / 3;
  groups.forEach((g, i) => {
    const x = M + (i % 3) * (cw7 + 0.2);
    const y = 1.72 + Math.floor(i / 3) * 2.42;
    card(s, x, y, cw7, i < 3 ? 2.24 : 1.88, i === 0 ? COPPER_T : PANEL);
    s.addText(g[0], { x: x + 0.32, y: y + 0.18, w: cw7 - 0.64, h: 0.36, fontSize: 13.5, bold: true,
      color: i === 0 ? COPPER : INK, fontFace: FJ, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(g[1].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < g[1].length - 1 } })),
      { x: x + 0.32, y: y + 0.62, w: cw7 - 0.64, h: 1.46, fontSize: 11, color: BODY,
        fontFace: FJ, isTextBox: true, margin: 0, paraSpaceAfter: 5, valign: "top" });
  });
  pageNo(s, 16);
}

pres.writeFile({ fileName: "/tmp/claude-0/-home-user--/0d54d14c-c43b-5ebc-9be8-4bddd15462bc/scratchpad/showa-spring-proposal.pptx" })
  .then(f => console.log("written:", f));
