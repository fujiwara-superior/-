const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "株式会社スペリオル";
pres.company = "株式会社スペリオル";
pres.title = "建設業 特別教育 オンライン受講プラットフォーム構築のご提案";

// ---- Palette (construction / safety) ----
const NAVY  = "16324F";
const DARK  = "0C1D2E";
const STEEL = "3E6B8C";
const AMBER = "E8960C";
const AMBERD= "B87400";
const AMBERL= "FCF1DC";
const LIGHT = "F3F5F7";
const CARD  = "FFFFFF";
const INK   = "1C2A38";
const MUTED = "6B7B8A";
const LINE  = "DCE3E9";
const GREEN = "2E7D5B";
const RED   = "B3402F";
const JP    = "Meiryo";

const W = 13.333, H = 7.5, M = 0.6, CW = W - M * 2;
let pageNo = 1;

const sh = () => ({ type: "outer", angle: 90, blur: 9, offset: 1, color: "8FA3B4", opacity: 0.3 });

function footer(s) {
  pageNo += 1;
  s.addText("株式会社聖建 御中 ｜ 建設業 特別教育 オンライン受講プラットフォーム ご提案", {
    x: M, y: 6.98, w: 9.5, h: 0.28, fontSize: 9, color: MUTED, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addText(String(pageNo), {
    x: W - M - 0.7, y: 6.98, w: 0.7, h: 0.28, align: "right", fontSize: 10, bold: true, color: STEEL,
    fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
}

function contentSlide(tag, title, lead) {
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  s.addText(tag, { x: M, y: 0.34, w: 8, h: 0.3, fontSize: 11, bold: true, color: AMBERD,
    fontFace: JP, isTextBox: true, margin: 0, charSpacing: 1.5, valign: "middle" });
  s.addText(title, { x: M, y: 0.66, w: CW, h: 0.6, fontSize: 27, bold: true, color: NAVY,
    fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  if (lead) {
    s.addText(lead, { x: M, y: 1.3, w: CW, h: 0.36, fontSize: 12.5, color: MUTED,
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  }
  footer(s);
  return s;
}

function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: fill || CARD },
    rectRadius: 0.05, line: { color: LINE, width: 0.5 }, shadow: sh() });
}

function circle(s, x, y, d, label, fill, txt, fs) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill || AMBER } });
  s.addText(label, { x, y, w: d, h: d, align: "center", valign: "middle", fontSize: fs || 13,
    bold: true, color: txt || "FFFFFF", fontFace: JP, isTextBox: true, margin: 0 });
}

function tbl(s, rows, opts) {
  s.addTable(rows, Object.assign({
    x: M, w: CW, fontFace: JP, fontSize: 11, color: INK, valign: "middle",
    border: [{ type: "solid", color: LINE, pt: 0.5 }, { type: "solid", color: LINE, pt: 0.5 },
             { type: "solid", color: LINE, pt: 0.5 }, { type: "solid", color: LINE, pt: 0.5 }],
    autoPage: false
  }, opts));
}

const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 11.5, align: "center" } });

// =====================================================================
// 1. Cover
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };
  s.addShape(pres.ShapeType.ellipse, { x: 10.4, y: -1.5, w: 5.2, h: 5.2, fill: { color: NAVY } });
  s.addShape(pres.ShapeType.ellipse, { x: 11.9, y: 5.3, w: 2.6, h: 2.6, fill: { color: NAVY } });
  s.addShape(pres.ShapeType.roundRect, { x: 0.9, y: 1.02, w: 2.6, h: 0.4, rectRadius: 0.2,
    fill: { color: AMBER } });
  s.addText("株式会社聖建 御中", { x: 0.9, y: 1.02, w: 2.6, h: 0.4, align: "center", valign: "middle",
    fontSize: 12, bold: true, color: DARK, fontFace: JP, isTextBox: true, margin: 0 });

  s.addText("建設業 特別教育", { x: 0.9, y: 1.85, w: 10.5, h: 0.95, fontSize: 44, bold: true,
    color: "FFFFFF", fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("オンライン受講プラットフォーム 構築のご提案", { x: 0.9, y: 2.82, w: 10.8, h: 0.8,
    fontSize: 30, bold: true, color: AMBER, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });

  s.addText("学科eラーニング  ×  本人確認  ×  実技対面 ─ 新規事業としての立ち上げ", {
    x: 0.9, y: 3.95, w: 10.5, h: 0.4, fontSize: 15, color: "C3D2DF", fontFace: JP, isTextBox: true, margin: 0 });

  const feats = ["受講の事実を証明できる仕組み", "法令要件を運用でなく仕組みで担保", "段階投資で始める新規事業"];
  feats.forEach((t, i) => {
    const x = 0.9 + i * 3.75;
    s.addShape(pres.ShapeType.roundRect, { x, y: 4.72, w: 3.5, h: 0.62, rectRadius: 0.05,
      fill: { color: NAVY } });
    s.addText(t, { x: x + 0.15, y: 4.72, w: 3.2, h: 0.62, fontSize: 11.5, color: "E8F0F7",
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  });

  s.addText("2026年8月30日", { x: 0.9, y: 6.15, w: 5, h: 0.3, fontSize: 12, color: "8FA6BA",
    fontFace: JP, isTextBox: true, margin: 0 });
  s.addText("株式会社スペリオル", { x: 0.9, y: 6.48, w: 5, h: 0.35, fontSize: 15, bold: true,
    color: "FFFFFF", fontFace: JP, isTextBox: true, margin: 0 });
  s.addNotes("本日は、聖建様が新規事業として建設業の特別教育をオンライン提供するための仕組みをご提案します。結論から申し上げると、実現可能です。");
}

// =====================================================================
// 2. Executive summary
// =====================================================================
{
  const s = contentSlide("EXECUTIVE SUMMARY", "ご提案の結論", "いただいたご回答をもとに、実現可能な形へ落とし込みました。");
  const items = [
    ["01", "結論：実現できます", "特別教育は、技能講習と違い行政の登録が不要です。聖建様が「教育を受ける側」から「教育を提供する側」へ回ることができます。"],
    ["02", "学科はオンライン、実技は対面", "厚生労働省の通達により学科のeラーニング化は適法です。実技は半田市の自社会場で実施し、ワンストップで完結させます。"],
    ["03", "本人確認そのものが商品価値", "「誰が・いつ・何時間・どの科目を受けたか」を証拠で示せること。それが、価格ではなく信頼で選ばれる理由になります。"],
    ["04", "小さく始めて、広げる", "まず1科目・学科のみで開始し、実績を作ってから科目と実技会場を広げます。初期投資を抑えた段階導入を推奨します。"]
  ];
  const cw = (CW - 0.3) / 2, ch = 1.85;
  items.forEach((it, i) => {
    const x = M + (i % 2) * (cw + 0.3);
    const y = 1.85 + Math.floor(i / 2) * (ch + 0.28);
    card(s, x, y, cw, ch);
    circle(s, x + 0.28, y + 0.3, 0.52, it[0], AMBER, "FFFFFF", 13);
    s.addText(it[1], { x: x + 0.95, y: y + 0.28, w: cw - 1.25, h: 0.42, fontSize: 15, bold: true,
      color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(it[2], { x: x + 0.95, y: y + 0.74, w: cw - 1.25, h: 0.95, fontSize: 11.5, color: INK,
      fontFace: JP, isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.25 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 6.05, w: CW, h: 0.75, rectRadius: 0.05, fill: { color: NAVY } });
  s.addText([
    { text: "推奨ロードマップ： ", options: { bold: true, color: AMBER } },
    { text: "「フルハーネス（またはテールゲートリフター）1科目・学科オンライン」から開始し、実技提供と科目追加で単価と継続率を上げていく", options: { color: "FFFFFF" } }
  ], { x: M + 0.25, y: 6.05, w: CW - 0.5, h: 0.75, fontSize: 12, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addNotes("4点に要約しています。特に03が今回の肝で、お困りごととして挙がった本人確認が、そのまま商品の差別化要素になります。");
}

// =====================================================================
// 3. Hearing results
// =====================================================================
{
  const s = contentSlide("STEP 0", "ヒアリング結果の確認", "8月30日にご回答いただいた内容と、本提案への反映箇所です。");
  const rows = [
    [hdr("ご確認事項"), hdr("聖建様のご回答"), hdr("本提案への反映")],
    ["受講される方の範囲", { text: "C. 外部の建設会社へ有料提供（新規事業）", options: { bold: true, color: NAVY } },
     "決済・企業アカウント・修了証発行を含む「事業パッケージ」として設計（P.8・P.14）"],
    ["講師と教材の確保", { text: "D. 未定。選択肢の比較から相談したい", options: { bold: true, color: NAVY } },
     "3案を費用・利益率・リスクで比較し、推奨シナリオをご提示（P.12・P.13）"],
    ["本人確認・受講確認の水準", { text: "A. 標準案（書類＋顔写真＋ランダム撮影＋テスト）", options: { bold: true, color: NAVY } },
     "5層の受講確認モデルとして具体化。法令要件との対応も明示（P.7）"],
    ["対象とする講習科目", { text: "B. 学科はオンライン、実技は自社で対面実施", options: { bold: true, color: NAVY } },
     "学科と実技を分離した2プラン構成。商圏の広さと単価を両立（P.14）"],
    ["予算・稼働時期・記録管理", { text: "いずれも未定", options: { bold: true, color: MUTED } },
     "段階投資型のフェーズ計画（P.16）と補助金の活用可能性（P.19）をご提示"]
  ];
  tbl(s, rows, { y: 1.85, colW: [2.7, 4.0, 5.433], rowH: [0.42, 0.78, 0.78, 0.78, 0.78, 0.78],
    fill: { color: CARD }, fontSize: 11 });
  s.addText("※ 「弊社でも調べたうえで当日ご質問したい」とのご連絡をいただいております。想定される論点はP.20にまとめました。",
    { x: M, y: 6.35, w: CW, h: 0.3, fontSize: 10.5, color: MUTED, fontFace: JP, isTextBox: true, margin: 0 });
  s.addNotes("まず認識のズレがないか確認させてください。ご回答はすべて本提案に反映しています。");
}

// =====================================================================
// 4. Market opportunity
// =====================================================================
{
  const s = contentSlide("MARKET", "なぜ、いま参入できるのか", "商圏と需要は、すでに数字で確認できます。");
  const stats = [
    ["28,423者", "愛知県内の建設業許可業者数", "2026年3月末時点。前年比 +0.8% と増加傾向", 26],
    ["8,000〜15,000円", "特別教育1科目あたりの受講料相場", "1名あたり。科目・実施形態により変動", 20],
    ["2024年2月", "テールゲートリフター特別教育が義務化", "学科4時間＋実技2時間。新たな受講需要が発生", 26]
  ];
  const cw = (CW - 0.5) / 3;
  stats.forEach((st, i) => {
    const x = M + i * (cw + 0.25);
    card(s, x, 1.85, cw, 2.0);
    s.addText(st[0], { x: x + 0.28, y: 2.05, w: cw - 0.56, h: 0.62, fontSize: st[3], bold: true,
      color: AMBERD, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(st[1], { x: x + 0.28, y: 2.78, w: cw - 0.56, h: 0.6, fontSize: 12.5, bold: true,
      color: NAVY, fontFace: JP, isTextBox: true, margin: 0 });
    s.addText(st[2], { x: x + 0.28, y: 3.38, w: cw - 0.56, h: 0.4, fontSize: 10, color: MUTED,
      fontFace: JP, isTextBox: true, margin: 0 });
  });
  card(s, M, 4.1, CW, 2.4);
  s.addText("外部委託のニーズが生まれる構造", { x: M + 0.35, y: 4.28, w: 6, h: 0.4, fontSize: 15,
    bold: true, color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  const pts = [
    "特別教育の実施義務は「事業者」にあり、未実施の場合の罰則も事業者に科されます",
    "一方で、従業員1〜29名規模の事業者が講師・教材・記録管理を自前で整えるのは負担が大きい",
    "半田市は知多半島・名古屋港臨海部・中部国際空港圏を商圏に持ち、建設・物流の需要が厚い",
    "学科をオンライン化すれば商圏は全国へ。実技は東海圏の来場需要を単価の高い商品として取り込める"
  ];
  pts.forEach((p, i) => {
    const y = 4.78 + i * 0.4;
    s.addShape(pres.ShapeType.ellipse, { x: M + 0.4, y: y + 0.11, w: 0.13, h: 0.13, fill: { color: AMBER } });
    s.addText(p, { x: M + 0.68, y: y, w: CW - 1.1, h: 0.36, fontSize: 11.5, color: INK,
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  });
  s.addNotes("市場は十分にあります。特にテールゲートリフターは2024年に義務化されたばかりで、まだ受講が行き渡っていません。");
}

// =====================================================================
// 5. Why Seiken can enter (comparison)
// =====================================================================
{
  const s = contentSlide("LEGAL BASIS", "特別教育は「登録なし」で提供できます", "免許・技能講習との決定的な違いが、聖建様の参入余地です。");
  const rows = [
    [hdr(""), hdr("免許"), hdr("技能講習"), hdr("特別教育（今回の対象）")],
    [{ text: "実施できる者", options: { bold: true, fill: { color: LIGHT } } },
     "都道府県労働局が試験を実施", "労働局長の登録を受けた登録教習機関のみ",
     { text: "事業者、または事業者から委託を受けた者", options: { bold: true, color: NAVY, fill: { color: AMBERL } } }],
    [{ text: "行政への登録", options: { bold: true, fill: { color: LIGHT } } },
     "─", "必要（登録教習機関）",
     { text: "不要", options: { bold: true, color: GREEN, fill: { color: AMBERL } } }],
    [{ text: "講師の要件", options: { bold: true, fill: { color: LIGHT } } },
     "─", "法令で定める要件あり",
     { text: "法定の資格要件なし。ただし「十分な知識と経験を有する者」であることが必要",
       options: { color: NAVY, fill: { color: AMBERL } } }],
    [{ text: "聖建様の参入", options: { bold: true, fill: { color: LIGHT } } },
     { text: "参入不可", options: { color: MUTED } }, { text: "登録のハードルが高い", options: { color: MUTED } },
     { text: "参入できます", options: { bold: true, color: GREEN, fill: { color: AMBERL } } }]
  ];
  tbl(s, rows, { y: 1.85, colW: [2.2, 2.6, 3.2, 4.133], rowH: [0.4, 0.72, 0.55, 0.85, 0.55],
    fill: { color: CARD }, fontSize: 11 });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 5.2, w: CW, h: 1.22, rectRadius: 0.05, fill: { color: NAVY } });
  circle(s, M + 0.3, 5.5, 0.6, "!", AMBER, DARK, 20);
  s.addText("登録制度がない領域だからこそ、「証明できること」が競争軸になります", {
    x: M + 1.1, y: 5.38, w: CW - 1.5, h: 0.38, fontSize: 14.5, bold: true, color: AMBER,
    fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("誰でも名乗れる以上、発注元や監督機関から「本当に受講させたのか」を問われた際に即答できるかどうかが、選ばれる事業者との分かれ目になります。",
    { x: M + 1.1, y: 5.78, w: CW - 1.5, h: 0.55, fontSize: 11.5, color: "D6E2EC",
      fontFace: JP, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  s.addNotes("技能講習と混同されがちですが、特別教育は登録不要です。ここが参入できる根拠であり、同時に競合も入りやすい領域だという意味でもあります。");
}

// =====================================================================
// 6. Three legal requirements
// =====================================================================
{
  const s = contentSlide("REQUIREMENTS", "越えるべき法令上の3つの要件", "この3つを、担当者の努力ではなく「仕組み」で満たします。");
  const blocks = [
    ["01", "学科をeラーニングで行う要件", "令和3年1月25日 基安安発0125第2号 ほか", [
      "法定の科目の範囲・教育時間・講師要件を満たすこと",
      "教本等、必要な教材を用いて行うこと",
      "受講者が受講した事実を適切に確認すること",
      "受講者からの質問に対応できる体制を設けること（実務上の要請）"
    ]],
    ["02", "実技は対面で実施", "聖建様のご選択：B（実技は自社で対面）", [
      "実技科目はオンラインで代替できません",
      "実施日・場所・指導者・受講者を記録する必要があります",
      "半田市の自社会場で実施し、学科と紐づけて一元管理します",
      "遠方の受講者には「学科のみ」プランを用意して取りこぼしを防ぎます"
    ]],
    ["03", "記録を3年間保存", "労働安全衛生規則 第38条", [
      "受講者氏名・科目・実施日・講師名などを記録",
      "保存義務は事業者側にありますが、教育機関の記録が実務上の証明になります",
      "システムで自動保存し、いつでもCSV／PDFで出力できるようにします"
    ]]
  ];
  const cw = (CW - 0.5) / 3;
  blocks.forEach((b, i) => {
    const x = M + i * (cw + 0.25);
    card(s, x, 1.85, cw, 4.55);
    circle(s, x + 0.3, 2.1, 0.56, b[0], NAVY, "FFFFFF", 14);
    s.addText(b[1], { x: x + 0.3, y: 2.82, w: cw - 0.6, h: 0.6, fontSize: 14.5, bold: true,
      color: NAVY, fontFace: JP, isTextBox: true, margin: 0 });
    s.addText(b[2], { x: x + 0.3, y: 3.44, w: cw - 0.6, h: 0.5, fontSize: 9.5, color: AMBERD,
      fontFace: JP, isTextBox: true, margin: 0 });
    s.addText(b[3].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < b[3].length - 1 } })),
      { x: x + 0.32, y: 4.0, w: cw - 0.62, h: 2.25, fontSize: 10.5, color: INK, fontFace: JP,
        isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 7, lineSpacingMultiple: 1.15 });
  });
  s.addNotes("この3つがクリアできれば、法令上は問題なく提供できます。3番目の記録保存はご相談時に未定とのことでしたので、システム側で自動化する前提で設計しています。");
}

// =====================================================================
// 7. Five-layer verification (core)
// =====================================================================
{
  const s = contentSlide("CORE", "本人確認・受講確認の5層モデル", "ご選択いただいた「標準案A」を、具体的な仕組みに落とし込みました。");
  const layers = [
    ["1", "受講前", "本人確認書類＋顔写真の登録", "運転免許証等をアップロード。顔写真と照合し、運営が承認するまで受講を開始できません。"],
    ["2", "受講中", "ランダムなタイミングでの自撮り撮影", "学科視聴中、予告なくカメラ撮影。登録顔写真と照合し、離席・交代を検知します。"],
    ["3", "受講中", "視聴ログの秒単位記録", "スキップ・早送りを禁止。法定の教育時間を満たしたことをログで証明します。"],
    ["4", "章末", "章ごとの理解度テスト", "合格しなければ次章に進めません。「再生しているだけ」の受講を排除します。"],
    ["5", "修了時", "修了確認テスト＋質問受付", "合格者にのみ修了証を発行。質問受付により法令上の体制要件も満たします。"]
  ];
  const lw = 7.55;
  layers.forEach((l, i) => {
    const y = 1.85 + i * 0.93;
    card(s, M, y, lw, 0.82);
    circle(s, M + 0.2, y + 0.16, 0.5, l[0], AMBER, "FFFFFF", 14);
    s.addShape(pres.ShapeType.roundRect, { x: M + 0.82, y: y + 0.14, w: 0.78, h: 0.28, rectRadius: 0.14,
      fill: { color: LIGHT } });
    s.addText(l[1], { x: M + 0.82, y: y + 0.14, w: 0.78, h: 0.28, align: "center", valign: "middle",
      fontSize: 9, bold: true, color: STEEL, fontFace: JP, isTextBox: true, margin: 0 });
    s.addText(l[2], { x: M + 1.7, y: y + 0.11, w: lw - 1.9, h: 0.32, fontSize: 12.5, bold: true,
      color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(l[3], { x: M + 0.84, y: y + 0.44, w: lw - 1.05, h: 0.32, fontSize: 10, color: INK,
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  });
  const rx = M + lw + 0.28, rw = CW - lw - 0.28;
  s.addShape(pres.ShapeType.roundRect, { x: rx, y: 1.85, w: rw, h: 4.5, rectRadius: 0.05, fill: { color: NAVY } });
  s.addText("この仕組みで、こう答えられます", { x: rx + 0.28, y: 2.05, w: rw - 0.56, h: 0.4,
    fontSize: 14, bold: true, color: AMBER, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addShape(pres.ShapeType.roundRect, { x: rx + 0.28, y: 2.55, w: rw - 0.56, h: 1.55, rectRadius: 0.05,
    fill: { color: DARK } });
  s.addText("「2026年11月12日、○○建設の△△様が、フルハーネス特別教育の学科を4時間52分受講。受講中の顔照合12回はすべて本人と一致。修了確認テスト92点で合格。」",
    { x: rx + 0.45, y: 2.68, w: rw - 0.9, h: 1.3, fontSize: 11.5, color: "FFFFFF", fontFace: JP,
      isTextBox: true, margin: 0, italic: true, lineSpacingMultiple: 1.3 });
  const benes = ["発注元・元請への説明責任を果たせる", "監督署の調査にその場で応じられる", "受講者本人の「受けた証明」になる", "他社との差別化がそのまま価格に反映できる"];
  benes.forEach((b, i) => {
    const y = 4.3 + i * 0.44;
    s.addShape(pres.ShapeType.ellipse, { x: rx + 0.32, y: y + 0.11, w: 0.13, h: 0.13, fill: { color: AMBER } });
    s.addText(b, { x: rx + 0.58, y: y, w: rw - 0.9, h: 0.38, fontSize: 11, color: "E2ECF4",
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  });
  s.addNotes("ここが本提案の中核です。お困りごととして挙げられた本人確認と受講確認は、法令要件そのものであり、同時に商品の価値でもあります。");
}

// =====================================================================
// 8. System overview
// =====================================================================
{
  const s = contentSlide("SYSTEM", "システム全体像", "3者それぞれに専用の画面を用意し、運営の手間を最小化します。");
  const users = [
    ["受講者（作業員）", "スマートフォンで受講", ["本人確認書類・顔写真の登録", "学科動画の視聴とテスト", "実技日程の予約", "修了証の閲覧・ダウンロード"]],
    ["発注企業のご担当者", "従業員の受講を管理", ["従業員をまとめて申込", "受講進捗のリアルタイム確認", "修了者一覧のCSV出力", "請求書の受領"]],
    ["聖建様（運営）", "事業の運営・管理", ["本人確認の承認", "顔照合アラートの確認", "実技日程・出欠の管理", "修了証発行・売上管理"]]
  ];
  const cw = (CW - 0.5) / 3;
  users.forEach((u, i) => {
    const x = M + i * (cw + 0.25);
    const isOwner = i === 2;
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.8, w: cw, h: 1.95, rectRadius: 0.05,
      fill: { color: isOwner ? NAVY : CARD }, line: { color: isOwner ? NAVY : LINE, width: 0.5 }, shadow: sh() });
    s.addText(u[0], { x: x + 0.28, y: 1.95, w: cw - 0.56, h: 0.36, fontSize: 14, bold: true,
      color: isOwner ? AMBER : NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(u[1], { x: x + 0.28, y: 2.29, w: cw - 0.56, h: 0.28, fontSize: 10, 
      color: isOwner ? "AFC3D4" : MUTED, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(u[2].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < u[2].length - 1 } })),
      { x: x + 0.3, y: 2.62, w: cw - 0.6, h: 1.05, fontSize: 10, color: isOwner ? "E2ECF4" : INK,
        fontFace: JP, isTextBox: true, margin: 0, paraSpaceAfter: 3 });
  });
  s.addText("プラットフォームが提供する主要機能", { x: M, y: 4.05, w: 8, h: 0.35, fontSize: 14,
    bold: true, color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  const funcs = ["学科動画の配信", "本人確認・顔照合", "視聴ログの記録", "テスト・自動採点",
                 "実技予約・出欠管理", "修了証PDF発行", "決済（カード・請求書）", "記録の3年以上保存"];
  const fw = (CW - 0.6) / 4;
  funcs.forEach((f, i) => {
    const x = M + (i % 4) * (fw + 0.2);
    const y = 4.58 + Math.floor(i / 4) * 0.96;
    card(s, x, y, fw, 0.84);
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.22, y: y + 0.33, w: 0.34, h: 0.18, rectRadius: 0.09,
      fill: { color: AMBER } });
    s.addText(f, { x: x + 0.68, y: y, w: fw - 0.9, h: 0.84, fontSize: 11, bold: true, color: INK,
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  });
  s.addNotes("聖建様の運営負荷を下げることを重視しています。特に本人確認の承認と実技日程の管理は、日々の運用で最も手間がかかる部分です。");
}

// =====================================================================
// 9. Learner flow
// =====================================================================
{
  const s = contentSlide("FLOW", "受講の流れ", "学科はオンラインで完結し、実技の実施形態でプランが分岐します。");
  const steps = ["①\nお申込み\n（企業一括／個人）", "②\nアカウント登録", "③\n本人確認書類・\n顔写真の登録",
                 "④\n運営による承認", "⑤\n学科eラーニング\n受講", "⑥\n修了確認テスト\n合格"];
  const bw = (CW - 5 * 0.16) / 6;
  steps.forEach((t, i) => {
    const x = M + i * (bw + 0.16);
    card(s, x, 1.85, bw, 1.15);
    s.addText(t, { x: x + 0.08, y: 1.85, w: bw - 0.16, h: 1.15, align: "center", valign: "middle",
      fontSize: 10.5, bold: true, color: NAVY, fontFace: JP, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    if (i < 5) {
      s.addShape(pres.ShapeType.triangle, { x: x + bw + 0.035, y: 2.32, w: 0.09, h: 0.2,
        fill: { color: AMBER }, rotate: 90 });
    }
  });
  const bw2 = (CW - 0.3) / 2;
  const branches = [
    ["プランA｜学科のみ", STEEL, ["⑦ 実技は受講者の所属事業者が実施", "⑧ 実技実施報告をシステムにアップロード", "→ 商圏は全国。実技会場の制約を受けません"]],
    ["プランB｜学科＋実技", AMBERD, ["⑦ 半田会場の実技日程を予約", "⑧ 来場して実技を受講（講師が実施記録を登録）", "→ 単価が高く、他社にない一括提供が差別化に"]]
  ];
  branches.forEach((b, i) => {
    const x = M + i * (bw2 + 0.3);
    card(s, x, 3.35, bw2, 1.62);
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.25, y: 3.52, w: 2.5, h: 0.34, rectRadius: 0.17,
      fill: { color: b[1] } });
    s.addText(b[0], { x: x + 0.25, y: 3.52, w: 2.5, h: 0.34, align: "center", valign: "middle",
      fontSize: 11, bold: true, color: "FFFFFF", fontFace: JP, isTextBox: true, margin: 0 });
    s.addText(b[2].map((t, j) => ({ text: t, options: { breakLine: j < b[2].length - 1,
      bold: j === 2, color: j === 2 ? b[1] : INK } })),
      { x: x + 0.28, y: 3.96, w: bw2 - 0.56, h: 0.92, fontSize: 10.5, fontFace: JP,
        isTextBox: true, margin: 0, paraSpaceAfter: 4 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 5.3, w: CW, h: 1.1, rectRadius: 0.05, fill: { color: NAVY } });
  const fin = [["⑨", "修了証の発行", "PDF／カードで発行。受講者・企業双方がいつでも再取得できます"],
               ["⑩", "記録を3年以上保存", "労働安全衛生規則第38条に対応。CSV一括出力にも対応します"]];
  fin.forEach((f, i) => {
    const x = M + 0.3 + i * 6.0;
    circle(s, x, 5.55, 0.6, f[0], AMBER, DARK, 15);
    s.addText(f[1], { x: x + 0.75, y: 5.5, w: 5.0, h: 0.32, fontSize: 13, bold: true, color: "FFFFFF",
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(f[2], { x: x + 0.75, y: 5.82, w: 5.0, h: 0.34, fontSize: 10, color: "C3D2DF",
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  });
  s.addNotes("プランAを用意することで、実技会場のキャパシティに売上が縛られなくなります。これは事業の伸ばし方として重要な設計です。");
}

// =====================================================================
// 10. Screen composition
// =====================================================================
{
  const s = contentSlide("SCREENS", "画面構成のイメージ", "3つのポータルを用意します。デザインは要件定義後に確定します。");
  const screens = [
    ["受講者ポータル", "スマートフォン対応", ["マイページ（受講状況の一覧）", "本人確認書類・顔写真の登録", "学科動画の視聴画面", "章末テスト・修了確認テスト", "実技日程の予約", "修了証のダウンロード"]],
    ["企業管理ポータル", "発注企業のご担当者向け", ["従業員一覧の管理", "複数名の一括申込", "受講進捗のリアルタイム確認", "修了者一覧のCSV出力", "請求書・領収書の確認", "追加受講の申込"]],
    ["運営管理画面", "聖建様向け", ["本人確認の承認キュー", "顔照合アラートの確認", "受講ログの照会・出力", "実技日程と出欠の管理", "修了証の発行・再発行", "売上・入金の管理"]]
  ];
  const cw = (CW - 0.5) / 3;
  screens.forEach((sc, i) => {
    const x = M + i * (cw + 0.25);
    card(s, x, 1.85, cw, 4.25);
    s.addShape(pres.ShapeType.roundRect, { x: x, y: 1.85, w: cw, h: 0.8, rectRadius: 0.05,
      fill: { color: i === 2 ? NAVY : STEEL } });
    s.addText(sc[0], { x: x + 0.28, y: 1.93, w: cw - 0.56, h: 0.36, fontSize: 14.5, bold: true,
      color: "FFFFFF", fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(sc[1], { x: x + 0.28, y: 2.28, w: cw - 0.56, h: 0.28, fontSize: 10, color: "D2E0EA",
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    sc[2].forEach((t, j) => {
      const y = 2.85 + j * 0.52;
      s.addShape(pres.ShapeType.roundRect, { x: x + 0.24, y: y, w: cw - 0.48, h: 0.42,
        rectRadius: 0.04, fill: { color: LIGHT } });
      s.addText(t, { x: x + 0.4, y: y, w: cw - 0.7, h: 0.42, fontSize: 10.5, color: INK,
        fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    });
  });
  s.addText("※ 本ページは機能構成のイメージです。実際の画面デザイン・項目はフェーズ0（要件定義）で確定します。",
    { x: M, y: 6.25, w: CW, h: 0.3, fontSize: 10.5, color: MUTED, fontFace: JP, isTextBox: true, margin: 0 });
  s.addNotes("企業管理ポータルがあると、発注企業のご担当者が自分で進捗を見られるため、聖建様への問い合わせ対応が大きく減ります。");
}

// =====================================================================
// 11. Audit response
// =====================================================================
{
  const s = contentSlide("VALUE", "「本当に受講したのか」に、証拠で答える", "受講確認の仕組みは、そのまま聖建様の営業材料になります。");
  const items = [
    ["受講証明の即時出力", "受講者ごとにPDFの受講証明を発行。科目・実施日・視聴時間・顔照合の結果・テスト得点までを1枚に記載します。"],
    ["企業単位での一覧管理", "発注企業のご担当者が、自社従業員の修了状況をいつでも確認できます。更新時期の管理にも使えます。"],
    ["記録の自動保存", "法定の3年を超えて自動保管。監督署の調査や元請からの要請にも、その場で応じられます。"],
    ["異常の自動検知", "顔照合の不一致、視聴時間の不足、テストの不自然な回答パターンを検知してアラートを出します。"]
  ];
  items.forEach((it, i) => {
    const y = 1.85 + i * 1.02;
    card(s, M, y, CW, 0.9);
    circle(s, M + 0.3, y + 0.19, 0.52, String(i + 1), NAVY, "FFFFFF", 14);
    s.addText(it[0], { x: M + 1.0, y: y + 0.13, w: 3.4, h: 0.34, fontSize: 14, bold: true, color: NAVY,
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(it[1], { x: M + 1.0, y: y + 0.47, w: CW - 1.4, h: 0.34, fontSize: 11, color: INK,
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 6.0, w: CW, h: 0.62, rectRadius: 0.05, fill: { color: AMBERL } });
  s.addText("競合が「安さ」で競うなかで、聖建様は「証明できること」で競う。これが価格を守る唯一の方法です。",
    { x: M + 0.3, y: 6.0, w: CW - 0.6, h: 0.62, fontSize: 12.5, bold: true, color: AMBERD,
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addNotes("ここは営業トークとしてそのまま使える部分です。他社が価格を下げてきても、証明力で差をつけられます。");
}

// =====================================================================
// 12. Instructor / material options
// =====================================================================
{
  const s = contentSlide("ANSWER TO Q2", "講師と教材の確保 ─ 3案の比較", "「未定・比較から相談したい」とのご回答を受け、費用・利益率・リスクで整理しました。");
  const rows = [
    [hdr("案"), hdr("内容"), hdr("初期費用"), hdr("継続費用"), hdr("強み"), hdr("留意点")],
    [{ text: "A\n自社講師", options: { bold: true, color: NAVY, fill: { color: LIGHT }, align: "center" } },
     "聖建様の有資格者・実務経験者が講師を務め、講義を撮影して教材化する",
     { text: "中\n（撮影・編集）", options: { align: "center" } }, { text: "低", options: { align: "center", color: GREEN, bold: true } },
     "利益率が最も高い。現場感のある独自教材になり、他社が真似できない",
     "「十分な知識と経験」を説明できる資料の整備が必要。教材制作の工数がかかる"],
    [{ text: "B\n外部講師に委託\n◎ 推奨", options: { bold: true, color: NAVY, fill: { color: AMBERL }, align: "center" } },
     "労働安全衛生コンサルタント等に講師を依頼し、講義を収録して教材化する",
     { text: "中〜高\n（講師料＋撮影）", options: { align: "center" } }, { text: "低\n（教材は資産化）", options: { align: "center", color: GREEN, bold: true } },
     "講師要件の説明が最も明確。立ち上げ期の信頼性を確保しやすい",
     "科目を増やすたびに講師の手配が必要になる"],
    [{ text: "C\n市販教材を調達", options: { bold: true, color: NAVY, fill: { color: LIGHT }, align: "center" } },
     "既存のeラーニング教材をライセンス調達し、システム側に注力する",
     { text: "低", options: { align: "center", color: GREEN, bold: true } }, { text: "高\n（1名ごとの\nライセンス料）", options: { align: "center", color: RED, bold: true } },
     "最短で開始できる。教材制作の負担がない",
     { text: "利益率が薄い。再販（外部提供）が許諾されない契約が多く、今回の事業モデルとは相性が悪い",
       options: { color: RED, bold: true } }]
  ];
  tbl(s, rows, { y: 1.9, colW: [1.5, 2.9, 1.35, 1.35, 2.5, 2.533], rowH: [0.42, 1.35, 1.35, 1.35],
    fill: { color: CARD }, fontSize: 10 });
  s.addText("※ 特別教育の講師に法定の資格要件はありませんが、「十分な知識と経験を有する者」であることが求められます。いずれの案でも講師の経歴を記録に残す運用が必要です。",
    { x: M, y: 6.4, w: CW, h: 0.32, fontSize: 10, color: MUTED, fontFace: JP, isTextBox: true, margin: 0 });
  s.addNotes("C案は一見安く見えますが、外販が前提の今回のモデルでは契約上使えない可能性が高いです。ここは必ず確認が必要な論点です。");
}

// =====================================================================
// 13. Recommended scenario
// =====================================================================
{
  const s = contentSlide("RECOMMENDATION", "推奨：B案で始め、A案へ移行する", "立ち上げ期は信頼性を、軌道に乗ってからは利益率を取りにいきます。");
  const steps = [
    ["STEP 1", "1〜3か月", "外部講師を招いて第1科目を収録", "講師要件の担保が最も明確な形で立ち上げます。収録した講義は聖建様の資産として残ります。"],
    ["STEP 2", "4〜9か月", "収録教材で運営開始・実績づくり", "受講者の反応、問い合わせの内容、実技会場の回し方など、運用ノウハウを蓄積します。"],
    ["STEP 3", "10か月〜", "2科目目以降は自社講師へ移行", "社内のベテランが講師を務め、講師料を圧縮。利益率を引き上げながら科目を増やします。"]
  ];
  const cw = (CW - 0.5) / 3;
  steps.forEach((st, i) => {
    const x = M + i * (cw + 0.25);
    card(s, x, 1.85, cw, 2.85);
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.28, y: 2.08, w: 1.15, h: 0.32, rectRadius: 0.16,
      fill: { color: AMBER } });
    s.addText(st[0], { x: x + 0.28, y: 2.08, w: 1.15, h: 0.32, align: "center", valign: "middle",
      fontSize: 10, bold: true, color: DARK, fontFace: JP, isTextBox: true, margin: 0 });
    s.addText(st[1], { x: x + 1.52, y: 2.08, w: cw - 1.8, h: 0.32, fontSize: 11, bold: true,
      color: MUTED, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(st[2], { x: x + 0.28, y: 2.55, w: cw - 0.56, h: 0.85, fontSize: 14, bold: true,
      color: NAVY, fontFace: JP, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
    s.addText(st[3], { x: x + 0.28, y: 3.45, w: cw - 0.56, h: 1.1, fontSize: 11, color: INK,
      fontFace: JP, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
    if (i < 2) {
      s.addShape(pres.ShapeType.triangle, { x: x + cw + 0.055, y: 3.15, w: 0.11, h: 0.24,
        fill: { color: STEEL }, rotate: 90 });
    }
  });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 4.95, w: CW, h: 1.45, rectRadius: 0.05, fill: { color: NAVY } });
  s.addText("なぜC案（市販教材の調達）を推奨しないのか", { x: M + 0.35, y: 5.15, w: 8, h: 0.36,
    fontSize: 14, bold: true, color: AMBER, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("市販のeラーニング教材は「自社の従業員に受講させる」ことを前提にライセンスされているものが大半で、第三者への再販・有料提供が許諾されないケースが多く見られます。外部提供を新規事業とする今回のモデルでは、契約上の制約と価格決定権を持てないことが致命的になります。まずは自社の資産となる教材を作ることをお勧めします。",
    { x: M + 0.35, y: 5.55, w: CW - 0.7, h: 0.75, fontSize: 11, color: "D6E2EC", fontFace: JP,
      isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  s.addNotes("結論はB案スタートです。ただし講師の当てがすでにあるならA案から始めても構いません。当日ご相談させてください。");
}

// =====================================================================
// 14. Product design
// =====================================================================
{
  const s = contentSlide("PRODUCT", "商品設計 ─ 3つのプラン", "学科と実技を分けることで、商圏の広さと単価の高さを両立します。");
  const plans = [
    ["プランA", "学科オンラインのみ", "8,000〜10,000円", "／ 1名", STEEL,
      ["実技は受講者の所属事業者が実施", "商圏：全国", "粗利率が高い主力商品", "フルハーネス等で一般的な提供形態"]],
    ["プランB", "学科＋実技（半田会場）", "15,000〜20,000円", "／ 1名", AMBERD,
      ["実技まで一括で完結", "商圏：愛知・岐阜・三重", "単価が高く、差別化になる", "会場のキャパシティが上限になる"]],
    ["プランC", "企業一括・年間契約", "個別お見積り", "／ 年間", NAVY,
      ["複数科目・複数名をまとめて契約", "企業管理ポータル付き", "継続収益（ストック）化できる", "更新時期の案内で解約を防ぐ"]]
  ];
  const cw = (CW - 0.5) / 3;
  plans.forEach((p, i) => {
    const x = M + i * (cw + 0.25);
    card(s, x, 1.85, cw, 3.45);
    s.addShape(pres.ShapeType.roundRect, { x: x, y: 1.85, w: cw, h: 0.62, rectRadius: 0.05, fill: { color: p[4] } });
    s.addText(p[0], { x: x + 0.28, y: 1.85, w: cw - 0.56, h: 0.62, fontSize: 14, bold: true,
      color: "FFFFFF", fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(p[1], { x: x + 0.28, y: 2.62, w: cw - 0.56, h: 0.4, fontSize: 14, bold: true,
      color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(p[2], { x: x + 0.28, y: 3.02, w: cw - 0.56, h: 0.44, fontSize: 21, bold: true,
      color: AMBERD, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(p[3].replace("／", "／ "), { x: x + 0.28, y: 3.46, w: cw - 0.56, h: 0.26, fontSize: 10.5,
      color: MUTED, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(p[5].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < p[5].length - 1 } })),
      { x: x + 0.3, y: 3.86, w: cw - 0.6, h: 1.34, fontSize: 10.5, color: INK, fontFace: JP,
        isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 6 });
  });
  card(s, M, 5.5, CW, 0.9);
  s.addText([{ text: "オプション商品　", options: { bold: true, color: NAVY, fontSize: 12.5 } },
             { text: "修了証カードの発行　／　受講記録の代行保管　／　企業専用コースの制作　／　更新時期のリマインド配信", options: { color: INK, fontSize: 11.5 } }],
    { x: M + 0.35, y: 5.6, w: CW - 0.7, h: 0.34, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("※ 価格は特別教育の一般的な受講料相場（8,000〜15,000円／1名）を踏まえた想定です。確定はフェーズ0で行います。",
    { x: M + 0.35, y: 5.98, w: CW - 0.7, h: 0.32, fontSize: 10, color: MUTED, fontFace: JP,
      isTextBox: true, margin: 0, valign: "middle" });
  s.addNotes("プランCのストック化が、事業を安定させる鍵です。単発の受講販売だけだと、毎年ゼロから営業することになります。");
}

// =====================================================================
// 15. Financial simulation
// =====================================================================
{
  const s = contentSlide("SIMULATION", "収支シミュレーション（仮試算）", "前提を置いた試算です。実数値はフェーズ0で聖建様と一緒に詰めます。");
  s.addChart(pres.ChartType.bar, [
    { name: "プランA 学科オンライン", labels: ["1年目", "2年目", "3年目"], values: [250, 600, 1100] },
    { name: "プランB 学科＋実技", labels: ["1年目", "2年目", "3年目"], values: [90, 360, 720] }
  ], {
    x: M, y: 1.9, w: 7.4, h: 4.45,
    barDir: "col", barGrouping: "stacked", barGapWidthPct: 55,
    chartColors: [NAVY, STEEL],
    showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF",
    dataLabelFontFace: JP, dataLabelFontSize: 11, dataLabelFormatCode: '#,##0"万円"',
    showLegend: true, legendPos: "b", legendFontFace: JP, legendFontSize: 10, legendColor: INK,
    showTitle: false,
    catAxisLabelColor: INK, catAxisLabelFontFace: JP, catAxisLabelFontSize: 12,
    valAxisLabelColor: MUTED, valAxisLabelFontFace: JP, valAxisLabelFontSize: 10,
    valAxisLabelFormatCode: '#,##0',
    valGridLine: { color: "E3E8EC", size: 1 }, catGridLine: { style: "none" },
    catAxisLineShow: false, valAxisLineShow: false
  });
  const rx = M + 7.65, rw = CW - 7.65;
  card(s, rx, 1.9, rw, 1.55);
  s.addText("試算の前提", { x: rx + 0.28, y: 2.02, w: rw - 0.56, h: 0.3, fontSize: 12.5, bold: true,
    color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addText([
    { text: "客単価：学科のみ 10,000円／学科＋実技 18,000円", options: { breakLine: true } },
    { text: "変動費率：学科 15%（決済手数料・本人確認API・配信費）", options: { breakLine: true } },
    { text: "　　　　　実技 45%（講師人件費・会場費）", options: {} }
  ], { x: rx + 0.28, y: 2.36, w: rw - 0.56, h: 0.95, fontSize: 10, color: INK, fontFace: JP,
    isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  s.addShape(pres.ShapeType.roundRect, { x: rx, y: 3.62, w: rw, h: 1.65, rectRadius: 0.05, fill: { color: NAVY } });
  s.addText("想定粗利", { x: rx + 0.28, y: 3.74, w: rw - 0.56, h: 0.3, fontSize: 12.5, bold: true,
    color: AMBER, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  [["1年目", "262万円"], ["2年目", "708万円"], ["3年目", "1,331万円"]].forEach((g, i) => {
    const y = 4.12 + i * 0.36;
    s.addText(g[0], { x: rx + 0.3, y: y, w: 1.2, h: 0.32, fontSize: 11, color: "C3D2DF",
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(g[1], { x: rx + 1.4, y: y, w: rw - 1.7, h: 0.32, align: "right", fontSize: 14, bold: true,
      color: "FFFFFF", fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  });
  s.addShape(pres.ShapeType.roundRect, { x: rx, y: 5.44, w: rw, h: 0.91, rectRadius: 0.05, fill: { color: AMBERL } });
  s.addText("初期投資を500万円と仮定した場合、2年目の途中で累計粗利が投資額を上回る試算です。",
    { x: rx + 0.28, y: 5.44, w: rw - 0.56, h: 0.91, fontSize: 11, bold: true, color: AMBERD,
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.2 });
  s.addText("※ 本試算は一般的な受講料相場に基づく仮置きであり、売上を保証するものではありません。",
    { x: M, y: 6.45, w: CW, h: 0.28, fontSize: 10, color: MUTED, fontFace: JP, isTextBox: true, margin: 0 });
  s.addNotes("あくまで仮試算です。ここは当日、聖建様の営業リソースや既存の取引先の数を伺ったうえで、現実的な数字に置き換えたいと考えています。");
}

// =====================================================================
// 16. Schedule
// =====================================================================
{
  const s = contentSlide("SCHEDULE", "導入スケジュール", "稼働時期は未定とのことでしたので、段階投資型の計画をご提案します。");
  const phases = [
    ["フェーズ0", "1か月", "事業設計・要件定義", "第1弾科目の決定／講師方針の確定／価格・利用規約・特定商取引法表記の整備／個人情報の取扱い方針の策定", NAVY, 2.3],
    ["フェーズ1", "2〜4か月", "MVP構築・テスト運用", "受講者ポータル／本人確認・顔照合／学科動画配信／視聴ログ／テスト／修了証発行／決済", STEEL, 3.6],
    ["フェーズ2", "5〜7か月", "本格運用・機能拡張", "企業管理ポータル／実技予約・出欠管理／一括請求／CSV出力／記録の自動保存", STEEL, 3.6],
    ["フェーズ3", "8か月〜", "事業の拡大", "科目の追加／AI顔認証の強化／他県への展開／企業年間契約の獲得", AMBERD, 2.6]
  ];
  phases.forEach((p, i) => {
    const y = 1.9 + i * 1.15;
    card(s, M, y, CW, 1.0);
    s.addShape(pres.ShapeType.roundRect, { x: M + 0.22, y: y + 0.2, w: 1.28, h: 0.6, rectRadius: 0.05,
      fill: { color: p[4] } });
    s.addText(p[0], { x: M + 0.22, y: y + 0.2, w: 1.28, h: 0.6, align: "center", valign: "middle",
      fontSize: 12, bold: true, color: "FFFFFF", fontFace: JP, isTextBox: true, margin: 0 });
    s.addText(p[1], { x: M + 1.66, y: y + 0.16, w: 1.1, h: 0.3, fontSize: 11, bold: true, color: AMBERD,
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(p[2], { x: M + 2.8, y: y + 0.14, w: 4.0, h: 0.34, fontSize: 14, bold: true, color: NAVY,
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(p[3], { x: M + 1.66, y: y + 0.52, w: CW - 2.0, h: 0.36, fontSize: 10.5, color: INK,
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  });
  s.addText("※ 各フェーズは前フェーズの結果を見て継続を判断できます。フェーズ1で反応が読めなければ、そこで一度立ち止まる判断も可能です。",
    { x: M, y: 6.5, w: CW, h: 0.3, fontSize: 10.5, color: MUTED, fontFace: JP, isTextBox: true, margin: 0 });
  s.addNotes("一括で作り切らず、フェーズごとに判断できる形にしています。予算が未定という状況では、この進め方が最もリスクが小さいと考えます。");
}

// =====================================================================
// 17. Cost
// =====================================================================
{
  const s = contentSlide("COST", "概算費用", "要件が固まる前の概算レンジです。正式なお見積りは要件確定後にご提示します。");
  const rows = [
    [hdr("項目"), hdr("概算レンジ"), hdr("含まれるもの")],
    [{ text: "フェーズ0　事業設計・要件定義", options: { bold: true, color: NAVY } },
     { text: "別途ご相談", options: { align: "center", color: MUTED } },
     "ヒアリング、業務フロー設計、要件定義書、画面設計、規約・表記類の整備支援"],
    [{ text: "フェーズ1　MVP構築", options: { bold: true, color: NAVY } },
     { text: "250〜400万円", options: { align: "center", bold: true, color: AMBERD } },
     "受講者ポータル、本人確認・顔照合、学科動画配信、視聴ログ、テスト、修了証発行、決済"],
    [{ text: "フェーズ2　機能拡張", options: { bold: true, color: NAVY } },
     { text: "150〜250万円", options: { align: "center", bold: true, color: AMBERD } },
     "企業管理ポータル、実技予約・出欠管理、一括請求、CSV出力、記録の自動保存"],
    [{ text: "教材制作（推奨B案の場合）", options: { bold: true, color: NAVY } },
     { text: "40〜80万円／科目", options: { align: "center", bold: true, color: AMBERD } },
     "外部講師の講師料、撮影、編集、テスト問題の作成、教材PDFの制作"],
    [{ text: "月額運用費", options: { bold: true, color: NAVY } },
     { text: "5〜10万円＋従量", options: { align: "center", bold: true, color: AMBERD } },
     "サーバー費、保守・障害対応、本人確認API（従量）、動画配信費（従量）"]
  ];
  tbl(s, rows, { y: 1.9, colW: [3.2, 2.3, 6.633], rowH: [0.42, 0.72, 0.72, 0.72, 0.72, 0.72],
    fill: { color: CARD }, fontSize: 11 });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 5.85, w: CW, h: 0.85, rectRadius: 0.05, fill: { color: AMBERL } });
  s.addText("※ 上記は同種システムの一般的な相場に基づく概算レンジです。機能の絞り込みによって下振れも可能ですし、AI顔認証をリアルタイム常時照合（厳格案B）に引き上げる場合は上振れします。正式なお見積りはフェーズ0の完了後に別途ご提示いたします。",
    { x: M + 0.3, y: 5.85, w: CW - 0.6, h: 0.85, fontSize: 11, color: AMBERD, fontFace: JP,
      isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.25 });
  s.addNotes("金額は概算です。まずフェーズ0で要件を固めてから、正式なお見積りをお出しします。");
}

// =====================================================================
// 18. Risks
// =====================================================================
{
  const s = contentSlide("RISK", "想定されるリスクと対策", "外部への有料提供だからこそ、先に手を打っておくべき点です。");
  const risks = [
    ["個人情報の取扱い", "本人確認書類と顔写真は、外部の受講者からお預かりする個人情報です。",
     "暗号化して保存し、アクセス権限を限定。保存期間のポリシーを定め、プライバシーポリシーと利用規約をフェーズ0で整備します。"],
    ["講師要件の説明責任", "「十分な知識と経験を有する者」の説明を求められる可能性があります。",
     "講師の経歴書を整備し、教材内と修了証に講師名を明示。B案で外部の有資格者を起用すれば、説明はさらに容易になります。"],
    ["実技会場のキャパシティ", "半田会場の定員が、そのままプランBの売上上限になります。",
     "予約制と定員管理で平準化し、繁忙期は追加日程を設定。プランA（学科のみ）を併売して会場に依存しない売上を作ります。"],
    ["競合との差別化", "大手の講習機関がすでに全国でオンライン提供を行っています。",
     "価格競争を避け、「本人確認の厳格さ」「学科から実技までワンストップ」「東海地区の地域密着」の3点で勝負します。"]
  ];
  const cw = (CW - 0.3) / 2, ch = 2.2;
  risks.forEach((r, i) => {
    const x = M + (i % 2) * (cw + 0.3);
    const y = 1.85 + Math.floor(i / 2) * (ch + 0.28);
    card(s, x, y, cw, ch);
    circle(s, x + 0.28, y + 0.26, 0.5, String(i + 1), RED, "FFFFFF", 13);
    s.addText(r[0], { x: x + 0.92, y: y + 0.24, w: cw - 1.2, h: 0.36, fontSize: 14.5, bold: true,
      color: NAVY, fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(r[1], { x: x + 0.3, y: y + 0.7, w: cw - 0.6, h: 0.5, fontSize: 11, color: RED,
      fontFace: JP, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.3, y: y + 1.22, w: cw - 0.6, h: 0.82,
      rectRadius: 0.04, fill: { color: LIGHT } });
    s.addText([{ text: "対策　", options: { bold: true, color: GREEN } }, { text: r[2], options: { color: INK } }],
      { x: x + 0.45, y: y + 1.22, w: cw - 0.9, h: 0.82, fontSize: 10.5, fontFace: JP,
        isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.2 });
  });
  s.addNotes("特に1番目の個人情報は、社内利用と外部提供では求められる水準が違います。ここは最初に固めておくべき部分です。");
}

// =====================================================================
// 19. Subsidies
// =====================================================================
{
  const s = contentSlide("FUNDING", "補助金の活用可能性", "予算が未定とのことでしたので、初期投資を軽くする選択肢をご提示します。");
  const subs = [
    ["新事業進出・ものづくり\n商業サービス補助金", "◎ 最も親和性が高い", GREEN,
     "新製品・新サービスの開発が対象で、システム構築費が補助対象になります。補助率は40〜60%程度。今回の「講習事業という新規事業の立ち上げ」と最も相性がよい制度です。"],
    ["小規模事業者\n持続化補助金", "○ 併用を検討", STEEL,
     "販路開拓が対象。上限は50万円程度と小さいものの、申請手続きが簡素です。サービスサイトの制作や広告出稿に充てやすい制度です。"],
    ["デジタル化・AI導入補助金\n（旧 IT導入補助金）", "△ 注意が必要", RED,
     "2026年度に名称変更されました。自社利用のITツール導入が主な対象のため、外部へ販売するシステムの開発費には原則として使えません。"]
  ];
  const cw = (CW - 0.5) / 3;
  subs.forEach((sb, i) => {
    const x = M + i * (cw + 0.25);
    card(s, x, 1.85, cw, 3.5);
    s.addText(sb[0], { x: x + 0.28, y: 2.05, w: cw - 0.56, h: 0.85, fontSize: 14.5, bold: true,
      color: NAVY, fontFace: JP, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.28, y: 2.98, w: cw - 0.56, h: 0.4, rectRadius: 0.05,
      fill: { color: sb[2] } });
    s.addText(sb[1], { x: x + 0.28, y: 2.98, w: cw - 0.56, h: 0.4, align: "center", valign: "middle",
      fontSize: 11.5, bold: true, color: "FFFFFF", fontFace: JP, isTextBox: true, margin: 0 });
    s.addText(sb[3], { x: x + 0.28, y: 3.55, w: cw - 0.56, h: 1.65, fontSize: 11, color: INK,
      fontFace: JP, isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.3 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 5.55, w: CW, h: 1.0, rectRadius: 0.05, fill: { color: NAVY } });
  circle(s, M + 0.3, 5.79, 0.52, "!", AMBER, DARK, 18);
  s.addText("補助金の制度内容・補助率・公募回次は毎年変更されます。本ページは2026年8月時点で公表されている情報に基づく参考情報です。実際に申請される場合は、認定支援機関を交えて個別に要件をご確認ください。",
    { x: M + 1.02, y: 5.55, w: CW - 1.4, h: 1.0, fontSize: 11, color: "D6E2EC", fontFace: JP,
      isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.25 });
  s.addNotes("補助金ありきで事業を設計するのは危険ですが、フェーズ1の初期投資を軽くする手段としては有効です。");
}

// =====================================================================
// 20. Next actions
// =====================================================================
{
  const s = contentSlide("NEXT", "次のアクションと、当日ご相談したい論点", "「当日ご質問したい」とのことでしたので、想定される論点を先にお出しします。");
  const lw = (CW - 0.3) / 2;
  card(s, M, 1.85, lw, 4.55);
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 1.85, w: lw, h: 0.68, rectRadius: 0.05, fill: { color: STEEL } });
  s.addText("弊社が次に行うこと", { x: M + 0.3, y: 1.85, w: lw - 0.6, h: 0.68, fontSize: 15, bold: true,
    color: "FFFFFF", fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  const ours = [
    ["本提案のご説明とご質問への回答", "本日いただくご質問に、その場でお答えします"],
    ["第1弾科目の選定支援", "自社に必要な科目と、外販で売れる科目の両面から整理します"],
    ["フェーズ0（要件定義）のご提案", "業務フローと画面を確定し、正式なお見積りをご提示します"],
    ["補助金活用の可否確認", "認定支援機関と連携し、申請できるかを確認します"]
  ];
  ours.forEach((o, i) => {
    const y = 2.75 + i * 0.92;
    circle(s, M + 0.3, y + 0.06, 0.46, String(i + 1), STEEL, "FFFFFF", 12);
    s.addText(o[0], { x: M + 0.9, y: y, w: lw - 1.2, h: 0.34, fontSize: 12.5, bold: true, color: NAVY,
      fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(o[1], { x: M + 0.9, y: y + 0.34, w: lw - 1.2, h: 0.4, fontSize: 10.5, color: MUTED,
      fontFace: JP, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  });

  const rx = M + lw + 0.3;
  card(s, rx, 1.85, lw, 4.55);
  s.addShape(pres.ShapeType.roundRect, { x: rx, y: 1.85, w: lw, h: 0.68, rectRadius: 0.05, fill: { color: AMBERD } });
  s.addText("当日ご相談したい論点", { x: rx + 0.3, y: 1.85, w: lw - 0.6, h: 0.68, fontSize: 15, bold: true,
    color: "FFFFFF", fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  const theirs = [
    "第1弾として提供する科目（推奨：フルハーネス型墜落制止用器具、またはテールゲートリフター）",
    "実技会場となる半田市の工場のキャパシティと、実技指導者の確保",
    "既存の取引先・協力会社への先行案内が可能か（初期の受講者確保の見通し）",
    "サービス名・ブランドの方向性（聖建様の社名を前面に出すか）",
    "受講記録の保存と日々の運用を、社内のどなたが担当されるか"
  ];
  theirs.forEach((t, i) => {
    const y = 2.72 + i * 0.72;
    s.addShape(pres.ShapeType.roundRect, { x: rx + 0.26, y: y, w: lw - 0.52, h: 0.62,
      rectRadius: 0.04, fill: { color: LIGHT } });
    s.addShape(pres.ShapeType.ellipse, { x: rx + 0.42, y: y + 0.245, w: 0.13, h: 0.13, fill: { color: AMBER } });
    s.addText(t, { x: rx + 0.68, y: y, w: lw - 1.0, h: 0.62, fontSize: 10.5, color: INK, fontFace: JP,
      isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.15 });
  });
  s.addNotes("当日は、右側の5点を中心に議論できればと思います。特に1番目の科目選定と2番目の実技会場は、事業の形を決める部分です。");
}

// =====================================================================
// 21. Closing
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };
  s.addShape(pres.ShapeType.ellipse, { x: -2.9, y: 5.3, w: 4.0, h: 4.0, fill: { color: NAVY } });
  s.addShape(pres.ShapeType.ellipse, { x: 11.6, y: -1.4, w: 3.4, h: 3.4, fill: { color: NAVY } });
  s.addText("特別教育は、行政の登録がなくても始められます。", { x: 1.5, y: 1.95, w: 10.5, h: 0.62,
    fontSize: 25, bold: true, color: "FFFFFF", fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("だからこそ、「きちんと受講したことを証明できる」ことが、\nそのまま商品になります。",
    { x: 1.5, y: 2.68, w: 10.5, h: 1.3, fontSize: 25, bold: true, color: AMBER, fontFace: JP,
      isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.25 });
  s.addText("聖建様の現場の知見を、東海地区の建設業全体へ届ける仕組みをご一緒に作らせてください。",
    { x: 1.5, y: 4.15, w: 10.5, h: 0.4, fontSize: 14, color: "C3D2DF", fontFace: JP, isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.roundRect, { x: 1.5, y: 5.0, w: 4.6, h: 1.15, rectRadius: 0.05, fill: { color: NAVY } });
  s.addText("株式会社スペリオル", { x: 1.8, y: 5.16, w: 4.0, h: 0.38, fontSize: 15, bold: true,
    color: "FFFFFF", fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addText("担当：藤原", { x: 1.8, y: 5.56, w: 4.0, h: 0.34, fontSize: 12, color: "AFC3D4",
    fontFace: JP, isTextBox: true, margin: 0, valign: "middle" });
  s.addNotes("本日はありがとうございました。ご質問をお受けします。");
}

pres.writeFile({ fileName: "/home/user/-/deck/聖建様_特別教育オンラインプラットフォーム_ご提案.pptx" })
  .then(f => console.log("WROTE:", f));
