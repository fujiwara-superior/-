// 聖建様 確認事項へのご回答（Word形式）
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  LevelFormat, convertInchesToTwip,
} = require('docx');

const FONT = 'Meiryo';
const NAVY = '12444B';
const ACCENT = 'BC4517';
const INK2 = '485C63';
const W = 9000;          // 表の全幅（DXA）

const P = (text, o = {}) => new Paragraph({
  alignment: o.align,
  spacing: { before: o.before ?? 0, after: o.after ?? 120, line: 300 },
  indent: o.indent,
  border: o.border,
  shading: o.shading,
  children: (Array.isArray(text) ? text : [text]).map(t =>
    typeof t === 'string'
      ? new TextRun({ text: t, font: FONT, size: o.size ?? 21, bold: o.bold, color: o.color })
      : t),
});
const T = (text, o = {}) => new TextRun({
  text, font: FONT, size: o.size ?? 21, bold: o.bold, color: o.color, italics: o.italics,
});

// 見出し（第n項）
const H = (no, title) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 360, after: 140 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 4 } },
  children: [
    new TextRun({ text: `${no}. `, font: FONT, size: 22, bold: true, color: ACCENT }),
    new TextRun({ text: title, font: FONT, size: 22, bold: true, color: NAVY }),
  ],
});

// 先方の確認事項（引用）
const Q = (text) => new Paragraph({
  spacing: { before: 60, after: 160, line: 280 },
  indent: { left: 340, right: 340 },
  shading: { type: ShadingType.CLEAR, fill: 'F1F5F5' },
  border: { left: { style: BorderStyle.SINGLE, size: 12, color: 'B6C4C4', space: 8 } },
  children: [
    new TextRun({ text: 'ご確認事項　', font: FONT, size: 17, bold: true, color: INK2 }),
    new TextRun({ text, font: FONT, size: 18, color: INK2 }),
  ],
});

// 箇条書き
const LI = (text, o = {}) => new Paragraph({
  numbering: { reference: 'dot', level: 0 },
  spacing: { after: 70, line: 300 },
  children: (Array.isArray(text) ? text : [text]).map(t =>
    typeof t === 'string' ? T(t, o) : t),
});

// 枠付きの注記
const NOTE = (title, paras, fill = 'F6E4DB', color = ACCENT) => {
  const kids = [new Paragraph({
    spacing: { after: 90, line: 290 },
    children: [new TextRun({ text: title, font: FONT, size: 21, bold: true, color })],
  })];
  paras.forEach((ps, i) => kids.push(new Paragraph({
    spacing: { after: i === paras.length - 1 ? 0 : 110, line: 300 },
    children: (Array.isArray(ps) ? ps : [ps]).map(t => typeof t === 'string' ? T(t) : t),
  })));
  return new Table({
    columnWidths: [W],
    width: { size: W, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color },
      bottom: { style: BorderStyle.SINGLE, size: 6, color },
      left: { style: BorderStyle.SINGLE, size: 6, color },
      right: { style: BorderStyle.SINGLE, size: 6, color },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: [new TableRow({ children: [new TableCell({
      width: { size: W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill },
      margins: { top: 140, bottom: 140, left: 200, right: 200 },
      children: kids,
    })] })],
  });
};

// 2列の表
const KV = (rows, w1 = 2300) => new Table({
  columnWidths: [w1, W - w1],
  width: { size: W, type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 4, color: 'D3DDDD' },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: 'D3DDDD' },
    left: { style: BorderStyle.SINGLE, size: 4, color: 'D3DDDD' },
    right: { style: BorderStyle.SINGLE, size: 4, color: 'D3DDDD' },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: 'D3DDDD' },
    insideVertical: { style: BorderStyle.SINGLE, size: 4, color: 'D3DDDD' },
  },
  rows: rows.map(([k, v]) => new TableRow({ children: [
    new TableCell({
      width: { size: w1, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: 'F1F5F5' },
      margins: { top: 90, bottom: 90, left: 140, right: 140 },
      children: [new Paragraph({ spacing: { after: 0, line: 280 },
        children: [T(k, { size: 19, bold: true })] })],
    }),
    new TableCell({
      width: { size: W - w1, type: WidthType.DXA },
      margins: { top: 90, bottom: 90, left: 140, right: 140 },
      children: [new Paragraph({ spacing: { after: 0, line: 280 },
        children: (Array.isArray(v) ? v : [v]).map(t => typeof t === 'string' ? T(t, { size: 19 }) : t) })],
    }),
  ] })),
});


// 金額表（3列）
const TBL3 = (rows, w = [3400, 2200, 3400]) => new Table({
  columnWidths: w,
  width: { size: W, type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 4, color: 'D3DDDD' },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: 'D3DDDD' },
    left: { style: BorderStyle.SINGLE, size: 4, color: 'D3DDDD' },
    right: { style: BorderStyle.SINGLE, size: 4, color: 'D3DDDD' },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: 'D3DDDD' },
    insideVertical: { style: BorderStyle.SINGLE, size: 4, color: 'D3DDDD' },
  },
  rows: rows.map(r => new TableRow({ children: r.cells.map((c, i) => new TableCell({
    width: { size: w[i], type: WidthType.DXA },
    shading: r.head ? { type: ShadingType.CLEAR, fill: 'F1F5F5' } : undefined,
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    children: [new Paragraph({
      alignment: i === 0 ? undefined : AlignmentType.RIGHT,
      spacing: { after: 0, line: 270 },
      children: [T(c, { size: 19, bold: r.head || r.bold })],
    })],
  })) })),
});

const body = [];

body.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 400 },
  children: [T('お見積金額に関するご説明', { size: 32, bold: true })],
}));

body.push(new Table({
  columnWidths: [5500, W - 5500],
  width: { size: W, type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
    left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
    insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
  },
  rows: [new TableRow({ children: [
    new TableCell({
      width: { size: 5500, type: WidthType.DXA },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children: [
        P(T('株式会社聖建　御中', { size: 26, bold: true }), { after: 60 }),
        P(T('津隈 佑己 様', { size: 20, color: INK2 }), { after: 0 }),
      ],
    }),
    new TableCell({
      width: { size: W - 5500, type: WidthType.DXA },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children: [
        P(T('2026年9月15日', { size: 19, color: INK2 }), { align: AlignmentType.RIGHT, after: 60 }),
        P(T('株式会社スペリオル', { size: 23, bold: true }), { align: AlignmentType.RIGHT, after: 40 }),
        P(T('〒464-0075', { size: 18, color: INK2 }), { align: AlignmentType.RIGHT, after: 10 }),
        P(T('名古屋市千種区内山3-18-10', { size: 18, color: INK2 }), { align: AlignmentType.RIGHT, after: 20 }),
        P(T('TEL 052-745-6607　／　担当　藤原', { size: 18, color: INK2 }), { align: AlignmentType.RIGHT, after: 0 }),
      ],
    }),
  ] })],
}));
body.push(new Paragraph({ spacing: { after: 60 }, border: {
  bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 6 } }, children: [] }));

body.push(P('お見積金額につきましてお尋ねをいただきましたので、弊社の考え方をご説明申し上げます。', { before: 260, after: 200 }));

// ---------- 1 ----------
body.push(H(1, 'お見積金額の積算について'));
body.push(P('今回のお見積金額 7,700,000円（税別）は、下記の積算に基づくものです。'));
body.push(TBL3([
  { head: true, cells: ['区分', '数量', '金額（税別）'] },
  { cells: ['システム開発', '125人日', '6,250,000円'] },
  { cells: ['eラーニング教材制作', '6科目', '1,800,000円'] },
  { bold: true, cells: ['小計', '', '8,050,000円'] },
  { cells: ['貴社特別値引き', '', '▲350,000円'] },
  { bold: true, cells: ['差引（税別）', '', '7,700,000円'] },
]));
body.push(P(['前回のお見積り 6,600,000円から増額となりましたのは、',
  T('在席確認（カメラ前からの離席検知）および受講中のランダム顔照合を新たに範囲に加えたため', { bold: true }),
  'です。内訳は次のとおりです。'], { before: 160 }));
body.push(TBL3([
  { head: true, cells: ['追加・除外した項目', '人日', '金額（税別）'] },
  { cells: ['在席確認（カメラ前からの離席検知）', '＋18人日', '＋900,000円'] },
  { cells: ['受講中のランダム顔照合', '＋4人日', '＋200,000円'] },
  { cells: ['上記にともなう既存機能の改修', '＋10人日', '＋500,000円'] },
  { cells: ['決済機能（範囲外といたしました）', '▲10人日', '▲500,000円'] },
  { bold: true, cells: ['差引', '＋22人日', '＋1,100,000円'] },
]));
body.push(P('お見積書および見積明細にも同じ内訳を記載しております。金額は積算の結果であり、いずれの項目も実際に必要となる作業に対応するものです。', { before: 160 }));

// ---------- 2 ----------
body.push(H(2, 'ご契約およびご請求の方法について'));
body.push(P('弊社は、お見積書の内訳に基づいて契約書および請求書を作成しております。'));
body.push(P('つきましては、誠に恐れ入りますが、次の点につきましてはお受けいたしかねます。'));
body.push(LI(T('お見積書の内訳と異なる金額でのご契約またはご請求', { bold: true })));
body.push(LI(T('ご契約の当事者以外の方へのお支払い、および名目のいかんを問わない金銭のお渡し', { bold: true })));
body.push(NOTE('本件に限ったものではございません', [
  ['上記は弊社が', T('すべてのお取引において同様に取り扱っている', { bold: true }),
   'ものであり、本件について特別に申し上げているものではございません。'],
  ['弊社は帳簿と実際の取引を一致させる必要があり、これを欠きますと、税務上の取扱いにおいて弊社および貴社の双方にご迷惑をおかけすることとなります。何卒ご理解を賜りますようお願い申し上げます。'],
], 'F1F5F5', NAVY));

// ---------- 3 ----------
body.push(H(3, 'ご予算に合わせたお見積りの出し直しについて'));
body.push(P(['ご予算が ', T('6,600,000円', { bold: true }),
  ' とのことでございましたら、その範囲に収まるよう、機能を絞り込んだ内容で改めてお見積りをご提出いたします。']));
body.push(P('今回の増額分は在席確認に関する部分ですので、こちらを範囲から外す形でのご提案が可能です。その場合、月額の保守・運用サービスもあわせて見直しとなります。'));
body.push(NOTE('ご判断をお聞かせください', [
  [T('(1) 在席確認を含めたまま進める場合', { bold: true }), '　… 7,700,000円（税別）。現在のお見積りのとおりです。'],
  [T('(2) ご予算 6,600,000円 に合わせる場合', { bold: true }), '　… 在席確認および受講中のランダム顔照合を範囲から外し、再見積をご提出いたします。'],
  ['いずれをご希望か、お聞かせいただけますでしょうか。ご連絡をいただき次第、速やかに書類をご用意いたします。'],
], 'DDEDE5', '26694F'));

body.push(P('本件は貴社の事業の立ち上げに関わる重要なご契約と存じております。良い形でお受けできますよう努めてまいりますので、何卒よろしくお願い申し上げます。', { before: 240 }));

body.push(new Paragraph({ spacing: { before: 400, after: 100 }, border: {
  top: { style: BorderStyle.SINGLE, size: 6, color: 'D3DDDD', space: 8 } }, children: [] }));
body.push(P(T('株式会社スペリオル　〒464-0075 名古屋市千種区内山3-18-10　TEL 052-745-6607　担当 藤原', { size: 18, color: INK2 })));

const doc = new Document({
  creator: '株式会社スペリオル',
  title: 'お見積金額に関するご説明',
  numbering: { config: [{
    reference: 'dot',
    levels: [{
      level: 0, format: LevelFormat.BULLET, text: '・', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 340, hanging: 220 } }, run: { font: FONT, size: 21 } },
    }],
  }] },
  sections: [{
    properties: { page: { margin: {
      top: convertInchesToTwip(0.85), bottom: convertInchesToTwip(0.85),
      left: convertInchesToTwip(0.9), right: convertInchesToTwip(0.9),
    } } },
    children: body,
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('/home/user/-/contract/聖建様_お見積金額に関するご説明.docx', b);
  console.log('saved');
});
