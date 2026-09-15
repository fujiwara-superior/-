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


// 条文
const ART = (no, title) => new Paragraph({
  spacing: { before: 300, after: 110 },
  children: [
    new TextRun({ text: `第${no}条　`, font: FONT, size: 22, bold: true, color: ACCENT }),
    new TextRun({ text: `（${title}）`, font: FONT, size: 22, bold: true, color: NAVY }),
  ],
});
// 条文の項
const CL = (n, text) => new Paragraph({
  spacing: { after: 90, line: 300 },
  indent: { left: 340, hanging: 280 },
  children: [new TextRun({ text: `${n}　`, font: FONT, size: 20, color: INK2 })].concat(
    (Array.isArray(text) ? text : [text]).map(t => typeof t === 'string' ? T(t, { size: 20 }) : t)),
});

const body = [];

body.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 400 },
  children: [T('ご契約条件に関する申し入れ', { size: 32, bold: true })],
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

body.push(P('お電話にてご相談いただきました件につきまして、弊社の考えを申し述べます。', { before: 260, after: 60 }));
body.push(P('ご要望の趣旨は、前払いいただく代金が、成果物の伴わないまま失われることのないようにしたい、というものと理解しております。その点につきましては弊社も同じ考えであり、より確実な方法をご提案申し上げます。'));

// ---------- 1 ----------
body.push(H(1, '個人保証について'));
body.push(P(['誠に恐れ入りますが、', T('弊社役員個人による保証はご容赦いただきたく存じます。', { bold: true })]));
body.push(P('本件は法人として受注するものであり、その履行責任は法人が負うべきものと考えております。また個人保証は本来、金銭の貸借において用いられる手段であり、請負契約にはなじまないものと存じます。'));
body.push(NOTE('保証よりも確実な方法をご提案いたします', [
  ['個人保証は、', T('万一のときに回収できるかもしれない', { bold: true }),
   'という手段にすぎず、貴社が本当に必要とされている「システムが完成すること」そのものを保証するものではありません。'],
  ['弊社からは、', T('何をもって完成とするかを契約前に確定し、その達成を客観的に検証できる仕組み', { bold: true }),
   'をご提案いたします。次項以降のとおりです。これは保証よりも、貴社にとって実効性のある担保になると考えております。'],
]));

// ---------- 2 ----------
body.push(H(2, '「完成」の定義を、契約前に確定します'));
body.push(P(['争いが生じるとすれば、その原因は', T('「完成したか否かの判断が分かれること」', { bold: true }),
  'にほかなりません。この点を曖昧にしたまま契約を結ぶことが、双方にとって最大のリスクです。']));
body.push(P(['そこで弊社は、要件定義の完了時に', T('検収テスト仕様書', { bold: true }),
  'を作成し、貴社のご承認をいただきます。']));
body.push(KV([
  ['検収テスト仕様書とは', [T('本件システムが備えるべき機能・動作を1件ずつ列挙し、それぞれについて「どのような操作をすれば、どのような結果になれば合格か」を具体的に記した一覧表です', { size: 19 })]],
  ['判定の方法', [T('各項目を実際に操作し、合格・不合格を判定します。全項目が合格した状態をもって「完成」とします', { size: 19 })]],
  ['意見の余地', [T('主観が入る余地をなくすため、「使いやすいこと」といった判断の分かれる表現は用いません', { size: 19 })]],
  ['記載外の事項', [T('この一覧表に記載のない事項は、完成の判断には影響しないものといたします', { size: 19 })]],
]));
body.push(P('この一覧表は、要件定義の成果物として貴社にご確認いただき、書面でご承認をいただいたうえで確定させます。以後、双方はこの一覧表を基準として完成の有無を判断することになります。', { before: 80 }));

// ---------- 3 ----------
body.push(H(3, '完成の判断が分かれた場合は、第三者による検証に委ねます'));
body.push(P(['それでもなお見解が一致しない場合に備え、', T('甲乙いずれからも独立した第三者による検証', { bold: true }),
  'を求めることができる旨を契約に定めることをご提案いたします。']));
body.push(P('検証は、上記の検収テスト仕様書の各項目について合否を判定する方法により行います。検証者が独自の基準を持ち込むことはなく、双方が事前に合意した基準のみで判断されます。'));
body.push(NOTE('検証費用の負担', [
  ['検証に要する費用は、', T('検証の結果、仕様を満たしていると判定された場合は貴社のご負担、満たしていないと判定された場合は弊社の負担', { bold: true }),
   'とすることをご提案いたします。'],
  ['判定を誤った側が費用を負うこととすることで、双方とも軽々に検証を求めることがなくなり、まずは話し合いで解決しようという動機が働きます。'],
], 'DDEDE5', '26694F'));

// ---------- 4 ----------
body.push(H(4, '返金について'));
body.push(P('返金に関しましては、次の2点を契約に定めていただきたく存じます。いずれも民法の定めに沿った順序です。'));
body.push(P(T('(1) 返金の前に、まず修補の機会をいただきたい', { bold: true, size: 22, color: NAVY }), { before: 200, after: 100 }));
body.push(P('仕様に適合しない点が判明した場合、まず相当の期間を定めて弊社に修補をご請求ください。弊社がその期間内に修補を完了しないときに限り、代金の減額または契約の解除をお求めいただく、という順序といたします。'));
body.push(P(T('(2) 返金の範囲は、適合しない部分に対応する金額に限る', { bold: true, size: 22, color: NAVY }), { before: 200, after: 100 }));
body.push(P('本件は、eラーニング教材6科目と配信システムからなり、それぞれ個別に検収が可能です。一部に不適合があった場合に代金全額の返還を求められることのないよう、返金の範囲は、適合しないと判定された部分に対応する金額に限るものといたします。'));

// ---------- 別紙 ----------
body.push(new Paragraph({ spacing: { before: 400, after: 200 },
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: 'D3DDDD', space: 10 } },
  children: [] }));
body.push(P(T('別紙　契約書にご反映いただきたい条項案', { size: 26, bold: true, color: NAVY }), { after: 140 }));
body.push(P('貴社にて契約書をご作成いただくにあたり、参考となるよう条文の案を記します。文言はご自由にご調整ください。', { after: 200 }));

body.push(ART('○', '仕事の完成'));
body.push(CL('1', '本契約における「仕事の完成」とは、要件定義の完了時に甲乙が書面により承認した検収テスト仕様書に定めるすべての項目について、合格の判定が得られた状態をいう。'));
body.push(CL('2', '検収テスト仕様書は、要件定義書に定める機能および性能に基づき乙が作成し、甲の承認を得るものとする。'));
body.push(CL('3', '検収テスト仕様書に記載のない事項は、本契約における完成の判断に影響しないものとする。'));

body.push(ART('○', '検収'));
body.push(CL('1', '本件業務の成果物は、eラーニング教材および配信システムからなり、それぞれ個別に検収するものとする。'));
body.push(CL('2', '乙は、成果物の完成後、甲に対し検収を依頼する。'));
body.push(CL('3', '甲は、検収依頼の日から30日以内に、検収テスト仕様書に基づく確認を行い、その合否を書面により乙に通知する。'));
body.push(CL('4', '甲が前項の期間内に通知を行わないときは、当該成果物は検収に合格したものとみなす。'));
body.push(CL('5', '甲は、検収テスト仕様書に定める項目以外の事由をもって検収を拒むことができない。'));

body.push(ART('○', '第三者による検証'));
body.push(CL('1', '検収の合否について甲乙の見解が一致しないときは、甲乙いずれの当事者も、第三者による検証を求めることができる。'));
body.push(CL('2', '検証者は、システム開発の検収実務に知見を有し、かつ甲乙いずれからも独立した者とし、甲乙協議のうえ選任する。'));
body.push(CL('3', '検証は、検収テスト仕様書の各項目について合否を判定する方法により行うものとし、甲乙は当該判定に従う。'));
body.push(CL('4', '検証に要する費用は、検証の結果、成果物が検収テスト仕様書に適合すると判定された場合は甲の負担とし、適合しないと判定された場合は乙の負担とする。'));

body.push(ART('○', '契約不適合があった場合の措置'));
body.push(CL('1', '検収後に成果物が検収テスト仕様書に適合しないことが判明した場合、甲は、まず乙に対し、相当の期間を定めて修補を請求するものとする。'));
body.push(CL('2', '乙が前項の期間内に修補を完了しないときに限り、甲は代金の減額または本契約の解除を求めることができる。'));
body.push(CL('3', '前項により代金の返還が生じる場合、その範囲は、検収テスト仕様書に適合しないと判定された部分に対応する金額に限るものとする。'));
body.push(CL('4', '本条に基づく甲の請求は、検収完了の日から6か月以内に書面により行われたものに限る。'));

body.push(ART('○', '責任の主体'));
body.push(CL('1', '本契約に基づく乙の債務については、乙がその責任を負うものとし、乙の役員または従業員による個人保証は付さないものとする。'));

body.push(new Paragraph({ spacing: { before: 400, after: 100 }, border: {
  top: { style: BorderStyle.SINGLE, size: 6, color: 'D3DDDD', space: 8 } }, children: [] }));
body.push(P([T('本書について　', { size: 18, bold: true, color: INK2 }),
  T('本書は弊社の考えをお伝えするものであり、契約書に代わるものではありません。最終的な権利関係および責任範囲は、双方が署名する契約書の定めによるものといたします。', { size: 18, color: INK2 })]));
body.push(P(T('株式会社スペリオル　〒464-0075 名古屋市千種区内山3-18-10　TEL 052-745-6607　担当 藤原', { size: 18, color: INK2 })));

const doc = new Document({
  creator: '株式会社スペリオル',
  title: 'ご契約条件に関する申し入れ',
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
  fs.writeFileSync('/home/user/-/contract/聖建様_ご契約条件に関する申し入れ.docx', b);
  console.log('saved');
});
