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

const body = [];

// ---------- 標題 ----------
body.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 400 },
  children: [T('確認事項へのご回答', { size: 32, bold: true })],
}));

// ---------- 宛先・発行者 ----------
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

// ---------- 前文 ----------
body.push(P('平素より格別のお引き立てを賜り、厚く御礼申し上げます。', { before: 260, after: 60 }));
body.push(P([
  'いただきました確認事項につきまして、下記のとおりご回答申し上げます。',
  T('5番および6番につきましては、弊社の事業継続に関わる事項のため、考え方をあわせてご説明いたします。', { bold: true }),
  'ご不明な点がございましたら、契約書のご署名前にお打ち合わせのお時間を頂戴できますと幸いです。',
], { after: 200 }));

// ---------- 1 ----------
body.push(H(1, '受講履歴のCSV・PDFダウンロード'));
body.push(Q('対象の受講生のログイン時間、視聴時間、ログアウト時間等の受講履歴詳細をCSVやPDFでダウンロードできる機能は搭載されますでしょうか'));
body.push(P('搭載いたします。運営管理画面から、受講者単位・期間単位でCSV出力できるようにします。出力項目は次のとおりです。'));
body.push(LI('受講者氏名、科目、受講開始日時、修了日時'));
body.push(LI('ログイン日時、ログアウト日時、セッションごとの再生時間'));
body.push(LI([T('有効視聴時間', { bold: true }), '（早送り・離席・タブ非表示を除外した、実際に視聴された時間）']));
body.push(LI('離席の回数と合計時間'));
body.push(LI('顔照合の実施回数と一致回数'));
body.push(LI('確認テストの得点、受験回数'));
body.push(P('PDFは修了証を出力します。受講履歴そのものは、集計・加工のしやすさからCSVを標準といたします。', { before: 80 }));
body.push(P(T('※ 受講者の所属企業ごとに、その企業の担当者が自社分だけを一括出力する機能（企業管理ポータル）は今回の範囲外です。貴社の運営側から全件を出力する機能は含まれます。', { size: 18, color: INK2 })));

// ---------- 2 ----------
body.push(H(2, 'テスト期間'));
body.push(Q('テスト期間を3週間〜1ヶ月設けていただけることは可能でしょうか'));
body.push(P('承知いたしました。3週間から1ヶ月のテスト期間を設けます。'));
body.push(P([ 'あわせて、この期間を', T('検収期間', { bold: true }),
  'と位置づけさせてください。期間内に書面でご指摘いただいた仕様との不適合は、無償で修正いたします。期間経過後に特段のご連絡がない場合は、検収完了とみなさせていただきます。' ]));
body.push(P('なお、テスト期間中に生じた仕様変更・機能追加のご要望は、別途お見積りのうえ対応いたします（7番をご参照ください）。'));

// ---------- 3 ----------
body.push(H(3, 'ドメインのご契約'));
body.push(Q('サーバーの契約は他社様と直接とのことでしたが、ドメインの契約も弊社が直接させていただけますでしょうか'));
body.push(P(['はい、', T('貴社で直接ご契約ください。', { bold: true }),
  'そのほうが望ましいと考えます。ドメインは貴社の資産であり、弊社名義で保有すべきものではありません。']));
body.push(P('DNSの設定およびSSL証明書の設定は弊社が行います。設定に必要な権限のみ、作業時に一時的にお預かりいたします。'));

// ---------- 4 ----------
body.push(H(4, '仕様を満たしていない場合の無償修正'));
body.push(Q('本サイトを構築するにあたり、仕様を満たしていない場合は無償での修正対応をお願いいたします。'));
body.push(P('承諾いたします。'));
body.push(P('あわせて、以下の2点を契約書に明記させてください。認識の相違を防ぐためのものです。'));
body.push(LI([T('「仕様」とは、要件定義の完了時に双方が書面で確定した仕様書を指します。', { bold: true }),
  '仕様書に記載のない事項は「仕様を満たしていない」には該当せず、仕様変更・機能追加として別途お見積りの対象となります。']));
body.push(LI(['契約不適合責任の期間は、', T('検収完了後6ヶ月間', { bold: true }), 'とさせてください。']));

// ---------- 5 ----------
body.push(H(5, '所有権および著作権の移転タイミング'));
body.push(Q('所有権及び著作権の移転タイミングをご教授願います'));
body.push(P('はじめに、法律上の前提を1点だけ共有させてください。'));
body.push(NOTE('プログラムの著作権は、制作した者に発生します', [
  ['請負契約で制作した場合も同様です。',
   T('契約で移転を定めないかぎり、著作権は制作した弊社に留まります', { bold: true }),
   '（著作権法第17条、第61条）。したがって「いつ移転するか」は、契約でどう定めるかの問題となります。そのうえで、弊社からは次の切り分けをご提案いたします。'],
], 'F1F5F5', NAVY));
body.push(P(T('(1) 貴社へ譲渡するもの — 本件のために新たに制作した部分', { bold: true, size: 22, color: NAVY }), { before: 240, after: 100 }));
body.push(LI('eラーニング教材（動画、構成台本、スライド、図解、確認問題）'));
body.push(LI('画面デザイン'));
body.push(LI('貴社の業務に固有の処理（受講管理、修了判定、修了証の様式 など）'));
body.push(LI('データベース設計、および貴社が本システムに蓄積されたデータ'));
body.push(P(['これらの著作権（', T('著作権法第27条および第28条の権利を含みます', { bold: true }), '）は、',
  T('検収完了かつ代金完済の時点', { bold: true }),
  'で貴社へ移転いたします。弊社は、貴社に対し、著作者人格権を行使いたしません。'], { before: 80 }));
body.push(P(T('(2) 弊社に留保するもの — 汎用的な部分', { bold: true, size: 22, color: NAVY }), { before: 200, after: 100 }));
body.push(LI('本人確認、顔照合、在席確認、視聴整合性の判定など、弊社が従前から保有し、今後も他の案件で用いる共通モジュール'));
body.push(LI('開発基盤、共通ライブラリ、社内ツール'));
body.push(P(['これらにつきましては弊社に権利を留保させていただき、貴社に対し、',
  T('本システムの利用・運用・保守・改修に必要な範囲で、無償・永続・再許諾不可の利用許諾', { bold: true }),
  'を付与いたします。貴社が本システムをお使いになるうえで、実務上の制約は一切生じません。'], { before: 80 }));
body.push(NOTE('この切り分けをお願いする理由', [
  ['すべてを譲渡いたしますと、', T('弊社は同種のシステムを今後一切ご提供できなくなります。', { bold: true }),
   '長年かけて蓄積した汎用部品まで貴社の著作物となるためです。それでは弊社の事業が成り立たず、結果として本システムの保守を続けることもできなくなります。'],
  ['一方、貴社にとって重要なのは', T('「このシステムを制約なく使い続けられること」「弊社に万一のことがあっても事業が止まらないこと」', { bold: true }),
   'であって、弊社の汎用部品の権利を保有されることではないと存じます。その目的は、上記の利用許諾と、6番でご提案するエスクローによって過不足なく満たすことができます。'],
]));
body.push(P(T('※ 「所有権」につきまして、プログラムは無体物のため、厳密には所有権の対象となりません。納品する媒体および書類の所有権は引渡し時に移転いたします。実質的な論点は上記の著作権とご理解ください。', { size: 18, color: INK2 }), { before: 140 }));

// ---------- 6 ----------
body.push(H(6, '進捗状況の確認方法とソースコードの取扱い'));
body.push(Q('進捗状況確認のため、ギットハブ等で管理・運用させていただくことは可能でしょうか。社内規定上難しい場合は、最新のソースコード一式とデザインデータをzipファイルかギガファイルで毎週共有していただくことは可能でしょうか'));
body.push(P(['進捗を随時ご確認になりたいというご要望は、当然のことと受け止めております。一方で、',
  T('開発途中のソースコードをお渡しすることにつきましては、お受けいたしかねます。', { bold: true }), '理由は5番に記したとおりです。']));
body.push(P('ご要望の目的を「進捗の可視化」と「継続性の担保」の2つに分け、それぞれ別の手段でお応えいたします。'));

body.push(P(T('(1) 進捗の可視化', { bold: true, size: 22, color: NAVY }), { before: 200, after: 100 }));
body.push(KV([
  ['月2回の進捗報告書', ['月初と月中の2回、機能単位の進捗率、前回からの完了項目、次回までの予定、課題と対応方針を書面でご報告します']],
  ['フェーズごとのデモ', [T('検証環境で実際に操作していただきます。', { size: 19, bold: true }), T('要件定義、基本設計、各機能の実装、結合テストの区切りごとに実施します。画面と挙動を直接ご確認いただけますので、資料上の進捗率より確実です', { size: 19 })]],
  ['作業ボードの常時閲覧', [T('弊社が日々更新している作業一覧を、いつでもご覧いただけます（下記）', { size: 19 })]],
]));

body.push(NOTE('GitHub（ギットハブ）の閲覧権限について', [
  ['GitHubは、', T('開発の作業管理', { bold: true }), 'と', T('プログラムの保管', { bold: true }),
   'を、同じ場所で行うための仕組みです。このうち', T('作業管理の部分だけ', { bold: true }), 'を、貴社にご覧いただけるようにいたします。'],
  [T('■ ご覧いただけるもの — 作業一覧と進捗ボード', { bold: true })],
  ['「何を作るか」を1件ずつ登録した作業一覧と、それを「未着手／作業中／確認待ち／完了」の列に並べたボードです。付箋を貼り替えていくホワイトボードのような画面とお考えください。担当者も表示されます。'],
  ['弊社が日々の作業で更新しているものを、そのままご覧いただきます。', T('ご報告用に作り直したものではありませんので、常に最新の状況がそのまま見えます。', { bold: true })],
  [T('■ ご覧いただけないもの — プログラム本体', { bold: true })],
  ['ソースコードは同じGitHub内の別の区画に保管されており、こちらは非公開のままとさせてください。'],
  ['工事に例えますと、', T('工程表はいつでもご覧いただけますが、設計図面の原本はお渡ししない', { bold: true }),
   'という関係になります。閲覧用のアカウントは弊社でご用意し、貴社のご担当者様へお渡しいたします。無料でご利用いただけます。'],
], 'F1F5F5', NAVY));

body.push(P(T('(2) 継続性の担保 — ソースコードエスクロー', { bold: true, size: 22, color: NAVY }), { before: 200, after: 100 }));
body.push(NOTE('「スペリオルに何かあったらどうするのか」へのお答え', [
  ['ソースコードを', T('中立的な第三者機関に預託（エスクロー）', { bold: true }),
   'いたします。預託先は、ご契約後に貴社とご相談のうえ決定いたします。'],
  ['次の場合に、預託されたソースコードが貴社へ開示されるよう定めます。'],
  ['　・弊社が解散、破産、または事業を停止した場合'],
  ['　・弊社が保守契約上の義務を、正当な理由なく履行しない場合'],
  ['これにより、', T('平常時は弊社が管理し、万一のときは貴社が確実にコードを入手できる', { bold: true }),
   '状態になります。貴社が負われるリスクは、ソースコードを直接お持ちいただく場合と実質的に変わりません。'],
], 'DDEDE5', '26694F'));

body.push(P(T('(3) 納品するドキュメント', { bold: true, size: 22, color: NAVY }), { before: 200, after: 100 }));
body.push(P('検収時に、次の資料を納品いたします。他社様への引き継ぎが必要になった場合も、これらがあれば対応可能です。'));
body.push(LI('基本設計書、画面設計書'));
body.push(LI('データベース定義書、API仕様書'));
body.push(LI('運用手順書、環境構築手順書'));

// ---------- 7 ----------
body.push(H(7, '追加費用'));
body.push(Q('今回の初期費用(税別770万円)に関しまして、デザイン・開発・テスト・公開・修正 全て込みで追加費用は一切発生しない認識となります'));
body.push(P(T('確定した仕様の範囲内でのデザイン・開発・テスト・公開・修正につきましては、初期費用 7,700,000円（税別）以外の費用は発生いたしません。', { bold: true })));
body.push(P('一方、次の場合は別途お見積りとなります。あらかじめご了承ください。'));
body.push(LI(['要件定義の完了後に生じた', T('仕様変更および機能追加', { bold: true })]));
body.push(LI('教材の追加制作、および法令改正にともなう改訂（1科目 300,000円）'));
body.push(LI('検収完了後の改修（月額の保守・運用サービスに含まれる軽微な改修を除きます）'));
body.push(P('また、次の費用はお見積りに含まれておりません（お見積書および前提条件に記載のとおりです）。', { before: 80 }));
body.push(LI('月額の保守・運用サービスおよび法令改正監視・通知サービス　合計 85,000円／月（税別）'));
body.push(LI('サーバー利用料　月額 20,000円程度（貴社にて直接ご契約）'));

// ---------- 8 ----------
body.push(H(8, '個人情報のセキュリティ対策'));
body.push(Q('個人情報のセキュリティ対策のご教授を願います(パスの暗号化での保存・クレカ決算の場合サイト内に保存しない機能の有無・サイト全体の通信暗号化・管理画面ログインの2段階認証 等)'));
body.push(KV([
  ['パスワード', [T('暗号化ではなく、ハッシュ化して保存します', { size: 19, bold: true }), T('（bcrypt 等）。暗号化は復号が可能なため、パスワードには用いません。ハッシュ化であれば、弊社を含め何人も元のパスワードを知ることができません。お忘れの際は再設定のみとなります', { size: 19 })]],
  ['クレジットカード情報', [T('決済機能は今回の範囲外のため、該当いたしません。将来ご導入の際は、決済代行会社の画面またはトークン方式を用い、カード番号を本システムに一切保存しない構成といたします', { size: 19 })]],
  ['通信の暗号化', [T('全ページを TLS 1.2 以上で暗号化します。HSTS を設定し、平文での接続を受け付けません', { size: 19 })]],
  ['管理画面のログイン', [T('2段階認証（認証アプリによるワンタイムコード）を実装します。', { size: 19, bold: true }), T('あわせて、IPアドレスによる接続制限も設定可能です', { size: 19 })]],
  ['本人確認書類・顔画像', [T('暗号化して保存し、閲覧権限を特定の担当者に限定します。閲覧操作そのものを監査ログに記録します。保存期間の経過後は自動的に削除します', { size: 19 })]],
  ['記録の保全', [T('受講記録・照合記録は追記のみの監査ログとし、ハッシュチェーンにより後からの改ざんができない構造とします', { size: 19 })]],
]));

// ---------- 9 ----------
body.push(H(9, '早送り・スキップ・バックグラウンド再生の禁止'));
body.push(Q('当初のお話通り、早送り・スキップ・バックグラウンド再生不可機能は必ず導入お願いいたします'));
body.push(P('いずれも実装いたします。'));
body.push(LI([T('前方へのシークを禁止します。', { bold: true }), '未視聴の位置へは進めません（視聴済み箇所への巻き戻しは可能とします）']));
body.push(LI([T('再生速度を1.0倍に固定します。', { bold: true }), '変更を検知した場合は1.0倍に戻します']));
body.push(LI(T('タブが背面に回った場合、画面がロックされた場合は再生を停止し、有効視聴時間に加算しません', { bold: true })));
body.push(LI(['あわせて、', T('在席確認', { bold: true }), 'によりカメラ前からの離席を検知し、再生を自動停止します。離席していた時間は有効視聴時間から除外します']));
body.push(P('これらの制御は、法定の学科時間を実質的に満たしたことを記録で示すための中核機能です。', { before: 80 }));

// ---------- 10 ----------
body.push(H(10, '貴社発行の契約書・返金特約'));
body.push(Q('今回、着手時・中間時・納品時での支払い方法ではないため、返金特約等を弊社の発行する契約書を設けさせていただきますので、確認・ご署名後に入金とさせていただきます。(御社提示の支払期日以内での契約)'));
body.push(P('本件の契約書につきまして、1点お願いがございます。'));
body.push(NOTE('署名の前に、契約書案を拝見させてください', [
  ['今回のお見積りは、', T('ご発注時に全額を一括でお支払いいただくことを前提に、貴社特別値引きを計上しております。', { bold: true }),
   '返金の条件が広く設定されますと、この前提が成り立たなくなり、お見積金額を見直させていただく必要が生じます。'],
  ['返金の対象となる事由、返金の範囲（全額か既履行分を控除した額か）、および請求可能な期間について、契約書案を拝見のうえご相談させていただければ幸いです。'],
], 'F7EAD3', '9C6510'));
body.push(P('なお、ご契約の当事者は株式会社聖建様との認識でおります。相違がございましたら、お手数ですがお知らせください。', { before: 140 }));

// ---------- 11 ----------
body.push(H(11, '再委託について'));
body.push(Q('当初のお話通り、本件は再委託でのご契約ではなく御社の開発・構築でのご契約となります'));
body.push(P(['承知いたしました。', T('本件の開発・構築は、原則として弊社が自ら行います。', { bold: true }),
  '案件全体を他社へ委ねることはいたしません。']));
body.push(P(['ただし1点、あらかじめ申し添えます。', T('サーバー環境の構築など、一部の専門的な作業につきましては、弊社の管理のもとで協力会社の支援を受ける場合がございます。', { bold: true }),
  'その場合も次のとおりといたしますので、ご安心ください。']));
body.push(LI('協力会社に対し、弊社と同等の秘密保持義務を課します'));
body.push(LI([T('成果物の品質および履行の責任は、すべて弊社が負います。', { bold: true }), '貴社の窓口は弊社に一本化されます']));
body.push(P('契約書に再委託の禁止条項を設けられる場合は、「事前の書面による承諾を得た場合はこの限りでない」旨を併記いただけますと幸いです。', { before: 80 }));

// ---------- 末尾 ----------
body.push(new Paragraph({ spacing: { before: 400, after: 100 }, border: {
  top: { style: BorderStyle.SINGLE, size: 6, color: 'D3DDDD', space: 8 } }, children: [] }));
body.push(P([T('本書について　', { size: 18, bold: true, color: INK2 }),
  T('本書は貴社よりお預かりした確認事項に対する弊社の回答であり、契約書に代わるものではありません。最終的な権利関係および責任範囲は、双方が署名する契約書の定めによるものといたします。', { size: 18, color: INK2 })]));
body.push(P(T('株式会社スペリオル　〒464-0075 名古屋市千種区内山3-18-10　TEL 052-745-6607　担当 藤原', { size: 18, color: INK2 })));

const doc = new Document({
  creator: '株式会社スペリオル',
  title: '確認事項へのご回答',
  numbering: { config: [{
    reference: 'dot',
    levels: [{
      level: 0, format: LevelFormat.BULLET, text: '・', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 340, hanging: 220 } },
               run: { font: FONT, size: 21 } },
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
  fs.writeFileSync('/home/user/-/contract/聖建様_確認事項へのご回答.docx', b);
  console.log('saved');
});
