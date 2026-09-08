# -*- coding: utf-8 -*-
"""株式会社聖建様向け 概算見積書・注文書"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

F = "Meiryo"
NAVY="1B3A57"; ACC="B05A1E"; LIGHT="F2F5F7"; SUB="DCE6EE"; YELF="FFF6D9"
INK="1A2430"; MUT="6C7A87"; LINE="C9D4DC"; OK="1E6B4A"; BAD="9E2F27"
YEL = PatternFill("solid", fgColor=YELF)
HDRF= PatternFill("solid", fgColor=NAVY)
SUBF= PatternFill("solid", fgColor=SUB)
LTF = PatternFill("solid", fgColor=LIGHT)
thin=Side(style="thin", color=LINE); med=Side(style="medium", color=NAVY)
BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
YEN='¥#,##0;(¥#,##0);-'
MD ='#,##0.0"人日";;-'

def C(ws,r,c,v,*,b=False,fmt=None,fill=None,al=None,col=INK,sz=10,wrap=False,bd=True,va="center"):
    x=ws.cell(row=r,column=c,value=v)
    x.font=Font(name=F,bold=b,size=sz,color=col)
    if fmt:x.number_format=fmt
    if fill:x.fill=fill
    if bd:x.border=BOX
    x.alignment=Alignment(horizontal=al or "left",vertical=va,wrap_text=wrap)
    return x

def HDR(ws,row,vals,widths=None,h=22):
    for i,v in enumerate(vals,1):
        x=ws.cell(row=row,column=i,value=v)
        x.font=Font(name=F,bold=True,size=9.5,color="FFFFFF")
        x.fill=HDRF; x.border=BOX
        x.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.row_dimensions[row].height=h
    if widths:
        for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w

wb=openpyxl.Workbook()

# =====================================================================
# 03_教材科目
# =====================================================================
ws=wb.create_sheet("03_教材科目")
ws["A1"]="eラーニング教材 制作科目一覧"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="出典：建設業向けeラーニング 商品化優先順位（2026年9月）。学科のみを対象とし、実技は本見積の範囲外です。"
ws["A2"].font=Font(name=F,size=9,color=MUT)

C(ws,4,1,"制作単価（1科目あたり）",b=True,fill=SUBF)
u=C(ws,4,3,300000,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
u.comment=Comment("聖建様のご指定単価。構成台本の書き起こし、スライド制作、図解、AI音声ナレーション、バーチャル講師映像、編集、章末問題・修了確認テストの作成までを含みます。撮影は行いません。","見積前提")
C(ws,4,2,"",fill=SUBF); C(ws,4,4,"円／科目",col=MUT,sz=9)
ws.merge_cells(start_row=4,start_column=1,end_row=4,end_column=2)

HDR(ws,6,["優先\n順位","特別教育の名称","学科\n(時間)","想定ターゲット","発注\nフェーズ","制作費"],
    widths=[7,34,8,26,9,14],h=30)

SUBJ=[
 (1,"フルハーネス型墜落制止用器具",4.5,"建築・鉄骨・足場・設備",2),
 (2,"足場の組立て・解体・変更",6.0,"鳶・足場・建築・改修",1),
 (3,"自由研削用といし",4.0,"鉄工・設備・解体・内装",3),
 (4,"低圧電気取扱業務",7.0,"電気・設備・保守",3),
 (5,"アーク溶接等",11.0,"鉄骨・配管・設備・橋梁",1),
 (6,"石綿使用建築物等の解体・破砕等",4.5,"解体・改修・リフォーム",1),
 (7,"特定粉じん作業",4.5,"解体・はつり・土木",2),
 (8,"第2種酸素欠乏危険作業",5.5,"下水・槽・地下設備",2),
 (9,"高所作業車（10m未満）",6.0,"設備・電気・外装",2),
 (10,"玉掛け（1t未満）",5.0,"鉄骨・設備・土木",3),
]
r=7; first=r
for no,name,h,tgt,ph in SUBJ:
    C(ws,r,1,no,al="center")
    C(ws,r,2,name,wrap=True)
    C(ws,r,3,h,fmt='0.0"h"',al="center")
    C(ws,r,4,tgt,sz=9,col=MUT,wrap=True)
    C(ws,r,5,ph,al="center",b=True,fill=YEL,col="0000FF")
    C(ws,r,6,f"=$C$4",fmt=YEN,al="right")
    ws.row_dimensions[r].height=20; r+=1
last=r-1
C(ws,r,1,"",fill=SUBF); C(ws,r,2,"合計（10科目）",b=True,fill=SUBF)
C(ws,r,3,f"=SUM(C{first}:C{last})",fmt='0.0"h"',al="center",b=True,fill=SUBF)
C(ws,r,4,"",fill=SUBF); C(ws,r,5,"",fill=SUBF)
C(ws,r,6,f"=SUM(F{first}:F{last})",fmt=YEN,al="right",b=True,fill=SUBF)
SUBJ_TOTAL=f"'03_教材科目'!F{r}"
tot_r=r
r+=2
for ph in (1,2,3):
    C(ws,r,2,f"うち フェーズ{ph} 分",bd=False,sz=9.5)
    C(ws,r,3,f'=COUNTIF(E{first}:E{last},{ph})&"科目"',al="center",bd=False,sz=9.5,col=MUT)
    C(ws,r,6,f"=SUMIF($E${first}:$E${last},{ph},$F${first}:$F${last})",fmt=YEN,al="right",bd=False,sz=9.5)
    globals()[f"SUBJ_P{ph}"]=f"'03_教材科目'!F{r}"
    r+=1
r+=1
for t in ["■ ご記入・ご変更いただける箇所（黄色のセル）",
          "・C4：1科目あたりの制作単価。ここを変えると全科目と見積書まで自動で再計算されます。",
          "・E列：各科目をどのフェーズで制作するか。1〜3で指定します。",
          "",
          "■ 備考",
          "・「アーク溶接等」は既存の制作資産があるとのことです。既存教材を再構成して費用を圧縮できる場合は、",
          "　当該行のF列を直接上書きしてご調整ください（例：150,000）。",
          "・実技は本見積の範囲外です。実技を伴う科目は、学科部分のみをeラーニングで提供する前提です。"]:
    C(ws,r,1,t,bd=False,sz=9.5,b=t.startswith("■"),col=NAVY if t.startswith("■") else MUT); r+=1
ws.freeze_panes="A7"

# =====================================================================
# 02_見積明細
# =====================================================================
ws=wb.create_sheet("02_見積明細")
ws["A1"]="見積明細（フェーズ別）"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="黄色のセルが編集箇所です。人日単価と各項目の人日を変更すると、見積書と注文書まで自動で再計算されます。"
ws["A2"].font=Font(name=F,size=9,color=MUT)

C(ws,4,1,"人日単価",b=True,fill=SUBF); C(ws,4,2,"",fill=SUBF)
ur=C(ws,4,3,60000,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
ur.comment=Comment("設計・実装・テストを含む標準人日単価。御社の基準単価に差し替えてください。","見積前提")
C(ws,4,4,"円／人日",col=MUT,sz=9)
ws.merge_cells(start_row=4,start_column=1,end_row=4,end_column=2)
RATE="$C$4"

HDR(ws,6,["区分","項目","内容","人日","単価","金額"],widths=[10,30,52,9,12,14],h=22)

P1=[("要件定義・基本設計・画面設計","業務フローの確定、データモデル設計、API設計、画面設計、利用規約・特定商取引法表記・個人情報の取扱い方針の整備支援",14),
("開発環境・インフラ構築","Nginx／PHP-FPM／PostgreSQL／Redis／オブジェクトストレージの構築、CI/CDパイプライン、監視の設定",10),
("動画配信基盤","HLS変換（FFmpeg）、短寿命の署名付きマニフェスト配信、章単位の受講制御、MP4直リンクの排除",16),
("視聴整合性・実視聴時間の集計","heartbeatの受信と検証、未視聴位置へのシーク制限、倍速検知、可視状態の判定、サーバー側での検証済み秒数の集計、異常検知",20),
("本人確認（登録時）","その場撮影UI（getUserMedia）、ライブネス判定、本人確認書類との突合、運営の審査画面",18),
("PC↔スマートフォン引き継ぎ","QRコードとワンタイムURLによる、撮影工程のスマートフォンへの引き継ぎ",3),
("顔照合基盤（受講開始時・受講中）","1:1照合APIの連携、3段階判定（合格／保留／失敗）、ランダム発火スケジューラ、判定ログの保存",18),
("確認問題・自動採点","章末問題、修了確認テスト、出題のランダム化、不合格後の再受講制御",11),
("修了判定・修了証発行","サーバー側での修了条件の再検証、修了記録の作成、修了証PDFの発行・再発行",10),
("受講者ポータル","申込、マイページ、受講画面、進捗表示、スマートフォン対応",14),
("結合テスト・E2E・受入支援","必須E2Eシナリオの実施、受入条件の確認、リリース判定資料の作成",13)]

P2=[("運営管理画面","本人確認の承認キュー、顔照合アラートの確認、受講ログの照会、受講者・組織の管理",18),
("企業管理ポータル","発注企業の担当者向け。従業員の一括申込、受講進捗の確認、修了者一覧のCSV出力",15),
("決済機能","クレジットカード決済、請求書払い、領収書の発行",12),
("質疑応答機能","受講画面からの質問登録、担当者による回答、対応履歴の保存（通達が求める質問対応体制）",7),
("監査ログ・監査レポート","追記のみの監査証跡、監査レポートの出力、記録の長期保存",11),
("テスト・受入支援","フェーズ2機能の結合テストおよび受入支援",7)]

P3=[("修了証の真正性検証","QRコードによる公開検証ページ、失効管理、元請からの照会への対応",8),
("教材の版管理・横展開","コース・章・設問の版管理、動画差し替え時の再現性確保、多科目への横展開対応",12),
("顔照合の閾値PoC・調整","実受講者20〜30名でのFRR（本人拒否率）測定、閾値の確定、運用チューニング",6),
("テスト・受入支援","フェーズ3機能の結合テストおよび受入支援",5)]

PH_TITLE={1:"フェーズ1｜基盤構築と第1弾リリース",2:"フェーズ2｜運用機能の拡充",3:"フェーズ3｜横展開と真正性検証"}
r=7; phase_rows={}
for idx,(ph,items) in enumerate([(1,P1),(2,P2),(3,P3)],1):
    C(ws,r,1,PH_TITLE[ph],b=True,fill=SUBF,col=NAVY)
    for c in range(2,7): C(ws,r,c,"",fill=SUBF)
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=6)
    ws.row_dimensions[r].height=20
    r+=1; s0=r
    for name,desc,md in items:
        C(ws,r,1,f"フェーズ{ph}",al="center",sz=9,col=MUT)
        C(ws,r,2,name,wrap=True,b=True)
        C(ws,r,3,desc,wrap=True,sz=9,col=MUT)
        C(ws,r,4,md,fmt=MD,al="center",fill=YEL,col="0000FF",b=True)
        C(ws,r,5,f"={RATE}",fmt=YEN,al="right")
        C(ws,r,6,f"=D{r}*E{r}",fmt=YEN,al="right")
        ws.row_dimensions[r].height=30
        r+=1
    s1=r-1
    # 教材制作行
    C(ws,r,1,f"フェーズ{ph}",al="center",sz=9,col=MUT)
    C(ws,r,2,"eラーニング教材の制作",wrap=True,b=True)
    C(ws,r,3,f'=\"特別教育 学科教材の制作（\"&COUNTIF(\'03_教材科目\'!$E$7:$E$16,{ph})&\"科目）　構成台本・スライド・図解・ナレーション・バーチャル講師映像・章末問題・修了確認テスト\"',wrap=True,sz=9,col=MUT)
    C(ws,r,4,f"=COUNTIF('03_教材科目'!$E$7:$E$16,{ph})",fmt='0"科目"',al="center",col=OK)
    C(ws,r,5,"='03_教材科目'!$C$4",fmt=YEN,al="right",col=OK)
    C(ws,r,6,f"=D{r}*E{r}",fmt=YEN,al="right")
    ws.row_dimensions[r].height=30
    mat_r=r; r+=1
    # 小計
    C(ws,r,1,"",fill=LTF); C(ws,r,2,f"フェーズ{ph} 小計",b=True,fill=LTF)
    C(ws,r,3,"",fill=LTF)
    C(ws,r,4,f"=SUM(D{s0}:D{s1})",fmt=MD,al="center",b=True,fill=LTF)
    C(ws,r,5,"",fill=LTF)
    C(ws,r,6,f"=SUM(F{s0}:F{mat_r})",fmt=YEN,al="right",b=True,fill=LTF)
    phase_rows[ph]=r
    ws.row_dimensions[r].height=20
    r+=2

# 合計
tr=r
C(ws,tr,1,"",fill=SUBF); C(ws,tr,2,"小計（税抜）",b=True,fill=SUBF,sz=11)
C(ws,tr,3,"",fill=SUBF)
C(ws,tr,4,f"=D{phase_rows[1]}+D{phase_rows[2]}+D{phase_rows[3]}",fmt=MD,al="center",b=True,fill=SUBF)
C(ws,tr,5,"",fill=SUBF)
C(ws,tr,6,f"=F{phase_rows[1]}+F{phase_rows[2]}+F{phase_rows[3]}",fmt=YEN,al="right",b=True,fill=SUBF,sz=11)
adj=tr+1
C(ws,adj,2,"調整額（お値引き等。マイナスで入力）",b=True)
C(ws,adj,3,"総額を調整される場合はここに入力してください",sz=9,col=MUT)
C(ws,adj,4,""); C(ws,adj,5,"")
C(ws,adj,6,0,fmt=YEN,al="right",fill=YEL,col="0000FF",b=True)
net=adj+1
C(ws,net,2,"差引 小計（税抜）",b=True); C(ws,net,3,""); C(ws,net,4,""); C(ws,net,5,"")
C(ws,net,6,f"=F{tr}+F{adj}",fmt=YEN,al="right",b=True)
tax=net+1
C(ws,tax,2,"消費税",b=True)
C(ws,tax,3,"税率",sz=9,col=MUT)
C(ws,tax,4,0.10,fmt="0%",al="center",fill=YEL,col="0000FF",b=True)
C(ws,tax,5,"")
C(ws,tax,6,f"=ROUND(F{net}*D{tax},0)",fmt=YEN,al="right",b=True)
gt=tax+1
C(ws,gt,1,"",fill=SUBF); C(ws,gt,2,"合計（税込）",b=True,fill=SUBF,sz=12)
C(ws,gt,3,"",fill=SUBF); C(ws,gt,4,"",fill=SUBF); C(ws,gt,5,"",fill=SUBF)
C(ws,gt,6,f"=F{net}+F{tax}",fmt=YEN,al="right",b=True,fill=SUBF,sz=12)
ws.row_dimensions[gt].height=24
Q=lambda a:f"'02_見積明細'!{a}"
NET_A,TAX_A,GT_A,SUB_A=Q(f"F{net}"),Q(f"F{tax}"),Q(f"F{gt}"),Q(f"F{tr}")
MD_A=Q(f"D{tr}")
PH_A={p:Q(f"F{phase_rows[p]}") for p in (1,2,3)}
PHMD_A={p:Q(f"D{phase_rows[p]}") for p in (1,2,3)}

r=gt+2
for t in ["■ この明細の根拠",
 "・システム開発の範囲は、次の2つの設計書に記載された機能をすべて対象としています。",
 "　（1）特別教育 eラーニング 本人確認・受講確認システム 詳細設計書 rev.1（弊社作成・2026年9月5日）",
 "　（2）アーク溶接等特別教育 配信システム基本設計書 1.1（2026年9月5日）",
 "・人日は上記2設計書の機能要件から積み上げた見込値です。要件確定により増減します。",
 "・教材制作費は、聖建様よりご指定の 1科目300,000円 を単価としています。",
 "",
 "■ 編集できる箇所（黄色のセル）",
 "・C4：人日単価　・D列：各項目の人日　・F列の調整額　・消費税率",
 "・科目数と教材単価はシート03、フェーズの割り当てもシート03のE列で変更できます。"]:
    C(ws,r,1,t,bd=False,sz=9.5,b=t.startswith("■"),col=NAVY if t.startswith("■") else MUT); r+=1
ws.freeze_panes="A7"

# =====================================================================
# 01_見積書（鑑）
# =====================================================================
ws=wb.create_sheet("01_見積書",0)
for col,w in zip("ABCDEFG",[3,20,20,16,16,18,3]): ws.column_dimensions[col].width=w
ws["B2"]="御　見　積　書"
ws["B2"].font=Font(name=F,bold=True,size=22,color=NAVY)
ws.merge_cells("B2:F2"); ws["B2"].alignment=Alignment(horizontal="center")
ws.row_dimensions[2].height=34

C(ws,4,2,"見積番号",sz=9,col=MUT,bd=False)
C(ws,4,3,"SP-2026-0001",sz=9,fill=YEL,col="0000FF",bd=False)
C(ws,5,2,"発行日",sz=9,col=MUT,bd=False)
C(ws,5,3,"2026年9月8日",sz=9,fill=YEL,col="0000FF",bd=False)
C(ws,6,2,"有効期限",sz=9,col=MUT,bd=False)
C(ws,6,3,"発行日より60日間",sz=9,fill=YEL,col="0000FF",bd=False)

C(ws,8,2,"株式会社聖建　御中",b=True,sz=15,bd=False)
ws.merge_cells("B8:D8")
C(ws,9,2,"〒475-0805　愛知県半田市浜田町1丁目7番地",sz=9,col=MUT,bd=False,fill=YEL)
ws.merge_cells("B9:D9")
C(ws,10,2,"営業部　津隈 佑己 様",sz=9.5,col=MUT,bd=False)
ws.merge_cells("B10:D10")

# 発行者
C(ws,8,5,"株式会社スペリオル",b=True,sz=12,bd=False); ws.merge_cells("E8:F8")
for i,(lbl,val) in enumerate([("所在地","〒　　　　　（ご記入ください）"),
                              ("TEL / FAX","（ご記入ください）"),
                              ("登録番号（適格請求書）","T　　　　　　　　　　　"),
                              ("担当","藤原")]):
    C(ws,9+i,5,lbl,sz=8.5,col=MUT,bd=False)
    C(ws,9+i,6,val,sz=8.5,bd=False,fill=YEL if i<3 else None,col="0000FF" if i<3 else INK)

C(ws,14,2,"下記のとおりお見積り申し上げます。",sz=10,bd=False); ws.merge_cells("B14:F14")

C(ws,16,2,"件名",b=True,fill=SUBF,al="center")
C(ws,16,3,"建設業 特別教育 eラーニング配信システム 構築 および 教材制作（10科目）一式",wrap=True)
ws.merge_cells("C16:F16"); ws.row_dimensions[16].height=32

C(ws,17,2,"御見積金額",b=True,fill=SUBF,al="center")
g=C(ws,17,3,f"={GT_A}",fmt='"¥"#,##0"  （税込）"',b=True,sz=18,col=NAVY,al="left")
ws.merge_cells("C17:F17"); ws.row_dimensions[17].height=36

rows=[("納期","フェーズ1：ご発注後 約4か月　／　フェーズ2：約3か月　／　フェーズ3：約3か月"),
      ("納入場所","貴社ご指定のサーバー環境（クラウド）"),
      ("検収条件","納品後14日以内に貴社にて検収。期間内にご連絡なき場合は検収完了とみなします。"),
      ("支払条件","フェーズごとに、着手時50%・検収後50%。検収月末締め翌月末日払い。"),
      ("備考","本書は概算見積です。要件確定後に正式見積を提出いたします。詳細はシート「05_前提条件」をご確認ください。")]
r=19
for lbl,val in rows:
    C(ws,r,2,lbl,b=True,fill=SUBF,al="center")
    C(ws,r,3,val,wrap=True,sz=9.5)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=6)
    ws.row_dimensions[r].height=42 if lbl=="備考" else 30
    r+=1

r+=1
C(ws,r,2,"内訳",b=True,col=NAVY,bd=False,sz=11); r+=1
HDR(ws,r,["","区分","内容","","人日","金額（税抜）"],h=20)
ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
hr=r; r+=1
for ph in (1,2,3):
    C(ws,r,2,f"フェーズ{ph}",b=True,al="center")
    C(ws,r,3,PH_TITLE[ph].split("｜")[1],sz=9.5)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
    C(ws,r,5,f"={PHMD_A[ph]}",fmt=MD,al="center")
    C(ws,r,6,f"={PH_A[ph]}",fmt=YEN,al="right")
    r+=1
C(ws,r,2,"小計",b=True,al="center",fill=LTF); C(ws,r,3,"（税抜）",sz=9,col=MUT,fill=LTF)
C(ws,r,4,"",fill=LTF)
C(ws,r,5,f"={MD_A}",fmt=MD,al="center",b=True,fill=LTF)
C(ws,r,6,f"={SUB_A}",fmt=YEN,al="right",b=True,fill=LTF); r+=1
C(ws,r,2,"差引小計",b=True,al="center"); C(ws,r,3,"調整額を反映",sz=9,col=MUT)
for c in (4,5): C(ws,r,c,"")
C(ws,r,6,f"={NET_A}",fmt=YEN,al="right",b=True); r+=1
C(ws,r,2,"消費税",b=True,al="center"); C(ws,r,3,"",sz=9)
for c in (4,5): C(ws,r,c,"")
C(ws,r,6,f"={TAX_A}",fmt=YEN,al="right",b=True); r+=1
C(ws,r,2,"合計",b=True,al="center",fill=SUBF,sz=11); C(ws,r,3,"（税込）",sz=9,col=MUT,fill=SUBF)
for c in (4,5): C(ws,r,c,"",fill=SUBF)
C(ws,r,6,f"={GT_A}",fmt=YEN,al="right",b=True,fill=SUBF,sz=11)
ws.row_dimensions[r].height=22
r+=2
C(ws,r,2,"※ 上記のほかに、月額の運用保守費および外部サービスの従量費が発生します（シート「05_前提条件」参照）。",sz=9,col=BAD,bd=False)
ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=6)

# =====================================================================
# 04_注文書
# =====================================================================
ws=wb.create_sheet("04_注文書")
for col,w in zip("ABCDEFG",[3,20,22,14,16,18,3]): ws.column_dimensions[col].width=w
ws["B2"]="注　文　書"
ws["B2"].font=Font(name=F,bold=True,size=22,color=NAVY)
ws.merge_cells("B2:F2"); ws["B2"].alignment=Alignment(horizontal="center")
ws.row_dimensions[2].height=34

C(ws,4,2,"注文番号",sz=9,col=MUT,bd=False); C(ws,4,3,"（ご記入ください）",sz=9,fill=YEL,col="0000FF",bd=False)
C(ws,5,2,"発行日",sz=9,col=MUT,bd=False);   C(ws,5,3,"　　年　　月　　日",sz=9,fill=YEL,col="0000FF",bd=False)

C(ws,7,2,"株式会社スペリオル　御中",b=True,sz=15,bd=False); ws.merge_cells("B7:D7")
C(ws,8,2,"担当　藤原 様",sz=9.5,col=MUT,bd=False)

C(ws,7,5,"発注者",sz=8.5,col=MUT,bd=False)
C(ws,8,5,"株式会社聖建",b=True,sz=12,bd=False); ws.merge_cells("E8:F8")
C(ws,9,5,"〒475-0805　愛知県半田市浜田町1丁目7番地",sz=8.5,col=MUT,bd=False); ws.merge_cells("E9:F9")
C(ws,10,5,"TEL 0569-84-9669",sz=8.5,col=MUT,bd=False)
C(ws,11,5,"ご担当　　　　　　　　　　　　　印",sz=9,bd=False); ws.merge_cells("E11:F11")

C(ws,13,2,"下記のとおり注文いたします。",sz=10,bd=False); ws.merge_cells("B13:F13")

C(ws,15,2,"件名",b=True,fill=SUBF,al="center")
C(ws,15,3,"建設業 特別教育 eラーニング配信システム 構築 および 教材制作（10科目）一式",wrap=True)
ws.merge_cells("C15:F15"); ws.row_dimensions[15].height=30
C(ws,16,2,"根拠見積",b=True,fill=SUBF,al="center")
C(ws,16,3,"='01_見積書'!C4&\"（\"&'01_見積書'!C5&\" 発行）\"",wrap=True,sz=9.5)
ws.merge_cells("C16:F16")

C(ws,18,2,"発注対象",b=True,col=NAVY,bd=False,sz=11)
C(ws,18,3,"発注するフェーズのD列に 1 を入力してください（既定はフェーズ1のみ）",sz=9,col=MUT,bd=False)
ws.merge_cells("C18:F18")
HDR(ws,19,["","フェーズ","内容","発注\n(1で選択)","人日","金額（税抜）"],h=26)
r=20; o0=r
for ph,flag in [(1,1),(2,0),(3,0)]:
    C(ws,r,2,f"フェーズ{ph}",b=True,al="center")
    C(ws,r,3,PH_TITLE[ph].split("｜")[1],sz=9.5,wrap=True)
    C(ws,r,4,flag,al="center",b=True,fill=YEL,col="0000FF")
    C(ws,r,5,f"=IF(D{r}=1,{PHMD_A[ph]},0)",fmt=MD,al="center")
    C(ws,r,6,f"=IF(D{r}=1,{PH_A[ph]},0)",fmt=YEN,al="right")
    ws.row_dimensions[r].height=26; r+=1
o1=r-1
C(ws,r,2,"",fill=LTF); C(ws,r,3,"発注小計（税抜）",b=True,fill=LTF)
C(ws,r,4,"",fill=LTF)
C(ws,r,5,f"=SUM(E{o0}:E{o1})",fmt=MD,al="center",b=True,fill=LTF)
C(ws,r,6,f"=SUM(F{o0}:F{o1})",fmt=YEN,al="right",b=True,fill=LTF)
osub=r; r+=1
C(ws,r,2,""); C(ws,r,3,"消費税"); C(ws,r,4,f"='02_見積明細'!D{tax}",fmt="0%",al="center"); C(ws,r,5,"")
C(ws,r,6,f"=ROUND(F{osub}*D{r},0)",fmt=YEN,al="right",b=True)
otax=r; r+=1
C(ws,r,2,"",fill=SUBF); C(ws,r,3,"ご注文金額（税込）",b=True,fill=SUBF,sz=12)
for c in (4,5): C(ws,r,c,"",fill=SUBF)
C(ws,r,6,f"=F{osub}+F{otax}",fmt=YEN,al="right",b=True,fill=SUBF,sz=13)
ws.row_dimensions[r].height=26
r+=2
for lbl,val in [("納期","フェーズ1：ご発注後 約4か月　／　フェーズ2：約3か月　／　フェーズ3：約3か月"),
                ("納入場所","貴社ご指定のサーバー環境（クラウド）"),
                ("検収条件","納品後14日以内に検収。期間内にご連絡なき場合は検収完了とみなします。"),
                ("支払条件","着手時50%・検収後50%。検収月末締め翌月末日払い。"),
                ("備考","フェーズ2以降は、前フェーズの検収後に別途発注書を発行するものとします。")]:
    C(ws,r,2,lbl,b=True,fill=SUBF,al="center")
    C(ws,r,3,val,wrap=True,sz=9.5)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=6)
    ws.row_dimensions[r].height=28; r+=1

# =====================================================================
# 05_前提条件
# =====================================================================
ws=wb.create_sheet("05_前提条件")
ws.column_dimensions["A"].width=3
ws.column_dimensions["B"].width=30
ws.column_dimensions["C"].width=88
ws["B1"]="見積の前提条件"
ws["B1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24

BLK=[("■ 1. 本見積に含まれるもの",NAVY,[
  ("設計書の範囲","次の2つの設計書に記載された機能の設計・実装・テスト・受入支援。"),
  ("",	"（1）特別教育 eラーニング 本人確認・受講確認システム 詳細設計書 rev.1（2026年9月5日）"),
  ("",	"（2）アーク溶接等特別教育 配信システム基本設計書 1.1（2026年9月5日）"),
  ("教材制作","特別教育 学科教材 10科目の制作。構成台本の書き起こし、スライド、図解、ナレーション、バーチャル講師映像、章末問題・修了確認テストの作成を含みます。"),
]),
("■ 2. 本見積に含まれないもの（別途費用）",BAD,[
  ("月額 運用保守費","サーバー・CDN費用 40,000〜80,000円／月、保守・障害対応 50,000〜80,000円／月（いずれも概算）。正式契約は別途締結します。"),
  ("外部サービスの従量費","顔照合API、ライブネス判定、動画配信、決済手数料。いずれも受講者数に比例するため、実績に応じたご請求となります。"),
  ("教材の監修","内容の監修は聖建様の社内有資格者が担われる前提です。外部の労働安全衛生コンサルタント等に監修を依頼される場合、1科目あたり350,000円程度が別途必要です。"),
  ("実技教育","実技はオンラインで代替できません。実技は受講者の所属事業者が実施する前提であり、本見積の範囲外です。"),
  ("集客・販売促進","ランディングページ、SEO、Web広告の運用。個人（一人親方）向けに販売される場合、別途ご検討をお勧めします。"),
  ("撮影","バーチャル講師とAI音声ナレーションを用いるため、実写撮影・スタジオ費用は含みません。"),
]),
("■ 3. 責任の分担",NAVY,[
  ("弊社が保証すること","教材が安全衛生特別教育規程の定める科目・範囲・時間を満たすこと（監修者名を明示）。受講記録を証明可能な形で保存・出力できること。"),
  ("弊社が保証しないこと","受講者の所属事業者が負う労働安全衛生法第59条第3項の実施義務が果たされたこと。労働災害が発生しないこと。行政機関が個別の事案において本教材を適法と判断すること。"),
  ("聖建様にお願いすること","教材内容の監修（社内有資格者による）。監修者の保有資格を修了証の写しで確認いただくこと。運営レビュー（本人確認の目視審査）の担当者のご手配。"),
  ("事前相談","所轄の労働基準監督署への事前相談を、章立ての設計が固まった段階で共同で実施することを推奨します（費用は発生しません）。"),
]),
("■ 4. 金額の前提",NAVY,[
  ("人日単価","60,000円／人日を標準としています（シート02のC4で変更可能）。"),
  ("教材制作単価","300,000円／科目。聖建様よりご指定の単価です（シート03のC4で変更可能）。"),
  ("概算である旨","本書は要件確定前の概算見積です。要件定義の完了後、正式なお見積りを提出いたします。人日は設計書の機能要件から積み上げた見込値であり、要件確定により増減します。"),
  ("有効期限","発行日より60日間。以降は再度お見積りいたします。"),
]),
("■ 5. 分割発注について",NAVY,[
  ("推奨","フェーズごとの分割発注を推奨します。フェーズ1の検収後に、フェーズ2以降の要否と内容をあらためてご判断いただけます。"),
  ("最小構成","フェーズ1のみでも、1科目を販売できる状態になります。まずここから着手されることをお勧めします。"),
])]
r=3
for head,col,items in BLK:
    C(ws,r,2,head,b=True,sz=11.5,col=col,fill=SUBF,bd=False)
    C(ws,r,3,"",fill=SUBF,bd=False)
    ws.row_dimensions[r].height=20; r+=1
    for lbl,txt in items:
        C(ws,r,2,lbl,b=True,sz=9.5,va="top",wrap=True)
        C(ws,r,3,txt,sz=9.5,wrap=True,va="top",col=INK)
        ws.row_dimensions[r].height=34 if len(txt)>75 else 22
        r+=1
    r+=1

if "Sheet" in wb.sheetnames: del wb["Sheet"]
wb._sheets = [wb[n] for n in ["01_見積書","02_見積明細","03_教材科目","04_注文書","05_前提条件"]]
wb.save("/home/user/-/quote/聖建様_概算見積書・注文書.xlsx")
print("saved")
