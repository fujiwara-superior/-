# -*- coding: utf-8 -*-
"""株式会社聖建様向け 概算見積書・注文書 rev.3（6科目・総額660万円・全額一括前払い）"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

F="Meiryo"; NAVY="1B3A57"; ACC="B05A1E"; LIGHT="F2F5F7"; SUB="DCE6EE"
INK="1A2430"; MUT="6C7A87"; LINE="C9D4DC"; OK="1E6B4A"; BAD="9E2F27"
YEL=PatternFill("solid",fgColor="FFF6D9"); HDRF=PatternFill("solid",fgColor=NAVY)
SUBF=PatternFill("solid",fgColor=SUB); LTF=PatternFill("solid",fgColor=LIGHT)
OKF=PatternFill("solid",fgColor="DFF0E6"); BADF=PatternFill("solid",fgColor="FBE3E1")
thin=Side(style="thin",color=LINE); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
YEN='¥#,##0;▲¥#,##0;-'; MD='#,##0.0"人日";"▲"#,##0.0"人日";-'

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
        x.font=Font(name=F,bold=True,size=9.5,color="FFFFFF"); x.fill=HDRF; x.border=BOX
        x.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.row_dimensions[row].height=h
    if widths:
        for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w

wb=openpyxl.Workbook()

# =====================================================================
# 04_教材科目（6科目）
# =====================================================================
ws=wb.create_sheet("04_教材科目")
ws["A1"]="eラーニング教材 制作科目一覧（6科目）"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="A列に 1 を入れた科目が制作対象になります。10科目から6科目へ絞り込んだ状態です。"
ws["A2"].font=Font(name=F,size=9,color=MUT)
C(ws,4,1,"",fill=SUBF); C(ws,4,2,"制作単価（1科目あたり）",b=True,fill=SUBF)
u=C(ws,4,4,300000,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
u.comment=Comment("聖建様ご指定の単価。構成台本、スライド、図解、ナレーション、バーチャル講師映像、章末問題・修了確認テストの作成まで。法令改正時の改訂も同額。","見積前提")
C(ws,4,3,"",fill=SUBF); C(ws,4,5,"円／科目",col=MUT,sz=9); C(ws,4,6,"",fill=SUBF)
HDR(ws,6,["制作\n(1で選択)","優先\n順位","特別教育の名称","学科\n(時間)","選定・除外の理由","制作費"],
    widths=[9,7,32,8,38,14],h=32)
SUBJ=[(1,1,"フルハーネス型墜落制止用器具",4.5,"対象職種が最も広く、初期商品群の中核"),
(1,2,"足場の組立て・解体・変更",6.0,"実技の定めがなく、学科のみで修了まで完結する"),
(1,3,"自由研削用といし",4.0,"幅広い工種で利用され、学科時間もコンパクト"),
(0,4,"低圧電気取扱業務",7.0,"学科7時間と最長級。制作負担が大きく第2次へ"),
(1,5,"アーク溶接等",11.0,"既存の制作資産を活用できるため優先"),
(1,6,"石綿使用建築物等の解体・破砕等",4.5,"実技なし。解体・改修の需要が継続的にある"),
(1,7,"特定粉じん作業",4.5,"実技なし。建設現場の有害業務として汎用性が高い"),
(0,8,"第2種酸素欠乏危険作業",5.5,"対象職種が設備・土木に限られるため第2次へ"),
(0,9,"高所作業車（10m未満）",6.0,"実技3時間を伴い、対象が限定的なため第2次へ"),
(0,10,"玉掛け（1t未満）",5.0,"実技4時間を伴い、対象が限定的なため第2次へ")]
r=7; first=r
for flag,no,name,h,why in SUBJ:
    C(ws,r,1,flag,al="center",b=True,fill=YEL,col="0000FF")
    C(ws,r,2,no,al="center",col=MUT)
    C(ws,r,3,name,wrap=True,col=INK if flag else MUT,b=bool(flag))
    C(ws,r,4,h,fmt='0.0"h"',al="center",col=INK if flag else MUT)
    C(ws,r,5,why,sz=9,col=MUT,wrap=True)
    C(ws,r,6,f"=IF(A{r}=1,$D$4,0)",fmt=YEN,al="right")
    ws.row_dimensions[r].height=26; r+=1
last=r-1
C(ws,r,1,f"=SUM(A{first}:A{last})",al="center",b=True,fill=SUBF,fmt='0"科目"')
C(ws,r,2,"",fill=SUBF); C(ws,r,3,"制作対象 合計",b=True,fill=SUBF)
C(ws,r,4,f"=SUMPRODUCT($A{first}:$A{last},$D{first}:$D{last})",fmt='0.0"h"',al="center",b=True,fill=SUBF)
C(ws,r,5,"",fill=SUBF)
C(ws,r,6,f"=SUM(F{first}:F{last})",fmt=YEN,al="right",b=True,fill=SUBF)
MATN=f"'04_教材科目'!$F${r}"; MATCNT=f"'04_教材科目'!$A${r}"; r+=2
for t in ["■ 選定の考え方",
"・実技の定めがない科目（足場・石綿・特定粉じん）を優先し、最短で販売を開始できるようにしています。",
"・アーク溶接等は既存の制作資産があるため、優先して商品化します。",
"・除外した4科目は、学科時間が長いか対象職種が限定的なものです。第2次発注でのご検討をお願いします。",
"・A列の 1 と 0 を入れ替えれば、対象科目を自由に変更できます（見積金額も自動で追随します）。"]:
    C(ws,r,1,t,bd=False,sz=9.5,b=t.startswith("■"),col=NAVY if t.startswith("■") else MUT); r+=1
ws.freeze_panes="A7"

# =====================================================================
# 03_見積明細
# =====================================================================
ws=wb.create_sheet("03_見積明細")
ws["A1"]="見積明細"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="黄色のセルが編集箇所です。人日単価と各項目の人日を変えると、見積書・増減内訳・注文書まで自動で再計算されます。"
ws["A2"].font=Font(name=F,size=9,color=MUT)
C(ws,4,1,"人日単価",b=True,fill=SUBF); C(ws,4,2,"",fill=SUBF)
C(ws,4,3,50000,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
C(ws,4,4,"円／人日",col=MUT,sz=9); ws.merge_cells(start_row=4,start_column=1,end_row=4,end_column=2)
RATE="$C$4"
HDR(ws,6,["区分","項目","内容","数量","単価","金額"],widths=[12,28,54,9,12,14],h=22)
SYS=[("要件定義・基本設計","業務フローの確定、データモデル設計、API設計、画面設計、利用規約・特定商取引法表記・個人情報の取扱い方針の整備支援",8),
("AWS本番環境の構築","VPC、EC2、RDS PostgreSQL（Multi-AZ・自動バックアップ・PITR）、ElastiCache、ALB、S3、CloudFront、CloudWatch の構築とCI/CD",7),
("動画配信基盤","HLS変換（FFmpeg）、短寿命の署名付きマニフェスト配信、章単位の受講制御、MP4直リンクの排除",12),
("視聴整合性・実視聴時間の集計","heartbeatの受信と検証、未視聴位置へのシーク制限、倍速検知、可視状態の判定、サーバー側での検証済み秒数の集計、異常検知",15),
("本人確認（登録時）","その場撮影UI（getUserMedia）、ライブネス判定、本人確認書類との突合、顔の1:1照合、運営の審査画面",11),
("PC↔スマートフォン引き継ぎ","QRコードとワンタイムURLによる撮影工程の引き継ぎ",3),
("顔照合（受講開始時のみ）","各章の再生開始時に1:1照合を実施。受講中のランダム照合は今回の範囲外とします",4),
("確認問題・自動採点","章末問題、修了確認テスト、出題のランダム化、不合格後の再受講制御",7),
("修了判定・修了証の発行","サーバー側での修了条件の再検証、修了記録の作成、修了証PDFの発行・再発行・ダウンロード",7),
("決済機能（ゼウス連携）","株式会社ゼウスのデータ転送（API）型による決済組み込み。申込から入金確認、受講開始までの自動化。領収書の発行",10),
("受講者ポータル","申込、マイページ、受講画面、進捗表示、スマートフォン対応",9),
("運営管理画面（基本機能）","本人確認の承認キュー、受講状況の一覧、受講ログの照会、修了証の発行管理",4),
("結合テスト・E2E・受入支援","必須E2Eシナリオの実施、受入条件の確認、リリース判定資料の作成",6)]
r=7
C(ws,r,1,"1. システム開発",b=True,fill=SUBF,col=NAVY)
for c in range(2,7): C(ws,r,c,"",fill=SUBF)
ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=6); r+=1; s0=r
for name,desc,md in SYS:
    C(ws,r,1,"システム",al="center",sz=9,col=MUT)
    C(ws,r,2,name,wrap=True,b=True); C(ws,r,3,desc,wrap=True,sz=9,col=MUT)
    C(ws,r,4,md,fmt=MD,al="center",fill=YEL,col="0000FF",b=True)
    C(ws,r,5,f"={RATE}",fmt=YEN,al="right"); C(ws,r,6,f"=D{r}*E{r}",fmt=YEN,al="right")
    ws.row_dimensions[r].height=30; r+=1
s1=r-1
C(ws,r,1,"",fill=LTF); C(ws,r,2,"システム開発 小計",b=True,fill=LTF); C(ws,r,3,"",fill=LTF)
C(ws,r,4,f"=SUM(D{s0}:D{s1})",fmt=MD,al="center",b=True,fill=LTF); C(ws,r,5,"",fill=LTF)
C(ws,r,6,f"=SUM(F{s0}:F{s1})",fmt=YEN,al="right",b=True,fill=LTF)
sysrow=r; r+=2
C(ws,r,1,"2. eラーニング教材の制作",b=True,fill=SUBF,col=NAVY)
for c in range(2,7): C(ws,r,c,"",fill=SUBF)
ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=6); r+=1
C(ws,r,1,"教材",al="center",sz=9,col=MUT)
C(ws,r,2,"特別教育 学科教材の制作",b=True,wrap=True)
C(ws,r,3,'="6科目・学科"&TEXT(\'04_教材科目\'!$D$17,"0.0")&"時間分。構成台本の書き起こし、スライド、図解、ナレーション、バーチャル講師映像、章末問題・修了確認テストの作成"',wrap=True,sz=9,col=MUT)
C(ws,r,4,f"={MATCNT}",fmt='0"科目"',al="center",col=OK)
C(ws,r,5,"='04_教材科目'!$D$4",fmt=YEN,al="right",col=OK)
C(ws,r,6,f"={MATN}",fmt=YEN,al="right")
ws.row_dimensions[r].height=36; matrow=r; r+=2
tr=r
C(ws,tr,1,"",fill=SUBF); C(ws,tr,2,"小計（税抜）",b=True,fill=SUBF,sz=11); C(ws,tr,3,"",fill=SUBF)
C(ws,tr,4,f"=D{sysrow}",fmt=MD,al="center",b=True,fill=SUBF); C(ws,tr,5,"",fill=SUBF)
C(ws,tr,6,f"=F{sysrow}+F{matrow}",fmt=YEN,al="right",b=True,fill=SUBF,sz=11)
adj=tr+1
C(ws,adj,2,"全額一括前払いによる調整",b=True,col=OK)
C(ws,adj,3,"ご発注時に全額をお支払いいただく条件に対する調整額です",sz=9,col=MUT)
C(ws,adj,4,""); C(ws,adj,5,"")
C(ws,adj,6,-350000,fmt=YEN,al="right",fill=YEL,col="0000FF",b=True)
net=adj+1
C(ws,net,2,"差引 小計（税抜）",b=True)
for c in (3,4,5): C(ws,net,c,"")
C(ws,net,6,f"=F{tr}+F{adj}",fmt=YEN,al="right",b=True)
tax=net+1
C(ws,tax,2,"消費税",b=True); C(ws,tax,3,"税率",sz=9,col=MUT)
C(ws,tax,4,0.10,fmt="0%",al="center",fill=YEL,col="0000FF",b=True); C(ws,tax,5,"")
C(ws,tax,6,f"=ROUND(F{net}*D{tax},0)",fmt=YEN,al="right",b=True)
gt=tax+1
C(ws,gt,1,"",fill=SUBF); C(ws,gt,2,"合計（税込）",b=True,fill=SUBF,sz=12)
for c in (3,4,5): C(ws,gt,c,"",fill=SUBF)
C(ws,gt,6,f"=F{net}+F{tax}",fmt=YEN,al="right",b=True,fill=SUBF,sz=12)
ws.row_dimensions[gt].height=24
Q=lambda a:f"'03_見積明細'!{a}"
NET_A,TAX_A,GT_A,SUB_A,ADJ_A=Q(f"F{net}"),Q(f"F{tax}"),Q(f"F{gt}"),Q(f"F{tr}"),Q(f"F{adj}")
SYS_A,MD_A=Q(f"F{sysrow}"),Q(f"D{sysrow}")
r=gt+2
for t in ["■ 今回の範囲から外したもの（第2次発注のご相談）",
 "・受講中のランダム顔照合（受講の途中で予告なく撮影し、本人が継続して受講しているかを確認する機能）",
 "・企業管理ポータル（従業員の一括申込、進捗確認、修了者一覧のCSV出力）",
 "・運営管理画面の拡張（監査レポート出力、組織管理、売上・入金管理）",
 "・受講画面からの質疑応答機能（当面はメールで対応）",
 "・修了証の真正性検証ページ（QRコード）",
 "・教材 4科目（低圧電気、第2種酸欠、高所作業車、玉掛け）",
 "",
 "■ 本見積に含まれないもの（別途）",
 "・月額の保守・運用サービスおよび法令改正監視・通知サービス（シート05）",
 "・AWSの利用料（聖建様のご負担）、株式会社ゼウスとのご契約（初期費用無料・月額3,000円・決済手数料）",
 "・法令改正にともなう教材の改訂（1科目300,000円・都度発注）"]:
    C(ws,r,1,t,bd=False,sz=9.5,b=t.startswith("■"),col=NAVY if t.startswith("■") else MUT); r+=1
ws.freeze_panes="A7"

# =====================================================================
# 02_増減内訳
# =====================================================================
ws=wb.create_sheet("02_増減内訳")
ws["A1"]="前回お見積りからの増減内訳"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="2026年9月8日付 SP-2026-0001（税抜 8,000,000円）から、今回のお見積りに至るまでの内訳です。"
ws["A2"].font=Font(name=F,size=9,color=MUT)
for c,w in zip("ABCDE",[6,40,16,14,56]): ws.column_dimensions[c].width=w
HDR(ws,4,["","項目","増減","人日","理由・内容"],h=22)
r=5
C(ws,r,1,"",fill=LTF); C(ws,r,2,"前回お見積り（2026年9月8日）",b=True,fill=LTF)
C(ws,r,3,8000000,fmt=YEN,al="right",b=True,fill=LTF)
C(ws,r,4,100,fmt=MD,al="center",fill=LTF,col=MUT)
C(ws,r,5,"教材10科目＋システム100人日",sz=9,col=MUT,fill=LTF)
base=r; r+=1
ITEMS=[("①","教材を10科目から6科目へ",-1200000,None,
  "低圧電気、第2種酸欠、高所作業車、玉掛けの4科目を第2次発注へ。1科目300,000円×4科目",BAD),
("②","決済機能（ゼウス連携）の追加",500000,10,
  "株式会社ゼウスのAPI型決済を組み込み。申込から入金確認、受講開始までを自動化。今回の新規ご要望",OK),
("③","受講中のランダム顔照合を除外",-350000,-7,
  "受講の途中で予告なく撮影する機能を今回の範囲から外します。登録時の本人確認と受講開始時の照合は維持",BAD),
("④","全額一括前払いによる調整",-350000,None,
  "ご発注時に全額をお支払いいただく条件に対する調整。分割回収に比べ資金負担と与信リスクが下がるため",BAD)]
for no,name,amt,md,why,col in ITEMS:
    C(ws,r,1,no,al="center",b=True,col=col)
    C(ws,r,2,name,b=True,wrap=True)
    C(ws,r,3,amt,fmt=YEN,al="right",b=True,col=col)
    C(ws,r,4,md if md is not None else "",fmt=MD,al="center",col=col if md else MUT)
    C(ws,r,5,why,sz=9,col=MUT,wrap=True)
    ws.row_dimensions[r].height=34; r+=1
last_i=r-1
C(ws,r,1,"",fill=SUBF); C(ws,r,2,"今回のお見積り（税抜）",b=True,fill=SUBF,sz=12)
C(ws,r,3,f"=C{base}+SUM(C{base+1}:C{last_i})",fmt=YEN,al="right",b=True,fill=SUBF,sz=12)
C(ws,r,4,f"={MD_A}",fmt=MD,al="center",b=True,fill=SUBF)
C(ws,r,5,"教材6科目＋システム103人日",sz=9,col=MUT,fill=SUBF)
ws.row_dimensions[r].height=26; totrow=r; r+=1
C(ws,r,2,"（検算）見積明細の小計＋調整額",sz=9,col=MUT,bd=False)
C(ws,r,3,f"={NET_A}",fmt=YEN,al="right",sz=9,col=MUT,bd=False)
r+=2

C(ws,r,2,"増減の内訳（金額の動き）",b=True,col=NAVY,bd=False,sz=11.5); r+=1
BRK=[("減額の要因",BAD,[("教材 4科目の削減","▲1,200,000"),("受講中ランダム照合の除外","▲350,000"),("全額一括前払いによる調整","▲350,000"),("小計","▲1,900,000")]),
("増額の要因",OK,[("決済機能（ゼウス連携）の追加","＋500,000"),("小計","＋500,000")])]
for head,col,rows_ in BRK:
    C(ws,r,2,head,b=True,col=col,fill=SUBF); C(ws,r,3,"",fill=SUBF); C(ws,r,4,"",fill=SUBF); C(ws,r,5,"",fill=SUBF)
    r+=1
    for lbl,val in rows_:
        C(ws,r,2,lbl,b=(lbl=="小計"),sz=9.5)
        C(ws,r,3,val,al="right",b=(lbl=="小計"),col=col,sz=9.5)
        C(ws,r,4,""); C(ws,r,5,"")
        r+=1
    r+=1
C(ws,r,2,"差引",b=True,fill=SUBF,sz=11); C(ws,r,3,"▲1,400,000",al="right",b=True,fill=SUBF,sz=11,col=BAD)
C(ws,r,4,"",fill=SUBF); C(ws,r,5,"8,000,000 − 1,400,000 ＝ 6,600,000",sz=9,col=MUT,fill=SUBF)
r+=2
NOTE=[("■ この表の使い方",NAVY),
("単純な値引きではなく、次の3つの対価として金額が下がっていることをご説明いただけます。",MUT),
("　① 納品する教材が4科目減ること（成果物の減少）",MUT),
("　② 受講中のランダム照合を実装しないこと（機能の減少）",MUT),
("　③ 全額を前払いいただくこと（支払条件の対価）",MUT),
("",None),
("一方で、決済機能（ゼウス連携）は新しいご要望のため増額しています。",MUT),
("減る理由と増える理由の両方を示すことで、価格の根拠が明確になります。",OK),
("",None),
("■ 月額費用の変更",NAVY),
("保守・運用サービス　80,000円 → 60,000円（▲20,000円）",MUT),
("　受講中のランダム照合を実装しないため、照合結果の目視確認という運用作業がなくなります。",MUT),
("　運用工数が実際に減るため、その分を月額に反映しています。",OK),
("法令改正監視・通知サービス　20,000円（変更なし）",MUT),
("　　　　　　　　　　　　　　合計 80,000円／月（サーバー費を除く・税抜）",MUT)]
for t,col in NOTE:
    C(ws,r,2,t,bd=False,sz=9.5,b=(col in (NAVY,)),col=col or MUT); r+=1

# =====================================================================
# 05_月額・別途費用
# =====================================================================
ws=wb.create_sheet("05_月額・別途費用")
ws["A1"]="月額費用と別途費用"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="初期費用（シート01）とは別に、下記が発生します。金額はすべて税抜です。"
ws["A2"].font=Font(name=F,size=9,color=MUT)
HDR(ws,4,["区分","項目","月額／単価","内容"],widths=[16,32,16,66],h=22)
r=5
C(ws,r,1,"弊社へ\nお支払い",al="center",b=True,wrap=True,fill=LTF)
C(ws,r,2,"保守・運用サービス",b=True)
C(ws,r,3,60000,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
C(ws,r,4,"システムの死活監視とエラー監視、障害対応、軽微な改修、受講者からの問い合わせ一次対応、月次レポート。受講中のランダム照合を実装しないため、照合結果の目視確認が不要となり、前回の80,000円から減額しています",wrap=True,sz=9.5)
ws.row_dimensions[r].height=46; m0=r; r+=1
C(ws,r,1,"",fill=LTF); C(ws,r,2,"法令改正監視・通知サービス",b=True)
C(ws,r,3,20000,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
C(ws,r,4,"月2回（月初・月中）、官報・厚生労働省の通達・パブリックコメントを確認。対象6科目に影響しうる改正を検知したら、聖建様と弊社の双方へ自動でメール通知。月次で「該当なし」も含めて報告",wrap=True,sz=9.5)
ws.row_dimensions[r].height=40; m1=r; r+=1
C(ws,r,1,"",fill=SUBF); C(ws,r,2,"月額 小計（サーバー費を除く）",b=True,fill=SUBF,sz=11)
C(ws,r,3,f"=SUM(C{m0}:C{m1})",fmt=YEN,al="right",b=True,fill=SUBF,sz=12)
C(ws,r,4,"聖建様のご指定額です",sz=9,col=MUT,fill=SUBF); MON=f"'05_月額・別途費用'!C{r}"; r+=2
C(ws,r,1,"お客様\nご負担",al="center",b=True,wrap=True,fill=LTF)
C(ws,r,2,"AWS 利用料（構成B・推奨）",b=True)
C(ws,r,3,22800,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
C(ws,r,4,"EC2、RDS PostgreSQL（Multi-AZ・自動バックアップ・PITR）、ElastiCache、ALB、S3、CloudFront、CloudWatch。聖建様名義でのご契約を推奨します",wrap=True,sz=9.5)
ws.row_dimensions[r].height=36; aws_r=r; r+=1
C(ws,r,1,"",fill=LTF); C(ws,r,2,"従量課金（顔照合・配信ほか）",b=True)
C(ws,r,3,"実費",al="right",b=True,col=MUT)
C(ws,r,4,"顔照合API、ライブネス判定、CDN転送、メール送信。受講者100名／月で約600円（AWS利用料に含まれます）",wrap=True,sz=9.5)
ws.row_dimensions[r].height=28; r+=1
C(ws,r,1,"",fill=LTF); C(ws,r,2,"株式会社ゼウス（決済代行）",b=True)
C(ws,r,3,3000,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
C(ws,r,4,"初期費用は無料。月額3,000円。クレジットカード決済手数料は最大3.5%（売上に応じて発生）。聖建様が直接ご契約いただきます",wrap=True,sz=9.5)
ws.row_dimensions[r].height=32; zeus_r=r; r+=2
C(ws,r,1,"都度\n発注",al="center",b=True,wrap=True,fill=LTF)
C(ws,r,2,"教材の改訂（法令改正対応）",b=True)
C(ws,r,3,"='04_教材科目'!D4",fmt=YEN,al="right",b=True,col=OK)
C(ws,r,4,"1科目あたり。新規制作と同額です。台本の修正、該当部分の動画の再生成と差し替え、教材の版の更新を含みます",wrap=True,sz=9.5)
ws.row_dimensions[r].height=30; r+=1
C(ws,r,1,"",fill=LTF); C(ws,r,2,"教材の追加制作（第2次）",b=True)
C(ws,r,3,"='04_教材科目'!D4",fmt=YEN,al="right",b=True,col=OK)
C(ws,r,4,"低圧電気、第2種酸欠、高所作業車、玉掛けの4科目。1科目あたり300,000円",wrap=True,sz=9.5)
ws.row_dimensions[r].height=28; r+=2
C(ws,r,1,"参考：月額の合計（聖建様のご負担総額）",b=True,col=NAVY,bd=False,sz=11); r+=1
C(ws,r,2,"弊社サービス ＋ AWS ＋ ゼウス月額",b=True)
C(ws,r,3,f"={MON}+C{aws_r}+C{zeus_r}",fmt=YEN,al="right",b=True,sz=12,col=ACC)
C(ws,r,4,"税抜。決済手数料は売上に応じて別途発生します",sz=9,col=MUT)
r+=2
for t,col in [("■ 月額80,000円の根拠",NAVY),
("聖建様のご指定額（サーバー費を除く月額80,000円）にあわせ、内訳を次のとおり組みました。",MUT),
("・保守・運用サービス　60,000円　… 受講中のランダム照合を実装しないため、照合結果の",MUT),
("　　目視確認という運用作業がなくなります。その分を減額しています。",OK),
("・法令改正監視・通知サービス　20,000円　… 前回から変更ありません。",MUT),
("",None),
("■ ゼウスについて",NAVY),
("・株式会社ゼウス（1994年設立・国内の決済代行事業者）との契約は聖建様が直接行っていただきます。",MUT),
("・初期費用は無料、月額3,000円、クレジットカード決済手数料は最大3.5%です（2026年9月時点）。",MUT),
("・接続方式はデータ転送（API）型を想定しています。申込から入金確認、受講開始までを自動化します。",MUT),
("・最新の料金と条件は、ご契約前にゼウス社へ直接ご確認ください。",BAD)]:
    C(ws,r,1,t,bd=False,sz=9.5,b=(col in (NAVY,)),col=col or MUT); r+=1

# =====================================================================
# 01_見積書
# =====================================================================
ws=wb.create_sheet("01_見積書",0)
for col,w in zip("ABCDEFG",[3,20,20,16,16,18,3]): ws.column_dimensions[col].width=w
ws["B2"]="御　見　積　書"; ws["B2"].font=Font(name=F,bold=True,size=22,color=NAVY)
ws.merge_cells("B2:F2"); ws["B2"].alignment=Alignment(horizontal="center"); ws.row_dimensions[2].height=34
C(ws,4,2,"見積番号",sz=9,col=MUT,bd=False); C(ws,4,3,"SP-2026-0002",sz=9,fill=YEL,col="0000FF",bd=False)
C(ws,5,2,"発行日",sz=9,col=MUT,bd=False);   C(ws,5,3,"2026年9月12日",sz=9,fill=YEL,col="0000FF",bd=False)
C(ws,6,2,"有効期限",sz=9,col=MUT,bd=False); C(ws,6,3,"発行日より30日間",sz=9,fill=YEL,col="0000FF",bd=False)
C(ws,8,2,"株式会社聖建　御中",b=True,sz=15,bd=False); ws.merge_cells("B8:D8")
C(ws,9,2,"〒475-0805　愛知県半田市浜田町1丁目7番地",sz=9,col=MUT,bd=False); ws.merge_cells("B9:D9")
C(ws,8,5,"株式会社スペリオル",b=True,sz=12,bd=False); ws.merge_cells("E8:F8")
for i,val in enumerate(["〒464-0075　名古屋市千種区内山3-18-10","TEL 052-745-6607",
                        "登録番号 T9180001048541","担当　藤原"]):
    C(ws,9+i,5,val,sz=8.5,col=MUT if i<3 else INK,bd=False)
    ws.merge_cells(start_row=9+i,start_column=5,end_row=9+i,end_column=6)
C(ws,14,2,"下記のとおりお見積り申し上げます。",sz=10,bd=False); ws.merge_cells("B14:F14")
C(ws,16,2,"件名",b=True,fill=SUBF,al="center")
C(ws,16,3,"建設業 特別教育 eラーニング配信システム 構築 および 教材制作（6科目）一式",wrap=True)
ws.merge_cells("C16:F16"); ws.row_dimensions[16].height=32
C(ws,17,2,"御見積金額",b=True,fill=SUBF,al="center")
C(ws,17,3,f"={NET_A}",fmt='"¥"#,##0"  （税抜）"',b=True,sz=18,col=NAVY,al="left")
ws.merge_cells("C17:F17"); ws.row_dimensions[17].height=36
C(ws,18,2,"",fill=SUBF)
C(ws,18,3,f'="消費税 ¥"&TEXT({TAX_A},"#,##0")&"　／　税込 ¥"&TEXT({GT_A},"#,##0")',sz=11,col=MUT,al="left")
ws.merge_cells("C18:F18"); ws.row_dimensions[18].height=20
rows=[("納期","システム：ご発注後 約4か月　／　教材：第1次3科目 約4か月、第2次3科目 約7か月"),
      ("納入場所","貴社ご指定のAWS環境"),
      ("検収条件","納品後14日以内に貴社にて検収。期間内にご連絡なき場合は検収完了とみなします。"),
      ("支払条件","ご発注時に全額を一括でお支払いください。本お見積りは当該条件を前提とした金額です。"),
      ("備考","月額費用・AWS利用料・ゼウス利用料・教材の改訂費は含みません（シート05）。前回お見積りからの増減はシート02をご参照ください。")]
r=20
for lbl,val in rows:
    C(ws,r,2,lbl,b=True,fill=SUBF,al="center")
    C(ws,r,3,val,wrap=True,sz=9.5)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=6)
    ws.row_dimensions[r].height=42 if lbl in("納期","備考","支払条件") else 30; r+=1
r+=1
C(ws,r,2,"内訳",b=True,col=NAVY,bd=False,sz=11); r+=1
HDR(ws,r,["","区分","内容","","数量","金額（税抜）"],h=20)
ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4); r+=1
C(ws,r,2,"システム開発",b=True,al="center")
C(ws,r,3,"配信基盤・視聴整合性・本人確認・修了証発行・決済",sz=9.5)
ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
C(ws,r,5,f"={MD_A}",fmt=MD,al="center"); C(ws,r,6,f"={SYS_A}",fmt=YEN,al="right"); r+=1
C(ws,r,2,"教材制作",b=True,al="center")
C(ws,r,3,"特別教育 学科教材 6科目",sz=9.5)
ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
C(ws,r,5,f"={MATCNT}",fmt='0"科目"',al="center"); C(ws,r,6,f"={MATN}",fmt=YEN,al="right"); r+=1
C(ws,r,2,"小計",b=True,al="center",fill=LTF); C(ws,r,3,"（税抜）",sz=9,col=MUT,fill=LTF)
C(ws,r,4,"",fill=LTF); C(ws,r,5,"",fill=LTF)
C(ws,r,6,f"={SUB_A}",fmt=YEN,al="right",b=True,fill=LTF); r+=1
C(ws,r,2,"調整",b=True,al="center"); C(ws,r,3,"全額一括前払いによる調整",sz=9,col=OK)
C(ws,r,4,""); C(ws,r,5,"")
C(ws,r,6,f"={ADJ_A}",fmt=YEN,al="right",b=True,col=OK); r+=1
C(ws,r,2,"差引小計",b=True,al="center"); C(ws,r,3,"（税抜）",sz=9,col=MUT)
C(ws,r,4,""); C(ws,r,5,"")
C(ws,r,6,f"={NET_A}",fmt=YEN,al="right",b=True); r+=1
C(ws,r,2,"消費税",b=True,al="center"); C(ws,r,3,""); C(ws,r,4,""); C(ws,r,5,"")
C(ws,r,6,f"={TAX_A}",fmt=YEN,al="right",b=True); r+=1
C(ws,r,2,"合計",b=True,al="center",fill=SUBF,sz=11); C(ws,r,3,"（税込）",sz=9,col=MUT,fill=SUBF)
C(ws,r,4,"",fill=SUBF); C(ws,r,5,"",fill=SUBF)
C(ws,r,6,f"={GT_A}",fmt=YEN,al="right",b=True,fill=SUBF,sz=11)
ws.row_dimensions[r].height=22; r+=2
C(ws,r,2,"■ 別途ご負担いただく費用（いずれも税抜）",b=True,col=NAVY,bd=False,sz=10); r+=1
for t in ["・月額 保守・運用サービス 60,000円 ＋ 法令改正監視・通知サービス 20,000円　合計 80,000円／月",
 "・AWS 利用料　月額 22,800円程度（推奨構成・受講者数により変動）。聖建様名義でのご契約を推奨します",
 "・株式会社ゼウス　月額 3,000円 ＋ 決済手数料 最大3.5%。聖建様が直接ご契約いただきます",
 "・法令改正にともなう教材の改訂　1科目 300,000円（都度発注）"]:
    C(ws,r,2,t,sz=9.5,bd=False); ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=6); r+=1

# =====================================================================
# 06_注文書
# =====================================================================
ws=wb.create_sheet("06_注文書")
for col,w in zip("ABCDEFG",[3,20,24,14,16,18,3]): ws.column_dimensions[col].width=w
ws["B2"]="注　文　書"; ws["B2"].font=Font(name=F,bold=True,size=22,color=NAVY)
ws.merge_cells("B2:F2"); ws["B2"].alignment=Alignment(horizontal="center"); ws.row_dimensions[2].height=34
C(ws,4,2,"注文番号",sz=9,col=MUT,bd=False); C(ws,4,3,"（ご記入ください）",sz=9,fill=YEL,col="0000FF",bd=False)
C(ws,5,2,"発行日",sz=9,col=MUT,bd=False);   C(ws,5,3,"　　年　　月　　日",sz=9,fill=YEL,col="0000FF",bd=False)
C(ws,7,2,"株式会社スペリオル　御中",b=True,sz=15,bd=False); ws.merge_cells("B7:D7")
C(ws,8,2,"〒464-0075　名古屋市千種区内山3-18-10",sz=9,col=MUT,bd=False); ws.merge_cells("B8:D8")
C(ws,9,2,"TEL 052-745-6607　／　登録番号 T9180001048541",sz=9,col=MUT,bd=False); ws.merge_cells("B9:D9")
C(ws,10,2,"担当　藤原 様",sz=9.5,col=MUT,bd=False)
C(ws,7,5,"発注者",sz=8.5,col=MUT,bd=False)
C(ws,8,5,"株式会社聖建",b=True,sz=13,bd=False); ws.merge_cells("E8:F8")
C(ws,9,5,"〒475-0805　愛知県半田市浜田町1丁目7番地",sz=8.5,col=MUT,bd=False); ws.merge_cells("E9:F9")
C(ws,10,5,"TEL 0569-84-9669",sz=8.5,col=MUT,bd=False); ws.merge_cells("E10:F10")
C(ws,11,5,"ご担当　　　　　　　　　　　　　　印",sz=9,bd=False); ws.merge_cells("E11:F11")
C(ws,13,2,"下記のとおり注文いたします。",sz=10,bd=False); ws.merge_cells("B13:F13")
C(ws,15,2,"件名",b=True,fill=SUBF,al="center")
C(ws,15,3,"建設業 特別教育 eラーニング配信システム 構築 および 教材制作（6科目）一式",wrap=True)
ws.merge_cells("C15:F15"); ws.row_dimensions[15].height=30
C(ws,16,2,"根拠見積",b=True,fill=SUBF,al="center")
C(ws,16,3,"='01_見積書'!C4&\"（\"&'01_見積書'!C5&\" 発行）\"",wrap=True,sz=9.5)
ws.merge_cells("C16:F16")
HDR(ws,18,["","区分","内容","数量","","金額（税抜）"],h=20)
r=19; o0=r
C(ws,r,2,"システム開発",b=True,al="center")
C(ws,r,3,"配信基盤・視聴整合性・本人確認・修了証発行・決済",sz=9,wrap=True)
C(ws,r,4,f"={MD_A}",fmt=MD,al="center"); C(ws,r,5,"")
C(ws,r,6,f"={SYS_A}",fmt=YEN,al="right"); ws.row_dimensions[r].height=28; r+=1
C(ws,r,2,"教材制作",b=True,al="center")
C(ws,r,3,"特別教育 学科教材 6科目",sz=9,wrap=True)
C(ws,r,4,f"={MATCNT}",fmt='0"科目"',al="center"); C(ws,r,5,"")
C(ws,r,6,f"={MATN}",fmt=YEN,al="right"); ws.row_dimensions[r].height=28; r+=1
o1=r-1
C(ws,r,2,"小計（税抜）",b=True,fill=LTF)
for c in (3,4,5): C(ws,r,c,"",fill=LTF)
C(ws,r,6,f"=SUM(F{o0}:F{o1})",fmt=YEN,al="right",b=True,fill=LTF); osub=r; r+=1
C(ws,r,2,"全額一括前払いによる調整",b=True,col=OK); C(ws,r,3,"",sz=9); C(ws,r,4,""); C(ws,r,5,"")
C(ws,r,6,f"={ADJ_A}",fmt=YEN,al="right",col=OK,b=True); oadj=r; r+=1
C(ws,r,2,"消費税",b=True); C(ws,r,3,"",sz=9)
C(ws,r,4,f"={Q(f'D{tax}')}",fmt="0%",al="center"); C(ws,r,5,"")
C(ws,r,6,f"=ROUND((F{osub}+F{oadj})*D{r},0)",fmt=YEN,al="right",b=True); otax=r; r+=1
C(ws,r,2,"ご注文金額（税込）",b=True,fill=SUBF,sz=12)
for c in (3,4,5): C(ws,r,c,"",fill=SUBF)
C(ws,r,6,f"=F{osub}+F{oadj}+F{otax}",fmt=YEN,al="right",b=True,fill=SUBF,sz=13)
ws.row_dimensions[r].height=26; r+=2
for lbl,val in [("納期","システム：ご発注後 約4か月　／　教材：第1次3科目 約4か月、第2次3科目 約7か月"),
                ("納入場所","貴社ご指定のAWS環境"),
                ("検収条件","納品後14日以内に検収。期間内にご連絡なき場合は検収完了とみなします。"),
                ("支払条件","本注文書の発行時に、上記ご注文金額の全額を一括でお支払いいたします。"),
                ("別途契約","月額の保守・運用サービスおよび法令改正監視・通知サービス（合計80,000円／月・税抜）は別途契約とします。AWS利用料および株式会社ゼウスとの契約は発注者が直接行うものとします。")]:
    C(ws,r,2,lbl,b=True,fill=SUBF,al="center")
    C(ws,r,3,val,wrap=True,sz=9.5)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=6)
    ws.row_dimensions[r].height=44 if lbl in("納期","別途契約") else 28; r+=1

# =====================================================================
# 07_前提条件
# =====================================================================
ws=wb.create_sheet("07_前提条件")
ws.column_dimensions["A"].width=3; ws.column_dimensions["B"].width=30; ws.column_dimensions["C"].width=88
ws["B1"]="見積の前提条件"; ws["B1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
BLK=[("■ 1. 本見積に含まれるもの",NAVY,[
 ("システム開発","配信基盤、視聴整合性、登録時の本人確認、受講開始時の顔照合、確認問題、修了判定、修了証の発行、決済（ゼウス連携）、受講者ポータル、運営管理画面の基本機能。"),
 ("教材制作","特別教育 学科教材 6科目。構成台本、スライド、図解、ナレーション、バーチャル講師映像、章末問題・修了確認テストの作成。"),
 ("修了証","修了証PDFの自動発行、再発行、受講者によるダウンロード。記載項目は受講者氏名、科目、法定学科時間、有効視聴時間、テスト得点、講師名、発行者、証明書番号。"),
 ("AWS環境の構築","本番運用に耐える構成で構築します。RDS PostgreSQL は Multi-AZ とし、自動バックアップとPITRを有効にします。"),
]),
("■ 2. 今回の範囲から外したもの",BAD,[
 ("受講中のランダム顔照合","受講の途中で予告なく撮影し、本人が継続して受講しているかを確認する機能。登録時の本人確認と受講開始時の照合は実装するため、替え玉への防御がゼロになるわけではありません。将来の追加実装を想定した設計としておきます。"),
 ("教材 4科目","低圧電気取扱業務、第2種酸素欠乏危険作業、高所作業車、玉掛け。第2次発注でのご検討をお願いします。"),
 ("企業管理ポータル","従業員の一括申込、進捗確認、修了者一覧のCSV出力。"),
 ("その他","運営管理画面の拡張、質疑応答機能、修了証の真正性検証ページ（QRコード）。"),
]),
("■ 3. 本見積に含まれないもの（別途）",BAD,[
 ("月額サービス","保守・運用サービス 60,000円／月、法令改正監視・通知サービス 20,000円／月。合計80,000円／月（税抜）。別途契約とします。"),
 ("AWS 利用料","聖建様のご負担とします。推奨構成で月額22,800円程度（税抜）。聖建様名義でのご契約を推奨します。"),
 ("ゼウス利用料","初期費用無料、月額3,000円、決済手数料 最大3.5%。聖建様が直接ご契約いただきます。最新の条件はゼウス社へご確認ください。"),
 ("教材の改訂・追加","1科目300,000円（都度発注）。"),
 ("教材の監修","内容の監修は聖建様の社内有資格者が担われる前提です。"),
 ("実技教育","実技はオンラインで代替できません。受講者の所属事業者が実施する前提です。"),
]),
("■ 4. 責任の分担",NAVY,[
 ("弊社が保証すること","教材が安全衛生特別教育規程の定める科目・範囲・時間を満たすこと（監修者名を明示）。受講記録を証明可能な形で保存・出力できること。"),
 ("弊社が保証しないこと","受講者の所属事業者が負う労働安全衛生法第59条第3項の実施義務が果たされたこと。労働災害が発生しないこと。行政機関が個別の事案において本教材を適法と判断すること。"),
 ("法令改正への対応","弊社は官報・通達を監視し、該当しうる改正を自動でメール通知します。その改正が教材のどこに影響するかの判断は、聖建様の監修者が行うものとします。"),
]),
("■ 5. 金額の前提",NAVY,[
 ("人日単価","50,000円／人日（シート03のC4で変更可能）。"),
 ("教材制作単価","300,000円／科目。聖建様ご指定の単価です。改訂・追加時も同額を適用します。"),
 ("支払条件","ご発注時に全額を一括でお支払いいただくことを前提とした金額です。分割払いの場合は金額が変わります。"),
 ("有効期限","発行日より30日間。"),
])]
r=3
for head,col,items in BLK:
    C(ws,r,2,head,b=True,sz=11.5,col=col,fill=SUBF,bd=False); C(ws,r,3,"",fill=SUBF,bd=False)
    ws.row_dimensions[r].height=20; r+=1
    for lbl,txt in items:
        C(ws,r,2,lbl,b=True,sz=9.5,va="top",wrap=True)
        C(ws,r,3,txt,sz=9.5,wrap=True,va="top")
        ws.row_dimensions[r].height=44 if len(txt)>95 else (34 if len(txt)>60 else 22); r+=1
    r+=1

if "Sheet" in wb.sheetnames: del wb["Sheet"]
wb._sheets=[wb[n] for n in ["01_見積書","02_増減内訳","03_見積明細","04_教材科目","05_月額・別途費用","06_注文書","07_前提条件"]]

# ---- 印刷設定（1ページに収める） ----
PAGE={
 "01_見積書":      ("A1:F41","portrait"),
 "02_増減内訳":    ("A1:E41","portrait"),
 "03_見積明細":    ("A1:F44","landscape"),
 "04_教材科目":    ("A1:F24","landscape"),
 "05_月額・別途費用":("A1:D30","landscape"),
 "06_注文書":      ("A1:F31","portrait"),
 "07_前提条件":    ("A1:C33","portrait"),
}
for name,(area,orient) in PAGE.items():
    w=wb[name]
    w.print_area=area
    w.page_setup.orientation=orient
    w.page_setup.paperSize=w.PAPERSIZE_A4
    w.page_setup.fitToWidth=1
    w.page_setup.fitToHeight=1
    w.sheet_properties.pageSetUpPr.fitToPage=True
    w.page_margins.left=w.page_margins.right=0.4
    w.page_margins.top=w.page_margins.bottom=0.4
    w.page_margins.header=w.page_margins.footer=0.2
    w.print_options.horizontalCentered=True

wb.save("/home/user/-/quote/聖建様_概算見積書・注文書.xlsx")
print("saved")
