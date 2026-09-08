# -*- coding: utf-8 -*-
"""株式会社聖建様向け 概算見積書・注文書 rev.2（総額800万円版）"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

F="Meiryo"; NAVY="1B3A57"; ACC="B05A1E"; LIGHT="F2F5F7"; SUB="DCE6EE"
INK="1A2430"; MUT="6C7A87"; LINE="C9D4DC"; OK="1E6B4A"; BAD="9E2F27"
YEL=PatternFill("solid",fgColor="FFF6D9"); HDRF=PatternFill("solid",fgColor=NAVY)
SUBF=PatternFill("solid",fgColor=SUB); LTF=PatternFill("solid",fgColor=LIGHT)
thin=Side(style="thin",color=LINE); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
YEN='¥#,##0;(¥#,##0);-'; MD='#,##0.0"人日";;-'

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
# 03_教材科目
# =====================================================================
ws=wb.create_sheet("03_教材科目")
ws["A1"]="eラーニング教材 制作科目一覧（10科目）"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="出典：建設業向けeラーニング 商品化優先順位。学科のみを対象とし、実技は本見積の範囲外です。"
ws["A2"].font=Font(name=F,size=9,color=MUT)
C(ws,4,1,"制作単価（1科目あたり）",b=True,fill=SUBF); C(ws,4,2,"",fill=SUBF)
u=C(ws,4,3,300000,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
u.comment=Comment("聖建様ご指定の単価。構成台本、スライド、図解、ナレーション、バーチャル講師映像、章末問題・修了確認テストの作成まで。撮影は行いません。法令改正時の改訂も同額を適用します。","見積前提")
C(ws,4,4,"円／科目",col=MUT,sz=9); ws.merge_cells(start_row=4,start_column=1,end_row=4,end_column=2)
HDR(ws,6,["優先\n順位","特別教育の名称","学科\n(時間)","想定ターゲット","納品\n時期","制作費"],
    widths=[7,34,8,26,9,14],h=30)
SUBJ=[(1,"フルハーネス型墜落制止用器具",4.5,"建築・鉄骨・足場・設備",2),
(2,"足場の組立て・解体・変更",6.0,"鳶・足場・建築・改修",1),
(3,"自由研削用といし",4.0,"鉄工・設備・解体・内装",3),
(4,"低圧電気取扱業務",7.0,"電気・設備・保守",3),
(5,"アーク溶接等",11.0,"鉄骨・配管・設備・橋梁",1),
(6,"石綿使用建築物等の解体・破砕等",4.5,"解体・改修・リフォーム",1),
(7,"特定粉じん作業",4.5,"解体・はつり・土木",2),
(8,"第2種酸素欠乏危険作業",5.5,"下水・槽・地下設備",2),
(9,"高所作業車（10m未満）",6.0,"設備・電気・外装",2),
(10,"玉掛け（1t未満）",5.0,"鉄骨・設備・土木",3)]
LOT={1:"第1次\n（〜4か月）",2:"第2次\n（〜7か月）",3:"第3次\n（〜10か月）"}
r=7; first=r
for no,name,h,tgt,lot in SUBJ:
    C(ws,r,1,no,al="center"); C(ws,r,2,name,wrap=True)
    C(ws,r,3,h,fmt='0.0"h"',al="center"); C(ws,r,4,tgt,sz=9,col=MUT,wrap=True)
    C(ws,r,5,lot,al="center",b=True,fill=YEL,col="0000FF",sz=8.5,wrap=True)
    C(ws,r,6,"=$C$4",fmt=YEN,al="right")
    ws.row_dimensions[r].height=26; r+=1
last=r-1
C(ws,r,1,"",fill=SUBF); C(ws,r,2,"合計（10科目）",b=True,fill=SUBF)
C(ws,r,3,f"=SUM(C{first}:C{last})",fmt='0.0"h"',al="center",b=True,fill=SUBF)
C(ws,r,4,"",fill=SUBF); C(ws,r,5,"",fill=SUBF)
C(ws,r,6,f"=SUM(F{first}:F{last})",fmt=YEN,al="right",b=True,fill=SUBF)
MAT=f"'03_教材科目'!F{r}"; MATN=f"'03_教材科目'!$F${r}"; r+=2
for t in ["■ 納品時期について",
 "・全10科目を同時には制作できません。3回に分けて納品します（第1次3科目、第2次4科目、第3次3科目）。",
 "・第1次は実技のない科目（足場・石綿）と既存資産のあるアーク溶接を優先し、最短で販売を開始できるようにします。",
 "・E列で納品時期を変更できます。",
 "",
 "■ 法令改正時の改訂",
 "・改正により教材の作り直しが必要になった場合は、1科目あたり300,000円（新規制作と同額）で改訂します。",
 "・改正の検知と通知は、月額の「法令改正監視・通知サービス」で行います（シート04）。",
 "・改正が教材のどこに影響するかの判断は、聖建様の監修者にお願いします。"]:
    C(ws,r,1,t,bd=False,sz=9.5,b=t.startswith("■"),col=NAVY if t.startswith("■") else MUT); r+=1
ws.freeze_panes="A7"

# =====================================================================
# 02_見積明細
# =====================================================================
ws=wb.create_sheet("02_見積明細")
ws["A1"]="見積明細"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="黄色のセルが編集箇所です。人日単価と各項目の人日を変えると、見積書と注文書まで自動で再計算されます。"
ws["A2"].font=Font(name=F,size=9,color=MUT)
C(ws,4,1,"人日単価",b=True,fill=SUBF); C(ws,4,2,"",fill=SUBF)
ur=C(ws,4,3,50000,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
ur.comment=Comment("設計・実装・テストを含む標準人日単価。御社の基準単価に差し替えてください。","見積前提")
C(ws,4,4,"円／人日",col=MUT,sz=9); ws.merge_cells(start_row=4,start_column=1,end_row=4,end_column=2)
RATE="$C$4"
HDR(ws,6,["区分","項目","内容","数量","単価","金額"],widths=[12,28,54,9,12,14],h=22)
SYS=[("要件定義・基本設計","業務フローの確定、データモデル設計、API設計、画面設計、利用規約・特定商取引法表記・個人情報の取扱い方針の整備支援",8),
("AWS本番環境の構築","VPC、EC2冗長構成、RDS PostgreSQL（Multi-AZ・自動バックアップ・PITR）、ElastiCache、ALB、S3、CloudFront、CloudWatch、AWS Backup の構築とCI/CD",7),
("動画配信基盤","HLS変換（FFmpeg）、短寿命の署名付きマニフェスト配信、章単位の受講制御、MP4直リンクの排除",12),
("視聴整合性・実視聴時間の集計","heartbeatの受信と検証、未視聴位置へのシーク制限、倍速検知、可視状態の判定、サーバー側での検証済み秒数の集計、異常検知",15),
("本人確認（登録時）","その場撮影UI（getUserMedia）、ライブネス判定、本人確認書類との突合、運営の審査画面",11),
("PC↔スマートフォン引き継ぎ","QRコードとワンタイムURLによる撮影工程の引き継ぎ。PCで受講される方の離脱を防ぎます",3),
("顔照合（受講開始時・受講中）","1:1照合APIの連携、3段階判定（合格／保留／失敗）、ランダム発火スケジューラ、判定ログの保存",11),
("確認問題・自動採点","章末問題、修了確認テスト、出題のランダム化、不合格後の再受講制御",7),
("修了判定・修了証発行","サーバー側での修了条件の再検証、修了記録の作成、修了証PDFの発行・再発行",7),
("受講者ポータル","申込、マイページ、受講画面、進捗表示、スマートフォン対応",9),
("運営管理画面（基本機能）","本人確認の承認キュー、顔照合アラートの確認、受講状況の一覧、受講ログの照会",4),
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
C(ws,r,3,"10科目・学科58時間分。構成台本の書き起こし、スライド、図解、ナレーション、バーチャル講師映像、章末問題・修了確認テストの作成。3回に分けて納品",wrap=True,sz=9,col=MUT)
C(ws,r,4,10,fmt='0"科目"',al="center",col=OK)
C(ws,r,5,"='03_教材科目'!$C$4",fmt=YEN,al="right",col=OK)
C(ws,r,6,f"={MATN}",fmt=YEN,al="right")
ws.row_dimensions[r].height=36; matrow=r; r+=1
C(ws,r,1,"",fill=LTF); C(ws,r,2,"教材制作 小計",b=True,fill=LTF); C(ws,r,3,"",fill=LTF)
C(ws,r,4,"",fill=LTF); C(ws,r,5,"",fill=LTF)
C(ws,r,6,f"=F{matrow}",fmt=YEN,al="right",b=True,fill=LTF)
matsub=r; r+=2
tr=r
C(ws,tr,1,"",fill=SUBF); C(ws,tr,2,"小計（税抜）",b=True,fill=SUBF,sz=11); C(ws,tr,3,"",fill=SUBF)
C(ws,tr,4,f"=D{sysrow}",fmt=MD,al="center",b=True,fill=SUBF); C(ws,tr,5,"",fill=SUBF)
C(ws,tr,6,f"=F{sysrow}+F{matsub}",fmt=YEN,al="right",b=True,fill=SUBF,sz=11)
adj=tr+1
C(ws,adj,2,"調整額（お値引き等。マイナスで入力）",b=True)
C(ws,adj,3,"総額を調整される場合はここに入力してください",sz=9,col=MUT)
C(ws,adj,4,""); C(ws,adj,5,"")
C(ws,adj,6,0,fmt=YEN,al="right",fill=YEL,col="0000FF",b=True)
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
Q=lambda a:f"'02_見積明細'!{a}"
NET_A,TAX_A,GT_A,SUB_A=Q(f"F{net}"),Q(f"F{tax}"),Q(f"F{gt}"),Q(f"F{tr}")
SYS_A,MAT_A,MD_A=Q(f"F{sysrow}"),Q(f"F{matsub}"),Q(f"D{sysrow}")
r=gt+2
for t in ["■ 本見積に含まれないもの（別途）",
 "・月額の保守・運用サービス、法令改正監視・通知サービス（シート04）",
 "・AWSの利用料（聖建様のご負担。聖建様名義でのご契約を推奨します）",
 "・法令改正にともなう教材の改訂（1科目300,000円・都度発注）",
 "",
 "■ 本見積に含まれない機能（第2次発注のご相談）",
 "・企業管理ポータル（従業員の一括申込、進捗確認、修了者一覧のCSV出力）",
 "・オンライン決済（当面は申込フォーム＋銀行振込での運用を想定）",
 "・運営管理画面の拡張（監査レポート出力、組織管理、売上・入金管理）",
 "・受講画面からの質疑応答機能（当面はメールで対応）",
 "・修了証の真正性検証ページ（QRコード）",
 "・教材の版管理の強化と多科目への横展開の自動化",
 "",
 "■ 編集できる箇所（黄色のセル）",
 "・C4 人日単価　・D列 各項目の人日　・調整額　・消費税率　・シート03の教材単価と納品時期"]:
    C(ws,r,1,t,bd=False,sz=9.5,b=t.startswith("■"),col=NAVY if t.startswith("■") else MUT); r+=1
ws.freeze_panes="A7"

# =====================================================================
# 04_月額・別途費用
# =====================================================================
ws=wb.create_sheet("04_月額・別途費用")
ws["A1"]="月額費用と別途費用"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="初期費用（シート01）とは別に、下記が発生します。"
ws["A2"].font=Font(name=F,size=9,color=MUT)
HDR(ws,4,["区分","項目","月額／単価","内容"],widths=[16,32,16,66],h=22)
r=5
C(ws,r,1,"弊社へ\nお支払い",al="center",b=True,wrap=True,fill=LTF)
C(ws,r,2,"保守・運用サービス",b=True)
C(ws,r,3,80000,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
C(ws,r,4,"システムの死活監視とエラー監視、障害対応、軽微な改修、受講者からの問い合わせ一次対応、月次レポートの提出",wrap=True,sz=9.5)
ws.row_dimensions[r].height=32; m0=r; r+=1
C(ws,r,1,"",fill=LTF)
C(ws,r,2,"法令改正監視・通知サービス",b=True)
C(ws,r,3,20000,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
C(ws,r,4,"官報・厚生労働省の通達・パブリックコメントを継続監視し、対象10科目に影響しうる改正を検知したら、聖建様と弊社の双方へ自動でメール通知。月次で「該当なし」も含めて報告",wrap=True,sz=9.5)
ws.row_dimensions[r].height=42; m1=r; r+=1
C(ws,r,1,"",fill=SUBF); C(ws,r,2,"月額 小計（税抜）",b=True,fill=SUBF)
C(ws,r,3,f"=SUM(C{m0}:C{m1})",fmt=YEN,al="right",b=True,fill=SUBF,sz=11)
C(ws,r,4,"",fill=SUBF); MON=f"'04_月額・別途費用'!C{r}"; r+=2

C(ws,r,1,"お客様\nご負担",al="center",b=True,wrap=True,fill=LTF)
C(ws,r,2,"AWS 利用料（本番構成）",b=True)
C(ws,r,3,36000,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
C(ws,r,4,"EC2冗長構成、RDS PostgreSQL（Multi-AZ）、ElastiCache、ALB、S3、CloudFront、CloudWatch、AWS Backup。聖建様名義でのご契約を推奨します。受講者数により変動します",wrap=True,sz=9.5)
ws.row_dimensions[r].height=42; aws_r=r; r+=1
C(ws,r,1,"",fill=LTF); C(ws,r,2,"従量課金（顔照合・配信ほか）",b=True)
C(ws,r,3,"実費",al="right",b=True,col=MUT)
C(ws,r,4,"顔照合API、ライブネス判定、CDN転送、メール送信。受講者100名／月で約830円、300名／月で約2,550円（AWS利用料に含まれます）",wrap=True,sz=9.5)
ws.row_dimensions[r].height=32; r+=1
C(ws,r,1,"",fill=LTF); C(ws,r,2,"決済手数料",b=True)
C(ws,r,3,"売上の3.6%",al="right",b=True,col=MUT)
C(ws,r,4,"オンライン決済を導入された場合。第2次発注の対象です",wrap=True,sz=9.5)
r+=2

C(ws,r,1,"都度\n発注",al="center",b=True,wrap=True,fill=LTF)
C(ws,r,2,"教材の改訂（法令改正対応）",b=True)
C(ws,r,3,"='03_教材科目'!C4",fmt=YEN,al="right",b=True,col=OK)
C(ws,r,4,"1科目あたり。新規制作と同額です。台本の修正、該当部分の動画の再生成と差し替え、教材の版の更新を含みます",wrap=True,sz=9.5)
ws.row_dimensions[r].height=32; r+=2

C(ws,r,1,"参考：月額の合計（聖建様のご負担総額）",b=True,col=NAVY,bd=False,sz=11); r+=1
C(ws,r,2,"弊社サービス ＋ AWS 利用料",b=True)
C(ws,r,3,f"={MON}+C{aws_r}",fmt=YEN,al="right",b=True,sz=12,col=ACC)
C(ws,r,4,"税抜。受講者数により従量分が変動します",sz=9,col=MUT)
r+=2

NOTE=[("■ 法令改正の自動通知について",NAVY),
("・監視の仕組みは弊社が共通サービスとして保有し、聖建様専用に開発するものではありません。",MUT),
("　そのため初期費用は発生せず、月額サービスとしてご提供します。",MUT),
("・検知した場合、聖建様のご担当者と弊社の双方へ同時に自動メールが送信されます。",MUT),
("・通知には、改正の名称、公布・施行日、該当しうる科目、参照URLを記載します。",MUT),
("",None),
("■ 責任の分担",NAVY),
("・官報・通達の機械的な監視と通知　　　　　→　弊社",MUT),
("・その改正が教材のどこに影響するかの判断　→　聖建様の監修者",BAD),
("・改訂作業の実施　　　　　　　　　　　　　→　弊社（都度発注）",MUT),
("・改訂後の内容の承認　　　　　　　　　　　→　聖建様の監修者",MUT),
("　弊社が法令を解釈して改訂の要否を判断することはできません。監修者の役割です。",BAD),
("",None),
("■ AWSをお客様のご負担とする理由",NAVY),
("・受講者が増えると転送量・顔照合の回数が増えます。定額に含めると、事業が成功するほど",MUT),
("　弊社が損をする構造になり、いずれ値上げのお願いが必要になります。",MUT),
("・聖建様名義でご契約いただければ、費用が明細で見え、原価管理ができます。",MUT),
("・万一お取引が終了しても、システムは聖建様の資産としてそのまま継続できます。",OK),
("・アカウント管理が難しい場合は、弊社が代行契約し、実費＋事務手数料10%でご請求する形も可能です。",MUT)]
for t,col in NOTE:
    C(ws,r,1,t,bd=False,sz=9.5,b=(col in (NAVY,)),col=col or MUT); r+=1

# =====================================================================
# 01_見積書
# =====================================================================
ws=wb.create_sheet("01_見積書",0)
for col,w in zip("ABCDEFG",[3,20,20,16,16,18,3]): ws.column_dimensions[col].width=w
ws["B2"]="御　見　積　書"; ws["B2"].font=Font(name=F,bold=True,size=22,color=NAVY)
ws.merge_cells("B2:F2"); ws["B2"].alignment=Alignment(horizontal="center"); ws.row_dimensions[2].height=34
C(ws,4,2,"見積番号",sz=9,col=MUT,bd=False); C(ws,4,3,"SP-2026-0001",sz=9,fill=YEL,col="0000FF",bd=False)
C(ws,5,2,"発行日",sz=9,col=MUT,bd=False);   C(ws,5,3,"2026年9月8日",sz=9,fill=YEL,col="0000FF",bd=False)
C(ws,6,2,"有効期限",sz=9,col=MUT,bd=False); C(ws,6,3,"発行日より60日間",sz=9,fill=YEL,col="0000FF",bd=False)
C(ws,8,2,"株式会社聖建　御中",b=True,sz=15,bd=False); ws.merge_cells("B8:D8")
C(ws,9,2,"〒475-0805　愛知県半田市浜田町1丁目7番地",sz=9,col=MUT,bd=False,fill=YEL); ws.merge_cells("B9:D9")
C(ws,10,2,"営業部　津隈 佑己 様",sz=9.5,col=MUT,bd=False); ws.merge_cells("B10:D10")
C(ws,8,5,"株式会社スペリオル",b=True,sz=12,bd=False); ws.merge_cells("E8:F8")
for i,(lbl,val) in enumerate([("所在地","〒　　　　　（ご記入ください）"),("TEL / FAX","（ご記入ください）"),
                              ("登録番号（適格請求書）","T　　　　　　　　　　　"),("担当","藤原")]):
    C(ws,9+i,5,lbl,sz=8.5,col=MUT,bd=False)
    C(ws,9+i,6,val,sz=8.5,bd=False,fill=YEL if i<3 else None,col="0000FF" if i<3 else INK)
C(ws,14,2,"下記のとおりお見積り申し上げます。",sz=10,bd=False); ws.merge_cells("B14:F14")
C(ws,16,2,"件名",b=True,fill=SUBF,al="center")
C(ws,16,3,"建設業 特別教育 eラーニング配信システム 構築 および 教材制作（10科目）一式",wrap=True)
ws.merge_cells("C16:F16"); ws.row_dimensions[16].height=32
C(ws,17,2,"御見積金額",b=True,fill=SUBF,al="center")
C(ws,17,3,f"={NET_A}",fmt='"¥"#,##0"  （税抜）"',b=True,sz=18,col=NAVY,al="left")
ws.merge_cells("C17:F17"); ws.row_dimensions[17].height=36
C(ws,18,2,"",fill=SUBF)
C(ws,18,3,f'="消費税 ¥"&TEXT({TAX_A},"#,##0")&"　／　税込 ¥"&TEXT({GT_A},"#,##0")',sz=11,col=MUT,al="left")
ws.merge_cells("C18:F18"); ws.row_dimensions[18].height=20
rows=[("納期","システム：ご発注後 約4か月　／　教材：第1次3科目 約4か月、第2次4科目 約7か月、第3次3科目 約10か月"),
      ("納入場所","貴社ご指定のAWS環境"),
      ("検収条件","納品後14日以内に貴社にて検収。期間内にご連絡なき場合は検収完了とみなします。"),
      ("支払条件","着手時40%・システム検収時30%・全教材納品時30%。検収月末締め翌月末日払い。"),
      ("備考","本書は概算見積です。月額費用・AWS利用料・教材の改訂費は含みません（シート04）。")]
r=20
for lbl,val in rows:
    C(ws,r,2,lbl,b=True,fill=SUBF,al="center")
    C(ws,r,3,val,wrap=True,sz=9.5)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=6)
    ws.row_dimensions[r].height=42 if lbl in("納期","備考") else 30; r+=1
r+=1
C(ws,r,2,"内訳",b=True,col=NAVY,bd=False,sz=11); r+=1
HDR(ws,r,["","区分","内容","","数量","金額（税抜）"],h=20)
ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4); r+=1
C(ws,r,2,"システム開発",b=True,al="center")
C(ws,r,3,"配信基盤・視聴整合性・本人確認・修了判定",sz=9.5)
ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
C(ws,r,5,f"={MD_A}",fmt=MD,al="center"); C(ws,r,6,f"={SYS_A}",fmt=YEN,al="right"); r+=1
C(ws,r,2,"教材制作",b=True,al="center")
C(ws,r,3,"特別教育 学科教材 10科目（学科58時間分）",sz=9.5)
ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
C(ws,r,5,10,fmt='0"科目"',al="center"); C(ws,r,6,f"={MAT_A}",fmt=YEN,al="right"); r+=1
C(ws,r,2,"小計",b=True,al="center",fill=LTF); C(ws,r,3,"（税抜）",sz=9,col=MUT,fill=LTF)
C(ws,r,4,"",fill=LTF); C(ws,r,5,"",fill=LTF)
C(ws,r,6,f"={SUB_A}",fmt=YEN,al="right",b=True,fill=LTF); r+=1
C(ws,r,2,"差引小計",b=True,al="center"); C(ws,r,3,"調整額を反映",sz=9,col=MUT)
C(ws,r,4,""); C(ws,r,5,"")
C(ws,r,6,f"={NET_A}",fmt=YEN,al="right",b=True); r+=1
C(ws,r,2,"消費税",b=True,al="center"); C(ws,r,3,""); C(ws,r,4,""); C(ws,r,5,"")
C(ws,r,6,f"={TAX_A}",fmt=YEN,al="right",b=True); r+=1
C(ws,r,2,"合計",b=True,al="center",fill=SUBF,sz=11); C(ws,r,3,"（税込）",sz=9,col=MUT,fill=SUBF)
C(ws,r,4,"",fill=SUBF); C(ws,r,5,"",fill=SUBF)
C(ws,r,6,f"={GT_A}",fmt=YEN,al="right",b=True,fill=SUBF,sz=11)
ws.row_dimensions[r].height=22; r+=2
C(ws,r,2,"■ 別途ご負担いただく費用",b=True,col=NAVY,bd=False,sz=10); r+=1
C(ws,r,2,f"・月額 保守・運用サービス および 法令改正監視・通知サービス　合計 100,000円／月（税抜）",sz=9.5,bd=False)
ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=6); r+=1
C(ws,r,2,"・AWS 利用料　月額 36,000円程度（税抜・本番構成・受講者数により変動）。聖建様名義でのご契約を推奨します",sz=9.5,bd=False)
ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=6); r+=1
C(ws,r,2,"・法令改正にともなう教材の改訂　1科目 300,000円（都度発注）",sz=9.5,bd=False)
ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=6)

# =====================================================================
# 05_注文書
# =====================================================================
ws=wb.create_sheet("05_注文書")
for col,w in zip("ABCDEFG",[3,20,24,14,16,18,3]): ws.column_dimensions[col].width=w
ws["B2"]="注　文　書"; ws["B2"].font=Font(name=F,bold=True,size=22,color=NAVY)
ws.merge_cells("B2:F2"); ws["B2"].alignment=Alignment(horizontal="center"); ws.row_dimensions[2].height=34
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
HDR(ws,18,["","区分","内容","数量","","金額（税抜）"],h=20)
ws.merge_cells(start_row=18,start_column=5,end_row=18,end_column=5)
r=19; o0=r
C(ws,r,2,"システム開発",b=True,al="center")
C(ws,r,3,"配信基盤・視聴整合性・本人確認・修了判定・受講者ポータル",sz=9,wrap=True)
C(ws,r,4,f"={MD_A}",fmt=MD,al="center"); C(ws,r,5,"")
C(ws,r,6,f"={SYS_A}",fmt=YEN,al="right"); ws.row_dimensions[r].height=28; r+=1
C(ws,r,2,"教材制作",b=True,al="center")
C(ws,r,3,"特別教育 学科教材 10科目（3回に分けて納品）",sz=9.5,wrap=True)
C(ws,r,4,10,fmt='0"科目"',al="center"); C(ws,r,5,"")
C(ws,r,6,f"={MAT_A}",fmt=YEN,al="right"); ws.row_dimensions[r].height=28; r+=1
o1=r-1
C(ws,r,2,"小計（税抜）",b=True,fill=LTF); C(ws,r,3,"",fill=LTF); C(ws,r,4,"",fill=LTF); C(ws,r,5,"",fill=LTF)
C(ws,r,6,f"=SUM(F{o0}:F{o1})",fmt=YEN,al="right",b=True,fill=LTF); osub=r; r+=1
C(ws,r,2,"調整額"); C(ws,r,3,"",sz=9); C(ws,r,4,""); C(ws,r,5,"")
C(ws,r,6,f"={Q(f'F{adj}')}",fmt=YEN,al="right"); oadj=r; r+=1
C(ws,r,2,"消費税",b=True); C(ws,r,3,"",sz=9)
C(ws,r,4,f"={Q(f'D{tax}')}",fmt="0%",al="center"); C(ws,r,5,"")
C(ws,r,6,f"=ROUND((F{osub}+F{oadj})*D{r},0)",fmt=YEN,al="right",b=True); otax=r; r+=1
C(ws,r,2,"ご注文金額（税込）",b=True,fill=SUBF,sz=12)
for c in (3,4,5): C(ws,r,c,"",fill=SUBF)
C(ws,r,6,f"=F{osub}+F{oadj}+F{otax}",fmt=YEN,al="right",b=True,fill=SUBF,sz=13)
ws.row_dimensions[r].height=26; r+=2
for lbl,val in [("納期","システム：ご発注後 約4か月　／　教材：第1次 約4か月、第2次 約7か月、第3次 約10か月"),
                ("納入場所","貴社ご指定のAWS環境"),
                ("検収条件","納品後14日以内に検収。期間内にご連絡なき場合は検収完了とみなします。"),
                ("支払条件","着手時40%・システム検収時30%・全教材納品時30%。検収月末締め翌月末日払い。"),
                ("別途契約","月額の保守・運用サービスおよび法令改正監視・通知サービス（合計100,000円／月・税抜）は別途契約とします。AWS利用料は発注者のご負担とします。")]:
    C(ws,r,2,lbl,b=True,fill=SUBF,al="center")
    C(ws,r,3,val,wrap=True,sz=9.5)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=6)
    ws.row_dimensions[r].height=40 if lbl in("納期","別途契約") else 28; r+=1

# =====================================================================
# 06_前提条件
# =====================================================================
ws=wb.create_sheet("06_前提条件")
ws.column_dimensions["A"].width=3; ws.column_dimensions["B"].width=30; ws.column_dimensions["C"].width=88
ws["B1"]="見積の前提条件"; ws["B1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
BLK=[("■ 1. 本見積に含まれるもの",NAVY,[
 ("設計書の範囲","次の2つの設計書に記載された機能のうち、シート02に明示した範囲の設計・実装・テスト・受入支援。"),
 ("","（1）特別教育 eラーニング 本人確認・受講確認システム 詳細設計書 rev.1（2026年9月5日）"),
 ("","（2）アーク溶接等特別教育 配信システム基本設計書 1.1（2026年9月5日）"),
 ("教材制作","特別教育 学科教材 10科目（学科58時間分）。3回に分けて納品します。"),
 ("AWS環境の構築","本番運用に耐える構成で構築します。RDS PostgreSQL は Multi-AZ とし、自動バックアップとPITRを有効にします。"),
]),
("■ 2. 本見積に含まれないもの",BAD,[
 ("月額サービス","保守・運用サービス 80,000円／月、法令改正監視・通知サービス 20,000円／月（いずれも税抜）。別途契約とします。"),
 ("AWS 利用料","聖建様のご負担とします。本番構成で月額36,000円程度（税抜）を見込みます。聖建様名義でのご契約を推奨します。"),
 ("教材の改訂","法令改正にともなう改訂は1科目300,000円（新規制作と同額）。都度ご発注ください。"),
 ("第2次発注の対象機能","企業管理ポータル、オンライン決済、運営管理画面の拡張（監査レポート出力等）、質疑応答機能、修了証の真正性検証ページ、教材の版管理の強化。"),
 ("教材の監修","内容の監修は聖建様の社内有資格者が担われる前提です。外部監修が必要な場合は別途となります。"),
 ("実技教育","実技はオンラインで代替できません。受講者の所属事業者が実施する前提です。"),
 ("集客・販売促進","ランディングページ、SEO、Web広告の運用。"),
]),
("■ 3. 責任の分担",NAVY,[
 ("弊社が保証すること","教材が安全衛生特別教育規程の定める科目・範囲・時間を満たすこと（監修者名を明示）。受講記録を証明可能な形で保存・出力できること。"),
 ("弊社が保証しないこと","受講者の所属事業者が負う労働安全衛生法第59条第3項の実施義務が果たされたこと。労働災害が発生しないこと。行政機関が個別の事案において本教材を適法と判断すること。"),
 ("法令改正への対応","弊社は官報・通達を機械的に監視し、該当しうる改正を自動でメール通知します。その改正が教材のどこに影響するかの判断は、聖建様の監修者が行うものとします。"),
 ("事前相談","所轄の労働基準監督署への事前相談を、章立ての設計が固まった段階で共同で実施することを推奨します（費用は発生しません）。"),
]),
("■ 4. 金額の前提",NAVY,[
 ("人日単価","50,000円／人日（シート02のC4で変更可能）。"),
 ("教材制作単価","300,000円／科目。聖建様ご指定の単価です。改訂時も同額を適用します。"),
 ("概算である旨","本書は要件確定前の概算見積です。要件定義の完了後、正式なお見積りを提出いたします。"),
 ("有効期限","発行日より60日間。"),
])]
r=3
for head,col,items in BLK:
    C(ws,r,2,head,b=True,sz=11.5,col=col,fill=SUBF,bd=False); C(ws,r,3,"",fill=SUBF,bd=False)
    ws.row_dimensions[r].height=20; r+=1
    for lbl,txt in items:
        C(ws,r,2,lbl,b=True,sz=9.5,va="top",wrap=True)
        C(ws,r,3,txt,sz=9.5,wrap=True,va="top")
        ws.row_dimensions[r].height=34 if len(txt)>75 else 22; r+=1
    r+=1

if "Sheet" in wb.sheetnames: del wb["Sheet"]
wb._sheets=[wb[n] for n in ["01_見積書","02_見積明細","03_教材科目","04_月額・別途費用","05_注文書","06_前提条件"]]
wb.save("/home/user/-/quote/聖建様_概算見積書・注文書.xlsx")
print("saved")
