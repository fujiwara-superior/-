# -*- coding: utf-8 -*-
"""聖建様向け ランニングコスト明細 rev.2（AWS本番構成・お客様負担／法令改正監視あり）"""
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
YEN='¥#,##0;(¥#,##0);-'

def C(ws,r,c,v,*,b=False,fmt=None,fill=None,al=None,col=INK,sz=10,wrap=False,bd=True,va="center"):
    x=ws.cell(row=r,column=c,value=v)
    x.font=Font(name=F,bold=b,size=sz,color=col)
    if fmt:x.number_format=fmt
    if fill:x.fill=fill
    if bd:x.border=BOX
    x.alignment=Alignment(horizontal=al or "left",vertical=va,wrap_text=wrap)
    return x
def HDR(ws,row,vals,widths=None,h=24):
    for i,v in enumerate(vals,1):
        x=ws.cell(row=row,column=i,value=v)
        x.font=Font(name=F,bold=True,size=9.5,color="FFFFFF"); x.fill=HDRF; x.border=BOX
        x.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.row_dimensions[row].height=h
    if widths:
        for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w

wb=openpyxl.Workbook()

# =====================================================================
# 02_単価と前提
# =====================================================================
ws=wb.create_sheet("02_単価と前提")
ws["A1"]="ランニングコストの単価と前提"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="黄色のセルが入力欄です。金額は2026年9月時点の公開情報に基づく想定値であり、確定値ではありません。"
ws["A2"].font=Font(name=F,size=9,color=MUT)
ws.column_dimensions["A"].width=34; ws.column_dimensions["B"].width=15
for c,w in zip("CDE",[34,20,20]): ws.column_dimensions[c].width=w

C(ws,4,1,"共通の前提",b=True,fill=SUBF,col=NAVY)
for c in range(2,6): C(ws,4,c,"",fill=SUBF)
ASM=[("為替レート",150,'¥#,##0"／USD"',"AWS等のドル建て費用の換算に使用します"),
("1受講あたりの動画転送量",3.5,'0.0" GB"',"学科4.5時間・720p（約1.5Mbps）＋再視聴分"),
("1受講あたりの顔照合回数",22,'0" 回"',"受講開始時＋受講中のランダム照合"),
("1受講あたりのライブネス判定",1,'0" 回"',"アカウント登録時の1回のみ"),
("受講料（参考）",10000,YEN,"売上と決済手数料の試算に使用"),
("決済手数料率",0.036,'0.0%',"オンライン決済を導入された場合"),
("CDN無料枠",1024,'#,##0" GB／月"',"CloudFrontの永年無料枠（月1TB）")]
r=5
for lbl,val,fmt,note in ASM:
    C(ws,r,1,lbl,b=True); C(ws,r,2,val,fmt=fmt,al="right",b=True,fill=YEL,col="0000FF")
    C(ws,r,3,note,sz=9,col=MUT,wrap=True); ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=5)
    ws.row_dimensions[r].height=22; r+=1
r+=1

C(ws,r,1,"A. 弊社へお支払いいただく月額（定額）",b=True,fill=SUBF,col=NAVY)
for c in range(2,6): C(ws,r,c,"",fill=SUBF)
r+=1; HDR(ws,r,["項目","月額","内容","","備考"],h=22); r+=1; a0=r
OURS=[("保守・運用サービス",80000,"システムの死活監視とエラー監視、障害対応、軽微な改修、受講者からの問い合わせ一次対応、月次レポート","受講者数によらず定額"),
("法令改正監視・通知サービス",20000,"官報・厚生労働省の通達・パブリックコメントを継続監視。対象10科目に影響しうる改正を検知したら、聖建様と弊社の双方へ自動でメール通知","監視の仕組みは弊社の共通基盤。初期開発費は不要")]
for lbl,val,desc,note in OURS:
    C(ws,r,1,lbl,wrap=True,b=True); C(ws,r,2,val,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
    C(ws,r,3,desc,sz=9,col=MUT,wrap=True); ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
    C(ws,r,5,note,sz=9,col=MUT,wrap=True); ws.row_dimensions[r].height=40; r+=1
a1=r-1
C(ws,r,1,"弊社サービス 小計",b=True,fill=LTF)
C(ws,r,2,f"=SUM(B{a0}:B{a1})",fmt=YEN,al="right",b=True,fill=LTF)
for c in (3,4,5): C(ws,r,c,"",fill=LTF)
OURTOT=f"'02_単価と前提'!$B${r}"; r+=2

C(ws,r,1,"B. AWS 利用料（お客様ご負担・本番構成）",b=True,fill=SUBF,col=ACC)
for c in range(2,6): C(ws,r,c,"",fill=SUBF)
r+=1; HDR(ws,r,["サービス","月額","用途","","備考"],h=22); r+=1; b0=r
AWSF=[("EC2（冗長構成 2台）",9000,"t3.medium×2。Laravel／Nginx／PHP-FPM／FFmpeg。ALB配下で冗長化","1台構成なら約4,500円"),
("RDS PostgreSQL（Multi-AZ）",11000,"db.t3.small 相当。自動バックアップ、PITR、自動フェイルオーバー","記録を守る中核。削減対象外"),
("ElastiCache Redis（レプリカ付き）",4000,"多重視聴の排他制御、キュー、レート制限","レプリカなしなら約2,000円"),
("Application Load Balancer",3000,"HTTPSの終端と負荷分散",""),
("NAT Gateway",6700,"プライベートサブネットからの外部通信","VPCエンドポイント活用で削減可"),
("S3（オブジェクトストレージ）",500,"教材動画、本人確認画像、バックアップの保管","従量。受講者増で微増"),
("Route 53 / ACM",200,"ドメインとSSL証明書","ACMは無料"),
("CloudWatch",1600,"メトリクス監視、ログ保管、アラート通知","")]
for lbl,val,desc,note in AWSF:
    C(ws,r,1,lbl,wrap=True); C(ws,r,2,val,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
    C(ws,r,3,desc,sz=9,col=MUT,wrap=True); ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
    C(ws,r,5,note,sz=9,col=MUT,wrap=True); ws.row_dimensions[r].height=28; r+=1
b1=r-1
C(ws,r,1,"AWS 固定費 小計",b=True,fill=LTF)
C(ws,r,2,f"=SUM(B{b0}:B{b1})",fmt=YEN,al="right",b=True,fill=LTF,col=ACC)
for c in (3,4,5): C(ws,r,c,"",fill=LTF)
AWSTOT=f"'02_単価と前提'!$B${r}"; r+=2

C(ws,r,1,"C. AWS 従量課金の単価（お客様ご負担）",b=True,fill=SUBF,col=ACC)
for c in range(2,6): C(ws,r,c,"",fill=SUBF)
r+=1; HDR(ws,r,["項目","単価","単位","算出根拠","備考"],h=22); r+=1
VR={}
VAR=[("動画配信（CDN転送）","=0.114*$B$5",'¥#,##0.0',"円／GB","CloudFront 日本 $0.114／GB","月1TBまで無料枠あり","cdn"),
("追加ストレージ","=0.025*$B$5",'¥#,##0.0',"円／GB／月","S3標準 $0.025／GB／月","本人確認画像の蓄積分","stg"),
("顔照合API","=0.001*$B$5",'¥#,##0.00',"円／回","Rekognition CompareFaces","1人あたり22回","fc"),
("ライブネス判定","=0.0195*$B$5",'¥#,##0.00',"円／回","Rekognition Face Liveness","1人あたり1回","lv"),
("メール送信",0.02,'¥#,##0.00',"円／通","SES","1人あたり約5通","ml")]
for lbl,val,fmt,unit,basis,note,key in VAR:
    C(ws,r,1,lbl,wrap=True); C(ws,r,2,val,fmt=fmt,al="right",b=True,col=OK)
    C(ws,r,3,unit,sz=9,al="center",col=MUT); C(ws,r,4,basis,sz=9,col=MUT,wrap=True)
    C(ws,r,5,note,sz=9,col=MUT,wrap=True)
    VR[key]=f"'02_単価と前提'!$B${r}"; ws.row_dimensions[r].height=24; r+=1
r+=1
C(ws,r,1,"D. 制作期間中のみ",b=True,fill=SUBF,col=NAVY)
for c in range(2,6): C(ws,r,c,"",fill=SUBF)
r+=1
C(ws,r,1,"アバター・音声合成サービス",wrap=True)
C(ws,r,2,5000,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
C(ws,r,3,"HeyGen等の動画生成サービスの月額利用料。10科目の制作が完了すれば解約できます",sz=9,col=MUT,wrap=True)
ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=5)
ws.row_dimensions[r].height=28; AVA=f"'02_単価と前提'!$B${r}"; r+=2
for t in ["■ 注意","・AWSの単価は2026年9月時点の公開情報に基づく想定です。実際の請求は構成と使用量で変動します。",
"・正式なご提示の前に、実構成でのAWS Pricing Calculatorによる試算をお勧めします。",
"・Savings Plans（1年契約）を利用すると、EC2とRDSでおおむね30%程度の削減が見込めます。"]:
    C(ws,r,1,t,bd=False,sz=9.5,b=t.startswith("■"),col=NAVY if t.startswith("■") else MUT); r+=1

# =====================================================================
# 01_月額サマリー
# =====================================================================
ws=wb.create_sheet("01_月額サマリー",0)
ws["A1"]="月額ランニングコスト（受講者数別）"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="受講者数（黄色セル）を変えると全体が再計算されます。金額はすべて税抜です。"
ws["A2"].font=Font(name=F,size=9,color=MUT)
for c,w in zip("ABCDEF",[36,15,15,15,38,4]): ws.column_dimensions[c].width=w
HDR(ws,4,["項目","ケース1","ケース2","ケース3","内容・算出方法"],h=22)
C(ws,5,1,"月間の受講者数",b=True,fill=SUBF)
for i,n in enumerate([50,100,300]): C(ws,5,2+i,n,fmt='#,##0"名"',al="center",b=True,fill=YEL,col="0000FF")
C(ws,5,5,"年間では600名／1,200名／3,600名",sz=9,col=MUT,fill=SUBF); ws.row_dimensions[5].height=22

r=7
C(ws,r,1,"A. 弊社へお支払い（定額）",b=True,fill=SUBF,col=NAVY)
for c in range(2,6): C(ws,r,c,"",fill=SUBF)
r+=1
C(ws,r,1,"保守・運用サービス",b=True)
for i in range(3): C(ws,r,2+i,"='02_単価と前提'!$B$15",fmt=YEN,al="right")
C(ws,r,5,"死活監視・エラー監視、障害対応、軽微改修、問い合わせ一次対応",sz=9,col=MUT,wrap=True)
ws.row_dimensions[r].height=24; r+=1
C(ws,r,1,"法令改正監視・通知サービス",b=True)
for i in range(3): C(ws,r,2+i,"='02_単価と前提'!$B$16",fmt=YEN,al="right")
C(ws,r,5,"改正を検知したら聖建様と弊社の双方へ自動メール通知",sz=9,col=MUT,wrap=True)
ws.row_dimensions[r].height=24; r+=1
ourrow=r
C(ws,r,1,"弊社サービス 小計",b=True,fill=LTF)
for i,col in enumerate("BCD"): C(ws,r,2+i,f"=SUM({col}{r-2}:{col}{r-1})",fmt=YEN,al="right",b=True,fill=LTF)
C(ws,r,5,"",fill=LTF); r+=2

C(ws,r,1,"B. AWS 利用料（お客様ご負担）",b=True,fill=SUBF,col=ACC)
for c in range(2,6): C(ws,r,c,"",fill=SUBF)
r+=1
C(ws,r,1,"AWS 固定費（本番構成）",b=True)
for i in range(3): C(ws,r,2+i,f"={AWSTOT}",fmt=YEN,al="right",col=ACC)
C(ws,r,5,"EC2冗長・RDS Multi-AZ・ElastiCache・ALB・NAT・S3・Route53・CloudWatch",sz=9,col=MUT,wrap=True)
ws.row_dimensions[r].height=26; r+=1; v0=r
VL=[("動画配信（CDN転送）",lambda n:f"=ROUND(MAX(0,{n}*'02_単価と前提'!$B$6-'02_単価と前提'!$B$11)*{VR['cdn']},0)","受講者数×3.5GB。月1TBの無料枠を控除"),
("追加ストレージ",lambda n:f"=ROUND({n}*0.02*{VR['stg']},0)","本人確認画像 1人あたり約20MB"),
("顔照合API",lambda n:f"=ROUND({n}*'02_単価と前提'!$B$7*{VR['fc']},0)","受講者数×22回×単価"),
("ライブネス判定",lambda n:f"=ROUND({n}*'02_単価と前提'!$B$8*{VR['lv']},0)","受講者数×1回×単価"),
("メール送信",lambda n:f"=ROUND({n}*5*{VR['ml']},0)","受講者数×5通×単価")]
for lbl,fn,note in VL:
    C(ws,r,1,lbl)
    for i,col in enumerate("BCD"): C(ws,r,2+i,fn(f"${col}$5"),fmt=YEN,al="right",col=ACC)
    C(ws,r,5,note,sz=9,col=MUT,wrap=True); ws.row_dimensions[r].height=22; r+=1
v1=r-1
awsrow=r
C(ws,r,1,"AWS 小計（固定＋従量）",b=True,fill=LTF,col=ACC)
for i,col in enumerate("BCD"): C(ws,r,2+i,f"={col}{v0-1}+SUM({col}{v0}:{col}{v1})",fmt=YEN,al="right",b=True,fill=LTF,col=ACC)
C(ws,r,5,"",fill=LTF); r+=2

C(ws,r,1,"C. 制作期間中のみ",b=True,fill=SUBF,col=NAVY)
for c in range(2,6): C(ws,r,c,"",fill=SUBF)
r+=1
C(ws,r,1,"アバター・音声合成サービス")
for i in range(3): C(ws,r,2+i,f"={AVA}",fmt=YEN,al="right",col=MUT)
C(ws,r,5,"10科目の制作完了後は解約できます（約10か月間）",sz=9,col=MUT,wrap=True)
avarow=r; r+=2

tot=r
C(ws,r,1,"月額 合計（制作期間中）",b=True,fill=SUBF,sz=11,col=NAVY)
for i,col in enumerate("BCD"): C(ws,r,2+i,f"={col}{ourrow}+{col}{awsrow}+{col}{avarow}",fmt=YEN,al="right",b=True,fill=SUBF,sz=11)
C(ws,r,5,"最初の約10か月",sz=9,col=MUT,fill=SUBF); ws.row_dimensions[r].height=24; r+=1
C(ws,r,1,"月額 合計（制作完了後）",b=True,fill=SUBF,sz=12,col=NAVY)
for i,col in enumerate("BCD"): C(ws,r,2+i,f"={col}{ourrow}+{col}{awsrow}",fmt=YEN,al="right",b=True,fill=SUBF,sz=12)
C(ws,r,5,"11か月目以降の定常状態",sz=9,col=MUT,fill=SUBF); ws.row_dimensions[r].height=26
steady=r; r+=1
C(ws,r,1,"うち 受講者1名あたり",b=True)
for i,col in enumerate("BCD"): C(ws,r,2+i,f"=ROUND({col}{steady}/{col}5,0)",fmt=YEN,al="right",b=True,col=ACC)
C(ws,r,5,"受講者が増えるほど1名あたりの負担は下がります",sz=9,col=MUT,wrap=True); r+=2

C(ws,r,1,"参考：売上との比較",b=True,fill=SUBF,col=NAVY)
for c in range(2,6): C(ws,r,c,"",fill=SUBF)
r+=1
C(ws,r,1,"月間の想定売上")
for i,col in enumerate("BCD"): C(ws,r,2+i,f"=${col}$5*'02_単価と前提'!$B$9",fmt=YEN,al="right",col=MUT)
C(ws,r,5,"受講料10,000円で試算",sz=9,col=MUT); sal=r; r+=1
C(ws,r,1,"決済手数料（導入された場合）")
for i,col in enumerate("BCD"): C(ws,r,2+i,f"=ROUND({col}{sal}*'02_単価と前提'!$B$10,0)",fmt=YEN,al="right",col=BAD)
C(ws,r,5,"第2次発注でオンライン決済を導入された場合",sz=9,col=MUT,wrap=True); fee=r; r+=1
C(ws,r,1,"差引（教材の償却前）",b=True)
for i,col in enumerate("BCD"): C(ws,r,2+i,f"={col}{sal}-{col}{steady}-{col}{fee}",fmt=YEN,al="right",b=True,col=OK)
C(ws,r,5,"ここが黒字になる受講者数が損益分岐です",sz=9,col=MUT,wrap=True); r+=2
for t in ["■ 読み方",
"・弊社サービス100,000円とAWS36,000円で、月136,000円が定常のランニングコストです。",
"・受講者が6倍（50名→300名）になっても、増えるのは2,500円程度です。CDNの無料枠が効いています。",
"・したがってこの事業の損益は「固定費を何名で割るか」でほぼ決まります。",
"・受講料10,000円なら、月15名前後で固定費を回収できる計算になります。"]:
    C(ws,r,1,t,bd=False,sz=9.5,b=t.startswith("■"),col=NAVY if t.startswith("■") else MUT); r+=1

# =====================================================================
# 03_サーバー構成の比較
# =====================================================================
ws=wb.create_sheet("03_サーバー構成の比較")
ws["A1"]="サーバー構成の比較"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="配信システム基本設計書 1.1 の前提（PHP 8.3／Laravel／PostgreSQL／Redis／HLS／FFmpeg／常駐ワーカー）に照らした評価です。"
ws["A2"].font=Font(name=F,size=9,color=MUT)
for c,w in zip("ABCD",[26,30,30,30]): ws.column_dimensions[c].width=w
HDR(ws,4,["評価項目","エックスサーバー（共用）","Xserver VPS","AWS（採用）"],h=26)
CMP=[("月額の目安","1,100円程度","4,000〜10,000円程度","36,000円程度（本番構成）"),
("PostgreSQL","✕ 非対応（MySQL系のみ）","○ 自分で導入","◎ RDS マネージド"),
("Redis","✕ 使えない","○ 自分で導入","◎ ElastiCache"),
("FFmpegでの動画変換","✕ 実行制限で不可","○ 可能","◎ 可能"),
("常駐ワーカー（Queue／Scheduler）","✕ 持てない","○ 可能","◎ 可能"),
("HLS配信・署名付きURL・CDN","✕ 実質不可","△ 別途CDNが必要","◎ S3＋CloudFront"),
("DBの自動バックアップ・PITR","✕","△ 自分で仕組みを作る","◎ 標準機能"),
("障害時の自動切替","✕","✕","◎ Multi-AZ で自動"),
("受講者増への対応","✕","△ 手作業でプラン変更","◎ スケールアウト可能"),
("顔照合API との親和性","―","△ 外部連携","◎ 同一基盤・国内リージョン"),
("総合評価","✕ 採用不可","△ 条件付きで可","◎ 本件で採用")]
r=5
for row in CMP:
    lbl=row[0]
    C(ws,r,1,lbl,b=True,wrap=True,fill=LTF if lbl=="総合評価" else None)
    for i,v in enumerate(row[1:],2):
        f=None; col=INK
        if v.startswith("✕"): f=BADF; col=BAD
        elif v.startswith("◎"): f=OKF; col=OK
        C(ws,r,i,v,wrap=True,sz=9.5,fill=f,col=col,b=(lbl=="総合評価"))
    ws.row_dimensions[r].height=26; r+=1
r+=1
for t,col in [("■ AWSを採用する理由",NAVY),
("エックスサーバーの共用レンタルサーバーでは、この設計は動きません。",BAD),
("　PostgreSQLもRedisも使えず、FFmpegでの動画変換も常駐ワーカーも持てません。",MUT),
("",None),
("Xserver VPSなら技術的には構築できますが、バックアップ・冗長化・監視・スケールを",INK),
("　すべて自前で作ることになり、障害時の復旧時間が読めません。",MUT),
("",None),
("本件でAWSを採用する理由は、費用ではなく記録を失わないためです。",OK),
("　本人確認と受講の記録には3年以上の保存義務があります。この記録を失うことは、",MUT),
("　事業の根幹である「証明できること」が崩れることを意味します。",MUT),
("　RDSのMulti-AZ・自動バックアップ・PITRは、この事業では必需品です。",MUT)]:
    C(ws,r,1,t,bd=False,sz=10,b=(col in (NAVY,BAD,OK)),col=col or MUT); r+=1

# =====================================================================
# 04_法令改正監視サービス
# =====================================================================
ws=wb.create_sheet("04_法令改正監視サービス")
ws["A1"]="法令改正監視・通知サービス"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="月額 20,000円（税抜）。初期開発費は発生しません。"
ws["A2"].font=Font(name=F,size=9,color=MUT)
for c,w in zip("ABC",[26,50,46]): ws.column_dimensions[c].width=w
HDR(ws,4,["項目","内容","備考"],h=22)
SV=[("監視の対象","官報、厚生労働省の通達・告示、パブリックコメント、安全衛生特別教育規程の改正","対象10科目に関係しうる範囲"),
("検知の方法","対象法令の版を定期取得して差分を検知。あわせてキーワード（各科目名、特別教育、安全衛生特別教育規程 等）で新着を抽出","毎日自動で実行"),
("通知","検知したら、聖建様のご担当者と弊社の双方へ同時に自動メールを送信","改正の名称、公布日、施行日、該当しうる科目、参照URLを記載"),
("月次報告","該当なしの月も含め、毎月末に監視結果をご報告","記録として残すことに意味があります"),
("通知先の設定","聖建様側の宛先は複数登録できます。担当者の変更にも対応します","管理画面から変更可能"),
("改正が見つかったら","影響範囲をご報告し、改訂が必要と判断された科目について都度お見積りします","1科目 300,000円（新規制作と同額）")]
r=5
for lbl,desc,note in SV:
    C(ws,r,1,lbl,b=True,wrap=True); C(ws,r,2,desc,wrap=True,sz=9.5); C(ws,r,3,note,wrap=True,sz=9,col=MUT)
    ws.row_dimensions[r].height=38; r+=1
r+=1
C(ws,r,1,"責任の分担",b=True,col=NAVY,bd=False,sz=11.5); r+=1
HDR(ws,r,["工程","担当","内容"],h=20); r+=1
RESP=[("官報・通達の機械的な監視と通知","弊社","毎日の自動監視と、検知時の自動メール通知"),
("その改正が教材のどこに影響するかの判断","聖建様の監修者","法令の解釈と影響範囲の確定。弊社は行いません"),
("改訂作業の実施","弊社","台本の修正、該当部分の動画の再生成と差し替え、版の更新"),
("改訂後の内容の承認","聖建様の監修者","改訂内容が現場の実態と合っているかの確認")]
for pr,who,desc in RESP:
    C(ws,r,1,pr,wrap=True,sz=9.5)
    C(ws,r,2,who,al="center",b=True,col=OK if who=="弊社" else BAD,
      fill=OKF if who=="弊社" else BADF)
    C(ws,r,3,desc,wrap=True,sz=9,col=MUT); ws.row_dimensions[r].height=30; r+=1
r+=1
for t,col in [("■ なぜこのサービスが必要か",NAVY),
("特別教育の科目・範囲・時間は安全衛生特別教育規程で定められています。",MUT),
("改正に気づかないまま古い教材を売り続けると、その修了証の有効性が問われます。",BAD),
("受講者は「現場に入るため」に購入されているため、修了証が無効になれば直接の損害になります。",BAD),
("",None),
("今回の10科目には、改正が比較的多い分野が含まれています。",MUT),
("　・石綿　　　　　　　　石綿障害予防規則の改正が続いている領域",MUT),
("　・特定粉じん　　　　　化学物質規制の見直しの影響を受けやすい",MUT),
("　・テールゲートリフター 2024年に新設されたばかりで運用面の通達が出る可能性",MUT),
("",None),
("■ 初期開発費が不要な理由",NAVY),
("監視の仕組みは弊社が共通サービスとして保有し、聖建様専用に開発するものではありません。",MUT),
("そのため初期費用は発生せず、月額でご提供できます。",OK)]:
    C(ws,r,1,t,bd=False,sz=10,b=(col in (NAVY,)),col=col or MUT); r+=1

# =====================================================================
# 05_契約の分け方
# =====================================================================
ws=wb.create_sheet("05_契約の分け方")
ws["A1"]="AWS利用料をお客様のご負担とする理由"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
for c,w in zip("ABC",[30,44,44]): ws.column_dimensions[c].width=w
HDR(ws,4,["","弊社の月額に含める場合","お客様のご負担とする場合（採用）"],h=22)
CC=[("受講者が増えたとき","転送量・顔照合・ストレージが増え、定額では赤字になります。事業が成功するほど弊社が損をします。","実費が受講者数に応じて増えるだけです。受講料で回収できます。"),
("原価の見え方","何にいくらかかっているか分かりません。","実費が明細で見えます。受講料の設定根拠になります。"),
("為替が動いたとき","ドル建て費用の上昇を弊社が吸収することになります。","実費として自動的に反映されます。"),
("お取引が終わるとき","契約主体が弊社のため、引き継ぎが煩雑になります。","聖建様名義であれば、そのまま継続できます。"),
("値上げ交渉","毎年お願いすることになります。","不要です。")]
r=5
for lbl,a,b in CC:
    C(ws,r,1,lbl,b=True,wrap=True)
    C(ws,r,2,a,wrap=True,sz=9.5,fill=BADF,col=BAD)
    C(ws,r,3,b,wrap=True,sz=9.5,fill=OKF,col=OK)
    ws.row_dimensions[r].height=44; r+=1
r+=1
for t,col in [("■ 契約の形",NAVY),("",None),
("① 弊社への月額 100,000円（税抜）",INK),
("　　保守・運用サービス 80,000円 ＋ 法令改正監視・通知サービス 20,000円。",MUT),
("　　人的サービスの対価であり、受講者数に左右されません。",MUT),
("",None),
("② AWS 利用料 ＝ 聖建様名義で直接ご契約（推奨）",INK),
("　　弊社は運用代行のみを行います。立替えが発生せず、費用が透明になります。",MUT),
("　　アカウント管理が難しい場合は、弊社が代行契約し、実費＋事務手数料10%でご請求する形も可能です。",MUT),
("",None),
("■ ご提示のしかた",NAVY),
("「月額10万円は保守・運用と法令監視の費用です。サーバーの実費は受講者数で変わるため、",INK),
("　聖建様のお名前で直接ご契約いただきたく存じます。いまの規模で月3万6千円程度、",INK),
("　受講者が月300名になっても月3万9千円ほどの見込みです。",INK),
("　定額に含めてしまうと、御社の受講者が増えたときに私どもが値上げをお願いすることになり、",INK),
("　かえってご迷惑をおかけします。」",INK),
("",None),
("この説明なら、値切りではなく合理性の話として受け取っていただけます。",OK)]:
    C(ws,r,1,t,bd=False,sz=10,b=(col in (NAVY,)),col=col or MUT); r+=1

if "Sheet" in wb.sheetnames: del wb["Sheet"]
wb._sheets=[wb[n] for n in ["01_月額サマリー","02_単価と前提","03_サーバー構成の比較","04_法令改正監視サービス","05_契約の分け方"]]
wb.save("/home/user/-/running/聖建様_ランニングコスト明細.xlsx")
print("saved")
