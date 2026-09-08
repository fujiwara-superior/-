# -*- coding: utf-8 -*-
"""聖建様向け ランニングコスト明細"""
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
YEN='¥#,##0;(¥#,##0);-'; YEN2='¥#,##0.00;;-'

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
ws["A2"]="黄色のセルが入力欄です。ここを変えると全シートが自動で再計算されます。金額は2026年9月時点の公開情報に基づく想定値です。"
ws["A2"].font=Font(name=F,size=9,color=MUT)

C(ws,4,1,"共通の前提",b=True,fill=SUBF,col=NAVY)
for c in range(2,6): C(ws,4,c,"",fill=SUBF)
ASM=[("為替レート",150,'¥#,##0"／USD"',"AWS等のドル建て費用の換算に使用します"),
     ("1受講あたりの動画転送量",3.5,'0.0" GB"',"学科4.5時間・720p（約1.5Mbps）＋再視聴分を見込んだ値"),
     ("1受講あたりの顔照合回数",22,'0" 回"',"受講開始時＋受講中のランダム照合（1時間あたり3回）"),
     ("1受講あたりのライブネス判定",1,'0" 回"',"アカウント登録時の1回のみ"),
     ("受講料（決済手数料の算出用）",10000,YEN,"プランAの想定単価。売上連動費の試算に使用します"),
     ("決済手数料率",0.036,'0.0%',"クレジットカード決済の一般的な料率"),
     ("CDN無料枠",1024,'#,##0" GB／月"',"CloudFrontの永年無料枠。月1TBまでインターネットへの転送が無料")]
r=5
for lbl,val,fmt,note in ASM:
    C(ws,r,1,lbl,b=True)
    C(ws,r,2,val,fmt=fmt,al="right",b=True,fill=YEL,col="0000FF")
    C(ws,r,3,note,sz=9,col=MUT,wrap=True)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=5)
    ws.row_dimensions[r].height=22; r+=1
ws.column_dimensions["A"].width=32; ws.column_dimensions["B"].width=15
for c,w in zip("CDE",[30,22,22]): ws.column_dimensions[c].width=w
RATE="$B$5"; GBP="$B$6"; FCN="$B$7"; LVN="$B$8"; FEE="$B$9"; FEER="$B$10"; FREE="$B$11"

r+=1
C(ws,r,1,"A. 固定費（受講者数に関係なく毎月発生）",b=True,fill=SUBF,col=NAVY)
for c in range(2,6): C(ws,r,c,"",fill=SUBF)
r+=1
HDR(ws,r,["項目","月額","内容","提供元","備考"],h=22)
r+=1; fx0=r
FIX=[("保守・運用サービス（弊社）",80000,"監視、障害対応、軽微な改修、受講者からの問い合わせ一次対応","株式会社スペリオル","聖建様よりご提示の金額"),
("アプリケーションサーバー",4500,"EC2 t3.medium 相当（2vCPU／4GB）。Laravel／Nginx／PHP-FPM","AWS","Xserver VPSなら約4,000〜8,000円"),
("データベース",5500,"RDS PostgreSQL db.t3.small 相当。自動バックアップとPITRを含む","AWS","自前構築なら安いが復旧責任を負う"),
("キャッシュ・キュー",2000,"ElastiCache Redis cache.t3.micro 相当。多重視聴の排他制御に使用","AWS","VPS同居なら0円"),
("ロードバランサ・固定IP・その他",3500,"ALB、Elastic IP、NAT等","AWS",""),
("ドメイン・SSL",200,"年間約2,000円を月割。SSLはACMで無料","AWS／レジストラ",""),
("監視・エラー追跡",3000,"Sentry等。障害と不正兆候の検知","外部SaaS","無料枠から開始すれば当面0円"),
("バックアップ保管",500,"S3標準〜低頻度アクセス。DBスナップショットと記録の長期保存","AWS","3年保存義務に対応"),
("アバター・音声合成サービス",5000,"教材を継続制作する期間のみ。制作を終えれば解約可能","外部SaaS","初年度のみ想定")]
for lbl,val,desc,src,note in FIX:
    C(ws,r,1,lbl,wrap=True)
    C(ws,r,2,val,fmt=YEN,al="right",b=True,fill=YEL,col="0000FF")
    C(ws,r,3,desc,sz=9,col=MUT,wrap=True)
    C(ws,r,4,src,sz=9,al="center",col=MUT)
    C(ws,r,5,note,sz=9,col=MUT,wrap=True)
    ws.row_dimensions[r].height=30; r+=1
fx1=r-1
C(ws,r,1,"固定費 小計",b=True,fill=LTF)
C(ws,r,2,f"=SUM(B{fx0}:B{fx1})",fmt=YEN,al="right",b=True,fill=LTF)
for c in (3,4,5): C(ws,r,c,"",fill=LTF)
FIXTOT=f"'02_単価と前提'!$B${r}"; r+=2

C(ws,r,1,"B. 変動費の単価（受講者数に比例）",b=True,fill=SUBF,col=NAVY)
for c in range(2,6): C(ws,r,c,"",fill=SUBF)
r+=1
HDR(ws,r,["項目","単価","単位","算出根拠","備考"],h=22)
r+=1
VARROWS={}
VAR=[("動画配信（CDN転送）",f"=0.114*{RATE}",'¥#,##0.0',"円／GB","CloudFront 日本 $0.114／GB","月1TBまで無料枠あり","cdn"),
("オブジェクトストレージ",f"=0.025*{RATE}",'¥#,##0.0',"円／GB／月","S3標準 $0.025／GB／月","動画58時間分＋本人確認画像","stg"),
("顔照合API",f"=0.001*{RATE}",'¥#,##0.00',"円／回","Rekognition CompareFaces $0.001／画像","受講者1人あたり22回","fc"),
("ライブネス判定",f"=0.0195*{RATE}",'¥#,##0.00',"円／回","Rekognition Face Liveness $0.0195／回","受講者1人あたり1回","lv"),
("メール送信",0.02,'¥#,##0.00',"円／通","SES等。1人あたり約5通","ほぼ無視できる水準","ml")]
for lbl,val,fmt,unit,basis,note,key in VAR:
    C(ws,r,1,lbl,wrap=True)
    C(ws,r,2,val,fmt=fmt,al="right",b=True,col=OK)
    C(ws,r,3,unit,sz=9,al="center",col=MUT)
    C(ws,r,4,basis,sz=9,col=MUT,wrap=True)
    C(ws,r,5,note,sz=9,col=MUT,wrap=True)
    VARROWS[key]=f"'02_単価と前提'!$B${r}"
    ws.row_dimensions[r].height=26; r+=1
r+=1
for t in ["■ この明細の考え方",
 "・固定費は、受講者が0人でも毎月発生します。事業の損益分岐を決めるのはこの金額です。",
 "・変動費は受講者数に比例します。受講料に含めて回収する性質の費用です。",
 "・決済手数料は売上に比例するため、別枠（シート01の C欄）で表示しています。",
 "",
 "■ 出典と注意",
 "・AWSの単価は2026年9月時点の公開情報に基づく想定です。実際の請求はリージョン・構成・使用量で変動します。",
 "・CloudFrontには月1TBの永年無料枠があります。立ち上げ期の配信費はほぼ発生しません。",
 "・正式なご提示の前に、実際の構成でAWS Pricing Calculatorによる試算をお勧めします。"]:
    C(ws,r,1,t,bd=False,sz=9.5,b=t.startswith("■"),col=NAVY if t.startswith("■") else MUT); r+=1

# =====================================================================
# 01_月額サマリー
# =====================================================================
ws=wb.create_sheet("01_月額サマリー",0)
ws["A1"]="月額ランニングコスト（受講者数別）"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="受講者数（3行目の黄色セル）を変えると全体が再計算されます。金額は税抜です。"
ws["A2"].font=Font(name=F,size=9,color=MUT)
for c,w in zip("ABCDEF",[34,16,16,16,34,4]): ws.column_dimensions[c].width=w

HDR(ws,4,["項目","ケース1","ケース2","ケース3","算出方法"],h=22)
C(ws,5,1,"月間の受講者数",b=True,fill=SUBF)
for i,(col,n) in enumerate(zip("BCD",[50,100,300])):
    C(ws,5,2+i,n,fmt='#,##0"名"',al="center",b=True,fill=YEL,col="0000FF")
C(ws,5,5,"想定される規模。年間では600名／1,200名／3,600名",sz=9,col=MUT,fill=SUBF)
ws.row_dimensions[5].height=22
N={0:"B5",1:"C5",2:"D5"}

r=7
C(ws,r,1,"A. 固定費",b=True,fill=SUBF,col=NAVY)
for c in range(2,6): C(ws,r,c,"",fill=SUBF)
r+=1
C(ws,r,1,"保守・運用サービス（弊社）",b=True)
for i in range(3): C(ws,r,2+i,"='02_単価と前提'!$B$15",fmt=YEN,al="right")
C(ws,r,5,"定額。監視・障害対応・軽微改修・問い合わせ一次対応",sz=9,col=MUT,wrap=True)
r+=1
C(ws,r,1,"AWS 固定費",b=True)
for i in range(3): C(ws,r,2+i,"=SUM('02_単価と前提'!$B$16:$B$20)+'02_単価と前提'!$B$22",fmt=YEN,al="right",col=ACC)
C(ws,r,5,"EC2・RDS・ElastiCache・ALB・S3・Route53。AWSに直接お支払いいただく分",sz=9,col=MUT,wrap=True)
awsfix=r; r+=1
C(ws,r,1,"その他SaaS 固定費",b=True)
for i in range(3): C(ws,r,2+i,"='02_単価と前提'!$B$21+'02_単価と前提'!$B$23",fmt=YEN,al="right")
C(ws,r,5,"監視サービス（無料枠から開始可）、アバター・音声合成（教材制作中のみ）",sz=9,col=MUT,wrap=True)
r+=1
fixrow=r
C(ws,r,1,"固定費 小計",b=True,fill=LTF)
for i,col in enumerate("BCD"): C(ws,r,2+i,f"=SUM({col}{r-3}:{col}{r-1})",fmt=YEN,al="right",b=True,fill=LTF)
C(ws,r,5,"",fill=LTF)
r+=2

C(ws,r,1,"B. 変動費（受講者数に比例）",b=True,fill=SUBF,col=NAVY)
for c in range(2,6): C(ws,r,c,"",fill=SUBF)
r+=1; v0=r
VLINES=[("動画配信（CDN転送）",lambda n:f"=ROUND(MAX(0,{n}*'02_単価と前提'!$B$6-'02_単価と前提'!$B$11)*{VARROWS['cdn']},0)",
         "受講者数×3.5GB。月1TBの無料枠を差し引いて課金"),
("オブジェクトストレージ",lambda n:f"=ROUND((50+{n}*0.02)*{VARROWS['stg']},0)",
 "教材動画 約50GB＋本人確認画像（1人あたり約20MB）"),
("顔照合API",lambda n:f"=ROUND({n}*'02_単価と前提'!$B$7*{VARROWS['fc']},0)",
 "受講者数×22回×単価"),
("ライブネス判定",lambda n:f"=ROUND({n}*'02_単価と前提'!$B$8*{VARROWS['lv']},0)",
 "受講者数×1回×単価"),
("メール送信",lambda n:f"=ROUND({n}*5*{VARROWS['ml']},0)","受講者数×5通×単価")]
for lbl,fn,note in VLINES:
    C(ws,r,1,lbl)
    for i,col in enumerate("BCD"): C(ws,r,2+i,fn(f"${col}$5"),fmt=YEN,al="right")
    C(ws,r,5,note,sz=9,col=MUT,wrap=True)
    ws.row_dimensions[r].height=24; r+=1
v1=r-1
varrow=r
C(ws,r,1,"変動費 小計",b=True,fill=LTF)
for i,col in enumerate("BCD"): C(ws,r,2+i,f"=SUM({col}{v0}:{col}{v1})",fmt=YEN,al="right",b=True,fill=LTF)
C(ws,r,5,"",fill=LTF)
r+=2

tot=r
C(ws,r,1,"月額ランニングコスト 合計（税抜）",b=True,fill=SUBF,sz=12,col=NAVY)
for i,col in enumerate("BCD"): C(ws,r,2+i,f"={col}{fixrow}+{col}{varrow}",fmt=YEN,al="right",b=True,fill=SUBF,sz=12)
C(ws,r,5,"決済手数料を除く",sz=9,col=MUT,fill=SUBF)
ws.row_dimensions[r].height=26
r+=1
C(ws,r,1,"うち 受講者1名あたり",b=True)
for i,col in enumerate("BCD"): C(ws,r,2+i,f"=ROUND({col}{tot}/{col}5,0)",fmt=YEN,al="right",b=True,col=ACC)
C(ws,r,5,"受講者が増えるほど1名あたりの負担は下がります",sz=9,col=MUT,wrap=True)
r+=1
C(ws,r,1,"うち AWSへのお支払い（固定＋変動）",b=True,fill=LTF,col=ACC)
for i,col in enumerate("BCD"): C(ws,r,2+i,f"={col}{awsfix}+{col}{varrow}",fmt=YEN,al="right",b=True,fill=LTF,col=ACC)
C(ws,r,5,"変動費はすべてAWS（CDN・S3・Rekognition・SES）です",sz=9,col=MUT,wrap=True,fill=LTF)
r+=2

C(ws,r,1,"C. 売上連動費（参考・上記合計には含みません）",b=True,fill=SUBF,col=BAD)
for c in range(2,6): C(ws,r,c,"",fill=SUBF)
r+=1
C(ws,r,1,"決済手数料")
for i,col in enumerate("BCD"):
    C(ws,r,2+i,f"=ROUND(${col}$5*'02_単価と前提'!$B$9*'02_単価と前提'!$B$10,0)",fmt=YEN,al="right",col=BAD)
C(ws,r,5,"受講者数×受講料10,000円×3.6%。売上から支払う費用です",sz=9,col=MUT,wrap=True)
ws.row_dimensions[r].height=24; r+=1
C(ws,r,1,"（参考）月間の想定売上")
for i,col in enumerate("BCD"): C(ws,r,2+i,f"=${col}$5*'02_単価と前提'!$B$9",fmt=YEN,al="right",col=MUT)
C(ws,r,5,"受講料10,000円で試算",sz=9,col=MUT)
r+=1
C(ws,r,1,"（参考）売上－ランニング費用",b=True)
for i,col in enumerate("BCD"): C(ws,r,2+i,f"={col}{r-1}-{col}{tot}-{col}{r-2}",fmt=YEN,al="right",b=True,col=OK)
C(ws,r,5,"教材の償却前。ここが黒字になる受講者数が損益分岐です",sz=9,col=MUT,wrap=True)
r+=2
for t in ["■ 読み方",
 "・固定費は受講者が0名でも発生します。月額約10万円が事業を維持する最低ラインです。",
 "・変動費は受講者数に比例しますが、規模の割に小さい金額です。CDNの無料枠が効いています。",
 "・したがって、この事業の損益は「固定費を何名で割るか」でほぼ決まります。",
 "・受講料10,000円なら、月11〜12名で固定費を回収できる計算になります。"]:
    C(ws,r,1,t,bd=False,sz=9.5,b=t.startswith("■"),col=NAVY if t.startswith("■") else MUT); r+=1

# =====================================================================
# 03_サーバー構成の比較
# =====================================================================
ws=wb.create_sheet("03_サーバー構成の比較")
ws["A1"]="サーバー構成の比較"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
ws["A2"]="配信システム基本設計書 1.1 が求める前提（PHP 8.3／Laravel／PostgreSQL／Redis／HLS／FFmpeg／常駐ワーカー）に照らした評価です。"
ws["A2"].font=Font(name=F,size=9,color=MUT)
for c,w in zip("ABCD",[26,30,30,30]): ws.column_dimensions[c].width=w
HDR(ws,4,["評価項目","エックスサーバー（共用）","Xserver VPS","AWS（推奨）"],h=26)
CMP=[("月額の目安","1,100円程度","4,000〜10,000円程度\n（2026年9月1日 改定後・要確認）","15,000〜20,000円程度"),
("PostgreSQL","✕ 非対応（MySQL系のみ）","○ 自分で導入","◎ RDSマネージド"),
("Redis","✕ 使えない","○ 自分で導入","◎ ElastiCache"),
("FFmpegでの動画変換","✕ 実行制限で不可","○ 可能","◎ 可能"),
("常駐ワーカー（Queue／Scheduler）","✕ 持てない","○ 可能","◎ 可能"),
("HLS配信・署名付きURL・CDN","✕ 実質不可","△ 別途CDNが必要","◎ S3＋CloudFront"),
("DBの自動バックアップ・PITR","✕","△ 自分で仕組みを作る","◎ 標準機能"),
("障害からの復旧","✕","△ 手作業。復旧時間が読めない","◎ スナップショットから復元"),
("受講者増への対応","✕","△ 手作業でプラン変更","◎ スケールアウト可能"),
("顔照合API（Rekognition）との親和性","―","△ 外部連携","◎ 同一基盤・国内リージョン"),
("運用にかかる手間","―","多い","少ない"),
("総合評価","✕ 採用不可","△ 条件付きで可","◎ 推奨")]
r=5
for row in CMP:
    lbl=row[0]
    C(ws,r,1,lbl,b=True,wrap=True,fill=LTF if lbl=="総合評価" else None)
    for i,v in enumerate(row[1:],2):
        f=None; col=INK
        if v.startswith("✕"): f=BADF; col=BAD
        elif v.startswith("◎"): f=OKF; col=OK
        C(ws,r,i,v,wrap=True,sz=9.5,fill=f,col=col,b=(lbl=="総合評価"))
    ws.row_dimensions[r].height=34 if "\n" in "".join(row) else 26
    r+=1
r+=1
GUIDE=[("■ 結論",NAVY),
("エックスサーバーの共用レンタルサーバーでは、この設計は動きません。",BAD),
("　PostgreSQLもRedisも使えず、FFmpegでの動画変換も常駐ワーカーも持てません。設計書の前提を満たしません。",MUT),
("",None),
("Xserver VPS なら技術的には構築できます。root権限があるため、必要なものは自分で導入できます。",INK),
("　ただし、バックアップ・冗長化・監視・スケールをすべて自分で作ることになります。",MUT),
("",None),
("弊社はAWSを推奨します。理由は費用ではなく、記録を失わないためです。",OK),
("　本人確認と受講の記録は3年以上の保存義務があります。この記録を失うことは、事業の根幹である",MUT),
("　「証明できること」が崩れることを意味します。RDSの自動バックアップとPITRは、この事業では",MUT),
("　贅沢品ではなく必需品です。月額の差は1万円程度で、その保険料としては安いと考えます。",MUT),
("",None),
("　顔照合にAWS Rekognitionを使うため、同じ基盤に載せるほうが構成が単純になり、",MUT),
("　個人データの国外移転の問題も生じにくくなります。",MUT),
("",None),
("■ 費用を抑える方法",NAVY),
("・監視サービスは無料枠から開始できます（当面0円）。",MUT),
("・アバター・音声合成のサブスクは、教材の制作が終われば解約できます。",MUT),
("・受講者が少ないうちはサーバーを小さく始め、増えてから上げれば十分です。",MUT),
("・CloudFrontの無料枠（月1TB）があるため、立ち上げ期の配信費はほぼ発生しません。",MUT)]
for t,col in GUIDE:
    C(ws,r,1,t,bd=False,sz=10,b=(col in (NAVY,BAD,OK)),col=col or MUT); r+=1

# =====================================================================
# 04_契約の分け方
# =====================================================================
ws=wb.create_sheet("04_契約の分け方")
ws["A1"]="月額8万円に、サーバー費を含めるべきか"
ws["A1"].font=Font(name=F,bold=True,size=15,color=NAVY); ws.row_dimensions[1].height=24
for c,w in zip("ABC",[30,44,44]): ws.column_dimensions[c].width=w
C(ws,3,1,"結論：含めないでください。実費は別途とすべきです。",b=True,sz=12,col=BAD,bd=False)
ws.merge_cells("A3:C3")
HDR(ws,5,["","含める場合","別途にする場合（推奨）"],h=22)
CC=[("受講者が増えたとき","転送量・顔照合・ストレージが増え、8万円では赤字になります。事業が成功するほど御社が損をします。","実費が受講者数に応じて増えるだけです。受講料で回収できます。"),
("聖建様側の見え方","何にいくらかかっているか分かりません。原価管理ができません。","実費が明細で見えます。受講料の設定根拠になります。"),
("為替が動いたとき","ドル建て費用が上がっても御社が吸収することになります。","実費として自動的に反映されます。"),
("契約が終わるとき","サーバーの契約主体が御社のため、引き継ぎが煩雑になります。","聖建様名義であれば、そのまま継続できます。"),
("値上げ交渉","毎年「上げてください」と言う必要があります。","不要です。")]
r=6
for lbl,a,b in CC:
    C(ws,r,1,lbl,b=True,wrap=True)
    C(ws,r,2,a,wrap=True,sz=9.5,fill=BADF,col=BAD)
    C(ws,r,3,b,wrap=True,sz=9.5,fill=OKF,col=OK)
    ws.row_dimensions[r].height=44; r+=1
r+=1
REC=[("■ 推奨する契約の形",NAVY),
("",None),
("① 月額 80,000円（税抜）＝ 保守・運用サービス",INK),
("　　監視、障害対応、軽微な改修、受講者からの問い合わせ一次対応、月次レポート。",MUT),
("　　人的サービスの対価であり、受講者数に左右されません。",MUT),
("",None),
("② インフラ実費 ＝ 聖建様名義でAWSと直接ご契約（推奨）",INK),
("　　御社は運用代行のみを行います。立替えが発生せず、費用が透明になります。",MUT),
("　　聖建様側でアカウント管理が難しい場合は、弊社が代行契約し、",MUT),
("　　実費＋事務手数料10%でご請求する形も可能です。",MUT),
("",None),
("③ 従量サービス（顔照合・ライブネス・決済）",INK),
("　　②と同じ扱いとします。受講者数に比例するため、定額に含めるべきではありません。",MUT),
("",None),
("■ ご提示のしかた",NAVY),
("「月額8万円は保守・運用の費用です。サーバーとAPIの実費は、受講者数で変わるため別途とさせてください。",INK),
("　いまの規模なら月2万円程度、受講者が月300名になっても月3万円台の見込みです。",INK),
("　定額に含めてしまうと、御社の受講者が増えたときに私どもが値上げをお願いすることになり、",INK),
("　かえってご迷惑をおかけします。」",INK),
("",None),
("この説明なら、値切りではなく合理性の話として受け取っていただけます。",OK)]
for t,col in REC:
    C(ws,r,1,t,bd=False,sz=10,b=(col in (NAVY,)),col=col or MUT); r+=1

if "Sheet" in wb.sheetnames: del wb["Sheet"]
wb._sheets=[wb[n] for n in ["01_月額サマリー","02_単価と前提","03_サーバー構成の比較","04_契約の分け方"]]
wb.save("/home/user/-/running/聖建様_ランニングコスト明細.xlsx")
print("saved")
