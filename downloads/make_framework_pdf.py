"""
ケース面接 論点整理フレームワーク - 無料相談参加者特典PDF生成
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak,
    Table, TableStyle
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from pathlib import Path

# フォント登録（Noto Sans JP 可変フォント）
FONT_PATH = r"C:/Windows/Fonts/NotoSansJP-VF.ttf"
pdfmetrics.registerFont(TTFont("NotoJP", FONT_PATH))

# カラー
TIFFANY = HexColor("#0ABAB5")
TIFFANY_DARK = HexColor("#058A85")
GRAY_900 = HexColor("#1F2937")
GRAY_700 = HexColor("#4B5563")
GRAY_500 = HexColor("#6B7280")
GRAY_200 = HexColor("#E5E7EB")
GRAY_50 = HexColor("#F9FAFB")

out = Path(__file__).parent / "case_interview_framework.pdf"

doc = SimpleDocTemplate(
    str(out), pagesize=A4,
    topMargin=20 * mm, bottomMargin=20 * mm,
    leftMargin=22 * mm, rightMargin=22 * mm,
    title="ケース面接 論点整理フレームワーク",
    author="SYMMETRY Lab株式会社",
)

styles = getSampleStyleSheet()

h1 = ParagraphStyle(
    "h1", parent=styles["Heading1"], fontName="NotoJP",
    fontSize=22, leading=30, textColor=GRAY_900,
    spaceAfter=4, alignment=TA_LEFT,
)
brand = ParagraphStyle(
    "brand", parent=styles["Normal"], fontName="NotoJP",
    fontSize=9, leading=12, textColor=TIFFANY_DARK,
    spaceAfter=24,
)
h2 = ParagraphStyle(
    "h2", parent=styles["Heading2"], fontName="NotoJP",
    fontSize=14, leading=22, textColor=GRAY_900,
    spaceBefore=18, spaceAfter=10,
    borderPadding=0,
)
h3 = ParagraphStyle(
    "h3", parent=styles["Heading3"], fontName="NotoJP",
    fontSize=11, leading=18, textColor=TIFFANY_DARK,
    spaceBefore=10, spaceAfter=4,
)
body = ParagraphStyle(
    "body", parent=styles["BodyText"], fontName="NotoJP",
    fontSize=9.5, leading=16, textColor=GRAY_700,
    spaceAfter=8, alignment=TA_LEFT,
)
small = ParagraphStyle(
    "small", parent=styles["BodyText"], fontName="NotoJP",
    fontSize=8, leading=13, textColor=GRAY_500,
)

story = []

# ============ 表紙 ============
story.append(Spacer(1, 40 * mm))
story.append(Paragraph("SYMMETRY LAB  ／  CASE INTERVIEW", brand))
story.append(Paragraph("ケース面接<br/>論点整理フレームワーク", h1))
story.append(Spacer(1, 8 * mm))
story.append(Paragraph(
    "戦略コンサル志望の学生が、本番で通用する<br/>"
    "思考プロセスを体系化するための実戦ガイド。",
    ParagraphStyle(
        "lead", parent=body, fontSize=11, leading=20, textColor=GRAY_700,
    ),
))
story.append(Spacer(1, 60 * mm))

# 表紙フッター
cover_footer = Table(
    [[Paragraph("無料相談参加者限定資料", small),
      Paragraph("SYMMETRY Lab 株式会社", small)]],
    colWidths=[80 * mm, 80 * mm],
)
cover_footer.setStyle(TableStyle([
    ("LINEABOVE", (0, 0), (-1, 0), 0.5, TIFFANY),
    ("TOPPADDING", (0, 0), (-1, -1), 10),
    ("ALIGN", (1, 0), (1, 0), "RIGHT"),
]))
story.append(cover_footer)
story.append(PageBreak())

# ============ イントロ ============
story.append(Paragraph("はじめに", h2))
story.append(Paragraph(
    "ケース面接で評価されるのは「答え」ではなく、そこに至る「思考プロセス」です。"
    "本書では、合格者が共通して持つ4つの思考ステップを、<b>Frame → Structure → "
    "Hypothesis → Verify</b> の順で解説します。",
    body,
))
story.append(Paragraph(
    "独学でケース対策を進めている方が陥りがちな「フレームワーク暗記」を脱し、"
    "面接官の評価軸に沿った思考を身につけるためのガイドとしてご活用ください。",
    body,
))

# ============ STEP 1 ============
story.append(Paragraph("STEP 01 ／ Frame — 論点設計", h2))
story.append(Paragraph(
    "ケース面接で最初に問われるのは「何を問われているか」を正確に言語化する力です。"
    "問題文の裏にあるクライアントの真の論点（Issue）を特定し、議論の枠組みを設計します。",
    body,
))
story.append(Paragraph("▸ チェックリスト", h3))
framing_items = [
    "・クライアントは誰か、最終的に何を決めたいのか",
    "・解くべき問い（Key Question）を1文で言えるか",
    "・検討の前提条件（時間軸・地域・対象セグメント）を確認したか",
    "・面接官と論点の合意形成ができているか",
]
for item in framing_items:
    story.append(Paragraph(item, body))

# ============ STEP 2 ============
story.append(Paragraph("STEP 02 ／ Structure — 構造化", h2))
story.append(Paragraph(
    "論点を MECE に分解し、検討すべき要素を網羅的に洗い出します。"
    "ここで汎用フレームワーク（3C / 4P / PEST）を機械的に当てはめるのは悪手。"
    "問題ごとに「自分の手で」構造を組み立てる訓練が必要です。",
    body,
))
story.append(Paragraph("▸ 構造化の3原則", h3))
structure_items = [
    "・<b>MECE</b>：漏れなくダブりなく、同じ階層を保つ",
    "・<b>Relevance</b>：論点に直結する切り口を選ぶ（もっとも差がつく軸）",
    "・<b>Depth</b>：分解は目的に対して十分な深さまで行う（2〜3階層）",
]
for item in structure_items:
    story.append(Paragraph(item, body))
story.append(PageBreak())

# ============ STEP 3 ============
story.append(Paragraph("STEP 03 ／ Hypothesis — 仮説構築", h2))
story.append(Paragraph(
    "構造化した要素の中で「最も重要な論点はどこか」を特定し、"
    "仮の答え（仮説）を立てます。"
    "ケース面接は「網羅性」よりも「深さ」が問われる場。"
    "全てを均等に検討する時間はなく、仮説に基づく優先順位付けが勝負を決めます。",
    body,
))
story.append(Paragraph("▸ 強い仮説の条件", h3))
hypothesis_items = [
    "・<b>Specific</b>：抽象論ではなく、数字や固有名詞を含む",
    "・<b>Actionable</b>：クライアントが意思決定できるレベルの具体性がある",
    "・<b>Testable</b>：次のステップで検証可能である",
]
for item in hypothesis_items:
    story.append(Paragraph(item, body))

# ============ STEP 4 ============
story.append(Paragraph("STEP 04 ／ Verify — 検証と結論", h2))
story.append(Paragraph(
    "仮説を検証するためのデータ・分析・ロジックを展開し、"
    "結論として「So What（だから何が言えるか）」を明確にします。"
    "面接官は、結論の正しさよりも「どう考えたか」「どこで詰まったか」を見ています。",
    body,
))
story.append(Paragraph("▸ 結論提示の型", h3))
verify_items = [
    "・<b>結論</b>：一文で、クライアントが取るべきアクションを示す",
    "・<b>根拠</b>：3点以内で、数字・定性情報を交えて補強",
    "・<b>リスク</b>：前提が崩れた場合に起こりうる影響を添える",
]
for item in verify_items:
    story.append(Paragraph(item, body))

# ============ 独学の限界 ============
story.append(Paragraph("独学の限界と、1対1指導の価値", h2))
story.append(Paragraph(
    "このフレームワーク自体は書籍でも解説されていますが、実際に使いこなせるようになるには「自分の思考のどこで詰まっているか」を外部から指摘してもらうことが不可欠です。",
    body,
))
story.append(Paragraph(
    "SYMMETRY Lab では、外資系戦略コンサル出身者が実案件ベースのオリジナルケースを用いて、"
    "あなたの思考プロセスを1対1でレビューします。まずは30分の無料相談で、"
    "現在地と志望ファームまでの距離をご一緒に整理させてください。",
    body,
))
story.append(Spacer(1, 8 * mm))

# CTAボックス
cta_content = [
    [Paragraph("無料相談のお申込み", ParagraphStyle(
        "ctah", parent=body, fontSize=11, textColor=white,
        fontName="NotoJP",
    ))],
    [Paragraph(
        "30分・オンライン開催・当日予約可<br/>"
        "https://symmetrylab.jp/lp-case.html",
        ParagraphStyle(
            "ctab", parent=body, fontSize=9, textColor=white,
            fontName="NotoJP", leading=15,
        ),
    )],
]
cta = Table(cta_content, colWidths=[160 * mm])
cta.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, -1), TIFFANY),
    ("TOPPADDING", (0, 0), (-1, -1), 12),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
    ("LEFTPADDING", (0, 0), (-1, -1), 20),
    ("RIGHTPADDING", (0, 0), (-1, -1), 20),
]))
story.append(cta)
story.append(Spacer(1, 20 * mm))

# フッター
story.append(Paragraph(
    "© SYMMETRY Lab 株式会社 ／ 本資料の無断転載・二次配布を禁じます。",
    small,
))

doc.build(story)
print(f"PDF generated: {out}")
print(f"Size: {out.stat().st_size} bytes")
