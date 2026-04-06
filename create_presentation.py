from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ─── カラーパレット ───
NAVY   = RGBColor(0x1F, 0x4E, 0x79)
BLUE   = RGBColor(0x2E, 0x75, 0xB6)
LBLUE  = RGBColor(0xD6, 0xE4, 0xF0)
ORANGE = RGBColor(0xC5, 0x50, 0x1B)
GREEN  = RGBColor(0x37, 0x86, 0x44)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY  = RGBColor(0xF2, 0xF2, 0xF2)
DGRAY  = RGBColor(0x40, 0x40, 0x40)
YELLOW = RGBColor(0xFF, 0xF2, 0xCC)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]  # 完全ブランク

def add_slide():
    return prs.slides.add_slide(BLANK)

def rect(slide, l, t, w, h, fill=None, line=None):
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.line.fill.background()
    if fill:
        shape.fill.solid(); shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
    else:
        shape.line.fill.background()
    return shape

def txbox(slide, text, l, t, w, h,
          size=18, bold=False, color=DGRAY, align=PP_ALIGN.LEFT,
          wrap=True):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return tb

def header_bar(slide, title, subtitle=""):
    rect(slide, 0, 0, 13.33, 1.3, fill=NAVY)
    txbox(slide, title, 0.4, 0.1, 10, 0.7,
          size=28, bold=True, color=WHITE)
    if subtitle:
        txbox(slide, subtitle, 0.4, 0.82, 10, 0.4,
              size=14, color=LBLUE)

def footer(slide, note="※ 本資料は作成時点の情報に基づきます。制度詳細は中小企業庁・IT導入補助金公式サイトをご確認ください。"):
    rect(slide, 0, 7.1, 13.33, 0.4, fill=NAVY)
    txbox(slide, note, 0.3, 7.13, 12.7, 0.35,
          size=9, color=LBLUE)

def bullet_box(slide, items, l, t, w, h, size=16, indent="　", color=DGRAY):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for item in items:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(4)
        run = p.add_run()
        run.text = indent + item
        run.font.size = Pt(size)
        run.font.color.rgb = color

# ══════════════════════════════════════════
# Slide 1: タイトル
# ══════════════════════════════════════════
s = add_slide()
rect(s, 0, 0, 13.33, 7.5, fill=NAVY)
rect(s, 0, 2.5, 13.33, 2.8, fill=BLUE)

txbox(s, "IT導入補助金　申請ガイド",
      0.6, 2.6, 12, 1.0, size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
txbox(s, "ホームページのオンラインショップ化　申請手順・記入例",
      0.6, 3.65, 12, 0.6, size=20, color=YELLOW, align=PP_ALIGN.CENTER)
txbox(s, "作成：社会保険労務士法人　　　対象：総務担当者・社長",
      0.6, 6.4, 12, 0.4, size=13, color=LBLUE, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════
# Slide 2: 結論（3行以内）
# ══════════════════════════════════════════
s = add_slide()
header_bar(s, "結　論", "まず3点を押さえてください")

rect(s, 0.4, 1.5, 12.5, 2.5, fill=YELLOW, line=ORANGE)
conclusions = [
    "① EC化（オンラインショップ）はIT導入補助金「通常枠」の対象となり得ます",
    "② 申請はITベンダー（登録済み）と一緒に行うため、まずベンダー選定が必須です",
    "③ gBizIDプライムの取得に2〜3週間かかるため、今すぐ申請開始を推奨します",
]
bullet_box(s, conclusions, 0.6, 1.55, 12.1, 2.3, size=17, indent="", color=DGRAY)

rect(s, 0.4, 4.2, 12.5, 1.0, fill=LBLUE)
txbox(s, "⚠ 重要：IT導入補助金の申請は「先着順」ではなく「公募期間制」です。\n    締切を必ず確認し、余裕を持ったスケジュールで進めてください。",
      0.6, 4.25, 12.1, 0.85, size=14, color=NAVY)

rect(s, 0.4, 5.4, 5.8, 1.5, fill=LGRAY)
txbox(s, "【補助率】", 0.6, 5.45, 2, 0.4, size=14, bold=True, color=NAVY)
txbox(s, "対象経費の1/2以内", 0.6, 5.85, 5, 0.5, size=16, color=DGRAY)

rect(s, 6.5, 5.4, 6.4, 1.5, fill=LGRAY)
txbox(s, "【補助額】", 6.7, 5.45, 3, 0.4, size=14, bold=True, color=NAVY)
txbox(s, "5万円〜450万円（類型により異なる）", 6.7, 5.85, 6, 0.5, size=16, color=DGRAY)

footer(s)

# ══════════════════════════════════════════
# Slide 3: IT導入補助金とは
# ══════════════════════════════════════════
s = add_slide()
header_bar(s, "IT導入補助金とは", "中小企業・小規模事業者向けITツール導入支援制度")

# 左カラム
rect(s, 0.4, 1.4, 6.0, 1.0, fill=BLUE)
txbox(s, "制度概要", 0.6, 1.45, 5.5, 0.55, size=16, bold=True, color=WHITE)
items_l = [
    "中小企業庁が所管する補助金制度",
    "業務効率化・売上向上に資するITツール導入費用を補助",
    "ソフトウェア・クラウドサービス・導入費用が対象",
    "EC（ネットショップ）構築も対象となる場合あり",
]
bullet_box(s, items_l, 0.6, 2.45, 5.6, 2.5, size=15, indent="▶ ")

# 右カラム
rect(s, 6.9, 1.4, 6.0, 1.0, fill=BLUE)
txbox(s, "主な申請枠（2024年度）", 7.1, 1.45, 5.5, 0.55, size=16, bold=True, color=WHITE)

rows = [
    ("通常枠（A・B類型）", "業務効率化・売上拡大向け。\nEC構築はこちら"),
    ("インボイス枠",       "インボイス対応ツール専用"),
    ("セキュリティ枠",     "サイバーセキュリティ対策"),
]
for i, (name, desc) in enumerate(rows):
    top = 2.5 + i * 1.05
    rect(s, 6.9, top, 2.2, 0.9, fill=LBLUE)
    txbox(s, name, 7.0, top+0.05, 2.0, 0.8, size=13, bold=True, color=NAVY)
    rect(s, 9.15, top, 3.7, 0.9, fill=LGRAY)
    txbox(s, desc, 9.25, top+0.05, 3.5, 0.8, size=13, color=DGRAY)

txbox(s, "※ 制度内容・補助率は年度・公募回により変更されます。申請前に公式サイトを必ず確認してください。",
      0.4, 6.5, 12.5, 0.5, size=11, color=ORANGE)
footer(s)

# ══════════════════════════════════════════
# Slide 4: 申請の流れ（7ステップ）
# ══════════════════════════════════════════
s = add_slide()
header_bar(s, "申請の流れ（7ステップ）", "ITベンダーと二人三脚で進めます")

steps = [
    ("STEP 1", "gBizIDプライム取得",      "商工会議所・GビズIDサイトで申請。\n発行まで2〜3週間かかるため最優先で着手", BLUE),
    ("STEP 2", "SECURITY ACTION宣言",    "IPAのサイトで「★一つ星」以上を宣言\n（無料・オンライン即日完了）", BLUE),
    ("STEP 3", "ITベンダー・ツール選定",   "IT導入補助金登録済みベンダーの中から\nEC構築ができる業者を選定", GREEN),
    ("STEP 4", "交付申請（ベンダーと共同）","事業計画・導入効果を記入して電子申請\n審査期間：約1〜2か月", ORANGE),
    ("STEP 5", "採択通知・交付決定",       "採択後に交付決定通知が届く\n※採択前の発注・支払いは補助対象外", ORANGE),
    ("STEP 6", "契約・発注・支払い",       "交付決定後に正式契約・支払いを実施", GREEN),
    ("STEP 7", "実績報告・補助金受領",     "完了報告をシステムで提出\n審査後に補助金が振り込まれる", NAVY),
]

cols = [0.3, 4.55, 8.8]
for i, (num, title, desc, col) in enumerate(steps):
    c = i % 3
    r = i // 3
    lft = cols[c]
    top = 1.4 + r * 2.6
    w = 4.0

    rect(s, lft, top, w, 0.55, fill=col)
    txbox(s, f"{num}　{title}", lft+0.1, top+0.05, w-0.2, 0.45,
          size=14, bold=True, color=WHITE)
    rect(s, lft, top+0.55, w, 1.5, fill=LGRAY)
    txbox(s, desc, lft+0.15, top+0.6, w-0.25, 1.35, size=13, color=DGRAY)

# STEP 7 は span
rect(s, 8.8, 1.4+2.6, 4.0, 0.55, fill=NAVY)
txbox(s, "STEP 7　実績報告・補助金受領", 8.9, 1.4+2.65, 3.8, 0.45,
      size=14, bold=True, color=WHITE)
rect(s, 8.8, 1.4+2.6+0.55, 4.0, 1.5, fill=LGRAY)
txbox(s, "完了報告をシステムで提出\n審査後に補助金が振り込まれる", 8.95, 1.4+2.6+0.6, 3.7, 1.35,
      size=13, color=DGRAY)

footer(s)

# ══════════════════════════════════════════
# Slide 5: 申請書　主な記入項目
# ══════════════════════════════════════════
s = add_slide()
header_bar(s, "申請書　主な記入項目", "交付申請フォーム（ITツール導入支援事業 申請マイページ）")

sections = [
    ("① 企業・事業者情報", [
        "法人名・代表者名・所在地・資本金・従業員数",
        "業種（日本標準産業分類）",
        "直近の決算情報（売上高・経常利益）",
    ]),
    ("② 導入するITツール情報", [
        "ベンダー名・ツール名（登録済みリストから選択）",
        "導入費用の内訳（ソフト費・クラウド費・導入支援費など）",
        "ツールの機能カテゴリ（EC・販売管理 等）",
    ]),
    ("③ 事業計画・導入効果", [
        "現状の課題（例：受注がすべて電話・FAXで非効率）",
        "ITツール導入後の業務改善内容",
        "定量的目標（売上○%向上・業務時間○時間削減 など）",
    ]),
    ("④ セキュリティ・誓約事項", [
        "SECURITY ACTION宣言済み証明",
        "労働生産性の伸び率目標（3年後）",
        "不正受給をしない旨の誓約",
    ]),
]

for i, (title, items) in enumerate(sections):
    c = i % 2
    r = i // 2
    lft = 0.4 + c * 6.5
    top = 1.4 + r * 2.7
    w = 6.1

    rect(s, lft, top, w, 0.5, fill=BLUE)
    txbox(s, title, lft+0.15, top+0.05, w-0.2, 0.42, size=15, bold=True, color=WHITE)
    rect(s, lft, top+0.5, w, 2.0, fill=LGRAY)
    bullet_box(s, items, lft+0.2, top+0.55, w-0.3, 1.8, size=14, indent="✔ ")

footer(s)

# ══════════════════════════════════════════
# Slide 6: 記入例（見本）
# ══════════════════════════════════════════
s = add_slide()
header_bar(s, "記入例（見本）", "③ 事業計画・導入効果　のイメージ")

rect(s, 0.4, 1.4, 12.5, 5.4, fill=WHITE, line=BLUE)
txbox(s, "【記入例】", 0.6, 1.45, 3, 0.4, size=14, bold=True, color=BLUE)

examples = [
    ("現状の課題",
     "当社のホームページは情報掲示のみで、商品購入は電話・FAXのみ対応。\n"
     "受注処理に1日平均2時間を要しており、休日・夜間の注文に対応できていない。\n"
     "機会損失が年間推定○○万円と試算される。"),
    ("導入するITツール",
     "○○社提供のECプラットフォーム「△△」を導入し、自社HPにEC機能を実装。\n"
     "商品管理・在庫管理・決済処理・顧客管理を一元化する。"),
    ("導入後の改善内容と目標",
     "・24時間365日の受注対応が可能になり、受注機会を拡大\n"
     "・受注処理時間を現状比 50%削減（2時間→1時間/日）\n"
     "・導入1年後に売上高 15%向上（運用差あり・仮定値）を目指す\n"
     "・労働生産性：3年後に 3%以上向上（申請要件）"),
]

top = 1.95
for label, text in examples:
    rect(s, 0.5, top, 2.8, 0.35, fill=NAVY)
    txbox(s, label, 0.6, top+0.02, 2.6, 0.3, size=13, bold=True, color=WHITE)
    rect(s, 3.35, top, 9.4, 0.35, fill=LBLUE)
    txbox(s, label, 3.45, top+0.02, 9.1, 0.3, size=13, bold=True, color=NAVY)
    rect(s, 0.5, top+0.35, 12.25, 0.95, fill=LGRAY)
    txbox(s, text, 0.65, top+0.37, 12.0, 0.9, size=13, color=DGRAY)
    top += 1.38

txbox(s, "※ 数値は仮定値です。実際の申請では自社の実績データを入力してください。",
      0.5, 6.55, 12.3, 0.35, size=11, color=ORANGE)
footer(s)

# ══════════════════════════════════════════
# Slide 7: 注意点・リスク
# ══════════════════════════════════════════
s = add_slide()
header_bar(s, "注意点・リスク", "申請前に必ず確認してください")

warnings = [
    ("採択前の発注・支払いは補助対象外",
     "交付決定通知を受け取る前に契約・支払いをした場合、全額自己負担になります。\nITベンダーへの発注は必ず採択通知後に行ってください。"),
    ("登録ITベンダー以外は対象外",
     "補助金対象のITツールは「IT導入支援事業者」として登録されたベンダーのツールのみです。\n見積もりを取るベンダーが登録済みか事前に確認してください。"),
    ("公募期間・締切の確認",
     "IT導入補助金は複数回の公募があり、締切を過ぎると次回まで申請できません。\n最新のスケジュールを公式サイトで確認してください。"),
    ("事業計画の定量目標が必須",
     "売上向上・業務時間削減などの数値目標が求められます。\n根拠のある数字を用意してください（「仮定」の場合はその旨記載）。"),
]

for i, (title, desc) in enumerate(warnings):
    top = 1.4 + i * 1.35
    rect(s, 0.4, top, 0.5, 1.1, fill=ORANGE)
    txbox(s, "!", 0.4, top+0.2, 0.5, 0.65, size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    rect(s, 0.95, top, 3.5, 0.45, fill=NAVY)
    txbox(s, title, 1.05, top+0.05, 3.3, 0.38, size=13, bold=True, color=WHITE)
    rect(s, 0.95, top+0.45, 11.95, 0.65, fill=LGRAY)
    txbox(s, desc, 1.1, top+0.48, 11.7, 0.6, size=13, color=DGRAY)

footer(s)

# ══════════════════════════════════════════
# Slide 8: 次のアクション（チェックリスト）
# ══════════════════════════════════════════
s = add_slide()
header_bar(s, "次のアクション　チェックリスト", "今日からできることから着手してください")

rect(s, 0.4, 1.4, 6.1, 5.5, fill=LGRAY)
txbox(s, "【貴社がやること】", 0.6, 1.45, 5.5, 0.45, size=15, bold=True, color=NAVY)
client_actions = [
    "□ gBizIDプライムを今すぐ申請する（2〜3週間）",
    "□ IPAサイトでSECURITY ACTION★一つ星を宣言する",
    "□ 自社の現状課題・売上データを整理する",
    "□ 登録ITベンダーから見積りを3社以上取得する",
    "□ 公募スケジュール（締切日）を確認し社内共有する",
    "□ 事業計画の数値目標を担当者と社長で合意する",
]
bullet_box(s, client_actions, 0.6, 1.95, 5.7, 4.7, size=14, indent="", color=DGRAY)

rect(s, 6.8, 1.4, 6.1, 5.5, fill=LBLUE)
txbox(s, "【当法人がサポートできること】", 7.0, 1.45, 5.7, 0.45, size=15, bold=True, color=NAVY)
support_actions = [
    "✔ 申請書の事業計画文章のレビュー",
    "✔ 労働生産性指標の計算サポート",
    "✔ 書類チェックリストの提供",
    "✔ gBizID取得手順のご案内",
    "",
    "※ IT導入補助金の申請実務は\n   ITベンダーまたは認定支援機関が\n   主体となります。",
    "   必要に応じて専門機関をご紹介します。",
]
bullet_box(s, support_actions, 7.0, 1.95, 5.7, 4.7, size=14, indent="", color=DGRAY)

footer(s)

# ══════════════════════════════════════════
# Slide 9: 参考リンク・お問い合わせ
# ══════════════════════════════════════════
s = add_slide()
header_bar(s, "参考情報・お問い合わせ先", "一次情報を必ずご確認ください")

links = [
    ("IT導入補助金 公式サイト（中小企業庁）",
     "https://it-shien.smrj.go.jp/",
     "最新の公募要領・対象ツール一覧・申請マイページはこちら"),
    ("GビズID（gBizIDプライム取得）",
     "https://gbiz-id.go.jp/",
     "申請にはgBizIDプライムが必須。今すぐ取得申請を"),
    ("SECURITY ACTION（IPA）",
     "https://www.ipa.go.jp/security/security-action/",
     "★一つ星の宣言はオンラインで即日完了"),
    ("中小企業基盤整備機構　相談窓口",
     "https://www.smrj.go.jp/",
     "IT導入補助金に関する無料相談が可能"),
]

for i, (name, url, desc) in enumerate(links):
    top = 1.4 + i * 1.35
    rect(s, 0.4, top, 12.5, 1.15, fill=LGRAY)
    rect(s, 0.4, top, 0.35, 1.15, fill=BLUE)
    txbox(s, name, 0.9, top+0.05, 11.8, 0.4, size=15, bold=True, color=NAVY)
    txbox(s, url,  0.9, top+0.42, 11.8, 0.3, size=12, color=BLUE)
    txbox(s, desc, 0.9, top+0.72, 11.8, 0.35, size=12, color=DGRAY)

footer(s, "本資料の内容は情報提供を目的としています。申請手続きの詳細は必ず公式情報をご確認のうえ、ITベンダーと連携して進めてください。")

# ─── 保存 ───
out = "/home/user/story20250803/IT導入補助金_申請ガイド_プレゼン.pptx"
prs.save(out)
print(f"保存完了: {out}")
