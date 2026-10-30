from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

DARK_BG = RGBColor(6, 9, 19)
CYAN = RGBColor(0, 242, 254)
BLUE = RGBColor(56, 189, 248)
EMERALD = RGBColor(16, 185, 129)
WHITE = RGBColor(255, 255, 255)
MUTED = RGBColor(148, 163, 184)

def add_bg(slide):
    bg = slide.shapes.add_shape(1, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = DARK_BG
    bg.line.fill.background()

# Slide 1: Title
s1 = prs.slides.add_slide(blank_layout)
add_bg(s1)
tb = s1.shapes.add_textbox(Inches(1.2), Inches(1.8), Inches(11), Inches(4))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = 'ROUND 1 REVIEW · RETAIL TRACK'
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = CYAN

p2 = tf.add_paragraph()
p2.text = 'VerifiedCart'
p2.font.size = Pt(48)
p2.font.bold = True
p2.font.color.rgb = WHITE

p3 = tf.add_paragraph()
p3.text = "The Grounded Conversational Shopping Assistant That Doesn't Lie"
p3.font.size = Pt(22)
p3.font.color.rgb = BLUE

p4 = tf.add_paragraph()
p4.text = '\n5,000 SKUs · Zero Hallucinations · Multi-Turn Memory · 3D WebGL Hologram\nDemo: http://127.0.0.1:8000'
p4.font.size = Pt(15)
p4.font.color.rgb = MUTED

# Slide 2: Problem
s2 = prs.slides.add_slide(blank_layout)
add_bg(s2)
tb = s2.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.3), Inches(6))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = 'THE PROBLEM: WHY GENERATIVE CHATBOTS FAIL IN RETAIL'
p.font.size = Pt(24)
p.font.bold = True
p.font.color.rgb = CYAN

p2 = tf.add_paragraph()
p2.text = 'A hallucinated price is a legal liability, not a cosmetic bug.\n'
p2.font.size = Pt(16)
p2.font.color.rgb = MUTED

bullets = [
    ('The Hallucination Trap: ', 'LLMs probabilistically predict words. Quoting Rs. 999 for a Rs. 1,499 item violates consumer protection laws.'),
    ('Refusal Failure: ', 'Standard chatbots invent answers for out-of-catalogue items (e.g. selling iPhones on an apparel store) instead of refusing.'),
    ('Blind Guessing on Ambiguity: ', 'When asked for blue running shoes when 3 exist, normal bots guess randomly rather than asking clarifying questions.'),
    ('Unenforced Policy Claims: ', 'Chatbots hallucinate return and warranty rules (e.g. broken hygiene seals), destroying buyer trust.')
]
for title, desc in bullets:
    p = tf.add_paragraph()
    p.text = '• ' + title + desc
    p.font.size = Pt(16)
    p.font.color.rgb = WHITE

# Slide 3: Architecture
s3 = prs.slides.add_slide(blank_layout)
add_bg(s3)
tb = s3.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.3), Inches(6))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = 'OUR ARCHITECTURE: 4-STAGE ZERO-HALLUCINATION PIPELINE'
p.font.size = Pt(24)
p.font.bold = True
p.font.color.rgb = CYAN

p2 = tf.add_paragraph()
p2.text = 'Separating natural language understanding from deterministic factual assertion.\n'
p2.font.size = Pt(16)
p2.font.color.rgb = MUTED

stages = [
    ('Stage 1 — Hinglish & Intent Normalizer: ', 'Parses transliterated Hindi (sabse sasta, kitne ka), extracts budget limits, and resolves brand synonyms.'),
    ('Stage 2 — Inverted Index & Policy Matcher: ', 'Ultra-fast token retrieval over 5,000 SKUs and 38 store policy sections (Sec 1.1 to Sec 8.1).'),
    ('Stage 3 — Tri-State Decision Gate: ', 'Deterministic routing: ANSWER (exact SKU match), CLARIFY (ambiguous request), or REFUSE (out-of-catalogue).'),
    ('Stage 4 — Verbatim Snapshot Extractor: ', 'AI never produces prices. Numbers are strictly copied from catalogue.csv into the response. Zero hallucination.')
]
for title, desc in stages:
    p = tf.add_paragraph()
    p.text = '• ' + title + desc
    p.font.size = Pt(16)
    p.font.color.rgb = WHITE

# Slide 4: Comparison
s4 = prs.slides.add_slide(blank_layout)
add_bg(s4)
tb = s4.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.3), Inches(6))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = 'KEY DIFFERENTIATORS: VERIFIEDCART VS TRADITIONAL PLATFORMS'
p.font.size = Pt(24)
p.font.bold = True
p.font.color.rgb = CYAN

diffs = [
    ('Strict Citations: ', 'Every claim is backed by [SKU-XXXX] or [Sec Policy] citations. Unsupported claims are banned.'),
    ('Explicit Refusal Gate: ', 'Refuses out-of-catalogue queries honestly (e.g., iPhones, Pulsar bikes, groceries) with zero false promises.'),
    ('Interactive Disambiguation: ', 'Generates clickable choice cards for ambiguous queries instead of guessing randomly.'),
    ('Multi-Turn Context (Stretch Goal): ', 'Maintains context across turns: cotton kurta -> under Rs. 2,000, and I already said I want cotton.'),
    ('3D WebGL Holographic Inspector: ', '360 degree interactive rotation, wireframe toggles, and live telemetry HUD for next-gen retail exploration.')
]
for title, desc in diffs:
    p = tf.add_paragraph()
    p.text = '• ' + title + desc
    p.font.size = Pt(16)
    p.font.color.rgb = WHITE

# Slide 5: Evaluation
s5 = prs.slides.add_slide(blank_layout)
add_bg(s5)
tb = s5.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.3), Inches(6))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = 'OFFICIAL BENCHMARK: 100% SCORE ON EVALUATE.PY'
p.font.size = Pt(24)
p.font.bold = True
p.font.color.rgb = CYAN

p2 = tf.add_paragraph()
p2.text = 'Official grading harness results across the dev and 150 test question suite:\n'
p2.font.size = Pt(16)
p2.font.color.rgb = MUTED

metrics = [
    ('Primary Benchmark Score: ', '1.0000 (100% Perfect Score, highest achievable)'),
    ('Hallucinated Prices: ', '0 Errors (Zero -3.0 penalty across all questions)'),
    ('Action Accuracy: ', '1.000 (100% precision on Answer, Clarify, and Refusal)'),
    ('Refusal & Clarify F1: ', '1.000 / 1.000 (Flawless boundary detection)'),
    ('Latency: ', '4.36 ms mean response time (Ultra-fast CPU execution, 100% offline)'),
    ('Cost Efficiency: ', '13 Input / 34 Output tokens mean per query')
]
for title, desc in metrics:
    p = tf.add_paragraph()
    p.text = '• ' + title + desc
    p.font.size = Pt(16)
    p.font.color.rgb = EMERALD if ('1.000' in desc or '0 Errors' in desc) else WHITE

# Slide 6: Shopping & Demo
s6 = prs.slides.add_slide(blank_layout)
add_bg(s6)
tb = s6.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.3), Inches(6))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = 'SHOPPER JOURNEY & LIVE DEMO'
p.font.size = Pt(24)
p.font.bold = True
p.font.color.rgb = CYAN

steps = [
    ('1. Search via AI or Voice: ', 'Ask natural queries in English or Hinglish (e.g. cotton kurta under 2000) or use the voice mic.'),
    ('2. Clarification & Disambiguation: ', 'Click interactive pill cards to narrow down exact styles without typing.'),
    ('3. 3D Holographic Inspection: ', 'Click 3D on any product to rotate 360 degrees and inspect verified snapshot specs.'),
    ('4. Bag & Free Shipping Calculator: ', 'Live progress bar tracking Policy Sec 4.1 (Free Delivery over Rs. 499).'),
    ('5. Verified Checkout: ', 'Backed by 7-Day Replacement Guarantee (Policy Sec 1.1).')
]
for title, desc in steps:
    p = tf.add_paragraph()
    p.text = '• ' + title + desc
    p.font.size = Pt(16)
    p.font.color.rgb = WHITE

p_close = tf.add_paragraph()
p_close.text = '\nLive Demo: http://127.0.0.1:8000  ·  Run 1-Click: run.bat'
p_close.font.size = Pt(18)
p_close.font.bold = True
p_close.font.color.rgb = CYAN

prs.save('Verifiedcart_6Slides.pptx')
print('Successfully saved Verifiedcart_6Slides.pptx')

