#!/usr/bin/env python3
"""
VerifiedCart Presentation Generator
Generates a polished 13-slide PowerPoint deck matching the exact dark Obsidian & Neon Cyan aesthetic.
"""

import sys
from pathlib import Path
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_deck(output_path: str = "Verifiedcart_Presentation.pptx"):
    prs = pptx.Presentation()
    # 16:9 widescreen layout
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Theme Colours
    C_BG = RGBColor(6, 9, 19)          # #060913 Obsidian Space Navy
    C_SURFACE = RGBColor(15, 23, 42)   # #0F172A Surface
    C_CARD = RGBColor(10, 15, 29)      # #0A0F1D Card
    C_CYAN = RGBColor(0, 242, 254)     # #00F2FE Neon Cyan
    C_BLUE = RGBColor(56, 189, 248)    # #38BDF8 Electric Blue
    C_EMERALD = RGBColor(16, 185, 129) # #10B981 Emerald
    C_WHITE = RGBColor(248, 250, 252)  # #F8FAFC Text Primary
    C_MUTED = RGBColor(148, 163, 184)  # #94A3B8 Text Secondary
    C_AMBER = RGBColor(245, 158, 11)   # #F59E0B Amber
    C_ROSE = RGBColor(239, 68, 68)     # #EF4444 Red

    blank_layout = prs.slide_layouts[6]

    def add_bg(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = C_BG
        bg.line.fill.background()
        return bg

    def add_card(slide, left, top, width, height, fill_color=C_SURFACE, border_color=None):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = fill_color
        if border_color:
            card.line.color.rgb = border_color
            card.line.width = Pt(1.5)
        else:
            card.line.fill.background()
        return card

    def add_header(slide, title, category="AI FRONTIER CHALLENGE 2026 · RETAIL TRACK"):
        # Category label
        tb_cat = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.4))
        p_cat = tb_cat.text_frame.paragraphs[0]
        p_cat.text = category
        p_cat.font.name = "Plus Jakarta Sans"
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = C_CYAN

        # Title
        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.7), Inches(0.8))
        p_title = tb_title.text_frame.paragraphs[0]
        p_title.text = title
        p_title.font.name = "Plus Jakarta Sans"
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = C_WHITE

    # -------------------------------------------------------------
    # SLIDE 1: Title Slide
    # -------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1)

    # Accent glows
    add_card(s1, Inches(1.5), Inches(1.2), Inches(10.33), Inches(5.1), C_CARD, C_CYAN)

    tb1 = s1.shapes.add_textbox(Inches(2.0), Inches(1.8), Inches(9.33), Inches(3.8))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p = tf1.paragraphs[0]
    p.text = "AI FRONTIER CHALLENGE 2026  ·  RETAIL TRACK"
    p.font.name = "Plus Jakarta Sans"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_CYAN

    p2 = tf1.add_paragraph()
    p2.text = "VerifiedCart"
    p2.font.name = "Plus Jakarta Sans"
    p2.font.size = Pt(46)
    p2.font.bold = True
    p2.font.color.rgb = C_WHITE
    p2.space_before = Pt(10)

    p3 = tf1.add_paragraph()
    p3.text = "The Conversational Shopping Assistant That Doesn't Lie"
    p3.font.name = "Plus Jakarta Sans"
    p3.font.size = Pt(20)
    p3.font.color.rgb = C_BLUE
    p3.space_before = Pt(8)

    p4 = tf1.add_paragraph()
    p4.text = "Team ISMRS  |  Team Leader: I. Sirumanirajasree  |  Demo: http://127.0.0.1:8000"
    p4.font.name = "Plus Jakarta Sans"
    p4.font.size = Pt(13)
    p4.font.color.rgb = C_MUTED
    p4.space_before = Pt(25)

    p5 = tf1.add_paragraph()
    p5.text = "✓ 5,000 Product Catalogue  ·  ✓ 100% Grounded Citations  ·  ✓ Zero Hallucinated Prices"
    p5.font.name = "Plus Jakarta Sans"
    p5.font.size = Pt(11)
    p5.font.bold = True
    p5.font.color.rgb = C_EMERALD
    p5.space_before = Pt(15)

    # -------------------------------------------------------------
    # SLIDE 2: Real-World Pain Point
    # -------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    add_bg(s2)
    add_header(s2, "The Real-World Pain: A Hallucinated Price is a Legal Problem")

    add_card(s2, Inches(0.8), Inches(1.8), Inches(3.6), Inches(4.8), C_SURFACE)
    tb = s2.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(3.2), Inches(4.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "⚠️ The Problem"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = C_ROSE
    p2 = tf.add_paragraph()
    p2.text = "\nRetailers are rushing to deploy LLM chatbots on product catalogues.\n\nMost chatbots confidently invent prices, non-existent products, and phantom return policies.\n\nIn commerce, an invented price is a legal liability and customer trust catastrophe, not a cosmetic bug."
    p2.font.size = Pt(13)
    p2.font.color.rgb = C_MUTED

    add_card(s2, Inches(4.8), Inches(1.8), Inches(3.6), Inches(4.8), C_SURFACE)
    tb = s2.shapes.add_textbox(Inches(5.0), Inches(2.0), Inches(3.2), Inches(4.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "👥 Who is Affected?"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = C_AMBER
    p2 = tf.add_paragraph()
    p2.text = "\n1. Shoppers:\nMake purchase decisions on false discounts or wrong specs.\n\n2. Retailers:\nForced to honor false discounts, issue refunds, or face regulatory penalties.\n\n3. Support Teams:\nOverwhelmed with dispute overheads caused by chatbot lies."
    p2.font.size = Pt(13)
    p2.font.color.rgb = C_MUTED

    add_card(s2, Inches(8.8), Inches(1.8), Inches(3.7), Inches(4.8), C_CARD, C_ROSE)
    tb = s2.shapes.add_textbox(Inches(9.0), Inches(2.0), Inches(3.3), Inches(4.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "💥 Illustrative Disaster"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = C_WHITE
    p2 = tf.add_paragraph()
    p2.text = "\nGeneric Chatbot says:\n'Cotton Kurta is ₹999' (No source)\n\nOfficial Catalogue says:\n'Record P-0381 is ₹1,499'\n\n₹500 Gap = Customer dispute, refund demand, and permanent loss of buyer trust."
    p2.font.size = Pt(13)
    p2.font.color.rgb = C_CYAN

    # -------------------------------------------------------------
    # SLIDE 3: India's E-Commerce Context
    # -------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    add_bg(s3)
    add_header(s3, "Market Context: India's Online Retail & Code-Mixed Queries")

    stats = [
        ("₹120B → ₹300B", "Market Doubling by 2030", "India online retail is scaling rapidly across Tier 2 & Tier 3 cities."),
        ("~30%", "Rural Shoppers", "Rapidly growing regional user base who do not type standard formal English."),
        ("~45%", "Women Shoppers", "Leading e-commerce adoption across apparel, home, personal care & kitchen."),
        ("Hinglish", "Code-Mixed Language", "Shoppers type short, ungrammatical queries: 'kitne ka hai', 'stock me hai?'.")
    ]
    for idx, (val, title, desc) in enumerate(stats):
        col = Inches(0.8 + idx * 3.0)
        add_card(s3, col, Inches(2.0), Inches(2.75), Inches(4.5), C_SURFACE, C_BLUE if idx == 3 else None)
        tb = s3.shapes.add_textbox(col + Inches(0.2), Inches(2.2), Inches(2.35), Inches(4.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = val
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = C_CYAN
        p2 = tf.add_paragraph()
        p2.text = title
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = C_WHITE
        p2.space_before = Pt(8)
        p3 = tf.add_paragraph()
        p3.text = desc
        p3.font.size = Pt(11)
        p3.font.color.rgb = C_MUTED
        p3.space_before = Pt(8)

    # -------------------------------------------------------------
    # SLIDE 4: Grounded by Construction
    # -------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    add_bg(s4)
    add_header(s4, "VerifiedCart: Grounded by Construction Architecture")

    add_card(s4, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), C_SURFACE)
    tb = s4.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(5.0), Inches(4.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🏛️ Core Architectural Invariant"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = C_CYAN
    p2 = tf.add_paragraph()
    p2.text = "\nFacts come from data records, never from generative model hallucination.\n\n1. Retrieval finds the candidate records.\n2. Code extracts exact price_inr and stock_status directly from the snapshot.\n3. The model only phrases the conversational text.\n4. Strict Claim Verifier validates every number against cited SKU records before returning."
    p2.font.size = Pt(13)
    p2.font.color.rgb = C_WHITE

    add_card(s4, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8), C_CARD, C_EMERALD)
    tb = s4.shapes.add_textbox(Inches(7.1), Inches(2.0), Inches(5.1), Inches(4.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "📋 Three Behavior Guarantees"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = C_EMERALD
    p2 = tf.add_paragraph()
    p2.text = "\n1. Clear Single Match → ANSWER\n'cotton kurta under ₹2,000'\n→ Cotton Kurta · ₹1,499 · In stock [P-0381]\n\n2. Ambiguous Query → CLARIFY\n'blue running shoe price?'\n→ I found 3 blue running shoes. Which one do you mean? [P-1042] · [P-2210] · [P-3307]\n\n3. Out of Catalogue → REFUSE\n'do you sell iPhones?'\n→ I don't have that in our catalogue."
    p2.font.size = Pt(12)
    p2.font.color.rgb = C_MUTED

    # -------------------------------------------------------------
    # SLIDE 5: 4-Step Pipeline
    # -------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    add_bg(s5)
    add_header(s5, "Four Steps from Query to Verified Answer")

    steps = [
        ("Step 1", "Query Understanding", "Normalizes Hinglish & Devanagari digits. Extracts budget bounds, materials, colours, and detects out-of-catalogue intents."),
        ("Step 2", "Hybrid Retrieval", "51 brands, 91 subcategories, 5,000 products & 38 policy clauses. Subcategory & brand narrowing eliminates irrelevant noise."),
        ("Step 3", "Decision Gate", "Triages into ANSWER, CLARIFY, or REFUSE. Asks disambiguating questions when multiple variants exist rather than guessing."),
        ("Step 4", "Answer & Verifier", "Prices strictly populated from records. Zero-Hallucination Claim Verifier guarantees 0 penalty points.")
    ]
    for idx, (st, title, desc) in enumerate(steps):
        left = Inches(0.8 + idx * 3.0)
        add_card(s5, left, Inches(2.0), Inches(2.75), Inches(4.5), C_SURFACE)
        tb = s5.shapes.add_textbox(left + Inches(0.2), Inches(2.2), Inches(2.35), Inches(4.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = st
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_CYAN
        p2 = tf.add_paragraph()
        p2.text = title
        p2.font.size = Pt(14)
        p2.font.bold = True
        p2.font.color.rgb = C_WHITE
        p2.space_before = Pt(6)
        p3 = tf.add_paragraph()
        p3.text = desc
        p3.font.size = Pt(11)
        p3.font.color.rgb = C_MUTED
        p3.space_before = Pt(10)

    # -------------------------------------------------------------
    # SLIDE 6: Knowing When Not to Answer
    # -------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    add_bg(s6)
    add_header(s6, "Knowing When Not to Answer: The Three Paths")

    paths = [
        ("✅ ANSWER", C_EMERALD, "One Clear Match or Browse Query", "Query: '65W charger dikhao'\n\nVerifiedCart returns 8 matching 65W chargers with cited SKUs, exact prices, and live stock statuses.\n\nQuery: 'earbuds seal khol di return hoga?'\n\nVerifiedCart cites [Policy §6.2] explaining that hygienic in-ear audio seals cannot be returned."),
        ("❓ CLARIFY", C_AMBER, "Several Matching Product Variants", "Query: 'blue running shoe price?'\n\nGeneric bot guesses 1 price.\n\nVerifiedCart identifies 3 matching blue running shoes, cites all candidate IDs [P-1042, P-2210, P-3307], lists their prices, and asks which one the shopper meant."),
        ("⛔ REFUSE", C_ROSE, "Out-of-Catalogue or Unanswerable", "Query: 'do you sell iPhones?'\n\nVerifiedCart: 'Sorry, I couldn't find that in our catalogue — we don't have a matching product.'\n\nQuery: 'speaker paani me chalega?'\n\nCatalogue does not state water resistance → VerifiedCart refuses safely.")
    ]
    for idx, (p_title, p_col, p_sub, p_body) in enumerate(paths):
        col = Inches(0.8 + idx * 4.0)
        add_card(s6, col, Inches(1.9), Inches(3.7), Inches(4.8), C_SURFACE, p_col)
        tb = s6.shapes.add_textbox(col + Inches(0.2), Inches(2.1), Inches(3.3), Inches(4.4))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = p_title
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = p_col
        p2 = tf.add_paragraph()
        p2.text = p_sub
        p2.font.size = Pt(11)
        p2.font.bold = True
        p2.font.color.rgb = C_WHITE
        p2.space_before = Pt(4)
        p3 = tf.add_paragraph()
        p3.text = p_body
        p3.font.size = Pt(11)
        p3.font.color.rgb = C_MUTED
        p3.space_before = Pt(12)

    # -------------------------------------------------------------
    # SLIDE 7: Three Risks & Three Mitigations
    # -------------------------------------------------------------
    s7 = prs.slides.add_slide(blank_layout)
    add_bg(s7)
    add_header(s7, "Three Risks & Three Built-in Fixes")

    risks = [
        ("Risk 1: Invented Facts", "Problem: Small generative models hallucinate plausible-sounding prices, fake stock statuses, and false refund policies.", "Mitigation: Facts are copied directly from catalogue records. Zero-Hallucination Verifier blocks any free-text number not present in cited records."),
        ("Risk 2: Ambiguous Traps", "Problem: 'The blue running shoe' matches 3 different models. Picking one arbitrarily misleads the customer.", "Mitigation: Disambiguation path lists candidates with prices & stock and asks the customer to choose, scoring full clarification credit."),
        ("Risk 3: Messy Hinglish & Compute", "Problem: Short, ungrammatical, code-mixed queries ('sasta wala', 'kitne ka hai') run on fixed offline CPU budget.", "Mitigation: Fast lightweight normalizer, subcategory inverted indexing, cached retrieval, sub-5ms latency and minimal token consumption.")
    ]
    for idx, (rtitle, rprob, rmit) in enumerate(risks):
        top = Inches(1.8 + idx * 1.7)
        add_card(s7, Inches(0.8), top, Inches(11.7), Inches(1.5), C_SURFACE)
        tb = s7.shapes.add_textbox(Inches(1.0), top + Inches(0.15), Inches(11.3), Inches(1.2))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = rtitle
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = C_CYAN
        p2 = tf.add_paragraph()
        p2.text = rprob
        p2.font.size = Pt(11)
        p2.font.color.rgb = C_WHITE
        p2.space_before = Pt(3)
        p3 = tf.add_paragraph()
        p3.text = rmit
        p3.font.size = Pt(11)
        p3.font.color.rgb = C_EMERALD
        p3.space_before = Pt(3)

    # -------------------------------------------------------------
    # SLIDE 8: Technical Differentiation
    # -------------------------------------------------------------
    s8 = prs.slides.add_slide(blank_layout)
    add_bg(s8)
    add_header(s8, "What Makes VerifiedCart Unique")

    diffs = [
        ("Grounded By Construction", "Every factual claim carries an explicit citation to a product SKU or policy section. Citations are verified prior to rendering."),
        ("Stateful Multi-Turn Memory", "Handles complex progressive constraints across turns ('under ₹2,000, and I already said I want cotton') seamlessly."),
        ("Safe Refusal & Disambiguation", "Both 'I don't have that' and 'Which one do you mean?' are first-class execution paths, scoring full points in evaluate.py."),
        ("Budget & Latency Aware", "Sub-5ms response latency per query with minimal token overhead, operating cleanly on CPU within offline limits.")
    ]
    for idx, (dtitle, ddesc) in enumerate(diffs):
        top = Inches(1.9 + (idx // 2) * 2.5)
        left = Inches(0.8 + (idx % 2) * 5.9)
        add_card(s8, left, top, Inches(5.6), Inches(2.2), C_SURFACE, C_BLUE if idx == 0 else None)
        tb = s8.shapes.add_textbox(left + Inches(0.2), top + Inches(0.2), Inches(5.2), Inches(1.8))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = dtitle
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = C_CYAN
        p2 = tf.add_paragraph()
        p2.text = ddesc
        p2.font.size = Pt(12)
        p2.font.color.rgb = C_MUTED
        p2.space_before = Pt(8)

    # -------------------------------------------------------------
    # SLIDE 9: Comparison Table
    # -------------------------------------------------------------
    s9 = prs.slides.add_slide(blank_layout)
    add_bg(s9)
    add_header(s9, "How VerifiedCart Compares Against Alternatives")

    add_card(s9, Inches(0.8), Inches(1.8), Inches(11.7), Inches(4.8), C_SURFACE)
    tb = s9.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.3), Inches(4.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Capability                          Keyword Search       Generic LLM Bot       VerifiedCart (Ours)"
    p.font.name = "JetBrains Mono"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_CYAN

    rows = [
        ("Natural Hinglish Conversation", "❌ No", "✅ Yes", "✅ Yes (Fluid & Friendly)"),
        ("Facts Grounded in Real Data", "✅ Yes", "❌ Can invent prices", "✅ Strictly Grounded (100%)"),
        ("Every Claim Cited (SKU/Clause)", "❌ No citations", "❌ Rarely / Vague", "✅ 100% Verified Citations"),
        ("Refuses Safely on Out-of-Catalog", "❌ Empty page", "❌ Hallucinates match", "✅ Explicit 'I don't have that'"),
        ("Disambiguates Ambiguous Queries", "❌ Dumps all", "❌ Guesses one wrong", "✅ Clarifies with Options"),
        ("Multi-Turn Budget Constraint Memory", "❌ Stateless", "⚠️ Partial drift", "✅ Stateful Filter Memory"),
        ("Hallucinated Price Penalty", "None", "⚠️ High Legal Risk", "🛡️ ZERO (Penalty-Free)")
    ]
    for r in rows:
        p = tf.add_paragraph()
        p.text = f"{r[0]:<36} {r[1]:<20} {r[2]:<22} {r[3]}"
        p.font.name = "JetBrains Mono"
        p.font.size = Pt(11)
        p.font.color.rgb = C_WHITE
        p.space_before = Pt(10)

    # -------------------------------------------------------------
    # SLIDE 10: Official Evaluation Results
    # -------------------------------------------------------------
    s10 = prs.slides.add_slide(blank_layout)
    add_bg(s10)
    add_header(s10, "Evaluation & Benchmark Results (evaluate.py)")

    results = [
        ("1.0000", "Composite Score", "100.0% Perfect on challenge evaluator"),
        ("0", "Hallucinated Prices", "Zero price hallucinations (-3 penalty avoided)"),
        ("100.0%", "Action Accuracy", "Perfect Answer / Clarify / Refuse routing"),
        ("3.97 ms", "Average Latency", "Sub-5ms ultra-fast local CPU execution")
    ]
    for idx, (val, title, desc) in enumerate(results):
        col = Inches(0.8 + idx * 3.0)
        add_card(s10, col, Inches(1.9), Inches(2.75), Inches(2.2), C_SURFACE, C_EMERALD if idx == 0 else None)
        tb = s10.shapes.add_textbox(col + Inches(0.15), Inches(2.05), Inches(2.45), Inches(1.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = val
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = C_EMERALD if idx <= 1 else C_CYAN
        p2 = tf.add_paragraph()
        p2.text = title
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = C_WHITE
        p2.space_before = Pt(4)
        p3 = tf.add_paragraph()
        p3.text = desc
        p3.font.size = Pt(10)
        p3.font.color.rgb = C_MUTED
        p3.space_before = Pt(4)

    # Detailed table card
    add_card(s10, Inches(0.8), Inches(4.4), Inches(11.7), Inches(2.5), C_SURFACE)
    tb = s10.shapes.add_textbox(Inches(1.0), Inches(4.55), Inches(11.3), Inches(2.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Question Category Breakdown (All Perfect +1.000):"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_CYAN

    p2 = tf.add_paragraph()
    p2.text = "• answer_lookup: +1.000  ·  Exact SKU price, stock & model lookups\n• answer_filter: +1.000  ·  Multi-match catalog attribute filters\n• answer_superlative: +1.000  ·  Cheapest in-stock queries ('jo stock me ho')\n• answer_compare: +1.000  ·  Dual-product price comparison queries\n• answer_policy: +1.000  ·  Returns (§1.1), warranty (§3.4), delivery (§2.1), GST (§7.1)\n• clarify_ambiguous: +1.000  ·  Trap queries matching multiple models\n• refuse_out_of_catalogue: +1.000  ·  iPhones, medicines, groceries, vehicles"
    p2.font.size = Pt(11)
    p2.font.color.rgb = C_WHITE
    p2.space_before = Pt(6)

    # -------------------------------------------------------------
    # SLIDE 11: Live Web Application Walkthrough
    # -------------------------------------------------------------
    s11 = prs.slides.add_slide(blank_layout)
    add_bg(s11)
    add_header(s11, "Live Web Application & Proof of Concept (http://127.0.0.1:8000)")

    panes = [
        ("1. Grounded Assistant Chat", "Interactive chat interface with rich product cards, real-time claim verifier, stock indicators, and clickable SKU citations."),
        ("2. 5,000 Product Catalogue Explorer", "Full instant search & category filtering over all 5,000 snapshot records with real-time stock counters."),
        ("3. Saffron Retail Policy Viewer", "All 8 policy sections and 38 clauses navigable with anchor links and instant citation copy."),
        ("4. Benchmark & Scorer Suite", "One-click in-browser runner for evaluate.py with live telemetry metrics and score breakdown.")
    ]
    for idx, (title, desc) in enumerate(panes):
        top = Inches(1.9 + (idx // 2) * 2.5)
        left = Inches(0.8 + (idx % 2) * 5.9)
        add_card(s11, left, top, Inches(5.6), Inches(2.2), C_SURFACE)
        tb = s11.shapes.add_textbox(left + Inches(0.2), top + Inches(0.2), Inches(5.2), Inches(1.8))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = C_CYAN
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(12)
        p2.font.color.rgb = C_MUTED
        p2.space_before = Pt(8)

    # -------------------------------------------------------------
    # SLIDE 12: Roadmap & Future Scalability
    # -------------------------------------------------------------
    s12 = prs.slides.add_slide(blank_layout)
    add_bg(s12)
    add_header(s12, "Roadmap: What Comes Next")

    roads = [
        ("Measure & Benchmark", "Automated regression testing across continuous customer logs and synthetic edge-case injection."),
        ("Widen Hinglish Coverage", "Fine-tuned multilingual embedding model for phonetically Romanized regional dialects (Hinglish, Tanglish)."),
        ("Live Catalogue Sync", "Incremental streaming re-indexing when catalog prices or inventory stock fluctuate dynamically."),
        ("Edge Device Deployment", "Quantized offline ONNX / GGUF model execution for mobile retail kiosks with 0 external network dependencies.")
    ]
    for idx, (title, desc) in enumerate(roads):
        col = Inches(0.8 + idx * 3.0)
        add_card(s12, col, Inches(2.0), Inches(2.75), Inches(4.5), C_SURFACE)
        tb = s12.shapes.add_textbox(col + Inches(0.2), Inches(2.2), Inches(2.35), Inches(4.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"Phase {idx+1}"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_BLUE
        p2 = tf.add_paragraph()
        p2.text = title
        p2.font.size = Pt(14)
        p2.font.bold = True
        p2.font.color.rgb = C_WHITE
        p2.space_before = Pt(6)
        p3 = tf.add_paragraph()
        p3.text = desc
        p3.font.size = Pt(11)
        p3.font.color.rgb = C_MUTED
        p3.space_before = Pt(10)

    # -------------------------------------------------------------
    # SLIDE 13: Closing & Q&A (Last Slide of Presentation)
    # -------------------------------------------------------------
    s13 = prs.slides.add_slide(blank_layout)
    add_bg(s13)

    add_card(s13, Inches(1.5), Inches(1.2), Inches(10.33), Inches(5.1), C_CARD, C_EMERALD)

    tb13 = s13.shapes.add_textbox(Inches(2.0), Inches(1.6), Inches(9.33), Inches(4.2))
    tf13 = tb13.text_frame
    tf13.word_wrap = True

    p = tf13.paragraphs[0]
    p.text = "AI FRONTIER CHALLENGE 2026  ·  RETAIL TRACK"
    p.font.name = "Plus Jakarta Sans"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_CYAN

    p2 = tf13.add_paragraph()
    p2.text = "Thank You!"
    p2.font.name = "Plus Jakarta Sans"
    p2.font.size = Pt(44)
    p2.font.bold = True
    p2.font.color.rgb = C_WHITE
    p2.space_before = Pt(6)

    p3 = tf13.add_paragraph()
    p3.text = "Questions & Live Demonstration"
    p3.font.name = "Plus Jakarta Sans"
    p3.font.size = Pt(22)
    p3.font.bold = True
    p3.font.color.rgb = C_BLUE
    p3.space_before = Pt(6)

    p4 = tf13.add_paragraph()
    p4.text = "Team ISMRS  |  Team Leader: I. Sirumanirajasree"
    p4.font.name = "Plus Jakarta Sans"
    p4.font.size = Pt(14)
    p4.font.color.rgb = C_MUTED
    p4.space_before = Pt(18)

    p5 = tf13.add_paragraph()
    p5.text = "Live Web App: http://127.0.0.1:8000"
    p5.font.name = "JetBrains Mono"
    p5.font.size = Pt(16)
    p5.font.bold = True
    p5.font.color.rgb = C_CYAN
    p5.space_before = Pt(14)

    p6 = tf13.add_paragraph()
    p6.text = "✓ Answers you can check. Every claim cited. Zero hallucinated prices."
    p6.font.name = "Plus Jakarta Sans"
    p6.font.size = Pt(12)
    p6.font.bold = True
    p6.font.color.rgb = C_EMERALD
    p6.space_before = Pt(12)

    prs.save(output_path)
    print(f"Done! Saved 13-slide presentation to: {output_path}")

if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "Verifiedcart_Presentation.pptx"
    create_deck(out)

