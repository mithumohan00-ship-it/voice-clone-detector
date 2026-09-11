import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# -------------------------------------------------------------
# PALETTE & CONSTANTS
# -------------------------------------------------------------
FONT_TITLE = "Helvetica"
FONT_BODY  = "Arial"

C_BG_DARK        = RGBColor(10, 14, 23)      # #0A0E17 Deep Obsidian Navy
C_CARD_BG        = RGBColor(17, 28, 51)      # #111C33 Card Slate Navy
C_CARD_BORDER    = RGBColor(30, 46, 78)      # #1E2E4E Subtle Border
C_CARD_ACTIVE    = RGBColor(24, 40, 72)      # Highlighted Card Fill

C_CYAN           = RGBColor(0, 210, 255)     # #00D2FF Electric Cyan Accent
C_BLUE_ACCENT    = RGBColor(59, 130, 246)    # #3B82F6 Vibrant Cobalt
C_BLUE_DARK      = RGBColor(30, 58, 138)     # #1E3A8A Dark Blue

C_SPOOF_RED      = RGBColor(239, 68, 68)     # #EF4444 Alert Red
C_SPOOF_BG       = RGBColor(45, 20, 28)      # Red Tinted Dark Card
C_SPOOF_BORDER   = RGBColor(127, 29, 29)     # Dark Red Border

C_BONAFIDE_GREEN = RGBColor(16, 185, 129)    # #10B981 Verified Green
C_BONAFIDE_BG    = RGBColor(16, 38, 30)      # Green Tinted Dark Card
C_BONAFIDE_BORDER= RGBColor(6, 95, 70)       # Dark Green Border

C_AMBER          = RGBColor(245, 158, 11)    # #F59E0B Warning / Notice
C_AMBER_BG       = RGBColor(45, 35, 15)      # Amber Tinted Dark Card
C_AMBER_BORDER   = RGBColor(146, 64, 14)     # Dark Amber Border

C_TEXT_WHITE     = RGBColor(248, 250, 252)   # #F8FAFC Primary White
C_TEXT_MUTED     = RGBColor(148, 163, 184)   # #94A3B8 Slate Gray Body
C_TEXT_DIM       = RGBColor(100, 116, 139)   # #64748B Secondary Detail

SLIDE_WIDTH  = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

def create_deck():
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    blank_layout = prs.slide_layouts[6]
    return prs, blank_layout

def apply_dark_background(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_BG_DARK
    bg.line.fill.background()
    return bg

def add_header(slide, title_text, category_text, slide_num):
    top_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.04))
    top_line.fill.solid()
    top_line.fill.fore_color.rgb = C_CYAN
    top_line.line.fill.background()

    tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.48), Inches(7.5), Inches(0.3))
    tf_tag = tag_box.text_frame
    tf_tag.word_wrap = True
    tf_tag.margin_left = tf_tag.margin_top = tf_tag.margin_right = tf_tag.margin_bottom = 0
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = f"SMART INDIA HACKATHON 2026  ·  {category_text.upper()}"
    p_tag.font.name = FONT_BODY
    p_tag.font.size = Pt(9.5)
    p_tag.font.bold = True
    p_tag.font.color.rgb = C_CYAN

    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.78), Inches(9.8), Inches(0.7))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    tf_title.margin_left = tf_title.margin_top = tf_title.margin_right = tf_title.margin_bottom = 0
    p_t = tf_title.paragraphs[0]
    p_t.text = title_text
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(21)
    p_t.font.bold = True
    p_t.font.color.rgb = C_TEXT_WHITE

    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.633), Inches(0.48), Inches(1.9), Inches(0.52))
    badge.fill.solid()
    badge.fill.fore_color.rgb = C_CARD_BG
    badge.line.color.rgb = C_CARD_BORDER
    badge.line.width = Pt(1)
    tf_b = badge.text_frame
    tf_b.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_b = tf_b.paragraphs[0]
    p_b.text = f"SIH-Aegis  |  {slide_num}/10"
    p_b.alignment = PP_ALIGN.CENTER
    p_b.font.name = FONT_BODY
    p_b.font.size = Pt(10.5)
    p_b.font.bold = True
    p_b.font.color.rgb = C_TEXT_MUTED

def set_speaker_notes(slide, meaning, script, qa_pairs):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.word_wrap = True
    
    full_text = []
    full_text.append("=== SLIDE MEANING & OBJECTIVE ===")
    full_text.append(meaning.strip())
    full_text.append("\n=== PRESENTER SCRIPT (20-40 SECONDS) ===")
    full_text.append(script.strip())
    full_text.append("\n=== LIKELY JUDGE QUESTIONS & DEFENSE ===")
    for q, a in qa_pairs:
        full_text.append(f"\nQ: {q}")
        full_text.append(f"A: {a}")
        
    tf.text = "\n".join(full_text)

def create_card(slide, left, top, width, height, border_color=C_CARD_BORDER, bg_color=C_CARD_BG):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1.2)
    return card

# -------------------------------------------------------------
# SLIDE BUILDERS
# -------------------------------------------------------------

def build_slide_1(prs, layout):
    slide = prs.slides.add_slide(layout)
    apply_dark_background(slide)

    # Decorative Cyan Accent Glow Bar
    glow_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.1), Inches(0.12), Inches(5.1))
    glow_bar.fill.solid()
    glow_bar.fill.fore_color.rgb = C_CYAN
    glow_bar.line.fill.background()

    # Top Category Pill
    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(1.1), Inches(5.2), Inches(0.42))
    pill.fill.solid()
    pill.fill.fore_color.rgb = C_CARD_BG
    pill.line.color.rgb = C_CYAN
    pill.line.width = Pt(1)
    tf_p = pill.text_frame
    tf_p.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_pill = tf_p.paragraphs[0]
    p_pill.text = "SMART INDIA HACKATHON 2026 · ROUND 1 PRESENTATION"
    p_pill.font.name = FONT_BODY
    p_pill.font.size = Pt(9.5)
    p_pill.font.bold = True
    p_pill.font.color.rgb = C_CYAN

    # Project Brand Name
    name_box = slide.shapes.add_textbox(Inches(1.2), Inches(1.65), Inches(10.5), Inches(1.1))
    tf_n = name_box.text_frame
    tf_n.word_wrap = True
    p_name = tf_n.paragraphs[0]
    p_name.text = "SIH-Aegis"
    p_name.font.name = FONT_TITLE
    p_name.font.size = Pt(46)
    p_name.font.bold = True
    p_name.font.color.rgb = C_TEXT_WHITE

    # Full Project Title
    title_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.75), Inches(11.0), Inches(1.2))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    p_title = tf_t.paragraphs[0]
    p_title.text = "AI-Powered Real-Time Detection and Prevention of\nVoice Cloning Impersonation Attacks"
    p_title.font.name = FONT_TITLE
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = C_CYAN

    # Core Value Tagline
    tagline_box = slide.shapes.add_textbox(Inches(1.2), Inches(4.05), Inches(11.0), Inches(0.6))
    tf_tag = tagline_box.text_frame
    tf_tag.word_wrap = True
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = "Transforming voice authenticity from a subjective human trust assumption into a measurable, objective security signal."
    p_tag.font.name = FONT_BODY
    p_tag.font.size = Pt(13)
    p_tag.font.color.rgb = C_TEXT_MUTED

    # 4 Metadata Information Badges at Bottom
    cards_info = [
        ("PROBLEM STATEMENT", "SIH26104 (ID: 104)"),
        ("ORGANIZATION", "AICTE · Blockchain & Cyber Security"),
        ("DEVELOPMENT STAGE", "Verified Working Baseline + Stage 1 Architecture"),
        ("TEAM IDENTITY", "SIH-Aegis · Software Track")
    ]
    card_w = Inches(2.78)
    card_h = Inches(1.15)
    top_pos = Inches(4.95)
    for i, (label, val) in enumerate(cards_info):
        left_pos = Inches(1.2) + i * Inches(2.95)
        c = create_card(slide, left_pos, top_pos, card_w, card_h, C_CARD_BORDER, C_CARD_BG)
        tf_c = c.text_frame
        tf_c.word_wrap = True
        tf_c.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_c.margin_left = Inches(0.18)
        
        p1 = tf_c.paragraphs[0]
        p1.text = label
        p1.font.name = FONT_BODY
        p1.font.size = Pt(8.5)
        p1.font.bold = True
        p1.font.color.rgb = C_CYAN
        p1.space_after = Pt(2)
        
        p2 = tf_c.add_paragraph()
        p2.text = val
        p2.font.name = FONT_BODY
        p2.font.size = Pt(10)
        p2.font.bold = True
        p2.font.color.rgb = C_TEXT_WHITE

    meaning = (
        "Establishes SIH-Aegis as a serious, technically disciplined cybersecurity initiative for SIH 2026. "
        "Sets the core paradigm shift: turning subjective human auditory belief into an objective, measurable security signal."
    )
    script = (
        "Good day, respected judges. We are Team SIH-Aegis, presenting our solution for Problem Statement SIH26104: "
        "AI-Powered Real-Time Detection and Prevention of Voice Cloning Impersonation Attacks. "
        "Today, generative audio allows attackers to clone any human voice with just seconds of reference audio. "
        "Our mission is to turn voice authenticity from an easily manipulated human trust assumption into a measurable, cryptographic-grade security signal. "
        "We have a validated, working end-to-end prototype and a clear engineering roadmap toward lightweight, mobile-capable deployment."
    )
    qa = [
        ("Why voice cloning detection now?", 
         "Because generative voice synthesis models have crossed the human perceptual threshold. Anyone can harvest 5 seconds of audio from social media and bypass phone-based authorizations or extort families. Traditional authentication assumes that if it sounds like the person, it is the person. That assumption is now broken.")
    ]
    set_speaker_notes(slide, meaning, script, qa)


def build_slide_2(prs, layout):
    slide = prs.slides.add_slide(layout)
    apply_dark_background(slide)
    add_header(slide, 'Your Voice Can Be Copied: The Generative Threat Landscape', 'Threat Context', 2)

    # 5-Step Attack Vector Flow
    steps = [
        ("1. Real Speaker", "3-5 sec voice harvested\nfrom social media or call", C_CARD_BORDER, C_CARD_BG, C_TEXT_WHITE),
        ("2. Voice Cloning", "Zero-shot neural TTS\nmodels (XTTS, ElevenLabs)", C_CARD_BORDER, C_CARD_BG, C_CYAN),
        ("3. Synthetic Speech", "Vocal timbre, emotion &\nintonation replicated", C_CARD_BORDER, C_CARD_BG, C_AMBER),
        ("4. Impersonation", "CEO fraud, fake emergency,\nKYC & voice auth bypass", C_SPOOF_BORDER, C_SPOOF_BG, C_SPOOF_RED),
        ("5. Severe Impact", "Financial theft, identity\ncompromise & social fraud", C_SPOOF_BORDER, C_SPOOF_BG, C_SPOOF_RED)
    ]

    card_w = Inches(2.15)
    card_h = Inches(1.6)
    card_top = Inches(1.7)
    
    for i, (title, desc, border, bg, color) in enumerate(steps):
        left_pos = Inches(0.8) + i * Inches(2.4)
        c = create_card(slide, left_pos, card_top, card_w, card_h, border, bg)
        tf_c = c.text_frame
        tf_c.word_wrap = True
        tf_c.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_c.margin_left = Inches(0.14)
        
        p1 = tf_c.paragraphs[0]
        p1.text = title
        p1.font.name = FONT_BODY
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = color
        p1.space_after = Pt(4)
        
        p2 = tf_c.add_paragraph()
        p2.text = desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_MUTED

        # Add connecting arrow if not last
        if i < len(steps) - 1:
            arr_box = slide.shapes.add_textbox(left_pos + card_w, card_top + Inches(0.55), Inches(0.25), Inches(0.5))
            tf_a = arr_box.text_frame
            tf_a.margin_left = tf_a.margin_top = tf_a.margin_right = tf_a.margin_bottom = 0
            p_a = tf_a.paragraphs[0]
            p_a.text = "→"
            p_a.font.name = FONT_TITLE
            p_a.font.size = Pt(18)
            p_a.font.bold = True
            p_a.font.color.rgb = C_CYAN
            p_a.alignment = PP_ALIGN.CENTER

    # 3 Analytical Deep-Dive Cards at Bottom
    impact_cards = [
        ("ZERO-SHOT ACCESSIBILITY", 
         "Attackers no longer need hours of studio speech. Diffusion and autoregressive neural vocoders generate indistinguishable clones from as little as 3 seconds of noisy audio harvested from Instagram, WhatsApp voice notes, or recorded outbound phone inquiries.", 
         C_CARD_BORDER, C_CYAN),
        ("COGNITIVE & AUDITORY BYPASS", 
         "Human hearing evaluates identity primarily through pitch, rhythm, and familiar vocabulary. Humans are chemically and emotionally conditioned to trust familiar voices, causing an estimated >40% error rate when detecting deepfakes over telephone codecs.", 
         C_CARD_BORDER, C_AMBER),
        ("CRITICAL SYSTEMIC TARGETS", 
         "The attack surface spans high-value banking authorization calls, remote video KYC onboarding, call-center verbal OTP verifications, executive wire-transfer authorizations, and emotionally coercive 'emergency bail/kidnapping' scams targeting families.", 
         C_SPOOF_BORDER, C_SPOOF_RED)
    ]

    card_w2 = Inches(3.72)
    card_h2 = Inches(3.1)
    card_top2 = Inches(3.6)

    for i, (title, text, border, color) in enumerate(impact_cards):
        left_pos = Inches(0.8) + i * Inches(4.0)
        c = create_card(slide, left_pos, card_top2, card_w2, card_h2, border, C_CARD_BG)
        tf_c = c.text_frame
        tf_c.word_wrap = True
        tf_c.vertical_anchor = MSO_ANCHOR.TOP
        tf_c.margin_left = tf_c.margin_right = Inches(0.22)
        tf_c.margin_top = Inches(0.2)
        
        p1 = tf_c.paragraphs[0]
        p1.text = title
        p1.font.name = FONT_BODY
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = color
        p1.space_after = Pt(8)
        
        p2 = tf_c.add_paragraph()
        p2.text = text
        p2.font.name = FONT_BODY
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_MUTED

    meaning = (
        "Visualizes how audio cloning has transformed from an academic curiosity into an industrialized threat vector. "
        "Shows judges that the team understands the full exploit chain: from voice harvesting to social engineering and financial fraud."
    )
    script = (
        "To understand the urgency of SIH-Aegis, we must look at how the threat has evolved. "
        "An attacker no longer requires studio equipment; 3 to 5 seconds of voice harvested from a social media reel is enough. "
        "Zero-shot neural speech synthesis clones the victim's exact pitch, cadence, and accent. "
        "The attacker then calls a family member, a bank, or a corporate helpdesk. "
        "Human auditory perception fails here because we naturally trust what sounds familiar. "
        "This is resulting in massive financial fraud, fake emergency extortion, and account takeovers globally."
    )
    qa = [
        ("Why can't humans detect cloned voices over the phone?", 
         "Standard cellular codecs compress audio heavily, cutting off frequencies above 3.4 kHz (narrowband) or 7 kHz (wideband). This compression strips away subtle acoustic clues, making synthetic voices sound virtually identical to genuine voices to human ears, especially under stressful emergency scenarios.")
    ]
    set_speaker_notes(slide, meaning, script, qa)


def build_slide_3(prs, layout):
    slide = prs.slides.add_slide(layout)
    apply_dark_background(slide)
    add_header(slide, 'Can We Trust a Familiar Voice? Measuring the Authenticity Signal', 'Problem Paradigm', 3)

    # 2 Comparison Cards (Traditional Trust vs SIH-Aegis Model)
    col_w = Inches(5.7)
    col_h = Inches(2.9)
    col_top = Inches(1.65)

    # Traditional Model (Red)
    c1 = create_card(slide, Inches(0.8), col_top, col_w, col_h, C_SPOOF_BORDER, C_SPOOF_BG)
    tf1 = c1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = tf1.margin_right = Inches(0.25)
    tf1.margin_top = Inches(0.2)
    p1 = tf1.paragraphs[0]
    p1.text = "TRADITIONAL TRUST MODEL  (UNSUSTAINABLE)"
    p1.font.name = FONT_BODY
    p1.font.size = Pt(12)
    p1.font.bold = True
    p1.font.color.rgb = C_SPOOF_RED
    p1.space_after = Pt(8)

    points1 = [
        ("Core Assumption", "If a voice sounds like someone, it can be trusted as that person."),
        ("Verification Method", "Subjective human auditory perception & cognitive familiarity."),
        ("Vulnerabilities", "Easily tricked by pitch-matching vocoders, emotional manipulation, and noisy lines."),
        ("Security Outcome", "Blind spot for social engineering, unauthorized wire transfers, and identity spoofing.")
    ]
    for bold_p, text_p in points1:
        p = tf1.add_paragraph()
        p.space_after = Pt(4)
        r1 = p.add_run()
        r1.text = f"• {bold_p}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = text_p
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # SIH-Aegis Model (Green)
    c2 = create_card(slide, Inches(6.8), col_top, col_w, col_h, C_BONAFIDE_BORDER, C_BONAFIDE_BG)
    tf2 = c2.text_frame
    tf2.word_wrap = True
    tf2.margin_left = tf2.margin_right = Inches(0.25)
    tf2.margin_top = Inches(0.2)
    p2 = tf2.paragraphs[0]
    p2.text = "SIH-AEGIS ACTIVE DEFENSE  (OBJECTIVE SECURITY)"
    p2.font.name = FONT_BODY
    p2.font.size = Pt(12)
    p2.font.bold = True
    p2.font.color.rgb = C_BONAFIDE_GREEN
    p2.space_after = Pt(8)

    points2 = [
        ("Core Principle", "Treat voice as an untrusted biometric payload until mathematically verified."),
        ("Verification Method", "Extract sub-perceptual acoustic artifacts, phase continuity, & spectral traces."),
        ("Capabilities", "Detects synthetic vocoder artifacts even when pitch and cadence perfectly match human speech."),
        ("Security Outcome", "Deterministic, measurable authenticity score (BONAFIDE vs SPOOF) with confidence.")
    ]
    for bold_p, text_p in points2:
        p = tf2.add_paragraph()
        p.space_after = Pt(4)
        r1 = p.add_run()
        r1.text = f"• {bold_p}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = text_p
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # 5 Real-World Deployment Sector Cards
    sec_title_box = slide.shapes.add_textbox(Inches(0.8), Inches(4.75), Inches(11.7), Inches(0.35))
    tf_st = sec_title_box.text_frame
    tf_st.word_wrap = True
    tf_st.margin_left = tf_st.margin_top = tf_st.margin_right = tf_st.margin_bottom = 0
    p_st = tf_st.paragraphs[0]
    p_st.text = "CRITICAL TARGET DOMAINS REQUIRING OBJECTIVE AUTHENTICITY"
    p_st.font.name = FONT_BODY
    p_st.font.size = Pt(10.5)
    p_st.font.bold = True
    p_st.font.color.rgb = C_CYAN

    sectors = [
        ("Banking & Wire Transfers", "Verbal authorization for high-value transactions"),
        ("Remote KYC & Onboarding", "Preventing synthetic injection into digital identity queues"),
        ("Enterprise Call Centers", "Inbound verification protecting customer service agents"),
        ("Emergency & Family Defense", "Real-time threat alerts on coercive extortion calls"),
        ("Voice Biometric MFA", "Anti-spoofing gateway shielding speaker recognition")
    ]
    sec_w = Inches(2.23)
    sec_h = Inches(1.5)
    sec_top = Inches(5.15)

    for i, (title_s, desc_s) in enumerate(sectors):
        left_s = Inches(0.8) + i * Inches(2.37)
        cs = create_card(slide, left_s, sec_top, sec_w, sec_h, C_CARD_BORDER, C_CARD_BG)
        tfs = cs.text_frame
        tfs.word_wrap = True
        tfs.vertical_anchor = MSO_ANCHOR.MIDDLE
        tfs.margin_left = tfs.margin_right = Inches(0.14)
        
        p_s1 = tfs.paragraphs[0]
        p_s1.text = title_s
        p_s1.font.name = FONT_BODY
        p_s1.font.size = Pt(10.5)
        p_s1.font.bold = True
        p_s1.font.color.rgb = C_TEXT_WHITE
        p_s1.space_after = Pt(4)
        
        p_s2 = tfs.add_paragraph()
        p_s2.text = desc_s
        p_s2.font.name = FONT_BODY
        p_s2.font.size = Pt(9)
        p_s2.font.color.rgb = C_TEXT_MUTED

    meaning = (
        "Defines the conceptual problem: replacing subjective human trust with objective, measurable AI verification. "
        "Shows judges the broad commercial and societal utility across banking, telecom, KYC, and personal security."
    )
    script = (
        "The core question SIH-Aegis solves is: 'Can a voice that sounds like someone actually be trusted as that person?' "
        "Historically, security relied on human judgement: 'It sounds like my boss or my mother, so I trust it.' "
        "SIH-Aegis replaces this broken assumption with active defense. "
        "Instead of asking whether it sounds real, our models measure physical and spectral authenticity: phase continuity, vocoder artifacts, and time-frequency consistency. "
        "This turns voice authenticity into a verifiable security signal for banking, KYC, call centers, and consumer defense."
    )
    qa = [
        ("Can this scale to banking and call centres?", 
         "Yes. In banking and call centers, SIH-Aegis functions as an automated pre-screening microservice. Inbound audio from telephony SIP trunks or WebRTC is streamed into our inference pipeline, returning a sub-second risk score before an agent authorizes a fund transfer or password reset.")
    ]
    set_speaker_notes(slide, meaning, script, qa)


def build_slide_4(prs, layout):
    slide = prs.slides.add_slide(layout)
    apply_dark_background(slide)
    add_header(slide, 'Current Working Pipeline: Verified Baseline & ONNX Parity', 'Technical Execution', 4)

    # Status Banner
    stat_banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.6), Inches(11.733), Inches(0.42))
    stat_banner.fill.solid()
    stat_banner.fill.fore_color.rgb = C_BONAFIDE_BG
    stat_banner.line.color.rgb = C_BONAFIDE_BORDER
    tf_sb = stat_banner.text_frame
    tf_sb.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_sb = tf_sb.paragraphs[0]
    p_sb.text = "✓ CURRENT VERIFIED PROTOTYPE: END-TO-END DETECTION OPERATIONAL ON PYTORCH & ONNX RUNTIME"
    p_sb.alignment = PP_ALIGN.CENTER
    p_sb.font.name = FONT_BODY
    p_sb.font.size = Pt(10)
    p_sb.font.bold = True
    p_sb.font.color.rgb = C_BONAFIDE_GREEN

    # 6 Pipeline Flow Steps
    steps = [
        ("1. Audio Input", ".wav / .mp3 / .flac\nSingle / streaming source"),
        ("2. Preprocessing", "16 kHz resampling, mono,\nlength normalization"),
        ("3. AASIST Baseline", "Pretrained anti-spoofing\nsinc-convolution backbone"),
        ("4. Logit Extraction", "Deterministic forward pass\nSoftmax probability"),
        ("5. Security Verdict", "SPOOF / BONAFIDE\nConfidence percentage"),
        ("6. Streamlit UI", "Real-time diagnostic\noperator dashboard")
    ]
    step_w = Inches(1.78)
    step_h = Inches(1.5)
    step_top = Inches(2.2)

    for i, (title, desc) in enumerate(steps):
        left_pos = Inches(0.8) + i * Inches(2.0)
        c = create_card(slide, left_pos, step_top, step_w, step_h, C_CARD_BORDER, C_CARD_BG)
        tfc = c.text_frame
        tfc.word_wrap = True
        tfc.vertical_anchor = MSO_ANCHOR.MIDDLE
        tfc.margin_left = tfc.margin_right = Inches(0.12)
        
        p1 = tfc.paragraphs[0]
        p1.text = title
        p1.font.name = FONT_BODY
        p1.font.size = Pt(10.5)
        p1.font.bold = True
        p1.font.color.rgb = C_CYAN
        p1.space_after = Pt(4)
        
        p2 = tfc.add_paragraph()
        p2.text = desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(9)
        p2.font.color.rgb = C_TEXT_MUTED

        if i < len(steps) - 1:
            arr_box = slide.shapes.add_textbox(left_pos + step_w, step_top + Inches(0.5), Inches(0.22), Inches(0.4))
            tf_a = arr_box.text_frame
            tf_a.margin_left = tf_a.margin_top = tf_a.margin_right = tf_a.margin_bottom = 0
            p_a = tf_a.paragraphs[0]
            p_a.text = "→"
            p_a.font.name = FONT_TITLE
            p_a.font.size = Pt(16)
            p_a.font.bold = True
            p_a.font.color.rgb = C_CYAN
            p_a.alignment = PP_ALIGN.CENTER

    # 2 Technical Verification Cards Below
    c_w = Inches(5.72)
    c_h = Inches(3.0)
    c_top = Inches(3.95)

    # Hardware & PyTorch Execution Card
    c_left = create_card(slide, Inches(0.8), c_top, c_w, c_h, C_CARD_BORDER, C_CARD_BG)
    tf_l = c_left.text_frame
    tf_l.word_wrap = True
    tf_l.margin_left = tf_l.margin_right = Inches(0.25)
    tf_l.margin_top = Inches(0.2)
    p_l1 = tf_l.paragraphs[0]
    p_l1.text = "PROTOTYPE ENVIRONMENT & GPU ACCELERATION"
    p_l1.font.name = FONT_BODY
    p_l1.font.size = Pt(11.5)
    p_l1.font.bold = True
    p_l1.font.color.rgb = C_CYAN
    p_l1.space_after = Pt(8)

    env_facts = [
        ("Framework", "PyTorch 2.x with native CUDA 12 support."),
        ("Host Hardware", "NVIDIA GeForce RTX 4050 Laptop GPU (6 GB VRAM)."),
        ("Model Checkpoint", "Pretrained AASIST (Audio Anti-Spoofing using Integrated Spectro-Temporal Graph Attention)."),
        ("Inference Mode", "Model.eval() deterministic evaluation mode verified."),
        ("Analyst Interface", "Interactive Streamlit web app supporting direct file upload, waveform display, and confidence readouts.")
    ]
    for b_txt, d_txt in env_facts:
        p = tf_l.add_paragraph()
        p.space_after = Pt(4)
        r1 = p.add_run()
        r1.text = f"• {b_txt}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = d_txt
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # ONNX Export & Equivalence Card
    c_right = create_card(slide, Inches(6.8), c_top, c_w, c_h, C_CARD_BORDER, C_CARD_BG)
    tf_r = c_right.text_frame
    tf_r.word_wrap = True
    tf_r.margin_left = tf_r.margin_right = Inches(0.25)
    tf_r.margin_top = Inches(0.2)
    p_r1 = tf_r.paragraphs[0]
    p_r1.text = "ONNX EXPORT & MATHEMATICAL PARITY VALIDATION"
    p_r1.font.name = FONT_BODY
    p_r1.font.size = Pt(11.5)
    p_r1.font.bold = True
    p_r1.font.color.rgb = C_BONAFIDE_GREEN
    p_r1.space_after = Pt(8)

    onnx_facts = [
        ("ONNX Conversion", "Successfully exported full AASIST graph to standard ONNX format for platform-agnostic runtime execution."),
        ("Runtime Engine", "Validated model execution using ONNX Runtime (CPU & CUDA execution providers)."),
        ("Numerical Parity", "Conducted rigorous output comparison against native PyTorch tensors."),
        ("Max Delta Verified", "Max absolute numerical difference: ~3.8e-06 (well within single-precision floating-point epsilon)."),
        ("Class Equivalence", "100% agreement on predicted class (BONAFIDE vs SPOOF) between PyTorch and ONNX Runtime engines.")
    ]
    for b_txt, d_txt in onnx_facts:
        p = tf_r.add_paragraph()
        p.space_after = Pt(4)
        r1 = p.add_run()
        r1.text = f"• {b_txt}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = d_txt
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    meaning = (
        "Provides undeniable technical proof of working code. Shows that the team has a fully functioning, GPU-accelerated "
        "AASIST prototype running in Streamlit, has exported it to ONNX, and validated mathematical parity to 3.8e-06."
    )
    script = (
        "Slide 4 demonstrates what is verified and working today in our codebase. "
        "We built a modular audio pipeline in Python: audio loading, 16 kHz mono resampling, feature extraction, and model evaluation. "
        "We successfully initialized and verified the AASIST anti-spoofing model on an NVIDIA RTX 4050 GPU in PyTorch. "
        "Crucially for deployment, we exported AASIST to ONNX and validated it in ONNX Runtime. "
        "Our parity tests proved a maximum numerical difference of just 3.8e-06 between PyTorch and ONNX, producing identical classification verdicts. "
        "This is exposed through an interactive Streamlit UI."
    )
    qa = [
        ("What is actually working today?", 
         "The end-to-end Python pipeline from audio input to AASIST inference, deterministic scoring, ONNX model export, ONNX Runtime validation with 3.8e-06 numerical parity, and our Streamlit diagnostic interface are all 100% verified and operational."),
        ("Why not just use AASIST as your final product?", 
         "AASIST is our working baseline, but as our field testing demonstrated, it suffers from severe domain mismatch outside clean academic datasets, which directly motivated our custom lightweight architecture.")
    ]
    set_speaker_notes(slide, meaning, script, qa)


def build_slide_5(prs, layout):
    slide = prs.slides.add_slide(layout)
    apply_dark_background(slide)
    add_header(slide, 'Why We Are Building Our Own Detector: The Baseline Insight', 'Key Innovation Insight', 5)

    # Central Engineering Insight Callout Banner
    ins_banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.6), Inches(11.733), Inches(0.72))
    ins_banner.fill.solid()
    ins_banner.fill.fore_color.rgb = C_CARD_BG
    ins_banner.line.color.rgb = C_CYAN
    ins_banner.line.width = Pt(1.5)
    tf_ib = ins_banner.text_frame
    tf_ib.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_ib1 = tf_ib.paragraphs[0]
    p_ib1.text = "CORE ENGINEERING INSIGHT: \"BENCHMARK PERFORMANCE ≠ DEPLOYMENT ROBUSTNESS\""
    p_ib1.alignment = PP_ALIGN.CENTER
    p_ib1.font.name = FONT_BODY
    p_ib1.font.size = Pt(13)
    p_ib1.font.bold = True
    p_ib1.font.color.rgb = C_CYAN
    p_ib2 = tf_ib.add_paragraph()
    p_ib2.text = "Validating AASIST under field conditions revealed critical domain-shift failure modes that motivated our custom detector."
    p_ib2.alignment = PP_ALIGN.CENTER
    p_ib2.font.name = FONT_BODY
    p_ib2.font.size = Pt(10)
    p_ib2.font.color.rgb = C_TEXT_MUTED

    # Two Branch Cards (Controlled Benchmark vs Real-World Audio)
    b_w = Inches(5.72)
    b_h = Inches(2.7)
    b_top = Inches(2.5)

    # Branch 1: Synthetic Benchmark (Success)
    cb1 = create_card(slide, Inches(0.8), b_top, b_w, b_h, C_BONAFIDE_BORDER, C_BONAFIDE_BG)
    tf_b1 = cb1.text_frame
    tf_b1.word_wrap = True
    tf_b1.margin_left = tf_b1.margin_right = Inches(0.25)
    tf_b1.margin_top = Inches(0.2)
    p_b1 = tf_b1.paragraphs[0]
    p_b1.text = "BRANCH 1: CONTROLLED ACADEMIC BENCHMARKS"
    p_b1.font.name = FONT_BODY
    p_b1.font.size = Pt(11.5)
    p_b1.font.bold = True
    p_b1.font.color.rgb = C_BONAFIDE_GREEN
    p_b1.space_after = Pt(6)

    b1_pts = [
        ("Evaluation Target", "ASVspoof 2019 Logical Access (LA) evaluation partitions."),
        ("Behavior Observed", "AASIST correctly and decisively classifies known synthetic samples as SPOOF."),
        ("Strengths", "Excels at detecting specific neural vocoders and synthesis artifacts represented in the training distribution."),
        ("False Sense of Security", "High paper accuracy suggests the problem is solved, masking deployment realities.")
    ]
    for b_t, d_t in b1_pts:
        p = tf_b1.add_paragraph()
        p.space_after = Pt(3)
        r1 = p.add_run()
        r1.text = f"• {b_t}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = d_t
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # Branch 2: Real-World Audio (Failure Mode / False Positives)
    cb2 = create_card(slide, Inches(6.8), b_top, b_w, b_h, C_SPOOF_BORDER, C_SPOOF_BG)
    tf_b2 = cb2.text_frame
    tf_b2.word_wrap = True
    tf_b2.margin_left = tf_b2.margin_right = Inches(0.25)
    tf_b2.margin_top = Inches(0.2)
    p_b2 = tf_b2.paragraphs[0]
    p_b2.text = "BRANCH 2: REAL-WORLD MICROPHONE & CALL RECORDINGS"
    p_b2.font.name = FONT_BODY
    p_b2.font.size = Pt(11.5)
    p_b2.font.bold = True
    p_b2.font.color.rgb = C_SPOOF_RED
    p_b2.space_after = Pt(6)

    b2_pts = [
        ("Evaluation Target", "Genuine smartphone voice notes, laptop microphones, and cellular recordings."),
        ("Behavior Observed", "AASIST frequently produced strong false SPOOF predictions on legitimate human speech!"),
        ("Root Cause Analysis", "Severe domain mismatch: room acoustics, background fan noise, and phone codecs are misclassified as synthetic artifacts."),
        ("Engineering Conclusion", "A baseline model cannot be deployed in production without domain-aware adaptation.")
    ]
    for b_t, d_t in b2_pts:
        p = tf_b2.add_paragraph()
        p.space_after = Pt(3)
        r1 = p.add_run()
        r1.text = f"• {b_t}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = d_t
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # Takeaway Action Card at Bottom
    c_take = create_card(slide, Inches(0.8), Inches(5.35), Inches(11.733), Inches(1.5), C_CARD_BORDER, C_CARD_BG)
    tf_tk = c_take.text_frame
    tf_tk.word_wrap = True
    tf_tk.margin_left = tf_tk.margin_right = Inches(0.25)
    tf_tk.margin_top = Inches(0.16)
    p_tk1 = tf_tk.paragraphs[0]
    p_tk1.text = "OUR STRATEGIC RESPONSE: DOMAIN-AWARE CUSTOM DETECTOR"
    p_tk1.font.name = FONT_BODY
    p_tk1.font.size = Pt(11)
    p_tk1.font.bold = True
    p_tk1.font.color.rgb = C_CYAN
    p_tk1.space_after = Pt(4)

    takeaway_pts = [
        ("We do NOT claim AASIST is useless", "AASIST is a strong theoretical baseline, but deploying it blindly would paralyze a bank with false alarm rejections."),
        ("Why a custom model is required", "We need a detector trained with acoustic augmentations (environmental noise, telecom compression, diverse Indian accents) to separate true synthesis artifacts from real-world acoustic variations."),
        ("Primary Development Direction", "We designed a lightweight Spectrogram CNN specifically structured to learn invariant time-frequency features while running efficiently on consumer and edge devices.")
    ]
    for b_t, d_t in takeaway_pts:
        p = tf_tk.add_paragraph()
        p.space_after = Pt(2)
        r1 = p.add_run()
        r1.text = f"→ {b_t}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = d_t
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    meaning = (
        "Demonstrates deep engineering maturity. Instead of pretending a pretrained model works magically everywhere, "
        "the team highlights empirical testing showing false alarms on genuine speech, directly justifying the custom CNN."
    )
    script = (
        "Slide 5 contains one of our most critical engineering discoveries. "
        "When we evaluated the AASIST baseline, it performed as expected on synthetic benchmarks: it correctly flagged spoofed voices. "
        "However, when we tested genuine, real-world voice notes recorded on smartphones and laptops, AASIST produced false SPOOF predictions on real people! "
        "This proved a fundamental industry lesson: benchmark performance does not equal deployment robustness. "
        "Academic models overfit to studio acoustics. Deploying AASIST in India would reject legitimate banking customers. "
        "This exact insight is why we are developing our own domain-aware Spectrogram CNN."
    )
    qa = [
        ("What makes your approach different from existing research?", 
         "Most academic projects stop at evaluating pretrained weights on ASVspoof clean partitions. We actively identified its real-world generalization failure on everyday microphones and designed our custom architecture and data augmentation pipeline specifically to mitigate channel and acoustic false positives."),
        ("How will you evaluate false positives?", 
         "We measure Equal Error Rate (EER) and false-rejection rate across cross-domain evaluation splits, deliberately including low-cost smartphone recordings, reverberant rooms, and cellular codecs.")
    ]
    set_speaker_notes(slide, meaning, script, qa)


def build_slide_6(prs, layout):
    slide = prs.slides.add_slide(layout)
    apply_dark_background(slide)
    add_header(slide, 'Level 3 Direction: Learning Time-Frequency Patterns with CNN', 'Custom Model Architecture', 6)

    # Left Column: Representation & Spectrogram (Width: 5.72)
    c_spec = create_card(slide, Inches(0.8), Inches(1.6), Inches(5.72), Inches(3.9), C_CARD_BORDER, C_CARD_BG)
    tf_s = c_spec.text_frame
    tf_s.word_wrap = True
    tf_s.margin_left = tf_s.margin_right = Inches(0.25)
    tf_s.margin_top = Inches(0.2)
    p_s1 = tf_s.paragraphs[0]
    p_s1.text = "ACOUSTIC REPRESENTATION: 128 × 128 LOG-MEL SPECTROGRAM"
    p_s1.font.name = FONT_BODY
    p_s1.font.size = Pt(11.5)
    p_s1.font.bold = True
    p_s1.font.color.rgb = C_CYAN
    p_s1.space_after = Pt(8)

    spec_details = [
        ("Why Spectrogram?", "Speech is non-stationary. Time-frequency representations capture harmonic structures and phase transitions where neural vocoders leave subtle mathematical flaws."),
        ("Input Sampling", "16,000 Hz single-channel mono audio."),
        ("STFT Configuration", "FFT window size: 512 | Hop length: 160 samples (10 ms frame advance)."),
        ("Filterbank Mapping", "128 Mel frequency bins scaled psychoacoustically."),
        ("Final Matrix Shape", "Fixed 128 × 128 float32 representation per 1.28-second temporal context."),
        ("Normalization", "Per-sample zero-mean, unit-variance standardization to eliminate recording volume bias.")
    ]
    for b_t, d_t in spec_details:
        p = tf_s.add_paragraph()
        p.space_after = Pt(4)
        r1 = p.add_run()
        r1.text = f"• {b_t}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = d_t
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # Right Column: CNN Architecture (Width: 5.72)
    c_cnn = create_card(slide, Inches(6.8), Inches(1.6), Inches(5.72), Inches(3.9), C_CARD_BORDER, C_CARD_BG)
    tf_c = c_cnn.text_frame
    tf_c.word_wrap = True
    tf_c.margin_left = tf_c.margin_right = Inches(0.25)
    tf_c.margin_top = Inches(0.2)
    p_c1 = tf_c.paragraphs[0]
    p_c1.text = "LIGHTWEIGHT CONVOLUTIONAL DETECTOR (~97.9K PARAMS)"
    p_c1.font.name = FONT_BODY
    p_c1.font.size = Pt(11.5)
    p_c1.font.bold = True
    p_c1.font.color.rgb = C_BONAFIDE_GREEN
    p_c1.space_after = Pt(8)

    arch_blocks = [
        ("Input Tensor", "Shape: 1 × 128 × 128 (Batch, Channel, Freq, Time)"),
        ("Conv Block 1", "Conv2d (1 → 16 channels) + BatchNorm + ReLU + MaxPool(2,2)"),
        ("Conv Block 2", "Conv2d (16 → 32 channels) + BatchNorm + ReLU + MaxPool(2,2)"),
        ("Conv Block 3", "Conv2d (32 → 64 channels) + BatchNorm + ReLU + MaxPool(2,2)"),
        ("Conv Block 4", "Conv2d (64 → 128 channels) + BatchNorm + ReLU (Dense spectral cues)"),
        ("Classification Head", "AdaptiveAvgPool2d(1,1) → Flatten → Dropout(0.3) → Linear(128 → 2)"),
        ("Model Footprint", "Exactly 97,890 trainable parameters (~390 KB memory footprint)")
    ]
    for b_t, d_t in arch_blocks:
        p = tf_c.add_paragraph()
        p.space_after = Pt(4)
        r1 = p.add_run()
        r1.text = f"• {b_t}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = d_t
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # Bottom Honesty & Technical Status Box
    c_stat = create_card(slide, Inches(0.8), Inches(5.65), Inches(11.733), Inches(1.3), C_AMBER_BORDER, C_AMBER_BG)
    tf_st = c_stat.text_frame
    tf_st.word_wrap = True
    tf_st.margin_left = tf_st.margin_right = Inches(0.25)
    tf_st.margin_top = Inches(0.14)
    p_st1 = tf_st.paragraphs[0]
    p_st1.text = "RIGOROUS TECHNICAL HONESTY STATEMENT"
    p_st1.font.name = FONT_BODY
    p_st1.font.size = Pt(10.5)
    p_st1.font.bold = True
    p_st1.font.color.rgb = C_AMBER
    p_st1.space_after = Pt(3)

    honesty_items = [
        ("Architecture Implemented & Verified", "The 4-block CNN graph has been coded in PyTorch; full tensor forward-pass compatibility has been verified on our RTX 4050 GPU."),
        ("Training Status (In Progress)", "The CNN has NOT yet been trained. There is NO trained checkpoint or fabricated accuracy score. Training dataset preparation with Indian accents and multi-TTS engines is actively underway.")
    ]
    for b_t, d_t in honesty_items:
        p = tf_st.add_paragraph()
        p.space_after = Pt(2)
        r1 = p.add_run()
        r1.text = f"✓ {b_t}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = d_t
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    meaning = (
        "Presents the team's primary new model direction: a lightweight, 97.9k parameter Spectrogram CNN. "
        "Highlights why CNN on spectrogram works, and explicitly confirms that the forward pass is verified while training is in progress."
    )
    script = (
        "Slide 6 presents our primary model direction: a lightweight Spectrogram CNN. "
        "Speech is non-stationary; neural vocoders leave characteristic frequency smearing and temporal phase discontinuities. "
        "We convert 16 kHz audio into 128 by 128 Log-Mel spectrograms via STFT. "
        "Our CNN uses 4 progressive convolution blocks, Batch Normalization, and Adaptive Pooling, totaling just 97,890 parameters—less than 400 kilobytes. "
        "We are technically transparent: the architecture is coded and GPU forward-pass verified on our RTX 4050, but it is not yet trained. "
        "We will not claim a fabricated accuracy score."
    )
    qa = [
        ("Why use a CNN rather than a transformer or RNN?", 
         "A 2D CNN treats the spectrogram as a time-frequency image, efficiently capturing localized acoustic anomalies (spectral cutoffs, harmonic smearing) with low parameter count. Transformers require heavy self-attention compute, which is prohibitive for real-time edge/mobile deployment."),
        ("Why spectrogram representation?", 
         "Raw audio contains heavy phase noise. Log-Mel spectrograms preserve perceptual frequency bands while exposing spectral energy anomalies left by neural vocoders like HiFi-GAN and WaveGlow."),
        ("Is the CNN trained already?", 
         "No. The architecture is fully implemented, verified, and ready for training. We are currently compiling our balanced Indian-speaker and multi-generator training dataset.")
    ]
    set_speaker_notes(slide, meaning, script, qa)


def build_slide_7(prs, layout):
    slide = prs.slides.add_slide(layout)
    apply_dark_background(slide)
    add_header(slide, 'System Architecture: Modular, Extensible Security Pipeline', 'Engineering Architecture', 7)

    # 6 Layer Horizontal Architecture Stack
    layers = [
        ("1. CLIENT & EDGE INTERACTION LAYER", 
         "Streamlit Diagnostic Web Dashboard (Working Prototype)  ·  Android Jetpack Compose Mobile App (Under Integration)",
         C_CYAN, C_CARD_BORDER, C_CARD_BG),
        ("2. PIPELINE INGESTION & ORCHESTRATION LAYER", 
         "Audio file ingestion (.wav, .mp3, .m4a)  ·  Header integrity verification  ·  3-second sliding chunk buffer  ·  Stream synchronizer",
         C_TEXT_WHITE, C_CARD_BORDER, C_CARD_BG),
        ("3. DIGITAL AUDIO PREPROCESSING LAYER", 
         "Dynamic 16 kHz resampling (librosa/soundfile)  ·  Mono downmix  ·  Silence/Energy threshold trimming  ·  RMS normalization",
         C_TEXT_WHITE, C_CARD_BORDER, C_CARD_BG),
        ("4. FEATURE EXTRACTION & TENSOR TRANSFORM", 
         "Path A: Raw Sinc-convolution waveform tensor slicing  ·  Path B: 128-Mel STFT Log-Mel Spectrogram generator (128 × 128)",
         C_CYAN, C_CARD_BORDER, C_CARD_ACTIVE),
        ("5. DUAL DETECTION ENGINE (MODULAR BACKBONE)", 
         "AASIST Benchmark Baseline (Pretrained, ONNX Runtime, 3.8e-06 parity)  ⇄  Lightweight Spectrogram CNN (97.9K params, primary dev)",
         C_BONAFIDE_GREEN, C_BONAFIDE_BORDER, C_BONAFIDE_BG),
        ("6. DECISION & EXPLAINABILITY LAYER", 
         "Softmax confidence scores  ·  Configurable thresholding  ·  Binary Classification (BONAFIDE / SPOOF)  ·  Audit log generator",
         C_AMBER, C_CARD_BORDER, C_CARD_BG)
    ]

    layer_w = Inches(11.733)
    layer_h = Inches(0.72)
    start_top = Inches(1.6)
    gap = Inches(0.82)

    for i, (title_l, desc_l, col_title, col_border, col_bg) in enumerate(layers):
        c = create_card(slide, Inches(0.8), start_top + i * gap, layer_w, layer_h, col_border, col_bg)
        tfc = c.text_frame
        tfc.word_wrap = True
        tfc.vertical_anchor = MSO_ANCHOR.MIDDLE
        tfc.margin_left = tfc.margin_right = Inches(0.2)
        
        p1 = tfc.paragraphs[0]
        p1.text = title_l
        p1.font.name = FONT_BODY
        p1.font.size = Pt(10.5)
        p1.font.bold = True
        p1.font.color.rgb = col_title
        p1.space_after = Pt(2)
        
        p2 = tfc.add_paragraph()
        p2.text = desc_l
        p2.font.name = FONT_BODY
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_MUTED

    # Bottom Callout
    c_foot = create_card(slide, Inches(0.8), Inches(6.55), Inches(11.733), Inches(0.52), C_CARD_BORDER, C_CARD_BG)
    tf_f = c_foot.text_frame
    tf_f.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_f.margin_left = Inches(0.2)
    p_f = tf_f.paragraphs[0]
    p_f.text = "MODULAR DESIGN PRINCIPLE: Features, models, and interfaces are decoupled. AASIST serves as our benchmark validator, while our 97.9K CNN is optimized for real-time mobile and edge microservices."
    p_f.font.name = FONT_BODY
    p_f.font.size = Pt(9.5)
    p_f.font.bold = True
    p_f.font.color.rgb = C_TEXT_WHITE

    meaning = (
        "Shows the full end-to-end layered architecture. Proves that SIH-Aegis is not just a standalone script, "
        "but a production-grade, modular software architecture with decoupled ingestion, preprocessing, modeling, and output layers."
    )
    script = (
        "Slide 7 illustrates the modular architecture of SIH-Aegis. "
        "The system is organized into clean functional layers: presentation via Streamlit and Android Compose, ingestion, preprocessing, dual feature extraction, and our detection engines. "
        "This decoupling is vital: our AASIST baseline and our custom Spectrogram CNN share the same audio preprocessing and validation pipelines. "
        "When our CNN completes training, it plugs directly into the existing ONNX Runtime engine without rewriting a single line of application code."
    )
    qa = [
        ("Why is the model designed to be so lightweight (~97.9k params)?", 
         "Because production telephony and smartphone detection requires low latency (<100 ms) and low battery/thermal impact. Heavy architectures cannot run alongside a real-time call. Our 97.9k CNN fits entirely in mobile L3 cache (~390 KB) and executes in milliseconds.")
    ]
    set_speaker_notes(slide, meaning, script, qa)


def build_slide_8(prs, layout):
    slide = prs.slides.add_slide(layout)
    apply_dark_background(slide)
    add_header(slide, 'Android Mobile Security: Current Milestone & Future Real-Time', 'Mobile Implementation', 8)

    # Left Panel: Current Verified Android Milestone
    c_curr = create_card(slide, Inches(0.8), Inches(1.6), Inches(5.72), Inches(3.95), C_BONAFIDE_BORDER, C_BONAFIDE_BG)
    tf_cur = c_curr.text_frame
    tf_cur.word_wrap = True
    tf_cur.margin_left = tf_cur.margin_right = Inches(0.25)
    tf_cur.margin_top = Inches(0.2)
    p_cur = tf_cur.paragraphs[0]
    p_cur.text = "CURRENT VERIFIED ANDROID MILESTONE (WORKING)"
    p_cur.font.name = FONT_BODY
    p_cur.font.size = Pt(11.5)
    p_cur.font.bold = True
    p_cur.font.color.rgb = C_BONAFIDE_GREEN
    p_cur.space_after = Pt(8)

    android_milestones = [
        ("Native Android Stack", "Modern Android Studio project written purely in Kotlin with Jetpack Compose & Material 3 design."),
        ("ONNX Runtime Mobile", "Integrated onnxruntime-android dependency into Gradle build."),
        ("Model Asset Loading", "Bundled AASIST ONNX model into app assets; verified successful initialization and in-memory session loading."),
        ("UI State Machine", "Built dynamic reactive Compose UI supporting 5 distinct states: READY → LISTENING → ANALYZING → BONAFIDE → SPOOF."),
        ("Current Integration Target", "Audio file upload-based inference pipeline currently under active integration for demonstration.")
    ]
    for b_t, d_t in android_milestones:
        p = tf_cur.add_paragraph()
        p.space_after = Pt(4)
        r1 = p.add_run()
        r1.text = f"✓ {b_t}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = d_t
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # Right Panel: Next Stage & Future Streaming Pipeline
    c_fut = create_card(slide, Inches(6.8), Inches(1.6), Inches(5.72), Inches(3.95), C_CARD_BORDER, C_CARD_BG)
    tf_fut = c_fut.text_frame
    tf_fut.word_wrap = True
    tf_fut.margin_left = tf_fut.margin_right = Inches(0.25)
    tf_fut.margin_top = Inches(0.2)
    p_fut = tf_fut.paragraphs[0]
    p_fut.text = "FUTURE MOBILE STREAMING PIPELINE (ROADMAP)"
    p_fut.font.name = FONT_BODY
    p_fut.font.size = Pt(11.5)
    p_fut.font.bold = True
    p_fut.font.color.rgb = C_CYAN
    p_fut.space_after = Pt(8)

    next_steps = [
        ("Microphone Audio Capture", "Streaming PCM audio buffer capture via Android AudioRecord API."),
        ("Sliding Window Chunking", "Slicing live speech into overlapping 2 to 3-second acoustic evaluation windows."),
        ("On-Device Preprocessing", "Native lightweight FFT Log-Mel extraction on Android CPU/NPU."),
        ("Low-Latency Inference", "Executing our 97.9k parameter CNN with sub-50ms latency using ONNX Runtime NNAPI accelerator."),
        ("Real-Time Visual Overlay", "Live visual security indicator alerting user during active audio calls.")
    ]
    for b_t, d_t in next_steps:
        p = tf_fut.add_paragraph()
        p.space_after = Pt(4)
        r1 = p.add_run()
        r1.text = f"→ {b_t}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = d_t
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # Bottom Boundary & Honesty Box
    c_bound = create_card(slide, Inches(0.8), Inches(5.7), Inches(11.733), Inches(1.3), C_AMBER_BORDER, C_AMBER_BG)
    tf_bd = c_bound.text_frame
    tf_bd.word_wrap = True
    tf_bd.margin_left = tf_bd.margin_right = Inches(0.25)
    tf_bd.margin_top = Inches(0.14)
    p_bd1 = tf_bd.paragraphs[0]
    p_bd1.text = "EXPLICIT OPERATING SYSTEM & SECURITY BOUNDARY NOTICE"
    p_bd1.font.name = FONT_BODY
    p_bd1.font.size = Pt(10.5)
    p_bd1.font.bold = True
    p_bd1.font.color.rgb = C_AMBER
    p_bd1.space_after = Pt(3)

    boundary_pts = [
        ("No False Claims on Cellular Interception", "Standard third-party Android applications CANNOT arbitrarily record cellular phone calls due to Android Scoped Storage and accessibility privacy restrictions introduced in Android 10+."),
        ("Our Realistic Target Architecture", "SIH-Aegis targets VoIP applications, messaging voice notes, consented recording workflows, enterprise customer support integrations, and Telecom SDK partnerships.")
    ]
    for b_t, d_t in boundary_pts:
        p = tf_bd.add_paragraph()
        p.space_after = Pt(2)
        r1 = p.add_run()
        r1.text = f"⚠ {b_t}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = d_t
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    meaning = (
        "Details mobile execution. Demonstrates that the Android application is real, built in Kotlin and Compose with ONNX Runtime, "
        "while honestly acknowledging Android OS limitations regarding cellular call recording."
    )
    script = (
        "Slide 8 addresses our mobile execution. "
        "We have set up an Android Studio project using Kotlin, Jetpack Compose, and Material 3. "
        "We successfully embedded ONNX Runtime Mobile and loaded our AASIST ONNX model into memory inside the Android environment. "
        "We have a responsive state machine transitioning between Ready, Listening, Analyzing, Bonafide, and Spoof. "
        "Equally important is our engineering integrity: we do not claim that our app intercepts live cellular calls in the background. "
        "Android OS privacy sandboxes prevent arbitrary call recording. "
        "Our roadmap targets VoIP, voice notes, user-consented audio, and enterprise telecom SDK partnerships."
    )
    qa = [
        ("Is Android already doing real-time cellular call detection?", 
         "No, and anyone claiming a normal 3rd-party Android app can freely intercept standard cellular calls is ignoring Android 10+ security policies. We are currently integrating file upload inference, and our next milestone is live microphone streaming for VoIP and messaging apps.")
    ]
    set_speaker_notes(slide, meaning, script, qa)


def build_slide_9(prs, layout):
    slide = prs.slides.add_slide(layout)
    apply_dark_background(slide)
    add_header(slide, 'Engineering & Data Strategy: Designed for the Real World', 'Data Strategy', 9)

    # 4 Challenge Blocks (Top Row)
    chall_title = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.7), Inches(0.3))
    tf_ct = chall_title.text_frame
    tf_ct.word_wrap = True
    tf_ct.margin_left = tf_ct.margin_top = tf_ct.margin_right = tf_ct.margin_bottom = 0
    p_ct = tf_ct.paragraphs[0]
    p_ct.text = "FOUR REAL-WORLD GENERALIZATION CHALLENGES"
    p_ct.font.name = FONT_BODY
    p_ct.font.size = Pt(10.5)
    p_ct.font.bold = True
    p_ct.font.color.rgb = C_SPOOF_RED

    challenges = [
        ("DATASET MISMATCH", "Academic datasets (ASVspoof) feature mostly Western studio voices, lacking Indian accents, regional dialects, and colloquial rhythms.", C_SPOOF_BORDER, C_SPOOF_BG),
        ("ACOUSTIC NOISE", "Real calls occur in traffic, crowded rooms, and near ceiling fans. Acoustic noise creates false spectral anomalies.", C_SPOOF_BORDER, C_SPOOF_BG),
        ("DEVICE & CODEC DRIFT", "Telecom codecs (AMR-NB, GSM, Opus) strip high frequencies, altering the acoustic signature seen by models.", C_SPOOF_BORDER, C_SPOOF_BG),
        ("UNSEEN TTS ENGINES", "Commercial voice synthesizers evolve weekly. A model that memorizes one vocoder fails against new cloning architectures.", C_SPOOF_BORDER, C_SPOOF_BG)
    ]
    ch_w = Inches(2.78)
    ch_h = Inches(1.8)
    ch_top = Inches(1.85)

    for i, (title, desc, border, bg) in enumerate(challenges):
        left_pos = Inches(0.8) + i * Inches(2.98)
        c = create_card(slide, left_pos, ch_top, ch_w, ch_h, border, bg)
        tfc = c.text_frame
        tfc.word_wrap = True
        tfc.vertical_anchor = MSO_ANCHOR.TOP
        tfc.margin_left = tfc.margin_right = Inches(0.18)
        tfc.margin_top = Inches(0.16)
        
        p1 = tfc.paragraphs[0]
        p1.text = title
        p1.font.name = FONT_BODY
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = C_SPOOF_RED
        p1.space_after = Pt(4)
        
        p2 = tfc.add_paragraph()
        p2.text = desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(9)
        p2.font.color.rgb = C_TEXT_MUTED

    # 4 Response Countermeasures (Bottom Row)
    resp_title = slide.shapes.add_textbox(Inches(0.8), Inches(3.8), Inches(11.7), Inches(0.3))
    tf_rt = resp_title.text_frame
    tf_rt.word_wrap = True
    tf_rt.margin_left = tf_rt.margin_top = tf_rt.margin_right = tf_rt.margin_bottom = 0
    p_rt = tf_rt.paragraphs[0]
    p_rt.text = "SIH-AEGIS STRATEGIC COUNTERMEASURES & DATA PIPELINE"
    p_rt.font.name = FONT_BODY
    p_rt.font.size = Pt(10.5)
    p_rt.font.bold = True
    p_rt.font.color.rgb = C_BONAFIDE_GREEN

    responses = [
        ("REPRESENTATIVE INDIAN DATA", "Compiling genuine speech across Indian English, Hindi, and regional languages to eliminate geographical demographic bias.", C_BONAFIDE_BORDER, C_BONAFIDE_BG),
        ("MULTI-GENERATOR SYNTHESIS", "Training with synthetic speech from multiple diverse TTS families: diffusion, autoregressive, and GAN vocoders (Bark, XTTS, ElevenLabs).", C_BONAFIDE_BORDER, C_BONAFIDE_BG),
        ("ACOUSTIC AUGMENTATIONS", "Injecting room impulse responses, background environmental noise, and lossy codec simulations during training to ensure noise invariance.", C_BONAFIDE_BORDER, C_BONAFIDE_BG),
        ("ANDROID CONSENT COLLECTOR", "Planned opt-in Android collection tool: collects 5-10s labeled samples with strict privacy-by-design, isolated metadata, and user consent.", C_CYAN, C_CARD_BG)
    ]
    re_top = Inches(4.15)

    for i, (title, desc, border, bg) in enumerate(responses):
        left_pos = Inches(0.8) + i * Inches(2.98)
        c = create_card(slide, left_pos, re_top, ch_w, ch_h + Inches(0.5), border, bg)
        tfc = c.text_frame
        tfc.word_wrap = True
        tfc.vertical_anchor = MSO_ANCHOR.TOP
        tfc.margin_left = tfc.margin_right = Inches(0.18)
        tfc.margin_top = Inches(0.16)
        
        p1 = tfc.paragraphs[0]
        p1.text = title
        p1.font.name = FONT_BODY
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = C_CYAN if "ANDROID" in title else C_BONAFIDE_GREEN
        p1.space_after = Pt(4)
        
        p2 = tfc.add_paragraph()
        p2.text = desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(9)
        p2.font.color.rgb = C_TEXT_MUTED

    # Bottom Privacy Banner
    c_priv = create_card(slide, Inches(0.8), Inches(6.55), Inches(11.733), Inches(0.55), C_CARD_BORDER, C_CARD_BG)
    tf_pr = c_priv.text_frame
    tf_pr.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_pr.margin_left = Inches(0.2)
    p_pr = tf_pr.paragraphs[0]
    p_pr.text = "PRIVACY-BY-DESIGN: Voice is biometric data. Collected samples require explicit user opt-in, store no personally identifiable information (PII), use isolated speaker UUIDs, and feature a strict deletion policy."
    p_pr.font.name = FONT_BODY
    p_pr.font.size = Pt(9.5)
    p_pr.font.bold = True
    p_pr.font.color.rgb = C_TEXT_WHITE

    meaning = (
        "Demonstrates a comprehensive real-world data strategy. Shows how the team tackles domain shift, Indian language accents, "
        "and unseen cloning models, while adhering to biometric privacy-by-design standards."
    )
    script = (
        "Slide 9 outlines our data and robustness strategy. "
        "A detector trained only on American or European voices will fail when deployed in India. "
        "Our strategy tackles four barriers: accent mismatch, acoustic noise, codec compression, and unseen synthesis generators. "
        "We are building our dataset to include Indian English and regional languages, multiple synthesis architectures, and aggressive noise augmentation. "
        "We also planned a consent-based Android data collection mode to gather opt-in Indian speech ethically. "
        "Voice is biometric data, so we enforce privacy-by-design: explicit user consent, strict metadata isolation, and zero personal identification linking."
    )
    qa = [
        ("How will this work for Indian speakers?", 
         "By deliberately including Indian English and regional linguistic datasets, and fine-tuning on speech containing typical Indian phoneme cadences and ambient acoustic environments, rather than relying exclusively on Western benchmarks."),
        ("What happens when a new cloning system appears?", 
         "Our CNN is trained on acoustic artifacts common to neural speech reconstruction (vocoder smoothing, phase discontinuities, and high-frequency roll-off), enabling cross-generator generalization rather than memorizing a specific model's watermarks."),
        ("How are you handling biometric privacy?", 
         "We follow privacy-by-design: explicit user opt-in consent, no collection of phone numbers or names, randomized speaker IDs, secure local storage, and immediate user-initiated deletion rights.")
    ]
    set_speaker_notes(slide, meaning, script, qa)


def build_slide_10(prs, layout):
    slide = prs.slides.add_slide(layout)
    apply_dark_background(slide)
    add_header(slide, 'Measurable Impact, Market Viability & Technical Roadmap', 'Impact & Roadmap', 10)

    # Left: 4 Application & Commercial Viability Pillars (Width: 5.72)
    c_app = create_card(slide, Inches(0.8), Inches(1.6), Inches(5.72), Inches(4.3), C_CARD_BORDER, C_CARD_BG)
    tf_a = c_app.text_frame
    tf_a.word_wrap = True
    tf_a.margin_left = tf_a.margin_right = Inches(0.25)
    tf_a.margin_top = Inches(0.2)
    p_a1 = tf_a.paragraphs[0]
    p_a1.text = "COMMERCIAL APPLICATION & SECURITY VALUE"
    p_a1.font.name = FONT_BODY
    p_a1.font.size = Pt(11.5)
    p_a1.font.bold = True
    p_a1.font.color.rgb = C_CYAN
    p_a1.space_after = Pt(8)

    app_pillars = [
        ("Banking & Financial Services", "Real-time verification of high-value verbal transaction authorizations, wire-transfer verifications, and voice-assisted banking bots."),
        ("Remote KYC & Identity Verification", "Anti-injection shield for government and commercial video-KYC queues, ensuring the applicant is speaking live."),
        ("Enterprise Call Centers & BPOs", "Inbound call protection preventing social-engineering attacks against customer service agents and IT helpdesks."),
        ("Citizen Defense & Public Safety", "Empowering telecom users and families with instant threat alerts against urgent extortion, fake kidnapping, and synthetic distress calls.")
    ]
    for b_t, d_t in app_pillars:
        p = tf_a.add_paragraph()
        p.space_after = Pt(5)
        r1 = p.add_run()
        r1.text = f"• {b_t}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_WHITE
        r2 = p.add_run()
        r2.text = d_t
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # Right: Phased Technical Roadmap (Width: 5.72)
    c_road = create_card(slide, Inches(6.8), Inches(1.6), Inches(5.72), Inches(4.3), C_CARD_BORDER, C_CARD_BG)
    tf_r = c_road.text_frame
    tf_r.word_wrap = True
    tf_r.margin_left = tf_r.margin_right = Inches(0.25)
    tf_r.margin_top = Inches(0.2)
    p_r1 = tf_r.paragraphs[0]
    p_r1.text = "PHASED DEVELOPMENT & DEPLOYMENT ROADMAP"
    p_r1.font.name = FONT_BODY
    p_r1.font.size = Pt(11.5)
    p_r1.font.bold = True
    p_r1.font.color.rgb = C_BONAFIDE_GREEN
    p_r1.space_after = Pt(8)

    phases = [
        ("PHASE 1 (COMPLETED)", "Python modular prototype, AASIST ONNX export, 3.8e-06 parity verification, 97.9k CNN GPU forward pass, Streamlit UI, Android base."),
        ("PHASE 2 (IN PROGRESS)", "Indian accent + multi-generator dataset compilation, CNN model training and validation, Android upload-based inference path."),
        ("PHASE 3 (NEXT MILESTONE)", "Android microphone streaming chunker, on-device mobile inference, false-positive reduction under telephony compression."),
        ("PHASE 4 (PRODUCTION SCALE)", "Multi-model ensemble (AASIST + CNN), enterprise SIP trunk integration, SDK packaging, and banking API gateways.")
    ]
    for b_t, d_t in phases:
        p = tf_r.add_paragraph()
        p.space_after = Pt(5)
        r1 = p.add_run()
        r1.text = f"✓ {b_t}: " if "COMPLETED" in b_t else f"→ {b_t}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = C_BONAFIDE_GREEN if "COMPLETED" in b_t else C_AMBER if "IN PROGRESS" in b_t else C_CYAN
        r2 = p.add_run()
        r2.text = d_t
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # Bottom Punchline Banner
    c_close = create_card(slide, Inches(0.8), Inches(6.05), Inches(11.733), Inches(0.95), C_CYAN, C_CARD_BG)
    tf_cl = c_close.text_frame
    tf_cl.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_cl.margin_left = Inches(0.3)
    p_cl1 = tf_cl.paragraphs[0]
    p_cl1.text = "THE SIH-AEGIS PROMISE"
    p_cl1.font.name = FONT_BODY
    p_cl1.font.size = Pt(9.5)
    p_cl1.font.bold = True
    p_cl1.font.color.rgb = C_CYAN
    p_cl1.space_after = Pt(2)
    p_cl2 = tf_cl.add_paragraph()
    p_cl2.text = "\"SIH-Aegis turns voice authenticity from a trust assumption into a measurable security signal.\""
    p_cl2.font.name = FONT_TITLE
    p_cl2.font.size = Pt(16)
    p_cl2.font.bold = True
    p_cl2.font.color.rgb = C_TEXT_WHITE

    meaning = (
        "Concludes the presentation with a powerful strategic vision. Demonstrates high commercial viability across banking, "
        "KYC, and citizen defense, presents an honest phased roadmap, and finishes on the memorable project punchline."
    )
    script = (
        "To conclude, SIH-Aegis is engineered for real-world viability and measurable impact. "
        "The market need spans banking authorizations, KYC onboarding verification, enterprise call center defense, and consumer scam protection. "
        "Our roadmap is clear: Phase 1 is verified with working baseline inference and ONNX parity; Phase 2 trains our lightweight CNN with diverse Indian speech; Phase 3 deploys on-device mobile streaming; and Phase 4 scales to enterprise telecom gateways. "
        "In an era where any voice can be forged in seconds, SIH-Aegis turns voice authenticity from a vulnerable trust assumption into a measurable, objective security signal. "
        "Thank you, and Team SIH-Aegis welcomes your questions."
    )
    qa = [
        ("What is your biggest competitive advantage?", 
         "Technical honesty, edge efficiency, and domain awareness. Rather than claiming impossible features, we discovered baseline real-world limitations and designed a lightweight 97.9k parameter detector tailored for low latency and Indian acoustic conditions."),
        ("How does this align with SIH 2026 evaluation criteria?", 
         "It hits all criteria: F1 Innovation (domain-aware spectrogram CNN), F2 Feasibility (verified ONNX runtime), F3 UX (Streamlit & Compose UI), F4 Impact (fraud defense), F5 Execution (RTX 4050 GPU pipeline), F6 Sustainability (open architecture), F7 Viability (banking/KYC applications), and F8 Privacy (consent-based design).")
    ]
    set_speaker_notes(slide, meaning, script, qa)


# -------------------------------------------------------------
# MAIN BUILD ORCHESTRATOR
# -------------------------------------------------------------
def build_presentation():
    print("Initializing SIH-Aegis Presentation Builder...")
    prs, layout = create_deck()

    print("Building Slide 1: Title Slide...")
    build_slide_1(prs, layout)

    print("Building Slide 2: The Threat...")
    build_slide_2(prs, layout)

    print("Building Slide 3: The Problem We Are Solving...")
    build_slide_3(prs, layout)

    print("Building Slide 4: Current Working Pipeline...")
    build_slide_4(prs, layout)

    print("Building Slide 5: Why We Are Building Our Own Detector...")
    build_slide_5(prs, layout)

    print("Building Slide 6: Level 3 Spectrogram CNN...")
    build_slide_6(prs, layout)

    print("Building Slide 7: System Architecture...")
    build_slide_7(prs, layout)

    print("Building Slide 8: Android Mobile Security...")
    build_slide_8(prs, layout)

    print("Building Slide 9: Engineering & Data Strategy...")
    build_slide_9(prs, layout)

    print("Building Slide 10: Impact & Roadmap...")
    build_slide_10(prs, layout)

    out_file = "/Users/mithumohan/Downloads/SIH/SIH-Aegis-Round1-Presentation.pptx"
    prs.save(out_file)
    print(f"✓ Presentation successfully saved to: {out_file}")

    # Also copy to /Users/mithumohan/Downloads/SIH-Aegis-Round1-Presentation.pptx
    backup_file = "/Users/mithumohan/Downloads/SIH-Aegis-Round1-Presentation.pptx"
    try:
        prs.save(backup_file)
        print(f"✓ Copy successfully saved to: {backup_file}")
    except Exception as e:
        print(f"Notice on backup save: {e}")

if __name__ == "__main__":
    build_presentation()
