import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

import deck_builder as db

def main():
    print("Loading base template...")
    prs = Presentation('template.pptx')
    blank_layout = prs.slide_layouts[6] # blank layout
    
    # ----------------------------------------------------
    # SLIDE 1: OFFICIAL TITLE PAGE
    # ----------------------------------------------------
    slide1 = prs.slides[0]
    # Update title textbox if available, and add project callout
    # Add stylish project box on bottom-left
    box1 = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 331286, 4550000, 5600000, 1650000)
    box1.fill.solid()
    box1.fill.fore_color.rgb = db.C_DARK_NAVY
    box1.line.color.rgb = db.C_CYAN
    box1.line.width = Pt(1.5)
    
    tf1 = box1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = int(Pt(16))
    tf1.margin_top = int(Pt(12))
    tf1.margin_right = int(Pt(16))
    
    p1 = tf1.paragraphs[0]
    p1.text = "SIH-AEGIS"
    p1.font.name = db.FONT_TITLE
    p1.font.size = Pt(22)
    p1.font.bold = True
    p1.font.color.rgb = db.C_CYAN
    p1.space_after = Pt(4)
    
    p2 = tf1.add_paragraph()
    p2.text = "AI-Powered Real-Time Detection and Prevention of\nVoice Cloning Impersonation Attacks"
    p2.font.name = db.FONT_BODY
    p2.font.size = Pt(12)
    p2.font.bold = True
    p2.font.color.rgb = db.C_WHITE
    p2.space_after = Pt(6)
    
    p3 = tf1.add_paragraph()
    p3.text = "ROUND 1 TECHNICAL EVALUATION & WORKING PROTOTYPE"
    p3.font.name = db.FONT_BODY
    p3.font.size = Pt(10)
    p3.font.bold = True
    p3.font.color.rgb = db.C_ACCENT_ORANGE
    
    db.set_notes(slide1, 
        "SPEAKER SCRIPT (30s):\n"
        "Respected judges, we represent Team SIH-Aegis. Today, we present our solution for the detection "
        "and prevention of voice cloning impersonation attacks. Voice cloning tools can now synthesize "
        "any human voice in seconds. Our mission is to turn voice authenticity from a fragile human assumption "
        "into a measurable, real-time security signal. In this Round 1 presentation, we demonstrate our working "
        "prototype, share our baseline engineering discoveries, and explain our custom, lightweight neural detector.\n\n"
        "LIKELY JUDGE QUESTION:\n"
        "Why is voice cloning detection critical right now?\n"
        "STRONG ANSWER:\n"
        "Generative audio models and diffusion vocoders have commoditized realistic speech synthesis. "
        "Traditional human verification and audio filters are defenseless against these attacks, leaving banking, "
        "executive telecom, and ordinary citizens exposed to devastating fraud."
    )
    print("Slide 1 configured.")

    # Remove template slides 2 through 7 (since we will build fresh slides 2 through 16)
    # In python-pptx, we delete slides from the end backwards
    for _ in range(len(prs.slides) - 1):
        rId = prs.slides._sldIdLst[1].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[1]
    
    print(f"Template cleared. Current slides: {len(prs.slides)}")
    
    # ----------------------------------------------------
    # SLIDE 2: PROJECT OVERVIEW (F1, F4)
    # ----------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s2, "IDEA TITLE", "Project Overview — Transforming Auditory Trust into Measurable Security", 2)
    
    card_w = 3600000
    card_h = 4200000
    top_pos = 1200000
    
    db.add_card(s2, 450000, top_pos, card_w, card_h,
        "THE EMERGING THREAT",
        [
            "Democratized Voice Synthesis: Diffusion models and neural vocoders can replicate human vocal timbre, cadence, and breath from mere seconds of audio.",
            "Collapse of Auditory Trust: 'Hearing someone's voice' is no longer proof of their genuine identity or presence.",
            "Severe Attack Impact: Fraudulent wire transfers, CEO spoofing, emergency family ransom scams, and biometric voiceprint theft are surging worldwide."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_RED, title_color=db.C_RED
    )
    
    db.add_card(s2, 4300000, top_pos, card_w, card_h,
        "THE SIH-AEGIS SOLUTION",
        [
            "Acoustic Authenticity Engine: A multi-stage deep learning pipeline that mathematically discriminates human vocal tracts from generative vocoders.",
            "Binary Verdict & Confidence: Generates an instant, reviewable BONAFIDE vs SPOOF classification with a calibrated probability score.",
            "Non-Invasive Verification: Operates passively on ingested audio streams before authorization decisions, money movement, or access grants occur."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_SIH_BLUE, title_color=db.C_SIH_BLUE
    )
    
    db.add_card(s2, 8150000, top_pos, card_w, card_h,
        "OUR ENGINEERING PHILOSOPHY",
        [
            "Baseline-First Validation: We empirically validated state-of-the-art academic models (AASIST) on real-world recordings before architecting our own solution.",
            "Uncovering Real-World Failure: Found significant false alarms on genuine telephone and noisy microphone audio, motivating domain-aware engineering.",
            "Lightweight & Deployable: Prioritizing compact architectures (~97.9k parameters) capable of edge and mobile inference over bloated academic models."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_DARK_NAVY, title_color=db.C_DARK_NAVY
    )
    
    banner = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 5550000, 11300000, 620000)
    banner.fill.solid()
    banner.fill.fore_color.rgb = db.C_DARK_NAVY
    banner.line.color.rgb = db.C_CYAN
    tf_b = banner.text_frame
    tf_b.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_b = tf_b.paragraphs[0]
    p_b.text = "SIH-Aegis turns voice authenticity from a trust assumption into a measurable security signal."
    p_b.alignment = PP_ALIGN.CENTER
    p_b.font.name = db.FONT_BODY
    p_b.font.size = Pt(13)
    p_b.font.bold = True
    p_b.font.color.rgb = db.C_WHITE

    db.set_notes(s2,
        "SPEAKER SCRIPT (30s):\n"
        "Our project addresses a critical cybersecurity frontier. With today's neural audio tools, "
        "anyone with five seconds of audio can clone a target's voice. We cannot rely on human hearing anymore. "
        "SIH-Aegis is an automated authenticity engine that inspects micro-acoustic artifacts left behind by "
        "generative models. Rather than assuming off-the-shelf models work everywhere, we followed an empirical "
        "baseline-first approach: we tested AASIST on real-world audio, uncovered its operational limitations, and "
        "are developing a domain-aware lightweight CNN for realistic deployment.\n\n"
        "LIKELY JUDGE QUESTION:\n"
        "What makes your approach different from conventional speaker recognition?\n"
        "STRONG ANSWER:\n"
        "Speaker recognition asks 'Who is speaking?' by comparing biometric voiceprints. SIH-Aegis asks 'Is this voice "
        "biologically human or synthetically generated?' It is an essential anti-spoofing defense layer that sits in front "
        "of speaker verification and identity workflows."
    )
    print("Slide 2 generated.")

    # ----------------------------------------------------
    # SLIDE 3: THE THREAT LANDSCAPE (F4, F8)
    # ----------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s3, "IDEA TITLE", "The Threat Landscape — 'Your Voice Can Be Copied' & Attack Vectors", 3)
    
    # Attack Chain Visual Box
    achain = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 1200000, 11300000, 950000)
    achain.fill.solid()
    achain.fill.fore_color.rgb = db.C_LIGHT_RED_BG
    achain.line.color.rgb = db.C_RED
    achain.line.width = Pt(1)
    tf_ac = achain.text_frame
    tf_ac.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_ac.margin_left = int(Pt(15))
    p_ac = tf_ac.paragraphs[0]
    p_ac.text = "ATTACK ANATOMY:  Real Person's Voice (3-5s sample)  ➔  Zero-Shot Neural Vocoder  ➔  Indistinguishable Clone  ➔  Target Exploitation"
    p_ac.font.name = db.FONT_BODY
    p_ac.font.size = Pt(12.5)
    p_ac.font.bold = True
    p_ac.font.color.rgb = db.C_RED
    
    # 5 Attack Vector Cards
    col_w = 2150000
    top_v = 2300000
    h_v = 3850000
    
    vectors = [
        ("FINANCIAL FRAUD", [
            "Executive Impersonation: CEO voice clones directing emergency wire transfers.",
            "Authorisation Bypass: Subverting phone-banking dual-factor approval.",
            "Scale: Global losses exceed hundreds of millions annually."
        ], db.C_RED),
        ("SOCIAL ENGINEERING", [
            "Fake Emergency Scams: 'Distress calls' to elderly parents claiming a crisis.",
            "Emotional Manipulation: Leverages biological familiarity to override caution.",
            "Speed: Demands instant money before victims can verify."
        ], db.C_ACCENT_ORANGE),
        ("VOICE AUTH SPOOFING", [
            "Biometric Spoofing: Fooling IVR systems that treat voice as a password.",
            "Account Takeover: Gaining unauthorized access to telecom & utility accounts.",
            "Silent Compromise: Victims unaware until funds or services are lost."
        ], db.C_SLATE_BLUE),
        ("CALL CENTRE TAKEOVER", [
            "Agent Deception: Synthetic callers manipulating tier-1 support staff.",
            "Credential Reset: Resetting 2FA passwords and contact numbers.",
            "Institutional Risk: Breaches customer databases at enterprise scale."
        ], db.C_DARK_NAVY),
        ("IDENTITY THEFT", [
            "Reputation Sabotage: Fabricating incriminating voice recordings.",
            "Digital Extortion: Blackmailing targets with synthetic compromising clips.",
            "Disinformation: Spreading fake statements attributed to officials."
        ], db.C_RED),
    ]
    
    for i, (title, bullets, acc_col) in enumerate(vectors):
        left_pos = 450000 + i * (col_w + 140000)
        db.add_card(s3, left_pos, top_v, col_w, h_v, title, bullets,
                    bg_color=db.C_CARD_BG, accent_color=acc_col, title_color=acc_col, font_size=9.5, title_size=11)

    db.set_notes(s3,
        "SPEAKER SCRIPT (30s):\n"
        "Voice cloning is not theoretical—it is active in the wild. Attackers harvest short audio clips from social "
        "media or phone calls, pass them through neural vocoders, and execute targeted exploitation. We have identified "
        "five distinct threat vectors: financial CEO fraud, emotional family distress scams, voice authentication bypass, "
        "call centre agent manipulation, and digital extortion. Each exploits the same core vulnerability: humans instinctively "
        "trust familiar vocal patterns.\n\n"
        "LIKELY JUDGE QUESTION:\n"
        "Can an ordinary listener tell if a modern voice clone is synthetic?\n"
        "STRONG ANSWER:\n"
        "No. Peer-reviewed acoustic studies show that modern neural vocoders fool human listeners in over 80% of blind "
        "trials. Human ears listen to semantics and melody; they cannot perceive the phase inconsistencies and high-frequency "
        "vocoder smearing that an AI detector identifies."
    )
    print("Slide 3 generated.")

    # ----------------------------------------------------
    # SLIDE 4: CORE PROBLEM DEFINITION (F1, F8)
    # ----------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s4, "IDEA TITLE", "Core Problem Definition — 'Can We Trust a Familiar Voice?'", 4)
    
    # 2 Comparison Columns
    col_w4 = 5500000
    top_4 = 1200000
    h_4 = 4000000
    
    db.add_card(s4, 450000, top_4, col_w4, h_4,
        "THE BROKEN LEGACY PARADIGM: HUMAN AUDITORY TRUST",
        [
            "Subjective Trust Assumption: 'It sounds like my boss, child, or colleague, so it must be genuine.'",
            "Physiological Blind Spots: Human auditory perception relies heavily on emotional tone, pitch, and familiarity, which generative AI synthesizes effortlessly.",
            "High Cognitive Vulnerability: Under panic, urgency, or authority pressure (e.g. 'Emergency wire transfer needed now'), human critical judgement completely breaks down.",
            "Traditional DSP Filters Fail: Handcrafted heuristic filters (simple pitch tracking, energy envelopes) cannot distinguish modern neural speech from biological phonation.",
            "Zero Mathematical Audit Trail: A human cannot log or prove why they believed a voice on an unrecorded phone call."
        ],
        bg_color=db.C_LIGHT_RED_BG, border_color=db.C_RED, accent_color=db.C_RED, title_color=db.C_RED, font_size=10.5
    )
    
    db.add_card(s4, 6250000, top_4, col_w4, h_4,
        "THE SIH-AEGIS PARADIGM: COMPUTABLE AUTHENTICITY",
        [
            "Objective Forensic Verification: Replaces subjective human belief with an automated mathematical classification.",
            "Micro-Acoustic Artifact Detection: Inspects phase coherence, spectral smearing, and synthesis anomalies invisible to the ear but evident in time-frequency representations.",
            "Quantified Security Metric: Outputs an auditable probability score alongside an unambiguous BONAFIDE vs SPOOF verdict.",
            "Passive Enterprise Integration: Sits upstream of banking authorizations, KYC verifications, and telecom customer support queues.",
            "Auditable Security Logs: Creates reviewable, timestamped authenticity records for high-stakes operational decisions."
        ],
        bg_color=db.C_LIGHT_GREEN_BG, border_color=db.C_GREEN, accent_color=db.C_GREEN, title_color=db.C_GREEN, font_size=10.5
    )
    
    # Bottom Center Callout
    c4 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 5400000, 11300000, 750000)
    c4.fill.solid()
    c4.fill.fore_color.rgb = db.C_DARK_NAVY
    c4.line.color.rgb = db.C_CYAN
    tf_c4 = c4.text_frame
    tf_c4.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_c4 = tf_c4.paragraphs[0]
    p_c4.text = "“Hearing a familiar voice is no longer enough to prove that the person is genuine.”"
    p_c4.alignment = PP_ALIGN.CENTER
    p_c4.font.name = db.FONT_TITLE
    p_c4.font.size = Pt(14)
    p_c4.font.bold = True
    p_c4.font.color.rgb = db.C_CYAN
    p_c4_sub = tf_c4.add_paragraph()
    p_c4_sub.text = "SIH-Aegis closes this critical security gap by computing objective authenticity before action is taken."
    p_c4_sub.alignment = PP_ALIGN.CENTER
    p_c4_sub.font.name = db.FONT_BODY
    p_c4_sub.font.size = Pt(11)
    p_c4_sub.font.color.rgb = db.C_WHITE

    db.set_notes(s4,
        "SPEAKER SCRIPT (30s):\n"
        "This slide captures the fundamental problem. For centuries, human society relied on voice as an instinctual "
        "identity signal. But neural generative models have broken this assumption. If an employee receives a call from "
        "their CFO demanding an urgent transfer, human instinct fails. SIH-Aegis introduces an objective, computable "
        "security layer that checks the physical and mathematical properties of the sound wave before the human ever acts.\n\n"
        "LIKELY JUDGE QUESTION:\n"
        "Why can't simple audio signal processing filters detect cloned voices?\n"
        "STRONG ANSWER:\n"
        "Basic DSP filters only check macro properties like fundamental frequency or silence ratios. Modern neural "
        "vocoders generate mathematically smooth, realistic waveforms that easily pass simple threshold checks. Detecting "
        "them requires deep representation learning across time and frequency."
    )
    print("Slide 4 generated.")

    # ----------------------------------------------------
    # SLIDE 5: CURRENT WORKING PROTOTYPE (F2, F5)
    # ----------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s5, "TECHNICAL APPROACH", "Current Working Prototype — 'Working Detection Pipeline' (Verified Flow)", 5)
    
    # 5 Horizontal Pipeline Step Boxes
    step_w = 2100000
    step_h = 1300000
    top_p = 1250000
    steps = [
        ("1. Audio Ingestion", "Upload WAV/MP3 or live mic recording via Streamlit", db.C_SLATE_BLUE),
        ("2. Preprocessing", "16 kHz Mono resampling, float32 normalization", db.C_SLATE_BLUE),
        ("3. AASIST Baseline", "Pretrained Graph Attention Network evaluated", db.C_CYAN),
        ("4. Deterministic Eval", "CUDA GPU inference on NVIDIA RTX 4050", db.C_CYAN),
        ("5. Security Output", "BONAFIDE / SPOOF verdict + calibrated confidence", db.C_GREEN),
    ]
    for i, (stitle, sdesc, scol) in enumerate(steps):
        left_s = 450000 + i * (step_w + 200000)
        box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_s, top_p, step_w, step_h)
        box.fill.solid()
        box.fill.fore_color.rgb = db.C_DARK_CARD_BG
        box.line.color.rgb = scol
        box.line.width = Pt(1.5)
        tf_s = box.text_frame
        tf_s.word_wrap = True
        tf_s.margin_left = int(Pt(8))
        tf_s.margin_right = int(Pt(8))
        tf_s.margin_top = int(Pt(8))
        p_st = tf_s.paragraphs[0]
        p_st.text = stitle
        p_st.font.name = db.FONT_BODY
        p_st.font.size = Pt(11)
        p_st.font.bold = True
        p_st.font.color.rgb = scol
        p_sd = tf_s.add_paragraph()
        p_sd.text = sdesc
        p_sd.font.name = db.FONT_BODY
        p_sd.font.size = Pt(9.5)
        p_sd.font.color.rgb = db.C_WHITE

    # 2 Detail Cards below pipeline
    top_cards = 2800000
    h_cards = 3350000
    w_card = 5500000
    
    db.add_card(s5, 450000, top_cards, w_card, h_cards,
        "VERIFIED PROTOTYPE CAPABILITIES (RUNNING TODAY)",
        [
            "Deep Learning Framework: PyTorch with verified CUDA acceleration on an NVIDIA RTX 4050 Laptop GPU (6 GB VRAM class).",
            "Baseline Model Active: Pretrained AASIST (Audio Anti-Spoofing using Integrated Spectro-Temporal Graph Attention Networks) checkpoint loads cleanly in deterministic eval mode.",
            "Feature Pipeline: Shared AudioLoader enforces 16,000 Hz mono standardization and float32 amplitude scaling [-1.0, 1.0].",
            "Interactive UI: Streamlit application running with upload, playback, spectrogram plotting, and live verdict display.",
            "Honest Classification: AASIST acts as our functional working prototype and reference benchmark, NOT our final primary model."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_SIH_BLUE, title_color=db.C_SIH_BLUE, font_size=10
    )
    
    db.add_card(s5, 6250000, top_cards, w_card, h_cards,
        "ONNX EXPORT & NUMERICAL EQUIVALENCE VERIFIED",
        [
            "Cross-Platform Export: AASIST graph successfully converted and exported to Open Neural Network Exchange (ONNX) format.",
            "Runtime Validation: ONNX Runtime engine executes the converted model with zero graph syntax or operator errors.",
            "Rigorous Equivalence Testing: PyTorch outputs and ONNX Runtime outputs evaluated on identical audio tensors.",
            "Max Numerical Divergence: 3.8e-06 (3.8 × 10⁻⁶) — proving strict mathematical equivalence across runtimes.",
            "Identical Class Predictions: Both execution engines produced identical BONAFIDE / SPOOF classifications.",
            "Mobile Readiness: Proves that our baseline inference path can execute on non-Python edge runtimes."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_GREEN, title_color=db.C_GREEN, font_size=10
    )

    db.set_notes(s5,
        "SPEAKER SCRIPT (30s):\n"
        "Here is what is running and verified today. Our end-to-end prototype ingests audio, resamples to 16 kHz mono, "
        "and runs deterministic inference through a pretrained AASIST baseline on an NVIDIA RTX 4050 Laptop GPU. "
        "Crucially for technical execution, we have already exported AASIST to ONNX and validated it in ONNX Runtime. "
        "The maximum numerical divergence between PyTorch and ONNX was only 3.8e-06, with 100% class agreement. "
        "This proves our engineering pipeline is mathematically sound and cross-platform capable.\n\n"
        "LIKELY JUDGE QUESTION:\n"
        "Is AASIST your final proposed model?\n"
        "STRONG ANSWER:\n"
        "No, AASIST is our working prototype and calibrated benchmark. As we show on the next slide, baseline testing "
        "uncovered critical domain-mismatch problems on real-world recordings, which is precisely why we are building "
        "our custom lightweight CNN detector."
    )
    print("Slide 5 generated.")

    # ----------------------------------------------------
    # SLIDE 6: WHAT WE LEARNED FROM THE BASELINE (F1, F2)
    # ----------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s6, "TECHNICAL APPROACH", "Baseline Observation — 'Why We Are Building Our Own Detector'", 6)
    
    # Two Branching Cards (Visual Contrast)
    col_w6 = 5500000
    top_6 = 1200000
    h_6a = 2100000
    
    db.add_card(s6, 450000, top_6, col_w6, h_6a,
        "TEST 1: KNOWN GENERATIVE SAMPLES (SYNTHETIC)",
        [
            "Input: High-fidelity speech samples generated by modern zero-shot TTS and voice-conversion models.",
            "Baseline AASIST Output: Accurately and decisively flagged as SPOOF with high probability.",
            "Observation: The baseline model performs well when evaluated against textbook synthetic audio."
        ],
        bg_color=db.C_LIGHT_GREEN_BG, border_color=db.C_GREEN, accent_color=db.C_GREEN, title_color=db.C_GREEN, font_size=10
    )
    
    db.add_card(s6, 6250000, top_6, col_w6, h_6a,
        "TEST 2: REAL-WORLD GENUINE MICROPHONE & CALL AUDIO",
        [
            "Input: Authentic human voices recorded on consumer smartphone microphones and phone call recordings.",
            "Baseline AASIST Output: Exhibited false-positive SPOOF predictions on genuine human speech!",
            "Observation: Baseline misclassifies benign hardware noise and room reverberation as synthetic artifacts."
        ],
        bg_color=db.C_LIGHT_RED_BG, border_color=db.C_RED, accent_color=db.C_RED, title_color=db.C_RED, font_size=10
    )
    
    # Central Insight Card
    c6_in = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 3450000, 11300000, 2700000)
    c6_in.fill.solid()
    c6_in.fill.fore_color.rgb = db.C_DARK_NAVY
    c6_in.line.color.rgb = db.C_ACCENT_ORANGE
    c6_in.line.width = Pt(1.5)
    tf_c6 = c6_in.text_frame
    tf_c6.word_wrap = True
    tf_c6.margin_left = int(Pt(18))
    tf_c6.margin_top = int(Pt(14))
    tf_c6.margin_right = int(Pt(18))
    
    p_ins1 = tf_c6.paragraphs[0]
    p_ins1.text = "CORE ENGINEERING INSIGHT:  Benchmark Performance ≠ Deployment Robustness"
    p_ins1.font.name = db.FONT_BODY
    p_ins1.font.size = Pt(14)
    p_ins1.font.bold = True
    p_ins1.font.color.rgb = db.C_ACCENT_ORANGE
    p_ins1.space_after = Pt(6)
    
    p_ins2 = tf_c6.add_paragraph()
    p_ins2.text = (
        "“We validated a pretrained anti-spoofing baseline (AASIST) under our target operational conditions and discovered "
        "that benchmark performance on academic datasets (like ASVspoof 2019 LA) does not automatically guarantee reliable "
        "deployment across real-world recording environments. Pretrained models overfit to studio acoustics and confuse "
        "microphone clipping, ambient reverb, and telephony codecs with generative synthesis. This critical discovery is "
        "what directly motivated the design of our custom, domain-aware Spectrogram CNN.”"
    )
    p_ins2.font.name = db.FONT_BODY
    p_ins2.font.size = Pt(11)
    p_ins2.font.color.rgb = db.C_WHITE
    p_ins2.space_after = Pt(8)
    
    p_ins3 = tf_c6.add_paragraph()
    p_ins3.text = "STRATEGIC TAKEAWAY: We do not discard AASIST—we retain it as a calibrated benchmark to measure our custom model against."
    p_ins3.font.name = db.FONT_BODY
    p_ins3.font.size = Pt(10.5)
    p_ins3.font.bold = True
    p_ins3.font.color.rgb = db.C_CYAN

    db.set_notes(s6,
        "SPEAKER SCRIPT (30s):\n"
        "This is our most important engineering finding for Round 1. We did not simply trust the literature. "
        "When we evaluated the state-of-the-art AASIST model on genuine phone and laptop recordings, it generated "
        "false-positive spoof alarms on real human beings. Academic datasets are recorded in clean acoustic conditions. "
        "In real life, phone mics, room echo, and cellular codecs cause distribution shift. This proves that an out-of-the-box "
        "benchmark cannot be deployed blindly. It demands a domain-aware, noise-augmented, lightweight custom detector.\n\n"
        "LIKELY JUDGE QUESTION:\n"
        "Why didn't you just keep AASIST if it's state-of-the-art?\n"
        "STRONG ANSWER:\n"
        "AASIST is designed for high-end servers and clean datasets. It has ~297k parameters and complex graph attention "
        "operations that are hard to optimize on mobile devices and vulnerable to acoustic distribution shifts. We retain "
        "AASIST as our benchmark reference, while engineering a lightweight ~97.9k CNN tailored for real-world robustness."
    )
    print("Slide 6 generated.")

    # ----------------------------------------------------
    # SLIDE 7: LEVEL 3 AI DIRECTION — SPECTROGRAM + CNN (F1, F2, F5)
    # ----------------------------------------------------
    s7 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s7, "TECHNICAL APPROACH", "Level 3 AI Direction — 'Learning Time-Frequency Patterns' (Spectrogram + CNN)", 7)
    
    # Feature Flow diagram box
    fflow = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 1200000, 11300000, 900000)
    fflow.fill.solid()
    fflow.fill.fore_color.rgb = db.C_LIGHT_BLUE_BG
    fflow.line.color.rgb = db.C_SIH_BLUE
    fflow.line.width = Pt(1)
    tf_ff = fflow.text_frame
    tf_ff.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_ff.margin_left = int(Pt(15))
    p_ff = tf_ff.paragraphs[0]
    p_ff.text = "TRANSFORMATION PIPELINE:  1D Raw Audio  ➔  STFT Fourier Transform  ➔  Log-Mel Filterbank (128 bins)  ➔  128×128 Tensor  ➔  Lightweight 2D-CNN"
    p_ff.font.name = db.FONT_BODY
    p_ff.font.size = Pt(11.5)
    p_ff.font.bold = True
    p_ff.font.color.rgb = db.C_SIH_BLUE
    
    # 2 Feature Detail Cards
    top_7 = 2250000
    h_7 = 3900000
    w_7 = 5500000
    
    db.add_card(s7, 450000, top_7, w_7, h_7,
        "WHY CONVERT VOICE INTO A LOG-MEL SPECTROGRAM?",
        [
            "Dimensional Transformation: Converts continuous 1D air pressure signals into rich 2D time-frequency energy images.",
            "Capturing Vocoder Artifacts: Neural synthesis engines leave subtle mathematical footprints—such as spectral smearing, unnatural harmonics, and phase discontinuities—clearly visible in spectrograms.",
            "Psychoacoustic Mel Scale: Mimics human cochlear frequency perception, concentrating analytical resolution in critical vocal formant bands (300 Hz – 3,400 Hz).",
            "Spatial Pattern Learning: Enables proven 2D Convolutional Neural Networks to extract local acoustic textures without massive 1D recurrent sequence models.",
            "Drastic Data Compression: Condenses thousands of raw waveform samples into a dense, standardized 128 × 128 float32 representation."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_SIH_BLUE, title_color=db.C_SIH_BLUE, font_size=10.5
    )
    
    db.add_card(s7, 6250000, top_7, w_7, h_7,
        "EXACT SPECTROGRAM EXTRACTION SPECIFICATIONS",
        [
            "Audio Sampling Rate: Standardized to 16,000 Hz mono via polyphase resampling.",
            "Mel Filterbank Bins: 128 frequency bands covering 0 Hz to 8,000 Hz (Nyquist limit).",
            "FFT Window Size (n_fft): 512 samples (~32 ms window) for optimal frequency resolution.",
            "Frame Hop Length: 160 samples (10 ms temporal stride) ensuring 68% frame overlap.",
            "Standardized Target Shape: 128 × 128 float32 tensor via dynamic padding/clipping.",
            "Per-Sample Standardization: Z-score normalization (μ=0, σ=1) applied per spectrogram to remove volume and mic gain disparities.",
            "Engineering Status: Spectrogram extractor fully implemented and unit-tested in `src/features/spectrogram.py`."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_DARK_NAVY, title_color=db.C_DARK_NAVY, font_size=10.5
    )

    db.set_notes(s7,
        "SPEAKER SCRIPT (30s):\n"
        "To build a detector that is fast and accurate, we convert speech into a 2D Log-Mel spectrogram. "
        "Sound is naturally a 1D wave, but generative cloning tools leave distinctive artifacts across frequency "
        "and time—such as unnatural harmonic continuity and high-frequency phase smearing. By converting 16 kHz audio "
        "into a 128x128 Log-Mel tensor, we can leverage 2D computer vision techniques to learn acoustic textures at a "
        "fraction of the computational cost of raw-waveform networks.\n\n"
        "LIKELY JUDGE QUESTION:\n"
        "Why use a 2D spectrogram CNN instead of raw audio or transformers?\n"
        "STRONG ANSWER:\n"
        "Raw audio transformers (like Wav2Vec2) require tens of millions of parameters and heavy GPU memory, making real-time "
        "mobile inference impossible. A 128x128 spectrogram captures the essential vocoder artifacts while allowing our CNN to "
        "operate with just ~97.9k parameters—ideal for sub-second, edge execution."
    )
    print("Slide 7 generated.")

    # ----------------------------------------------------
    # SLIDE 8: CNN ARCHITECTURE SPECIFICATION (F2, F5)
    # ----------------------------------------------------
    s8 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s8, "TECHNICAL APPROACH", "CNN Architecture Specification — 4-Block Neural Classifier (~97,890 Parameters)", 8)
    
    # Layer Breakdown Card
    top_8 = 1200000
    h_8 = 4000000
    w_8a = 5900000
    w_8b = 5100000
    
    db.add_card(s8, 450000, top_8, w_8a, h_8,
        "DETECTION MODEL: LAYER-BY-LAYER TOPOLOGY",
        [
            "Input Tensor: [Batch × 1 × 128 × 128] float32 normalized spectrogram.",
            "Conv Block 1: Conv2D (1 → 16, 3×3, pad 1) + BatchNorm2d + ReLU + MaxPool(2×2) ➔ Output: [16 × 64 × 64].",
            "Conv Block 2: Conv2D (16 → 32, 3×3, pad 1) + BatchNorm2d + ReLU + MaxPool(2×2) ➔ Output: [32 × 32 × 32].",
            "Conv Block 3: Conv2D (32 → 64, 3×3, pad 1) + BatchNorm2d + ReLU + MaxPool(2×2) ➔ Output: [64 × 16 × 16].",
            "Conv Block 4: Conv2D (64 → 128, 3×3, pad 1) + BatchNorm2d + ReLU ➔ Output: [128 × 16 × 16].",
            "Pooling Layer: AdaptiveAvgPool2d(1×1) ➔ Compresses spatial grid to [128 × 1 × 1].",
            "Classifier Head: Flatten ➔ Dropout (p=0.3) ➔ Linear (128 → 2 logits) ➔ Softmax (BONAFIDE / SPOOF).",
            "Deterministic Forward Pass: Verified on NVIDIA RTX 4050 Laptop GPU (CUDA) with zero tensor shape mismatch."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_CYAN, title_color=db.C_DARK_NAVY, font_size=10
    )
    
    db.add_card(s8, 6650000, top_8, w_8b, h_8,
        "ARCHITECTURAL EFFICIENCY & VERIFICATION",
        [
            "Ultra-Compact Parameter Count: Exactly 97,890 trainable parameters (~390 KB uncompressed weights size).",
            "Model Footprint Comparison:",
            "  • RawNet2: ~1,250,000 parameters (>5 MB)",
            "  • AASIST: ~297,000 parameters (~1.2 MB)",
            "  • SIH-Aegis CNN: ~97,890 parameters (~0.39 MB)",
            "VRAM Footprint: Less than 50 MB during GPU forward execution.",
            "Edge Suitability: Compact enough for real-time inference on mobile devices and entry-level server containers.",
            "Fast Retraining Cycles: Retraining on new cloning attacks takes hours on a single GPU instead of days on a cluster.",
            "CRITICAL HONEST STATUS: The CNN architecture is fully implemented and forward-pass verified on GPU; the training pipeline and dataset preparation are currently in progress."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_GREEN, title_color=db.C_GREEN, font_size=9.8
    )
    
    # Bottom Note
    b8 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 5350000, 11300000, 800000)
    b8.fill.solid()
    b8.fill.fore_color.rgb = db.C_LIGHT_BLUE_BG
    b8.line.color.rgb = db.C_SIH_BLUE
    tf_b8 = b8.text_frame
    tf_b8.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_b8 = tf_b8.paragraphs[0]
    p_b8.text = "TECHNICAL HONESTY: Architecture implemented & GPU forward pass PASS. NO trained checkpoint exists yet; NO fake accuracy is reported."
    p_b8.alignment = PP_ALIGN.CENTER
    p_b8.font.name = db.FONT_BODY
    p_b8.font.size = Pt(11)
    p_b8.font.bold = True
    p_b8.font.color.rgb = db.C_SIH_BLUE

    db.set_notes(s8,
        "SPEAKER SCRIPT (30s):\n"
        "Here is the exact topology of our primary model. It is a 4-block 2D CNN with Batch Normalization, ReLU, and MaxPool, "
        "funneling into an Adaptive Average Pool and a Dropout-regularized Linear classifier. The total footprint is just ~97,890 "
        "parameters—under 400 kilobytes. This is 3 times smaller than AASIST and 12 times smaller than RawNet2. We have verified "
        "the forward pass on our RTX 4050 GPU. Being technically honest, the architecture is coded and verified, while training "
        "and dataset preparation are actively underway.\n\n"
        "LIKELY JUDGE QUESTION:\n"
        "Is this CNN model trained already? What is its accuracy?\n"
        "STRONG ANSWER:\n"
        "We are technically honest: the CNN architecture is coded and GPU-verified, but it is NOT yet trained. We do not display "
        "invented accuracy numbers. Our current functional prototype uses the verified AASIST baseline while our CNN training "
        "pipeline is being prepared."
    )
    print("Slide 8 generated.")

    # ----------------------------------------------------
    # SLIDE 9: SYSTEM ARCHITECTURE (F2, F5)
    # ----------------------------------------------------
    s9 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s9, "TECHNICAL APPROACH", "System Architecture — 'Modular SIH-Aegis Architecture' (Separation of Concerns)", 9)
    
    # 6 Layer Horizontal Architecture Stack
    arch_layers = [
        ("1. Presentation Layer", "Streamlit Web Dashboard (live recording, file upload, waveform plotting)  |  Android Jetpack Compose UI", db.C_SLATE_BLUE),
        ("2. Orchestration Layer (src/pipeline)", "DetectorPipeline: coordinates audio ingestion, feature extraction, model inference, and score calibration", db.C_SLATE_BLUE),
        ("3. Signal Processing (src/audio)", "AudioLoader: format validation, polyphase 16 kHz resampling, PCM normalization, dynamic channel mixing", db.C_CYAN),
        ("4. Feature Engineering (src/features)", "Log-Mel Spectrogram Generator: STFT calculation, 128 Mel filterbanks, 128×128 tensor standardizer", db.C_CYAN),
        ("5. Model Inference (src/models)", "Spectrogram CNN (Primary Development Direction)  |  AASIST Pretrained Detector (Verified Baseline)", db.C_GREEN),
        ("6. Quality & Verification (tests/)", "Isolated unit tests: tests/audio, tests/features, tests/models, tests/pipeline ensuring zero regression", db.C_DARK_NAVY),
    ]
    
    top_arch = 1200000
    layer_h = 700000
    for i, (ltitle, ldesc, lcol) in enumerate(arch_layers):
        lbox = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, top_arch + i * (layer_h + 100000), 11300000, layer_h)
        lbox.fill.solid()
        lbox.fill.fore_color.rgb = db.C_CARD_BG
        lbox.line.color.rgb = lcol
        lbox.line.width = Pt(1.5)
        tf_l = lbox.text_frame
        tf_l.word_wrap = True
        tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_l.margin_left = int(Pt(14))
        p_lt = tf_l.paragraphs[0]
        p_lt.text = ltitle + " — "
        p_lt.font.name = db.FONT_BODY
        p_lt.font.size = Pt(11.5)
        p_lt.font.bold = True
        p_lt.font.color.rgb = lcol
        r_ld = p_lt.add_run()
        r_ld.text = ldesc
        r_ld.font.name = db.FONT_BODY
        r_ld.font.size = Pt(10.5)
        r_ld.font.bold = False
        r_ld.font.color.rgb = db.C_DARK_TEXT
        
    # Key Takeaway Box
    b9 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 5600000, 11300000, 550000)
    b9.fill.solid()
    b9.fill.fore_color.rgb = db.C_DARK_NAVY
    b9.line.fill.background()
    tf_b9 = b9.text_frame
    tf_b9.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_b9 = tf_b9.paragraphs[0]
    p_b9.text = "ARCHITECTURAL PRINCIPLE: The audio front-end and UI remain identical when models are upgraded or swapped behind the interface."
    p_b9.alignment = PP_ALIGN.CENTER
    p_b9.font.name = db.FONT_BODY
    p_b9.font.size = Pt(11)
    p_b9.font.bold = True
    p_b9.font.color.rgb = db.C_CYAN

    db.set_notes(s9,
        "SPEAKER SCRIPT (30s):\n"
        "Our codebase is designed around clean separation of concerns. The presentation layer (Streamlit and Android) "
        "interacts solely with the pipeline orchestrator. Audio loading, feature engineering, and model inference are "
        "completely modular. When we upgrade from our AASIST benchmark to our trained CNN, the audio loader and user "
        "interface require zero changes. Furthermore, our matching test tree isolates bugs immediately—preventing audio "
        "sampling errors from being mistaken for model bugs.\n\n"
        "LIKELY JUDGE QUESTION:\n"
        "How do you test and isolate bugs across this system?\n"
        "STRONG ANSWER:\n"
        "We have independent unit tests for each layer in tests/audio, tests/models, and tests/pipeline. We verify audio "
        "resampling and spectrogram shapes before any tensor touches a neural network."
    )
    print("Slide 9 generated.")

    # ----------------------------------------------------
    # SLIDE 10: TECHNOLOGY STACK & HARDWARE (F5)
    # ----------------------------------------------------
    s10 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s10, "TECHNICAL APPROACH", "Technology Stack & Hardware Acceleration — Production-Grade Engineering", 10)
    
    # 4 Quadrant Cards
    qw = 5500000
    qh = 2300000
    top_q1 = 1200000
    top_q2 = 3650000
    left_q1 = 450000
    left_q2 = 6250000
    
    db.add_card(s10, left_q1, top_q1, qw, qh,
        "CORE DEEP LEARNING & NUMERICS",
        [
            "PyTorch 2.x: Model graph construction, tensor autograd, GPU execution.",
            "NVIDIA CUDA 12.x: Hardware-accelerated matrix multiplication & convolution.",
            "ONNX Runtime: Cross-platform portable graph inference engine.",
            "NumPy & SciPy: Array operations, polyphase resampling, signal filtering."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_SIH_BLUE, title_color=db.C_SIH_BLUE, font_size=10
    )
    
    db.add_card(s10, left_q2, top_q1, qw, qh,
        "AUDIO & DIGITAL SIGNAL PROCESSING",
        [
            "librosa: Audio signal processing, Short-Time Fourier Transform (STFT), Mel filterbanks.",
            "SoundFile: High-performance multi-format audio I/O (WAV, FLAC, OGG).",
            "Standardization: 16,000 Hz mono standard, float32 PCM dynamic range.",
            "AASIST Pretrained Library: Baseline Graph Attention Network implementation."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_CYAN, title_color=db.C_CYAN, font_size=10
    )
    
    db.add_card(s10, left_q1, top_q2, qw, qh,
        "MOBILE & CLIENT INTERFACES",
        [
            "Streamlit: Functional web prototype for audio upload, real-time spectrogram plotting, and verdict display.",
            "Android Studio: Modern native mobile client development environment.",
            "Kotlin & Jetpack Compose: Declarative, reactive mobile user interface.",
            "Material 3: Adaptive theme and security status feedback widgets."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_DARK_NAVY, title_color=db.C_DARK_NAVY, font_size=10
    )
    
    db.add_card(s10, left_q2, top_q2, qw, qh,
        "VERIFIED HARDWARE & TOOLING",
        [
            "NVIDIA RTX 4050 Laptop GPU: 6 GB VRAM class, CUDA acceleration verified.",
            "Deterministic Evaluation Mode: Fixed random seeds and CUDA algorithms for reproducible evaluation.",
            "Version Control & CI: Git, GitHub repository with clean branch architecture.",
            "Pytest Suite: Automated test execution verifying tensor integrity across pipeline."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_GREEN, title_color=db.C_GREEN, font_size=10
    )

    db.set_notes(s10,
        "SPEAKER SCRIPT (30s):\n"
        "Our technology stack leverages enterprise-grade tools. In the core deep learning layer, we use PyTorch and CUDA "
        "on our development hardware—an NVIDIA RTX 4050 Laptop GPU with 6 GB VRAM. For audio processing, librosa and SoundFile "
        "handle mathematical signal transformation. For interfaces, we run Streamlit on desktop and Kotlin with Jetpack Compose "
        "on Android. Importantly, our entire codebase is tracked on GitHub with deterministic Pytest verification.\n\n"
        "LIKELY JUDGE QUESTION:\n"
        "What hardware is required to run your detector in production?\n"
        "STRONG ANSWER:\n"
        "Because our Spectrogram CNN requires only ~97.9k parameters, it does NOT require expensive server clusters. "
        "It can execute on a single commodity GPU or even standard modern server CPUs via ONNX Runtime with sub-second latency."
    )
    print("Slide 10 generated.")

    # ----------------------------------------------------
    # SLIDE 11: ANDROID IMPLEMENTATION (F3, F5)
    # ----------------------------------------------------
    s11 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s11, "TECHNICAL APPROACH", "Android Mobile Implementation — 'From Prototype to Mobile Security'", 11)
    
    top_11 = 1200000
    h_11 = 3900000
    w_11 = 5500000
    
    db.add_card(s11, 450000, top_11, w_11, h_11,
        "CURRENT ANDROID MILESTONES (COMPLETED)",
        [
            "Native Android Application: Built in modern Kotlin using Jetpack Compose and Material 3 design system.",
            "ONNX Runtime Integration: ONNX Runtime Mobile library integrated and initializing successfully on Android.",
            "Model Bundled into Assets: AASIST ONNX model successfully bundled inside Android assets and loaded into mobile memory.",
            "Reactive State Machine: UI supports four distinct security states: READY ➔ LISTENING ➔ ANALYZING ➔ BONAFIDE / SPOOF.",
            "Demonstration Target: Audio file selection ➔ ONNX Runtime inference on-device ➔ Instant security risk display.",
            "Next Stage: Integrating audio recording pipeline to feed microphone chunks into on-device ONNX inference."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_SIH_BLUE, title_color=db.C_SIH_BLUE, font_size=10.2
    )
    
    db.add_card(s11, 6250000, top_11, w_11, h_11,
        "TECHNICAL HONESTY: PLATFORM REALITIES & ROADMAP",
        [
            "CRITICAL HONEST FACT: Third-party Android applications CANNOT freely intercept arbitrary cellular-call audio streams due to OS-level security sandboxing and privacy restrictions.",
            "Target Operational Workflows: Focus is placed where third-party apps can operate legally and technically:",
            "  • Ingesting recorded voice notes (WhatsApp, Telegram audio).",
            "  • Screening uploaded audio attachments and voicemail files.",
            "  • VoIP and live speakerphone microphone audio capture.",
            "Enterprise / Carrier Vision: Full cellular stream inspection requires carrier-level telecom gateway deployment or OEM system-level privileges.",
            "Current Progress Status: Android UI and ONNX engine loading complete; live audio chunk streaming under active integration."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_ACCENT_ORANGE, title_color=db.C_ACCENT_ORANGE, font_size=10
    )
    
    # Bottom Callout
    b11 = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 5250000, 11300000, 900000)
    b11.fill.solid()
    b11.fill.fore_color.rgb = db.C_DARK_NAVY
    b11.line.color.rgb = db.C_CYAN
    tf_b11 = b11.text_frame
    tf_b11.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_b11 = tf_b11.paragraphs[0]
    p_b11.text = "CURRENT DEMONSTRATION: Audio Upload ➔ Android ONNX Model ➔ On-Device Security Verdict"
    p_b11.alignment = PP_ALIGN.CENTER
    p_b11.font.name = db.FONT_BODY
    p_b11.font.size = Pt(12)
    p_b11.font.bold = True
    p_b11.font.color.rgb = db.C_CYAN
    p_b11_sub = tf_b11.add_paragraph()
    p_b11_sub.text = "We do NOT claim cellular-call interception is solved. We demonstrate transparent, verified mobile progress."
    p_b11_sub.alignment = PP_ALIGN.CENTER
    p_b11_sub.font.name = db.FONT_BODY
    p_b11_sub.font.size = Pt(10)
    p_b11_sub.font.color.rgb = db.C_WHITE

    db.set_notes(s11,
        "SPEAKER SCRIPT (30s):\n"
        "For our mobile layer, we built a native Android app in Kotlin and Jetpack Compose. We successfully integrated "
        "the ONNX Runtime Mobile SDK and loaded our AASIST ONNX model into memory. Crucially, we maintain total technical "
        "honesty: standard Android apps cannot record cellular phone calls due to Android security restrictions. We are "
        "not making false claims about cellular call interception. Our target demonstration is on-device audio upload "
        "and microphone verification for voice notes, VoIP, and suspicious audio files."
    )
    print("Slide 11 generated.")

    # ----------------------------------------------------
    # SLIDE 12: ETHICAL DATA COLLECTION PROTOCOL (F2, F8)
    # ----------------------------------------------------
    s12 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s12, "TECHNICAL APPROACH", "Ethical Data Collection Protocol — 'Building Balanced Indian Speech Corpora'", 12)
    
    # 5-Stage Data Acquisition Workflow
    steps_12 = [
        ("1. Speaker Consent", "Explicit digital consent & privacy policy acceptance", db.C_SLATE_BLUE),
        ("2. Language Select", "Selection of native language & regional dialect", db.C_SLATE_BLUE),
        ("3. Guided Prompts", "Display of phonetically balanced 5-10s sentences", db.C_CYAN),
        ("4. Sample Recording", "High-fidelity recording with acoustic validation", db.C_CYAN),
        ("5. Encrypted Storage", "Anonymized storage linked only to session ID", db.C_GREEN),
    ]
    step_w12 = 2100000
    top_12 = 1250000
    for i, (stitle, sdesc, scol) in enumerate(steps_12):
        left_s = 450000 + i * (step_w12 + 200000)
        box = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_s, top_12, step_w12, 1200000)
        box.fill.solid()
        box.fill.fore_color.rgb = db.C_DARK_CARD_BG
        box.line.color.rgb = scol
        box.line.width = Pt(1.5)
        tf_s = box.text_frame
        tf_s.word_wrap = True
        tf_s.margin_left = int(Pt(8))
        tf_s.margin_top = int(Pt(8))
        p_st = tf_s.paragraphs[0]
        p_st.text = stitle
        p_st.font.name = db.FONT_BODY
        p_st.font.size = Pt(11)
        p_st.font.bold = True
        p_st.font.color.rgb = scol
        p_sd = tf_s.add_paragraph()
        p_sd.text = sdesc
        p_sd.font.name = db.FONT_BODY
        p_sd.font.size = Pt(9.5)
        p_sd.font.color.rgb = db.C_WHITE

    # 2 Detail Cards below workflow
    top_c12 = 2650000
    h_c12 = 3500000
    w_c12 = 5500000
    
    db.add_card(s12, 450000, top_c12, w_c12, h_c12,
        "STRUCTURED METADATA SCHEMA (ANONYMIZED)",
        [
            "Anonymized Session Identifier: Non-revisable hash decoupling biometric data from personal user identity.",
            "Linguistic Labeling: Primary language, regional dialect, and phonetic prompt identifier.",
            "Hardware Profile: Device model, internal mic hardware specs, input sampling frequency.",
            "Acoustic Environment Tag: Indoor office, outdoor street noise, quiet room, reverberant hall.",
            "Phonetic Diversity: Prompts designed to trigger varied fricatives, plosives, and nasals.",
            "Honest Status: Planned data-collection framework; large-scale dataset acquisition in active planning."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_SIH_BLUE, title_color=db.C_SIH_BLUE, font_size=10
    )
    
    db.add_card(s12, 6250000, top_c12, w_c12, h_c12,
        "CONTROLLED PAIRED SYNTHESIS GENERATION",
        [
            "Balanced Genuine-Spoof Pairs: Each collected genuine voice sample is cloned using leading open-source neural models (e.g. VITS, XTTS, Bark, YourTTS).",
            "Multi-Vocoder Exposure: Prevents model from memorizing a single synthesis footprint by utilizing diverse generative architectures.",
            "Ground Truth Calibration: Creates rigorously verified training pairs with identical semantic text, isolating acoustic synthesis artifacts.",
            "Zero Hallucinated Training Data: Every spoofed sample has a verifiable genuine counterpart for controlled EER (Equal Error Rate) evaluation."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_GREEN, title_color=db.C_GREEN, font_size=10
    )

    db.set_notes(s12,
        "SPEAKER SCRIPT (30s):\n"
        "To solve the dataset problem, we designed an ethical, consent-driven data collection protocol. "
        "Through a guided mode, contributors read standardized 5-10 second phonetic prompts in their native languages. "
        "All data is anonymized and metadata-tagged with device and noise profiles. To create balanced training pairs, "
        "we clone these genuine recordings using multiple open-source vocoders. This gives us balanced genuine-versus-spoof "
        "pairs with identical words, forcing the model to learn synthesis artifacts rather than phonetic shortcuts."
    )
    print("Slide 12 generated.")

    # ----------------------------------------------------
    # SLIDE 13: ENGINEERING & DATA STRATEGY (F2, F6)
    # ----------------------------------------------------
    s13 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s13, "FEASIBILITY AND VIABILITY", "Engineering & Data Strategy — 'Designed for the Real World'", 13)
    
    # 4 Challenge vs Strategy Cards
    w_13 = 2650000
    h_13 = 4000000
    top_13 = 1200000
    
    challenges = [
        ("DATASET MISMATCH", [
            "Challenge: Academic sets (ASVspoof) lack Indian accent and linguistic diversity.",
            "Strategy: Deliberately curating Indian English and regional language speech corpora.",
            "Result: Prevents models from flagging Indian phonetics as synthesis anomalies."
        ], db.C_RED),
        ("CHANNEL & NOISE", [
            "Challenge: Clean studio models fail in street traffic, fan hum, or room echo.",
            "Strategy: Dynamic noise augmentation (additive Gaussian, room impulse response simulation).",
            "Result: Model learns noise-invariant vocal tract features."
        ], db.C_ACCENT_ORANGE),
        ("CODEC COMPRESSION", [
            "Challenge: Cellular and VoIP audio uses lossy codecs (AMR-NB, G.711, Opus) that alter audio.",
            "Strategy: Injecting multi-codec re-compression during training batch pipelines.",
            "Result: High resilience across cellular networks and messaging apps."
        ], db.C_SLATE_BLUE),
        ("UNSEEN CLONING TOOLS", [
            "Challenge: New TTS algorithms appear weekly, bypassing static detectors.",
            "Strategy: Training against multiple diverse vocoders (diffusion, autoregressive, GAN).",
            "Result: Generalization across underlying vocoder phase artifacts."
        ], db.C_GREEN),
    ]
    
    for i, (ctitle, cbullets, ccol) in enumerate(challenges):
        left_c = 450000 + i * (w_13 + 230000)
        db.add_card(s13, left_c, top_13, w_13, h_13, ctitle, cbullets,
                    bg_color=db.C_CARD_BG, accent_color=ccol, title_color=ccol, font_size=10, title_size=12)
        
    # Formula Banner
    bf13 = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 5400000, 11300000, 750000)
    bf13.fill.solid()
    bf13.fill.fore_color.rgb = db.C_DARK_NAVY
    bf13.line.color.rgb = db.C_CYAN
    tf_bf13 = bf13.text_frame
    tf_bf13.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_bf13 = tf_bf13.paragraphs[0]
    p_bf13.text = "ROBUSTNESS EQUATION: Representative Speech + Multi-Vocoder Spoofs + Codec/Noise Augmentation = Real-World Generalization"
    p_bf13.alignment = PP_ALIGN.CENTER
    p_bf13.font.name = db.FONT_BODY
    p_bf13.font.size = Pt(11.5)
    p_bf13.font.bold = True
    p_bf13.font.color.rgb = db.C_WHITE

    db.set_notes(s13,
        "SPEAKER SCRIPT (30s):\n"
        "Real-world generalization is the hardest problem in audio anti-spoofing. Our strategy directly tackles four "
        "operational failure modes: dataset bias, background noise, codec compression, and zero-day cloning engines. "
        "Instead of hoping our model generalizes, we actively augment our training data with telephony codecs like AMR "
        "and G.711, inject room reverberation, and train across multiple synthesis engines. Our goal is not high accuracy "
        "on one benchmark; our goal is robustness across real Indian acoustic conditions.\n\n"
        "LIKELY JUDGE QUESTION:\n"
        "How will you evaluate false positives on genuine callers?\n"
        "STRONG ANSWER:\n"
        "We evaluate our system using Equal Error Rate (EER) and minimum Detection Cost Function (minDCF) across diverse "
        "unseen genuine speakers, setting conservative thresholds that strictly penalize false-positive rejections in financial flows."
    )
    print("Slide 13 generated.")

    # ----------------------------------------------------
    # SLIDE 14: SECURITY & PRIVACY BY DESIGN (F8)
    # ----------------------------------------------------
    s14 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s14, "FEASIBILITY AND VIABILITY", "Security & Privacy by Design — Zero-Trust Principles for Biometric Speech", 14)
    
    top_14 = 1200000
    h_14 = 4000000
    w_14 = 5500000
    
    db.add_card(s14, 450000, top_14, w_14, h_14,
        "SIX CORE BIOMETRIC PRIVACY SAFEGUARDS",
        [
            "1. Explicit Informed Consent: Mandatory opt-in before any audio is ingested, analyzed, or stored.",
            "2. Ephemeral In-Memory Processing: Raw audio waveforms are processed in volatile RAM and purged immediately post-inference.",
            "3. Non-Invertible Representation: 128×128 Log-Mel spectrograms mathematically strip linguistic semantics—raw private speech cannot be reconstructed.",
            "4. Cryptographic At-Rest Protection: All development and calibration samples are encrypted using AES-256.",
            "5. Strict Identity Decoupling: Biometric acoustic features are completely isolated from personal identifiers (PII).",
            "6. Right to Complete Deletion: Immediate data purge protocol for research participants upon request."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_SIH_BLUE, title_color=db.C_SIH_BLUE, font_size=10.5
    )
    
    db.add_card(s14, 6250000, top_14, w_14, h_14,
        "REGULATORY ALIGNMENT & ETHICAL GOVERNANCE",
        [
            "Digital Personal Data Protection (DPDPA 2023): Designed around Indian data governance requirements for processing sensitive biometric information.",
            "Zero Cloud Dependency: Edge and on-premise inference ensures private customer voice data never leaves the institution's secure network perimeter.",
            "Tamper-Evident Audit Logging: Security outputs (score, timestamp, model version) are logged cryptographically without retaining raw voice recordings.",
            "Non-Surveillance Architecture: System is designed strictly for incoming audio verification—never background passive wiretapping.",
            "Ethical Defense Focus: Purpose-built solely to defend individuals and enterprises against fraudulent exploitation."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_GREEN, title_color=db.C_GREEN, font_size=10.5
    )
    
    b14 = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 5350000, 11300000, 800000)
    b14.fill.solid()
    b14.fill.fore_color.rgb = db.C_DARK_NAVY
    b14.line.fill.background()
    tf_b14 = b14.text_frame
    tf_b14.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_b14 = tf_b14.paragraphs[0]
    p_b14.text = "PRIVACY PRINCIPLE: Voice is sensitive biometric data. We verify authenticity without storing or exposing private speech."
    p_b14.alignment = PP_ALIGN.CENTER
    p_b14.font.name = db.FONT_BODY
    p_b14.font.size = Pt(11.5)
    p_b14.font.bold = True
    p_b14.font.color.rgb = db.C_WHITE

    db.set_notes(s14,
        "SPEAKER SCRIPT (30s):\n"
        "Voice is biometric data, and protecting it is central to SIH-Aegis. We follow privacy-by-design under India's "
        "DPDPA 2023. First, we operate with ephemeral processing: audio is held in RAM only during inference and purged immediately. "
        "Second, our Log-Mel spectrograms are non-invertible—an attacker cannot reconstruct the conversation from the spectrogram. "
        "Third, because our model is lightweight, inference can run completely on-device or on-premise without transmitting voice data to the cloud."
    )
    print("Slide 14 generated.")

    # ----------------------------------------------------
    # SLIDE 15: IMPACT & BUSINESS VIABILITY (F4, F7)
    # ----------------------------------------------------
    s15 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s15, "IMPACT AND BENEFITS", "Real-World Impact & Use Cases — Operational Integration & Business Viability", 15)
    
    # 6 Application Cards (2 rows of 3)
    cw15 = 3600000
    ch15 = 1950000
    top_r1 = 1200000
    top_r2 = 3300000
    
    use_cases = [
        ("BANKING & FINANCIAL SECURITY", [
            "Screens voice-authorized wire transfers before funds move.",
            "Validates callback verifications for high-value transactions.",
            "Directly mitigates millions in authorized push-payment fraud."
        ], db.C_RED, 450000, top_r1),
        ("KYC & REMOTE ONBOARDING", [
            "Adds acoustic liveness checks to video/audio KYC.",
            "Prevents synthetic video avatars with cloned speech.",
            "Protects non-banking financial companies (NBFCs) & fintech."
        ], db.C_SIH_BLUE, 4300000, top_r1),
        ("CALL CENTRE & IVR DEFENSE", [
            "Flags synthetic callers attempting account takeover.",
            "Alerts support agents in real time to suspicious synthetic audio.",
            "Reduces manual audit overhead and agent social-engineering risk."
        ], db.C_DARK_NAVY, 8150000, top_r1),
        ("CITIZEN SCAM PREVENTION", [
            "Assists citizens in verifying suspicious emergency calls.",
            "Provides evidence triage for cybercrime reporting helplines.",
            "Protects elderly and vulnerable populations from extortion."
        ], db.C_ACCENT_ORANGE, 450000, top_r2),
        ("EXECUTIVE & ENTERPRISE DEFENSE", [
            "Protects corporate leadership from spear-phishing voice clones.",
            "Screens confidential voice messages and board directives.",
            "Hardens corporate communication against brand sabotage."
        ], db.C_SLATE_BLUE, 4300000, top_r2),
        ("DIGITAL IDENTITY PROTECTION", [
            "Provides individuals with forensic tools when their voice is cloned.",
            "Supports legal and law enforcement defamation disputes.",
            "Establishes a verifiable cryptographic standard for voice authenticity."
        ], db.C_GREEN, 8150000, top_r2),
    ]
    
    for title, bullets, acc, lpos, tpos in use_cases:
        db.add_card(s15, lpos, tpos, cw15, ch15, title, bullets,
                    bg_color=db.C_CARD_BG, accent_color=acc, title_color=acc, font_size=9.5, title_size=11)
        
    # Value Proposition Banner
    b15 = s15.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 5400000, 11300000, 750000)
    b15.fill.solid()
    b15.fill.fore_color.rgb = db.C_DARK_NAVY
    b15.line.color.rgb = db.C_CYAN
    tf_b15 = b15.text_frame
    tf_b15.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_b15 = tf_b15.paragraphs[0]
    p_b15.text = "BUSINESS VIABILITY: Extremely lightweight (~97.9k params) design allows ultra-low-cost deployment on existing infrastructure."
    p_b15.alignment = PP_ALIGN.CENTER
    p_b15.font.name = db.FONT_BODY
    p_b15.font.size = Pt(11.5)
    p_b15.font.bold = True
    p_b15.font.color.rgb = db.C_WHITE

    db.set_notes(s15,
        "SPEAKER SCRIPT (30s):\n"
        "SIH-Aegis offers immense business and social impact. In banking and call centers, it stops voice-authorized wire fraud "
        "and account takeovers before money moves. In citizen security, it protects families from panic extortion scams. "
        "From a business viability perspective, because our model is so compact, institutions can deploy it as a microservice "
        "inside their existing IVR and telephony stacks without purchasing multi-million-dollar GPU clusters. It is economically "
        "viable for regional rural banks as well as multinational telecom operators."
    )
    print("Slide 15 generated.")

    # ----------------------------------------------------
    # SLIDE 16: PROTOTYPE STATUS & HONEST ROADMAP (F2, F5, F6)
    # ----------------------------------------------------
    s16 = prs.slides.add_slide(blank_layout)
    db.add_slide_chrome(s16, "FEASIBILITY AND VIABILITY", "Prototype Status & Honest Engineering Reality (Transparent Evaluation)", 16)
    
    top_16 = 1200000
    h_16 = 4000000
    w_16 = 3600000
    
    db.add_card(s16, 450000, top_16, w_16, h_16,
        "COMPLETED & VERIFIED TODAY",
        [
            "Modular GitHub Repository: Structured `src/` tree with automated pytest scaffold.",
            "Audio Ingestion & Preprocessing: 16 kHz mono resampling, float32 normalization.",
            "AASIST Baseline Running: Verified PyTorch checkpoint running deterministically on RTX 4050 GPU.",
            "ONNX Export Validated: 3.8e-06 numerical equivalence confirmed across runtimes.",
            "Interactive Streamlit UI: Full upload, playback, and diagnostic dashboard.",
            "Spectrogram CNN Coded: 4-block architecture (~97.9k params) implemented.",
            "Forward Pass PASS: CUDA forward execution verified on NVIDIA GPU.",
            "Android Base Project: Kotlin + Compose UI + ONNX Runtime initialized with AASIST model loaded."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_GREEN, title_color=db.C_GREEN, font_size=9.8
    )
    
    db.add_card(s16, 4300000, top_16, w_16, h_16,
        "IN ACTIVE PROGRESS",
        [
            "Indian Speech Dataset Curation: Aggregating diverse Indian English and regional language audio samples.",
            "CNN Training Pipeline: Establishing loss curves, cross-entropy criteria, and learning rate schedules.",
            "Equal Error Rate (EER) Evaluation: Calibrating detection thresholds against AASIST baseline.",
            "Android Upload Inference: Connecting mobile ONNX execution directly to the user upload interface.",
            "Controlled Synthetic Generation: Producing paired cloned speech using open-source vocoders."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_ACCENT_ORANGE, title_color=db.C_ACCENT_ORANGE, font_size=10
    )
    
    db.add_card(s16, 8150000, top_16, w_16, h_16,
        "NEXT PLANNED MILESTONES",
        [
            "Validated CNN Checkpoint: Fully trained weights evaluated across unseen Indian speakers.",
            "Android Live Mic Inference: Sliding window chunking for real-time edge speech analysis.",
            "Multi-Model Ensemble: Combining the CNN spectro-temporal detector with the AASIST baseline.",
            "Telephony Gateway Integration: SIP/VoIP trunk connectors for enterprise call-center pilots."
        ],
        bg_color=db.C_CARD_BG, accent_color=db.C_DARK_NAVY, title_color=db.C_DARK_NAVY, font_size=10
    )
    
    # Bottom Banner
    b16 = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 5350000, 11300000, 800000)
    b16.fill.solid()
    b16.fill.fore_color.rgb = db.C_DARK_NAVY
    b16.line.color.rgb = db.C_CYAN
    tf_b16 = b16.text_frame
    tf_b16.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_b16 = tf_b16.paragraphs[0]
    p_b16.text = "SIH-Aegis turns voice authenticity from a trust assumption into a measurable security signal."
    p_b16.alignment = PP_ALIGN.CENTER
    p_b16.font.name = db.FONT_TITLE
    p_b16.font.size = Pt(13)
    p_b16.font.bold = True
    p_b16.font.color.rgb = db.C_CYAN
    p_b16_sub = tf_b16.add_paragraph()
    p_b16_sub.text = "A validated working prototype with an honest, robust engineering path toward lightweight, mobile-capable defense."
    p_b16_sub.alignment = PP_ALIGN.CENTER
    p_b16_sub.font.name = db.FONT_BODY
    p_b16_sub.font.size = Pt(10)
    p_b16_sub.font.color.rgb = db.C_WHITE

    db.set_notes(s16,
        "SPEAKER SCRIPT (30s):\n"
        "To conclude our Round 1 presentation, we present a completely transparent summary of our technical status. "
        "We have eight core milestones completed and verified today—including our working AASIST baseline, our PyTorch-to-ONNX "
        "pipeline, and our GPU-verified CNN architecture. Our CNN training and Indian dataset curation are in active progress, "
        "followed by our Android live microphone deployment. SIH-Aegis is an honest, rigorous engineering project designed to "
        "solve one of the most pressing cybersecurity crises in India and the world. Thank you, and we welcome your questions!\n\n"
        "LIKELY JUDGE QUESTION:\n"
        "Summarize in one sentence: what is actually working today?\n"
        "STRONG ANSWER:\n"
        "Today, we have a working end-to-end prototype that ingests audio, standardizes to 16 kHz mono, executes deterministic "
        "AASIST inference on GPU with verified ONNX equivalence, visualizes spectrograms in Streamlit, and has the lightweight "
        "CNN architecture coded and forward-pass verified."
    )
    print("Slide 16 generated.")

    # Save to files
    out_path1 = "SIH-Aegis-Expanded-Presentation.pptx"
    out_path2 = "/Users/mithumohan/Downloads/SIH-Aegis-Expanded-Final.pptx"
    prs.save(out_path1)
    print(f"Saved to {out_path1}")
    try:
        prs.save(out_path2)
        print(f"Saved to {out_path2}")
    except Exception as e:
        print(f"Note: could not save directly to Downloads root: {e}")

if __name__ == '__main__':
    main()
