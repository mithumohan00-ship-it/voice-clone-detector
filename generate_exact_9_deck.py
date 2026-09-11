import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

import build_exact_9_slides as b9

def build():
    print("Loading base template...")
    prs = Presentation('template.pptx')
    blank_layout = prs.slide_layouts[6]

    # ====================================================
    # SLIDE 1: TITLE PAGE
    # ====================================================
    slide1 = prs.slides[0]
    
    # Update TextBox 9 with provided details
    for shape in slide1.shapes:
        if shape.name == 'TextBox 9':
            tf = shape.text_frame
            tf.word_wrap = True
            # Clear existing text
            for p in tf.paragraphs:
                p.text = ""
            
            p0 = tf.paragraphs[0]
            p0.text = "• Problem Statement ID – SIH26104 (104)"
            p0.font.name = b9.FONT_BODY
            p0.font.size = Pt(13)
            p0.font.bold = True
            p0.font.color.rgb = RGBColor(20, 20, 20)
            p0.space_after = Pt(4)
            
            items = [
                ("Organization", "All India Council for Technical Education (AICTE)"),
                ("Problem Statement Title", "AI-Powered Real-Time Detection and Prevention of Voice Cloning Impersonation Attacks"),
                ("Theme", "Blockchain & Cybersecurity"),
                ("PS Category", "Software"),
                ("Team ID", "0"),
                ("Team Name (Registered on portal)", "SIH-Aegis")
            ]
            for label, val in items:
                p = tf.add_paragraph()
                r1 = p.add_run()
                r1.text = f"• {label} – "
                r1.font.name = b9.FONT_BODY
                r1.font.size = Pt(12)
                r1.font.bold = True
                r1.font.color.rgb = RGBColor(20, 20, 20)
                
                r2 = p.add_run()
                r2.text = val
                r2.font.name = b9.FONT_BODY
                r2.font.size = Pt(12)
                r2.font.bold = False
                r2.font.color.rgb = b9.C_SIH_BLUE if label in ["Theme", "Team Name (Registered on portal)"] else RGBColor(40, 40, 40)
                p.space_after = Pt(4)

    # Bottom-left dark blue project badge
    box1 = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 331286, 5200000, 5600000, 1250000)
    box1.fill.solid()
    box1.fill.fore_color.rgb = b9.C_NAVY_DARK
    box1.line.color.rgb = b9.C_SIH_BLUE
    box1.line.width = Pt(1.5)
    
    tf1 = box1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = int(Pt(14))
    tf1.margin_top = int(Pt(10))
    p_t1 = tf1.paragraphs[0]
    p_t1.text = "SIH-AEGIS"
    p_t1.font.name = b9.FONT_TITLE
    p_t1.font.size = Pt(20)
    p_t1.font.bold = True
    p_t1.font.color.rgb = b9.C_WHITE
    p_t1.space_after = Pt(2)
    
    p_t2 = tf1.add_paragraph()
    p_t2.text = "AI-Powered Real-Time Detection and Prevention of\nVoice Cloning Impersonation Attacks"
    p_t2.font.name = b9.FONT_BODY
    p_t2.font.size = Pt(11)
    p_t2.font.color.rgb = RGBColor(200, 225, 255)
    
    b9.set_notes(slide1,
        "SPEAKER SCRIPT (30s):\n"
        "Respected judges, we represent Team SIH-Aegis working on Problem Statement SIH26104 from AICTE under Blockchain & Cybersecurity. "
        "Our project addresses the urgent crisis of voice cloning impersonation attacks. With modern neural vocoders, attackers can clone "
        "any target's voice from just seconds of audio. Our solution, SIH-Aegis, transforms voice authenticity from an unverified human "
        "assumption into a measurable, real-time security signal. Today, we demonstrate our working prototype, share our empirical "
        "baseline findings, and outline our custom lightweight neural detector."
    )
    print("Slide 1 configured.")

    # Remove template slides 2-7
    for _ in range(len(prs.slides) - 1):
        rId = prs.slides._sldIdLst[1].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[1]

    # ====================================================
    # SLIDE 2: PROPOSED SOLUTION (Exact match to Page 2)
    # ====================================================
    s2 = prs.slides.add_slide(blank_layout)
    b9.add_chrome(s2, "IDEA TITLE", "Proposed Solution (Describe your Idea/Solution/Prototype)", 2)
    
    # Subtitle line
    sub_box = s2.shapes.add_textbox(450000, 1120000, 11300000, 480000)
    tf_sub = sub_box.text_frame
    tf_sub.word_wrap = True
    p_s1 = tf_sub.paragraphs[0]
    p_s1.text = "PROPOSED SOLUTION — SIH-Aegis voice authenticity engine"
    p_s1.font.name = b9.FONT_BODY
    p_s1.font.size = Pt(13)
    p_s1.font.bold = True
    p_s1.font.color.rgb = b9.C_NAVY_DARK
    p_s2 = tf_sub.add_paragraph()
    p_s2.text = "A speech sample is validated, converted to a log-Mel spectrogram and scored by a lightweight CNN, returning a spoof probability and a BONAFIDE / SPOOF verdict."
    p_s2.font.name = b9.FONT_BODY
    p_s2.font.size = Pt(10)
    p_s2.font.color.rgb = b9.C_TEXT_MUTED

    # Pipeline Flowchart Across Top
    steps_s2 = [
        ("Audio Input", "microphone / upload"),
        ("Validation +\nPreprocessing", "16 kHz mono"),
        ("Log-Mel\nSpectrogram", "128 mels · n_fft 512"),
        ("128 × 128\nRepresentation", "standardised float32"),
        ("Lightweight\nCNN", "≈ 97.9K parameters"),
        ("Spoof\nProbability", "BONAFIDE / SPOOF")
    ]
    sw = 1600000
    sh = 950000
    top_step = 1680000
    for i, (st, sd) in enumerate(steps_s2):
        lpos = 450000 + i * (sw + 320000)
        bg_col = b9.C_NAVY_DARK if i == 5 else (b9.C_SIH_BLUE if i == 4 else b9.C_FLOW_BOX_BG)
        txt_col = b9.C_WHITE if i in [4, 5] else b9.C_NAVY_DARK
        border_col = b9.C_SIH_BLUE if i in [4, 5] else b9.C_FLOW_BOX_BORDER
        b9.add_flow_step(s2, lpos, top_step, sw, sh, st, sd, bg_color=bg_col, border_color=border_col, text_color=txt_col)
        if i < len(steps_s2) - 1:
            b9.add_arrow(s2, lpos + sw + 45000, top_step + 380000, 230000, 180000)

    # Baseline Callout Below CNN Box
    callout = s2.shapes.add_textbox(7600000, 2680000, 4150000, 320000)
    tf_c = callout.text_frame
    p_c = tf_c.paragraphs[0]
    p_c.text = "AASIST (pretrained) — retained as secondary baseline / benchmark, not the primary detector"
    p_c.font.name = b9.FONT_BODY
    p_c.font.size = Pt(8.5)
    p_c.font.color.rgb = b9.C_SIH_BLUE
    p_c.alignment = PP_ALIGN.RIGHT

    # 3 Column Cards
    cw = 3600000
    ch = 3200000
    top_c = 3050000
    
    b9.add_column_card(s2, 450000, top_c, cw, ch,
        "Detailed explanation of the solution",
        [
            "Audio is validated, resampled and mixed down to a 16 kHz mono waveform by a shared AudioLoader.",
            "The waveform is converted to a log-Mel spectrogram (128 mels, n_fft 512, hop 160) and resized to a standardised 128 × 128 float32 tensor.",
            "A compact 4-block CNN (≈ 97,890 parameters) performs 2-class classification over that time–frequency image.",
            "The DetectorPipeline returns a spoof probability and a BONAFIDE / SPOOF verdict to a Streamlit interface."
        ],
        font_size=9.5
    )
    
    b9.add_column_card(s2, 4300000, top_c, cw, ch,
        "How it addresses the problem",
        [
            "Replaces the human assumption 'this sounds like them, so it is them' with a computed, reviewable authenticity score.",
            "Operates on any recorded speech sample, so it can sit in front of call-centre, KYC and voice-authentication workflows.",
            "The detector sits behind a common interface, so the audio and UI layers stay unchanged when the model is upgraded.",
            "Small model + GPU-verified inference path make near-real-time analysis a realistic next stage."
        ],
        font_size=9.5
    )
    
    b9.add_column_card(s2, 8150000, top_c, cw, ch,
        "Innovation and uniqueness",
        [
            "Baseline-first engineering: we measured a pretrained state-of-the-art model on real-world audio before trusting it.",
            "That test exposed false-positive spoof predictions on genuine recordings — the finding that motivated a custom, domain-aware detector.",
            "Deliberately lightweight (≈ 97.9K parameters) so it can be retrained cheaply as new cloning systems appear.",
            "Benchmark model kept alongside the primary model, enabling a future multi-model ensemble."
        ],
        font_size=9.5
    )
    
    b9.set_notes(s2,
        "SPEAKER SCRIPT (30s):\n"
        "Here is our proposed solution: the SIH-Aegis Voice Authenticity Engine. When audio is received, our AudioLoader "
        "standardizes it to 16 kHz mono and converts it into a 128x128 Log-Mel spectrogram. Our compact 4-block CNN "
        "evaluates time-frequency patterns to return a spoof probability and a BONAFIDE/SPOOF verdict. This replaces subjective "
        "human trust with an auditable security metric that can sit in front of banking, KYC, and call-center pipelines."
    )
    print("Slide 2 generated.")

    # ====================================================
    # SLIDE 3: TECHNICAL APPROACH (Exact match to Page 3)
    # ====================================================
    s3 = prs.slides.add_slide(blank_layout)
    b9.add_chrome(s3, "TECHNICAL APPROACH", "", 3)

    # Methodology subtitle
    sub_box3 = s3.shapes.add_textbox(450000, 1100000, 11300000, 350000)
    tf_sub3 = sub_box3.text_frame
    p_m = tf_sub3.paragraphs[0]
    p_m.text = "METHODOLOGY — detection model, layer by layer"
    p_m.font.name = b9.FONT_BODY
    p_m.font.size = Pt(13)
    p_m.font.bold = True
    p_m.font.color.rgb = b9.C_NAVY_DARK

    # Layer by layer flow boxes (8 steps + Output)
    layers = [
        ("Input", "1 × 128 × 128"),
        ("Conv2D 1→16", "BN · ReLU · MaxPool"),
        ("Conv2D 16→32", "BN · ReLU · MaxPool"),
        ("Conv2D 32→64", "BN · ReLU · MaxPool"),
        ("Conv2D 64→128", "BN · ReLU"),
        ("Adaptive", "Average Pool"),
        ("Flatten", "Dropout 0.3"),
        ("Linear", "128 → 2"),
        ("OUTPUT", "BONAFIDE / SPOOF")
    ]
    lw = 1080000
    lh = 800000
    top_l = 1500000
    for i, (lt, ld) in enumerate(layers):
        lpos = 450000 + i * (lw + 190000)
        is_out = (i == len(layers) - 1)
        bg = b9.C_NAVY_DARK if is_out else b9.C_FLOW_BOX_BG
        tc = b9.C_WHITE if is_out else b9.C_NAVY_DARK
        bc = b9.C_SIH_BLUE if is_out else b9.C_FLOW_BOX_BORDER
        b9.add_flow_step(s3, lpos, top_l, lw, lh, lt, ld, bg_color=bg, border_color=bc, text_color=tc)
        if i < len(layers) - 1:
            b9.add_arrow(s3, lpos + lw + 20000, top_l + 320000, 150000, 140000)

    # Param note below layer flow
    note_box = s3.shapes.add_textbox(450000, 2350000, 11300000, 320000)
    tf_nb = note_box.text_frame
    p_nb = tf_nb.paragraphs[0]
    p_nb.text = "≈ 97,890 parameters · deterministic evaluation mode · forward pass verified on NVIDIA RTX 4050 Laptop GPU (CUDA): batch × 1 × 128 × 128 → batch × 2 — PASS"
    p_nb.font.name = b9.FONT_BODY
    p_nb.font.size = Pt(9)
    p_nb.font.bold = True
    p_nb.font.color.rgb = b9.C_SIH_BLUE

    # Lower Half: Left = Tech Stack, Right = Architecture
    w_half = 5500000
    top_low = 2750000
    h_low = 3450000
    
    # Left Card: Technology Stack
    card_ts = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, top_low, w_half, h_low)
    card_ts.fill.solid()
    card_ts.fill.fore_color.rgb = b9.C_CARD_BG
    card_ts.line.color.rgb = b9.C_BORDER_GRAY
    card_ts.line.width = Pt(1)
    tf_ts = card_ts.text_frame
    tf_ts.word_wrap = True
    tf_ts.margin_left = int(Pt(12))
    tf_ts.margin_top = int(Pt(10))
    p_ts0 = tf_ts.paragraphs[0]
    p_ts0.text = "TECHNOLOGY STACK"
    p_ts0.font.name = b9.FONT_BODY
    p_ts0.font.size = Pt(12)
    p_ts0.font.bold = True
    p_ts0.font.color.rgb = b9.C_NAVY_DARK
    p_ts0.space_after = Pt(4)
    
    ts_items = [
        ("CORE", "Python · PyTorch · CUDA · NumPy · SciPy"),
        ("AUDIO & FEATURES", "librosa · SoundFile · Log-Mel spectrogram"),
        ("MODELS", "Spectrogram CNN (primary) · AASIST (baseline) · scikit-learn"),
        ("INTERFACE & TOOLING", "Streamlit · Git · GitHub"),
        ("HARDWARE", "NVIDIA RTX 4050 Laptop GPU · 6 GB class"),
        ("AUDIO CONFIGURATION", "16,000 Hz mono · 128 Mel bins · n_fft 512 · hop 160 · 128 × 128 target · per-sample standardisation · float32")
    ]
    for tag, val in ts_items:
        p = tf_ts.add_paragraph()
        r1 = p.add_run()
        r1.text = f"{tag}: "
        r1.font.name = b9.FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = b9.C_SIH_BLUE if tag == "AUDIO CONFIGURATION" else b9.C_TEXT_MUTED
        
        r2 = p.add_run()
        r2.text = val
        r2.font.name = b9.FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.bold = (tag == "AUDIO CONFIGURATION")
        r2.font.color.rgb = b9.C_TEXT_DARK
        p.space_after = Pt(3)

    # Right Card: System Architecture
    card_sa = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 6250000, top_low, w_half, h_low)
    card_sa.fill.solid()
    card_sa.fill.fore_color.rgb = b9.C_CARD_BG
    card_sa.line.color.rgb = b9.C_BORDER_GRAY
    card_sa.line.width = Pt(1)
    tf_sa = card_sa.text_frame
    tf_sa.word_wrap = True
    tf_sa.margin_left = int(Pt(12))
    tf_sa.margin_top = int(Pt(10))
    p_sa0 = tf_sa.paragraphs[0]
    p_sa0.text = "SYSTEM ARCHITECTURE — modular separation of concerns"
    p_sa0.font.name = b9.FONT_BODY
    p_sa0.font.size = Pt(12)
    p_sa0.font.bold = True
    p_sa0.font.color.rgb = b9.C_NAVY_DARK
    p_sa0.space_after = Pt(4)
    
    sa_items = [
        ("UI", "app.py — Streamlit upload / record and result display"),
        ("PIPELINE", "src/pipeline — detector_pipeline.py · result.py (orchestration, decision layer)"),
        ("AUDIO", "src/audio — audio_loader.py · preprocessing.py · recorder.py"),
        ("FEATURES", "src/features — spectrogram.py (log-Mel, 128 × 128)"),
        ("MODELS", "src/models — cnn/spectrogram_cnn.py (primary) · aasist_detector.py · ensemble.py"),
        ("TRAINING & TESTS", "src/training · tests/audio · tests/models · tests/pipeline")
    ]
    for tag, val in sa_items:
        p = tf_sa.add_paragraph()
        r1 = p.add_run()
        r1.text = f"[{tag}] "
        r1.font.name = b9.FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = b9.C_NAVY_DARK
        
        r2 = p.add_run()
        r2.text = val
        r2.font.name = b9.FONT_BODY
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = b9.C_TEXT_DARK
        p.space_after = Pt(3)

    b9.set_notes(s3,
        "SPEAKER SCRIPT (30s):\n"
        "Here is the technical methodology and system architecture. Our primary detection model is a compact 4-block CNN "
        "with exactly 97,890 parameters, verified with deterministic GPU forward execution on an NVIDIA RTX 4050 GPU. "
        "Our codebase follows strict separation of concerns in src/: the audio loader, feature engineering, and model "
        "backends are modular and independently unit-tested."
    )
    print("Slide 3 generated.")

    # ====================================================
    # SLIDE 4: NEW SLIDE 1 — BASELINE EXPERIMENTATION & ONNX
    # ====================================================
    s4 = prs.slides.add_slide(blank_layout)
    b9.add_chrome(s4, "TECHNICAL APPROACH", "Baseline Validation & ONNX Runtime Cross-Platform Verification", 4)
    
    # Subtitle line
    sub4 = s4.shapes.add_textbox(450000, 1120000, 11300000, 480000)
    tf_sub4 = sub4.text_frame
    p_sb4_1 = tf_sub4.paragraphs[0]
    p_sb4_1.text = "EMPIRICAL BASELINE VALIDATION — Discovering Domain Shift & Cross-Runtime Equivalence"
    p_sb4_1.font.name = b9.FONT_BODY
    p_sb4_1.font.size = Pt(13)
    p_sb4_1.font.bold = True
    p_sb4_1.font.color.rgb = b9.C_NAVY_DARK
    p_sb4_2 = tf_sub4.add_paragraph()
    p_sb4_2.text = "Testing a state-of-the-art model exposed real-world generalization limits, driving our custom CNN development while validating ONNX execution."
    p_sb4_2.font.name = b9.FONT_BODY
    p_sb4_2.font.size = Pt(10)
    p_sb4_2.font.color.rgb = b9.C_TEXT_MUTED

    # Flowchart Across Top
    steps_s4 = [
        ("AASIST Baseline", "Pretrained PyTorch model"),
        ("CUDA Evaluation", "Deterministic GPU testing"),
        ("Real-World Tests", "Synthetic vs genuine mic"),
        ("False Positives", "Exposed domain shift"),
        ("ONNX Export", "Cross-platform runtime"),
        ("Strict Equivalence", "Max diff: 3.8 × 10⁻⁶")
    ]
    for i, (st, sd) in enumerate(steps_s4):
        lpos = 450000 + i * (sw + 320000)
        bg = b9.C_ACCENT_ORANGE if i == 3 else (b9.C_ACCENT_GREEN if i == 5 else b9.C_FLOW_BOX_BG)
        tc = b9.C_WHITE if i in [3, 5] else b9.C_NAVY_DARK
        bc = b9.C_ACCENT_ORANGE if i == 3 else (b9.C_ACCENT_GREEN if i == 5 else b9.C_FLOW_BOX_BORDER)
        b9.add_flow_step(s4, lpos, top_step, sw, sh, st, sd, bg_color=bg, border_color=bc, text_color=tc)
        if i < len(steps_s4) - 1:
            b9.add_arrow(s4, lpos + sw + 45000, top_step + 380000, 230000, 180000)

    # 3 Column Cards
    b9.add_column_card(s4, 450000, top_c, cw, ch,
        "Empirical Baseline Testing & Finding",
        [
            "We integrated AASIST (ICASSP 2022) as our initial working prototype and evaluated it against known cloned speech.",
            "Synthetic Detection: Successfully flagged our generated zero-shot clones as SPOOF with decisive probability.",
            "Real-World False Alarms: Testing genuine phone calls and laptop recordings triggered false-positive SPOOF alerts.",
            "Crucial Insight: Benchmark performance on academic datasets (ASVspoof) does not equal real-world deployment robustness."
        ],
        font_size=9.5
    )
    
    b9.add_column_card(s4, 4300000, top_c, cw, ch,
        "ONNX Export & Numerical Verification",
        [
            "Cross-Platform Migration: Exported the PyTorch AASIST model graph to Open Neural Network Exchange (ONNX) format.",
            "Runtime Validation: Successfully executed inference inside the ONNX Runtime engine on CPU and GPU.",
            "Strict Numerical Equivalence: Verified maximum absolute output divergence of ~3.8e-06 between PyTorch and ONNX.",
            "Class Parity: Both execution engines produced identical classifications on validation audio batches."
        ],
        font_size=9.5
    )
    
    b9.add_column_card(s4, 8150000, top_c, cw, ch,
        "Engineering Decision & Model Strategy",
        [
            "Benchmark Retention: AASIST is retained as our measured reference point rather than discarded.",
            "Custom Domain-Aware CNN: Motivated the design of our lightweight ~97.9k parameter Spectrogram CNN.",
            "Fast Retraining Advantage: Compact architecture can be retrained in hours to absorb emerging cloning algorithms.",
            "Multi-Model Potential: Maintains both models behind a unified interface for future multi-model ensembling."
        ],
        font_size=9.5
    )
    
    b9.set_notes(s4,
        "SPEAKER SCRIPT (30s):\n"
        "This slide highlights one of our strongest engineering findings. Rather than blindly trusting an academic baseline, "
        "we tested AASIST on genuine phone and microphone audio and discovered false-positive spoof predictions. "
        "Acoustic channel mismatch and consumer mics cause distribution shift. We also verified cross-platform export: "
        "AASIST was converted to ONNX with an exact maximum numerical divergence of 3.8e-06. This baseline-first discipline "
        "directly motivated our lightweight, domain-aware CNN."
    )
    print("Slide 4 generated.")

    # ====================================================
    # SLIDE 5: NEW SLIDE 2 — ANDROID MOBILE IMPLEMENTATION
    # ====================================================
    s5 = prs.slides.add_slide(blank_layout)
    b9.add_chrome(s5, "TECHNICAL APPROACH", "Android Mobile Architecture — 'From Prototype to Mobile Security'", 5)
    
    sub5 = s5.shapes.add_textbox(450000, 1120000, 11300000, 480000)
    tf_sub5 = sub5.text_frame
    p_sb5_1 = tf_sub5.paragraphs[0]
    p_sb5_1.text = "MOBILE DEPLOYMENT LAYER — Native Android Client & Edge Inference Architecture"
    p_sb5_1.font.size = Pt(13)
    p_sb5_1.font.bold = True
    p_sb5_1.font.color.rgb = b9.C_NAVY_DARK
    p_sb5_2 = tf_sub5.add_paragraph()
    p_sb5_2.text = "Bringing voice authenticity to the endpoint via Kotlin, Jetpack Compose, and embedded ONNX Runtime inference."
    p_sb5_2.font.size = Pt(10)
    p_sb5_2.font.color.rgb = b9.C_TEXT_MUTED

    steps_s5 = [
        ("Android Client", "Kotlin · Jetpack Compose"),
        ("Audio Ingestion", "File upload / mic capture"),
        ("ONNX Mobile", "Embedded runtime engine"),
        ("Model in Assets", "AASIST ONNX loaded"),
        ("On-Device Scoring", "Local inference pipeline"),
        ("Instant Verdict", "BONAFIDE / SPOOF UI")
    ]
    for i, (st, sd) in enumerate(steps_s5):
        lpos = 450000 + i * (sw + 320000)
        bg = b9.C_NAVY_DARK if i == 5 else (b9.C_SIH_BLUE if i == 2 else b9.C_FLOW_BOX_BG)
        tc = b9.C_WHITE if i in [2, 5] else b9.C_NAVY_DARK
        bc = b9.C_SIH_BLUE if i in [2, 5] else b9.C_FLOW_BOX_BORDER
        b9.add_flow_step(s5, lpos, top_step, sw, sh, st, sd, bg_color=bg, border_color=bc, text_color=tc)
        if i < len(steps_s5) - 1:
            b9.add_arrow(s5, lpos + sw + 45000, top_step + 380000, 230000, 180000)

    b9.add_column_card(s5, 450000, top_c, cw, ch,
        "Native Android Progress (Completed)",
        [
            "Native Android Studio project built in Kotlin with Jetpack Compose and Material 3 design.",
            "ONNX Runtime Mobile SDK successfully integrated and initialized inside the Android application.",
            "The AASIST ONNX model is bundled into Android assets and loads into memory without errors.",
            "Reactive state machine implemented: READY ➔ LISTENING ➔ ANALYZING ➔ BONAFIDE / SPOOF."
        ],
        font_size=9.5
    )
    
    b9.add_column_card(s5, 4300000, top_c, cw, ch,
        "Platform Reality & Technical Honesty",
        [
            "CRITICAL HONEST FACT: Third-party Android apps CANNOT intercept live cellular-call audio due to OS security sandboxing.",
            "Target Demonstration Path: Audio file upload & voice note screening (WhatsApp/Telegram audio) evaluated on-device.",
            "Next Mobile Milestone: Microphone capture feeding sliding audio chunks into local ONNX inference for real-time verification.",
            "Enterprise Telecom Vision: Full cellular stream inspection requires carrier-level telephony gateway integration."
        ],
        font_size=9.5
    )
    
    b9.add_column_card(s5, 8150000, top_c, cw, ch,
        "Zero-Trust Biometric Voice Privacy",
        [
            "Zero Cloud Dependency: Local on-device execution ensures sensitive voice data never leaves the user's phone.",
            "Ephemeral Memory Processing: Audio buffers in RAM are purged immediately after calculating the security score.",
            "Non-Invertible Representation: Spectrogram features prevent raw private speech reconstruction.",
            "DPDPA 2023 Alignment: Complies with Indian data governance principles for sensitive biometric voiceprints."
        ],
        font_size=9.5
    )
    
    b9.set_notes(s5,
        "SPEAKER SCRIPT (30s):\n"
        "Our mobile layer demonstrates practical edge security. We built a native Android app using Kotlin and Jetpack Compose, "
        "integrated ONNX Runtime Mobile, and successfully bundled our AASIST model into Android assets. We are completely "
        "transparent: third-party Android apps cannot intercept cellular calls due to Android security sandboxing. Our focus "
        "is on-device upload screening, voice notes, and microphone chunking, running 100% locally with zero cloud exposure."
    )
    print("Slide 5 generated.")

    # ====================================================
    # SLIDE 6: NEW SLIDE 3 — INDIAN CONTEXT & DATA STRATEGY
    # ====================================================
    s6 = prs.slides.add_slide(blank_layout)
    b9.add_chrome(s6, "TECHNICAL APPROACH", "Data Strategy & Ethical Speech Collection for the Indian Context", 6)
    
    sub6 = s6.shapes.add_textbox(450000, 1120000, 11300000, 480000)
    tf_sub6 = sub6.text_frame
    p_sb6_1 = tf_sub6.paragraphs[0]
    p_sb6_1.text = "DATA STRATEGY — Overcoming Linguistic Bias & Ethical Guided Audio Acquisition"
    p_sb6_1.font.size = Pt(13)
    p_sb6_1.font.bold = True
    p_sb6_1.font.color.rgb = b9.C_NAVY_DARK
    p_sb6_2 = tf_sub6.add_paragraph()
    p_sb6_2.text = "Engineering robustness across India's linguistic diversity, consumer phone microphones, and cellular codecs."
    p_sb6_2.font.size = Pt(10)
    p_sb6_2.font.color.rgb = b9.C_TEXT_MUTED

    steps_s6 = [
        ("Indian Speakers", "Multilingual diversity"),
        ("Guided Prompts", "5-10s phonetic prompts"),
        ("Consent Protocol", "Explicit digital opt-in"),
        ("Codec Augment", "AMR-NB · G.711 · Opus"),
        ("Noise Injection", "Reverb · traffic · hum"),
        ("Paired Synthesis", "Cloned via VITS/XTTS")
    ]
    for i, (st, sd) in enumerate(steps_s6):
        lpos = 450000 + i * (sw + 320000)
        bg = b9.C_NAVY_DARK if i == 5 else (b9.C_ACCENT_ORANGE if i == 3 else b9.C_FLOW_BOX_BG)
        tc = b9.C_WHITE if i in [3, 5] else b9.C_NAVY_DARK
        bc = b9.C_NAVY_DARK if i == 5 else (b9.C_ACCENT_ORANGE if i == 3 else b9.C_FLOW_BOX_BORDER)
        b9.add_flow_step(s6, lpos, top_step, sw, sh, st, sd, bg_color=bg, border_color=bc, text_color=tc)
        if i < len(steps_s6) - 1:
            b9.add_arrow(s6, lpos + sw + 45000, top_step + 380000, 230000, 180000)

    b9.add_column_card(s6, 450000, top_c, cw, ch,
        "Designing for the Indian Speech Context",
        [
            "Linguistic Diversity Gap: Academic corpora (ASVspoof) lack Indian English, regional languages, and code-switching.",
            "Target Demographics: Training data strategy prioritizes Indian speakers across accents and phonetic structures.",
            "Hardware Variance: Capturing audio across budget and premium smartphone microphones to prevent hardware bias.",
            "Generalization Focus: System learns fundamental vocal tract properties rather than memorizing clean studio English."
        ],
        font_size=9.5
    )
    
    b9.add_column_card(s6, 4300000, top_c, cw, ch,
        "Environmental & Codec Augmentation",
        [
            "Acoustic Noise Augmentation: Injecting simulated background noise (traffic, indoor fan hum, room reverberation).",
            "Telephony Codec Simulation: Passing audio through lossy cellular codecs (AMR-NB, G.711, Opus) during training.",
            "Channel Invariance: Forces the CNN to focus on vocoder phase inconsistencies rather than high-frequency fidelity.",
            "Real-World Resilience: Prepares model for distribution shift before deployment in active telecom networks."
        ],
        font_size=9.5
    )
    
    b9.add_column_card(s6, 8150000, top_c, cw, ch,
        "Ethical Collection & Paired Synthesis",
        [
            "Consent-First Protocol: Planned Android guided recording mode collects consent-based, anonymized Indian speech.",
            "Balanced Paired Cloning: Genuine recordings are cloned using multiple open-source vocoders (VITS, XTTS, Bark).",
            "Identical Semantic Content: Genuine and spoof pairs share identical words, isolating synthesis artifacts.",
            "Multi-Vocoder Defense: Prevents model from overfitting to a single commercial cloning algorithm."
        ],
        font_size=9.5
    )
    
    b9.set_notes(s6,
        "SPEAKER SCRIPT (30s):\n"
        "Real-world generalization in India requires acoustic and linguistic domain awareness. Standard datasets ignore "
        "Indian accents and regional phonetics. Our data strategy combines consent-based Indian speech collection with "
        "heavy codec and noise augmentation—including simulated telephony compression like AMR and G.711. By pairing genuine "
        "speech with synthetic clones from multiple vocoders, we ensure our model generalizes rather than memorizing."
    )
    print("Slide 6 generated.")

    # ====================================================
    # SLIDE 7: FEASIBILITY AND VIABILITY (Exact match to Page 4)
    # ====================================================
    s7 = prs.slides.add_slide(blank_layout)
    b9.add_chrome(s7, "FEASIBILITY AND VIABILITY", "What already runs, what can go wrong, and how we handle it", 7)
    
    # 3 Large Columns
    cw7 = 3600000
    ch7 = 4900000
    top_7 = 1200000
    
    # Col 1: Feasibility — built and verified
    card_f1 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, top_7, cw7, ch7)
    card_f1.fill.solid()
    card_f1.fill.fore_color.rgb = b9.C_CARD_BG
    card_f1.line.color.rgb = b9.C_BORDER_GRAY
    tf_f1 = card_f1.text_frame
    tf_f1.word_wrap = True
    tf_f1.margin_left = int(Pt(10))
    tf_f1.margin_top = int(Pt(10))
    p_f1 = tf_f1.paragraphs[0]
    p_f1.text = "Feasibility — built and verified"
    p_f1.font.name = b9.FONT_BODY
    p_f1.font.size = Pt(12)
    p_f1.font.bold = True
    p_f1.font.color.rgb = b9.C_NAVY_DARK
    p_f1.space_after = Pt(4)
    
    f1_sections = [
        ("COMPLETED", [
            "Modular GitHub repository with test scaffold",
            "AASIST pretrained baseline loads and runs",
            "Shared AudioLoader + preprocessing (16 kHz mono)",
            "DetectorPipeline and result layer",
            "Streamlit UI running with AASIST integrated",
            "Log-Mel extractor → 128 × 128 representation",
            "Lightweight CNN architecture (≈ 97.9K params)",
            "CUDA verified · CNN GPU forward pass PASS",
            "ONNX export validated (max error: 3.8e-06)",
            "Android Compose UI + ONNX Runtime initialized"
        ], b9.C_ACCENT_GREEN),
        ("IN PROGRESS", [
            "Training dataset selection & Indian speech curation",
            "CNN dataset and training pipeline",
            "CNN evaluation against baseline",
            "Connecting the trained CNN to Streamlit",
            "Android audio upload inference integration"
        ], b9.C_ACCENT_ORANGE),
        ("NOT YET STARTED", [
            "Trained CNN checkpoint · real-time streaming cellular call analysis · large-scale benchmark · final deployment"
        ], b9.C_TEXT_MUTED),
        ("VERIFIED ENVIRONMENT", [
            "PyTorch + CUDA on an NVIDIA RTX 4050 Laptop GPU (6 GB class). AASIST checkpoint loads and runs in deterministic evaluation mode."
        ], b9.C_SIH_BLUE)
    ]
    for header, bullets, col in f1_sections:
        p_h = tf_f1.add_paragraph()
        p_h.text = header
        p_h.font.name = b9.FONT_BODY
        p_h.font.size = Pt(8.5)
        p_h.font.bold = True
        p_h.font.color.rgb = col
        p_h.space_after = Pt(2)
        for b in bullets:
            p_b = tf_f1.add_paragraph()
            p_b.text = "• " + b
            p_b.font.name = b9.FONT_BODY
            p_b.font.size = Pt(8)
            p_b.font.color.rgb = b9.C_TEXT_DARK
            p_b.space_after = Pt(1.5)

    # Col 2: Potential challenges and risks
    card_f2 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 4300000, top_7, cw7, ch7)
    card_f2.fill.solid()
    card_f2.fill.fore_color.rgb = b9.C_CARD_BG
    card_f2.line.color.rgb = b9.C_BORDER_GRAY
    tf_f2 = card_f2.text_frame
    tf_f2.word_wrap = True
    tf_f2.margin_left = int(Pt(10))
    tf_f2.margin_top = int(Pt(10))
    p_f2 = tf_f2.paragraphs[0]
    p_f2.text = "Potential challenges and risks"
    p_f2.font.name = b9.FONT_BODY
    p_f2.font.size = Pt(12)
    p_f2.font.bold = True
    p_f2.font.color.rgb = b9.C_NAVY_DARK
    p_f2.space_after = Pt(4)
    
    risks = [
        ("1", "Domain / dataset mismatch", "A detector trained on one corpus mis-scores audio from different devices, codecs and languages."),
        ("2", "Baseline false positives", "Our testing showed pretrained AASIST correctly flagged our generated clones, but also predicted spoof for several genuine microphone and call recordings."),
        ("3", "Unseen cloning systems", "New TTS and voice-conversion models appear faster than public spoof datasets are refreshed."),
        ("4", "Indian speech coverage", "ASVspoof2019 LA is a useful benchmark but does not represent Indian accents, languages or recording conditions."),
        ("5", "Real-time latency budget", "Streaming inference must stay well inside conversational delay to be usable in a live call."),
        ("6", "Single-GPU development", "All work currently runs on one 6 GB-class laptop GPU, which caps model and batch size.")
    ]
    for num, rtitle, rdesc in risks:
        p = tf_f2.add_paragraph()
        r1 = p.add_run()
        r1.text = f"[{num}] {rtitle}\n"
        r1.font.name = b9.FONT_BODY
        r1.font.size = Pt(9)
        r1.font.bold = True
        r1.font.color.rgb = b9.C_NAVY_DARK
        r2 = p.add_run()
        r2.text = rdesc
        r2.font.name = b9.FONT_BODY
        r2.font.size = Pt(8.2)
        r2.font.color.rgb = b9.C_TEXT_DARK
        p.space_after = Pt(3)

    # Col 3: Strategies for overcoming them
    card_f3 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 8150000, top_7, cw7, ch7)
    card_f3.fill.solid()
    card_f3.fill.fore_color.rgb = b9.C_CARD_BG
    card_f3.line.color.rgb = b9.C_BORDER_GRAY
    tf_f3 = card_f3.text_frame
    tf_f3.word_wrap = True
    tf_f3.margin_left = int(Pt(10))
    tf_f3.margin_top = int(Pt(10))
    p_f3 = tf_f3.paragraphs[0]
    p_f3.text = "Strategies for overcoming them"
    p_f3.font.name = b9.FONT_BODY
    p_f3.font.size = Pt(12)
    p_f3.font.bold = True
    p_f3.font.color.rgb = b9.C_NAVY_DARK
    p_f3.space_after = Pt(4)
    
    strats = [
        ("1", "Train on representative audio", "Future training set to target Indian speakers, Indian English and regional languages alongside real microphone and call recordings."),
        ("2", "Benchmark, don't defer", "AASIST stays as a measured reference point; the custom CNN is evaluated against it rather than assumed better."),
        ("3", "Augment for the real world", "Codec, noise, device and environment variation added during training so the model sees distribution shift before deployment does."),
        ("4", "Keep the model small", "A ≈ 97.9K-parameter network is cheap to retrain, so new attack families can be absorbed quickly."),
        ("5", "Shared front end, swappable model", "One log-Mel pipeline feeds every detector, so a new model drops in behind the same interface."),
        ("6", "Test each stage independently", "tests/audio, tests/models and tests/pipeline isolate failures instead of debugging the whole system at once.")
    ]
    for num, stitle, sdesc in strats:
        p = tf_f3.add_paragraph()
        r1 = p.add_run()
        r1.text = f"[{num}] {stitle}\n"
        r1.font.name = b9.FONT_BODY
        r1.font.size = Pt(9)
        r1.font.bold = True
        r1.font.color.rgb = b9.C_ACCENT_GREEN
        r2 = p.add_run()
        r2.text = sdesc
        r2.font.name = b9.FONT_BODY
        r2.font.size = Pt(8.2)
        r2.font.color.rgb = b9.C_TEXT_DARK
        p.space_after = Pt(3)

    b9.set_notes(s7,
        "SPEAKER SCRIPT (30s):\n"
        "Here is our feasibility and viability matrix. We are transparent: our AASIST baseline, ONNX export, and CNN forward pass "
        "are completed and verified on our RTX 4050 GPU, while our CNN training is in progress. We identified six specific risks—such as "
        "dataset mismatch, unseen cloning models, and latency budgets—and paired each with a concrete engineering mitigation, "
        "including acoustic augmentation, lightweight parameter efficiency, and swappable modular pipelines."
    )
    print("Slide 7 generated.")

    # ====================================================
    # SLIDE 8: IMPACT AND BENEFITS (Exact match to Page 5)
    # ====================================================
    s8 = prs.slides.add_slide(blank_layout)
    b9.add_chrome(s8, "IMPACT AND BENEFITS", "", 8)

    # Top Roadmap Bar (Current Position: Stage 1)
    sub8 = s8.shapes.add_textbox(450000, 1080000, 11300000, 300000)
    tf_sub8 = sub8.text_frame
    p_r0 = tf_sub8.paragraphs[0]
    p_r0.text = "ROADMAP — from prototype to deployable control"
    p_r0.font.name = b9.FONT_BODY
    p_r0.font.size = Pt(12)
    p_r0.font.bold = True
    p_r0.font.color.rgb = b9.C_NAVY_DARK

    stages = [
        ("STAGE 1", "Prototype (Verified)"),
        ("STAGE 2", "CNN training & evaluation"),
        ("STAGE 3", "Real-time microphone inference"),
        ("STAGE 4", "Robustness: codecs, noise"),
        ("STAGE 5", "Multi-model ensemble"),
        ("STAGE 6", "Deployment")
    ]
    st_w = 1650000
    st_h = 650000
    top_st = 1420000
    for i, (st, sd) in enumerate(stages):
        lpos = 450000 + i * (st_w + 280000)
        bg = b9.C_NAVY_DARK if i == 0 else (b9.C_FLOW_BOX_BG if i == 1 else b9.C_CARD_BG)
        tc = b9.C_WHITE if i == 0 else (b9.C_SIH_BLUE if i == 1 else b9.C_TEXT_MUTED)
        bc = b9.C_NAVY_DARK if i == 0 else b9.C_FLOW_BOX_BORDER
        b9.add_flow_step(s8, lpos, top_st, st_w, st_h, st, sd, bg_color=bg, border_color=bc, text_color=tc)
        if i < len(stages) - 1:
            b9.add_arrow(s8, lpos + st_w + 35000, top_st + 220000, 210000, 150000)

    # Roadmap disclaimer
    cur_pos = s8.shapes.add_textbox(450000, 2120000, 11300000, 250000)
    tf_cp = cur_pos.text_frame
    p_cp = tf_cp.paragraphs[0]
    p_cp.text = "Current position: Stage 1. No production readiness, deployment or accuracy claim is made — CNN training and evaluation are still in progress."
    p_cp.font.name = b9.FONT_BODY
    p_cp.font.size = Pt(8.5)
    p_cp.font.color.rgb = b9.C_TEXT_MUTED

    # Middle Section: Left = Target Audience (6 items), Right = Benefits (4 pills)
    w_aud = 5500000
    top_aud = 2450000
    h_aud = 3100000
    
    # Left Box: Potential Impact on Target Audience
    card_aud = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, top_aud, w_aud, h_aud)
    card_aud.fill.solid()
    card_aud.fill.fore_color.rgb = b9.C_CARD_BG
    card_aud.line.color.rgb = b9.C_BORDER_GRAY
    tf_aud = card_aud.text_frame
    tf_aud.word_wrap = True
    tf_aud.margin_left = int(Pt(12))
    tf_aud.margin_top = int(Pt(10))
    p_a0 = tf_aud.paragraphs[0]
    p_a0.text = "POTENTIAL IMPACT ON THE TARGET AUDIENCE"
    p_a0.font.name = b9.FONT_BODY
    p_a0.font.size = Pt(11)
    p_a0.font.bold = True
    p_a0.font.color.rgb = b9.C_NAVY_DARK
    p_a0.space_after = Pt(4)
    
    aud_items = [
        ("1", "Banking & fraud prevention", "Screens voice-authorised instructions before money moves."),
        ("2", "Voice authentication / KYC", "Adds a liveness-and-authenticity check to voice-based identity flows."),
        ("3", "Call-centre security", "Flags synthetic callers attempting agent-assisted account takeover."),
        ("4", "Social-engineering defence", "Catches the 'familiar voice in distress' pretext used against staff and families."),
        ("5", "Digital identity protection", "Gives individuals a way to contest a cloned recording of themselves."),
        ("6", "Financial scam detection", "Supports investigation and triage of reported voice-based scams.")
    ]
    for num, atitle, adesc in aud_items:
        p = tf_aud.add_paragraph()
        r1 = p.add_run()
        r1.text = f"({num}) {atitle}: "
        r1.font.name = b9.FONT_BODY
        r1.font.size = Pt(9)
        r1.font.bold = True
        r1.font.color.rgb = b9.C_SIH_BLUE
        r2 = p.add_run()
        r2.text = adesc
        r2.font.name = b9.FONT_BODY
        r2.font.size = Pt(8.8)
        r2.font.color.rgb = b9.C_TEXT_DARK
        p.space_after = Pt(2.5)

    # Right Box: Benefits of the Solution
    card_ben = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 6250000, top_aud, w_aud, h_aud)
    card_ben.fill.solid()
    card_ben.fill.fore_color.rgb = b9.C_CARD_BG
    card_ben.line.color.rgb = b9.C_BORDER_GRAY
    tf_ben = card_ben.text_frame
    tf_ben.word_wrap = True
    tf_ben.margin_left = int(Pt(12))
    tf_ben.margin_top = int(Pt(10))
    p_b0 = tf_ben.paragraphs[0]
    p_b0.text = "BENEFITS OF THE SOLUTION"
    p_b0.font.name = b9.FONT_BODY
    p_b0.font.size = Pt(11)
    p_b0.font.bold = True
    p_b0.font.color.rgb = b9.C_NAVY_DARK
    p_b0.space_after = Pt(4)
    
    bens = [
        ("SOCIAL", "Protects citizens — especially elderly and non-technical users — from impersonation attacks that exploit a familiar voice.", b9.C_SIH_BLUE),
        ("ECONOMIC", "Reduces losses from voice-authorised fraud and lowers the manual verification burden on institutions.", b9.C_NAVY_DARK),
        ("INSTITUTIONAL", "Replaces a subjective human judgement with an auditable authenticity signal that can be logged and reviewed.", b9.C_ACCENT_GREEN),
        ("TECHNICAL", "A compact model on commodity GPU hardware keeps the cost of adoption low for banks, telecoms and smaller enterprises.", b9.C_ACCENT_ORANGE)
    ]
    for tag, desc, col in bens:
        p = tf_ben.add_paragraph()
        r1 = p.add_run()
        r1.text = f"[{tag}] "
        r1.font.name = b9.FONT_BODY
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = col
        r2 = p.add_run()
        r2.text = desc
        r2.font.name = b9.FONT_BODY
        r2.font.size = Pt(8.8)
        r2.font.color.rgb = b9.C_TEXT_DARK
        p.space_after = Pt(4)

    # Bottom Banner
    bb8 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 5650000, 11300000, 560000)
    bb8.fill.solid()
    bb8.fill.fore_color.rgb = b9.C_NAVY_DARK
    bb8.line.color.rgb = b9.C_SIH_BLUE
    tf_bb8 = bb8.text_frame
    tf_bb8.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_bb8 = tf_bb8.paragraphs[0]
    p_bb8.text = "SIH-Aegis turns voice authenticity from a trust assumption into a measurable security signal."
    p_bb8.alignment = PP_ALIGN.CENTER
    p_bb8.font.name = b9.FONT_BODY
    p_bb8.font.size = Pt(11.5)
    p_bb8.font.bold = True
    p_bb8.font.color.rgb = b9.C_WHITE

    b9.set_notes(s8,
        "SPEAKER SCRIPT (30s):\n"
        "SIH-Aegis delivers multi-dimensional impact. In banking and call centers, it screens voice-authorized transfers "
        "before funds move, reducing massive financial fraud. For citizens, it provides defense against distress scams. "
        "Economically, our lightweight ~97.9k parameter architecture enables deployment on commodity edge servers without "
        "expensive cloud subscriptions, turning voice authenticity into a measurable, auditable signal."
    )
    print("Slide 8 generated.")

    # ====================================================
    # SLIDE 9: RESEARCH AND REFERENCES (Exact match to Page 6)
    # ====================================================
    s9 = prs.slides.add_slide(blank_layout)
    b9.add_chrome(s9, "RESEARCH  AND REFERENCES", "", 9)

    sub9 = s9.shapes.add_textbox(450000, 1080000, 11300000, 450000)
    tf_sub9 = sub9.text_frame
    p_s9_1 = tf_sub9.paragraphs[0]
    p_s9_1.text = "Sources behind the architecture and the design decisions"
    p_s9_1.font.name = b9.FONT_BODY
    p_s9_1.font.size = Pt(13)
    p_s9_1.font.bold = True
    p_s9_1.font.color.rgb = b9.C_NAVY_DARK
    p_s9_2 = tf_sub9.add_paragraph()
    p_s9_2.text = "Verify every citation and paste the final URL into each entry before submitting."
    p_s9_2.font.name = b9.FONT_BODY
    p_s9_2.font.size = Pt(9.5)
    p_s9_2.font.color.rgb = b9.C_TEXT_MUTED

    # 6 Reference Cards (2 rows of 3)
    cw9 = 3600000
    ch9 = 1750000
    top_r1_9 = 1580000
    top_r2_9 = 3450000
    
    refs = [
        ("ANTI-SPOOFING MODEL", "AASIST: Audio Anti-Spoofing using Integrated Spectro-Temporal Graph Attention Networks", "Jung et al., ICASSP 2022 — the pretrained baseline we integrated and then benchmarked against real-world recordings.", 450000, top_r1_9),
        ("BENCHMARK DATASET", "ASVspoof 2019 Logical Access database and evaluation plan", "Reference corpus for anti-spoofing research; used as a benchmark, noted as not representative of all Indian speech conditions.", 4300000, top_r1_9),
        ("CHALLENGE SERIES", "ASVspoof challenge series — spoofing attack taxonomy", "Defines the attack categories and evaluation methodology our detector is designed against.", 8150000, top_r1_9),
        ("FRAMEWORK & CUDA", "PyTorch — deep learning framework and CUDA execution", "Model definition, GPU acceleration and deterministic evaluation mode for the Spectrogram CNN and AASIST baseline.", 450000, top_r2_9),
        ("AUDIO PROCESSING", "librosa · SoundFile · SciPy · NumPy", "Audio I/O, resampling to 16 kHz mono and log-Mel spectrogram extraction in the shared feature pipeline.", 4300000, top_r2_9),
        ("INTERFACES & EDGE", "Streamlit (Web) & Jetpack Compose + ONNX (Android)", "Prototype interfaces for upload, on-device execution, and display of the BONAFIDE / SPOOF result.", 8150000, top_r2_9),
    ]
    for tag, title, desc, lpos, tpos in refs:
        card = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, lpos, tpos, cw9, ch9)
        card.fill.solid()
        card.fill.fore_color.rgb = b9.C_CARD_BG
        card.line.color.rgb = b9.C_BORDER_GRAY
        tf_rc = card.text_frame
        tf_rc.word_wrap = True
        tf_rc.margin_left = int(Pt(10))
        tf_rc.margin_top = int(Pt(8))
        
        p_t = tf_rc.paragraphs[0]
        r_tag = p_t.add_run()
        r_tag.text = f"[{tag}] "
        r_tag.font.name = b9.FONT_BODY
        r_tag.font.size = Pt(8)
        r_tag.font.bold = True
        r_tag.font.color.rgb = b9.C_SIH_BLUE
        
        r_ti = p_t.add_run()
        r_ti.text = title
        r_ti.font.name = b9.FONT_BODY
        r_ti.font.size = Pt(8.5)
        r_ti.font.bold = True
        r_ti.font.color.rgb = b9.C_NAVY_DARK
        p_t.space_after = Pt(2)
        
        p_d = tf_rc.add_paragraph()
        p_d.text = desc
        p_d.font.name = b9.FONT_BODY
        p_d.font.size = Pt(8)
        p_d.font.color.rgb = b9.C_TEXT_DARK
        p_d.space_after = Pt(2)
        
        p_l = tf_rc.add_paragraph()
        p_l.text = "Link: ______________________________________________"
        p_l.font.name = b9.FONT_BODY
        p_l.font.size = Pt(7.5)
        p_l.font.color.rgb = b9.C_TEXT_MUTED

    # Bottom Repository Box
    rep_box = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, 5350000, 11300000, 850000)
    rep_box.fill.solid()
    rep_box.fill.fore_color.rgb = b9.C_NAVY_DARK
    rep_box.line.color.rgb = b9.C_SIH_BLUE
    tf_rep = rep_box.text_frame
    tf_rep.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_rep.margin_left = int(Pt(14))
    p_rep1 = tf_rep.paragraphs[0]
    p_rep1.text = "PROJECT REPOSITORY"
    p_rep1.font.name = b9.FONT_BODY
    p_rep1.font.size = Pt(10.5)
    p_rep1.font.bold = True
    p_rep1.font.color.rgb = b9.C_CYAN if hasattr(b9, 'C_CYAN') else b9.C_SIH_BLUE
    p_rep1.space_after = Pt(2)
    p_rep2 = tf_rep.add_paragraph()
    p_rep2.text = "github.com/<team>/voice-clone-detector — modular src/ layout (audio · features · models · pipeline · training) with a matching tests/ tree. Stage 1 prototype; CNN training in progress."
    p_rep2.font.name = b9.FONT_BODY
    p_rep2.font.size = Pt(9.5)
    p_rep2.font.color.rgb = b9.C_WHITE

    b9.set_notes(s9,
        "SPEAKER SCRIPT (30s):\n"
        "Our architecture stands on peer-reviewed scientific foundations. We benchmark against AASIST from ICASSP 2022 and "
        "the ASVspoof challenge series. Our implementation is built on PyTorch, librosa, and ONNX Runtime, and hosted in a "
        "clean GitHub repository with matching test suites. Thank you for your time, and Team SIH-Aegis is ready for your questions!"
    )
    print("Slide 9 generated.")

    # Save to file
    out_file = "SIH-Aegis-Submission-9Slides.pptx"
    prs.save(out_file)
    print(f"Presentation saved successfully as {out_file}")

if __name__ == '__main__':
    build()
