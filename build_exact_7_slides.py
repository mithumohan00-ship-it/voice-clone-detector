import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

import build_exact_9_slides as b9

def build_7_slide_presentation():
    print("Loading base template.pptx...")
    prs = Presentation('template.pptx')
    blank_layout = prs.slide_layouts[6]

    # ====================================================
    # SLIDE 1: TITLE PAGE (Exact Match to Page 1)
    # ====================================================
    slide1 = prs.slides[0]
    
    # Update TextBox 9 with provided details
    for shape in slide1.shapes:
        if shape.name == 'TextBox 9':
            tf = shape.text_frame
            tf.word_wrap = True
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

    # Shared dimensions
    sw = 1600000
    sh = 950000
    top_step = 1680000
    cw = 3600000
    ch = 3200000
    top_c = 3050000

    # ====================================================
    # SLIDE 2: PROPOSED SOLUTION (Exact match to Page 2)
    # ====================================================
    s2 = prs.slides.add_slide(blank_layout)
    b9.add_chrome(s2, "IDEA TITLE", "Proposed Solution (Describe your Idea/Solution/Prototype)", 2)
    
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

    steps_s2 = [
        ("Audio Input", "microphone / upload"),
        ("Validation +\nPreprocessing", "16 kHz mono"),
        ("Log-Mel\nSpectrogram", "128 mels · n_fft 512"),
        ("128 × 128\nRepresentation", "standardised float32"),
        ("Lightweight\nCNN", "≈ 97.9K parameters"),
        ("Spoof\nProbability", "BONAFIDE / SPOOF")
    ]
    for i, (st, sd) in enumerate(steps_s2):
        lpos = 450000 + i * (sw + 320000)
        bg_col = b9.C_NAVY_DARK if i == 5 else (b9.C_SIH_BLUE if i == 4 else b9.C_FLOW_BOX_BG)
        txt_col = b9.C_WHITE if i in [4, 5] else b9.C_NAVY_DARK
        border_col = b9.C_SIH_BLUE if i in [4, 5] else b9.C_FLOW_BOX_BORDER
        b9.add_flow_step(s2, lpos, top_step, sw, sh, st, sd, bg_color=bg_col, border_color=border_col, text_color=txt_col)
        if i < len(steps_s2) - 1:
            b9.add_arrow(s2, lpos + sw + 45000, top_step + 380000, 230000, 180000)

    callout = s2.shapes.add_textbox(7600000, 2680000, 4150000, 320000)
    tf_c = callout.text_frame
    p_c = tf_c.paragraphs[0]
    p_c.text = "AASIST (pretrained) — retained as secondary baseline / benchmark, not the primary detector"
    p_c.font.name = b9.FONT_BODY
    p_c.font.size = Pt(8.5)
    p_c.font.color.rgb = b9.C_SIH_BLUE
    p_c.alignment = PP_ALIGN.RIGHT

    b9.add_column_card(s2, 450000, top_c, cw, ch, 
        "Detailed explanation of the solution",
        [
            "Audio is validated, resampled and mixed down to a 16 kHz mono waveform by a shared AudioLoader.",
            "The waveform is converted to a log-Mel spectrogram (128 mels, n_fft 512, hop 160) and resized to a standardised 128 × 128 float32 tensor.",
            "A compact 4-block CNN (≈ 97,890 parameters) performs 2-class classification over that time–frequency image.",
            "The DetectorPipeline returns a spoof probability and a BONAFIDE / SPOOF verdict to a Streamlit interface."
        ]
    )
    b9.add_column_card(s2, 4300000, top_c, cw, ch,
        "How it addresses the problem",
        [
            "Replaces the human assumption 'this sounds like them, so it is them' with a computed, reviewable authenticity score.",
            "Operates on any recorded speech sample, so it can sit in front of call-centre, KYC and voice-authentication workflows.",
            "The detector sits behind a common interface, so the audio and UI layers stay unchanged when the model is upgraded.",
            "Small model + GPU-verified inference path make near-real-time analysis a realistic next stage."
        ]
    )
    b9.add_column_card(s2, 8150000, top_c, cw, ch,
        "Innovation and uniqueness",
        [
            "Baseline-first engineering: we measured a pretrained state-of-the-art model on real-world audio before trusting it.",
            "That test exposed false-positive spoof predictions on genuine recordings — the finding that motivated a custom, domain-aware detector.",
            "Deliberately lightweight (≈ 97.9K parameters) so it can be retrained cheaply as new cloning systems appear.",
            "Benchmark model kept alongside the primary model, enabling a future multi-model ensemble."
        ]
    )
    b9.set_notes(s2,
        "SPEAKER SCRIPT (30s):\n"
        "Here is the SIH-Aegis solution architecture. When an audio sample arrives, our shared AudioLoader validates and resamples "
        "it to 16 kHz mono. It computes a log-Mel spectrogram with 128 Mel bins, standardizing it to a 128 by 128 float32 tensor. "
        "Our primary detector is a lightweight 4-block CNN with roughly 97,900 parameters that classifies speech as Bonafide or Spoof. "
        "Crucially, we retain the pretrained AASIST model as a measured baseline benchmark. This replaces subjective human belief with an "
        "objective, mathematical authenticity score."
    )
    print("Slide 2 generated.")

    # ====================================================
    # SLIDE 3: TECHNICAL APPROACH — METHODOLOGY (Exact match to Page 3)
    # ====================================================
    s3 = prs.slides.add_slide(blank_layout)
    b9.add_chrome(s3, "TECHNICAL APPROACH", "Methodology — detection model, layer by layer", 3)
    
    sub_box3 = s3.shapes.add_textbox(450000, 1120000, 11300000, 360000)
    tf_sub3 = sub_box3.text_frame
    p_s3 = tf_sub3.paragraphs[0]
    p_s3.text = "METHODOLOGY — detection model, layer by layer"
    p_s3.font.name = b9.FONT_BODY
    p_s3.font.size = Pt(13)
    p_s3.font.bold = True
    p_s3.font.color.rgb = b9.C_NAVY_DARK

    layers = [
        ("Input", "1 × 128 × 128"),
        ("Conv2D\n1→16", "BN · ReLU ·\nMaxPool"),
        ("Conv2D\n16→32", "BN · ReLU ·\nMaxPool"),
        ("Conv2D\n32→64", "BN · ReLU ·\nMaxPool"),
        ("Conv2D\n64→128", "BN · ReLU"),
        ("Adaptive\nAverage Pool", ""),
        ("Flatten", "Dropout 0.3"),
        ("Linear\n128 → 2", ""),
        ("OUTPUT", "BONAFIDE /\nSPOOF")
    ]
    lw = 1100000
    lh = 950000
    top_l = 1580000
    for i, (lt, ld) in enumerate(layers):
        lpos = 450000 + i * (lw + 195000)
        bg = b9.C_NAVY_DARK if i == 8 else (b9.C_SIH_BLUE if i == 4 else b9.C_FLOW_BOX_BG)
        tc = b9.C_WHITE if i in [4, 8] else b9.C_NAVY_DARK
        bc = b9.C_SIH_BLUE if i in [4, 8] else b9.C_FLOW_BOX_BORDER
        b9.add_flow_step(s3, lpos, top_l, lw, lh, lt, ld, bg_color=bg, border_color=bc, text_color=tc)
        if i < len(layers) - 1:
            b9.add_arrow(s3, lpos + lw + 20000, top_l + 380000, 155000, 160000)

    # Verification banner below flow
    vbox = s3.shapes.add_textbox(450000, 2580000, 11300000, 320000)
    tf_v = vbox.text_frame
    p_v = tf_v.paragraphs[0]
    p_v.text = "≈ 97,890 parameters · deterministic evaluation mode · forward pass verified on NVIDIA RTX 4050 Laptop GPU (CUDA): batch × 1 × 128 × 128 → batch × 2 — PASS"
    p_v.font.name = b9.FONT_BODY
    p_v.font.size = Pt(8.5)
    p_v.font.bold = True
    p_v.font.color.rgb = b9.C_SIH_BLUE
    p_v.alignment = PP_ALIGN.CENTER

    # 2 Big Cards for Tech Stack & System Architecture
    w_large = 5500000
    h_large = 3400000
    top_large = 2950000

    card_tech = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 450000, top_large, w_large, h_large)
    card_tech.fill.solid()
    card_tech.fill.fore_color.rgb = b9.C_CARD_BG
    card_tech.line.color.rgb = b9.C_BORDER_GRAY
    tf_t = card_tech.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = int(Pt(14))
    tf_t.margin_top = int(Pt(12))
    p_th = tf_t.paragraphs[0]
    p_th.text = "TECHNOLOGY STACK"
    p_th.font.name = b9.FONT_BODY
    p_th.font.size = Pt(12)
    p_th.font.bold = True
    p_th.font.color.rgb = b9.C_NAVY_DARK
    p_th.space_after = Pt(4)

    tech_items = [
        ("CORE", "Python · PyTorch · CUDA · NumPy · SciPy"),
        ("AUDIO & FEATURES", "librosa · SoundFile · Log-Mel spectrogram"),
        ("MODELS", "Spectrogram CNN (primary) · AASIST (baseline) · scikit-learn"),
        ("INTERFACE & TOOLING", "Streamlit · Git · GitHub"),
        ("HARDWARE", "NVIDIA RTX 4050 Laptop GPU · 6 GB class"),
        ("AUDIO CONFIGURATION", "16,000 Hz mono · 128 Mel bins · n_fft 512 · hop 160 · 128 × 128 target · per-sample standardisation · float32")
    ]
    for label, val in tech_items:
        p = tf_t.add_paragraph()
        r1 = p.add_run()
        r1.text = f"{label}: "
        r1.font.name = b9.FONT_BODY
        r1.font.size = Pt(8.8)
        r1.font.bold = True
        r1.font.color.rgb = b9.C_SIH_BLUE if label == "AUDIO CONFIGURATION" else b9.C_NAVY_DARK
        
        r2 = p.add_run()
        r2.text = val
        r2.font.name = b9.FONT_BODY
        r2.font.size = Pt(8.8)
        r2.font.bold = (label == "AUDIO CONFIGURATION")
        r2.font.color.rgb = b9.C_NAVY_DARK if label == "AUDIO CONFIGURATION" else b9.C_TEXT_MUTED
        p.space_after = Pt(2)

    card_arch = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 6250000, top_large, w_large, h_large)
    card_arch.fill.solid()
    card_arch.fill.fore_color.rgb = b9.C_CARD_BG
    card_arch.line.color.rgb = b9.C_BORDER_GRAY
    tf_a = card_arch.text_frame
    tf_a.word_wrap = True
    tf_a.margin_left = int(Pt(14))
    tf_a.margin_top = int(Pt(12))
    p_ah = tf_a.paragraphs[0]
    p_ah.text = "SYSTEM ARCHITECTURE — modular separation of concerns"
    p_ah.font.name = b9.FONT_BODY
    p_ah.font.size = Pt(12)
    p_ah.font.bold = True
    p_ah.font.color.rgb = b9.C_NAVY_DARK
    p_ah.space_after = Pt(4)

    arch_items = [
        ("[UI]", "app.py — Streamlit upload / record and result display"),
        ("[PIPELINE]", "src/pipeline — detector_pipeline.py · result.py (orchestration, decision layer)"),
        ("[AUDIO]", "src/audio — audio_loader.py · preprocessing.py · recorder.py"),
        ("[FEATURES]", "src/features — spectrogram.py (log-Mel, 128 × 128)"),
        ("[MODELS]", "src/models — cnn/spectrogram_cnn.py (primary) · aasist_detector.py · ensemble.py"),
        ("[TRAINING & TESTS]", "src/training · tests/audio · tests/models · tests/pipeline")
    ]
    for label, val in arch_items:
        p = tf_a.add_paragraph()
        r1 = p.add_run()
        r1.text = f"{label} "
        r1.font.name = b9.FONT_BODY
        r1.font.size = Pt(8.8)
        r1.font.bold = True
        r1.font.color.rgb = b9.C_NAVY_DARK
        
        r2 = p.add_run()
        r2.text = val
        r2.font.name = b9.FONT_BODY
        r2.font.size = Pt(8.8)
        r2.font.color.rgb = b9.C_TEXT_MUTED
        p.space_after = Pt(2)

    b9.set_notes(s3,
        "SPEAKER SCRIPT (30s):\n"
        "This slide breaks down our model and system architecture. The Spectrogram CNN uses 4 Conv2D blocks scaling from 1 to 128 channels, "
        "with Batch Normalization, ReLU, and MaxPool, followed by Adaptive Average Pooling and Dropout. It totals 97,890 parameters. "
        "We have verified the full forward pass on our NVIDIA RTX 4050 GPU with PyTorch and CUDA. The codebase has clean modular separation "
        "of concerns across UI, pipeline orchestration, audio I/O, feature extraction, and models with matching unit tests."
    )
    print("Slide 3 generated.")

    # ====================================================
    # SLIDE 4: TECHNICAL APPROACH — BASELINE VALIDATION & ONNX (Exact match to Page 4)
    # ====================================================
    s4 = prs.slides.add_slide(blank_layout)
    b9.add_chrome(s4, "TECHNICAL APPROACH", "Baseline Validation & ONNX Runtime Cross-Platform Verification", 4)
    
    sub4 = s4.shapes.add_textbox(450000, 1120000, 11300000, 480000)
    tf_sub4 = sub4.text_frame
    p_sb4_1 = tf_sub4.paragraphs[0]
    p_sb4_1.text = "EMPIRICAL BASELINE VALIDATION — Discovering Domain Shift & Cross-Runtime Equivalence"
    p_sb4_1.font.size = Pt(13)
    p_sb4_1.font.bold = True
    p_sb4_1.font.color.rgb = b9.C_NAVY_DARK
    p_sb4_2 = tf_sub4.add_paragraph()
    p_sb4_2.text = "Testing a state-of-the-art model exposed real-world generalization limits, driving our custom CNN development while validating ONNX execution."
    p_sb4_2.font.size = Pt(10)
    p_sb4_2.font.color.rgb = b9.C_TEXT_MUTED

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
        bg = b9.C_NAVY_DARK if i == 5 else (b9.C_ACCENT_ORANGE if i == 3 else b9.C_FLOW_BOX_BG)
        tc = b9.C_WHITE if i in [3, 5] else b9.C_NAVY_DARK
        bc = b9.C_NAVY_DARK if i == 5 else (b9.C_ACCENT_ORANGE if i == 3 else b9.C_FLOW_BOX_BORDER)
        b9.add_flow_step(s4, lpos, top_step, sw, sh, st, sd, bg_color=bg, border_color=bc, text_color=tc)
        if i < len(steps_s4) - 1:
            b9.add_arrow(s4, lpos + sw + 45000, top_step + 380000, 230000, 180000)

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
        "Slide 4 captures a key engineering achievement and insight. We exported our AASIST baseline to ONNX and validated numerical parity: "
        "the maximum divergence between PyTorch and ONNX Runtime is approximately 3.8e-06 with identical class predictions. "
        "More importantly, testing real-world microphone and phone recordings revealed false-positive spoof alerts on genuine voices. "
        "This domain shift proved that academic benchmark accuracy does not translate directly to deployment robustness, directly motivating "
        "our custom lightweight domain-aware detector."
    )
    print("Slide 4 generated.")

    # ====================================================
    # SLIDE 5: TECHNICAL APPROACH — ANDROID MOBILE (Exact match to Page 5)
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
    # SLIDE 6: TECHNICAL APPROACH — DATA STRATEGY (Exact match to Page 6)
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
    # SLIDE 7: FEASIBILITY, VIABILITY & IMPACT (Comprehensive Synthesis of Pages 7 & 8)
    # ====================================================
    s7 = prs.slides.add_slide(blank_layout)
    b9.add_chrome(s7, "FEASIBILITY AND VIABILITY", "What already runs, what can go wrong, and how it delivers measurable impact", 7)
    
    cw7 = 3600000
    ch7 = 4300000
    top_7 = 1150000
    
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
    p_f1.font.size = Pt(11.5)
    p_f1.font.bold = True
    p_f1.font.color.rgb = b9.C_NAVY_DARK
    p_f1.space_after = Pt(3)
    
    f1_sections = [
        ("COMPLETED", [
            "Modular GitHub repository with test scaffold",
            "AASIST pretrained baseline loads and runs",
            "Shared AudioLoader + preprocessing (16 kHz mono)",
            "DetectorPipeline and result layer in Streamlit UI",
            "Log-Mel extractor → 128 × 128 representation",
            "Lightweight CNN architecture (≈ 97.9K params)",
            "CUDA verified · CNN GPU forward pass PASS",
            "ONNX export validated (max error: 3.8e-06)",
            "Android Compose UI + ONNX Runtime initialized"
        ], b9.C_ACCENT_GREEN),
        ("IN PROGRESS", [
            "Training dataset selection & Indian speech curation",
            "CNN training pipeline & baseline evaluation",
            "Android audio upload inference integration"
        ], b9.C_ACCENT_ORANGE),
        ("NOT YET STARTED", [
            "Trained CNN checkpoint · real-time cellular call analysis"
        ], b9.C_TEXT_MUTED),
        ("VERIFIED ENVIRONMENT", [
            "PyTorch + CUDA on an NVIDIA RTX 4050 GPU (6 GB). Deterministic evaluation verified."
        ], b9.C_SIH_BLUE)
    ]
    for header, bullets, col in f1_sections:
        p_h = tf_f1.add_paragraph()
        p_h.text = header
        p_h.font.name = b9.FONT_BODY
        p_h.font.size = Pt(8.2)
        p_h.font.bold = True
        p_h.font.color.rgb = col
        p_h.space_after = Pt(1)
        for b in bullets:
            p_b = tf_f1.add_paragraph()
            p_b.text = "• " + b
            p_b.font.name = b9.FONT_BODY
            p_b.font.size = Pt(7.6)
            p_b.font.color.rgb = b9.C_TEXT_DARK
            p_b.space_after = Pt(1)

    # Col 2: Potential challenges & Mitigation Strategies
    card_f2 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 4300000, top_7, cw7, ch7)
    card_f2.fill.solid()
    card_f2.fill.fore_color.rgb = b9.C_CARD_BG
    card_f2.line.color.rgb = b9.C_BORDER_GRAY
    tf_f2 = card_f2.text_frame
    tf_f2.word_wrap = True
    tf_f2.margin_left = int(Pt(10))
    tf_f2.margin_top = int(Pt(10))
    p_f2 = tf_f2.paragraphs[0]
    p_f2.text = "Challenges & Mitigation Strategies"
    p_f2.font.name = b9.FONT_BODY
    p_f2.font.size = Pt(11.5)
    p_f2.font.bold = True
    p_f2.font.color.rgb = b9.C_NAVY_DARK
    p_f2.space_after = Pt(3)
    
    paired_risks = [
        ("[1] Domain / dataset mismatch", "Risk: Cross-device / language mis-scoring.\n→ Strategy: Train on representative Indian English & regional audio."),
        ("[2] Baseline false positives", "Risk: AASIST false alarms on real mics.\n→ Strategy: Custom domain-aware CNN trained with acoustic noise invariance."),
        ("[3] Unseen cloning systems", "Risk: Attackers deploy new commercial TTS.\n→ Strategy: Codec, noise, and multi-vocoder variation during training."),
        ("[4] Indian speech coverage", "Risk: ASVspoof lacks Indian accents.\n→ Strategy: Consent-based opt-in mobile collection protocol."),
        ("[5] Real-time latency budget", "Risk: Audio must be processed within live call delay.\n→ Strategy: Ultra-compact ≈97.9K parameter network (<50ms)."),
        ("[6] Modular separation", "Risk: Complex monolithic pipeline failures.\n→ Strategy: Isolated unit tests (tests/audio, models, pipeline).")
    ]
    for rtitle, rdesc in paired_risks:
        p = tf_f2.add_paragraph()
        r1 = p.add_run()
        r1.text = f"{rtitle}\n"
        r1.font.name = b9.FONT_BODY
        r1.font.size = Pt(8.2)
        r1.font.bold = True
        r1.font.color.rgb = b9.C_SIH_BLUE
        r2 = p.add_run()
        r2.text = rdesc
        r2.font.name = b9.FONT_BODY
        r2.font.size = Pt(7.6)
        r2.font.color.rgb = b9.C_TEXT_DARK
        p.space_after = Pt(2)

    # Col 3: Impact, Benefits & Strategic Roadmap
    card_f3 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 8150000, top_7, cw7, ch7)
    card_f3.fill.solid()
    card_f3.fill.fore_color.rgb = b9.C_CARD_BG
    card_f3.line.color.rgb = b9.C_BORDER_GRAY
    tf_f3 = card_f3.text_frame
    tf_f3.word_wrap = True
    tf_f3.margin_left = int(Pt(10))
    tf_f3.margin_top = int(Pt(10))
    p_f3 = tf_f3.paragraphs[0]
    p_f3.text = "Impact, Benefits & Roadmap"
    p_f3.font.name = b9.FONT_BODY
    p_f3.font.size = Pt(11.5)
    p_f3.font.bold = True
    p_f3.font.color.rgb = b9.C_NAVY_DARK
    p_f3.space_after = Pt(3)

    impact_items = [
        ("TARGET AUDIENCE IMPACT", [
            "Banking & Fraud: Screens voice-authorized fund moves.",
            "Voice Auth / KYC: Liveness & authenticity check for remote identity.",
            "Call Centers: Shields agents against vishing & account takeover.",
            "Citizen Defense: Shields families from emergency bail/kidnap scams."
        ], b9.C_NAVY_DARK),
        ("SOLUTION BENEFITS", [
            "[Social] Protects citizens from impersonation extortion.",
            "[Economic] Reduces fraud losses & manual review overhead.",
            "[Institutional] Auditable authenticity signal for logs.",
            "[Technical] Commodity GPU & edge mobile inference."
        ], b9.C_ACCENT_GREEN),
        ("ROADMAP FROM PROTOTYPE TO CONTROL", [
            "Stage 1: Prototype (Verified) ➔ 2: CNN Training ➔ 3: Real-Time Mic ➔ 4: Telephony Codecs ➔ 5: Ensemble ➔ 6: Deployment."
        ], b9.C_SIH_BLUE)
    ]
    for header, bullets, col in impact_items:
        p_h = tf_f3.add_paragraph()
        p_h.text = header
        p_h.font.name = b9.FONT_BODY
        p_h.font.size = Pt(8.2)
        p_h.font.bold = True
        p_h.font.color.rgb = col
        p_h.space_after = Pt(1)
        for b in bullets:
            p_b = tf_f3.add_paragraph()
            p_b.text = "• " + b if not b.startswith("Stage") else b
            p_b.font.name = b9.FONT_BODY
            p_b.font.size = Pt(7.6)
            p_b.font.color.rgb = b9.C_TEXT_DARK
            p_b.space_after = Pt(1.2)

    # Bottom Blue Punchline Box on Slide 7 (matching Page 8 style)
    punch_box = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, 450000, 5600000, 11300000, 560000)
    punch_box.fill.solid()
    punch_box.fill.fore_color.rgb = b9.C_NAVY_DARK
    punch_box.line.color.rgb = b9.C_SIH_BLUE
    punch_box.line.width = Pt(1.5)
    tf_pb = punch_box.text_frame
    tf_pb.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_pb = tf_pb.paragraphs[0]
    p_pb.text = "SIH-Aegis turns voice authenticity from a trust assumption into a measurable security signal."
    p_pb.alignment = PP_ALIGN.CENTER
    p_pb.font.name = b9.FONT_BODY
    p_pb.font.size = Pt(13)
    p_pb.font.bold = True
    p_pb.font.color.rgb = b9.C_WHITE

    b9.set_notes(s7,
        "SPEAKER SCRIPT (30s):\n"
        "Here is our feasibility, viability, and impact summary. We are technically transparent: our AASIST baseline, ONNX export, "
        "and 97.9k CNN forward pass are completed and verified on our RTX 4050 GPU, while our CNN training is in progress. "
        "We identified six specific risks—such as dataset mismatch, unseen cloning models, and latency budgets—and paired each with a "
        "concrete engineering mitigation, including acoustic augmentation, lightweight parameter efficiency, and swappable modular pipelines. "
        "Our impact spans banking authorizations, KYC protection, and family extortion scam defense across a phased 6-stage roadmap. "
        "SIH-Aegis turns voice authenticity from a trust assumption into a measurable security signal. Thank you!"
    )
    print("Slide 7 generated.")

    # Save to final file
    out_file = "/Users/mithumohan/Downloads/SIH/SIH-Aegis-Submission-7Slides.pptx"
    prs.save(out_file)
    print(f"Presentation successfully saved as {out_file} (Total Slides: {len(prs.slides)})")

if __name__ == '__main__':
    build_7_slide_presentation()
