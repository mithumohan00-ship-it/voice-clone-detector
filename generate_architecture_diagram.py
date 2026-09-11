import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches

# Set high DPI for publication/presentation quality
fig = plt.figure(figsize=(16, 10.5), dpi=300)
ax = fig.add_subplot(111)
ax.set_xlim(0, 16)
ax.set_ylim(0, 10.5)
ax.axis('off')

# Color Palette (SIH Cybersecurity Theme)
BG_COLOR = "#0A1128"       # Deep Cyber Navy
BOX_BG = "#131E3A"         # Slate Box Fill
BOX_BORDER = "#2B3D5E"     # Border
CYAN_ACCENT = "#00E5FF"    # Tech Cyan
BLUE_ACCENT = "#0070C0"    # SIH Blue
GREEN_ACCENT = "#10B981"   # Bonafide Emerald
ORANGE_ACCENT = "#F59E0B"  # Amber Orange
TEXT_WHITE = "#FFFFFF"
TEXT_MUTED = "#94A3B8"

fig.patch.set_facecolor(BG_COLOR)

# Title Banner
ax.text(8.0, 10.05, "SIH-AEGIS: END-TO-END SYSTEM ARCHITECTURE", 
        fontsize=20, fontweight='bold', ha='center', va='center', color=TEXT_WHITE, fontfamily='sans-serif')
ax.text(8.0, 9.7, "AI-Powered Real-Time Detection and Prevention of Voice Cloning Impersonation Attacks", 
        fontsize=11.5, ha='center', va='center', color=CYAN_ACCENT, fontfamily='sans-serif')

def draw_layer_card(x, y, w, h, title, border_col, fill_col=BOX_BG, title_col=CYAN_ACCENT):
    rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.08,rounding_size=0.15",
                                  facecolor=fill_col, edgecolor=border_col, linewidth=1.5, zorder=2)
    ax.add_patch(rect)
    # Title placed clearly inside the top padding
    ax.text(x + 0.35, y + h - 0.28, title, fontsize=11, fontweight='bold', color=title_col, va='center', zorder=3)

def draw_sub_box(x, y, w, h, title, subtitle, fill_col="#17254A", border_col=BOX_BORDER, title_col=TEXT_WHITE):
    rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.04,rounding_size=0.1",
                                  facecolor=fill_col, edgecolor=border_col, linewidth=1.0, zorder=3)
    ax.add_patch(rect)
    ax.text(x + w/2, y + h/2 + 0.12, title, fontsize=9.2, fontweight='bold', color=title_col, ha='center', va='center', zorder=4)
    if subtitle:
        ax.text(x + w/2, y + h/2 - 0.14, subtitle, fontsize=7.8, color=TEXT_MUTED, ha='center', va='center', zorder=4)

def draw_arrow(x1, y1, x2, y2, col=CYAN_ACCENT, width=1.6):
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle="->", color=col, lw=width, shrinkA=3, shrinkB=3), zorder=5)

# --- LAYER 1: CLIENT & INGESTION (y = 7.6 to 9.3) ---
draw_layer_card(0.6, 7.55, 14.8, 1.8, "1. INGESTION & CLIENT PRESENTATION LAYER", CYAN_ACCENT, title_col=CYAN_ACCENT)
draw_sub_box(0.9, 7.75, 3.2, 1.15, "Streamlit Dashboard", "Web UI · Mic Capture · Uploads", fill_col="#15264D", border_col=CYAN_ACCENT)
draw_sub_box(4.5, 7.75, 3.2, 1.15, "Native Android Client", "Kotlin · Jetpack Compose · M3", fill_col="#15264D", border_col=CYAN_ACCENT)
draw_sub_box(8.1, 7.75, 3.4, 1.15, "Voice Notes & Audio Files", "WhatsApp · Telegram · WAV/MP3", fill_col="#15264D", border_col=BLUE_ACCENT)
draw_sub_box(11.9, 7.75, 3.2, 1.15, "Telecom / IVR Trunk (Target)", "SIP Gateway · Support Queue", fill_col="#15264D", border_col=BOX_BORDER, title_col=TEXT_MUTED)

# Downward Arrows from L1 to L2
draw_arrow(2.5, 7.55, 2.5, 7.05)
draw_arrow(6.1, 7.55, 6.1, 7.05)
draw_arrow(9.8, 7.55, 9.8, 7.05)

# --- LAYER 2: AUDIO PROCESSING & FEATURE EXTRACTION (y = 5.3 to 7.05) ---
draw_layer_card(0.6, 5.35, 14.8, 1.7, "2. AUDIO CONDITIONING & SPECTROGRAM PIPELINE (src/audio & src/features)", BLUE_ACCENT, title_col=BLUE_ACCENT)
draw_sub_box(0.9, 5.55, 3.2, 1.05, "AudioLoader & Resampler", "16,000 Hz Mono · Float32 PCM", border_col=BLUE_ACCENT)
draw_sub_box(4.5, 5.55, 3.2, 1.05, "STFT Engine (scipy/librosa)", "n_fft=512 (~32ms) · hop=160 (10ms)", border_col=BLUE_ACCENT)
draw_sub_box(8.1, 5.55, 3.4, 1.05, "Log-Mel Filterbank", "128 Mel Bins · 0–8000 Hz Nyquist", border_col=BLUE_ACCENT)
draw_sub_box(11.9, 5.55, 3.2, 1.05, "Tensor Standardizer", "128 × 128 Shape · Per-Sample Z-Norm", border_col=CYAN_ACCENT)

# Horizontal Pipeline Arrows
draw_arrow(4.1, 6.07, 4.5, 6.07)
draw_arrow(7.7, 6.07, 8.1, 6.07)
draw_arrow(11.5, 6.07, 11.9, 6.07)

# Downward Arrow from L2 to L3
draw_arrow(8.0, 5.35, 8.0, 4.85)

# --- LAYER 3: DUAL-MODEL INFERENCE ENGINE (y = 2.7 to 4.85) ---
draw_layer_card(0.6, 2.70, 14.8, 2.15, "3. DUAL-MODEL INFERENCE ENGINE (src/models)", ORANGE_ACCENT, title_col=ORANGE_ACCENT)

# Primary Model Card
rect_p = patches.FancyBboxPatch((0.9, 2.90), 7.6, 1.55, boxstyle="round,pad=0.04,rounding_size=0.1",
                                facecolor="#0E2338", edgecolor=CYAN_ACCENT, linewidth=1.3, zorder=3)
ax.add_patch(rect_p)
ax.text(1.1, 4.18, "PRIMARY: Lightweight Spectrogram CNN (~97.9K Parameters)", fontsize=9.8, fontweight='bold', color=CYAN_ACCENT, zorder=4)
ax.text(1.1, 3.85, "Topology: 4 × Conv2D (16, 32, 64, 128) + BatchNorm + ReLU + MaxPool2D", fontsize=8.4, color=TEXT_WHITE, zorder=4)
ax.text(1.1, 3.55, "Classifier Head: AdaptiveAvgPool2d(1x1) ➔ Flatten ➔ Dropout(0.3) ➔ Linear(128➔2)", fontsize=8.4, color=TEXT_WHITE, zorder=4)
ax.text(1.1, 3.22, "Status: Architecture Coded · Forward Pass PASS on RTX 4050 GPU (Training in progress)", fontsize=8.2, color=ORANGE_ACCENT, zorder=4)

# Baseline Model Card
rect_b = patches.FancyBboxPatch((8.8, 2.90), 6.3, 1.55, boxstyle="round,pad=0.04,rounding_size=0.1",
                                facecolor="#0E2338", edgecolor=GREEN_ACCENT, linewidth=1.3, zorder=3)
ax.add_patch(rect_b)
ax.text(9.0, 4.18, "BASELINE / BENCHMARK: AASIST (~297K Parameters)", fontsize=9.8, fontweight='bold', color=GREEN_ACCENT, zorder=4)
ax.text(9.0, 3.85, "Framework: Pretrained Graph Attention Network (ICASSP 2022)", fontsize=8.4, color=TEXT_WHITE, zorder=4)
ax.text(9.0, 3.55, "Cross-Platform: Exported ONNX Graph · Validated in ONNX Runtime", fontsize=8.4, color=TEXT_WHITE, zorder=4)
ax.text(9.0, 3.22, "Status: Verified Working Prototype (Max Error 3.8e-06 · Bundled in Android)", fontsize=8.2, color=GREEN_ACCENT, zorder=4)

# Downward Arrow from L3 to L4
draw_arrow(8.0, 2.70, 8.0, 2.20)

# --- LAYER 4: ORCHESTRATION, VERDICT & SECURITY (y = 0.45 to 2.20) ---
draw_layer_card(0.6, 0.45, 14.8, 1.75, "4. ORCHESTRATION, RISK SCORING & ZERO-TRUST PRIVACY (src/pipeline)", GREEN_ACCENT, title_col=GREEN_ACCENT)

draw_sub_box(0.9, 0.65, 3.2, 1.15, "DetectorPipeline", "Thresholding · Logit Calibration", border_col=GREEN_ACCENT)
draw_sub_box(4.5, 0.65, 3.2, 1.15, "Security Verdict", "BONAFIDE vs SPOOF · Confidence %", fill_col="#132E27", border_col=GREEN_ACCENT, title_col=GREEN_ACCENT)
draw_sub_box(8.1, 0.65, 3.4, 1.15, "Zero-Trust Privacy Guard", "Ephemeral RAM · No Raw Voice Storage", border_col=ORANGE_ACCENT)
draw_sub_box(11.9, 0.65, 3.2, 1.15, "Audit Logging (DPDPA)", "Session Hash · Timestamp · Score", border_col=BOX_BORDER)

# Horizontal Flow Arrows in L4
draw_arrow(4.1, 1.22, 4.5, 1.22)
draw_arrow(7.7, 1.22, 8.1, 1.22)
draw_arrow(11.5, 1.22, 11.9, 1.22)

# Save high-resolution PNG
out_img = "sih_aegis_architecture.png"
plt.savefig(out_img, bbox_inches='tight', facecolor=fig.get_facecolor(), edgecolor='none', dpi=300)
plt.close()
print("Updated architecture diagram saved.")
