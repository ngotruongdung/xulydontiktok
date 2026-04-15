import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
from datetime import datetime

# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Order Studio", page_icon="📦", layout="wide")

# ─── GLOBAL STYLES — Soft Pastel Gradient ─────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

/* ════════════════════════════════════════════════════════════
   BASE RESET & FOUNDATION
════════════════════════════════════════════════════════════ */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
html { -webkit-font-smoothing: antialiased; text-rendering: optimizeLegibility; }

.stApp {
    font-family: 'Inter', system-ui, -apple-system, sans-serif;
    background: #FAFBFE;
    min-height: 100vh;
    position: relative;
    overflow-x: hidden;
}
/* Floating gradient orbs */
.stApp::before {
    content: '';
    position: fixed;
    top: -120px; right: -80px;
    width: 380px; height: 380px;
    background: radial-gradient(circle, rgba(232,210,255,0.45) 0%, rgba(232,210,255,0) 70%);
    border-radius: 50%;
    pointer-events: none;
    z-index: 0;
    animation: orbFloat1 18s ease-in-out infinite;
}
.stApp::after {
    content: '';
    position: fixed;
    bottom: -60px; left: -100px;
    width: 340px; height: 340px;
    background: radial-gradient(circle, rgba(187,247,208,0.35) 0%, rgba(187,247,208,0) 70%);
    border-radius: 50%;
    pointer-events: none;
    z-index: 0;
    animation: orbFloat2 22s ease-in-out infinite;
}
@keyframes orbFloat1 {
    0%, 100% { transform: translate(0, 0) scale(1); }
    33% { transform: translate(-40px, 30px) scale(1.05); }
    66% { transform: translate(20px, -20px) scale(0.95); }
}
@keyframes orbFloat2 {
    0%, 100% { transform: translate(0, 0) scale(1); }
    50% { transform: translate(50px, -40px) scale(1.08); }
}

#MainMenu, footer, header { visibility: hidden; display: none; height: 0; overflow: hidden; }
/* Ẩn branding "Created by" trên Streamlit Community Cloud */
[data-testid="stAppDeployButton"] { display: none !important; }
.viewerBadge_container__r5tak,
.viewerBadge_link__qRIco,
#root > div:nth-child(1) > div.withScreencast > div > div > div > section > div.block-container > div:nth-child(1) > div > div:nth-child(1) > div > div:nth-child(3) > div { display: none !important; }
a[href*="streamlit.io/cloud"], a[href*="share.streamlit.io"] { display: none !important; }
.block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 5rem !important;
    max-width: 100% !important;
    padding-left: 3rem !important;
    padding-right: 3rem !important;
    position: relative;
    z-index: 1;
}
/* Giới hạn nội dung chính để không bị loãng trên màn hình rộng */
.stMainBlockContainer > div {
    max-width: 1200px;
    margin-left: auto;
    margin-right: auto;
}

/* ════════════════════════════════════════════════════════════
   TOPBAR — Minimal greeting
════════════════════════════════════════════════════════════ */
.os-topbar {
    display: flex; align-items: center; justify-content: space-between;
    padding: 20px 0 24px;
    margin-bottom: 4px;
}
.os-topbar-brand {
    display: flex; align-items: center; gap: 16px;
}
.os-topbar-logo {
    width: 44px; height: 44px;
    background: linear-gradient(145deg, #e0e7ff, #c7d2fe);
    border-radius: 14px;
    display: flex; align-items: center; justify-content: center;
    font-size: 20px;
    box-shadow: 0 2px 12px rgba(129,140,248,0.18);
    flex-shrink: 0;
}
.os-topbar-name {
    font-size: 22px; font-weight: 300;
    color: #6366f1; letter-spacing: -0.3px;
}
.os-topbar-name strong {
    font-weight: 700; color: #4338ca;
}
.os-topbar-sub {
    font-size: 12.5px; color: #a5a8b8; font-weight: 400;
    margin-top: 2px; letter-spacing: 0.1px;
}
.os-topbar-status {
    display: flex; align-items: center; gap: 7px;
    font-size: 12px; font-weight: 500; color: #9ca3af;
    background: rgba(255,255,255,0.80);
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    border: 1px solid rgba(229,231,235,0.60);
    border-radius: 100px;
    padding: 7px 16px;
}
.os-status-dot {
    width: 7px; height: 7px; border-radius: 50%;
    background: #34d399;
    display: inline-block;
    animation: sdot 2.5s ease-in-out infinite;
}
@keyframes sdot {
    0%, 100% { opacity: 1; box-shadow: 0 0 0 0 rgba(52,211,153,0.4); }
    50% { opacity: 0.6; box-shadow: 0 0 0 4px rgba(52,211,153,0); }
}

/* ════════════════════════════════════════════════════════════
   SECTION LABEL — Ultra minimal
════════════════════════════════════════════════════════════ */
.os-section {
    display: flex; align-items: center; gap: 12px;
    margin: 32px 0 16px;
    padding-bottom: 0;
    border-bottom: none;
}
.os-section-num {
    display: inline-flex; align-items: center; justify-content: center;
    width: 28px; height: 28px;
    background: linear-gradient(135deg, #e0e7ff, #c7d2fe);
    border-radius: 9px;
    font-size: 11px; font-weight: 700;
    font-family: 'JetBrains Mono', monospace;
    color: #6366f1;
    flex-shrink: 0;
    transition: transform 0.25s ease, box-shadow 0.25s ease;
}
.os-section:hover .os-section-num {
    transform: scale(1.1) rotate(-3deg);
    box-shadow: 0 4px 14px rgba(99,102,241,0.20);
}
.os-section-title {
    font-size: 15px; font-weight: 600; color: #64748b;
    letter-spacing: -0.1px;
}

/* ════════════════════════════════════════════════════════════
   GLASS CARD — Shared card mixin
════════════════════════════════════════════════════════════ */
.os-card {
    background: rgba(255,255,255,0.72);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    border: 1px solid rgba(229,231,235,0.50);
    border-radius: 20px;
    padding: 24px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02), 0 8px 32px rgba(99,102,241,0.04);
    transition: transform 0.25s ease, box-shadow 0.25s ease;
}
.os-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 2px 6px rgba(0,0,0,0.03), 0 12px 40px rgba(99,102,241,0.08);
}

/* ════════════════════════════════════════════════════════════
   SETTINGS SELECTS
════════════════════════════════════════════════════════════ */
.stSelectbox > div > div {
    border: 1px solid rgba(229,231,235,0.70) !important;
    border-radius: 14px !important;
    background: rgba(255,255,255,0.80) !important;
    backdrop-filter: blur(12px) !important;
    font-size: 14px !important;
    font-family: 'Inter', sans-serif !important;
    transition: all 0.2s ease !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.02) !important;
}
.stSelectbox > div > div:hover {
    border-color: #c7d2fe !important;
    box-shadow: 0 2px 12px rgba(99,102,241,0.08) !important;
}
.stSelectbox > div > div:focus-within {
    border-color: #a5b4fc !important;
    box-shadow: 0 0 0 3px rgba(99,102,241,0.08) !important;
}
.stSelectbox label {
    font-size: 12.5px !important;
    font-weight: 500 !important;
    color: #94a3b8 !important;
    text-transform: none !important;
    letter-spacing: 0.2px !important;
    margin-bottom: 8px !important;
}

/* ════════════════════════════════════════════════════════════
   FILE UPLOADER — Glass style (Complete rewrite)
════════════════════════════════════════════════════════════ */

/* ── Container chính ── */
[data-testid="stFileUploader"] {
    background: rgba(255,255,255,0.65);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    border: 1.5px dashed rgba(196,206,224,0.60) !important;
    border-radius: 20px !important;
    padding: 0;
    transition: all 0.25s ease;
    margin-top: 6px !important;
    overflow: hidden;
}
[data-testid="stFileUploader"]:hover {
    border-color: #c7d2fe !important;
    background: rgba(255,255,255,0.80);
    box-shadow: 0 4px 20px rgba(99,102,241,0.06);
}

/* ── Label (tiêu đề uploader) ── */
[data-testid="stFileUploader"] label {
    font-size: 12.5px !important;
    font-weight: 600 !important;
    color: #64748b !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
    text-align: center !important;
    display: block !important;
    padding: 14px 16px 0 !important;
}

/* ── Ẩn icon help (ⓘ) ── */
[data-testid="stFileUploader"] [data-testid="stTooltipHoverTarget"] {
    display: none !important;
}

/* ── Ẩn TOÀN BỘ text gốc tiếng Anh trong instructions ── */
[data-testid="stFileUploaderDropzoneInstructions"] {
    font-size: 0 !important;
    line-height: 0 !important;
    height: 0 !important;
    min-height: 0 !important;
    overflow: hidden !important;
    margin: 0 !important;
    padding: 0 !important;
    visibility: hidden !important;
    position: absolute !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] * {
    display: none !important;
}

/* ════════════════════════════════════════════════════════════
   DROPZONE — TRẠNG THÁI CHƯA UPLOAD (có instructions)
════════════════════════════════════════════════════════════ */
[data-testid="stFileUploaderDropzone"] {
    border: none !important;
    background: transparent !important;
    width: 100% !important;
    min-width: 0 !important;
}

[data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDropzoneInstructions"]) {
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    gap: 0 !important;
    padding: 24px 20px 28px !important;
    text-align: center !important;
}

/* Text hướng dẫn tiếng Việt */
[data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDropzoneInstructions"])::before {
    content: 'Kéo thả file vào đây';
    display: block !important;
    width: 100% !important;
    font-size: 13px;
    color: #b0b8c9;
    font-weight: 400;
    font-family: 'Inter', sans-serif;
    text-align: center !important;
    line-height: 1;
    margin-bottom: 16px;
    order: -2;
}

/* Ẩn tất cả inner wrapper elements mặc định */
[data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDropzoneInstructions"]) > div {
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    width: 100% !important;
}

/* ── Nút "Chọn file" (trạng thái chưa upload) ── */
[data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDropzoneInstructions"]) button {
    /* Ẩn text gốc hoàn toàn */
    font-size: 0 !important;
    line-height: 0 !important;
    color: transparent !important;
    /* Style nút */
    min-width: 120px !important;
    padding: 10px 24px !important;
    border-radius: 12px !important;
    border: 1px solid rgba(199,210,254,0.50) !important;
    background: linear-gradient(135deg, #f8faff, #f0f4ff) !important;
    transition: all 0.25s ease !important;
    cursor: pointer !important;
    position: relative !important;
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
}
[data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDropzoneInstructions"]) button:hover {
    border-color: #c7d2fe !important;
    background: linear-gradient(135deg, #eef2ff, #e0e7ff) !important;
    box-shadow: 0 4px 16px rgba(99,102,241,0.10) !important;
    transform: translateY(-1px) !important;
}
[data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDropzoneInstructions"]) button::after {
    content: 'Chọn file' !important;
    font-size: 13px !important;
    font-weight: 600 !important;
    font-family: 'Inter', sans-serif !important;
    color: #818cf8 !important;
    line-height: normal !important;
}
/* Ẩn mọi child bên trong button (icon, span text gốc) */
[data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDropzoneInstructions"]) button > * {
    display: none !important;
}

/* ════════════════════════════════════════════════════════════
   DROPZONE — TRẠNG THÁI ĐÃ UPLOAD (compact)
════════════════════════════════════════════════════════════ */
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) {
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    gap: 0 !important;
    padding: 14px 16px !important;
}

/* ── File item row (đã upload) ── */
[data-testid="stFileUploaderFile"] {
    padding: 8px 12px !important;
    overflow: hidden !important;
    max-width: 100% !important;
    min-width: 0 !important;
    display: flex !important;
    align-items: center !important;
    gap: 8px !important;
    flex-wrap: nowrap !important;
    background: rgba(241,245,249,0.60) !important;
    border-radius: 12px !important;
    border: 1px solid rgba(229,231,235,0.40) !important;
    margin-bottom: 8px !important;
    width: 100% !important;
}

/* Icon file */
[data-testid="stFileUploaderFile"] [data-testid="stFileUploaderFileIcon"] {
    flex-shrink: 0 !important;
    opacity: 0.6 !important;
}

/* Tên file truncate */
[data-testid="stFileUploaderFile"] > div:first-child {
    overflow: hidden !important;
    text-overflow: ellipsis !important;
    white-space: nowrap !important;
    min-width: 0 !important;
    flex: 1 !important;
}
[data-testid="stFileUploaderFile"] small {
    overflow: hidden !important;
    text-overflow: ellipsis !important;
    white-space: nowrap !important;
    max-width: 160px !important;
    display: inline-block !important;
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 11.5px !important;
    color: #64748b !important;
}

/* Nút xóa file (X) */
[data-testid="stFileUploaderFile"] button {
    font-size: inherit !important;
    min-width: unset !important;
    padding: 4px !important;
    border-radius: 50% !important;
    border: none !important;
    background: transparent !important;
    flex-shrink: 0 !important;
    opacity: 0.5 !important;
    transition: opacity 0.2s ease !important;
}
[data-testid="stFileUploaderFile"] button:hover {
    opacity: 1 !important;
    background: rgba(239,68,68,0.08) !important;
}
[data-testid="stFileUploaderFile"] button::after {
    content: none !important;
}

/* ── Nút "Đổi file" (compact, sau khi upload) ── */
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) > button,
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) > span > button {
    /* Ẩn text gốc hoàn toàn */
    font-size: 0 !important;
    line-height: 0 !important;
    color: transparent !important;
    /* Style nút */
    padding: 7px 16px !important;
    border-radius: 10px !important;
    border: 1px solid rgba(199,210,254,0.50) !important;
    background: linear-gradient(135deg, #f8faff, #f0f4ff) !important;
    min-width: unset !important;
    cursor: pointer !important;
    transition: all 0.2s ease !important;
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    margin-top: 2px !important;
}
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) > button:hover,
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) > span > button:hover {
    border-color: #c7d2fe !important;
    background: linear-gradient(135deg, #eef2ff, #e0e7ff) !important;
    box-shadow: 0 2px 10px rgba(99,102,241,0.08) !important;
}
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) > button::after,
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) > span > button::after {
    content: 'Đổi file' !important;
    font-size: 12px !important;
    font-weight: 500 !important;
    font-family: 'Inter', sans-serif !important;
    color: #818cf8 !important;
    line-height: normal !important;
}
/* Ẩn mọi child text bên trong nút compact */
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) > button > *,
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) > span > button > * {
    display: none !important;
}

/* ── Ẩn nút browse dư thừa / nút "+" thừa ── */
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) > span:has(button) {
    display: inline-flex !important;
    justify-content: center !important;
    width: auto !important;
}
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) > span:nth-of-type(n+2) {
    display: none !important;
}

/* Platform badge — Pill style */
.badge {
    display: inline-flex; align-items: center; gap: 4px;
    font-size: 11px; font-weight: 600;
    padding: 4px 12px; border-radius: 100px;
    letter-spacing: 0.3px;
    font-family: 'Inter', sans-serif;
    vertical-align: middle;
    margin-left: 6px;
}
.badge-tiktok {
    background: linear-gradient(135deg, #1e1b4b, #312e81);
    color: #e0e7ff;
    box-shadow: 0 2px 8px rgba(30,27,75,0.15);
}
.badge-shopee {
    background: linear-gradient(135deg, #fff1f0, #ffe4e1);
    color: #dc2626;
    border: 1px solid rgba(220,38,38,0.12);
}

/* ════════════════════════════════════════════════════════════
   METRIC CARDS — Soft glass with gradient orb
════════════════════════════════════════════════════════════ */
.os-metrics {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 14px;
    margin: 20px 0;
}
.os-metric {
    background: rgba(255,255,255,0.72);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    border: 1px solid rgba(229,231,235,0.50);
    border-radius: 22px;
    padding: 24px 22px;
    position: relative; overflow: hidden;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02), 0 8px 32px rgba(99,102,241,0.04);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}
.os-metric:hover {
    transform: translateY(-3px);
    box-shadow: 0 4px 8px rgba(0,0,0,0.03), 0 16px 48px rgba(99,102,241,0.08);
}
/* Gradient orb in each metric card */
.os-metric:nth-child(1)::before {
    content: '';
    position: absolute; top: -30px; right: -30px;
    width: 100px; height: 100px;
    background: radial-gradient(circle, rgba(196,181,253,0.35) 0%, transparent 70%);
    border-radius: 50%;
}
.os-metric:nth-child(2)::before {
    content: '';
    position: absolute; top: -30px; right: -30px;
    width: 100px; height: 100px;
    background: radial-gradient(circle, rgba(253,186,205,0.30) 0%, transparent 70%);
    border-radius: 50%;
}
.os-metric:nth-child(3)::before {
    content: '';
    position: absolute; top: -30px; right: -30px;
    width: 100px; height: 100px;
    background: radial-gradient(circle, rgba(167,243,208,0.35) 0%, transparent 70%);
    border-radius: 50%;
}
.os-metric-icon {
    width: 38px; height: 38px;
    background: rgba(241,245,249,0.80);
    border-radius: 12px;
    display: flex; align-items: center; justify-content: center;
    font-size: 17px; margin-bottom: 16px;
    border: 1px solid rgba(229,231,235,0.40);
}
.os-metric-label {
    font-size: 11.5px; font-weight: 500;
    color: #94a3b8; text-transform: uppercase;
    letter-spacing: 0.8px; margin-bottom: 8px;
}
.os-metric-value {
    font-size: 38px; font-weight: 800;
    color: #1e293b; letter-spacing: -2px;
    line-height: 1; margin-bottom: 6px;
    font-variant-numeric: tabular-nums;
}
.os-metric-note {
    font-size: 12.5px; color: #b0b8c9; font-weight: 400;
}

/* ════════════════════════════════════════════════════════════
   DEDUP BLOCK — Glass card
════════════════════════════════════════════════════════════ */
.os-dedup {
    background: rgba(255,255,255,0.72);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    border: 1px solid rgba(229,231,235,0.50);
    border-radius: 20px;
    padding: 20px 24px;
    display: flex; align-items: center; gap: 20px;
    margin: 14px 0 22px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02), 0 8px 32px rgba(99,102,241,0.04);
}
.os-dedup.has-removed {
    border-left: 4px solid #fbbf24;
    background: linear-gradient(135deg, rgba(255,251,235,0.80) 0%, rgba(255,255,255,0.72) 40%);
}
.os-dedup.no-removed {
    border-left: 4px solid #34d399;
    background: linear-gradient(135deg, rgba(236,253,245,0.80) 0%, rgba(255,255,255,0.72) 40%);
}
.os-dedup-icon { font-size: 24px; flex-shrink: 0; }
.os-dedup-body { flex: 1; }
.os-dedup-title {
    font-size: 14px; font-weight: 600; color: #334155;
    margin-bottom: 10px;
}
.os-dedup-pills { display: flex; gap: 6px; flex-wrap: wrap; }
.os-dedup-pill {
    font-size: 11.5px; font-weight: 500;
    font-family: 'Inter', sans-serif;
    padding: 4px 12px; border-radius: 100px;
    background: rgba(241,245,249,0.80); color: #94a3b8;
    border: 1px solid rgba(229,231,235,0.40);
}
.os-dedup-pill.removed {
    background: rgba(254,249,195,0.60); color: #b45309;
    border-color: rgba(253,224,71,0.30);
}
.os-dedup-pill.kept {
    background: rgba(209,250,229,0.60); color: #059669;
    border-color: rgba(52,211,153,0.20);
}
.os-dedup-count {
    text-align: right; flex-shrink: 0;
    font-size: 34px; font-weight: 800;
    letter-spacing: -1.5px;
    font-variant-numeric: tabular-nums;
}
.os-dedup-count.warn { color: #f59e0b; }
.os-dedup-count.ok   { color: #34d399; }
.os-dedup-count-lbl {
    font-size: 11.5px; color: #b0b8c9;
    margin-top: 2px; text-align: right; font-weight: 400;
}

/* ════════════════════════════════════════════════════════════
   SKU HEADER ROW — Glass pill
════════════════════════════════════════════════════════════ */
.os-sku-header {
    display: flex; align-items: center; justify-content: space-between;
    padding: 14px 20px;
    background: rgba(255,255,255,0.72);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    border: 1px solid rgba(229,231,235,0.50);
    border-radius: 16px;
    margin: 20px 0 8px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02), 0 4px 16px rgba(99,102,241,0.03);
}
.os-sku-left { display: flex; align-items: center; gap: 12px; }
.os-sku-code {
    font-family: 'JetBrains Mono', monospace;
    font-size: 12.5px; font-weight: 600;
    color: #6366f1;
    background: linear-gradient(135deg, #eef2ff, #e0e7ff);
    padding: 5px 14px; border-radius: 10px;
    letter-spacing: 0.3px;
}
.os-sku-label { font-size: 12.5px; color: #b0b8c9; font-weight: 400; }
.os-sku-total {
    font-size: 15px; font-weight: 700;
    color: #334155; letter-spacing: -0.3px;
    font-variant-numeric: tabular-nums;
}

/* ════════════════════════════════════════════════════════════
   DOWNLOAD BUTTON — Soft gradient
════════════════════════════════════════════════════════════ */
.stDownloadButton > button {
    background: linear-gradient(135deg, #818cf8, #a78bfa) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 16px !important;
    padding: 15px 28px !important;
    font-weight: 600 !important;
    font-size: 13.5px !important;
    letter-spacing: 0.1px !important;
    font-family: 'Inter', sans-serif !important;
    box-shadow: 0 4px 24px rgba(129,140,248,0.30) !important;
    transition: all 0.3s cubic-bezier(.34,1.56,.64,1) !important;
    width: 100% !important;
}
.stDownloadButton > button:hover {
    box-shadow: 0 8px 36px rgba(129,140,248,0.45) !important;
    transform: translateY(-2px) !important;
    background: linear-gradient(135deg, #6366f1, #818cf8) !important;
}
.stDownloadButton > button:active {
    transform: translateY(0) !important;
    box-shadow: 0 2px 12px rgba(129,140,248,0.25) !important;
}

/* ════════════════════════════════════════════════════════════
   DOWNLOAD WRAPPER — Glass
════════════════════════════════════════════════════════════ */
.os-dl-wrap {
    border: 1px solid rgba(229,231,235,0.50);
    border-radius: 22px;
    padding: 22px;
    background: rgba(255,255,255,0.72);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    margin: 20px 0;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02), 0 8px 32px rgba(99,102,241,0.04);
}
.os-dl-meta {
    display: flex; align-items: center; justify-content: space-between;
    margin-top: 14px; padding-top: 14px;
    border-top: 1px solid rgba(241,245,249,0.80);
}
.os-dl-filename {
    font-size: 12px; color: #b0b8c9;
    font-family: 'JetBrains Mono', monospace;
}
.os-dl-stats {
    font-size: 12.5px; color: #94a3b8; font-weight: 500;
}

/* ════════════════════════════════════════════════════════════
   EMPTY STATE — Floating card with orb
════════════════════════════════════════════════════════════ */
.os-empty {
    text-align: center;
    padding: 64px 32px;
    border: 1px solid rgba(229,231,235,0.50);
    border-radius: 28px;
    background: rgba(255,255,255,0.72);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    margin: 14px 0 36px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02), 0 8px 32px rgba(99,102,241,0.04);
    position: relative;
    overflow: hidden;
}
.os-empty::before {
    content: '';
    position: absolute;
    top: 30px; left: 50%;
    transform: translateX(-50%);
    width: 180px; height: 180px;
    background: radial-gradient(circle, rgba(196,181,253,0.22) 0%, transparent 70%);
    border-radius: 50%;
    animation: emptyOrb 6s ease-in-out infinite;
}
@keyframes emptyOrb {
    0%, 100% { transform: translateX(-50%) scale(1); opacity: 0.7; }
    50% { transform: translateX(-50%) scale(1.15); opacity: 1; }
}
.os-empty-icon {
    font-size: 36px; margin-bottom: 20px; display: block;
    position: relative;
    animation: emptyFloat 3.5s ease-in-out infinite;
}
@keyframes emptyFloat {
    0%,100% { transform: translateY(0); }
    50%      { transform: translateY(-6px); }
}
.os-empty-title {
    font-size: 18px; font-weight: 700; color: #334155;
    margin-bottom: 10px; letter-spacing: -0.3px;
    position: relative;
}
.os-empty-sub {
    font-size: 14px; color: #94a3b8;
    max-width: 380px; margin: 0 auto 36px;
    line-height: 1.75; font-weight: 400;
    position: relative;
}
.os-steps-row {
    display: flex; align-items: flex-start;
    justify-content: center; gap: 0;
    position: relative;
}
.os-step-item {
    display: flex; flex-direction: column; align-items: center;
    gap: 8px; padding: 0 20px;
    position: relative;
}
.os-step-item:not(:last-child)::after {
    content: '';
    position: absolute;
    top: 15px; right: -4px;
    width: 28px; height: 2px;
    background: linear-gradient(90deg, rgba(196,181,253,0.40), rgba(196,181,253,0.10));
    border-radius: 1px;
}
.os-step-num {
    width: 32px; height: 32px;
    background: linear-gradient(135deg, #eef2ff, #e0e7ff);
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-size: 12px; font-weight: 700; color: #818cf8;
    font-family: 'JetBrains Mono', monospace;
    border: 1px solid rgba(199,210,254,0.50);
}
.os-step-txt {
    font-size: 11.5px; color: #b0b8c9; font-weight: 500;
    white-space: nowrap;
}

/* ════════════════════════════════════════════════════════════
   UPLOAD NOTE — Minimal
════════════════════════════════════════════════════════════ */
.os-note {
    font-size: 12.5px; color: #b0b8c9;
    margin-top: 12px; line-height: 1.7;
    padding: 2px 4px;
}
.os-note strong { color: #818cf8; font-weight: 600; }
.os-note::before { content: ''; }

/* ════════════════════════════════════════════════════════════
   DIVIDER — Soft gradient
════════════════════════════════════════════════════════════ */
.os-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(196,181,253,0.20), transparent);
    margin: 28px 0;
}

/* ════════════════════════════════════════════════════════════
   DATAFRAME — Glass
════════════════════════════════════════════════════════════ */
.stDataFrame {
    border: 1px solid rgba(229,231,235,0.50) !important;
    border-radius: 16px !important;
    overflow: hidden !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02) !important;
}
iframe { border-radius: 16px; }

/* ════════════════════════════════════════════════════════════
   SECTION TAG
════════════════════════════════════════════════════════════ */
.os-section-tag {
    font-size: 10px; font-weight: 600;
    color: #818cf8;
    background: linear-gradient(135deg, #eef2ff, #e0e7ff);
    padding: 4px 12px; border-radius: 100px;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    font-family: 'JetBrains Mono', monospace;
    margin-left: auto;
    border: 1px solid rgba(199,210,254,0.40);
}
.os-section-sep { flex: 1; }



/* ════════════════════════════════════════════════════════════
   THIRD GRADIENT ORB (via JS-injected element)
════════════════════════════════════════════════════════════ */
.os-orb-pink {
    position: fixed;
    top: 40%; right: 5%;
    width: 260px; height: 260px;
    background: radial-gradient(circle, rgba(253,186,205,0.25) 0%, transparent 70%);
    border-radius: 50%;
    pointer-events: none;
    z-index: 0;
    animation: orbFloat3 20s ease-in-out infinite;
}
@keyframes orbFloat3 {
    0%, 100% { transform: translate(0, 0) scale(1); }
    40% { transform: translate(-30px, 40px) scale(1.06); }
    80% { transform: translate(20px, -20px) scale(0.94); }
}

/* ════════════════════════════════════════════════════════════
   RESPONSIVE
════════════════════════════════════════════════════════════ */
@media (max-width: 768px) {
    .os-topbar { flex-wrap: wrap; gap: 12px; }
    .os-topbar-status { margin-left: auto; }
    .os-metrics { grid-template-columns: 1fr; }
    .os-dedup { flex-direction: column; text-align: center; }
    .os-dedup-count { text-align: center; }
    .os-dedup-count-lbl { text-align: center; }
    .os-steps-row { flex-wrap: wrap; gap: 12px; }
    .os-step-item:not(:last-child)::after { display: none; }
    .os-topbar-name { font-size: 19px; }
}
</style>
""", unsafe_allow_html=True)

# Inject extra gradient orb
st.markdown('<div class="os-orb-pink"></div>', unsafe_allow_html=True)

# ─── HÀM TÁCH SKU ─────────────────────────────────────────────────────────────
def parse_sku_from_col_g(val):
    val = str(val).strip()
    return val.split('_')[0] if '_' in val else val

# ─── HÀM TÁCH MÀU & SIZE ──────────────────────────────────────────────────────
def parse_color_size_from_col_i(val):
    val = str(val).strip()
    if ',' in val:
        last_comma_index = val.rfind(',')
        color = val[:last_comma_index].strip()
        size = val[last_comma_index + 1:].strip() if last_comma_index < len(val) - 1 else "F"
        if ':' in size:
            size = size.split(':')[0].strip()
        color = ' '.join(color.replace(',', ' ').split())
    else:
        color = val
        size = "F"
    if ':' in size:
        size = size.split(':')[0].strip()
    return pd.Series([color, size])

# ─── HÀM HELPER FIELD CODE ────────────────────────────────────────────────────
def add_page_number_field(paragraph):
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._element.append(fldChar1)
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'PAGE'
    run._element.append(instrText)
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    run._element.append(fldChar2)
    t = OxmlElement('w:t')
    run._element.append(t)
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._element.append(fldChar3)
    return run

def add_num_pages_field(paragraph):
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._element.append(fldChar1)
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'NUMPAGES'
    run._element.append(instrText)
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    run._element.append(fldChar2)
    t = OxmlElement('w:t')
    run._element.append(t)
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._element.append(fldChar3)
    return run

# ─── HÀM HELPER DOCX ──────────────────────────────────────────────────────────
def set_cell_shading(cell, color_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color_hex)
    shd.set(qn('w:val'), 'clear')
    tcPr.append(shd)

def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            element = OxmlElement(f'w:{edge}')
            element.set(qn('w:val'), val.get('val', 'single'))
            element.set(qn('w:sz'), val.get('sz', '4'))
            element.set(qn('w:color'), val.get('color', '000000'))
            element.set(qn('w:space'), '0')
            tcBorders.append(element)
    tcPr.append(tcBorders)

def set_row_height(row, height_cm=0.6):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(int(height_cm * 567)))
    trHeight.set(qn('w:hRule'), 'atLeast')
    trPr.append(trHeight)

def set_cell_margins(cell, top=30, bottom=30, left=60, right=60):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in [('top', top), ('bottom', bottom), ('start', left), ('end', right)]:
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:w'), str(val))
        el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)

def format_cell(cell, text, font_size=10, bold=False, align='left', font_name='Arial', color=None, indent=False):
    cell.text = ''
    para = cell.paragraphs[0]
    para_format = para.paragraph_format
    para_format.space_before = Pt(1)
    para_format.space_after = Pt(1)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    prefix = '    ' if indent else ''
    run = para.add_run(prefix + str(text))
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.name = font_name
    if color:
        run.font.color.rgb = RGBColor(*color)
    if align == 'center':
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right':
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT

# ─── HÀM XUẤT WORD ────────────────────────────────────────────────────────────
def export_to_word(detail_summary, total_orders, total_items, shop_name="TITIKID", platform="SHOPEE", shift="CHIỀU"):
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.4)

    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_para.clear()
    run_page = add_page_number_field(footer_para)
    run_page.font.size = Pt(9)
    run_page.font.name = 'Arial'
    run_sep = footer_para.add_run(' / ')
    run_sep.font.size = Pt(9)
    run_sep.font.name = 'Arial'
    run_total = add_num_pages_field(footer_para)
    run_total.font.size = Pt(9)
    run_total.font.name = 'Arial'

    current_date = datetime.now().strftime('%d.%m')
    title_text = f"{shop_name} - {platform} - {shift} - {current_date} - {total_orders} ĐƠN - {total_items} ÁO"
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(title_text)
    title_run.font.size = Pt(13)
    title_run.font.bold = True
    title_run.font.name = 'Arial'
    title_run.font.color.rgb = RGBColor(0, 0, 0)
    title_para.paragraph_format.space_after = Pt(6)

    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    table.autofit = False
    col_widths = [Inches(1.5), Inches(2.5), Inches(2.2), Inches(1.0)]

    hdr_cells = table.rows[0].cells
    headers = ['SKU sản phẩm', 'Màu', 'Size', 'ÁO']
    header_bg = 'C5D0DC'
    for i, header_text in enumerate(headers):
        format_cell(hdr_cells[i], header_text, font_size=10, bold=True, align='left' if i < 3 else 'right', font_name='Arial')
        set_cell_shading(hdr_cells[i], header_bg)
    set_row_height(table.rows[0], 0.9)

    top_total_cells = table.add_row().cells
    top_total_bg = 'DCE6F1'
    format_cell(top_total_cells[0], 'Tổng số', font_size=10, bold=True, font_name='Arial')
    format_cell(top_total_cells[1], '', font_size=10, font_name='Arial')
    format_cell(top_total_cells[2], '', font_size=10, font_name='Arial')
    format_cell(top_total_cells[3], '0', font_size=10, bold=True, align='right', font_name='Arial', color=(0, 0, 139))
    for cell in top_total_cells:
        set_cell_shading(cell, top_total_bg)
    set_row_height(table.rows[-1], 0.85)

    unique_skus = detail_summary['SKU'].unique()
    for sku in unique_skus:
        sku_data = detail_summary[detail_summary['SKU'] == sku].copy()
        total_sku = int(sku_data['SL'].sum())
        sku_data = sku_data.sort_values(by=['Phân loại', 'Size'])
        for _, data_row in sku_data.iterrows():
            row_cells = table.add_row().cells
            format_cell(row_cells[0], sku, font_size=10, bold=True, font_name='Arial')
            format_cell(row_cells[1], str(data_row['Phân loại']), font_size=10, bold=False, font_name='Arial')
            format_cell(row_cells[2], str(data_row['Size']), font_size=10, font_name='Arial')
            format_cell(row_cells[3], str(int(data_row['SL'])), font_size=10, align='right', font_name='Arial')
            for cell in row_cells:
                set_cell_shading(cell, 'FFFFFF')
            set_row_height(table.rows[-1], 0.85)

        total_row_cells = table.add_row().cells
        format_cell(total_row_cells[0], f'Tổng số {sku}', font_size=10, bold=True, font_name='Arial')
        format_cell(total_row_cells[1], '', font_size=10, font_name='Arial')
        format_cell(total_row_cells[2], '', font_size=10, font_name='Arial')
        format_cell(total_row_cells[3], str(total_sku), font_size=10, bold=True, align='right', font_name='Arial', color=(0, 0, 139))
        for cell in total_row_cells:
            set_cell_shading(cell, 'DCE6F1')
        set_row_height(table.rows[-1], 0.85)

    grand_total_cells = table.add_row().cells
    format_cell(grand_total_cells[0], 'Tổng cộng', font_size=10, bold=True, font_name='Arial')
    format_cell(grand_total_cells[1], '', font_size=10, font_name='Arial')
    format_cell(grand_total_cells[2], '', font_size=10, font_name='Arial')
    format_cell(grand_total_cells[3], str(total_items), font_size=10, bold=True, align='right', font_name='Arial', color=(0, 0, 139))
    for cell in grand_total_cells:
        set_cell_shading(cell, 'DCE6F1')
    set_row_height(table.rows[-1], 0.85)

    for row in table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = w

    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    for row_idx, row in enumerate(table.rows):
        if row_idx == 0:
            trPr = row._tr.get_or_add_trPr()
            tblHeader = OxmlElement('w:tblHeader')
            trPr.append(tblHeader)

    target = io.BytesIO()
    doc.save(target)
    return target.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
#  TIÊU ĐỀ
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('''
<div class="os-topbar">
    <div class="os-topbar-brand">
        <div class="os-topbar-logo">📦</div>
        <div>
            <div class="os-topbar-name">Xin chào, <strong>Order Studio</strong></div>
            <div class="os-topbar-sub">TikTok & Shopee · Tổng hợp · Lọc trùng · Xuất Word</div>
        </div>
    </div>
    <div class="os-topbar-status">
        <span class="os-status-dot"></span> Sẵn sàng
    </div>
</div>
''', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  01 — CÀI ĐẶT
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('''
<div class="os-section">
    <span class="os-section-num">01</span>
    <span class="os-section-title">Cài đặt</span>
</div>
''', unsafe_allow_html=True)


current_hour = datetime.now().hour
default_shift_index = 0 if current_hour < 12 else 1

col_s1, col_s2, col_s3 = st.columns(3)
with col_s1:
    shop_name = st.selectbox("Tên shop", ["TITIKID", "GIMME"], key="shop_name")
with col_s2:
    platform = st.selectbox("Sàn bán hàng", ["TIKTOK", "SHOPEE"], key="platform")
with col_s3:
    shift = st.selectbox("Ca làm việc", ["SÁNG", "CHIỀU"], index=default_shift_index, key="shift")


# ══════════════════════════════════════════════════════════════════════════════
#  02 — UPLOAD FILE
# ══════════════════════════════════════════════════════════════════════════════
platform_badge = (
    '<span class="badge badge-tiktok">TikTok</span>'
    if platform == "TIKTOK"
    else '<span class="badge badge-shopee">Shopee</span>'
)

st.markdown(f'''
<div class="os-section" style="margin-top: 20px;">
    <span class="os-section-num">02</span>
    <span class="os-section-title">Tải file đơn hàng {platform_badge}</span>
</div>
''', unsafe_allow_html=True)

col_f1, col_f2 = st.columns(2)
with col_f1:
    uploaded_file = st.file_uploader(
        f"📄 File ca hiện tại ({platform}) · Bắt buộc",
        type=["csv", "xlsx"],
        key="file_current",
        help="File đơn hàng ca này cần tổng hợp"
    )
with col_f2:
    prev_file = st.file_uploader(
        "📂 File ca trước · Tùy chọn (lọc đơn trùng)",
        type=["csv", "xlsx"],
        key="file_prev",
        help="Tải lên để loại bỏ đơn đã soạn ở ca trước"
    )

st.markdown('''
<p class="os-note">
    Chỉ cần <strong>file ca hiện tại</strong> là đủ để tổng hợp và xuất Word.
    Muốn <strong>lọc đơn trùng</strong> thì tải thêm file ca trước.
</p>
''', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  03 — KẾT QUẢ
# ══════════════════════════════════════════════════════════════════════════════
if not uploaded_file:
    st.markdown('''
    <div class="os-section" style="margin-top:28px;">
        <span class="os-section-num">03</span>
        <span class="os-section-title">Kết quả</span>
    </div>
    <div class="os-empty">
        <span class="os-empty-icon">📂</span>
        <div class="os-empty-title">Chưa có dữ liệu</div>
        <div class="os-empty-sub">Tải lên file đơn hàng phía trên để bắt đầu tổng hợp, lọc trùng và xuất file Word soạn hàng.</div>
        <div class="os-steps-row">
            <div class="os-step-item">
                <div class="os-step-num">1</div>
                <div class="os-step-txt">Chọn cửa hàng</div>
            </div>
            <div class="os-step-item">
                <div class="os-step-num">2</div>
                <div class="os-step-txt">Tải file lên</div>
            </div>
            <div class="os-step-item">
                <div class="os-step-num">3</div>
                <div class="os-step-txt">Lọc trùng ca trước</div>
            </div>
            <div class="os-step-item">
                <div class="os-step-num">4</div>
                <div class="os-step-txt">Tải Word soạn hàng</div>
            </div>
        </div>
    </div>
    ''', unsafe_allow_html=True)

else:
    try:
        # ── Đọc file hiện tại ────────────────────────────────────────────────
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, low_memory=False, dtype=str)
        else:
            df = pd.read_excel(uploaded_file, engine='calamine', dtype=str)
        df = df.dropna(how='all').reset_index(drop=True)

        # ── Mapping cột theo sàn ─────────────────────────────────────────────
        if platform == "TIKTOK":
            sku_col_index, variation_col_index, qty_col_index = 6, 8, 9
        else:
            # Shopee: cột khác nhau theo shop
            if shop_name == "GIMME":
                # GIMME: T(19)=SKU, U(20)=Phân loại, AA(26)=Số lượng
                sku_col_index, variation_col_index, qty_col_index = 19, 20, 26
            else:
                # TITIKID: S(18)=SKU, T(19)=Phân loại, Z(25)=Số lượng
                sku_col_index, variation_col_index, qty_col_index = 18, 19, 25

        max_col_needed = max(sku_col_index, variation_col_index, qty_col_index)
        if len(df.columns) <= max_col_needed:
            st.error(f"❌ File không đủ cột cho sàn {platform}. Cần ít nhất {max_col_needed + 1} cột, file chỉ có {len(df.columns)} cột.")
            st.stop()

        col_sku       = df.columns[sku_col_index]
        col_variation = df.columns[variation_col_index]
        col_qty       = df.columns[qty_col_index]

        # Shopee: dùng cột G (mã vận đơn, index 6) thay vì cột A (mã đơn hàng)
        if platform == "SHOPEE":
            if len(df.columns) <= 6:
                st.error("❌ File Shopee không đủ cột. Cần ít nhất cột G (mã vận đơn).")
                st.stop()
            id_col = df.columns[6]  # Cột G = mã vận đơn
            # Bỏ những đơn không có mã vận đơn
            no_tracking_count = df[id_col].isna().sum() + (df[id_col].astype(str).str.strip() == '').sum()
            df = df[df[id_col].astype(str).str.strip().ne('') & df[id_col].notna()].reset_index(drop=True)
            if no_tracking_count > 0:
                st.info(f"ℹ️ Đã bỏ **{no_tracking_count}** đơn không có mã vận đơn (cột G trống).")
        else:
            id_col = df.columns[0]  # TikTok: vẫn dùng cột A

        total_raw     = df[id_col].nunique()
        removed_count = 0
        prev_total    = 0

        # ── Nhãn mục 03 ─────────────────────────────────────────────────
        st.markdown('''
        <div class="os-section" style="margin-top:28px;">
            <span class="os-section-num">03</span>
            <span class="os-section-title">Kết quả</span>
            <span class="os-section-sep"></span>
            <span class="os-section-tag">xuất dữ liệu</span>
        </div>
        ''', unsafe_allow_html=True)

        # ── Lọc trùng với file buổi trước ───────────────────────────────────
        if prev_file is not None:
            try:
                if prev_file.name.endswith('.csv'):
                    df_prev = pd.read_csv(prev_file, low_memory=False, dtype=str)
                else:
                    df_prev = pd.read_excel(prev_file, engine='calamine', dtype=str)
                df_prev = df_prev.dropna(how='all').reset_index(drop=True)

                # Shopee: dùng cột G (mã vận đơn) cho file ca trước
                if platform == "SHOPEE" and len(df_prev.columns) > 6:
                    prev_id_col = df_prev.columns[6]
                    # Bỏ dòng không có mã vận đơn trong file ca trước
                    df_prev = df_prev[df_prev[prev_id_col].astype(str).str.strip().ne('') & df_prev[prev_id_col].notna()].reset_index(drop=True)
                else:
                    prev_id_col = df_prev.columns[0]

                prev_ids     = set(df_prev[prev_id_col].dropna().astype(str).str.strip())
                current_ids  = set(df[id_col].dropna().astype(str).str.strip())
                duplicated   = current_ids & prev_ids
                removed_count = len(duplicated)
                prev_total   = len(prev_ids)

                df = df[~df[id_col].astype(str).str.strip().isin(duplicated)].reset_index(drop=True)
                kept_count = df[id_col].nunique()

                if removed_count > 0:
                    st.markdown(f'''
                    <div class="os-dedup has-removed">
                        <div class="os-dedup-icon">⇄</div>
                        <div class="os-dedup-body">
                            <div class="os-dedup-title">Đã lọc trùng với file ca trước</div>
                            <div class="os-dedup-pills">
                                <span class="os-dedup-pill">Ca này: {total_raw} đơn</span>
                                <span class="os-dedup-pill">Ca trước: {prev_total} đơn</span>
                                <span class="os-dedup-pill removed">Loại: {removed_count} trùng</span>
                                <span class="os-dedup-pill kept">Giữ: {kept_count} mới</span>
                            </div>
                        </div>
                        <div>
                            <div class="os-dedup-count warn">−{removed_count}</div>
                            <div class="os-dedup-count-lbl">đơn trùng</div>
                        </div>
                    </div>
                    ''', unsafe_allow_html=True)
                else:
                    st.markdown(f'''
                    <div class="os-dedup no-removed">
                        <div class="os-dedup-icon">✓</div>
                        <div class="os-dedup-body">
                            <div class="os-dedup-title">Không có đơn trùng</div>
                            <div class="os-dedup-pills">
                                <span class="os-dedup-pill">Ca này: {total_raw} đơn</span>
                                <span class="os-dedup-pill">Ca trước: {prev_total} đơn</span>
                                <span class="os-dedup-pill kept">Tất cả đều là đơn mới</span>
                            </div>
                        </div>
                        <div>
                            <div class="os-dedup-count ok">0</div>
                            <div class="os-dedup-count-lbl">đơn trùng</div>
                        </div>
                    </div>
                    ''', unsafe_allow_html=True)
            except Exception as e_prev:
                st.warning(f"⚠️ Không đọc được file buổi trước: {e_prev}")

        # ── Xử lý data ───────────────────────────────────────────────────────
        df['SKU_ID'] = df[col_sku].apply(parse_sku_from_col_g)
        df[['PL', 'SZ']] = df[col_variation].apply(parse_color_size_from_col_i)
        df['SL'] = pd.to_numeric(df[col_qty], errors='coerce').fillna(0).astype(int)

        total_items  = int(df['SL'].sum())
        total_orders = df[id_col].nunique()

        detail_summary = df.groupby(['SKU_ID', 'PL', 'SZ'])['SL'].sum().reset_index()
        detail_summary.columns = ['SKU', 'Phân loại', 'Size', 'SL']
        unique_skus_count = detail_summary['SKU'].nunique()

        # ── Metric Cards ─────────────────────────────────────────────────────
        dedup_note = f"sau khi lọc {removed_count} trùng" if removed_count > 0 else "đơn hàng duy nhất"
        st.markdown(f'''
        <div class="os-metrics">
            <div class="os-metric">
                <div class="os-metric-icon">📦</div>
                <div class="os-metric-label">Tổng đơn hàng</div>
                <div class="os-metric-value">{total_orders}</div>
                <div class="os-metric-note">{dedup_note}</div>
            </div>
            <div class="os-metric">
                <div class="os-metric-icon">👕</div>
                <div class="os-metric-label">Tổng sản phẩm</div>
                <div class="os-metric-value">{total_items}</div>
                <div class="os-metric-note">tổng số lượng áo</div>
            </div>
            <div class="os-metric">
                <div class="os-metric-icon">🏷️</div>
                <div class="os-metric-label">Loại SKU</div>
                <div class="os-metric-value">{unique_skus_count}</div>
                <div class="os-metric-note">mã sản phẩm khác nhau</div>
            </div>
        </div>
        ''', unsafe_allow_html=True)

        # ── Download Word ─────────────────────────────────────────────────────
        current_date_file = datetime.now().strftime('%d.%m')
        word_filename = f"{shop_name}_{platform}_{shift}_{current_date_file}.docx"
        word_data = export_to_word(detail_summary, total_orders, total_items, shop_name, platform, shift)

        st.markdown('<div class="os-dl-wrap">', unsafe_allow_html=True)
        st.download_button(
            f"⬇️  Tải xuống phiếu soạn hàng  —  {total_orders} đơn · {total_items} áo",
            word_data,
            word_filename,
            use_container_width=True
        )
        st.markdown(f'''
            <div class="os-dl-meta">
                <span class="os-dl-filename">{word_filename}</span>
                <span class="os-dl-stats">{total_orders} đơn · {total_items} sản phẩm · {unique_skus_count} mã SKU</span>
            </div>
        </div>''', unsafe_allow_html=True)

        # ── Chi tiết theo SKU ─────────────────────────────────────────────────
        st.markdown('<div class="os-divider"></div>', unsafe_allow_html=True)

        unique_skus = detail_summary['SKU'].unique()
        for sku in unique_skus:
            sku_data  = detail_summary[detail_summary['SKU'] == sku].sort_values(by='Size')
            total_sku = int(sku_data['SL'].sum())

            st.markdown(f'''
            <div class="os-sku-header">
                <div class="os-sku-left">
                    <span class="os-sku-code">{sku}</span>
                    <span class="os-sku-label">sản phẩm</span>
                </div>
                <span class="os-sku-total">{total_sku} cái</span>
            </div>
            ''', unsafe_allow_html=True)

            st.dataframe(
                sku_data[['Phân loại', 'Size', 'SL']],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Phân loại": st.column_config.TextColumn("Phân loại / Màu", width="large"),
                    "Size": st.column_config.TextColumn("Size", width="medium"),
                    "SL": st.column_config.NumberColumn("SL", width="small", format="%d")
                }
            )

    except Exception as e:
        st.error(f"❌ Lỗi xử lý file: {e}")