# -*- coding: utf-8 -*-
import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches

os.makedirs("assets", exist_ok=True)

# 1. Diagram Bab 2: Arsitektur 80:20 (Winning Engine vs Asymmetric Sandbox)
def generate_bab2_diagram():
    fig, ax = plt.subplots(figsize=(10, 5.5), dpi=200)
    fig.patch.set_facecolor('#0F172A')
    ax.set_facecolor('#0F172A')
    
    # Title
    ax.text(0.5, 0.92, "ARSITEKTUR KONTEN 80:20 (100 KONTEN)", 
            color='#FFFFFF', fontsize=16, fontweight='bold', ha='center', va='center')
    ax.text(0.5, 0.85, "Keseimbangan Antara Skala Eksploitasi & Eksplorasi Asimetris", 
            color='#94A3B8', fontsize=11, ha='center', va='center')

    # Box 80%
    box80 = patches.FancyBboxPatch((0.08, 0.22), 0.52, 0.52,
                                   boxstyle="round,pad=0.02,rounding_size=0.03",
                                   facecolor='#1E293B', edgecolor='#D9531E', linewidth=2.5)
    ax.add_patch(box80)
    
    ax.text(0.34, 0.65, "80% WINNING ENGINE", color='#D9531E', fontsize=14, fontweight='bold', ha='center')
    ax.text(0.34, 0.57, "(80 Konten Eksploitasi Format Juara)", color='#F8FAFC', fontsize=11, fontweight='semibold', ha='center')
    
    details_80 = [
        "• 20 Konten: Bedah Website Brand / Figur Terkenal",
        "• 20 Konten: 3 Kesalahan Fatal & Solusi Praktis",
        "• 20 Konten: Analisis Mikro (Headline & CTA)",
        "• 20 Konten: Live Redesign Before & After"
    ]
    y_pos = 0.47
    for d in details_80:
        ax.text(0.12, y_pos, d, color='#CBD5E1', fontsize=9.5, ha='left')
        y_pos -= 0.06
        
    ax.text(0.34, 0.26, "Tujuan: Memompa Reach, Pengikut Baru & Otoritas", 
            color='#38BDF8', fontsize=9.5, fontstyle='italic', ha='center')

    # Box 20%
    box20 = patches.FancyBboxPatch((0.64, 0.22), 0.28, 0.52,
                                   boxstyle="round,pad=0.02,rounding_size=0.03",
                                   facecolor='#1E293B', edgecolor='#10B981', linewidth=2.5)
    ax.add_patch(box20)
    
    ax.text(0.78, 0.65, "20% SANDBOX", color='#10B981', fontsize=14, fontweight='bold', ha='center')
    ax.text(0.78, 0.57, "(20 Konten Laboratorium)", color='#F8FAFC', fontsize=11, fontweight='semibold', ha='center')
    
    details_20 = [
        "• 8 Konten: Jual Jasa & Studi Kasus",
        "• 6 Konten: Personal Brand / Vlog",
        "• 6 Konten: Eksperimen Format Liar"
    ]
    y_pos = 0.47
    for d in details_20:
        ax.text(0.66, y_pos, d, color='#CBD5E1', fontsize=9.5, ha='left')
        y_pos -= 0.06
        
    ax.text(0.78, 0.26, "Tujuan: Konversi Klien & R&D Masa Depan", 
            color='#34D399', fontsize=9.5, fontstyle='italic', ha='center')

    # Bottom summary tag
    ax.text(0.5, 0.08, "Struktur Risiko Asimetris: Kerugian Terbatas (Capped Downside), Peluang Tak Terbatas (Unlimited Upside)",
            color='#E2E8F0', fontsize=9.5, ha='center', va='center',
            bbox=dict(boxstyle="round,pad=0.5", facecolor='#334155', edgecolor='none'))

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')
    
    plt.tight_layout()
    plt.savefig('assets/diagram_bab2_arsitektur.png', dpi=200, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Diagram Bab 2 tersimpan di assets/diagram_bab2_arsitektur.png")

# 2. Diagram Bab 6: Kalender Mingguan Rasio 4:1
def generate_bab6_diagram():
    fig, ax = plt.subplots(figsize=(10, 5), dpi=200)
    fig.patch.set_facecolor('#0F172A')
    ax.set_facecolor('#0F172A')
    
    ax.text(0.5, 0.90, "JADWAL DISTRIBUSI MINGGUAN RASIO 4:1", 
            color='#FFFFFF', fontsize=15, fontweight='bold', ha='center')
    ax.text(0.5, 0.82, "Alur Eksekusi Taktis Tanpa Mengorbankan Waktu & Kewarasan", 
            color='#94A3B8', fontsize=10.5, ha='center')

    days = [
        ("SENIN", "80% Pilar", "Bedah Brand Besar", "#D9531E"),
        ("SELASA", "80% Pilar", "3 Kesalahan Fatal", "#D9531E"),
        ("RABU", "80% Pilar", "Analisis Mikro CTA", "#D9531E"),
        ("KAMIS", "80% Pilar", "Live Redesign", "#D9531E"),
        ("JUMAT", "20% Jual", "Studi Kasus Jasa", "#10B981"),
        ("SABTU", "20% Lab", "Vlog / Eksperimen", "#38BDF8"),
        ("MINGGU", "Review", "Analisis Data", "#A855F7")
    ]
    
    col_width = 0.115
    start_x = 0.06
    spacing = 0.018
    
    for i, (day, cat, desc, color) in enumerate(days):
        x = start_x + i * (col_width + spacing)
        # Day card
        card = patches.FancyBboxPatch((x, 0.22), col_width, 0.50,
                                      boxstyle="round,pad=0.015,rounding_size=0.02",
                                      facecolor='#1E293B', edgecolor=color, linewidth=2)
        ax.add_patch(card)
        
        # Day header
        ax.text(x + col_width/2, 0.65, day, color='#F8FAFC', fontsize=11, fontweight='bold', ha='center')
        # Category badge
        ax.text(x + col_width/2, 0.55, cat, color=color, fontsize=9.5, fontweight='semibold', ha='center')
        # Description
        ax.text(x + col_width/2, 0.38, desc, color='#CBD5E1', fontsize=8.5, ha='center', va='center', multialignment='center')

    # Bottom note
    ax.text(0.5, 0.10, "⚡ Sistem Batching: Riset (2 Jam) + Rekam (2 Jam) + Edit (2 Jam) = Beres 1 Pekan Penuh",
            color='#FDE047', fontsize=10, ha='center', va='center',
            bbox=dict(boxstyle="round,pad=0.4", facecolor='#334155', edgecolor='none'))

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')
    
    plt.tight_layout()
    plt.savefig('assets/diagram_bab6_kalender.png', dpi=200, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Diagram Bab 6 tersimpan di assets/diagram_bab6_kalender.png")

if __name__ == "__main__":
    generate_bab2_diagram()
    generate_bab6_diagram()
