"""
Génère presentation.pptx — Mémoire M1 MIASHS Ardacham Mahamat Teguene
Style CM : fond blanc, titres bleu foncé, texte noir, graphiques inclus.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.enum.dml import MSO_THEME_COLOR
import copy
from lxml import etree

# ── Constantes couleurs ───────────────────────────────────────
BLEU      = RGBColor(0,  70, 140)
BLEU_ACC  = RGBColor(0, 102, 204)
BLEU_FOND = RGBColor(235, 242, 252)
GRIS_FOND = RGBColor(245, 247, 250)
NOIR      = RGBColor(0,   0,   0)
BLANC     = RGBColor(255, 255, 255)
ROUGE_ACC = RGBColor(180,  30,  30)

# ── Dimensions (16:9) ─────────────────────────────────────────
W = Inches(13.33)
H = Inches(7.5)
FIGS = "/Users/ardachammahamatteguene/Desktop/Bureau - MacBook Pro de Ardacham/Memoire/latex/figures/"

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

blank_layout = prs.slide_layouts[6]   # complètement vide

# ─────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────

def add_slide():
    return prs.slides.add_slide(blank_layout)

def rect(slide, x, y, w, h, fill_rgb, alpha=None):
    """Ajoute un rectangle plein."""
    shape = slide.shapes.add_shape(1, x, y, w, h)  # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.line.fill.background()
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    shape.line.color.rgb = fill_rgb
    return shape

def title_bar(slide, text, sub=None):
    """Bande titre bleue en haut + texte blanc."""
    rect(slide, 0, 0, W, Inches(1.0), BLEU)
    tb = slide.shapes.add_textbox(Inches(0.3), Inches(0.08), W - Inches(0.6), Inches(0.8))
    tf = tb.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = BLANC
    if sub:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.LEFT
        r2 = p2.add_run()
        r2.text = sub
        r2.font.size = Pt(14)
        r2.font.bold = False
        r2.font.color.rgb = RGBColor(200, 220, 255)

def textbox(slide, x, y, w, h, text, size=16, bold=False, color=NOIR,
            align=PP_ALIGN.LEFT, italic=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return tb

def bullet_box(slide, x, y, w, h, items, size=15, title=None, title_color=BLEU,
               body_color=NOIR, title_size=16):
    """Boîte avec titre optionnel et liste à puces."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    if title:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = title
        r.font.size = Pt(title_size)
        r.font.bold = True
        r.font.color.rgb = title_color
    for item in items:
        p = tf.paragraphs[0] if (first and not title) else tf.add_paragraph()
        first = False
        p.alignment = PP_ALIGN.LEFT
        p.level = 0
        r = p.add_run()
        r.text = "▸  " + item
        r.font.size = Pt(size)
        r.font.color.rgb = body_color

def blue_box(slide, x, y, w, h, title, items, size=14):
    """Bloc coloré style Beamer block."""
    rect(slide, x, y, w, Inches(0.38), BLEU)
    tb = slide.shapes.add_textbox(x + Inches(0.1), y + Inches(0.04),
                                   w - Inches(0.2), Inches(0.32))
    tf = tb.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = title
    r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = BLANC
    rect(slide, x, y + Inches(0.38), w, h - Inches(0.38), BLEU_FOND)
    tb2 = slide.shapes.add_textbox(x + Inches(0.12), y + Inches(0.42),
                                    w - Inches(0.24), h - Inches(0.5))
    tf2 = tb2.text_frame; tf2.word_wrap = True
    first = True
    for item in items:
        p2 = tf2.paragraphs[0] if first else tf2.add_paragraph()
        first = False
        p2.alignment = PP_ALIGN.LEFT
        r2 = p2.add_run()
        r2.text = "▸  " + item
        r2.font.size = Pt(size); r2.font.color.rgb = NOIR

def image(slide, path, x, y, w=None, h=None):
    if w and h:
        slide.shapes.add_picture(path, x, y, w, h)
    elif w:
        slide.shapes.add_picture(path, x, y, width=w)
    elif h:
        slide.shapes.add_picture(path, x, y, height=h)
    else:
        slide.shapes.add_picture(path, x, y)

def hline(slide, y, color=BLEU_ACC, width_pt=1.5):
    """Trait horizontal bleu."""
    from pptx.util import Pt as pt_unit
    ln = slide.shapes.add_shape(1, 0, y, W, Pt(1))
    ln.fill.background()
    ln.line.color.rgb = color
    ln.line.width = Pt(width_pt)

def add_table(slide, x, y, w, h, headers, rows, font_size=13, col_widths=None):
    """Ajoute un tableau stylé bleu/blanc/noir."""
    cols = len(headers)
    table = slide.shapes.add_table(len(rows)+1, cols, x, y, w, h).table
    if col_widths:
        for i, cw in enumerate(col_widths):
            table.columns[i].width = cw
    # En-têtes
    for j, hdr in enumerate(headers):
        cell = table.cell(0, j)
        cell.fill.solid(); cell.fill.fore_color.rgb = BLEU
        tf = cell.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = hdr
        r.font.size = Pt(font_size); r.font.bold = True; r.font.color.rgb = BLANC
    # Lignes
    for i, row in enumerate(rows):
        bg = BLANC if i % 2 == 0 else GRIS_FOND
        for j, val in enumerate(row):
            cell = table.cell(i+1, j)
            cell.fill.solid(); cell.fill.fore_color.rgb = bg
            tf = cell.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if j > 0 else PP_ALIGN.LEFT
            r = p.add_run(); r.text = str(val)
            r.font.size = Pt(font_size)
            r.font.color.rgb = NOIR
    return table

# ══════════════════════════════════════════════════════════════
# DIAPO 1 — TITRE
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
rect(sl, 0, Inches(2.6), W, Inches(2.5), BLEU)

textbox(sl, Inches(0.6), Inches(2.72),
        W - Inches(1.2), Inches(1.1),
        "Modélisation des comportements de santé",
        size=32, bold=True, color=BLANC, align=PP_ALIGN.CENTER)
textbox(sl, Inches(0.6), Inches(3.55),
        W - Inches(1.2), Inches(0.6),
        "par apprentissage automatique",
        size=26, bold=False, color=RGBColor(200,220,255), align=PP_ALIGN.CENTER)

textbox(sl, Inches(0.6), Inches(4.5),
        W - Inches(1.2), Inches(0.4),
        "Régression, classification et clustering sur données synthétiques",
        size=17, italic=True, color=RGBColor(80,80,80), align=PP_ALIGN.CENTER)

rect(sl, Inches(3.5), Inches(5.25), W - Inches(7.0), Pt(2), BLEU_ACC)

textbox(sl, Inches(0.6), Inches(5.5),
        W - Inches(1.2), Inches(0.35),
        "Ardacham Mahamat Teguene  —  Master 1 MIASHS  —  Université de Montpellier Paul Valéry",
        size=15, bold=True, color=BLEU, align=PP_ALIGN.CENTER)
textbox(sl, Inches(0.6), Inches(5.9),
        W - Inches(1.2), Inches(0.35),
        "Tuteurs : Sophie Lèbre & Jérôme Pasquet  —  Juin 2026",
        size=13, color=RGBColor(100,100,100), align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════
# DIAPO 2 — PLAN
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
title_bar(sl, "Plan de la présentation")

sections = [
    ("1", "Contexte et problématique"),
    ("2", "Jeu de données & pipeline CRISP-DM"),
    ("3", "Analyse exploratoire (EDA)"),
    ("4", "Modélisation supervisée — Régression"),
    ("5", "Modélisation supervisée — Classification"),
    ("6", "Clustering et interprétabilité"),
    ("7", "Bilan & Perspectives"),
]
for i, (num, sec) in enumerate(sections):
    y = Inches(1.25) + i * Inches(0.75)
    rect(sl, Inches(0.4), y, Inches(0.45), Inches(0.52), BLEU)
    textbox(sl, Inches(0.4), y + Inches(0.06), Inches(0.45), Inches(0.4),
            num, size=18, bold=True, color=BLANC, align=PP_ALIGN.CENTER)
    textbox(sl, Inches(1.05), y + Inches(0.08), Inches(10), Inches(0.4),
            sec, size=18, color=NOIR)

# ══════════════════════════════════════════════════════════════
# DIAPO 3 — CONTEXTE & PROBLÉMATIQUE
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
title_bar(sl, "Contexte et problématique")

# Bloc problématique
rect(sl, Inches(0.4), Inches(1.15), W - Inches(0.8), Inches(1.05), BLEU_FOND)
rect(sl, Inches(0.4), Inches(1.15), Inches(0.08), Inches(1.05), BLEU_ACC)
textbox(sl, Inches(0.6), Inches(1.2), W - Inches(1.1), Inches(0.3),
        "PROBLÉMATIQUE", size=13, bold=True, color=BLEU_ACC)
textbox(sl, Inches(0.6), Inches(1.48), W - Inches(1.1), Inches(0.6),
        "Quels facteurs du mode de vie influencent le plus l'état de santé d'un individu "
        "— et peut-on les identifier, les modéliser et les expliquer de manière rigoureuse ?",
        size=15, bold=True, color=BLEU)

bullet_box(sl, Inches(0.4), Inches(2.4), Inches(5.8), Inches(2.4),
    ["Les maladies chroniques résultent de comportements quotidiens mesurables",
     "ML peut capturer des relations complexes inaccessibles à la statistique classique",
     "Double enjeu : prédire ET expliquer (interprétabilité)"],
    title="Contexte data science & santé", size=15)

bullet_box(sl, Inches(6.6), Inches(2.4), Inches(6.3), Inches(2.4),
    ["Obésité : IMC > 30  (seuil OMS)",
     "Hypercholestérolémie : cholestérol > 200 mg/dL  (NCEP)",
     "Hypertension : pression systolique ≥ 140 mmHg  (ESC)"],
    title="Seuils médicaux de référence", size=15)

textbox(sl, Inches(0.4), Inches(5.05), W - Inches(0.8), Inches(0.35),
        "→  Variable cible : disease_risk (0/1)  |  Indicateur continu : cholesterol (mg/dL)",
        size=14, italic=True, color=RGBColor(80,80,80))

# ══════════════════════════════════════════════════════════════
# DIAPO 4 — JEU DE DONNÉES
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
title_bar(sl, "Jeu de données — Health & Lifestyle Dataset")

add_table(sl, Inches(0.4), Inches(1.15), Inches(5.5), Inches(3.2),
    ["Type", "Variables"],
    [["Continues (8)", "age, bmi, daily_steps, sleep_hours, cholesterol,\nsystolic_bp, diastolic_bp, resting_hr"],
     ["Binaires (4)", "smoker, alcohol, family_history, disease_risk"],
     ["Catégorielle (1)", "gender"],
     ["Cible régression", "cholesterol (mg/dL)"],
     ["Cible classification", "disease_risk (0 = sain / 1 = à risque)"]],
    font_size=13, col_widths=[Inches(1.8), Inches(3.7)])

bullet_box(sl, Inches(0.4), Inches(4.5), Inches(5.5), Inches(2.5),
    ["100 000 individus, 15 variables",
     "Dataset synthétique (Kaggle) — pas de contrainte RGPD",
     "Aucune valeur manquante",
     "Distributions quasi-uniformes (skewness < |0.02|)"],
    title="Points clés", size=14)

image(sl, FIGS+"disease_risk_distribution.png",
      Inches(6.3), Inches(1.1), w=Inches(6.6))
textbox(sl, Inches(6.3), Inches(5.05), Inches(6.6), Inches(0.35),
        "Déséquilibre 75% / 25% → modèle naïf = 75% accuracy sans rien apprendre",
        size=13, italic=True, color=RGBColor(80,80,80))

# ══════════════════════════════════════════════════════════════
# DIAPO 5 — PIPELINE CRISP-DM
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
title_bar(sl, "Pipeline CRISP-DM")

etapes = [
    ("EDA", "Corrélations < 0.01\nDéséquilibre 75/25\nPas d'outliers"),
    ("Feature\nEngineering", "Encodage gender\n3 variables métier\nData leakage détecté"),
    ("Régression", "6 modèles\nRandomizedSearchCV\nK-Fold k=5"),
    ("Classification", "3 modèles\nSMOTE vs class_weight\nStratified K-Fold"),
    ("Clustering", "K-Means k=3\nCAH Ward\nVisualisaton ACP"),
    ("Interprétabilité", "Coefficients LR/Ridge\nFI Gini RF/XGB\nSHAP"),
]

box_w = Inches(1.85)
box_h = Inches(1.5)
gap = Inches(0.22)
start_x = Inches(0.35)
y_box = Inches(1.3)
arrow_y = y_box + box_h / 2

for i, (titre, desc) in enumerate(etapes):
    x = start_x + i * (box_w + gap)
    rect(sl, x, y_box, box_w, Inches(0.42), BLEU)
    textbox(sl, x, y_box + Inches(0.06), box_w, Inches(0.3),
            titre, size=14, bold=True, color=BLANC, align=PP_ALIGN.CENTER)
    rect(sl, x, y_box + Inches(0.42), box_w, box_h - Inches(0.42), BLEU_FOND)
    textbox(sl, x + Inches(0.08), y_box + Inches(0.5),
            box_w - Inches(0.16), box_h - Inches(0.6),
            desc, size=11, color=NOIR)
    if i < len(etapes) - 1:
        ax = x + box_w + Inches(0.04)
        rect(sl, ax, arrow_y - Pt(3), gap - Inches(0.04), Pt(6), BLEU_ACC)

textbox(sl, Inches(0.35), Inches(3.1), W - Inches(0.7), Inches(0.35),
        "Approche itérative : chaque étape informe la suivante (EDA → leakage détecté → feature engineering révisé → modèles)",
        size=13, italic=True, color=RGBColor(80,80,80))

# Résultats synthèse
add_table(sl, Inches(0.35), Inches(3.6), W - Inches(0.7), Inches(2.8),
    ["Tâche", "Meilleur modèle", "R² / AUC", "Silhouette", "Conclusion"],
    [["Régression", "XGBoost tuné", "R² = 0.0001", "—", "Aucun signal — tous équivalents"],
     ["Classification", "Logistic Regression", "AUC = 0.499", "—", "Équivalent tirage aléatoire"],
     ["Clustering", "K-Means k=3", "—", "0.103", "Profils biologiques interprétables"]],
    font_size=13)

# ══════════════════════════════════════════════════════════════
# DIAPO 6 — EDA : DISTRIBUTIONS & CORRÉLATIONS
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
title_bar(sl, "Analyse exploratoire — EDA")

textbox(sl, Inches(0.4), Inches(1.1), Inches(12.5), Inches(0.3),
        "Distributions des variables continues", size=15, bold=True, color=BLEU)
image(sl, FIGS+"distributions_continues.png", Inches(0.4), Inches(1.4), w=Inches(7.8))

textbox(sl, Inches(8.4), Inches(1.1), Inches(4.5), Inches(0.3),
        "Corrélations avec la cible", size=15, bold=True, color=BLEU)
image(sl, FIGS+"correlations_target.png", Inches(8.4), Inches(1.4), w=Inches(4.5))

rect(sl, Inches(0.4), Inches(5.4), W - Inches(0.8), Inches(1.7), BLEU_FOND)
rect(sl, Inches(0.4), Inches(5.4), Inches(0.08), Inches(1.7), ROUGE_ACC)
textbox(sl, Inches(0.6), Inches(5.45), W - Inches(1.0), Inches(0.3),
        "Résultat central", size=13, bold=True, color=ROUGE_ACC)
textbox(sl, Inches(0.6), Inches(5.75), W - Inches(1.0), Inches(1.1),
        "Toutes les corrélations sont < |0.01|. La variable la plus corrélée à disease_risk "
        "est resting_hr avec r = 0.005 — statistiquement non significatif sur 100 000 individus. "
        "Conclusion : disease_risk a été généré indépendamment. Justification pour méthodes non-linéaires (RF, XGBoost).",
        size=14, color=NOIR)

# ══════════════════════════════════════════════════════════════
# DIAPO 7 — FEATURE ENGINEERING
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
title_bar(sl, "Feature Engineering — 16 variables explicatives")

bullet_box(sl, Inches(0.4), Inches(1.2), Inches(5.8), Inches(1.5),
    ["gender → gender_enc (0/1)",
     "Suppression de id (identifiant sans pouvoir prédictif)"],
    title="Encodage", size=14)

bullet_box(sl, Inches(0.4), Inches(2.85), Inches(5.8), Inches(2.5),
    ["hypertension : systolic_bp ≥ 140 ou diastolic_bp ≥ 90  (seuil ESC)",
     "bmi_cat : 0=sous-poids / 1=normal / 2=surpoids / 3=obèse  (seuil OMS)",
     "lifestyle_score : score composite (steps + sleep + water – smoker – alcohol – bmi – calories)"],
    title="3 variables métier créées (connaissances médicales)", size=14)

blue_box(sl, Inches(6.5), Inches(1.2), Inches(6.4), Inches(2.2),
    "⚠  Data Leakage détecté et corrigé",
    ["hypercholesterol = 1 si cholesterol > 200 mg/dL",
     "Directement dérivée de la variable cible → 'donner la réponse' au modèle",
     "R² artificiel = 0.675 avec leakage  →  R² ≈ 0 après suppression",
     "Exclue de la régression, conservée uniquement pour le clustering"])

bullet_box(sl, Inches(6.5), Inches(3.55), Inches(6.4), Inches(1.8),
    ["Standardisation z-score pour les modèles linéaires",
     "Normalisation min-max pour K-Means (distance euclidienne)"],
    title="Mise à l'échelle", size=14)

textbox(sl, Inches(0.4), Inches(5.5), W - Inches(0.8), Inches(0.35),
        "Dataset final : 15 variables → 16 features explicatives  (14 − 1 + 3 = 16)",
        size=14, bold=True, color=BLEU)

# ══════════════════════════════════════════════════════════════
# DIAPO 8 — RÉGRESSION : PROTOCOLE & MODÈLES
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
title_bar(sl, "Régression — Protocole & justification des modèles")

bullet_box(sl, Inches(0.4), Inches(1.15), Inches(5.5), Inches(1.6),
    ["Split 80/20  (random_state=42)",
     "Validation croisée K-Fold k=5  →  stabilité de l'estimation",
     "Tuning : RandomizedSearchCV n_iter=20  (plus rapide que GridSearch)",
     "Métriques : R², RMSE, MAE"],
    title="Protocole expérimental", size=14)

add_table(sl, Inches(0.4), Inches(2.9), Inches(5.6), Inches(3.55),
    ["Modèle", "Justification"],
    [["Régression linéaire", "Baseline — référence obligatoire"],
     ["Ridge (L2)", "Réduit les coefficients sans les annuler — stabilise si variables peu informatives"],
     ["Lasso (L1)", "Force des coefficients à 0 → sélection automatique de variables"],
     ["ElasticNet", "Combine L1 + L2 — robuste quand on ne sait pas lequel convient"],
     ["Random Forest", "Ensembliste non-linéaire, parallèle — capte interactions"],
     ["XGBoost", "Ensembliste séquentiel — corrige les erreurs itérativement"]],
    font_size=12, col_widths=[Inches(2.0), Inches(3.6)])

textbox(sl, Inches(6.3), Inches(1.15), Inches(6.6), Inches(0.35),
        "Logique : spectre de complexité croissante", size=15, bold=True, color=BLEU)

bullet_box(sl, Inches(6.3), Inches(1.6), Inches(6.6), Inches(1.6),
    ["Baseline linéaire → plancher de performance",
     "Régularisées → teste si contraindre les coeff. aide",
     "Ensemblistes → capture les non-linéarités potentielles",
     "Si complexe ≤ baseline → complexité non justifiée (Occam)"],
    size=14)

image(sl, FIGS+"distributions_continues.png",
      Inches(6.3), Inches(3.3), w=Inches(6.6))

# ══════════════════════════════════════════════════════════════
# DIAPO 9 — RÉGRESSION : RÉSULTATS
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
title_bar(sl, "Régression — Résultats")

add_table(sl, Inches(0.4), Inches(1.15), Inches(7.5), Inches(2.85),
    ["Modèle", "R²", "RMSE (mg/dL)", "MAE (mg/dL)", "CV-R²"],
    [["Linéaire (baseline)",  "0.0001",  "43.33", "37.59", "≈ 0"],
     ["Ridge (λ=7197, tuné)", "0.0001",  "43.33", "37.59", "≈ 0"],
     ["Lasso (λ=9.54, tuné)", "−0.000",  "43.33", "37.60", "≈ 0"],
     ["ElasticNet",           "−0.000",  "43.33", "37.60", "≈ 0"],
     ["Random Forest (tuné)", "−0.005",  "43.44", "37.67", "≈ 0"],
     ["XGBoost (tuné)",       "0.0001",  "43.33", "37.59", "≈ 0"]],
    font_size=13, col_widths=[Inches(2.4), Inches(1.0), Inches(1.5), Inches(1.5), Inches(1.0)])

bullet_box(sl, Inches(0.4), Inches(4.2), Inches(7.5), Inches(2.8),
    ["Lasso (tuné) : 16/16 coefficients = 0 → prédit uniquement la moyenne (224.3 mg/dL)",
     "Random Forest non-tuné : R² < 0 → pire que prédire la moyenne (surajustement du bruit)",
     "Data leakage corrigé : R² artificiel 0.675 → ≈ 0 après suppression de hypercholesterol",
     "Aucun modèle retenu : différences entre −0.005 et 0.0001 sont statistiquement négligeables",
     "Cause fondamentale : variables générées indépendamment — signal absent par construction"],
    title="Analyse critique", size=13)

image(sl, FIGS+"correlation_matrix.png",
      Inches(8.1), Inches(1.15), w=Inches(4.8))
textbox(sl, Inches(8.1), Inches(5.0), Inches(4.8), Inches(0.35),
        "Matrice de corrélation — toutes valeurs < |0.01|",
        size=12, italic=True, color=RGBColor(80,80,80), align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════
# DIAPO 10 — CLASSIFICATION : RÉSULTATS
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
title_bar(sl, "Classification — Résultats")

add_table(sl, Inches(0.4), Inches(1.15), Inches(6.5), Inches(2.2),
    ["Modèle", "Accuracy", "Précision", "Rappel", "F1", "AUC-ROC"],
    [["Naïf (tout à 0)", "0.752", "n/a",   "0.000", "0.000", "0.500"],
     ["Logistic Regression", "0.498", "0.249", "0.509", "0.335", "0.499"],
     ["Random Forest",       "0.568", "0.245", "0.355", "0.290", "0.494"],
     ["XGBoost",             "0.525", "0.246", "0.441", "0.316", "0.499"]],
    font_size=13, col_widths=[Inches(2.1), Inches(0.85), Inches(0.85), Inches(0.85), Inches(0.7), Inches(0.9)])

bullet_box(sl, Inches(0.4), Inches(3.5), Inches(6.5), Inches(2.65),
    ["AUC ≈ 0.5 pour tous = équivalent à un tirage aléatoire (pile ou face)",
     "class_weight='balanced' : rappel monte mais accuracy < 50% → sur-correction",
     "SMOTE : même résultat — ne peut pas créer un signal absent",
     "Accuracy seule trompeuse sur données déséquilibrées (75/25) → F1 + AUC obligatoires",
     "Cohérent avec régression : disease_risk généré indépendamment de toutes les variables"],
    title="Analyse critique", size=13)

image(sl, FIGS+"roc_curves.png", Inches(6.9), Inches(1.05), w=Inches(3.2))
textbox(sl, Inches(6.9), Inches(4.15), Inches(3.2), Inches(0.3),
        "Courbes ROC — AUC ≈ 0.5",
        size=12, italic=True, color=RGBColor(80,80,80), align=PP_ALIGN.CENTER)

image(sl, FIGS+"confusion_matrices.png", Inches(10.1), Inches(1.05), w=Inches(2.9))
textbox(sl, Inches(10.1), Inches(4.15), Inches(2.9), Inches(0.3),
        "Matrices de confusion",
        size=12, italic=True, color=RGBColor(80,80,80), align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════
# DIAPO 11 — CLUSTERING
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
title_bar(sl, "Clustering — K-Means & CAH")

textbox(sl, Inches(0.4), Inches(1.1), Inches(5.5), Inches(0.3),
        "K-Means  (k=3, 100 000 individus)", size=15, bold=True, color=BLEU)

add_table(sl, Inches(0.4), Inches(1.45), Inches(5.5), Inches(1.55),
    ["Cluster", "Profil", "IMC moy.", "Syst. (mmHg)"],
    [["0", "Obèse hypertendu",     "35.0", "142"],
     ["1", "Surpoids normotendu",  "29.1", "115"],
     ["2", "Normal hypertendu",    "24.0", "142"]],
    font_size=13, col_widths=[Inches(0.75), Inches(2.15), Inches(1.25), Inches(1.35)])

textbox(sl, Inches(0.4), Inches(3.1), Inches(5.5), Inches(0.35),
        "Partition selon IMC + pression systolique (variables à plus grande variance)",
        size=13, italic=True, color=RGBColor(80,80,80))

textbox(sl, Inches(0.4), Inches(3.55), Inches(5.5), Inches(0.3),
        "CAH Ward  (n=5 000 — complexité O(n²))", size=15, bold=True, color=BLEU)

add_table(sl, Inches(0.4), Inches(3.9), Inches(5.5), Inches(1.55),
    ["Cluster", "Profil", "% Fumeurs", "Syst."],
    [["0", "Non-fumeurs, HTA",     "2%",    "142"],
     ["1", "Fumeurs modérés",      "18%",   "115"],
     ["2", "Fumeurs intenses, HTA","100%",  "142"]],
    font_size=13, col_widths=[Inches(0.75), Inches(2.15), Inches(1.25), Inches(1.35)])

textbox(sl, Inches(0.4), Inches(5.55), Inches(5.5), Inches(0.35),
        "CAH isole le groupe fumeurs intenses — K-Means les aurait dilués",
        size=13, italic=True, color=RGBColor(80,80,80))

image(sl, FIGS+"kmeans_selection_k.png", Inches(6.0), Inches(1.05), w=Inches(3.6))
textbox(sl, Inches(6.0), Inches(4.15), Inches(3.6), Inches(0.3),
        "Sélection k — k=3 maximise la silhouette",
        size=12, italic=True, color=RGBColor(80,80,80), align=PP_ALIGN.CENTER)

image(sl, FIGS+"kmeans_pca.png", Inches(9.8), Inches(1.05), w=Inches(3.3))
textbox(sl, Inches(9.8), Inches(4.15), Inches(3.3), Inches(0.35),
        "Clusters K-Means — ACP 2D (24.3% variance)",
        size=12, italic=True, color=RGBColor(80,80,80), align=PP_ALIGN.CENTER)

add_table(sl, Inches(6.0), Inches(4.6), Inches(7.0), Inches(1.55),
    ["Méthode", "Silhouette", "Davies-Bouldin", "Note"],
    [["K-Means", "0.103", "2.655", "Faible — attendu sur données sans structure causale"],
     ["CAH", "0.105", "2.721", "ACP 2D = 24.3% var. (15 dims quasi-indépendantes)"]],
    font_size=12, col_widths=[Inches(1.3), Inches(1.1), Inches(1.6), Inches(3.0)])

# ══════════════════════════════════════════════════════════════
# DIAPO 12 — INTERPRÉTABILITÉ
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
title_bar(sl, "Interprétabilité — 5 méthodes comparées")

textbox(sl, Inches(0.4), Inches(1.1), Inches(12.5), Inches(0.35),
        "Méthodes : coefficients LR & Ridge  |  Feature Importance Gini (RF & XGBoost)  |  SHAP (XGBoost)",
        size=14, color=NOIR)

add_table(sl, Inches(0.4), Inches(1.55), Inches(5.5), Inches(2.0),
    ["Variable", "LR", "RF", "XGB", "SHAP", "Ridge", "Score"],
    [["systolic_bp",    "✓", "✓", "✓", "✓", "✓", "5/5"],
     ["daily_steps",    "✓", "✓", "·", "✓", "·", "3/5"],
     ["family_history", "✓", "·", "✓", "·", "✓", "3/5"],
     ["bmi",            "·", "✓", "·", "✓", "·", "2/5"]],
    font_size=13, col_widths=[Inches(1.7), Inches(0.55), Inches(0.55), Inches(0.55), Inches(0.7), Inches(0.7), Inches(0.7)])

blue_box(sl, Inches(0.4), Inches(3.7), Inches(5.5), Inches(2.5),
    "Analyse critique",
    ["Les 5 méthodes ne convergent pas → mesurent du bruit aléatoire",
     "systolic_bp seule présente partout : grande variance, pas signal causal",
     "Divergence = résultat analytique en soi (Rudin 2019)",
     "Sur données réelles : cholestérol, tabagisme, IMC domineraient clairement"])

image(sl, FIGS+"rf_feature_importance.png", Inches(6.1), Inches(1.1), w=Inches(3.5))
textbox(sl, Inches(6.1), Inches(4.35), Inches(3.5), Inches(0.3),
        "Feature importance (Gini) — Random Forest",
        size=12, italic=True, color=RGBColor(80,80,80), align=PP_ALIGN.CENTER)

image(sl, FIGS+"shap_summary.png", Inches(9.7), Inches(1.1), w=Inches(3.4))
textbox(sl, Inches(9.7), Inches(4.35), Inches(3.4), Inches(0.3),
        "SHAP summary plot — XGBoost",
        size=12, italic=True, color=RGBColor(80,80,80), align=PP_ALIGN.CENTER)

image(sl, FIGS+"lr_coefficients.png", Inches(6.1), Inches(4.7), w=Inches(7.0))
textbox(sl, Inches(6.1), Inches(6.35), Inches(7.0), Inches(0.3),
        "Coefficients Logistic Regression (variables standardisées)",
        size=12, italic=True, color=RGBColor(80,80,80), align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════
# DIAPO 13 — BILAN
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
title_bar(sl, "Bilan — Résultat central")

rect(sl, Inches(0.4), Inches(1.15), W - Inches(0.8), Inches(1.0), BLEU_FOND)
rect(sl, Inches(0.4), Inches(1.15), Inches(0.08), Inches(1.0), ROUGE_ACC)
textbox(sl, Inches(0.6), Inches(1.2), W - Inches(1.0), Inches(0.3),
        "Limite fondamentale du dataset synthétique", size=14, bold=True, color=ROUGE_ACC)
textbox(sl, Inches(0.6), Inches(1.5), W - Inches(1.0), Inches(0.55),
        "Les variables ont été générées indépendamment, sans relation causale encodée. "
        "Aucune technique ne peut extraire un signal qui n'existe pas.",
        size=14, color=NOIR)

bullet_box(sl, Inches(0.4), Inches(2.35), Inches(5.8), Inches(2.5),
    ["Corrélations < 0.01 en EDA le prédisait",
     "Lasso : 16/16 coefficients = 0",
     "R² ≈ 0 et AUC ≈ 0.5 pour tous les modèles",
     "Data leakage détecté (R² 0.675 → 0)",
     "Interprétabilité divergente (5 méthodes)"],
    title="Preuves convergentes", size=14, title_color=ROUGE_ACC)

bullet_box(sl, Inches(6.6), Inches(2.35), Inches(6.3), Inches(2.5),
    ["Pipeline complet et rigoureux (CRISP-DM)",
     "Bonne pratique détection data leakage",
     "Validation croisée stratifiée systématique",
     "Deux stratégies d'imbalance comparées",
     "Directement transposable à des données réelles"],
    title="Ce n'est pas un échec — valeur méthodologique", size=14, title_color=BLEU)

rect(sl, Inches(0.4), Inches(5.1), W - Inches(0.8), Inches(1.1), BLEU)
textbox(sl, Inches(0.6), Inches(5.2), W - Inches(1.0), Inches(0.8),
        "« Savoir diagnostiquer pourquoi un modèle échoue est aussi précieux "
        "que d'en construire un qui réussit. »",
        size=16, bold=True, color=BLANC, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════
# DIAPO 14 — PERSPECTIVES
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
title_bar(sl, "Perspectives")

persp = [
    ("1", "Changer de dataset",
     "Cohortes Framingham ou NHANES — vraies relations biologiques, même pipeline"),
    ("2", "Améliorer la classification",
     "Stacking, réseaux de neurones (Deep Learning, M2) — SMOTE déjà intégré"),
    ("3", "Déploiement",
     "API REST (FastAPI) : profil individuel → prédiction + valeurs SHAP"),
    ("4", "Interprétabilité avancée",
     "Partial Dependence Plots (PDP), ICE — limites de SHAP agrégé"),
    ("5", "Data drift",
     "Suivi AUC sur nouvelles cohortes, réentraînement périodique"),
    ("6", "Master 2",
     "Causalité, fairness algorithmique, apprentissage fédéré"),
]

for i, (num, titre, desc) in enumerate(persp):
    col = 0 if i < 3 else 1
    row = i % 3
    x = Inches(0.4) + col * Inches(6.5)
    y = Inches(1.25) + row * Inches(1.75)
    rect(sl, x, y, Inches(6.1), Inches(1.55), GRIS_FOND)
    rect(sl, x, y, Inches(0.42), Inches(1.55), BLEU)
    textbox(sl, x, y + Inches(0.5), Inches(0.42), Inches(0.5),
            num, size=18, bold=True, color=BLANC, align=PP_ALIGN.CENTER)
    textbox(sl, x + Inches(0.5), y + Inches(0.1), Inches(5.5), Inches(0.38),
            titre, size=15, bold=True, color=BLEU)
    textbox(sl, x + Inches(0.5), y + Inches(0.5), Inches(5.5), Inches(0.9),
            desc, size=13, color=NOIR)

# ══════════════════════════════════════════════════════════════
# DIAPO 15 — QUESTIONS
# ══════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, BLANC)
rect(sl, 0, Inches(2.8), W, Inches(2.0), BLEU)
textbox(sl, Inches(0.6), Inches(3.0), W - Inches(1.2), Inches(0.9),
        "Merci pour votre attention", size=36, bold=True,
        color=BLANC, align=PP_ALIGN.CENTER)
textbox(sl, Inches(0.6), Inches(3.95), W - Inches(1.2), Inches(0.6),
        "Questions ?", size=28, color=RGBColor(200,220,255), align=PP_ALIGN.CENTER)

textbox(sl, Inches(0.6), Inches(5.4), W - Inches(1.2), Inches(0.35),
        "Ardacham Mahamat Teguene  —  M1 MIASHS  —  Université de Montpellier Paul Valéry  —  Juin 2026",
        size=14, color=RGBColor(120,120,120), align=PP_ALIGN.CENTER)

# ── Sauvegarde ───────────────────────────────────────────────
OUT = "/Users/ardachammahamatteguene/Desktop/Bureau - MacBook Pro de Ardacham/Memoire/presentation.pptx"
prs.save(OUT)
print(f"Saved: {OUT}")
