# Modélisation des comportements de santé par apprentissage automatique

**Mémoire de fin de première année de Master**
**ARDACHAM Mahamat Teguene — Master 1 MIASHS — Université de Montpellier Paul Valéry**
**Année universitaire 2025-2026**
**Tuteurs universitaires : Sophie Lèbre & Jérôme Pasquet**

---

## Résumé

Projet de data science appliqué à la modélisation des comportements de santé par apprentissage automatique, s'appuyant sur le [*Health & Lifestyle Dataset*](https://www.kaggle.com/datasets/chik0di/health-and-lifestyle-dataset) — un jeu de données synthétique de 100 000 individus décrits par 15 variables (habitudes de vie + indicateurs biologiques).

Pipeline complet conduit selon la méthodologie **CRISP-DM** :

- **EDA** : corrélations < 0.01 entre toutes les variables, absence de signal linéaire
- **Feature engineering** : encodage, création de variables métier (score de style de vie composite, hypertension, catégories IMC), standardisation — 18 variables au total (15 originales + 3 nouvelles : hypertension, bmi_cat, lifestyle_score)
- **Régression** : 6 modèles comparés (Linéaire, Ridge, Lasso, ElasticNet, Random Forest, XGBoost) — tuning par `RandomizedSearchCV` (n_iter=20), validation croisée K-Fold (k=5)
- **Classification** : 3 modèles (LR, RF, XGBoost) + comparaison `class_weight='balanced'` vs SMOTE
- **Clustering** : K-Means (k=3) + CAH liaison Ward sur sous-échantillon 5 000 individus + visualisation ACP (24.3% de variance en 2D)
- **Interprétabilité** : coefficients LR/Ridge, feature importances Gini (RF, XGBoost), SHAP (complément visuel)

**Résultat central** : le dataset synthétique ne contient aucun signal — variables générées indépendamment. Tous les modèles atteignent R² ≈ 0 et AUC ≈ 0.5. La valeur du projet réside dans la rigueur méthodologique du pipeline, directement transposable à des données épidémiologiques réelles (cohortes Framingham, NHANES).

---

## Rapport final

Le PDF compilé est disponible dans ce dépôt : [`latex/Memoire_M1_ARDACHAM_Mahamat_Teguene_2026.pdf`](latex/Memoire_M1_ARDACHAM_Mahamat_Teguene_2026.pdf)

## Support de soutenance

- [`Support Mahamat - Memoire.pdf`](Support%20Mahamat%20-%20Memoire.pdf) — Présentation 28 slides (PDF)
- [`Support Mahamat - Memoire.pptx`](Support%20Mahamat%20-%20Memoire.pptx) — Présentation modifiable (PPTX)
- [`script_oral_complet.txt`](script_oral_complet.txt) — Script oral intégral : Partie A (15 min slide par slide) + Partie B (36 Q&R thématiques)

---

## Structure du dépôt

```text
Memoire/
├── Support Mahamat - Memoire.pdf     # Présentation soutenance (28 slides)
├── Support Mahamat - Memoire.pptx    # Présentation modifiable
├── script_oral_complet.txt           # Script oral + 36 Q&R
├── notebooks/
│   ├── 01_EDA.ipynb                  # Analyse exploratoire
│   ├── 02_feature_engineering.ipynb  # Encodage, variables métier, standardisation
│   ├── 03_regression.ipynb           # 6 modèles de régression + tuning
│   ├── 04_classification.ipynb       # 3 modèles + class_weight + SMOTE
│   ├── 05_clustering.ipynb           # K-Means + CAH + ACP
│   └── 06_interpretabilite.ipynb     # Coefficients + FI Gini + SHAP
├── latex/
│   ├── Memoire_M1_ARDACHAM_Mahamat_Teguene_2026.tex   # Document principal
│   ├── Memoire_M1_ARDACHAM_Mahamat_Teguene_2026.pdf   # PDF compilé (tectonic)
│   ├── references.bib                # 25 sources bibliographiques
│   ├── front_page.tex
│   ├── chapters/
│   │   ├── abstract.tex              # Résumé
│   │   ├── introduction.tex          # Introduction + Contexte (data science + métier)
│   │   ├── chapter_1.tex             # État de l'art (ML santé, SMOTE, SHAP)
│   │   ├── chapter_2.tex             # EDA + Feature engineering
│   │   ├── chapter_3.tex             # Régression + Classification
│   │   ├── chapter_4.tex             # Clustering + Interprétabilité
│   │   ├── chapter_5.tex             # Méthodologie, retour d'expérience, perspectives
│   │   ├── conclusion.tex
│   │   ├── abbreviations.tex
│   │   ├── lexique.tex
│   │   └── annexes.tex
│   ├── figures/                      # Graphiques intégrés dans le LaTeX
│   └── logos/
│       └── upv_new.png
├── figures/                          # Graphiques exportés des notebooks
├── health_lifestyle_dataset.csv      # Dataset brut (100 000 lignes, ignoré par git) — source : Kaggle (lien ci-dessus)
└── data_processed.csv                # Dataset après feature engineering (ignoré par git)
```

---

## Stack technique

| Catégorie | Outil |
| --- | --- |
| Langage | Python 3 |
| Manipulation données | pandas, numpy |
| Modélisation | scikit-learn, xgboost |
| Déséquilibre classes | imbalanced-learn (SMOTE) |
| Interprétabilité | shap |
| Réduction de dim. | scikit-learn (PCA) |
| Versionnement | Git / GitHub |
| Rédaction | LaTeX — `Memoire_M1_ARDACHAM_Mahamat_Teguene_2026.tex` (compilé avec tectonic) |

---

## Points méthodologiques clés

- **Data leakage détecté et corrigé** : `hypercholesterol` = f(`cholesterol`) → R² artificiel de 0.675, corrigé à ≈ 0 après suppression (variable temporaire dans 03_regression.ipynb uniquement)
- **Plus-value** : pipeline de diagnostic — 5 preuves convergentes d'absence de signal (corrélations, R²≈0, AUC≈0.5, silhouette, interprétabilité divergente), détection du data leakage, accuracy trompeuse sur déséquilibre 75/25
- **SVM exclu** : complexité O(n²) incompatible avec 100 000 observations (≈ 10 milliards d'opérations)
- **Lasso** : 16/16 coefficients réduits à 0 — confirme l'absence totale de signal linéaire
- **SMOTE vs class_weight** : les deux stratégies confirment qu'aucune ne peut extraire un signal absent
- **Clustering** : silhouettes ≈ 0.10 (faibles mais attendues sur données synthétiques sans structure causale)
- **ACP** : 24.3% de variance expliquée en 2D — forme rectangulaire due aux variables discrètes (`bmi_cat` sur PC1, `hypertension` binaire sur PC2)
- **Interprétabilité** : divergence entre les 5 méthodes — elle-même un résultat analytique sur données sans signal
- **Corrélations non-linéaires** : Pearson ≤ 0.01 ; RF/XGBoost (non-linéaires) échouent aussi → absence de signal confirmée même au-delà du linéaire

---

## Résultats chiffrés

| Tâche | Meilleur modèle | Métrique principale |
| --- | --- | --- |
| Régression (cholestérol) | XGBoost tuné | R² = 0.0001 |
| Classification (disease_risk) | Logistic Regression | AUC = 0.499 |
| Clustering | K-Means k=3 | Silhouette = 0.103 |

---

## Note sur l'utilisation de l'IA

Dans le cadre de ce mémoire, j'ai utilisé des outils d'intelligence artificielle générative (principalement **Claude** d'Anthropic) comme aide à la recherche, à la structuration du document, à la relecture et au débogage du code Python et LaTeX.

Tout le contenu scientifique — choix des modèles, interprétation des résultats, analyse critique, conclusions — a été compris, vérifié et validé par l'auteur. L'IA a servi d'outil d'assistance, pas de substitut au raisonnement analytique.

---

## Mots-clés

data science · santé · machine learning · régression · classification · clustering · interprétabilité · SHAP · SMOTE · données synthétiques · data leakage · CRISP-DM
