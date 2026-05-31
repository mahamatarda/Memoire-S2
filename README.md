# Modélisation des comportements de santé par apprentissage automatique

**Mémoire de fin de première année de Master**
**ARDACHAM Mahamat Teguene — Master 1 MIASHS — Université de Montpellier Paul Valéry**
**Année universitaire 2025-2026**
**Tuteurs universitaires : Sophie Lèbre & Jérôme Pasquet**

---

## Résumé

Projet de data science appliqué à la modélisation des comportements de santé par apprentissage automatique, s'appuyant sur le *Health & Lifestyle Dataset* — un jeu de données synthétique de 100 000 individus décrits par 15 variables (habitudes de vie + indicateurs biologiques).

Pipeline complet conduit selon la méthodologie **CRISP-DM** :
- **EDA** : corrélations < 0.01 entre toutes les variables, absence de signal linéaire
- **Feature engineering** : encodage, création de variables métier (score de style de vie composite, hypertension, hypercholestérol), standardisation
- **Régression** : 6 modèles comparés (Linéaire, Ridge, Lasso, ElasticNet, Random Forest, XGBoost) — tuning par `RandomizedSearchCV`, validation croisée K-Fold (k=5)
- **Classification** : 3 modèles (LR, RF, XGBoost) + comparaison `class_weight='balanced'` vs SMOTE
- **Clustering** : K-Means (k=3) + CAH liaison Ward sur sous-échantillon 5 000 individus + visualisation ACP
- **Interprétabilité** : coefficients LR/Ridge, feature importances Gini, SHAP (complément visuel)

**Résultat central** : le dataset synthétique ne contient aucun signal — variables générées indépendamment. Tous les modèles atteignent R² ≈ 0 et AUC ≈ 0.5. La valeur du projet réside dans la rigueur méthodologique du pipeline, directement transposable à des données épidémiologiques réelles (Framingham, NHANES).

---

## Structure du dépôt

```
Memoire/
├── notebooks/
│   ├── 01_EDA.ipynb                  # Analyse exploratoire
│   ├── 02_feature_engineering.ipynb  # Encodage, variables métier, standardisation
│   ├── 03_regression.ipynb           # 6 modèles de régression + tuning
│   ├── 04_classification.ipynb       # 3 modèles + class_weight + SMOTE
│   ├── 05_clustering.ipynb           # K-Means + CAH + ACP
│   └── 06_interpretabilite.ipynb     # Coefficients + FI Gini + SHAP
├── latex/
│   ├── manuscript.tex                # Document principal
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
│   └── logos/
│       └── upv_new.png
├── figures/                          # Graphiques exportés des notebooks
├── health_lifestyle_dataset.csv      # Dataset brut (100 000 lignes)
└── data_processed.csv                # Dataset après feature engineering
```

---

## Stack technique

| Catégorie | Outil |
|---|---|
| Langage | Python 3 |
| Manipulation données | pandas, numpy |
| Modélisation | scikit-learn, xgboost |
| Déséquilibre classes | imbalanced-learn (SMOTE) |
| Interprétabilité | shap |
| Réduction de dim. | scikit-learn (PCA) |
| Versionnement | Git / GitHub |
| Rédaction | LaTeX |

---

## Points méthodologiques clés

- **Data leakage détecté et corrigé** : `hypercholesterol` = f(`cholesterol`) → R² artificiel de 0.675, corrigé à ≈ 0 après suppression
- **SVM exclu** : complexité O(n²) incompatible avec 100 000 observations (≈ 10 milliards d'opérations)
- **Lasso** : 16/16 coefficients réduits à 0 — confirme l'absence totale de signal
- **SMOTE vs class_weight** : les deux stratégies confirment qu'aucune ne peut extraire un signal absent
- **Clustering** : silhouettes ≈ 0.10 (faibles mais attendues sur données synthétiques sans structure causale)
- **ACP** : 24.3% de variance expliquée en 2D sur 15 dimensions quasi-indépendantes

---

## Mots-clés

data science · santé · machine learning · régression · classification · clustering · interprétabilité · SHAP · SMOTE · données synthétiques · data leakage · CRISP-DM
