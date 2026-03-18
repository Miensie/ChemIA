# ChemAI — Add-in Excel de Chimiométrie Multivariée

> **PCA · PLS · Clustering · Anomalies · IA · Rapport**  
> Analyse multivariée professionnelle directement dans Microsoft Excel.

🌐 **Demo live** : [https://VOTRE-USERNAME.github.io/chemai-addin](https://VOTRE-USERNAME.github.io/chemai-addin)

---

## Déploiement sur GitHub Pages (guide complet)

### Étape 1 — Créer le dépôt GitHub

1. Aller sur [github.com/new](https://github.com/new)
2. Nom du dépôt : `chemai-addin`
3. Visibilité : **Public** (requis pour GitHub Pages gratuit)
4. Cliquer **Create repository**

---

### Étape 2 — Pousser les fichiers

```bash
# Dans le dossier chemai-github-pages/ (ce dossier)
git init
git add .
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/Miensie/ChemIA.git
git push -u origin main
```

---

### Étape 3 — Activer GitHub Pages

1. Aller sur votre dépôt GitHub → **Settings**
2. Section **Pages** (menu gauche)
3. Source : `Deploy from a branch`
4. Branch : `main` / `/ (root)`
5. Cliquer **Save**

⏳ Attendre 1-2 minutes → votre site est disponible sur :  
`https://VOTRE-USERNAME.github.io/chemai-addin`

---

### Étape 4 — Personnaliser le manifest

Éditez `addin/manifest-ghpages.xml` : remplacez **toutes** les occurrences de
`VOTRE-USERNAME` par votre vrai nom d'utilisateur GitHub.

```xml
<!-- Exemple avec username "labochimie" -->
<IconUrl DefaultValue="https://labochimie.github.io/chemai-addin/assets/icon-32.png"/>
<SourceLocation DefaultValue="https://labochimie.github.io/chemai-addin/addin/taskpane.html"/>
```

Puis re-pushez :
```bash
git add addin/manifest-ghpages.xml
git commit -m "fix: update manifest with real GitHub username"
git push
```

---

### Étape 5 — Charger l'Add-in dans Excel

#### Excel Online (le plus simple)
1. Ouvrir [excel.office.com](https://excel.office.com)
2. **Insérer** → **Compléments Office** → **Charger mon complément**
3. Sélectionner `addin/manifest-ghpages.xml`

#### Excel Desktop (Windows / macOS)
1. **Insérer** → **Compléments** → **Mes compléments**
2. **Télécharger un complément** → sélectionner `addin/manifest-ghpages.xml`
3. Le bouton **ChemAI Analyse** apparaît dans l'onglet Accueil

---

## Structure du dépôt

```
chemai-addin/
│
├── index.html                   ← Page d'accueil GitHub Pages
├── README.md                    ← Ce fichier
│
├── addin/
│   ├── taskpane.html            ← Interface principale de l'Add-in
│   ├── taskpane-standalone.js   ← Orchestrateur UI
│   └── manifest-ghpages.xml    ← Manifest Office (à personnaliser)
│
├── scripts/
│   ├── chemMath.js              ← Moteur de calcul JS pur (PCA, PLS, K-Means…)
│   ├── chemCharts.js            ← Graphiques SVG scientifiques
│   ├── chemExcel.js             ← Interactions Office.js / Excel
│   ├── chemAI.js                ← Client Gemini API
│   ├── chemBackend.js           ← Client HTTP backend Python (optionnel)
│   └── chemReport.js            ← Générateur de rapport HTML
│
└── styles/
    └── taskpane.css             ← Thème laboratoire sombre
```

---

## Fonctionnement sur GitHub Pages

### Mode autonome (JS pur) ✅
Tous les calculs s'exécutent **dans le navigateur** — aucun serveur requis :

| Module | Algorithme | JS pur |
|--------|-----------|--------|
| PCA | SVD par déflation itérative | ✅ |
| PLS-1 | NIPALS | ✅ |
| K-Means | K-Means++ | ✅ |
| Hiérarchique | Ward, Complete, Average | ✅ |
| Mahalanobis | Distance Chi² | ✅ |
| T² Hotelling | F-distribution | ✅ |
| Q résiduel | Jackson-Mudholkar | ✅ |

### Mode backend Python (optionnel) 🔧
Pour activer Isolation Forest, LOF, SVM, Random Forest :

```bash
cd backend/
pip install -r requirements.txt
python main.py
# → http://localhost:8000
```

Puis dans l'Add-in, la section backend se connecte automatiquement à `localhost:8000`.

---

## Configuration de l'IA (Gemini)

1. Obtenir une clé gratuite sur [aistudio.google.com](https://aistudio.google.com)
2. Dans l'Add-in → panneau **Rapport** → coller la clé API
3. Cliquer 💾 pour sauvegarder en session

L'IA génère des interprétations scientifiques en français pour chaque analyse.

---

## Données de test

Copiez ce tableau dans Excel pour tester :

```
Sample	V1	V2	V3	V4	V5	Y
S001	2.34	1.12	4.56	0.78	3.21	12.4
S002	2.67	1.34	4.23	0.92	3.45	13.1
S003	1.89	0.98	5.12	0.65	2.87	11.8
S004	3.12	1.56	3.89	1.12	3.78	14.2
S005	2.45	1.23	4.67	0.83	3.12	12.7
S006	8.91	0.45	4.12	0.77	3.34	13.0
S007	2.78	1.45	4.34	0.95	3.56	13.5
S008	2.23	1.08	4.89	0.71	2.98	12.1
S009	2.56	1.28	4.45	0.87	3.23	12.9
S010	3.34	1.67	3.67	1.18	3.89	14.6
S011	2.12	1.02	4.78	0.69	2.91	11.9
S012	2.89	1.48	4.23	0.98	3.67	13.7
```

`S006` est un outlier intentionnel (V1 = 8.91).

---

## Références scientifiques

- **Wold, S. et al.** (2001). PLS-regression. *Chemometrics and ILS*, 58(2), 109–130.
- **Jolliffe, I.T.** (2002). *Principal Component Analysis*, 2nd ed. Springer.
- **Jackson, J.E., Mudholkar, G.S.** (1979). Control procedures for residuals. *Technometrics*, 21(3).
- **Ward, J.H.** (1963). Hierarchical grouping. *JASA*, 58(301), 236–244.
- **Arthur, D., Vassilvitskii, S.** (2007). k-means++. *SODA 2007*.

---

*ChemAI v2.0 — Open Source — Office.js + JS pur + Python FastAPI (optionnel)*
