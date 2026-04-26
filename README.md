# 📊 Liasse Fiscale Tunisienne — Convertisseur Excel vers XML

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Python](https://img.shields.io/badge/Python-3.8+-green)
![Flask](https://img.shields.io/badge/Flask-2.3.3-red)
![License](https://img.shields.io/badge/License-MIT-yellow)

Application professionnelle pour **comptables et experts-comptables**, conforme aux normes fiscales tunisiennes 🇹🇳

---

## Table des matières

- [Aperçu](#aperçu)
- [Fonctionnalités](#fonctionnalités)
- [Formulaires supportés](#formulaires-supportés)
- [Structure du fichier Excel](#structure-du-fichier-excel)
- [Installation](#installation)
- [Utilisation](#utilisation)
- [Dépannage](#dépannage)
- [Structure du projet](#structure-du-projet)
- [Technologies utilisées](#technologies-utilisées)
- [Licence](#licence)

---

## Aperçu

Cette application permet aux professionnels de la comptabilité de **convertir automatiquement** leurs fichiers Excel au format XML, conformément aux exigences de la liasse fiscale tunisienne.

Plus besoin de ressaisir manuellement les données. Glissez votre fichier Excel, cliquez sur **Convertir**, et obtenez vos fichiers XML prêts à être intégrés.

---

## Fonctionnalités

- 🚀 **Conversion en 1 clic** — Transformez vos fichiers Excel en XML instantanément
- 📑 **6 formulaires fiscaux** — Support complet de F6001, F6002, F6003, F6004, F6005 et F6007
- 🔄 **Deux modèles F6004** — Choix entre le modèle Autorisé (synthétique) et le modèle Référence (détaillé)
- 🎯 **Détection intelligente** — Reconnaissance automatique des feuilles et des codes
- 📝 **Extraction automatique** — Récupération des informations du déclarant (matricule, raison sociale, etc.)
- 🖱️ **Interface glisser-déposer** — Simple et intuitive
- 📦 **Export ZIP** — Tous les fichiers XML regroupés dans une seule archive
- 💻 **Application portable** — Version `.exe` disponible, sans installation Python requise
- 🌐 **100 % hors ligne** — Aucune connexion internet nécessaire

---

## Formulaires supportés

| Formulaire | Description |
|------------|-------------|
| **F6001** | Bilan — Actif |
| **F6002** | Bilan — Passif |
| **F6003** | État de résultat |
| **F6004** | Flux de trésorerie (modèle Autorisé ou Référence) |
| **F6005** | Résultat fiscal |
| **F6007** | Autres informations |

---

## Structure du fichier Excel

### Feuilles requises

| Nom de la feuille | Formulaire | Description | Obligatoire |
|-------------------|------------|-------------|:-----------:|
| `entet` | — | Informations du déclarant | ✅ |
| `actif` | F6001 | Bilan — Actif | ✅ |
| `passif` | F6002 | Bilan — Passif | ✅ |
| `resultat` | F6003 | État de résultat | ✅ |
| `E flux auto` | F6004 | Flux — Modèle Autorisé | ⚠️ Un des deux |
| `et flux ref` | F6004 | Flux — Modèle Référence | ⚠️ Un des deux |
| `rsultat fiscal` | F6005 | Résultat fiscal | ✅ |
| `informations` | F6007 | Autres informations | ❌ Optionnel |

> **Note :** Si les deux feuilles F6004 sont présentes, l'application vous demandera de choisir le modèle à utiliser.

### Format des colonnes

Chaque feuille de données doit contenir **exactement 3 colonnes** :

| Colonne | Contenu | Exemple |
|---------|---------|---------|
| A | Code du formulaire | `F60010001` |
| B | Libellé descriptif | `Actifs non courants (Brut)` |
| C | Valeur en Dinars | `1595640.954` |

### Feuille `entet` — Informations du déclarant

| Champ | Exemple |
|-------|---------|
| Matricule Fiscal | `960046GAM000` |
| Nomet Prenom ou Raison Sociale | `STE ETTADHAMEN DU COMMERCE` |
| Activite | `COMMERCE EN GROS MATERIAUX DE CONSTRUCTION` |
| Adresse | `RUE SIDI KOUBINE 5050 MOKNINE` |
| Exercice | `2024` |
| Date Debut Exercice | `01/01/2024` |
| Date Cloture Exercice | `31/12/2024` |
| Acte DeDepot | `0` → 0 = Spontané, 1 = Rectification, 2 = Régularisation |
| Nature Depot | `P` → P = Provisoire, D = Définitif |

> ⚠️ Écrivez **"Activite"** sans accent pour éviter les problèmes de détection.

### Conversion des montants

Les montants sont automatiquement convertis en millimes sur 13 chiffres :

| Montant en Dinars | Format XML généré |
|-------------------|-------------------|
| 1,000 DT | `0000000001000` |
| 1 595 640,954 DT | `0001595640954` |
| 212 343 243 DT | `0212343243000` |

---

## Installation

### Option 1 — Application portable `.exe` *(recommandée pour les comptables)*

Aucune installation technique requise.

1. Téléchargez le fichier `LiasseFiscale.exe`
2. Double-cliquez dessus
3. Une fenêtre console s'ouvre — **ne la fermez pas**, c'est le serveur local
4. Le navigateur s'ouvre automatiquement
5. L'application est prête à l'emploi

Pour arrêter l'application, fermez la fenêtre console ou appuyez sur `Ctrl + C`.

> ⚠️ Si votre antivirus bloque le fichier, il s'agit d'un faux positif fréquent avec les `.exe` générés par PyInstaller. Ajoutez simplement une exception.

---

### Option 2 — Version développement (Python)

**Prérequis :** Python 3.8 ou supérieur → [Télécharger Python](https://www.python.org/downloads/)

> ⚠️ Pendant l'installation de Python, cochez **"Add Python to PATH"**.

```bash
# 1. Cloner le dépôt
git clone https://github.com/votre-username/liasse-fiscale-converter.git
cd liasse-fiscale-converter

# 2. Créer les dossiers nécessaires
mkdir templates static

# 3. Déposer les fichiers :
#    → dashboard.html dans templates/
#    → template_liasse_fiscale.xlsx dans static/ (optionnel)

# 4. Installer les dépendances
pip install -r requirements.txt

# 5. Lancer l'application
python app.py

# 6. Ouvrir dans le navigateur
#    → http://127.0.0.1:5000
```

**Contenu de `requirements.txt` :**

```
Flask==2.3.3
flask-cors==4.0.0
pandas==2.0.3
openpyxl==3.1.2
```

---

## Utilisation

| Étape | Action |
|-------|--------|
| 1 | Lancez l'application (`.exe` ou `python app.py`) |
| 2 | Attendez que le navigateur s'ouvre |
| 3 | Glissez-déposez votre fichier Excel |
| 4 | Cliquez sur **"Analyser le fichier"** |
| 5 | Vérifiez les formulaires détectés |
| 6 | Si deux modèles F6004 sont présents, choisissez-en un |
| 7 | Cliquez sur **"Générer les XML"** |
| 8 | Le fichier `.zip` contenant tous les XML est téléchargé automatiquement |

---

## Dépannage

| Problème | Cause probable | Solution |
|----------|----------------|----------|
| `Python n'est pas reconnu` | Python non installé ou PATH manquant | Utilisez la version `.exe` |
| L'antivirus bloque le `.exe` | Faux positif (PyInstaller) | Ajoutez une exception antivirus |
| Le navigateur ne s'ouvre pas | Problème `webbrowser` | Ouvrez manuellement `http://127.0.0.1:5000` |
| `"Template non trouvé"` | Dossier `static/` absent | Le template est optionnel, ignorez |
| `"Activité"` non détecté | Accent dans le libellé | Écrivez `Activite` sans accent |
| Erreur de conversion | Format Excel incorrect | Vérifiez les 3 colonnes (Code, Libellé, Valeur) |
| Port 5000 déjà utilisé | Autre application sur ce port | Modifiez `app.py` : `app.run(port=8080)` |
| Feuille non détectée | Nom de feuille incorrect | Vérifiez l'orthographe exacte des noms de feuilles |
| F6004 non généré | Les deux modèles sont présents | Choisissez un modèle dans l'interface |

### Erreurs courantes

**`ModuleNotFoundError: No module named 'flask'`**
```bash
pip install flask pandas openpyxl flask-cors
```

**`TemplateNotFound: dashboard.html`**
```bash
# Vérifiez que dashboard.html est bien dans le dossier templates/
ls templates/dashboard.html
```

---

## Structure du projet

```
liasse-fiscale-converter/
│
├── app.py                              ← Application principale Flask
├── requirements.txt                    ← Dépendances Python
├── README.md                           ← Ce fichier
│
├── templates/
│   └── dashboard.html                  ← Interface utilisateur
│
├── static/
│   └── template_liasse_fiscale.xlsx    ← Template Excel (optionnel)
│
├── dist/
│   └── LiasseFiscale.exe               ← Application portable (généré)
│
└── build/                              ← Dossier temporaire PyInstaller (à ignorer)
```

---

## Technologies utilisées

| Technologie | Version | Rôle |
|-------------|---------|------|
| Python | 3.8+ | Langage principal |
| Flask | 2.3.3 | Framework web |
| Pandas | 2.0.3 | Lecture des fichiers Excel |
| OpenPyXL | 3.1.2 | Moteur Excel |
| PyInstaller | 5.13.0 | Génération de l'`.exe` portable |

---

## Licence

Ce projet est distribué sous licence **MIT**. Vous êtes libre de l'utiliser, le modifier et le redistribuer, à condition de mentionner la licence originale.

---

🇹🇳 Conforme à la législation fiscale tunisienne &nbsp;|&nbsp; Développé avec ❤️ pour les comptables tunisiens
