# Happy Post — Outils d'expédition

**[>>> Utiliser l'app en ligne <<<](https://happy-post.streamlit.app/)**

![Avant / Après](preview.png)

## Fonctionnalités

### 1. Reformatage d'étiquettes

Les PDF d'étiquettes Happy Post utilisent **une page entière par étiquette**. Cet outil regroupe **4 étiquettes par feuille A4** — ~70% de papier économisé.

### 2. Génération du fichier d'import

Génère le fichier d'import Happy Post (.xlsx) directement depuis les bons de commande Amazon (PDF). Extraction automatique des adresses Belgique, calcul du poids selon le type de lot.

## Utilisation

### App en ligne (recommandé)

Rendez-vous sur **[happy-post.streamlit.app](https://happy-post.streamlit.app/)** — deux onglets, uploadez votre PDF, téléchargez le résultat.

### En local

```bash
pip install -r requirements.txt
streamlit run app.py
```

### Ligne de commande

```bash
# Reformater les étiquettes
python reformat_etiquettes.py <etiquettes_happypost.pdf> [output.pdf]

# Générer le fichier d'import
python generate_import.py <commandes_amazon.pdf> [YYYY-MM-DD]
```

## Dépendances

```bash
pip install -r requirements.txt
```

---

Un outil par [FURGO](https://shop.furgo.fr)
