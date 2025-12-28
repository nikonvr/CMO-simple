[README_STREAMLIT.md](https://github.com/user-attachments/files/24360224/README_STREAMLIT.md)
# Application Streamlit - Calcul d'Empilement de Couches Minces

Version web ultra ergonomique de l'application de calcul d'empilement de couches minces.

## Installation

1. Installer les dépendances :
```bash
pip install -r requirements_streamlit.txt
```

Ou manuellement :
```bash
pip install streamlit numpy matplotlib pandas xlsxwriter
```

## Lancement

Pour lancer l'application Streamlit :

```bash
streamlit run app_streamlit.py
```

L'application s'ouvrira automatiquement dans votre navigateur web à l'adresse `http://localhost:8501`

## Fonctionnalités

### ✅ Interface moderne et ergonomique
- Design épuré avec CSS personnalisé
- Layout responsive (large)
- Sidebar organisée par sections
- Onglets pour les différents graphiques

### ✅ Toutes les fonctionnalités de la version PyQt6
- Calcul des propriétés optiques (Rs, Rp, Ts, Tp)
- Graphique spectral (en fonction de la longueur d'onde)
- Graphique angulaire (en fonction de l'angle d'incidence)
- Visualisation de l'empilement (profil d'indice)
- Export Excel avec toutes les données

### ✅ Système UNDO/REDO (5 niveaux)
- Boutons UNDO/REDO dans la sidebar
- Historique de 5 états maximum
- Sauvegarde automatique des modifications

### ✅ Export de données
- Téléchargement des graphiques (PNG)
- Export Excel avec paramètres et données calculées
- Format professionnel prêt à utiliser

## Utilisation

1. **Configurez les paramètres** dans la barre latérale :
   - Matériaux (indices de réfraction)
   - Configuration de l'empilement
   - Paramètres spectraux et angulaires
   - Options d'affichage

2. **Cliquez sur "Calculer"** pour lancer le calcul

3. **Consultez les résultats** dans les onglets :
   - Graphique Spectral
   - Graphique Angulaire
   - Visualisation Empilement

4. **Téléchargez les résultats** :
   - Graphiques en PNG
   - Données en Excel (si activé)

## Structure du code

- `app_streamlit.py` : Application Streamlit principale
- `cm_simple7.py` : Module de calcul (importé)
- `requirements_streamlit.txt` : Dépendances Python

## Avantages de la version Streamlit

- ✅ Accessible depuis n'importe quel navigateur
- ✅ Pas d'installation nécessaire pour l'utilisateur final (si déployée)
- ✅ Interface moderne et intuitive
- ✅ Partage facile (lien URL)
- ✅ Compatible avec tous les systèmes d'exploitation

## Déploiement

Pour déployer l'application en ligne, vous pouvez utiliser :
- Streamlit Cloud (gratuit) : https://streamlit.io/cloud
- Heroku
- AWS/GCP/Azure
- Votre propre serveur

