# Application Streamlit - Échéanciers ACD vers Pennylane

Cette application transforme un échéancier d'emprunt ACD dans la structure du classeur Pennylane fourni : paramètres de l'emprunt en haut du fichier et échéancier détaillé à partir de la ligne 5.

## Entrées acceptées

- fichier CSV, TXT, TSV ou XLSX exporté depuis ACD ;
- tableau copié-collé directement depuis ACD ou Excel.

## Sorties

- classeur XLSX recommandé pour Pennylane ;
- fichier CSV conservant la même disposition ;
- texte tabulé à coller dans Excel à partir de la cellule A1.

## Démarrage sous Windows

1. Décompressez le dossier.
2. Vérifiez que Python 3.11 ou une version plus récente est installé.
3. Double-cliquez sur `lancer_application.bat`.
4. L'application Streamlit s'ouvre dans le navigateur.

## Démarrage manuel

```powershell
py -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m streamlit run app.py
```

## Principales règles

- La ligne ACD `N° 0` est reconnue comme ligne d'ouverture puis retirée de l'échéancier.
- Le capital est reconstitué en additionnant tous les déblocages successifs : `nouveau solde - ancien solde + capital remboursé`.
- Lorsqu'un déblocage figure dans la colonne ACD `Autre` et explique la hausse du restant dû, il n'est pas traité comme un frais.
- Les lignes à zéro sont conservées par défaut : elles représentent généralement un différé ou un moratoire.
- Le taux périodique ACD est annualisé selon la fréquence détectée.
- L'assurance reste distincte et la commission est regroupée avec la part de `Autre` qui ne correspond pas à un déblocage.
- L'échéance Pennylane est calculée par défaut comme `Annuité ACD + Assurance + Commission + Autre hors déblocage`.
- Le solde est contrôlé à chaque ligne selon `solde précédent + déblocage - capital remboursé`.
- Les téléchargements sont bloqués en présence d'une erreur de date, de calcul ou de solde.

## Fichiers du projet

- `app.py` : interface Streamlit ;
- `converter.py` : lecture, conversion, contrôles et génération des fichiers ;
- `requirements.txt` : dépendances Python ;
- `lancer_application.bat` : raccourci Windows.
