# index-excel-file

Une macro VBA pour générer automatiquement un sommaire cliquable des feuilles d'un classeur Microsoft Excel.

## Fonctionnalités

- Supprime le précédent onglet `Sommaire`, s'il existe.
- Crée un nouvel onglet `Sommaire` en première position.
- Ajoute un lien hypertexte vers la cellule `A1` de chaque autre feuille.
- Applique une présentation simple : titre, liens bleus soulignés et largeur ajustée.
- Affiche un message lorsque le sommaire est créé.

## Installation et utilisation

1. Ouvrir le classeur dans **Microsoft Excel** (version de bureau avec prise en charge de VBA).
2. Enregistrer le classeur au format **`.xlsm`** pour conserver les macros.
3. Ouvrir l'éditeur VBA avec **Alt + F11** sur Windows.
4. Dans l'éditeur VBA, sélectionner **Fichier > Importer un fichier...**, puis choisir `CreerSommaire.bas`. Autre possibilité : créer un module et y coller la procédure.
5. Exécuter la macro **`CréerSommaire`** via **Alt + F8** dans Excel, si les paramètres de sécurité autorisent les macros.

## À savoir

- Le sommaire ne se met pas à jour automatiquement : relancer la macro après avoir ajouté, supprimé ou renommé des feuilles.
- Toute modification manuelle de l'ancien onglet `Sommaire` est perdue lors de sa recréation.
- Les feuilles masquées sont également incluses dans la liste.
- La macro agit sur le classeur contenant le code (`ThisWorkbook`).
- Le code n'a pas été testé automatiquement dans Excel dans le cadre de cette publication.

## Licence

Distribué sous licence MIT. Voir [LICENSE](LICENSE).
