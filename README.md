# Dispatcher SNA

Range le dossier **Prod IA / À déposer** (un sous-dossier par commune : Excel
rempli + `cas_analyse_<ID>.pptx`) dans l'arborescence SharePoint
`DepXX/<commune>/`.

## Modes

- **Commune complète** : l'Excel `audit_<commune>.xlsx` remplace celui de
  SharePoint, tous les PPT sont copiés dans `analyse/`.
- **ID erreur** : on colle des ID erreur (`74143_1, 74143_2, 38068_207`, comme
  dans export_lot_by_id). Pour chaque commune :
  - **Excel** : seules les lignes de ces ID sont mises à jour dans
    `DepXX/<commune>/audit_<commune>.xlsx`, et seulement pour les **colonnes
    cochées** dans *Configuration → Colonnes Excel à mettre à jour* (colonnes
    associées par nom d'en-tête ; valeurs, hyperliens, mise en forme, formules
    recalées). Le reste du fichier n'est pas touché. Un ID absent du fichier
    SharePoint n'est pas ajouté (signalé dans le journal).
    - Les colonnes sont détectées automatiquement au démarrage et à chaque
      changement de chemin (en-têtes de l'onglet `Audit` du 1er Excel de
      Prod IA, sinon d'un `audit_<commune>.xlsx` SharePoint) ; bouton
      *🔍 Détecter* pour relancer. Le choix est sauvegardé dans
      `config_dispatcher.json` (`colonnes_maj`).
    - `ID erreur` n'est pas proposée : elle sert à retrouver les lignes.
    - Aucune colonne cochée → l'Excel n'est pas modifié (seuls les PPT le sont).
  - **PPT** (cases exclusives) :
    - *Remplacer les PPT* : `cas_analyse_<ID>.pptx` est écrasé par celui de Prod IA ;
    - *Modifier les PPT* : dans le PPT existant, la zone « Analyse » reçoit la
      colonne `Remarque` et la ligne « Lien Street View : » la colonne
      `Lien Street View` (hyperlien de la cellule prioritaire). Une valeur vide
      laisse la zone inchangée. Si le PPT n'existe pas encore sur SharePoint,
      celui de Prod IA est copié ;
    - aucune case : Excel seul.
  - Le dossier Prod IA de la commune n'est supprimé que si tout a réussi **et**
    qu'il ne contient pas d'autres ID que ceux saisis.

## Lancer

```bash
uv run dispatch_data.py
```

## Tests

```bash
uv run pytest
```
