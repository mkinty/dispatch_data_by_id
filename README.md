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

### Modifications persistantes quand le fichier est ouvert par un collègue

Réécrire le fichier entier puis laisser OneDrive l'envoyer échoue dès qu'un autre
utilisateur a le classeur ouvert : OneDrive ne sait pas fusionner un fichier
produit hors d'Office (conflit, copie `-NOMDUPC`, ou envoi refusé). Les
modifications restaient alors dans le OneDrive local.

Case **🔐 Écrire via Excel (co-édition SharePoint)**, cochée par défaut
(`services/excel_com.py`, nécessite Excel + `pywin32`) :

1. une instance Excel dédiée et invisible est lancée. L'Excel déjà ouvert sur le
   poste n'est jamais touché ;
2. le classeur est ouvert **par son adresse SharePoint**, retrouvée à partir
   des dossiers synchronisés par OneDrive (`services/onedrive.py`), donc dans
   la session de co-édition. Si l'adresse n'est pas trouvée, il est ouvert par
   le chemin local ;
3. seules les cellules (ID × colonnes cochées) sont écrites : valeurs, formules
   recalées, liens, formats de nombre. Les couleurs et polices ne sont pas
   recopiées ;
4. après l'enregistrement, chaque cellule est **relue** : un ID n'est compté
   comme mis à jour que si toutes ses valeurs sont en place ;
5. si le classeur est ouvert en **mode exclusif** ailleurs, Excel l'ouvre en
   lecture seule : rien n'est écrit, et l'ID est signalé avec le nom de la
   personne qui bloque.

**ID en attente** : tout ID non appliqué (fichier bloqué, valeur non confirmée,
ID introuvable, PPT en erreur…) est enregistré avec sa raison dans
`ids_en_attente.json`. Le bouton *📋 Charger les ID en attente* les remet dans
la zone de saisie pour les rejouer. Ceux qui réussissent sont retirés de la
liste. Le dossier Prod IA d'une commune en échec est conservé, ce qui permet
de rejouer.

Décocher la case revient à l'ancienne réécriture du fichier (openpyxl).
`"excel_visible": true` dans `config_dispatcher.json` affiche Excel pendant le
traitement (diagnostic).

## Structure

```
main.py                      # Point d'entrée : lance la fenêtre
dispatch_sna/
├── config.py                # Chargement/sauvegarde config_dispatcher.json
├── services/                # Logique métier PURE — aucune dépendance à Tkinter
│   ├── constantes.py        #   préfixes Dep/audit_/cas_analyse_, noms de colonnes
│   ├── error_ids.py         #   extraction des ID erreur, regroupés par commune
│   ├── excel_rows.py        #   en-têtes, lecture et remplacement des lignes (openpyxl)
│   ├── excel_com.py         #   écriture cellule par cellule via Excel, en co-édition + vérification
│   ├── onedrive.py          #   dossier synchronisé local → adresse SharePoint
│   ├── attente.py           #   liste persistante des ID non appliqués
│   ├── ppt_editor.py        #   réécriture Analyse / Lien Street View (XML du .pptx)
│   ├── chemins.py           #   arborescence Prod IA / SharePoint, copies, détection des colonnes
│   └── dispatcher.py        #   orchestration : ranger_commune / ranger_ids
└── gui/                     # Présentation uniquement — délègue aux services
    ├── theme.py             #   couleurs et polices
    └── app.py               #   fenêtre Tkinter (thread de travail + _poll)
tests/                       # un fichier de test par service (+ fichiers.py : fabriques de fichiers)
```

**Règle** : `services/` n'importe jamais `gui/`. Les services rendent compte de
leur progression via deux callbacks (`log_fn(message, couleur)` et
`phase_fn(étape, état)`), ce qui permet de tout tester sans ouvrir de fenêtre.

## Lancer

```bash
uv run main.py
```

## Tests

```bash
uv run pytest
```
