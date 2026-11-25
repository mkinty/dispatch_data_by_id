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
    - Une colonne cochée remplace toujours l'ancienne valeur, **même si la valeur
      de Prod IA est vide** : la cellule SharePoint est alors effacée, lien et
      formule compris, et le journal le signale (🧽).
    - `ID erreur` n'est pas proposée : elle sert à retrouver les lignes.
    - Détection : Prod IA jusqu'à 3 niveaux de sous-dossiers (`À déposer/LOTxx/<commune>/`
      compris), puis `DepXX/<commune>/audit_*.xlsx`. Jusqu'à 12 fichiers sont essayés,
      ceux déjà présents sur le disque d'abord. Un fichier n'est retenu que si son onglet
      Audit contient « ID erreur ». En cas d'échec, la raison s'affiche pour chaque
      fichier écarté, et *📂 Choisir un fichier…* permet d'indiquer un Excel à la main.
    - Tant que les colonnes ne sont pas connues, la mise à jour par ID refuse de se
      lancer : jamais « toutes les colonnes » à l'aveugle.
    - Onglet Audit, pour la lecture comme pour l'écriture :
      1. l'onglet nommé « Audit » (casse et espaces ignorés), sinon l'unique onglet
         dont le nom contient « audit » (« Audit SNA »…). Plusieurs candidats = erreur ;
      2. sinon la **1re feuille** du classeur, si sa ligne 1 contient « ID erreur ».
         Le journal le signale (📄). Sans « ID erreur », le fichier est refusé plutôt
         que d'écrire dans une feuille qui n'est pas la bonne.
    - Aucune colonne cochée → l'Excel n'est pas modifié (seuls les PPT le sont).
  - **PPT** (cases exclusives) :
    - *Remplacer les PPT* : `cas_analyse_<ID>.pptx` est écrasé par celui de Prod IA ;
    - *Modifier les PPT* : dans le PPT existant, la zone « Analyse » reçoit la
      colonne `Remarque` et la ligne « Lien Street View : » la colonne
      `Lien Street View` (hyperlien de la cellule prioritaire). Une valeur vide
      vide la zone si la colonne est cochée, et la laisse inchangée sinon. Si le
      PPT n'existe pas encore sur SharePoint,
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

**Excel occupé (« L'appel a été rejeté par l'appelé », -2147418111)** : quand
Excel est en train de charger un classeur SharePoint, de se synchroniser avec
les co-éditeurs, de recalculer ou d'enregistrer, il refuse les commandes
(`RPC_E_CALL_REJECTED`). L'outil s'en protège de trois façons :

1. un **filtre de messages COM** (`_FiltreMessages`, `IMessageFilter` via
   ctypes) fait réessayer chaque appel refusé toutes les 250 ms, jusqu'à
   `excel_delai_occupe_s` secondes (120 par défaut) ;
2. le **calcul automatique est suspendu** pendant l'écriture des cellules
   (un seul recalcul à la fin, avant l'enregistrement) et l'outil attend
   qu'Excel soit prêt après l'ouverture et après l'enregistrement ;
3. si Excel reste occupé malgré tout, le classeur est **retenté entièrement**
   (3 essais). En dernier recours, les ID partent dans la liste d'attente avec
   la raison « Excel est resté occupé ».

**Excel bloqué sur une fenêtre invisible** : si SharePoint affiche une boîte
modale (« Nous rechargeons votre classeur pour terminer la fusion »…), personne
ne peut cliquer dessus dans l'Excel invisible, et l'appel en cours ne revient
jamais. La session est donc **surveillée** (`SessionExcelSurveillee`) :
- Excel tourne sur un fil dédié, propriétaire de sa session COM ;
- l'outil surveille ses signes de vie : réponses, refus « occupé » réessayés,
  avancement cellule par cellule ;
- sans signe de vie pendant `excel_delai_blocage_s` (120 s), le processus de
  **cette** instance est tué (jamais l'Excel de l'utilisateur), l'outil attend sa
  disparition effective, en relance une neuve et retente le classeur ;
- une instance Excel disparue (plantage) est relancée de la même façon.

L'outil ne gèle donc plus : au pire, une commune échoue et ses ID rejoignent la
liste « ID en attente ». L'enregistrement automatique est coupé pour l'instance de
l'outil, pour éviter des fusions à répétition.

**Vérification après enregistrement** : elle repose uniquement sur la relecture
des cellules écrites. Le drapeau `Saved` du classeur n'est plus utilisé : la
moindre activité de co-édition d'un collègue le remet à faux, et il faisait
rejeter à tort tous les ID d'une commune.

**PPT seuls** (aucune colonne cochée) : aucun Excel n'est écrit, et Excel n'est
même pas lancé. L'Excel Prod IA ne sert qu'à lire Remarque et Lien Street View
en mode « Modifier ». Les ID absents de cet Excel ne sont donc plus des erreurs
en mode « Remplacer ». Le dossier Prod IA de la commune est supprimé dès que les
PPT ont été traités sans erreur, même s'il contient l'Excel. Seul cas où il est
conservé : il contient des PPT d'autres ID, non saisis, qui ne sont pas encore
livrés.

**Refus OneDrive (« [WinError 389] L'opération de cloud a échoué »)** : les PPT
déjà livrés sont « disponibles en ligne uniquement », car l'outil les désépingle
pour libérer de la place. OneDrive refuse souvent de *renommer par-dessus* un tel
fichier. Toutes les lectures et écritures passent donc par
`services/fichiers_cloud.py` :
- les erreurs cloud passagères sont réessayées (pauses de 1, 3, 8 puis 15 s) ;
- seuls les octets sont copiés, jamais les attributs du fichier source (lecture
  seule, attributs OneDrive) ;
- le remplacement se fait en trois temps : remplacement direct ; sinon
  rapatriement de l'ancien fichier puis remplacement ; sinon suppression puis
  dépôt, avec remise en place de l'ancien fichier si le dépôt échoue. La
  destination n'est jamais laissée vide ;
- la taille écrite est vérifiée, et l'erreur finale nomme l'étape en cause.

**Disjoncteurs** :
- **par commune** : 3 refus OneDrive d'affilée → les PPT restants sont mis en
  attente sans être tentés ;
- **par lot** : 3 communes d'affilée dans ce cas → pause de 60 s, puis arrêt du
  lot si cela continue. Les communes restantes passent dans « ID en attente ».

La libération d'espace (désépinglage) est faite après tous les PPT de la commune.

**Filtres et colonnes masquées** : à chaque mise à jour par ID, dans l'onglet
`Audit` uniquement, l'outil retire les critères de filtre (tableaux Excel
compris) et réaffiche les lignes et colonnes masquées, côté Prod IA comme côté
SharePoint. Les flèches de filtre restent en place. Le journal indique chaque
retrait (🧹). Sur un onglet **protégé** :
- le filtre est retiré quand même (`AutoFilter.ShowAllData`, autorisé là où
  `Worksheet.ShowAllData` échoue) ;
- les lignes et colonnes masquées à la main restent masquées, et le journal le
  signale (🔒). Cela n'a pas d'effet sur la recherche des ID ;
- un lien cliquable ne peut pas être créé : l'adresse est écrite en texte (🔗) ;
- une cellule verrouillée à mettre à jour donne une erreur claire. La
  protection n'est jamais retirée par l'outil.

Si aucune colonne n'est cochée, le fichier SharePoint n'est pas
ouvert et ses filtres restent. L'onglet `Audit` est obligatoire : s'il manque,
l'erreur le dit et aucun autre onglet n'est utilisé à sa place.

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
