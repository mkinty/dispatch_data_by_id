# Dispatcher SNA

Range le dossier **Prod IA / À déposer** (un sous-dossier par commune : Excel
rempli + `cas_analyse_<ID>.pptx`) dans l'arborescence SharePoint
`DepXX/<commune>/`.

## Modes

- **Commune complète** : l'Excel `audit_<commune>.xlsx` remplace celui de
  SharePoint, tous les PPT sont copiés dans `analyse/`.
- **ID imm** : on colle des **ID imm** (`ID-3806800000IW`, un par ligne), comme
  dans audit_validation_finale. La commune est donnée par les **5 caractères après
  « ID- »** (`38068`, ou `2B030` pour `ID-2B030000006D`). Minuscules acceptées,
  en-tête « ID imm » collé avec les valeurs ignoré, doublons retirés.
  - **Pourquoi l'ID imm** : un **ID erreur** peut couvrir plusieurs logements, donc
    plusieurs lignes. Seul l'**ID imm** est unique par ligne. La mise à jour vise
    donc exactement la ligne saisie, jamais un autre logement du même ID erreur. Un
    ID imm en double (erreur de saisie) : toutes ses lignes sont mises à jour.
  - **PPT** : ils restent nommés par ID erreur (`cas_analyse_<ID erreur>.pptx`).
    L'ID erreur de chaque ID imm est lu dans l'Excel Prod IA, à défaut dans l'Excel
    SharePoint, ce qui suffit pour les PPT seuls. Plusieurs ID imm du même ID erreur
    partagent son PPT, traité une seule fois.
  - **ID erreur collés par erreur** : ils sont reconnus et signalés, sans être
    utilisés. Les entrées « ID en attente » de l'ancienne version, au format ID
    erreur, ne sont pas rechargées : il faut recoller leurs ID imm.

  Pour chaque commune :
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
    - `ID imm` et `ID erreur` ne sont pas proposées : la première sert à retrouver
      les lignes, la seconde à nommer les PPT.
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

### Fichier ouvert : jamais écrasé (comme audit_validation_finale)

La voie d'écriture est choisie **fichier par fichier** (`services/ecriture.py`).
L'outil ne remplace jamais un fichier entier quand quelqu'un l'a ouvert :

| Situation | Ce que fait l'outil |
|---|---|
| `🔐 Écrire via Excel` coché, ou fichier **synchronisé avec SharePoint** (dossier OneDrive) | Excel rejoint la **co-édition**, n'écrit que les cellules des ID imm, enregistre puis relit. |
| Fichier déjà ouvert **dans votre propre Excel** | Les cellules sont écrites **dans ce classeur ouvert**, qui est enregistré et reste ouvert (vos modifications en cours sont enregistrées avec). Avant : « verrouillé par vous-même ». |
| Fichier local **ouvert**, case décochée | Le fichier verrou `~$…` donne le nom de la personne ; Excel est démarré **à la demande** pour écrire en co-édition. |
| Fichier ouvert et Excel indisponible | **Rien n'est écrit** : les ID partent en attente, avec le nom de la personne et la raison. |
| Fichier local libre, case décochée | Réécriture directe (openpyxl). |
| PPT ouvert (verrou `~$` de PowerPoint) | PPT non écrasé ; ID en attente avec le nom de la personne. |
| Mode commune, Excel ou PPT ouvert | Fichier non remplacé, dossier Prod IA conservé. |

Deux cas ne peuvent pas être forcés : un classeur ouvert en **mode exclusif**
(ancien Excel, enregistrement automatique désactivé) et un fichier **extrait**
(check-out). C'est à la personne qui a le fichier de les débloquer.

### Plantages et interface figée

- **Tkinter jamais appelé depuis le fil de traitement** : la case « Supprimer le
  dossier » est lue avant le lancement, et plus pendant le lot. Lire une case
  Tkinter depuis un autre fil peut faire planter l'outil.
- **Objets Excel libérés sur leur propre fil** : la trace d'une erreur Excel est
  vidée sur le fil Excel avant d'être transmise (`_liberer_objets_com`, repris
  d'audit_validation_finale). Sinon, ses objets COM seraient libérés depuis un
  autre fil, ce qui peut faire planter Python.
- **Rien n'attend OneDrive dans le fil de l'interface** : l'aperçu (liste de Prod
  IA, dossier SharePoint de chaque commune) et le bilan préalable au lancement se
  font en arrière-plan. Mesuré avec 475 communes collées : 42 ms de gel au plus.
- **Journal** : au plus 300 lignes affichées par rafraîchissement, 6 000 lignes
  gardées à l'écran. Le **journal complet** du lot est écrit dans
  `rapports/journal_<date>.txt`, dont le chemin est donné en fin de traitement.

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

## Sources SharePoint par lien (connexion avec votre compte)

Les deux champs **Racine SharePoint** et **Dossier À déposer** acceptent, au choix :

- un **dossier local** synchronisé par OneDrive (fonctionnement habituel) ;
- le **lien SharePoint du dossier**, comme dans audit_validation_finale :
  - l'adresse copiée dans la barre du navigateur, sur le dossier ouvert (`https://…sharepoint.com/sites/…/Audit`, ou `…/Forms/AllItems.aspx?id=…`) ;
  - un lien de partage « au chemin lisible » (`…/:f:/r/sites/…`) ;
  - un chemin WebDAV (`\\…sharepoint.com@SSL\DavWWWRoot\…`).

  Un lien de partage opaque (`/:f:/s/…/EaBc…`) ne contient pas le chemin : il est refusé, et l'outil indique la marche à suivre.

On peut mélanger : une source en lien et l'autre en dossier local.

**Connexion** (même processus qu'audit_validation_finale) :

1. Collez le lien, puis quittez le champ ou appuyez sur Entrée.
2. Cliquez sur **🔗 Se connecter à SharePoint**. L'outil essaie d'abord sans fenêtre, au cas où la session du profil dédié serait encore valide. Sinon, une fenêtre **Edge** (ou Chrome) s'ouvre sur le dossier : identifiez-vous normalement, avec votre compte professionnel et la MFA. La fenêtre se ferme dès que la session ouvre le dossier.
3. Au démarrage suivant, la reconnexion est silencieuse si la session est encore valide.

L'outil n'enregistre ni mot de passe ni cookie. La session reste dans un profil de navigateur dédié (`%LOCALAPPDATA%\DispatcherSNA\navigateur`), chiffré par Windows. **🚪 Oublier la connexion** efface ce profil.

**Traitement d'une commune** (`services/sharepoint/sources.py`). La logique de rangement reste la même. Un dossier de travail temporaire reçoit uniquement ce qui sert :

| Étape | Source par lien |
|---|---|
| Dossier commune de « À déposer » | téléchargé ; une fois rangé avec succès, **mis à la corbeille du site** (restaurable), jamais supprimé définitivement. Le dossier « À déposer » lui-même n'est jamais supprimé. |
| Excel audit (mise à jour par ID imm) | téléchargé avec sa version (ETag), lignes écrites, puis renvoyé **seulement si personne ne l'a enregistré entre-temps** (If-Match). Sinon il est relu et la mise à jour est refaite (3 essais), donc le travail du collègue n'est jamais perdu. |
| Excel ouvert par quelqu'un sur SharePoint (Excel, Excel Online, Teams) | SharePoint refuse de remplacer le fichier (verrou 423). L'outil l'ouvre alors dans Excel par son adresse SharePoint, rejoint la co-édition, écrit **cellule par cellule** les colonnes cochées des lignes des ID imm, enregistre et relit les valeurs. SharePoint fusionne avec le travail du collègue. Sans Excel, rien n'est écrit et l'ID passe en attente, avec le nom de la personne. |
| Classeur ouvert en mode **exclusif** (sans co-édition) | Excel ne l'obtient qu'en lecture seule : rien n'est écrit, le nom de la personne est donné, avec la marche à suivre (le fermer, ou l'ouvrir dans Excel Online / Microsoft 365 avec l'enregistrement automatique). |
| Excel extrait (check-out) | rien n'est écrit, et le nom de la personne qui l'a extrait est donné. |
| PPT à modifier | téléchargés avec leur ETag, modifiés, renvoyés (If-Match). |
| PPT remplacés, fichiers neufs | déposés ; le dossier `analyse` est créé au besoin. |
| PPT **ouvert** par quelqu'un (sur SharePoint, ou fichier verrou `~$` en local) | écrit **en co-édition PowerPoint**, par son adresse (`services/ppt_com.py`). « Modifier » : seules la zone Analyse (la même que celle choisie par l'édition XML) et la ligne « Lien Street View : » (texte et hyperlien). « Remplacer » : les diapositives du nouveau PPT sont insérées et les anciennes retirées, dans la présentation ouverte. Le résultat est relu après l'enregistrement. Si vous avez déjà ouvert la présentation dans votre PowerPoint, l'outil écrit dans celle-ci, qui reste ouverte. Le PowerPoint de l'utilisateur n'est jamais fermé. |
| « 🔐 Écrire via Excel » coché | l'Excel est écrit en co-édition, ouvert par son adresse SharePoint. Si Excel ne peut pas ouvrir ce fichier, l'outil passe par l'API. |
| Envoi par l'API | la version du serveur est **relue** après l'envoi : un ID n'est compté qu'une fois ses valeurs retrouvées sur SharePoint. |

**Connexion Microsoft 365 demandée par Excel** (première ouverture d'un fichier du site) : l'outil la détecte tout de suite, au lieu d'attendre le délai anti-blocage. Il ouvre alors une fenêtre Excel visible pour vous connecter une fois avec votre compte professionnel, puis reprend.

**Fichiers ouverts : nouvel essai.** En fin de lot, les ID refusés parce que leur fichier était ouvert (hors fichier extrait) sont retentés une fois, après 60 s. Les réglages sont `reprises_verrou` et `delai_reprise_s` dans `config_dispatcher.json`.

**Limite.** Avec une racine en **dossier local OneDrive**, un PPT ouvert par un collègue *sur SharePoint* est invisible depuis le poste : seul le fichier verrou `~$` local se voit. Pour être protégé dans tous les cas, utilisez le **lien SharePoint** de la racine.

Si la session expire en cours de lot, celui-ci s'arrête proprement : les communes restantes passent dans « ID en attente ». Il suffit de se reconnecter puis de cliquer sur « Charger les ID en attente ».

Si l'organisation interdit le téléchargement depuis un poste non géré (accès conditionnel), l'outil le signale sans le contourner.

## Structure

```
main.py                      # Point d'entrée : lance la fenêtre
dispatch_sna/
├── config.py                # Chargement/sauvegarde config_dispatcher.json
├── services/                # Logique métier PURE — aucune dépendance à Tkinter
│   ├── constantes.py        #   préfixes Dep/audit_/cas_analyse_, noms de colonnes
│   ├── error_ids.py         #   extraction des ID imm (et ID erreur), regroupés par commune
│   ├── ecriture.py          #   voie d'écriture par fichier : co-édition si ouvert ou synchronisé
│   ├── excel_rows.py        #   en-têtes, lecture et remplacement des lignes (openpyxl)
│   ├── excel_com.py         #   écriture cellule par cellule via Excel, en co-édition + vérification
│   ├── onedrive.py          #   dossier synchronisé local → adresse SharePoint
│   ├── attente.py           #   liste persistante des ID non appliqués
│   ├── ppt_editor.py        #   réécriture Analyse / Lien Street View (XML du .pptx)
│   ├── ppt_com.py           #   PPT ouvert : écriture via PowerPoint, en co-édition + relecture
│   ├── chemins.py           #   arborescence Prod IA / SharePoint, copies, détection des colonnes
│   ├── sharepoint/          #   sources par LIEN (repris d'audit_validation_finale)
│   │   ├── navigateur.py    #     Edge/Chrome piloté (profil dédié) : identification, cookies
│   │   ├── connexion.py     #     liens acceptés, connexion silencieuse puis fenêtre, sessions
│   │   ├── rest.py          #     API SharePoint : lecture, ETag/If-Match, dépôt, corbeille
│   │   └── sources.py       #     dossier de travail par commune + envoi des fichiers écrits
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
