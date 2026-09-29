"""
═══════════════════════════════════════════════════════════════════
  DISPATCHER SNA — Organisation (PC SFR / OneDrive)
═══════════════════════════════════════════════════════════════════
Lit le dossier "Prod IA" (déposé depuis le PC Babayaga : un sous-dossier
par commune contenant l'Excel rempli + les PPT), puis range chaque fichier
au bon endroit dans l'arborescence SharePoint :

  Prod IA/81016/audit_81016.xlsx        →  Dep81/81016/audit_81016.xlsx  (remplace)
  Prod IA/81016/cas_analyse_81016_*.pptx →  Dep81/81016/analyse/

Une fois une commune rangée, son dossier dans Prod IA est SUPPRIMÉ.

Mode « ID erreur » : l'utilisateur colle des ID erreur (74143_1, 38068_207...).
Seules les lignes de ces ID sont remplacées dans Dep/commune/audit_commune.xlsx
(le reste du fichier SharePoint est conservé), et pour les PPT au choix :
  - Remplacer les PPT : cas_analyse_<ID>.pptx écrasé par celui de Prod IA
  - Modifier les PPT  : zones « Analyse » et « Lien Street View » du PPT
                        existant réécrites avec les colonnes « Remarque » et
                        « Lien Street View » de la ligne de l'ID

Sécurités :
- Ne touche rien si le dossier destination DepXX/commune n'existe pas.
- Remplace proprement le fichier Excel vierge du client.
- Ne supprime de Prod IA que si tout a réussi.

Lancement : python main.py (ou uv run main.py)
═══════════════════════════════════════════════════════════════════
"""
