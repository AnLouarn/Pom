&ANALYZE-SUSPEND _VERSION-NUMBER UIB_v8r12 GUI ADM1
&ANALYZE-RESUME
/* Connected Databases 
*/
&Scoped-define WINDOW-NAME W-Win
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _DEFINITIONS W-Win 
CREATE WIDGET-POOL.
 
 
 
 
 
 
 
 

/*--- COMMENTAIRE MODIF PROGRAMME POSITIONNE PAR BLOCAGE ---
---> 08/03/18 - 11:01:15
---------------------------------------------------------------
Correction import pour:
- traiter proprement la date si elle est mal renseign‚e
- autoriser les lignes blanches
---------------------------------------------------------------
---> 11/07/17 - 17:10:00
---------------------------------------------------------------
Correction de problŠme de A Supprimer qui restait syst‚matiquement aprŠs une erreur d'import
---------------------------------------------------------------
---> 04/07/17 - 11:59:20
---------------------------------------------------------------
Correction problŠme obligatoire non coch‚ … tort
---------------------------------------------------------------
---> 29/12/16 - 18:00:54
---------------------------------------------------------------
Gestion du champ EAN
---------------------------------------------------------------
---> 28/10/16 - 14:15:04
---------------------------------------------------------------
Correction sur les calculs de dates PF + import des clients
--- COMMENTAIRE MODIF PROGRAMME POSITIONNE PAR BLOCAGE ---*/
 
/* ***************************  Definitions  ************************** */
 
/* Parameters Definitions ---                                           */
 
/* Local Variable Definitions ---                                       */
 
&IF DEFINED(UIB_is_Running) NE 0 
&THEN
  DEF VAR Menu_Utilisateur  AS CHAR  NO-UNDO INIT "NDE".
  DEF VAR Menu_Soc          AS ROWID NO-UNDO.
 
  FIND FIRST gpi_cnf WHERE gpi_cnf.cnf_code = "666" NO-LOCK NO-ERROR.
  IF AVAILABLE gpi_cnf THEN
      MENU_soc = ROWID(gpi_cnf).
&ELSE
  DEF NEW GLOBAL SHARED VAR Menu_Utilisateur    AS CHAR  NO-UNDO.
  DEF NEW GLOBAL SHARED VAR Menu_Soc            AS ROWID NO-UNDO.
&ENDIF
/*
DEFINE SHARED TEMP-TABLE tt_pf
  FIELD tt_num        AS CHARACTER
  FIELD tt_nom        AS CHARACTER
  FIELD tt_depAdr     AS CHARACTER
  FIELD tt_villeAdr   AS CHARACTER
  FIELD tt_dtLivPrev  AS DATE
  FIELD tt_estAppr    AS LOGICAL      /* True si au moins un jour d'approche pr‚vu */
  FIELD tt_approche   AS LOGICAL EXTENT 5
  INDEX idx_pf tt_num
  .
*/
FIND FIRST GPI_CNF WHERE ROWID ( GPI_CNF ) = Menu_Soc NO-LOCK NO-ERROR.
&IF NOT DEFINED(UIB_is_Running) EQ 0 
&THEN
    FIND FIRST GPI_CNF NO-LOCK.
    Menu_Soc = ROWID ( GPI_CNF ).
&ENDIF
 
FIND FIRST GPI_CNFP NO-LOCK.
 
DEFINE STREAM st_article.
DEFINE STREAM st_ol.
DEFINE STREAM st_param.
/*DEFINE STREAM st_log.*/
 
{objets\apiwindows\resize-auto-var.i}
 
 
/*DEF VAR ii              AS INT  NO-UNDO. */
DEF VAR plein-ecran     AS LOG  NO-UNDO.
DEF VAR Nom-Programme   AS CHAR NO-UNDO.
DEF VAR IsLocked        AS INT  NO-UNDO.
 
DEFINE VARIABLE gi_cpt AS INTEGER NO-UNDO.
 
DEFINE VARIABLE gl_modeBatch AS LOGICAL     NO-UNDO.
/* Le r‚pertoire et le nom du fichier Excel de log archiv‚ (si diff‚rent du r‚pertoire courant) */
DEFINE VARIABLE gc_replog    AS CHARACTER   NO-UNDO.
DEFINE VARIABLE gc_ficlog    AS CHARACTER   NO-UNDO.
 
/* Pour l'int‚gration */
DEFINE TEMP-TABLE tt_article NO-UNDO LIKE GPI_PGARTI USE-INDEX idx_arti_num_article
/*    FIELD arti_atco     AS CHARACTER    /* Code compos‚ */  */
/*    FIELD arti_famille  AS CHARACTER    /* Code famille */  */
 
    FIELD tt_dt_export  AS DATETIME
    FIELD tt_action     AS CHARACTER
    FIELD tt_dimensions AS CHARACTER    /* H * l * P - Pour affichage */
    FIELD tt_fichier    AS CHARACTER
    FIELD tt_numligne   AS INTEGER
    FIELD tt_controle   AS LOGICAL  /* Pour v‚rification des doublons ou erreurs (modif d'article inexistant, cr‚ation d'article d‚j… existant, suppression d'article inexistant, ... */
    INDEX idx_tri tt_fichier DESC tt_numligne DESC. /* Pour des recherches ordonn‚es. Ralentira les recherche, mais permet d'avoir le dernier ‚l‚ment cr‚‚ quelle que soit la recherche (en cherchant le premier) */
 
DEFINE BUFFER bf_article  FOR tt_article.
DEFINE BUFFER bf_article2 FOR tt_article.
 
DEFINE TEMP-TABLE tt_composant    NO-UNDO LIKE GPI_PGARTC
    FIELD tt_fichier  AS CHARACTER
    FIELD tt_numligne AS INTEGER.
 
DEFINE BUFFER bf_composant FOR tt_composant.
 
DEFINE TEMP-TABLE tt_ol         NO-UNDO LIKE GPI_PGOL
/*    FIELD ol_date_liv_souhait   AS DATE */
    FIELD tt_dt_export          AS DATETIME
    FIELD tt_fichier            AS CHARACTER
    FIELD tt_numligne           AS INTEGER.
 
DEFINE BUFFER bf_ol FOR tt_ol.
DEFINE BUFFER bf_pgol FOR GPI_PGOL.
 
DEFINE TEMP-TABLE tt_ligneOL    NO-UNDO LIKE GPI_PGDETAILOL
    FIELD tt_fichier  AS CHARACTER
    FIELD tt_numligne AS INTEGER.
 
DEFINE BUFFER bf_ligneOL FOR tt_ligneOL.
 
/* Ajout des interlocuteurs */
DEFINE TEMP-TABLE tt_interClient NO-UNDO LIKE gpi_pgic
    FIELD tt_fichier  AS CHARACTER
    FIELD tt_numligne AS INTEGER.
DEFINE BUFFER bf_interClient FOR tt_interClient.
 
DEFINE TEMP-TABLE tt_interEber NO-UNDO LIKE gpi_pgie
    FIELD tt_fichier  AS CHARACTER
    FIELD tt_numligne AS INTEGER
    INDEX idx_tri tt_fichier tt_numligne pgie_idts.
 
DEFINE BUFFER bf_interEber FOR tt_interEber.
 
DEFINE TEMP-TABLE tt_plan LIKE GPI_PGPT
    FIELD tt_fichier  AS CHARACTER
    FIELD tt_numligne AS INTEGER.
 
DEFINE BUFFER bf_plan FOR tt_plan.
 
DEFINE TEMP-TABLE tt_erreur NO-UNDO
    FIELD date_heure_erreur AS DATETIME
    FIELD fichier           AS CHARACTER
    FIELD numLigne          AS INTEGER
    FIELD libErreur         AS CHARACTER.
 
DEFINE TEMP-TABLE tt_fichiers NO-UNDO
    FIELD tt_fichier        AS CHARACTER
    FIELD tt_rep_origine    AS CHARACTER.
 
PROCEDURE LockWindowUpdate EXTERNAL "user32.dll" :
  DEFINE INPUT  PARAMETER hWndLock AS LONG.
  DEFINE RETURN PARAMETER IsLocked AS LONG.
END PROCEDURE.
 
DEFINE VARIABLE gl_outlook AS LOGICAL     NO-UNDO.
 
/* ****************************************** */
/*  Pour le calcul des dates de livraison PF  */
/* ****************************************** */
DEFINE VARIABLE gc_typePlateforme AS CHARACTER NO-UNDO.
/*DEFINE VARIABLE gl_tournee AS LOGICAL     NO-UNDO EXTENT 5. */
 
DEFINE TEMP-TABLE tt_plateforme
    FIELD tt_num         AS INTEGER      /* Pour ‚viter les conversions dans les recherches */
    FIELD tt_code        AS CHARACTER    /* Pour ‚viter les conversions dans les recherches */
    FIELD tt_livrable    AS LOGICAL      /* Vrai si au moins un jour de la tourn‚e … O */
    FIELD tt_dtLivPrev   AS DATE         /* La premiŠre date de livraison possible vers la plateforme - PrevDtCharg */
    /* M????? - NID le 03/07/17 - Pour une tentative de compr‚hension de l'algo, il s'agit des jours d'approche, pas de tourn‚e */
    /*FIELD tt_tournee     AS LOGICAL EXTENT 5 */
    FIELD tt_approche    AS LOGICAL EXTENT 5
    /* Fin M????? - NID le 03/07/17 */
    /* Pour la gestion du mode obligatoire */
    FIELD tt_dtArriveePF AS DATE         /* = DTCHARG2 */
    FIELD tt_dtCharg     AS DATE         /* Le premier jour de chargement possible depuis la plateforme vers le client final = DTCHARG */
    FIELD tt_dtLivCourt  AS DATE         /* La premiŠre date de livraison possible, sans +2 de la livraison 72h = DTL */
    FIELD tt_dtLivLong   AS DATE         /* La premiŠre date de livraison possible, avec le +2 de la livraison 72h = DTL */
    FIELD tt_dtLiv2Jours AS DATE         /* La date de livraison si T2J = 2 (calcul particulier) = DTL */
 
    INDEX idx_num tt_num
    INDEX idx_code tt_code
    .
 
DEFINE TEMP-TABLE tt_client LIKE GPI_PGCLI
    FIELD tt_origine  AS CHARACTER  /* CLI ou OL */
    FIELD tt_fichier  AS CHARACTER
    FIELD tt_numligne AS INTEGER
    INDEX idxcli cli_num_client cli_cp tt_fichier
    INDEX idxRech tt_origine cli_num_client cli_cp tt_fichier.
 
DEFINE TEMP-TABLE tt_cliint LIKE gpi_pgcliint
    FIELD tt_fichier  AS CHARACTER
    FIELD tt_numligne AS INTEGER
    INDEX idxcliint cli_num_client cli_cp cliint_cle tt_fichier.
 
DEFINE TEMP-TABLE TT_OL_EXCLUS NO-UNDO
    FIELD num_ol AS CHAR 
    INDEX idx_num_ol num_ol.
 
DEFINE VARIABLE gl_importOL AS LOGICAL     NO-UNDO.
 
DEFINE VARIABLE gdt_dateImport  AS DATE     NO-UNDO.
DEFINE VARIABLE gi_heureImport  AS INTEGER  NO-UNDO.
DEFINE VARIABLE gi_minImport    AS INTEGER  NO-UNDO.
/* 
Champs disponibles
 - cli_num_client   : Nø Client (cl‚)
 - cli_cp           : CP Livraison (cl‚)
 - cli_hayon        : Hayon (O/N)
 - cli_mail         : Mail Auto (O/N)
 - cli_dderdv       : Demande de RDV (O/N)
 - cli_confrdv      : Confirmation de RDV (O/N)
 - cli_lot          : lot (O/N)
 - cli_com          : observation
 Seuls cli_hayon, cli_dderdv, cli_lot et cli_com peuvent ˆtre modifi‚s par l'import pour un mˆme couple nø client/CP Livraison
 */
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&ANALYZE-SUSPEND _UIB-PREPROCESSOR-BLOCK 
 
/* ********************  Preprocessor Definitions  ******************** */
 
&Scoped-define PROCEDURE-TYPE SmartWindow
&Scoped-define DB-AWARE no
 
&Scoped-define ADM-CONTAINER WINDOW
 
/* Name of designated FRAME-NAME and/or first browse and/or first query */
&Scoped-define FRAME-NAME F-Main
&Scoped-define BROWSE-NAME Br_Articles
 
/* Internal Tables (found by Frame, Query & Browse Queries)             */
&Scoped-define INTERNAL-TABLES tt_article tt_client tt_composant tt_erreur ~
tt_interEber tt_ligneOL tt_ol tt_plan
 
/* Definitions for BROWSE Br_Articles                                   */
&Scoped-define FIELDS-IN-QUERY-Br_Articles tt_article.tt_action tt_article.ARTI_NUM_ARTICLE tt_article.ARTI_CODE_CONSTRUCTEUR tt_article.ARTI_REFERENCE tt_article.ARTI_DESIGNATION tt_article.ARTI_POIDS_BRUT tt_article.tt_dimensions tt_article.ARTI_COEFFICIENT tt_article.ARTI_COEF_GERBAGE tt_article.ARTI_DIVISION tt_article.ARTI_CODE_ENSEMBLE   
&Scoped-define ENABLED-FIELDS-IN-QUERY-Br_Articles   
&Scoped-define SELF-NAME Br_Articles
&Scoped-define QUERY-STRING-Br_Articles FOR EACH tt_article BY tt_article.tt_fichier BY tt_article.tt_numligne
&Scoped-define OPEN-QUERY-Br_Articles OPEN QUERY {&SELF-NAME} FOR EACH tt_article BY tt_article.tt_fichier BY tt_article.tt_numligne.
&Scoped-define TABLES-IN-QUERY-Br_Articles tt_article
&Scoped-define FIRST-TABLE-IN-QUERY-Br_Articles tt_article
 
 
/* Definitions for BROWSE Br_Clients                                    */
&Scoped-define FIELDS-IN-QUERY-Br_Clients tt_client.cli_num_client tt_client.cli_cp tt_client.cli_confrdv tt_client.cli_dderdv tt_client.cli_hayon tt_client.cli_lot tt_client.cli_mail tt_client.cli_com   
&Scoped-define ENABLED-FIELDS-IN-QUERY-Br_Clients   
&Scoped-define SELF-NAME Br_Clients
&Scoped-define OPEN-QUERY-Br_Clients /* OPEN QUERY {&SELF-NAME} FOR EACH tt_client BY tt_client.tt_fichier BY tt_client.tt_numligne. */ OPEN QUERY {&SELF-NAME} FOR EACH tt_client BY tt_client.tt_fichier BY tt_client.cli_num_client BY tt_client.cli_cp.
&Scoped-define TABLES-IN-QUERY-Br_Clients tt_client
&Scoped-define FIRST-TABLE-IN-QUERY-Br_Clients tt_client
 
 
/* Definitions for BROWSE Br_Corresp                                    */
&Scoped-define FIELDS-IN-QUERY-Br_Corresp tt_composant.ARTI_NUM_ARTICLE tt_composant.ARTC_NUM_COMPOSANT tt_composant.ARTC_QUANTITE   
&Scoped-define ENABLED-FIELDS-IN-QUERY-Br_Corresp   
&Scoped-define SELF-NAME Br_Corresp
&Scoped-define QUERY-STRING-Br_Corresp FOR EACH tt_composant BY tt_composant.tt_fichier BY tt_composant.tt_numligne
&Scoped-define OPEN-QUERY-Br_Corresp OPEN QUERY {&SELF-NAME} FOR EACH tt_composant BY tt_composant.tt_fichier BY tt_composant.tt_numligne.
&Scoped-define TABLES-IN-QUERY-Br_Corresp tt_composant
&Scoped-define FIRST-TABLE-IN-QUERY-Br_Corresp tt_composant
 
 
/* Definitions for BROWSE Br_Erreurs                                    */
&Scoped-define FIELDS-IN-QUERY-Br_Erreurs tt_erreur.date_heure_erreur tt_erreur.fichier tt_erreur.numLigne tt_erreur.libErreur   
&Scoped-define ENABLED-FIELDS-IN-QUERY-Br_Erreurs   
&Scoped-define SELF-NAME Br_Erreurs
&Scoped-define QUERY-STRING-Br_Erreurs FOR EACH tt_erreur BY tt_erreur.date_heure_erreur
&Scoped-define OPEN-QUERY-Br_Erreurs OPEN QUERY {&SELF-NAME} FOR EACH tt_erreur BY tt_erreur.date_heure_erreur.
&Scoped-define TABLES-IN-QUERY-Br_Erreurs tt_erreur
&Scoped-define FIRST-TABLE-IN-QUERY-Br_Erreurs tt_erreur
 
 
/* Definitions for BROWSE Br_InterEber                                  */
&Scoped-define FIELDS-IN-QUERY-Br_InterEber tt_interEber.pgie_idts tt_interEber.pgie_prenom tt_interEber.pgie_nom tt_interEber.pgie_mail tt_interEber.pgie_tel   
&Scoped-define ENABLED-FIELDS-IN-QUERY-Br_InterEber   
&Scoped-define SELF-NAME Br_InterEber
&Scoped-define QUERY-STRING-Br_InterEber FOR EACH tt_interEber BY tt_interEber.tt_fichier BY tt_interEber.tt_numligne
&Scoped-define OPEN-QUERY-Br_InterEber OPEN QUERY {&SELF-NAME} FOR EACH tt_interEber BY tt_interEber.tt_fichier BY tt_interEber.tt_numligne.
&Scoped-define TABLES-IN-QUERY-Br_InterEber tt_interEber
&Scoped-define FIRST-TABLE-IN-QUERY-Br_InterEber tt_interEber
 
 
/* Definitions for BROWSE Br_LigneOL                                    */
&Scoped-define FIELDS-IN-QUERY-Br_LigneOL tt_ligneOL.OL_NUM_OL tt_ligneOL.DETAILOL_NUM_LIGNE tt_ligneOL.ARTI_NUM_ARTICLE tt_ligneOL.DETAILOL_QUANTITE   
&Scoped-define ENABLED-FIELDS-IN-QUERY-Br_LigneOL   
&Scoped-define SELF-NAME Br_LigneOL
&Scoped-define QUERY-STRING-Br_LigneOL FOR EACH tt_ligneOL WHERE tt_ligneOL.ol_num_ol = tt_ol.ol_num_ol AND tt_ligneOL.tt_fichier = tt_ol.tt_fichier BY tt_ligneOL.tt_fichier BY tt_ligneOL.tt_numligne
&Scoped-define OPEN-QUERY-Br_LigneOL OPEN QUERY {&SELF-NAME} FOR EACH tt_ligneOL WHERE tt_ligneOL.ol_num_ol = tt_ol.ol_num_ol AND tt_ligneOL.tt_fichier = tt_ol.tt_fichier BY tt_ligneOL.tt_fichier BY tt_ligneOL.tt_numligne.
&Scoped-define TABLES-IN-QUERY-Br_LigneOL tt_ligneOL
&Scoped-define FIRST-TABLE-IN-QUERY-Br_LigneOL tt_ligneOL
 
 
/* Definitions for BROWSE Br_OL                                         */
&Scoped-define FIELDS-IN-QUERY-Br_OL tt_ol.OL_NUM_OL tt_ol.OL_GESTIONNAIRE_COMMANDES tt_ol.OL_NUM_CLIENT tt_ol.OL_REF_CLIENT tt_ol.OL_MODE_TRANSPORT tt_ol.OL_TITRE_COMMANDE tt_ol.OL_NOM_COMMANDE tt_ol.OL_ADRESSE1_COMMANDE tt_ol.OL_ADRESSE2_COMMANDE tt_ol.OL_CP_COMMANDE tt_ol.OL_VILLE_COMMANDE tt_ol.OL_TITRE_LIVRAISON tt_ol.OL_NOM_LIVRAISON tt_ol.OL_ADRESSE1_LIVRAISON tt_ol.OL_ADRESSE2_LIVRAISON tt_ol.OL_CP_LIVRAISON tt_ol.OL_VILLE_LIVRAISON tt_ol.OL_LOT_OBLIGATOIRE tt_ol.OL_RDV_OBLIGATOIRE tt_ol.OL_HAYON_OBLIGATOIRE /* tt_ol.OL_RDV_OBLIGATOIRE tt_ol.OL_HAYON_OBLIGATOIRE */   
&Scoped-define ENABLED-FIELDS-IN-QUERY-Br_OL   
&Scoped-define SELF-NAME Br_OL
&Scoped-define QUERY-STRING-Br_OL FOR EACH tt_ol BY tt_OL.tt_fichier BY tt_OL.tt_numligne
&Scoped-define OPEN-QUERY-Br_OL OPEN QUERY {&SELF-NAME} FOR EACH tt_ol BY tt_OL.tt_fichier BY tt_OL.tt_numligne.
&Scoped-define TABLES-IN-QUERY-Br_OL tt_ol
&Scoped-define FIRST-TABLE-IN-QUERY-Br_OL tt_ol
 
 
/* Definitions for BROWSE Br_PlanTournee                                */
&Scoped-define FIELDS-IN-QUERY-Br_PlanTournee tt_plan.pgpt_codpos tt_plan.pgpt_cpidx tt_plan.pgpt_insee tt_plan.pgpt_ville tt_plan.pgpt_vil20 tt_plan.pgpt_pays tt_plan.pgpt_tournee[1] tt_plan.pgpt_tournee[2] tt_plan.pgpt_tournee[3] tt_plan.pgpt_tournee[4] tt_plan.pgpt_tournee[5] tt_plan.pgpt_tournee[6] tt_plan.pgpt_montagne tt_plan.pgpt_trfpf tt_plan.pgpt_dept tt_plan.pgpt_plat tt_plan.pgpt_to2j tt_plan.pgpt_par   
&Scoped-define ENABLED-FIELDS-IN-QUERY-Br_PlanTournee   
&Scoped-define SELF-NAME Br_PlanTournee
&Scoped-define QUERY-STRING-Br_PlanTournee FOR EACH tt_plan BY tt_plan.tt_fichier BY tt_plan.tt_numligne
&Scoped-define OPEN-QUERY-Br_PlanTournee OPEN QUERY {&SELF-NAME} FOR EACH tt_plan BY tt_plan.tt_fichier BY tt_plan.tt_numligne.
&Scoped-define TABLES-IN-QUERY-Br_PlanTournee tt_plan
&Scoped-define FIRST-TABLE-IN-QUERY-Br_PlanTournee tt_plan
 
 
/* Definitions for FRAME F-Erreurs                                      */
&Scoped-define OPEN-BROWSERS-IN-QUERY-F-Erreurs ~
    ~{&OPEN-QUERY-Br_Erreurs}
 
/* Definitions for FRAME F-Main                                         */
&Scoped-define OPEN-BROWSERS-IN-QUERY-F-Main ~
    ~{&OPEN-QUERY-Br_Articles}~
    ~{&OPEN-QUERY-Br_Clients}~
    ~{&OPEN-QUERY-Br_Corresp}~
    ~{&OPEN-QUERY-Br_InterEber}~
    ~{&OPEN-QUERY-Br_LigneOL}~
    ~{&OPEN-QUERY-Br_OL}~
    ~{&OPEN-QUERY-Br_PlanTournee}
 
/* Standard List Definitions                                            */
&Scoped-Define ENABLED-OBJECTS Tg_reinitClients rs_visu Br_Articles ~
Br_Clients Br_Corresp Br_InterEber Br_OL Br_PlanTournee ED_Commentaire ~
Br_LigneOL liste-1 liste-3 liste-4 liste-2 Btn_Importer Btn_Test ~
Btn_Quitter lbl_intresutl 
&Scoped-Define DISPLAYED-OBJECTS Tg_reinitClients rs_visu ED_Commentaire ~
liste-1 liste-3 liste-4 liste-2 lbl_intresutl 
 
/* Custom List Definitions                                              */
/* List-1,List-2,List-3,List-4,List-5,List-6                            */
 
/* _UIB-PREPROCESSOR-BLOCK-END */
&ANALYZE-RESUME
 
 
/* ************************  Function Prototypes ********************** */
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION-FORWARD fi_ferie W-Win 
FUNCTION fi_ferie RETURNS LOGICAL
  ( INPUT idt_jour AS DATE )  FORWARD.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION-FORWARD fi_listeMailsValide W-Win 
FUNCTION fi_listeMailsValide RETURNS LOGICAL
  (listeEmails AS CHARACTER)  FORWARD.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION-FORWARD fi_mailValide W-Win 
FUNCTION fi_mailValide RETURNS LOGICAL
  (email AS CHARACTER )  FORWARD.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION-FORWARD fi_recupDateHeure W-Win 
FUNCTION fi_recupDateHeure RETURNS DATETIME
  ( ic_fichier AS CHARACTER )  FORWARD.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
/* ***********************  Control Definitions  ********************** */
 
/* Define the widget handle for the window                              */
DEFINE VAR W-Win AS WIDGET-HANDLE NO-UNDO.
 
/* Definitions of handles for SmartObjects                              */
DEFINE VARIABLE h_jauge-ocx AS HANDLE NO-UNDO.
 
/* Definitions of the field level widgets                               */
DEFINE BUTTON Btn_Executer 
     IMAGE-UP FILE "image\gpi_exec.bmp":U
     IMAGE-INSENSITIVE FILE "image\gpi_exe1.bmp":U
     LABEL "&Executer" 
     SIZE 9.29 BY 1
     FONT 10.
 
DEFINE BUTTON btn_rech_articles 
     IMAGE-UP FILE "IMAGE\gpi_cher":U
     IMAGE-INSENSITIVE FILE "image/gpi_che1.bmp":U
     LABEL "" 
     SIZE 3.14 BY 1.
 
DEFINE BUTTON btn_rech_OL 
     IMAGE-UP FILE "IMAGE\gpi_cher":U
     IMAGE-INSENSITIVE FILE "image/gpi_che1.bmp":U
     LABEL "" 
     SIZE 3.14 BY 1.
 
DEFINE BUTTON btn_rech_Param 
     IMAGE-UP FILE "IMAGE\gpi_cher":U
     IMAGE-INSENSITIVE FILE "image/gpi_che1.bmp":U
     LABEL "" 
     SIZE 3.14 BY 1.
 
DEFINE BUTTON btn_rech_svg 
     IMAGE-UP FILE "IMAGE\gpi_cher":U
     IMAGE-INSENSITIVE FILE "image/gpi_che1.bmp":U
     LABEL "" 
     SIZE 3.14 BY 1.
 
DEFINE VARIABLE fi_email AS CHARACTER FORMAT "X(256)":U 
     LABEL "Email supervision" 
     VIEW-AS FILL-IN 
     SIZE 82.43 BY .81 TOOLTIP "Adresse mail pour envois du log d'import" NO-UNDO.
 
DEFINE VARIABLE fi_rep_integration_articles AS CHARACTER FORMAT "X(256)":U 
     LABEL "R‚pertoire Articles" 
     VIEW-AS FILL-IN 
     SIZE 82.43 BY .81 TOOLTIP "R‚pertoire d'int‚gration" NO-UNDO.
 
DEFINE VARIABLE fi_rep_integration_ol AS CHARACTER FORMAT "X(256)":U 
     LABEL "R‚pertoire OL" 
     VIEW-AS FILL-IN 
     SIZE 82.43 BY .81 TOOLTIP "R‚pertoire d'int‚gration" NO-UNDO.
 
DEFINE VARIABLE fi_rep_integration_param AS CHARACTER FORMAT "X(256)":U 
     LABEL "R‚pertoire ParamŠtres" 
     VIEW-AS FILL-IN 
     SIZE 82.43 BY .81 TOOLTIP "R‚pertoire d'int‚gration" NO-UNDO.
 
DEFINE VARIABLE fi_rep_svg AS CHARACTER FORMAT "X(256)":U 
     LABEL "R‚pertoire Sauvegarde" 
     VIEW-AS FILL-IN 
     SIZE 82.43 BY .81 TOOLTIP "R‚pertoire d'int‚gration" NO-UNDO.
 
DEFINE BUTTON Btn_Excel 
     IMAGE-UP FILE "image/gpi_excel3.bmp":U
     IMAGE-INSENSITIVE FILE "image/gpi_excel4.bmp":U
     LABEL "&Excel" 
     SIZE 9.29 BY 1
     FONT 10.
 
DEFINE BUTTON Btn_Fermer 
     IMAGE-UP FILE "image\gpi_ok.bmp":U
     IMAGE-INSENSITIVE FILE "image\gpi_ok1.bmp":U
     LABEL "&Fermer" 
     SIZE 9.29 BY 1
     FONT 10.
 
DEFINE BUTTON Btn_Importer 
     IMAGE-UP FILE "image\gpi_impo.bmp":U
     IMAGE-INSENSITIVE FILE "image\gpi_impo1.bmp":U
     LABEL "&Importer" 
     SIZE 9.29 BY 1
     FONT 10.
 
DEFINE BUTTON Btn_Quitter 
     IMAGE-UP FILE "image\gpi_quit.bmp":U
     IMAGE-INSENSITIVE FILE "image\gpi_qui1.bmp":U
     LABEL "&Quitter" 
     SIZE 9.29 BY 1
     FONT 10.
 
DEFINE BUTTON Btn_Test 
     LABEL "Test Repart" 
     SIZE 9.29 BY 1.
 
DEFINE VARIABLE ED_Commentaire AS CHARACTER 
     VIEW-AS EDITOR NO-WORD-WRAP SCROLLBAR-HORIZONTAL SCROLLBAR-VERTICAL
     SIZE 111.14 BY 2.73 NO-UNDO.
 
DEFINE VARIABLE lbl_intresutl AS CHARACTER FORMAT "X(256)":U 
      VIEW-AS TEXT 
     SIZE 19.86 BY .54
     FONT 6 NO-UNDO.
 
DEFINE VARIABLE rs_visu AS CHARACTER 
     VIEW-AS RADIO-SET HORIZONTAL
     RADIO-BUTTONS 
          "Articles", "A",
"Corresp. Articles", "C",
"OL", "O",
"Plan de Tourn‚e", "P",
"Interlocuteurs Eberhardt", "IE",
"Clients", "CL"
     SIZE 70.72 BY .81 NO-UNDO.
 
DEFINE VARIABLE liste-1 AS CHARACTER 
     VIEW-AS SELECTION-LIST SINGLE SCROLLBAR-VERTICAL 
     SIZE 3.29 BY .31 NO-UNDO.
 
DEFINE VARIABLE liste-2 AS CHARACTER 
     VIEW-AS SELECTION-LIST SINGLE SCROLLBAR-VERTICAL 
     SIZE 3.29 BY .31 NO-UNDO.
 
DEFINE VARIABLE liste-3 AS CHARACTER 
     VIEW-AS SELECTION-LIST SINGLE SCROLLBAR-VERTICAL 
     SIZE 3.29 BY .31 NO-UNDO.
 
DEFINE VARIABLE liste-4 AS CHARACTER 
     VIEW-AS SELECTION-LIST SINGLE SCROLLBAR-VERTICAL 
     SIZE 3.29 BY .31 NO-UNDO.
 
DEFINE VARIABLE Tg_reinitClients AS LOGICAL INITIAL NO 
     LABEL "R‚initialisation clients" 
     VIEW-AS TOGGLE-BOX
     SIZE 17.43 BY .77 TOOLTIP "Vide les tables Clients et Interlocuteurs avant l'import" NO-UNDO.
 
DEFINE BUTTON Btn_param_annuler 
     IMAGE-UP FILE "image\gpi_annu.bmp":U
     IMAGE-INSENSITIVE FILE "image\gpi_ann1.bmp":U
     LABEL "&Annuler" 
     SIZE 9.29 BY 1
     FONT 10.
 
DEFINE BUTTON Btn_param_ok 
     IMAGE-UP FILE "image\gpi_ok.bmp":U
     IMAGE-INSENSITIVE FILE "image\gpi_ok1.bmp":U
     LABEL "&OK" 
     SIZE 9.29 BY 1
     FONT 10.
 
DEFINE VARIABLE ed_lstExclusions AS CHARACTER 
     VIEW-AS EDITOR NO-WORD-WRAP MAX-CHARS 31900 SCROLLBAR-VERTICAL
     SIZE 54.57 BY 19.81 NO-UNDO.
 
DEFINE IMAGE icone-banniere
     FILENAME "bourse/images/planning/find.bmp":U
     STRETCH-TO-FIT RETAIN-SHAPE TRANSPARENT
     SIZE 5.43 BY 1.46.
 
DEFINE IMAGE IMAGE-banniere
     FILENAME "image/menu/fonds/fond-colonne.jpg":U
     SIZE 6 BY 24.35.
 
DEFINE IMAGE img_errImport
     FILENAME "bourse/images/planning/find.bmp":U
     STRETCH-TO-FIT RETAIN-SHAPE TRANSPARENT
     SIZE 5.43 BY 1.46 TOOLTIP "Erreurs Import".
 
DEFINE IMAGE img_param
     FILENAME "bourse/images/planning/find.bmp":U
     STRETCH-TO-FIT RETAIN-SHAPE TRANSPARENT
     SIZE 5.43 BY 1.46 TOOLTIP "Erreurs Import".
 
/* Query definitions                                                    */
&ANALYZE-SUSPEND
DEFINE QUERY Br_Articles FOR 
      tt_article SCROLLING.
 
DEFINE QUERY Br_Clients FOR 
      tt_client SCROLLING.
 
DEFINE QUERY Br_Corresp FOR 
      tt_composant SCROLLING.
 
DEFINE QUERY Br_Erreurs FOR 
      tt_erreur SCROLLING.
 
DEFINE QUERY Br_InterEber FOR 
      tt_interEber SCROLLING.
 
DEFINE QUERY Br_LigneOL FOR 
      tt_ligneOL SCROLLING.
 
DEFINE QUERY Br_OL FOR 
      tt_ol SCROLLING.
 
DEFINE QUERY Br_PlanTournee FOR 
      tt_plan SCROLLING.
&ANALYZE-RESUME
 
/* Browse definitions                                                   */
DEFINE BROWSE Br_Articles
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS Br_Articles W-Win _FREEFORM
  QUERY Br_Articles DISPLAY
      tt_article.tt_action              FORMAT "X"          COLUMN-LABEL "Action"
      tt_article.ARTI_NUM_ARTICLE       FORMAT ">>>>>9"     COLUMN-LABEL "Nø Art."
      tt_article.ARTI_CODE_CONSTRUCTEUR FORMAT "X(2)"       COLUMN-LABEL "Const."
      tt_article.ARTI_REFERENCE         FORMAT "X(20)"      COLUMN-LABEL "Ref‚rence"
      tt_article.ARTI_DESIGNATION       FORMAT "X(40)"      COLUMN-LABEL "D‚signation"
      tt_article.ARTI_POIDS_BRUT        FORMAT ">>,>>>,>>9" COLUMN-LABEL "Poid brut (en g)"
      tt_article.tt_dimensions          FORMAT "X(20)"      COLUMN-LABEL "Dimensions (h/l/p)"
      tt_article.ARTI_COEFFICIENT       FORMAT ">>>9.999"   COLUMN-LABEL "Coefficient"
      tt_article.ARTI_COEF_GERBAGE      FORMAT ">9"         COLUMN-LABEL "Coef. Gerb."
      tt_article.ARTI_DIVISION          FORMAT "X"          COLUMN-LABEL "Div."
      tt_article.ARTI_CODE_ENSEMBLE     FORMAT "X"          COLUMN-LABEL "Ens."
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS SIZE 111.14 BY 14.31
         FONT 8 FIT-LAST-COLUMN.
 
DEFINE BROWSE Br_Clients
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS Br_Clients W-Win _FREEFORM
  QUERY Br_Clients DISPLAY
      tt_client.cli_num_client    FORMAT "999999" COLUMN-LABEL "Nø Client"
tt_client.cli_cp            FORMAT "X(5)"   COLUMN-LABEL "CP IDX"
tt_client.cli_confrdv                       COLUMN-LABEL "Conf!RDV" VIEW-AS TOGGLE-BOX
tt_client.cli_dderdv                        COLUMN-LABEL "Dde!RDV"  VIEW-AS TOGGLE-BOX
tt_client.cli_hayon                         COLUMN-LABEL "Hayon"    VIEW-AS TOGGLE-BOX
tt_client.cli_lot                           COLUMN-LABEL "Lot"      VIEW-AS TOGGLE-BOX
tt_client.cli_mail                          COLUMN-LABEL "Mail"     VIEW-AS TOGGLE-BOX
tt_client.cli_com           FORMAT "X(256)" COLUMN-LABEL "Commentaire" WIDTH 40
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS SIZE 111.14 BY 14.38
         FONT 8 FIT-LAST-COLUMN.
 
DEFINE BROWSE Br_Corresp
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS Br_Corresp W-Win _FREEFORM
  QUERY Br_Corresp DISPLAY
      tt_composant.ARTI_NUM_ARTICLE       FORMAT ">>>>>9"     COLUMN-LABEL "Article"
      tt_composant.ARTC_NUM_COMPOSANT     FORMAT ">>>>>9"     COLUMN-LABEL "Composant"
      tt_composant.ARTC_QUANTITE          FORMAT ">>>9"       COLUMN-LABEL "Qte"
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS SIZE 111.14 BY 13.5
         FONT 8 FIT-LAST-COLUMN.
 
DEFINE BROWSE Br_Erreurs
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS Br_Erreurs W-Win _FREEFORM
  QUERY Br_Erreurs DISPLAY
      tt_erreur.date_heure_erreur     FORMAT "99/99/9999 HH:MM:SS.SSS"    WIDTH 24 COLUMN-LABEL "Date/Heure!Erreur"
tt_erreur.fichier               FORMAT "X(120)"                      WIDTH 40 COLUMN-LABEL "Fichier"
tt_erreur.numLigne              FORMAT ">>>>>>"                     WIDTH 6  COLUMN-LABEL "Ligne"
tt_erreur.libErreur             FORMAT "X(256)"                     WIDTH 50 COLUMN-LABEL "Erreur"
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS SIZE 110.86 BY 8.69
         FONT 8 FIT-LAST-COLUMN.
 
DEFINE BROWSE Br_InterEber
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS Br_InterEber W-Win _FREEFORM
  QUERY Br_InterEber DISPLAY
      tt_interEber.pgie_idts      FORMAT "X(2)"   COLUMN-LABEL "IDTS":C
tt_interEber.pgie_prenom    FORMAT "X(15)"  COLUMN-LABEL "Pr‚nom":C
tt_interEber.pgie_nom       FORMAT "X(25)"  COLUMN-LABEL "Nom":C
tt_interEber.pgie_mail      FORMAT "X(40)"  COLUMN-LABEL "Mail":C
tt_interEber.pgie_tel       FORMAT "X(14)"  COLUMN-LABEL "T‚l.":C
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS SIZE 111.14 BY 14.38
         FONT 8 FIT-LAST-COLUMN.
 
DEFINE BROWSE Br_LigneOL
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS Br_LigneOL W-Win _FREEFORM
  QUERY Br_LigneOL DISPLAY
      tt_ligneOL.OL_NUM_OL              FORMAT "X(10)"      COLUMN-LABEL "Nø O.L."
      tt_ligneOL.DETAILOL_NUM_LIGNE     FORMAT ">>9"        COLUMN-LABEL "Nø Ligne"
      tt_ligneOL.ARTI_NUM_ARTICLE       FORMAT ">>>>>9"     COLUMN-LABEL "Nø Article"
      tt_ligneOL.DETAILOL_QUANTITE      FORMAT ">>>>>9"     COLUMN-LABEL "Quantit‚"
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS SIZE 111.14 BY 5.77
         FONT 8 FIT-LAST-COLUMN.
 
DEFINE BROWSE Br_OL
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS Br_OL W-Win _FREEFORM
  QUERY Br_OL DISPLAY
      tt_ol.OL_NUM_OL                   FORMAT "X(10)"  COLUMN-LABEL "Nø O.L."
      tt_ol.OL_GESTIONNAIRE_COMMANDES   FORMAT "X(2)"   COLUMN-LABEL "Gest."
      tt_ol.OL_NUM_CLIENT               FORMAT "999999" COLUMN-LABEL "Nø Client"
      tt_ol.OL_REF_CLIENT               FORMAT "X(25)"  COLUMN-LABEL "Ref. Commande Client"
      tt_ol.OL_MODE_TRANSPORT           FORMAT "X(2)"   COLUMN-LABEL "Mode!Trsp"
      tt_ol.OL_TITRE_COMMANDE           FORMAT "X(5)"   COLUMN-LABEL "Titre Commande"
      tt_ol.OL_NOM_COMMANDE             FORMAT "X(25)"  COLUMN-LABEL "Nom Commande"
      tt_ol.OL_ADRESSE1_COMMANDE        FORMAT "X(25)"  COLUMN-LABEL "Adresse Commande"
      tt_ol.OL_ADRESSE2_COMMANDE        FORMAT "X(25)"  COLUMN-LABEL "Compl‚ment adresse"
      tt_ol.OL_CP_COMMANDE              FORMAT "X(5)"   COLUMN-LABEL "CP"
      tt_ol.OL_VILLE_COMMANDE           FORMAT "X(35)"  COLUMN-LABEL "Ville Commande"
      tt_ol.OL_TITRE_LIVRAISON          FORMAT "X(5)"   COLUMN-LABEL "Titre Livraison"
      tt_ol.OL_NOM_LIVRAISON            FORMAT "X(25)"  COLUMN-LABEL "Nom Livraison"
      tt_ol.OL_ADRESSE1_LIVRAISON       FORMAT "X(25)"  COLUMN-LABEL "Adresse Livraison"
      tt_ol.OL_ADRESSE2_LIVRAISON       FORMAT "X(25)"  COLUMN-LABEL "Compl‚ment adresse"
      tt_ol.OL_CP_LIVRAISON             FORMAT "X(5)"   COLUMN-LABEL "CP"
      tt_ol.OL_VILLE_LIVRAISON          FORMAT "X(35)"  COLUMN-LABEL "Ville Livraison"
      tt_ol.OL_LOT_OBLIGATOIRE          FORMAT "X"      COLUMN-LABEL "Lot"
      tt_ol.OL_RDV_OBLIGATOIRE          FORMAT "X"      COLUMN-LABEL "RDV"
      tt_ol.OL_HAYON_OBLIGATOIRE        FORMAT "X"      COLUMN-LABEL "Hayon"
      /*
      tt_ol.OL_RDV_OBLIGATOIRE          FORMAT "O/N"    COLUMN-LABEL "RDV"
      tt_ol.OL_HAYON_OBLIGATOIRE        FORMAT "O/N"    COLUMN-LABEL "Hayon"
      */
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS SIZE 111.14 BY 5.69
         FONT 8 FIT-LAST-COLUMN.
 
DEFINE BROWSE Br_PlanTournee
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS Br_PlanTournee W-Win _FREEFORM
  QUERY Br_PlanTournee DISPLAY
      tt_plan.pgpt_codpos       FORMAT "X(5)"
      tt_plan.pgpt_cpidx        FORMAT "99"
      tt_plan.pgpt_insee        FORMAT "X(5)"
      tt_plan.pgpt_ville        FORMAT "X(35)"
      tt_plan.pgpt_vil20        FORMAT "X(20)"
      tt_plan.pgpt_pays         FORMAT "X(2)"
      tt_plan.pgpt_tournee[1]               COLUMN-LABEL "Lundi"
      tt_plan.pgpt_tournee[2]               COLUMN-LABEL "Mardi"
      tt_plan.pgpt_tournee[3]               COLUMN-LABEL "Mercredi"
      tt_plan.pgpt_tournee[4]               COLUMN-LABEL "Jeudi"
      tt_plan.pgpt_tournee[5]               COLUMN-LABEL "Vendredi"
      tt_plan.pgpt_tournee[6]               COLUMN-LABEL "Samedi"
      tt_plan.pgpt_montagne
      tt_plan.pgpt_trfpf        FORMAT "X(10)"
      tt_plan.pgpt_dept         FORMAT "X(2)"
      tt_plan.pgpt_plat         FORMAT "99"
      tt_plan.pgpt_to2j         FORMAT "9"
      tt_plan.pgpt_par          FORMAT "X(10)"
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS SIZE 111.14 BY 14.38
         FONT 8 FIT-LAST-COLUMN.
 
 
/* ************************  Frame Definitions  *********************** */
 
DEFINE FRAME F-Main
     Tg_reinitClients AT ROW 7.69 COL 81.57
     rs_visu AT ROW 7.73 COL 9.57 NO-LABEL
     Br_Articles AT ROW 8.73 COL 9.57
     Br_Clients AT ROW 8.73 COL 9.57
     Br_Corresp AT ROW 8.73 COL 9.57
     Br_InterEber AT ROW 8.73 COL 9.57
     Br_OL AT ROW 8.73 COL 9.57
     Br_PlanTournee AT ROW 8.73 COL 9.57
     ED_Commentaire AT ROW 14.46 COL 9.57 NO-LABEL
     Br_LigneOL AT ROW 17.27 COL 9.57
     liste-1 AT ROW 17.54 COL 1 NO-LABEL
     liste-3 AT ROW 17.54 COL 2.43 NO-LABEL
     liste-4 AT ROW 17.54 COL 2.43 NO-LABEL
     liste-2 AT ROW 17.54 COL 2.43 NO-LABEL
     Btn_Importer AT ROW 23.35 COL 9.57
     Btn_Test AT ROW 23.35 COL 21.86
     Btn_Quitter AT ROW 23.35 COL 111.57
     lbl_intresutl AT ROW 7.88 COL 97.57 COLON-ALIGNED NO-LABEL
    WITH 1 DOWN NO-BOX KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 1 ROW 1
         SIZE 320 BY 320
         FONT 8.
 
DEFINE FRAME FRAME-banniere
     icone-banniere AT ROW 1.19 COL 1.29
     IMAGE-banniere AT ROW 1 COL 1
     img_errImport AT ROW 12.46 COL 1.29
     img_param AT ROW 14.96 COL 1.29
    WITH 1 DOWN NO-BOX KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 1 ROW 1
         SIZE 6 BY 24.35.
 
DEFINE FRAME F-Criteres
     btn_rech_articles AT ROW 1.08 COL 102.57
     Btn_Executer AT ROW 1.08 COL 107.57
     fi_rep_integration_articles AT ROW 1.15 COL 17.86 COLON-ALIGNED
     btn_rech_OL AT ROW 2.27 COL 102.57
     fi_rep_integration_ol AT ROW 2.35 COL 17.86 COLON-ALIGNED
     btn_rech_Param AT ROW 3.46 COL 102.57
     fi_rep_integration_param AT ROW 3.54 COL 17.86 COLON-ALIGNED
     btn_rech_svg AT ROW 4.65 COL 102.57
     fi_rep_svg AT ROW 4.77 COL 17.86 COLON-ALIGNED
     fi_email AT ROW 5.96 COL 17.86 COLON-ALIGNED
    WITH 1 DOWN NO-BOX KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 7.14 ROW 1.08
         SIZE 116.14 BY 5.85
         FONT 8.
 
DEFINE FRAME F-jauge
    WITH 1 DOWN NO-BOX KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 7.14 ROW 1.04
         SIZE 116.14 BY 5.85.
 
DEFINE FRAME F-Erreurs
     Br_Erreurs AT ROW 1.96 COL 2.14
     Btn_Fermer AT ROW 11.46 COL 47.57
     Btn_Excel AT ROW 11.46 COL 58.86
    WITH 1 DOWN KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 8.29 ROW 7.23
         SIZE 114 BY 12.85
         FONT 8
         TITLE "Erreurs d'importation".
 
DEFINE FRAME F-PARAM
     ed_lstExclusions AT ROW 2.15 COL 6.86 NO-LABEL
     Btn_param_ok AT ROW 22.15 COL 23.86
     Btn_param_annuler AT ROW 22.15 COL 35.57
     "Liste des OL … exclure de l'int‚gration ( un par ligne )" VIEW-AS TEXT
          SIZE 43.72 BY .73 AT ROW 1.38 COL 7.14
          FONT 10
    WITH 1 DOWN KEEP-TAB-ORDER 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 27.72 ROW 1.15
         SIZE 66.43 BY 23.38
         FONT 8
         TITLE "ParamŠtres".
 
 
/* *********************** Procedure Settings ************************ */
 
&ANALYZE-SUSPEND _PROCEDURE-SETTINGS
/* Settings for THIS-PROCEDURE
   Type: SmartWindow
   Allow: Basic,Browse,DB-Fields,Query,Smart,Window
   Other Settings: COMPILE
 */
&ANALYZE-RESUME _END-PROCEDURE-SETTINGS
 
/* *************************  Create Window  ************************** */
 
&ANALYZE-SUSPEND _CREATE-WINDOW
IF SESSION:DISPLAY-TYPE = "GUI":U THEN
  CREATE WINDOW W-Win ASSIGN
         HIDDEN             = YES
         TITLE              = "Import PGI"
         HEIGHT             = 24.23
         WIDTH              = 122.86
         MAX-HEIGHT         = 320
         MAX-WIDTH          = 320
         VIRTUAL-HEIGHT     = 320
         VIRTUAL-WIDTH      = 320
         RESIZE             = YES
         SCROLL-BARS        = NO
         STATUS-AREA        = NO
         BGCOLOR            = ?
         FGCOLOR            = ?
         THREE-D            = YES
         MESSAGE-AREA       = NO
         SENSITIVE          = YES.
ELSE {&WINDOW-NAME} = CURRENT-WINDOW.
/* END WINDOW DEFINITION                                                */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _INCLUDED-LIB W-Win 
/* ************************* Included-Libraries *********************** */
 
{src/adm/method/window.i}
{src/adm/method/containr.i}
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
 
 
/* ***********  Runtime Attributes and AppBuilder Settings  *********** */
 
&ANALYZE-SUSPEND _RUN-TIME-ATTRIBUTES
/* SETTINGS FOR WINDOW W-Win
  VISIBLE,,RUN-PERSISTENT                                               */
/* REPARENT FRAME */
ASSIGN FRAME F-Criteres:FRAME = FRAME F-Main:HANDLE
       FRAME F-Erreurs:FRAME = FRAME F-Main:HANDLE
       FRAME F-jauge:FRAME = FRAME F-Main:HANDLE
       FRAME F-PARAM:FRAME = FRAME F-Main:HANDLE
       FRAME FRAME-banniere:FRAME = FRAME F-Main:HANDLE.
 
/* SETTINGS FOR FRAME F-Criteres
                                                                        */
/* SETTINGS FOR FRAME F-Erreurs
                                                                        */
/* BROWSE-TAB Br_Erreurs 1 F-Erreurs */
ASSIGN 
       Br_Erreurs:COLUMN-RESIZABLE IN FRAME F-Erreurs       = TRUE.
 
/* SETTINGS FOR FRAME F-jauge
                                                                        */
/* SETTINGS FOR FRAME F-Main
   FRAME-NAME                                                           */
 
DEFINE VARIABLE XXTABVALXX AS LOGICAL NO-UNDO.
 
ASSIGN XXTABVALXX = FRAME F-Erreurs:MOVE-BEFORE-TAB-ITEM (Tg_reinitClients:HANDLE IN FRAME F-Main)
       XXTABVALXX = FRAME F-PARAM:MOVE-BEFORE-TAB-ITEM (FRAME F-Erreurs:HANDLE)
       XXTABVALXX = FRAME F-Criteres:MOVE-BEFORE-TAB-ITEM (FRAME F-PARAM:HANDLE)
       XXTABVALXX = FRAME F-jauge:MOVE-BEFORE-TAB-ITEM (FRAME F-Criteres:HANDLE)
       XXTABVALXX = FRAME FRAME-banniere:MOVE-BEFORE-TAB-ITEM (FRAME F-jauge:HANDLE)
/* END-ASSIGN-TABS */.
 
/* BROWSE-TAB Br_Articles rs_visu F-Main */
/* BROWSE-TAB Br_Clients Br_Articles F-Main */
/* BROWSE-TAB Br_Corresp Br_Clients F-Main */
/* BROWSE-TAB Br_InterEber Br_Corresp F-Main */
/* BROWSE-TAB Br_OL Br_InterEber F-Main */
/* BROWSE-TAB Br_PlanTournee Br_OL F-Main */
/* BROWSE-TAB Br_LigneOL ED_Commentaire F-Main */
ASSIGN 
       Br_Articles:COLUMN-RESIZABLE IN FRAME F-Main       = TRUE.
 
ASSIGN 
       Br_Clients:COLUMN-RESIZABLE IN FRAME F-Main       = TRUE.
 
ASSIGN 
       Br_Corresp:COLUMN-RESIZABLE IN FRAME F-Main       = TRUE.
 
ASSIGN 
       Br_InterEber:COLUMN-RESIZABLE IN FRAME F-Main       = TRUE.
 
ASSIGN 
       Br_LigneOL:COLUMN-RESIZABLE IN FRAME F-Main       = TRUE.
 
ASSIGN 
       Br_OL:COLUMN-RESIZABLE IN FRAME F-Main       = TRUE.
 
ASSIGN 
       Br_PlanTournee:COLUMN-RESIZABLE IN FRAME F-Main       = TRUE.
 
ASSIGN 
       ED_Commentaire:READ-ONLY IN FRAME F-Main        = TRUE.
 
/* SETTINGS FOR FRAME F-PARAM
   NOT-VISIBLE                                                          */
ASSIGN 
       FRAME F-PARAM:HIDDEN           = TRUE
       FRAME F-PARAM:SENSITIVE        = FALSE.
 
/* SETTINGS FOR FRAME FRAME-banniere
                                                                        */
ASSIGN 
       FRAME FRAME-banniere:HIDDEN           = TRUE.
 
IF SESSION:DISPLAY-TYPE = "GUI":U AND VALID-HANDLE(W-Win)
THEN W-Win:HIDDEN = YES.
 
/* _RUN-TIME-ATTRIBUTES-END */
&ANALYZE-RESUME
 
 
/* Setting information for Queries and Browse Widgets fields            */
 
&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE Br_Articles
/* Query rebuild information for BROWSE Br_Articles
     _START_FREEFORM
OPEN QUERY {&SELF-NAME} FOR EACH tt_article BY tt_article.tt_fichier BY tt_article.tt_numligne.
     _END_FREEFORM
     _Query            is OPENED
*/  /* BROWSE Br_Articles */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE Br_Clients
/* Query rebuild information for BROWSE Br_Clients
     _START_FREEFORM
/*
OPEN QUERY {&SELF-NAME} FOR EACH tt_client BY tt_client.tt_fichier BY tt_client.tt_numligne.
*/
OPEN QUERY {&SELF-NAME} FOR EACH tt_client BY tt_client.tt_fichier BY tt_client.cli_num_client BY tt_client.cli_cp.
     _END_FREEFORM
     _Query            is OPENED
*/  /* BROWSE Br_Clients */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE Br_Corresp
/* Query rebuild information for BROWSE Br_Corresp
     _START_FREEFORM
OPEN QUERY {&SELF-NAME} FOR EACH tt_composant BY tt_composant.tt_fichier BY tt_composant.tt_numligne.
     _END_FREEFORM
     _Query            is OPENED
*/  /* BROWSE Br_Corresp */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE Br_Erreurs
/* Query rebuild information for BROWSE Br_Erreurs
     _START_FREEFORM
OPEN QUERY {&SELF-NAME} FOR EACH tt_erreur BY tt_erreur.date_heure_erreur.
     _END_FREEFORM
     _Query            is OPENED
*/  /* BROWSE Br_Erreurs */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE Br_InterEber
/* Query rebuild information for BROWSE Br_InterEber
     _START_FREEFORM
OPEN QUERY {&SELF-NAME} FOR EACH tt_interEber BY tt_interEber.tt_fichier BY tt_interEber.tt_numligne.
     _END_FREEFORM
     _Query            is OPENED
*/  /* BROWSE Br_InterEber */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE Br_LigneOL
/* Query rebuild information for BROWSE Br_LigneOL
     _START_FREEFORM
OPEN QUERY {&SELF-NAME} FOR EACH tt_ligneOL WHERE tt_ligneOL.ol_num_ol = tt_ol.ol_num_ol AND tt_ligneOL.tt_fichier = tt_ol.tt_fichier BY tt_ligneOL.tt_fichier BY tt_ligneOL.tt_numligne.
     _END_FREEFORM
     _Query            is OPENED
*/  /* BROWSE Br_LigneOL */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE Br_OL
/* Query rebuild information for BROWSE Br_OL
     _START_FREEFORM
OPEN QUERY {&SELF-NAME} FOR EACH tt_ol BY tt_OL.tt_fichier BY tt_OL.tt_numligne.
     _END_FREEFORM
     _Query            is OPENED
*/  /* BROWSE Br_OL */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE Br_PlanTournee
/* Query rebuild information for BROWSE Br_PlanTournee
     _START_FREEFORM
OPEN QUERY {&SELF-NAME} FOR EACH tt_plan BY tt_plan.tt_fichier BY tt_plan.tt_numligne.
     _END_FREEFORM
     _Query            is OPENED
*/  /* BROWSE Br_PlanTournee */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _QUERY-BLOCK FRAME F-PARAM
/* Query rebuild information for FRAME F-PARAM
     _Query            is NOT OPENED
*/  /* FRAME F-PARAM */
&ANALYZE-RESUME
 
 
 
 
 
/* ************************  Control Triggers  ************************ */
 
&Scoped-define SELF-NAME W-Win
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL W-Win W-Win
ON END-ERROR OF W-Win /* Import PGI */
OR ENDKEY OF {&WINDOW-NAME} ANYWHERE DO:
  /* This case occurs when the user presses the "Esc" key.
     In a persistently run window, just ignore this.  If we did not, the
     application would exit. */
  IF THIS-PROCEDURE:PERSISTENT THEN RETURN NO-APPLY.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL W-Win W-Win
ON WINDOW-CLOSE OF W-Win /* Import PGI */
DO:
  /* This ADM code must be left here in order for the SmartWindow
     and its descendents to terminate properly on exit. */
  RUN dispatch IN THIS-PROCEDURE ('exit':U).
  RETURN NO-APPLY.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL W-Win W-Win
ON WINDOW-MAXIMIZED OF W-Win /* Import PGI */
DO:
    plein-ecran = TRUE.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL W-Win W-Win
ON WINDOW-RESIZED OF W-Win /* Import PGI */
DO:
    RUN RESIZE.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL W-Win W-Win
ON WINDOW-RESTORED OF W-Win /* Import PGI */
DO:
    plein-ecran = FALSE.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define BROWSE-NAME Br_OL
&Scoped-define SELF-NAME Br_OL
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL Br_OL W-Win
ON VALUE-CHANGED OF Br_OL IN FRAME F-Main
DO:
    {&OPEN-QUERY-Br_LigneOL}
    IF AVAILABLE tt_ol THEN
        ASSIGN ED_Commentaire:SCREEN-VALUE = tt_ol.OL_COMMENTAIRE.
    ELSE
        ASSIGN ED_Commentaire:SCREEN-VALUE = "".
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define FRAME-NAME F-Erreurs
&Scoped-define SELF-NAME Btn_Excel
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL Btn_Excel W-Win
ON CHOOSE OF Btn_Excel IN FRAME F-Erreurs /* Excel */
DO:
    DEFINE VARIABLE lc_fichier AS CHARACTER NO-UNDO.    /* On ne va pas s'en servir */ 
 
    RUN edite_browse (OUTPUT lc_fichier).
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define FRAME-NAME F-Criteres
&Scoped-define SELF-NAME Btn_Executer
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL Btn_Executer W-Win
ON CHOOSE OF Btn_Executer IN FRAME F-Criteres /* Executer */
DO:
    DISABLE Btn_Importer WITH FRAME F-Main.
 
    EMPTY TEMP-TABLE tt_erreur.
 
    gi_cpt = 0.
 
    /* Initialise les tables client */
    RUN initTTClient.
 
    RUN integrationArticles(fi_rep_integration_articles:HANDLE IN FRAME F-Criteres) NO-ERROR.
    IF ERROR-STATUS:ERROR THEN
        RETURN NO-APPLY.
 
    RUN integrationOL(fi_rep_integration_OL:HANDLE IN FRAME F-Criteres) NO-ERROR.
    IF ERROR-STATUS:ERROR THEN
        RETURN NO-APPLY.
 
    RUN integrationParam(fi_rep_integration_param:HANDLE IN FRAME F-Criteres) NO-ERROR.
    IF ERROR-STATUS:ERROR THEN
        RETURN NO-APPLY.
 
    APPLY "VALUE-CHANGED" TO rs_visu.
 
    IF CAN-FIND(FIRST tt_article) OR CAN-FIND(FIRST tt_composant) OR CAN-FIND(FIRST tt_ol) OR CAN-FIND(FIRST tt_plan) OR CAN-FIND(FIRST tt_interEber) OR CAN-FIND(FIRST tt_client) OR CAN-FIND(FIRST tt_cliint) THEN
        ENABLE Btn_Importer WITH FRAME F-Main.
    ELSE
        DISABLE Btn_Importer WITH FRAME F-Main.
 
    IF NOT gl_modeBatch AND CAN-FIND(FIRST tt_erreur) THEN
        RUN afficheErreurs(TRUE).
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define FRAME-NAME F-Erreurs
&Scoped-define SELF-NAME Btn_Fermer
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL Btn_Fermer W-Win
ON CHOOSE OF Btn_Fermer IN FRAME F-Erreurs /* Fermer */
DO:
    RUN afficheErreurs(FALSE).
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define FRAME-NAME F-Main
&Scoped-define SELF-NAME Btn_Importer
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL Btn_Importer W-Win
ON CHOOSE OF Btn_Importer IN FRAME F-Main /* Importer */
DO:
    /* Traite les tables temporaires : Articles, composants, O.L., Detail OL, plan de tourn‚e, intervenants Eberhardt, clients, etc... et g‚nŠre les donn‚es en base */
    RUN traitement NO-ERROR.
    IF ERROR-STATUS:ERROR THEN
        RETURN NO-APPLY.
 
    APPLY "VALUE-CHANGED" TO rs_visu.
 
    IF NOT gl_modeBatch AND CAN-FIND(FIRST tt_erreur) THEN
        RUN afficheErreurs(TRUE).
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define FRAME-NAME F-PARAM
&Scoped-define SELF-NAME Btn_param_annuler
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL Btn_param_annuler W-Win
ON CHOOSE OF Btn_param_annuler IN FRAME F-PARAM /* Annuler */
DO:
    DISPLAY ed_lstExclusions WITH FRAME F-Param.
    FRAME F-Param:HIDDEN = TRUE.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define SELF-NAME Btn_param_ok
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL Btn_param_ok W-Win
ON CHOOSE OF Btn_param_ok IN FRAME F-PARAM /* OK */
DO:
    ASSIGN ed_lstExclusions.
 
    RUN sauve_param.
    RUN create_TT_OL_EXCLUS.
 
    FRAME F-Param:HIDDEN = TRUE.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define FRAME-NAME F-Main
&Scoped-define SELF-NAME Btn_Quitter
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL Btn_Quitter W-Win
ON CHOOSE OF Btn_Quitter IN FRAME F-Main /* Quitter */
DO:
    RUN dispatch ('exit').
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define FRAME-NAME F-Criteres
&Scoped-define SELF-NAME btn_rech_articles
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL btn_rech_articles W-Win
ON CHOOSE OF btn_rech_articles IN FRAME F-Criteres
DO:
    DEF VAR OK  AS  LOG     NO-UNDO.
    DEF VAR rep AS  CHAR    NO-UNDO.
 
    /* rep = trim(fi_rep_integration:screen-value,"\"). */
    rep = fi_rep_integration_articles:SCREEN-VALUE IN FRAME F-Criteres.
 
    RUN Objets\files\files-tools.r ( INPUT "select-directory",
                                     INPUT CURRENT-WINDOW:HANDLE,
                                     INPUT "S‚lectionner le r‚pertoire d'int‚gration des articles...",
                                     INPUT rep,
                                     INPUT "",
                                     OUTPUT rep,
                                     OUTPUT OK ).
 
    IF OK THEN
    DO:
        fi_rep_integration_articles = rep.
        DISPLAY fi_rep_integration_articles WITH FRAME F-Criteres.
    END.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define SELF-NAME btn_rech_OL
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL btn_rech_OL W-Win
ON CHOOSE OF btn_rech_OL IN FRAME F-Criteres
DO:
    DEF VAR OK  AS  LOG     NO-UNDO.
    DEF VAR rep AS  CHAR    NO-UNDO.
 
    /* rep = trim(fi_rep_integration:screen-value,"\"). */
    rep = fi_rep_integration_ol:SCREEN-VALUE IN FRAME F-Criteres.
 
    RUN Objets\files\files-tools.r ( INPUT "select-directory",
                                     INPUT CURRENT-WINDOW:HANDLE,
                                     INPUT "S‚lectionner le r‚pertoire d'int‚gration des OL ...",
                                     INPUT rep,
                                     INPUT "",
                                     OUTPUT rep,
                                     OUTPUT OK ).
 
    IF OK THEN
    DO:
        fi_rep_integration_ol = rep.
        DISPLAY fi_rep_integration_ol WITH FRAME F-Criteres.
    END.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define SELF-NAME btn_rech_Param
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL btn_rech_Param W-Win
ON CHOOSE OF btn_rech_Param IN FRAME F-Criteres
DO:
    DEF VAR OK  AS  LOG     NO-UNDO.
    DEF VAR rep AS  CHAR    NO-UNDO.
 
    /* rep = trim(fi_rep_integration:screen-value,"\"). */
    rep = fi_rep_integration_param:SCREEN-VALUE IN FRAME F-Criteres.
 
    RUN Objets\files\files-tools.r ( INPUT "select-directory",
                                     INPUT CURRENT-WINDOW:HANDLE,
                                     INPUT "S‚lectionner le r‚pertoire d'int‚gration des paramŠtres ...",
                                     INPUT rep,
                                     INPUT "",
                                     OUTPUT rep,
                                     OUTPUT OK ).
 
    IF OK THEN
    DO:
        fi_rep_integration_param = rep.
        DISPLAY fi_rep_integration_param WITH FRAME F-Criteres.
    END.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define SELF-NAME btn_rech_svg
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL btn_rech_svg W-Win
ON CHOOSE OF btn_rech_svg IN FRAME F-Criteres
DO:
    DEF VAR OK  AS  LOG     NO-UNDO.
    DEF VAR rep AS  CHAR    NO-UNDO.
 
    rep = fi_rep_svg:SCREEN-VALUE.
 
    RUN Objets\files\files-tools.r ( INPUT "select-directory",
                                     INPUT CURRENT-WINDOW:HANDLE,
                                     INPUT "S‚lectionner le r‚pertoire d'int‚gration des articles...",
                                     INPUT rep,
                                     INPUT "",
                                     OUTPUT rep,
                                     OUTPUT OK ).
 
    IF OK THEN
    DO:
        fi_rep_svg = rep.
        DISPLAY fi_rep_svg WITH FRAME F-Criteres.
    END.
 
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define FRAME-NAME F-Main
&Scoped-define SELF-NAME Btn_Test
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL Btn_Test W-Win
ON CHOOSE OF Btn_Test IN FRAME F-Main /* Test Repart */
DO:
    DEFINE VARIABLE li_jauge AS INTEGER     NO-UNDO.
 
    SELECT COUNT(*) INTO li_jauge FROM gpi_pgol WHERE GPI_PGOL.ol_date_dernier_envoi_eberhardt = TODAY AND NOT gpi_pgol.ol_a_supprimer.
 
    /* Initialise les premiŠre dates de livraison PF possibles */
    RUN jauge-set-libelle IN h_jauge-ocx (INPUT "R‚partition Lot/PF. Veuillez patienter ...").
    RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
    /* Une fois l'import termin², on lance la r²partition PF/Lot */
    RUN specif\PGI-EBERHARDT\p-PF_ou_lot.p.
 
    /* Initialise les premiŠre dates de livraison PF possibles */
    RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Initialisation plateformes. Veuillez patienter ...").
    RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
    RUN initPF.
 
    /* Calcul de la date de livraison pour le OL affect‚s en PF */
    RUN initDateLivraison.
 
    /* Initialisation du statut Obligatoire pour les OL affect‚s en PF */
    RUN majObligatoire.
 
    RUN jauge-fin IN h_jauge-ocx.
 
    FRAME F-jauge:MOVE-TO-BOTTOM().
    FRAME F-jauge:VISIBLE = FALSE.
 
    FRAME F-Criteres:VISIBLE = TRUE.
    FRAME F-Criteres:MOVE-TO-TOP().
 
    RETURN "".
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define FRAME-NAME F-Criteres
&Scoped-define SELF-NAME fi_email
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL fi_email W-Win
ON LEAVE OF fi_email IN FRAME F-Criteres /* Email supervision */
DO:
    ASSIGN fi_email.
    IF fi_email <> "" AND NOT fi_listeMailsValide(fi_email) THEN
    DO:
        MESSAGE "Merci de saisir une ou des adresses mail valides, s‚par‚es par , ou ; " VIEW-AS ALERT-BOX WARNING.
        RETURN NO-APPLY.
    END.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define SELF-NAME fi_rep_integration_articles
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL fi_rep_integration_articles W-Win
ON LEAVE OF fi_rep_integration_articles IN FRAME F-Criteres /* R‚pertoire Articles */
DO:
    DEFINE VARIABLE lok AS LOGICAL    NO-UNDO.
    ASSIGN fi_rep_integration_articles.
    IF fi_rep_integration_articles <> "" THEN
    DO:    
        RUN Objets\Files\verifie_repertoire.r ( INPUT fi_rep_integration_articles, INPUT TRUE, INPUT " des fichiers d'int‚gration", INPUT " acc‚der aux fichiers", INPUT TRUE, OUTPUT lOK ).
        IF NOT lok THEN
            RETURN NO-APPLY.
    END.
 
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define SELF-NAME fi_rep_integration_ol
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL fi_rep_integration_ol W-Win
ON LEAVE OF fi_rep_integration_ol IN FRAME F-Criteres /* R‚pertoire OL */
DO:
    DEFINE VARIABLE lok AS LOGICAL    NO-UNDO.
    ASSIGN fi_rep_integration_ol.
    IF fi_rep_integration_ol <> "" THEN
    DO:    
        RUN Objets\Files\verifie_repertoire.r ( INPUT fi_rep_integration_ol, INPUT TRUE, INPUT " des fichiers d'int‚gration", INPUT " acc‚der aux fichiers", INPUT TRUE, OUTPUT lOK ).
        IF NOT lok THEN
            RETURN NO-APPLY.
    END.
 
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define SELF-NAME fi_rep_integration_param
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL fi_rep_integration_param W-Win
ON LEAVE OF fi_rep_integration_param IN FRAME F-Criteres /* R‚pertoire ParamŠtres */
DO:
    DEFINE VARIABLE lok AS LOGICAL    NO-UNDO.
    ASSIGN fi_rep_integration_param.
    IF fi_rep_integration_param <> "" THEN
    DO:    
        RUN Objets\Files\verifie_repertoire.r ( INPUT fi_rep_integration_param, INPUT TRUE, INPUT " des fichiers d'int‚gration", INPUT " acc‚der aux fichiers", INPUT TRUE, OUTPUT lOK ).
        IF NOT lok THEN
            RETURN NO-APPLY.
    END.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define SELF-NAME fi_rep_svg
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL fi_rep_svg W-Win
ON LEAVE OF fi_rep_svg IN FRAME F-Criteres /* R‚pertoire Sauvegarde */
DO:
    DEFINE VARIABLE lok AS LOGICAL    NO-UNDO.
 
    ASSIGN fi_rep_svg.
    IF fi_rep_svg <> "" THEN
    DO:    
        RUN Objets\Files\verifie_repertoire.r ( INPUT fi_rep_svg, INPUT TRUE, INPUT " des fichiers d'int‚gration", INPUT " acc‚der aux fichiers", INPUT TRUE, OUTPUT lOK ).
        IF NOT lok THEN
            RETURN NO-APPLY.
    END.
 
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define FRAME-NAME FRAME-banniere
&Scoped-define SELF-NAME icone-banniere
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL icone-banniere W-Win
ON MOUSE-SELECT-DBLCLICK OF icone-banniere IN FRAME FRAME-banniere
DO:
DEF VAR card AS CHAR NO-UNDO.
 
    /*------- 4 paramŠtres (chr(9)) pouvant ˆtre renseign‚s ---------------------------------------------------------------
 
 
        PROGRAMME    = Nom du programme en cours (automatique avec entry(2,PROGRAM-NAME(1)," ")
 
        VERSION      = Indication sur la version (date indice ...)
 
        PARAMETER    = Liste (chr(10)) des gpi_par utilis‚s dans ce programme suivis pour chacun du type entre parenthŠses
                       Ex: "PARAMETER=gpi_gcout:size (utilisateur)" + chr(10) + "Autre param (Soci‚t‚)" ...
 
        LOCALISATION = A partir de quelle fonctionnalit‚ est lanc‚ le programme.
                       Ex: LOCALISATION=Gestion des tiers" + chr(10) + "Saisie" ...
 
        DESCRIPTION  = Description succinte du programme.
 
    ---------------------------------------------------------------------------------------------------------------------*/
 
    card = "PROGRAMME="     + entry(2,PROGRAM-NAME(1)," ") + CHR(9) + 
           "VERSION=Nø 140-120-142-523"                    + CHR(9) +
           "PARAMETER="                                    + CHR(9) +
           "LOCALISATION="                                 + chr(9) +
           "DESCRIPTION="   + w-win:TITLE.
 
    FILE-INFO:FILE-NAME = "objets\FILEs\file-card.r".
    IF FILE-INFO:FULL-PATHNAME <> ? THEN
        RUN objets\FILEs\file-card.r (card).
    ELSE
        MESSAGE ENTRY(2,PROGRAM-NAME(1)," ") VIEW-AS ALERT-BOX INFORMATION TITLE "Programme en cours".
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define SELF-NAME img_errImport
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL img_errImport W-Win
ON MOUSE-SELECT-CLICK OF img_errImport IN FRAME FRAME-banniere
DO:
    IF FRAME F-Erreurs:VISIBLE THEN
        RUN afficheErreurs (FALSE).
    ELSE
        RUN afficheErreurs (TRUE).
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define SELF-NAME img_param
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL img_param W-Win
ON MOUSE-SELECT-CLICK OF img_param IN FRAME FRAME-banniere
DO:
    FRAME F-Param:VISIBLE = TRUE.
    FRAME F-Param:MOVE-TO-TOP().
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define FRAME-NAME F-Main
&Scoped-define SELF-NAME rs_visu
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL rs_visu W-Win
ON VALUE-CHANGED OF rs_visu IN FRAME F-Main
DO:
    ASSIGN rs_visu.
    HIDE Br_Articles Br_Corresp Br_OL Br_LigneOL ED_Commentaire BR_PlanTournee Br_InterEber Br_Clients TG_reinitClients IN FRAME F-Main.
    CASE rs_visu:
        WHEN "A" THEN
        DO:
            VIEW Br_Articles IN FRAME F-Main.
            {&OPEN-QUERY-Br_Articles}
        END.
        WHEN "C" THEN
        DO:
            VIEW Br_Corresp IN FRAME F-Main.
            {&OPEN-QUERY-Br_Corresp}
        END.
        WHEN "O" THEN
        DO:
            VIEW Br_OL Br_LigneOL ED_Commentaire IN FRAME F-Main.
            {&OPEN-QUERY-Br_OL}
            APPLY "VALUE-CHANGED" TO Br_OL.
        END.
        WHEN "P" THEN
        DO:
            VIEW Br_PlanTournee IN FRAME F-Main.
            {&OPEN-QUERY-Br_PlanTournee}
        END.
        WHEN "IE" THEN
        DO:
            VIEW Br_InterEber IN FRAME F-Main.
            {&OPEN-QUERY-Br_InterEber}
        END.
        WHEN "Cl" THEN
        DO:
            VIEW Br_Clients Tg_reinitClients IN FRAME F-Main.
            {&OPEN-QUERY-Br_Clients}
        END.
    END CASE.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define SELF-NAME Tg_reinitClients
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL Tg_reinitClients W-Win
ON VALUE-CHANGED OF Tg_reinitClients IN FRAME F-Main /* R‚initialisation clients */
DO:
    ASSIGN Tg_reinitClients.
END.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&Scoped-define BROWSE-NAME Br_Articles
&UNDEFINE SELF-NAME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _MAIN-BLOCK W-Win 
 
 
/*----- TRIGGER CREE PAR LE BLOCAGE - NE PAS MODIFIER / DEPLACER -----*/
ON ALT-F1 OF {&WINDOW-NAME}
ANYWHERE DO:
RUN appli\version.r ( "Programme : " + ENTRY ( 2, PROGRAM-NAME ( 1 ), " " ) + CHR(10) + CHR(10) + "Version = Nø 140-120-142-523" ). /* Nicolas */
END.
/*----- TRIGGER CREE PAR LE BLOCAGE - NE PAS MODIFIER / DEPLACER -----*/
/* ***************************  Main Block  *************************** */
 
    ASSIGN  w-win:X = 0
            w-win:Y = 0
            .
 
    Nom-Programme = PROGRAM-NAME(1).
    IF NUM-ENTRIES ( Nom-Programme, " " ) > 1 THEN
            Nom-Programme = ENTRY ( 2, Nom-Programme, " " ).
 
    ASSIGN  Nom-Programme = REPLACE ( Nom-Programme, "/", "\" )
            Nom-Programme = ENTRY ( NUM-ENTRIES ( Nom-Programme, "\" ), Nom-Programme, "\" )
            .
 
    {objets\apiwindows\resize-auto.i}
 
    IF NOT AVAILABLE GPI_CNF THEN
        gl_modeBatch = TRUE.
 
    /*
        Image en haut de la banniŠre … modifier    
    */
    FILE-INFO:FILE-NAME = "image\importation.bmp".
    IF FILE-INFO:FULL-PATHNAME <> ? THEN
        icone-banniere:LOAD-IMAGE(FILE-INFO:FULL-PATHNAME) IN FRAME frame-banniere.
    ELSE
        icone-banniere:LOAD-IMAGE("Bourse\images\planning\find.bmp") IN FRAME frame-banniere.
    icone-banniere:MOVE-TO-TOP().
 
 
    FILE-INFO:FILE-NAME = "image\param.bmp".
    IF FILE-INFO:FULL-PATHNAME <> ? THEN
        img_param:LOAD-IMAGE(FILE-INFO:FULL-PATHNAME) IN FRAME frame-banniere.
    ELSE
        img_param:LOAD-IMAGE("Bourse\images\planning\find.bmp") IN FRAME frame-banniere.
    img_param:MOVE-TO-TOP().
 
 
    FILE-INFO:FILE-NAME = "image\err_import.bmp".
    IF FILE-INFO:FULL-PATHNAME <> ? THEN
        img_errImport:LOAD-IMAGE(FILE-INFO:FULL-PATHNAME) IN FRAME frame-banniere.
    ELSE
        img_errImport:LOAD-IMAGE("Bourse\images\planning\find.bmp") IN FRAME frame-banniere.
    img_errImport:MOVE-TO-TOP().
 
 
    IF AVAILABLE GPI_CNF THEN
        W-Win:TITLE = W-Win:TITLE + " --- Utilisateur " + Menu_Utilisateur + " --- Agence " + GPI_CNF.cnf_nom + " ---".
 
 
    ASSIGN
    W-Win:MIN-HEIGHT = W-Win:HEIGHT
    W-Win:MIN-WIDTH  = W-Win:WIDTH.
 
    /* Include custom  Main Block code for SmartWindows. */
    {src/adm/template/windowmn.i}
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
/* **********************  Internal Procedures  *********************** */
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE adm-create-objects W-Win  _ADM-CREATE-OBJECTS
PROCEDURE adm-create-objects :
/*------------------------------------------------------------------------------
  Purpose:     Create handles for all SmartObjects used in this procedure.
               After SmartObjects are initialized, then SmartLinks are added.
  Parameters:  <none>
------------------------------------------------------------------------------*/
  DEFINE VARIABLE adm-current-page  AS INTEGER NO-UNDO.
 
  RUN get-attribute IN THIS-PROCEDURE ('Current-Page':U).
  ASSIGN adm-current-page = INTEGER(RETURN-VALUE).
 
  CASE adm-current-page: 
 
    WHEN 0 THEN DO:
       RUN init-object IN THIS-PROCEDURE (
             INPUT  'objets/jauge/jauge-ocx.r':U ,
             INPUT  FRAME F-jauge:HANDLE ,
             INPUT  'backcolor = 255|255|255,
                     barcolor = 095|230|095,
                     Fontcolor = 000|000|000,
                     alignement = 2,
                     bordure = 1,
                     captiontyp = 3,
                     dotwidth = 10,
                     GapWidth = 2,
                     caption = Votre texte,
                     cacher_100 = N,
                     cachee = N,
                     HORIZONTAL = O,
                     Marges = 2,
                     attente = N':U ,
             OUTPUT h_jauge-ocx ).
       RUN set-position IN h_jauge-ocx ( 2.38 , 1.72 ) NO-ERROR.
       RUN set-size IN h_jauge-ocx ( 2.77 , 114.72 ) NO-ERROR.
 
       /* Adjust the tab order of the smart objects. */
    END. /* Page 0 */
 
  END CASE.
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE adm-row-available W-Win  _ADM-ROW-AVAILABLE
PROCEDURE adm-row-available :
/*------------------------------------------------------------------------------
  Purpose:     Dispatched to this procedure when the Record-
               Source has a new row available.  This procedure
               tries to get the new row (or foriegn keys) from
               the Record-Source and process it.
  Parameters:  <none>
------------------------------------------------------------------------------*/
 
  /* Define variables needed by this internal procedure.             */
  {src/adm/template/row-head.i}
 
  /* Process the newly available records (i.e. display fields,
     open queries, and/or pass records on to any RECORD-TARGETS).    */
  {src/adm/template/row-end.i}
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE afficheErreurs W-Win 
PROCEDURE afficheErreurs :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT PARAMETER il_affiche AS LOGICAL NO-UNDO.   /* True pour l'afficher, false pour la masquer */
 
IF il_affiche THEN
DO:
    /* D‚sactive la frame principale et affiche la frame des erreurs */
    DISABLE ALL WITH FRAME F-Main.
    ASSIGN
    FRAME F-Erreurs:VISIBLE = TRUE
    FRAME F-Erreurs:SENSITIVE = TRUE.
    FRAME F-Erreurs:MOVE-TO-TOP().
    ENABLE ALL WITH FRAME F-Erreurs.
 
    {&OPEN-QUERY-Br_Erreurs}
END.
ELSE
DO:
    /* D‚sactive et masque la frame des erreur, et active la frame principale */
    DISABLE ALL WITH FRAME F-Erreurs.
    ASSIGN
    FRAME F-Erreurs:SENSITIVE = FALSE
    FRAME F-Erreurs:VISIBLE = FALSE.
 
    ENABLE ALL WITH FRAME F-Main.
    APPLY "VALUE-CHANGED" TO rs_visu.
 
    IF CAN-FIND(FIRST tt_article) OR CAN-FIND(FIRST tt_composant) OR CAN-FIND(FIRST tt_ol) OR CAN-FIND(FIRST tt_plan) OR CAN-FIND(FIRST tt_interEber) OR CAN-FIND(FIRST tt_client) OR CAN-FIND(FIRST tt_cliint) THEN
        ENABLE Btn_Importer WITH FRAME F-Main.
    ELSE
        DISABLE Btn_Importer WITH FRAME F-Main.
END.
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE controleArticles W-Win 
PROCEDURE controleArticles :
/*------------------------------------------------------------------------------
  Purpose:     Contr“le la table d'articles int‚gr‚es pour v‚rifier qu'il n'y a 
               pas d'incoh‚rences
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ii_numArticles AS INTEGER     NO-UNDO.
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Contr“le des articles en cours. Veuillez patienter ...").
 
RUN jauge-init IN h_jauge-ocx ( INPUT ii_numArticles ).
 
FOR EACH bf_article BY bf_article.tt_fichier BY bf_article.tt_numligne: /* Traite les articles dans leur ordre d'‚mission */
    RUN jauge-next IN h_jauge-ocx.
    CASE bf_article.tt_action:
        WHEN "C" THEN
        DO:
            /* On v‚rifie que l'article n'existe pas, sois en base, sois plus haut dans les fichiers */
            /* S'il existe, cela deviens une modification */
            FIND FIRST GPI_PGARTI WHERE GPI_PGARTI.ARTI_NUM_ARTICLE = bf_article.ARTI_NUM_ARTICLE AND GPI_PGARTI.ARTI_DATE-HEURE_SUPPRESSION = ? NO-LOCK NO-ERROR.  /* Si l'article a ‚t‚ supprim‚, on peut le recr‚er */
            IF AVAILABLE GPI_PGARTI THEN
            DO:
                /* On v‚rifie qu'il n'a pas ‚t‚ supprim‚ plus haut dans la liste. */
                /* On regarde la derniŠre modification  */
                /* Si c'est C, c'est qu'il a ‚t‚ recr‚‚ (donc probablement supprim‚ plus haut), on passe en M */
                /* Si c'est M, mˆme chose */
                /* Si c'est S, on peut laisser la cr‚ation de l'article */
                FIND FIRST bf_article2 WHERE bf_article2.arti_num_article = bf_article.ARTI_NUM_ARTICLE AND bf_article2.tt_controle USE-INDEX idx_tri NO-ERROR.    /* Pour avoir le dernier article correspondant int‚gr‚. idx_tri tri par fichier/nø de ligne en ordre d‚croissant */
                IF NOT AVAILABLE bf_article2 OR (AVAILABLE bf_article2 AND bf_article2.tt_action <> "S") THEN
                    ASSIGN bf_article.tt_action   = "M".
            END.
            ELSE
            DO:
                FIND FIRST bf_article2 WHERE bf_article2.arti_num_article = bf_article.ARTI_NUM_ARTICLE AND bf_article2.tt_controle USE-INDEX idx_tri NO-ERROR.    /* Pour avoir le dernier article correspondant int‚gr‚. idx_tri tri par fichier/nø de ligne en ordre d‚croissant */
                IF AVAILABLE bf_article2 AND bf_article2.tt_action <> "S" THEN
                    ASSIGN bf_article.tt_action   = "M".
            END.
            ASSIGN bf_article.tt_controle = TRUE.
        END. /* WHEN "C" */
        WHEN "M" THEN
        DO TRANSACTION:
            /* On v‚rifie que l'article existe en base. Si ce n'est pas le cas, et que l'article n'est pas cr‚‚ (et non supprim‚) plus haut dans les fichiers, on passe en cr‚ation */
            FIND FIRST GPI_PGARTI WHERE GPI_PGARTI.ARTI_NUM_ARTICLE = bf_article.ARTI_NUM_ARTICLE AND GPI_PGARTI.ARTI_DATE-HEURE_SUPPRESSION = ? NO-LOCK NO-ERROR.  /* Si l'article a ‚t‚ supprim‚, on peut le recr‚er */
            IF AVAILABLE GPI_PGARTI THEN
            DO:
                /*
                 On v‚rifie qu'il n'a pas ‚t‚ supprim‚ plus haut dans la liste.
                 On regarde la derniŠre modification 
                 Si c'est C, c'est qu'il a ‚t‚ recr‚‚ (donc probablement supprim‚ plus haut), on passe en M
                 Si c'est M, mˆme chose
                 Si c'est S, on peut laisser la cr‚ation de l'article
                */
                FIND FIRST bf_article2 WHERE bf_article2.arti_num_article = bf_article.ARTI_NUM_ARTICLE AND bf_article2.tt_controle USE-INDEX idx_tri NO-ERROR.    /* Pour avoir le dernier article correspondant int‚gr‚. idx_tri tri par fichier/nø de ligne en ordre d‚croissant */
                IF AVAILABLE bf_article2 AND bf_article2.tt_action = "S" THEN
                    ASSIGN bf_article.tt_action   = "C".
            END.
            ELSE
            DO:
                FIND FIRST bf_article2 WHERE bf_article2.arti_num_article = bf_article.ARTI_NUM_ARTICLE AND bf_article2.tt_controle USE-INDEX idx_tri NO-ERROR.    /* Pour avoir le dernier article correspondant int‚gr‚. idx_tri tri par fichier/nø de ligne en ordre d‚croissant */
                IF NOT AVAILABLE bf_article2 OR (AVAILABLE bf_article2 AND bf_article2.tt_action = "S") THEN
                    ASSIGN bf_article.tt_action   = "C".
            END.
            ASSIGN bf_article.tt_controle = TRUE.
        END. /* WHEN "M" */
        WHEN "S" THEN
        DO TRANSACTION:
            /*
             L'article doit exister, sinon on supprime purement et simplement la ligne de suppression
             On v‚rifie que l'article existe en base. Si ce n'est pas le cas, et que l'article n'est pas cr‚‚ (et non supprim‚) plus haut dans les fichiers, on passe en cr‚ation
            */
            FIND FIRST GPI_PGARTI WHERE GPI_PGARTI.ARTI_NUM_ARTICLE = bf_article.ARTI_NUM_ARTICLE AND GPI_PGARTI.ARTI_DATE-HEURE_SUPPRESSION = ? NO-LOCK NO-ERROR.  /* Si l'article a ‚t‚ supprim‚, on peut le recr‚er */
            IF AVAILABLE GPI_PGARTI THEN
            DO:
                /*
                 On v‚rifie qu'il n'a pas ‚t‚ supprim‚ plus haut dans la liste.
                 On regarde la derniŠre modification 
                 Si c'est C, c'est qu'il a ‚t‚ recr‚‚ (donc probablement supprim‚ plus haut), on passe en M
                 Si c'est M, mˆme chose
                 Si c'est S, on peut laisser la cr‚ation de l'article
                */
                FIND FIRST bf_article2 WHERE bf_article2.arti_num_article = bf_article.ARTI_NUM_ARTICLE AND bf_article2.tt_controle USE-INDEX idx_tri NO-ERROR.    /* Pour avoir le dernier article correspondant int‚gr‚. idx_tri tri par fichier/nø de ligne en ordre d‚croissant */
                IF AVAILABLE bf_article2 AND bf_article2.tt_action = "S" THEN
                    DELETE bf_article.
            END.
            ELSE
            DO:
                FIND FIRST bf_article2 WHERE bf_article2.arti_num_article = bf_article.ARTI_NUM_ARTICLE AND bf_article2.tt_controle USE-INDEX idx_tri NO-ERROR.    /* Pour avoir le dernier article correspondant int‚gr‚. idx_tri tri par fichier/nø de ligne en ordre d‚croissant */
                IF NOT AVAILABLE bf_article2 OR (AVAILABLE bf_article2 AND bf_article2.tt_action = "S") THEN
                    DELETE bf_article.
            END.
            IF AVAILABLE bf_article THEN
                ASSIGN bf_article.tt_controle = TRUE.
        END. /* WHEN "S" */
        OTHERWISE
            DELETE bf_article.
    END CASE.
END.
 
/* fin de traitement */
RUN jauge-init IN  h_jauge-ocx ( INPUT ii_numArticles).
RUN jauge-fin IN h_jauge-ocx.
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE create_TT_OL_EXCLUS W-Win 
PROCEDURE create_TT_OL_EXCLUS :
/*--*/
 
    DEF VAR ii AS INT NO-UNDO.
 
    EMPTY TEMP-TABLE TT_OL_EXCLUS.
 
    DO ii=1 TO NUM-ENTRIES ( ed_LstExclusions, CHR(10) ):
        CREATE TT_OL_EXCLUS.
        TT_OL_EXCLUS.num_ol = ENTRY ( ii, ed_LstExclusions, CHR(10) ).
    END.
 
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE createFichier W-Win
PROCEDURE createFichier:
/*------------------------------------------------------------------------------
 Purpose: CrŠe un enregistrement dans la table tt_fichier pour archivage ult‚rieur
 Notes:
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ic_fichier AS CHARACTER NO-UNDO.
DEFINE INPUT  PARAMETER ic_rep     AS CHARACTER NO-UNDO.
 
CREATE tt_fichiers.
ASSIGN
tt_fichiers.tt_fichier     = ic_fichier
tt_fichiers.tt_rep_origine = ic_rep.
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
 
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE disable_UI W-Win  _DEFAULT-DISABLE
PROCEDURE disable_UI :
/*------------------------------------------------------------------------------
  Purpose:     DISABLE the User Interface
  Parameters:  <none>
  Notes:       Here we clean-up the user-interface by deleting
               dynamic widgets we have created and/or hide 
               frames.  This procedure is usually called when
               we are ready to "clean-up" after running.
------------------------------------------------------------------------------*/
  /* Delete the WINDOW we created */
  IF SESSION:DISPLAY-TYPE = "GUI":U AND VALID-HANDLE(W-Win)
  THEN DELETE WIDGET W-Win.
  IF THIS-PROCEDURE:PERSISTENT THEN DELETE PROCEDURE THIS-PROCEDURE.
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE edite_browse W-Win 
PROCEDURE edite_browse :
/*--*/
 
    /*DEF INPUT PARAMETER quoi AS CHAR NO-UNDO.*/
    DEFINE OUTPUT PARAMETER oc_nomFic   AS CHARACTER NO-UNDO.
 
    DEFINE VARIABLE lc_repert AS CHARACTER NO-UNDO.
    DEFINE VARIABLE lc_param  AS CHARACTER NO-UNDO.
 
    IF br_Erreurs:NUM-ITERATIONS IN FRAME F-Erreurs = 0 THEN
        RETURN.
 
    /* Si mode batch enregistre directement le fichier :          */
    /* - Dans le r‚pertoire Sauvegarde\ann‚e\mois\jour si existe  */
    /* - sinon, dans le r‚pertoire courant du serveur             */
    /* Si le r‚pertoire est diff‚rent du r‚pertoire courant, on   */
    /* l'enregistre dans la variable globale gc_replog, pour ˆtre */
    /* enregistr‚ d‚finitivement et indiqu‚ dans le mail envoy‚.  */
 
    /* Il y a quelque chose … ‚diter, on va d‚finir dans quel r‚pertoire enregistrer le fichier */
    ASSIGN
    gc_replog = ""
    gc_ficlog = ""
    lc_param  = "Excel".
    /* On n'enregistre le fichier que si l'on est en mode batch, et que le mail est valide */
    IF gl_modeBatch AND fi_email <> "" AND fi_listeMailsValide(fi_email) THEN
    DO:
 
        gc_ficlog = "log_" + STRING(YEAR(TODAY),"9999") + STRING(MONTH(TODAY), "99") + STRING(DAY(TODAY), "99") + REPLACE(STRING(TIME, "hh:mm:ss"), ":", "") + ".xls".
 
        RUN genereRepSauvegarde(OUTPUT lc_repert) NO-ERROR.
        IF ERROR-STATUS:ERROR OR lc_repert = "" OR lc_repert = ? THEN
        DO:
            ASSIGN
            FILE-INFORMATION:FILE-NAME = "."
            lc_repert = RIGHT-TRIM(RIGHT-TRIM(FILE-INFORMATION:FULL-PATHNAME,"/"),"\")
            gc_replog = "".
        END.
        ELSE
            ASSIGN gc_replog = lc_repert.
 
        oc_nomfic = lc_repert + "\" + gc_ficlog.
 
        lc_param = lc_param + CHR(8) + oc_nomfic + CHR(8) + "Non".
    END.
 
    br_Erreurs:QUERY:REPOSITION-TO-ROW(10000) NO-ERROR.
    br_Erreurs:QUERY:REPOSITION-TO-ROW(1) NO-ERROR.
 
    RUN PROCEDUR\p-edite_browse-2.p ( INPUT br_erreurs:HANDLE,
                                      INPUT "tout"                                              + CHR(8) +
                                            "oui"                                               + CHR(8) +
                                            ""                                                  + CHR(8) +
                                            "Ajuster",                                                    
                                      INPUT lc_param,
                                      INPUT "Erreurs d'importation"                             + CHR(8) + 
                                            ""                                        + CHR(8) +
                                            "Non"                                               + CHR(8) +
                                            "",                                               
                                      INPUT "Oui"                                               + CHR(8) +
                                            "Oui"                                               + CHR(8) +
                                            "Oui"                                               + CHR(8) +
                                            "Oui").                                                       
 
    /*
        Les paramŠtres ci-dessous
 
 
        DEF INPUT PARAMETER Handle_Browse           AS WIDGET-HANDLE    NO-UNDO.
        /*
            Handle du browse … parcourir
        */
 
        DEF INPUT PARAMETER parametrage_colonnes    AS CHAR             NO-UNDO.
        /*
            8 entries CHR(8) :
                Liste Colonnes A Editer         Liste des noms des champs du browse … ‚diter dans l'ordre ( tout=toutes les colonnes VISIBLES )
                Multi-Pages Oui/Non             Si Oui, ‚dite sur une nouvelle page ce qui "d‚borde" sur la droite de la page ( Xprint seulement )
                Colonnes Multi-Pages            Dans le cas du Multi-Page, contient une liste les num‚ros de colonnes … r‚-‚diter sur la page qui d‚borde
                Choix Multi-Pages               Si "Oui" : Si l'‚dition "d‚borde" en largeur, demande "Continuer sur la seconde page" OU "1 page seulement" OU "Ajuster" OU "Abandon"
                                                           ( Dans ce cas, ne tient pas compte du 2Šme paramŠtre : Multi-Pages, ni du 2Šme : Colonnes Multi-Pages )
                                                           UNIQUEMENT DANS LE CAS OU ON RECUPERE LES POLICES DU BROWSE
                                                Si "Non" : Edite tout ce qui est possible d'‚diter sur une page, en coupant ce qui d‚borde
                                                Si "Ajuster" : Force l'ajustement sur une page
                                                Si "ajusterExcel": Force l'ajustement sur une page pour l'impression Excel
                Conserver format dec(excel)        Num‚ros des colonnes s‚par‚s par des virgules pour lesquelles il faut conserver le format decimal d‚finit dans le browse
                type traits verticaux,horiz(excel) S‚par‚ par une virgule, si non renseign‚ 1,1 (3 pour idem que l'entˆte)
                Alignement sp‚cifique(excel)       Num‚ro de colonne:[left,center,right],Num‚ro de colonne....
                Orientation de la page(excel)      portrait/paysage, par d‚faut portrait
        */
 
        DEF INPUT PARAMETER sortie                  AS CHAR             NO-UNDO.
        /*                                      
            Si XPrint, 3 entries CHR(8) :           
                Fichier_xpr                     Nom du fichier Xprint de sortie qui sera plac‚ dans le r‚pertoire de d‚marrage de D'Artagnan
                Imprimante                      ? pour avoir le choix de l'imprimante, ou alors son nom
                Zoom                            no = edition, ZoomToWidth ou valeur du zoom ( 100 = pleine page ), ou "Pas d'aper‡u"
                                                Si "Pdf", cr‚‚ un fichier Pdf sans aper‡u, 
 
            Si Excel, 3 entries CHR(8) :
                "Excel"                         Transforme le browse en fiche Excel plutot qu'en ‚dition Xprint
                Chemin Complet Fichier          Enregistrer le fichier Excel avec ce nom ( Types de fichiers g‚r‚s : .XLS et .CSV )
                Aper‡u Oui/Non                  Afficher le fichier g‚n‚r‚ dans Excel ( Par d‚faut … "Oui" )
        */                
 
        DEF INPUT PARAMETER titres_totaux_pied      AS CHAR             NO-UNDO.
        /*
            4 entries CHR(8) :
                titre                           Titre principal en gras ‚dit‚ avec Xprint ou repris sous Excel ( d‚coup‚ … 60 mm ou 160 mm )
                sous-titre                      Sous-titre multi-lignes ( Balises xpr accept‚es en DEBUT de ligne. Ex: "lig1" + chr(10) + "<B>lig2"
                Totaux Oui/Non                  Afficher ou non un total des INT et DEC … la fin de l'‚dition ou de la feuille Excel  OU  Liste des noms des colonnes … cumuler
                Pied de Page                    Affiche un texte en gras en pied de page de la derniere page de l'‚dition ou de la feuille Excel
 
                Titre de colonne suppl‚mentaire Affiche une case au dessus du titre d'une colonne avec un texte
                                                Il est possible de faire une case sur plusieurs colonnes, voir ci-dessous
                                                ----------------------       -------------------
                                                |   Titre Col 123    |       | Titre Col5 Lig1 |
                                                |                    |       | Titre Col5 Lig2 |
                                                ------------------------------------------------
                                                | Col1 | Col2 | Col3 |  Col4 |            Col5 |
                                                ------------------------------------------------
                                                | txt1 | txt3 | txt5 |  txt9 |            txt7 |
                                                | txt2 | txt4 | txt6 | txt10 |            txt8 |
 
                                                Exemple de syntaxe du paramŠtre : "1-3:Titre Col 123" + CHR(9) + "5:Titre Col5 Lig 1" + CHR(10) + "Titre Col5 Lig2"
 
                MARCHE PAS 
                Champ … traitement particulier  Nom d'un champ + ":" + action : Traitement particulier si le champ est vide
                                                "Ne pas afficher" : La ligne n'est pas imprim‚e ( pour ne pas imprimer les lignes de sous-totaux )
                                                "Ne pas cumuler"  : La ligne n'est pas cumul‚e ( ‚galement pour les sous-totaux )
        */
 
        DEF INPUT PARAMETER recup_couleurs_polices  AS CHAR             NO-UNDO.
        /*                                      
            4 entries CHR(8) :                  
                Recup Couleurs Entˆte Oui/Non           Reprend les couleurs des entˆtes des colonnes
                Recup Couleurs Lignes Oui/Non           Reprend les couleurs des lignes du browse
                Recup Polices  Entˆte Oui/Non/Non Gras  Reprend les polices des entˆtes des colonnes
                                                        Si "Non Gras", n'‚dite pas les titres de colonnes en gras
                                                        pour pouvoir ‚diter des colonnes de faible largeur
                Recup Polices  Lignes Oui/Non           Reprend les polices des lignes du browse
 
            MODIFS THR - 08/08/12, int‚gration possibilit‚ de cr‚er des graphiques.
            Entr‚es graphiques (fonctionne seulement avec Excel) CHR(8) pour s‚parer des entr‚es pr‚c‚dentes, puis CHR(5)
                Type de graphique                       "51":histogramme, "1":Area, "4":courbe, "-4120":donnuts...cf function "create_graph" dans "Excel4.i"
                Titres                                  Titres graphique CHR(9) Titre abscisse CHR(9) Titre ordonn‚e
                Colonne abscisse                        Un seul ‚l‚ment
                Colonnes ordonnees                      s‚par‚es par CHR(9)
                Titre colonnes                          s‚par‚s par CHR(9)
                [Nombre ligne]                          Facultatif, si non renseign‚ se base sur toutes les lignes du browse
                [Type de graphique.....]                Pour plusieurs graphiques s‚par‚s par CHR(3)
        */                                           
 
 
    */
 
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE enable_UI W-Win  _DEFAULT-ENABLE
PROCEDURE enable_UI :
/*------------------------------------------------------------------------------
  Purpose:     ENABLE the User Interface
  Parameters:  <none>
  Notes:       Here we display/view/enable the widgets in the
               user-interface.  In addition, OPEN all queries
               associated with each FRAME and BROWSE.
               These statements here are based on the "Other 
               Settings" section of the widget Property Sheets.
------------------------------------------------------------------------------*/
  DISPLAY Tg_reinitClients rs_visu ED_Commentaire liste-1 liste-3 liste-4 
          liste-2 lbl_intresutl 
      WITH FRAME F-Main IN WINDOW W-Win.
  ENABLE Tg_reinitClients rs_visu Br_Articles Br_Clients Br_Corresp 
         Br_InterEber Br_OL Br_PlanTournee ED_Commentaire Br_LigneOL liste-1 
         liste-3 liste-4 liste-2 Btn_Importer Btn_Test Btn_Quitter 
         lbl_intresutl 
      WITH FRAME F-Main IN WINDOW W-Win.
  {&OPEN-BROWSERS-IN-QUERY-F-Main}
  ENABLE icone-banniere IMAGE-banniere img_errImport img_param 
      WITH FRAME FRAME-banniere IN WINDOW W-Win.
  VIEW FRAME FRAME-banniere IN WINDOW W-Win.
  {&OPEN-BROWSERS-IN-QUERY-FRAME-banniere}
  VIEW FRAME F-jauge IN WINDOW W-Win.
  {&OPEN-BROWSERS-IN-QUERY-F-jauge}
  DISPLAY fi_rep_integration_articles fi_rep_integration_ol 
          fi_rep_integration_param fi_rep_svg fi_email 
      WITH FRAME F-Criteres IN WINDOW W-Win.
  ENABLE btn_rech_articles Btn_Executer fi_rep_integration_articles btn_rech_OL 
         fi_rep_integration_ol btn_rech_Param fi_rep_integration_param 
         btn_rech_svg fi_rep_svg fi_email 
      WITH FRAME F-Criteres IN WINDOW W-Win.
  {&OPEN-BROWSERS-IN-QUERY-F-Criteres}
  DISPLAY ed_lstExclusions 
      WITH FRAME F-PARAM IN WINDOW W-Win.
  ENABLE ed_lstExclusions Btn_param_ok Btn_param_annuler 
      WITH FRAME F-PARAM IN WINDOW W-Win.
  {&OPEN-BROWSERS-IN-QUERY-F-PARAM}
  FRAME F-PARAM:SENSITIVE = NO.
  ENABLE Br_Erreurs Btn_Fermer Btn_Excel 
      WITH FRAME F-Erreurs IN WINDOW W-Win.
  {&OPEN-BROWSERS-IN-QUERY-F-Erreurs}
  VIEW W-Win.
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE erreur W-Win 
PROCEDURE erreur :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ic_fichier  AS CHARACTER   NO-UNDO.
DEFINE INPUT  PARAMETER ii_numLigne AS INTEGER     NO-UNDO.
DEFINE INPUT  PARAMETER ic_message  AS CHARACTER   NO-UNDO.
 
CREATE tt_erreur.
ASSIGN
tt_erreur.date_heure_erreur = NOW
tt_erreur.fichier           = ic_fichier
tt_erreur.numLigne          = ii_numLigne
tt_erreur.libErreur         = ic_message.
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE genereMail W-Win 
PROCEDURE genereMail :
/*------------------------------------------------------------------------------
  Purpose:     Envoie un mail d'avertissement en cas d'erreur, avec le log en piŠce jointe
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT PARAMETER ic_nomFic AS CHARACTER NO-UNDO.
DEFINE OUTPUT PARAMETER ol_ok    AS LOGICAL   NO-UNDO.
 
DEFINE VARIABLE lc_listeMails AS CHARACTER NO-UNDO.
DEFINE VARIABLE lc_objetMail  AS CHARACTER NO-UNDO.
DEFINE VARIABLE lc_corpsMail  AS CHARACTER NO-UNDO.
 
ASSIGN
lc_listeMails = "TO:" + REPLACE(fi_email:SCREEN-VALUE IN FRAME F-Criteres, ",", ";") + " BCC:nid@g-p-i.fr"
lc_objetMail  = "Anomalie import PGI Eberhardt du " + STRING(TODAY) + " … " + STRING(TIME, "hh:mm")
lc_corpsMail  = "Des anomalies se sont produites … l'import des fichiers Eberhardt." + CHR(10) +
                "Vous trouverez le d‚tail de ces anomalies dans le fichier joint " + gc_ficlog + CHR(10).
 
IF gc_replog <> "" THEN
    IF gc_replog BEGINS "\\" THEN
        lc_corpsMail = lc_corpsMail + CHR(10) + 
                       "Ce fichier est disponible dans le r‚pertoire " + gc_replog + ".".
    ELSE
        lc_corpsMail = lc_corpsMail + CHR(10) + 
                       "Ce fichier est disponible dans le r‚pertoire " + gc_replog + " du serveur.".
 
RUN objets\files\p-envoi_Mail( INPUT lc_listeMails, /*   Sous la forme : "TO:mail@dest1;mail@dest2 CC:mail@CarbonCopy BCC:mail@BlindCarbonCopy"   */
                               INPUT   lc_objetMail,    /*   Objet du message                                                                         */
                               INPUT   lc_corpsMail,    /*   Corps du message                                                                         */
                               INPUT   ic_nomFic,       /*   Chemins complets des fichiers : Chaque fichier XPR sera transform‚ en PDF avant envoi    */
                               INPUT   FALSE,            /*   Si vrai, la messagerie s'ouvrira, sinon, envoi en arriŠre-plan         */
                               INPUT   FALSE,            /*   Si vrai, les messages d'erreurs sont affich‚s                                            */
                               INPUT   "",              /*   Si Vide envoi par p-ENVOI_MAIL - Sinon envoi par CDO - pas besoin de OUTLOOK ou autre    */
                               /*
                               INPUT   IF gl_outlook
                                       THEN ""
                                       ELSE "HTML",  /* Pas d'utilisation d'Outlook en cas d'erreur */
                               */
                               OUTPUT  ol_ok ).
 
IF ol_ok AND gc_replog = "" THEN
    OS-DELETE VALUE(ic_nomFic).
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE genereRepSauvegarde W-Win 
PROCEDURE genereRepSauvegarde :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE OUTPUT PARAMETER rep_cible       AS CHARACTER NO-UNDO.
 
DEFINE VARIABLE rep_sauvegarde  AS CHARACTER NO-UNDO.
 
FILE-INFO:FILE-NAME = fi_rep_svg .
rep_sauvegarde = FILE-INFO:FULL-PATHNAME .
 
/* V‚rification */
IF rep_sauvegarde = ? OR rep_sauvegarde = "" THEN 
DO:
    RUN erreur ("", 0, "Le r‚pertoire de sauvegarde n'est pas renseign‚ !").
    APPLY "ENTRY" TO fi_rep_svg IN FRAME F-Criteres. 
    RETURN ERROR "".
END.
 
/*............................................................*/
/*    pr‚paration des r‚pertoires aaaa/mm/jj pour d‚placement */
/*............................................................*/
rep_cible = RIGHT-TRIM(RIGHT-TRIM(rep_sauvegarde, "\"), "/").
 
rep_cible = rep_cible + "\" + STRING(YEAR(TODAY),"9999").
FILE-INFO:FILE-NAME = rep_cible.
IF FILE-INFO:FULL-PATHNAME = ? THEN
    OS-CREATE-DIR VALUE ( FILE-INFO:FILE-NAME ).
 
rep_cible = rep_cible + "\" + STRING(MONTH(TODAY),"99").
FILE-INFO:FILE-NAME = rep_cible.
IF FILE-INFO:FULL-PATHNAME = ? THEN
    OS-CREATE-DIR VALUE ( FILE-INFO:FILE-NAME ).
 
rep_cible = rep_cible + "\" + STRING(DAY(TODAY),"99").
FILE-INFO:FILE-NAME = rep_cible.
IF FILE-INFO:FULL-PATHNAME = ? THEN
    OS-CREATE-DIR VALUE ( FILE-INFO:FILE-NAME ).
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE get_os_error W-Win 
PROCEDURE get_os_error :
DEF INPUT  PARAMETER num_error AS INT   NO-UNDO.
DEF OUTPUT PARAMETER lib_error AS CHAR  NO-UNDO.
 
CASE num_error :
 
    WHEN   0 THEN   lib_error = "No error".
    WHEN   1 THEN   lib_error = "Not owner".
    WHEN   2 THEN   lib_error = "No such file or directory".
    WHEN   3 THEN   lib_error = "Interrupted system call".
    WHEN   4 THEN   lib_error = "I/O error".
    WHEN   5 THEN   lib_error = "Bad file number".
    WHEN   6 THEN   lib_error = "No more processes".
    WHEN   7 THEN   lib_error = "Not enough core memory".
    WHEN   8 THEN   lib_error = "Permission denied".
    WHEN   9 THEN   lib_error = "Bad address".
    WHEN  10 THEN   lib_error = "File exists".
    WHEN  11 THEN   lib_error = "No such device".
    WHEN  12 THEN   lib_error = "Not a directory".
    WHEN  13 THEN   lib_error = "Is a directory".
    WHEN  14 THEN   lib_error = "File table overflow".
    WHEN  15 THEN   lib_error = "Too many open files".
    WHEN  16 THEN   lib_error = "File too large".
    WHEN  17 THEN   lib_error = "No space left on device".
    WHEN  18 THEN   lib_error = "Directory not empty".
    WHEN 999 THEN   lib_error = "Unmapped error (Progress default)".
    OTHERWISE       lib_error = "Unknown !".
END CASE.
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE initDateLivraison W-Win 
PROCEDURE initDateLivraison :
/*------------------------------------------------------------------------------
  Purpose:     Initialise la date de livraison pr‚vue sur les OL en PF
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE VARIABLE ldt_liv     AS DATE     NO-UNDO.
DEFINE VARIABLE ldt_charg   AS DATE     NO-UNDO.
DEFINE VARIABLE ldt_arrPf   AS DATE     NO-UNDO.
DEFINE VARIABLE ldt_livRef  AS DATE     NO-UNDO.
DEFINE VARIABLE ll_ok       AS LOGICAL  NO-UNDO.
DEFINE VARIABLE li_nbNon    AS INTEGER  NO-UNDO.
DEFINE VARIABLE li_numSem   AS INTEGER  NO-UNDO.
DEFINE VARIABLE ldt_test    AS DATE     NO-UNDO.
 
DEFINE VARIABLE li_nbEnreg  AS INTEGER  NO-UNDO.
 
/*SELECT COUNT(*) INTO li_nbEnreg FROM gpi_pgol WHERE gpi_pgol.ol_mode_liv = "PF" AND gpi_pgol.ol_num_pf <> "" AND gpi_pgol.ol_date_liv IS NULL AND gpi_pgol.ol_date_retour_Eber IS NULL. */
SELECT COUNT(*) INTO li_nbEnreg FROM gpi_pgol WHERE gpi_pgol.ol_mode_liv = "PF" AND gpi_pgol.ol_num_pf <> "" /*AND gpi_pgol.ol_date_liv IS NULL*/ AND gpi_pgol.ol_date_retour_Eber IS NULL AND NOT gpi_pgol.ol_a_supprimer AND NOT GPI_PGOL.ol_ancien_ol. /* V‚rification faite, pas de contr“le sur la date de livraison. Ajout du contr“le sur ancien_ol */
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Initialisation des dates livraison pr‚visionnelles. Veuillez patienter ... (<p>%)").
 
RUN jauge-init IN h_jauge-ocx ( INPUT li_nbEnreg ).
 
/*OUTPUT TO VALUE("c:\temp\2016\10\28\20161028_" + REPLACE(STRING(TIME),":","") + "_initDateLivraison.txt"). */
 
/*
FOR EACH gpi_pgol WHERE gpi_pgol.ol_mode_liv = "PF" AND gpi_pgol.ol_date_liv = ? NO-LOCK,
FOR EACH gpi_pgol WHERE gpi_pgol.ol_mode_liv = "PF" AND gpi_pgol.ol_date_liv = ? AND NOT gpi_pgol.ol_a_supprimer NO-LOCK,
*/
FOR EACH gpi_pgol WHERE gpi_pgol.ol_mode_liv = "PF" 
                    AND gpi_pgol.ol_date_retour_Eber  = ? 
                    AND NOT gpi_pgol.ol_a_supprimer AND 
                    NOT GPI_PGOL.ol_ancien_ol NO-LOCK,
    FIRST tt_plateforme WHERE tt_plateforme.tt_code = gpi_pgol.ol_num_pf:
 
    RUN jauge-next IN h_jauge-ocx.
 
    /* La premiŠre date de livraison disponible */
    ASSIGN
    ldt_charg = tt_plateforme.tt_dtLivPrev
    ldt_liv   = ldt_charg + 1.
 
    /*PUT UNFORMATTED "OL " gpi_pgol.ol_num_ol " - PF " gpi_pgol.ol_num_pf " - Etape 1 - Date Chg = " ldt_charg SKIP.*/
 
    DO WHILE fi_ferie(ldt_liv):
        ASSIGN ldt_liv = ldt_liv + 1.
    END.
 
    /* Date d'arriv‚e … la plateforme */
    ASSIGN ldt_arrPf = ldt_liv.
 
    /*PUT UNFORMATTED "OL " gpi_pgol.ol_num_ol " - PF " gpi_pgol.ol_num_pf " - Etape 2 - Date arriv‚e PF = " ldt_arrPF SKIP. */
 
    /* Calcul de la date th‚orique de livraison client */
    /* Cherche le premier jour travaill‚ suivant la date d'arriv‚e sur la plateforme */
    ASSIGN ldt_liv = ldt_liv + 1.
    DO WHILE fi_ferie(ldt_liv):
        ASSIGN ldt_liv = ldt_liv + 1.
    END.
 
    /* Cette date deviens la date de r‚f‚rence */
    ASSIGN ldt_livRef = ldt_liv.
 
    /*PUT UNFORMATTED "OL " gpi_pgol.ol_num_ol " - PF " gpi_pgol.ol_num_pf " - Etape 3 - Date livraison de r‚f‚rence = " ldt_livRef SKIP. */
    FIND FIRST gpi_pgpt WHERE gpi_pgpt.pgpt_codpos = gpi_pgol.ol_cp_livraison AND gpi_pgpt.pgpt_vil20 = gpi_pgol.ol_ville_livraison NO-LOCK NO-ERROR.
    IF AVAILABLE gpi_pgpt THEN
    DO:
        /* NID le 27/10/16 - On va tester tout ‡a diff‚rement */
        /* On commence par v‚rifier qu'au moins un jour est diff‚rent de false (soit ?, soit true) */
        DEFINE VARIABLE ll_test AS LOGICAL     NO-UNDO.
 
        /*PUT UNFORMATTED "OL " gpi_pgol.ol_num_ol " - PF " gpi_pgol.ol_num_pf " - Etape 4 - Test : " ll_test " - Tournee : L=" gpi_pgpt.pgpt_tournee[1] " M=" gpi_pgpt.pgpt_tournee[2] " Me=" gpi_pgpt.pgpt_tournee[3] " J=" gpi_pgpt.pgpt_tournee[4] " V=" gpi_pgpt.pgpt_tournee[5] SKIP. */
 
        /* Vaux faux si tout est … faux, ? si on n' que du faux ou du ?, vrai si au moins un est … vrai, que les autres soient ? ou faux */
        ASSIGN ll_test = gpi_pgpt.pgpt_tournee[1] OR gpi_pgpt.pgpt_tournee[2] OR gpi_pgpt.pgpt_tournee[3] OR gpi_pgpt.pgpt_tournee[4] OR gpi_pgpt.pgpt_tournee[5].
 
        /* Au moins un jour (entre lundi et vendredi) est … vrai ou ?. Dans le cas contraire, la date reste la date de r‚f‚rence */
        IF ll_test <> FALSE THEN
        DO:
            ASSIGN ll_ok = FALSE.
            DO WHILE NOT ll_ok:
                CASE gpi_pgpt.pgpt_tournee[WEEKDAY(ldt_liv) - 1]:
                    WHEN TRUE THEN
                    DO:
                        /* Si tourn‚e OK on v‚rifie le jour de d‚part th‚orique en fonction du type de la tourn‚e (1 ou 2j) */
                        ASSIGN ldt_test = ldt_liv - GPI_PGPT.pgpt_to2j.
                        DO WHILE fi_ferie(ldt_test):
                            ASSIGN ldt_test = ldt_test - 1.
                        END.
                        IF ldt_test >= ldt_arrPf THEN
                            ASSIGN ll_ok = TRUE.
                        ELSE
                            ASSIGN ldt_liv = ldt_liv + 1.
                    END.
                    WHEN FALSE THEN
                        ASSIGN ldt_liv = ldt_liv + 1.
                    WHEN ? THEN
                        ASSIGN
                        ldt_liv = ldt_liv + 2
                        ll_ok   = TRUE.
                END.
 
                DO WHILE fi_ferie(ldt_liv):
                    ASSIGN ldt_liv = ldt_liv + 1.
                END.
                /*PUT UNFORMATTED "OL " gpi_pgol.ol_num_ol " - PF " gpi_pgol.ol_num_pf " - Etape 5 - Date livraison = " ldt_liv SKIP. */
            END.
        END.
 
        /*
        /*PUT UNFORMATTED "Tourn‚e trouv‚e pour " gpi_pgol.ol_cp_livraison " " gpi_pgol.ol_ville_livraison SKIP. */
        ASSIGN
        ll_ok     = FALSE
        li_nbNon  = 1
        /* Le nø jour dans la semaine, de 1 (lundi) … 7 (dimanche) */
        li_numSem = IF WEEKDAY(ldt_liv) = 1 THEN 7 ELSE WEEKDAY(ldt_liv) - 1. /* Progress retourne un jour de la semaine entre dimanche (1) et samedi (6), alors qu'il faudrait entre lundi et dimanche */
 
        DO WHILE NOT ll_ok:
 
            /*PUT UNFORMATTED "Tourn‚ee pour le " gdt_liv " : " gpi_pgpt.pgpt_tournee[gi_numsem] SKIP. */
 
            /* Si tourn‚e OK on v‚rifie le jour de d‚part th‚orique en fonction du type de la tourn‚e (1 ou 2j) */
            IF gpi_pgpt.pgpt_tournee[li_numsem] = TRUE THEN
            DO:
                ASSIGN ldt_test = ldt_liv - gpi_pgpt.pgpt_to2j.
                DO WHILE fi_ferie(ldt_test):
                    ASSIGN ldt_test = ldt_test - 1.
                END.
                /*PUT UNFORMATTED "Etape 4 - Date test = " gdt_test SKIP. */
            END.
 
            IF (gpi_pgpt.pgpt_tournee[li_numsem] = FALSE AND li_nbNon < 6) OR
               (gpi_pgpt.pgpt_tournee[li_numsem] = TRUE AND ldt_test < ldt_arrPf) THEN
            DO:
                ASSIGN
                li_nbNon = li_nbNon + 1
                ldt_liv = ldt_liv + 1.
                DO WHILE fi_ferie(ldt_liv):
                    ASSIGN ldt_liv = ldt_liv + 1.
                END.
                ASSIGN li_numSem = IF WEEKDAY(ldt_liv) = 1 THEN 7 ELSE WEEKDAY(ldt_liv) - 1. /* Progress retourne un jour de la semaine entre dimanche (1) et samedi (6), alors qu'il faudrais entre lundi et dimanche */
                /*PUT UNFORMATTED "Etape 5 - Poursuite traitement avec " gdt_liv SKIP. */
            END.
            ELSE
            DO:
                IF li_nbNon >= 6 THEN
                    ASSIGN ldt_liv = ldt_livRef.
 
                ASSIGN li_numSem = IF WEEKDAY(ldt_liv) = 1 THEN 7 ELSE WEEKDAY(ldt_liv) - 1. /* Progress retourne un jour de la semaine entre dimanche (1) et samedi (6), alors qu'il faudrais entre lundi et dimanche */
 
                /*PUT UNFORMATTED "Etape 6 - Traitement arr‚t‚ - Date livraison : " gdt_liv " - Nb Non : " gi_nbNon " - Tourn‚e : " gpi_pgpt.pgpt_tournee[gi_numsem] SKIP. */
 
                IF gpi_pgpt.pgpt_tournee[li_numsem] = ? THEN
                    ASSIGN ldt_liv = ldt_liv + 2.
 
                DO WHILE fi_ferie(ldt_liv):
                    ASSIGN ldt_liv = ldt_liv + 1.
                END.
                /*PUT UNFORMATTED "Etape 7 - Date Liv = " gdt_liv SKIP. */
                ASSIGN ll_ok = TRUE.
            END.
        END.
        */
    END.
    /*
    ELSE
        PUT UNFORMATTED "Tourn‚e NON trouv‚e pour " gpi_pgol.ol_cp_livraison " " gpi_pgol.ol_ville_livraison SKIP.
 
    PUT UNFORMATTED "Etape 8 - Date Liv = " gdt_liv SKIP.
    PUT UNFORMATTED "===========================================================================================================================================" SKIP.
    */
 
    /*PUT UNFORMATTED "OL " gpi_pgol.ol_num_ol " - PF " gpi_pgol.ol_num_pf " - Etape 6 - Date livraison = " ldt_liv SKIP. */
    IF ldt_liv <> ? AND ldt_liv <> gpi_pgol.ol_date_liv THEN
    DO TRANSACTION:
        FIND FIRST bf_pgol WHERE bf_pgol.ol_num_ol = gpi_pgol.ol_num_ol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        IF AVAILABLE bf_pgol THEN
        DO:
            ASSIGN bf_pgol.ol_date_liv = ldt_liv.
            RELEASE bf_pgol.
        END.
    END.
END.
/*OUTPUT CLOSE. */ 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE initPF W-Win 
PROCEDURE initPF :
/*------------------------------------------------------------------------------
  Purpose:     Initialise la table des PF, et la date de livraison previsionnelle
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE VARIABLE li_ii     AS INTEGER    NO-UNDO.
 
/* Pour la recherche de la journ‚e */
DEFINE VARIABLE li_jSem   AS INTEGER     NO-UNDO.
DEFINE VARIABLE ldt_jour  AS DATE        NO-UNDO.
 
/* Remplis la table tt_plateforme a partir des donn‚es des OL */
EMPTY TEMP-TABLE tt_plateforme.
 
/*OUTPUT TO VALUE("C:\Temp\2016\10\27\20161027_" + REPLACE(STRING(TIME, "hh:mm:ss"),":","") + "_majObligatoire.txt"). */
/* On cr‚e un enregistrement par plateforme */
FOR EACH gpi_pgpt NO-LOCK BREAK BY gpi_pgpt.pgpt_plat:
    IF FIRST-OF (gpi_pgpt.pgpt_plat) THEN
    DO:
        IF gpi_pgpt.pgpt_plat = 0 THEN NEXT.
 
        CREATE tt_plateforme.
        ASSIGN
        tt_plateforme.tt_num       = gpi_pgpt.pgpt_plat
        tt_plateforme.tt_code      = STRING(gpi_pgpt.pgpt_plat)
        /* M????? - NID le 03/07/17 */
        /* tt_plateforme.tt_tournee   = FALSE */
        tt_plateforme.tt_approche  = FALSE
        /* Fin M????? - NID le 03/07/17 */
        tt_plateforme.tt_livrable  = FALSE
        tt_plateforme.tt_dtLivPrev = ?.
 
        /* R‚cup‚ration des jours d'approche */
        DO li_ii = 1 TO LENGTH(GPI_PGPT.pgpt_trfpf):
            IF SUBSTRING(GPI_PGPT.pgpt_trfpf, li_ii, 1) = "O" THEN
                ASSIGN
                tt_plateforme.tt_livrable       = TRUE
                /* M????? - NID le 03/07/17 */
                /* tt_plateforme.tt_tournee[li_ii] = TRUE. */
                tt_plateforme.tt_approche[li_ii] = TRUE.
                /* Fin M????? - NID le 03/07/17 */
        END.
 
        ASSIGN ldt_jour = TODAY + 1.
 
        /* La date de livraison de r‚f‚rence est le premier jour non feri‚ qui suis la date du jour */
        DO WHILE fi_ferie(ldt_jour) :
            ldt_jour = ldt_jour + 1.
        END.
 
        /* Si tous les jours d'approche sont … N, pas la peine d'aller plus loin, on affecte le jour de r‚f‚rence */
        IF TRIM(GPI_PGPT.pgpt_trfpf) <> "NNNNN" THEN
            /* Sinon, on cherche ensuite le premier jour non feri‚ correspondant aux jours d'approche */
            /* M????? - NID le 03/07/17 */
            /* DO WHILE (fi_ferie(ldt_jour) OR tt_plateforme.tt_tournee[WEEKDAY(ldt_jour) - 1] <> TRUE): /* ferie retourne False si weekday(ldt_jour) = 1 ou 7, donc pas de soucis de d‚passement de capacit‚ sur le tableau, WEEKDAY(ldt_jour) ‚tant compris entre 2 et 6 */ */
            DO WHILE (fi_ferie(ldt_jour) OR tt_plateforme.tt_approche[WEEKDAY(ldt_jour) - 1] <> TRUE): /* ferie retourne False si weekday(ldt_jour) = 1 ou 7, donc pas de soucis de d‚passement de capacit‚ sur le tableau, WEEKDAY(ldt_jour) ‚tant compris entre 2 et 6 */
            /* Fin M????? - NID le 03/07/17 */
                ldt_jour = ldt_jour + 1.
            END.
 
        ASSIGN tt_plateforme.tt_dtLivPrev = ldt_jour.
 
        /*
        FIELD tt_dtCharg    AS DATE         /* Le premier jour de chargement possible depuis la plateforme vers le client final */
        FIELD tt_dtLivCourt AS DATE         /* La premiŠre date de livraison possible, sans +2 de la livraison 72h */
        FIELD tt_dtLivLong  AS DATE         /* La premiŠre date de livraison possible, avec le +2 de la livraison 72h */
        FIELD tt_dtLiv2Jours AS DATE
        FIELD tt_dtArriveePF AS DATE
        */
 
        /* Calcule les trois dates de livraisons possible pour la plate forme (72h ou non) */
        /* Utilis‚e pour d‚finir le caractŠre obligatoire des OL */
 
        /* Calcule d'abord le premier jour de chargement possible depuis PLA2E */
        ASSIGN ldt_jour = ldt_jour + 1.
 
        /* Si aucun jour d'approche disponible, on prend le premier jour non feri‚ */
        IF tt_plateforme.tt_livrable THEN
            /*DO WHILE fi_ferie(ldt_jour) OR tt_plateforme.tt_tournee[WEEKDAY(ldt_jour) - 1] <> TRUE: /* ferie retourne false si dimanche(1) ou samedi(7), weekday vaux donc entre 2 et 6, et la recherche se fait entre 1 et 5 */ */
            DO WHILE fi_ferie(ldt_jour) OR tt_plateforme.tt_approche[WEEKDAY(ldt_jour) - 1] <> TRUE: /* ferie retourne false si dimanche(1) ou samedi(7), weekday vaux donc entre 2 et 6, et la recherche se fait entre 1 et 5 */
                ASSIGN ldt_jour = ldt_jour + 1.
            END.
        ELSE
            DO WHILE fi_ferie(ldt_jour):
                ASSIGN ldt_jour = ldt_jour + 1.
            END.
 
        ASSIGN tt_plateforme.tt_dtcharg = ldt_jour.
 
        /* Calcule la date d'arriv‚e … la PF */
        ASSIGN ldt_jour = ldt_jour + 1.
        DO WHILE fi_ferie(ldt_jour):
            ASSIGN ldt_jour = ldt_jour + 1.
        END.
 
        ASSIGN tt_plateforme.tt_dtArriveePF = ldt_jour.
 
        /* La premiŠre date de livraison possible, sans +2 de la livraison 72h */
        ASSIGN ldt_jour = ldt_jour + 1.
        DO WHILE fi_ferie(ldt_jour):
            ASSIGN ldt_jour = ldt_jour + 1.
        END.
 
        ASSIGN tt_plateforme.tt_dtLivCourt = ldt_jour.
 
        ASSIGN ldt_jour = ldt_jour + 2.
 
        DO WHILE fi_ferie(ldt_jour):
            ASSIGN ldt_jour = ldt_jour + 1.
        END.
 
        ASSIGN tt_plateforme.tt_dtLivLong = ldt_jour.
 
        /* Ici, la deuxiŠme boucle un peu compliqu‚e pour le calcul de la date de livraison si T2J = 2 */
        ASSIGN ldt_jour = tt_plateforme.tt_dtCharg + 1.
 
        /* Si aucun jour d'approche disponible, on prend le premier jour non feri‚ */
        IF tt_plateforme.tt_livrable THEN
            /*DO WHILE fi_ferie(ldt_jour) OR tt_plateforme.tt_tournee[WEEKDAY(ldt_jour) - 1] <> TRUE: /* ferie retourne false si dimanche(1) ou samedi(7), weekday vaux donc entre 2 et 6, et la recherche se fait entre 1 et 5 */ */
            DO WHILE fi_ferie(ldt_jour) OR tt_plateforme.tt_approche[WEEKDAY(ldt_jour) - 1] <> TRUE: /* ferie retourne false si dimanche(1) ou samedi(7), weekday vaux donc entre 2 et 6, et la recherche se fait entre 1 et 5 */
                ASSIGN ldt_jour = ldt_jour + 1.
            END.
        ELSE
            DO WHILE fi_ferie(ldt_jour):
                ASSIGN ldt_jour = ldt_jour + 1.
            END.
 
        /* Calcul d'une nouvelle date d'arrivage */
        ASSIGN ldt_jour = ldt_jour + 1.
        DO WHILE fi_ferie(ldt_jour):
            ASSIGN ldt_jour = ldt_jour + 1.
        END.
 
        ASSIGN ldt_jour = ldt_jour + 1.
        DO WHILE fi_ferie(ldt_jour):
            ASSIGN ldt_jour = ldt_jour + 1.
        END.
 
        ASSIGN tt_plateforme.tt_dtLiv2Jours = ldt_jour.
 
        /*
        PUT UNFORMATTED "Plateforme : " tt_plateforme.tt_num SKIP
                        "Date Liv Prev : " tt_plateforme.tt_dtLivPrev SKIP
                        "Date d‚part Pla2E : " tt_plateforme.tt_dtcharg SKIP
                        "Date arriv‚e PF : " tt_plateforme.tt_dtArriveePF SKIP
                        "Dt Liv Calc 1 : " tt_plateforme.tt_dtLivCourt  SKIP
                        "Dt Liv Calc 2 : " tt_plateforme.tt_dtLivLong SKIP
                        "Dt Liv Calc 2 jours : " tt_plateforme.tt_dtLiv2Jours SKIP
                        "" SKIP
                        "===================================================================" SKIP
                        "" SKIP.
        */
    END.
END.
/*OUTPUT CLOSE. */
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE initTTClient W-Win 
PROCEDURE initTTClient :
/*------------------------------------------------------------------------------
  Purpose:     Remplis les tt tt_client et tt_cliint avec les donn‚es en base avant l'import
               Finalement inutile, on garde pour l'init des tt
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
/*
 
DEFINE TEMP-TABLE tt_client LIKE gpi_pgcli
    FIELD tt_fichier  AS CHARACTER
    FIELD tt_numligne AS INTEGER
    INDEX idxcli cli_num_client cli_cp tt_fichier.
 
DEFINE TEMP-TABLE tt_cliint LIKE gpi_pgcliint
    FIELD tt_fichier  AS CHARACTER
    FIELD tt_numligne AS INTEGER
    INDEX idxcliint cli_num_client cli_cp cliint_cle tt_fichier.
 
DEFINE VARIABLE gl_importClient AS LOGICAL     NO-UNDO. /* La table contient toujours des donn‚es, donc on met un flag si elle a ‚t‚ modifi‚e */
*/
 
EMPTY TEMP-TABLE tt_client.
EMPTY TEMP-TABLE tt_cliint.
/*
FOR EACH gpi_pgcli NO-LOCK:
    CREATE tt_client.
    BUFFER-COPY gpi_pgcli TO tt_client.
END.
 
FOR EACH gpi_pgcliint NO-LOCK:
    CREATE tt_cliint.
    BUFFER-COPY gpi_pgcliint TO tt_cliint.
END.
*/
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE integrationArticles W-Win 
PROCEDURE integrationArticles :
DEFINE INPUT PARAMETER ih_repertoire AS HANDLE NO-UNDO.
 
DEFINE VARIABLE lok             AS LOGICAL   NO-UNDO.
/*DEFINE VARIABLE rep_cible       AS CHARACTER NO-UNDO.*/
DEFINE VARIABLE rep_origine     AS CHARACTER NO-UNDO.
DEFINE VARIABLE liste_dir       AS CHARACTER NO-UNDO.
DEFINE VARIABLE liste_fic       AS CHARACTER NO-UNDO.
DEFINE VARIABLE fichier         AS CHARACTER NO-UNDO.
 
DEFINE VARIABLE ii              AS INTEGER   NO-UNDO. 
 
DEFINE VARIABLE li_numLigne     AS INTEGER   NO-UNDO.
DEFINE VARIABLE lc_ligne        AS CHARACTER NO-UNDO.
 
DEFINE VARIABLE lc_format       AS CHARACTER NO-UNDO.   /* Format du fichier en cours : PGARTI/PGARTC */
DEFINE VARIABLE iMax            AS INTEGER   NO-UNDO.
 
DEFINE VARIABLE lc_repert AS CHARACTER NO-UNDO.
 
DEFINE VARIABLE lc_ficLog AS CHARACTER NO-UNDO.
 
DEFINE VARIABLE ldt_dt_integration AS DATETIME NO-UNDO.
 
DEFINE VARIABLE li_numArticles  AS INTEGER     NO-UNDO.
 
DEFINE VARIABLE llc_file AS LONGCHAR NO-UNDO.   /* M????? - NID le 15/02/19 - Astr'In */
 
ASSIGN lc_repert = ih_repertoire:SCREEN-VALUE.
 
IF TRIM(lc_repert) = ""  THEN
DO:
    /*IF lTraitementBatch = FALSE THEN /* 22/01/15 ne pas executer le msg en mode auto/batch */*/
    IF NOT gl_modeBatch THEN
        MESSAGE "Vous devez renseigner le r‚pertoire d'int‚gration des articles." VIEW-AS ALERT-BOX INFO BUTTONS OK.
    ELSE
        RUN erreur ("", 0, "Vous devez renseigner le r‚pertoire d'int‚gration des articles.").
    APPLY "ENTRY" TO ih_repertoire. 
    RETURN ERROR "".
END.
ELSE
DO:       
    RUN Objets\Files\verifie_repertoire.r ( INPUT lc_repert, INPUT TRUE, INPUT " des fichiers d'int‚gration", INPUT " acc‚der aux fichiers", INPUT TRUE, OUTPUT lOK ).
    IF NOT lok THEN
        RETURN ERROR "".
END.
 
EMPTY TEMP-TABLE tt_article.
EMPTY TEMP-TABLE tt_composant.
 
/*......................................*/
/*  r‚cup information (date/heure       */
/*......................................*/
FILE-INFO:FILE-NAME = lc_repert.
rep_origine = FILE-INFO:FULL-PATHNAME .
 
/* on r‚cupŠre tous les fichier */
RUN Objets\Files\proc-recup-files.r (  INPUT   rep_origine                ,          /* repertoire d'origine                  */
                                       INPUT   "*"                        ,          /* type de repertoire *.*  a*            */
                                       INPUT   TRUE                       ,          /* fournir les fichiers                  */
                                       INPUT   FALSE                      ,          /* fournir que les noms des fichiers     */
                                       INPUT   "*.*"                      ,          /* type de fichiers *.*  a*.txt          */
                                       INPUT   0                          ,          /* nb niv de srep 0 aucun ou ? = tous    */
                                       INPUT   liste-1:HANDLE IN FRAME f-main,       /* handle d'une selection liste sort 1   */
                                       INPUT   liste-2:HANDLE             ,          /* handle d'une selection liste sort 2   */
                                       INPUT   liste-3:HANDLE             ,          /* handle d'une selection liste sort 3   */
                                       INPUT   liste-4:HANDLE             ,          /* handle d'une selection liste sort 4   */
                                       OUTPUT  liste_dir                  ,          /* Liste des dirs                        */
                                       OUTPUT  liste_fic                  ,          /* Liste des fichiers                    */
                                       OUTPUT  lok ).                                /* traitement ok ?                       */
 
 
IF NUM-ENTRIES ( liste_fic ) = 0 OR NUM-ENTRIES ( liste_fic ) =  ? THEN
DO:  
    IF NOT gl_modeBatch THEN
        MESSAGE "Vous n'avez pas de fichier articles … int‚grer dans le r‚pertoire " + rep_origine VIEW-AS ALERT-BOX.
 
    RETURN.
END.
 
FRAME F-Criteres:MOVE-TO-BOTTOM().
FRAME F-Criteres:VISIBLE = FALSE.
 
FRAME F-jauge:VISIBLE = TRUE.
FRAME F-jauge:MOVE-TO-TOP().
 
/* jauge.*/
iMax = NUM-ENTRIES ( liste_fic ) + 1.
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Int‚gration en cours. Veuillez patienter ...").
 
RUN jauge-init IN h_jauge-ocx ( INPUT iMax ).
 
ASSIGN li_numArticles = 0.
DO ii = 1 TO NUM-ENTRIES(liste_fic) :
 
    ASSIGN
    fichier = ENTRY ( ii, liste_fic )
    li_numLigne = 0
    lok = TRUE
    ldt_dt_integration = fi_recupDateHeure(fichier).
 
    RUN jauge-next IN h_jauge-ocx.
 
    ASSIGN lc_format = "".
 
    INPUT STREAM st_article FROM VALUE(fichier) CONVERT TARGET "ibm850".
    REPEAT:
        IMPORT STREAM st_article UNFORMATTED lc_ligne.
        li_numLigne = li_numLigne + 1.
 
        IF li_numLigne = 1 THEN
        DO:
            IF NUM-ENTRIES(lc_ligne, "|") = 3 OR NUM-ENTRIES(lc_ligne, "|") = 4 THEN
                ASSIGN lc_format = "PGARTC".
            /*IF NUM-ENTRIES(lc_ligne, "|") = 13 OR NUM-ENTRIES(lc_ligne, "|") = 14 THEN */
            IF NUM-ENTRIES(lc_ligne, "|") >= 13 THEN
                ASSIGN lc_format = "PGARTI".
 
            IF lc_format = "" THEN
            DO:
                ASSIGN lok = FALSE.
                /* M????? - NID le 15/02/19 - Astr'In */
                /* On va v‚rifier si le fichier est complŠtement vide. Si c'est le cas, on l'enregistre pour archivage sans le remonter en erreur */
                COPY-LOB FROM FILE fichier TO llc_file NO-ERROR.
                IF TRIM(REPLACE(REPLACE(REPLACE(REPLACE(llc_file, " ", ""), CHR(10), ""), CHR(9), ""), CHR(13), "")) = "" THEN
                    RUN createFichier(fichier, rep_origine).
                ELSE
                /* Fin M????? - NID le 15/02/19 - Astr'In */
                    RUN erreur(fichier, 1, "Format du fichier non reconnu.").
            END.
        END.
 
        CASE lc_format:
            /* Articles */
            WHEN "PGARTI" THEN
            DO:
                IF NUM-ENTRIES(lc_ligne, "|") < 13 THEN
                DO:
                    RUN erreur(fichier, li_numLigne, "Erreur de format pour la ligne d'article " + lc_ligne).
                    ASSIGN lok = FALSE.
                    NEXT.
                END.
 
                CREATE tt_article.
                ASSIGN
                tt_article.tt_action                = ENTRY(1, lc_ligne, "|")
                tt_article.ARTI_NUM_ARTICLE         = INTEGER(ENTRY(2, lc_ligne, "|"))
                tt_article.ARTI_CODE_CONSTRUCTEUR   = ENTRY(3, lc_ligne, "|")
                tt_article.ARTI_REFERENCE           = ENTRY(4, lc_ligne, "|")
                tt_article.arti_designation         = ENTRY(5, lc_ligne, "|")
                tt_article.arti_poids_brut          = INTEGER(ENTRY(6, lc_ligne, "|"))
                tt_article.arti_hauteur             = INTEGER(ENTRY(7, lc_ligne, "|"))
                tt_article.arti_largeur             = INTEGER(ENTRY(8, lc_ligne, "|"))
                tt_article.arti_profondeur          = INTEGER(ENTRY(9, lc_ligne, "|"))
                tt_article.arti_coefficient         = DECIMAL(ENTRY(10, lc_ligne, "|"))
                tt_article.arti_coef_gerbage        = INTEGER(ENTRY(11, lc_ligne, "|"))
                tt_article.arti_division            = ENTRY(12, lc_ligne, "|")
                tt_article.arti_code_ensemble       = ENTRY(13, lc_ligne, "|")
                /* Ajout - NID le 06/10/16 - Int‚gration des nouveaux champs */
                tt_article.arti_atco                = IF NUM-ENTRIES(lc_ligne, "|") > 13 THEN ENTRY(14, lc_ligne, "|") ELSE ""
                tt_article.arti_famille             = IF NUM-ENTRIES(lc_ligne, "|") > 14 THEN ENTRY(15, lc_ligne, "|") ELSE ""
                /* Fin Ajout - NID le 06/10/16 */
                /* Ajout - NID le 19/12/16 - M00015347 - Ajout du code EAN */
                tt_article.arti_ean                 = IF NUM-ENTRIES(lc_ligne, "|") > 15 AND TRIM(ENTRY(16, lc_ligne, "|")) <> "" THEN ENTRY(16, lc_ligne, "|") ELSE ENTRY(2, lc_ligne, "|")
                /* Fin Ajout - NID le 19/12/16 */
                tt_article.tt_dimensions            = STRING(tt_article.ARTI_HAUTEUR) + " x " + STRING(tt_article.ARTI_LARGEUR) + " x " + STRING(tt_article.ARTI_PROFONDEUR) 
                tt_article.tt_dt_export             = ldt_dt_integration
                tt_article.tt_fichier               = fichier   
                tt_article.tt_numligne              = li_numLigne   NO-ERROR.
 
                IF ERROR-STATUS:ERROR THEN
                DO:
                    ASSIGN lok = FALSE.
                    RUN erreur (fichier, li_numLigne, "Erreur … l'import de la ligne d'article " + lc_ligne).
                    DELETE tt_article.
                END.
                ELSE
                    ASSIGN li_numArticles = li_numArticles + 1.
            END.
            /* Articles compos‚s */
            WHEN "PGARTC" THEN
            DO:
                IF NUM-ENTRIES(lc_ligne, "|") < 3 THEN
                DO:
                    RUN erreur (fichier, li_numLigne, "Erreur de format pour la ligne de composant article " + lc_ligne).
                    ASSIGN lok = FALSE.
                    NEXT.
                END.
 
                CREATE tt_composant.
                ASSIGN
                tt_composant.ARTI_NUM_ARTICLE     = INTEGER(ENTRY(1, lc_ligne, "|"))
                tt_composant.ARTC_NUM_COMPOSANT   = INTEGER(ENTRY(2, lc_ligne, "|"))
                tt_composant.ARTC_QUANTITE        = INTEGER(ENTRY(3, lc_ligne, "|")) 
                tt_composant.tt_fichier           = fichier   
                tt_composant.tt_numligne          = li_numLigne   NO-ERROR.
                IF ERROR-STATUS:ERROR THEN
                DO:
                    /* A revoir */
                    ASSIGN lok = FALSE.
                    RUN erreur (fichier, li_numLigne, "Erreur … l'import de la ligne de composant article " + lc_ligne).
                    DELETE tt_composant.
                END.
            END.
        END CASE.
    END.
    INPUT STREAM st_article CLOSE.
END.
 
/* fin de traitement */
RUN jauge-init IN  h_jauge-ocx ( INPUT iMax).
RUN jauge-fin IN h_jauge-ocx.
 
RUN controleArticles (li_numArticles).
 
FRAME F-jauge:MOVE-TO-BOTTOM().
FRAME F-jauge:VISIBLE = FALSE.
 
FRAME F-Criteres:VISIBLE = TRUE.
FRAME F-Criteres:MOVE-TO-TOP().
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE integrationOL W-Win 
PROCEDURE integrationOL :
DEFINE INPUT PARAMETER ih_repertoire AS HANDLE NO-UNDO.
 
DEFINE VARIABLE lok             AS LOGICAL   NO-UNDO.
DEFINE VARIABLE rep_origine     AS CHARACTER NO-UNDO.
DEFINE VARIABLE liste_dir       AS CHARACTER NO-UNDO.
DEFINE VARIABLE liste_fic       AS CHARACTER NO-UNDO.
DEFINE VARIABLE fichier         AS CHARACTER NO-UNDO.
 
DEFINE VARIABLE ii              AS INTEGER   NO-UNDO. 
 
DEFINE VARIABLE li_numLigne     AS INTEGER   NO-UNDO.
DEFINE VARIABLE lc_ligne        AS CHARACTER NO-UNDO.
 
DEFINE VARIABLE lc_format       AS CHARACTER NO-UNDO.   /* Format du fichier en cours : PGARTI/PGARTC */
DEFINE VARIABLE iMax            AS INTEGER   NO-UNDO.
 
DEFINE VARIABLE lc_repert AS CHARACTER NO-UNDO.
 
DEFINE VARIABLE ldt_dt_integration AS DATETIME NO-UNDO.
 
/*DEFINE VARIABLE ldt_dt_liv_souhait AS DATE        NO-UNDO. */
/*DEFINE VARIABLE li_dtLivSouhait    AS INTEGER     NO-UNDO. */
/*DEFINE VARIABLE lcDtLivSouhait     AS CHARACTER   NO-UNDO. */
 
DEF VAR integrer_detail     AS  LOG NO-UNDO.
 
DEFINE VARIABLE ll_mode89   AS LOGICAL     NO-UNDO.
 
/* M17110 - NID le 25/04/18 - Astr'In */
DEFINE VARIABLE lc_date  AS CHARACTER   NO-UNDO.
DEFINE VARIABLE li_date  AS INTEGER     NO-UNDO.
DEFINE VARIABLE lc_heure AS CHARACTER   NO-UNDO.
DEFINE VARIABLE li_heure AS INTEGER     NO-UNDO.
/* Fin M17110 - NID le 25/04/18 - Astr'In */
 
DEFINE VARIABLE llc_file AS LONGCHAR NO-UNDO.   /* M????? - NID le 15/02/19 - Astr'In */
 
ASSIGN lc_repert = ih_repertoire:SCREEN-VALUE.
 
IF TRIM(lc_repert) = ""  THEN
DO:
    IF NOT gl_modeBatch THEN
        MESSAGE "Vous devez renseigner le r‚pertoire d'import des OL." VIEW-AS ALERT-BOX INFO BUTTONS OK.
    ELSE
        RUN erreur ("", 0, "Vous devez renseigner le r‚pertoire d'import des OL.").
    APPLY "ENTRY" TO ih_repertoire. 
    RETURN ERROR "".
END.
ELSE
DO:       
    RUN Objets\Files\verifie_repertoire.r ( INPUT lc_repert, INPUT TRUE, INPUT " des fichiers d'int‚gration", INPUT " acc‚der aux fichiers", INPUT TRUE, OUTPUT lOK ).
    IF NOT lok THEN
        RETURN ERROR "".
END.
 
EMPTY TEMP-TABLE tt_ol.
EMPTY TEMP-TABLE tt_ligneOL.
EMPTY TEMP-TABLE tt_interClient.    /* Les interlocuteurs Eberhardt */
/*EMPTY TEMP-TABLE tt_client.   -- On ne la vide plus, elle contient les donn‚es en base */
 
/*......................................*/
/*  r‚cup information (date/heure       */
/*......................................*/
FILE-INFO:FILE-NAME = lc_repert.
rep_origine = FILE-INFO:FULL-PATHNAME.
 
/* on r‚cupŠre tous les fichier */
RUN Objets\Files\proc-recup-files.r (  INPUT   rep_origine                ,          /* repertoire d'origine                  */
                                       INPUT   "*"                        ,          /* type de repertoire *.*  a*            */
                                       INPUT   TRUE                       ,          /* fournir les fichiers                  */
                                       INPUT   FALSE                      ,          /* fournir que les noms des fichiers     */
                                       INPUT   "*.*"                      ,          /* type de fichiers *.*  a*.txt          */
                                       INPUT   0                          ,          /* nb niv de srep 0 aucun ou ? = tous    */
                                       INPUT   liste-1:HANDLE IN FRAME f-main,       /* handle d'une selection liste sort 1   */
                                       INPUT   liste-2:HANDLE             ,          /* handle d'une selection liste sort 2   */
                                       INPUT   liste-3:HANDLE             ,          /* handle d'une selection liste sort 3   */
                                       INPUT   liste-4:HANDLE             ,          /* handle d'une selection liste sort 4   */
                                       OUTPUT  liste_dir                  ,          /* Liste des dirs                        */
                                       OUTPUT  liste_fic                  ,          /* Liste des fichiers                    */
                                       OUTPUT  lok ).                                /* traitement ok ?                       */
 
 
IF NUM-ENTRIES ( liste_fic ) = 0 OR NUM-ENTRIES ( liste_fic ) =  ? THEN
DO:  
    IF NOT gl_modeBatch THEN
        MESSAGE "Vous n'avez pas de fichier O.L. … int‚grer dans le r‚pertoire " rep_origine VIEW-AS ALERT-BOX INFO BUTTONS OK.
    RETURN.
END.
 
/*.jauge.*/
iMax = NUM-ENTRIES ( liste_fic ) + 1.
 
FRAME F-Criteres:MOVE-TO-BOTTOM().
FRAME F-Criteres:VISIBLE = FALSE.
 
FRAME F-jauge:VISIBLE = TRUE.
FRAME F-jauge:MOVE-TO-TOP().
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Int‚gration en cours. Veuillez patienter ...").
 
RUN jauge-init IN h_jauge-ocx ( INPUT iMax ).
 
DO ii = 1 TO NUM-ENTRIES(liste_fic) :
 
    ASSIGN
    fichier = ENTRY ( ii, liste_fic )
    li_numLigne = 0
    ldt_dt_integration = fi_recupDateHeure(fichier).
 
    IF ldt_dt_integration = ? THEN ASSIGN ldt_dt_integration = NOW. /* M17110 - NID le 25/04/18 - Pour avoir toujours une date heure valide */
 
    RUN jauge-next IN h_jauge-ocx.
 
    ASSIGN li_numLigne = 0.
    INPUT STREAM st_article FROM VALUE(fichier) CONVERT TARGET "ibm850".
    REPEAT:
        IMPORT STREAM st_article UNFORMATTED lc_ligne.
        li_numLigne = li_numLigne + 1.
 
        /* Contr“le si la premiŠre entr‚e correspond a un format connu, afin d'‚viter de lire tout un fichier qui n'a rien … voir */
        IF li_numLigne = 1 AND ENTRY(1, lc_ligne, "|") <> "ENT" AND 
                               ENTRY(1, lc_ligne, "|") <> "LIG" AND 
                               ENTRY(1, lc_ligne, "|") <> "COM" AND 
                               ENTRY(1, lc_ligne, "|") <> "INT" THEN
        DO:
            /* M????? - NID le 15/02/19 - Astr'In */
            /* On va v‚rifier si le fichier est complŠtement vide. Si c'est le cas, on l'enregistre pour archivage sans le remonter en erreur */
            COPY-LOB FROM FILE fichier TO llc_file NO-ERROR.
            IF TRIM(REPLACE(REPLACE(REPLACE(REPLACE(llc_file, " ", ""), CHR(10), ""), CHR(9), ""), CHR(13), "")) = "" THEN
                RUN createFichier(fichier, rep_origine).
            ELSE
            /* Fin M????? - NID le 15/02/19 - Astr'In */
                RUN erreur (fichier, 1, "Format du fichier non reconnu.").
            LEAVE.
        END.
 
        IF TRIM(lc_ligne) = "" THEN NEXT.
 
        ASSIGN lc_format = ENTRY(1, lc_ligne, "|").
 
        CASE lc_format:
            /* Entˆte */
            WHEN "ENT" THEN
                RUN integrationOl_ENT(INPUT  lc_ligne,
                                      INPUT  fichier,
                                      INPUT  li_numLigne,
                                      INPUT  ldt_dt_integration,
                                      OUTPUT integrer_detail,
                                      OUTPUT ll_mode89).
 
            /* Detail */
            WHEN "LIG" THEN
                RUN integrationOL_LIG(INPUT  integrer_detail,
                                      INPUT  ll_mode89,
                                      INPUT  fichier,
                                      INPUT  li_numLigne,
                                      INPUT  lc_ligne).
 
            /* Commentaire */
            WHEN "COM" THEN
                RUN integrationOL_COM(INPUT  integrer_detail,
                                      INPUT  ll_mode89,
                                      INPUT  fichier,
                                      INPUT  li_numLigne,
                                      INPUT  lc_ligne).
 
            /* Intervenants */
            WHEN "INT" THEN
                RUN integrationOL_INT(INPUT  integrer_detail,
                                      INPUT  ll_mode89,
                                      INPUT  fichier,
                                      INPUT  li_numLigne,
                                      INPUT  lc_ligne).
 
            OTHERWISE
                RUN erreur (fichier, li_numLigne, "Format de la ligne d'OL " + QUOTER(lc_ligne) + " non reconnu.").
        END CASE.
    END.
 
    INPUT STREAM st_article CLOSE.
END.
 
/* fin de traitement */
RUN jauge-init IN  h_jauge-ocx ( INPUT iMax).
RUN jauge-fin IN h_jauge-ocx.
 
FRAME F-jauge:MOVE-TO-BOTTOM().
FRAME F-jauge:VISIBLE = FALSE.
 
FRAME F-Criteres:VISIBLE = TRUE.
FRAME F-Criteres:MOVE-TO-TOP().
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE integrationOL_COM W-Win 
PROCEDURE integrationOL_COM :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER integrer_detail AS LOGICAL     NO-UNDO.
DEFINE INPUT  PARAMETER il_mode89       AS LOGICAL     NO-UNDO. /* Si mode tprs 89, on n'integre pas l'OL, mais on ne g‚nŠre pas d'erreur */
DEFINE INPUT  PARAMETER fichier         AS CHARACTER   NO-UNDO.
DEFINE INPUT  PARAMETER li_numLigne     AS INTEGER     NO-UNDO.
DEFINE INPUT  PARAMETER lc_ligne        AS CHARACTER   NO-UNDO.
 
/* M00014463 - NID le 23/05/16 - Astr'In - Nouveau format pour les commentaires */
/*
IF NUM-ENTRIES(lc_ligne, "|") < 2 THEN
DO:
    RUN erreur (fichier, li_numLigne, "Erreur de format pour la ligne de commentaire " + lc_ligne ).
    NEXT.
END.
 
IF NOT AVAILABLE tt_ol THEN
DO:
    RUN erreur (fichier, li_numLigne, "Aucune commande en cours de traitement, impossible d'affecter le commentaire " + lc_ligne).
    NEXT.
END.
*/
 
IF il_mode89 THEN RETURN.
 
IF NUM-ENTRIES(lc_ligne, "|") < 4 THEN
DO:
    RUN erreur (fichier, li_numLigne, "Erreur de format pour la ligne de commentaire " + lc_ligne ).
    RETURN.
    /*NEXT. */
END.
 
IF NOT integrer_detail THEN
DO:
    RUN erreur (fichier, li_numLigne, "Cette ligne de d‚tail est li‚e … un OL exclus " + lc_ligne).
    RETURN.
    /*NEXT. */
END.
 
/* FIND FIRST tt_ol WHERE tt_ol.OL_NUM_OL = ENTRY(2, lc_ligne, "|") NO-ERROR. */
FIND FIRST tt_ol WHERE tt_ol.OL_NUM_OL = ENTRY(2, lc_ligne, "|") AND tt_ol.tt_fichier = fichier NO-ERROR.
 
IF NOT AVAILABLE tt_ol THEN
DO:
    RUN erreur (fichier, li_numLigne, "L'OL Nø " +  ENTRY(2, lc_ligne, "|") + "n'est pas en cours de cr‚ation, impossible d'affecter le commentaire " + lc_ligne).
    RETURN.
    /*NEXT. */
END.
/* Fin M00014463 - NID le 23/05/16 - Astr'In */
 
ASSIGN tt_ol.OL_COMMENTAIRE = IF tt_ol.OL_COMMENTAIRE = ""
                              THEN ENTRY(4, lc_ligne, "|")
                              ELSE tt_ol.OL_COMMENTAIRE + CHR(10) + ENTRY(4, lc_ligne, "|") NO-ERROR.
IF ERROR-STATUS:ERROR THEN
    RUN erreur (fichier, li_numLigne, "Erreur … l'import de la ligne de commentaire " + lc_ligne).
 
/* Ajout NID le 20/10/16 - Surtout pas, ‡a va ‚craser le commentaire client modifiable. Il va falloir r‚cup‚rer le commentaire directement dans l'OL */
/*
FIND FIRST tt_client WHERE tt_client.cli_num_client = tt_ol.ol_num_client AND tt_client.cli_cp = tt_ol.ol_cp_livraison AND tt_client.tt_fichier = fichier NO-ERROR.
IF NOT AVAILABLE tt_client THEN
DO:
    CREATE tt_client.
    ASSIGN
    tt_client.cli_num_client = tt_ol.ol_num_client
    tt_client.cli_cp         = tt_ol.ol_cp_livraison
    tt_client.tt_fichier     = fichier.
    tt_client.tt_numligne    = li_numLigne   NO-ERROR.
END.
ASSIGN
gl_importClient      = TRUE
tt_client.cli_com = tt_ol.ol_commentaire.
*/
/* Fin Ajout NID le 20/10/16 */
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE integrationOL_ENT W-Win 
PROCEDURE integrationOL_ENT :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER lc_ligne            AS CHARACTER   NO-UNDO.
DEFINE INPUT  PARAMETER fichier             AS CHARACTER   NO-UNDO.
DEFINE INPUT  PARAMETER li_numLigne         AS INTEGER     NO-UNDO.
/* M17110 - NID le 25/04/18 - Astr'In */
/*
DEFINE INPUT  PARAMETER ldt_dt_integration  AS DATE        NO-UNDO.
*/
DEFINE INPUT  PARAMETER ldt_dt_integration  AS DATETIME    NO-UNDO.
/* Fin M17110 - NID le 25/04/18 - Astr'In */
DEFINE OUTPUT PARAMETER integrer_detail     AS LOGICAL     NO-UNDO.
DEFINE OUTPUT PARAMETER ol_mode89           AS LOGICAL     NO-UNDO.
 
DEFINE VARIABLE ldt_dt_liv_souhait          AS DATE        NO-UNDO.
DEFINE VARIABLE li_dtLivSouhait             AS INTEGER     NO-UNDO.
DEFINE VARIABLE lcDtLivSouhait              AS CHARACTER   NO-UNDO.

DEFINE VARIABLE ll_messagerie AS LOGICAL NO-UNDO.   /* M19873 - NID le 06/04/20 - Astr'In */

ASSIGN
ol_mode89       = FALSE
integrer_detail = TRUE
ll_messagerie   = FALSE /* M19873 - NID le 06/04/20 - Astr'In */
.
 
/* Modification du 30/05/2016 - Suppression du champ en double en 7Šme position, d‚calage des entry de 8->22 … 7->21 */
/*
IF NUM-ENTRIES(lc_ligne, "|") < 22 THEN
*/
IF NUM-ENTRIES(lc_ligne, "|") < 21 THEN
DO:
    RUN erreur (fichier, li_numLigne, "Erreur de format pour la ligne d'entˆte " + lc_ligne).
    /*NEXT. */
    RETURN.
END.
 
/*
    CHC 21/10/2016
    V‚rification de la pr‚sence ou non du num‚ro de l'OL dans la liste des exclusions
*/
/* Ajout NID le 26/10/16 - Si OL exclus, on crŠs quand mˆme l'OL en initialisant le champ ol_ancien_ol *
IF CAN-FIND ( FIRST TT_OL_EXCLUS WHERE TT_OL_EXCLUS.num_ol = ENTRY(2, lc_ligne, "|") ) THEN
DO:
    integrer_detail = FALSE.
 
    RUN erreur (fichier, li_numLigne, "Cette ligne correspond … un OL exclus " + lc_ligne).
    /*NEXT. */
    RETURN.
END.
*/
 
IF ENTRY(6, lc_ligne, "|") = "89" THEN
DO:
    ASSIGN ol_mode89 = TRUE.
 
    RETURN.
END.

/* M19873 - NID le 06/04/20 - Astr'In */
IF ENTRY(6, lc_ligne, "|") = "91" THEN
    ASSIGN ll_messagerie = TRUE.
/* Fin M19873 - NID le 06/04/20 - Astr'In */
 
ASSIGN  ldt_dt_liv_souhait  = ?
        lcDtLivSouhait      = IF NUM-ENTRIES(lc_ligne, "|") > 22 THEN ENTRY(23, lc_ligne, "|") ELSE "".  /* Ajout NID le 06/10/16 - M00014894 - Astr'In - Date de livraison souhait‚e */
 
IF lcDtLivSouhait <> "" THEN
DO:
    ASSIGN li_dtLivSouhait = INTEGER(lcDtLivSouhait) NO-ERROR.
    /* NID le 08/03/18 - Astr'In */
    /*
    IF NOT ERROR-STATUS:ERROR AND LENGTH(lcDtLivSouhait) = 8 THEN
        ASSIGN ldt_dt_liv_souhait = DATE(INTEGER(SUBSTRING(lcDtLivSouhait,5, 2)), INTEGER(SUBSTRING(lcDtLivSouhait,7, 2)), INTEGER(SUBSTRING(lcDtLivSouhait,1, 4))).
    */
    IF NOT ERROR-STATUS:ERROR AND LENGTH(lcDtLivSouhait) = 8 THEN
    DO:
        IF INTEGER(SUBSTRING(lcDtLivSouhait,1, 4)) > YEAR(TODAY) + 5 OR INTEGER(SUBSTRING(lcDtLivSouhait,1, 4)) < YEAR(TODAY) - 5 THEN
        DO:
            RUN erreur (fichier, li_numLigne, "Erreur … l'import de la ligne d'entˆte " + lc_ligne + " - Ann‚e de date incoh‚rente.").
            RETURN "".
        END.
        ELSE
            IF INTEGER(SUBSTRING(lcDtLivSouhait,5, 2)) > 12 OR INTEGER(SUBSTRING(lcDtLivSouhait,5, 2)) < 1 THEN
            DO:
                RUN erreur (fichier, li_numLigne, "Erreur … l'import de la ligne d'entˆte " + lc_ligne + " - Mois de date incoh‚rent.").
                RETURN "".
            END.
            ELSE
                IF INTEGER(SUBSTRING(lcDtLivSouhait,7, 2)) > 31 OR INTEGER(SUBSTRING(lcDtLivSouhait,7, 2)) < 1 THEN
                DO:
                    RUN erreur (fichier, li_numLigne, "Erreur … l'import de la ligne d'entˆte " + lc_ligne + " - Jour de date incoh‚rente.").
                    RETURN "".
                END.
            ELSE
                ASSIGN ldt_dt_liv_souhait = DATE(INTEGER(SUBSTRING(lcDtLivSouhait,5, 2)), INTEGER(SUBSTRING(lcDtLivSouhait,7, 2)), INTEGER(SUBSTRING(lcDtLivSouhait,1, 4))).
    END.        
END.
 
CREATE tt_ol.
ASSIGN
tt_ol.OL_NUM_OL                 = ENTRY(2, lc_ligne, "|")
tt_ol.OL_GESTIONNAIRE_COMMANDES = ENTRY(3, lc_ligne, "|")
/* Devrait ˆtre en Integer, ne correspond pas au dossier *
tt_ol.OL_NUM_CLIENT             = INTEGER(ENTRY(4, lc_ligne, "|"))
*/
tt_ol.OL_NUM_CLIENT             = ENTRY(4, lc_ligne, "|")
tt_ol.OL_REF_CLIENT             = ENTRY(5, lc_ligne, "|")
tt_ol.OL_MODE_TRANSPORT         = ENTRY(6, lc_ligne, "|")
/* Modification du 30/05/2016 - Suppression du champ en double en 7Šme position, d‚calage des entry de 8->22 … 7->21 */
tt_ol.OL_TITRE_COMMANDE         = ENTRY(7, lc_ligne, "|")
tt_ol.OL_NOM_COMMANDE           = ENTRY(8, lc_ligne, "|")
tt_ol.OL_ADRESSE1_COMMANDE      = ENTRY(9, lc_ligne, "|")
tt_ol.OL_ADRESSE2_COMMANDE      = ENTRY(10, lc_ligne, "|")
tt_ol.OL_CP_COMMANDE            = ENTRY(11, lc_ligne, "|")
tt_ol.OL_VILLE_COMMANDE         = ENTRY(12, lc_ligne, "|")
tt_ol.OL_TITRE_LIVRAISON        = ENTRY(13, lc_ligne, "|")
tt_ol.OL_NOM_LIVRAISON          = ENTRY(14, lc_ligne, "|")
tt_ol.OL_ADRESSE1_LIVRAISON     = ENTRY(15, lc_ligne, "|")
tt_ol.OL_ADRESSE2_LIVRAISON     = ENTRY(16, lc_ligne, "|")
tt_ol.OL_CP_LIVRAISON           = ENTRY(17, lc_ligne, "|")
tt_ol.OL_VILLE_LIVRAISON        = ENTRY(18, lc_ligne, "|")
/* Ces trois champs devraient ˆtre des Logical, mais sont d‚clar‚s en CHAR */
tt_ol.OL_RDV_OBLIGATOIRE        = ENTRY(19, lc_ligne, "|")
tt_ol.OL_HAYON_OBLIGATOIRE      = ENTRY(20, lc_ligne, "|")
/* Ajout NID le 26/10/16 - On ne tiens plus compte de cette information - le 27/10/16 - aprŠs r‚flexion, si, si le client n'existe pas */
tt_ol.OL_LOT_OBLIGATOIRE        = ENTRY(21, lc_ligne, "|")
/*tt_ol.OL_LOT_OBLIGATOIRE        = "N" */
/* Fin Ajout NID le 26/10/16 */ 
tt_ol.OL_CROSSDOCK              = IF NUM-ENTRIES(lc_ligne, "|") > 21 THEN ENTRY(22, lc_ligne, "|") ELSE ""  /* Ajout NID le 29/09/16 - M00014894 - Astr'In - Gestion Crossdock */
 
tt_ol.ol_date_liv_souhait       = ldt_dt_liv_souhait
 
/* tt_ol.ol_ancien_ol              = CAN-FIND ( FIRST TT_OL_EXCLUS WHERE TT_OL_EXCLUS.num_ol = tt_ol.ol_num_ol )   /* Ajout NID le 26/10/16 */ */
tt_ol.ol_ancien_ol              = CAN-FIND ( FIRST GPI_PGI_ENTPREP WHERE GPI_PGI_ENTPREP.NUM_OL = tt_ol.ol_num_ol)  /* Ajout NID le 28/10/16 */
 
/*
tt_ol.OL_RDV_OBLIGATOIRE        = LOGICAL(ENTRY(20, lc_ligne, "|"), "O/N")
tt_ol.OL_HAYON_OBLIGATOIRE      = LOGICAL(ENTRY(21, lc_ligne, "|"), "O/N")
tt_ol.OL_LOT_OBLIGATOIRE        = LOGICAL(ENTRY(22, lc_ligne, "|"), "O/N")
*/
tt_ol.ACT_ID                    = ?
/*
tt_ol.OL_DATE                   = TODAY
*/
tt_ol.ol_date-heure_creation    = NOW
tt_ol.OL_COMMENTAIRE            = "" 
 
/* Ajout NID le 26/10/16, mˆme si je suis s–r de l'avoir fait hier */
/*
tt_ol.ol_cp_idx                 = ""
*/
tt_ol.ol_cp_idx                 = SUBSTRING(tt_ol.OL_CP_LIVRAISON, 1, 3)
tt_ol.ol_dep_livraison          = SUBSTRING(tt_ol.OL_CP_LIVRAISON, 1, 2)
/* Fin Ajout NID le 26/10/16 */
tt_ol.ol_date_chg               = ?
tt_ol.ol_date_RDV               = ?
/* tt_ol.ol_dep_livraison          = "" */
tt_ol.ol_heure_chg              = ""
tt_ol.ol_heure_RDV              = ""
tt_ol.ol_jour_livraison         = ?
tt_ol.ol_ml                     = 0
tt_ol.ol_mode_liv               = ""
tt_ol.ol_num_affaire            = ""
tt_ol.ol_num_pf                 = ""
tt_ol.ol_poids                  = 0
tt_ol.ol_qte                    = 0
tt_ol.ol_type_int               = ""
 
tt_ol.tt_dt_export              = ldt_dt_integration
/* M17110 - NID le 25/04/18 - Astr'In - S‚parer la date du fichier de la date r‚elle de cr‚ation. On mettait en date de cr‚ation les date/heure du fichier. Plus maintenant */
tt_ol.ol_date_fichier           = DATE(ldt_dt_integration)
tt_ol.ol_heure_fichier          = STRING(INTEGER(TRUNCATE(MTIME(ldt_dt_integration) / 1000, 0)), "hh:mm")
/* Fin M17110 - NID le 25/04/18 - Astr'In */
tt_ol.tt_fichier                = fichier   
tt_ol.tt_numligne               = li_numLigne   NO-ERROR.
 
IF ERROR-STATUS:ERROR THEN
DO:
    RUN erreur (fichier, li_numLigne, "Erreur … l'import de la ligne d'entˆte " + lc_ligne).
    DELETE tt_ol.
END.
/* Ajout NID le 26/10/16 - On ne met plus … jour les clients de puis les OL */
/*
ELSE
DO:
    /* L'identifiant du client (merci l'ancien systŠme) est le cp_idx, un CP raccourcis */
    /* FIND FIRST tt_client WHERE tt_client.cli_num_client = tt_ol.ol_num_client AND tt_client.cli_cp = tt_ol.ol_cp_livraison AND tt_client.tt_fichier = fichier NO-ERROR. */
    FIND FIRST tt_client WHERE tt_client.cli_num_client = tt_ol.ol_num_client AND tt_client.cli_cp = tt_ol.ol_cp_idx AND tt_client.tt_fichier = fichier NO-ERROR.
    IF NOT AVAILABLE tt_client THEN
    DO:
        CREATE tt_client.
        ASSIGN
        tt_client.cli_num_client = tt_ol.ol_num_client
        /*tt_client.cli_cp         = tt_ol.ol_cp_livraison */
        tt_client.cli_cp         = tt_ol.ol_cp_idx
        tt_client.tt_origine     = "OL"
        tt_client.tt_fichier     = fichier
        tt_client.tt_numligne    = li_numLigne   NO-ERROR.
    END.
    ASSIGN
    tt_client.cli_hayon  = (tt_ol.ol_hayon_obligatoire = "O")
    tt_client.cli_dderdv = (tt_ol.ol_rdv_obligatoire = "O")
    tt_client.cli_lot    = (tt_ol.ol_lot_obligatoire = "O").
END.
*/
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE integrationOL_INT W-Win 
PROCEDURE integrationOL_INT :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER integrer_detail AS LOGICAL     NO-UNDO. 
DEFINE INPUT  PARAMETER il_mode89       AS LOGICAL     NO-UNDO. /* Si mode tprs 89, on n'integre pas l'OL, mais on ne g‚nŠre pas d'erreur */
DEFINE INPUT  PARAMETER fichier         AS CHARACTER   NO-UNDO.
DEFINE INPUT  PARAMETER li_numLigne     AS INTEGER     NO-UNDO.
DEFINE INPUT  PARAMETER lc_ligne        AS CHARACTER   NO-UNDO.
 
IF il_mode89 THEN RETURN.
 
IF NOT integrer_detail THEN
DO:
    RUN erreur (fichier, li_numLigne, "Cette ligne de d‚tail est li‚e … un OL exclus " + lc_ligne).
    RETURN .
    /* NEXT. */
END.
 
/* V‚rifie que l'intervenant n'est pas d‚j… pr‚sent dans le fichier */
FIND FIRST tt_interClient WHERE tt_interClient.pgic_num_client  = ENTRY(2, lc_ligne, "|")
                            AND tt_interClient.pgic_type_int    = ENTRY(3, lc_ligne, "|")
                            AND tt_interClient.tt_fichier       = fichier NO-ERROR.
 
IF NOT AVAILABLE tt_interClient THEN
DO:
    CREATE tt_interClient.
    ASSIGN
    tt_interClient.pgic_num_client  = ENTRY(2, lc_ligne, "|")
    tt_interClient.pgic_type_int    = ENTRY(3, lc_ligne, "|")
    tt_interClient.pgic_nom         = ENTRY(4, lc_ligne, "|")
    tt_interClient.pgic_tel         = ENTRY(5, lc_ligne, "|")
    tt_interClient.pgic_fax         = ENTRY(6, lc_ligne, "|")
    tt_interClient.pgic_mail        = ENTRY(7, lc_ligne, "|")
    tt_interClient.pgic_portable    = ENTRY(8, lc_ligne, "|")
    tt_interClient.pgic_fonction    = ENTRY(9, lc_ligne, "|")
    tt_interClient.tt_fichier       = fichier
    tt_interClient.tt_numligne      = li_numLigne   NO-ERROR.
END.
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE integrationOL_LIG W-Win 
PROCEDURE integrationOL_LIG :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER integrer_detail AS LOGICAL     NO-UNDO.
DEFINE INPUT  PARAMETER il_mode89       AS LOGICAL     NO-UNDO. /* Si mode tprs 89, on n'integre pas l'OL, mais on ne g‚nŠre pas d'erreur */
DEFINE INPUT  PARAMETER fichier         AS CHARACTER   NO-UNDO.
DEFINE INPUT  PARAMETER li_numLigne     AS INTEGER     NO-UNDO.
DEFINE INPUT  PARAMETER lc_ligne        AS CHARACTER   NO-UNDO.
 
IF il_mode89 THEN RETURN.
 
IF NUM-ENTRIES(lc_ligne, "|") < 5 THEN
DO:
    RUN erreur (fichier, li_numLigne, "Erreur de format pour la ligne de d‚tail " + lc_ligne).
    /*NEXT. */
    RETURN.
END.
 
IF NOT integrer_detail THEN
DO:
    RUN erreur (fichier, li_numLigne, "Cette ligne de d‚tail est li‚e … un OL exclus " + lc_ligne).
    /*NEXT. */
    RETURN.
END.
 
CREATE tt_ligneOL.
 
ASSIGN
tt_ligneOL.OL_NUM_OL            = ENTRY(2, lc_ligne, "|")
tt_ligneOL.DETAILOL_NUM_LIGNE   = INTEGER(ENTRY(3, lc_ligne, "|"))
tt_ligneOL.ARTI_NUM_ARTICLE     = INTEGER(ENTRY(4, lc_ligne, "|"))
tt_ligneOL.DETAILOL_QUANTITE    = INTEGER(ENTRY(5, lc_ligne, "|")) 
tt_ligneOL.tt_fichier           = fichier
tt_ligneOL.tt_numligne          = li_numLigne   NO-ERROR.
 
IF ERROR-STATUS:ERROR THEN
DO:
    RUN erreur (fichier, li_numLigne, "Erreur … l'import de la ligne de d‚tail " + lc_ligne).
    DELETE tt_ligneOL.
END.
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE integrationParam W-Win 
PROCEDURE integrationParam :
DEFINE INPUT PARAMETER ih_repertoire AS HANDLE NO-UNDO.
 
DEFINE VARIABLE ll_ok             AS LOGICAL   NO-UNDO.
DEFINE VARIABLE lc_rep_origine    AS CHARACTER NO-UNDO.
DEFINE VARIABLE lc_liste_dir      AS CHARACTER NO-UNDO.
DEFINE VARIABLE lc_liste_fic      AS CHARACTER NO-UNDO.
DEFINE VARIABLE lc_fichier        AS CHARACTER NO-UNDO.
DEFINE VARIABLE lc_fichComplet    AS CHARACTER NO-UNDO.
 
 
DEFINE VARIABLE li_ii             AS INTEGER   NO-UNDO. 
 
DEFINE VARIABLE li_numLigne       AS INTEGER   NO-UNDO.
DEFINE VARIABLE lc_ligne          AS CHARACTER NO-UNDO.
 
DEFINE VARIABLE lc_format         AS CHARACTER   NO-UNDO.
DEFINE VARIABLE li_Max            AS INTEGER   NO-UNDO.
 
DEFINE VARIABLE lc_repert AS CHARACTER NO-UNDO.
DEFINE VARIABLE lc_ficLog AS CHARACTER NO-UNDO.
 
DEFINE VARIABLE llc_file AS LONGCHAR NO-UNDO.   /* M????? - NID le 15/02/19 - Astr'In */
 
ASSIGN lc_repert = ih_repertoire:SCREEN-VALUE.
 
IF TRIM(lc_repert) = ""  THEN
DO:
    IF NOT gl_modeBatch THEN
        MESSAGE "Vous devez renseigner le r‚pertoire d'int‚gration des paramŠtres." VIEW-AS ALERT-BOX INFO BUTTONS OK.
    ELSE
        RUN erreur ("", 0, "Vous devez renseigner le r‚pertoire d'int‚gration des paramŠtres.").
    APPLY "ENTRY" TO ih_repertoire.
    RETURN ERROR "".
END.
ELSE
DO:       
    RUN Objets\Files\verifie_repertoire.r ( INPUT lc_repert, INPUT TRUE, INPUT " des fichiers d'int‚gration", INPUT " acc‚der aux fichiers", INPUT TRUE, OUTPUT ll_OK ).
    IF NOT ll_ok THEN
        RETURN ERROR "".
END.
 
EMPTY TEMP-TABLE tt_plan.
EMPTY TEMP-TABLE tt_interEber.
 
/* Ajout NID le 20/10/16 - Ajout de l'import des clients/interlocuteurs */
 
/*......................................*/
/*  r‚cup information (date/heure       */
/*......................................*/
FILE-INFO:FILE-NAME = lc_repert.
lc_rep_origine = FILE-INFO:FULL-PATHNAME .
 
/* on r‚cupŠre tous les fichier */
RUN Objets\Files\proc-recup-files.r (  INPUT   lc_rep_origine             ,          /* repertoire d'origine                  */
                                       INPUT   "*"                        ,          /* type de repertoire *.*  a*            */
                                       INPUT   TRUE                       ,          /* fournir les fichiers                  */
                                       INPUT   FALSE                      ,          /* fournir que les noms des fichiers     */
                                       INPUT   "*.*"                      ,          /* type de fichiers *.*  a*.txt          */
                                       INPUT   0                          ,          /* nb niv de srep 0 aucun ou ? = tous    */
                                       INPUT   liste-1:HANDLE IN FRAME f-main,       /* handle d'une selection liste sort 1   */
                                       INPUT   liste-2:HANDLE             ,          /* handle d'une selection liste sort 2   */
                                       INPUT   liste-3:HANDLE             ,          /* handle d'une selection liste sort 3   */
                                       INPUT   liste-4:HANDLE             ,          /* handle d'une selection liste sort 4   */
                                       OUTPUT  lc_liste_dir               ,          /* Liste des dirs                        */
                                       OUTPUT  lc_liste_fic               ,          /* Liste des fichiers                    */
                                       OUTPUT  ll_ok ).                              /* traitement ok ?                       */
 
IF NUM-ENTRIES ( lc_liste_fic ) = 0 OR NUM-ENTRIES ( lc_liste_fic ) =  ? THEN
DO:
    /*
    IF NOT gl_modeBatch THEN
        MESSAGE "Vous n'avez pas de fichier paramŠtres … int‚grer dans le r‚pertoire " + rep_origine VIEW-AS ALERT-BOX.
    */
    RETURN.
END.
 
FRAME F-Criteres:MOVE-TO-BOTTOM().
FRAME F-Criteres:VISIBLE = FALSE.
 
FRAME F-jauge:VISIBLE = TRUE.
FRAME F-jauge:MOVE-TO-TOP().
 
/* jauge.*/
li_Max = NUM-ENTRIES ( lc_liste_fic ) + 1.
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Int‚gration en cours. Veuillez patienter ...").
 
RUN jauge-init IN h_jauge-ocx ( INPUT li_Max ).
 
DO li_ii = 1 TO NUM-ENTRIES(lc_liste_fic) :
 
    ASSIGN
    lc_fichComplet = ENTRY ( li_ii, lc_liste_fic )
    lc_fichier     = ENTRY(NUM-ENTRIES(lc_fichComplet, "\"), lc_fichComplet, "\")
    li_numLigne = 0
    ll_ok = TRUE.
 
    RUN jauge-next IN h_jauge-ocx.
 
    ASSIGN lc_format = IF lc_fichier BEGINS "PGPACP"
                       THEN "Plan Tourn‚e"
                       ELSE IF lc_fichier BEGINS "PGINEB" 
                            THEN "Interlocuteurs Eberhardt"
                            /* Ajout NID le 20/10/16 - Gestion des clients */
                            ELSE IF lc_fichier BEGINS "PGCLIN"
                                 THEN "Clients Astrin"
                            /* Fin Ajout NID le 20/10/16 - Gestion des clients */
                                 ELSE "".
    IF lc_format <> "" THEN
    DO:
        INPUT STREAM st_param FROM VALUE(lc_fichComplet) CONVERT TARGET "ibm850".
        REPEAT:
            IMPORT STREAM st_param UNFORMATTED lc_ligne.
            li_numLigne = li_numLigne + 1.
 
            CASE lc_format:
                /********************/
                /* Plan de tourn‚es */
                /********************/
                WHEN "Plan Tourn‚e" THEN
                DO:
                    IF NUM-ENTRIES(lc_ligne, "|") < 18 THEN
                    DO:
                        RUN erreur(lc_fichier, li_numLigne, "Erreur de format pour le plan de tourn‚e " + lc_ligne).
                        ASSIGN ll_ok = FALSE.
                        NEXT.
                    END.
 
                    CREATE tt_plan.
                    ASSIGN
                    tt_plan.pgpt_codpos     = ENTRY(1, lc_ligne, "|")                   /* X(5) */
                    tt_plan.pgpt_cpidx      = INTEGER(ENTRY(2, lc_ligne, "|"))          /* 99 */
                    tt_plan.pgpt_insee      = ENTRY(3, lc_ligne, "|")                   /* X(5) */
                    tt_plan.pgpt_ville      = ENTRY(4, lc_ligne, "|")                   /* X(35) */
                    tt_plan.pgpt_vil20      = ENTRY(5, lc_ligne, "|")                   /* X(20) */
                    tt_plan.pgpt_pays       = ENTRY(6, lc_ligne, "|")                   /* X(2) */
                    tt_plan.pgpt_tournee[1] = LOGICAL(ENTRY(7, lc_ligne, "|"), "O/N")   /* Lundi */
                    tt_plan.pgpt_tournee[2] = LOGICAL(ENTRY(8, lc_ligne, "|"), "O/N")   /* Mardi */
                    tt_plan.pgpt_tournee[3] = LOGICAL(ENTRY(9, lc_ligne, "|"), "O/N")   /* Mercredi */
                    tt_plan.pgpt_tournee[4] = LOGICAL(ENTRY(10, lc_ligne, "|"), "O/N")  /* Jeudi */
                    tt_plan.pgpt_tournee[5] = LOGICAL(ENTRY(11, lc_ligne, "|"), "O/N")  /* Vendredi */
                    tt_plan.pgpt_tournee[6] = LOGICAL(ENTRY(12, lc_ligne, "|"), "O/N")  /* Samedi */
                    tt_plan.pgpt_montagne   = LOGICAL(ENTRY(13, lc_ligne, "|"), "O/N")  /* X */
                    tt_plan.pgpt_trfpf      = SUBSTRING(ENTRY(14, lc_ligne, "|"), 1, 5) /* X(10) - Jours transfert plateforme */
                    tt_plan.pgpt_dept       = ENTRY(15, lc_ligne, "|")                  /* X(2) */
                    tt_plan.pgpt_plat       = INTEGER(ENTRY(16, lc_ligne, "|"))         /* 99 - Nø Plateforme */
                    tt_plan.pgpt_to2j       = INTEGER(ENTRY(17, lc_ligne, "|"))         /* 9 - Tourn‚e 2 jours */
                    tt_plan.pgpt_par        = ENTRY(18, lc_ligne, "|")                  /* X(10) */
                    tt_plan.tt_fichier      = lc_fichComplet
                    tt_plan.tt_numligne     = li_numLigne NO-ERROR.
                    IF ERROR-STATUS:ERROR THEN
                    DO:
                        ASSIGN ll_ok = FALSE.
                        RUN erreur (lc_fichier, li_numLigne, "Erreur … l'import de la ligne de plan de tourn‚e " + lc_ligne).
                        DELETE tt_plan.
                    END.
                END.
                /****************************/
                /* INTERLOCUTEURS EBERHARDT */
                /****************************/
                WHEN "Interlocuteurs Eberhardt" THEN
                DO:
                    IF NUM-ENTRIES(lc_ligne, "|") < 5 THEN
                    DO:
                        RUN erreur(lc_fichier, li_numLigne, "Erreur de format pour l'interlocuteur Eberhardt " + lc_ligne).
                        ASSIGN ll_ok = FALSE.
                        NEXT.
                    END.
 
                    CREATE tt_interEber.
                    ASSIGN
                    tt_interEber.pgie_prenom = ENTRY(1, lc_ligne, "|")
                    tt_interEber.pgie_nom    = ENTRY(2, lc_ligne, "|")
                    tt_interEber.pgie_mail   = ENTRY(3, lc_ligne, "|")
                    tt_interEber.pgie_idts   = ENTRY(4, lc_ligne, "|")
                    tt_interEber.pgie_tel    = ENTRY(5, lc_ligne, "|")
                    tt_interEber.tt_fichier  = lc_fichComplet
                    tt_interEber.tt_numligne = li_numLigne NO-ERROR.
 
                    IF ERROR-STATUS:ERROR THEN
                    DO:
                        ASSIGN ll_ok = FALSE.
                        RUN erreur (lc_fichier, li_numLigne, "Erreur … l'import de la ligne d'interlocuteur Eberhardt " + lc_ligne).
                        DELETE tt_interEber.
                    END.
                END.
                /* Ajout NID le 20/10/16 - Gestion des clients */
                /************************************/
                /* CLIENTS ET INTERLOCUTEURS ASTRIN */
                /************************************/
                WHEN "Clients Astrin" THEN
                DO:
                    IF TRIM(lc_ligne) = "" THEN NEXT.
 
                    DEFINE VARIABLE li_numInter  AS INTEGER   NO-UNDO.
                    DEFINE VARIABLE lc_numClient AS CHARACTER NO-UNDO.
                    /*DEFINE VARIABLE lc_CodePos   AS CHARACTER NO-UNDO. */
                    DEFINE VARIABLE lc_cpIdx     AS CHARACTER NO-UNDO.
 
                    IF NUM-ENTRIES(lc_ligne, "|") < 14 THEN
                    DO:
                        RUN erreur(lc_fichier, li_numLigne, "Erreur de format de ligne pour le client " + lc_ligne).
                        ASSIGN ll_ok = FALSE.
                        NEXT.
                    END.
 
                    ASSIGN
                    lc_numClient = TRIM(ENTRY(1, lc_ligne, "|"))
                    /*lc_CodePos   = TRIM(ENTRY(2, lc_ligne, "|")) */
                    lc_cpIdx     = TRIM(ENTRY(3, lc_ligne, "|"))
                    li_numInter  = INTEGER(TRIM(ENTRY(14, lc_ligne, "|"))) NO-ERROR.
 
                    IF ERROR-STATUS:ERROR THEN
                    DO:
                        RUN erreur(lc_fichier, li_numLigne, "Erreur de format de donn‚es pour le client " + lc_ligne).
                        ASSIGN ll_ok = FALSE.
                        NEXT.
                    END.
 
                    IF lc_numClient = "" THEN
                        NEXT.
 
                    IF lc_cpIdx = "" THEN
                        NEXT.
 
                    ASSIGN lc_numClient = FILL("0", 6 - LENGTH(lc_numClient)) + lc_numClient.
 
                    IF LENGTH(lc_numClient) <> 6 THEN
                    DO:
                        RUN erreur(lc_fichier, li_numLigne, "Erreur de longueur pour nø de client  " + lc_numClient).
                        ASSIGN ll_ok = FALSE.
                        NEXT.
                    END.
 
                    IF LENGTH(lc_cpIdx) <> 3 THEN
                    /*IF LENGTH(lc_CodePos) <> 5 THEN */
                    DO:
                        RUN erreur(lc_fichier, li_numLigne, "Erreur de longueur pour l'index  " + lc_cpIdx).
                        /*ASSIGN ll_ok = FALSE. */
                        NEXT.
                    END.
                    /* Le CP n'est pas un identifiant utile, car il est calcul‚ a partir du CP_IDX. */
                    /* Pour Boulanger par exemple, on a comme CP 01100, 13600, 49700, 62100 alors que les adresses de livraison ont comme CP (par exemple) 13690 et 62110 */
                    /* On doit travailler avec le CP_IDX, qui constitue les trois premiers chiffres du CP */
 
                    /* On va bien le trouver plusieurs fois dans le fichier, mais on ne le cr‚e qu'une fois */
/*                    FIND FIRST tt_client WHERE tt_client.cli_num_client = lc_numClient AND tt_client.cli_cp = lc_CodePos AND tt_client.tt_fichier = lc_fichComplet NO-ERROR. */
                    FIND FIRST tt_client WHERE tt_client.cli_num_client = lc_numClient AND tt_client.cli_cp = lc_cpIdx AND tt_client.tt_fichier = lc_fichComplet NO-ERROR.
                    IF NOT AVAILABLE tt_client THEN
                    DO:
                        CREATE tt_client.
                        ASSIGN
                        tt_client.cli_num_client = lc_numClient
/*                        tt_client.cli_cp         = tt_ol.ol_cp_livraison */
                        tt_client.cli_cp         = lc_cpIdx
                        tt_client.cli_hayon      = TRIM(ENTRY(8, lc_ligne, "|"))  = "Oui" OR TRIM(ENTRY(8, lc_ligne, "|"))  = "1"
                        tt_client.cli_dderdv     = TRIM(ENTRY(10, lc_ligne, "|")) = "Oui" OR TRIM(ENTRY(10, lc_ligne, "|")) = "1"
                        tt_client.cli_lot        = TRIM(ENTRY(12, lc_ligne, "|")) = "Oui" OR TRIM(ENTRY(12, lc_ligne, "|")) = "1"
                        tt_client.cli_mail       = TRIM(ENTRY(9, lc_ligne, "|"))  = "Oui" OR TRIM(ENTRY(9, lc_ligne, "|"))  = "1"
                        tt_client.cli_confrdv    = TRIM(ENTRY(11, lc_ligne, "|")) = "Oui" OR TRIM(ENTRY(11, lc_ligne, "|")) = "1"
                        tt_client.cli_com        = TRIM(ENTRY(13, lc_ligne, "|"))
                        tt_client.tt_origine     = "CLI"
                        tt_client.tt_fichier     = lc_fichComplet
                        tt_client.tt_numligne    = li_numLigne   NO-ERROR.
                    END.
 
                    FIND FIRST tt_cliint WHERE tt_cliint.cli_num_client = lc_numClient 
/*                                           AND tt_cliint.cli_cp         = lc_CodePos */
                                           AND tt_cliint.cli_cp         = lc_cpIdx
                                           AND tt_cliint.cliint_cle     = li_numInter 
                                           AND tt_cliint.tt_fichier     = lc_fichComplet 
                                           NO-ERROR.
                    IF NOT AVAILABLE tt_cliint THEN
                    DO:
                        CREATE tt_cliint.
                        ASSIGN
                        tt_cliint.cli_num_client     = lc_numClient
                        tt_cliint.cli_cp             = lc_cpIdx
                        /*
                        tt_cliint.cli_cp             = lc_CodePos
                        */
                        tt_cliint.cliint_cle         = li_numInter
                        tt_cliint.cliint_intervenant = TRIM(ENTRY(4, lc_ligne, "|"))
                        tt_cliint.cliint_tel         = TRIM(ENTRY(5, lc_ligne, "|"))
                        tt_cliint.cliint_fax         = TRIM(ENTRY(6, lc_ligne, "|"))
                        tt_cliint.cliint_mail        = TRIM(ENTRY(7, lc_ligne, "|"))
                        tt_cliint.tt_fichier         = lc_fichComplet
                        tt_cliint.tt_numligne        = li_numLigne   NO-ERROR.
                    END.
                END.
            END CASE.
        END.
        INPUT STREAM st_param CLOSE.
    END.
    ELSE
    /* M????? - NID le 15/02/19 - Astr'In */
    DO:
        /* On va v‚rifier si le fichier est complŠtement vide. Si c'est le cas, on l'enregistre pour archivage sans le remonter en erreur */
        COPY-LOB FROM FILE lc_fichier TO llc_file NO-ERROR.
        IF TRIM(REPLACE(REPLACE(REPLACE(REPLACE(llc_file, " ", ""), CHR(10), ""), CHR(9), ""), CHR(13), "")) = "" THEN
            RUN createFichier(lc_fichier, lc_rep_origine).
    END.    
    /* Fin M????? - NID le 15/02/19 - Astr'In */
END.
 
/* fin de traitement */
RUN jauge-init IN  h_jauge-ocx ( INPUT li_Max).
RUN jauge-fin IN h_jauge-ocx.
 
FRAME F-jauge:MOVE-TO-BOTTOM().
FRAME F-jauge:VISIBLE = FALSE.
 
FRAME F-Criteres:VISIBLE = TRUE.
FRAME F-Criteres:MOVE-TO-TOP().
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE lecture_param W-Win 
PROCEDURE lecture_param :
/**/
 
    FIND FIRST GPI_PAR WHERE GPI_PAR.uti_abrege = "IMPORT:EXCLUSIONS ANCIENS OL"
                         AND GPI_PAR.par_param  = "PGI EBERHARDT" NO-LOCK NO-ERROR.
 
    IF AVAILABLE GPI_PAR THEN
    DO:
        ed_LstExclusions = GPI_PAR.par_val.
        DISPLAY ed_LstExclusions WITH FRAME F-Param.
 
        RUN create_TT_OL_EXCLUS.
    END.
 
 
    FIND FIRST GPI_PAR WHERE GPI_PAR.PAR_PARAM = "import_pgi:parametres" AND
                             GPI_PAR.UTI_ABREGE = "" NO-LOCK NO-ERROR.
 
    IF AVAILABLE GPI_PAR THEN
    DO:
        ASSIGN
        fi_rep_integration_articles = GPI_PAR.PAR_VAL2[1]
        fi_rep_integration_ol       = GPI_PAR.PAR_VAL2[2]
        fi_rep_svg                  = GPI_PAR.PAR_VAL2[3]
        fi_email                    = GPI_PAR.PAR_VAL2[4]
        fi_rep_integration_param    = GPI_PAR.PAR_VAL2[5]
        .
 
        DISPLAY fi_rep_integration_articles fi_rep_integration_ol fi_rep_integration_param fi_rep_svg fi_email WITH FRAME F-Criteres.
    END.
 
 
 
    FIND FIRST gpi_par WHERE gpi_par.par_param = "PGI Eberhardt"
                         AND gpi_par.uti_abrege = "gpi" NO-LOCK NO-ERROR.
 
    IF AVAILABLE gpi_par THEN
        ASSIGN
        gc_typePlateforme = gpi_par.par_val2[1]
        gl_outlook        = gpi_par.par_val2[8] = "O".
    ELSE
        ASSIGN 
        gc_typePlateforme = "FCPL"
        gl_outlook        = FALSE.
 
    FIND CURRENT GPI_PAR NO-LOCK.
 
 
    /*---- A decommenter pour gerer l'enregistrement automatique de la taille ---------------------
        &IF DEFINED(UIB_is_Running) EQ 0 
        &THEN
            FIND FIRST GPI_PAR WHERE GPI_PAR.par_param  = Nom-Programme + ":size"
                                 AND GPI_PAR.uti_abrege = Menu_Utilisateur NO-LOCK NO-ERROR.
 
            IF AVAILABLE GPI_PAR THEN 
            DO:
                ASSIGN  w-win:X             = INT ( GPI_PAR.par_val2 [ 1 ] )
                        w-win:Y             = INT ( GPI_PAR.par_val2 [ 2 ] )
                        w-win:WIDTH-PIXELS  = INT ( GPI_PAR.par_val2 [ 3 ] )
                        w-win:HEIGHT-PIXELS = INT ( GPI_PAR.par_val2 [ 4 ] )
                        plein-ecran         =     ( GPI_PAR.par_val2 [ 5 ] = "O" )
                        .
 
                RUN Resize.
 
                IF plein-ecran THEN
                DO:
                    w-win:WINDOW-STATE = 1.
                    RUN Resize.
                END.
            END.
        &ENDIF
    ----*/
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE local-exit W-Win 
PROCEDURE local-exit :
/* -----------------------------------------------------------
  Purpose:  Starts an "exit" by APPLYing CLOSE event, which starts "destroy".
  Parameters:  <none>
  Notes:    If activated, should APPLY CLOSE, *not* dispatch adm-exit.   
-------------------------------------------------------------*/
 
    RUN maj_param.
    IF gl_modeBatch THEN
        QUIT.   /* Pour ‚viter l'ouverture du procedure editor */
    ELSE
    DO:
        APPLY "CLOSE":U TO THIS-PROCEDURE.
        RETURN.
    END.
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE local-initialize W-Win 
PROCEDURE local-initialize :
/*------------------------------------------------------------------------------
  Purpose:     Override standard ADM method
  Notes:       
------------------------------------------------------------------------------*/
DEFINE VARIABLE lc_nomFic AS CHARACTER NO-UNDO.
DEFINE VARIABLE ll_ok     AS LOGICAL   NO-UNDO.
 
DEFINE VARIABLE lc_heureMinImport AS CHARACTER NO-UNDO.
 
RUN LockWindowUpdate ( w-win:HWND, OUTPUT IsLocked).
 
/* Code placed here will execute PRIOR to standard behavior. */
 
/* Dispatch standard ADM method.                             */
RUN dispatch IN THIS-PROCEDURE ( INPUT 'initialize':U ) .
 
/* Code placed here will execute AFTER standard behavior.    */
RUN lecture_param.
 
APPLY "VALUE-CHANGED" TO rs_visu IN FRAME F-Main.
 
DISABLE Btn_Importer WITH FRAME F-Main.
 
RUN afficheErreurs(FALSE).
 
FRAME F-jauge:MOVE-TO-BOTTOM().
FRAME F-jauge:VISIBLE = FALSE.
 
FRAME F-Criteres:VISIBLE = TRUE.
FRAME F-Criteres:MOVE-TO-TOP().
 
ASSIGN
gdt_dateImport      = TODAY
lc_heureMinImport   = STRING(TIME, "HH:MM:SS")
gi_heureImport      = INTEGER(ENTRY(1, lc_heureMinImport, ":"))
gi_minImport        = INTEGER(ENTRY(2, lc_heureMinImport, ":")).
 
IF MENU_utilisateur = "gpi" THEN
    ASSIGN
    Btn_Test:SENSITIVE = TRUE
    Btn_Test:VISIBLE   = TRUE.
ELSE
    ASSIGN
    Btn_Test:SENSITIVE = FALSE
    Btn_Test:VISIBLE   = FALSE.
 
IF gl_modeBatch THEN
DO:
    ASSIGN TG_reinitClients = FALSE.
 
    APPLY "CHOOSE" TO Btn_Executer IN FRAME F-Criteres.
    IF Btn_Importer:SENSITIVE THEN
        APPLY "CHOOSE" TO Btn_Importer IN FRAME F-Main.
 
    IF CAN-FIND(FIRST tt_erreur) THEN
    DO:
        RUN afficheErreurs(TRUE).
        RUN edite_browse(OUTPUT lc_nomFic).
 
        IF fi_email <> "" AND fi_listeMailsValide(fi_email) THEN
            RUN genereMail(lc_nomFic, OUTPUT ll_ok).
        RUN afficheErreurs(FALSE).
        /*
        APPLY "CHOOSE" TO Btn_Excel IN FRAME F-Erreurs.
        */
    END.
    APPLY "CHOOSE" TO Btn_Quitter.
END.
 
/* Code placed here will execute AFTER standard behavior.    */
IF islocked <> 0 THEN
    RUN LockWindowUpdate ( 0, OUTPUT IsLocked ).
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE majObligatoire W-Win 
PROCEDURE majObligatoire :
/*------------------------------------------------------------------------------
  Purpose:     Initialise le statut obligatoire d'un OL en plate forme, si sa date de livraison pr‚visionnelle 
               est inf‚rieure … la date de livraison la plus int‚ressante pour la plateforme et la destination
               (cf ChoixOLObligatoire)
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE VARIABLE ll_obligatoire AS LOGICAL     NO-UNDO.
 
DEFINE VARIABLE li_nbEnreg  AS INTEGER  NO-UNDO.
DEFINE VARIABLE ldt_dtLivTest AS DATE NO-UNDO.
 
/*SELECT COUNT(*) INTO li_nbEnreg FROM gpi_pgol WHERE gpi_pgol.ol_num_pf <> "" AND gpi_pgol.ol_date_retour_Eber IS NULL. */
SELECT COUNT(*) INTO li_nbEnreg FROM gpi_pgol WHERE gpi_pgol.ol_num_pf <> "" AND gpi_pgol.ol_date_retour_Eber IS NULL AND NOT gpi_pgol.ol_a_supprimer AND NOT GPI_PGOL.ol_ancien_ol.
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Mise … jour du status obligatoire. Veuillez patienter ... (<p>%)").
 
RUN jauge-init IN h_jauge-ocx ( INPUT li_nbEnreg ).
 
/*FOR EACH gpi_pgol WHERE gpi_pgol.ol_mode_liv = "PF" AND gpi_pgol.ol_date_retour_Eber <> ? NO-LOCK,*/
/*FOR EACH gpi_pgol WHERE gpi_pgol.ol_num_pf <> "" AND gpi_pgol.ol_date_retour_Eber = ? NO-LOCK, */
FOR EACH gpi_pgol WHERE gpi_pgol.ol_num_pf <> "" AND gpi_pgol.ol_date_retour_Eber = ? AND NOT gpi_pgol.ol_a_supprimer AND NOT GPI_PGOL.ol_ancien_ol NO-LOCK,
    FIRST tt_plateforme WHERE tt_plateforme.tt_code = gpi_pgol.ol_num_pf:
 
    RUN jauge-next IN h_jauge-ocx.
 
    ASSIGN ll_obligatoire = FALSE.  
 
    FIND FIRST gpi_pgpt WHERE gpi_pgpt.pgpt_codpos = gpi_pgol.ol_cp_livraison AND gpi_pgpt.pgpt_vil20 = gpi_pgol.ol_ville_livraison NO-LOCK NO-ERROR.
    /*
    IF AVAILABLE gpi_pgpt AND gpi_pgpt.pgpt_tournee[WEEKDAY(tt_plateforme.tt_dtCharg) - 1] = ? AND gpi_pgpt.pgpt_to2j < 2 THEN
    DO:
        IF gpi_pgol.ol_date_liv < tt_plateforme.tt_dtLivLong THEN /* Livraison 72h sans tourn‚e 2 jours, on compare avec la date + 2j */
            ASSIGN ll_obligatoire = TRUE.
    END.
    ELSE
    DO:
        IF gpi_pgol.ol_date_liv < tt_plateforme.tt_dtLivCourt THEN /* Pas de livraison 72h ou tourn‚e 2 jours, on compare avec la date sans + 2 j */
            ASSIGN ll_obligatoire = TRUE.
 
    END.
    */
 
    IF NOT AVAILABLE GPI_PGPT OR (GPI_PGPT.pgpt_to2j < 2 AND GPI_PGPT.pgpt_tournee[WEEKDAY(tt_plateforme.tt_dtCharg) - 1] <> ?) THEN
        ASSIGN ldt_dtLivTest = tt_plateforme.tt_dtLivCourt.
    ELSE
        IF GPI_PGPT.pgpt_to2j = 2 THEN
            ASSIGN ldt_dtLivTest = tt_plateforme.tt_dtLiv2Jours.
        ELSE
            ASSIGN ldt_dtLivTest = tt_plateforme.tt_dtLivLong.
 
    /* M????? - NID le 04/07/17 - Astr'In - Pb de coch‚ obligatoire */
    IF (GPI_PGOL.ol_date_liv_souhait <> ? AND GPI_PGOL.ol_date_liv_souhait < ldt_dtLivTest) OR GPI_PGOL.ol_date_liv < ldt_dtLivTest THEN
    /*
    IF GPI_PGOL.ol_date_liv < ldt_dtLivTest THEN
    */
    /* Fin M????? - NID le 04/07/17 - Astr'In */
        ASSIGN ll_obligatoire = TRUE.
 
    /* Gestion du champ Obligatoire */
    IF ll_obligatoire <> gpi_pgol.ol_obligatoire THEN
    DO TRANSACTION:
        FIND FIRST bf_pgol WHERE bf_pgol.ol_num_ol = gpi_pgol.ol_num_ol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        IF AVAILABLE bf_pgol THEN
        DO:
            ASSIGN bf_pgol.ol_obligatoire = ll_obligatoire.
            RELEASE bf_pgol.
        END.
    END.
END.
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE maj_param W-Win 
PROCEDURE maj_param :
/*--*/
    FIND FIRST GPI_PAR WHERE GPI_PAR.PAR_PARAM = "import_pgi:parametres" AND
                             GPI_PAR.UTI_ABREGE = "" EXCLUSIVE-LOCK NO-ERROR.
    IF NOT AVAILABLE GPI_PAR AND NOT LOCKED GPI_PAR THEN
    DO:
        CREATE GPI_PAR.
        ASSIGN
        GPI_PAR.PAR_PARAM = "import_pgi:parametres"
        GPI_PAR.UTI_ABREGE = "".
    END.
 
    IF AVAILABLE GPI_PAR THEN
    DO:
        ASSIGN
        GPI_PAR.PAR_VAL2[1] = fi_rep_integration_articles:SCREEN-VALUE IN FRAME F-Criteres
        GPI_PAR.PAR_VAL2[2] = fi_rep_integration_ol:SCREEN-VALUE
        GPI_PAR.PAR_VAL2[3] = fi_rep_svg:SCREEN-VALUE
        GPI_PAR.PAR_VAL2[4] = fi_email:SCREEN-VALUE
        GPI_PAR.PAR_VAL2[5] = fi_rep_integration_param:SCREEN-VALUE
        .
    END.
 
    /*---- A decommenter pour gerer l'enregistrement automatique de la taille ---------------------
    &IF DEFINED(UIB_is_Running) EQ 0 
    &THEN
        FIND FIRST GPI_PAR WHERE GPI_PAR.par_param  = Nom-Programme + ":size"
                             AND GPI_PAR.uti_abrege = Menu_Utilisateur NO-ERROR.
 
        IF NOT LOCKED GPI_PAR THEN 
        DO:
            IF NOT AVAILABLE GPI_PAR THEN 
            DO:
                CREATE  GPI_PAR.
                ASSIGN  GPI_PAR.par_param       = Nom-Programme + ":size"
                        GPI_PAR.uti_abrege      = Menu_Utilisateur  
                        GPI_PAR.par_val2 [ 1 ]  = STRING ( 0                       ) 
                        GPI_PAR.par_val2 [ 2 ]  = STRING ( 0                       ) 
                        GPI_PAR.par_val2 [ 3 ]  = STRING ( w-win:MIN-WIDTH-PIXELS  ) 
                        GPI_PAR.par_val2 [ 4 ]  = STRING ( w-win:MIN-HEIGHT-PIXELS ) 
                        GPI_PAR.par_val2 [ 5 ]  = "N"
                        .
            END.
 
            IF NOT plein-ecran THEN
                ASSIGN  GPI_PAR.par_val2 [ 1 ] = STRING ( w-win:X             )
                        GPI_PAR.par_val2 [ 2 ] = STRING ( w-win:Y             )
                        GPI_PAR.par_val2 [ 3 ] = STRING ( w-win:WIDTH-PIXELS  )
                        GPI_PAR.par_val2 [ 4 ] = STRING ( w-win:HEIGHT-PIXELS )
                        .
 
            GPI_PAR.par_val2 [ 5 ] = STRING ( plein-ecran, "O/N" ). 
            RELEASE GPI_PAR.
        END.
    &ENDIF
    ----*/
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE resize W-Win 
PROCEDURE resize :
/*--*/
 
    DEF VAR liste_hdl AS CHAR NO-UNDO.
 
    /*
        Syntaxe de la liste:
        -------------------
 
            handle,par,par1  chr(9)  handle,par,par1  chr(9) ...
 
            handle = string du handle de l'objet
            par    = constitu& des lettres R C H W
            par1   = L si traitement automatique du label associ‚ (SIDE-LABEL-HANDLE)
 
            R = row     C = Col     W = Width     H = height          
 
            Chaque lettre peut ˆtre suivi d'un diviseur (1 chiffre de 1 a 9)
            Ce diviseur permet de d‚caller des objets d'une mˆme ligne ou colonne
 
              Exemple de paramÀtre: RC2 D‚callage vertical (Row) du diffŠrentiel
            --------------------        D²‚allage horizontal (Col) du diffŠrentiel / 2
                                        Multiplicateur * exemple C4*2  = 1/4 du diffŠrentiel multipli‚ par 2 soit un demi
                                        Pour appliquer par exemple 1/2 puis 1/4 puis 1/4
    */
 
 
    liste_hdl =    STRING ( image-banniere       :HANDLE IN FRAME frame-banniere )   + "," + "H"
        + CHR(9) + STRING ( FRAME frame-banniere :HANDLE )                           + "," + "H"
        + CHR(9) + STRING ( Btn_Importer         :HANDLE IN FRAME F-Main )           + "," + "R"
 
        + CHR(9) + STRING ( Btn_Quitter          :HANDLE IN FRAME F-Main )           + "," + "RC"
 
        /* Browses */
        + CHR(9) + STRING ( BROWSE Br_Articles   :HANDLE)                            + "," + "WH"
        + CHR(9) + STRING ( BROWSE Br_Corresp    :HANDLE)                            + "," + "WH"
        + CHR(9) + STRING ( BROWSE Br_InterEber  :HANDLE)                            + "," + "WH"
        + CHR(9) + STRING ( BROWSE Br_LigneOL    :HANDLE)                            + "," + "WH2R2"
        + CHR(9) + STRING ( BROWSE Br_OL         :HANDLE)                            + "," + "WH2"
        + CHR(9) + STRING ( BROWSE Br_PlanTournee:HANDLE)                            + "," + "WH"
 
        + CHR(9) + STRING (ED_Commentaire        :HANDLE IN FRAME F-Main )           + "," + "WR2"
 
        + CHR(9) + STRING ( FRAME F-Erreurs      :HANDLE )                           + "," + "Z"
        .
 
    RUN Resize-Auto ( INPUT        w-win:HANDLE,
                      INPUT        liste_hdl,
                      INPUT        "",
                      INPUT-OUTPUT anch,
                      INPUT-OUTPUT ancw,
                      OUTPUT       difh, 
                      OUTPUT       difw). 
 
    RUN maj_param.
 
END PROCEDURE.
 
 
PROCEDURE proc-expand:
 
    DEF INPUT PARAMETER difh AS DEC NO-UNDO.
    DEF INPUT PARAMETER difw AS DEC NO-UNDO.
 
    /*---- Ajouter les objets non-progress ( OCX, Jauge, Date, Adresse... ---------------------
        /*     DEF VAR h AS DEC NO-UNDO.                                 */
        /*     DEF VAR w AS DEC NO-UNDO.                                 */
        /*                                                               */
        /*     RUN get-size IN h_jauge-ocx ( OUTPUT h, OUTPUT w).        */
        /*     RUN set-size IN h_jauge-ocx ( INPUT  h, input  w + difw). */
    ----*/
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE sauveFichier W-Win 
PROCEDURE sauveFichier :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
 
DEFINE INPUT PARAMETER ic_fichier       AS CHARACTER NO-UNDO.
DEFINE INPUT PARAMETER ic_repOrigine    AS CHARACTER NO-UNDO.
DEFINE INPUT PARAMETER ic_repCible      AS CHARACTER NO-UNDO.
 
DEFINE VARIABLE lc_lib AS CHARACTER NO-UNDO.
 
/* Dans le cas ou le dossier de sauvegarde est le mˆme dossier que celui d'origine, il ne faut pas supprimer ni faire la sauvegarde */
IF ic_fichier <> "" AND ic_repCible <> ic_repOrigine AND ic_repCible <> ic_repOrigine + "\*"  THEN
DO:
    /* Copie du fichier */
    OS-COPY VALUE ( ic_fichier ) VALUE ( ic_repCible ).
    IF OS-ERROR <> 0 THEN
    DO:
        RUN get_os_error ( OS-ERROR, OUTPUT lc_lib ).
        RUN erreur(ic_fichier, 0, "Erreur de copie du fichier " + ic_fichier + ". Erreur " + STRING ( OS-ERROR ) + " -> " + lc_lib).
        /*
        /*IF lTraitementBatch = FALSE THEN /* 22/01/15 ne pas executer le msg en mode auto/batch */*/
            MESSAGE "Erreur de copie du fichier" SKIP
                    ic_fichier                   SKIP(2)
                    "Erreur " + STRING ( OS-ERROR ) + " -> " + lc_lib
                    VIEW-AS ALERT-BOX.
        */
    END.
    ELSE
    DO:
        /* Suppression du fichier */
        OS-DELETE VALUE (ic_fichier ).
        IF OS-ERROR <> 0 THEN
        DO:
            RUN get_os_error ( OS-ERROR, OUTPUT lc_lib ).
            RUN erreur(ic_fichier, 0, "Erreur de suppression du fichier " + ic_fichier + ". Erreur " + STRING ( OS-ERROR ) + " -> " + lc_lib).
            /*
            /*IF lTraitementBatch = FALSE THEN /* 22/01/15 ne pas executer le msg en mode auto/batch */*/
                MESSAGE "Erreur de suppression du fichier"  SKIP
                        ic_fichier                          SKIP(2)
                        "Erreur " + STRING ( OS-ERROR ) + " -> " + lc_lib
                        VIEW-AS ALERT-BOX.
            */
        END.
    END.
END.
 
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE sauve_param W-Win 
PROCEDURE sauve_param :
/*--*/
 
    FIND FIRST GPI_PAR WHERE GPI_PAR.uti_abrege = "IMPORT:EXCLUSIONS ANCIENS OL"
                         AND GPI_PAR.par_param  = "PGI EBERHARDT" EXCLUSIVE-LOCK NO-ERROR.
 
    IF NOT AVAILABLE GPI_PAR THEN 
        IF LOCKED GPI_PAR THEN
        DO:
            MESSAGE "Erreur !" SKIP
                    " " SKIP
                    "Les paramŠtres sont verrouill‚s par un autre utilisateur !" SKIP
                    "Aucune mise … jour des paramŠtres effectu‚e" VIEW-AS ALERT-BOX.
 
            RETURN "erreur".
        END.
        ELSE
        DO:
            CREATE  GPI_PAR.
            ASSIGN  GPI_PAR.uti_abrege  = "IMPORT:EXCLUSIONS ANCIENS OL" 
                    GPI_PAR.par_param   = "PGI EBERHARDT" 
                    .
        END.
 
    FIND CURRENT GPI_PAR EXCLUSIVE-LOCK.
 
    GPI_PAR.par_val = ed_LstExclusions.
 
    FIND CURRENT GPI_PAR NO-LOCK.
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE send-records W-Win  _ADM-SEND-RECORDS
PROCEDURE send-records :
/*------------------------------------------------------------------------------
  Purpose:     Send record ROWID's for all tables used by
               this file.
  Parameters:  see template/snd-head.i
------------------------------------------------------------------------------*/
 
  /* Define variables needed by this internal procedure.               */
  {src/adm/template/snd-head.i}
 
  /* For each requested table, put it's ROWID in the output list.      */
  {src/adm/template/snd-list.i "tt_plan"}
  {src/adm/template/snd-list.i "tt_ol"}
  {src/adm/template/snd-list.i "tt_ligneOL"}
  {src/adm/template/snd-list.i "tt_interEber"}
  {src/adm/template/snd-list.i "tt_composant"}
  {src/adm/template/snd-list.i "tt_client"}
  {src/adm/template/snd-list.i "tt_article"}
  {src/adm/template/snd-list.i "tt_erreur"}
 
  /* Deal with any unexpected table requests before closing.           */
  {src/adm/template/snd-end.i}
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE state-changed W-Win 
PROCEDURE state-changed :
/* -----------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
-------------------------------------------------------------*/
  DEFINE INPUT PARAMETER p-issuer-hdl AS HANDLE NO-UNDO.
  DEFINE INPUT PARAMETER p-state AS CHARACTER NO-UNDO.
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE supOL W-Win 
PROCEDURE supOL :
/*------------------------------------------------------------------------------
  Purpose:     V‚rifie si des OL en cours (non encore renvoy‚s … Eberhrdt)
               ont ‚t‚ supprim‚s du flux de donn‚es
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
/*
 Si l'enregistrement n'a pas ‚t‚ renvoy‚ … Eberhardt, et qu'il n'est pas dans le fichier d'int‚gration du jour
 c'est qu'il a ‚t‚ annul‚ par Eberhardt. On doit le supprimer physiquement (s'il n'est affect‚ … aucune affaire)
 ou logiquement (si une affaire est affect‚e)
*/
FOR EACH gpi_pgol WHERE gpi_pgol.ol_date_retour_eber = ?
                    AND (gpi_pgol.ol_date_dernier_envoi_eberhardt = ? OR 
                         gpi_pgol.ol_date_dernier_envoi_eberhardt < TODAY OR
                         (GPI_PGOL.ol_date_dernier_envoi_eberhardt = TODAY AND
                          (GPI_PGOL.ol_heure_dernier_envoi_eberhardt < gi_heureImport OR 
                           (GPI_PGOL.ol_heure_dernier_envoi_eberhardt = gi_heureImport AND 
                            GPI_PGOL.ol_minute_dernier_envoi_eberhard < gi_minImport)))) EXCLUSIVE-LOCK:
 
    /* Contr“le l'existence de la saisie */
    FIND FIRST gpi_act WHERE gpi_act.act_id = gpi_pgol.act_id NO-LOCK NO-ERROR.
    IF NOT AVAILABLE gpi_act OR gpi_act.act_type_document = "SS" THEN
    DO:
        /* Un magnifique bug supprimait les entˆtes sans supprimer les lignes. On a donc plein de lignes parasites … supprimer */
        FIND FIRST GPI_PGDETAILOL WHERE GPI_PGDETAILOL.ol_num_ol = gpi_pgol.ol_num_ol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        DO WHILE AVAILABLE GPI_PGDETAILOL:
            DELETE GPI_PGDETAILOL NO-ERROR.
            IF ERROR-STATUS:ERROR THEN
            DO:
                RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "Erreur lors de la mise … jours de l'OL " + bf_ol.OL_NUM_OL + ". Erreur " + ERROR-STATUS:GET-MESSAGE(ERROR-STATUS:NUM-MESSAGES)).
                UNDO, NEXT.
            END.
            FIND NEXT GPI_PGDETAILOL WHERE GPI_PGDETAILOL.ol_num_ol = bf_ol.ol_num_ol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        END.
        IF LOCKED GPI_PGDETAILOL THEN
        DO:
            RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "OL " + bf_ol.OL_NUM_OL + " en cours de modification par un autre utilisateur.").
            UNDO, NEXT.    /* D‚fait le bloc transaction en cours. */
        END.
 
        /* Suppression physique */
        DELETE gpi_pgol.
    END.
    ELSE
    DO:
        /* Suppression physique */
        ASSIGN gpi_pgol.ol_a_supprimer = TRUE.
        IF gpi_pgol.ol_mode_liv = "PF" THEN
            RUN erreur ("", 0, "L'OL " + gpi_pgol.ol_num_ol + " pour la plateforme " + gpi_act.act_nom_destinataire + " n'est plus … traiter, mais est affect‚ … l'affaire " + gpi_act.act_affaire + ".").
        ELSE
            RUN erreur ("", 0, "L'OL " + gpi_pgol.ol_num_ol + " pour le client " + gpi_pgol.ol_nom_livraison + " " + gpi_pgol.ol_cp_livraison + gpi_pgol.ol_ville_livraison + " (" + gpi_pgol.ol_num_client + ") n'est plus … traiter, mais est affect‚ … l'affaire " + gpi_act.act_affaire + ".").
    END.
END.
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE traitement W-Win 
PROCEDURE traitement :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
/* NID le 25/10/16 - Refonte du traitement des OL. */
/* On doit appliquer les modifications dans tout les cas, mˆme si l'article est en affaire ou renvoy‚ … Eberhardt */
 
DEFINE VARIABLE lc_rep_articles  AS CHARACTER NO-UNDO.
DEFINE VARIABLE lc_rep_ol        AS CHARACTER NO-UNDO.
DEFINE VARIABLE lc_rep_param     AS CHARACTER NO-UNDO.
 
DEFINE VARIABLE lc_rep_svg       AS CHARACTER NO-UNDO.
DEFINE VARIABLE li_jauge         AS INTEGER   NO-UNDO.
 
DEFINE VARIABLE ll_ok            AS LOGICAL   NO-UNDO.
/*DEFINE VARIABLE ll_enAffaire     AS LOGICAL   NO-UNDO.  /* Ajout NDE le 06/10/16 */ */
DEFINE VARIABLE li_heureTrt      AS INTEGER   NO-UNDO.
DEFINE VARIABLE li_minTrt        AS INTEGER   NO-UNDO.
DEFINE VARIABLE lc_heureMinTrt   AS CHARACTER NO-UNDO.
 
EMPTY TEMP-TABLE tt_fichiers.
 
/* Initialisation du r‚pertoire d'origine des articles */
ASSIGN
gl_importOL = FALSE
FILE-INFORMATION:FILE-NAME = fi_rep_integration_articles:SCREEN-VALUE IN FRAME F-Criteres
lc_rep_articles = FILE-INFORMATION:FULL-PATHNAME.
 
/* Initialisation du r‚pertoire d'origine des ol */
ASSIGN
FILE-INFORMATION:FILE-NAME = fi_rep_integration_ol:SCREEN-VALUE IN FRAME F-Criteres
lc_rep_ol = FILE-INFORMATION:FULL-PATHNAME.
 
/* Initialisation du r‚pertoire d'origine des paramŠtres */
ASSIGN
FILE-INFORMATION:FILE-NAME = fi_rep_integration_param:SCREEN-VALUE IN FRAME F-Criteres
lc_rep_param = FILE-INFORMATION:FULL-PATHNAME.
 
/* Initialisation du r‚pertoire de sauvegarde */
RUN genereRepSauvegarde(OUTPUT lc_rep_svg).
 
/* On va parcourir les diff‚rentes tables temporaire pour cr‚er les enregistrements en base. */
/* Si une erreur se produit, on cr‚e un enregistrement dans la table des erreurs.            */
/* Le fichier n'est d‚plac‚ que si aucune erreur ne s'est produite lors de son traitement.   */
 
/****************/
/*** ARTICLES ***/
/****************/
RUN trt_articles (INPUT lc_rep_articles).
 
/******************/
/*** COMPOSANTS ***/
/******************/
RUN trt_composants (INPUT lc_rep_articles).
 
/***************************/
/*** ORDRES DE LIVRAISON ***/
/***************************/
ASSIGN li_jauge = 0.
FOR EACH bf_ol:
    ASSIGN li_jauge = li_jauge + 1.
END.
 
ASSIGN
lc_heureMinTrt   = STRING(TIME, "HH:MM:SS")
li_heureTrt      = INTEGER(ENTRY(1, lc_heureMinTrt, ":"))
li_minTrt        = INTEGER(ENTRY(2, lc_heureMinTrt, ":")).
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Traitement des O.L. Veuillez patienter ... (<p>%)").
 
RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
FOR EACH bf_ol BY bf_ol.tt_fichier BY bf_ol.tt_numligne:
 
    RUN jauge-next IN h_jauge-ocx.
 
    FIND FIRST tt_fichiers WHERE tt_fichiers.tt_fichier = bf_ol.tt_fichier NO-ERROR.
    IF NOT AVAILABLE tt_fichiers THEN
        RUN createFichier(bf_ol.tt_fichier, lc_rep_ol).

    /* Le 28/07/16 - Pour les OL, on est en annule et remplace, sauf si l'OL est d‚j… affect‚ … une affaire */
    FIND FIRST GPI_PGOL WHERE GPI_PGOL.OL_NUM_OL = bf_ol.OL_NUM_OL NO-LOCK NO-ERROR.
    IF AVAILABLE GPI_PGOL THEN
    DO:
        DO TRANSACTION:
            /*ASSIGN ll_enAffaire = FALSE. */
            FIND CURRENT gpi_pgol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
                /*
                IF AVAILABLE GPI_PGOL THEN
                    MESSAGE "OL bloqu‚ avec succ‚s." VIEW-AS ALERT-BOX.
                ELSE
                    MESSAGE "Impossible de bloquer l'OL." VIEW-AS ALERT-BOX.
                */

            IF NOT AVAILABLE gpi_pgol THEN
            DO:
                IF LOCKED gpi_pgol THEN
                DO:
                    RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "OL " + bf_ol.OL_NUM_OL + " en cours de modification par un autre utilisateur.").
                    UNDO, LEAVE.    /* Defait le bloc transaction en cours. */
                END.
                ELSE
                DO:
                    CREATE GPI_PGOL.
                    BUFFER-COPY bf_ol EXCEPT bf_ol.ACT_ID bf_ol.OL_DATE-HEURE_CREATION
                             /* M17110 - NID le 25/04/18 - Astr'In - S‚parer la date du fichier de la date r‚elle de cr‚ation. La date d'int‚gration du fichier est d‚j… stock‚e en base plus haut */
                             /*
                             TO GPI_PGOL ASSIGN GPI_PGOL.OL_DATE-HEURE_CREATION = IF bf_ol.tt_dt_export = ? THEN NOW ELSE bf_ol.tt_dt_export
                             */
                             TO GPI_PGOL ASSIGN GPI_PGOL.OL_DATE-HEURE_CREATION = NOW
                             /* Fin M17110 - NID le 25/04/18 - Astr'In */ 
                                                GPI_PGOL.ol_date_dernier_envoi_eberhardt = TODAY
                                                GPI_PGOL.ol_heure_dernier_envoi_eberhardt = li_heureTrt
                                                GPI_PGOL.ol_minute_dernier_envoi_eberhard = li_minTrt.
                    RELEASE GPI_PGOL.
                END. /* ELSE (IF LOCKED gpi_pgol) */
            END. /* IF NOT AVAILABLE gpi_pgol */
            ELSE
            DO:
                /* NID le 05/05/21 - Si l'OL a + de deux ans, c'est qu'on r‚cupŠre un vieil OL d‚ja utilis‚. On ‚crase tout */
                IF ADD-INTERVAL(GPI_PGOL.ol_date-heure_creation, 3, "years") <= NOW THEN
                DO:
                    BUFFER-COPY bf_ol EXCEPT bf_ol.ACT_ID bf_ol.OL_DATE-HEURE_CREATION
                             TO GPI_PGOL ASSIGN GPI_PGOL.OL_DATE-HEURE_CREATION           = NOW
                                                GPI_PGOL.ol_date_dernier_envoi_eberhardt  = TODAY
                                                GPI_PGOL.ol_heure_dernier_envoi_eberhardt = li_heureTrt
                                                GPI_PGOL.ol_minute_dernier_envoi_eberhard = li_minTrt.
                END.
                ELSE
                DO:
                /* Fin NID le 05/05/21 */
                    /* Ajout NID le 25/10/16 - On modifie syst‚matiquement, mais uniquement les champs qui on pu ˆtre mis … jour */
                    /*
                    /* On contr“le s'il faut faire la modification. Si l'OL n'est pas encore en pr‚paration, on peut le mettre … jour */
                    IF gpi_pgol.ol_date_chg = ? AND GPI_PGOL.ol_date_RDV = ? AND gpi_pgol.ol_num_affaire = "" AND GPI_PGOL.ol_date_retour_Eber = ? THEN
                        BUFFER-COPY bf_ol EXCEPT bf_ol.ol_num_ol bf_ol.ACT_ID bf_ol.OL_DATE-HEURE_CREATION TO GPI_PGOL.
                    ELSE
                        ASSIGN ll_enAffaire = TRUE.
                    */
                    ASSIGN
                    /* Ajout NID le 25/10/16 */
                    GPI_PGOL.ol_gestionnaire_commandes          = bf_ol.ol_gestionnaire_commandes
                    GPI_PGOL.ol_num_client                      = bf_ol.ol_num_client
                    GPI_PGOL.ol_ref_client                      = bf_ol.ol_ref_client
                    GPI_PGOL.ol_mode_transport                  = bf_ol.ol_mode_transport
                    GPI_PGOL.ol_titre_commande                  = bf_ol.ol_titre_commande
                    GPI_PGOL.ol_nom_commande                    = bf_ol.ol_nom_commande
                    GPI_PGOL.ol_adresse1_commande               = bf_ol.ol_adresse1_commande
                    GPI_PGOL.ol_adresse2_commande               = bf_ol.ol_adresse2_commande
                    GPI_PGOL.ol_cp_commande                     = bf_ol.ol_cp_commande
                    GPI_PGOL.ol_ville_commande                  = bf_ol.ol_ville_commande
                    GPI_PGOL.ol_titre_livraison                 = bf_ol.ol_titre_livraison
                    GPI_PGOL.ol_nom_livraison                   = bf_ol.ol_nom_livraison
                    GPI_PGOL.ol_adresse1_livraison              = bf_ol.ol_adresse1_livraison
                    GPI_PGOL.ol_adresse2_livraison              = bf_ol.ol_adresse2_livraison
                    GPI_PGOL.ol_cp_livraison                    = bf_ol.ol_cp_livraison
                    GPI_PGOL.ol_ville_livraison                 = bf_ol.ol_ville_livraison
                    GPI_PGOL.ol_rdv_obligatoire                 = bf_ol.ol_rdv_obligatoire
                    GPI_PGOL.ol_hayon_obligatoire               = bf_ol.ol_hayon_obligatoire
                    GPI_PGOL.ol_lot_obligatoire                 = bf_ol.ol_lot_obligatoire
                    GPI_PGOL.ol_crossdock                       = bf_ol.ol_crossdock
                    GPI_PGOL.ol_date_liv_souhait                = bf_ol.ol_date_liv_souhait
                    GPI_PGOL.ol_cp_idx                          = bf_ol.ol_cp_idx
                    GPI_PGOL.ol_dep_livraison                   = bf_ol.ol_dep_livraison
                    /* Fin Ajout NID le 25/10/16 */
                    /* NID le 11/07/17 - pb OL a supprimer par erreur */
                    GPI_PGOL.ol_a_supprimer                     = FALSE /* Si l'OL est dans le fichier, c'est que visiblement il n'‚tait pas a supprimer */
                    GPI_PGOL.ol_date_dernier_envoi_eberhardt    = TODAY
                    GPI_PGOL.ol_heure_dernier_envoi_eberhardt   = li_heureTrt
                    GPI_PGOL.ol_minute_dernier_envoi_eberhard   = li_minTrt.
                END.
                FIND CURRENT GPI_PGOL NO-LOCK NO-ERROR.
/*                RELEASE GPI_PGOL.*/
            END.    /* ELSE (IF NOT AVAILABLE gpi_pgol) */
 
            /*
            /* On ne le fait que si l'entˆte a ‚t‚ modifi‚e */
            IF NOT ll_enAffaire THEN
            DO:
            */
                /* Suppression des d‚tails OL li‚s pour pouvoir les recr‚er. On d‚fait la transaction si une erreur se produit */
                FIND FIRST GPI_PGDETAILOL WHERE GPI_PGDETAILOL.ol_num_ol = bf_ol.ol_num_ol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
                DO WHILE AVAILABLE GPI_PGDETAILOL:
                    DELETE GPI_PGDETAILOL NO-ERROR.
                    IF ERROR-STATUS:ERROR THEN
                    DO:
                        RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "Erreur lors de la mise … jours de l'OL " + bf_ol.OL_NUM_OL + ". Erreur " + ERROR-STATUS:GET-MESSAGE(ERROR-STATUS:NUM-MESSAGES)).
                        UNDO, LEAVE.
                    END.
                    FIND NEXT GPI_PGDETAILOL WHERE GPI_PGDETAILOL.ol_num_ol = bf_ol.ol_num_ol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
                END.
                IF LOCKED GPI_PGDETAILOL THEN
                DO:
                    RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "OL " + bf_ol.OL_NUM_OL + " en cours de modification par un autre utilisateur.").
                    UNDO, LEAVE.    /* Defait le bloc transaction en cours. */
                END.
 
                /* Recr‚e les nouvelles lignes de d‚tail */
                RUN trt_detailOl (INPUT  bf_ol.ol_num_ol,
                                  INPUT  bf_ol.tt_fichier,
                                  INPUT  lc_rep_ol,
                                  OUTPUT ll_ok).
                IF NOT ll_ok THEN
                    UNDO, LEAVE.    /* D‚fait le bloc transaction en cours. */
            /* END. */
            DELETE bf_ol.
        END. /* IF gpi_pgol.ol_date_chargement = ? AND gpi_pgol.ol_num_affaire = "" - TRANSACTION */
        FIND CURRENT GPI_PGOL NO-LOCK NO-ERROR.
 
        /*
            /* Sinon, on supprime simplement bf_ol, et les bf_ligneOL li‚s */
            ELSE
            DO: 
                FOR EACH bf_ligneOL WHERE bf_ligneOL.ol_num_ol = bf_ol.ol_num_ol AND bf_ligneOL.tt_fichier = bf_ol.tt_fichier:
                    DELETE bf_ligneOL.
                END.
                DELETE bf_ol.
            END.
            /*RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "OL " + bf_ol.OL_NUM_OL + " existe d‚j….").*/
        */
    END.
    ELSE
    DO TRANSACTION:
        CREATE GPI_PGOL.
        BUFFER-COPY bf_ol EXCEPT bf_ol.ACT_ID bf_ol.OL_DATE-HEURE_CREATION
                 /* M17110 - NID le 25/04/18 - Astr'In - S‚parer la date du fichier de la date r‚elle de cr‚ation. La date d'int‚gration du fichier est d‚j… stock‚e en base plus haut */
                 /*
                 TO GPI_PGOL ASSIGN GPI_PGOL.OL_DATE-HEURE_CREATION = IF bf_ol.tt_dt_export = ? THEN NOW ELSE bf_ol.tt_dt_export
                 */
                 TO GPI_PGOL ASSIGN GPI_PGOL.OL_DATE-HEURE_CREATION = NOW
                 /* Fin M17110 - NID le 25/04/18 - Astr'In */ 
                                    GPI_PGOL.ol_date_dernier_envoi_eberhardt  = TODAY
                                    GPI_PGOL.ol_heure_dernier_envoi_eberhardt = li_heureTrt
                                    GPI_PGOL.ol_minute_dernier_envoi_eberhard = li_minTrt.
        RELEASE GPI_PGOL.
 
        /* Suppression des d‚tails OL li‚s pour pouvoir les recr‚er. On d‚fait la transaction si une erreur se produit */
        /* Un magnifique bug supprimait les entˆtes sans supprimer les lignes. On a donc plein de lignes parasites … supprimer */
        FIND FIRST GPI_PGDETAILOL WHERE GPI_PGDETAILOL.ol_num_ol = bf_ol.ol_num_ol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        DO WHILE AVAILABLE GPI_PGDETAILOL:
            DELETE GPI_PGDETAILOL NO-ERROR.
            IF ERROR-STATUS:ERROR THEN
            DO:
                RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "Erreur lors de la mise … jours de l'OL " + bf_ol.OL_NUM_OL + ". Erreur " + ERROR-STATUS:GET-MESSAGE(ERROR-STATUS:NUM-MESSAGES)).
                UNDO, LEAVE.
            END.
            FIND NEXT GPI_PGDETAILOL WHERE GPI_PGDETAILOL.ol_num_ol = bf_ol.ol_num_ol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        END.
        IF LOCKED GPI_PGDETAILOL THEN
        DO:
            RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "OL " + bf_ol.OL_NUM_OL + " en cours de modification par un autre utilisateur.").
            UNDO, LEAVE.    /* Defait le bloc transaction en cours. */
        END.
 
        RUN trt_detailOl (INPUT  bf_ol.ol_num_ol,
                          INPUT  bf_ol.tt_fichier,
                          INPUT  lc_rep_ol,
                          OUTPUT ll_ok).
        IF NOT ll_ok THEN
            UNDO, LEAVE.    /* Defait le bloc transaction en cours. */
 
        DELETE bf_ol.
    END. /* ELSE (IF AVAILABLE GPI_PGOL) - TRANSACTION */
    ASSIGN gl_importOL = TRUE.
END. /* FOR EACH bf_ol */
 
/* Traiter s‚par‚ment OL et lignes pose plusieurs problŠmes, problŠmes de lignes dupliqu‚es, transaction incorrectement d‚faites, etc... *
/********************/
/*** DETAILS O.L. ***/
/********************/
ASSIGN li_jauge = 0.
FOR EACH bf_ligneOL:
    ASSIGN li_jauge = li_jauge + 1.
END.
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Traitement des d‚tails O.L. Veuillez patienter ... (<p>%)").
 
RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
FOR EACH bf_ligneOL BY bf_ligneOL.tt_fichier BY bf_ligneOL.tt_numligne:
 
    RUN jauge-next IN h_jauge-ocx.
 
    FIND FIRST tt_fichiers WHERE tt_fichiers.tt_fichier = bf_ligneOL.tt_fichier NO-ERROR.
    IF NOT AVAILABLE tt_fichiers THEN
    DO:
        CREATE tt_fichiers.
        ASSIGN
        tt_fichiers.tt_fichier     = bf_ligneOL.tt_fichier
        tt_fichiers.tt_rep_origine = lc_rep_ol.
    END. /* IF NOT AVAILABLE tt_fichiers */
 
    FIND FIRST GPI_PGDETAILOL WHERE GPI_PGDETAILOL.OL_NUM_OL = bf_ligneOL.OL_NUM_OL AND 
                                    GPI_PGDETAILOL.DETAILOL_NUM_LIGNE = bf_ligneOL.DETAILOL_NUM_LIGNE NO-LOCK NO-ERROR.
    IF AVAILABLE GPI_PGDETAILOL THEN
        RUN erreur (bf_ligneOL.tt_fichier, bf_ligneOL.tt_numligne, "La ligne d‚tail " + STRING(bf_ligneOL.DETAILOL_NUM_LIGNE) + " pour l'OL " + bf_ligneOL.OL_NUM_OL + " existe d‚j….").
    ELSE
    DO:
        /* Contr“le que l'OL existe */
        FIND FIRST GPI_PGOL WHERE GPI_PGOL.OL_NUM_OL = bf_ligneOL.OL_NUM_OL NO-LOCK NO-ERROR.
        IF NOT AVAILABLE GPI_PGOL THEN
            RUN erreur (bf_ligneOL.tt_fichier, bf_ligneOL.tt_numligne, "Cr‚ation de ligne impossible, l'OL " + bf_ligneOL.OL_NUM_OL + " n'existe pas.").
        ELSE
        DO:
            /* Contr“le que l'article existe */
            FIND FIRST GPI_PGARTI WHERE GPI_PGARTI.ARTI_NUM_ARTICLE = bf_ligneOL.ARTI_NUM_ARTICLE NO-LOCK NO-ERROR.
            IF NOT AVAILABLE GPI_PGARTI THEN
                RUN erreur (bf_ligneOL.tt_fichier, bf_ligneOL.tt_numligne, "Cr‚ation de ligne impossible, l'article " + STRING(bf_ligneOL.ARTI_NUM_ARTICLE) + " n'existe pas.").
            ELSE
            DO:
                CREATE GPI_PGDETAILOL.
                BUFFER-COPY bf_ligneOL TO GPI_PGDETAILOL.
                RELEASE GPI_PGDETAILOL.
                DELETE bf_ligneOL.
            END. /* ELSE (IF NOT AVAILABLE GPI_PGARTI) */
        END. /* ELSE (IF NOT AVAILABLE GPI_PGOL) */
    END. /* IF AVAILABLE GPI_PGDETAILOL */
END. /* FOR EACH bf_ligneOL */
*/
 
/*****************/
/**** CLIENTS ****/
/*****************/
RUN trt_clients (INPUT lc_rep_ol).
 
/**********************/
/*** INTERVENANTS ***/
/**********************/
/*
 Pour les intervenants, on a une difficult‚e, qui repose sur le fait que on a du annule et remplace, si plusieurs fichiers sont int‚gr‚s avec le mˆme client, c'est le dernier fichier qui sers de r‚f‚rence. 
 Il faut donc supprimer les intervenants existants si le client est pr‚sent, et ne les cr‚er que pour le dernier fichier du client
*/
FOR EACH tt_interClient BREAK BY tt_interClient.pgic_num_client BY tt_interClient.tt_fichier:
    /* Le dernier fichier est celui qui contient les bonnes informations */
    IF LAST-OF (tt_interClient.pgic_num_client) THEN
    DO TRANSACTION:
        /* On sais quel est le dernier fichier */
        /* Suppression des intervenants */
        FIND FIRST GPI_PGIC WHERE GPI_PGIC.pgic_num_client = tt_interClient.pgic_num_client EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        DO WHILE AVAILABLE GPI_PGIC:
            DELETE GPI_PGIC.
            FIND NEXT GPI_PGIC WHERE GPI_PGIC.pgic_num_client = tt_interClient.pgic_num_client EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        END.
        IF NOT AVAILABLE GPI_PGIC AND LOCKED GPI_PGIC THEN
        DO:
            RUN erreur (tt_interClient.tt_fichier, 0, "Erreur lors de la suppression des intervenants pour le client " + tt_interClient.pgic_num_client + ". Enregistrement en cours d'utilisation par un autre utilisateur.").
            UNDO, LEAVE.
        END.
 
        /* Cr‚ation des nouveaux */
        FOR EACH bf_interClient WHERE bf_interClient.pgic_num_client = tt_interClient.pgic_num_client AND bf_interClient.tt_fichier = tt_interClient.tt_fichier:
            /* Cr‚e l'intervenant */
            CREATE GPI_PGIC.
            BUFFER-COPY bf_interClient TO GPI_PGIC.
            RELEASE GPI_PGIC.
        END.
        /* On supprime les intervenants du client dans la tt */
        FOR EACH bf_interClient WHERE bf_interClient.pgic_num_client = tt_interClient.pgic_num_client:
            DELETE bf_interClient.
        END.
    END. /* TRANSACTION */
END.
 
/*************************/
/*** PLANS DE TOURNEES ***/
/*************************/
RUN trt_PlanTournee (INPUT lc_rep_param).
 
/******************************/
/*** INTERVENANTS EBERHARDT ***/
/******************************/
RUN trt_InterEber (INPUT lc_rep_param).

FOR EACH tt_fichiers NO-LOCK:
    /*
    FIND FIRST tt_erreur WHERE tt_erreur.fichier = tt_fichiers.tt_fichier NO-LOCK NO-ERROR.
    */
    FIND FIRST tt_erreur WHERE tt_erreur.fichier = tt_fichiers.tt_fichier AND tt_erreur.libErreur MATCHES "*OL * en cours de modification par un autre utilisateur*" NO-LOCK NO-ERROR.
    /* Si pas d'erreur, on peut d‚placer le fichier */
    /* NID le 07/03/21 - Soucis si les fichier restent trop longtemps, des OL qui n'existent plus sont toujour int‚gr‚s */
    /* De plus, tant qu'un OL n'est pas supprim‚ par Eberhadt ou trait‚ par Astr'In, il continue a ˆtre envoy‚. Inutile de le garder */
    /* Une exception, si l'erreur t lie … un blocage utilisateur, c'est passager */  
    IF NOT AVAILABLE tt_erreur THEN
    /*
    IF menu_utilisateur = "gpi" AND NOT gl_modeBatch THEN
        MESSAGE AVAILABLE tt_erreur SKIP
                tt_fichiers.tt_fichier SKIP 
                tt_fichiers.tt_rep_origine SKIP
                lc_rep_svg VIEW-AS ALERT-BOX.
    */
        RUN sauveFichier(tt_fichiers.tt_fichier, tt_fichiers.tt_rep_origine, lc_rep_svg).
END.

/* Ajout NID le 03/10/16 - Gestion des OL supprim‚s */
IF gl_importOL THEN
    RUN supOL.

/* Initialise les premiŠre dates de livraison PF possibles */
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "R‚partition Lot/PF. Veuillez patienter ...").
RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).

/* Une fois l'import termin‚, on lance la r‚partition PF/Lot */
RUN specif\PGI-EBERHARDT\p-PF_ou_lot.r.

/* Initialise les premiŠre dates de livraison PF possibles */
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Initialisation plateformes. Veuillez patienter ...").
RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).

RUN initPF.

/* Calcul de la date de livraison pour le OL affect‚s en PF */
RUN initDateLivraison.

/* Initialisation du statut Obligatoire pour les OL affect‚s en PF */
RUN majObligatoire.

RUN jauge-fin IN h_jauge-ocx.

FRAME F-jauge:MOVE-TO-BOTTOM().
FRAME F-jauge:VISIBLE = FALSE.

FRAME F-Criteres:VISIBLE = TRUE.
FRAME F-Criteres:MOVE-TO-TOP().

RETURN "".

END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE traitement_svg_161025 W-Win 
PROCEDURE traitement_svg_161025 :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
/*
DEFINE VARIABLE lc_rep_articles  AS CHARACTER NO-UNDO.
DEFINE VARIABLE lc_rep_ol        AS CHARACTER NO-UNDO.
DEFINE VARIABLE lc_rep_param     AS CHARACTER NO-UNDO.
 
DEFINE VARIABLE lc_rep_svg       AS CHARACTER NO-UNDO.
DEFINE VARIABLE li_jauge         AS INTEGER   NO-UNDO.
 
DEFINE VARIABLE ll_ok            AS LOGICAL   NO-UNDO.
DEFINE VARIABLE ll_enAffaire     AS LOGICAL   NO-UNDO.  /* Ajout NDE le 06/10/16 */
DEFINE VARIABLE li_heureTrt      AS INTEGER   NO-UNDO.
DEFINE VARIABLE li_minTrt        AS INTEGER   NO-UNDO.
DEFINE VARIABLE lc_heureMinTrt   AS CHARACTER NO-UNDO.
 
EMPTY TEMP-TABLE tt_fichiers.
 
/* Initialisation du r‚pertoire d'origine des articles */
ASSIGN
gl_importOL = FALSE
FILE-INFORMATION:FILE-NAME = fi_rep_integration_articles:SCREEN-VALUE IN FRAME F-Criteres
lc_rep_articles = FILE-INFORMATION:FULL-PATHNAME.
 
/* Initialisation du r‚pertoire d'origine des ol */
ASSIGN
FILE-INFORMATION:FILE-NAME = fi_rep_integration_ol:SCREEN-VALUE IN FRAME F-Criteres
lc_rep_ol = FILE-INFORMATION:FULL-PATHNAME.
 
/* Initialisation du r‚pertoire d'origine des paramŠtres */
ASSIGN
FILE-INFORMATION:FILE-NAME = fi_rep_integration_param:SCREEN-VALUE IN FRAME F-Criteres
lc_rep_param = FILE-INFORMATION:FULL-PATHNAME.
 
/* Initialisation du r‚pertoire de sauvegarde */
RUN genereRepSauvegarde(OUTPUT lc_rep_svg).
 
/* On va parcourir les diff‚rentes tables temporaire pour cr‚er les enregistrements en base. */
/* Si une erreur se produit, on cr‚e un enregistrement dans la table des erreurs.            */
/* Le fichier n'est d‚plac‚ que si aucune erreur ne s'est produite lors de son traitement.   */
 
/****************/
/*** ARTICLES ***/
/****************/
RUN trt_articles (INPUT lc_rep_articles).
 
/******************/
/*** COMPOSANTS ***/
/******************/
RUN trt_composants (INPUT lc_rep_articles).
 
/***************************/
/*** ORDRES DE LIVRAISON ***/
/***************************/
ASSIGN li_jauge = 0.
FOR EACH bf_ol:
    ASSIGN li_jauge = li_jauge + 1.
END.
 
ASSIGN
lc_heureMinTrt   = STRING(TIME, "HH:MM:SS")
li_heureTrt      = INTEGER(ENTRY(1, lc_heureMinTrt, ":"))
li_minTrt        = INTEGER(ENTRY(2, lc_heureMinTrt, ":")).
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Traitement des O.L. Veuillez patienter ... (<p>%)").
 
RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
FOR EACH bf_ol BY bf_ol.tt_fichier BY bf_ol.tt_numligne:
 
    RUN jauge-next IN h_jauge-ocx.
 
    FIND FIRST tt_fichiers WHERE tt_fichiers.tt_fichier = bf_ol.tt_fichier NO-ERROR.
    IF NOT AVAILABLE tt_fichiers THEN
    DO:
        CREATE tt_fichiers.
        ASSIGN
        tt_fichiers.tt_fichier     = bf_ol.tt_fichier
        tt_fichiers.tt_rep_origine = lc_rep_ol.
    END. /* IF NOT AVAILABLE tt_fichiers */
 
    /* Le 28/07/16 - Pour les OL, on est en annule et remplace, sauf si l'OL est d‚j… affect‚ … une affaire */
    FIND FIRST GPI_PGOL WHERE GPI_PGOL.OL_NUM_OL = bf_ol.OL_NUM_OL NO-LOCK NO-ERROR.
    IF AVAILABLE GPI_PGOL THEN
    DO:
        DO TRANSACTION:
            ASSIGN ll_enAffaire = FALSE.
            FIND CURRENT gpi_pgol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
            IF NOT AVAILABLE gpi_pgol THEN
            DO:
                IF LOCKED gpi_pgol THEN
                DO:
                    RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "OL " + bf_ol.OL_NUM_OL + " en cours de modification par un autre utilisateur.").
                    UNDO, LEAVE.    /* Defait le bloc transaction en cours. */
                END.
                ELSE
                DO:
                    CREATE GPI_PGOL.
                    BUFFER-COPY bf_ol EXCEPT bf_ol.ACT_ID bf_ol.OL_DATE-HEURE_CREATION
                             TO GPI_PGOL ASSIGN GPI_PGOL.OL_DATE-HEURE_CREATION = IF bf_ol.tt_dt_export = ? THEN NOW ELSE bf_ol.tt_dt_export
                                                GPI_PGOL.ol_date_dernier_envoi_eberhardt = TODAY
                                                GPI_PGOL.ol_heure_dernier_envoi_eberhardt = li_heureTrt
                                                GPI_PGOL.ol_minute_dernier_envoi_eberhard = li_minTrt.
                    RELEASE GPI_PGOL.
                END. /* ELSE (IF LOCKED gpi_pgol) */
            END. /* IF NOT AVAILABLE gpi_pgol */
            ELSE
            DO:
                /* On contr“le s'il faut faire la modification. Si l'OL n'est pas encore en pr‚paration, on peut le mettre … jour */
                IF gpi_pgol.ol_date_chg = ? AND GPI_PGOL.ol_date_RDV = ? AND gpi_pgol.ol_num_affaire = "" AND GPI_PGOL.ol_date_retour_Eber = ? THEN
                    BUFFER-COPY bf_ol EXCEPT bf_ol.ACT_ID bf_ol.OL_DATE-HEURE_CREATION TO GPI_PGOL.
                ELSE
                    ASSIGN ll_enAffaire = TRUE.
 
                ASSIGN
                GPI_PGOL.ol_date_dernier_envoi_eberhardt = TODAY
                GPI_PGOL.ol_heure_dernier_envoi_eberhardt = li_heureTrt
                GPI_PGOL.ol_minute_dernier_envoi_eberhard = li_minTrt.
 
                RELEASE GPI_PGOL.
            END.    /* ELSE (IF NOT AVAILABLE gpi_pgol) */
 
            /* On ne le fait que si l'entˆte a ‚t‚ modifi‚e */
            IF NOT ll_enAffaire THEN
            DO:
                /* Suppression des d‚tails OL li‚s pour pouvoir les recr‚er. On d‚fait la transaction si une erreur se produit */
                FIND FIRST GPI_PGDETAILOL WHERE GPI_PGDETAILOL.ol_num_ol = bf_ol.ol_num_ol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
                DO WHILE AVAILABLE GPI_PGDETAILOL:
                    DELETE GPI_PGDETAILOL NO-ERROR.
                    IF ERROR-STATUS:ERROR THEN
                    DO:
                        RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "Erreur lors de la mise … jours de l'OL " + bf_ol.OL_NUM_OL + ". Erreur " + ERROR-STATUS:GET-MESSAGE(ERROR-STATUS:NUM-MESSAGES)).
                        UNDO, LEAVE.
                    END.
                    FIND NEXT GPI_PGDETAILOL WHERE GPI_PGDETAILOL.ol_num_ol = bf_ol.ol_num_ol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
                END.
                IF LOCKED GPI_PGDETAILOL THEN
                DO:
                    RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "OL " + bf_ol.OL_NUM_OL + " en cours de modification par un autre utilisateur.").
                    UNDO, LEAVE.    /* Defait le bloc transaction en cours. */
                END.
 
                /* Recr‚e les nouvelles lignes de d‚tail */
                RUN trt_detailOl (INPUT  bf_ol.ol_num_ol,
                                  INPUT  bf_ol.tt_fichier,
                                  INPUT  lc_rep_ol,
                                  OUTPUT ll_ok).
                IF NOT ll_ok THEN
                    UNDO, LEAVE.    /* D‚fait le bloc transaction en cours. */
            END.
            DELETE bf_ol.
        END. /* IF gpi_pgol.ol_date_chargement = ? AND gpi_pgol.ol_num_affaire = "" - TRANSACTION */
 
        /*
            /* Sinon, on supprime simplement bf_ol, et les bf_ligneOL li‚s */
            ELSE
            DO: 
                FOR EACH bf_ligneOL WHERE bf_ligneOL.ol_num_ol = bf_ol.ol_num_ol AND bf_ligneOL.tt_fichier = bf_ol.tt_fichier:
                    DELETE bf_ligneOL.
                END.
                DELETE bf_ol.
            END.
            /*RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "OL " + bf_ol.OL_NUM_OL + " existe d‚j….").*/
        */
    END.
    ELSE
    DO TRANSACTION:
        CREATE GPI_PGOL.
        BUFFER-COPY bf_ol EXCEPT bf_ol.ACT_ID bf_ol.OL_DATE-HEURE_CREATION
                 TO GPI_PGOL ASSIGN GPI_PGOL.OL_DATE-HEURE_CREATION = IF bf_ol.tt_dt_export = ? THEN NOW ELSE bf_ol.tt_dt_export
                                    GPI_PGOL.ol_date_dernier_envoi_eberhardt  = TODAY
                                    GPI_PGOL.ol_heure_dernier_envoi_eberhardt = li_heureTrt
                                    GPI_PGOL.ol_minute_dernier_envoi_eberhard = li_minTrt.
        RELEASE GPI_PGOL.
 
        /* Suppression des d‚tails OL li‚s pour pouvoir les recr‚er. On d‚fait la transaction si une erreur se produit */
        /* Un magnifique bug supprimait les entˆtes sans supprimer les lignes. On a donc plein de lignes parasites … supprimer */
        FIND FIRST GPI_PGDETAILOL WHERE GPI_PGDETAILOL.ol_num_ol = bf_ol.ol_num_ol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        DO WHILE AVAILABLE GPI_PGDETAILOL:
            DELETE GPI_PGDETAILOL NO-ERROR.
            IF ERROR-STATUS:ERROR THEN
            DO:
                RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "Erreur lors de la mise … jours de l'OL " + bf_ol.OL_NUM_OL + ". Erreur " + ERROR-STATUS:GET-MESSAGE(ERROR-STATUS:NUM-MESSAGES)).
                UNDO, LEAVE.
            END.
            FIND NEXT GPI_PGDETAILOL WHERE GPI_PGDETAILOL.ol_num_ol = bf_ol.ol_num_ol EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        END.
        IF LOCKED GPI_PGDETAILOL THEN
        DO:
            RUN erreur (bf_ol.tt_fichier, bf_ol.tt_numligne, "OL " + bf_ol.OL_NUM_OL + " en cours de modification par un autre utilisateur.").
            UNDO, LEAVE.    /* Defait le bloc transaction en cours. */
        END.
 
        RUN trt_detailOl (INPUT  bf_ol.ol_num_ol,
                          INPUT  bf_ol.tt_fichier,
                          INPUT  lc_rep_ol,
                          OUTPUT ll_ok).
        IF NOT ll_ok THEN
            UNDO, LEAVE.    /* Defait le bloc transaction en cours. */
 
        DELETE bf_ol.
    END. /* ELSE (IF AVAILABLE GPI_PGOL) */
    ASSIGN gl_importOL = TRUE.
END. /* FOR EACH bf_ol */
 
/* Traiter s‚par‚ment OL et lignes pose plusieurs problŠmes, problŠmes de lignes dupliqu‚es, transaction incorrectement d‚faites, etc... *
/********************/
/*** DETAILS O.L. ***/
/********************/
ASSIGN li_jauge = 0.
FOR EACH bf_ligneOL:
    ASSIGN li_jauge = li_jauge + 1.
END.
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Traitement des d‚tails O.L. Veuillez patienter ... (<p>%)").
 
RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
FOR EACH bf_ligneOL BY bf_ligneOL.tt_fichier BY bf_ligneOL.tt_numligne:
 
    RUN jauge-next IN h_jauge-ocx.
 
    FIND FIRST tt_fichiers WHERE tt_fichiers.tt_fichier = bf_ligneOL.tt_fichier NO-ERROR.
    IF NOT AVAILABLE tt_fichiers THEN
    DO:
        CREATE tt_fichiers.
        ASSIGN
        tt_fichiers.tt_fichier     = bf_ligneOL.tt_fichier
        tt_fichiers.tt_rep_origine = lc_rep_ol.
    END. /* IF NOT AVAILABLE tt_fichiers */
 
    FIND FIRST GPI_PGDETAILOL WHERE GPI_PGDETAILOL.OL_NUM_OL = bf_ligneOL.OL_NUM_OL AND 
                                    GPI_PGDETAILOL.DETAILOL_NUM_LIGNE = bf_ligneOL.DETAILOL_NUM_LIGNE NO-LOCK NO-ERROR.
    IF AVAILABLE GPI_PGDETAILOL THEN
        RUN erreur (bf_ligneOL.tt_fichier, bf_ligneOL.tt_numligne, "La ligne d‚tail " + STRING(bf_ligneOL.DETAILOL_NUM_LIGNE) + " pour l'OL " + bf_ligneOL.OL_NUM_OL + " existe d‚j….").
    ELSE
    DO:
        /* Contr“le que l'OL existe */
        FIND FIRST GPI_PGOL WHERE GPI_PGOL.OL_NUM_OL = bf_ligneOL.OL_NUM_OL NO-LOCK NO-ERROR.
        IF NOT AVAILABLE GPI_PGOL THEN
            RUN erreur (bf_ligneOL.tt_fichier, bf_ligneOL.tt_numligne, "Cr‚ation de ligne impossible, l'OL " + bf_ligneOL.OL_NUM_OL + " n'existe pas.").
        ELSE
        DO:
            /* Contr“le que l'article existe */
            FIND FIRST GPI_PGARTI WHERE GPI_PGARTI.ARTI_NUM_ARTICLE = bf_ligneOL.ARTI_NUM_ARTICLE NO-LOCK NO-ERROR.
            IF NOT AVAILABLE GPI_PGARTI THEN
                RUN erreur (bf_ligneOL.tt_fichier, bf_ligneOL.tt_numligne, "Cr‚ation de ligne impossible, l'article " + STRING(bf_ligneOL.ARTI_NUM_ARTICLE) + " n'existe pas.").
            ELSE
            DO:
                CREATE GPI_PGDETAILOL.
                BUFFER-COPY bf_ligneOL TO GPI_PGDETAILOL.
                RELEASE GPI_PGDETAILOL.
                DELETE bf_ligneOL.
            END. /* ELSE (IF NOT AVAILABLE GPI_PGARTI) */
        END. /* ELSE (IF NOT AVAILABLE GPI_PGOL) */
    END. /* IF AVAILABLE GPI_PGDETAILOL */
END. /* FOR EACH bf_ligneOL */
*/
 
/****************/
/**** CLIENTS ***/
/****************/
RUN trt_clients (INPUT lc_rep_ol).
 
/**********************/
/*** INTERVENANTS ***/
/**********************/
/*
 Pour les intervenants, on a une difficult‚e, qui repose sur le fait que on a du annule et remplace, si plusieurs fichiers sont int‚gr‚s avec le mˆme client, c'est le dernier fichier qui sers de r‚f‚rence. 
 Il faut donc supprimer les intervenants existants si le client est pr‚sent, et ne les cr‚er que pour le dernier fichier du client
*/
FOR EACH tt_interClient BREAK BY tt_interClient.pgic_num_client BY tt_interClient.tt_fichier:
    /* Le dernier fichier est celui qui contient les bonnes informations */
    IF LAST-OF (tt_interClient.pgic_num_client) THEN
    DO TRANSACTION:
        /* On sais quel est le dernier fichier */
        /* Suppression des intervenants */
        FIND FIRST GPI_PGIC WHERE GPI_PGIC.pgic_num_client = tt_interClient.pgic_num_client EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        DO WHILE AVAILABLE GPI_PGIC:
            DELETE GPI_PGIC.
            FIND NEXT GPI_PGIC WHERE GPI_PGIC.pgic_num_client = tt_interClient.pgic_num_client EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        END.
        IF NOT AVAILABLE GPI_PGIC AND LOCKED GPI_PGIC THEN
        DO:
            RUN erreur (tt_interClient.tt_fichier, 0, "Erreur lors de la suppression des intervenants pour le client " + tt_interClient.pgic_num_client + ". Enregistrement en cours d'utilisation par un autre utilisateur.").
            UNDO, LEAVE.
        END.
 
        /* Cr‚ation des nouveaux */
        FOR EACH bf_interClient WHERE bf_interClient.pgic_num_client = tt_interClient.pgic_num_client AND bf_interClient.tt_fichier = tt_interClient.tt_fichier:
            /* Cr‚e l'intervenant */
            CREATE GPI_PGIC.
            BUFFER-COPY bf_interClient TO GPI_PGIC.
            RELEASE GPI_PGIC.
        END.
        /* On supprime les intervenants du client dans la tt */
        FOR EACH bf_interClient WHERE bf_interClient.pgic_num_client = tt_interClient.pgic_num_client:
            DELETE bf_interClient.
        END.
    END.
END.
 
/*************************/
/*** PLANS DE TOURNEES ***/
/*************************/
RUN trt_PlanTournee (INPUT lc_rep_param).
 
/******************************/
/*** INTERVENANTS EBERHARDT ***/
/******************************/
RUN trt_InterEber (INPUT lc_rep_param).
 
FOR EACH tt_fichiers NO-LOCK:
    FIND FIRST tt_erreur WHERE tt_erreur.fichier = tt_fichiers.tt_fichier NO-LOCK NO-ERROR.
    /* Si pas d'erreur, on peut d‚placer le fichier */
    IF NOT AVAILABLE tt_erreur THEN
        RUN sauveFichier(tt_fichiers.tt_fichier, tt_fichiers.tt_rep_origine, lc_rep_svg).
END.
 
/* Ajout NID le 03/10/16 - Gestion des OL supprim‚s */
IF gl_importOL THEN
    RUN supOL.
 
/* Initialise les premiŠre dates de livraison PF possibles */
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "R‚partition Lot/PF. Veuillez patienter ...").
RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
/* Une fois l'import termin‚, on lance la r‚partition PF/Lot */
RUN specif\PGI-EBERHARDT\p-PF_ou_lot.r.
 
/* Initialise les premiŠre dates de livraison PF possibles */
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Initialisation plateformes. Veuillez patienter ...").
RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
RUN initPF.
 
/* Calcul de la date de livraison pour le OL affect‚s en PF */
RUN initDateLivraison.
 
/* Initialisation du statut Obligatoire pour les OL affect‚s en PF */
RUN majObligatoire.
 
RUN jauge-fin IN h_jauge-ocx.
 
FRAME F-jauge:MOVE-TO-BOTTOM().
FRAME F-jauge:VISIBLE = FALSE.
 
FRAME F-Criteres:VISIBLE = TRUE.
FRAME F-Criteres:MOVE-TO-TOP().
 
RETURN "".
*/    /* Pourquoi suis je oblig‚ d'en mettre 2 ? */
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE trt_articles W-Win 
PROCEDURE trt_articles :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ic_rep_articles AS CHARACTER   NO-UNDO.
 
DEFINE VARIABLE li_jauge         AS INTEGER   NO-UNDO.
 
ASSIGN li_jauge = 0.
FOR EACH bf_article BY bf_article.tt_fichier BY bf_article.tt_numligne:
    ASSIGN li_jauge = li_jauge + 1.
END.
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Traitement des articles. Veuillez patienter ... (<p>%)").
 
RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
FOR EACH bf_article BY bf_article.tt_fichier BY bf_article.tt_numligne:
 
    RUN jauge-next IN h_jauge-ocx.
 
    FIND FIRST tt_fichiers WHERE tt_fichiers.tt_fichier = bf_article.tt_fichier NO-ERROR.
    IF NOT AVAILABLE tt_fichiers THEN
        RUN createFichier(bf_article.tt_fichier, ic_rep_articles).
 
    CASE bf_article.tt_action:
        WHEN "C" THEN
        DO:
            FIND FIRST GPI_PGARTI WHERE GPI_PGARTI.ARTI_NUM_ARTICLE = bf_article.ARTI_NUM_ARTICLE NO-LOCK NO-ERROR.
            IF NOT AVAILABLE GPI_PGARTI THEN
            DO TRANSACTION:
                CREATE GPI_PGARTI.
                BUFFER-COPY bf_article EXCEPT bf_article.ARTI_DATE-HEURE_CREATION bf_article.ARTI_DATE-HEURE_MODIFICATION bf_article.ARTI_DATE-HEURE_SUPPRESSION
                         TO GPI_PGARTI ASSIGN GPI_PGARTI.ARTI_DATE-HEURE_CREATION = IF bf_article.tt_dt_export = ? THEN NOW ELSE bf_article.tt_dt_export.
                RELEASE GPI_PGARTI.
                DELETE bf_article.
            END. /* IF NOT AVAILABLE GPI_PGARTI */
            ELSE
            /* Cas particulier de la cr‚ation d'un article supprim‚ auparavant */
            IF GPI_PGARTI.ARTI_DATE-HEURE_SUPPRESSION <> ? THEN
            DO TRANSACTION:
                FIND CURRENT GPI_PGARTI EXCLUSIVE-LOCK NO-ERROR.
                IF LOCKED GPI_PGARTI THEN
                    RUN erreur (bf_article.tt_fichier, bf_article.tt_numligne, "Modification de l'article nø " + STRING(bf_article.ARTI_NUM_ARTICLE) + " impossible : article en cours de modification par un autre utilisateur.").
                ELSE
                DO:
                    BUFFER-COPY bf_article EXCEPT bf_article.ARTI_DATE-HEURE_CREATION bf_article.ARTI_DATE-HEURE_MODIFICATION bf_article.ARTI_DATE-HEURE_SUPPRESSION
                             TO GPI_PGARTI ASSIGN GPI_PGARTI.ARTI_DATE-HEURE_MODIFICATION = IF bf_article.tt_dt_export = ? THEN NOW ELSE bf_article.tt_dt_export.
                                                  GPI_PGARTI.ARTI_DATE-HEURE_SUPPRESSION  = ?.
                    RELEASE GPI_PGARTI.
                    DELETE bf_article.
                END. /* ELSE (IF LOCKED GPI_PGARTI) */
            END. /* IF GPI_PGARTI.ARTI_DATE-HEURE_SUPPRESSION <> ? */
            ELSE
                RUN erreur (bf_article.tt_fichier, bf_article.tt_numligne, "Cr‚ation impossible : Article nø " + STRING(bf_article.ARTI_NUM_ARTICLE) + " existe d‚j….").
        END. /* WHEN "C" */
        WHEN "M" THEN
        DO TRANSACTION:
            FIND FIRST GPI_PGARTI WHERE GPI_PGARTI.ARTI_NUM_ARTICLE = bf_article.ARTI_NUM_ARTICLE EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
            IF AVAILABLE GPI_PGARTI THEN
            DO:
                BUFFER-COPY bf_article EXCEPT bf_article.ARTI_DATE-HEURE_CREATION bf_article.ARTI_DATE-HEURE_MODIFICATION bf_article.ARTI_DATE-HEURE_SUPPRESSION
                         TO GPI_PGARTI ASSIGN GPI_PGARTI.ARTI_DATE-HEURE_MODIFICATION = IF bf_article.tt_dt_export = ? THEN NOW ELSE bf_article.tt_dt_export.
                                              bf_article.ARTI_DATE-HEURE_SUPPRESSION  = ?.
                RELEASE GPI_PGARTI.
                DELETE bf_article.
            END. /* IF AVAILABLE GPI_PGARTI */
            ELSE
            DO:
                IF LOCKED GPI_PGARTI THEN
                    RUN erreur (bf_article.tt_fichier, bf_article.tt_numligne, "Modification de l'article nø " + STRING(bf_article.ARTI_NUM_ARTICLE) + " impossible : article en cours de modification par un autre utilisateur.").
                ELSE
                    RUN erreur (bf_article.tt_fichier, bf_article.tt_numligne, "Modification de l'article nø " + STRING(bf_article.ARTI_NUM_ARTICLE) + " impossible : l'article n'existe pas.").
            END. /* ELSE (IF AVAILABLE GPI_PGARTI) */
        END. /* WHEN "M" */
        WHEN "S" THEN
        DO TRANSACTION:
            FIND FIRST GPI_PGARTI WHERE GPI_PGARTI.ARTI_NUM_ARTICLE = bf_article.ARTI_NUM_ARTICLE EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
            IF AVAILABLE GPI_PGARTI THEN
            DO:
                IF GPI_PGARTI.ARTI_DATE-HEURE_SUPPRESSION <> ? THEN
                DO:
                    RUN erreur (bf_article.tt_fichier, bf_article.tt_numligne, "Suppression de l'article nø " + STRING(bf_article.ARTI_NUM_ARTICLE) + " impossible : article d‚j… supprim‚ le " + STRING(GPI_PGARTI.ARTI_DATE-HEURE_SUPPRESSION, "99/99/9999")).
                    RELEASE GPI_PGARTI.
                END. /* IF GPI_PGARTI.ARTI_DATE-HEURE_SUPPRESSION <> ? */
                ELSE
                DO:
                    ASSIGN GPI_PGARTI.ARTI_DATE-HEURE_SUPPRESSION = IF bf_article.tt_dt_export = ? THEN NOW ELSE bf_article.tt_dt_export.
 
                    RELEASE GPI_PGARTI.
                    DELETE bf_article.
                END. /* ELSE (IF GPI_PGARTI.ARTI_DATE-HEURE_SUPPRESSION <> ?) */
            END. /* IF AVAILABLE GPI_PGARTI */
            ELSE
            DO:
                IF LOCKED GPI_PGARTI THEN
                    RUN erreur (bf_article.tt_fichier, bf_article.tt_numligne, "Suppression de l'article nø " + STRING(bf_article.ARTI_NUM_ARTICLE) + " impossible : article en cours de modification par un autre utilisateur.").
                ELSE
                    RUN erreur (bf_article.tt_fichier, bf_article.tt_numligne, "Suppression de l'article nø " + STRING(bf_article.ARTI_NUM_ARTICLE) + " impossible : l'article n'existe pas.").
            END. /* ELSE (IF AVAILABLE GPI_PGARTI) */
        END. /* WHEN "S" */
        OTHERWISE
            RUN erreur (bf_article.tt_fichier, bf_article.tt_numligne, "Action " + bf_article.tt_action + " non g‚r‚e.").
    END CASE. /* CASE bf_article.tt_action */
END. /* FOR EACH bf_article */
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE trt_clients W-Win 
PROCEDURE trt_clients :
/*------------------------------------------------------------------------------
  Purpose:     Int‚gration des clients Astr'In, issus de la base Access, et modification par import OL
  Parameters:  <none>
  Notes:       Modifi‚e le 20/10/16 pour g‚rer l'import des clients depuis la base Access
               Ca complique un poil l'algo :)
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ic_rep_ol  AS CHARACTER   NO-UNDO.
 
DEFINE VARIABLE li_jauge AS INTEGER     NO-UNDO.
 
/* Vidage de la table si demand‚ */
IF tg_reinitClients THEN
DO:
    ASSIGN li_jauge = 0.
    SELECT COUNT(*) INTO li_jauge FROM gpi_pgcli.
 
    RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Suppression des clients. Veuillez patienter ... (<p>%)").
 
    RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
    FOR EACH gpi_pgcli EXCLUSIVE-LOCK:
        RUN jauge-next IN h_jauge-ocx.
        DELETE gpi_pgcli.
    END.
 
    ASSIGN li_jauge = 0.
    SELECT COUNT(*) INTO li_jauge FROM gpi_pgcliint.
 
    RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Suppression des interlocuteurs. Veuillez patienter ... (<p>%)").
 
    RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
    FOR EACH gpi_pgcliint EXCLUSIVE-LOCK:
        RUN jauge-next IN h_jauge-ocx.
        DELETE gpi_pgcliint.
    END.
END.
 
ASSIGN li_jauge = 0.
FOR EACH tt_client:
    ASSIGN li_jauge = li_jauge + 1.
END.
FOR EACH tt_cliint:
    ASSIGN li_jauge = li_jauge + 1.
END.
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Traitement des clients. Veuillez patienter ... (<p>%)").
 
RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
/* On va d'abord traiter l'int‚gration depuis le fichier client */
FOR EACH tt_client WHERE tt_client.tt_origine = "CLI" BREAK BY tt_client.cli_num_client BY tt_client.cli_cp BY tt_client.tt_fichier:
 
    RUN jauge-next IN h_jauge-ocx.
 
    /* On prend la derniŠre occurrence du client tout fichiers confondus */
    IF LAST-OF (tt_client.cli_num_client) OR LAST-OF (tt_client.cli_cp) THEN
    DO TRANSACTION:
        FIND FIRST tt_fichiers WHERE tt_fichiers.tt_fichier = tt_client.tt_fichier NO-ERROR.
        IF NOT AVAILABLE tt_fichiers THEN
            RUN createFichier(tt_client.tt_fichier, ic_rep_ol).
 
        FIND FIRST gpi_pgcli WHERE gpi_pgcli.cli_num_client = tt_client.cli_num_client AND gpi_pgcli.cli_cp = tt_client.cli_cp EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        IF NOT AVAILABLE gpi_pgcli THEN
        DO:
            /* Soit il est bloqu‚ par un autre utilisateur, et on remonte une erreur */
            IF LOCKED gpi_pgcli THEN
                RUN erreur (tt_client.tt_fichier, tt_client.tt_numligne, "Le client nø" + tt_client.cli_num_client + " pour le CP " + tt_client.cli_cp + " est en cours de modification par un autre utilisateur.").
            ELSE
            DO:
                CREATE gpi_pgcli.
                BUFFER-COPY tt_client TO gpi_pgcli.
            END.
        END.
        ELSE
            ASSIGN
            GPI_PGCLI.cli_lot     = tt_client.cli_lot
            GPI_PGCLI.cli_hayon   = tt_client.cli_hayon
            GPI_PGCLI.cli_ddeRDV  = tt_client.cli_ddeRDV
            GPI_PGCLI.cli_mail    = tt_client.cli_mail
            GPI_PGCLI.cli_confrdv = tt_client.cli_confrdv
            GPI_PGCLI.cli_com     = tt_client.cli_com.
        RELEASE gpi_pgcli.
    END.
END. /* FOR EACH tt_client */
 
/* Ajout NID le 27/10/16 - On n'en tiens plus compte */
/*
/* Ensuite, on traite les cr‚ations/modifications depuis le fichier des OL */
FOR EACH tt_client WHERE tt_client.tt_origine = "OL" BREAK BY tt_client.cli_num_client BY tt_client.cli_cp BY tt_client.tt_fichier:
 
    RUN jauge-next IN h_jauge-ocx.
 
    /* On prend la derniŠre occurrence du client tout fichiers confondus */
    IF LAST-OF (tt_client.cli_num_client) OR LAST-OF (tt_client.cli_cp) THEN
    DO TRANSACTION:
        FIND FIRST tt_fichiers WHERE tt_fichiers.tt_fichier = tt_client.tt_fichier NO-ERROR.
        IF NOT AVAILABLE tt_fichiers THEN
        DO:
            CREATE tt_fichiers.
            ASSIGN
            tt_fichiers.tt_fichier     = tt_client.tt_fichier
            tt_fichiers.tt_rep_origine = ic_rep_ol.
        END. /* IF NOT AVAILABLE tt_fichiers */
 
        FIND FIRST gpi_pgcli WHERE gpi_pgcli.cli_num_client = tt_client.cli_num_client AND gpi_pgcli.cli_cp = tt_client.cli_cp EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        IF NOT AVAILABLE gpi_pgcli THEN
        DO:
            /* Soit il est bloqu‚ par un autre utilisateur, et on remonte une erreur */
            IF LOCKED gpi_pgcli THEN
                RUN erreur (tt_client.tt_fichier, tt_client.tt_numligne, "Le client nø" + tt_client.cli_num_client + " pour le CP " + tt_client.cli_cp + " est en cours de modification par un autre utilisateur.").
            ELSE
            DO:
                CREATE gpi_pgcli.
                BUFFER-COPY tt_client TO gpi_pgcli.
            END.
        END.
        ELSE
            ASSIGN
            gpi_pgcli.cli_hayon  = tt_client.cli_hayon
            gpi_pgcli.cli_dderdv = tt_client.cli_dderdv
            gpi_pgcli.cli_lot    = tt_client.cli_lot.
        RELEASE gpi_pgcli.
    END.
END. /* FOR EACH tt_client */
*/
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Traitement des interlocuteurs. Veuillez patienter ... (<p>%)").
/* Traitement des interlocuteurs */
FOR EACH tt_cliint BREAK BY tt_cliint.cli_num_client BY tt_cliint.cli_cp BY tt_cliint.cliint_cle BY tt_cliint.tt_fichier:
 
    RUN jauge-next IN h_jauge-ocx.
 
    /* On prend la derniŠre occurence du client tout fichiers confondus */
    IF LAST-OF (tt_cliint.cli_num_client) OR LAST-OF (tt_cliint.cli_cp) OR LAST-OF(tt_cliint.cliint_cle) THEN
    DO TRANSACTION:
        FIND FIRST tt_fichiers WHERE tt_fichiers.tt_fichier = tt_cliint.tt_fichier NO-ERROR.
        IF NOT AVAILABLE tt_fichiers THEN
            RUN createFichier(tt_cliint.tt_fichier, ic_rep_ol).
 
        FIND FIRST gpi_pgcliint WHERE GPI_PGCLIINT.cli_num_client = tt_cliint.cli_num_client 
                                  AND GPI_PGCLIINT.cli_cp = tt_cliint.cli_cp 
                                  AND GPI_PGCLIINT.cliint_cle = tt_cliint.cliint_cle EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        IF NOT AVAILABLE GPI_PGCLIINT THEN
        DO:
            /* Soit il est bloqu‚ par un autre utilisateur, et on remonte une erreur */
            IF LOCKED GPI_PGCLIINT THEN
                RUN erreur (tt_cliint.tt_fichier, tt_cliint.tt_numligne, "L'interlocuteur " + tt_cliint.cliint_intervenant + " pour le client nø" + tt_cliint.cli_num_client + " et le CP " + tt_cliint.cli_cp + " est en cours de modification par un autre utilisateur.").
            ELSE
            DO:
                CREATE GPI_PGCLIINT.
                BUFFER-COPY tt_cliint TO GPI_PGCLIINT.
            END.
        END.
        ELSE
            ASSIGN
            GPI_PGCLIINT.cliint_intervenant = tt_cliint.cliint_intervenant
            GPI_PGCLIINT.cliint_tel         = tt_cliint.cliint_tel
            GPI_PGCLIINT.cliint_fax         = tt_cliint.cliint_fax
            GPI_PGCLIINT.cliint_mail        = tt_cliint.cliint_mail.
        RELEASE GPI_PGCLIINT.
    END.
END. /* FOR EACH tt_cliint */
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE trt_composants W-Win 
PROCEDURE trt_composants :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ic_rep_articles AS CHARACTER   NO-UNDO.
 
DEFINE VARIABLE li_jauge AS INTEGER     NO-UNDO.
DEFINE BUFFER bf_artc FOR GPI_PGARTC.

/* NID le 05/05/21 - Mise … jour des composants. Si un article est compos‚, et qu'on lui transmet des composants, ils annulent et remplacent tout composant ant‚rieur.
Il faut donc supprimer les composants existants des articles pour lesquels des composants ont ‚t‚ envoy‚ */
ASSIGN li_jauge = 0.
FOR EACH bf_composant BREAK BY bf_composant.arti_num_article:
    
    IF FIRST-OF(bf_composant.arti_num_article) THEN
        ASSIGN li_jauge = li_jauge + 1.
END.
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Suppression des composants. Veuillez patienter ...  (<p>%)").
 
RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).

FOR EACH bf_composant BREAK BY bf_composant.arti_num_article:
    
    IF FIRST-OF(bf_composant.arti_num_article) THEN
    DO:
        RUN jauge-next IN h_jauge-ocx.
        FOR EACH GPI_PGARTC WHERE GPI_PGARTC.ARTI_NUM_ARTICLE = bf_composant.ARTI_NUM_ARTICLE NO-LOCK:
            DO TRANSACTION:
                FIND FIRST bf_artc WHERE ROWID(bf_artc) = ROWID(GPI_PGARTC) EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
                IF AVAILABLE bf_artc THEN
                    DELETE bf_artc.
            END.
        END.
    END.
END.
/* Fin NID le 05/05/21 */

ASSIGN li_jauge = 0.
FOR EACH bf_composant:
    ASSIGN li_jauge = li_jauge + 1.
END.
 
RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Traitement des composants. Veuillez patienter ...  (<p>%)").
 
RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
FOR EACH bf_composant BY bf_composant.tt_fichier BY bf_composant.tt_numligne:

    RUN jauge-next IN h_jauge-ocx.
 
    FIND FIRST tt_fichiers WHERE tt_fichiers.tt_fichier = bf_composant.tt_fichier NO-ERROR.
    IF NOT AVAILABLE tt_fichiers THEN
        RUN createFichier(bf_composant.tt_fichier, ic_rep_articles).
 
    DO TRANSACTION:
        FIND FIRST GPI_PGARTC WHERE GPI_PGARTC.ARTI_NUM_ARTICLE = bf_composant.ARTI_NUM_ARTICLE AND
                                    GPI_PGARTC.ARTC_NUM_COMPOSANT = bf_composant.ARTC_NUM_COMPOSANT EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        IF AVAILABLE GPI_PGARTC THEN
        DO:
            ASSIGN GPI_PGARTC.ARTC_QUANTITE = bf_composant.ARTC_QUANTITE.
            RELEASE GPI_PGARTC.
            DELETE bf_composant.
        END. /* IF AVAILABLE GPI_PGARTC */
        ELSE
        DO:
            IF LOCKED GPI_PGARTC THEN
                RUN erreur(bf_composant.tt_fichier, bf_composant.tt_numligne, "Composant nø " + STRING(bf_composant.ARTC_NUM_COMPOSANT) + " de l'article " + STRING(bf_composant.ARTI_NUM_ARTICLE)+ " en cours de modification par un autre utilisateur.").
            ELSE
            DO:
                /* Avant de cr‚er le composant, on v‚rifie que l'article existe */
                FIND FIRST GPI_PGARTI WHERE GPI_PGARTI.ARTI_NUM_ARTICLE = bf_composant.ARTI_NUM_ARTICLE NO-LOCK NO-ERROR.
                IF NOT AVAILABLE GPI_PGARTI OR GPI_PGARTI.ARTI_DATE-HEURE_SUPPRESSION <> ? THEN
                    RUN erreur (bf_composant.tt_fichier, bf_composant.tt_numligne, "Cr‚ation composant impossible, l'article compos‚ nø" + STRING(bf_composant.ARTI_NUM_ARTICLE) + " n'existe pas ou a ‚t‚ supprim‚.").
                ELSE
                DO:
                    FIND FIRST GPI_PGARTI WHERE GPI_PGARTI.ARTI_NUM_ARTICLE = bf_composant.ARTC_NUM_COMPOSANT NO-LOCK NO-ERROR.
                    IF NOT AVAILABLE GPI_PGARTI OR GPI_PGARTI.ARTI_DATE-HEURE_SUPPRESSION <> ? THEN
                        RUN erreur (bf_composant.tt_fichier, bf_composant.tt_numligne, "Cr‚ation composant impossible, l'article composant nø" + STRING(bf_composant.ARTC_NUM_COMPOSANT) + " n'existe pas ou a ‚t‚ supprim‚.").
                    ELSE
                    DO:
                        CREATE GPI_PGARTC.
                        ASSIGN
                        GPI_PGARTC.ARTI_NUM_ARTICLE     = bf_composant.ARTI_NUM_ARTICLE
                        GPI_PGARTC.ARTC_NUM_COMPOSANT   = bf_composant.ARTC_NUM_COMPOSANT
                        GPI_PGARTC.ARTC_QUANTITE        = bf_composant.ARTC_QUANTITE.
                        RELEASE GPI_PGARTC.
                        DELETE bf_composant.
                    END. /* IF NOT AVAILABLE GPI_PGARTI */
                END. /* ELSE (IF NOT AVAILABLE GPI_PGARTI) */
            END. /* ELSE (IF LOCKED GPI_PGARTC) */
        END. /* ELSE (IF AVAILABLE GPI_PGARTC) */
    END. /* DO TRANSACTION */
END. /* FOR EACH bf_composant */
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE trt_detailOl W-Win 
PROCEDURE trt_detailOl :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ic_num_ol  AS CHARACTER   NO-UNDO.
DEFINE INPUT  PARAMETER ic_fichier AS CHARACTER   NO-UNDO.
DEFINE INPUT  PARAMETER ic_rep_ol  AS CHARACTER   NO-UNDO.
DEFINE OUTPUT PARAMETER ol_ok      AS LOGICAL     NO-UNDO.
 
ASSIGN ol_ok = TRUE.
FOR EACH bf_ligneOL WHERE bf_ligneOl.ol_num_ol = ic_num_ol
                      AND bf_ligneOL.tt_fichier = ic_fichier
                    BY bf_ligneOL.tt_numligne:
 
    FIND FIRST tt_fichiers WHERE tt_fichiers.tt_fichier = bf_ligneOL.tt_fichier NO-ERROR.
    IF NOT AVAILABLE tt_fichiers THEN
        RUN createFichier(bf_ligneOL.tt_fichier, ic_rep_ol).
 
    FIND FIRST GPI_PGDETAILOL WHERE GPI_PGDETAILOL.OL_NUM_OL = bf_ligneOL.OL_NUM_OL AND 
                                    GPI_PGDETAILOL.DETAILOL_NUM_LIGNE = bf_ligneOL.DETAILOL_NUM_LIGNE NO-LOCK NO-ERROR.
    IF AVAILABLE GPI_PGDETAILOL THEN
    DO:
        RUN erreur (bf_ligneOL.tt_fichier, bf_ligneOL.tt_numligne, "La ligne d‚tail " + STRING(bf_ligneOL.DETAILOL_NUM_LIGNE) + " pour l'OL " + bf_ligneOL.OL_NUM_OL + " existe d‚j….").
        ASSIGN ol_ok = FALSE.
    END.
    ELSE
    DO:
        /* Contr“le que l'OL existe */
        /*FIND FIRST GPI_PGOL WHERE GPI_PGOL.OL_NUM_OL = bf_ligneOL.OL_NUM_OL AND GPI_PGOL.ol_date_retour_Eber = ? NO-LOCK NO-ERROR. */
        FIND FIRST GPI_PGOL WHERE GPI_PGOL.OL_NUM_OL = bf_ligneOL.OL_NUM_OL NO-LOCK NO-ERROR.
        IF NOT AVAILABLE GPI_PGOL THEN
        DO:
            RUN erreur (bf_ligneOL.tt_fichier, bf_ligneOL.tt_numligne, "Cr‚ation de ligne impossible, l'OL " + bf_ligneOL.OL_NUM_OL + " n'existe pas.").
            ASSIGN ol_ok = FALSE.
        END.
        ELSE
        DO:
            /* Contr“le que l'article existe */
            FIND FIRST GPI_PGARTI WHERE GPI_PGARTI.ARTI_NUM_ARTICLE = bf_ligneOL.ARTI_NUM_ARTICLE NO-LOCK NO-ERROR.
            IF NOT AVAILABLE GPI_PGARTI THEN
            DO:
                RUN erreur (bf_ligneOL.tt_fichier, bf_ligneOL.tt_numligne, "Cr‚ation de ligne impossible pour l'OL " + bf_ligneOL.OL_NUM_OL + ", l'article " + STRING(bf_ligneOL.ARTI_NUM_ARTICLE) + " n'existe pas.").
                ASSIGN ol_ok = FALSE.
            END.
            ELSE
            DO:
                CREATE GPI_PGDETAILOL.
                BUFFER-COPY bf_ligneOL TO GPI_PGDETAILOL.
                RELEASE GPI_PGDETAILOL.
                DELETE bf_ligneOL.
            END. /* ELSE (IF NOT AVAILABLE GPI_PGARTI) */
        END. /* ELSE (IF NOT AVAILABLE GPI_PGOL) */
    END. /* IF AVAILABLE GPI_PGDETAILOL */
END. /* FOR EACH bf_ligneOL */
 
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE trt_InterEber W-Win 
PROCEDURE trt_InterEber :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ic_rep_param AS CHARACTER   NO-UNDO.
 
DEFINE VARIABLE li_jauge AS INTEGER     NO-UNDO.
DEFINE VARIABLE ll_ok    AS LOGICAL     NO-UNDO.
 
DEFINE VARIABLE lc_fichier AS CHARACTER NO-UNDO.
 
/* Traitement … ne faire que si on a quelque chose … int‚grer */
IF CAN-FIND (FIRST bf_interEber) THEN
DO:
    /*
     Fonctionne sur le mode Annule et remplace, on supprime d'abord tout avant de recharger la table
     Si plusieurs fichiers sont disponibles, seul le dernier est int‚gr‚
    */
    FIND LAST bf_interEber USE-INDEX idx_tri NO-ERROR.
    ASSIGN lc_fichier = bf_interEber.tt_fichier.
 
    ASSIGN li_jauge = 0.
    FOR EACH GPI_PGIE NO-LOCK:
        ASSIGN li_jauge = li_jauge + 1.
    END.
 
    RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Suppression des interlocuteurs Eberhardt. Veuillez patienter ... (<p>%)").
 
    RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
    /* En cas de problŠme de suppression, on d‚fait tout et on ne lance pas l'import */
    ASSIGN ll_ok = TRUE.
    suppression_pgie:
    DO TRANSACTION:
        FIND FIRST GPI_PGIE EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        DO WHILE AVAILABLE GPI_PGIE:
 
            RUN jauge-next IN h_jauge-ocx.
 
            DELETE GPI_PGIE NO-ERROR.
            IF ERROR-STATUS:ERROR THEN
            DO:
                ASSIGN ll_ok = FALSE.
                UNDO suppression_pgie, LEAVE suppression_pgie.
            END.
 
            FIND NEXT GPI_PGIE EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        END.
        IF LOCKED GPI_PGIE THEN
        DO:
            ASSIGN ll_ok = FALSE.
            UNDO suppression_pgie, LEAVE suppression_pgie.
        END.
    END.
 
    /* Tout s'est bien pass‚ */
    IF ll_ok THEN
    DO:
        ASSIGN li_jauge = 0.
        FOR EACH bf_interEber:
            ASSIGN li_jauge = li_jauge + 1.
        END.
 
        RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Traitement des interlocuteurs Eberhardt. Veuillez patienter ... (<p>%)").
 
        RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
        FOR EACH bf_interEber WHERE bf_interEber.tt_fichier = lc_fichier BY bf_interEber.tt_fichier BY bf_interEber.tt_numligne:
 
            RUN jauge-next IN h_jauge-ocx.
 
            FIND FIRST tt_fichiers WHERE tt_fichiers.tt_fichier = bf_interEber.tt_fichier NO-ERROR.
            IF NOT AVAILABLE tt_fichiers THEN
                RUN createFichier(bf_interEber.tt_fichier, ic_rep_param).
 
            FIND FIRST GPI_PGIE WHERE GPI_PGIE.pgie_idts = bf_interEber.pgie_idts NO-LOCK NO-ERROR.
            IF AVAILABLE GPI_PGIE THEN
                RUN erreur (bf_interEber.tt_fichier, bf_interEber.tt_numligne, "L'interlocuteur " + bf_interEber.pgie_prenom + " " + bf_interEber.pgie_nom + " de code " + bf_interEber.pgie_idts + " est pr‚sent plusieurs fois dans le fichier.").
            ELSE
            DO:
                CREATE GPI_PGIE.
                BUFFER-COPY bf_interEber TO GPI_PGIE.
                RELEASE GPI_PGIE.
                DELETE bf_interEber.
            END. /* ELSE (IF AVAILABLE GPI_PGOL) */
        END. /* FOR EACH bf_interEber */
    END. /* IF ll_ok */
END.    /* IF CAN-FIND (FIRST bf_interEber) */
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE trt_PlanTournee W-Win 
PROCEDURE trt_PlanTournee :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ic_rep_param AS CHARACTER   NO-UNDO.
 
DEFINE VARIABLE li_jauge AS INTEGER     NO-UNDO.
DEFINE VARIABLE ll_ok    AS LOGICAL     NO-UNDO.
 
/* Traitement … ne faire que si on a quelque chose … int‚grer */
IF CAN-FIND (FIRST bf_plan) THEN
DO:
    /* Fonctionne sur le mode Annule et remplace, on supprime d'abord tout avant de recharger la table */
    ASSIGN li_jauge = 0.
    FOR EACH GPI_PGPT NO-LOCK:
        ASSIGN li_jauge = li_jauge + 1.
    END.
 
    RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Suppression du plan de tourn‚es. Veuillez patienter ... (<p>%)").
 
    RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
    /* En cas de problŠme de suppression, on d‚fait tout et on ne lance pas l'import */
    ASSIGN ll_ok = TRUE.
    suppression_pgpt:
    DO TRANSACTION:
        FIND FIRST GPI_PGPT EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        DO WHILE AVAILABLE GPI_PGPT:
 
            RUN jauge-next IN h_jauge-ocx.
 
            DELETE GPI_PGPT NO-ERROR.
            IF ERROR-STATUS:ERROR THEN
            DO:
                ASSIGN ll_ok = FALSE.
                UNDO suppression_pgpt, LEAVE suppression_pgpt.
            END.
 
            FIND NEXT GPI_PGPT EXCLUSIVE-LOCK NO-WAIT NO-ERROR.
        END.
        IF LOCKED GPI_PGPT THEN
        DO:
            ASSIGN ll_ok = FALSE.
            UNDO suppression_pgpt, LEAVE suppression_pgpt.
        END.
    END.
 
    /* Tout s'est bien pass‚ */
    IF ll_ok THEN
    DO:
        ASSIGN li_jauge = 0.
        FOR EACH bf_plan:
            ASSIGN li_jauge = li_jauge + 1.
        END.
 
        RUN jauge-set-libelle IN h_jauge-ocx (INPUT "Traitement des plans de tourn‚es. Veuillez patienter ... (<p>%)").
 
        RUN jauge-init IN h_jauge-ocx ( INPUT li_jauge ).
 
        FOR EACH bf_plan BY bf_plan.tt_fichier BY bf_plan.tt_numligne:
 
            RUN jauge-next IN h_jauge-ocx.
 
            FIND FIRST tt_fichiers WHERE tt_fichiers.tt_fichier = bf_plan.tt_fichier NO-ERROR.
            IF NOT AVAILABLE tt_fichiers THEN
                RUN createFichier(bf_plan.tt_fichier, ic_rep_param).
 
            FIND FIRST GPI_PGPT WHERE GPI_PGPT.pgpt_codpos = bf_plan.pgpt_codpos 
                                  AND GPI_PGPT.pgpt_cpidx  = bf_plan.pgpt_cpidx  NO-LOCK NO-ERROR.
            IF AVAILABLE GPI_PGPT THEN
                RUN erreur (bf_plan.tt_fichier, bf_plan.tt_numligne, "La ville " + bf_plan.pgpt_ville + " de CP " + bf_plan.pgpt_codpos + " et nø Idx " + STRING(bf_plan.pgpt_cpidx) + " existe d‚j….").
            ELSE
            DO:
                CREATE GPI_PGPT.
                BUFFER-COPY bf_plan TO GPI_PGPT.
                RELEASE GPI_PGPT.
                DELETE bf_plan.
            END. /* ELSE (IF AVAILABLE GPI_PGOL) */
        END. /* FOR EACH bf_plan */
    END. /* IF ll_ok */
END.    /* IF CAN-FIND (FIRST bf_plan) */
 
END PROCEDURE.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
/* ************************  Function Implementations ***************** */
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION fi_ferie W-Win 
FUNCTION fi_ferie RETURNS LOGICAL
  ( INPUT idt_jour AS DATE ) :
/*------------------------------------------------------------------------------
  Purpose:  Retourne vrai si jour non travaill‚, c…d d‚clar‚ comme feri‚, ou samedi/dimanche
    Notes:  
------------------------------------------------------------------------------*/
 
FIND FIRST gpi_ferie WHERE GPI_FERIE.ferie_date = idt_jour NO-LOCK NO-ERROR.
IF AVAILABLE gpi_ferie OR WEEKDAY(idt_jour) = 1 OR WEEKDAY(idt_jour) = 7 THEN
    RETURN TRUE.
 
RETURN FALSE.
 
END FUNCTION.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION fi_listeMailsValide W-Win 
FUNCTION fi_listeMailsValide RETURNS LOGICAL
  (listeEmails AS CHARACTER) :
/*------------------------------------------------------------------------------
  Purpose:  Controle si la chaine pass‚e en paramŠtre est une liste de mails valides.
            Chacun des mails doit ˆtre valide (voir fi_mailValide), et chaque mail peut 
            ˆtre s‚par‚ par les caractŠres , ou ;
    Notes:  
------------------------------------------------------------------------------*/
DEFINE VARIABLE l_listeMailCorrige AS CHARACTER   NO-UNDO.  /* La liste des emails corrig‚e, dans laquelle on a remplac‚ le , par des ; */
DEFINE VARIABLE l_mail             AS CHARACTER   NO-UNDO. 
DEFINE VARIABLE l_ii               AS INTEGER     NO-UNDO. 
DEFINE VARIABLE l_ok               AS LOGICAL     NO-UNDO.
 
ASSIGN listeEmails = TRIM(listeEmails).
 
IF listeEmails = ? OR listeEmails = "" THEN RETURN FALSE.
 
/* On commence par remplacer les virgules par des ; */
ASSIGN l_listeMailCorrige = REPLACE(listeEmails, ",", ";").
 
l_ok = TRUE.
DO l_ii = 1 TO NUM-ENTRIES(l_listeMailCorrige, ";"):
 
    ASSIGN l_mail = ENTRY(l_ii, l_listeMailCorrige, ";").
 
    IF NOT fi_mailValide(l_mail) THEN
        l_ok = FALSE.
END.
RETURN l_ok.
 
END FUNCTION.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION fi_mailValide W-Win 
FUNCTION fi_mailValide RETURNS LOGICAL
  (email AS CHARACTER ) :
/*------------------------------------------------------------------------------
  Purpose:  Contr“le que la chaine pass‚e en entr‚e est une adresse mail valide
            c'est a dire non vide, et au format a@b.c
            a peut contenir des ., -, _ etc, mais une adresse mail ne contien qu'une 
            et une seule @, et les diff‚rents ‚l‚ments a, b et c peuvent contenir des points, 
            mais il faut au moins un points sur la partie qui suis l'@.
    Notes:  
------------------------------------------------------------------------------*/
 
email = TRIM(email).
 
IF email = ? OR email = "" THEN RETURN FALSE.
 
IF NUM-ENTRIES(email, "@") <> 2 THEN RETURN FALSE.
 
IF NUM-ENTRIES(ENTRY(2, email, "@"), ".") < 2 THEN RETURN FALSE.
 
RETURN TRUE.
 
END FUNCTION.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION fi_recupDateHeure W-Win 
FUNCTION fi_recupDateHeure RETURNS DATETIME
  ( ic_fichier AS CHARACTER ) :
/*------------------------------------------------------------------------------
  Purpose:  Extrait les dates et heure de g‚n‚ration du fichier lorsque c'est possible
    Notes:  
------------------------------------------------------------------------------*/
DEFINE VARIABLE ldt_integration AS DATETIME    NO-UNDO.
DEFINE VARIABLE li_testDate     AS INTEGER     NO-UNDO.
 
DEFINE VARIABLE lcDateHeure     AS CHARACTER   NO-UNDO.
DEFINE VARIABLE lcFichier       AS CHARACTER   NO-UNDO.
 
ASSIGN
lcFichier = ENTRY(NUM-ENTRIES(ic_fichier, "\"), ic_fichier, "\")
lcFichier = ENTRY(1, lcFichier, ".")
lcDateHeure = SUBSTRING(lcFichier, 7, 10).  /* Le fichier commence par 6 caractères alpha (PGARTI, PGPACP, PGARTC, PGEMIW) suivi de AAMMJJHHMM. */
 
ASSIGN li_testDate = INTEGER(lcDateHeure) NO-ERROR. /* On v‚rifie qu'il s'agit bien d'un nombre. On aura un problŠme le 01/01/2022 */
 
/* Il ne s'agit pas d'un nombre */
IF ERROR-STATUS:ERROR THEN RETURN ?.
 
/* On construit la date */
ASSIGN ldt_integration = DATETIME(INTEGER(SUBSTRING(lcDateHeure, 3, 2)),            /* Mois */
                                  INTEGER(SUBSTRING(lcDateHeure, 5, 2)),            /* Jour */
                                  2000 + INTEGER(SUBSTRING(lcDateHeure, 1, 2)),     /* Ann‚e */
                                  INTEGER(SUBSTRING(lcDateHeure, 7, 2)),            /* Heure */
                                  INTEGER(SUBSTRING(lcDateHeure, 9, 2))) NO-ERROR.  /* Minutes */
 
IF ERROR-STATUS:ERROR THEN
    RETURN ?.
ELSE
    RETURN ldt_integration.
 
END FUNCTION.
 
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
 
