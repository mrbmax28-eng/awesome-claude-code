Attribute VB_Name = "Module_Ajout_Pieces_Techniques"
Option Explicit

' ================================================================
' AJOUT DES PIECES TECHNIQUES - CREO PARAMETRIC 10
'
' Point d'entree public : Ajout_Pieces_Techniques
'
' Principe :
'   - parcours recursif de l'assemblage actif et de ses sous-assemblages ;
'   - detection de ROLL_CENTER, ROLL2_CENTER, ROLL3_CENTER, etc. ;
'   - lecture de Roll_Diam et Largeur dans Excel, controlees avec Creo ;
'   - creation/reutilisation d'une variante par diametre et largeur ;
'   - placement CSYS sur CSYS dans l'assemblage directement proprietaire ;
'   - choix global Moteur_L/Moteur_R puis placement de l'assemblage moteur ;
'   - squelette du sous-assemblage : lignes de l'esquisse de charpente ;
'   - FOOT_HEIGHT = distance verticale entre le centre du rouleau et la
'     premiere ligne du squelette rencontree en descendant (V15) ;
'   - instances des tables de famille pied et moteur par hauteur ;
'   - mise a jour et suppression differentielles grace a une feuille d'etat.
'
' Important : Roll.prt et ses variantes sont exclus de la recherche afin
' d'eviter l'ajout recursif d'un rouleau sur un autre rouleau.
'
' V14 (2026.10.05.14) - corrections, sans changement du comportement nominal :
'   1. AddRow / AddColumn des tables de famille appeles avec 1 seul argument.
'      Le second argument (Values) est obligatoire dans pfcls : en VBA il
'      faut passer Null. Cause de l'echec de creation des pieds FTF_H.../
'      moteurs FTM_H... (erreur 449 ou 450).
'   2. Recherche des colonnes de famille : par symbole ET par l'element
'      reference (RefItem). Evite d'ajouter une colonne deja presente
'      (XToolkitNoChange) quand la colonne s'appelle FOOT_HEIGHT ou M##.
'   3. Cellule restee "*" (par defaut) apres un essai interrompu : la
'      macro la renseigne au lieu de bloquer jusqu'au redemarrage de Creo.
'   4. Gestionnaires d'erreur : On Error GoTo -1 avant le nettoyage, sinon
'      une erreur de nettoyage masquait l'erreur reelle (plantage VBA brut).
'   5. Noms d'instances compares via InstanceName en plus de FileName.
'   6. Lecture numerique tolerante des cellules (entier ou reel).
'
' V15 (2026.10.05.15) - hauteur des pieds depuis le squelette :
'   Roll_Y n'est plus utilise. Pour chaque rouleau, la macro lit le
'   squelette du sous-assemblage qui le contient, descend verticalement
'   depuis ROLL_CENTER et retient la premiere ligne rencontree.
'   FOOT_HEIGHT = Y(centre) - Y(ligne). Aucune ligne : erreur et arret.
'   Reglages : TECH_AXE_VERTICAL, TECH_AXE_PROFONDEUR, TECH_TOLERANCE_LIGNE,
'   TECH_PIED_DECALAGE_LIGNE (ci-dessous).
'
' V16 (2026.10.05.16) :
'   - tables de famille en liaison precoce (types pfcls) : corrige
'     l'erreur 424 "Objet requis" sur CreateDimensionColumn ;
'   - FOOT_HEIGHT traitee uniquement comme une cote (repli parametre retire) ;
'   - noms FTF_H / FTM_H sans "P" final (ex. FTF_H1380 et non FTF_H1380P).
'
' V17 (2026.10.05.17) :
'   - suppression des pre-controles CheckIsSaveAllowed avant modification
'     des tables de famille (resultat dependant du contexte d'interface
'     Creo selon PTC) ; la sauvegarde est tentee et l'erreur reelle de
'     Creo est remontee avec l'etat d'ecriture du modele ;
'   - nouvelle macro publique Diagnostic_Pieces_Techniques : controle
'     complet de la chaine, sans aucune modification, rapport dans la
'     feuille DIAG_PIECES_TECH.
'
' V18 (2026.10.05.18) - pack pilote HEATING, comportement nominal inchange :
'   - TECH_ExecuterAuto(cote, chemin moteur, message) : meme traitement
'     qu'Ajout_Pieces_Techniques, sans aucune fenetre, avec compte rendu
'     et indicateur de reussite (utilise par M_PILOTE_HEATING) ;
'   - TECH_PUB_EstModeleRoll / TECH_PUB_EstModeleMoteur : regles de
'     reconnaissance des rouleaux et moteurs partagees avec le module
'     paliers, pour que les deux modules voient exactement les memes
'     pieces porteuses ;
'   - TECH_VersionTexte : version lisible par le pilote.
' ================================================================

Private Const TECH_ROLL_FICHIER As String = "Roll.prt"
Private Const TECH_ROLL_CSYS As String = "ROLL_CENTER"
Private Const TECH_ROLL_DIM As String = "Roll_Diam"
Private Const TECH_ROLL_WIDTH As String = "Largeur"
Private Const TECH_VERSION As String = "2026.10.05.18"
Private Const TECH_MOTEUR_FICHIER As String = _
    "ASSEMBLAGE_MOTEUR_CYLINDER.asm"
Private Const TECH_MOTEUR_FICHIER_HISTORIQUE As String = _
    "30002434__SR18_DU01_DRIVE_STA.asm"
Private Const TECH_MOTEUR_CSYS As String = "MOTEUR"
Private Const TECH_MOTEUR_GAUCHE As String = "Moteur_L"
Private Const TECH_MOTEUR_DROIT As String = "Moteur_R"
Private Const TECH_PIED_FICHIER As String = _
    "30002434__SR18_DU01_DRIVE_ST_2.prt"
Private Const TECH_PIED_DIM As String = "FOOT_HEIGHT"
' V15 : axes du repere du sous-assemblage (0 = X, 1 = Y, 2 = Z).
' Verticale = Y positif vers le haut. Profondeur = axe ignore pour croiser
' le rayon descendant avec l'esquisse de charpente (plan de l'esquisse).
' Mettre TECH_AXE_PROFONDEUR = -1 pour exiger un croisement en 3D strict.
Private Const TECH_AXE_VERTICAL As Long = 1
Private Const TECH_AXE_PROFONDEUR As Long = 2
Private Const TECH_TOLERANCE_LIGNE As Double = 0.5
Private Const TECH_PIED_DECALAGE_LIGNE As Double = 0#
Private Const TECH_FEUILLE_ETAT As String = "__HUB_ETAT"
Private Const TECH_PREFIXE_VARIANTE As String = "ROLL_D"
Private Const TECH_PREFIXE_PIED As String = "FTF_H"
Private Const TECH_PREFIXE_PIED_COPIE As String = "MCP_H"
Private Const TECH_PREFIXE_MOTEUR As String = "FTM_H"
Private Const TECH_PREFIXE_MOTEUR_COPIE As String = "MCA_H"
Private Const TECH_PREFIXE_PIED_HISTORIQUE As String = "AMF_H"
Private Const TECH_PREFIXE_MOTEUR_HISTORIQUE As String = "AMC_H"

' Colonnes T a AM de __HUB_ETAT. Les colonnes A a R restent reservees au HUB.
Private Const TECH_COL_ASSEMBLAGE As Long = 20
Private Const TECH_COL_CLE As Long = 21
Private Const TECH_COL_SOURCE_ID As Long = 22
Private Const TECH_COL_SOURCE_MODELE As Long = 23
Private Const TECH_COL_SOURCE_PIECE As Long = 24
Private Const TECH_COL_OCCURRENCE As Long = 25
Private Const TECH_COL_REPERE As Long = 26
Private Const TECH_COL_DIAMETRE As Long = 27
Private Const TECH_COL_ROLL_MODELE As Long = 28
Private Const TECH_COL_ROLL_ID As Long = 29
Private Const TECH_COL_STATUT As Long = 30
Private Const TECH_COL_DATE As Long = 31
Private Const TECH_COL_LARGEUR As Long = 32
Private Const TECH_COL_MOTEUR_COTE As Long = 33
Private Const TECH_COL_MOTEUR_MODELE As Long = 34
Private Const TECH_COL_MOTEUR_ID As Long = 35
Private Const TECH_COL_ROLL_HAUTEUR As Long = 36
Private Const TECH_COL_FOOT_HEIGHT As Long = 37
Private Const TECH_COL_PIED_MODELE As Long = 38
Private Const TECH_COL_MOTEUR_SOURCE As Long = 39

Private m_TECH_Etape As String
Private m_TECH_MoteurFichier As String
Private m_TECH_MoteurDossier As String
Private m_TECH_LecteurReseauTemp As String
Private m_TECH_PartageReseauTemp As String
Private m_TECH_DossierCreoInitial As String

Public Sub Version_Pieces_Techniques()
    MsgBox "Module_Ajout_Pieces_Techniques" & vbCrLf & _
        "Version : " & TECH_VERSION, vbInformation, _
        "Version Pieces techniques"
End Sub

Public Sub Ajout_Pieces_Techniques()
    Dim coteMoteur As String
    Dim message As String

    coteMoteur = TECH_ChoisirCoteMoteur()
    If coteMoteur = "" Then Exit Sub
    TECH_Executer coteMoteur, "", False, message
End Sub

' V18 : execution sans fenetre pour le pilote. coteMoteur = Moteur_L ou
' Moteur_R ; cheminMoteur = chemin complet de l'assemblage moteur-cylindre.
Public Function TECH_ExecuterAuto(ByVal coteMoteur As String, _
    ByVal cheminMoteur As String, ByRef message As String) As Boolean

    message = ""
    If StrComp(coteMoteur, TECH_MOTEUR_GAUCHE, vbTextCompare) = 0 Then
        coteMoteur = TECH_MOTEUR_GAUCHE
    ElseIf StrComp(coteMoteur, TECH_MOTEUR_DROIT, vbTextCompare) = 0 Then
        coteMoteur = TECH_MOTEUR_DROIT
    Else
        message = "Cote moteur invalide : '" & coteMoteur & "' (" & _
            TECH_MOTEUR_GAUCHE & " ou " & TECH_MOTEUR_DROIT & ")."
        Exit Function
    End If
    If Len(Trim$(cheminMoteur)) = 0 Then
        message = "Chemin de l'assemblage moteur non renseigne."
        Exit Function
    End If
    If Len(Dir$(cheminMoteur)) = 0 Then
        message = "Assemblage moteur introuvable : " & cheminMoteur
        Exit Function
    End If
    TECH_ExecuterAuto = TECH_Executer(coteMoteur, cheminMoteur, True, message)
End Function

Public Function TECH_VersionTexte() As String
    TECH_VersionTexte = TECH_VERSION
End Function

' V18 : regles partagees avec M_FRAME_HEATING_PALIERS.
Public Function TECH_PUB_EstModeleRoll(ByVal nomModele As String) As Boolean
    TECH_PUB_EstModeleRoll = TECH_EstModeleRoll(nomModele)
End Function

Public Function TECH_PUB_EstModeleMoteur(ByVal nomModele As String, _
    Optional ByVal fichierMoteur As String = "") As Boolean
    Dim ancien As String
    ancien = m_TECH_MoteurFichier
    If Len(fichierMoteur) > 0 Then
        m_TECH_MoteurFichier = TECH_NomModeleLogique(fichierMoteur)
    End If
    TECH_PUB_EstModeleMoteur = TECH_EstModeleMoteur(nomModele)
    m_TECH_MoteurFichier = ancien
End Function

Private Function TECH_Executer(ByVal coteMoteur As String, _
    ByVal cheminMoteur As String, ByVal silencieux As Boolean, _
    ByRef message As String) As Boolean
    Dim cConnexion As pfcls.CCpfcAsyncConnection
    Dim connexion As pfcls.IpfcAsyncConnection
    Dim session As pfcls.IpfcBaseSession
    Dim modeleRacine As pfcls.IpfcModel
    Dim asmRacine As pfcls.IpfcAssembly
    Dim modeleRoll As pfcls.IpfcModel
    Dim solidRoll As pfcls.IpfcSolid
    Dim modeleMoteur As pfcls.IpfcModel
    Dim solidMoteur As pfcls.IpfcSolid
    Dim wsEtat As Worksheet
    Dim visites As Object
    Dim cacheVariantes As Object
    Dim cacheVariantesMoteur As Object
    Dim dossierPieces As String
    Dim connexionCreee As Boolean
    Dim nbAssemblages As Long
    Dim nbAjoutes As Long
    Dim nbMaj As Long
    Dim nbSupprimes As Long
    Dim nbInchanges As Long
    Dim controleItem As pfcls.IpfcModelItem
    Dim controleDimension As pfcls.IpfcBaseDimension
    Dim nbMoteursAjoutes As Long
    Dim nbMoteursDeplaces As Long
    Dim nbMoteursInchanges As Long
    Dim errNumero As Long
    Dim errDescription As String
    Dim errSource As String
    Dim ancienScreenUpdating As Boolean
    Dim ancienEnableEvents As Boolean
    Dim ancienDisplayAlerts As Boolean

    On Error GoTo Erreur

    Debug.Print "AJOUT PIECES TECHNIQUES - VERSION " & TECH_VERSION

    message = ""
    m_TECH_MoteurFichier = TECH_MOTEUR_FICHIER
    m_TECH_MoteurDossier = ""
    m_TECH_LecteurReseauTemp = ""
    m_TECH_PartageReseauTemp = ""
    m_TECH_DossierCreoInitial = ""

    ancienScreenUpdating = Application.ScreenUpdating
    ancienEnableEvents = Application.EnableEvents
    ancienDisplayAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    m_TECH_Etape = "Connexion a Creo"
    TECH_ConnecterCreo cConnexion, connexion, session, connexionCreee
    m_TECH_DossierCreoInitial = session.GetCurrentDirectory()

    m_TECH_Etape = "Controle de l'assemblage actif"
    Set modeleRacine = session.GetActiveModel()
    If modeleRacine Is Nothing Then
        Err.Raise vbObjectError + 15000, , "Aucun modele actif dans Creo."
    End If
    If modeleRacine.Type <> pfcls.EpfcMDL_ASSEMBLY Then
        Err.Raise vbObjectError + 15001, , _
            "Le modele actif doit etre l'assemblage principal."
    End If
    Set asmRacine = modeleRacine
    If asmRacine Is Nothing Then
        Err.Raise vbObjectError + 15002, , _
            "Le modele actif n'est pas exploitable comme assemblage."
    End If

    m_TECH_Etape = "Resolution du dossier de travail"
    dossierPieces = TECH_ObtenirDossierPieces(session)
    m_TECH_Etape = "Selection de l'assemblage moteur-cylindre"
    TECH_ResoudreFichierMoteur dossierPieces, m_TECH_MoteurFichier, _
        m_TECH_MoteurDossier, cheminMoteur

    m_TECH_Etape = "Chargement et controle de Roll.prt"
    Set modeleRoll = TECH_ChargerModele(session, TECH_ROLL_FICHIER, dossierPieces)
    If modeleRoll Is Nothing Then
        Err.Raise vbObjectError + 15003, , _
            TECH_ROLL_FICHIER & " est introuvable dans le dossier de travail :" & _
            vbCrLf & dossierPieces
    End If
    Set solidRoll = modeleRoll
    If solidRoll Is Nothing Then
        Err.Raise vbObjectError + 15004, , TECH_ROLL_FICHIER & " n'est pas une piece."
    End If
    Set controleItem = TECH_TrouverCsysExact(modeleRoll, TECH_ROLL_CSYS)
    If controleItem Is Nothing Then
        Err.Raise vbObjectError + 15005, , _
            "Le repere " & TECH_ROLL_CSYS & " est absent de " & TECH_ROLL_FICHIER & "."
    End If
    Set controleDimension = TECH_TrouverDimension(solidRoll, TECH_ROLL_DIM)
    If controleDimension Is Nothing Then
        Err.Raise vbObjectError + 15006, , _
            "La cote " & TECH_ROLL_DIM & " est absente de " & TECH_ROLL_FICHIER & "."
    End If
    Set controleDimension = TECH_TrouverDimension(solidRoll, TECH_ROLL_WIDTH)
    If controleDimension Is Nothing Then
        Err.Raise vbObjectError + 15007, , _
            "La cote " & TECH_ROLL_WIDTH & " est absente de " & TECH_ROLL_FICHIER & "."
    End If
    TECH_ValiderSolid solidRoll, "Controle de " & TECH_ROLL_FICHIER

    Set wsEtat = TECH_ObtenirFeuilleEtat(True)
    Set visites = CreateObject("Scripting.Dictionary")
    visites.CompareMode = vbTextCompare
    Set cacheVariantes = CreateObject("Scripting.Dictionary")
    cacheVariantes.CompareMode = vbTextCompare
    Set cacheVariantesMoteur = CreateObject("Scripting.Dictionary")
    cacheVariantesMoteur.CompareMode = vbTextCompare

    m_TECH_Etape = "Parcours de l'assemblage"
    TECH_TraiterAssemblage session, modeleRacine, modeleRoll, dossierPieces, _
        wsEtat, visites, cacheVariantes, nbAssemblages, nbAjoutes, nbMaj, _
        nbSupprimes, nbInchanges

    m_TECH_Etape = "Chargement et controle de l'assemblage moteur"
    Set modeleMoteur = TECH_ChargerModele(session, m_TECH_MoteurFichier, _
        m_TECH_MoteurDossier)
    If modeleMoteur Is Nothing Then
        Err.Raise vbObjectError + 15008, , _
            m_TECH_MoteurFichier & " est introuvable dans le dossier choisi."
    End If
    If modeleMoteur.Type <> pfcls.EpfcMDL_ASSEMBLY Then
        Err.Raise vbObjectError + 15009, , _
            m_TECH_MoteurFichier & " n'est pas un assemblage Creo."
    End If
    Set solidMoteur = modeleMoteur
    If solidMoteur Is Nothing Then
        Err.Raise vbObjectError + 15010, , _
            m_TECH_MoteurFichier & " n'est pas exploitable comme solide."
    End If
    Set controleItem = TECH_TrouverCsysExact(modeleMoteur, TECH_MOTEUR_CSYS)
    If controleItem Is Nothing Then
        Err.Raise vbObjectError + 15011, , _
            "Le repere " & TECH_MOTEUR_CSYS & " est absent de " & _
            m_TECH_MoteurFichier & "."
    End If
    TECH_ValiderSolid solidMoteur, "Controle de " & m_TECH_MoteurFichier
    TECH_ValiderModeleMoteurSource session, modeleMoteur

    m_TECH_Etape = "Placement des moteurs sur " & coteMoteur
    visites.RemoveAll
    TECH_TraiterMoteurs session, modeleRacine, modeleMoteur, _
        m_TECH_MoteurDossier, cacheVariantesMoteur, coteMoteur, wsEtat, _
        visites, True, nbMoteursAjoutes, nbMoteursDeplaces, _
        nbMoteursInchanges
    visites.RemoveAll
    TECH_TraiterMoteurs session, modeleRacine, modeleMoteur, _
        m_TECH_MoteurDossier, cacheVariantesMoteur, coteMoteur, wsEtat, _
        visites, False, nbMoteursAjoutes, nbMoteursDeplaces, _
        nbMoteursInchanges

    m_TECH_Etape = "Retour a l'assemblage principal"
    TECH_ActiverModele session, modeleRacine

    m_TECH_Etape = "Enregistrement de l'etat Excel"
    wsEtat.Visible = xlSheetVeryHidden
    ThisWorkbook.Save

    Application.DisplayAlerts = ancienDisplayAlerts
    Application.EnableEvents = ancienEnableEvents
    Application.ScreenUpdating = ancienScreenUpdating

    On Error Resume Next
    TECH_LibererLecteurReseauTemp session
    If connexionCreee Then connexion.Disconnect 5
    On Error GoTo 0

    message = "Assemblages analyses : " & CStr(nbAssemblages) & vbCrLf & _
        "Rouleaux ajoutes : " & CStr(nbAjoutes) & vbCrLf & _
        "Rouleaux mis a jour : " & CStr(nbMaj) & vbCrLf & _
        "Rouleaux supprimes : " & CStr(nbSupprimes) & vbCrLf & _
        "Rouleaux inchanges : " & CStr(nbInchanges) & vbCrLf & vbCrLf & _
        "Moteurs ajoutes : " & CStr(nbMoteursAjoutes) & vbCrLf & _
        "Moteurs deplaces : " & CStr(nbMoteursDeplaces) & vbCrLf & _
        "Moteurs inchanges : " & CStr(nbMoteursInchanges)
    TECH_Executer = True
    If Not silencieux Then MsgBox "Ajout des pieces techniques termine." & vbCrLf & _
        "Version macro : " & TECH_VERSION & vbCrLf & vbCrLf & message, _
        vbInformation, "Ajout Pieces techniques"
    Exit Function

Erreur:
    errNumero = Err.Number
    errDescription = Err.description
    errSource = Err.source

    Application.DisplayAlerts = ancienDisplayAlerts
    Application.EnableEvents = ancienEnableEvents
    Application.ScreenUpdating = ancienScreenUpdating

    On Error GoTo -1 ' V14 : libere le gestionnaire actif avant nettoyage
    On Error Resume Next
    If Not session Is Nothing And Not modeleRacine Is Nothing Then _
        TECH_ActiverModele session, modeleRacine
    TECH_LibererLecteurReseauTemp session
    If connexionCreee Then connexion.Disconnect 5
    On Error GoTo 0

    message = "Etape : " & m_TECH_Etape & vbCrLf & _
        "Erreur : " & CStr(errNumero) & vbCrLf & _
        "Source : " & errSource & vbCrLf & errDescription
    If Not silencieux Then MsgBox "ERREUR pendant l'ajout des pieces techniques." & _
        vbCrLf & "Version macro : " & TECH_VERSION & vbCrLf & vbCrLf & message, _
        vbCritical, "Ajout Pieces techniques"
End Function

Private Function TECH_ChoisirCoteMoteur() As String
    Dim choix As VbMsgBoxResult

    choix = MsgBox( _
        "Choisissez le cote pour TOUS les moteurs :" & vbCrLf & vbCrLf & _
        "OUI  = cote gauche (Moteur_L)" & vbCrLf & _
        "NON = cote droit (Moteur_R)" & vbCrLf & _
        "ANNULER = quitter sans modification", _
        vbYesNoCancel + vbQuestion + vbDefaultButton1, _
        "Ajout Pieces techniques - Cote moteur")

    If choix = vbYes Then
        TECH_ChoisirCoteMoteur = TECH_MOTEUR_GAUCHE
    ElseIf choix = vbNo Then
        TECH_ChoisirCoteMoteur = TECH_MOTEUR_DROIT
    End If
End Function

Private Sub TECH_ResoudreFichierMoteur( _
    ByVal dossierParDefaut As String, _
    ByRef nomFichier As String, _
    ByRef dossierMoteur As String, _
    Optional ByVal cheminImpose As String = "")

    Dim dossierNormalise As String
    Dim fichierTrouve As String
    Dim dialogue As Object
    Dim cheminChoisi As String
    Dim positionSeparateur As Long
    Dim nomPhysique As String

    dossierNormalise = Replace(Trim$(dossierParDefaut), "/", "\")
    If dossierNormalise <> "" Then
        If Right$(dossierNormalise, 1) <> "\" Then _
            dossierNormalise = dossierNormalise & "\"
    End If

    If Len(Trim$(cheminImpose)) > 0 Then
        ' V18 : chemin fourni par le pilote, sans fenetre de selection.
        cheminChoisi = Replace(Trim$(cheminImpose), "/", "\")
        If Len(Dir$(cheminChoisi)) = 0 Then
            Err.Raise vbObjectError + 15015, , _
                "Assemblage moteur introuvable : " & cheminChoisi
        End If
        GoTo CheminConnu
    End If

    fichierTrouve = Dir$(dossierNormalise & _
        TECH_NomModeleLogique(nomFichier) & "*")

    Set dialogue = Application.FileDialog(3)
    With dialogue
        .title = "Selectionner ASSEMBLAGE_MOTEUR_CYLINDER.ASM"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Assemblages Creo", "*.asm*"
        If fichierTrouve <> "" Then
            .InitialFileName = dossierNormalise & fichierTrouve
        Else
            .InitialFileName = dossierNormalise
        End If
        If .Show <> -1 Then
            Err.Raise vbObjectError + 15012, , _
                "Selection de l'assemblage moteur annulee."
        End If
        cheminChoisi = CStr(.SelectedItems(1))
    End With

CheminConnu:

    positionSeparateur = InStrRev(cheminChoisi, "\", -1, vbTextCompare)
    If positionSeparateur <= 0 Then
        Err.Raise vbObjectError + 15013, , _
            "Chemin moteur invalide : " & cheminChoisi
    End If
    dossierMoteur = Left$(cheminChoisi, positionSeparateur)
    nomPhysique = Mid$(cheminChoisi, positionSeparateur + 1)
    nomFichier = TECH_NomModeleLogique(nomPhysique)
    If InStrRev(nomFichier, ".asm", -1, vbTextCompare) <= 0 Then
        Err.Raise vbObjectError + 15014, , _
            "Le fichier choisi n'est pas un assemblage Creo : " & nomPhysique
    End If
End Sub

Private Sub TECH_ConnecterCreo( _
    ByRef cConnexion As pfcls.CCpfcAsyncConnection, _
    ByRef connexion As pfcls.IpfcAsyncConnection, _
    ByRef session As pfcls.IpfcBaseSession, _
    ByRef connexionCreee As Boolean)

    Set cConnexion = New pfcls.CCpfcAsyncConnection
    connexionCreee = False

    On Error Resume Next
    Set connexion = cConnexion.GetActiveConnection()
    Err.Clear
    On Error GoTo Erreur

    If connexion Is Nothing Then
        Set connexion = cConnexion.Connect("", "", "", 30)
        connexionCreee = True
    End If
    If connexion Is Nothing Then
        Err.Raise vbObjectError + 15010, , _
            "Connexion a Creo impossible. Verifie que Creo est ouvert."
    End If

    Set session = connexion.session
    If session Is Nothing Then
        Err.Raise vbObjectError + 15011, , "La session Creo est introuvable."
    End If
    Exit Sub

Erreur:
    Err.Raise Err.Number, , "TECH_ConnecterCreo :" & vbCrLf & Err.description
End Sub

Private Sub TECH_TraiterAssemblage( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal modeleAssemblage As pfcls.IpfcModel, _
    ByVal modeleRollSource As pfcls.IpfcModel, _
    ByVal dossierPieces As String, _
    ByVal wsEtat As Worksheet, _
    ByVal visites As Object, _
    ByVal cacheVariantes As Object, _
    ByRef nbAssemblages As Long, _
    ByRef nbAjoutes As Long, _
    ByRef nbMaj As Long, _
    ByRef nbSupprimes As Long, _
    ByRef nbInchanges As Long)

    Dim nomAssemblage As String
    Dim asm As pfcls.IpfcAssembly
    Dim solidAsm As pfcls.IpfcSolid
    Dim features As pfcls.IpfcFeatures
    Dim feat As pfcls.IpfcFeature
    Dim featItem As pfcls.IpfcModelItem
    Dim composant As pfcls.IpfcComponentFeat
    Dim modeleComposant As pfcls.IpfcModel
    Dim solidComposant As pfcls.IpfcSolid
    Dim sousAssemblages As Collection
    Dim reperes As Collection
    Dim repere As pfcls.IpfcModelItem
    Dim souhaites As Object
    Dim informations As Variant
    Dim sourceID As Long
    Dim nomRepere As String
    Dim nomPieceSource As String
    Dim occurrence As Long
    Dim diametre As Double
    Dim largeur As Double
    Dim cle As String
    Dim i As Long
    Dim numeroRepere As Long

    If modeleAssemblage Is Nothing Then Exit Sub
    nomAssemblage = TECH_NomModeleLogique(modeleAssemblage.Filename)
    If visites.Exists(nomAssemblage) Then Exit Sub
    visites.Add nomAssemblage, True

    Set asm = modeleAssemblage
    Set solidAsm = asm
    If asm Is Nothing Or solidAsm Is Nothing Then
        Err.Raise vbObjectError + 15020, , _
            "Assemblage inexploitable : " & nomAssemblage
    End If

    nbAssemblages = nbAssemblages + 1
    m_TECH_Etape = "Analyse de " & nomAssemblage
    Set sousAssemblages = New Collection
    Set souhaites = CreateObject("Scripting.Dictionary")
    souhaites.CompareMode = vbTextCompare

    Set features = solidAsm.ListFeaturesByType(False, pfcls.EpfcFEATTYPE_COMPONENT)
    If Not features Is Nothing Then
        For i = 0 To features.Count - 1
            Set feat = Nothing
            Set featItem = Nothing
            Set composant = Nothing
            Set modeleComposant = Nothing
            Set solidComposant = Nothing

            Set feat = features.item(i)
            If Not feat Is Nothing Then Set featItem = feat
            If Not feat Is Nothing Then Set composant = feat

            If Not composant Is Nothing Then
                Set modeleComposant = TECH_ModeleDepuisComposant(session, composant)
                If modeleComposant Is Nothing Then
                    Err.Raise vbObjectError + 15021, , _
                        "Impossible de charger un composant actif de " & nomAssemblage & _
                        ". Feature ID : " & CStr(featItem.id)
                End If

                If modeleComposant.Type = pfcls.EpfcMDL_ASSEMBLY Then
                    If Not TECH_EstModeleMoteur(modeleComposant.Filename) Then
                        sousAssemblages.Add modeleComposant
                    End If
                ElseIf Not TECH_EstModeleRoll(modeleComposant.Filename) Then
                    Set solidComposant = modeleComposant
                    If Not solidComposant Is Nothing Then
                        Set reperes = TECH_ListerReperesRoll(modeleComposant)
                        If Not reperes Is Nothing Then
                            If reperes.Count > 0 Then
                                sourceID = featItem.id
                                diametre = TECH_LireDiametreSource( _
                                    modeleComposant, solidComposant, nomAssemblage, _
                                    sourceID, nomPieceSource, occurrence)
                                largeur = TECH_LireLargeurSource( _
                                    modeleComposant, solidComposant, _
                                    nomPieceSource, occurrence)

                                For numeroRepere = 1 To reperes.Count
                                    Set repere = reperes.item(numeroRepere)
                                    nomRepere = UCase$(Trim$(repere.GetName()))
                                    cle = TECH_CleCible(nomAssemblage, sourceID, nomRepere)
                                    informations = Array(sourceID, modeleComposant.Filename, _
                                        nomPieceSource, occurrence, nomRepere, diametre, largeur)
                                    souhaites(cle) = informations
                                Next numeroRepere
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    End If

    TECH_SynchroniserAssemblage session, modeleAssemblage, modeleRollSource, _
        dossierPieces, wsEtat, cacheVariantes, souhaites, nbAjoutes, nbMaj, _
        nbSupprimes, nbInchanges

    For i = 1 To sousAssemblages.Count
        TECH_TraiterAssemblage session, sousAssemblages.item(i), modeleRollSource, _
            dossierPieces, wsEtat, visites, cacheVariantes, nbAssemblages, _
            nbAjoutes, nbMaj, nbSupprimes, nbInchanges
    Next i
End Sub

Private Sub TECH_SynchroniserAssemblage( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal modeleAssemblage As pfcls.IpfcModel, _
    ByVal modeleRollSource As pfcls.IpfcModel, _
    ByVal dossierPieces As String, _
    ByVal wsEtat As Worksheet, _
    ByVal cacheVariantes As Object, _
    ByVal souhaites As Object, _
    ByRef nbAjoutes As Long, _
    ByRef nbMaj As Long, _
    ByRef nbSupprimes As Long, _
    ByRef nbInchanges As Long)

    Dim nomAssemblage As String
    Dim asm As pfcls.IpfcAssembly
    Dim solidAsm As pfcls.IpfcSolid
    Dim modeleVariante As pfcls.IpfcModel
    Dim solidVariante As pfcls.IpfcSolid
    Dim modeleActuel As pfcls.IpfcModel
    Dim cle As Variant
    Dim informations As Variant
    Dim ligneEtat As Long
    Dim ancienID As Long
    Dim nouveauID As Long
    Dim nomVariante As String
    Dim ancienExiste As Boolean
    Dim modifie As Boolean
    Dim lignesASupprimer As Collection
    Dim r As Long
    Dim derniereLigne As Long
    Dim i As Long
    Dim ligneTemp As Long
    Dim numeroRemplacement As Long
    Dim descriptionRemplacement As String

    nomAssemblage = TECH_NomModeleLogique(modeleAssemblage.Filename)
    Set asm = modeleAssemblage
    Set solidAsm = asm
    Set lignesASupprimer = New Collection

    For Each cle In souhaites.Keys
        informations = souhaites(cle)
        nomVariante = TECH_NomVarianteRoll(CDbl(informations(5)), _
            CDbl(informations(6))) & ".prt"
        ligneEtat = TECH_TrouverLigneEtat(wsEtat, CStr(cle))
        ancienID = 0
        ancienExiste = False
        Set modeleActuel = Nothing

        If ligneEtat > 0 Then
            ancienID = CLng(Val(CStr(wsEtat.Cells(ligneEtat, TECH_COL_ROLL_ID).value)))
            If ancienID > 0 And _
                TECH_FeatureComposantActive(solidAsm, ancienID) Then
                Set modeleActuel = TECH_ModeleComposantParID(session, solidAsm, ancienID)
                ancienExiste = Not modeleActuel Is Nothing
            End If
        End If

        If ancienExiste Then
            If TECH_RollConforme(modeleActuel, nomVariante, _
                CDbl(informations(5)), CDbl(informations(6))) And _
                TECH_OccurrenceRollPlacee(solidAsm, ancienID) And _
                TECH_OccurrenceRollCiblee( _
                    solidAsm, ancienID, CLng(informations(0))) Then
                nbInchanges = nbInchanges + 1
                GoTo CibleSuivante
            End If
        End If

        m_TECH_Etape = "Preparation de " & nomVariante
        Set modeleVariante = TECH_ObtenirVarianteRoll(session, modeleRollSource, _
            dossierPieces, CDbl(informations(5)), CDbl(informations(6)), _
            cacheVariantes)
        Set solidVariante = modeleVariante
        If solidVariante Is Nothing Then
            Err.Raise vbObjectError + 15030, , _
                "La variante " & nomVariante & " n'est pas une piece."
        End If

        m_TECH_Etape = "Ajout sur " & CStr(informations(4)) & _
            " dans " & nomAssemblage
        nouveauID = TECH_AjouterRoll(session, asm, solidAsm, solidVariante, _
            CLng(informations(0)), CStr(informations(4)))

        If ancienExiste Then
            On Error GoTo RemplacementEchoue
            TECH_SupprimerMoteurEtat session, wsEtat, ligneEtat, asm, solidAsm
            TECH_SupprimerRollSuivi session, asm, solidAsm, ancienID
            On Error GoTo 0
            nbMaj = nbMaj + 1
        Else
            If ligneEtat > 0 Then
                TECH_SupprimerMoteurEtat session, wsEtat, ligneEtat, asm, solidAsm
            End If
            nbAjoutes = nbAjoutes + 1
        End If

        TECH_EcrireEtat wsEtat, ligneEtat, nomAssemblage, CStr(cle), _
            informations, modeleVariante.Filename, nouveauID
        modifie = True
        GoTo CibleSuivante

RemplacementEchoue:
        numeroRemplacement = Err.Number
        descriptionRemplacement = Err.description
        On Error GoTo -1 ' V14 : libere le gestionnaire actif avant nettoyage
        On Error Resume Next
        TECH_SupprimerComposant session, asm, solidAsm, nouveauID
        On Error GoTo 0
        Err.Raise numeroRemplacement, , _
            "Mise a jour du rouleau impossible. L'ancien rouleau a ete conserve." & _
            vbCrLf & descriptionRemplacement

CibleSuivante:
    Next cle

    derniereLigne = wsEtat.Cells(wsEtat.Rows.Count, TECH_COL_CLE).End(xlUp).Row
    For r = 2 To derniereLigne
        If StrComp(Trim$(CStr(wsEtat.Cells(r, TECH_COL_ASSEMBLAGE).value)), _
            nomAssemblage, vbTextCompare) = 0 Then

            cle = Trim$(CStr(wsEtat.Cells(r, TECH_COL_CLE).value))
            If Not souhaites.Exists(CStr(cle)) Then
                ancienID = CLng(Val(CStr(wsEtat.Cells(r, TECH_COL_ROLL_ID).value)))
                If ancienID > 0 Then
                    If TECH_FeatureComposantActive(solidAsm, ancienID) Then
                        Set modeleActuel = TECH_ModeleComposantParID(session, solidAsm, ancienID)
                        If Not modeleActuel Is Nothing Then
                            m_TECH_Etape = "Suppression d'un rouleau devenu inutile dans " & _
                                nomAssemblage
                            If TECH_EstModeleRoll(modeleActuel.Filename) Then
                                TECH_SupprimerMoteurEtat session, wsEtat, r, asm, solidAsm
                                TECH_SupprimerRollSuivi session, asm, solidAsm, ancienID
                                nbSupprimes = nbSupprimes + 1
                                modifie = True
                            Else
                                Debug.Print "ETAT ROULEAU PERIME : ID " & CStr(ancienID) & _
                                    " pointe vers " & modeleActuel.Filename & _
                                    " -> composant conserve, ligne d'etat oubliee."
                            End If
                        End If
                    Else
                        Debug.Print "ETAT ROULEAU PERIME : ID " & CStr(ancienID) & _
                            " n'est pas un composant actif de " & nomAssemblage & _
                            " -> aucune suppression Creo."
                    End If
                End If
                lignesASupprimer.Add r
            End If
        End If
    Next r

    If modifie Then
        m_TECH_Etape = "Validation de " & nomAssemblage
        TECH_ValiderSolid solidAsm, "Assemblage " & nomAssemblage
        If Not modeleAssemblage.CheckIsSaveAllowed(False) Then
            Err.Raise vbObjectError + 15031, , _
                "Creo refuse la sauvegarde de " & nomAssemblage & "."
        End If
        modeleAssemblage.Save
    End If

    For i = lignesASupprimer.Count To 1 Step -1
        ligneTemp = CLng(lignesASupprimer.item(i))
        wsEtat.Range(wsEtat.Cells(ligneTemp, TECH_COL_ASSEMBLAGE), _
            wsEtat.Cells(ligneTemp, TECH_COL_MOTEUR_SOURCE)).ClearContents
    Next i
End Sub

Private Sub TECH_TraiterMoteurs( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal modeleAssemblage As pfcls.IpfcModel, _
    ByVal modeleMoteurSource As pfcls.IpfcModel, _
    ByVal dossierMoteur As String, _
    ByVal cacheVariantesMoteur As Object, _
    ByVal coteMoteur As String, _
    ByVal wsEtat As Worksheet, _
    ByVal visites As Object, _
    ByVal validerSeulement As Boolean, _
    ByRef nbAjoutes As Long, _
    ByRef nbDeplaces As Long, _
    ByRef nbInchanges As Long)

    Dim nomAssemblage As String
    Dim asm As pfcls.IpfcAssembly
    Dim solidAsm As pfcls.IpfcSolid
    Dim features As pfcls.IpfcFeatures
    Dim feature As pfcls.IpfcFeature
    Dim item As pfcls.IpfcModelItem
    Dim composant As pfcls.IpfcComponentFeat
    Dim modeleComposant As pfcls.IpfcModel
    Dim sousAssemblages As Collection
    Dim rollIDs As Collection
    Dim ligneEtat As Long
    Dim modeleRollControle As pfcls.IpfcModel
    Dim repereControle As pfcls.IpfcModelItem
    Dim repereGauche As pfcls.IpfcModelItem
    Dim repereDroit As pfcls.IpfcModelItem
    Dim modeleMoteurVariante As pfcls.IpfcModel
    Dim solidMoteurVariante As pfcls.IpfcSolid
    Dim hauteurRoll As Double
    Dim footHeight As Double
    Dim nomPiedVariante As String
    Dim modifie As Boolean
    Dim i As Long
    Dim segmentsSquelette As Collection
    Dim nomSquelette As String
    Dim centreRoll As Variant
    Dim ligneY As Double
    Dim nomRepereEtat As String

    If modeleAssemblage Is Nothing Then Exit Sub
    nomAssemblage = TECH_NomModeleLogique(modeleAssemblage.Filename)
    If visites.Exists(nomAssemblage) Then Exit Sub
    visites.Add nomAssemblage, True

    Set asm = modeleAssemblage
    Set solidAsm = asm
    If asm Is Nothing Or solidAsm Is Nothing Then
        Err.Raise vbObjectError + 15140, , _
            "Assemblage inexploitable pour les moteurs : " & nomAssemblage
    End If

    Set sousAssemblages = New Collection
    Set rollIDs = New Collection
    Set features = solidAsm.ListFeaturesByType(False, pfcls.EpfcFEATTYPE_COMPONENT)
    If Not features Is Nothing Then
        For i = 0 To features.Count - 1
            Set feature = features.item(i)
            Set item = Nothing
            Set composant = Nothing
            Set modeleComposant = Nothing
            If Not feature Is Nothing Then Set item = feature
            If Not feature Is Nothing Then Set composant = feature
            If Not composant Is Nothing Then
                Set modeleComposant = TECH_ModeleDepuisComposant(session, composant)
                If Not modeleComposant Is Nothing Then
                    If TECH_EstModeleRoll(modeleComposant.Filename) Then
                        rollIDs.Add CLng(item.id)
                    ElseIf modeleComposant.Type = pfcls.EpfcMDL_ASSEMBLY Then
                        If Not TECH_EstModeleMoteur(modeleComposant.Filename) Then
                            sousAssemblages.Add modeleComposant
                        End If
                    End If
                End If
            End If
        Next i
    End If

    For i = 1 To rollIDs.Count
        ligneEtat = TECH_TrouverLigneEtatRoll(wsEtat, nomAssemblage, _
            CLng(rollIDs.item(i)))
        If ligneEtat > 0 Then
            Set modeleRollControle = TECH_ModeleComposantParID(session, solidAsm, _
                CLng(rollIDs.item(i)))
            Set repereControle = Nothing
            Set repereGauche = Nothing
            Set repereDroit = Nothing
            If Not modeleRollControle Is Nothing Then
                Set repereControle = TECH_TrouverCsysExact(modeleRollControle, coteMoteur)
                Set repereGauche = TECH_TrouverCsysExact( _
                    modeleRollControle, TECH_MOTEUR_GAUCHE)
                Set repereDroit = TECH_TrouverCsysExact( _
                    modeleRollControle, TECH_MOTEUR_DROIT)
            End If
            If repereGauche Is Nothing Or repereDroit Is Nothing Then
                Err.Raise vbObjectError + 15142, , _
                    "Le rouleau ID " & CStr(rollIDs.item(i)) & _
                    " doit contenir les deux reperes moteur." & vbCrLf & _
                    "Moteur_L : " & IIf(repereGauche Is Nothing, _
                        "ABSENT", "OK") & vbCrLf & _
                    "Moteur_R : " & IIf(repereDroit Is Nothing, _
                        "ABSENT", "OK") & vbCrLf & _
                    "Assemblage : " & nomAssemblage
            ElseIf repereControle Is Nothing Then
                Err.Raise vbObjectError + 15144, , _
                    "Le repere choisi " & coteMoteur & _
                    " est inaccessible sur le rouleau ID " & _
                    CStr(rollIDs.item(i)) & "."
            End If

            ' V15 : hauteur du pied mesuree sur le squelette.
            nomRepereEtat = Trim$(CStr(wsEtat.Cells(ligneEtat, TECH_COL_REPERE).value))
            m_TECH_Etape = "Squelette de " & nomAssemblage & " / " & nomRepereEtat
            If segmentsSquelette Is Nothing Then
                Set segmentsSquelette = TECH_ListerLignesSquelette( _
                    session, asm, solidAsm, nomAssemblage, nomSquelette)
            End If
            centreRoll = TECH_CentreRouleau(asm, CLng(rollIDs.item(i)), _
                modeleRollControle)
            hauteurRoll = CDbl(centreRoll(TECH_AXE_VERTICAL))
            If Not TECH_TrouverLigneSousCentre(centreRoll, segmentsSquelette, _
                ligneY) Then
                Err.Raise vbObjectError + 15305, , _
                    "Aucune ligne du squelette sous le centre du rouleau." & _
                    vbCrLf & "Assemblage : " & nomAssemblage & _
                    vbCrLf & "Squelette : " & nomSquelette & _
                    vbCrLf & "Repere : " & nomRepereEtat & _
                    " / rouleau ID " & CStr(rollIDs.item(i)) & _
                    vbCrLf & "Centre (X ; Y ; Z) : " & _
                    Format$(centreRoll(0), "0.###") & " ; " & _
                    Format$(centreRoll(1), "0.###") & " ; " & _
                    Format$(centreRoll(2), "0.###") & _
                    vbCrLf & "Tolerance horizontale : " & _
                    CStr(TECH_TOLERANCE_LIGNE) & " mm"
            End If
            footHeight = TECH_CalculerFootHeight(hauteurRoll, ligneY)
            Debug.Print "PIED MOTEUR : " & nomRepereEtat & _
                " / centre Y=" & CStr(hauteurRoll) & _
                " / ligne Y=" & CStr(ligneY) & _
                " / FOOT_HEIGHT=" & CStr(footHeight)
            If Not validerSeulement Then
                Set modeleMoteurVariante = TECH_ObtenirVarianteMoteur( _
                    session, modeleMoteurSource, dossierMoteur, footHeight, _
                    cacheVariantesMoteur, nomPiedVariante)
                Set solidMoteurVariante = modeleMoteurVariante
                If solidMoteurVariante Is Nothing Then
                    Err.Raise vbObjectError + 15143, , _
                        "La variante moteur n'est pas exploitable comme solide : " & _
                        modeleMoteurVariante.Filename
                End If
                If TECH_SynchroniserMoteur(session, asm, solidAsm, _
                    solidMoteurVariante, CLng(rollIDs.item(i)), coteMoteur, _
                    wsEtat, ligneEtat, hauteurRoll, footHeight, _
                    nomPiedVariante, nbAjoutes, nbDeplaces, nbInchanges) Then
                    modifie = True
                End If
            End If
        End If
    Next i

    If modifie Then
        m_TECH_Etape = "Validation des moteurs dans " & nomAssemblage
        TECH_ValiderSolid solidAsm, "Assemblage moteur " & nomAssemblage
        TECH_SauvegarderModele modeleAssemblage, _
            "Assemblage " & nomAssemblage & " apres placement des moteurs"
    End If

    For i = 1 To sousAssemblages.Count
        TECH_TraiterMoteurs session, sousAssemblages.item(i), _
            modeleMoteurSource, dossierMoteur, cacheVariantesMoteur, _
            coteMoteur, wsEtat, visites, validerSeulement, nbAjoutes, _
            nbDeplaces, nbInchanges
    Next i
End Sub

Private Function TECH_SynchroniserMoteur( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal asm As pfcls.IpfcAssembly, _
    ByVal solidAsm As pfcls.IpfcSolid, _
    ByVal solidMoteur As pfcls.IpfcSolid, _
    ByVal rollFeatureID As Long, _
    ByVal coteMoteur As String, _
    ByVal wsEtat As Worksheet, _
    ByVal ligneEtat As Long, _
    ByVal hauteurRoll As Double, _
    ByVal footHeight As Double, _
    ByVal nomPiedVariante As String, _
    ByRef nbAjoutes As Long, _
    ByRef nbDeplaces As Long, _
    ByRef nbInchanges As Long) As Boolean

    Dim ancienID As Long
    Dim nouveauID As Long
    Dim ancienExiste As Boolean
    Dim modeleAncien As pfcls.IpfcModel
    Dim modeleMoteurAttendu As pfcls.IpfcModel
    Dim nomMoteurAttendu As String
    Dim numeroErreur As Long
    Dim descriptionErreur As String

    Set modeleMoteurAttendu = solidMoteur
    nomMoteurAttendu = TECH_NomEffectif(modeleMoteurAttendu)
    ancienID = CLng(Val(CStr(wsEtat.Cells(ligneEtat, TECH_COL_MOTEUR_ID).value)))
    If ancienID > 0 And _
        TECH_FeatureComposantActive(solidAsm, ancienID) Then
        Set modeleAncien = TECH_ModeleComposantParID(session, solidAsm, ancienID)
        ancienExiste = Not modeleAncien Is Nothing
    End If

    If ancienExiste Then
        If StrComp(TECH_NomEffectif(modeleAncien), _
            nomMoteurAttendu, vbTextCompare) = 0 And _
            StrComp(Trim$(CStr(wsEtat.Cells(ligneEtat, TECH_COL_MOTEUR_COTE).value)), _
                coteMoteur, vbTextCompare) = 0 And _
            TECH_OccurrenceRollPlacee(solidAsm, ancienID) Then
            TECH_EcrireEtatMoteur wsEtat, ligneEtat, coteMoteur, _
                nomMoteurAttendu, ancienID, hauteurRoll, footHeight, _
                nomPiedVariante
            nbInchanges = nbInchanges + 1
            Exit Function
        End If
    End If

    m_TECH_Etape = "Placement moteur sur " & coteMoteur & _
        " / rouleau ID " & CStr(rollFeatureID)
    nouveauID = TECH_AjouterMoteur(session, asm, solidAsm, solidMoteur, _
        rollFeatureID, coteMoteur)

    If ancienExiste Then
        On Error GoTo RemplacementEchoue
        If TECH_EstModeleMoteur(modeleAncien.Filename) Then
            TECH_SupprimerComposant session, asm, solidAsm, ancienID
        End If
        On Error GoTo 0
        nbDeplaces = nbDeplaces + 1
    Else
        nbAjoutes = nbAjoutes + 1
    End If

    TECH_EcrireEtatMoteur wsEtat, ligneEtat, coteMoteur, _
        nomMoteurAttendu, nouveauID, hauteurRoll, footHeight, _
        nomPiedVariante
    TECH_SynchroniserMoteur = True
    Exit Function

RemplacementEchoue:
    numeroErreur = Err.Number
    descriptionErreur = Err.description
    On Error GoTo -1 ' V14 : libere le gestionnaire actif avant nettoyage
    On Error Resume Next
    TECH_SupprimerComposant session, asm, solidAsm, nouveauID
    On Error GoTo 0
    Err.Raise numeroErreur, , _
        "Deplacement du moteur impossible. L'ancien moteur a ete conserve." & _
        vbCrLf & descriptionErreur
End Function

Private Sub TECH_EcrireEtatMoteur( _
    ByVal wsEtat As Worksheet, _
    ByVal ligneEtat As Long, _
    ByVal coteMoteur As String, _
    ByVal nomMoteur As String, _
    ByVal moteurFeatureID As Long, _
    ByVal hauteurRoll As Double, _
    ByVal footHeight As Double, _
    ByVal nomPied As String)

    wsEtat.Cells(ligneEtat, TECH_COL_MOTEUR_COTE).value = coteMoteur
    wsEtat.Cells(ligneEtat, TECH_COL_MOTEUR_MODELE).value = nomMoteur
    wsEtat.Cells(ligneEtat, TECH_COL_MOTEUR_ID).value = moteurFeatureID
    wsEtat.Cells(ligneEtat, TECH_COL_ROLL_HAUTEUR).value = hauteurRoll
    wsEtat.Cells(ligneEtat, TECH_COL_FOOT_HEIGHT).value = footHeight
    wsEtat.Cells(ligneEtat, TECH_COL_PIED_MODELE).value = nomPied
    wsEtat.Cells(ligneEtat, TECH_COL_MOTEUR_SOURCE).value = _
        m_TECH_MoteurFichier
End Sub

Private Function TECH_CalculerFootHeight( _
    ByVal centreY As Double, _
    ByVal ligneY As Double) As Double

    Dim resultat As Double

    ' V15 : le pied descend du centre du rouleau jusqu'a la premiere ligne
    ' du squelette rencontree sous ce centre (charpente).
    resultat = Round(centreY - ligneY + TECH_PIED_DECALAGE_LIGNE, 6)
    If resultat <= 0# Then
        Err.Raise vbObjectError + 15205, , _
            "FOOT_HEIGHT calcule est invalide." & vbCrLf & _
            "Centre rouleau Y : " & CStr(centreY) & " mm" & vbCrLf & _
            "Ligne squelette Y : " & CStr(ligneY) & " mm" & vbCrLf & _
            "FOOT_HEIGHT : " & CStr(resultat) & " mm"
    End If
    TECH_CalculerFootHeight = resultat
End Function

Private Function TECH_ListerLignesSquelette( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal asm As pfcls.IpfcAssembly, _
    ByVal solidAsm As pfcls.IpfcSolid, _
    ByVal nomAssemblage As String, _
    ByRef nomSquelette As String) As Collection

    Dim asmObjet As Object
    Dim squeletteDeclare As Object
    Dim nomDeclare As String
    Dim features As pfcls.IpfcFeatures
    Dim feature As pfcls.IpfcFeature
    Dim item As pfcls.IpfcModelItem
    Dim composant As pfcls.IpfcComponentFeat
    Dim modele As pfcls.IpfcModel
    Dim modeleObjet As Object
    Dim modeleSquelette As pfcls.IpfcModel
    Dim idSquelette As Long
    Dim nbCandidats As Long
    Dim listeCandidats As String
    Dim estSquelette As Boolean
    Dim path As pfcls.IpfcComponentPath
    Dim transformation As pfcls.IpfcTransform3D
    Dim proprietaire As pfcls.IpfcModelItemOwner
    Dim courbes As pfcls.IpfcModelItems
    Dim segments As Collection
    Dim i As Long

    nomSquelette = ""

    ' 1. Squelette declare de l'assemblage (Creo : IpfcAssembly.GetSkeleton).
    On Error Resume Next
    Set asmObjet = asm
    Set squeletteDeclare = asmObjet.GetSkeleton()
    If Not squeletteDeclare Is Nothing Then
        nomDeclare = TECH_NomModeleLogique(CStr(squeletteDeclare.Filename))
    End If
    Err.Clear
    On Error GoTo 0

    ' 2. Composant correspondant (ID necessaire pour la transformation).
    Set features = solidAsm.ListFeaturesByType(False, pfcls.EpfcFEATTYPE_COMPONENT)
    If Not features Is Nothing Then
        For i = 0 To features.Count - 1
            Set feature = features.item(i)
            Set item = Nothing
            Set composant = Nothing
            Set modele = Nothing
            If Not feature Is Nothing Then Set item = feature
            If Not feature Is Nothing Then Set composant = feature
            If Not composant Is Nothing Then
                Set modele = TECH_ModeleDepuisComposant(session, composant)
            End If
            If Not modele Is Nothing Then
                If modele.Type = pfcls.EpfcMDL_PART Then
                    estSquelette = False
                    If nomDeclare <> "" Then
                        estSquelette = (StrComp(TECH_NomEffectif(modele), _
                            nomDeclare, vbTextCompare) = 0)
                    Else
                        On Error Resume Next
                        Set modeleObjet = modele
                        estSquelette = CBool(modeleObjet.IsSkeleton)
                        Err.Clear
                        On Error GoTo 0
                    End If
                    If estSquelette Then
                        nbCandidats = nbCandidats + 1
                        Set modeleSquelette = modele
                        idSquelette = CLng(item.id)
                        If listeCandidats <> "" Then _
                            listeCandidats = listeCandidats & vbCrLf
                        listeCandidats = listeCandidats & "- " & _
                            TECH_NomEffectif(modele) & " (ID " & CStr(item.id) & ")"
                    End If
                End If
            End If
        Next i
    End If

    If nbCandidats = 0 Then
        Err.Raise vbObjectError + 15300, , _
            "Aucun squelette trouve dans le sous-assemblage " & _
            nomAssemblage & "." & vbCrLf & _
            "La hauteur des pieds moteur est calculee a partir des lignes " & _
            "de charpente du squelette de ce sous-assemblage."
    ElseIf nbCandidats > 1 Then
        Err.Raise vbObjectError + 15301, , _
            "Plusieurs squelettes dans " & nomAssemblage & " :" & vbCrLf & _
            listeCandidats
    End If
    nomSquelette = TECH_NomEffectif(modeleSquelette)

    ' 3. Lignes du squelette exprimees dans le repere du sous-assemblage.
    Set path = TECH_CreerPath(asm, idSquelette)
    Set transformation = path.GetTransform(True)
    If transformation Is Nothing Then
        Err.Raise vbObjectError + 15302, , _
            "Position du squelette " & nomSquelette & " illisible."
    End If

    Set segments = New Collection
    Set proprietaire = modeleSquelette
    Set courbes = proprietaire.ListItems(pfcls.EpfcITEM_CURVE)
    If Not courbes Is Nothing Then
        For i = 0 To courbes.Count - 1
            TECH_AjouterSegmentsCourbe courbes.item(i), transformation, segments
        Next i
    End If

    If segments.Count = 0 Then
        Err.Raise vbObjectError + 15303, , _
            "Le squelette " & nomSquelette & " de " & nomAssemblage & _
            " ne contient aucune ligne droite exploitable."
    End If
    Debug.Print "SQUELETTE " & nomSquelette & " (" & nomAssemblage & ") : " & _
        CStr(segments.Count) & " ligne(s)"
    Set TECH_ListerLignesSquelette = segments
End Function

Private Sub TECH_AjouterSegmentsCourbe( _
    ByVal itemCourbe As pfcls.IpfcModelItem, _
    ByVal transformation As pfcls.IpfcTransform3D, _
    ByVal segments As Collection)

    Dim courbe As pfcls.IpfcGeomCurve
    Dim descripteur As pfcls.IpfcCurveDescriptor
    Dim visible As Boolean

    If itemCourbe Is Nothing Then Exit Sub
    On Error Resume Next
    Set courbe = itemCourbe
    If courbe Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    visible = True
    visible = CBool(courbe.IsVisible)
    Err.Clear
    Set descripteur = courbe.GetCurveDescriptor()
    Err.Clear
    On Error GoTo 0
    If Not visible Then Exit Sub

    If Not descripteur Is Nothing Then
        If TECH_AjouterSegmentsDescripteur(descripteur, transformation, _
            segments, 0) Then Exit Sub
    End If
    ' Repli : evaluation parametrique, retenue seulement si la courbe est droite.
    TECH_AjouterSegmentParEvaluation courbe, transformation, segments
End Sub

Private Function TECH_AjouterSegmentsDescripteur( _
    ByVal descripteur As pfcls.IpfcCurveDescriptor, _
    ByVal transformation As pfcls.IpfcTransform3D, _
    ByVal segments As Collection, _
    ByVal profondeur As Long) As Boolean

    Dim ligne As pfcls.IpfcLineDescriptor
    Dim composite As pfcls.IpfcCompositeCurveDescriptor
    Dim compositeObjet As Object
    Dim elements As Object
    Dim element As pfcls.IpfcCurveDescriptor
    Dim nb As Long
    Dim i As Long

    If descripteur Is Nothing Or profondeur > 10 Then Exit Function

    On Error Resume Next
    Err.Clear
    Set ligne = descripteur
    If Err.Number <> 0 Then Set ligne = Nothing
    Err.Clear
    On Error GoTo 0
    If Not ligne Is Nothing Then
        TECH_AjouterSegment ligne.End1, ligne.End2, transformation, segments
        TECH_AjouterSegmentsDescripteur = True
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    Set composite = descripteur
    If Err.Number <> 0 Then Set composite = Nothing
    Err.Clear
    If Not composite Is Nothing Then
        Set compositeObjet = composite
        Set elements = compositeObjet.Elements
        nb = 0
        If Not elements Is Nothing Then nb = elements.Count
        Err.Clear
    End If
    On Error GoTo 0
    If composite Is Nothing Then Exit Function

    For i = 0 To nb - 1
        Set element = Nothing
        On Error Resume Next
        Set element = elements.Item(i)
        Err.Clear
        On Error GoTo 0
        If Not element Is Nothing Then
            TECH_AjouterSegmentsDescripteur element, transformation, _
                segments, profondeur + 1
        End If
    Next i
    TECH_AjouterSegmentsDescripteur = True
End Function

Private Sub TECH_AjouterSegmentParEvaluation( _
    ByVal courbe As pfcls.IpfcGeomCurve, _
    ByVal transformation As pfcls.IpfcTransform3D, _
    ByVal segments As Collection)

    Dim p0 As pfcls.IpfcPoint3D
    Dim pm As pfcls.IpfcPoint3D
    Dim p1 As pfcls.IpfcPoint3D
    Dim a As Variant
    Dim m As Variant
    Dim b As Variant
    Dim longueur2 As Double
    Dim t As Double
    Dim ecart2 As Double
    Dim k As Long

    On Error Resume Next
    Err.Clear
    Set p0 = courbe.Eval3DData(0#).Point
    Set pm = courbe.Eval3DData(0.5).Point
    Set p1 = courbe.Eval3DData(1#).Point
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    If p0 Is Nothing Or pm Is Nothing Or p1 Is Nothing Then Exit Sub

    a = TECH_LirePoint(p0)
    m = TECH_LirePoint(pm)
    b = TECH_LirePoint(p1)
    For k = 0 To 2
        longueur2 = longueur2 + (b(k) - a(k)) ^ 2
    Next k
    If longueur2 < 0.000001 Then Exit Sub
    For k = 0 To 2
        t = t + (m(k) - a(k)) * (b(k) - a(k))
    Next k
    t = t / longueur2
    For k = 0 To 2
        ecart2 = ecart2 + (m(k) - (a(k) + t * (b(k) - a(k)))) ^ 2
    Next k
    If Sqr(ecart2) > TECH_TOLERANCE_LIGNE Then Exit Sub
    TECH_AjouterSegment p0, p1, transformation, segments
End Sub

Private Sub TECH_AjouterSegment( _
    ByVal pointA As pfcls.IpfcPoint3D, _
    ByVal pointB As pfcls.IpfcPoint3D, _
    ByVal transformation As pfcls.IpfcTransform3D, _
    ByVal segments As Collection)

    Dim a As Variant
    Dim b As Variant

    If pointA Is Nothing Or pointB Is Nothing Then Exit Sub
    a = TECH_LirePoint(transformation.TransformPoint(pointA))
    b = TECH_LirePoint(transformation.TransformPoint(pointB))
    segments.Add Array(a(0), a(1), a(2), b(0), b(1), b(2))
End Sub

Private Function TECH_LirePoint(ByVal point As pfcls.IpfcPoint3D) As Variant
    TECH_LirePoint = Array(CDbl(point.item(0)), CDbl(point.item(1)), _
        CDbl(point.item(2)))
End Function

Private Function TECH_CentreRouleau( _
    ByVal asm As pfcls.IpfcAssembly, _
    ByVal rollFeatureID As Long, _
    ByVal modeleRoll As pfcls.IpfcModel) As Variant

    Dim itemCsys As pfcls.IpfcModelItem
    Dim csys As pfcls.IpfcCoordSystem
    Dim origine As pfcls.IpfcPoint3D
    Dim path As pfcls.IpfcComponentPath
    Dim transformation As pfcls.IpfcTransform3D

    Set itemCsys = TECH_TrouverCsysExact(modeleRoll, TECH_ROLL_CSYS)
    If itemCsys Is Nothing Then
        Err.Raise vbObjectError + 15304, , _
            "Repere " & TECH_ROLL_CSYS & " absent du rouleau ID " & _
            CStr(rollFeatureID) & "."
    End If
    Set csys = itemCsys
    Set origine = csys.CoordSys.GetOrigin()
    Set path = TECH_CreerPath(asm, rollFeatureID)
    Set transformation = path.GetTransform(True)
    TECH_CentreRouleau = TECH_LirePoint(transformation.TransformPoint(origine))
End Function

Private Function TECH_TrouverLigneSousCentre( _
    ByVal centre As Variant, _
    ByVal segments As Collection, _
    ByRef ligneY As Double) As Boolean

    ' Rayon vertical descendant depuis le centre (axe TECH_AXE_VERTICAL).
    ' La distance horizontale ignore TECH_AXE_PROFONDEUR (plan de l'esquisse
    ' de charpente). La ligne retenue est la plus haute strictement sous le
    ' centre : c'est la premiere rencontree en descendant.
    Dim segment As Variant
    Dim axesH(0 To 1) As Long
    Dim nbAxes As Long
    Dim k As Long
    Dim h As Long
    Dim p1(0 To 2) As Double
    Dim p2(0 To 2) As Double
    Dim longueur2 As Double
    Dim t As Double
    Dim distance2 As Double
    Dim impactY As Double
    Dim meilleurY As Double
    Dim trouve As Boolean
    Dim v As Long

    v = TECH_AXE_VERTICAL
    For k = 0 To 2
        If k <> v And k <> TECH_AXE_PROFONDEUR Then
            axesH(nbAxes) = k
            nbAxes = nbAxes + 1
        End If
    Next k

    For Each segment In segments
        For k = 0 To 2
            p1(k) = segment(k)
            p2(k) = segment(k + 3)
        Next k

        longueur2 = 0#
        For h = 0 To nbAxes - 1
            longueur2 = longueur2 + (p2(axesH(h)) - p1(axesH(h))) ^ 2
        Next h

        If longueur2 < 0.000001 Then
            ' Ligne verticale (ou point) : premier contact a son extremite haute.
            t = 0#
            If p2(v) > p1(v) Then t = 1#
        Else
            t = 0#
            For h = 0 To nbAxes - 1
                t = t + (centre(axesH(h)) - p1(axesH(h))) * _
                    (p2(axesH(h)) - p1(axesH(h)))
            Next h
            t = t / longueur2
            If t < 0# Then t = 0#
            If t > 1# Then t = 1#
        End If

        distance2 = 0#
        For h = 0 To nbAxes - 1
            distance2 = distance2 + _
                (p1(axesH(h)) + t * (p2(axesH(h)) - p1(axesH(h))) - _
                centre(axesH(h))) ^ 2
        Next h

        If Sqr(distance2) <= TECH_TOLERANCE_LIGNE Then
            impactY = p1(v) + t * (p2(v) - p1(v))
            If impactY < centre(v) - 0.01 Then
                If Not trouve Or impactY > meilleurY Then
                    meilleurY = impactY
                    trouve = True
                End If
            End If
        End If
    Next segment

    ligneY = meilleurY
    TECH_TrouverLigneSousCentre = trouve
End Function

Private Sub TECH_ValiderModeleMoteurSource( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal modeleMoteurSource As pfcls.IpfcModel)

    Dim solidMoteurSource As pfcls.IpfcSolid
    Dim featurePied As pfcls.IpfcFeature
    Dim composantPied As pfcls.IpfcComponentFeat
    Dim modelePied As pfcls.IpfcModel
    Dim solidPied As pfcls.IpfcSolid
    Dim featureID As Long
    Dim dimension As pfcls.IpfcBaseDimension

    If modeleMoteurSource Is Nothing Then
        Err.Raise vbObjectError + 15206, , _
            "Assemblage moteur source Nothing."
    End If
    Set solidMoteurSource = modeleMoteurSource
    If solidMoteurSource Is Nothing Then
        Err.Raise vbObjectError + 15207, , _
            "Assemblage moteur source non exploitable."
    End If

    TECH_TrouverComposantPied session, solidMoteurSource, featurePied, _
        composantPied, modelePied, featureID
    Debug.Print "PIED TROUVE DANS L'ASSEMBLAGE MOTEUR : " & _
        modelePied.Filename & " / feature " & CStr(featureID)

    Set solidPied = modelePied
    If solidPied Is Nothing Then
        Err.Raise vbObjectError + 15209, , _
            TECH_PIED_FICHIER & " n'est pas exploitable comme piece."
    End If
    Set dimension = TECH_TrouverDimension(solidPied, TECH_PIED_DIM)
    If dimension Is Nothing Then
        Err.Raise vbObjectError + 15210, , _
            "La cote " & TECH_PIED_DIM & " est absente de " & _
            TECH_PIED_FICHIER & "."
    End If
End Sub

Private Function TECH_ObtenirVarianteMoteur( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal modeleMoteurSource As pfcls.IpfcModel, _
    ByVal dossierMoteur As String, _
    ByVal footHeight As Double, _
    ByVal cacheVariantes As Object, _
    ByRef nomPiedVariante As String) As pfcls.IpfcModel

    Dim nomBase As String
    Dim nomFichier As String
    Dim cle As String
    Dim modele As pfcls.IpfcModel
    Dim solideSource As pfcls.IpfcSolid
    Dim solideInstance As pfcls.IpfcSolid
    Dim featurePied As pfcls.IpfcFeature
    Dim composantPied As pfcls.IpfcComponentFeat
    Dim modelePiedSource As pfcls.IpfcModel
    Dim modelePiedInstance As pfcls.IpfcModel
    Dim piedID As Long
    Dim membre As pfcls.IpfcFamilyMember
    Dim colonne As pfcls.IpfcFamilyTableColumn
    Dim ligne As pfcls.IpfcFamilyTableRow
    Dim fabriqueValeur As pfcls.CMpfcModelItem
    Dim valeur As pfcls.IpfcParamValue
    Dim nomColonne As String
    Dim colonneComposant As pfcls.IpfcFamColComp
    Dim nouvelleLigne As Boolean
    Dim colonneCreee As Boolean
    Dim celluleReparee As Boolean
    Dim etapeLocale As String
    Dim errNumero As Long
    Dim errDescription As String
    Dim csys As pfcls.IpfcModelItem

    On Error GoTo Erreur
    etapeLocale = "Recherche du pied de l'assemblage moteur source"
    Set solideSource = modeleMoteurSource
    TECH_TrouverComposantPied session, solideSource, featurePied, _
        composantPied, modelePiedSource, piedID

    nomPiedVariante = TECH_NomVariantePied( _
        modelePiedSource.Filename, footHeight) & ".prt"
    nomBase = TECH_NomVarianteMoteur(modeleMoteurSource.Filename, footHeight)
    nomFichier = nomBase & ".asm"
    cle = "FTM|" & UCase$(modeleMoteurSource.Filename) & "|" & UCase$(nomFichier)
    If cacheVariantes.Exists(cle) Then
        Set TECH_ObtenirVarianteMoteur = cacheVariantes(cle)
        Exit Function
    End If

    etapeLocale = "Instance de pied " & nomPiedVariante
    Set modelePiedInstance = TECH_ObtenirInstancePiedFamille( _
        session, modelePiedSource, footHeight, cacheVariantes)

    etapeLocale = "Colonne du composant pied dans la table moteur"
    Set membre = TECH_MembreFamilleDepuisModele(modeleMoteurSource)
    nomColonne = "m" & CStr(piedID)
    Set colonne = TECH_TrouverColonneFamille(membre, _
        Array(nomColonne, "M" & CStr(piedID)), _
        pfcls.EpfcITEM_FEATURE, piedID, _
        Array("mm" & CStr(piedID), "MM" & CStr(piedID)))
    If colonne Is Nothing Then
        Set colonneComposant = membre.CreateComponentColumn(featurePied)
        Set colonne = TECH_AjouterColonneFamille(membre, colonneComposant)
        colonneCreee = True
    End If
    If colonne Is Nothing Then
        Err.Raise vbObjectError + 15261, , _
            "Colonne " & nomColonne & " non disponible dans " & _
            modeleMoteurSource.Filename
    End If

    etapeLocale = "Ligne moteur " & nomBase
    Set ligne = TECH_LigneFamille(membre, nomBase)
    If ligne Is Nothing Then
        Set ligne = TECH_AjouterLigneFamille(membre, nomBase)
        nouvelleLigne = True
    End If
    If ligne Is Nothing Then
        Err.Raise vbObjectError + 15262, , _
            "Impossible de creer l'instance moteur " & nomBase
    End If

    If Not nouvelleLigne Then
        ' V14 : une ligne creee lors d'un essai interrompu peut rester "*".
        If TECH_CelluleParDefaut(membre, colonne, ligne) Then
            celluleReparee = True
            Debug.Print "TABLE MOTEUR : cellule par defaut renseignee pour " & nomBase
        End If
    End If
    If nouvelleLigne Or celluleReparee Then
        etapeLocale = "Selection du pied " & nomPiedVariante
        Set fabriqueValeur = New pfcls.CMpfcModelItem
        Set valeur = fabriqueValeur.CreateStringParamValue( _
            TECH_BaseSansExtension(nomPiedVariante))
        membre.SetCell colonne, ligne, valeur
    Else
        etapeLocale = "Controle de la variante moteur existante"
        Set valeur = membre.GetCell(colonne, ligne)
        If valeur Is Nothing Then
            Err.Raise vbObjectError + 15264, , _
                "Cellule de pied vide pour l'instance " & nomBase
        End If
        If StrComp(TECH_BaseSansExtension(valeur.StringValue), _
            TECH_BaseSansExtension(nomPiedVariante), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 15265, , _
                "L'instance moteur " & nomBase & _
                " pointe vers " & valeur.StringValue & _
                " au lieu de " & nomPiedVariante
        End If
    End If

    If nouvelleLigne Or colonneCreee Or celluleReparee Then
        etapeLocale = "Enregistrement de la table moteur"
        TECH_SauvegarderModele modeleMoteurSource, _
            "Table de famille du moteur " & modeleMoteurSource.Filename
    End If

    etapeLocale = "Ouverture de l'instance moteur " & nomFichier
    On Error Resume Next
    Set modele = session.GetModelFromFileName(nomFichier)
    Err.Clear
    On Error GoTo Erreur
    If modele Is Nothing Then Set modele = ligne.CreateInstance()
    If modele Is Nothing Then
        Err.Raise vbObjectError + 15267, , _
            "Creo n'a pas genere l'instance moteur " & nomBase
    End If

    etapeLocale = "Verification du pied de l'instance moteur"
    Set solideInstance = modele
    TECH_TrouverComposantPied session, solideInstance, featurePied, _
        composantPied, modelePiedInstance, piedID
    If StrComp(TECH_NomEffectif(modelePiedInstance), _
        TECH_NomModeleLogique(nomPiedVariante), vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 15268, , _
            "La table de famille a conserve " & _
            modelePiedInstance.Filename & " au lieu de " & nomPiedVariante
    End If

    Set csys = TECH_TrouverCsysExact(modele, TECH_MOTEUR_CSYS)
    If csys Is Nothing Then
        Err.Raise vbObjectError + 15269, , _
            "CSYS MOTEUR absent de l'instance " & nomFichier
    End If
    TECH_ValiderSolid solideInstance, "Instance moteur " & nomFichier

    cacheVariantes.Add cle, modele
    Set TECH_ObtenirVarianteMoteur = modele
    Exit Function

Erreur:
    errNumero = Err.Number
    errDescription = Err.description
    If InStr(1, errDescription, "CheckIsModifiable", vbTextCompare) = 0 Then _
        errDescription = errDescription & vbCrLf & TECH_EtatEcriture(modeleMoteurSource)
    Err.Raise errNumero, , _
        "TECH_ObtenirVarianteMoteur :" & vbCrLf & _
        "Etape interne : " & etapeLocale & vbCrLf & errDescription
End Function

Private Function TECH_ObtenirInstancePiedFamille( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal modeleSource As pfcls.IpfcModel, _
    ByVal footHeight As Double, _
    ByVal cacheVariantes As Object) As pfcls.IpfcModel

    Dim nomBase As String
    Dim nomFichier As String
    Dim cle As String
    Dim membre As pfcls.IpfcFamilyMember
    Dim solide As pfcls.IpfcSolid
    Dim dimensionBase As pfcls.IpfcBaseDimension
    Dim dimension As pfcls.IpfcDimension
    Dim itemDimension As pfcls.IpfcModelItem
    Dim colonne As pfcls.IpfcFamilyTableColumn
    Dim ligne As pfcls.IpfcFamilyTableRow
    Dim fabriqueValeur As pfcls.CMpfcModelItem
    Dim valeur As pfcls.IpfcParamValue
    Dim instance As pfcls.IpfcModel
    Dim solideInstance As pfcls.IpfcSolid
    Dim dimensionInstance As pfcls.IpfcBaseDimension
    Dim nouvelleLigne As Boolean
    Dim colonneCreee As Boolean
    Dim celluleReparee As Boolean
    Dim etapeLocale As String
    Dim errNumero As Long
    Dim errDescription As String

    On Error GoTo Erreur
    nomBase = TECH_NomVariantePied(modeleSource.Filename, footHeight)
    nomFichier = nomBase & ".prt"
    cle = "FTF|" & UCase$(modeleSource.Filename) & "|" & UCase$(nomFichier)
    If cacheVariantes.Exists(cle) Then
        Set TECH_ObtenirInstancePiedFamille = cacheVariantes(cle)
        Exit Function
    End If

    etapeLocale = "Recherche de la cote FOOT_HEIGHT du pied generique"
    Set solide = modeleSource
    Set dimensionBase = TECH_TrouverDimension(solide, TECH_PIED_DIM)
    If dimensionBase Is Nothing Then
        Err.Raise vbObjectError + 15270, , _
            "FOOT_HEIGHT absent de " & modeleSource.Filename
    End If
    On Error Resume Next
    Set dimension = dimensionBase
    Err.Clear
    On Error GoTo Erreur
    If dimension Is Nothing Then
        Err.Raise vbObjectError + 15271, , _
            "FOOT_HEIGHT doit etre une cote pilotable, pas une cote de reference."
    End If
    Set itemDimension = dimension
    Set membre = TECH_MembreFamilleDepuisModele(modeleSource)

    etapeLocale = "Colonne FOOT_HEIGHT de la table du pied"
    Set colonne = TECH_TrouverColonneFamille(membre, _
        Array("d" & CStr(itemDimension.id), "D" & CStr(itemDimension.id)), _
        pfcls.EpfcITEM_DIMENSION, itemDimension.id, _
        Array(dimensionBase.Symbol, TECH_PIED_DIM))
    If colonne Is Nothing Then
        Set colonne = TECH_CreerColonneHauteurPied(membre, modeleSource, _
            dimension, itemDimension.id, dimensionBase.Symbol)
        colonneCreee = True
    End If
    If colonne Is Nothing Then
        Err.Raise vbObjectError + 15272, , _
            "Colonne de FOOT_HEIGHT indisponible dans la table du pied."
    End If

    etapeLocale = "Ligne du pied " & nomBase
    Set ligne = TECH_LigneFamille(membre, nomBase)
    If ligne Is Nothing Then
        Set ligne = TECH_AjouterLigneFamille(membre, nomBase)
        nouvelleLigne = True
    End If
    If ligne Is Nothing Then
        Err.Raise vbObjectError + 15273, , _
            "Impossible de creer l'instance de pied " & nomBase
    End If

    If Not nouvelleLigne Then
        ' V14 : une ligne creee lors d'un essai interrompu peut rester "*".
        If TECH_CelluleParDefaut(membre, colonne, ligne) Then
            celluleReparee = True
            Debug.Print "TABLE PIED : cellule par defaut renseignee pour " & nomBase
        End If
    End If
    If nouvelleLigne Or celluleReparee Then
        etapeLocale = "FOOT_HEIGHT de " & nomBase
        Set fabriqueValeur = New pfcls.CMpfcModelItem
        Set valeur = fabriqueValeur.CreateDoubleParamValue(footHeight)
        membre.SetCell colonne, ligne, valeur
    Else
        etapeLocale = "Controle de l'instance de pied existante"
        Set valeur = membre.GetCell(colonne, ligne)
        If valeur Is Nothing Then
            Err.Raise vbObjectError + 15275, , _
                "Cellule FOOT_HEIGHT vide pour " & nomBase
        End If
        If Not TECH_DoublesEgaux(TECH_ValeurNumeriqueParam(valeur), footHeight) Then
            Err.Raise vbObjectError + 15276, , _
                "L'instance " & nomBase & " porte " & _
                CStr(TECH_ValeurNumeriqueParam(valeur)) & " mm au lieu de " & _
                CStr(footHeight) & " mm."
        End If
    End If

    If nouvelleLigne Or colonneCreee Or celluleReparee Then
        etapeLocale = "Enregistrement de la table du pied"
        TECH_SauvegarderModele modeleSource, _
            "Table de famille du pied " & modeleSource.Filename
    End If

    etapeLocale = "Ouverture de l'instance de pied " & nomFichier
    On Error Resume Next
    Set instance = session.GetModelFromFileName(nomFichier)
    Err.Clear
    On Error GoTo Erreur
    If instance Is Nothing Then Set instance = ligne.CreateInstance()
    If instance Is Nothing Then
        Err.Raise vbObjectError + 15278, , _
            "Creo n'a pas genere l'instance de pied " & nomBase
    End If
    Set solideInstance = instance
    Set dimensionInstance = TECH_TrouverDimension( _
        solideInstance, TECH_PIED_DIM)
    If dimensionInstance Is Nothing Then
        Err.Raise vbObjectError + 15279, , _
            "FOOT_HEIGHT absent de l'instance " & nomBase
    End If
    TECH_ValiderSolid solideInstance, "Instance pied " & nomFichier
    If Not TECH_DoublesEgaux( _
        TECH_ValeurDimensionPositive(dimensionInstance), footHeight) Then
        Err.Raise vbObjectError + 15280, , _
            "FOOT_HEIGHT genere pour " & nomBase & " vaut " & _
            CStr(TECH_ValeurDimensionPositive(dimensionInstance)) & _
            " mm ; attendu : " & CStr(footHeight) & " mm."
    End If

    cacheVariantes.Add cle, instance
    Set TECH_ObtenirInstancePiedFamille = instance
    Exit Function

Erreur:
    errNumero = Err.Number
    errDescription = Err.description
    If InStr(1, errDescription, "CheckIsModifiable", vbTextCompare) = 0 Then _
        errDescription = errDescription & vbCrLf & TECH_EtatEcriture(modeleSource)
    Err.Raise errNumero, , _
        "TECH_ObtenirInstancePiedFamille :" & vbCrLf & _
        "Etape interne : " & etapeLocale & vbCrLf & errDescription
End Function

Private Function TECH_MembreFamilleDepuisModele( _
    ByVal modele As pfcls.IpfcModel) As pfcls.IpfcFamilyMember

    Dim membreTypage As pfcls.IpfcFamilyMember

    ' V16 : liaison precoce (types pfcls) pour toutes les operations de
    ' table de famille. L'appel tardif via Object provoquait l'erreur 424
    ' "Objet requis" sur CreateDimensionColumn.
    On Error Resume Next
    Set membreTypage = modele
    Err.Clear
    On Error GoTo 0
    If membreTypage Is Nothing Then
        Err.Raise vbObjectError + 15282, , _
            "Modele non exploitable comme membre de famille : " & modele.Filename
    End If
    Set TECH_MembreFamilleDepuisModele = membreTypage
End Function

Private Function TECH_EssayerAjouterColonneFamille( _
    ByVal membre As pfcls.IpfcFamilyMember, _
    ByVal colonneNouvelle As pfcls.IpfcFamilyTableColumn, _
    ByRef messageErreur As String) As pfcls.IpfcFamilyTableColumn

    Dim resultat As pfcls.IpfcFamilyTableColumn
    Dim messageNull As String

    ' AddColumn(Column, Values) : le second argument doit etre fourni.
    ' PTC : en VBA, passer Null. Nothing est essaye en second recours.
    messageErreur = ""
    If colonneNouvelle Is Nothing Then
        messageErreur = "Objet colonne Nothing."
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    Set resultat = membre.AddColumn(colonneNouvelle, Null)
    If Err.Number <> 0 Then
        messageNull = CStr(Err.Number) & " " & Err.Description
        Err.Clear
        Set resultat = Nothing
        Set resultat = membre.AddColumn(colonneNouvelle, Nothing)
        If Err.Number <> 0 Then
            messageErreur = "AddColumn(Null) : " & messageNull & vbCrLf & _
                "AddColumn(Nothing) : " & CStr(Err.Number) & " " & _
                Err.Description
            Set resultat = Nothing
        End If
    End If
    Err.Clear
    On Error GoTo 0
    Set TECH_EssayerAjouterColonneFamille = resultat
End Function

Private Function TECH_AjouterColonneFamille( _
    ByVal membre As pfcls.IpfcFamilyMember, _
    ByVal colonneNouvelle As pfcls.IpfcFamilyTableColumn) _
    As pfcls.IpfcFamilyTableColumn

    Dim resultat As pfcls.IpfcFamilyTableColumn
    Dim messageErreur As String
    Dim symbole As String

    On Error Resume Next
    symbole = CStr(colonneNouvelle.Symbol)
    Err.Clear
    On Error GoTo 0

    Set resultat = TECH_EssayerAjouterColonneFamille( _
        membre, colonneNouvelle, messageErreur)
    If resultat Is Nothing And symbole <> "" Then
        ' Colonne deja presente (XToolkitNoChange) ou retour vide : relecture.
        Set resultat = TECH_ColonneFamille(membre, symbole)
    End If
    If resultat Is Nothing Then
        Err.Raise vbObjectError + 15281, , _
            "La colonne de table de famille n'a pas ete creee." & _
            IIf(symbole <> "", vbCrLf & "Symbole : " & symbole, "") & _
            IIf(messageErreur <> "", vbCrLf & messageErreur, "")
    End If
    Set TECH_AjouterColonneFamille = resultat
End Function

Private Function TECH_AjouterLigneFamille( _
    ByVal membre As pfcls.IpfcFamilyMember, _
    ByVal nomInstance As String) As pfcls.IpfcFamilyTableRow

    Dim resultat As pfcls.IpfcFamilyTableRow
    Dim messageNull As String
    Dim messageNothing As String

    ' AddRow(InstanceName, Values) : le second argument doit etre fourni.
    On Error Resume Next
    Err.Clear
    Set resultat = membre.AddRow(nomInstance, Null)
    If Err.Number <> 0 Then
        messageNull = CStr(Err.Number) & " " & Err.Description
        Err.Clear
        Set resultat = Nothing
        Set resultat = membre.AddRow(nomInstance, Nothing)
        If Err.Number <> 0 Then
            messageNothing = CStr(Err.Number) & " " & Err.Description
            Set resultat = Nothing
        End If
    End If
    Err.Clear
    On Error GoTo 0

    If resultat Is Nothing Then Set resultat = TECH_LigneFamille(membre, nomInstance)
    If resultat Is Nothing Then
        Err.Raise vbObjectError + 15283, , _
            "Impossible d'ajouter la ligne " & nomInstance & _
            " a la table de famille." & vbCrLf & _
            "AddRow(Null) : " & messageNull & vbCrLf & _
            "AddRow(Nothing) : " & messageNothing & vbCrLf & _
            "Si l'erreur cite XToolkitAbort ou XToolkitFound, un modele " & _
            "portant ce nom existe deja (fichier ou session Creo)."
    End If
    Set TECH_AjouterLigneFamille = resultat
End Function

Private Function TECH_TrouverColonneFamille( _
    ByVal membre As pfcls.IpfcFamilyMember, _
    ByVal symbolesPrioritaires As Variant, _
    ByVal typeElement As Long, _
    ByVal idElement As Long, _
    ByVal symbolesSecondaires As Variant) As pfcls.IpfcFamilyTableColumn

    Dim symbole As Variant
    Dim colonnes As pfcls.IpfcFamilyTableColumns
    Dim colonne As pfcls.IpfcFamilyTableColumn
    Dim colonneItem As pfcls.IpfcFamColModelItem
    Dim reference As pfcls.IpfcModelItem
    Dim trouveParSymbole As pfcls.IpfcFamilyTableColumn
    Dim symboleColonne As String
    Dim typeReference As Long
    Dim idReference As Long
    Dim lectureOK As Boolean
    Dim nb As Long
    Dim i As Long

    ' 1. Symboles generes par l'API (d#, m#).
    For Each symbole In symbolesPrioritaires
        If Trim$(CStr(symbole)) <> "" Then
            Set colonne = TECH_ColonneFamille(membre, Trim$(CStr(symbole)))
            If Not colonne Is Nothing Then
                Set TECH_TrouverColonneFamille = colonne
                Exit Function
            End If
        End If
    Next symbole

    ' 2. Parcours des colonnes : element reference (type + ID), puis
    '    symbole secondaire (cote renommee FOOT_HEIGHT, colonne mm#...).
    On Error Resume Next
    Err.Clear
    Set colonnes = membre.ListColumns()
    nb = 0
    If Not colonnes Is Nothing Then nb = colonnes.Count
    Err.Clear
    For i = 0 To nb - 1
        Set colonne = Nothing
        Set colonneItem = Nothing
        Set reference = Nothing
        Set colonne = colonnes.item(i)
        Err.Clear
        If Not colonne Is Nothing Then
            Set colonneItem = colonne
            If Err.Number <> 0 Then Set colonneItem = Nothing
            Err.Clear
            If Not colonneItem Is Nothing Then
                Set reference = colonneItem.RefItem
                Err.Clear
            End If
            If Not reference Is Nothing Then
                typeReference = -1
                idReference = -1
                typeReference = CLng(reference.Type)
                idReference = CLng(reference.id)
                lectureOK = (Err.Number = 0)
                Err.Clear
                If lectureOK Then
                    If typeReference = typeElement And idReference = idElement Then
                        Set TECH_TrouverColonneFamille = colonne
                        On Error GoTo 0
                        Exit Function
                    End If
                End If
            End If
            If trouveParSymbole Is Nothing Then
                symboleColonne = ""
                symboleColonne = Trim$(CStr(colonne.Symbol))
                Err.Clear
                If symboleColonne <> "" Then
                    For Each symbole In symbolesSecondaires
                        If Trim$(CStr(symbole)) <> "" Then
                            If StrComp(symboleColonne, Trim$(CStr(symbole)), _
                                vbTextCompare) = 0 Then
                                Set trouveParSymbole = colonne
                            End If
                        End If
                    Next symbole
                End If
            End If
        End If
        Err.Clear
    Next i
    Err.Clear
    On Error GoTo 0
    Set TECH_TrouverColonneFamille = trouveParSymbole
End Function

Private Function TECH_CreerColonneHauteurPied( _
    ByVal membre As pfcls.IpfcFamilyMember, _
    ByVal modeleSource As pfcls.IpfcModel, _
    ByVal dimension As pfcls.IpfcDimension, _
    ByVal idDimension As Long, _
    ByVal symboleDimension As String) As pfcls.IpfcFamilyTableColumn

    Dim colonneDimension As pfcls.IpfcFamColDimension
    Dim resultat As pfcls.IpfcFamilyTableColumn
    Dim messageCreation As String
    Dim messageAjout As String

    ' Colonne de COTE FOOT_HEIGHT (liaison precoce).
    On Error Resume Next
    Err.Clear
    Set colonneDimension = membre.CreateDimensionColumn(dimension)
    If Err.Number <> 0 Then
        messageCreation = CStr(Err.Number) & " " & Err.Description
        Set colonneDimension = Nothing
    End If
    Err.Clear
    On Error GoTo 0

    If Not colonneDimension Is Nothing Then
        Set resultat = TECH_EssayerAjouterColonneFamille( _
            membre, colonneDimension, messageAjout)
    End If
    If resultat Is Nothing Then
        ' Colonne peut-etre deja presente : relecture.
        Set resultat = TECH_TrouverColonneFamille(membre, _
            Array("d" & CStr(idDimension), "D" & CStr(idDimension)), _
            pfcls.EpfcITEM_DIMENSION, idDimension, _
            Array(symboleDimension, TECH_PIED_DIM))
    End If
    If Not resultat Is Nothing Then
        Set TECH_CreerColonneHauteurPied = resultat
        Exit Function
    End If

    Err.Raise vbObjectError + 15284, , _
        "Impossible de creer la colonne de cote " & TECH_PIED_DIM & _
        " (d" & CStr(idDimension) & ") dans la table du pied " & _
        modeleSource.Filename & "." & vbCrLf & _
        IIf(messageCreation <> "", "CreateDimensionColumn : " & _
            messageCreation & vbCrLf, "") & _
        IIf(messageAjout <> "", messageAjout & vbCrLf, "") & _
        "Verifier que la cote n'est pas pilotee par une relation " & _
        "et que la piece generique est modifiable."
End Function

Private Function TECH_CelluleParDefaut( _
    ByVal membre As pfcls.IpfcFamilyMember, _
    ByVal colonne As pfcls.IpfcFamilyTableColumn, _
    ByVal ligne As pfcls.IpfcFamilyTableRow) As Boolean

    Dim resultat As Boolean

    On Error Resume Next
    Err.Clear
    resultat = CBool(membre.GetCellIsDefault(colonne, ligne))
    ' Cellule illisible : on la renseigne ; SetCell signalera un vrai refus.
    If Err.Number <> 0 Then resultat = True
    Err.Clear
    On Error GoTo 0
    TECH_CelluleParDefaut = resultat
End Function

Private Function TECH_ValeurNumeriqueParam( _
    ByVal valeur As pfcls.IpfcParamValue) As Double

    Dim valeurReelle As Double
    Dim valeurEntiere As Double
    Dim reelleOK As Boolean
    Dim entiereOK As Boolean

    If valeur Is Nothing Then
        Err.Raise vbObjectError + 15285, , "Cellule de table de famille vide."
    End If
    On Error Resume Next
    Err.Clear
    valeurReelle = CDbl(valeur.DoubleValue)
    reelleOK = (Err.Number = 0)
    Err.Clear
    valeurEntiere = CDbl(valeur.IntValue)
    entiereOK = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    If reelleOK And (valeurReelle <> 0# Or Not entiereOK) Then
        TECH_ValeurNumeriqueParam = valeurReelle
    ElseIf entiereOK Then
        TECH_ValeurNumeriqueParam = valeurEntiere
    Else
        Err.Raise vbObjectError + 15285, , _
            "Valeur de cellule non numerique dans la table de famille."
    End If
End Function

Private Function TECH_CheminFichierModele( _
    ByVal modele As pfcls.IpfcModel) As String

    Dim dossier As String
    Dim nomLogique As String
    Dim fichier As String
    Dim meilleur As String
    Dim meilleurNumero As Long
    Dim numero As Long
    Dim suffixe As String
    Dim position As Long

    ' Fichier disque le plus recent (nom.prt ou nom.prt.N) du dossier d'origine.
    If modele Is Nothing Then Exit Function
    On Error Resume Next
    dossier = CStr(modele.Origin)
    nomLogique = TECH_NomModeleLogique(modele.Filename)
    Err.Clear
    dossier = Replace(Trim$(dossier), "/", "\")
    If dossier = "" Or nomLogique = "" Then Exit Function
    position = InStr(1, dossier, nomLogique, vbTextCompare)
    If position > 0 Then dossier = Left$(dossier, position - 1)
    If Right$(dossier, 1) <> "\" Then dossier = dossier & "\"

    meilleurNumero = -1
    fichier = Dir$(dossier & nomLogique & "*")
    Do While fichier <> ""
        suffixe = Mid$(fichier, Len(nomLogique) + 1)
        numero = -1
        If suffixe = "" Then
            numero = 0
        ElseIf Left$(suffixe, 1) = "." And IsNumeric(Mid$(suffixe, 2)) Then
            numero = CLng(Val(Mid$(suffixe, 2)))
        End If
        If numero > meilleurNumero Then
            meilleurNumero = numero
            meilleur = fichier
        End If
        fichier = Dir$()
    Loop
    Err.Clear
    On Error GoTo 0
    If meilleur <> "" Then TECH_CheminFichierModele = dossier & meilleur
End Function

Private Function TECH_EtatEcriture( _
    ByVal modele As pfcls.IpfcModel) As String

    Dim texte As String
    Dim valeur As String
    Dim fichier As String
    Dim attributs As Long

    ' Etat d'ecriture d'un modele, pour les messages d'erreur et le
    ' diagnostic. Ne declenche jamais d'erreur.
    If modele Is Nothing Then Exit Function
    On Error Resume Next
    texte = "Modele : " & modele.Filename
    Err.Clear
    valeur = ""
    valeur = CStr(modele.Origin)
    If Err.Number <> 0 Then valeur = "illisible"
    Err.Clear
    texte = texte & vbCrLf & "Origine Creo : " & valeur

    valeur = ""
    valeur = IIf(modele.CheckIsModifiable(False), "OUI", "NON")
    If Err.Number <> 0 Then valeur = "illisible (" & Err.Description & ")"
    Err.Clear
    texte = texte & vbCrLf & "Modifiable (CheckIsModifiable) : " & valeur

    valeur = ""
    valeur = IIf(modele.CheckIsSaveAllowed(False), "OUI", "NON")
    If Err.Number <> 0 Then valeur = "illisible (" & Err.Description & ")"
    Err.Clear
    texte = texte & vbCrLf & "Sauvegarde autorisee (CheckIsSaveAllowed) : " & valeur

    fichier = TECH_CheminFichierModele(modele)
    If fichier = "" Then
        texte = texte & vbCrLf & "Fichier disque : introuvable"
    Else
        attributs = GetAttr(fichier)
        If Err.Number <> 0 Then
            texte = texte & vbCrLf & "Fichier disque : " & fichier & " (attributs illisibles)"
        ElseIf (attributs And vbReadOnly) <> 0 Then
            texte = texte & vbCrLf & "Fichier disque : " & fichier & _
                " -> LECTURE SEULE (Windows)"
        Else
            texte = texte & vbCrLf & "Fichier disque : " & fichier & " (ecriture permise)"
        End If
    End If
    Err.Clear
    On Error GoTo 0
    TECH_EtatEcriture = texte
End Function

Private Sub TECH_SauvegarderModele( _
    ByVal modele As pfcls.IpfcModel, _
    ByVal contexte As String)

    Dim numeroErreur As Long
    Dim descriptionErreur As String
    Dim sourceErreur As String

    ' V17 : pas de pre-controle CheckIsSaveAllowed. PTC precise que ce
    ' controle depend du contexte d'interface Creo du moment et peut
    ' renvoyer Faux sans raison durable. On sauvegarde ; si Creo refuse,
    ' l'erreur reelle est remontee avec l'etat d'ecriture du modele.
    On Error Resume Next
    Err.Clear
    modele.Save
    numeroErreur = Err.Number
    descriptionErreur = Err.Description
    sourceErreur = Err.Source
    Err.Clear
    On Error GoTo 0
    If numeroErreur <> 0 Then
        Err.Raise numeroErreur, sourceErreur, _
            contexte & " : Creo refuse la sauvegarde." & vbCrLf & _
            descriptionErreur & vbCrLf & TECH_EtatEcriture(modele)
    End If
End Sub

Private Function TECH_DiagRelationCote( _
    ByVal modele As pfcls.IpfcModel, _
    ByVal nomCote As String, _
    ByRef pilotee As Boolean) As String

    Dim proprietaire As pfcls.IpfcRelationOwner
    Dim relations As pfcls.Istringseq
    Dim ligne As String
    Dim compacte As String
    Dim resultat As String
    Dim nb As Long
    Dim i As Long

    pilotee = False
    On Error Resume Next
    Err.Clear
    Set proprietaire = modele
    Set relations = proprietaire.Relations
    If Err.Number <> 0 Then
        TECH_DiagRelationCote = "relations illisibles : " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    nb = 0
    If Not relations Is Nothing Then nb = relations.Count
    For i = 0 To nb - 1
        ligne = CStr(relations.item(i))
        compacte = UCase$(Replace(Replace(ligne, " ", ""), vbTab, ""))
        If Left$(compacte, Len(nomCote) + 1) = UCase$(nomCote) & "=" Then
            pilotee = True
            resultat = "PILOTEE PAR RELATION : " & Trim$(ligne)
        End If
    Next i
    Err.Clear
    On Error GoTo 0
    If resultat = "" Then
        resultat = "aucune relation n'impose " & nomCote & _
            " (" & CStr(nb) & " ligne(s) de relations)"
    End If
    TECH_DiagRelationCote = resultat
End Function

Private Function TECH_DiagListerColonnes( _
    ByVal membre As pfcls.IpfcFamilyMember) As String

    Dim colonnes As pfcls.IpfcFamilyTableColumns
    Dim texte As String
    Dim i As Long

    On Error Resume Next
    Set colonnes = membre.ListColumns()
    If Not colonnes Is Nothing Then
        For i = 0 To colonnes.Count - 1
            If texte <> "" Then texte = texte & " ; "
            texte = texte & CStr(colonnes.item(i).Symbol)
        Next i
    End If
    If Err.Number <> 0 Then texte = texte & " (lecture interrompue : " & _
        Err.Description & ")"
    Err.Clear
    On Error GoTo 0
    If texte = "" Then texte = "aucune colonne (table vide)"
    TECH_DiagListerColonnes = texte
End Function

Private Function TECH_DiagListerLignes( _
    ByVal membre As pfcls.IpfcFamilyMember) As String

    Dim lignes As pfcls.IpfcFamilyTableRows
    Dim texte As String
    Dim i As Long

    On Error Resume Next
    Set lignes = membre.ListRows()
    If Not lignes Is Nothing Then
        For i = 0 To lignes.Count - 1
            If texte <> "" Then texte = texte & " ; "
            texte = texte & CStr(lignes.item(i).InstanceName)
        Next i
    End If
    If Err.Number <> 0 Then texte = texte & " (lecture interrompue : " & _
        Err.Description & ")"
    Err.Clear
    On Error GoTo 0
    If texte = "" Then texte = "aucune instance"
    TECH_DiagListerLignes = texte
End Function

Private Sub TECH_DiagAjouter( _
    ByVal rapport As Collection, _
    ByVal zone As String, _
    ByVal objet As String, _
    ByVal controle As String, _
    ByVal resultat As String, _
    ByVal detail As String)

    rapport.Add Array(zone, objet, controle, resultat, _
        Replace(detail, vbCrLf, " | "))
End Sub

Public Sub Diagnostic_Pieces_Techniques()
    Dim cConnexion As pfcls.CCpfcAsyncConnection
    Dim connexion As pfcls.IpfcAsyncConnection
    Dim session As pfcls.IpfcBaseSession
    Dim connexionCreee As Boolean
    Dim modeleRacine As pfcls.IpfcModel
    Dim dossierPieces As String
    Dim modeleMoteur As pfcls.IpfcModel
    Dim wsEtat As Worksheet
    Dim visites As Object
    Dim rapport As Collection
    Dim n As Long
    Dim d As String
    Dim estAssemblage As Boolean

    ' Controle complet SANS modification des modeles Creo ni de l'etat
    ' __HUB_ETAT. Charge seulement les modeles necessaires en session.
    ' Resultat : feuille DIAG_PIECES_TECH.
    Set rapport = New Collection
    m_TECH_MoteurFichier = TECH_MOTEUR_FICHIER
    m_TECH_MoteurDossier = ""
    m_TECH_LecteurReseauTemp = ""
    m_TECH_PartageReseauTemp = ""
    m_TECH_DossierCreoInitial = ""
    TECH_DiagAjouter rapport, "Macro", "Version", "Version du module", "OK", TECH_VERSION

    On Error Resume Next
    Err.Clear
    TECH_ConnecterCreo cConnexion, connexion, session, connexionCreee
    n = Err.Number: d = Err.Description: Err.Clear
    On Error GoTo 0
    If n <> 0 Or session Is Nothing Then
        TECH_DiagAjouter rapport, "Creo", "Session", "Connexion", "KO", d
        GoTo Fin
    End If
    TECH_DiagAjouter rapport, "Creo", "Session", "Connexion", "OK", ""

    On Error Resume Next
    m_TECH_DossierCreoInitial = session.GetCurrentDirectory()
    Set modeleRacine = session.GetActiveModel()
    estAssemblage = False
    If Not modeleRacine Is Nothing Then _
        estAssemblage = (modeleRacine.Type = pfcls.EpfcMDL_ASSEMBLY)
    Err.Clear
    On Error GoTo 0
    If Not estAssemblage Then
        TECH_DiagAjouter rapport, "Creo", "Modele actif", _
            "Assemblage principal actif", "KO", _
            "Ouvrir et activer l'assemblage principal dans Creo."
        GoTo Fin
    End If
    TECH_DiagAjouter rapport, "Creo", modeleRacine.Filename, _
        "Assemblage principal actif", "OK", ""

    On Error Resume Next
    Err.Clear
    dossierPieces = TECH_ObtenirDossierPieces(session)
    n = Err.Number: d = Err.Description: Err.Clear
    On Error GoTo 0
    TECH_DiagAjouter rapport, "Excel", "Dossier de travail", "Dossier des pieces", _
        IIf(n = 0, "OK", "KO"), IIf(n = 0, dossierPieces, d)

    ' --- Assemblage moteur et pied --------------------------------------
    On Error Resume Next
    Err.Clear
    TECH_ResoudreFichierMoteur dossierPieces, m_TECH_MoteurFichier, _
        m_TECH_MoteurDossier
    n = Err.Number: d = Err.Description: Err.Clear
    If n = 0 Then
        Set modeleMoteur = TECH_ChargerModele(session, m_TECH_MoteurFichier, _
            m_TECH_MoteurDossier)
        n = Err.Number: d = Err.Description: Err.Clear
    End If
    On Error GoTo 0
    If n <> 0 Or modeleMoteur Is Nothing Then
        TECH_DiagAjouter rapport, "Moteur", m_TECH_MoteurFichier, _
            "Chargement", "KO", d
    Else
        TECH_DiagAjouter rapport, "Moteur", modeleMoteur.Filename, _
            "Chargement", "OK", m_TECH_MoteurDossier
        On Error Resume Next
        Err.Clear
        TECH_DiagMoteurEtPied rapport, session, modeleMoteur
        If Err.Number <> 0 Then TECH_DiagAjouter rapport, "Moteur", _
            modeleMoteur.Filename, "Controle interrompu", "KO", _
            CStr(Err.Number) & " " & Err.Description
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Sous-assemblages, squelettes, rouleaux -------------------------
    Set wsEtat = TECH_ObtenirFeuilleEtat(False)
    If wsEtat Is Nothing Then
        TECH_DiagAjouter rapport, "Excel", TECH_FEUILLE_ETAT, _
            "Feuille d'etat", "ALERTE", _
            "Absente : lancer d'abord l'ajout des rouleaux."
    End If
    Set visites = CreateObject("Scripting.Dictionary")
    visites.CompareMode = vbTextCompare
    On Error Resume Next
    Err.Clear
    TECH_DiagAssemblage rapport, session, modeleRacine, wsEtat, visites
    If Err.Number <> 0 Then TECH_DiagAjouter rapport, "Assemblages", _
        m_TECH_Etape, "Parcours interrompu", "KO", _
        CStr(Err.Number) & " " & Err.Description
    Err.Clear
    On Error GoTo 0

Fin:
    On Error Resume Next
    TECH_DiagEcrireRapport rapport
    If Not session Is Nothing Then TECH_LibererLecteurReseauTemp session
    If connexionCreee Then connexion.Disconnect 5
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TECH_DiagMoteurEtPied( _
    ByVal rapport As Collection, _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal modeleMoteur As pfcls.IpfcModel)

    Dim solidMoteur As pfcls.IpfcSolid
    Dim featurePied As pfcls.IpfcFeature
    Dim composantPied As pfcls.IpfcComponentFeat
    Dim modelePied As pfcls.IpfcModel
    Dim solidPied As pfcls.IpfcSolid
    Dim piedID As Long
    Dim dimensionBase As pfcls.IpfcBaseDimension
    Dim dimension As pfcls.IpfcDimension
    Dim itemDimension As pfcls.IpfcModelItem
    Dim membrePied As pfcls.IpfcFamilyMember
    Dim membreMoteur As pfcls.IpfcFamilyMember
    Dim colonne As pfcls.IpfcFamilyTableColumn
    Dim colonneDimension As pfcls.IpfcFamColDimension
    Dim colonneComposant As pfcls.IpfcFamColComp
    Dim texteRelation As String
    Dim pilotee As Boolean
    Dim n As Long
    Dim d As String
    Dim idDimension As Long
    Dim symbole As String
    Dim valeur As Double
    Dim nomPied As String

    Set solidMoteur = modeleMoteur
    TECH_DiagAjouter rapport, "Moteur", modeleMoteur.Filename, _
        "Repere " & TECH_MOTEUR_CSYS, _
        IIf(TECH_TrouverCsysExact(modeleMoteur, TECH_MOTEUR_CSYS) Is Nothing, _
            "KO", "OK"), ""
    TECH_DiagAjouter rapport, "Moteur", modeleMoteur.Filename, _
        "Etat d'ecriture", "INFO", TECH_EtatEcriture(modeleMoteur)

    On Error Resume Next
    Err.Clear
    TECH_TrouverComposantPied session, solidMoteur, featurePied, _
        composantPied, modelePied, piedID
    n = Err.Number: d = Err.Description: Err.Clear
    On Error GoTo 0
    If n <> 0 Or modelePied Is Nothing Then
        TECH_DiagAjouter rapport, "Pied", "?", "Composant pied dans le moteur", "KO", d
        Exit Sub
    End If
    nomPied = modelePied.Filename
    TECH_DiagAjouter rapport, "Pied", nomPied, "Composant pied dans le moteur", _
        "OK", "Feature ID " & CStr(piedID)
    TECH_DiagAjouter rapport, "Pied", nomPied, "Etat d'ecriture", "INFO", _
        TECH_EtatEcriture(modelePied)

    ' Cote FOOT_HEIGHT
    Set solidPied = modelePied
    Set dimensionBase = TECH_TrouverDimension(solidPied, TECH_PIED_DIM)
    If dimensionBase Is Nothing Then
        TECH_DiagAjouter rapport, "Pied", nomPied, "Cote " & TECH_PIED_DIM, "KO", _
            "Cote introuvable dans la piece."
        Exit Sub
    End If
    On Error Resume Next
    Err.Clear
    Set dimension = dimensionBase
    n = Err.Number: Err.Clear
    symbole = dimensionBase.Symbol
    valeur = dimensionBase.DimValue
    Set itemDimension = dimensionBase
    idDimension = itemDimension.id
    Err.Clear
    On Error GoTo 0
    If n <> 0 Or dimension Is Nothing Then
        TECH_DiagAjouter rapport, "Pied", nomPied, "Cote " & TECH_PIED_DIM, "KO", _
            "Cote de reference (non pilotable) : une table de famille ne peut pas la piloter."
        Exit Sub
    End If
    TECH_DiagAjouter rapport, "Pied", nomPied, "Cote " & TECH_PIED_DIM, "OK", _
        "Symbole " & symbole & " / d" & CStr(idDimension) & " / valeur " & _
        CStr(valeur) & " mm"

    texteRelation = TECH_DiagRelationCote(modelePied, symbole, pilotee)
    TECH_DiagAjouter rapport, "Pied", nomPied, "Relations sur la cote", _
        IIf(pilotee, "KO", "OK"), texteRelation

    ' Table de famille du pied
    On Error Resume Next
    Err.Clear
    Set membrePied = TECH_MembreFamilleDepuisModele(modelePied)
    n = Err.Number: d = Err.Description: Err.Clear
    On Error GoTo 0
    If membrePied Is Nothing Then
        TECH_DiagAjouter rapport, "Pied", nomPied, "Acces table de famille", "KO", d
    Else
        TECH_DiagAjouter rapport, "Pied", nomPied, "Colonnes existantes", "INFO", _
            TECH_DiagListerColonnes(membrePied)
        TECH_DiagAjouter rapport, "Pied", nomPied, "Instances existantes", "INFO", _
            TECH_DiagListerLignes(membrePied)
        Set colonne = TECH_TrouverColonneFamille(membrePied, _
            Array("d" & CStr(idDimension), "D" & CStr(idDimension)), _
            pfcls.EpfcITEM_DIMENSION, idDimension, Array(symbole, TECH_PIED_DIM))
        If Not colonne Is Nothing Then
            TECH_DiagAjouter rapport, "Pied", nomPied, "Colonne FOOT_HEIGHT", "OK", _
                "Deja presente : " & colonne.Symbol
        Else
            On Error Resume Next
            Err.Clear
            Set colonneDimension = membrePied.CreateDimensionColumn(dimension)
            n = Err.Number: d = Err.Description: Err.Clear
            On Error GoTo 0
            TECH_DiagAjouter rapport, "Pied", nomPied, _
                "Colonne FOOT_HEIGHT (test, non ajoutee)", _
                IIf(n = 0 And Not colonneDimension Is Nothing, "OK", "KO"), _
                IIf(n = 0, "Colonne absente ; elle pourra etre creee.", d)
        End If
    End If

    ' Table de famille du moteur
    On Error Resume Next
    Err.Clear
    Set membreMoteur = TECH_MembreFamilleDepuisModele(modeleMoteur)
    n = Err.Number: d = Err.Description: Err.Clear
    On Error GoTo 0
    If membreMoteur Is Nothing Then
        TECH_DiagAjouter rapport, "Moteur", modeleMoteur.Filename, _
            "Acces table de famille", "KO", d
        Exit Sub
    End If
    TECH_DiagAjouter rapport, "Moteur", modeleMoteur.Filename, _
        "Colonnes existantes", "INFO", TECH_DiagListerColonnes(membreMoteur)
    TECH_DiagAjouter rapport, "Moteur", modeleMoteur.Filename, _
        "Instances existantes", "INFO", TECH_DiagListerLignes(membreMoteur)
    Set colonne = TECH_TrouverColonneFamille(membreMoteur, _
        Array("m" & CStr(piedID), "M" & CStr(piedID)), _
        pfcls.EpfcITEM_FEATURE, piedID, _
        Array("mm" & CStr(piedID), "MM" & CStr(piedID)))
    If Not colonne Is Nothing Then
        TECH_DiagAjouter rapport, "Moteur", modeleMoteur.Filename, _
            "Colonne composant pied", "OK", "Deja presente : " & colonne.Symbol
    Else
        On Error Resume Next
        Err.Clear
        Set colonneComposant = membreMoteur.CreateComponentColumn(featurePied)
        n = Err.Number: d = Err.Description: Err.Clear
        On Error GoTo 0
        TECH_DiagAjouter rapport, "Moteur", modeleMoteur.Filename, _
            "Colonne composant pied (test, non ajoutee)", _
            IIf(n = 0 And Not colonneComposant Is Nothing, "OK", "KO"), _
            IIf(n = 0, "Colonne absente ; elle pourra etre creee.", d)
    End If
End Sub

Private Sub TECH_DiagAssemblage( _
    ByVal rapport As Collection, _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal modeleAssemblage As pfcls.IpfcModel, _
    ByVal wsEtat As Worksheet, _
    ByVal visites As Object)

    Dim nomAssemblage As String
    Dim asm As pfcls.IpfcAssembly
    Dim solidAsm As pfcls.IpfcSolid
    Dim features As pfcls.IpfcFeatures
    Dim feature As pfcls.IpfcFeature
    Dim item As pfcls.IpfcModelItem
    Dim composant As pfcls.IpfcComponentFeat
    Dim modeleComposant As pfcls.IpfcModel
    Dim modeleRoll As pfcls.IpfcModel
    Dim sousAssemblages As Collection
    Dim rollIDs As Collection
    Dim segments As Collection
    Dim nomSquelette As String
    Dim centre As Variant
    Dim ligneY As Double
    Dim footHeight As Double
    Dim ligneEtat As Long
    Dim nomRoll As String
    Dim n As Long
    Dim d As String
    Dim i As Long
    Dim rollID As Long

    If modeleAssemblage Is Nothing Then Exit Sub
    nomAssemblage = TECH_NomModeleLogique(modeleAssemblage.Filename)
    If visites.Exists(nomAssemblage) Then Exit Sub
    visites.Add nomAssemblage, True
    Set asm = modeleAssemblage
    Set solidAsm = asm
    Set sousAssemblages = New Collection
    Set rollIDs = New Collection

    Set features = solidAsm.ListFeaturesByType(False, pfcls.EpfcFEATTYPE_COMPONENT)
    If Not features Is Nothing Then
        For i = 0 To features.Count - 1
            Set feature = features.item(i)
            Set item = feature
            Set composant = feature
            Set modeleComposant = TECH_ModeleDepuisComposant(session, composant)
            If modeleComposant Is Nothing Then
                TECH_DiagAjouter rapport, nomAssemblage, "Feature " & CStr(item.id), _
                    "Composant charge", "ALERTE", "Modele du composant non charge."
            ElseIf TECH_EstModeleRoll(modeleComposant.Filename) Then
                rollIDs.Add CLng(item.id)
            ElseIf modeleComposant.Type = pfcls.EpfcMDL_ASSEMBLY Then
                If Not TECH_EstModeleMoteur(modeleComposant.Filename) Then _
                    sousAssemblages.Add modeleComposant
            End If
        Next i
    End If

    If rollIDs.Count > 0 Then
        TECH_DiagAjouter rapport, nomAssemblage, "Rouleaux", "Nombre", "INFO", _
            CStr(rollIDs.Count)
        On Error Resume Next
        Err.Clear
        Set segments = TECH_ListerLignesSquelette(session, asm, solidAsm, _
            nomAssemblage, nomSquelette)
        n = Err.Number: d = Err.Description: Err.Clear
        On Error GoTo 0
        If n <> 0 Or segments Is Nothing Then
            TECH_DiagAjouter rapport, nomAssemblage, "Squelette", _
                "Lignes de charpente", "KO", d
        Else
            TECH_DiagAjouter rapport, nomAssemblage, nomSquelette, _
                "Lignes de charpente", "OK", CStr(segments.Count) & " ligne(s)"
        End If

        For i = 1 To rollIDs.Count
            rollID = CLng(rollIDs.item(i))
            Set modeleRoll = TECH_ModeleComposantParID(session, solidAsm, rollID)
            nomRoll = "Rouleau ID " & CStr(rollID)
            ligneEtat = 0
            If Not wsEtat Is Nothing Then _
                ligneEtat = TECH_TrouverLigneEtatRoll(wsEtat, nomAssemblage, rollID)
            If ligneEtat > 0 Then
                nomRoll = nomRoll & " / " & _
                    Trim$(CStr(wsEtat.Cells(ligneEtat, TECH_COL_REPERE).value))
                TECH_DiagAjouter rapport, nomAssemblage, nomRoll, _
                    "Suivi __HUB_ETAT", "OK", "Ligne " & CStr(ligneEtat)
            Else
                TECH_DiagAjouter rapport, nomAssemblage, nomRoll, _
                    "Suivi __HUB_ETAT", "ALERTE", _
                    "Rouleau non suivi : aucun moteur ne sera place dessus."
            End If
            If Not modeleRoll Is Nothing Then
                TECH_DiagAjouter rapport, nomAssemblage, nomRoll, _
                    "Reperes Moteur_L / Moteur_R", _
                    IIf(TECH_TrouverCsysExact(modeleRoll, TECH_MOTEUR_GAUCHE) Is Nothing _
                        Or TECH_TrouverCsysExact(modeleRoll, TECH_MOTEUR_DROIT) Is Nothing, _
                        "KO", "OK"), ""
                On Error Resume Next
                Err.Clear
                centre = TECH_CentreRouleau(asm, rollID, modeleRoll)
                n = Err.Number: d = Err.Description: Err.Clear
                On Error GoTo 0
                If n <> 0 Then
                    TECH_DiagAjouter rapport, nomAssemblage, nomRoll, _
                        "Centre du rouleau", "KO", d
                ElseIf Not segments Is Nothing Then
                    If TECH_TrouverLigneSousCentre(centre, segments, ligneY) Then
                        On Error Resume Next
                        Err.Clear
                        footHeight = TECH_CalculerFootHeight( _
                            CDbl(centre(TECH_AXE_VERTICAL)), ligneY)
                        n = Err.Number: d = Err.Description: Err.Clear
                        On Error GoTo 0
                        TECH_DiagAjouter rapport, nomAssemblage, nomRoll, _
                            "Ligne sous le centre / FOOT_HEIGHT", _
                            IIf(n = 0, "OK", "KO"), _
                            "Centre Y " & Format$(centre(TECH_AXE_VERTICAL), "0.###") & _
                            " / ligne Y " & Format$(ligneY, "0.###") & _
                            IIf(n = 0, " / FOOT_HEIGHT " & Format$(footHeight, "0.###") & _
                                " mm -> " & TECH_NomVariantePied("", footHeight), " / " & d)
                    Else
                        TECH_DiagAjouter rapport, nomAssemblage, nomRoll, _
                            "Ligne sous le centre / FOOT_HEIGHT", "KO", _
                            "Aucune ligne sous le centre (" & _
                            Format$(centre(0), "0.###") & " ; " & _
                            Format$(centre(1), "0.###") & " ; " & _
                            Format$(centre(2), "0.###") & ")"
                    End If
                End If
            End If
        Next i
    End If

    For i = 1 To sousAssemblages.Count
        TECH_DiagAssemblage rapport, session, sousAssemblages.item(i), _
            wsEtat, visites
    Next i
End Sub

Private Sub TECH_DiagEcrireRapport(ByVal rapport As Collection)
    Dim ws As Worksheet
    Dim ligne As Long
    Dim element As Variant
    Dim nbKO As Long
    Dim nbAlerte As Long
    Dim k As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("DIAG_PIECES_TECH")
    Err.Clear
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = "DIAG_PIECES_TECH"
    End If
    ws.Visible = xlSheetVisible
    ws.Cells.Clear
    ws.Range("A1:E1").value = Array("Zone", "Objet", "Controle", "Resultat", "Detail")
    ws.Range("A1:E1").Font.Bold = True
    ligne = 2
    For Each element In rapport
        For k = 0 To 4
            ws.Cells(ligne, k + 1).value = CStr(element(k))
        Next k
        Select Case CStr(element(3))
            Case "KO"
                ws.Cells(ligne, 4).Interior.Color = RGB(255, 199, 206)
                nbKO = nbKO + 1
            Case "ALERTE"
                ws.Cells(ligne, 4).Interior.Color = RGB(255, 235, 156)
                nbAlerte = nbAlerte + 1
            Case "OK"
                ws.Cells(ligne, 4).Interior.Color = RGB(198, 239, 206)
        End Select
        ligne = ligne + 1
    Next element
    ws.Columns("A:D").AutoFit
    ws.Columns("E").ColumnWidth = 120
    ws.Columns("E").WrapText = True
    ws.Activate
    MsgBox "Diagnostic termine (aucune modification Creo)." & vbCrLf & _
        "KO : " & CStr(nbKO) & vbCrLf & "Alertes : " & CStr(nbAlerte) & _
        vbCrLf & vbCrLf & "Detail : feuille DIAG_PIECES_TECH.", _
        IIf(nbKO > 0, vbExclamation, vbInformation), "Diagnostic Pieces techniques"
End Sub

Private Function TECH_NomEffectif( _
    ByVal modele As pfcls.IpfcModel) As String

    Dim nomFichier As String
    Dim nomInstance As String
    Dim extension As String
    Dim positionPoint As Long

    ' Nom logique d'un modele, instance de famille comprise.
    ' Identique a FileName pour un modele ordinaire.
    If modele Is Nothing Then Exit Function
    nomFichier = TECH_NomModeleLogique(modele.Filename)
    TECH_NomEffectif = nomFichier
    On Error Resume Next
    nomInstance = Trim$(CStr(modele.InstanceName))
    Err.Clear
    On Error GoTo 0
    If nomInstance = "" Then Exit Function
    nomInstance = TECH_BaseSansExtension(nomInstance)
    If StrComp(TECH_BaseSansExtension(nomFichier), nomInstance, _
        vbTextCompare) = 0 Then Exit Function

    positionPoint = InStrRev(nomFichier, ".")
    If positionPoint > 0 Then
        extension = Mid$(nomFichier, positionPoint)
    ElseIf modele.Type = pfcls.EpfcMDL_ASSEMBLY Then
        extension = ".asm"
    Else
        extension = ".prt"
    End If
    TECH_NomEffectif = nomInstance & extension
End Function

Private Function TECH_LigneFamille( _
    ByVal membre As pfcls.IpfcFamilyMember, _
    ByVal nom As String) As pfcls.IpfcFamilyTableRow

    On Error Resume Next
    Set TECH_LigneFamille = membre.GetRow(nom)
    Err.Clear
    On Error GoTo 0
End Function

Private Function TECH_ColonneFamille( _
    ByVal membre As pfcls.IpfcFamilyMember, _
    ByVal nom As String) As pfcls.IpfcFamilyTableColumn

    On Error Resume Next
    Set TECH_ColonneFamille = membre.GetColumn(nom)
    Err.Clear
    On Error GoTo 0
End Function

Private Function TECH_CopierEtRetrouverModele( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal modeleSource As pfcls.IpfcModel, _
    ByVal nomBase As String, _
    ByVal nomFichier As String, _
    ByVal dossier As String) As pfcls.IpfcModel

    Dim copie As pfcls.IpfcModel
    Dim fichierPhysique As String
    Dim erreurCopieRetrieve As String
    Dim erreurCopieSimple As String

    ' Methode normale recommandee par l'API Creo.
    On Error Resume Next
    Err.Clear
    Set copie = modeleSource.CopyAndRetrieve(nomBase, Nothing)
    erreurCopieRetrieve = Err.source & " " & Err.description
    Err.Clear
    On Error GoTo 0
    If Not copie Is Nothing Then
        Set TECH_CopierEtRetrouverModele = copie
        Exit Function
    End If

    ' Certaines configurations creent le fichier avant de renvoyer une erreur.
    fichierPhysique = Dir$(dossier & nomFichier & "*")
    If fichierPhysique <> "" Then
        Set copie = TECH_ChargerModele(session, nomFichier, dossier, False)
        If Not copie Is Nothing Then
            Set TECH_CopierEtRetrouverModele = copie
            Exit Function
        End If
    End If

    ' Repli : copie disque, puis chargement explicite dans la session.
    On Error Resume Next
    Err.Clear
    Call modeleSource.Copy(nomBase, Nothing)
    erreurCopieSimple = Err.source & " " & Err.description
    Err.Clear
    On Error GoTo 0
    fichierPhysique = Dir$(dossier & nomFichier & "*")
    If fichierPhysique <> "" Then
        Set copie = TECH_ChargerModele(session, nomFichier, dossier, False)
    End If
    If Not copie Is Nothing Then
        Set TECH_CopierEtRetrouverModele = copie
        Exit Function
    End If

    Err.Raise vbObjectError + 15226, , _
        "Impossible de copier le modele vers le nom Creo : " & nomBase & _
        vbCrLf & "CopyAndRetrieve : " & erreurCopieRetrieve & _
        vbCrLf & "Copy : " & erreurCopieSimple
End Function

Private Sub TECH_TrouverComposantPied( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal solidAssemblage As pfcls.IpfcSolid, _
    ByRef featurePied As pfcls.IpfcFeature, _
    ByRef composantPied As pfcls.IpfcComponentFeat, _
    ByRef modelePied As pfcls.IpfcModel, _
    ByRef featureID As Long)

    Dim features As pfcls.IpfcFeatures
    Dim feature As pfcls.IpfcFeature
    Dim composant As pfcls.IpfcComponentFeat
    Dim modele As pfcls.IpfcModel
    Dim item As pfcls.IpfcModelItem
    Dim base As String
    Dim baseSource As String
    Dim prefixe As String
    Dim prefixeHistorique As String
    Dim solidCandidat As pfcls.IpfcSolid
    Dim dimensionPied As pfcls.IpfcBaseDimension
    Dim featureExacte As pfcls.IpfcFeature
    Dim composantExact As pfcls.IpfcComponentFeat
    Dim modeleExact As pfcls.IpfcModel
    Dim idExact As Long
    Dim nbExacts As Long
    Dim featureParCote As pfcls.IpfcFeature
    Dim composantParCote As pfcls.IpfcComponentFeat
    Dim modeleParCote As pfcls.IpfcModel
    Dim idParCote As Long
    Dim nbAvecCote As Long
    Dim listeAvecCote As String
    Dim i As Long

    Set featurePied = Nothing
    Set composantPied = Nothing
    Set modelePied = Nothing
    featureID = 0
    baseSource = UCase$(TECH_BaseSansExtension(TECH_PIED_FICHIER))
    prefixe = UCase$(TECH_PREFIXE_PIED)
    prefixeHistorique = UCase$(TECH_PREFIXE_PIED_HISTORIQUE)
    Set features = solidAssemblage.ListFeaturesByType( _
        False, pfcls.EpfcFEATTYPE_COMPONENT)
    If Not features Is Nothing Then
        For i = 0 To features.Count - 1
            Set feature = features.item(i)
            Set composant = Nothing
            Set modele = Nothing
            Set item = Nothing
            If Not feature Is Nothing Then Set composant = feature
            If Not feature Is Nothing Then Set item = feature
            If Not composant Is Nothing Then
                Set modele = TECH_ModeleDepuisComposant(session, composant)
                If Not modele Is Nothing Then
                    base = UCase$(TECH_BaseSansExtension(TECH_NomEffectif(modele)))
                    If base = baseSource Or _
                        Left$(base, Len(prefixe)) = prefixe Or _
                        Left$(base, Len(prefixeHistorique)) = _
                            prefixeHistorique Or _
                        Left$(base, Len(TECH_PREFIXE_PIED_COPIE)) = _
                            TECH_PREFIXE_PIED_COPIE Then
                        nbExacts = nbExacts + 1
                        Set featureExacte = feature
                        Set composantExact = composant
                        Set modeleExact = modele
                        idExact = CLng(item.id)
                    End If

                    ' Le nom physique peut varier. La vraie signature du pied
                    ' est la presence de la cote FOOT_HEIGHT dans une piece
                    ' directement contenue dans l'assemblage selectionne.
                    If modele.Type = pfcls.EpfcMDL_PART Then
                        Set solidCandidat = Nothing
                        Set dimensionPied = Nothing
                        Set solidCandidat = modele
                        If Not solidCandidat Is Nothing Then
                            Set dimensionPied = TECH_TrouverDimension( _
                                solidCandidat, TECH_PIED_DIM)
                        End If
                        If Not dimensionPied Is Nothing Then
                            nbAvecCote = nbAvecCote + 1
                            Set featureParCote = feature
                            Set composantParCote = composant
                            Set modeleParCote = modele
                            idParCote = CLng(item.id)
                            If listeAvecCote <> "" Then _
                                listeAvecCote = listeAvecCote & vbCrLf
                            listeAvecCote = listeAvecCote & "- " & _
                                TECH_NomModeleLogique(modele.Filename)
                        End If
                    End If
                End If
            End If
        Next i
    End If

    If nbExacts = 1 Then
        Set featurePied = featureExacte
        Set composantPied = composantExact
        Set modelePied = modeleExact
        featureID = idExact
        Exit Sub
    ElseIf nbExacts > 1 Then
        Err.Raise vbObjectError + 15231, , _
            "Plusieurs composants portent le nom du pied dans " & _
            "l'assemblage moteur." & vbCrLf & _
            "Un seul composant pilotable est attendu."
    End If

    If nbAvecCote = 1 Then
        Set featurePied = featureParCote
        Set composantPied = composantParCote
        Set modelePied = modeleParCote
        featureID = idParCote
        Exit Sub
    ElseIf nbAvecCote = 0 Then
        Err.Raise vbObjectError + 15230, , _
            "Aucune piece contenant la cote " & TECH_PIED_DIM & _
            " n'a ete trouvee directement dans l'assemblage moteur." & _
            vbCrLf & "La recherche ne depend plus du nom du fichier." & _
            vbCrLf & "Assemblage : " & m_TECH_MoteurFichier
    Else
        Err.Raise vbObjectError + 15231, , _
            "Plusieurs pieces de l'assemblage contiennent la cote " & _
            TECH_PIED_DIM & "." & vbCrLf & _
            "La macro ne peut pas choisir automatiquement :" & vbCrLf & _
            listeAvecCote
    End If
End Sub

Private Sub TECH_ValiderComposant( _
    ByVal composant As pfcls.IpfcComponentFeat, _
    ByVal contexte As String)

    If composant Is Nothing Then
        Err.Raise vbObjectError + 15236, , _
            contexte & " : composant Nothing."
    End If
    If Not composant.IsPlaced Then
        Err.Raise vbObjectError + 15237, , _
            contexte & " : composant non place."
    End If
    If composant.IsPackaged Then
        Err.Raise vbObjectError + 15238, , _
            contexte & " : composant packaged."
    End If
    If composant.IsFrozen Then
        Err.Raise vbObjectError + 15239, , _
            contexte & " : composant frozen."
    End If
End Sub

Private Function TECH_AjouterMoteur( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal asm As pfcls.IpfcAssembly, _
    ByVal solidAsm As pfcls.IpfcSolid, _
    ByVal solidMoteur As pfcls.IpfcSolid, _
    ByVal rollFeatureID As Long, _
    ByVal coteMoteur As String) As Long

    Dim featureRoll As pfcls.IpfcFeature
    Dim composantRoll As pfcls.IpfcComponentFeat
    Dim modeleRoll As pfcls.IpfcModel
    Dim modeleMoteur As pfcls.IpfcModel
    Dim csysRoll As pfcls.IpfcModelItem
    Dim csysMoteur As pfcls.IpfcModelItem
    Dim pathRoll As pfcls.IpfcComponentPath
    Dim pathMoteur As pfcls.IpfcComponentPath
    Dim pathAsm As pfcls.IpfcComponentPath
    Dim selectionRoll As pfcls.IpfcSelection
    Dim selectionMoteur As pfcls.IpfcSelection
    Dim featureMoteur As pfcls.IpfcFeature
    Dim composantMoteur As pfcls.IpfcComponentFeat
    Dim itemMoteur As pfcls.IpfcModelItem
    Dim contraintes As pfcls.IpfcComponentConstraints
    Dim nouvelID As Long
    Dim numeroErreur As Long
    Dim descriptionErreur As String

    On Error GoTo Erreur
    Set featureRoll = solidAsm.GetFeatureById(rollFeatureID)
    If featureRoll Is Nothing Then
        Err.Raise vbObjectError + 15150, , _
            "Rouleau introuvable. Feature ID : " & CStr(rollFeatureID)
    End If
    Set composantRoll = featureRoll
    If composantRoll Is Nothing Then
        Err.Raise vbObjectError + 15151, , _
            "La feature du rouleau n'est pas un composant."
    End If
    Set modeleRoll = TECH_ModeleDepuisComposant(session, composantRoll)
    If modeleRoll Is Nothing Then
        Err.Raise vbObjectError + 15152, , "Modele du rouleau inaccessible."
    End If
    Set csysRoll = TECH_TrouverCsysExact(modeleRoll, coteMoteur)
    If csysRoll Is Nothing Then
        Err.Raise vbObjectError + 15153, , _
            "Le repere " & coteMoteur & " est absent du rouleau."
    End If
    Set modeleMoteur = solidMoteur
    Set csysMoteur = TECH_TrouverCsysExact(modeleMoteur, TECH_MOTEUR_CSYS)
    If csysMoteur Is Nothing Then
        Err.Raise vbObjectError + 15154, , _
            "Le repere " & TECH_MOTEUR_CSYS & " est absent de l'assemblage moteur."
    End If

    Set pathRoll = TECH_CreerPath(asm, rollFeatureID)
    Set selectionRoll = TECH_CreerSelectionCsys(csysRoll, pathRoll)
    Set featureMoteur = asm.AssembleComponent(modeleMoteur, Null)
    If featureMoteur Is Nothing Then
        Err.Raise vbObjectError + 15156, , _
            "AssembleComponent n'a retourne aucune feature moteur."
    End If
    Set composantMoteur = featureMoteur
    Set itemMoteur = featureMoteur
    nouvelID = itemMoteur.id
    Set pathMoteur = TECH_CreerPath(asm, nouvelID)
    Set pathAsm = TECH_CreerPathRacine(asm)
    Set selectionMoteur = TECH_CreerSelectionCsys(csysMoteur, pathMoteur)
    Set contraintes = TECH_CreerContrainteCsys(selectionRoll, selectionMoteur)
    composantMoteur.SetConstraints contraintes, pathAsm
    composantMoteur.Regenerate
    TECH_RegenererSolideFiable solidAsm, "Assemblage apres ajout du moteur"

    If Not composantMoteur.IsPlaced Or composantMoteur.IsPackaged Or _
        composantMoteur.IsUnderconstrained Or composantMoteur.IsFrozen Then
        Err.Raise vbObjectError + 15155, , "Le moteur ajoute n'est pas place correctement."
    End If
    TECH_AjouterMoteur = nouvelID
    Exit Function

Erreur:
    numeroErreur = Err.Number
    descriptionErreur = Err.description
    On Error GoTo -1 ' V14 : libere le gestionnaire actif avant nettoyage
    On Error Resume Next
    If nouvelID > 0 Then TECH_SupprimerComposant session, asm, solidAsm, nouvelID
    On Error GoTo 0
    Err.Raise numeroErreur, , "TECH_AjouterMoteur :" & vbCrLf & descriptionErreur
End Function

Private Function TECH_AjouterRoll( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal asm As pfcls.IpfcAssembly, _
    ByVal solidAsm As pfcls.IpfcSolid, _
    ByVal solidRoll As pfcls.IpfcSolid, _
    ByVal sourceFeatureID As Long, _
    ByVal nomRepereSource As String) As Long

    Dim featureSource As pfcls.IpfcFeature
    Dim composantSource As pfcls.IpfcComponentFeat
    Dim modeleSource As pfcls.IpfcModel
    Dim csysSource As pfcls.IpfcModelItem
    Dim pathSource As pfcls.IpfcComponentPath
    Dim selectionSource As pfcls.IpfcSelection
    Dim modeleRoll As pfcls.IpfcModel
    Dim csysRoll As pfcls.IpfcModelItem
    Dim featureRoll As pfcls.IpfcFeature
    Dim composantRoll As pfcls.IpfcComponentFeat
    Dim itemRoll As pfcls.IpfcModelItem
    Dim pathRoll As pfcls.IpfcComponentPath
    Dim pathAsm As pfcls.IpfcComponentPath
    Dim selectionRoll As pfcls.IpfcSelection
    Dim contraintes As pfcls.IpfcComponentConstraints
    Dim nouvelID As Long
    Dim errNumero As Long
    Dim errDescription As String

    On Error GoTo Erreur

    Set featureSource = solidAsm.GetFeatureById(sourceFeatureID)
    If featureSource Is Nothing Then
        Err.Raise vbObjectError + 15040, , _
            "La piece porteuse a disparu. Feature ID : " & CStr(sourceFeatureID)
    End If
    Set composantSource = featureSource
    If composantSource Is Nothing Then
        Err.Raise vbObjectError + 15041, , _
            "La feature porteuse n'est pas un composant. ID : " & CStr(sourceFeatureID)
    End If
    Set modeleSource = TECH_ModeleDepuisComposant(session, composantSource)
    If modeleSource Is Nothing Then
        Err.Raise vbObjectError + 15042, , "Modele de la piece porteuse introuvable."
    End If
    Set csysSource = TECH_TrouverCsysExact(modeleSource, nomRepereSource)
    If csysSource Is Nothing Then
        Err.Raise vbObjectError + 15043, , _
            "Le repere " & nomRepereSource & " a disparu de " & modeleSource.Filename & "."
    End If

    Set pathSource = TECH_CreerPath(asm, sourceFeatureID)
    Set selectionSource = TECH_CreerSelectionCsys(csysSource, pathSource)

    Set modeleRoll = solidRoll
    Set csysRoll = TECH_TrouverCsysExact(modeleRoll, TECH_ROLL_CSYS)
    If csysRoll Is Nothing Then
        Err.Raise vbObjectError + 15044, , _
            "Le repere " & TECH_ROLL_CSYS & " est absent de " & modeleRoll.Filename & "."
    End If

    Set featureRoll = asm.AssembleComponent(solidRoll, Null)
    If featureRoll Is Nothing Then
        Err.Raise vbObjectError + 15045, , "AssembleComponent n'a retourne aucune feature."
    End If
    Set composantRoll = featureRoll
    Set itemRoll = featureRoll
    nouvelID = itemRoll.id
    If nouvelID <= 0 Then
        Err.Raise vbObjectError + 15046, , "Feature ID invalide apres ajout du rouleau."
    End If

    Set pathRoll = TECH_CreerPath(asm, nouvelID)
    Set pathAsm = TECH_CreerPathRacine(asm)
    Set selectionRoll = TECH_CreerSelectionCsys(csysRoll, pathRoll)
    Set contraintes = TECH_CreerContrainteCsys(selectionSource, selectionRoll)
    composantRoll.SetConstraints contraintes, pathAsm
    composantRoll.Regenerate
    TECH_RegenererSolideFiable solidAsm, "Assemblage apres ajout du rouleau"

    If Not composantRoll.IsPlaced Then
        Err.Raise vbObjectError + 15047, , "Le rouleau ajoute n'est pas place."
    End If
    If composantRoll.IsPackaged Then
        Err.Raise vbObjectError + 15048, , "Le rouleau ajoute est encore packaged."
    End If
    If composantRoll.IsUnderconstrained Then
        Err.Raise vbObjectError + 15049, , "Le rouleau ajoute est sous-contraint."
    End If
    If composantRoll.IsFrozen Then
        Err.Raise vbObjectError + 15050, , "Le rouleau ajoute est frozen."
    End If

    TECH_AjouterRoll = nouvelID
    Exit Function

Erreur:
    errNumero = Err.Number
    errDescription = Err.description
    On Error GoTo -1 ' V14 : libere le gestionnaire actif avant nettoyage
    On Error Resume Next
    If Not featureRoll Is Nothing Then _
        TECH_SupprimerComposant session, asm, solidAsm, nouvelID
    On Error GoTo 0
    Err.Raise errNumero, , "TECH_AjouterRoll :" & vbCrLf & errDescription
End Function

Private Function TECH_FeatureComposantActive( _
    ByVal solidAsm As pfcls.IpfcSolid, _
    ByVal featureID As Long) As Boolean

    Dim features As pfcls.IpfcFeatures
    Dim feature As pfcls.IpfcFeature
    Dim item As pfcls.IpfcModelItem
    Dim i As Long

    TECH_FeatureComposantActive = False
    If featureID <= 0 Then Exit Function
    If solidAsm Is Nothing Then
        Err.Raise vbObjectError + 15057, , _
            "Assemblage inaccessible pendant le controle de suppression."
    End If

    Set features = solidAsm.ListFeaturesByType(False, pfcls.EpfcFEATTYPE_COMPONENT)
    If features Is Nothing Then
        Err.Raise vbObjectError + 15058, , _
            "Lecture des composants actifs impossible."
    End If
    For i = 0 To features.Count - 1
        Set feature = features.item(i)
        Set item = Nothing
        If Not feature Is Nothing Then Set item = feature
        If Not item Is Nothing Then
            If item.id = featureID Then
                TECH_FeatureComposantActive = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub TECH_ActiverModele( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal modele As pfcls.IpfcModel)

    Dim fenetre As pfcls.IpfcWindow
    Dim actif As pfcls.IpfcModel

    If session Is Nothing Or modele Is Nothing Then
        Err.Raise vbObjectError + 15059, , _
            "Modele ou session indisponible pour l'activation Creo."
    End If
    Set fenetre = Nothing
    On Error Resume Next
    Set fenetre = session.GetModelWindow(modele)
    Err.Clear
    On Error GoTo Erreur
    If fenetre Is Nothing Then
        Set fenetre = session.CreateModelWindow(modele)
        modele.Display
    End If
    If fenetre Is Nothing Then
        Err.Raise vbObjectError + 15067, , "Fenetre Creo indisponible."
    End If
    fenetre.Activate
    Set actif = session.CurrentModel
    If actif Is Nothing Then
        Err.Raise vbObjectError + 15068, , _
            "Aucun modele actif apres l'activation Creo."
    End If
    If StrComp(TECH_NomModeleLogique(actif.Filename), _
        TECH_NomModeleLogique(modele.Filename), vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 15069, , _
            "Activation du modele non confirmee. Attendu : " & modele.Filename & _
            " / Actif : " & actif.Filename
    End If
    Exit Sub

Erreur:
    Err.Raise Err.Number, , "TECH_ActiverModele :" & vbCrLf & Err.description
End Sub

Private Sub TECH_AttendreMs(ByVal millisecondes As Long)
    Dim depart As Double
    Dim ecoule As Double

    If millisecondes <= 0 Then Exit Sub
    depart = Timer
    Do
        DoEvents
        ecoule = (Timer - depart) * 1000#
        If ecoule < 0# Then ecoule = ecoule + 86400000#
    Loop While ecoule < CDbl(millisecondes)
End Sub

Private Sub TECH_SupprimerComposant( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal asm As pfcls.IpfcAssembly, _
    ByVal solidAsm As pfcls.IpfcSolid, _
    ByVal featureID As Long)

    Dim sessionUI As pfcls.IpfcSession
    Dim modeleAsm As pfcls.IpfcModel
    Dim cAssembly As pfcls.CMpfcAssembly
    Dim cSelect As pfcls.CMpfcSelect
    Dim idsPath As pfcls.Iintseq
    Dim pathComposant As pfcls.IpfcComponentPath
    Dim selectionComposant As pfcls.IpfcSelection
    Dim buffer As pfcls.IpfcSelectionBuffer
    Dim macroSuppression As String
    Dim tentative As Long
    Dim encoreActif As Boolean
    Dim numeroErreur As Long
    Dim descriptionErreur As String
    Dim nomAsmErreur As String

    If featureID <= 0 Then Exit Sub
    If session Is Nothing Then
        Err.Raise vbObjectError + 15060, , "Session Creo indisponible."
    End If
    If asm Is Nothing Or solidAsm Is Nothing Then
        Err.Raise vbObjectError + 15061, , "Assemblage Creo indisponible."
    End If
    If Not TECH_FeatureComposantActive(solidAsm, featureID) Then Exit Sub

    Set modeleAsm = asm
    TECH_ActiverModele session, modeleAsm
    Set sessionUI = session
    If sessionUI Is Nothing Then
        Err.Raise vbObjectError + 15062, , "Session interactive Creo indisponible."
    End If

    Set cAssembly = New pfcls.CMpfcAssembly
    Set cSelect = New pfcls.CMpfcSelect
    Set buffer = sessionUI.CurrentSelectionBuffer
    If buffer Is Nothing Then
        Err.Raise vbObjectError + 15063, , "SelectionBuffer Creo indisponible."
    End If

    Set idsPath = New pfcls.Cintseq
    idsPath.Append featureID
    Set pathComposant = cAssembly.CreateComponentPath(asm, idsPath)
    If pathComposant Is Nothing Then
        Err.Raise vbObjectError + 15064, , _
            "ComponentPath impossible pour la feature " & CStr(featureID) & "."
    End If
    Set selectionComposant = cSelect.CreateComponentSelection(pathComposant)
    If selectionComposant Is Nothing Then
        Err.Raise vbObjectError + 15065, , _
            "Selection impossible pour la feature " & CStr(featureID) & "."
    End If

    buffer.Clear
    buffer.AddSelection selectionComposant
    macroSuppression = "~ Command `ProCmdEditDelete`;" & _
        "~ Activate `del_sup_msg` `ok`;"
    m_TECH_Etape = "Suppression Creo dans " & modeleAsm.Filename & _
        " / Feature ID " & CStr(featureID)
    session.RunMacro macroSuppression
    On Error Resume Next
    session.FlushCurrentWindow
    Err.Clear
    On Error GoTo Erreur

    For tentative = 1 To 24
        DoEvents
        TECH_AttendreMs 150
        encoreActif = TECH_FeatureComposantActive(solidAsm, featureID)
        If Not encoreActif Then Exit For
    Next tentative

    On Error Resume Next
    buffer.Clear
    Err.Clear
    On Error GoTo Erreur
    If encoreActif Then
        Err.Raise vbObjectError + 15066, , _
            "Creo n'a pas supprime le composant actif."
    End If
    Exit Sub

Erreur:
    numeroErreur = Err.Number
    descriptionErreur = Err.description
    On Error GoTo -1 ' V14 : libere le gestionnaire actif avant nettoyage
    On Error Resume Next
    nomAsmErreur = "inconnu"
    If Not modeleAsm Is Nothing Then nomAsmErreur = modeleAsm.Filename
    If Not buffer Is Nothing Then buffer.Clear
    Err.Clear
    On Error GoTo 0
    Err.Raise numeroErreur, , _
        "TECH_SupprimerComposant :" & vbCrLf & _
        "Assemblage : " & nomAsmErreur & vbCrLf & _
        "Feature ID : " & CStr(featureID) & vbCrLf & descriptionErreur
End Sub

Private Sub TECH_SupprimerRollSuivi( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal asm As pfcls.IpfcAssembly, _
    ByVal solidAsm As pfcls.IpfcSolid, _
    ByVal featureID As Long)

    Dim modele As pfcls.IpfcModel
    Set modele = TECH_ModeleComposantParID(session, solidAsm, featureID)
    If modele Is Nothing Then Exit Sub
    If Not TECH_EstModeleRoll(modele.Filename) Then
        Err.Raise vbObjectError + 15062, , _
            "Suppression refusee : la feature " & CStr(featureID) & _
            " pointe vers " & modele.Filename & " et non vers un rouleau suivi."
    End If
    TECH_SupprimerComposant session, asm, solidAsm, featureID
End Sub

Private Sub TECH_SupprimerMoteurEtat( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal wsEtat As Worksheet, _
    ByVal ligneEtat As Long, _
    ByVal asm As pfcls.IpfcAssembly, _
    ByVal solidAsm As pfcls.IpfcSolid)

    Dim moteurID As Long
    Dim modele As pfcls.IpfcModel

    If ligneEtat <= 0 Then Exit Sub
    moteurID = CLng(Val(CStr(wsEtat.Cells(ligneEtat, TECH_COL_MOTEUR_ID).value)))
    If moteurID > 0 And TECH_FeatureComposantActive(solidAsm, moteurID) Then
        Set modele = TECH_ModeleComposantParID(session, solidAsm, moteurID)
        If Not modele Is Nothing Then
            If TECH_EstModeleMoteur(modele.Filename) Then
                TECH_SupprimerComposant session, asm, solidAsm, moteurID
            End If
        End If
    End If
    wsEtat.Range(wsEtat.Cells(ligneEtat, TECH_COL_MOTEUR_COTE), _
        wsEtat.Cells(ligneEtat, TECH_COL_MOTEUR_SOURCE)).ClearContents
End Sub

Private Function TECH_RollConforme( _
    ByVal modele As pfcls.IpfcModel, _
    ByVal nomVarianteAttendu As String, _
    ByVal diametreAttendu As Double, _
    ByVal largeurAttendue As Double) As Boolean

    Dim solid As pfcls.IpfcSolid
    Dim dimension As pfcls.IpfcBaseDimension
    Dim csys As pfcls.IpfcModelItem

    TECH_RollConforme = False
    If modele Is Nothing Then Exit Function
    If StrComp(TECH_NomModeleLogique(modele.Filename), _
        TECH_NomModeleLogique(nomVarianteAttendu), vbTextCompare) <> 0 Then Exit Function
    Set solid = modele
    If solid Is Nothing Then Exit Function
    Set dimension = TECH_TrouverDimension(solid, TECH_ROLL_DIM)
    If dimension Is Nothing Then Exit Function
    If Not TECH_DoublesEgaux(TECH_ValeurDimensionPositive(dimension), _
        diametreAttendu) Then Exit Function
    Set dimension = TECH_TrouverDimension(solid, TECH_ROLL_WIDTH)
    If dimension Is Nothing Then Exit Function
    If Not TECH_DoublesEgaux(TECH_ValeurDimensionPositive(dimension), _
        largeurAttendue) Then Exit Function
    Set csys = TECH_TrouverCsysExact(modele, TECH_ROLL_CSYS)
    If csys Is Nothing Then Exit Function
    TECH_RollConforme = True
End Function

Private Function TECH_OccurrenceRollPlacee( _
    ByVal solidAsm As pfcls.IpfcSolid, _
    ByVal featureID As Long) As Boolean

    Dim feature As pfcls.IpfcFeature
    Dim composant As pfcls.IpfcComponentFeat

    On Error GoTo NonConforme
    Set feature = solidAsm.GetFeatureById(featureID)
    If feature Is Nothing Then Exit Function
    Set composant = feature
    If composant Is Nothing Then Exit Function
    If Not composant.IsPlaced Then Exit Function
    If composant.IsPackaged Then Exit Function
    If composant.IsUnderconstrained Then Exit Function
    If composant.IsFrozen Then Exit Function
    TECH_OccurrenceRollPlacee = True
    Exit Function

NonConforme:
    Err.Clear
    TECH_OccurrenceRollPlacee = False
End Function

Private Function TECH_OccurrenceRollCiblee( _
    ByVal solidAsm As pfcls.IpfcSolid, _
    ByVal rollFeatureID As Long, _
    ByVal sourceFeatureIDAttendu As Long) As Boolean

    Dim featureRoll As pfcls.IpfcFeature
    Dim composant As pfcls.IpfcComponentFeat
    Dim contraintes As pfcls.IpfcComponentConstraints
    Dim contrainte As pfcls.IpfcComponentConstraint
    Dim selectionSource As pfcls.IpfcSelection
    Dim pathSource As pfcls.IpfcComponentPath
    Dim idsSource As pfcls.Iintseq

    TECH_OccurrenceRollCiblee = False
    If solidAsm Is Nothing Then Exit Function
    If rollFeatureID <= 0 Or sourceFeatureIDAttendu <= 0 Then Exit Function

    On Error GoTo NonConforme
    Set featureRoll = solidAsm.GetFeatureById(rollFeatureID)
    If featureRoll Is Nothing Then Exit Function
    Set composant = featureRoll
    If composant Is Nothing Then Exit Function

    Set contraintes = composant.GetConstraints()
    If contraintes Is Nothing Then Exit Function
    If contraintes.Count <> 1 Then Exit Function
    Set contrainte = contraintes.item(0)
    If contrainte Is Nothing Then Exit Function
    If contrainte.Type <> pfcls.EpfcASM_CONSTRAINT_CSYS Then Exit Function

    Set selectionSource = contrainte.AssemblyReference
    If selectionSource Is Nothing Then Exit Function
    Set pathSource = selectionSource.path
    If pathSource Is Nothing Then Exit Function
    Set idsSource = pathSource.ComponentIds
    If idsSource Is Nothing Then Exit Function
    If idsSource.Count <> 1 Then Exit Function

    TECH_OccurrenceRollCiblee = _
        (idsSource.item(0) = sourceFeatureIDAttendu)
    Exit Function

NonConforme:
    Err.Clear
    TECH_OccurrenceRollCiblee = False
End Function

Private Function TECH_ObtenirVarianteRoll( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal modeleSource As pfcls.IpfcModel, _
    ByVal dossierPieces As String, _
    ByVal diametre As Double, _
    ByVal largeur As Double, _
    ByVal cacheVariantes As Object) As pfcls.IpfcModel

    Dim nomBase As String
    Dim nomFichier As String
    Dim cle As String
    Dim modele As pfcls.IpfcModel
    Dim solid As pfcls.IpfcSolid
    Dim dimension As pfcls.IpfcBaseDimension
    Dim dimensionLargeur As pfcls.IpfcBaseDimension
    Dim dossierInitial As String
    Dim modeleCree As Boolean
    Dim errNumero As Long
    Dim errDescription As String
    Dim csysControle As pfcls.IpfcModelItem

    On Error GoTo Erreur
    If diametre <= 0# Then
        Err.Raise vbObjectError + 15070, , _
            "Diametre de rouleau invalide : " & CStr(diametre)
    End If
    If largeur <= 0# Then
        Err.Raise vbObjectError + 15077, , _
            "Largeur de rouleau invalide : " & CStr(largeur)
    End If

    nomBase = TECH_NomVarianteRoll(diametre, largeur)
    nomFichier = nomBase & ".prt"
    cle = UCase$(nomFichier)

    If cacheVariantes.Exists(cle) Then
        Set modele = cacheVariantes(cle)
        Set TECH_ObtenirVarianteRoll = modele
        Exit Function
    End If

    Set modele = TECH_ChargerModele(session, nomFichier, dossierPieces, True)
    If modele Is Nothing Then
        dossierInitial = session.GetCurrentDirectory()
        TECH_ChangerDossierCreo session, dossierPieces, _
            "Copie de la variante rouleau"
        Set modele = modeleSource.CopyAndRetrieve(nomBase, Nothing)
        modeleCree = True
        If modele Is Nothing Then
            Err.Raise vbObjectError + 15071, , _
                "CopyAndRetrieve a echoue pour " & nomFichier & "."
        End If
    End If

    Set solid = modele
    If solid Is Nothing Then
        Err.Raise vbObjectError + 15072, , nomFichier & " n'est pas une piece."
    End If
    Set dimension = TECH_TrouverDimension(solid, TECH_ROLL_DIM)
    If dimension Is Nothing Then
        Err.Raise vbObjectError + 15073, , _
            "La cote " & TECH_ROLL_DIM & " est absente de " & nomFichier & "."
    End If
    Set dimensionLargeur = TECH_TrouverDimension(solid, TECH_ROLL_WIDTH)
    If dimensionLargeur Is Nothing Then
        Err.Raise vbObjectError + 15078, , _
            "La cote " & TECH_ROLL_WIDTH & " est absente de " & nomFichier & "."
    End If

    If modeleCree Then
        dimension.DimValue = diametre
        dimensionLargeur.DimValue = largeur
        TECH_RegenererSolideFiable solid, "Mise a la cote de " & nomFichier
        TECH_ValiderSolid solid, "Variante " & nomFichier
        If Not modele.CheckIsSaveAllowed(False) Then
            Err.Raise vbObjectError + 15074, , _
                "Creo refuse la sauvegarde de " & nomFichier & "."
        End If
        modele.Save
    ElseIf Not TECH_DoublesEgaux(TECH_ValeurDimensionPositive(dimension), diametre) Or _
        Not TECH_DoublesEgaux(TECH_ValeurDimensionPositive(dimensionLargeur), largeur) Then
        Err.Raise vbObjectError + 15075, , _
            "La variante existante " & nomFichier & " n'a pas les bonnes cotes." & _
            vbCrLf & TECH_ROLL_DIM & " attendu : " & CStr(diametre) & _
            vbCrLf & TECH_ROLL_WIDTH & " attendue : " & CStr(largeur)
    End If

    Set csysControle = TECH_TrouverCsysExact(modele, TECH_ROLL_CSYS)
    If csysControle Is Nothing Then
        Err.Raise vbObjectError + 15076, , _
            "Le repere " & TECH_ROLL_CSYS & " est absent de " & nomFichier & "."
    End If

    If dossierInitial <> "" Then _
        TECH_ChangerDossierCreo session, dossierInitial, _
            "Restauration apres variante rouleau"
    cacheVariantes.Add cle, modele
    Set TECH_ObtenirVarianteRoll = modele
    Exit Function

Erreur:
    errNumero = Err.Number
    errDescription = Err.description
    On Error GoTo -1 ' V14 : libere le gestionnaire actif avant nettoyage
    On Error Resume Next
    If dossierInitial <> "" Then _
        TECH_ChangerDossierCreo session, dossierInitial, _
            "Restauration apres erreur variante rouleau"
    If modeleCree Then
        If Not modele Is Nothing Then modele.Delete
    End If
    On Error GoTo 0
    Err.Raise errNumero, , "TECH_ObtenirVarianteRoll :" & vbCrLf & errDescription
End Function

Private Function TECH_LireDiametreSource( _
    ByVal modeleSource As pfcls.IpfcModel, _
    ByVal solidSource As pfcls.IpfcSolid, _
    ByVal nomAssemblage As String, _
    ByVal sourceFeatureID As Long, _
    ByRef nomPieceSource As String, _
    ByRef occurrence As Long) As Double

    Dim dimension As pfcls.IpfcBaseDimension
    Dim valeurCreo As Double
    Dim valeurExcel As Double
    Dim excelTrouve As Boolean

    Set dimension = TECH_TrouverDimension(solidSource, TECH_ROLL_DIM)
    If dimension Is Nothing Then
        Err.Raise vbObjectError + 15080, , _
            "La piece " & modeleSource.Filename & " contient un repere ROLL..._CENTER" & _
            " mais ne contient pas la cote " & TECH_ROLL_DIM & "."
    End If
    valeurCreo = TECH_ValeurDimensionPositive(dimension)
    If valeurCreo <= 0# Then
        Err.Raise vbObjectError + 15081, , _
            "La cote " & TECH_ROLL_DIM & " est invalide dans " & modeleSource.Filename & "."
    End If

    valeurExcel = TECH_LireDiametreExcel(nomAssemblage, sourceFeatureID, _
        modeleSource.Filename, valeurCreo, nomPieceSource, occurrence, excelTrouve)

    If excelTrouve Then
        If Not TECH_DoublesEgaux(valeurExcel, valeurCreo) Then
            Err.Raise vbObjectError + 15082, , _
                "Excel et Creo ne donnent pas le meme " & TECH_ROLL_DIM & "." & vbCrLf & _
                "Piece : " & nomPieceSource & vbCrLf & _
                "Occurrence : " & CStr(occurrence) & vbCrLf & _
                "Excel : " & CStr(valeurExcel) & vbCrLf & _
                "Creo : " & CStr(valeurCreo) & vbCrLf & _
                "Rafraichis d'abord la piece ou l'assemblage."
        End If
        TECH_LireDiametreSource = valeurExcel
    Else
        nomPieceSource = TECH_NomModeleLogique(modeleSource.Filename)
        occurrence = 0
        Debug.Print "AJOUT PIECES TECHNIQUES : mapping Excel non trouve pour " & _
            modeleSource.Filename & ". Valeur Creo utilisee : " & CStr(valeurCreo)
        TECH_LireDiametreSource = valeurCreo
    End If
End Function

Private Function TECH_LireDiametreExcel( _
    ByVal nomAssemblage As String, _
    ByVal sourceFeatureID As Long, _
    ByVal modelePhysique As String, _
    ByVal valeurCreo As Double, _
    ByRef nomPieceSource As String, _
    ByRef occurrence As Long, _
    ByRef trouve As Boolean) As Double

    Dim wsEtatSource As Worksheet
    Dim wsParam As Worksheet
    Dim ligneEtat As Long
    Dim ligneDimension As Long
    Dim colonneOccurrence As Long
    Dim valeur As Variant

    trouve = False
    nomPieceSource = ""
    occurrence = 0

    Set wsEtatSource = TECH_TrouverFeuilleEtatAssemblage(nomAssemblage)
    If Not wsEtatSource Is Nothing Then
        ligneEtat = TECH_TrouverMappingDansEtat(wsEtatSource, sourceFeatureID, modelePhysique)
        If ligneEtat > 0 Then
            nomPieceSource = Trim$(CStr(wsEtatSource.Cells(ligneEtat, 2).value))
            occurrence = CLng(Val(CStr(wsEtatSource.Cells(ligneEtat, 3).value)))
            Set wsParam = TECH_TrouverFeuilleParametres(nomPieceSource)
            If wsParam Is Nothing Then
                Err.Raise vbObjectError + 15090, , _
                    "Feuille Excel introuvable pour " & nomPieceSource & "."
            End If
            ligneDimension = TECH_TrouverLigneNom(wsParam, TECH_ROLL_DIM)
            If ligneDimension = 0 Then
                Err.Raise vbObjectError + 15091, , _
                    "La ligne " & TECH_ROLL_DIM & " est absente de la feuille " & _
                    wsParam.name & "."
            End If
            colonneOccurrence = TECH_TrouverColonneOccurrence(wsParam, occurrence)
            valeur = wsParam.Cells(ligneDimension, colonneOccurrence).value
            If IsError(valeur) Or Not IsNumeric(valeur) Then
                Err.Raise vbObjectError + 15092, , _
                    "Valeur Excel non numerique pour " & TECH_ROLL_DIM & _
                    " dans " & wsParam.name & "."
            End If
            TECH_LireDiametreExcel = Abs(CDbl(valeur))
            trouve = True
            Exit Function
        End If
    End If

    TECH_RechercherValeurExcelParModele modelePhysique, valeurCreo, wsParam, _
        colonneOccurrence, ligneDimension
    If Not wsParam Is Nothing Then
        valeur = wsParam.Cells(ligneDimension, colonneOccurrence).value
        nomPieceSource = TECH_NomModeleLogique(CStr(wsParam.Range("B2").value))
        occurrence = TECH_NumeroOccurrenceColonne(wsParam, colonneOccurrence)
        TECH_LireDiametreExcel = Abs(CDbl(valeur))
        trouve = True
    End If
End Function

Private Function TECH_LireLargeurSource( _
    ByVal modeleSource As pfcls.IpfcModel, _
    ByVal solidSource As pfcls.IpfcSolid, _
    ByVal nomPieceSource As String, _
    ByVal occurrence As Long) As Double

    Dim dimension As pfcls.IpfcBaseDimension
    Dim valeurCreo As Double
    Dim valeurExcel As Double
    Dim wsParam As Worksheet
    Dim ligneLargeur As Long
    Dim colonneOccurrence As Long
    Dim valeur As Variant

    Set dimension = TECH_TrouverDimension(solidSource, TECH_ROLL_WIDTH)
    If dimension Is Nothing Then
        Err.Raise vbObjectError + 15083, , _
            "La piece " & modeleSource.Filename & " contient un repere ROLL..._CENTER" & _
            " mais ne contient pas la cote " & TECH_ROLL_WIDTH & "."
    End If
    valeurCreo = TECH_ValeurDimensionPositive(dimension)
    If valeurCreo <= 0# Then
        Err.Raise vbObjectError + 15084, , _
            "La cote " & TECH_ROLL_WIDTH & " est invalide dans " & _
            modeleSource.Filename & "."
    End If

    If occurrence <= 0 Or Len(Trim$(nomPieceSource)) = 0 Then
        Debug.Print "AJOUT PIECES TECHNIQUES : mapping Excel non trouve pour " & _
            TECH_ROLL_WIDTH & " de " & modeleSource.Filename & _
            ". Valeur Creo utilisee : " & CStr(valeurCreo)
        TECH_LireLargeurSource = valeurCreo
        Exit Function
    End If

    Set wsParam = TECH_TrouverFeuilleParametres(nomPieceSource)
    If wsParam Is Nothing Then
        Err.Raise vbObjectError + 15085, , _
            "Feuille Excel introuvable pour " & nomPieceSource & "."
    End If
    ligneLargeur = TECH_TrouverLigneNom(wsParam, TECH_ROLL_WIDTH)
    If ligneLargeur = 0 Then
        Err.Raise vbObjectError + 15086, , _
            "La ligne " & TECH_ROLL_WIDTH & " est absente de la feuille " & _
            wsParam.name & "."
    End If
    colonneOccurrence = TECH_TrouverColonneOccurrence(wsParam, occurrence)
    valeur = wsParam.Cells(ligneLargeur, colonneOccurrence).value
    If IsError(valeur) Or Not IsNumeric(valeur) Then
        Err.Raise vbObjectError + 15087, , _
            "Valeur Excel non numerique pour " & TECH_ROLL_WIDTH & _
            " dans " & wsParam.name & "."
    End If
    valeurExcel = Abs(CDbl(valeur))
    If valeurExcel <= 0# Then
        Err.Raise vbObjectError + 15088, , _
            "Valeur Excel invalide pour " & TECH_ROLL_WIDTH & _
            " dans " & wsParam.name & "."
    End If
    If Not TECH_DoublesEgaux(valeurExcel, valeurCreo) Then
        Err.Raise vbObjectError + 15089, , _
            "Excel et Creo ne donnent pas la meme " & TECH_ROLL_WIDTH & "." & vbCrLf & _
            "Piece : " & nomPieceSource & vbCrLf & _
            "Occurrence : " & CStr(occurrence) & vbCrLf & _
            "Excel : " & CStr(valeurExcel) & vbCrLf & _
            "Creo : " & CStr(valeurCreo) & vbCrLf & _
            "Rafraichis d'abord la piece ou l'assemblage."
    End If
    TECH_LireLargeurSource = valeurExcel
End Function

Private Sub TECH_RechercherValeurExcelParModele( _
    ByVal modelePhysique As String, _
    ByVal valeurCreo As Double, _
    ByRef wsTrouvee As Worksheet, _
    ByRef colonneTrouvee As Long, _
    ByRef ligneTrouvee As Long)

    Dim ws As Worksheet
    Dim nomPiece As String
    Dim ligne As Long
    Dim derniereColonne As Long
    Dim col As Long
    Dim valeur As Variant
    Dim nbCorrespondances As Long

    Set wsTrouvee = Nothing
    colonneTrouvee = 0
    ligneTrouvee = 0

    For Each ws In ThisWorkbook.Worksheets
        nomPiece = Trim$(CStr(ws.Range("B2").value))
        If InStr(1, nomPiece, ".prt", vbTextCompare) > 0 Then
            If TECH_ModeleCorrespondPiece(modelePhysique, nomPiece) Then
                ligne = TECH_TrouverLigneNom(ws, TECH_ROLL_DIM)
                If ligne > 0 Then
                    derniereColonne = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
                    For col = 2 To derniereColonne
                        valeur = ws.Cells(ligne, col).value
                        If Not IsError(valeur) Then
                            If IsNumeric(valeur) Then
                                If TECH_DoublesEgaux(Abs(CDbl(valeur)), valeurCreo) Then
                                    nbCorrespondances = nbCorrespondances + 1
                                    Set wsTrouvee = ws
                                    colonneTrouvee = col
                                    ligneTrouvee = ligne
                                End If
                            End If
                        End If
                    Next col
                End If
            End If
        End If
    Next ws

    If nbCorrespondances <> 1 Then
        Set wsTrouvee = Nothing
        colonneTrouvee = 0
        ligneTrouvee = 0
    End If
End Sub

Private Function TECH_TrouverFeuilleEtatAssemblage( _
    ByVal nomAssemblage As String) As Worksheet

    Dim candidats As Variant
    Dim candidat As Variant
    Dim base As String
    Dim ws As Worksheet

    base = TECH_BaseSansExtension(nomAssemblage)
    candidats = Array(TECH_NomFeuilleEtatSection(base), _
        TECH_NomFeuilleEtatSection(nomAssemblage), "_ETAT_RAFRAICHISSEMENT")

    For Each candidat In candidats
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(candidat))
        Err.Clear
        On Error GoTo 0
        If Not ws Is Nothing Then
            Set TECH_TrouverFeuilleEtatAssemblage = ws
            Exit Function
        End If
    Next candidat
End Function

Private Function TECH_TrouverMappingDansEtat( _
    ByVal ws As Worksheet, _
    ByVal featureID As Long, _
    ByVal modelePhysique As String) As Long

    Dim derniereLigne As Long
    Dim r As Long
    Dim modeleEtat As String

    derniereLigne = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    For r = 2 To derniereLigne
        If CLng(Val(CStr(ws.Cells(r, 6).value))) = featureID Then
            modeleEtat = Trim$(CStr(ws.Cells(r, 7).value))
            If StrComp(TECH_NomModeleLogique(modeleEtat), _
                TECH_NomModeleLogique(modelePhysique), vbTextCompare) = 0 Then
                TECH_TrouverMappingDansEtat = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Function TECH_TrouverFeuilleParametres( _
    ByVal nomPiece As String) As Worksheet

    Dim ws As Worksheet
    Dim cible As String
    Dim courant As String

    cible = TECH_NomModeleLogique(nomPiece)
    For Each ws In ThisWorkbook.Worksheets
        courant = Trim$(CStr(ws.Range("B2").value))
        If InStr(1, courant, ".prt", vbTextCompare) > 0 Then
            If StrComp(TECH_NomModeleLogique(courant), cible, vbTextCompare) = 0 Then
                Set TECH_TrouverFeuilleParametres = ws
                Exit Function
            End If
        End If
    Next ws
End Function

Private Function TECH_TrouverLigneNom( _
    ByVal ws As Worksheet, _
    ByVal nomRecherche As String) As Long

    Dim derniereLigne As Long
    Dim r As Long
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 4 To derniereLigne
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), _
            nomRecherche, vbTextCompare) = 0 Then
            TECH_TrouverLigneNom = r
            Exit Function
        End If
    Next r
End Function

Private Function TECH_TrouverColonneOccurrence( _
    ByVal ws As Worksheet, _
    ByVal occurrence As Long) As Long

    Dim derniereColonne As Long
    Dim col As Long
    Dim colonneValue As Long
    Dim valeur As Variant

    derniereColonne = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    For col = 2 To derniereColonne
        valeur = ws.Cells(3, col).value
        If StrComp(Trim$(CStr(valeur)), "VALUE", vbTextCompare) = 0 Then
            colonneValue = col
        ElseIf IsNumeric(valeur) Then
            If CLng(valeur) = occurrence Then
                TECH_TrouverColonneOccurrence = col
                Exit Function
            End If
        End If
    Next col

    If colonneValue > 0 Then
        TECH_TrouverColonneOccurrence = colonneValue
    Else
        Err.Raise vbObjectError + 15100, , _
            "Colonne d'occurrence " & CStr(occurrence) & _
            " introuvable dans la feuille " & ws.name & "."
    End If
End Function

Private Function TECH_NumeroOccurrenceColonne( _
    ByVal ws As Worksheet, _
    ByVal colonne As Long) As Long

    Dim valeur As Variant
    valeur = ws.Cells(3, colonne).value
    If IsNumeric(valeur) Then
        TECH_NumeroOccurrenceColonne = CLng(valeur)
    Else
        TECH_NumeroOccurrenceColonne = 1
    End If
End Function

Private Function TECH_ListerReperesRoll( _
    ByVal modele As pfcls.IpfcModel) As Collection

    Dim resultat As New Collection
    Dim owner As pfcls.IpfcModelItemOwner
    Dim items As pfcls.IpfcModelItems
    Dim item As pfcls.IpfcModelItem
    Dim nom As String
    Dim i As Long

    Set owner = modele
    If owner Is Nothing Then
        Set TECH_ListerReperesRoll = resultat
        Exit Function
    End If
    Set items = owner.ListItems(pfcls.EpfcITEM_COORD_SYS)
    If items Is Nothing Then
        Set TECH_ListerReperesRoll = resultat
        Exit Function
    End If

    For i = 0 To items.Count - 1
        Set item = items.item(i)
        If Not item Is Nothing Then
            nom = ""
            On Error Resume Next
            nom = UCase$(Trim$(item.GetName()))
            Err.Clear
            On Error GoTo 0
            If TECH_EstNomRepereRoll(nom) Then resultat.Add item
        End If
    Next i
    Set TECH_ListerReperesRoll = resultat
End Function

Private Function TECH_EstNomRepereRoll(ByVal nom As String) As Boolean
    Dim milieu As String
    Dim i As Long
    Dim ch As String

    nom = UCase$(Trim$(nom))
    If nom = "ROLL_CENTER" Then
        TECH_EstNomRepereRoll = True
        Exit Function
    End If
    If Left$(nom, 4) <> "ROLL" Then Exit Function
    If Right$(nom, 7) <> "_CENTER" Then Exit Function
    If Len(nom) <= 11 Then Exit Function

    milieu = Mid$(nom, 5, Len(nom) - 11)
    If milieu = "" Then Exit Function
    For i = 1 To Len(milieu)
        ch = Mid$(milieu, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    TECH_EstNomRepereRoll = True
End Function

Private Function TECH_TrouverCsysExact( _
    ByVal modele As pfcls.IpfcModel, _
    ByVal nomRecherche As String) As pfcls.IpfcModelItem

    Dim owner As pfcls.IpfcModelItemOwner
    Dim item As pfcls.IpfcModelItem
    Dim items As pfcls.IpfcModelItems
    Dim i As Long
    Dim nom As String

    Set owner = modele
    If owner Is Nothing Then Exit Function
    Set items = owner.ListItems(pfcls.EpfcITEM_COORD_SYS)
    If items Is Nothing Then Exit Function

    For i = 0 To items.Count - 1
        Set item = items.item(i)
        If Not item Is Nothing Then
            nom = ""
            On Error Resume Next
            nom = Trim$(item.GetName())
            Err.Clear
            On Error GoTo 0
            If StrComp(nom, nomRecherche, vbTextCompare) = 0 Then
                Set TECH_TrouverCsysExact = item
                Exit Function
            End If
        End If
    Next i
End Function

Private Function TECH_TrouverDimension( _
    ByVal proprietaire As pfcls.IpfcModelItemOwner, _
    ByVal nomRecherche As String) As pfcls.IpfcBaseDimension

    Dim item As pfcls.IpfcModelItem
    Dim items As pfcls.IpfcModelItems
    Dim dimension As pfcls.IpfcBaseDimension
    Dim nom As String
    Dim i As Long

    On Error Resume Next
    Set item = proprietaire.GetItemByName(pfcls.EpfcITEM_DIMENSION, nomRecherche)
    If item Is Nothing Then
        Err.Clear
        Set item = proprietaire.GetItemByName(pfcls.EpfcITEM_DIMENSION, _
            UCase$(nomRecherche))
    End If
    If Not item Is Nothing Then Set dimension = item
    Err.Clear
    On Error GoTo 0
    If Not dimension Is Nothing Then
        Set TECH_TrouverDimension = dimension
        Exit Function
    End If

    Set items = proprietaire.ListItems(pfcls.EpfcITEM_DIMENSION)
    If items Is Nothing Then Exit Function
    For i = 0 To items.Count - 1
        Set item = items.item(i)
        Set dimension = Nothing
        On Error Resume Next
        Set dimension = item
        nom = ""
        If Not dimension Is Nothing Then nom = dimension.Symbol
        If nom = "" Then nom = item.GetName()
        Err.Clear
        On Error GoTo 0
        If StrComp(Trim$(nom), nomRecherche, vbTextCompare) = 0 Then
            Set TECH_TrouverDimension = dimension
            Exit Function
        End If
    Next i
End Function

Private Function TECH_ValeurDimensionPositive( _
    ByVal dimension As pfcls.IpfcBaseDimension) As Double

    TECH_ValeurDimensionPositive = Abs(CDbl(dimension.DimValue))
End Function

Private Function TECH_CreerPath( _
    ByVal asm As pfcls.IpfcAssembly, _
    ByVal featureID As Long) As pfcls.IpfcComponentPath

    Dim ids As pfcls.Iintseq
    Dim factory As pfcls.CMpfcAssembly
    Set ids = New pfcls.Cintseq
    ids.Append featureID
    Set factory = New pfcls.CMpfcAssembly
    Set TECH_CreerPath = factory.CreateComponentPath(asm, ids)
    If TECH_CreerPath Is Nothing Then
        Err.Raise vbObjectError + 15110, , _
            "ComponentPath impossible pour la feature " & CStr(featureID) & "."
    End If
End Function

Private Function TECH_CreerPathRacine( _
    ByVal asm As pfcls.IpfcAssembly) As pfcls.IpfcComponentPath

    Dim ids As pfcls.Iintseq
    Dim factory As pfcls.CMpfcAssembly
    Set ids = New pfcls.Cintseq
    Set factory = New pfcls.CMpfcAssembly
    Set TECH_CreerPathRacine = factory.CreateComponentPath(asm, ids)
End Function

Private Function TECH_CreerSelectionCsys( _
    ByVal csys As pfcls.IpfcModelItem, _
    ByVal path As pfcls.IpfcComponentPath) As pfcls.IpfcSelection

    Dim factory As pfcls.CMpfcSelect
    Set factory = New pfcls.CMpfcSelect
    Set TECH_CreerSelectionCsys = factory.CreateModelItemSelection(csys, path)
    If TECH_CreerSelectionCsys Is Nothing Then
        Err.Raise vbObjectError + 15111, , "Selection CSYS impossible."
    End If
End Function

Private Function TECH_CreerContrainteCsys( _
    ByVal selectionAssemblage As pfcls.IpfcSelection, _
    ByVal selectionComposant As pfcls.IpfcSelection) As pfcls.IpfcComponentConstraints

    Dim factory As pfcls.CCpfcComponentConstraint
    Dim contrainte As pfcls.IpfcComponentConstraint
    Dim contraintes As pfcls.IpfcComponentConstraints

    Set factory = New pfcls.CCpfcComponentConstraint
    Set contrainte = factory.Create(pfcls.EpfcASM_CONSTRAINT_CSYS)
    contrainte.AssemblyReference = selectionAssemblage
    contrainte.ComponentReference = selectionComposant
    Set contraintes = New pfcls.CpfcComponentConstraints
    contraintes.Append contrainte
    Set TECH_CreerContrainteCsys = contraintes
End Function

Private Function TECH_ModeleDepuisComposant( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal composant As pfcls.IpfcComponentFeat) As pfcls.IpfcModel

    On Error Resume Next
    Set TECH_ModeleDepuisComposant = session.GetModelFromDescr(composant.ModelDescr)
    Err.Clear
    On Error GoTo 0
End Function

Private Function TECH_ModeleComposantParID( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal solidAsm As pfcls.IpfcSolid, _
    ByVal featureID As Long) As pfcls.IpfcModel

    Dim feature As pfcls.IpfcFeature
    Dim composant As pfcls.IpfcComponentFeat
    If featureID <= 0 Then Exit Function
    On Error Resume Next
    Set feature = solidAsm.GetFeatureById(featureID)
    If Not feature Is Nothing Then Set composant = feature
    If Not composant Is Nothing Then
        Set TECH_ModeleComposantParID = session.GetModelFromDescr(composant.ModelDescr)
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function TECH_ChargerModele( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal nomFichier As String, _
    ByVal dossier As String, _
    Optional ByVal autoriserAbsent As Boolean = False) As pfcls.IpfcModel

    Dim modele As pfcls.IpfcModel
    Dim cDescr As pfcls.CCpfcModelDescriptor
    Dim descr As pfcls.IpfcModelDescriptor
    Dim cOptions As pfcls.CCpfcRetrieveModelOptions
    Dim options As pfcls.IpfcRetrieveModelOptions
    Dim nomLogique As String
    Dim dossierNormalise As String
    Dim fichierTrouve As String
    Dim tentative As Long
    Dim numeroErreur As Long
    Dim descriptionErreur As String
    Dim sourceErreur As String
    Dim dossierInitial As String
    Dim dossierChange As Boolean
    Dim etapeLocale As String

    On Error GoTo Erreur
    etapeLocale = "Normalisation du nom Creo"
    nomLogique = TECH_NomModeleLogique(nomFichier)
    etapeLocale = "Recherche du modele dans la session"
    On Error Resume Next
    Set modele = session.GetModelFromFileName(nomLogique)
    Err.Clear
    On Error GoTo Erreur
    If Not modele Is Nothing Then
        Set TECH_ChargerModele = modele
        Exit Function
    End If

    dossierNormalise = Replace(Trim$(dossier), "/", "\")
    If dossierNormalise <> "" Then
        If Right$(dossierNormalise, 1) <> "\" Then _
            dossierNormalise = dossierNormalise & "\"
    End If

    etapeLocale = "Recherche du fichier physique"
    fichierTrouve = Dir$(dossierNormalise & nomLogique & "*")
    If fichierTrouve = "" Then
        If autoriserAbsent Then
            Set TECH_ChargerModele = Nothing
            Exit Function
        End If
        Err.Raise vbObjectError + 15120, , _
            "Aucun fichier physique correspondant a " & nomLogique & "*" & vbCrLf & _
            "Dossier : " & dossierNormalise
    End If

    ' RetrieveModelWithOpts ne tient pas toujours compte du chemin porte par
    ' le descripteur. C'est notamment le cas lorsque l'assemblage moteur a ete
    ' choisi dans un dossier different du repertoire courant de Creo.
    ' Basculer temporairement dans le dossier du fichier rend le chargement
    ' identique pour Moteur_L et Moteur_R, puis le contexte initial est restaure.
    etapeLocale = "Bascule vers le dossier du modele"
    dossierInitial = session.GetCurrentDirectory()
    If dossierNormalise <> "" Then
        If StrComp(TECH_DossierComparable(dossierInitial), _
            TECH_DossierComparable(dossierNormalise), vbTextCompare) <> 0 Then
            TECH_ChangerDossierCreo session, dossierNormalise, _
                "Chargement du modele " & nomLogique
            dossierChange = True
        End If
    End If

    etapeLocale = "Creation des options de chargement"
    Set cOptions = New pfcls.CCpfcRetrieveModelOptions
    Set options = cOptions.Create
    For tentative = 1 To 5
        etapeLocale = "Creation du descripteur Creo"
        Set cDescr = New pfcls.CCpfcModelDescriptor
        Set descr = cDescr.CreateFromFileName(nomLogique)
        Set modele = Nothing
        numeroErreur = 0
        descriptionErreur = ""
        sourceErreur = ""

        etapeLocale = "Chargement du modele dans Creo"
        On Error Resume Next
        Err.Clear
        Set modele = session.RetrieveModelWithOpts(descr, options)
        numeroErreur = Err.Number
        descriptionErreur = Err.description
        sourceErreur = Err.source
        Err.Clear
        If modele Is Nothing Then Set modele = session.GetModelFromDescr(descr)
        Err.Clear
        On Error GoTo Erreur

        If Not modele Is Nothing Then Exit For
        Debug.Print "CHARGEMENT CREO ECHEC essai " & CStr(tentative) & _
            "/5 : " & nomLogique & " / " & descriptionErreur
        DoEvents
    Next tentative

    If modele Is Nothing Then
        Err.Raise vbObjectError + 15121, sourceErreur, _
            "Le fichier existe sur le disque mais Creo ne peut pas le charger." & vbCrLf & _
            "Fichier detecte : " & fichierTrouve & vbCrLf & _
            "Dossier : " & dossierNormalise & _
            IIf(descriptionErreur <> "", vbCrLf & "Erreur Creo : " & _
                descriptionErreur, "")
    End If
    If dossierChange Then _
        TECH_ChangerDossierCreo session, dossierInitial, _
            "Restauration apres chargement de " & nomLogique
    Set TECH_ChargerModele = modele
    Exit Function

Erreur:
    numeroErreur = Err.Number
    descriptionErreur = Err.description
    sourceErreur = Err.source
    On Error GoTo -1 ' V14 : libere le gestionnaire actif avant nettoyage
    On Error Resume Next
    If dossierChange Then _
        TECH_ChangerDossierCreo session, dossierInitial, _
            "Restauration apres erreur de chargement"
    On Error GoTo 0
    If autoriserAbsent And fichierTrouve = "" Then
        Set TECH_ChargerModele = Nothing
    Else
        Err.Raise numeroErreur, sourceErreur, _
            "TECH_ChargerModele :" & vbCrLf & _
            "Etape interne : " & etapeLocale & vbCrLf & _
            "Modele : " & nomLogique & vbCrLf & _
            "Dossier : " & dossierNormalise & vbCrLf & descriptionErreur
    End If
End Function

Private Function TECH_DossierComparable(ByVal dossier As String) As String
    dossier = Replace(Trim$(dossier), "/", "\")
    Do While Len(dossier) > 3 And Right$(dossier, 1) = "\"
        dossier = Left$(dossier, Len(dossier) - 1)
    Loop
    TECH_DossierComparable = dossier
End Function

Private Sub TECH_ChangerDossierCreo( _
    ByVal session As pfcls.IpfcBaseSession, _
    ByVal dossier As String, _
    ByVal contexte As String)

    Dim dossierCreo As String
    Dim dossierCourt As String
    Dim fso As Object
    Dim numeroErreur As Long
    Dim sourceErreur As String
    Dim descriptionErreur As String

    dossierCreo = TECH_ResoudreDossierCreo(dossier)
    If dossierCreo = "" Then
        Err.Raise vbObjectError + 15229, , _
            contexte & " : dossier Creo vide."
    End If

    On Error Resume Next
    Err.Clear
    session.ChangeDirectory dossierCreo
    numeroErreur = Err.Number
    sourceErreur = Err.source
    descriptionErreur = Err.description
    Err.Clear
    On Error GoTo 0
    If numeroErreur = 0 Then Exit Sub

    ' Dernier repli pour les installations Creo qui refusent les espaces :
    ' utiliser le chemin court Windows lorsqu'il est disponible.
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If fso.FolderExists(dossierCreo) Then _
        dossierCourt = CStr(fso.GetFolder(dossierCreo).ShortPath)
    Err.Clear
    On Error GoTo 0
    If dossierCourt <> "" Then
        If StrComp(dossierCourt, dossierCreo, vbTextCompare) <> 0 Then
            On Error Resume Next
            Err.Clear
            session.ChangeDirectory dossierCourt
            numeroErreur = Err.Number
            sourceErreur = Err.source
            descriptionErreur = Err.description
            Err.Clear
            On Error GoTo 0
            If numeroErreur = 0 Then Exit Sub
        End If
    End If

    Err.Raise numeroErreur, sourceErreur, _
        contexte & " : changement de dossier Creo impossible." & vbCrLf & _
        "Dossier transmis : " & dossier & vbCrLf & _
        "Dossier resolu pour Creo : " & dossierCreo & vbCrLf & _
        IIf(dossierCourt <> "", "Chemin court essaye : " & _
            dossierCourt & vbCrLf, "") & _
        descriptionErreur
End Sub

Private Function TECH_ResoudreDossierCreo(ByVal dossier As String) As String
    Dim chemin As String
    Dim reseau As Object
    Dim lecteurs As Object
    Dim fso As Object
    Dim lecteur As String
    Dim partage As String
    Dim meilleurLecteur As String
    Dim meilleurPartage As String
    Dim dossierResolu As String
    Dim i As Long
    Dim numeroErreur As Long
    Dim sourceErreur As String
    Dim descriptionErreur As String

    chemin = TECH_DossierComparable(dossier)
    If Left$(chemin, 2) <> "\\" Then
        TECH_ResoudreDossierCreo = chemin
        Exit Function
    End If

    Set reseau = CreateObject("WScript.Network")
    Set lecteurs = reseau.EnumNetworkDrives
    For i = 0 To lecteurs.Count - 1 Step 2
        lecteur = CStr(lecteurs.item(i))
        partage = TECH_DossierComparable(CStr(lecteurs.item(i + 1)))
        If TECH_CheminDansPartage(chemin, partage) Then
            If Len(partage) > Len(meilleurPartage) Then
                meilleurPartage = partage
                meilleurLecteur = lecteur
            End If
        End If
    Next i

    If meilleurLecteur = "" Then
        partage = TECH_RacinePartageUNC(chemin)
        lecteur = TECH_TrouverLecteurLibre()
        If partage = "" Or lecteur = "" Then
            Err.Raise vbObjectError + 15230, , _
                "Impossible de preparer le chemin reseau pour Creo : " & chemin
        End If

        On Error Resume Next
        Err.Clear
        reseau.MapNetworkDrive lecteur, partage, False
        numeroErreur = Err.Number
        sourceErreur = Err.source
        descriptionErreur = Err.description
        Err.Clear
        On Error GoTo 0
        If numeroErreur <> 0 Then
            Err.Raise numeroErreur, sourceErreur, _
                "Creo ne peut pas utiliser directement le chemin UNC et " & _
                "la creation automatique d'un lecteur reseau a echoue." & _
                vbCrLf & "Partage : " & partage & vbCrLf & _
                descriptionErreur
        End If
        m_TECH_LecteurReseauTemp = lecteur
        m_TECH_PartageReseauTemp = partage
        meilleurLecteur = lecteur
        meilleurPartage = partage
    End If

    dossierResolu = TECH_ComposerCheminMappe( _
        chemin, meilleurPartage, meilleurLecteur)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dossierResolu) Then
        Err.Raise vbObjectError + 15231, , _
            "Le dossier reseau converti est introuvable." & vbCrLf & _
            "UNC : " & chemin & vbCrLf & _
            "Chemin mappe : " & dossierResolu
    End If
    TECH_ResoudreDossierCreo = dossierResolu
End Function

Private Function TECH_CheminDansPartage( _
    ByVal chemin As String, _
    ByVal partage As String) As Boolean

    If partage = "" Or Len(chemin) < Len(partage) Then Exit Function
    If StrComp(Left$(chemin, Len(partage)), partage, _
        vbTextCompare) <> 0 Then Exit Function
    If Len(chemin) = Len(partage) Then
        TECH_CheminDansPartage = True
    Else
        TECH_CheminDansPartage = _
            (Mid$(chemin, Len(partage) + 1, 1) = "\")
    End If
End Function

Private Function TECH_RacinePartageUNC(ByVal chemin As String) As String
    Dim positionServeur As Long
    Dim positionPartage As Long

    If Left$(chemin, 2) <> "\\" Then Exit Function
    positionServeur = InStr(3, chemin, "\", vbBinaryCompare)
    If positionServeur = 0 Then Exit Function
    positionPartage = InStr(positionServeur + 1, chemin, "\", vbBinaryCompare)
    If positionPartage = 0 Then
        TECH_RacinePartageUNC = chemin
    Else
        TECH_RacinePartageUNC = Left$(chemin, positionPartage - 1)
    End If
End Function

Private Function TECH_TrouverLecteurLibre() As String
    Dim fso As Object
    Dim codeLettre As Long
    Dim lecteur As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    For codeLettre = 90 To 68 Step -1
        lecteur = Chr$(codeLettre) & ":"
        If Not fso.DriveExists(lecteur) Then
            TECH_TrouverLecteurLibre = lecteur
            Exit Function
        End If
    Next codeLettre
End Function

Private Function TECH_ComposerCheminMappe( _
    ByVal cheminUNC As String, _
    ByVal partageUNC As String, _
    ByVal lecteur As String) As String

    Dim suffixe As String
    suffixe = Mid$(cheminUNC, Len(partageUNC) + 1)
    If suffixe = "" Then suffixe = "\"
    If Left$(suffixe, 1) <> "\" Then suffixe = "\" & suffixe
    TECH_ComposerCheminMappe = lecteur & suffixe
End Function

Private Sub TECH_LibererLecteurReseauTemp( _
    ByVal session As pfcls.IpfcBaseSession)

    Dim dossierCourant As String
    Dim reseau As Object

    If m_TECH_LecteurReseauTemp = "" Then Exit Sub
    On Error Resume Next
    dossierCourant = session.GetCurrentDirectory()
    If StrComp(Left$(dossierCourant, 2), _
        m_TECH_LecteurReseauTemp, vbTextCompare) = 0 Then
        If m_TECH_DossierCreoInitial <> "" And _
            Left$(m_TECH_DossierCreoInitial, 2) <> "\\" Then
            session.ChangeDirectory _
                TECH_DossierComparable(m_TECH_DossierCreoInitial)
            dossierCourant = session.GetCurrentDirectory()
        End If
    End If

    If StrComp(Left$(dossierCourant, 2), _
        m_TECH_LecteurReseauTemp, vbTextCompare) <> 0 Then
        Set reseau = CreateObject("WScript.Network")
        reseau.RemoveNetworkDrive m_TECH_LecteurReseauTemp, True, False
        m_TECH_LecteurReseauTemp = ""
        m_TECH_PartageReseauTemp = ""
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function TECH_ObtenirDossierPieces( _
    ByVal session As pfcls.IpfcBaseSession) As String

    Dim dossier As String
    Dim ws As Worksheet
    Dim fso As Object

    On Error Resume Next
    dossier = Trim$(CStr(ThisWorkbook.Worksheets("ASSEMBLAGE").Range("G2").value))
    Err.Clear
    On Error GoTo 0

    If dossier = "" Then
        For Each ws In ThisWorkbook.Worksheets
            If StrComp(TECH_NomModeleLogique(CStr(ws.Range("B2").value)), _
                TECH_ROLL_FICHIER, vbTextCompare) = 0 Then
                dossier = Trim$(CStr(ws.Range("B1").value))
                Exit For
            End If
        Next ws
    End If
    If dossier = "" Then dossier = session.GetCurrentDirectory()

    dossier = Replace(dossier, "/", "\")
    If Right$(dossier, 1) <> "\" Then dossier = dossier & "\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dossier) Then
        Err.Raise vbObjectError + 15121, , _
            "Dossier de travail introuvable :" & vbCrLf & dossier
    End If
    TECH_ObtenirDossierPieces = dossier
End Function

Private Function TECH_ObtenirFeuilleEtat( _
    ByVal creer As Boolean) As Worksheet

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TECH_FEUILLE_ETAT)
    Err.Clear
    On Error GoTo 0

    If ws Is Nothing And creer Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = TECH_FEUILLE_ETAT
        ws.Cells(1, TECH_COL_ASSEMBLAGE).value = "ASSEMBLAGE"
        ws.Cells(1, TECH_COL_CLE).value = "CLE_CIBLE"
        ws.Cells(1, TECH_COL_SOURCE_ID).value = "SOURCE_FEATURE_ID"
        ws.Cells(1, TECH_COL_SOURCE_MODELE).value = "SOURCE_MODELE_CREO"
        ws.Cells(1, TECH_COL_SOURCE_PIECE).value = "SOURCE_PIECE_EXCEL"
        ws.Cells(1, TECH_COL_OCCURRENCE).value = "OCCURRENCE_EXCEL"
        ws.Cells(1, TECH_COL_REPERE).value = "REPERE"
        ws.Cells(1, TECH_COL_DIAMETRE).value = "ROLL_DIAM"
        ws.Cells(1, TECH_COL_ROLL_MODELE).value = "ROLL_MODELE_CREO"
        ws.Cells(1, TECH_COL_ROLL_ID).value = "ROLL_FEATURE_ID"
        ws.Cells(1, TECH_COL_STATUT).value = "STATUT"
        ws.Cells(1, TECH_COL_DATE).value = "DERNIERE_MAJ"
    End If
    If Not ws Is Nothing Then
        ws.Cells(1, TECH_COL_LARGEUR).value = "ROLL_LARGEUR"
        ws.Cells(1, TECH_COL_MOTEUR_COTE).value = "MOTEUR_COTE"
        ws.Cells(1, TECH_COL_MOTEUR_MODELE).value = "MOTEUR_MODELE_CREO"
        ws.Cells(1, TECH_COL_MOTEUR_ID).value = "MOTEUR_FEATURE_ID"
        ws.Cells(1, TECH_COL_ROLL_HAUTEUR).value = "ROLL_HAUTEUR_Y"
        ws.Cells(1, TECH_COL_FOOT_HEIGHT).value = "FOOT_HEIGHT"
        ws.Cells(1, TECH_COL_PIED_MODELE).value = "PIED_MODELE_CREO"
        ws.Cells(1, TECH_COL_MOTEUR_SOURCE).value = "MOTEUR_SOURCE_CREO"
        ws.Visible = xlSheetVeryHidden
    End If
    Set TECH_ObtenirFeuilleEtat = ws
End Function

Private Function TECH_TrouverLigneEtatRoll( _
    ByVal ws As Worksheet, _
    ByVal nomAssemblage As String, _
    ByVal rollFeatureID As Long) As Long

    Dim derniereLigne As Long
    Dim r As Long
    derniereLigne = ws.Cells(ws.Rows.Count, TECH_COL_CLE).End(xlUp).Row
    For r = 2 To derniereLigne
        If StrComp(Trim$(CStr(ws.Cells(r, TECH_COL_ASSEMBLAGE).value)), _
            nomAssemblage, vbTextCompare) = 0 Then
            If CLng(Val(CStr(ws.Cells(r, TECH_COL_ROLL_ID).value))) = rollFeatureID Then
                TECH_TrouverLigneEtatRoll = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Function TECH_TrouverLigneEtat( _
    ByVal ws As Worksheet, _
    ByVal cle As String) As Long

    Dim derniereLigne As Long
    Dim r As Long
    derniereLigne = ws.Cells(ws.Rows.Count, TECH_COL_CLE).End(xlUp).Row
    For r = 2 To derniereLigne
        If StrComp(Trim$(CStr(ws.Cells(r, TECH_COL_CLE).value)), _
            cle, vbTextCompare) = 0 Then
            TECH_TrouverLigneEtat = r
            Exit Function
        End If
    Next r
End Function

Private Sub TECH_EcrireEtat( _
    ByVal ws As Worksheet, _
    ByVal ligne As Long, _
    ByVal nomAssemblage As String, _
    ByVal cle As String, _
    ByVal informations As Variant, _
    ByVal nomRoll As String, _
    ByVal rollFeatureID As Long)

    If ligne <= 0 Then
        ligne = ws.Cells(ws.Rows.Count, TECH_COL_CLE).End(xlUp).Row + 1
        If ligne < 2 Then ligne = 2
    End If
    ws.Cells(ligne, TECH_COL_ASSEMBLAGE).value = nomAssemblage
    ws.Cells(ligne, TECH_COL_CLE).value = cle
    ws.Cells(ligne, TECH_COL_SOURCE_ID).value = CLng(informations(0))
    ws.Cells(ligne, TECH_COL_SOURCE_MODELE).value = CStr(informations(1))
    ws.Cells(ligne, TECH_COL_SOURCE_PIECE).value = CStr(informations(2))
    ws.Cells(ligne, TECH_COL_OCCURRENCE).value = CLng(informations(3))
    ws.Cells(ligne, TECH_COL_REPERE).value = CStr(informations(4))
    ws.Cells(ligne, TECH_COL_DIAMETRE).value = CDbl(informations(5))
    ws.Cells(ligne, TECH_COL_LARGEUR).value = CDbl(informations(6))
    ws.Cells(ligne, TECH_COL_ROLL_MODELE).value = nomRoll
    ws.Cells(ligne, TECH_COL_ROLL_ID).value = rollFeatureID
    ws.Cells(ligne, TECH_COL_STATUT).value = "OK"
    ws.Cells(ligne, TECH_COL_DATE).value = Now
End Sub

Private Sub TECH_ValiderSolid( _
    ByVal solid As pfcls.IpfcSolid, _
    ByVal contexte As String)

    Dim failed As pfcls.IpfcFeatures
    If solid Is Nothing Then
        Err.Raise vbObjectError + 15130, , contexte & " : Solid Nothing."
    End If
    TECH_RegenererSolideFiable solid, contexte
    Set failed = Nothing
    On Error Resume Next
    Set failed = solid.ListFailedFeatures()
    Err.Clear
    On Error GoTo Erreur
    If Not failed Is Nothing Then
        If failed.Count > 0 Then
            Err.Raise vbObjectError + 15131, , _
                contexte & " : " & CStr(failed.Count) & " feature(s) en echec."
        End If
    End If
    Exit Sub

Erreur:
    Err.Raise Err.Number, , "TECH_ValiderSolid :" & vbCrLf & Err.description
End Sub

Private Sub TECH_RegenererSolideFiable( _
    ByVal solid As pfcls.IpfcSolid, _
    ByVal contexte As String)

    Dim tentative As Long
    Dim numeroErreur As Long
    Dim descriptionErreur As String
    Dim sourceErreur As String

    If solid Is Nothing Then
        Err.Raise vbObjectError + 15132, , _
            contexte & " : regeneration impossible, Solid Nothing."
    End If

    ' Creo renvoie XToolkitRegenerateAgain lorsqu'une premiere passe a
    ' modifie le modele et qu'une nouvelle regeneration est necessaire.
    ' La logique est identique a Plan_RegenererSolideFiable du pack source.
    For tentative = 1 To 3
        numeroErreur = 0
        descriptionErreur = ""
        sourceErreur = ""

        On Error Resume Next
        Err.Clear
        solid.Regenerate Nothing
        numeroErreur = Err.Number
        descriptionErreur = Err.description
        sourceErreur = Err.source
        Err.Clear
        On Error GoTo 0

        If numeroErreur = 0 Then Exit Sub

        Debug.Print "REGEN TECHNIQUE ECHEC essai " & CStr(tentative) & _
            "/3 : " & contexte & " / " & descriptionErreur
        DoEvents
    Next tentative

    ' PTC definit XToolkitRegenerateAgain comme une regeneration reussie :
    ' seules les relations demandent une passe supplementaire. Certains
    ' modeles peuvent renvoyer cet etat a chaque appel. Apres les trois
    ' passes, on poursuit donc le controle des features en echec.
    If TECH_EstRegenerateAgain(sourceErreur, descriptionErreur) Then
        Debug.Print "REGEN TECHNIQUE ACCEPTEE apres 3 passes : " & contexte
        Exit Sub
    End If

    Err.Raise numeroErreur, sourceErreur, _
        contexte & " : regeneration impossible apres 3 passes." & _
        IIf(descriptionErreur <> "", vbCrLf & descriptionErreur, "")
End Sub

Private Function TECH_EstRegenerateAgain( _
    ByVal sourceErreur As String, _
    ByVal descriptionErreur As String) As Boolean

    TECH_EstRegenerateAgain = _
        (InStr(1, sourceErreur, "XToolkitRegenerateAgain", vbTextCompare) > 0) Or _
        (InStr(1, descriptionErreur, "XToolkitRegenerateAgain", vbTextCompare) > 0)
End Function

Private Function TECH_CleCible( _
    ByVal nomAssemblage As String, _
    ByVal sourceFeatureID As Long, _
    ByVal nomRepere As String) As String

    TECH_CleCible = LCase$(TECH_NomModeleLogique(nomAssemblage)) & "|" & _
        CStr(sourceFeatureID) & "|" & UCase$(Trim$(nomRepere))
End Function

Private Function TECH_NomVarianteRoll( _
    ByVal diametre As Double, _
    ByVal largeur As Double) As String

    TECH_NomVarianteRoll = TECH_PREFIXE_VARIANTE & _
        TECH_TexteNomNombre(diametre) & "_L" & TECH_TexteNomNombre(largeur)
End Function

Private Function TECH_NomVariantePied( _
    ByVal nomModeleSource As String, _
    ByVal footHeight As Double) As String

    ' Nom volontairement court et independant du nom de la piece interne.
    ' Cela evite les caracteres invisibles, les noms PDM trop longs et les
    ' anciennes variantes source qui provoquaient XToolkitInvalidName.
    TECH_NomVariantePied = TECH_PREFIXE_PIED & _
        TECH_TexteNomHauteur(footHeight)
End Function

Private Function TECH_NomVarianteMoteur( _
    ByVal nomModeleSource As String, _
    ByVal footHeight As Double) As String

    TECH_NomVarianteMoteur = TECH_PREFIXE_MOTEUR & _
        TECH_TexteNomHauteur(footHeight)
End Function

Private Function TECH_NomVarianteDepuisSource( _
    ByVal nomModeleSource As String, _
    ByVal suffixe As String) As String

    Dim base As String
    base = TECH_BaseSansExtension(nomModeleSource)
    If base = "" Then
        Err.Raise vbObjectError + 15227, , _
            "Nom du modele source vide pour la creation de variante."
    End If
    If Len(suffixe) >= 31 Then
        Err.Raise vbObjectError + 15228, , _
            "Suffixe de variante trop long : " & suffixe
    End If
    If Len(base) + Len(suffixe) > 31 Then
        base = Left$(base, 31 - Len(suffixe))
    End If
    TECH_NomVarianteDepuisSource = base & suffixe
End Function

Private Function TECH_TexteNomNombre(ByVal valeur As Double) As String
    Dim texte As String
    texte = Format$(Round(valeur, 6), "0.######")
    texte = Replace(texte, ",", "P")
    texte = Replace(texte, ".", "P")
    texte = Replace(texte, "-", "M")
    TECH_TexteNomNombre = texte
End Function

Private Function TECH_TexteNomHauteur(ByVal valeur As Double) As String
    Dim texte As String

    ' V16 : sans separateur final. Format$ "0.######" laisse "1380," en
    ' parametres regionaux francais, d'ou le nom FTF_H1380P. Le nommage
    ' des rouleaux (TECH_TexteNomNombre) est volontairement inchange pour
    ' retrouver les variantes ROLL_D deja creees.
    texte = TECH_TexteNomNombre(valeur)
    Do While Right$(texte, 1) = "P" And Len(texte) > 1
        texte = Left$(texte, Len(texte) - 1)
    Loop
    TECH_TexteNomHauteur = texte
End Function

Private Function TECH_EstModeleRoll(ByVal nomModele As String) As Boolean
    Dim base As String
    Dim suffixe As String
    Dim positionLargeur As Long
    base = UCase$(TECH_BaseSansExtension(nomModele))
    If base = "ROLL" Then
        TECH_EstModeleRoll = True
        Exit Function
    End If
    If Left$(base, Len(TECH_PREFIXE_VARIANTE)) <> TECH_PREFIXE_VARIANTE Then Exit Function
    suffixe = Mid$(base, Len(TECH_PREFIXE_VARIANTE) + 1)
    If suffixe = "" Then Exit Function
    positionLargeur = InStr(1, suffixe, "_L", vbTextCompare)
    If positionLargeur > 0 Then
        If Not TECH_EstSuffixeNombre(Left$(suffixe, positionLargeur - 1)) Then Exit Function
        If Not TECH_EstSuffixeNombre(Mid$(suffixe, positionLargeur + 2)) Then Exit Function
    Else
        ' Compatibilite avec les variantes creees par la premiere version.
        If Not TECH_EstSuffixeNombre(suffixe) Then Exit Function
    End If
    TECH_EstModeleRoll = True
End Function

Private Function TECH_EstModeleMoteur(ByVal nomModele As String) As Boolean
    Dim base As String
    Dim positionSuffixe As Long
    Dim nomAttendu As String
    base = UCase$(TECH_BaseSansExtension(nomModele))
    If Left$(base, Len(TECH_PREFIXE_MOTEUR)) = _
        UCase$(TECH_PREFIXE_MOTEUR) Then
        TECH_EstModeleMoteur = _
            TECH_EstSuffixeNombre(Mid$(base, Len(TECH_PREFIXE_MOTEUR) + 1))
        Exit Function
    End If
    If Left$(base, Len(TECH_PREFIXE_MOTEUR_COPIE)) = _
        UCase$(TECH_PREFIXE_MOTEUR_COPIE) Then
        TECH_EstModeleMoteur = TECH_EstSuffixeNombre( _
            Mid$(base, Len(TECH_PREFIXE_MOTEUR_COPIE) + 1))
        Exit Function
    End If
    If Left$(base, Len(TECH_PREFIXE_MOTEUR_HISTORIQUE)) = _
        UCase$(TECH_PREFIXE_MOTEUR_HISTORIQUE) Then
        TECH_EstModeleMoteur = TECH_EstSuffixeNombre( _
            Mid$(base, Len(TECH_PREFIXE_MOTEUR_HISTORIQUE) + 1))
        Exit Function
    End If
    positionSuffixe = InStrRev(base, "_MH", -1, vbTextCompare)
    If positionSuffixe > 1 Then
        If TECH_EstSuffixeNombre(Mid$(base, positionSuffixe + 3)) Then
            TECH_EstModeleMoteur = True
            Exit Function
        End If
    End If
    nomAttendu = m_TECH_MoteurFichier
    If nomAttendu = "" Then nomAttendu = TECH_MOTEUR_FICHIER
    TECH_EstModeleMoteur = _
        (StrComp(TECH_BaseSansExtension(nomModele), _
            TECH_BaseSansExtension(nomAttendu), vbTextCompare) = 0) Or _
        (StrComp(TECH_BaseSansExtension(nomModele), _
            TECH_BaseSansExtension(TECH_MOTEUR_FICHIER_HISTORIQUE), _
            vbTextCompare) = 0)
End Function

Private Function TECH_EstSuffixeNombre(ByVal suffixe As String) As Boolean
    Dim i As Long
    Dim ch As String
    If suffixe = "" Then Exit Function
    For i = 1 To Len(suffixe)
        ch = Mid$(suffixe, i, 1)
        If (ch < "0" Or ch > "9") And ch <> "P" And ch <> "M" Then Exit Function
    Next i
    TECH_EstSuffixeNombre = True
End Function

Private Function TECH_ModeleCorrespondPiece( _
    ByVal modelePhysique As String, _
    ByVal pieceExcel As String) As Boolean

    Dim modeleBase As String
    Dim pieceBase As String
    Dim pieceCourte As String
    Dim suffixes As Variant
    Dim suffixe As Variant

    modeleBase = LCase$(TECH_BaseSansExtension(modelePhysique))
    pieceBase = LCase$(TECH_BaseSansExtension(pieceExcel))
    If modeleBase = "" Or pieceBase = "" Then Exit Function
    If modeleBase = pieceBase Then
        TECH_ModeleCorrespondPiece = True
        Exit Function
    End If
    pieceCourte = pieceBase
    If Len(pieceCourte) > 20 Then pieceCourte = Left$(pieceCourte, 20)
    suffixes = Array("_v", "_s", "_r")
    For Each suffixe In suffixes
        If Left$(modeleBase, Len(pieceBase) + 2) = pieceBase & CStr(suffixe) Then
            TECH_ModeleCorrespondPiece = True
            Exit Function
        End If
        If Left$(modeleBase, Len(pieceCourte) + 2) = pieceCourte & CStr(suffixe) Then
            TECH_ModeleCorrespondPiece = True
            Exit Function
        End If
    Next suffixe
End Function

Private Function TECH_NomModeleLogique(ByVal nomModele As String) As String
    Dim positionPrt As Long
    Dim positionAsm As Long
    Dim positionDrw As Long
    Dim positionExtension As Long

    nomModele = Trim$(nomModele)
    positionPrt = InStrRev(nomModele, ".prt", -1, vbTextCompare)
    positionAsm = InStrRev(nomModele, ".asm", -1, vbTextCompare)
    positionDrw = InStrRev(nomModele, ".drw", -1, vbTextCompare)
    positionExtension = positionPrt
    If positionAsm > positionExtension Then positionExtension = positionAsm
    If positionDrw > positionExtension Then positionExtension = positionDrw
    If positionExtension > 0 Then
        TECH_NomModeleLogique = Left$(nomModele, positionExtension + 3)
    Else
        TECH_NomModeleLogique = nomModele
    End If
End Function

Private Function TECH_BaseSansExtension(ByVal nomModele As String) As String
    Dim logique As String
    Dim positionExtension As Long
    logique = TECH_NomModeleLogique(nomModele)
    positionExtension = InStrRev(logique, ".", -1, vbTextCompare)
    If positionExtension > 1 Then
        TECH_BaseSansExtension = Left$(logique, positionExtension - 1)
    Else
        TECH_BaseSansExtension = logique
    End If
End Function

Private Function TECH_NomFeuilleEtatSection(ByVal nomSection As String) As String
    Dim base As String
    Dim resultat As String
    Dim i As Long
    Dim code As Double
    Dim ch As String

    base = TECH_BaseSansExtension(nomSection)
    base = Replace(base, "\", "_")
    base = Replace(base, "/", "_")
    base = Replace(base, ":", "_")
    base = Replace(base, "*", "_")
    base = Replace(base, "?", "_")
    base = Replace(base, "[", "_")
    base = Replace(base, "]", "_")
    base = Replace(base, Chr$(34), "_")
    code = 0#
    For i = 1 To Len(nomSection)
        ch = Mid$(nomSection, i, 1)
        code = (code * 131# + AscW(ch)) Mod 1000000#
    Next i
    If Len(base) > 20 Then base = Left$(base, 20)
    resultat = "_ER_" & base & "_" & Format$(CLng(code), "000000")
    If Len(resultat) > 31 Then resultat = Left$(resultat, 31)
    TECH_NomFeuilleEtatSection = resultat
End Function

Private Function TECH_DoublesEgaux(ByVal a As Double, ByVal b As Double) As Boolean
    Dim tolerance As Double
    tolerance = 0.0000001 * (1# + Abs(a) + Abs(b))
    TECH_DoublesEgaux = (Abs(a - b) <= tolerance)
End Function
