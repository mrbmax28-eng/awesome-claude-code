Attribute VB_Name = "M_FRAME_HEATING_PALIERS"
Option Explicit

' ================================================================
' HEATING - PALIERS SOUS LES ROLL_CENTER  (version PAL_VERSION)
'
' Points d'entree publics :
'   HEAT_PlacerPaliersRouleaux     place les paliers (heating.asm actif)
'   HEAT_ControlerSqueletteHeating controle Creo, ne modifie rien
'   HEAT_VersionPaliers            affiche la version importee
'
' Regle garantie : chaque palier est a PAL_DISTANCE_PALIER (1500 mm) sous
' le centre du rouleau le plus bas de son groupe, et c'est la PREMIERE
' ligne du squelette sous chaque rouleau du groupe.
'
'   1. Lecture des reperes ROLL_CENTER / ROLLn_CENTER des pieces de
'      HEATING (sous-assemblages compris), dans le repere de heating.asm.
'   2. Mesure, dans Creo, du decalage d entre le Y Excel et le Y reel du
'      squelette dans heating.asm : les niveaux deja calcules
'      (Niveaux_HEATING, paliers exclus) sont retrouves parmi les lignes
'      horizontales du squelette. d couvre la position du squelette et tout
'      decalage de l'export IGES. Y palier Excel = Y rouleau - 1500 - d :
'      le palier est a 1500 mm SOUS le centre dans Creo, quel que soit d.
'      (Sans compensation, un d de 1694 mm donnait le palier 194 mm
'      au-dessus du rouleau.)
'   3. Groupes : rouleaux a 2000 mm maxi du plus bas du groupe. Un rouleau
'      seul rejoint le groupe le plus proche. Reference du palier = le
'      rouleau le plus bas du groupe FINAL (rouleau seul compris) : aucun
'      rouleau ne peut se retrouver sous son palier.
'   4. Y palier = Y reference - 1500. Pas de palier sous le pied des
'      poteaux (Structure_HEATING!G13).
'   5. NiveauxY_HEATING : un niveau PAL_xx par groupe. Sont desactives
'      (D = 0, marque [OFF PALIER]) les niveaux a moins de l'ecart mini
'      (G15) du palier ET ceux situes entre le palier et les rouleaux du
'      groupe : sinon le pied se calerait sur eux et non sur le palier.
'      Ils sont listes AVANT confirmation et reactives au lancement
'      suivant s'ils ne genent plus.
'   6. HEAT_Recalculer, puis verification dans Niveaux_HEATING : chaque
'      palier doit y etre, sans niveau entre lui et ses rouleaux. Sinon
'      tout est remis dans l'etat initial.
'   7. Export IGES propose.
'
' Aucun modele Creo n'est modifie par ce module.
'
' 2026.10.05.09 (pack pilote) :
'   - pieces porteuses reconnues avec les MEMES regles que le module
'     pieds moteurs (TECH_PUB_EstModeleRoll / TECH_PUB_EstModeleMoteur) :
'     un rouleau equipe d'un pied a forcement son palier ;
'   - arret si un composant n'est pas charge dans Creo (liste nommee) ;
'   - HEAT_PalPlacerAuto, HEAT_PalControlerAuto, HEAT_PalInventaire :
'     memes traitements sans fenetre, pour M_PILOTE_HEATING ;
'   - controle : rouleau porte par un sous-assemblage (son pied se calerait
'     sur un autre squelette) et repere ROLL dans un squelette = KO.
' Necessite Module_Ajout_Pieces_Techniques V18 et M_FRAME_HEATING du pack.
' ================================================================

Private Const PAL_VERSION As String = "2026.10.05.09"
Private Const PAL_ERR As Long = vbObjectError + 8200
Private Const PAL_FEUILLE As String = "PALIERS_HEATING"
Private Const PAL_LIGNE_DONNEES As Long = 7
Private Const PAL_TOLERANCE_DEFAUT As Double = 2000#
' Distance fixe, vers le bas, entre le centre du rouleau le plus bas d'un
' groupe et son palier.
Private Const PAL_DISTANCE_PALIER As Double = 1500#
Private Const PAL_DOUBLON As Double = 0.5
' Decalage vertical tolere du squelette dans heating.asm (mm).
Private Const PAL_TOL_SQUELETTE As Double = 0.01
' Ecart horizontal tolere pour qu'une ligne croise la verticale d'un
' rouleau (meme regle que le module pieds moteurs).
Private Const PAL_TOL_LIGNE As Double = 0.5
' Ecart de distance accepte par le controle (mm).
Private Const PAL_TOL_CONTROLE As Double = 1#
Private Const PAL_MARQUE As String = "[OFF PALIER] "
Private Const PAL_PREFIXE_ID As String = "PAL_"
Private Const PAL_NIV_FIRST As Long = 7
Private Const PAL_NIV_LAST As Long = 206
Private Const PAL_CALC_FIRST As Long = 7
Private Const PAL_CALC_LAST As Long = 86
Private Const PAL_MSG_MAX As Long = 760
' Configuration du pilote : chemin de l'assemblage moteur (regles moteur).
Private Const PAL_FEUILLE_PILOTE As String = "PILOTE_HEATING"
Private Const PAL_CELLULE_MOTEUR As String = "C6"

' Bilan du dernier parcours (PalParcourir).
Private m_PalNonCharges As String
Private m_PalReperesSquelette As Long

Public Sub HEAT_VersionPaliers()
    MsgBox "M_FRAME_HEATING_PALIERS version " & PAL_VERSION, _
        vbInformation, "Paliers HEATING"
End Sub

Public Sub HEAT_PlacerPaliersRouleaux()
    Dim message As String
    PalPlacer False, True, message
End Sub

' Pilote : placement sans fenetre (sauf la confirmation si confirmer).
' Pas d'export IGES ici : le pilote l'enchaine.
Public Function HEAT_PalPlacerAuto(ByVal confirmer As Boolean, _
    ByRef message As String) As Boolean
    HEAT_PalPlacerAuto = PalPlacer(True, confirmer, message)
End Function

Public Function HEAT_PalVersionTexte() As String
    HEAT_PalVersionTexte = PAL_VERSION
End Function

' Pilote : inventaire avant tout traitement, sans rien modifier.
' Vrai si l'inventaire a pu etre fait ; nbBloquants = problemes qui
' empecheraient des paliers et des pieds justes ; texte = compte rendu.
Public Function HEAT_PalInventaire(ByRef nbBloquants As Long, _
    ByRef texte As String) As Boolean
    Dim session As pfcls.IpfcBaseSession
    Dim modele As pfcls.IpfcModel
    Dim asm As pfcls.IpfcAssembly
    Dim solide As pfcls.IpfcSolid
    Dim squelette As pfcls.IpfcSolid
    Dim pose As pfcls.IpfcTransform3D
    Dim nomSquelette As String
    Dim ids As pfcls.Iintseq
    Dim rouleaux As Collection
    Dim ignores As Long
    Dim reperesAsm As Long
    Dim nbSousAsm As Long
    Dim listeSousAsm As String
    Dim i As Long

    On Error GoTo Echec
    nbBloquants = 0
    texte = ""
    Set session = PalSession()
    Set modele = session.GetActiveModel()
    If modele Is Nothing Then Err.Raise PAL_ERR + 1, , "Aucun modele actif dans Creo."
    If modele.Type <> pfcls.EpfcMDL_ASSEMBLY Then Err.Raise PAL_ERR + 2, , _
        "Le modele actif doit etre l'assemblage HEATING (.asm)."
    Set asm = modele
    Set solide = modele
    PalAjouterLigne texte, "Assemblage actif : " & modele.Filename
    If Not PalEstHeating(modele.Filename) Then
        nbBloquants = nbBloquants + 1
        PalAjouterLigne texte, "KO  " & modele.Filename & " n'est pas reconnu comme " & _
            "l'assemblage HEATING (nom ou ZONES_FOUR!X8)."
    End If
    PalExigerMillimetres solide, modele.Filename
    If PalPoseSquelette(session, asm, solide, squelette, nomSquelette, pose) Then
        PalAjouterLigne texte, "OK  squelette " & nomSquelette & " (Y " & _
            PalFmt(PalOrigine(pose, 1)) & " mm dans l'assemblage)"
    Else
        nbBloquants = nbBloquants + 1
        PalAjouterLigne texte, "KO  aucun squelette dans " & modele.Filename & _
            " : le creer (HEAT_CreerSqueletteCreo) avant de lancer le pilote."
    End If

    Set rouleaux = New Collection
    Set ids = New pfcls.Cintseq
    reperesAsm = PalCompterReperesAsm(modele)
    PalDebutParcours
    PalParcourir session, asm, solide, ids, 0, rouleaux, ignores, reperesAsm, _
        modele.Filename
    If ignores > 0 Then
        nbBloquants = nbBloquants + 1
        PalAjouterLigne texte, "KO  " & ignores & " composant(s) non charge(s) : " & _
            PalCouper(m_PalNonCharges, 300)
    Else
        PalAjouterLigne texte, "OK  tous les composants sont charges"
    End If
    If m_PalReperesSquelette > 0 Then
        nbBloquants = nbBloquants + 1
        PalAjouterLigne texte, "KO  " & m_PalReperesSquelette & _
            " repere(s) ROLL dans un squelette (a supprimer)"
    End If
    If rouleaux.Count = 0 Then
        nbBloquants = nbBloquants + 1
        PalAjouterLigne texte, "KO  aucun repere ROLL_CENTER dans les pieces"
    Else
        PalAjouterLigne texte, "OK  " & rouleaux.Count & " repere(s) ROLL_CENTER trouves"
    End If
    For i = 1 To rouleaux.Count
        If StrComp(PalChampTexte(rouleaux, i, 6), modele.Filename, vbTextCompare) <> 0 Then
            nbSousAsm = nbSousAsm + 1
            If nbSousAsm <= 3 Then
                If listeSousAsm <> "" Then listeSousAsm = listeSousAsm & ", "
                listeSousAsm = listeSousAsm & LCase$(PalBase(PalChampTexte(rouleaux, i, 0))) & _
                    " dans " & PalChampTexte(rouleaux, i, 6)
            End If
        End If
    Next i
    If nbSousAsm > 0 Then
        nbBloquants = nbBloquants + 1
        PalAjouterLigne texte, "KO  " & nbSousAsm & " rouleau(x) porte(s) par un " & _
            "sous-assemblage : leur pied se calerait sur un autre squelette que " & _
            "celui des paliers (" & listeSousAsm & IIf(nbSousAsm > 3, ", ...", "") & ")"
    End If
    If reperesAsm > 0 Then PalAjouterLigne texte, "ALERTE  " & reperesAsm & _
        " repere(s) ROLL_CENTER poses dans un assemblage : ignores par les deux modules"
    HEAT_PalInventaire = True
    Exit Function
Echec:
    nbBloquants = nbBloquants + 1
    PalAjouterLigne texte, "KO  inventaire interrompu : " & Err.Description
End Function

' Pilote : le squelette de l'assemblage actif correspond-il aux niveaux
' Excel (Niveaux_HEATING, hors paliers) ? d = decalage Excel -> Creo.
Public Function HEAT_PalDecalageMesurable(ByRef d As Double, _
    ByRef probleme As String) As Boolean
    Dim session As pfcls.IpfcBaseSession
    Dim modele As pfcls.IpfcModel
    Dim asm As pfcls.IpfcAssembly
    Dim solide As pfcls.IpfcSolid
    Dim squelette As pfcls.IpfcSolid
    Dim pose As pfcls.IpfcTransform3D
    Dim nomSquelette As String
    Dim lignes As Collection
    Dim autres As Long
    Dim nbRef As Long

    On Error GoTo Echec
    probleme = ""
    Set session = PalSession()
    Set modele = session.GetActiveModel()
    If modele Is Nothing Then Err.Raise PAL_ERR + 1, , "Aucun modele actif dans Creo."
    If modele.Type <> pfcls.EpfcMDL_ASSEMBLY Then Err.Raise PAL_ERR + 2, , _
        "Le modele actif doit etre l'assemblage HEATING (.asm)."
    Set asm = modele
    Set solide = modele
    If Not PalPoseSquelette(session, asm, solide, squelette, nomSquelette, pose) Then
        probleme = "aucun squelette dans " & modele.Filename & "."
        Exit Function
    End If
    Set lignes = PalLignesSquelette(squelette, pose, autres)
    HEAT_PalDecalageMesurable = PalMesurerDecalage(lignes, d, nbRef, probleme)
    If HEAT_PalDecalageMesurable Then probleme = "squelette " & nomSquelette & _
        " a jour (" & nbRef & " niveau(x) retrouve(s), d = " & PalFmt(d) & " mm)."
    Exit Function
Echec:
    probleme = Err.Description
End Function

Private Sub PalAjouterLigne(ByRef texte As String, ByVal ligne As String)
    If texte <> "" Then texte = texte & vbCrLf
    texte = texte & ligne
End Sub

Private Function PalPlacer(ByVal auto As Boolean, ByVal confirmer As Boolean, _
    ByRef message As String) As Boolean
    Dim session As pfcls.IpfcBaseSession
    Dim modele As pfcls.IpfcModel
    Dim asm As pfcls.IpfcAssembly
    Dim solide As pfcls.IpfcSolid
    Dim ids As pfcls.Iintseq
    Dim rouleaux As Collection
    Dim ignores As Long
    Dim reperesAsm As Long
    Dim n As Long
    Dim ys() As Double
    Dim idx() As Long
    Dim debutGroupe() As Long
    Dim finGroupe() As Long
    Dim nbGroupes As Long
    Dim yRef() As Double
    Dim yHaut() As Double
    Dim abaisse() As Boolean
    Dim squelette As pfcls.IpfcSolid
    Dim pose As pfcls.IpfcTransform3D
    Dim nomSquelette As String
    Dim aSquelette As Boolean
    Dim decalageSquelette As Double
    Dim decalageY As Double
    Dim nbReferences As Long
    Dim lignesSq As Collection
    Dim autresSq As Long
    Dim origineDecalage As String
    Dim actif() As Boolean
    Dim nbActifs As Long
    Dim distance() As Double
    Dim yPalier() As Double
    Dim lignesOff As Collection
    Dim offGroupe() As String
    Dim nbDesactives As Long
    Dim listeDesactives As String
    Dim ws As Worksheet
    Dim niveaux As Worksheet
    Dim structure As Worksheet
    Dim bottomY As Double
    Dim topY As Double
    Dim topOK As Boolean
    Dim minGap As Double
    Dim instantane As Variant
    Dim instantanePris As Boolean
    Dim feuilleAvant As Variant
    Dim feuillePrise As Boolean
    Dim recalcLance As Boolean
    Dim probleme As String
    Dim choix As VbMsgBoxResult
    Dim i As Long
    Dim g As Long
    Dim etape As String
    Dim nomAssemblage As String
    Dim errNumero As Long
    Dim errTexte As String

    On Error GoTo Echec
    message = ""
    If auto Then HEAT_DefinirSilencieux True

    etape = "Preparation des feuilles HEATING"
    HEAT_Initialiser
    Set structure = ThisWorkbook.Worksheets("Structure_HEATING")
    Set niveaux = ThisWorkbook.Worksheets("NiveauxY_HEATING")
    PalLireStructure structure, bottomY, topY, topOK, minGap

    etape = "Connexion Creo"
    Set session = PalSession()
    Set modele = session.GetActiveModel()
    If modele Is Nothing Then Err.Raise PAL_ERR + 1, , _
        "Activer l'assemblage HEATING dans Creo."
    If modele.Type <> pfcls.EpfcMDL_ASSEMBLY Then Err.Raise PAL_ERR + 2, , _
        "Le modele actif doit etre l'assemblage HEATING (.asm)."
    nomAssemblage = modele.Filename
    If Not PalEstHeating(nomAssemblage) Then
        If auto Then Err.Raise PAL_ERR + 5, , "Le modele actif " & nomAssemblage & _
            " n'est pas reconnu comme l'assemblage HEATING (ZONES_FOUR!X8)."
        If MsgBox("Le modele actif est " & nomAssemblage & "." & vbCrLf & _
            "Il n'est pas reconnu comme l'assemblage HEATING." & vbCrLf & _
            "Les Y seront lus dans le repere de cet assemblage. Continuer ?", _
            vbYesNo + vbQuestion, "Paliers HEATING") <> vbYes Then GoTo Abandon
    End If
    Set asm = modele
    Set solide = modele
    PalExigerMillimetres solide, nomAssemblage

    ' Les ROLL_CENTER sont lus dans le repere de l'assemblage, les paliers
    ' sont ecrits en Y Excel. Le decalage d entre les deux est MESURE dans
    ' le squelette, puis compense : Y palier Excel = Y rouleau - 1500 - d.
    etape = "Mesure du decalage Excel -> Creo dans " & nomAssemblage
    decalageY = 0#
    origineDecalage = "pas de squelette : d = 0 suppose (squelette a creer a l'origine)"
    aSquelette = PalPoseSquelette(session, asm, solide, squelette, nomSquelette, pose)
    If aSquelette Then
        decalageSquelette = PalOrigine(pose, 1)
        Set lignesSq = PalLignesSquelette(squelette, pose, autresSq)
        If PalMesurerDecalage(lignesSq, decalageY, nbReferences, probleme) Then
            origineDecalage = "mesure sur " & nbReferences & " niveau(x) du squelette " & _
                nomSquelette
        Else
            If auto Then Err.Raise PAL_ERR + 63, , _
                "Decalage Excel -> Creo impossible a mesurer dans " & nomSquelette & _
                " : " & probleme & " Recharger l'IGES dans le squelette puis relancer."
            If MsgBox("Decalage Excel -> Creo impossible a mesurer dans " & _
                nomSquelette & " :" & vbCrLf & probleme & vbCrLf & vbCrLf & _
                "Solution sure : NON, puis HEAT_Recalculer, export IGES, recharger " & _
                "l'IGES dans le squelette et relancer." & vbCrLf & vbCrLf & _
                "OUI = continuer avec la seule position du squelette (Y " & _
                PalFmt(decalageSquelette) & " mm)." & vbCrLf & _
                "NON = arreter sans rien modifier", vbYesNo + vbExclamation, _
                "Paliers HEATING") <> vbYes Then GoTo Abandon
            decalageY = decalageSquelette
            origineDecalage = "NON MESURE : position du squelette seule"
        End If
    End If

    etape = "Lecture des ROLL_CENTER dans " & nomAssemblage
    Set rouleaux = New Collection
    Set ids = New pfcls.Cintseq
    reperesAsm = PalCompterReperesAsm(modele)
    PalDebutParcours
    PalParcourir session, asm, solide, ids, 0, rouleaux, ignores, reperesAsm, _
        nomAssemblage
    ' Aucun rouleau ne doit echapper aux paliers : tout doit etre charge.
    If ignores > 0 Then Err.Raise PAL_ERR + 6, , ignores & _
        " composant(s) non charge(s) dans Creo : " & PalCouper(m_PalNonCharges, 400) & _
        ". Ouvrir " & nomAssemblage & " avec tous ses composants, puis relancer. " & _
        "Rien n'a ete modifie."
    If m_PalReperesSquelette > 0 Then Err.Raise PAL_ERR + 7, , _
        m_PalReperesSquelette & " repere(s) ROLL_CENTER dans un squelette : le " & _
        "module moteur y poserait un rouleau. Les supprimer du squelette. " & _
        "Rien n'a ete modifie."
    If rouleaux.Count = 0 Then Err.Raise PAL_ERR + 3, , _
        "Aucun repere ROLL_CENTER / ROLLn_CENTER dans les pieces de " & _
        nomAssemblage & "."

    etape = "Tri des ROLL_CENTER"
    n = rouleaux.Count
    ReDim ys(1 To n)
    ReDim idx(1 To n)
    For i = 1 To n
        idx(i) = i
        ys(i) = PalChamp(rouleaux, i, 3)
    Next i
    PalTrier ys, idx, n

    etape = "Regroupement des rouleaux proches en hauteur"
    Set ws = PalFeuille()
    PalGrouper ys, idx, n, PAL_TOLERANCE_DEFAUT, debutGroupe, finGroupe, _
        nbGroupes, yRef, abaisse

    etape = "Calcul des paliers"
    ReDim yHaut(1 To nbGroupes)
    ReDim distance(1 To nbGroupes)
    ReDim yPalier(1 To nbGroupes)
    ReDim actif(1 To nbGroupes)
    nbActifs = 0
    For g = 1 To nbGroupes
        distance(g) = PAL_DISTANCE_PALIER
        ' yRef et ys : repere de l'assemblage. yPalier et yHaut : Y Excel.
        yHaut(g) = ys(finGroupe(g)) - decalageY
        yPalier(g) = Round(yRef(g) - distance(g) - decalageY, 6)
        ' Securite reelle, dans le repere Creo : le palier est exactement a
        ' la distance voulue SOUS le plus bas, et au moins a cette distance
        ' sous CHAQUE rouleau du groupe.
        If Abs(yRef(g) - (yPalier(g) + decalageY) - distance(g)) > 0.001 Then _
            Err.Raise PAL_ERR + 33, , "Controle de securite : distance du palier de '" & _
            PalNomRangee(g, nbGroupes) & "' differente de " & PalFmt(distance(g)) & _
            " mm. Aucune modification faite."
        For i = debutGroupe(g) To finGroupe(g)
            If ys(i) - (yPalier(g) + decalageY) < distance(g) - 0.001 Then Err.Raise PAL_ERR + 34, , _
                "Controle de securite : le palier de '" & PalNomRangee(g, nbGroupes) & _
                "' serait a moins de " & PalFmt(distance(g)) & " mm sous " & _
                PalChampTexte(rouleaux, idx(i), 0) & " (" & _
                PalChampTexte(rouleaux, idx(i), 1) & "). Aucune modification faite."
        Next i
        actif(g) = (yPalier(g) >= bottomY - 0.000001)
        If actif(g) Then nbActifs = nbActifs + 1
    Next g

    etape = "Controle des paliers"
    PalControler yPalier, yHaut, actif, nbGroupes, topY, topOK, minGap

    etape = "Niveaux genants dans NiveauxY_HEATING"
    Set lignesOff = PalPlanifierNiveaux(niveaux, yPalier, yHaut, actif, _
        nbGroupes, minGap, offGroupe, nbDesactives, listeDesactives)

    etape = "Proposition"
    ' La feuille est remise en etat si l'on arrete ou si l'on annule :
    ' elle doit toujours decrire les paliers presents dans NiveauxY_HEATING.
    feuilleAvant = ws.Range("A2:K" & (PAL_LIGNE_DONNEES + 500)).Formula
    feuillePrise = True
    PalEcrireProposition ws, rouleaux, ys, idx, debutGroupe, finGroupe, _
        nbGroupes, nomAssemblage, yRef, yPalier, distance, actif, offGroupe
    choix = vbYes
    If confirmer Then
    ws.Activate
    choix = MsgBox(PalTexteConfirmation(rouleaux, idx, debutGroupe, finGroupe, _
        nbGroupes, yRef, yPalier, actif, abaisse, offGroupe, ignores, _
        reperesAsm, nomAssemblage, aSquelette, bottomY, decalageY, _
        origineDecalage), _
        vbYesNo + vbQuestion, "Paliers HEATING - verification")
    End If
    If choix <> vbYes Then
        PalRestaurerFeuille ws, feuilleAvant, "ARRETE par l'utilisateur"
        feuillePrise = False
        message = "Arret demande a la confirmation : rien n'a ete modifie."
        If Not auto Then MsgBox message, vbInformation, "Paliers HEATING"
        GoTo Abandon
    End If

    etape = "Ecriture de NiveauxY_HEATING"
    instantane = niveaux.Range(niveaux.Cells(PAL_NIV_FIRST, 1), _
        niveaux.Cells(PAL_NIV_LAST, 5)).Formula
    instantanePris = True
    PalEcrireNiveaux niveaux, yPalier, distance, actif, nbGroupes, lignesOff
    PalEcrireResultats ws, distance, yPalier, actif, nbGroupes, decalageY

    etape = "Recalcul HEATING"
    recalcLance = True
    HEAT_Recalculer
    If Not HEAT_CalculValide() Then
        PalAnnuler niveaux, instantane, ws, feuilleAvant, _
            "ANNULE : HEATING non recalcule"
        instantanePris = False
        feuillePrise = False
        message = "Le recalcul HEATING a echoue avec les paliers (" & _
            HEAT_DerniereErreur() & "). NiveauxY_HEATING est remis dans son etat " & _
            "initial et HEATING a ete recalcule tel qu'avant."
        If Not auto Then MsgBox message, vbExclamation, "Paliers HEATING"
        GoTo Abandon
    End If

    etape = "Verification de Niveaux_HEATING"
    If Not PalVerifierCalcul(yPalier, yHaut, actif, nbGroupes, probleme) Then
        PalAnnuler niveaux, instantane, ws, feuilleAvant, "ANNULE : " & probleme
        instantanePris = False
        feuillePrise = False
        message = "Verification apres recalcul : " & probleme & _
            " NiveauxY_HEATING est remis dans son etat initial et HEATING " & _
            "a ete recalcule tel qu'avant."
        If Not auto Then MsgBox message, vbExclamation, "Paliers HEATING"
        GoTo Abandon
    End If
    instantanePris = False
    feuillePrise = False

    message = nbActifs & " palier(s) a " & PalFmt(PAL_DISTANCE_PALIER) & _
        " mm sous le rouleau le plus bas de leur groupe (d = " & PalFmt(decalageY) & _
        " mm, " & origineDecalage & "). " & n & " rouleau(x), " & nbGroupes & _
        " groupe(s), " & nbDesactives & " niveau(x) desactive(s)" & _
        IIf(nbDesactives > 0, " : " & PalCouper(listeDesactives, 300), "") & "."
    PalPlacer = True
    If auto Then
        HEAT_DefinirSilencieux False
        Exit Function
    End If

    If MsgBox(nbActifs & " palier(s) place(s) a " & PalFmt(PAL_DISTANCE_PALIER) & _
        " mm sous le rouleau le plus bas de leur groupe" & _
        IIf(nbGroupes > nbActifs, " (" & (nbGroupes - nbActifs) & _
            " groupe(s) sans palier : trop bas).", ".") & _
        vbCrLf & "Verifie dans Niveaux_HEATING apres recalcul." & _
        vbCrLf & nbDesactives & " niveau(x) desactive(s)" & _
        IIf(nbDesactives > 0, " : " & PalCouper(listeDesactives, 300), ".") & _
        vbCrLf & vbCrLf & "Exporter maintenant l'IGES des axes HEATING ?" & _
        vbCrLf & "(Recharger ensuite l'IGES dans la fonction importee du " & _
        "squelette existant, puis lancer HEAT_ControlerSqueletteHeating.)", _
        vbYesNo + vbQuestion, "Paliers HEATING") = vbYes Then
        HEAT_ExporterAxesIGES
    End If
    Exit Function

Abandon:
    If auto Then HEAT_DefinirSilencieux False
    If message = "" Then message = "Arret demande : rien n'a ete modifie."
    Exit Function

Echec:
    errNumero = Err.Number
    errTexte = Err.Description
    On Error GoTo -1
    On Error Resume Next
    If instantanePris Then
        niveaux.Range(niveaux.Cells(PAL_NIV_FIRST, 1), _
            niveaux.Cells(PAL_NIV_LAST, 5)).Formula = instantane
        If recalcLance Then HEAT_Recalculer
    End If
    If feuillePrise Then PalRestaurerFeuille ws, feuilleAvant, "ANNULE : " & etape
    If auto Then HEAT_DefinirSilencieux False
    On Error GoTo 0
    message = "Etape : " & etape & vbCrLf & _
        "Erreur : " & CStr(errNumero) & vbCrLf & errTexte & _
        IIf(instantanePris, vbCrLf & _
            "NiveauxY_HEATING a ete remis dans son etat initial" & _
            IIf(recalcLance, " et HEATING recalcule.", "."), "")
    If Not auto Then MsgBox "Paliers HEATING interrompus." & vbCrLf & message, _
        vbCritical, "Paliers HEATING"
End Function

Private Sub PalAnnuler(ByVal niveaux As Worksheet, ByRef instantane As Variant, _
    ByVal ws As Worksheet, ByRef feuilleAvant As Variant, ByVal texte As String)
    niveaux.Range(niveaux.Cells(PAL_NIV_FIRST, 1), _
        niveaux.Cells(PAL_NIV_LAST, 5)).Formula = instantane
    PalRestaurerFeuille ws, feuilleAvant, texte
    HEAT_Recalculer
End Sub

' ---------------------------------------------------------------- Creo

Private Function PalSession() As pfcls.IpfcBaseSession
    Dim fabrique As pfcls.CCpfcAsyncConnection
    Dim connexion As pfcls.IpfcAsyncConnection

    Set fabrique = New pfcls.CCpfcAsyncConnection
    On Error Resume Next
    Set connexion = fabrique.GetActiveConnection()
    If connexion Is Nothing Then
        Err.Clear
        Set connexion = fabrique.Connect("", "", "", 30)
    End If
    Err.Clear
    On Error GoTo 0
    If connexion Is Nothing Then Err.Raise PAL_ERR + 20, , _
        "Creo indisponible ; verifier pfcls et la connexion."
    Set PalSession = connexion.session
End Function

Private Sub PalExigerMillimetres(ByVal solide As pfcls.IpfcSolid, _
    ByVal nom As String)
    Dim systeme As pfcls.IpfcUnitSystem
    Dim unite As pfcls.IpfcUnit

    Set systeme = solide.GetPrincipalUnits()
    Set unite = systeme.GetUnit(pfcls.EpfcUNIT_LENGTH)
    If UCase$(unite.name) <> "MM" Then Err.Raise PAL_ERR + 4, , _
        nom & " : unite de longueur " & unite.name & " ; millimetres requis."
End Sub

Private Sub PalDebutParcours()
    m_PalNonCharges = ""
    m_PalReperesSquelette = 0
End Sub

' Parcours de l'assemblage : memes regles que le module pieds moteurs
' (assemblages moteurs sautes, rouleaux Roll/ROLL_D ignores, toute autre
' piece portant un repere ROLL est une piece porteuse).
Private Sub PalParcourir(ByVal session As pfcls.IpfcBaseSession, _
    ByVal racine As pfcls.IpfcAssembly, _
    ByVal courant As pfcls.IpfcSolid, _
    ByVal ids As pfcls.Iintseq, _
    ByVal profondeur As Long, _
    ByVal rouleaux As Collection, _
    ByRef ignores As Long, _
    ByRef reperesAsm As Long, _
    ByVal nomCourant As String)

    Dim features As pfcls.IpfcFeatures
    Dim feature As pfcls.IpfcFeature
    Dim item As pfcls.IpfcModelItem
    Dim composant As pfcls.IpfcComponentFeat
    Dim modele As pfcls.IpfcModel
    Dim enfant As pfcls.Iintseq
    Dim solideEnfant As pfcls.IpfcSolid
    Dim base As String
    Dim nomManquant As String
    Dim i As Long

    If profondeur > 32 Then Err.Raise PAL_ERR + 10, , _
        "Assemblages imbriques sur plus de 32 niveaux."
    Set features = courant.ListFeaturesByType(False, pfcls.EpfcFEATTYPE_COMPONENT)
    If features Is Nothing Then Exit Sub
    For i = 0 To features.Count - 1
        Set feature = features.item(i)
        Set item = feature
        Set composant = feature
        Set modele = Nothing
        On Error Resume Next
        Set modele = session.GetModelFromDescr(composant.ModelDescr)
        Err.Clear
        On Error GoTo 0
        If modele Is Nothing Then
            ignores = ignores + 1
            nomManquant = ""
            On Error Resume Next
            nomManquant = composant.ModelDescr.GetFileName()
            Err.Clear
            On Error GoTo 0
            If nomManquant = "" Then nomManquant = "composant ID " & item.id
            If m_PalNonCharges <> "" Then m_PalNonCharges = m_PalNonCharges & ", "
            m_PalNonCharges = m_PalNonCharges & nomManquant & " (dans " & nomCourant & ")"
        Else
            Set enfant = PalCopierIds(ids, item.id)
            base = UCase$(PalBase(modele.Filename))
            If modele.Type = pfcls.EpfcMDL_ASSEMBLY Then
                If Not PalEstMoteur(modele.Filename) Then
                    reperesAsm = reperesAsm + PalCompterReperesAsm(modele)
                    Set solideEnfant = modele
                    PalParcourir session, racine, solideEnfant, enfant, _
                        profondeur + 1, rouleaux, ignores, reperesAsm, modele.Filename
                End If
            ElseIf modele.Type = pfcls.EpfcMDL_PART Then
                If TECH_PUB_EstModeleRoll(modele.Filename) Then
                    ' Rouleau ajoute par le module moteur : pas une piece porteuse.
                ElseIf PalEstSquelette(modele) Then
                    m_PalReperesSquelette = m_PalReperesSquelette + _
                        PalCompterReperesAsm(modele)
                Else
                    PalLireReperes racine, modele, enfant, rouleaux, nomCourant
                End If
            End If
        End If
    Next i
End Sub

' Reperes ROLL_CENTER poses au niveau d'un ASSEMBLAGE : ils ne sont pas
' utilises (seuls ceux des pieces comptent), mais ils sont signales.
Private Function PalCompterReperesAsm(ByVal modele As pfcls.IpfcModel) As Long
    Dim proprietaire As pfcls.IpfcModelItemOwner
    Dim items As pfcls.IpfcModelItems
    Dim nom As String
    Dim i As Long

    On Error Resume Next
    Set proprietaire = modele
    Set items = proprietaire.ListItems(pfcls.EpfcITEM_COORD_SYS)
    If Err.Number <> 0 Then Set items = Nothing
    Err.Clear
    On Error GoTo 0
    If items Is Nothing Then Exit Function
    For i = 0 To items.Count - 1
        nom = ""
        On Error Resume Next
        nom = items.item(i).GetName()
        Err.Clear
        On Error GoTo 0
        If PalEstRepereRoll(nom) Then PalCompterReperesAsm = PalCompterReperesAsm + 1
    Next i
End Function

Private Sub PalLireReperes(ByVal racine As pfcls.IpfcAssembly, _
    ByVal modele As pfcls.IpfcModel, _
    ByVal ids As pfcls.Iintseq, _
    ByVal rouleaux As Collection, _
    ByVal proprietaire As String)

    Dim possesseur As pfcls.IpfcModelItemOwner
    Dim items As pfcls.IpfcModelItems
    Dim item As pfcls.IpfcModelItem
    Dim csys As pfcls.IpfcCoordSystem
    Dim fabrique As pfcls.CMpfcAssembly
    Dim chemin As pfcls.IpfcComponentPath
    Dim transfo As pfcls.IpfcTransform3D
    Dim point As pfcls.IpfcPoint3D
    Dim nom As String
    Dim i As Long

    Set possesseur = modele
    Set items = possesseur.ListItems(pfcls.EpfcITEM_COORD_SYS)
    If items Is Nothing Then Exit Sub
    For i = 0 To items.Count - 1
        Set item = items.item(i)
        nom = ""
        On Error Resume Next
        nom = UCase$(Trim$(item.GetName()))
        Err.Clear
        On Error GoTo 0
        If PalEstRepereRoll(nom) Then
            If transfo Is Nothing Then
                Set fabrique = New pfcls.CMpfcAssembly
                Set chemin = fabrique.CreateComponentPath(racine, ids)
                ' Vrai = transformation composant -> assemblage racine.
                Set transfo = chemin.GetTransform(True)
            End If
            Set csys = item
            Set point = transfo.TransformPoint(csys.CoordSys.GetOrigin())
            PalAjouterRouleau rouleaux, modele.Filename, nom, _
                CDbl(point.item(0)), CDbl(point.item(1)), CDbl(point.item(2)), ids, _
                proprietaire
        End If
    Next i
End Sub

Private Sub PalAjouterRouleau(ByVal rouleaux As Collection, _
    ByVal fichier As String, ByVal repere As String, _
    ByVal x As Double, ByVal y As Double, ByVal z As Double, _
    ByVal ids As pfcls.Iintseq, ByVal proprietaire As String)

    Dim i As Long
    Dim chemin As String

    ' Un meme repere vu deux fois au meme endroit n'est compte qu'une fois.
    For i = 1 To rouleaux.Count
        If StrComp(PalChampTexte(rouleaux, i, 1), repere, vbTextCompare) = 0 Then
            If Abs(PalChamp(rouleaux, i, 2) - x) <= PAL_DOUBLON And _
                Abs(PalChamp(rouleaux, i, 3) - y) <= PAL_DOUBLON And _
                Abs(PalChamp(rouleaux, i, 4) - z) <= PAL_DOUBLON Then Exit Sub
        End If
    Next i
    For i = 0 To ids.Count - 1
        If chemin <> "" Then chemin = chemin & "/"
        chemin = chemin & CStr(ids.item(i))
    Next i
    ' 6 = assemblage proprietaire : celui dont le squelette donne le pied.
    rouleaux.Add Array(fichier, repere, x, y, z, chemin, proprietaire)
End Sub

Private Function PalCopierIds(ByVal ids As pfcls.Iintseq, _
    ByVal id As Long) As pfcls.Iintseq

    Dim copie As pfcls.Iintseq
    Dim i As Long

    Set copie = New pfcls.Cintseq
    For i = 0 To ids.Count - 1
        copie.Append ids.item(i)
    Next i
    copie.Append id
    Set PalCopierIds = copie
End Function

Private Function PalEstRepereRoll(ByVal nom As String) As Boolean
    Dim milieu As String
    Dim i As Long
    Dim c As String

    nom = UCase$(Trim$(nom))
    If nom = "ROLL_CENTER" Then
        PalEstRepereRoll = True
        Exit Function
    End If
    If Left$(nom, 4) <> "ROLL" Or Right$(nom, 7) <> "_CENTER" Then Exit Function
    If Len(nom) <= 11 Then Exit Function
    milieu = Mid$(nom, 5, Len(nom) - 11)
    For i = 1 To Len(milieu)
        c = Mid$(milieu, i, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next i
    PalEstRepereRoll = True
End Function

Private Function PalEstSquelette(ByVal modele As pfcls.IpfcModel) As Boolean
    Dim objet As Object
    On Error Resume Next
    Set objet = modele
    PalEstSquelette = CBool(objet.IsSkeleton)
    Err.Clear
    On Error GoTo 0
End Function

' Memes regles que le module pieds moteurs (source unique).
Private Function PalEstMoteur(ByVal fichier As String) As Boolean
    PalEstMoteur = TECH_PUB_EstModeleMoteur(fichier, PalFichierMoteur())
End Function

Private Function PalFichierMoteur() As String
    Dim chemin As String
    Dim p As Long
    On Error Resume Next
    chemin = Trim$(PalTexte(ThisWorkbook.Worksheets(PAL_FEUILLE_PILOTE). _
        Range(PAL_CELLULE_MOTEUR).value))
    Err.Clear
    On Error GoTo 0
    p = InStrRev(Replace(chemin, "/", "\"), "\")
    If p > 0 Then chemin = Mid$(chemin, p + 1)
    PalFichierMoteur = chemin
End Function

Private Function PalEstHeating(ByVal fichier As String) As Boolean
    Dim nom As String
    Dim inscrit As String

    nom = UCase$(fichier)
    If InStr(1, nom, "HEATING") > 0 And InStr(1, nom, "PREHEAT") = 0 Then
        PalEstHeating = True
        Exit Function
    End If
    On Error Resume Next
    inscrit = PalTexte(ThisWorkbook.Worksheets("ZONES_FOUR").Range("X8").value)
    Err.Clear
    On Error GoTo 0
    If inscrit <> "" Then PalEstHeating = (StrComp(inscrit, fichier, vbTextCompare) = 0)
End Function

Private Function PalBase(ByVal fichier As String) As String
    Dim p As Long
    fichier = Trim$(fichier)
    p = InStr(1, fichier, ".")
    If p > 1 Then PalBase = Left$(fichier, p - 1) Else PalBase = fichier
End Function

' ---------------------------------------------------------------- Calculs

Private Function PalChamp(ByVal rouleaux As Collection, ByVal i As Long, _
    ByVal k As Long) As Double
    Dim v As Variant
    v = rouleaux.item(i)
    PalChamp = CDbl(v(k))
End Function

Private Function PalChampTexte(ByVal rouleaux As Collection, ByVal i As Long, _
    ByVal k As Long) As String
    Dim v As Variant
    v = rouleaux.item(i)
    PalChampTexte = CStr(v(k))
End Function

Private Sub PalTrier(ByRef ys() As Double, ByRef idx() As Long, ByVal n As Long)
    Dim i As Long
    Dim j As Long
    Dim y As Double
    Dim k As Long

    For i = 2 To n
        y = ys(i): k = idx(i)
        j = i - 1
        Do While j >= 1
            If ys(j) <= y Then Exit Do
            ys(j + 1) = ys(j): idx(j + 1) = idx(j)
            j = j - 1
        Loop
        ys(j + 1) = y: idx(j + 1) = k
    Next i
End Sub

Private Sub PalGrouper(ByRef ys() As Double, ByRef idx() As Long, _
    ByVal n As Long, ByVal ecartMaxi As Double, ByRef debutGroupe() As Long, _
    ByRef finGroupe() As Long, ByRef nbGroupes As Long, _
    ByRef yReference() As Double, ByRef abaisse() As Boolean)

    ' Regle de regroupement (Y tries par ordre croissant en entree) :
    '   1. groupes de rouleaux proches en hauteur : tous les rouleaux a
    '      ecartMaxi ou moins du plus bas du groupe ;
    '   2. un rouleau seul rejoint le groupe de plusieurs rouleaux le plus
    '      proche en hauteur (s'il n'existe aucun groupe de plusieurs
    '      rouleaux, chaque rouleau seul garde son propre groupe) ;
    '   3. groupes renumerotes du bas vers le haut ; ys et idx sont
    '      reordonnes pour que chaque groupe soit contigu ;
    '   4. Y de reference du palier = rouleau le plus bas du groupe FINAL.
    '      Un rouleau seul situe SOUS le groupe qu'il rejoint abaisse donc
    '      le palier (abaisse = Vrai) : sinon le palier serait au-dessus de
    '      lui (un rouleau seul est par construction a plus de ecartMaxi).
    Dim brut() As Long
    Dim nbBrut As Long
    Dim effectif() As Long
    Dim yMin() As Double
    Dim yMax() As Double
    Dim groupeFinal() As Long
    Dim rang() As Long
    Dim ysTrie() As Double
    Dim idxTrie() As Long
    Dim nbMulti As Long
    Dim meilleur As Long
    Dim meilleureDistance As Double
    Dim d As Double
    Dim i As Long
    Dim c As Long
    Dim k As Long
    Dim pos As Long

    nbGroupes = 0
    If n < 1 Then Exit Sub
    ReDim brut(1 To n)

    ' 1. Groupes bruts.
    nbBrut = 1
    brut(1) = 1
    ReDim yMin(1 To 1)
    yMin(1) = ys(1)
    For i = 2 To n
        If ys(i) - yMin(nbBrut) > ecartMaxi + 0.000001 Then
            nbBrut = nbBrut + 1
            ReDim Preserve yMin(1 To nbBrut)
            yMin(nbBrut) = ys(i)
        End If
        brut(i) = nbBrut
    Next i
    ReDim effectif(1 To nbBrut)
    ReDim yMax(1 To nbBrut)
    For i = 1 To n
        effectif(brut(i)) = effectif(brut(i)) + 1
        yMax(brut(i)) = ys(i)
    Next i
    For c = 1 To nbBrut
        If effectif(c) >= 2 Then nbMulti = nbMulti + 1
    Next c

    ' 2. Rouleaux seuls rattaches au groupe le plus proche.
    ReDim groupeFinal(1 To n)
    For i = 1 To n
        groupeFinal(i) = brut(i)
        If effectif(brut(i)) = 1 And nbMulti > 0 Then
            meilleur = 0
            meilleureDistance = 1E+30
            For c = 1 To nbBrut
                If effectif(c) >= 2 Then
                    If ys(i) < yMin(c) Then
                        d = yMin(c) - ys(i)
                    Else
                        d = ys(i) - yMax(c)
                    End If
                    If d < meilleureDistance Then
                        meilleureDistance = d
                        meilleur = c
                    End If
                End If
            Next c
            groupeFinal(i) = meilleur
        End If
    Next i

    ' 3. Numerotation du bas vers le haut (ordre du premier rouleau rencontre,
    '    ys etant croissant) puis tri stable par groupe.
    ReDim rang(1 To nbBrut)
    For i = 1 To n
        If rang(groupeFinal(i)) = 0 Then
            nbGroupes = nbGroupes + 1
            rang(groupeFinal(i)) = nbGroupes
        End If
    Next i
    ReDim ysTrie(1 To n)
    ReDim idxTrie(1 To n)
    ReDim debutGroupe(1 To nbGroupes)
    ReDim finGroupe(1 To nbGroupes)
    ReDim yReference(1 To nbGroupes)
    ReDim abaisse(1 To nbGroupes)
    pos = 0
    For k = 1 To nbGroupes
        debutGroupe(k) = pos + 1
        For i = 1 To n
            If rang(groupeFinal(i)) = k Then
                pos = pos + 1
                ysTrie(pos) = ys(i)
                idxTrie(pos) = idx(i)
                ' 4. Premier rouleau du groupe = le plus bas.
                If pos = debutGroupe(k) Then
                    yReference(k) = ys(i)
                    abaisse(k) = (effectif(brut(i)) = 1 And nbMulti > 0)
                End If
            End If
        Next i
        finGroupe(k) = pos
    Next k
    For i = 1 To n
        ys(i) = ysTrie(i)
        idx(i) = idxTrie(i)
    Next i
End Sub

Private Sub PalControler(ByRef yPalier() As Double, ByRef yHaut() As Double, _
    ByRef actif() As Boolean, ByVal nb As Long, ByVal topY As Double, _
    ByVal topOK As Boolean, ByVal minGap As Double)

    Dim g As Long
    Dim h As Long

    ' Les paliers trop bas sont deja ecartes (actif = Faux).
    For g = 1 To nb
        If actif(g) And topOK Then
            If yPalier(g) > topY + 0.000001 Then Err.Raise PAL_ERR + 31, , _
                PalNomRangee(g, nb) & " : palier Y " & PalFmt(yPalier(g)) & _
                " au-dessus du sommet des poteaux (Structure_HEATING!G14 = " & _
                PalFmt(topY) & ")."
        End If
    Next g
    For g = 1 To nb - 1
        For h = g + 1 To nb
            If actif(g) And actif(h) Then
                If Abs(yPalier(g) - yPalier(h)) < minGap - 0.000001 Then _
                    Err.Raise PAL_ERR + 32, , _
                        "Les paliers des groupes '" & PalNomRangee(g, nb) & _
                        "' (Y " & PalFmt(yPalier(g)) & ") et '" & _
                        PalNomRangee(h, nb) & "' (Y " & PalFmt(yPalier(h)) & _
                        ") sont a moins de l'ecart mini (" & PalFmt(minGap) & _
                        " mm, Structure_HEATING!G15)."
                ' Groupes tries du bas vers le haut : le palier du groupe h
                ' ne doit pas passer sous un rouleau du groupe g.
                If yPalier(h) < yHaut(g) - 0.000001 Then Err.Raise PAL_ERR + 35, , _
                    "Le palier de '" & PalNomRangee(h, nb) & "' (Y " & _
                    PalFmt(yPalier(h)) & ") passe sous le rouleau le plus haut de '" & _
                    PalNomRangee(g, nb) & "' (Y " & PalFmt(yHaut(g)) & ") : ces " & _
                    "rouleaux se caleraient sur lui et non sur leur palier. " & _
                    "Aucune modification faite."
            End If
        Next h
    Next g
End Sub

' Lignes de NiveauxY_HEATING a desactiver, calculees AVANT toute ecriture
' (etat vu apres reactivation des [OFF PALIER] et retrait des PAL_xx).
' Un niveau gene un palier s'il est a moins de l'ecart mini du palier, ou
' entre le palier et le rouleau le plus haut du groupe : le pied du
' rouleau se calerait sur lui au lieu du palier.
Private Function PalPlanifierNiveaux(ByVal a As Worksheet, ByRef yPalier() As Double, _
    ByRef yHaut() As Double, ByRef actif() As Boolean, ByVal nb As Long, _
    ByVal minGap As Double, ByRef offGroupe() As String, _
    ByRef nbDesactives As Long, ByRef listeDesactives As String) As Collection

    Dim lignes As Collection
    Dim r As Long
    Dim g As Long
    Dim id As String
    Dim note As String
    Dim y As Double
    Dim haut As Double
    Dim actifLigne As Boolean
    Dim libelle As String

    Set lignes = New Collection
    ReDim offGroupe(1 To nb)
    nbDesactives = 0
    listeDesactives = ""
    For r = PAL_NIV_FIRST To PAL_NIV_LAST
        If Application.WorksheetFunction.CountA( _
            a.Range(a.Cells(r, 1), a.Cells(r, 5))) > 0 Then
            id = Trim$(PalTexte(a.Cells(r, 1).value))
            note = PalTexte(a.Cells(r, 5).value)
            If UCase$(Left$(id, Len(PAL_PREFIXE_ID))) <> PAL_PREFIXE_ID And _
                PalNumerique(a.Cells(r, 3).value) And _
                PalNumerique(a.Cells(r, 4).value) Then
                actifLigne = (CDbl(a.Cells(r, 4).value) = 1) Or _
                    (Left$(note, Len(PAL_MARQUE)) = PAL_MARQUE)
                If actifLigne Then
                    y = CDbl(a.Cells(r, 3).value)
                    For g = 1 To nb
                        If actif(g) Then
                            haut = yHaut(g)
                            If haut < yPalier(g) + minGap Then haut = yPalier(g) + minGap
                            If y > yPalier(g) - minGap + 0.000001 And _
                                y < haut - 0.000001 Then
                                lignes.Add r
                                nbDesactives = nbDesactives + 1
                                libelle = IIf(id = "", "ligne " & r, id) & _
                                    " (Y " & PalFmt(y) & ")"
                                If offGroupe(g) <> "" Then offGroupe(g) = offGroupe(g) & ", "
                                offGroupe(g) = offGroupe(g) & libelle
                                If listeDesactives <> "" Then listeDesactives = listeDesactives & ", "
                                listeDesactives = listeDesactives & libelle
                                Exit For
                            End If
                        End If
                    Next g
                End If
            End If
        End If
    Next r
    Set PalPlanifierNiveaux = lignes
End Function

' Decalage d = Y Creo (repere de l'assemblage) - Y Excel, mesure sur les
' niveaux de Niveaux_HEATING retrouves dans le squelette. Les paliers
' (PAL_xx de NiveauxY_HEATING) sont exclus : ce sont eux qu'on corrige.
' Le decalage retenu doit retrouver TOUS les niveaux de reference, et
' etre unique ; sinon le squelette n'est pas a jour et rien n'est mesure.
Private Function PalMesurerDecalage(ByVal lignes As Collection, ByRef d As Double, _
    ByRef nbRef As Long, ByRef probleme As String) As Boolean

    Dim calc As Worksheet
    Dim saisie As Worksheet
    Dim refs() As Double
    Dim nRefs As Long
    Dim yl() As Double
    Dim nL As Long
    Dim r As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim y As Double
    Dim a As Variant
    Dim existe As Boolean
    Dim cand As Double
    Dim nbOK As Long
    Dim somme As Double
    Dim meilleur As Double
    Dim ecartK As Double
    Dim nbSol As Long
    Dim dSol As Double

    probleme = ""
    nbRef = 0
    Set calc = ThisWorkbook.Worksheets("Niveaux_HEATING")
    Set saisie = ThisWorkbook.Worksheets("NiveauxY_HEATING")
    ReDim refs(1 To PAL_CALC_LAST - PAL_CALC_FIRST + 1)
    For r = PAL_CALC_FIRST To PAL_CALC_LAST
        If PalNumerique(calc.Cells(r, 4).value) Then
            y = CDbl(calc.Cells(r, 4).value)
            If Not PalEstPalierSaisi(saisie, y) Then
                nRefs = nRefs + 1
                refs(nRefs) = y
            End If
        End If
    Next r
    If nRefs = 0 Then
        probleme = "aucun niveau calcule (hors paliers) dans Niveaux_HEATING."
        Exit Function
    End If

    ' Y distincts des lignes horizontales du squelette.
    If lignes Is Nothing Then
        probleme = "squelette sans ligne lisible."
        Exit Function
    End If
    ReDim yl(1 To lignes.Count + 1)
    For i = 1 To lignes.Count
        a = lignes.item(i)
        If a(0) = "X" Or a(0) = "Z" Then
            existe = False
            For k = 1 To nL
                If Abs(yl(k) - a(2)) <= 0.05 Then
                    existe = True
                    Exit For
                End If
            Next k
            If Not existe Then
                nL = nL + 1
                yl(nL) = a(2)
            End If
        End If
    Next i
    If nL = 0 Then
        probleme = "aucune ligne horizontale dans le squelette."
        Exit Function
    End If

    ' Le vrai decalage amene le premier niveau sur une des lignes : chaque
    ' ligne donne un candidat, garde s'il retrouve tous les niveaux.
    For i = 1 To nL
        cand = yl(i) - refs(1)
        nbOK = 0
        somme = 0#
        For j = 1 To nRefs
            meilleur = 1E+30
            For k = 1 To nL
                ecartK = yl(k) - (refs(j) + cand)
                If Abs(ecartK) < Abs(meilleur) Then meilleur = ecartK
            Next k
            If Abs(meilleur) <= PAL_TOL_LIGNE Then
                nbOK = nbOK + 1
                somme = somme + cand + meilleur
            End If
        Next j
        If nbOK = nRefs Then
            If nbSol = 0 Then
                nbSol = 1
                dSol = somme / nRefs
            ElseIf Abs(somme / nRefs - dSol) > PAL_TOL_LIGNE Then
                nbSol = nbSol + 1
            End If
        End If
    Next i
    If nbSol = 0 Then
        probleme = "les " & nRefs & " niveau(x) de Niveaux_HEATING ne se retrouvent " & _
            "pas tous dans le squelette avec un meme decalage : squelette pas a " & _
            "jour (IGES non recharge) ou axe Y inverse."
        Exit Function
    End If
    If nbSol > 1 Then
        probleme = "decalage ambigu : " & nRefs & " niveau(x) de reference seulement, " & _
            "plusieurs decalages possibles."
        Exit Function
    End If
    d = Round(dSol, 3)
    nbRef = nRefs
    PalMesurerDecalage = True
End Function

Private Function PalEstPalierSaisi(ByVal saisie As Worksheet, ByVal y As Double) As Boolean
    Dim r As Long
    For r = PAL_NIV_FIRST To PAL_NIV_LAST
        If UCase$(Left$(Trim$(PalTexte(saisie.Cells(r, 1).value)), Len(PAL_PREFIXE_ID))) = _
            PAL_PREFIXE_ID Then
            If PalNumerique(saisie.Cells(r, 3).value) Then
                If Abs(CDbl(saisie.Cells(r, 3).value) - y) <= PAL_TOL_LIGNE Then
                    PalEstPalierSaisi = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

' Apres HEAT_Recalculer : chaque palier doit figurer dans Niveaux_HEATING
' (colonne D) et aucun niveau calcule ne doit rester entre lui et ses
' rouleaux. Couvre un HEATING en mode automatique qui ignorerait
' NiveauxY_HEATING, ou un niveau ajoute par le calcul.
Private Function PalVerifierCalcul(ByRef yPalier() As Double, _
    ByRef yHaut() As Double, ByRef actif() As Boolean, ByVal nb As Long, _
    ByRef probleme As String) As Boolean

    Dim calc As Worksheet
    Dim r As Long
    Dim g As Long
    Dim y As Double
    Dim trouve As Boolean

    probleme = ""
    Set calc = ThisWorkbook.Worksheets("Niveaux_HEATING")
    For g = 1 To nb
        If actif(g) Then
            trouve = False
            For r = PAL_CALC_FIRST To PAL_CALC_LAST
                If PalNumerique(calc.Cells(r, 4).value) Then
                    y = CDbl(calc.Cells(r, 4).value)
                    If Abs(y - yPalier(g)) <= 0.01 Then
                        trouve = True
                    ElseIf y > yPalier(g) + 0.01 And y < yHaut(g) - 0.01 Then
                        probleme = "niveau calcule Y " & PalFmt(y) & _
                            " (Niveaux_HEATING ligne " & r & ") entre le palier Y " & _
                            PalFmt(yPalier(g)) & " et les rouleaux de '" & _
                            PalNomRangee(g, nb) & "' : le pied se calerait dessus."
                        Exit Function
                    End If
                End If
            Next r
            If Not trouve Then
                probleme = "le palier Y " & PalFmt(yPalier(g)) & " de '" & _
                    PalNomRangee(g, nb) & "' est absent de Niveaux_HEATING apres " & _
                    "recalcul. HEATING est-il en mode automatique (NiveauxY_HEATING " & _
                    "alors ignore) ?"
                Exit Function
            End If
        End If
    Next g
    PalVerifierCalcul = True
End Function

' ---------------------------------------------------------------- Excel

Private Sub PalLireStructure(ByVal s As Worksheet, ByRef bottomY As Double, _
    ByRef topY As Double, ByRef topOK As Boolean, ByRef minGap As Double)

    If Not PalNumerique(s.Range("G13").value) Then _
        Err.Raise PAL_ERR + 40, , "Structure_HEATING!G13 (pied des poteaux) doit etre numerique."
    If Not PalNumerique(s.Range("G15").value) Then _
        Err.Raise PAL_ERR + 41, , "Structure_HEATING!G15 (ecart mini) doit etre numerique."
    bottomY = CDbl(s.Range("G13").value)
    minGap = CDbl(s.Range("G15").value)
    If minGap < 0 Then Err.Raise PAL_ERR + 42, , _
        "Structure_HEATING!G15 (ecart mini) doit etre positif ou nul."
    topOK = False
    If PalNumerique(s.Range("G14").value) Then
        topY = CDbl(s.Range("G14").value)
        topOK = True
    End If
End Sub

Private Function PalFeuille() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PAL_FEUILLE)
    Err.Clear
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = PAL_FEUILLE
        ws.Tab.Color = RGB(223, 113, 36)
    End If
    ws.Range("A1").value = "HEATING - paliers sous les ROLL_CENTER (version " & _
        PAL_VERSION & ")"
    ws.Range("A3").value = "Ecart maxi en hauteur dans un groupe (mm, fixe)"
    ws.Range("C3").value = PAL_TOLERANCE_DEFAUT
    ws.Range("C3").Interior.ColorIndex = xlNone
    ws.Range("E3:G3").ClearContents
    ws.Range("A4").value = "Y palier Excel = Y du centre le plus bas du groupe - " & _
        PalFmt(PAL_DISTANCE_PALIER) & " mm - d (d = decalage Excel -> Creo mesure)." & _
        " Y reference en repere Creo, Y palier en Y Excel."
    ws.Range("A5").value = "Decalage d Excel -> Creo au placement (mm)"
    ws.Range("A6:K6").value = Array("Groupe", "Nb rouleaux", "Y min", "Y max", _
        "Y reference (rouleau le plus bas du groupe)", "Distance vers le bas (mm)", _
        "Y palier", "ID niveau", "Etat", "Reperes (fichier : repere : Y)", _
        "Niveaux desactives par ce palier")
    ws.Range("A6:K6").Font.Bold = True
    Set PalFeuille = ws
End Function

Private Sub PalEcrireProposition(ByVal ws As Worksheet, ByVal rouleaux As Collection, _
    ByRef ys() As Double, ByRef idx() As Long, ByRef debutGroupe() As Long, _
    ByRef finGroupe() As Long, ByVal nbGroupes As Long, _
    ByVal nomAssemblage As String, ByRef yRef() As Double, _
    ByRef yPalier() As Double, ByRef distance() As Double, _
    ByRef actif() As Boolean, ByRef offGroupe() As String)

    Dim g As Long
    Dim i As Long
    Dim r As Long
    Dim derniere As Long
    Dim liste As String

    ws.Range("A2").value = "Assemblage : " & nomAssemblage & " / releve du " & _
        Format$(Now, "yyyy-mm-dd hh:nn")
    ws.Range("C5").ClearContents
    derniere = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If derniere >= PAL_LIGNE_DONNEES Then _
        ws.Range(ws.Cells(PAL_LIGNE_DONNEES, 1), ws.Cells(derniere, 11)).ClearContents
    ws.Range(ws.Cells(PAL_LIGNE_DONNEES, 9), ws.Cells(PAL_LIGNE_DONNEES + 500, 9)).Interior.ColorIndex = xlNone
    For g = 1 To nbGroupes
        r = PAL_LIGNE_DONNEES + g - 1
        liste = ""
        For i = debutGroupe(g) To finGroupe(g)
            If liste <> "" Then liste = liste & " ; "
            liste = liste & PalChampTexte(rouleaux, idx(i), 0) & " : " & _
                PalChampTexte(rouleaux, idx(i), 1) & " : " & PalFmt(ys(i))
        Next i
        If Len(liste) > 32000 Then liste = Left$(liste, 32000) & " ..."
        ws.Cells(r, 1).value = PalNomRangee(g, nbGroupes)
        ws.Cells(r, 2).value = finGroupe(g) - debutGroupe(g) + 1
        ws.Cells(r, 3).value = ys(debutGroupe(g))
        ws.Cells(r, 4).value = ys(finGroupe(g))
        ws.Cells(r, 5).value = yRef(g)
        ws.Cells(r, 6).value = distance(g)
        ws.Cells(r, 7).value = yPalier(g)
        ws.Cells(r, 9).value = IIf(actif(g), "PROPOSE", _
            "PROPOSE : pas de palier (sous le pied des poteaux, G13)")
        ws.Cells(r, 10).value = liste
        ws.Cells(r, 11).value = offGroupe(g)
    Next g
    ws.Columns("A:I").AutoFit
    ws.Columns("J").ColumnWidth = 100
    ws.Columns("K").ColumnWidth = 60
End Sub

Private Sub PalEcrireResultats(ByVal ws As Worksheet, _
    ByRef distance() As Double, ByRef yPalier() As Double, _
    ByRef actif() As Boolean, ByVal nb As Long, ByVal decalageY As Double)

    Dim g As Long
    Dim r As Long

    ws.Range("C5").value = decalageY
    ws.Range("I3").ClearContents
    ws.Range("I3").Interior.ColorIndex = xlNone
    For g = 1 To nb
        r = PAL_LIGNE_DONNEES + g - 1
        ws.Cells(r, 6).value = distance(g)
        ws.Cells(r, 7).value = yPalier(g)
        If actif(g) Then
            ws.Cells(r, 8).value = PAL_PREFIXE_ID & Format$(g, "00")
            ws.Cells(r, 9).value = "OK " & Format$(Now, "yyyy-mm-dd hh:nn")
            ws.Cells(r, 9).Interior.Color = RGB(198, 239, 206)
        Else
            ws.Cells(r, 8).value = ""
            ws.Cells(r, 9).value = "PAS DE PALIER : sous le pied des poteaux (G13)"
            ws.Cells(r, 9).Interior.Color = RGB(255, 235, 156)
        End If
    Next g
End Sub

' Remet PALIERS_HEATING tel qu'avant le lancement et note pourquoi en I3.
Private Sub PalRestaurerFeuille(ByVal ws As Worksheet, ByRef avant As Variant, _
    ByVal texte As String)
    ws.Range("A2:K" & (PAL_LIGNE_DONNEES + 500)).Formula = avant
    ws.Range(ws.Cells(PAL_LIGNE_DONNEES, 9), _
        ws.Cells(PAL_LIGNE_DONNEES + 500, 9)).Interior.ColorIndex = xlNone
    ws.Range("I3").value = "Dernier lancement " & Format$(Now, "yyyy-mm-dd hh:nn") & _
        " : " & texte & ". Paliers ci-dessous et NiveauxY_HEATING inchanges."
    ws.Range("I3").Interior.Color = RGB(255, 199, 206)
End Sub

Private Sub PalEcrireNiveaux(ByVal a As Worksheet, ByRef yPalier() As Double, _
    ByRef distance() As Double, ByRef actif() As Boolean, ByVal nb As Long, _
    ByVal lignesOff As Collection)

    Dim r As Long
    Dim g As Long
    Dim note As String
    Dim id As String
    Dim libre As Long
    Dim v As Variant

    ' 1. Reactiver les niveaux desactives par un lancement precedent.
    For r = PAL_NIV_FIRST To PAL_NIV_LAST
        note = PalTexte(a.Cells(r, 5).value)
        If Left$(note, Len(PAL_MARQUE)) = PAL_MARQUE Then
            a.Cells(r, 4).value = 1
            a.Cells(r, 5).value = Mid$(note, Len(PAL_MARQUE) + 1)
        End If
    Next r

    ' 2. Retirer les paliers d'un lancement precedent.
    For r = PAL_NIV_FIRST To PAL_NIV_LAST
        id = UCase$(Trim$(PalTexte(a.Cells(r, 1).value)))
        If Left$(id, Len(PAL_PREFIXE_ID)) = PAL_PREFIXE_ID Then
            a.Range(a.Cells(r, 1), a.Cells(r, 5)).ClearContents
        End If
    Next r

    ' 3. Desactiver les niveaux genants (liste montree avant confirmation).
    For Each v In lignesOff
        r = CLng(v)
        a.Cells(r, 4).value = 0
        a.Cells(r, 5).value = PAL_MARQUE & PalTexte(a.Cells(r, 5).value)
    Next v

    ' 4. Ecrire un palier par groupe dans les lignes libres.
    libre = PAL_NIV_FIRST
    For g = 1 To nb
        If Not actif(g) Then GoTo GroupeSuivant
        Do While libre <= PAL_NIV_LAST
            If Application.WorksheetFunction.CountA( _
                a.Range(a.Cells(libre, 1), a.Cells(libre, 5))) = 0 Then Exit Do
            libre = libre + 1
        Loop
        If libre > PAL_NIV_LAST Then Err.Raise PAL_ERR + 50, , _
            "NiveauxY_HEATING est plein (lignes " & PAL_NIV_FIRST & " a " & _
            PAL_NIV_LAST & ")."
        a.Cells(libre, 1).value = PAL_PREFIXE_ID & Format$(g, "00")
        a.Cells(libre, 2).value = "PRINCIPAL"
        a.Cells(libre, 3).value = yPalier(g)
        a.Cells(libre, 4).value = 1
        a.Cells(libre, 5).value = "Palier " & LCase$(PalNomRangee(g, nb)) & _
            " : " & PalFmt(distance(g)) & " mm sous le rouleau le plus bas"
        libre = libre + 1
GroupeSuivant:
    Next g
End Sub

Private Function PalTexteConfirmation(ByVal rouleaux As Collection, _
    ByRef idx() As Long, ByRef debutGroupe() As Long, _
    ByRef finGroupe() As Long, ByVal nbGroupes As Long, ByRef yRef() As Double, _
    ByRef yPalier() As Double, ByRef actif() As Boolean, _
    ByRef abaisse() As Boolean, ByRef offGroupe() As String, _
    ByVal ignores As Long, ByVal reperesAsm As Long, _
    ByVal nomAssemblage As String, ByVal aSquelette As Boolean, _
    ByVal bottomY As Double, ByVal decalageY As Double, _
    ByVal origineDecalage As String) As String

    Dim texte As String
    Dim g As Long

    texte = "Decalage Excel -> Creo d = " & PalFmt(decalageY) & " mm (" & _
        origineDecalage & ")." & vbCrLf
    If Abs(decalageY) > PAL_TOL_SQUELETTE Then texte = texte & _
        "! Compense : palier a " & PalFmt(PAL_DISTANCE_PALIER) & " mm sous les " & _
        "rouleaux DANS CREO. Ne plus deplacer ni recreer le squelette." & vbCrLf
    texte = texte & vbCrLf & "Rouleaux trouves dans " & nomAssemblage & " :" & vbCrLf
    For g = nbGroupes To 1 Step -1
        texte = texte & vbCrLf & UCase$(PalNomRangee(g, nbGroupes)) & _
            " (" & (finGroupe(g) - debutGroupe(g) + 1) & ") : plus bas Y " & _
            PalFmt(yRef(g))
        If actif(g) Then
            texte = texte & " -> palier Y Excel " & PalFmt(yPalier(g)) & _
                " (Creo " & PalFmt(yPalier(g) + decalageY) & ")" & vbCrLf
        Else
            texte = texte & " -> PAS DE PALIER (Y " & PalFmt(yPalier(g)) & _
                " sous le pied " & PalFmt(bottomY) & ")" & vbCrLf
        End If
        If abaisse(g) And actif(g) Then texte = texte & _
            "   ! palier abaisse par un rouleau seul situe sous le groupe" & vbCrLf
        texte = texte & PalListePieces(rouleaux, idx, debutGroupe(g), finGroupe(g), 4)
        If offGroupe(g) <> "" Then texte = texte & "   niveaux desactives : " & _
            PalCouper(offGroupe(g), 120) & vbCrLf
    Next g
    If ignores > 0 Then texte = texte & vbCrLf & ignores & _
        " composant(s) non charge(s) dans Creo, non verifie(s)."
    If reperesAsm > 0 Then texte = texte & vbCrLf & reperesAsm & _
        " repere(s) ROLL_CENTER pose(s) dans un assemblage : ignore(s), " & _
        "seuls ceux des pieces comptent."
    If Not aSquelette Then texte = texte & vbCrLf & _
        "Pas encore de squelette : il devra etre cree a l'origine, puis " & _
        "relancer cette macro pour mesurer d."
    PalTexteConfirmation = PalCouper(texte, PAL_MSG_MAX) & vbCrLf & vbCrLf & _
        "Palier a " & PalFmt(PAL_DISTANCE_PALIER) & " mm sous le rouleau le " & _
        "plus bas de chaque groupe. Detail : feuille " & PAL_FEUILLE & "." & _
        vbCrLf & "OUI = placer les paliers" & vbCrLf & _
        "NON = arreter sans rien modifier"
End Function

Private Function PalListePieces(ByVal rouleaux As Collection, _
    ByRef idx() As Long, ByVal premier As Long, ByVal dernier As Long, _
    ByVal maxi As Long) As String

    Dim noms() As String
    Dim nombres() As Long
    Dim nb As Long
    Dim i As Long
    Dim k As Long
    Dim nom As String
    Dim trouve As Boolean
    Dim texte As String

    ' Une ligne par piece et repere ; les occurrences identiques sont comptees.
    ReDim noms(1 To dernier - premier + 1)
    ReDim nombres(1 To dernier - premier + 1)
    For i = premier To dernier
        nom = LCase$(PalBase(PalChampTexte(rouleaux, idx(i), 0))) & _
            " (" & PalChampTexte(rouleaux, idx(i), 1) & ")"
        trouve = False
        For k = 1 To nb
            If noms(k) = nom Then
                nombres(k) = nombres(k) + 1
                trouve = True
                Exit For
            End If
        Next k
        If Not trouve Then
            nb = nb + 1
            noms(nb) = nom
            nombres(nb) = 1
        End If
    Next i
    For k = 1 To nb
        If k > maxi Then
            texte = texte & "   ... et " & (nb - maxi) & " autre(s)" & vbCrLf
            Exit For
        End If
        texte = texte & "   - " & noms(k) & IIf(nombres(k) > 1, _
            "  x" & nombres(k), "") & vbCrLf
    Next k
    PalListePieces = texte
End Function

Private Function PalNomRangee(ByVal g As Long, ByVal nb As Long) As String
    ' Groupes tries du bas vers le haut : 1 = le plus bas.
    If nb = 1 Then
        PalNomRangee = "Rouleaux"
    ElseIf g = 1 Then
        PalNomRangee = "Rouleaux du dessous"
    ElseIf g = nb Then
        PalNomRangee = "Rouleaux du dessus"
    ElseIf nb = 3 Then
        PalNomRangee = "Rouleaux du milieu"
    Else
        PalNomRangee = "Rouleaux intermediaires " & CStr(g - 1)
    End If
End Function

Private Function PalFmt(ByVal valeur As Double) As String
    PalFmt = Format$(valeur, "0.###")
    If Right$(PalFmt, 1) = "," Or Right$(PalFmt, 1) = "." Then _
        PalFmt = Left$(PalFmt, Len(PalFmt) - 1)
End Function

Private Function PalCouper(ByVal texte As String, ByVal maxi As Long) As String
    If Len(texte) <= maxi Then
        PalCouper = texte
    Else
        PalCouper = Left$(texte, maxi) & " ..."
    End If
End Function

' Texte d'une cellule, vide si la cellule contient une erreur Excel.
Private Function PalTexte(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    PalTexte = CStr(v)
End Function

Private Function PalNumerique(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    PalNumerique = IsNumeric(v)
End Function

' ---------------------------------------------------------------- Squelette

Private Function PalPoseSquelette(ByVal session As pfcls.IpfcBaseSession, _
    ByVal asm As pfcls.IpfcAssembly, ByVal solide As pfcls.IpfcSolid, _
    ByRef squelette As pfcls.IpfcSolid, ByRef nomSquelette As String, _
    ByRef pose As pfcls.IpfcTransform3D) As Boolean

    ' Retourne Vrai si l'assemblage a un squelette ; pose = transformation
    ' EFFECTIVE du squelette dans l'assemblage (repere squelette -> asm).
    Dim objetAsm As Object
    Dim modeleSq As pfcls.IpfcModel
    Dim features As pfcls.IpfcFeatures
    Dim feature As pfcls.IpfcFeature
    Dim item As pfcls.IpfcModelItem
    Dim composant As pfcls.IpfcComponentFeat
    Dim modele As pfcls.IpfcModel
    Dim ids As pfcls.Iintseq
    Dim fabrique As pfcls.CMpfcAssembly
    Dim chemin As pfcls.IpfcComponentPath
    Dim nb As Long
    Dim i As Long

    Set squelette = Nothing
    Set pose = Nothing
    nomSquelette = ""
    On Error Resume Next
    Set objetAsm = asm
    Set squelette = objetAsm.GetSkeleton()
    Err.Clear
    On Error GoTo 0
    If squelette Is Nothing Then Exit Function
    Set modeleSq = squelette
    nomSquelette = modeleSq.Filename

    Set features = solide.ListFeaturesByType(False, pfcls.EpfcFEATTYPE_COMPONENT)
    If Not features Is Nothing Then
        For i = 0 To features.Count - 1
            Set feature = features.item(i)
            Set item = feature
            Set composant = feature
            Set modele = Nothing
            On Error Resume Next
            Set modele = session.GetModelFromDescr(composant.ModelDescr)
            Err.Clear
            On Error GoTo 0
            If Not modele Is Nothing Then
                If StrComp(modele.Filename, nomSquelette, vbTextCompare) = 0 Then
                    nb = nb + 1
                    Set ids = New pfcls.Cintseq
                    ids.Append item.id
                    Set fabrique = New pfcls.CMpfcAssembly
                    Set chemin = fabrique.CreateComponentPath(asm, ids)
                    Set pose = chemin.GetTransform(True)
                End If
            End If
        Next i
    End If
    If nb <> 1 Then Err.Raise PAL_ERR + 60, , _
        "Impossible d'identifier une seule occurrence du squelette " & _
        nomSquelette & " (trouve : " & nb & ")."
    If Not PalPoseAlignee(pose) Then Err.Raise PAL_ERR + 61, , _
        "Le squelette " & nomSquelette & " est tourne dans l'assemblage : " & _
        "ses axes X/Y/Z doivent etre ceux de l'assemblage."
    PalPoseSquelette = True
End Function

Private Function PalPoseAlignee(ByVal pose As pfcls.IpfcTransform3D) As Boolean
    Dim vx As pfcls.IpfcVector3D
    Dim vy As pfcls.IpfcVector3D
    Dim vz As pfcls.IpfcVector3D
    Set vx = pose.GetXAxis(): Set vy = pose.GetYAxis(): Set vz = pose.GetZAxis()
    If Abs(CDbl(vx.item(0)) - 1#) > 0.0001 Then Exit Function
    If Abs(CDbl(vx.item(1))) > 0.0001 Then Exit Function
    If Abs(CDbl(vx.item(2))) > 0.0001 Then Exit Function
    If Abs(CDbl(vy.item(0))) > 0.0001 Then Exit Function
    If Abs(CDbl(vy.item(1)) - 1#) > 0.0001 Then Exit Function
    If Abs(CDbl(vy.item(2))) > 0.0001 Then Exit Function
    If Abs(CDbl(vz.item(0))) > 0.0001 Then Exit Function
    If Abs(CDbl(vz.item(1))) > 0.0001 Then Exit Function
    If Abs(CDbl(vz.item(2)) - 1#) > 0.0001 Then Exit Function
    PalPoseAlignee = True
End Function

Private Function PalOrigine(ByVal pose As pfcls.IpfcTransform3D, _
    ByVal k As Long) As Double
    Dim o As pfcls.IpfcPoint3D
    If pose Is Nothing Then Exit Function
    Set o = pose.GetOrigin()
    PalOrigine = CDbl(o.item(k))
End Function

' Array : axe ("X", "Y", "Z"), x1, y1, z1, x2, y2, z2 dans le repere de
' l'assemblage. Les courbes non droites ou inclinees sont comptees a part.
Private Function PalLignesSquelette(ByVal squelette As pfcls.IpfcSolid, _
    ByVal pose As pfcls.IpfcTransform3D, ByRef autres As Long) As Collection

    Dim lignes As Collection
    Dim proprietaire As pfcls.IpfcModelItemOwner
    Dim items As pfcls.IpfcModelItems
    Dim features As pfcls.IpfcModelItems
    Dim feat As pfcls.IpfcFeature
    Dim i As Long

    Set lignes = New Collection
    autres = 0
    Set proprietaire = squelette
    Set items = proprietaire.ListItems(pfcls.EpfcITEM_CURVE)
    If Not items Is Nothing Then PalAjouterLignes lignes, items, pose, autres
    If lignes.Count = 0 Then
        Set features = proprietaire.ListItems(pfcls.EpfcITEM_FEATURE)
        If Not features Is Nothing Then
            For i = 0 To features.Count - 1
                Set feat = features.item(i)
                Set items = Nothing
                On Error Resume Next
                Set items = feat.ListSubItems(pfcls.EpfcITEM_CURVE)
                Err.Clear
                On Error GoTo 0
                If Not items Is Nothing Then PalAjouterLignes lignes, items, pose, autres
            Next i
        End If
    End If
    Set PalLignesSquelette = lignes
End Function

Private Sub PalAjouterLignes(ByVal lignes As Collection, _
    ByVal items As pfcls.IpfcModelItems, ByVal pose As pfcls.IpfcTransform3D, _
    ByRef autres As Long)

    Dim i As Long
    Dim item As pfcls.IpfcModelItem
    Dim courbe As pfcls.IpfcGeomCurve
    Dim descripteur As pfcls.IpfcCurveDescriptor
    Dim segment As pfcls.IpfcLineDescriptor
    Dim p As pfcls.IpfcPoint3D
    Dim q As pfcls.IpfcPoint3D
    Dim a(0 To 2) As Double
    Dim b(0 To 2) As Double
    Dim k As Long
    Dim axe As String
    Dim lu As Boolean
    Dim t As Double

    For i = 0 To items.Count - 1
        lu = False
        Set segment = Nothing
        On Error Resume Next
        Set item = items.item(i)
        Set courbe = item
        Set descripteur = courbe.GetCurveDescriptor()
        Set segment = descripteur
        If Err.Number = 0 And Not segment Is Nothing Then
            Set p = pose.TransformPoint(segment.End1)
            Set q = pose.TransformPoint(segment.End2)
            For k = 0 To 2
                a(k) = CDbl(p.item(k))
                b(k) = CDbl(q.item(k))
            Next k
            lu = (Err.Number = 0)
        End If
        Err.Clear
        On Error GoTo 0
        axe = ""
        If lu Then
            If Abs(b(0) - a(0)) > 0.05 And Abs(b(1) - a(1)) <= 0.01 And _
                Abs(b(2) - a(2)) <= 0.01 Then axe = "X"
            If Abs(b(1) - a(1)) > 0.05 And Abs(b(0) - a(0)) <= 0.01 And _
                Abs(b(2) - a(2)) <= 0.01 Then axe = "Y"
            If Abs(b(2) - a(2)) > 0.05 And Abs(b(0) - a(0)) <= 0.01 And _
                Abs(b(1) - a(1)) <= 0.01 Then axe = "Z"
        End If
        If axe = "" Then
            autres = autres + 1
        Else
            ' Points ordonnes dans le sens de l'axe.
            k = InStr(1, "XYZ", axe) - 1
            If a(k) > b(k) Then
                For k = 0 To 2
                    t = a(k): a(k) = b(k): b(k) = t
                Next k
            End If
            lignes.Add Array(axe, a(0), a(1), a(2), b(0), b(1), b(2))
        End If
    Next i
End Sub

Private Sub PalEnveloppe(ByVal session As pfcls.IpfcBaseSession, _
    ByVal asm As pfcls.IpfcAssembly, ByVal solide As pfcls.IpfcSolid, _
    ByVal nomSquelette As String, ByRef mini() As Double, _
    ByRef maxi() As Double, ByRef nbPieces As Long, ByRef nbIgnores As Long)

    ' Enveloppe des composants directs de l'assemblage (squelette exclu),
    ' dans le repere de l'assemblage.
    Dim features As pfcls.IpfcFeatures
    Dim feature As pfcls.IpfcFeature
    Dim item As pfcls.IpfcModelItem
    Dim composant As pfcls.IpfcComponentFeat
    Dim modele As pfcls.IpfcModel
    Dim piece As pfcls.IpfcSolid
    Dim boite As pfcls.IpfcOutline3D
    Dim ids As pfcls.Iintseq
    Dim fabrique As pfcls.CMpfcAssembly
    Dim chemin As pfcls.IpfcComponentPath
    Dim pose As pfcls.IpfcTransform3D
    Dim coin As pfcls.IpfcPoint3D
    Dim resultat As pfcls.IpfcPoint3D
    Dim c As Long
    Dim k As Long
    Dim i As Long
    Dim ok As Boolean

    ReDim mini(0 To 2)
    ReDim maxi(0 To 2)
    For k = 0 To 2
        mini(k) = 1E+30
        maxi(k) = -1E+30
    Next k
    nbPieces = 0
    nbIgnores = 0
    Set features = solide.ListFeaturesByType(False, pfcls.EpfcFEATTYPE_COMPONENT)
    If features Is Nothing Then Exit Sub
    For i = 0 To features.Count - 1
        Set feature = features.item(i)
        Set item = feature
        Set composant = feature
        Set modele = Nothing
        On Error Resume Next
        Set modele = session.GetModelFromDescr(composant.ModelDescr)
        Err.Clear
        On Error GoTo 0
        If modele Is Nothing Then
            nbIgnores = nbIgnores + 1
        ElseIf StrComp(modele.Filename, nomSquelette, vbTextCompare) <> 0 Then
            ok = False
            On Error Resume Next
            Err.Clear
            Set piece = modele
            Set boite = piece.GeomOutline
            Set ids = New pfcls.Cintseq
            ids.Append item.id
            Set fabrique = New pfcls.CMpfcAssembly
            Set chemin = fabrique.CreateComponentPath(asm, ids)
            Set pose = chemin.GetTransform(True)
            ok = (Err.Number = 0 And Not boite Is Nothing And Not pose Is Nothing)
            Err.Clear
            On Error GoTo 0
            If ok Then
                ' Les 8 coins de la boite locale, transformes dans l'assemblage.
                For c = 0 To 7
                    Set coin = New pfcls.CpfcPoint3D
                    coin.Set 0, CDbl(boite.item(c Mod 2).item(0))
                    coin.Set 1, CDbl(boite.item((c \ 2) Mod 2).item(1))
                    coin.Set 2, CDbl(boite.item((c \ 4) Mod 2).item(2))
                    Set resultat = pose.TransformPoint(coin)
                    For k = 0 To 2
                        If CDbl(resultat.item(k)) < mini(k) Then mini(k) = CDbl(resultat.item(k))
                        If CDbl(resultat.item(k)) > maxi(k) Then maxi(k) = CDbl(resultat.item(k))
                    Next k
                Next c
                nbPieces = nbPieces + 1
            Else
                nbIgnores = nbIgnores + 1
            End If
        End If
    Next i
End Sub

' Ligne horizontale la plus haute strictement sous (x, y) qui croise la
' verticale du point : ligne X dont l'etendue couvre x, ou ligne Z passant
' par x. La profondeur Z est ignoree, comme dans le module pieds moteurs.
Private Function PalLigneSousPoint(ByVal lignes As Collection, _
    ByVal x As Double, ByVal y As Double, ByRef yLigne As Double) As Boolean
    Dim i As Long
    Dim a As Variant
    Dim trouve As Boolean
    Dim impact As Double
    For i = 1 To lignes.Count
        a = lignes.item(i)
        impact = 1E+30
        If PalCroiseVerticale(a, x) Then
            impact = a(2)
        ElseIf a(0) = "Y" Then
            ' Poteau dans la verticale : le module moteur retient son sommet.
            If Abs(a(1) - x) <= PAL_TOL_LIGNE Then impact = a(5)
        End If
        If impact < y - 0.01 Then
            If Not trouve Or impact > yLigne Then
                yLigne = impact
                trouve = True
            End If
        End If
    Next i
    PalLigneSousPoint = trouve
End Function

Private Function PalLigneA(ByVal lignes As Collection, _
    ByVal x As Double, ByVal y As Double) As Boolean
    Dim i As Long
    Dim a As Variant
    For i = 1 To lignes.Count
        a = lignes.item(i)
        If PalCroiseVerticale(a, x) Then
            If Abs(a(2) - y) <= PAL_TOL_CONTROLE Then
                PalLigneA = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function PalCroiseVerticale(ByVal a As Variant, ByVal x As Double) As Boolean
    If a(0) = "X" Then
        PalCroiseVerticale = (x >= a(1) - PAL_TOL_LIGNE And x <= a(4) + PAL_TOL_LIGNE)
    ElseIf a(0) = "Z" Then
        PalCroiseVerticale = (Abs(a(1) - x) <= PAL_TOL_LIGNE)
    End If
End Function

Private Function PalNiveauPresent(ByVal lignes As Collection, _
    ByVal y As Double) As Boolean
    Dim i As Long
    Dim a As Variant
    For i = 1 To lignes.Count
        a = lignes.item(i)
        If a(0) = "X" Or a(0) = "Z" Then
            If Abs(a(2) - y) <= 0.5 Then
                PalNiveauPresent = True
                Exit Function
            End If
        End If
    Next i
End Function

' ---------------------------------------------------------------- Controle

Public Sub HEAT_ControlerSqueletteHeating()
    Dim nbKO As Long
    Dim resume As String
    PalControlerTout False, nbKO, resume
End Sub

' Pilote : controle sans fenetre. Vrai si le controle a pu aller au bout ;
' nbKO = nombre de KO bloquants (les ALERTE ne bloquent pas).
Public Function HEAT_PalControlerAuto(ByRef nbKO As Long, _
    ByRef resume As String) As Boolean
    HEAT_PalControlerAuto = PalControlerTout(True, nbKO, resume)
End Function

Private Function PalControlerTout(ByVal silencieux As Boolean, ByRef nbKO As Long, _
    ByRef resume As String) As Boolean
    Dim session As pfcls.IpfcBaseSession
    Dim modele As pfcls.IpfcModel
    Dim asm As pfcls.IpfcAssembly
    Dim solide As pfcls.IpfcSolid
    Dim squelette As pfcls.IpfcSolid
    Dim pose As pfcls.IpfcTransform3D
    Dim nomSquelette As String
    Dim aSquelette As Boolean
    Dim lignes As Collection
    Dim autres As Long
    Dim s As Worksheet
    Dim n As Worksheet
    Dim rapport As Collection
    Dim ox As Double, oy As Double, oz As Double
    Dim exMin(0 To 2) As Double
    Dim exMax(0 To 2) As Double
    Dim crMin(0 To 2) As Double
    Dim crMax(0 To 2) As Double
    Dim envMin() As Double
    Dim envMax() As Double
    Dim nbPieces As Long
    Dim nbIgnores As Long
    Dim a As Variant
    Dim i As Long
    Dim k As Long
    Dim r As Long
    Dim y As Double
    Dim manquants As String
    Dim nbManquants As Long
    Dim rouleaux As Collection
    Dim ignores As Long
    Dim reperesAsm As Long
    Dim ids As pfcls.Iintseq
    Dim nR As Long
    Dim ys() As Double
    Dim idx() As Long
    Dim debutGroupe() As Long
    Dim finGroupe() As Long
    Dim nbGroupes As Long
    Dim yRefGroupe() As Double
    Dim abaisse() As Boolean
    Dim g As Long
    Dim yLigne As Double
    Dim yPalAsm As Double
    Dim xR As Double
    Dim ecart As Double
    Dim nbFaux As Long
    Dim listeFaux As String
    Dim surPalier As Boolean
    Dim nomG As String
    Dim dCtl As Double
    Dim nbRefCtl As Long
    Dim problemeCtl As String
    Dim dMesure As Boolean
    Dim bottomY As Double, topY As Double, topOK As Boolean, minGap As Double
    Dim nomAxe As String
    Dim centreEnv As Double
    Dim centreCadre As Double
    Dim entoure As Boolean
    Dim etape As String

    On Error GoTo Echec
    nbKO = 0
    resume = ""
    Set rapport = New Collection

    etape = "Lecture Excel"
    HEAT_Initialiser
    Set s = ThisWorkbook.Worksheets("Structure_HEATING")
    Set n = ThisWorkbook.Worksheets("Niveaux_HEATING")
    PalLireStructure s, bottomY, topY, topOK, minGap
    exMin(0) = CDbl(s.Range("G7").value)
    exMax(0) = exMin(0) + CDbl(s.Range("G4").value)
    exMin(2) = CDbl(s.Range("G8").value)
    exMax(2) = exMin(2) + CDbl(s.Range("G5").value)
    exMin(1) = bottomY
    If PalNumerique(s.Range("B6").value) Then
        exMax(1) = CDbl(s.Range("B6").value)
    ElseIf topOK Then
        exMax(1) = topY
    End If

    etape = "Connexion Creo"
    Set session = PalSession()
    Set modele = session.GetActiveModel()
    If modele Is Nothing Then Err.Raise PAL_ERR + 1, , "Activer l'assemblage HEATING dans Creo."
    If modele.Type <> pfcls.EpfcMDL_ASSEMBLY Then Err.Raise PAL_ERR + 2, , _
        "Le modele actif doit etre l'assemblage HEATING (.asm)."
    Set asm = modele
    Set solide = modele

    etape = "Lecture du squelette"
    aSquelette = PalPoseSquelette(session, asm, solide, squelette, nomSquelette, pose)
    If Not aSquelette Then
        rapport.Add Array("Squelette", "present", "absent", "", "KO", _
            "Aucun squelette dans " & modele.Filename & ".")
    Else
        ox = PalOrigine(pose, 0): oy = PalOrigine(pose, 1): oz = PalOrigine(pose, 2)
        rapport.Add Array("Position du squelette " & nomSquelette, "0 ; 0 ; 0", _
            PalFmt(ox) & " ; " & PalFmt(oy) & " ; " & PalFmt(oz), "", _
            IIf(Abs(ox) + Abs(oy) + Abs(oz) <= 0.01, "OK", "ALERTE"), _
            IIf(Abs(ox) + Abs(oy) + Abs(oz) <= 0.01, "Squelette a l'origine de l'assemblage.", _
            "Squelette decale dans l'assemblage : toutes ses lignes sont decalees " & _
            "d'autant. Les paliers le compensent (decalage mesure) ; ne plus le deplacer."))
        Set lignes = PalLignesSquelette(squelette, pose, autres)
        ' Decalage reel Excel -> Creo (position du squelette + export IGES).
        dCtl = oy
        dMesure = PalMesurerDecalage(lignes, dCtl, nbRefCtl, problemeCtl)
        If Not dMesure Then dCtl = oy
        rapport.Add Array("Decalage Excel -> Creo (niveaux retrouves)", "0", _
            IIf(dMesure, PalFmt(dCtl), "non mesurable"), _
            IIf(dMesure, PalFmt(dCtl), ""), _
            IIf(dMesure, IIf(Abs(dCtl) <= PAL_TOL_SQUELETTE, "OK", "ALERTE"), "KO"), _
            IIf(dMesure, IIf(Abs(dCtl) <= PAL_TOL_SQUELETTE, _
            "Les niveaux Excel sont a la meme hauteur dans Creo.", _
            "Les niveaux Excel sont " & PalFmt(dCtl) & " mm plus haut dans Creo " & _
            "(dont position du squelette " & PalFmt(oy) & " mm). Les paliers le " & _
            "compensent s'ils ont ete places avec ce meme d."), problemeCtl))
        If lignes.Count = 0 Then
            rapport.Add Array("Lignes du squelette", "", "0", "", "KO", _
                "Aucune ligne droite lisible dans le squelette.")
        Else
            ' Cadre reel dans Creo : poteaux = lignes verticales.
            For k = 0 To 2
                crMin(k) = 1E+30: crMax(k) = -1E+30
            Next k
            For i = 1 To lignes.Count
                a = lignes.item(i)
                If a(0) = "Y" Then
                    If a(1) < crMin(0) Then crMin(0) = a(1)
                    If a(1) > crMax(0) Then crMax(0) = a(1)
                    If a(3) < crMin(2) Then crMin(2) = a(3)
                    If a(3) > crMax(2) Then crMax(2) = a(3)
                    If a(2) < crMin(1) Then crMin(1) = a(2)
                    If a(5) > crMax(1) Then crMax(1) = a(5)
                End If
            Next i
            For k = 0 To 2
                nomAxe = Mid$("XYZ", k + 1, 1)
                rapport.Add Array("Cadre " & nomAxe & " : squelette Creo / calcul Excel", _
                    PalFmt(exMin(k) + PalOrigine(pose, k)) & " a " & _
                    PalFmt(exMax(k) + PalOrigine(pose, k)), _
                    PalFmt(crMin(k)) & " a " & PalFmt(crMax(k)), _
                    PalFmt(crMin(k) - exMin(k) - PalOrigine(pose, k)), _
                    IIf(Abs(crMin(k) - exMin(k) - PalOrigine(pose, k)) <= 1 And _
                        Abs(crMax(k) - exMax(k) - PalOrigine(pose, k)) <= 1, "OK", "KO"), _
                    "Si KO : le squelette Creo ne correspond pas au dernier calcul " & _
                    "Excel. Recharger l'IGES AXES_HEATING_3D.igs dans le squelette.")
            Next k

            ' Niveaux calcules (Niveaux_HEATING) presents dans le squelette ?
            For r = PAL_CALC_FIRST To PAL_CALC_LAST
                If PalNumerique(n.Cells(r, 4).value) Then
                    y = CDbl(n.Cells(r, 4).value) + dCtl
                    If Not PalNiveauPresent(lignes, y) Then
                        nbManquants = nbManquants + 1
                        If nbManquants <= 10 Then manquants = manquants & _
                            IIf(manquants = "", "", ", ") & _
                            Trim$(PalTexte(n.Cells(r, 2).value)) & " (Y " & PalFmt(y) & ")"
                    End If
                End If
            Next r
            rapport.Add Array("Etages Excel presents dans le squelette Creo", _
                "tous", IIf(nbManquants = 0, "tous", nbManquants & " absent(s)"), "", _
                IIf(nbManquants = 0, "OK", "KO"), _
                IIf(nbManquants = 0, "Le squelette Creo est a jour.", _
                "Absents : " & manquants & ". Recharger l'IGES dans le squelette."))
        End If
    End If

    etape = "Enveloppe de l'assemblage"
    PalEnveloppe session, asm, solide, nomSquelette, envMin, envMax, nbPieces, nbIgnores
    If nbPieces > 0 And aSquelette Then
        If lignes.Count > 0 Then
            For k = 0 To 2
                If k <> 1 Then
                    nomAxe = Mid$("XYZ", k + 1, 1)
                    centreEnv = (envMin(k) + envMax(k)) / 2#
                    centreCadre = (crMin(k) + crMax(k)) / 2#
                    entoure = (crMin(k) <= envMin(k) + 0.5 And crMax(k) >= envMax(k) - 0.5)
                    rapport.Add Array("Cadre autour de l'assemblage en " & nomAxe, _
                        "assemblage " & PalFmt(envMin(k)) & " a " & PalFmt(envMax(k)), _
                        "cadre " & PalFmt(crMin(k)) & " a " & PalFmt(crMax(k)), _
                        PalFmt(centreCadre - centreEnv), IIf(entoure, "OK", "ALERTE"), _
                        IIf(entoure, "Le cadre entoure l'assemblage.", _
                        "Le cadre est decale de " & PalFmt(centreCadre - centreEnv) & _
                        " mm en " & nomAxe & " par rapport au centre de l'assemblage."))
                End If
            Next k
        End If
    End If

    etape = "Paliers enregistres"
    If aSquelette Then PalControlerPaliersEnregistres rapport, dCtl

    etape = "Paliers sous les rouleaux"
    Set rouleaux = New Collection
    Set ids = New pfcls.Cintseq
    reperesAsm = PalCompterReperesAsm(modele)
    PalDebutParcours
    PalParcourir session, asm, solide, ids, 0, rouleaux, ignores, reperesAsm, _
        modele.Filename
    nR = rouleaux.Count
    rapport.Add Array("Composants charges dans Creo", "tous", _
        IIf(ignores = 0, "tous", ignores & " non charge(s)"), "", _
        IIf(ignores = 0, "OK", "KO"), IIf(ignores = 0, _
        "Tous les composants sont charges : aucun rouleau oublie.", _
        "Non charges (rouleaux non verifies) : " & PalCouper(m_PalNonCharges, 400)))
    rapport.Add Array("Reperes ROLL dans un squelette", "0", _
        CStr(m_PalReperesSquelette), "", IIf(m_PalReperesSquelette = 0, "OK", "KO"), _
        IIf(m_PalReperesSquelette = 0, "Aucun.", _
        "Le module moteur poserait un rouleau sur le squelette : supprimer ces reperes."))
    If nR > 0 Then
        nbFaux = 0
        listeFaux = ""
        For i = 1 To nR
            If StrComp(PalChampTexte(rouleaux, i, 6), modele.Filename, vbTextCompare) <> 0 Then
                nbFaux = nbFaux + 1
                If nbFaux <= 3 Then
                    If listeFaux <> "" Then listeFaux = listeFaux & ", "
                    listeFaux = listeFaux & LCase$(PalBase(PalChampTexte(rouleaux, i, 0))) & _
                        " dans " & PalChampTexte(rouleaux, i, 6)
                End If
            End If
        Next i
        rapport.Add Array("Rouleaux portes directement par " & modele.Filename, _
            CStr(nR), CStr(nR - nbFaux), "", IIf(nbFaux = 0, "OK", "KO"), _
            IIf(nbFaux = 0, "Les pieds se caleront sur le squelette HEATING (paliers).", _
            nbFaux & " rouleau(x) dans un sous-assemblage : leur pied se calerait sur " & _
            "le squelette de ce sous-assemblage, pas sur les paliers. " & listeFaux & _
            IIf(nbFaux > 3, ", ...", "")))
    End If
    If nR > 0 And aSquelette Then
        ReDim ys(1 To nR)
        ReDim idx(1 To nR)
        For i = 1 To nR
            idx(i) = i
            ys(i) = PalChamp(rouleaux, i, 3)
        Next i
        PalTrier ys, idx, nR
        PalGrouper ys, idx, nR, PAL_TOLERANCE_DEFAUT, debutGroupe, finGroupe, _
            nbGroupes, yRefGroupe, abaisse
        For g = nbGroupes To 1 Step -1
            nomG = PalNomRangee(g, nbGroupes)
            ' Palier attendu, dans le repere de l'assemblage.
            yPalAsm = yRefGroupe(g) - PAL_DISTANCE_PALIER
            If yPalAsm - dCtl < bottomY - 0.000001 Then
                rapport.Add Array(nomG, "pas de palier", "", "", "OK", _
                    "Palier sous le pied des poteaux : non cree (regle).")
            ElseIf lignes.Count = 0 Then
                rapport.Add Array(nomG, PalFmt(PAL_DISTANCE_PALIER) & " mm", _
                    "aucune ligne", "", "KO", "Squelette sans ligne lisible.")
            Else
                ' 1. Rouleau le plus bas : distance a la premiere ligne dessous.
                i = debutGroupe(g)
                xR = PalChamp(rouleaux, idx(i), 2)
                If PalLigneSousPoint(lignes, xR, ys(i), yLigne) Then
                    ecart = ys(i) - yLigne - PAL_DISTANCE_PALIER
                    rapport.Add Array(nomG & " : rouleau le plus bas -> premiere ligne dessous", _
                        PalFmt(PAL_DISTANCE_PALIER) & " mm", PalFmt(ys(i) - yLigne) & " mm", _
                        PalFmt(ecart), IIf(Abs(ecart) <= PAL_TOL_CONTROLE, "OK", "KO"), _
                        "Centre Y " & PalFmt(ys(i)) & " ; ligne Y " & PalFmt(yLigne) & _
                        IIf(Abs(ecart) <= PAL_TOL_CONTROLE, "", " : " & _
                        PalCauseEcart(lignes, xR, ys(i), yLigne, yPalAsm, dCtl)))
                Else
                    rapport.Add Array(nomG & " : rouleau le plus bas -> premiere ligne dessous", _
                        PalFmt(PAL_DISTANCE_PALIER) & " mm", "aucune ligne", "", "KO", _
                        "Aucune ligne du squelette ne croise la verticale du rouleau (X " & _
                        PalFmt(xR) & ") sous Y " & PalFmt(ys(i)) & _
                        " : cadre decale en X par rapport aux rouleaux ?")
                End If

                ' 2. Chaque rouleau du groupe doit trouver le palier en premier.
                nbFaux = 0
                listeFaux = ""
                For i = debutGroupe(g) To finGroupe(g)
                    xR = PalChamp(rouleaux, idx(i), 2)
                    surPalier = PalLigneSousPoint(lignes, xR, ys(i), yLigne)
                    If surPalier Then surPalier = (Abs(yLigne - yPalAsm) <= PAL_TOL_CONTROLE)
                    If Not surPalier Then
                        nbFaux = nbFaux + 1
                        If nbFaux <= 3 Then
                            If listeFaux <> "" Then listeFaux = listeFaux & ", "
                            listeFaux = listeFaux & LCase$(PalBase(PalChampTexte(rouleaux, idx(i), 0))) & _
                                " Y " & PalFmt(ys(i))
                        End If
                    End If
                Next i
                rapport.Add Array(nomG & " : chaque rouleau se cale sur le palier", _
                    (finGroupe(g) - debutGroupe(g) + 1) & " rouleau(x)", _
                    (finGroupe(g) - debutGroupe(g) + 1 - nbFaux) & " sur le palier", "", _
                    IIf(nbFaux = 0, "OK", "KO"), _
                    IIf(nbFaux = 0, "Aucune ligne entre les rouleaux et le palier Y " & _
                        PalFmt(yPalAsm) & ".", _
                    nbFaux & " rouleau(x) trouvent une autre ligne (ou aucune) avant le palier : " & _
                        listeFaux & IIf(nbFaux > 3, ", ...", "") & _
                        ". Relancer HEAT_PlacerPaliersRouleaux puis recharger l'IGES."))
            End If
        Next g
    End If
    If reperesAsm > 0 Then rapport.Add Array("Reperes ROLL_CENTER d'assemblage", _
        "0", CStr(reperesAsm), "", "INFO", _
        "Reperes poses dans un assemblage et non dans une piece : ignores par les paliers.")

    etape = "Ecriture du rapport"
    PalEcrireRapport rapport, modele.Filename, silencieux, nbKO, resume
    PalControlerTout = True
    Exit Function

Echec:
    resume = "Controle interrompu. Etape : " & etape & " ; erreur " & _
        CStr(Err.Number) & " : " & Err.Description
    If Not silencieux Then MsgBox "Controle HEATING interrompu." & vbCrLf & _
        "Etape : " & etape & vbCrLf & "Erreur : " & CStr(Err.Number) & vbCrLf & _
        Err.Description, vbCritical, "Controle HEATING"
End Function

' Explication d'une distance rouleau -> premiere ligne differente de 1500.
Private Function PalCauseEcart(ByVal lignes As Collection, ByVal x As Double, _
    ByVal yCentre As Double, ByVal yLigne As Double, ByVal yPalAsm As Double, _
    ByVal oy As Double) As String
    If Abs(oy) > PAL_TOL_SQUELETTE And _
        Abs((yCentre - yLigne) - (PAL_DISTANCE_PALIER - oy)) <= PAL_TOL_CONTROLE Then
        PalCauseEcart = "ecart = decalage Excel -> Creo (" & PalFmt(oy) & _
            " mm) non compense : relancer HEAT_PlacerPaliersRouleaux (il le mesure), " & _
            "puis recharger l'IGES."
    ElseIf yLigne > yPalAsm + PAL_TOL_CONTROLE Then
        PalCauseEcart = "une ligne Y " & PalFmt(yLigne) & " est entre le rouleau et " & _
            "le palier Y " & PalFmt(yPalAsm) & " (niveau actif). Relancer les paliers, " & _
            "recharger l'IGES."
    ElseIf Not PalLigneA(lignes, x, yPalAsm) Then
        PalCauseEcart = "aucune ligne a Y " & PalFmt(yPalAsm) & _
            " : palier non ecrit ou IGES non recharge dans le squelette."
    Else
        PalCauseEcart = "relancer HEAT_PlacerPaliersRouleaux puis recharger l'IGES."
    End If
End Function

Private Sub PalEcrireRapport(ByVal rapport As Collection, ByVal nomAssemblage As String, _
    ByVal silencieux As Boolean, ByRef nbKO As Long, ByRef texteKO As String)
    Dim ws As Worksheet
    Dim ligne As Long
    Dim e As Variant
    Dim k As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CONTROLE_HEATING")
    Err.Clear
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = "CONTROLE_HEATING"
        ws.Tab.Color = RGB(223, 113, 36)
    End If
    ws.Cells.Clear
    ws.Range("A1").value = "Controle du squelette HEATING dans " & nomAssemblage & _
        " - " & Format$(Now, "yyyy-mm-dd hh:nn") & " (aucune modification, version " & _
        PAL_VERSION & ")"
    ws.Range("A2").value = "KO = bloquant (a corriger). ALERTE = a connaitre, ne bloque pas."
    ws.Range("A3:F3").value = Array("Controle", "Attendu", "Mesure dans Creo", _
        "Ecart (mm)", "Resultat", "Explication")
    ws.Range("A3:F3").Font.Bold = True
    ligne = 4
    For Each e In rapport
        For k = 0 To 5
            ws.Cells(ligne, k + 1).value = CStr(e(k))
        Next k
        Select Case CStr(e(4))
            Case "OK": ws.Cells(ligne, 5).Interior.Color = RGB(198, 239, 206)
            Case "KO"
                ws.Cells(ligne, 5).Interior.Color = RGB(255, 199, 206)
                nbKO = nbKO + 1
                If nbKO <= 6 Then texteKO = texteKO & "- " & CStr(e(0)) & " : " & _
                    PalCouper(CStr(e(5)), 160) & vbCrLf
            Case Else: ws.Cells(ligne, 5).Interior.Color = RGB(255, 235, 156)
        End Select
        ligne = ligne + 1
    Next e
    ws.Columns("A:E").AutoFit
    ws.Columns("F").ColumnWidth = 90
    ws.Columns("F").WrapText = True
    If silencieux Then Exit Sub
    ws.Activate
    If nbKO = 0 Then
        MsgBox "Aucun KO bloquant : paliers a " & PalFmt(PAL_DISTANCE_PALIER) & _
            " mm sous les rouleaux, squelette a jour. Voir les ALERTE eventuelles " & _
            "dans CONTROLE_HEATING.", vbInformation, "Controle HEATING"
    Else
        MsgBox nbKO & " probleme(s) :" & vbCrLf & vbCrLf & texteKO & vbCrLf & _
            "Detail : feuille CONTROLE_HEATING.", vbExclamation, "Controle HEATING"
    End If
End Sub

Private Sub PalControlerPaliersEnregistres(ByVal rapport As Collection, _
    ByVal oy As Double)

    ' Relit PALIERS_HEATING : chaque palier cree doit etre a la distance
    ' enregistree sous le rouleau le plus bas de son groupe, avec la
    ' position ACTUELLE du squelette.
    Dim ws As Worksheet
    Dim r As Long
    Dim derniere As Long
    Dim yRef As Double
    Dim yLocal As Double
    Dim yAsm As Double
    Dim dist As Double
    Dim oyPlacement As Double
    Dim nom As String
    Dim ok As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PAL_FEUILLE)
    Err.Clear
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If PalNumerique(ws.Range("C5").value) Then
        oyPlacement = CDbl(ws.Range("C5").value)
        rapport.Add Array("Decalage Excel -> Creo : au placement / maintenant", _
            PalFmt(oyPlacement), PalFmt(oy), PalFmt(oy - oyPlacement), _
            IIf(Abs(oy - oyPlacement) <= PAL_TOL_SQUELETTE, "OK", "KO"), _
            IIf(Abs(oy - oyPlacement) <= PAL_TOL_SQUELETTE, _
            "Le squelette n'a pas bouge depuis le placement des paliers.", _
            "Le squelette a bouge depuis le placement : tous les paliers sont " & _
            "decales d'autant. Relancer HEAT_PlacerPaliersRouleaux puis recharger l'IGES."))
    End If

    derniere = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    For r = PAL_LIGNE_DONNEES To derniere
        If Left$(PalTexte(ws.Cells(r, 8).value), Len(PAL_PREFIXE_ID)) = PAL_PREFIXE_ID Then
            If PalNumerique(ws.Cells(r, 5).value) And PalNumerique(ws.Cells(r, 7).value) Then
                nom = PalTexte(ws.Cells(r, 1).value)
                yRef = CDbl(ws.Cells(r, 5).value)
                yLocal = CDbl(ws.Cells(r, 7).value)
                dist = PAL_DISTANCE_PALIER
                If PalNumerique(ws.Cells(r, 6).value) Then dist = CDbl(ws.Cells(r, 6).value)
                yAsm = yLocal + oy
                ok = (Abs((yRef - yAsm) - dist) <= PAL_TOL_CONTROLE)
                rapport.Add Array(nom & " : palier " & PalTexte(ws.Cells(r, 8).value) & _
                    " sous le rouleau le plus bas", _
                    PalFmt(dist) & " mm sous " & PalFmt(yRef), _
                    PalFmt(yRef - yAsm) & " mm (Y " & PalFmt(yAsm) & ")", _
                    PalFmt(yRef - yAsm - dist), IIf(ok, "OK", "KO"), _
                    IIf(ok, PalFmt(yRef - yAsm) & " mm sous le centre.", _
                    IIf(yAsm >= yRef, "PALIER AU-DESSUS DU ROULEAU. ", "") & _
                    "Relancer HEAT_PlacerPaliersRouleaux (il mesure le decalage), " & _
                    "puis recharger l'IGES."))
            End If
        End If
    Next r
End Sub
