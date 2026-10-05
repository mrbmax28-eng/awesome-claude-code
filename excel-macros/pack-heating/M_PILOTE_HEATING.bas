Attribute VB_Name = "M_PILOTE_HEATING"
Option Explicit

' ================================================================
' PILOTE HEATING - paliers + pieds moteurs en deux lancements
'
'   PILOTE_Configurer        une fois : cote moteur et assemblage moteur
'   PILOTE_Verifier          controle de depart seul, ne modifie rien
'   PILOTE_1_PaliersEtIGES   recalcul, paliers, export IGES
'        -> dans Creo : recharger AXES_HEATING_3D.igs dans le squelette
'   PILOTE_2_ControleEtPieds controle Creo, pieds moteurs, controle final
'
' Chaque etape est verifiee avant de passer a la suivante. Au premier
' probleme, le pilote s'arrete, dit quoi corriger, et note tout dans la
' feuille JOURNAL_PILOTE. L'assemblage HEATING doit etre actif dans Creo.
'
' Une seule etape reste manuelle : recharger l'IGES dans la fonction
' importee du squelette (aucune macro ne peut redefinir cette fonction
' sans supprimer le squelette, ce qui casserait les pieces qui s'y
' appuient). La phase 2 verifie que ce rechargement a bien ete fait.
'
' Modules requis (pack) : M_FRAME_HEATING, M_FRAME_HEATING_PALIERS
' 2026.10.05.09, Module_Ajout_Pieces_Techniques 2026.10.05.18.
' ================================================================

Private Const PIL_VERSION As String = "2026.10.05.01"
Private Const PIL_VERSION_PALIERS As String = "2026.10.05.09"
Private Const PIL_VERSION_MOTEUR As String = "2026.10.05.18"
Private Const PIL_FEUILLE As String = "PILOTE_HEATING"
Private Const PIL_JOURNAL As String = "JOURNAL_PILOTE"
Private Const PIL_ETAT_TECH As String = "__HUB_ETAT"
Private Const PIL_FEUILLE_PALIERS As String = "PALIERS_HEATING"
Private Const PIL_TOL As Double = 1#
' Colonnes de __HUB_ETAT (Module_Ajout_Pieces_Techniques).
Private Const PIL_COL_ASSEMBLAGE As Long = 20
Private Const PIL_COL_CLE As Long = 21
Private Const PIL_COL_REPERE As Long = 26
Private Const PIL_COL_SOURCE_MODELE As Long = 23
Private Const PIL_COL_MOTEUR_COTE As Long = 33
Private Const PIL_COL_MOTEUR_ID As Long = 35
Private Const PIL_COL_ROLL_HAUTEUR As Long = 36
Private Const PIL_COL_FOOT_HEIGHT As Long = 37

' ---------------------------------------------------------------- Public

Public Sub PILOTE_Configurer()
    Dim ws As Worksheet
    Dim choix As VbMsgBoxResult
    Dim dialogue As Object

    On Error GoTo Echec
    Set ws = PilFeuille()
    choix = MsgBox("Cote pour TOUS les moteurs :" & vbCrLf & vbCrLf & _
        "OUI = gauche (Moteur_L)" & vbCrLf & "NON = droite (Moteur_R)" & vbCrLf & _
        "ANNULER = ne rien changer", vbYesNoCancel + vbQuestion, "Pilote HEATING")
    If choix = vbCancel Then Exit Sub
    ws.Range("C5").value = IIf(choix = vbYes, "Moteur_L", "Moteur_R")

    Set dialogue = Application.FileDialog(3)
    With dialogue
        .Title = "Selectionner l'assemblage moteur-cylindre (.asm)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Assemblages Creo", "*.asm*"
        If Len(PilTexte(ws.Range("C6").value)) > 0 Then .InitialFileName = PilTexte(ws.Range("C6").value)
        If .Show = -1 Then ws.Range("C6").value = CStr(.SelectedItems(1))
    End With
    ws.Activate
    PilJournal "CONFIG", "Configuration", "OK", "Cote " & PilTexte(ws.Range("C5").value) & _
        " ; moteur " & PilTexte(ws.Range("C6").value) & " ; confirmation paliers " & _
        PilTexte(ws.Range("C7").value)
    MsgBox "Configuration enregistree dans " & PIL_FEUILLE & "." & vbCrLf & _
        "Cote : " & PilTexte(ws.Range("C5").value) & vbCrLf & _
        "Moteur : " & PilTexte(ws.Range("C6").value), vbInformation, "Pilote HEATING"
    Exit Sub
Echec:
    MsgBox "Configuration interrompue : " & Err.Description, vbExclamation, "Pilote HEATING"
End Sub

Public Sub PILOTE_Verifier()
    Dim rapport As String
    On Error GoTo Echec
    If PilPreVol("VERIF", rapport) Then
        MsgBox "Tout est pret." & vbCrLf & vbCrLf & PilCouper(rapport, 800), _
            vbInformation, "Pilote HEATING - verification"
    Else
        MsgBox "A corriger avant de lancer le pilote :" & vbCrLf & vbCrLf & _
            PilCouper(rapport, 800) & vbCrLf & vbCrLf & "Detail : " & PIL_JOURNAL & ".", _
            vbExclamation, "Pilote HEATING - verification"
    End If
    Exit Sub
Echec:
    PilErreur "VERIF", "Verification", Err.Number, Err.Description
End Sub

Public Sub PILOTE_1_PaliersEtIGES()
    Dim rapport As String
    Dim message As String
    Dim fichier As String
    Dim d As Double
    Dim probleme As String
    Dim etape As String
    Dim ws As Worksheet
    Dim errNumero As Long
    Dim errTexte As String

    On Error GoTo Echec
    Set ws = PilFeuille()
    etape = "Controle de depart"
    If Not PilPreVol("PHASE 1", rapport) Then
        PilArret "PHASE 1", etape, rapport
        Exit Sub
    End If

    etape = "Recalcul HEATING"
    HEAT_DefinirSilencieux True
    HEAT_Recalculer
    HEAT_DefinirSilencieux False
    If Not HEAT_CalculValide() Then
        PilArret "PHASE 1", etape, "HEATING ne se recalcule pas : " & HEAT_DerniereErreur() & _
            vbCrLf & "Corriger NiveauxY_HEATING / Structure_HEATING, puis relancer."
        Exit Sub
    End If
    PilJournal "PHASE 1", etape, "OK", "Niveaux_HEATING a jour"

    ' Le decalage Excel -> Creo se mesure sur un squelette a jour.
    etape = "Squelette a jour avant paliers"
    If Not HEAT_PalDecalageMesurable(d, probleme) Then
        If HEAT_ExporterIGESAuto(fichier, message) Then
            PilEtat "PHASE 1 EN ATTENTE : recharger l'IGES puis relancer PILOTE_1", fichier
            PilArret "PHASE 1", etape, "Le squelette ne correspond pas aux niveaux Excel (" & _
                probleme & ")." & vbCrLf & vbCrLf & "L'IGES a jour vient d'etre exporte :" & _
                vbCrLf & fichier & vbCrLf & vbCrLf & "Dans Creo : recharger ce fichier dans la " & _
                "fonction importee du squelette HEATING, regenerer, puis relancer PILOTE_1."
        Else
            PilArret "PHASE 1", etape, probleme & vbCrLf & message
        End If
        Exit Sub
    End If
    PilJournal "PHASE 1", etape, "OK", probleme

    etape = "Placement des paliers"
    If Not HEAT_PalPlacerAuto(PilConfirmerPaliers(), message) Then
        PilArret "PHASE 1", etape, message
        Exit Sub
    End If
    PilJournal "PHASE 1", etape, "OK", message

    etape = "Export IGES"
    If Not HEAT_ExporterIGESAuto(fichier, message) Then
        PilArret "PHASE 1", etape, message & vbCrLf & _
            "Les paliers sont ecrits dans Excel : relancer PILOTE_1 apres correction."
        Exit Sub
    End If
    PilJournal "PHASE 1", etape, "OK", fichier

    PilEtat "PHASE 1 TERMINEE " & Format$(Now, "yyyy-mm-dd hh:nn") & _
        " : recharger l'IGES puis lancer PILOTE_2", fichier
    ThisWorkbook.Save
    MsgBox "Phase 1 terminee." & vbCrLf & vbCrLf & PilCouper(message, 300) & vbCrLf & _
        vbCrLf & "A FAIRE dans Creo :" & vbCrLf & "1. Recharger ce fichier dans la " & _
        "fonction importee du squelette HEATING :" & vbCrLf & fichier & vbCrLf & _
        "2. Regenerer et enregistrer." & vbCrLf & vbCrLf & _
        "Puis lancer PILOTE_2_ControleEtPieds.", vbInformation, "Pilote HEATING"
    Exit Sub
Echec:
    errNumero = Err.Number
    errTexte = Err.Description
    On Error Resume Next
    HEAT_DefinirSilencieux False
    On Error GoTo 0
    PilErreur "PHASE 1", etape, errNumero, errTexte
End Sub

Public Sub PILOTE_2_ControleEtPieds()
    Dim rapport As String
    Dim message As String
    Dim bilan As String
    Dim nbKO As Long
    Dim etape As String
    Dim nomAssemblage As String
    Dim ws As Worksheet

    On Error GoTo Echec
    Set ws = PilFeuille()
    etape = "Controle de depart"
    If Not PilPreVol("PHASE 2", rapport) Then
        PilArret "PHASE 2", etape, rapport
        Exit Sub
    End If
    nomAssemblage = PilNomAssemblageActif()

    etape = "Controle des paliers dans Creo"
    If Not HEAT_PalControlerAuto(nbKO, bilan) Then
        PilArret "PHASE 2", etape, bilan
        Exit Sub
    End If
    If nbKO > 0 Then
        PilArret "PHASE 2", etape, nbKO & " KO bloquant(s) dans CONTROLE_HEATING :" & _
            vbCrLf & PilCouper(bilan, 600) & vbCrLf & _
            "Le plus souvent : IGES non recharge dans le squelette. Le recharger, " & _
            "regenerer, puis relancer PILOTE_2."
        ThisWorkbook.Worksheets("CONTROLE_HEATING").Activate
        Exit Sub
    End If
    PilJournal "PHASE 2", etape, "OK", "Aucun KO bloquant (CONTROLE_HEATING)"

    etape = "Pieds moteurs"
    If Not TECH_ExecuterAuto(PilTexte(ws.Range("C5").value), _
        PilTexte(ws.Range("C6").value), message) Then
        PilArret "PHASE 2", etape, message
        Exit Sub
    End If
    PilJournal "PHASE 2", etape, "OK", Replace(message, vbCrLf, " ; ")

    etape = "Controle final des pieds"
    If Not PilControlerPieds(nomAssemblage, bilan) Then
        PilArret "PHASE 2", etape, bilan
        Exit Sub
    End If
    PilJournal "PHASE 2", etape, "OK", bilan

    PilEtat "TERMINE " & Format$(Now, "yyyy-mm-dd hh:nn") & _
        " : paliers et pieds verifies", PilTexte(ws.Range("C10").value)
    ThisWorkbook.Save
    MsgBox "HEATING termine et verifie." & vbCrLf & vbCrLf & bilan & vbCrLf & vbCrLf & _
        Replace(message, vbCrLf & vbCrLf, vbCrLf), vbInformation, "Pilote HEATING"
    Exit Sub
Echec:
    PilErreur "PHASE 2", etape, Err.Number, Err.Description
End Sub

' ---------------------------------------------------------------- Controles

' Controle de depart commun : versions, classeur, configuration, Creo.
Private Function PilPreVol(ByVal phase As String, ByRef rapport As String) As Boolean
    Dim ws As Worksheet
    Dim nbBloquants As Long
    Dim nbInventaire As Long
    Dim texte As String
    Dim cote As String
    Dim chemin As String

    rapport = ""
    Set ws = PilFeuille()

    PilVerifierVersion rapport, nbBloquants, "M_FRAME_HEATING_PALIERS", _
        HEAT_PalVersionTexte(), PIL_VERSION_PALIERS
    PilVerifierVersion rapport, nbBloquants, "Module_Ajout_Pieces_Techniques", _
        TECH_VersionTexte(), PIL_VERSION_MOTEUR

    If Len(ThisWorkbook.Path) = 0 Then
        nbBloquants = nbBloquants + 1
        PilLigne rapport, "KO  classeur jamais enregistre : l'enregistrer en .xlsm"
    ElseIf ThisWorkbook.ReadOnly Then
        nbBloquants = nbBloquants + 1
        PilLigne rapport, "KO  classeur ouvert en lecture seule"
    Else
        PilLigne rapport, "OK  classeur " & ThisWorkbook.Name
    End If

    cote = PilTexte(ws.Range("C5").value)
    chemin = PilTexte(ws.Range("C6").value)
    If StrComp(cote, "Moteur_L", vbTextCompare) <> 0 And _
        StrComp(cote, "Moteur_R", vbTextCompare) <> 0 Then
        nbBloquants = nbBloquants + 1
        PilLigne rapport, "KO  cote moteur non configure (" & PIL_FEUILLE & "!C5) : " & _
            "lancer PILOTE_Configurer"
    ElseIf Len(chemin) = 0 Then
        nbBloquants = nbBloquants + 1
        PilLigne rapport, "KO  assemblage moteur non configure (" & PIL_FEUILLE & _
            "!C6) : lancer PILOTE_Configurer"
    ElseIf Len(Dir$(chemin)) = 0 Then
        nbBloquants = nbBloquants + 1
        PilLigne rapport, "KO  assemblage moteur introuvable : " & chemin
    Else
        PilLigne rapport, "OK  moteurs " & cote & " ; " & chemin
    End If

    texte = ""
    If Not HEAT_PalInventaire(nbInventaire, texte) Then
        If nbInventaire = 0 Then nbInventaire = 1
    End If
    nbBloquants = nbBloquants + nbInventaire
    PilLigne rapport, texte

    PilPreVol = (nbBloquants = 0)
    PilJournal phase, "Controle de depart", IIf(PilPreVol, "OK", "KO"), _
        Replace(rapport, vbCrLf, " | ")
End Function

Private Sub PilVerifierVersion(ByRef rapport As String, ByRef nbBloquants As Long, _
    ByVal nomModule As String, ByVal trouvee As String, ByVal attendue As String)
    If StrComp(trouvee, attendue, vbTextCompare) = 0 Then
        PilLigne rapport, "OK  " & nomModule & " " & trouvee
    Else
        nbBloquants = nbBloquants + 1
        PilLigne rapport, "KO  " & nomModule & " " & trouvee & " importe, " & attendue & _
            " attendu : supprimer l'ancien module puis importer celui du pack"
    End If
End Sub

' Apres le module moteur : chaque rouleau HEATING a son moteur, son pied,
' et ce pied descend jusqu'au palier de son groupe (quand il en a un).
Private Function PilControlerPieds(ByVal nomAssemblage As String, _
    ByRef bilan As String) As Boolean
    Dim etat As Worksheet
    Dim pal As Worksheet
    Dim r As Long
    Dim g As Long
    Dim derniere As Long
    Dim derniereP As Long
    Dim nbLignes As Long
    Dim nbAttendus As Long
    Dim nbFaux As Long
    Dim nbSurPalier As Long
    Dim liste As String
    Dim centre As Double
    Dim pied As Double
    Dim ligne As Double
    Dim decalage As Double
    Dim yMin As Double
    Dim yMax As Double
    Dim palierAsm As Double
    Dim cote As String
    Dim nomRouleau As String
    Dim faute As String

    bilan = ""
    On Error Resume Next
    Set etat = ThisWorkbook.Worksheets(PIL_ETAT_TECH)
    Set pal = ThisWorkbook.Worksheets(PIL_FEUILLE_PALIERS)
    Err.Clear
    On Error GoTo 0
    If etat Is Nothing Then
        bilan = "Feuille " & PIL_ETAT_TECH & " absente : le module moteur n'a rien enregistre."
        Exit Function
    End If
    If pal Is Nothing Then
        bilan = "Feuille " & PIL_FEUILLE_PALIERS & " absente : relancer PILOTE_1."
        Exit Function
    End If
    If PilNumerique(pal.Range("C5").value) Then decalage = CDbl(pal.Range("C5").value)
    cote = PilTexte(PilFeuille().Range("C5").value)

    derniereP = pal.Cells(pal.Rows.Count, 1).End(xlUp).Row
    For g = 7 To derniereP
        If PilNumerique(pal.Cells(g, 2).value) Then nbAttendus = nbAttendus + CLng(pal.Cells(g, 2).value)
    Next g

    derniere = etat.Cells(etat.Rows.Count, PIL_COL_CLE).End(xlUp).Row
    For r = 2 To derniere
        If StrComp(PilNomLogique(PilTexte(etat.Cells(r, PIL_COL_ASSEMBLAGE).value)), _
            PilNomLogique(nomAssemblage), vbTextCompare) = 0 And _
            Len(PilTexte(etat.Cells(r, PIL_COL_CLE).value)) > 0 Then
            nbLignes = nbLignes + 1
            nomRouleau = PilTexte(etat.Cells(r, PIL_COL_SOURCE_MODELE).value) & " / " & _
                PilTexte(etat.Cells(r, PIL_COL_REPERE).value)
            faute = ""
            If Not PilNumerique(etat.Cells(r, PIL_COL_MOTEUR_ID).value) Then
                faute = "pas de moteur"
            ElseIf CDbl(etat.Cells(r, PIL_COL_MOTEUR_ID).value) <= 0 Then
                faute = "pas de moteur"
            ElseIf StrComp(PilTexte(etat.Cells(r, PIL_COL_MOTEUR_COTE).value), cote, _
                vbTextCompare) <> 0 Then
                faute = "moteur sur " & PilTexte(etat.Cells(r, PIL_COL_MOTEUR_COTE).value)
            ElseIf Not PilNumerique(etat.Cells(r, PIL_COL_FOOT_HEIGHT).value) Or _
                Not PilNumerique(etat.Cells(r, PIL_COL_ROLL_HAUTEUR).value) Then
                faute = "pied non calcule"
            Else
                centre = CDbl(etat.Cells(r, PIL_COL_ROLL_HAUTEUR).value)
                pied = CDbl(etat.Cells(r, PIL_COL_FOOT_HEIGHT).value)
                ligne = centre - pied
                ' Groupe du rouleau dans PALIERS_HEATING (Y min / Y max, repere Creo).
                For g = 7 To derniereP
                    If Left$(PilTexte(pal.Cells(g, 8).value), 4) = "PAL_" And _
                        PilNumerique(pal.Cells(g, 3).value) And _
                        PilNumerique(pal.Cells(g, 4).value) And _
                        PilNumerique(pal.Cells(g, 7).value) Then
                        yMin = CDbl(pal.Cells(g, 3).value)
                        yMax = CDbl(pal.Cells(g, 4).value)
                        If centre >= yMin - 0.5 And centre <= yMax + 0.5 Then
                            palierAsm = CDbl(pal.Cells(g, 7).value) + decalage
                            If Abs(ligne - palierAsm) <= PIL_TOL Then
                                nbSurPalier = nbSurPalier + 1
                            Else
                                faute = "pied " & PilFmt(pied) & " mm jusqu'a Y " & _
                                    PilFmt(ligne) & " au lieu du palier Y " & PilFmt(palierAsm)
                            End If
                            Exit For
                        End If
                    End If
                Next g
            End If
            If faute <> "" Then
                nbFaux = nbFaux + 1
                If nbFaux <= 5 Then PilLigne liste, "- " & nomRouleau & " : " & faute
            End If
        End If
    Next r

    If nbLignes <> nbAttendus Then
        bilan = nbLignes & " rouleau(x) equipe(s) dans " & nomAssemblage & " pour " & _
            nbAttendus & " detecte(s) par les paliers : une piece a ete oubliee ou " & _
            "comptee deux fois. Comparer " & PIL_ETAT_TECH & " et " & PIL_FEUILLE_PALIERS & "."
        If nbFaux > 0 Then bilan = bilan & vbCrLf & liste
        Exit Function
    End If
    If nbFaux > 0 Then
        bilan = nbFaux & " rouleau(x) incorrect(s) sur " & nbLignes & " :" & vbCrLf & liste
        Exit Function
    End If
    bilan = nbLignes & " rouleau(x) equipe(s) d'un moteur " & cote & " ; " & _
        nbSurPalier & " pied(s) poses sur leur palier, les autres sur la charpente " & _
        "(groupes sans palier)."
    PilControlerPieds = True
End Function

' ---------------------------------------------------------------- Outils

Private Function PilFeuille() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PIL_FEUILLE)
    Err.Clear
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PIL_FEUILLE
        ws.Tab.Color = RGB(223, 113, 36)
        ws.Range("C7").value = "OUI"
    End If
    ws.Range("A1").value = "PILOTE HEATING (version " & PIL_VERSION & ")"
    ws.Range("A2").value = "1. PILOTE_Configurer (une fois)   2. PILOTE_1_PaliersEtIGES   " & _
        "3. Creo : recharger l'IGES dans le squelette   4. PILOTE_2_ControleEtPieds"
    ws.Range("B5").value = "Cote moteur (Moteur_L / Moteur_R)"
    ws.Range("B6").value = "Assemblage moteur-cylindre (chemin .asm)"
    ws.Range("B7").value = "Confirmer les paliers avant ecriture (OUI / NON)"
    ws.Range("B9").value = "Etat"
    ws.Range("B10").value = "IGES a recharger dans le squelette"
    ws.Range("C5:C7").Interior.Color = RGB(255, 240, 191)
    ws.Columns("B").ColumnWidth = 46
    ws.Columns("C").ColumnWidth = 90
    Set PilFeuille = ws
End Function

Private Function PilConfirmerPaliers() As Boolean
    PilConfirmerPaliers = (UCase$(Trim$(PilTexte(PilFeuille().Range("C7").value))) <> "NON")
End Function

Private Sub PilEtat(ByVal etat As String, ByVal fichier As String)
    Dim ws As Worksheet
    Set ws = PilFeuille()
    ws.Range("C9").value = etat
    ws.Range("C10").value = fichier
End Sub

Private Sub PilJournal(ByVal phase As String, ByVal etape As String, _
    ByVal resultat As String, ByVal detail As String)
    Dim ws As Worksheet
    Dim r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PIL_JOURNAL)
    Err.Clear
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PIL_JOURNAL
        ws.Tab.Color = RGB(223, 113, 36)
        ws.Range("A1:E1").value = Array("Date", "Phase", "Etape", "Resultat", "Detail")
        ws.Range("A1:E1").Font.Bold = True
        ws.Columns("A").ColumnWidth = 17
        ws.Columns("B:C").ColumnWidth = 30
        ws.Columns("D").ColumnWidth = 10
        ws.Columns("E").ColumnWidth = 120
    End If
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, 2).value = phase
    ws.Cells(r, 3).value = etape
    ws.Cells(r, 4).value = resultat
    ws.Cells(r, 5).value = Left$(detail, 32000)
    Select Case resultat
        Case "OK": ws.Cells(r, 4).Interior.Color = RGB(198, 239, 206)
        Case "KO", "ERREUR": ws.Cells(r, 4).Interior.Color = RGB(255, 199, 206)
        Case Else: ws.Cells(r, 4).Interior.Color = RGB(255, 235, 156)
    End Select
End Sub

Private Sub PilArret(ByVal phase As String, ByVal etape As String, ByVal texte As String)
    PilJournal phase, etape, "KO", Replace(texte, vbCrLf, " | ")
    PilEtat phase & " ARRETEE a l'etape '" & etape & "' (" & _
        Format$(Now, "yyyy-mm-dd hh:nn") & ") : voir " & PIL_JOURNAL, _
        PilTexte(PilFeuille().Range("C10").value)
    MsgBox phase & " arretee." & vbCrLf & "Etape : " & etape & vbCrLf & vbCrLf & _
        PilCouper(texte, 750) & vbCrLf & vbCrLf & "Detail : feuille " & PIL_JOURNAL & ".", _
        vbExclamation, "Pilote HEATING"
End Sub

Private Sub PilErreur(ByVal phase As String, ByVal etape As String, _
    ByVal numero As Long, ByVal description As String)
    Dim texte As String
    texte = "Erreur " & CStr(numero) & " : " & description
    On Error Resume Next
    PilJournal phase, etape, "ERREUR", texte
    PilEtat phase & " INTERROMPUE a l'etape '" & etape & "' : voir " & PIL_JOURNAL, _
        PilTexte(PilFeuille().Range("C10").value)
    On Error GoTo 0
    MsgBox phase & " interrompue." & vbCrLf & "Etape : " & etape & vbCrLf & texte, _
        vbCritical, "Pilote HEATING"
End Sub

Private Function PilNomAssemblageActif() As String
    Dim fabrique As pfcls.CCpfcAsyncConnection
    Dim connexion As pfcls.IpfcAsyncConnection
    Dim modele As pfcls.IpfcModel

    Set fabrique = New pfcls.CCpfcAsyncConnection
    On Error Resume Next
    Set connexion = fabrique.GetActiveConnection()
    If connexion Is Nothing Then
        Err.Clear
        Set connexion = fabrique.Connect("", "", "", 30)
    End If
    Err.Clear
    On Error GoTo 0
    If connexion Is Nothing Then Err.Raise vbObjectError + 8900, , "Creo indisponible."
    Set modele = connexion.Session.GetActiveModel()
    If modele Is Nothing Then Err.Raise vbObjectError + 8901, , "Aucun modele actif dans Creo."
    PilNomAssemblageActif = modele.Filename
End Function

' Nom sans numero de version Creo (heating.asm.12 -> heating.asm).
Private Function PilNomLogique(ByVal nom As String) As String
    Dim p As Long
    nom = LCase$(Trim$(nom))
    p = InStrRev(nom, ".asm")
    If p > 0 Then nom = Left$(nom, p + 3)
    PilNomLogique = nom
End Function

Private Sub PilLigne(ByRef texte As String, ByVal ligne As String)
    If Len(ligne) = 0 Then Exit Sub
    If texte <> "" Then texte = texte & vbCrLf
    texte = texte & ligne
End Sub

Private Function PilTexte(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    PilTexte = Trim$(CStr(v))
End Function

Private Function PilNumerique(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    PilNumerique = IsNumeric(v)
End Function

Private Function PilFmt(ByVal valeur As Double) As String
    PilFmt = Format$(valeur, "0.###")
    If Right$(PilFmt, 1) = "," Or Right$(PilFmt, 1) = "." Then _
        PilFmt = Left$(PilFmt, Len(PilFmt) - 1)
End Function

Private Function PilCouper(ByVal texte As String, ByVal maxi As Long) As String
    If Len(texte) <= maxi Then
        PilCouper = texte
    Else
        PilCouper = Left$(texte, maxi) & " ..."
    End If
End Function
