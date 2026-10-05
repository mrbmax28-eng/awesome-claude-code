Attribute VB_Name = "M_FRAME_HEATING"
Option Explicit

' Zone HEATING : niveaux propres, position X/Z liee a ZONES_FOUR si activee.
' Repere Creo : X longitudinal, Y vertical, Z transversal ; unites mm.
' Le composant squelette doit etre cree dans le sous-assemblage HEATING actif.
' 2026.10.05 (pack pilote) : mode silencieux (HEAT_DefinirSilencieux),
' derniere erreur lisible (HEAT_DerniereErreur) et export IGES sans
' fenetre (HEAT_ExporterIGESAuto). Calculs inchanges.

Private Const HEAT_ERR As Long = vbObjectError + 8100
Private Const HEAT_FIRST As Long = 7
Private Const HEAT_LAST As Long = 206
Private HeatLastOK As Boolean
Private HeatLastSkeletonOK As Boolean
Private HeatSilencieux As Boolean
Private HeatDerniereErreur As String

' Pilote : True = aucune fenetre dans HEAT_Recalculer ; l'erreur reste
' lisible par HEAT_DerniereErreur. Toujours remettre a False ensuite.
Public Sub HEAT_DefinirSilencieux(ByVal actif As Boolean)
    HeatSilencieux = actif
End Sub

Public Function HEAT_DerniereErreur() As String
    HEAT_DerniereErreur = HeatDerniereErreur
End Function

' Pilote : recalcul puis ecriture de AXES_HEATING_3D.igs, sans fenetre.
Public Function HEAT_ExporterIGESAuto(ByRef fichier As String, _
    ByRef message As String) As Boolean
    Dim dossier As String
    Dim ancien As Boolean
    On Error GoTo Echec
    message = ""
    fichier = ""
    ancien = HeatSilencieux
    HeatSilencieux = True
    HEAT_Recalculer
    HeatSilencieux = ancien
    If Not HeatLastOK Then
        message = "Recalcul HEATING : " & HeatDerniereErreur
        Exit Function
    End If
    If Len(ThisWorkbook.path) = 0 Then Err.Raise HEAT_ERR + 20, , "Enregistrer le classeur en .xlsm."
    dossier = ThisWorkbook.path & Application.PathSeparator & "Esquisses_HEATING"
    If Dir$(dossier, vbDirectory) = "" Then MkDir dossier
    fichier = dossier & Application.PathSeparator & "AXES_HEATING_3D.igs"
    HeatWriteIges fichier
    HEAT_ExporterIGESAuto = True
    Exit Function
Echec:
    HeatSilencieux = ancien
    message = "Export IGES HEATING : " & Err.description
End Function

Public Sub HEAT_Initialiser()
    Dim s As Worksheet, n As Worksheet, t As Worksheet, p As Worksheet, a As Worksheet
    Dim ys As Variant, i As Long
    Set s = HeatSheet("Structure_HEATING")
    Set a = HeatSheet("NiveauxY_HEATING")
    Set n = HeatSheet("Niveaux_HEATING")
    Set t = HeatSheet("Traverses_HEATING")
    Set p = HeatSheet("Points3D_HEATING")
    s.Tab.Color = RGB(223, 113, 36)
    a.Tab.Color = RGB(223, 113, 36)
    n.Tab.Color = RGB(223, 113, 36)
    t.Tab.Color = RGB(223, 113, 36)
    p.Tab.Color = RGB(223, 113, 36)

    If Len(CStr(s.Range("A2").value)) = 0 Then
        s.Range("A2").value = "HEATING - squelette independant (mm)"
        s.Range("A4:A13").value = Application.Transpose(Array( _
            "Zone", "Y pied poteaux", "Y sommet poteaux", "Longueur poteaux", _
            "Niveaux actifs", "Niveaux principaux", "Niveaux intermediaires", _
            "Traverses X et Z", "Y premiere traverse", "Y derniere traverse"))
        s.Range("B4").value = "HEATING"
        s.Range("F4:F15").value = Application.Transpose(Array( _
            "Travee X entre colonnes", "Travee Z entre colonnes", _
            "Cote de section (modifiable)", "X colonne cote -", "Z colonne cote -", _
            "Debord poutre X-", "Debord poutre X+", "Debord poutre Z-", _
            "Debord poutre Z+", "Y pied poteaux / EXIT", _
            "Y sommet poteaux / EXIT", "Ecart mini entre niveaux"))
        s.Range("G4:G15").value = Application.Transpose(Array( _
            7600, 9575, 300, -3800, -4787.5, 0, 0, 0, 0, -540, 25515, 210))
        s.Range("F18").value = "X colonne cote +"
        s.Range("F19").value = "Z colonne cote +"
        s.Range("G18").Formula = "=G7+G4"
        s.Range("G19").Formula = "=G8+G5"
        s.Range("A16").value = "Les Y se reglent dans NiveauxY_HEATING ; recalculer puis reexporter l'IGES."
        s.Range("A17").value = "La position X/Z est un point de depart centre ; ajuster G7/G8 dans l'assemblage reel."
        s.Range("A18").value = "Colonne CGL_FUS_HSG_CAG_D : detection et X/Z issus de l'assemblage HEATING."
        s.Range("G4:G15").Interior.Color = RGB(255, 240, 191)
        s.Columns("A").ColumnWidth = 38
        s.Columns("F").ColumnWidth = 34
        s.Columns("G").ColumnWidth = 18
    End If
    s.Range("F20:F27").value = Application.Transpose(Array( _
        "Double steering actif 1/0", "X colonne double cote -", _
        "Z colonnes doubles", "Y pied double (zero commun)", _
        "Y fin double (avant dernier niveau)", "X colonne double cote +", _
        "Z colonne double cote +", "Y traverse double"))
    s.Range("G20:G27").Interior.Color = RGB(222, 238, 247)
    If Len(CStr(a.Range("A6").value)) = 0 Then
        a.Range("A2").value = "HEATING - niveaux signes par rapport a EXIT = Y 0"
        a.Range("A3").value = "Poteaux de -540 a 25515 ; premier niveau a 0, ecarts 3780, 5 x 2880, puis 7335."
        a.Range("A4").value = "A ID ; B PRINCIPAL/INTERMEDIAIRE ; C Y signe en mm ; D 1 actif / 0 ignore."
        a.Range("A6:E6").value = Array("ID niveau", "Type", "Y / EXIT (mm)", "Actif 1/0", "Note")
        ys = Array(0, 3780, 6660, 9540, 12420, 15300, 18180, 25515)
        For i = 0 To UBound(ys)
            a.Cells(HEAT_FIRST + i, 1).value = "H_" & CStr(i)
            a.Cells(HEAT_FIRST + i, 2).value = "PRINCIPAL"
            a.Cells(HEAT_FIRST + i, 3).value = ys(i)
            a.Cells(HEAT_FIRST + i, 4).value = 1
            a.Cells(HEAT_FIRST + i, 5).value = IIf(i = 0, "Premier niveau / EXIT", _
                IIf(i = 1, "Ecart bas 3780", IIf(i = UBound(ys), "Ecart haut 7335", "Ecart 2880")))
        Next i
        a.Range("A7:E206").Interior.Color = RGB(255, 240, 191)
        a.Columns("A").ColumnWidth = 21
        a.Columns("B").ColumnWidth = 20
        a.Columns("C:D").ColumnWidth = 18
        a.Columns("E").ColumnWidth = 34
    End If
    If Len(CStr(n.Range("A6").value)) = 0 Then
        n.Range("A2").value = "HEATING - niveaux actifs tries selon Y"
        n.Range("A6:G6").value = Array("Ordre", "ID source", "Type", "Y / EXIT", _
            "Ecart depuis precedent", "Controle", "Ligne source")
        n.Columns("A:G").ColumnWidth = 21
    End If
    If Len(CStr(t.Range("A6").value)) = 0 Then
        t.Range("A2").value = "HEATING - quatre axes de traverse par niveau (deux X, deux Z)"
        t.Range("A6:M6").value = Array("ID", "Ordre", "Face", "X depart", _
            "Y / EXIT", "Z depart", "Z fin", "Longueur", "Cote", _
            "Ecart Y", "X fin", "Axe", "Niveau source")
        t.Columns("A:M").ColumnWidth = 17
    End If
    If Len(CStr(p.Range("A6").value)) = 0 Then
        p.Range("A2").value = "HEATING - points 3D dans le repere Creo"
        p.Range("A6:E6").value = Array("ID", "X", "Y", "Z", "Role")
        p.Columns("A:E").ColumnWidth = 21
    End If
End Sub

Public Sub HEAT_Recalculer()
    On Error GoTo Echec
    Dim s As Worksheet, a As Worksheet, n As Worksheet, t As Worksheet, p As Worksheet
    Dim v(4 To 15) As Double, i As Long, j As Long, r As Long, count As Long
    Dim activeCount As Long, principalCount As Long, intermCount As Long
    Dim levelY(1 To 80) As Double, levelID(1 To 80) As String
    Dim levelType(1 To 80) As String, sourceRow(1 To 80) As Long
    Dim swapY As Double, swapID As String, swapType As String, swapRow As Long
    Dim seen As Object, indicator As Variant, typeName As String, idName As String
    Dim bottomY As Double, topY As Double, minGap As Double, previousY As Double
    Dim xMin As Double, xMax As Double, zMin As Double, zMax As Double
    Dim beamRow As Long, pointRow As Long, lastExisting As Long
    Dim phase As String, errorNumber As Long, errorText As String
    Dim doubleActive As Double, doubleX As Double, doubleX2 As Double, doubleZ As Double
    Dim doubleStart As Double, doubleEnd As Double, cagX As Double, cagZ As Double
    Dim cagXMin As Double, cagXMax As Double
    Dim correctionX As Double, correctionZ As Double
    Dim cornerX As Boolean, cornerZ As Boolean

    HeatLastOK = False
    HeatDerniereErreur = ""
    phase = "initialisation des feuilles HEATING"
    HEAT_Initialiser
    Set s = ThisWorkbook.Worksheets("Structure_HEATING")
    Set a = ThisWorkbook.Worksheets("NiveauxY_HEATING")
    Set n = ThisWorkbook.Worksheets("Niveaux_HEATING")
    Set t = ThisWorkbook.Worksheets("Traverses_HEATING")
    Set p = ThisWorkbook.Worksheets("Points3D_HEATING")
    phase = "lecture des dimensions G4:G15"
    For i = 4 To 15
        If Not IsNumeric(s.Cells(i, 7).value) Then _
            Err.Raise HEAT_ERR + 1, , "Structure_HEATING!G" & i & " doit etre numerique."
        v(i) = CDbl(s.Cells(i, 7).value)
    Next i
    If v(4) <= 0 Or v(5) <= 0 Or v(6) <= 0 Then _
        Err.Raise HEAT_ERR + 2, , "Travees X/Z et cote de section doivent etre positifs."
    For i = 9 To 12
        If v(i) < 0 Then Err.Raise HEAT_ERR + 3, , "Debord G" & i & " negatif."
    Next i
    bottomY = v(13): topY = v(14): minGap = v(15)
    If bottomY >= topY Or minGap < 0 Then _
        Err.Raise HEAT_ERR + 4, , "Verifier pied G13, sommet G14 et ecart mini G15."
    xMin = v(7): xMax = xMin + v(4)
    zMin = v(8): zMax = zMin + v(5)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    phase = "lecture et tri des niveaux Y"
    For r = HEAT_FIRST To HEAT_LAST
        count = Application.WorksheetFunction.CountA(a.Range(a.Cells(r, 1), a.Cells(r, 5)))
        If count > 0 Then
            indicator = a.Cells(r, 4).value
            If Not IsNumeric(indicator) Then Err.Raise HEAT_ERR + 5, , _
                "NiveauxY_HEATING!D" & r & " : mettre 1 ou 0."
            If CDbl(indicator) <> 0 And CDbl(indicator) <> 1 Then Err.Raise HEAT_ERR + 6, , _
                "NiveauxY_HEATING!D" & r & " : mettre 1 ou 0."
            If CDbl(indicator) = 1 Then
                If activeCount = 80 Then Err.Raise HEAT_ERR + 7, , "80 niveaux maximum."
                If Not IsNumeric(a.Cells(r, 3).value) Then Err.Raise HEAT_ERR + 8, , _
                    "NiveauxY_HEATING!C" & r & " : Y numerique requis."
                idName = Trim$(CStr(a.Cells(r, 1).value))
                typeName = UCase$(Trim$(CStr(a.Cells(r, 2).value)))
                If Len(idName) = 0 Or (typeName <> "PRINCIPAL" And typeName <> "INTERMEDIAIRE") Then _
                    Err.Raise HEAT_ERR + 9, , "ID et type invalides, ligne " & r & "."
                If seen.Exists(idName) Then Err.Raise HEAT_ERR + 10, , "ID duplique : " & idName
                seen.Add idName, True
                activeCount = activeCount + 1
                levelY(activeCount) = CDbl(a.Cells(r, 3).value)
                levelID(activeCount) = idName
                levelType(activeCount) = typeName
                sourceRow(activeCount) = r
            End If
        End If
    Next r
    If activeCount = 0 Then Err.Raise HEAT_ERR + 11, , "Activer au moins un niveau dans NiveauxY_HEATING."
    For i = 2 To activeCount
        swapY = levelY(i): swapID = levelID(i)
        swapType = levelType(i): swapRow = sourceRow(i)
        j = i - 1
        Do While j >= 1
            If levelY(j) <= swapY Then Exit Do
            levelY(j + 1) = levelY(j)
            levelID(j + 1) = levelID(j)
            levelType(j + 1) = levelType(j)
            sourceRow(j + 1) = sourceRow(j)
            j = j - 1
        Loop
        levelY(j + 1) = swapY
        levelID(j + 1) = swapID
        levelType(j + 1) = swapType
        sourceRow(j + 1) = swapRow
    Next i
    If levelY(1) < bottomY Or levelY(activeCount) > topY Then _
        Err.Raise HEAT_ERR + 12, , "Un niveau actif est hors des poteaux (G13:G14)."
    For i = 2 To activeCount
        If levelY(i) - levelY(i - 1) < minGap - 0.000001 Then _
            Err.Raise HEAT_ERR + 13, , "Ecart insuffisant entre " & levelID(i - 1) & " et " & levelID(i) & "."
    Next i
    For i = 1 To activeCount
        If levelType(i) = "PRINCIPAL" Then
            principalCount = principalCount + 1
        Else
            intermCount = intermCount + 1
        End If
    Next i

    lastExisting = t.Cells(t.Rows.count, 1).End(xlUp).row
    If lastExisting >= HEAT_FIRST Then t.Range("A7:M" & lastExisting).ClearContents
    lastExisting = p.Cells(p.Rows.count, 1).End(xlUp).row
    If lastExisting >= HEAT_FIRST Then p.Range("A7:E" & lastExisting).ClearContents
    n.Range("A7:G86").ClearContents
    beamRow = HEAT_FIRST: pointRow = HEAT_FIRST
    phase = "construction des traverses et points 3D"
    HeatCorners p, pointRow, bottomY, xMin, xMax, zMin, zMax, "PIED_COLONNE"
    previousY = bottomY
    For i = 1 To activeCount
        r = HEAT_FIRST + i - 1
        n.Cells(r, 1).value = i
        n.Cells(r, 2).value = levelID(i)
        n.Cells(r, 3).value = levelType(i)
        n.Cells(r, 4).value = levelY(i)
        n.Cells(r, 5).value = levelY(i) - previousY
        n.Cells(r, 6).value = "OK"
        n.Cells(r, 7).value = sourceRow(i)
        HeatBeams t, beamRow, i, levelID(i), levelY(i), levelY(i) - previousY, _
            xMin, xMax, zMin, zMax, v(6), v(9), v(10), v(11), v(12)
        If levelY(i) > previousY Then _
            HeatCorners p, pointRow, levelY(i), xMin, xMax, zMin, zMax, "TRAVERSE"
        previousY = levelY(i)
    Next i
    If topY > previousY Then _
        HeatCorners p, pointRow, topY, xMin, xMax, zMin, zMax, "SOMMET_COLONNE"
    phase = "lecture de la detection CAG_D"
    With ThisWorkbook.Worksheets("ZONES_FOUR")
        If Not IsNumeric(.Range("AS8").value) Then Err.Raise HEAT_ERR + 60, , _
            "ZONES_FOUR!AS8 : detection double steering invalide."
        doubleActive = CDbl(.Range("AS8").value)
        s.Range("G20").value = doubleActive
        If doubleActive = 1 Then
            If activeCount < 2 Then Err.Raise HEAT_ERR + 61, , _
                "La colonne double requiert au moins deux niveaux HEATING actifs."
            phase = "validation de AT:AW8 et BI:BJ8"
            ' Ne pas regrouper ces appels dans une grande expression VBA :
            ' certaines versions d'Excel renvoient l'erreur 16 « Expression
            ' too complex » lorsque plusieurs proprietes Range sont combinees.
            If Not IsNumeric(.Range("AT8").value) Then _
                Err.Raise HEAT_ERR + 62, , "ZONES_FOUR!AT8 doit etre numerique."
            If Not IsNumeric(.Range("AU8").value) Then _
                Err.Raise HEAT_ERR + 62, , "ZONES_FOUR!AU8 doit etre numerique."
            If Not IsNumeric(.Range("AV8").value) Then _
                Err.Raise HEAT_ERR + 62, , "ZONES_FOUR!AV8 doit etre numerique."
            If Not IsNumeric(.Range("AW8").value) Then _
                Err.Raise HEAT_ERR + 62, , "ZONES_FOUR!AW8 doit etre numerique."
            If Not IsNumeric(.Range("BI8").value) Then _
                Err.Raise HEAT_ERR + 62, , "ZONES_FOUR!BI8 doit etre numerique. Relancer le releve Creo."
            If Not IsNumeric(.Range("BJ8").value) Then _
                Err.Raise HEAT_ERR + 62, , "ZONES_FOUR!BJ8 doit etre numerique. Relancer le releve Creo."
            ' Les deux poteaux reprennent les deux bords X physiques de CAG_D.
            ' La correction X est appliquee aux deux bords ; la correction Z
            ' s'applique au centre transversal de la piece.
            phase = "positionnement des deux colonnes sous CAG_D"
            correctionX = CDbl(.Range("AV8").value)
            correctionZ = CDbl(.Range("AW8").value)
            cagX = CDbl(.Range("AT8").value)
            cagZ = CDbl(.Range("AU8").value)
            cagXMin = CDbl(.Range("BI8").value)
            cagXMax = CDbl(.Range("BJ8").value)
            If cagXMin >= cagXMax Then _
                Err.Raise HEAT_ERR + 67, , "Les bornes X BI8/BJ8 de CAG_D doivent etre croissantes."
            doubleX = cagXMin + correctionX
            doubleX2 = cagXMax + correctionX
            doubleZ = cagZ + correctionZ
            doubleStart = bottomY
            doubleEnd = levelY(activeCount - 1)
            If doubleStart >= doubleEnd Then _
                Err.Raise HEAT_ERR + 63, , _
                    "Le pied de colonne G13 doit etre sous l'avant-dernier niveau HEATING."
            If doubleX < xMin Or doubleX > xMax Then _
                Err.Raise HEAT_ERR + 64, , _
                    "Le bord X cote - de CAG_D est hors travee HEATING : verifier " & _
                    "LONGUEURS_HEATING, G7/G4 et AV8."
            If doubleX2 < xMin Or doubleX2 > xMax Then _
                Err.Raise HEAT_ERR + 64, , _
                    "Le bord X cote + de CAG_D est hors travee HEATING : verifier " & _
                    "LONGUEURS_HEATING, G7/G4 et AV8."
            If doubleZ < zMin Or doubleZ > zMax Then _
                Err.Raise HEAT_ERR + 64, , _
                    "Le centre Z de CAG_D est hors travee HEATING : verifier " & _
                    "LONGUEURS_HEATING, G7/G4 et les corrections AV/AW8."
            cornerX = Abs(doubleX - xMin) < 0.001
            If Not cornerX Then cornerX = Abs(doubleX - xMax) < 0.001
            If Not cornerX Then cornerX = Abs(doubleX2 - xMin) < 0.001
            If Not cornerX Then cornerX = Abs(doubleX2 - xMax) < 0.001
            cornerZ = Abs(doubleZ - zMin) < 0.001
            If Not cornerZ Then cornerZ = Abs(doubleZ - zMax) < 0.001
            If cornerX And cornerZ Then _
                Err.Raise HEAT_ERR + 66, , _
                    "Une colonne double se superpose a un poteau d'angle : ajuster AV/AW8."
            s.Range("G21").Formula = "=ZONES_FOUR!BI8+ZONES_FOUR!AV8"
            s.Range("G22").Formula = "=ZONES_FOUR!AU8+ZONES_FOUR!AW8"
            s.Range("G23").Formula = "=G13"
            s.Range("G24").value = doubleEnd
            s.Range("G25").Formula = "=ZONES_FOUR!BJ8+ZONES_FOUR!AV8"
            s.Range("G26").Formula = "=G22"
            s.Range("G27").Formula = "=G24"
            .Range("BK8").Formula = "=Structure_HEATING!G27"
            phase = "ecriture des points et de la traverse CAG_D"
            p.Cells(pointRow, 1).Resize(1, 5).value = Array("DOUBLE_A_PIED", _
                doubleX, doubleStart, doubleZ, "DOUBLE_STEERING")
            pointRow = pointRow + 1
            p.Cells(pointRow, 1).Resize(1, 5).value = Array("DOUBLE_A_FIN", _
                doubleX, doubleEnd, doubleZ, "DOUBLE_STEERING")
            pointRow = pointRow + 1
            p.Cells(pointRow, 1).Resize(1, 5).value = Array("DOUBLE_B_PIED", _
                doubleX2, doubleStart, doubleZ, "DOUBLE_STEERING")
            pointRow = pointRow + 1
            p.Cells(pointRow, 1).Resize(1, 5).value = Array("DOUBLE_B_FIN", _
                doubleX2, doubleEnd, doubleZ, "DOUBLE_STEERING")
            pointRow = pointRow + 1
            p.Cells(pointRow, 1).Resize(1, 5).value = Array("DOUBLE_TRAVERSE_X-", _
                xMin, doubleEnd, doubleZ, "DOUBLE_TRAVERSE")
            pointRow = pointRow + 1
            p.Cells(pointRow, 1).Resize(1, 5).value = Array("DOUBLE_TRAVERSE_X+", _
                xMax, doubleEnd, doubleZ, "DOUBLE_TRAVERSE")
            HeatBeamRow t, beamRow, activeCount, "DOUBLE_TRAVERSE", doubleEnd, 0#, _
                v(6), "DOUBLE_TRAVERSE", xMin, xMax, doubleZ, doubleZ, "X"
        ElseIf doubleActive = 0 Then
            s.Range("G21:G27").ClearContents
            .Range("BK8").ClearContents
        Else
            Err.Raise HEAT_ERR + 65, , "ZONES_FOUR!AS8 doit valoir 0 ou 1."
        End If
    End With
    s.Range("B5").value = bottomY
    s.Range("B6").value = topY
    s.Range("B7").value = topY - bottomY
    s.Range("B8").value = activeCount
    s.Range("B9").value = principalCount
    s.Range("B10").value = intermCount
    s.Range("B11").value = beamRow - HEAT_FIRST
    s.Range("B12").value = levelY(1)
    s.Range("B13").value = levelY(activeCount)
    phase = "recalcul Excel des feuilles liees"
    Application.Calculate
    HeatLastOK = True
    If Not HeatSilencieux Then MsgBox "HEATING : " & activeCount & " niveaux, " & (beamRow - HEAT_FIRST) & _
        " traverses ; poteaux Y=" & HeatDecimal(bottomY) & " a " & HeatDecimal(topY) & _
        " mm. Position X/Z : " & HeatDecimal(xMin) & " / " & HeatDecimal(zMin) & ".", _
        vbInformation, "HEATING"
    Exit Sub
Echec:
    errorNumber = Err.number: errorText = Err.description
    HeatDerniereErreur = phase & " (erreur " & errorNumber & ") : " & errorText
    On Error Resume Next
    ThisWorkbook.Worksheets("ZONES_FOUR").Range("AH8").value = _
        "ERREUR : HEATING " & phase & " (" & errorNumber & ") : " & errorText
    ThisWorkbook.Worksheets("ZONES_FOUR").Range("AJ8").value = 0
    On Error GoTo 0
    If Not HeatSilencieux Then MsgBox "HEATING non calcule (" & phase & ", erreur " & _
        errorNumber & ") : " & errorText, vbExclamation, "HEATING"
End Sub

' Permet au pilote multi-zones de stopper un export si le recalcul echoue.
Public Function HEAT_CalculValide() As Boolean
    HEAT_CalculValide = HeatLastOK
End Function

Public Function HEAT_CreationValide() As Boolean
    HEAT_CreationValide = HeatLastSkeletonOK
End Function

Public Sub HEAT_ExporterAxesIGES()
    On Error GoTo Echec
    Dim dossier As String, fichier As String
    HEAT_Recalculer
    If Not HeatLastOK Then Exit Sub
    If Len(ThisWorkbook.path) = 0 Then Err.Raise HEAT_ERR + 20, , "Enregistrer le classeur en .xlsm."
    dossier = ThisWorkbook.path & Application.PathSeparator & "Esquisses_HEATING"
    If Dir$(dossier, vbDirectory) = "" Then MkDir dossier
    fichier = dossier & Application.PathSeparator & "AXES_HEATING_3D.igs"
    HeatWriteIges fichier
    MsgBox "IGES HEATING exporte : " & fichier & vbCrLf & _
        "Pour un squelette existant, recharger sa fonction importee dans Creo.", vbInformation, "HEATING"
    Exit Sub
Echec:
    MsgBox "Export HEATING arrete : " & Err.description, vbExclamation, "HEATING"
End Sub

' Dans l'assemblage actif, un seul squelette est autorise par cette macro.
' Pour deux zones independantes, activer un sous-assemblage HEATING sans squelette.
Public Sub HEAT_CreerSqueletteCreo()
    On Error GoTo Echec
    Dim connectionFactory As pfcls.CCpfcAsyncConnection
    Dim connection As pfcls.IpfcAsyncConnection
    Dim session As pfcls.IpfcBaseSession, activeModel As pfcls.IpfcModel
    Dim assembly As pfcls.IpfcAssembly, assemblySolid As pfcls.IpfcSolid
    Dim skeleton As pfcls.IpfcSolid, skeletonModel As pfcls.IpfcModel
    Dim imported As pfcls.IpfcModel, sourcePart As pfcls.IpfcSolid
    Dim regenFactory As pfcls.CCpfcRegenInstructions, regen As pfcls.IpfcRegenInstructions
    Dim dossier As String, fichier As String, name As String, tempName As String
    Dim stepName As String, errorText As String, rollback As Boolean
    Dim readError As Long, readText As String
    HeatLastSkeletonOK = False
    stepName = "Calcul HEATING"
    HEAT_Recalculer
    If Not HeatLastOK Then Exit Sub
    If Len(ThisWorkbook.path) = 0 Then Err.Raise HEAT_ERR + 30, , "Enregistrer le classeur en .xlsm."
    dossier = ThisWorkbook.path & Application.PathSeparator & "Esquisses_HEATING"
    If Dir$(dossier, vbDirectory) = "" Then MkDir dossier
    fichier = dossier & Application.PathSeparator & "AXES_HEATING_3D.igs"
    stepName = "Export IGES"
    HeatWriteIges fichier
    stepName = "Connexion Creo"
    Set connectionFactory = New pfcls.CCpfcAsyncConnection
    On Error Resume Next
    Set connection = connectionFactory.GetActiveConnection()
    If connection Is Nothing Then
        Err.Clear
        Set connection = connectionFactory.Connect("", "", "", 30)
    End If
    Err.Clear
    On Error GoTo Echec
    If connection Is Nothing Then Err.Raise HEAT_ERR + 31, , "Creo indisponible ; verifier pfcls et la connexion."
    Set session = connection.session
    Set activeModel = session.GetActiveModel()
    If activeModel Is Nothing Then Err.Raise HEAT_ERR + 32, , "Activer l'assemblage HEATING dans Creo."
    If activeModel.Type <> pfcls.EpfcMDL_ASSEMBLY Then _
        Err.Raise HEAT_ERR + 33, , "Le modele actif doit etre un assemblage .ASM."
    Set assembly = activeModel
    Set assemblySolid = activeModel
    stepName = "Verification du squelette"
    On Error Resume Next
    Set skeleton = assembly.GetSkeleton()
    readError = Err.number: readText = Err.description
    Err.Clear
    On Error GoTo Echec
    If readError <> 0 And InStr(1, readText, "XToolkitNotFound", vbTextCompare) = 0 Then _
        Err.Raise HEAT_ERR + 34, , "Squelette non lisible : " & readText
    If Not skeleton Is Nothing Then
        HeatLastSkeletonOK = True
        MsgBox "Cet assemblage possede deja un squelette. Il est inchange. " & _
            "Activer le sous-assemblage HEATING sans squelette, ou recharger " & _
            "l'IGES dans une fonction importee existante : " & fichier, vbInformation, "HEATING"
        Exit Sub
    End If
    stepName = "Import IGES"
    name = HeatFreeName(session, "SK_HEATING_" & HeatCleanName(activeModel.filename))
    tempName = HeatFreeName(session, "HI_" & Format$(Now, "yymmddhhnnss"))
    Set imported = session.ImportNewModel(fichier, pfcls.EpfcIMPORT_NEW_IGES, _
        pfcls.EpfcMDL_PART, tempName, Nothing)
    If imported Is Nothing Then Err.Raise HEAT_ERR + 35, , "Import IGES sans modele."
    If imported.Type <> pfcls.EpfcMDL_PART Then Err.Raise HEAT_ERR + 36, , "IGES non importe en piece."
    Set sourcePart = imported
    stepName = "Creation squelette"
    assembly.AssembleSkeletonByCopy name, sourcePart
    rollback = True
    Set skeleton = assembly.GetSkeleton()
    If skeleton Is Nothing Then Err.Raise HEAT_ERR + 37, , "Squelette introuvable apres copie."
    If Not skeleton.IsSkeleton Then Err.Raise HEAT_ERR + 38, , "Composant non reconnu comme squelette."
    Set skeletonModel = skeleton
    rollback = False
    Set regenFactory = New pfcls.CCpfcRegenInstructions
    Set regen = regenFactory.Create(False, False, Nothing)
    stepName = "Regeneration"
    skeleton.Regenerate regen
    assemblySolid.Regenerate regen
    skeletonModel.Save
    activeModel.Save
    HeatLastSkeletonOK = True
    MsgBox "Squelette HEATING cree dans " & activeModel.filename & ". " & _
        "Verifier les courbes et les unites mm. Source : " & fichier, vbInformation, "HEATING"
    Exit Sub
Echec:
    errorText = Err.description
    On Error Resume Next
    If rollback Then assembly.DeleteSkeleton
    MsgBox "HEATING non termine (" & stepName & ") : " & errorText & vbCrLf & _
        "IGES conserve : " & fichier, vbExclamation, "HEATING"
End Sub

Private Function HeatSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set HeatSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If HeatSheet Is Nothing Then
        Set HeatSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        HeatSheet.name = sheetName
    End If
End Function

Private Sub HeatBeams(ByVal ws As Worksheet, ByRef outRow As Long, _
    ByVal order As Long, ByVal id As String, ByVal y As Double, ByVal gap As Double, _
    ByVal xMin As Double, ByVal xMax As Double, _
    ByVal zMin As Double, ByVal zMax As Double, ByVal side As Double, _
    ByVal overXm As Double, ByVal overXp As Double, _
    ByVal overZm As Double, ByVal overZp As Double)
    Dim face As Long, x As Double, z As Double
    For face = 0 To 1
        x = IIf(face = 0, xMin, xMax)
        HeatBeamRow ws, outRow, order, id, y, gap, side, _
            IIf(face = 0, "X_MIN", "X_MAX"), x, x, zMin - overZm, zMax + overZp, "Z"
    Next face
    For face = 0 To 1
        z = IIf(face = 0, zMin, zMax)
        HeatBeamRow ws, outRow, order, id, y, gap, side, _
            IIf(face = 0, "Z_MIN", "Z_MAX"), xMin - overXm, xMax + overXp, z, z, "X"
    Next face
End Sub

Private Sub HeatBeamRow(ByVal ws As Worksheet, ByRef outRow As Long, _
    ByVal order As Long, ByVal id As String, ByVal y As Double, _
    ByVal gap As Double, ByVal side As Double, ByVal face As String, _
    ByVal x1 As Double, ByVal x2 As Double, ByVal z1 As Double, _
    ByVal z2 As Double, ByVal axis As String)
    ws.Cells(outRow, 1).value = "HTR_" & Format$(outRow - HEAT_FIRST + 1, "000")
    ws.Cells(outRow, 2).value = order
    ws.Cells(outRow, 3).value = face
    ws.Cells(outRow, 4).value = x1
    ws.Cells(outRow, 5).value = y
    ws.Cells(outRow, 6).value = z1
    ws.Cells(outRow, 7).value = z2
    ws.Cells(outRow, 8).value = IIf(axis = "X", x2 - x1, z2 - z1)
    ws.Cells(outRow, 9).value = side
    ws.Cells(outRow, 10).value = gap
    ws.Cells(outRow, 11).value = x2
    ws.Cells(outRow, 12).value = axis
    ws.Cells(outRow, 13).value = id
    outRow = outRow + 1
End Sub

Private Sub HeatCorners(ByVal ws As Worksheet, ByRef outRow As Long, _
    ByVal y As Double, ByVal xMin As Double, ByVal xMax As Double, _
    ByVal zMin As Double, ByVal zMax As Double, ByVal role As String)
    Dim ix As Long, iz As Long
    For ix = 0 To 1
        For iz = 0 To 1
            ws.Cells(outRow, 1).value = "HN_" & Format$(outRow - HEAT_FIRST + 1, "000")
            ws.Cells(outRow, 2).value = IIf(ix = 0, xMin, xMax)
            ws.Cells(outRow, 3).value = y
            ws.Cells(outRow, 4).value = IIf(iz = 0, zMin, zMax)
            ws.Cells(outRow, 5).value = role
            outRow = outRow + 1
        Next iz
    Next ix
End Sub

' IGES ASCII type 110 : 4 axes de colonnes et 4 axes par niveau.
Private Sub HeatWriteIges(ByVal filename As String)
    Dim s As Worksheet, t As Worksheet, f As Integer, i As Long, r As Long, iz As Long
    Dim lastRow As Long, nSegments As Long, nGlobal As Long, dimensionMax As Double
    Dim xMin As Double, xMax As Double, zMin As Double, zMax As Double
    Dim bottomY As Double, topY As Double, x As Double, z As Double
    Dim header As String, stamp As String, opened As Boolean, detail As String
    Set s = ThisWorkbook.Worksheets("Structure_HEATING")
    Set t = ThisWorkbook.Worksheets("Traverses_HEATING")
    xMin = CDbl(s.Range("G7").value)
    xMax = xMin + CDbl(s.Range("G4").value)
    zMin = CDbl(s.Range("G8").value)
    zMax = zMin + CDbl(s.Range("G5").value)
    bottomY = CDbl(s.Range("B5").value)
    topY = CDbl(s.Range("B6").value)
    lastRow = t.Cells(t.Rows.count, 1).End(xlUp).row
    If lastRow < HEAT_FIRST Or topY <= bottomY Then _
        Err.Raise HEAT_ERR + 40, , "HEATING non calcule ou sans traverses."
    nSegments = 4 + lastRow - HEAT_FIRST + 1
    If CDbl(s.Range("G20").value) = 1 Then nSegments = nSegments + 2
    dimensionMax = Application.WorksheetFunction.Max( _
        Abs(xMin - CDbl(s.Range("G9").value)), _
        Abs(xMax + CDbl(s.Range("G10").value)), _
        Abs(zMin - CDbl(s.Range("G11").value)), _
        Abs(zMax + CDbl(s.Range("G12").value)), Abs(bottomY), Abs(topY))
    stamp = Format$(Now, "yyyymmdd.hhnnss")
    header = ",," & HeatHollerith("FRAME HEATING") & "," & _
        HeatHollerith("HEATING_AXES.IGS") & "," & HeatHollerith("FRAME HEATING") & _
        "," & HeatHollerith("FRAME HEATING") & _
        ",32,308,15,308,15,,1.,2," & HeatHollerith("MM") & ",1,0.01," & _
        HeatHollerith(stamp) & ",1E-07," & HeatDecimal(dimensionMax) & _
        "," & HeatHollerith("HEATING") & ",,11,0," & HeatHollerith(stamp) & ",;"
    nGlobal = HeatCeil(Len(header), 72)
    f = FreeFile
    On Error GoTo FermetureErreur
    Open filename For Output As #f
    opened = True
    Print #f, Space$(72) & "S0000001"
    For i = 1 To nGlobal
        HeatIgesLine f, Mid$(header, (i - 1) * 72 + 1, 72), "G", i
    Next i
    For i = 1 To nSegments
        HeatIgesDirectory f, i
    Next i
    i = 0
    For r = 0 To 1
        x = IIf(r = 0, xMin, xMax)
        For iz = 0 To 1
            z = IIf(iz = 0, zMin, zMax)
            i = i + 1
            HeatIgesSegment f, i, x, bottomY, z, x, topY, z
        Next iz
    Next r
    If CDbl(s.Range("G20").value) = 1 Then
        i = i + 1
        HeatIgesSegment f, i, CDbl(s.Range("G21").value), CDbl(s.Range("G23").value), _
            CDbl(s.Range("G22").value), CDbl(s.Range("G21").value), _
            CDbl(s.Range("G24").value), CDbl(s.Range("G22").value)
        i = i + 1
        HeatIgesSegment f, i, CDbl(s.Range("G25").value), CDbl(s.Range("G23").value), _
            CDbl(s.Range("G26").value), CDbl(s.Range("G25").value), _
            CDbl(s.Range("G24").value), CDbl(s.Range("G26").value)
    End If
    For r = HEAT_FIRST To lastRow
        i = i + 1
        HeatIgesSegment f, i, CDbl(t.Cells(r, 4).value), CDbl(t.Cells(r, 5).value), _
            CDbl(t.Cells(r, 6).value), CDbl(t.Cells(r, 11).value), _
            CDbl(t.Cells(r, 5).value), CDbl(t.Cells(r, 7).value)
    Next r
    HeatIgesLine f, "S" & Right$(Space$(7) & "1", 7) & _
        "G" & Right$(Space$(7) & nGlobal, 7) & _
        "D" & Right$(Space$(7) & 2 * nSegments, 7) & _
        "P" & Right$(Space$(7) & nSegments, 7), "T", 1
    Close #f
    Exit Sub
FermetureErreur:
    detail = Err.description
    On Error Resume Next
    If opened Then Close #f
    Err.Raise HEAT_ERR + 41, , "Ecriture IGES HEATING : " & detail
End Sub

Private Function HeatHollerith(ByVal value As String) As String
    HeatHollerith = CStr(Len(value)) & "H" & value
End Function

Private Function HeatDecimal(ByVal value As Double) As String
    HeatDecimal = Replace$(Format$(value, "0.########"), ",", ".")
End Function

Private Function HeatCeil(ByVal value As Double, ByVal stepSize As Double) As Long
    HeatCeil = CLng(-Int(-value / stepSize + 0.000000001))
End Function

Private Sub HeatIgesLine(ByVal f As Integer, ByVal payload As String, _
    ByVal section As String, ByVal index As Long)
    If Len(payload) > 72 Then Err.Raise HEAT_ERR + 42, , "Ligne IGES trop longue."
    Print #f, Left$(payload & Space$(72), 72) & section & Format$(index, "0000000")
End Sub

Private Function HeatIgesField(ByVal n As Long) As String
    HeatIgesField = Right$(Space$(8) & CStr(n), 8)
End Function

Private Sub HeatIgesDirectory(ByVal f As Integer, ByVal n As Long)
    Dim first As String, second As String
    first = HeatIgesField(110) & HeatIgesField(n) & _
        HeatIgesField(0) & HeatIgesField(0) & HeatIgesField(0) & _
        HeatIgesField(0) & HeatIgesField(0) & HeatIgesField(0) & "00000000"
    second = HeatIgesField(110) & HeatIgesField(0) & HeatIgesField(0) & _
        HeatIgesField(1) & HeatIgesField(0) & Space$(24) & HeatIgesField(0)
    HeatIgesLine f, first, "D", 2 * n - 1
    HeatIgesLine f, second, "D", 2 * n
End Sub

Private Sub HeatIgesSegment(ByVal f As Integer, ByVal n As Long, _
    ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, _
    ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double)
    Dim payload As String
    payload = "110," & HeatDecimal(x1) & "," & HeatDecimal(y1) & "," & _
        HeatDecimal(z1) & "," & HeatDecimal(x2) & "," & _
        HeatDecimal(y2) & "," & HeatDecimal(z2) & ";"
    If Len(payload) > 64 Then Err.Raise HEAT_ERR + 43, , "Segment IGES trop long : " & n
    HeatIgesLine f, Left$(payload & Space$(64), 64) & _
        HeatIgesField(2 * n - 1), "P", n
End Sub

Private Function HeatCleanName(ByVal filename As String) As String
    Dim base As String, i As Long, c As String, cleaned As String
    base = UCase$(Trim$(filename))
    i = InStr(1, base, ".ASM", vbTextCompare)
    If i > 0 Then base = Left$(base, i - 1)
    For i = 1 To Len(base)
        c = Mid$(base, i, 1)
        If c Like "[A-Z0-9_]" Then cleaned = cleaned & c Else cleaned = cleaned & "_"
    Next i
    HeatCleanName = Left$(cleaned, 17)
End Function

Private Function HeatFreeName(ByVal session As pfcls.IpfcBaseSession, _
    ByVal preferred As String) As String
    Dim index As Long, candidate As String, model As pfcls.IpfcModel
    Dim errNumber As Long, errText As String, folder As String
    folder = session.GetCurrentDirectory()
    For index = 0 To 999
        If index = 0 Then
            candidate = preferred
        Else
            candidate = Left$(preferred, 26) & "_" & Format$(index, "000")
        End If
        Set model = Nothing
        On Error Resume Next
        Set model = session.GetModel(candidate, pfcls.EpfcMDL_PART)
        errNumber = Err.number: errText = Err.description
        Err.Clear
        On Error GoTo 0
        If errNumber <> 0 And InStr(1, errText, "XToolkitNotFound", vbTextCompare) = 0 Then _
            Err.Raise HEAT_ERR + 44, , "Recherche du modele " & candidate & " : " & errText
        If model Is Nothing Then
            If Not HeatPartExists(folder, candidate) And _
                Not HeatPartExists(ThisWorkbook.path, candidate) Then
                HeatFreeName = candidate
                Exit Function
            End If
        End If
    Next index
    Err.Raise HEAT_ERR + 45, , "Aucun nom de piece libre pour " & preferred
End Function

Private Function HeatPartExists(ByVal folder As String, ByVal name As String) As Boolean
    If Len(folder) = 0 Then Exit Function
    HeatPartExists = Len(Dir$(folder & Application.PathSeparator & name & ".prt*")) > 0
End Function
