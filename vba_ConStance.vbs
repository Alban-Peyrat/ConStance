Option Explicit

Dim CSX As String
Dim colIndex As String
Dim rowIndex As Integer
Dim inputPPN As String
Dim folderPath As String
Dim fileName As String
Dim mainWorkBook As Workbook
Sub read_Sudoc_Data()
    'Pour tout ce qui touche au XML en VB et au Sudoc MarcXML
    'http://documentation.abes.fr/sudoc/manuels/administration/aidewebservices/#SudocMarcXML
    'https://excel-macro.tutorialhorizon.com/vba-excel-read-data-from-xml-file/
    'https://www.mrexcel.com/board/threads/reading-xml-into-excel-with-vba.822719/
    'https://software-solutions-online.com/excel-vba-get-data-from-web-using-msxml/
    'https://stackoverflow.com/questions/19194544/selectsinglenode-using-vbscript#answer-19195587
    Dim nbPPN As Integer, count As Integer, jj As Integer
    Dim oXMLFile As Object, XMLFileName As String
           
    mainWorkBook.Worksheets("Introduction").Activate
    nbPPN = Application.WorksheetFunction.CountA(Range("I:I"))
    
    count = 0
    For jj = 2 To nbPPN
        If Cells(jj + count, 9).Value <> "" Then
            'Rajoute les 0 devant le PPN si nécessaire
            inputPPN = Right(Cells(jj + count, 9).Value, 9)
            While Len(inputPPN) < 9
                inputPPN = "0" & inputPPN
            Wend
                
            'Crée l'URL et la charge
            Dim URL As String
            Select Case CSX
                Case "[CS2]", "[CS3]", "[CS4]", "[CS5]", "[CS6]", "[CS7]"
                    URL = "https://www.sudoc.fr/" & inputPPN & ".xml"
                Case "[CS1]"
                    URL = "https://www.idref.fr/" & inputPPN & ".xml"
                Case Else
                    MsgBox "La valeur entrée en H2 ne devrait pas être possible"
            End Select
                
            Set oXMLFile = CreateObject("Microsoft.XMLDOM")
            XMLFileName = URL
            oXMLFile.async = False
            oXMLFile.Load (XMLFileName)
    
            'Lance le bon script
            Select Case CSX
                Case "[CS2]"
                    ctrlUB700S3 oXMLFile, mainWorkBook, jj
                Case "[CS1]"
                    ctrlUA103eqUA200f oXMLFile, mainWorkBook, jj
                Case "[CS3]"
                    ctrlUB7XXS3 oXMLFile, mainWorkBook, jj
                Case "[CS4]"
                    getUBAge oXMLFile, mainWorkBook, jj
                Case "[CS5]"
                    getUBAge100 oXMLFile, mainWorkBook, jj
                Case "[CS6]"
                    getEdition oXMLFile, mainWorkBook, jj
                Case "[CS7]"
                    getUBAgeDD oXMLFile, mainWorkBook, jj
                Case Else
                    MsgBox "La valeur entrée en H2 ne devrait pas être possible"
            End Select
            
        Else
            'Gère la boucle en cas de cellule vide
            count = count + 1
            jj = jj - 1
        End If
    Next
    
    'Formattage cellules
    With mainWorkBook.Sheets("Résultats").Range("A2:" & colIndex & (jj + count - 1))
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        .Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    mainWorkBook.Worksheets("Résultats").Activate
    Range("A:" & colIndex).RemoveDuplicates Columns:=Array(1), Header:=xlYes
    
    rowIndex = jj - 1
    
End Sub
Sub formatEnTetes()
    'https://www.automateexcel.com/vba/format-cells/
    'Crée les en-têtes pour la feuille "Résultats"
    
    mainWorkBook.Worksheets("Résultats").Activate
    Select Case CSX
        Case "[CS2]"
            Range("A1").Value = "PPN"
            Range("B1").Value = "UB700[0]"
            Range("C1").Value = "Résultat"
            colIndex = "C"
        Case "[CS1]"
            Range("A1").Value = "PPN"
            Range("B1").Value = "UA103a"
            Range("C1").Value = "UA200f1"
            Range("D1").Value = "Note birth date"
            Range("E1").Value = "UA103b"
            Range("F1").Value = "UA200f2"
            Range("G1").Value = "Note death date"
            Range("H1").Value = "Résultat"
            Range("I1").Value = "Note"
            colIndex = "I"
        Case "[CS3]"
            Range("A1").Value = "PPN"
            Range("B1").Value = "UB7XX[0]"
            Range("C1").Value = "Résultat"
            colIndex = "C"
        Case "[CS4]", "[CS5]"
            Range("A1").Value = "PPN"
            Range("B1").Value = "Année (la plus élevée)"
            Range("D1").Value = "Année moyenne"
            Range("E1").Value = "Année médianne"
            Range("F1").Value = "Nb titres exclus"
            colIndex = "B"
        Case "[CS6]"
            Range("A1").Value = "PPN"
            Range("B1").Value = "Édition"
            Range("C1").Value = "Titre"
            Range("D1").Value = "Clef de titre"
            Range("E1").Value = "Cat. clef titre"
            Range("F1").Value = "PPN ment. resp."
            Range("G1").Value = "Résultats"
            colIndex = "G"
        Case "[CS7]"
            Range("A1").Value = "PPN"
            Range("B1").Value = "Année 21X (la plus élevée)"
            Range("C1").Value = "Année 100 (la plus élevée)"
            Range("E1").Value = "Année moyenne 21X"
            Range("F1").Value = "Année médianne 21X"
            Range("G1").Value = "Nb titres exclus 21X"
            Range("E3").Value = "Année moyenne 100"
            Range("F3").Value = "Année médianne 100"
            Range("G3").Value = "Nb titres exclus 100"
            colIndex = "C"
        Case Else
            MsgBox "La valeur entrée en H2 ne devrait pas être possible"
    End Select
    With mainWorkBook.Worksheets("Résultats").Range("A1:" & colIndex & "1")
        .Interior.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(255, 255, 255)
    End With
    
    'Pour éviter que les PPN deviennent des nombres
    Range("A:" & colIndex).NumberFormat = "@"
    
End Sub
Sub ctrlUB700S3(oXMLFile, mainWorkBook, jj)
    
    Dim dollar As String, output As String, PPN, PPNval As String
    Dim UB700Node

    'Récupère le PPN et sort de la fonction s'il y a une erreur liée au PPN
    Set PPN = oXMLFile.SelectSingleNode("/record/controlfield[@tag='001']/text()")
    If Not PPN Is Nothing Then
        PPNval = PPN.NodeValue
    Else
        PPNval = "ERREUR [entrée n°" & jj & " : " & Chr(10) & inputPPN & "]"
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Interior.Color = RGB(0, 0, 0)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Font.Color = RGB(255, 255, 255)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Value = PPNval
        Exit Sub
    End If
    
    Set UB700Node = oXMLFile.SelectSingleNode("/record/datafield[@tag='700']/subfield[0]")
    If Not UB700Node Is Nothing Then
        dollar = UB700Node.getAttribute("code")
    Else
        dollar = "PAS DE 700"
    End If
    output = ""
    
    If dollar = "3" Then
        output = "OK"
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Interior.Color = RGB(146, 208, 80)
    Else
        output = "Problème"
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Interior.Color = RGB(255, 0, 0)
    End If
        
    mainWorkBook.Sheets("Résultats").Range("B" & jj).Value = dollar
    mainWorkBook.Sheets("Résultats").Range("A" & jj).Value = PPNval
    mainWorkBook.Sheets("Résultats").Range("C" & jj).Value = output
    
    'Unload oXMLFile(XMLFileName)
    
End Sub
Sub ctrlUB7XXS3(oXMLFile, mainWorkBook, jj)
    
    Dim dollar As String, output As String, PPN, PPNval As String, mismatchCount As Integer
    Dim UB7XXNodes, kk As Integer, parentNode

    'Récupère le PPN et sort de la fonction s'il y a une erreur liée au PPN
    Set PPN = oXMLFile.SelectSingleNode("/record/controlfield[@tag='001']/text()")
    If Not PPN Is Nothing Then
        PPNval = PPN.NodeValue
    Else
        PPNval = "ERREUR [entrée n°" & jj & " : " & Chr(10) & inputPPN & "]"
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Interior.Color = RGB(0, 0, 0)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Font.Color = RGB(255, 255, 255)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Value = PPNval
        Exit Sub
    End If
    
    mismatchCount = 0
    
    Set UB7XXNodes = oXMLFile.SelectNodes("/record/datafield[@tag>=700 and @tag<800]/subfield[0]")
    
    If Not UB7XXNodes Is Nothing Then
        For kk = 0 To (UB7XXNodes.Length - 1)
            Set parentNode = UB7XXNodes(kk).parentNode
            dollar = appendNote(dollar, "[" & kk & "] " & parentNode.getAttribute("tag") & " 1er $ : " & UB7XXNodes(kk).getAttribute("code"))
            If Right(dollar, 1) <> "3" Then
                If InStr(output, "Problème") = 0 Then
                    output = "Problème"
                End If
                output = appendNote(output, "occ. [" & kk & "]")
                mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Interior.Color = RGB(255, 0, 0)
            End If
        Next
    Else
        dollar = "PAS DE 7XX"
    End If
    
    If InStr(output, "Problème") = 0 Then
        output = "OK"
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Interior.Color = RGB(146, 208, 80)
    'Permet de dire quels champs sont pbatiques
    Else
    
    End If
        
    mainWorkBook.Sheets("Résultats").Range("B" & jj).Value = dollar
    mainWorkBook.Sheets("Résultats").Range("A" & jj).Value = PPNval
    mainWorkBook.Sheets("Résultats").Range("C" & jj).Value = output
    
    'Unload oXMLFile(XMLFileName)
    
End Sub
Sub ctrlUA103eqUA200f(oXMLFile, mainWorkBook, jj)
    'ATTENTION ELLE NE PREND QUE LE 1ER UA200 S'IL Y EN A PLUSIEURS
    'à terme je rajouterai date de mort je pense
    'et un meilleur handling de
    Dim UA103a As String, UA103b As String, UA200f As String, UA200f1 As String, UA200f2 As String, moveRight As Integer, UA103aNode, UA103bNode, UA200fNode
    Dim result As String, note As String, birthNote As String, deathNote As String, PPN, PPNval As String
        
    'Récupère le PPN et sort de la fonction s'il y a une erreur liée au PPN
    Set PPN = oXMLFile.SelectSingleNode("/record/controlfield[@tag='001']/text()")
    If Not PPN Is Nothing Then
        PPNval = PPN.NodeValue
    Else
        PPNval = "ERREUR [entrée n°" & jj & " : " & Chr(10) & inputPPN & "]"
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":I" & jj).Interior.Color = RGB(0, 0, 0)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":I" & jj).Font.Color = RGB(255, 255, 255)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":I" & jj).Value = PPNval
        Exit Sub
    End If
        
    note = ""
    'Récupération 103a
    Set UA103aNode = oXMLFile.SelectSingleNode("/record/datafield[@tag='103']/subfield[@code='a']/text()")
    If Not UA103aNode Is Nothing Then
        'Apparement il y a des espaces en 103a
        UA103a = Replace(UA103aNode.NodeValue, " ", "")
        If Left(UA103a, 1) = "-" Then
            birthNote = appendNote(birthNote, "Date av. J.-C. en UA103a")
            UA103a = Mid(UA103a, 2, 4)
        Else
            UA103a = Left(UA103a, 4)
        End If
        If InStr(UA103aNode.NodeValue, "?") > 0 Then
            birthNote = appendNote(birthNote, "Date incertaine en UA103a")
            UA103a = UA103a
        End If
    Else
        UA103a = "PAS DE UA103a"
    End If
    'Récupération 103b
    Set UA103bNode = oXMLFile.SelectSingleNode("/record/datafield[@tag='103']/subfield[@code='b']/text()")
    If Not UA103bNode Is Nothing Then
        UA103b = Replace(UA103bNode.NodeValue, " ", "")
        If Left(UA103b, 1) = "-" Then
            deathNote = appendNote(deathNote, "Date av. J.-C. en UA103b")
            UA103b = Mid(UA103b, 2, 4)
        Else
            UA103b = Left(UA103b, 4)
        End If
        If InStr(UA103bNode.NodeValue, "?") > 0 Then
            deathNote = appendNote(deathNote, "Date incertaine en UA103b")
            UA103b = UA103b
        End If
    Else
        UA103b = "PAS DE UA103b"
    End If
    
    'Récupération des 200f
    Set UA200fNode = oXMLFile.SelectSingleNode("/record/datafield[@tag='200']/subfield[@code='f']/text()")
    If Not UA200fNode Is Nothing Then
        moveRight = 0
    'Naissance
        UA200f = UA200fNode.NodeValue
        If InStr(UA200f, "av") > 0 Then
            birthNote = appendNote(birthNote, "Date av. J.-C. en UA200f1")
            deathNote = appendNote(deathNote, "Date av. J.-C. en UA200f2")
        End If
        If Left(UA200f, 1) = "-" Then
            birthNote = appendNote(birthNote, "Date av. J.-C. en UA200f1")
            moveRight = moveRight + 1
        End If
        UA200f1 = Mid(UA200f, 1 + moveRight, 4)
        If Mid(UA200f, 5 + moveRight, 1) = "?" Then
            birthNote = appendNote(birthNote, "Date incertaine en UA200f1")
            moveRight = moveRight + 1
        End If
    'Mort
        If Len(UA200f) >= (9 + moveRight) Then
            If Mid(UA200f, 10 + moveRight, 1) = "?" Then
                deathNote = appendNote(deathNote, "Date incertaine en UA200f2")
            End If
            UA200f2 = Mid(UA200f, 6 + moveRight, 4)
        Else
                deathNote = appendNote(deathNote, "Pb death date UA200f")
                mainWorkBook.Sheets("Résultats").Range("E" & jj).Interior.Color = RGB(0, 176, 240)
        End If
    Else
        UA200f1 = "PAS DE UA200f"
        UA200f2 = "PAS DE UA200f"
    End If
    
    'Comparaison des années de naissance
    If Replace(Replace(UA103a, ".", "X"), "?", "X") = Replace(Replace(UA200f1, ".", "X"), "?", "X") Then
        result = "OK"
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":I" & jj).Interior.Color = RGB(146, 208, 80)
    ElseIf (InStr(UA103a, "PAS DE") > 0) And ((InStr(UA200f1, "PAS DE") > 0) Or (UA200f1 = "....")) Then
        result = "OK"
        note = appendNote(note, "Pas de birth date")
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":I" & jj).Interior.Color = RGB(146, 208, 80)
    Else
        result = "Diff."
        note = appendNote(note, "Pas de corresp. birth date")
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":I" & jj).Interior.Color = RGB(255, 0, 0)
    End If
    'Comparaison des années de mort
    If (InStr(UA103b, "PAS DE") > 0) And ((InStr(UA200f2, "PAS DE") > 0) Or (UA200f2 = "....")) Then
        note = appendNote(note, "Pas de death date")
    ElseIf Replace(Replace(UA103b, "X", "."), "?", ".") <> Replace(Replace(UA200f2, "X", "."), "?", ".") Then
        result = "Diff."
        note = appendNote(note, "Pas de corresp. death date")
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":I" & jj).Interior.Color = RGB(255, 0, 0)
    End If
    
    'Vérification du format des données
    If UA103a <> Replace(Replace(UA103a, ".", "X"), "?", "X") Then
        birthNote = appendNote(birthNote, "Format d'année incorrect en UA103a")
        mainWorkBook.Sheets("Résultats").Range("B" & jj).Interior.Color = RGB(0, 176, 240)
    End If
    If UA200f1 <> Replace(Replace(UA200f1, "X", "."), "?", ".") Then
        birthNote = appendNote(birthNote, "Format d'année incorrect en UA200f")
        mainWorkBook.Sheets("Résultats").Range("C" & jj).Interior.Color = RGB(0, 176, 240)
    End If
    If UA103b <> Replace(Replace(UA103b, ".", "X"), "?", "X") Then
        deathNote = appendNote(deathNote, "Format d'année incorrect en UA103b")
        mainWorkBook.Sheets("Résultats").Range("E" & jj).Interior.Color = RGB(0, 176, 240)
    End If
    If UA200f2 <> Replace(Replace(UA200f2, "X", "."), "?", ".") Then
        deathNote = appendNote(deathNote, "Format d'année incorrect en UA200f")
        mainWorkBook.Sheets("Résultats").Range("F" & jj).Interior.Color = RGB(0, 176, 240)
    End If

    If note <> "" And InStr(note, "corresp.") = 0 Then
        mainWorkBook.Sheets("Résultats").Range("I" & jj).Interior.Color = RGB(0, 176, 240)
    ElseIf InStr(note, "corresp.") > 0 Then
        mainWorkBook.Sheets("Résultats").Range("I" & jj).Interior.Color = RGB(255, 0, 0)
    End If
    If birthNote <> "" Then
        'Si les deux dates sont marquées comme incertaines
        If Abs(InStr(birthNote, "incertain") - Len(birthNote)) <> InStrRev(birthNote, "incertain") And _
        InStr(birthNote, "incertain") <> 0 Then
            birthNote = appendNote(Replace(Replace(birthNote, "Date incertaine en UA200f1", ""), "Date incertaine en UA103a", ""), "Birth date incertaine")
        End If
        'Si les deux dates sont marquées comme av. J.-C.
        If Abs(InStr(birthNote, "av. J.-C.") - Len(birthNote)) <> InStrRev(birthNote, "av. J.-C.") And _
        InStr(birthNote, "av. J.-C.") <> 0 Then
            birthNote = appendNote(Replace(Replace(birthNote, "Date av. J.-C. en UA200f1", ""), "Date av. J.-C. en UA103a", ""), "Birth date av. J.-C.")
        End If
        mainWorkBook.Sheets("Résultats").Range("D" & jj).Interior.Color = RGB(0, 176, 240)
        'Enlève les retours à la ligne résiduels
        While InStr(birthNote, Chr(10) & Chr(10)) > 0
            birthNote = Replace(birthNote, Chr(10) & Chr(10), Chr(10))
        Wend
        If Left(birthNote, 1) = Chr(10) Then
            birthNote = Right(birthNote, Len(birthNote) - 1)
        End If
    End If
    If deathNote <> "" Then
        'Si les deux dates sont marquées comme incertaines
        If Abs(InStr(deathNote, "incertain") - Len(deathNote)) <> InStrRev(deathNote, "incertain") And _
        InStr(deathNote, "incertain") <> 0 Then
            deathNote = appendNote(Replace(Replace(deathNote, "Date incertaine en UA200f2", ""), "Date incertaine en UA103b", ""), "Death date incertaine")
        End If
        'Si les deux dates sont marquées comme av. J.-C.
        If Abs(InStr(deathNote, "av. J.-C.") - Len(deathNote)) <> InStrRev(deathNote, "av. J.-C.") And _
        InStr(deathNote, "av. J.-C.") <> 0 Then
            deathNote = appendNote(Replace(Replace(deathNote, "Date av. J.-C. en UA200f2", ""), "Date av. J.-C. en UA103b", ""), "Death date av. J.-C.")
        End If
        mainWorkBook.Sheets("Résultats").Range("G" & jj).Interior.Color = RGB(0, 176, 240)
        'Enlève les retours à la ligne résiduels
        While InStr(deathNote, Chr(10) & Chr(10)) > 0
            deathNote = Replace(deathNote, Chr(10) & Chr(10), Chr(10))
        Wend
        If Left(deathNote, 1) = Chr(10) Then
            deathNote = Right(deathNote, Len(deathNote) - 1)
        End If
    End If
        
    mainWorkBook.Sheets("Résultats").Range("A" & jj).Value = PPNval
    mainWorkBook.Sheets("Résultats").Range("B" & jj).Value = UA103a
    mainWorkBook.Sheets("Résultats").Range("C" & jj).Value = UA200f1
    mainWorkBook.Sheets("Résultats").Range("D" & jj).Value = birthNote
    mainWorkBook.Sheets("Résultats").Range("E" & jj).Value = UA103b
    mainWorkBook.Sheets("Résultats").Range("F" & jj).Value = UA200f2
    mainWorkBook.Sheets("Résultats").Range("G" & jj).Value = deathNote
    mainWorkBook.Sheets("Résultats").Range("H" & jj).Value = result
    mainWorkBook.Sheets("Résultats").Range("I" & jj).Value = note
    
    'Unload oXMLFile(XMLFileName)
    
End Sub
Sub getUBAge(oXMLFile, mainWorkBook, jj)
    
    Dim annee As String, anneeReW As String, output As String, PPN, PPNval As String
    Dim UB21XNodes, kk As Integer, ll As Integer

    'Récupère le PPN et sort de la fonction s'il y a une erreur liée au PPN
    Set PPN = oXMLFile.SelectSingleNode("/record/controlfield[@tag='001']/text()")
    If Not PPN Is Nothing Then
        PPNval = PPN.NodeValue
    Else
        PPNval = "ERREUR [entrée n°" & jj & " : " & Chr(10) & inputPPN & "]"
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":B" & jj).Interior.Color = RGB(0, 0, 0)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":B" & jj).Font.Color = RGB(255, 255, 255)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":B" & jj).Value = PPNval
        Exit Sub
    End If
    
    Set UB21XNodes = oXMLFile.SelectNodes("/record/datafield[@tag=214 or @tag=210]/subfield[@code='d']")
    
    output = ""
    If Not UB21XNodes Is Nothing Then
        For kk = 0 To (UB21XNodes.Length - 1)
            annee = UB21XNodes(kk).text()
            annee = Replace(annee, " ", "")
            annee = Replace(annee, "DL", "")
            annee = Replace(annee, "C", "")
            annee = Replace(annee, "cop", "")
            annee = Replace(annee, "P", "")
            annee = Replace(annee, ".", "")
            annee = Replace(annee, ",", "")
            annee = Replace(annee, "-", "")
            annee = Replace(annee, "?", "")
            annee = Replace(annee, "(", "")
            annee = Replace(annee, ")", "")
            
            'Si après cette première vague, du texte est encore présent, une détection caratère à caractère est effectuée
            If IsNumeric(annee) = False Then
                For ll = 1 To Len(annee)
                    If IsNumeric(Mid(annee, ll, 1)) = True Then
                        anneeReW = anneeReW & Mid(annee, ll, 1)
                        If Len(anneeReW) = 8 Then
                            If Left(anneeReW, 4) < Right(anneeReW, 4) Then
                               anneeReW = Right(anneeReW, 4)
                            Else
                                anneeReW = Left(anneeReW, 4)
                            End If
                        End If
                    End If
                Next
                'Donne 0 à la valeur annnee pour éviter de poser problèmes plus tard
                If anneeReW <> "" Then
                    annee = anneeReW
                Else
                    annee = "0"
                End If
            End If
            
            'Vérifie si l'année fait 4 chiffres (0 exclus), sinon essaye de voir si les 4 premiers chiffres font une année entre 1000 et 9999, sinon laisse vide
            If CDbl(annee) > 2030 Or CDbl(annee) < 1900 Then
                If CLng(Right(annee, 4)) < 2030 And CLng(Right(annee, 4)) > 1900 Then
                    annee = Right(annee, 4)
                ElseIf CLng(Left(annee, 4)) < 2030 And CLng(Left(annee, 4)) > 1900 Then
                    annee = Left(annee, 4)
                ElseIf CLng(annee) < 2030 And CLng(annee) > 1900 Then
                    annee = annee
                Else
                    output = ""
                End If
            End If
            
            'Regarde si une valeur est déjà présente. Si oui, regarde laquelle est la plus grande
            If output <> "" Then
                If CLng(annee) > CLng(output) Then
                    output = annee
                End If
            Else
                output = annee
            End If
        Next
    Else
        output = ""
    End If

'Colore les PPN qui n'ont pas eu d'années associées
    If output = "" Or output = "0" Then
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":B" & jj).Interior.Color = RGB(255, 0, 0)
    Else
        mainWorkBook.Sheets("Résultats").Range("B" & jj).Value = CLng(output)
    End If
        
    mainWorkBook.Sheets("Résultats").Range("A" & jj).Value = PPNval
    
    'Unload oXMLFile(XMLFileName)
    
End Sub
Sub getUBAgeDD(oXMLFile, mainWorkBook, jj)
    
    Dim annee1 As String, annee2 As String, output As String, PPN, PPNval As String
    Dim annee As String, anneeReW As String, output2 As String
    Dim UB100Node, UB21XNodes, kk As Integer, ll As Integer
'100 est prioritaire


    'Récupère le PPN et sort de la fonction s'il y a une erreur liée au PPN
    Set PPN = oXMLFile.SelectSingleNode("/record/controlfield[@tag='001']/text()")
    If Not PPN Is Nothing Then
        PPNval = PPN.NodeValue
    Else
        PPNval = "ERREUR [entrée n°" & jj & " : " & Chr(10) & inputPPN & "]"
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Interior.Color = RGB(0, 0, 0)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Font.Color = RGB(255, 255, 255)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Value = PPNval
        Exit Sub
    End If
    
    
 '100
    Set UB100Node = oXMLFile.SelectSingleNode("/record/datafield[@tag=100]/subfield[@code='a']/text()")
    
    output = ""
    If Not UB100Node Is Nothing Then
        annee1 = Mid(UB100Node.NodeValue, 10, 4)
        annee2 = Mid(UB100Node.NodeValue, 14, 4)
        If IsNumeric(annee2) = True Then
            If CInt(annee2) > CInt(annee1) And CInt(annee2) > 1900 And CInt(annee2) < 2030 Then
                output = annee2
            End If
        End If
        If CInt(annee1) > 1900 And CInt(annee1) < 2030 And output = "" Then
            output = annee1
        End If
    End If

'Colore les PPN qui n'ont pas eu d'années associées en 100
    If output = "" Or output = "0" Then
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":C" & jj).Interior.Color = RGB(255, 0, 0)
    Else
        mainWorkBook.Sheets("Résultats").Range("C" & jj).Value = CInt(output)
    End If


'21X
    Set UB21XNodes = oXMLFile.SelectNodes("/record/datafield[@tag=214 or @tag=210]/subfield[@code='d']")
    
    output2 = ""
    If Not UB21XNodes Is Nothing Then
        For kk = 0 To (UB21XNodes.Length - 1)
            annee = UB21XNodes(kk).text()
            annee = Replace(annee, " ", "")
            annee = Replace(annee, "DL", "")
            annee = Replace(annee, "C", "")
            annee = Replace(annee, "cop", "")
            annee = Replace(annee, "P", "")
            annee = Replace(annee, ".", "")
            annee = Replace(annee, ",", "")
            annee = Replace(annee, "-", "")
            annee = Replace(annee, "?", "")
            annee = Replace(annee, "(", "")
            annee = Replace(annee, ")", "")
            
            'Si après cette première vague, du texte est encore présent, une détection caratère à caractère est effectuée
            If IsNumeric(annee) = False Then
                For ll = 1 To Len(annee)
                    If IsNumeric(Mid(annee, ll, 1)) = True Then
                        anneeReW = anneeReW & Mid(annee, ll, 1)
                        If Len(anneeReW) = 8 Then
                            If Left(anneeReW, 4) < Right(anneeReW, 4) Then
                               anneeReW = Right(anneeReW, 4)
                            Else
                                anneeReW = Left(anneeReW, 4)
                            End If
                        End If
                    End If
                Next
                'Donne 0 à la valeur annnee pour éviter de poser problèmes plus tard
                If anneeReW <> "" Then
                    annee = anneeReW
                Else
                    annee = "0"
                End If
            End If
            
            'Vérifie si l'année fait 4 chiffres (0 exclus), sinon essaye de voir si les 4 premiers chiffres font une année entre 1000 et 9999, sinon laisse vide
            If CDbl(annee) > 2030 Or CDbl(annee) < 1900 Then
                If CLng(Right(annee, 4)) < 2030 And CLng(Right(annee, 4)) > 1900 Then
                    annee = Right(annee, 4)
                ElseIf CLng(Left(annee, 4)) < 2030 And CLng(Left(annee, 4)) > 1900 Then
                    annee = Left(annee, 4)
                ElseIf CLng(annee) < 2030 And CLng(annee) > 1900 Then
                    annee = annee
                Else
                    output2 = ""
                End If
            End If
            
            'Regarde si une valeur est déjà présente. Si oui, regarde laquelle est la plus grande
            If output2 <> "" Then
                If CLng(annee) > CLng(output2) Then
                    output2 = annee
                End If
            Else
                output2 = annee
            End If
        Next
    Else
        output2 = ""
    End If

'Colore le 21X des PPN qui n'ont pas eu d'années associées en 21X
    If output2 = "" Or output2 = "0" Then
        mainWorkBook.Sheets("Résultats").Range("B" & jj & ":B" & jj).Interior.Color = RGB(255, 0, 0)
    Else
        mainWorkBook.Sheets("Résultats").Range("B" & jj).Value = CLng(output2)
    End If
        
    mainWorkBook.Sheets("Résultats").Range("A" & jj).Value = PPNval
    
    'Unload oXMLFile(XMLFileName)
    
End Sub
Sub getUBAge100(oXMLFile, mainWorkBook, jj)
    
    Dim annee1 As String, annee2 As String, output As String, PPN, PPNval As String
    Dim UB100Node, kk As Integer, ll As Integer

    'Récupère le PPN et sort de la fonction s'il y a une erreur liée au PPN
    Set PPN = oXMLFile.SelectSingleNode("/record/controlfield[@tag='001']/text()")
    If Not PPN Is Nothing Then
        PPNval = PPN.NodeValue
    Else
        PPNval = "ERREUR [entrée n°" & jj & " : " & Chr(10) & inputPPN & "]"
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":B" & jj).Interior.Color = RGB(0, 0, 0)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":B" & jj).Font.Color = RGB(255, 255, 255)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":B" & jj).Value = PPNval
        Exit Sub
    End If
    
    Set UB100Node = oXMLFile.SelectSingleNode("/record/datafield[@tag=100]/subfield[@code='a']/text()")
    
    output = ""
    If Not UB100Node Is Nothing Then
        annee1 = Mid(UB100Node.NodeValue, 10, 4)
        annee2 = Mid(UB100Node.NodeValue, 14, 4)
        If IsNumeric(annee2) = True Then
            If CInt(annee2) > CInt(annee1) And CInt(annee2) > 1900 And CInt(annee2) < 2030 Then
                output = annee2
            End If
        End If
        If CInt(annee1) > 1900 And CInt(annee1) < 2030 And output = "" Then
            output = annee1
        End If
    End If

'Colore les PPN qui n'ont pas eu d'années associées
    If output = "" Or output = "0" Then
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":B" & jj).Interior.Color = RGB(255, 0, 0)
    Else
        mainWorkBook.Sheets("Résultats").Range("B" & jj).Value = CInt(output)
    End If
        
    mainWorkBook.Sheets("Résultats").Range("A" & jj).Value = PPNval
    
    'Unload oXMLFile(XMLFileName)
    
End Sub
Sub getStatsUBAge()
    Dim nonEmptyRows As Integer
    Dim yearDataRange As Range
    
    mainWorkBook.Sheets("Résultats").Activate
    Range("A:B").Sort key1:=Cells(2, 2), order1:=xlAscending, Header:=xlYes
    nonEmptyRows = Application.WorksheetFunction.CountA(Range("B:B"))
    Set yearDataRange = Range("B2:B" & nonEmptyRows)
    
    Range("D2") = Application.WorksheetFunction.Average(yearDataRange)
    Range("E2") = Application.WorksheetFunction.Median(yearDataRange)
    
    Range("F2") = rowIndex - nonEmptyRows
End Sub
Sub getStatsUBAgeDD()
    Dim nonEmptyRows As Integer
    Dim yearDataRange As Range
    
    mainWorkBook.Sheets("Résultats").Activate
    Range("A:C").Sort key1:=Cells(2, 3), order1:=xlAscending, Header:=xlYes
    '21X
    nonEmptyRows = Application.WorksheetFunction.CountA(Range("B:B"))
    Set yearDataRange = Range("B2:B" & nonEmptyRows)
    
    Range("E2") = Application.WorksheetFunction.Average(yearDataRange)
    Range("F2") = Application.WorksheetFunction.Median(yearDataRange)
    
    Range("G2") = rowIndex - nonEmptyRows
    '100
    nonEmptyRows = Application.WorksheetFunction.CountA(Range("C:C"))
    Set yearDataRange = Range("C2:C" & nonEmptyRows)
    
    Range("E4") = Application.WorksheetFunction.Average(yearDataRange)
    Range("F4") = Application.WorksheetFunction.Median(yearDataRange)
    
    Range("G4") = rowIndex - nonEmptyRows
End Sub
Sub getEdition(oXMLFile, mainWorkBook, jj)
    
    Dim kk As Integer, ll As Integer
    Dim PPN, PPNval As String
    Dim UB205, edition As String
    Dim UB200a, UB200eNodes, titre As String, UB451Nodes, UB451oNodes, titre451 As String
    Dim UB7XXNodes, UB7XXPPN, mentR As String
    Dim clefTitre As String, catClefTitre As String

    'Récupère le PPN et sort de la fonction s'il y a une erreur liée au PPN
    Set PPN = oXMLFile.SelectSingleNode("/record/controlfield[@tag='001']/text()")
    If Not PPN Is Nothing Then
        PPNval = PPN.NodeValue
    Else
        PPNval = "ERREUR [entrée n°" & jj & " : " & Chr(10) & inputPPN & "]"
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":B" & jj).Interior.Color = RGB(0, 0, 0)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":B" & jj).Font.Color = RGB(255, 255, 255)
        mainWorkBook.Sheets("Résultats").Range("A" & jj & ":B" & jj).Value = PPNval
        Exit Sub
    End If
    
    Set UB205 = oXMLFile.SelectSingleNode("/record/datafield[@tag='205']/subfield[@code='a']/text()")
    If Not UB205 Is Nothing Then
        edition = UB205.NodeValue
    Else
        edition = ""
    End If
        
    Set UB200a = oXMLFile.SelectSingleNode("/record/datafield[@tag='200']/subfield[@code='a']/text()")
    If Not UB200a Is Nothing Then
        titre = UB200a.NodeValue
        Set UB200eNodes = oXMLFile.SelectNodes("/record/datafield[@tag='200']/subfield[@code='e']/text()")
        If Not UB200eNodes Is Nothing Then
            For kk = 0 To (UB200eNodes.Length - 1)
                titre = titre & " : " & UB200eNodes(kk).NodeValue
            Next
        End If
    Else
        titre = ""
    End If
    
    Set UB451Nodes = oXMLFile.SelectNodes("/record/datafield[@tag='451']")
    If Not UB451Nodes Is Nothing Then
        For kk = 0 To (UB451Nodes.Length - 1)
            Set UB451oNodes = UB451Nodes(kk).ChildNodes
            titre451 = ""
            For ll = 0 To (UB451oNodes.Length - 1)
                If UB451oNodes(ll).getAttribute("code") = "t" And titre451 = "" Then
                    titre451 = UB451oNodes(ll).text
                ElseIf UB451oNodes(ll).getAttribute("code") = "o" Then
                    titre451 = titre451 & " : " & UB451oNodes(ll).text
                End If
            Next
            titre = appendNote(titre, titre451)
        Next
    End If
    
    Set UB7XXPPN = oXMLFile.SelectNodes("/record/datafield[@tag>=700 and @tag<800]/subfield[@code='3']")
    mentR = ""
    If Not UB7XXPPN Is Nothing Then
        For kk = 0 To (UB7XXPPN.Length - 1)
            Set UB7XXNodes = UB7XXPPN(kk).parentNode
            If UB7XXNodes.getAttribute("tag") < 710 Then
                mentR = appendNote(mentR, UB7XXPPN(kk).text)
            Else
                mentR = appendNote(mentR, "C" & UB7XXPPN(kk).text)
            End If
        Next
    End If

    'Génère la clef et attribue sa catégorie
    clefTitre = UCase(titre)
    If Left(clefTitre, 4) = "LES " Or Left(clefTitre, 4) = "UNE " Or Left(clefTitre, 4) = "DES " Then
        clefTitre = Right(clefTitre, Len(clefTitre) - 4)
    ElseIf Left(clefTitre, 3) = "LE " Or Left(clefTitre, 3) = "LA " Or Left(clefTitre, 3) = "UN " Then
        clefTitre = Right(clefTitre, Len(clefTitre) - 3)
    ElseIf Left(clefTitre, 2) = "L'" Then
        clefTitre = Right(clefTitre, Len(clefTitre) - 2)
    End If
    clefTitre = Replace(clefTitre, Chr(10) & "LE ", Chr(10))
    clefTitre = Replace(clefTitre, Chr(10) & "LES ", Chr(10))
    clefTitre = Replace(clefTitre, Chr(10) & "LA ", Chr(10))
    clefTitre = Replace(clefTitre, Chr(10) & "UN ", Chr(10))
    clefTitre = Replace(clefTitre, Chr(10) & "UNE ", Chr(10))
    clefTitre = Replace(clefTitre, Chr(10) & "DES ", Chr(10))
    clefTitre = Replace(clefTitre, Chr(10) & "L'", Chr(10))
    clefTitre = Replace(clefTitre, " LE ", " ")
    clefTitre = Replace(clefTitre, " LES ", " ")
    clefTitre = Replace(clefTitre, " LA ", " ")
    clefTitre = Replace(clefTitre, " L'", " ")
    clefTitre = Replace(clefTitre, " UN ", " ")
    clefTitre = Replace(clefTitre, " UNE ", " ")
    clefTitre = Replace(clefTitre, " DES ", " ")
    clefTitre = Replace(clefTitre, " À ", " ")
    clefTitre = Replace(clefTitre, " DE ", " ")
    clefTitre = Replace(clefTitre, ",", "")
    clefTitre = Replace(clefTitre, ":", "")
    clefTitre = Replace(clefTitre, ";", "")
    clefTitre = Replace(clefTitre, "?", "")
    clefTitre = Replace(clefTitre, "!", "")
    clefTitre = Replace(clefTitre, ".", "")
    clefTitre = Replace(clefTitre, "(", "")
    clefTitre = Replace(clefTitre, ")", "")
    clefTitre = Replace(clefTitre, "[", "")
    clefTitre = Replace(clefTitre, "]", "")
    clefTitre = Replace(clefTitre, """", "")
    clefTitre = Replace(clefTitre, "&", "")
    clefTitre = Replace(clefTitre, "=", "")
    clefTitre = Replace(clefTitre, Chr(171), "")
    clefTitre = Replace(clefTitre, Chr(187), "")
    clefTitre = Replace(clefTitre, "'", "")
    clefTitre = Replace(clefTitre, "-", " ")
    While InStr(clefTitre, "  ") > 0
        clefTitre = Replace(clefTitre, "  ", " ")
    Wend
    clefTitre = Trim(clefTitre)
    clefTitre = Replace(clefTitre, " ", "_")
    
    'Je passe la détection de la catégorie de la clef en fonction des titres maintenant (si il reste encore une clef)
    'Select Case UBound(Split(clefTitre, "_")) + 1
    '    Case 0 To 3
    '        catClefTitre = "1 : >66 %"
    '    Case 4 To 7
    '        catClefTitre = "2 : >70 %"
    '    Case 8 To 12
    '        catClefTitre = "3 : >75 %"
    '    Case Else
    '        catClefTitre = "4 : >80 %"
    'End Select
        
    mainWorkBook.Sheets("Résultats").Range("A" & jj).Value = PPNval
    mainWorkBook.Sheets("Résultats").Range("B" & jj).Value = edition
    mainWorkBook.Sheets("Résultats").Range("C" & jj).Value = titre
    mainWorkBook.Sheets("Résultats").Range("D" & jj).Value = clefTitre
    'mainWorkBook.Sheets("Résultats").Range("E" & jj).Value = catClefTitre
    mainWorkBook.Sheets("Résultats").Range("F" & jj).Value = mentR
    'Unload oXMLFile(XMLFileName)
    
End Sub
Sub getEditionFind()
    Dim zz As Integer, yy As Integer, xx As Integer
    Dim fullClefTitreOr, fullClefTitreDup, clefTitreOr, clefTitreDup, clefMot As Variant, clefMotCount As Integer
    Dim clefPPNor, clefPPNdup, clefPPN As Variant, clefPPNCount As Integer
    Dim tableMatch(99, 3), tableCount As Integer, output As String
    Dim catClefTitre As Integer, percMin As Integer, matchEffCul As Integer, matchNb As Integer, skipPerc As Boolean
    
    
    mainWorkBook.Sheets("Résultats").Activate
    
    For zz = 2 To rowIndex
        If Range("B" & zz).Value <> "" Then
            tableCount = 0
            'catClefTitre = CInt(Left(Range("E" & zz).Value, 1))
            fullClefTitreOr = Split(Range("D" & zz).Value, Chr(10))
            For Each clefTitreOr In fullClefTitreOr
                clefTitreOr = Split(clefTitreOr, "_")
                For yy = 2 To rowIndex
    'Check la clef du titre
                    fullClefTitreDup = Split(Range("D" & yy).Value, Chr(10))
                    For Each clefTitreDup In fullClefTitreDup
                        clefTitreDup = Split(clefTitreDup, "_")
                        clefMotCount = 0
                        matchEffCul = 0
                        skipPerc = False
                        'comment ej fais s ke dup est plus long
                        If UBound(clefTitreOr) <= UBound(clefTitreDup) Then
                            For Each clefMot In clefTitreOr
                                If Left(clefMot, 4) = Left(clefTitreDup(clefMotCount), 4) And Right(clefMot, 4) = Right(clefTitreDup(clefMotCount), 4) Then
                                    matchEffCul = matchEffCul + 4
                                ElseIf Left(clefMot, 4) = Left(clefTitreDup(clefMotCount), 4) Or Right(clefMot, 4) = Right(clefTitreDup(clefMotCount), 4) Then
                                    matchEffCul = matchEffCul + 3
                                End If
                                clefMotCount = clefMotCount + 1
                            Next
                        Else
                            For Each clefMot In clefTitreDup
                                If Left(clefMot, 4) = Left(clefTitreOr(clefMotCount), 4) And Right(clefMot, 4) = Right(clefTitreOr(clefMotCount), 4) Then
                                    matchEffCul = matchEffCul + 4
                                ElseIf Left(clefMot, 4) = Left(clefTitreOr(clefMotCount), 4) Or Right(clefMot, 4) = Right(clefTitreOr(clefMotCount), 4) Then
                                    matchEffCul = matchEffCul + 3
                                End If
                                clefMotCount = clefMotCount + 1
                            Next
                        End If
                        matchNb = matchEffCul / ((UBound(clefTitreOr) + 1) * 2 + (UBound(clefTitreDup) + 1) * 2) * 100
                        If matchEffCul = clefMotCount * 4 Then
                            skipPerc = True
                        End If
                        
                        percMin = 80
                        'Select Case catClefTitre
                        '    Case 1
                        '        percMin = 66
                        '    Case 2
                        '        percMin = 70
                        '    Case 3
                        '        percMin = 75
                        '    Case 4
                        '        percMin = 80
                        'End Select
                        
        'Si ya match, il va regarder les auteurs
                        If (matchNb >= percMin Or skipPerc = True) And Range("A" & zz).Value <> Range("A" & yy).Value And Range("A" & yy).Value <> "" Then
                            tableMatch(tableCount, 0) = Range("A" & yy).Value
                            tableMatch(tableCount, 1) = matchNb
                            tableMatch(tableCount, 2) = 0
                            tableMatch(tableCount, 3) = 0
                            
                            clefPPNor = Split(Range("F" & zz).Value, Chr(10))
                            clefPPNdup = Split(Range("F" & yy).Value, Chr(10))
                            For Each clefPPN In clefPPNor
                                For xx = 0 To UBound(clefPPNdup)
                                    If clefPPN = clefPPNdup(xx) Then
                                        tableMatch(tableCount, 3) = tableMatch(tableCount, 3) + 2
                                        If Left(clefPPN, 1) <> "C" Then
                                            tableMatch(tableCount, 2) = tableMatch(tableCount, 2) + 2
                                        End If
                                    End If
                                Next
                            Next
                            If UBound(clefPPNor) + UBound(clefPPNdup) + 2 >= 0 Then
                                tableMatch(tableCount, 2) = CInt(tableMatch(tableCount, 2) / (UBound(clefPPNor) + UBound(clefPPNdup) + 2) * 100)
                                tableMatch(tableCount, 3) = CInt(tableMatch(tableCount, 3) / (UBound(clefPPNor) + UBound(clefPPNdup) + 2) * 100)
                            Else
                                tableMatch(tableCount, 2) = "Ø"
                                tableMatch(tableCount, 3) = "Ø"
                            End If
                            tableCount = tableCount + 1
                        End If
                        
                    Next
                Next
            Next

            If tableCount > 0 Then
            output = "Double éd. possible :"
                For yy = 0 To tableCount - 1
                    output = appendNote(output, tableMatch(yy, 0) & " : " & tableMatch(yy, 1) & "%, " & tableMatch(yy, 2) & "%, " & tableMatch(yy, 3) & "%")
                    If tableMatch(yy, 2) > 0 Then
                        mainWorkBook.Sheets("Résultats").Range("A" & zz & ":G" & zz).Interior.Color = RGB(255, 0, 0)
                    ElseIf tableMatch(yy, 3) + 1 > tableMatch(yy, 2) And _
                    mainWorkBook.Sheets("Résultats").Range("A" & zz & ":G" & zz).Interior.Color <> RGB(255, 0, 0) Then
                        mainWorkBook.Sheets("Résultats").Range("A" & zz & ":G" & zz).Interior.Color = RGB(255, 192, 0)
                    ElseIf mainWorkBook.Sheets("Résultats").Range("A" & zz & ":G" & zz).Interior.Color <> RGB(255, 0, 0) And _
                    mainWorkBook.Sheets("Résultats").Range("A" & zz & ":G" & zz).Interior.Color <> RGB(255, 192, 0) Then
                        mainWorkBook.Sheets("Résultats").Range("A" & zz & ":G" & zz).Interior.Color = RGB(0, 176, 240)
                    End If
                Next
            Else
                output = "Aucune détection automatique"
                mainWorkBook.Sheets("Résultats").Range("A" & zz & ":G" & zz).Interior.Color = RGB(146, 208, 80)
            End If
            Range("G" & zz) = output
        End If
    Next
End Sub
Function appendNote(var As String, text As String)
    If var = "" Then
        var = text
    Else
        var = var & Chr(10) & text
    End If
    appendNote = var
End Function
Sub cleanData()
    Worksheets("Résultats").Activate
    Range("A:ZZ").Delete
    Worksheets("Introduction").Activate
    Range("I2:J999999").ClearContents
    Range("I2").Select
End Sub
Sub imp_PPN_Alma()
    
    
Dim nbRow As Integer
Dim exportAlma As Workbook

Set mainWorkBook = ActiveWorkbook
folderPath = Application.ActiveWorkbook.Path
Workbooks.Open fileName:=folderPath & "\export_alma_ConStance.xlsx"
Set exportAlma = Workbooks("export_alma_ConStance.xlsx")

nbRow = Cells(Rows.count, "K").End(xlUp).Row

'Récupère les données
Dim PPN
    
For rowIndex = 2 To nbRow
    PPN = exportAlma.Worksheets("Results").Cells(rowIndex, 10).Value
    PPN = Right(Mid(PPN, InStr(PPN, "(PPN)"), 14), 9)
    mainWorkBook.Worksheets("Introduction").Cells(rowIndex, 9).Value = PPN
Next
Workbooks("export_alma_ConStance.xlsx").Close
mainWorkBook.Worksheets("Introduction").Activate
Range("A2").Select
    
End Sub
Sub Main()
'Timer : https://www.thespreadsheetguru.com/the-code-vault/2015/1/28/vba-calculate-macro-run-time

'Timer : début
Dim StartTime As Double
Dim MinutesElapsed As String
StartTime = Timer

Set mainWorkBook = ActiveWorkbook

mainWorkBook.Worksheets("Introduction").Activate
CSX = Right(Range("H2").Value, 5)
Range("I:I").Sort key1:=Cells(2, 9), order1:=xlAscending, Header:=xlYes

formatEnTetes

read_Sudoc_Data

'Lance un script additionnel si nécessaire
Select Case CSX
    Case "[CS4]", "[CS5]"
        getStatsUBAge
    Case "[CS7]"
        getStatsUBAgeDD
    Case "[CS6]"
        getEditionFind
End Select

mainWorkBook.Worksheets("Résultats").Activate
Columns("A:" & colIndex).AutoFit
Rows("1:" & rowIndex).AutoFit

'Formattage spéciaux pour un script
Select Case CSX
    Case "[CS4]", "[CS5]"
        With mainWorkBook.Worksheets("Résultats").Range("D1:F1")
            .Interior.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlCenter
            .Font.Color = RGB(255, 255, 255)
        End With
        With mainWorkBook.Sheets("Résultats").Range("D2:F2")
            .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
            .Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        Columns("D:F").AutoFit
    Case "[CS7]"
        With mainWorkBook.Worksheets("Résultats").Range("E1:G1")
            .Interior.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlCenter
            .Font.Color = RGB(255, 255, 255)
        End With
        With mainWorkBook.Worksheets("Résultats").Range("E3:G3")
            .Interior.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlCenter
            .Font.Color = RGB(255, 255, 255)
        End With
        With mainWorkBook.Sheets("Résultats").Range("E2:G2")
            .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
            .Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With mainWorkBook.Sheets("Résultats").Range("E4:G4")
            .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
            .Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        Columns("E:G").AutoFit
    Case "[CS6]"
        mainWorkBook.Sheets("Résultats").Range("B:D").HorizontalAlignment = xlLeft
        mainWorkBook.Sheets("Résultats").Range("C:C").ColumnWidth = 65
        mainWorkBook.Sheets("Résultats").Range("D:D").ColumnWidth = 55
        mainWorkBook.Sheets("Résultats").Range("B:B").ColumnWidth = 30
        mainWorkBook.Sheets("Résultats").Range("B1:D1").HorizontalAlignment = xlCenter
        mainWorkBook.Sheets("Résultats").Range("E:E").ColumnWidth = 0
        mainWorkBook.Sheets("Résultats").Range("A:G").Sort key1:=Cells(2, 4), order1:=xlAscending, Header:=xlYes
End Select

Range("A1").Select

'Timer suite & fin
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
MsgBox "Exécution terminée en " & MinutesElapsed & "."

End Sub
