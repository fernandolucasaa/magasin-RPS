Attribute VB_Name = "Module5"
Option Explicit
Dim ws As Worksheet
Dim X As Integer
Dim l As Byte
Dim TabSheet() As Variant
Dim myText(15) As OLEObject

Function FichOuvert(F As String) As Boolean

On Error Resume Next
FichOuvert = Not Workbooks(F) Is Nothing
    
End Function

'Enregistrer le bon de prêt
Sub Nouveau_bon()

'Calculer le temps de calcul
'Dim start As Double, seconds As Double
'start = Timer

'Variables pour les fichiers
Dim bonPret As String, dispocockpitGererique As String, chemin As String, tampon As String, pieces As String, numeroPret As String, pret As String

pret = "pret.xlsm"
bonPret = "Bon_pret.xlsm"
dispocockpitGererique = "DISPOCOCKPIT GENERIQUE.xlsx"
tampon = "Tampon.xlsm"
pieces = "PIECES.xlsm"
numeroPret = "Numero_pret.xlsm"

Windows(bonPret).Activate
chemin = Application.ActiveWorkbook.Path

'____________________________________________________________________________________________________________________'
            'Tester l'entrée des données
                    
'Tester si les cellules C3 (CM) et C4 (quantité) sont vides
If Range("C3") = "" Or Range("C4") = "" Then
    MsgBox ("Veuillez remplir le numéro du CMS, la quantité empruntée, le nom de l'emprunteur et l'observation")
    Exit Sub
End If

'Tester si la cellule C3 (CM) est un nombre, qu'il est composé de 10 chiffres et qu'il existe
Dim CMS As String
CMS = Range("C3").Value
    
If Not IsNumeric(CMS) Then
    MsgBox ("Veuillez entrer un CMS composé de 10 chiffres")
    Exit Sub
Else
    Dim longueurCMS As Integer
    longueurCMS = Len(Range("C3").Value) 'Comptage des caractères
            
    'Tester si la cellule contient 10 caractères
    If longueurCMS <> 10 Then
        MsgBox ("Veuillez entrer un CMS composé de 10 chiffres")
        Exit Sub
    End If
    
    'Tester si le CMS existe
    If IsError(Range("E3")) Then
        MsgBox ("Le CMS indiqué n'existe pas")
        ' ? Range("E3") = "=IF(RC[-2]="""","""",VLOOKUP(RC[-2],Piece!C[-4]:C[-2],2,FALSE))" 'RC[-2] -> CM
        ' ? Range("C3").Select
        ' ? Application.CutCopyMode = False
        Exit Sub
    End If
End If

'Tester si la cellule C4 (quantité) est un nombre
Dim Quantite As Long
Quantite = Range("C4").Value

If Not IsNumeric(Quantite) Then
    MsgBox ("Veuillez entrer le nombre de pièce à sortir")
    Exit Sub
End If

'Tester si le emprunteur est dans la base des données
If IsError(Range("E5")) And Range("E8").Value = "" Then
    MsgBox ("Le nom rempli n'est pas dans la liste. Verifiez le nome et/ou indiquer votre nom et le nom de votre responsable dans la case Commentaires.")
    Exit Sub
End If

'___________________________________________________________________________________________________________________'
                'Message pour valider la création du bon de prêt'

If MsgBox("Etes-vous sur de vouloir créer le bon de prêt ?", vbYesNo, "Demande de confirmation") = vbYes Then

    'Masquer les opérations de la macro
    Application.ScreenUpdating = False
                        
    Range("C3").Select 'Active la cellule C3 (CMS)
    Selection.NumberFormat = "0" 'Change le type de cellule en "nombre"
    
    '____________________________________________________________________________________________________________________'
                'Copier les informations du fichier "Bon_pret.xlms" dans le fichier "Tampon.xlms"
    
    'Tester si le fichier "Tampon" est ouvert et l'ouvrir si nécessaire
    If Not (FichOuvert(tampon)) Then
        Workbooks.Open Filename:=chemin & "\" & tampon
    End If
    
    Windows(tampon).Activate
    
    'Retirer la proctection du fichier "Tampon.xlsm" pour permettre faire des modifications
    ActiveSheet.Unprotect "spr"
    
    'Ajout d'une ligne dans l'onglet "Pret" du fichier "Tampon"
    Range("A2").Activate
    ActiveCell.EntireRow.Insert Shift:=xlDown 'Insert une ligne au dessus de la case A2
    
    'Formater cellules (DATE et CMS)
    Range("B2").Activate
    Selection.NumberFormat = "m/d/yyyy" 'Change le type de cellule en "date courte"
    Range("C2").Select
    Selection.NumberFormat = "0" 'Change le type de cellule en "nombre"
    
    'Copie de la date (B2), du CMS (C3), de la quantité (C4), du responsable de l'emprunteur (C6), du numéro de série (C5),
    'du nom de l'emprunteur (E6), de la destination de la pièce (Unité) (E5), du numéro de téléphone (E7) , du type de prêt (C8)
    'et du commentaire (E8) dans le fichier "Tampon" dans l'onglet "Pret"
    
    Workbooks(tampon).Sheets("Pret").Cells(2, 2).Value = Workbooks(bonPret).Sheets("Bon_pret").Cells(2, 2) 'date
    Workbooks(tampon).Sheets("Pret").Cells(2, 3).Value = Workbooks(bonPret).Sheets("Bon_pret").Cells(3, 3) 'CMS
    Workbooks(tampon).Sheets("Pret").Cells(2, 7).Value = Workbooks(bonPret).Sheets("Bon_pret").Cells(4, 3) 'quantité
    Workbooks(tampon).Sheets("Pret").Cells(2, 9).Value = Workbooks(bonPret).Sheets("Bon_pret").Cells(6, 3) 'responsable de l'emprunteur
    Workbooks(tampon).Sheets("Pret").Cells(2, 5).Value = Workbooks(bonPret).Sheets("Bon_pret").Cells(5, 3) 'numéro de série
    Workbooks(tampon).Sheets("Pret").Cells(2, 10).Value = Workbooks(bonPret).Sheets("Bon_pret").Cells(6, 5) 'nom de l'emprunteur
    Workbooks(tampon).Sheets("Pret").Cells(2, 8).Value = Workbooks(bonPret).Sheets("Bon_pret").Cells(5, 5) 'unité
    Workbooks(tampon).Sheets("Pret").Cells(2, 11).Value = Workbooks(bonPret).Sheets("Bon_pret").Cells(7, 5) 'numéro de téléphone
    Workbooks(tampon).Sheets("Pret").Cells(2, 12).Value = Workbooks(bonPret).Sheets("Bon_pret").Cells(8, 3) 'type de prêt
    Workbooks(tampon).Sheets("Pret").Cells(2, 24).Value = Workbooks(bonPret).Sheets("Bon_pret").Cells(8, 5) 'commentaire

    'Remplir la cellule D2 (Désignation)
    Workbooks(tampon).Activate
    Sheets("Pret").Select
    'Recherche V de la valeur de la cellule C2 (CMS) pour la cellule D2 (Désignation) dans l'onglet "Piece" du fichier "Tampon.xlsm"
    Range("D2").FormulaR1C1 = "=VLOOKUP(RC[-1],Piece!C[-3]:C[2],2,FALSE)"
    
    'Mettre les commentaires en évidence
    If Workbooks(tampon).Sheets("Pret").Range("X2").Value <> "" Then
        Range("X2").Interior.Color = RGB(255, 0, 0)
    End If
    
    '____________________________________________________________________________________________________________________'
                'Mise à jour du fichier "Tampon.xlsm" en utilisant les fichiers "DISPOCOCKPIT GENERIQUE.xlsx" et "PIECES.xlsm"
    
    If Not (FichOuvert(dispocockpitGererique)) Then
        Workbooks.Open Filename:=chemin & "\" & dispocockpitGererique
    End If
    
    If Not (FichOuvert(pieces)) Then
        Workbooks.Open Filename:=chemin & "\" & pieces
    End If
    
    '____________________________________________________________________________________________________________________'
                'Mise à jour des données dans le fichier de création de pret ("Tampon.xlsm")

    'Remplir la cellule F2 (Empla.)
    Workbooks(tampon).Activate
    Sheets("Pret").Select
    'Recherche V de la valeur de la cellule C2 (CMS) pour la cellule F2 (Empla.) dans l'onglet "resultat" du fichier "PIECES.xlsm"
    Range("F2").FormulaR1C1 = "=VLOOKUP(RC[-3],[PIECES.xlsm]resultat!C1:C6,4,FALSE)"
    
    'Sélectionne la cellule P2 (Valeur Stock) dans le fichier
    Range("P2").FormulaR1C1 = _
        "=VLOOKUP(RC[-13],'[DISPOCOCKPIT GENERIQUE.xlsx]MPR PILOTAGE'!A:P,13,FALSE)"
    
    'Sélectionne la cellule Q2 (Quant. en stock SAP) dans le fichier
    Range("Q2").FormulaR1C1 = _
        "=VLOOKUP(RC[-14],'[DISPOCOCKPIT GENERIQUE.xlsx]MPR PILOTAGE'!C1:C16,12,FALSE)"
    
    'Mise en forme de la cellule O2 (Delta jour)
    Range("O2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Active la feuille "Pret"
    Sheets("Pret").Select
    
    'Sélectionne la cellule P2 (Valuer Stock)
    Range("P2").FormulaR1C1 = _
        "=VLOOKUP(RC[-13],'[DISPOCOCKPIT GENERIQUE.xlsx]MPR PILOTAGE'!C1:C16,5,FALSE)"
    
    'Mise en forme de la cellule Q2 (Quant. en stock SAP)
    Range("Q2").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="0", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    'Calcul conditionnel
    'Delta jour
    Range("O2").FormulaR1C1 = "=IF(RC[-2]=0,TODAY()-RC[-13],RC[-2]-RC[-13])"
    Range("S2").FormulaR1C1 = "=IF(RC[-6]<30,1,0)"
    'stock deporté
    Range("T2").FormulaR1C1 = "=IF(AND(RC[-7]<60,RC[-7]>29),1,0)"
    Range("U2").FormulaR1C1 = "=IF(RC[-8]>60,1,0)"
    Range("V2").FormulaR1C1 = "=IF(AND(RC[-9]<60,RC[-9]>29),1,0)"
    Range("W2").FormulaR1C1 = "=+IF(RC[-12]>0,0,1)"
    'Q. en stock physique
    Range("R2").FormulaR1C1 = "=RC[-1]-RC[-11]"  'Q. en stock SAP - Q.prélevée
    
    Rows("2:2").Select
    Selection.Font.Bold = False
    Selection.EntireRow.AutoFit
    
    '____________________________________________________________________________________________________________________'
                'Fin macro de la copie des informations pour la création de prêt
    
    'Fermeture des fichiers
    Windows(dispocockpitGererique).Activate
    Workbooks(dispocockpitGererique).Close SaveChanges:=False
    Workbooks(pieces).Activate
    Workbooks(pieces).Close SaveChanges:=False

    '____________________________________________________________________________________________________________________'
                'Début macro pour faire le calcul du nouveau numéro de ligne

    If Not (FichOuvert(numeroPret)) Then
        Workbooks.Open Filename:=chemin & "\Numero_pret\" & numeroPret
    End If
        
    Dim numero As Integer
    
    Windows(numeroPret).Activate
    numero = Range("A1").Value + 1
    Range("A1").Value = numero
    ActiveWorkbook.Save
    ActiveWindow.Close
    
    Windows(tampon).Activate
    Worksheets("Pret").Select
    Range("A2").Value = numero
    
    'fin macro nouveau numéro de ligne
    
    '____________________________________________________________________________________________________________________'
                'Fin du test de si le technicien veut créer un bon de prêt'

Else

    '? Range("E3").Select
    '? Range("E3") = "=IF(RC[-2]="""","""",VLOOKUP(RC[-2],Piece!C[-4]:C[-2],2,FALSE))"
    Range("C3:C5,C8,E6,E8").Select
    Selection.ClearContents
    Range("C3").Select
    Exit Sub
    
End If

'____________________________________________________________________________________________________________________'
                'Fermer le fichier "Tampon.xlsm" en sauvegardant
            
Windows(tampon).Activate
Range("A2").Select

With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
    
Selection.NumberFormat = "General"

'Activer autre fois la protection du fichier
Windows(tampon).Activate
ActiveSheet.Protect "spr", UserInterfaceOnly:=True, AllowFiltering:=True

Workbooks(tampon).Close SaveChanges:=True

'____________________________________________________________________________________________________________________'

'Afficher les opérations de la macro
Application.ScreenUpdating = True

'Temps d'execution
'seconds = Round(Timer - start, 2)
'MsgBox "Temps d'exécution : " & seconds & " secondes"

'Message box pour confirmer la création du bon de prêt
MsgBox ("Le bon de prêt a bien été enregistré.")

'Fermer le fichier "Bon_pret.xlsm" sans sauvegarder
Workbooks(bonPret).Activate
Range("C3:C5,C8,E6,E8").Select
Selection.ClearContents
Range("C3").Select

Worksheets("Bon_pret").Protect UserInterfaceOnly:=True, Password:="spr", AllowFiltering:=True
'Workbooks("Bon_pret.xlsm").Close SaveChanges:=False

'Revenir à l'accueil excel
If Not (FichOuvert(pret)) Then
    Workbooks.Open Filename:=chemin & "\" & pret
End If
Windows(pret).Activate

Workbooks(bonPret).Close SaveChanges:=False

End Sub
