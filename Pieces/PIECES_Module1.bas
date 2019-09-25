Attribute VB_Name = "Module1"
Option Explicit
Dim Ws As Worksheet
Dim X As Integer
Dim l As Byte
Dim TabSheet() As Variant
Dim myText(15) As OLEObject

Function FichOuvert(F As String) As Boolean

On Error Resume Next
FichOuvert = Not Workbooks(F) Is Nothing

End Function

'Retourner à l'accueil
Sub Accueil()

Worksheets("resultat").Protect userinterfaceonly:=True, Password:="spr", AllowFiltering:=True, AllowUsingPivotTables:=True

'Variables pour les fichiers
Dim pieces As String, pret As String, chemin As String

pieces = "PIECES.xlsm"
pret = "pret.xlsm"

Windows(pieces).Activate
chemin = Application.ActiveWorkbook.Path

If Not (FichOuvert(pret)) Then
    Workbooks.Open Filename:=chemin & "\" & pret
End If

''Creer les filtres de la première ligne
'Rows("1:1").Select
'Selection.AutoFilter
Range("A1").Select

Windows(pret).Activate
Workbooks(pieces).Close SaveChanges:=False
'Workbooks(pieces).Close SaveChanges:=True

End Sub

'Mise à jour des donnée en utilisant le fichier "PIECE GENERIQUE.xlsx"
Sub MAJ()

Worksheets("resultat").Protect userinterfaceonly:=True, Password:="spr", AllowFiltering:=True, AllowUsingPivotTables:=True

'Afficher les opérations de la macro
Application.ScreenUpdating = False

'Variables pour les fichiers
Dim pieces As String, piecesGenerique As String, chemin As String

pieces = "PIECES.xlsm"
piecesGenerique = "PIECES GENERIQUE.xlsx"

Windows(pieces).Activate
chemin = Application.ActiveWorkbook.Path

If Not (FichOuvert(piecesGenerique)) Then
    Workbooks.Open Filename:=chemin & "\" & piecesGenerique
End If

'Calculer le nombre des lignes utilisées
Dim derlignes As Long

Windows(piecesGenerique).Activate
Sheets("resultat").Select
Range("A1").Select

derlignes = Cells(Rows.Count, 1).End(xlUp).Row 'Calcul de nombre de ligne dans le tableau
    
'Copier
Windows(piecesGenerique).Activate
Sheets("resultat").Select
Range("A2", "F" & derlignes).Copy

'Coller
Windows(pieces).Activate
Sheets("resultat").Select
Range("A2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Liberer le presse papier
Application.CutCopyMode = False

''Filtre
'Rows("1:1").Select
'Selection.AutoFilter
'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
'        , AllowSorting:=True, AllowFiltering:=True
'Range("A1").Select

'Fermer et sauvegarder
Workbooks(piecesGenerique).Close SaveChanges:=False
Windows(pieces).Activate
Range("A1").Select
Workbooks(pieces).Save

MsgBox "Le fichier 'PIECES.xlsm' a été mis à jour en utilisant le fichier 'PIECES GENERIQUE.xlsx'"
    
'Afficher les opérations de la macro
Application.ScreenUpdating = True

End Sub

'Faire un bon de prêt
Sub SortiePieces()

Worksheets("resultat").Protect userinterfaceonly:=True, Password:="spr", AllowFiltering:=True, AllowUsingPivotTables:=True

Dim r As Integer
r = Selection.Row

If MsgBox("Voullez vous faire une sortie du CMS " & Sheets("resultat").Range("A" & r).Value & " ?", vbYesNo, "RPS") = vbYes Then
    
    'Ouverture de fichier
    Dim pieces As String, bonPret As String, chemin As String

    pieces = "PIECES.xlsm"
    bonPret = "Bon_pret.xlsm"

    Windows(pieces).Activate
    chemin = Application.ActiveWorkbook.Path

    If Not (FichOuvert(bonPret)) Then
        Workbooks.Open Filename:=chemin & "\" & bonPret
    End If
    
    'Préremplir les informations (CMS et n° série)
    Workbooks(bonPret).Sheets("Bon_pret").Range("C3").Value = Workbooks(pieces).Sheets("resultat").Range("A" & r).Value 'Le CMS
    
    Windows(bonPret).Activate
    'Workbooks("PIECES.xlsm").Close SaveChanges:=True
    Workbooks("PIECES.xlsm").Close SaveChanges:=False
    
End If

End Sub

'Supprimer l'enregistrement des modifications lors de la fermeture
Sub Auto_Close()

ThisWorkbook.Saved = True 'Excel répond comme si le classeur a déjà été enregistré

End Sub
