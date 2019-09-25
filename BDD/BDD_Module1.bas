Attribute VB_Name = "Module1"
Function FichOuvert(F As String) As Boolean

On Error Resume Next
FichOuvert = Not Workbooks(F) Is Nothing

End Function

'Mettre à jour les feuils "BDD" des fichiers "Bon_pret.xlsm" et "Retour_pret.xlsm"
Sub MAJ()

'Calculer le temps de calcul
Dim start As Double, seconds As Double
start = Timer

'Variables pour les fichiers
Dim bonPret As String, retourPret As String, bdd As String, chemin As String

bonPret = "Bon_pret.xlsm"
retourPret = "Retour_pret.xlsm"
bdd = "BDD2.xlsm"

Windows(bdd).Activate
chemin = Application.ActiveWorkbook.Path
'MsgBox chemin

'Augmenter la vitesse de calcul
Application.ScreenUpdating = False

'_____________________________________________________________________________________________________________'
'                    'Mettre à jour la base des donées de "Bon_pret.xlsm"

If Not (FichOuvert(bonPret)) Then
    Workbooks.Open Filename:=chemin & "\" & bonPret
End If

If Workbooks(bonPret).ReadOnly = True Then
    MsgBox ("Attention le fichier 'Bon_pret.xlsm' est en lecture seule, merci de fermer le fichier sur le poste concerné avant de continuer, car on va faire des modifications dans celui-ci")
    Exit Sub
End If

'Effacer les cellules
Windows(bonPret).Activate
Worksheets("BDD").Visible = True
Worksheets("BDD").Select

Dim derlignes As Long, derlignes2 As Long
derlignes = Cells(Rows.Count, 1).End(xlUp).Row 'Trouver la derniere ligne non vide

Range("A2", "H" & derlignes).ClearContents

'Copier
Windows(bdd).Activate
Worksheets("BDD").Select
derlignes2 = Cells(Rows.Count, 1).End(xlUp).Row 'Trouver la derniere ligne non vide
Range("A3", "H" & derlignes2).Copy

'Coller
Windows(bonPret).Activate
Worksheets("BDD").Select
Range("A2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Cacher la feuille et sauvagarder
Range("A1").Select
Worksheets("BDD").Visible = xlSheetVeryHidden
Worksheets("Bon_pret").Select
Range("C3").Select
Workbooks(bonPret).Save
Workbooks(bonPret).Close

Windows(bdd).Activate
Range("A1").Select
Application.CutCopyMode = False 'vider le presse-papier

'_____________________________________________________________________________________________________________'
'                    'Mettre à jour la base des donées de "Retour_pret.xlsm"

If Not (FichOuvert(retourPret)) Then
    Workbooks.Open Filename:=chemin & "\" & retourPret
End If

If Workbooks(retourPret).ReadOnly = True Then
    MsgBox ("Attention le fichier 'Retour_pret.xlsm' est en lecture seule, merci de fermer le fichier sur le poste concerné avant de continuer, car on va faire des modifications dans celui-ci")
    Exit Sub
End If

'Effacer les cellules
Windows(retourPret).Activate
Worksheets("BDD").Visible = True
Worksheets("BDD").Select

derlignes = Cells(Rows.Count, 1).End(xlUp).Row 'Trouver la derniere ligne non vide

Range("A2", "H" & derlignes).ClearContents

'Copier
Windows(bdd).Activate
Worksheets("BDD").Select
derlignes2 = Cells(Rows.Count, 1).End(xlUp).Row 'Trouver la derniere ligne non vide
Range("A3", "H" & derlignes2).Copy

'Coller
Windows(retourPret).Activate
Worksheets("BDD").Select
Range("A2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Cacher la feuille et sauvagarder
Range("A1").Select
Worksheets("BDD").Visible = xlSheetVeryHidden
Worksheets("Retour_pret").Select
Range("C3").Select
Workbooks(retourPret).Save
Workbooks(retourPret).Close

Windows(bdd).Activate
Range("A1").Select
Application.CutCopyMode = False 'vider le presse-papier

'_____________________________________________________________________________________________________________'
'

Workbooks(bdd).Save

'Afficher les opérations de la macro
Application.ScreenUpdating = True

'Worksheets("BDD").Protect userinterfaceonly:=True, Password:="spr"

seconds = Round(Timer - start, 2)

MsgBox ("Les bases des données des fichiers 'Bon_pret' et 'Retour_pret' ont été actualisés!")
MsgBox "Temps d'exécution : " & seconds & " secondes"

End Sub

'Supprimer l'enregistrement des modifications lors de la fermeture
Sub Auto_Close()

ThisWorkbook.Saved = True 'Excel répond comme si le classeur a déjà été enregistré

End Sub


