Attribute VB_Name = "Module1"
Function FichOuvert(F As String) As Boolean

On Error Resume Next
FichOuvert = Not Workbooks(F) Is Nothing

End Function

'Retourner à l'accueil
Sub Accueil()

Worksheets("Pret").Protect userinterfaceonly:=True, Password:="spr", AllowFiltering:=True, AllowSorting:=True

'Variables pour les fichiers
Dim tampon As String, pret As String, chemin As String

tampon = "Tampon.xlsm"
pret = "pret.xlsm"

Windows(tampon).Activate
chemin = Application.ActiveWorkbook.Path

If Not (FichOuvert(pret)) Then
    Workbooks.Open Filename:=chemin & "\" & pret
End If

''Creer les filtres de la première ligne
'Rows("1:1").Select
'Selection.AutoFilter
'Range("A1").Select

'Fermer
Windows(pret).Activate
'Workbooks(tampon).Close SaveChanges:=True
Workbooks(tampon).Close SaveChanges:=False


End Sub

'Faire le retour de prêt en utilisant le fichier "Retour_pret.xslm"
Sub retourPret()

Dim r As Integer
r = Selection.Row

If MsgBox("Voulez-vous faire le retour du CMS " & Sheets("Pret").Range("C" & r).Value & " ?", vbYesNo, "RPS") = vbYes Then
    
    'Variables pour les fichiers
    Dim tampon As String, retourPret As String, chemin As String

    tampon = "Tampon.xlsm"
    retourPret = "Retour_pret.xlsm"

    Windows(tampon).Activate
    chemin = Application.ActiveWorkbook.Path
    
    If Not (FichOuvert(retourPret)) Then
        Workbooks.Open Filename:=chemin & "\" & retourPret
    End If
    
    'Préremplir les informations
    Workbooks(retourPret).Sheets("Retour_pret").Range("C3").Value = Workbooks(tampon).Sheets("Pret").Range("C" & r).Value 'Le CMS
    Workbooks(retourPret).Sheets("Retour_pret").Range("C4").Value = Workbooks(tampon).Sheets("Pret").Range("G" & r).Value 'La quantité
    Workbooks(retourPret).Sheets("Retour_pret").Range("E6").Value = Workbooks(tampon).Sheets("Pret").Range("J" & r).Value 'L'emprunteur
        
    Workbooks(retourPret).Sheets("Retour_pret").Range("C8").Select 'Pré selection de la case
    
    'Fermer
    Windows(retourPret).Activate
    'Workbooks("Tampon.xlsm").Close SaveChanges:=True
    Workbooks(tampon).Close SaveChanges:=False
    
End If

End Sub

'Supprimer l'enregistrement des modifications lors de la fermeture
Sub Auto_Close()

ThisWorkbook.Saved = True 'Excel répond comme si le classeur a déjà été enregistré

End Sub
