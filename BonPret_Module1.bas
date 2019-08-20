Attribute VB_Name = "Module1"
'Retourner à l'accueil
Sub Accueil()

ActiveSheet.Protect UserInterfaceOnly:=True, Password:="spr"

'Variables pour les fichiers
Dim bonPret As String, pret As String, chemin As String

bonPret = "Bon_pret.xlsm"
pret = "pret.xlsm"

Windows(bonPret).Activate
chemin = Application.ActiveWorkbook.Path

'Effacer les données des cellules
Windows(bonPret).Activate
Range("C3:C5,C8,E6,E8").Select
Selection.ClearContents

Range("C3").Select

If Not (FichOuvert(pret)) Then
    Workbooks.Open Filename:=chemin & "\" & pret
End If

Workbooks(bonPret).Close SaveChanges:=False
Windows(pret).Activate

End Sub

'Supprimer l'enregistrement des modifications lors de la fermeture
Sub Auto_Close()

ThisWorkbook.Saved = True 'Excel répond comme si le classeur a déjà été enregistré

End Sub

