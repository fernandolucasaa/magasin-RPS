Attribute VB_Name = "Module1"
'Retourner à l'accueil
Sub Accueil()

'Variables pour les fichiers
Dim retourPret As String, pret As String, chemin As String

retourPret = "Retour_pret.xlsm"
pret = "pret.xlsm"

Windows(retourPret).Activate
chemin = Application.ActiveWorkbook.Path

'Effacer les données des cellules
Windows(retourPret).Activate
Range("C3,C4,C8,E6").Select
Selection.ClearContents

Range("C3").Select

If Not (FichOuvert(pret)) Then
    Workbooks.Open Filename:=chemin & "\" & pret
End If

Workbooks(retourPret).Close SaveChanges:=False
Windows(pret).Activate

End Sub

'Supprimer l'enregistrement des modifications lors de la fermeture
Sub Auto_Close()

ThisWorkbook.Saved = True 'Excel répond comme si le classeur a déjà été enregistré
'Call Accueil

End Sub
