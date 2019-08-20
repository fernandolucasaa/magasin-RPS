VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Numéro_de_prêt 
   Caption         =   "Numéro de prêt"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "RetourPret_Numéro_de_prêt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Numéro_de_prêt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then cancel = True
End Sub

Private Sub CommandButton1_Click()

ligne = Numéro_de_prêt.ComboBox1.Value
    
    'Active la fenêtre Pret'
    Windows("Tampon.xlsm").Activate
    'active la feuille "pret"'
    Sheets("Pret").Select
    
    '____________________________________________________________________________________________________________________'
                    'tester si la valeur entrée est numérique'
    If IsNumeric(ligne) Then

    'Enregistre la valeur rentrée dans la message box'
    Range("AA1").Value = ligne
     If ligne <> "" Then 'Si la valeur est différente de "" '

    'Recherche le numéro de ligne correspondant au retour de prêt'
    Valeur_reCherchee = Range("AA1").Value
    'Active la feuille "Prêt"'
    Sheets("Pret").Select
    'Dans la plage de données avec tous les doublons, on recherche le numéro de ligne de prêt'
    Set PlageDeRecherchededonnees = ActiveSheet.Columns(1)
    'méthode find, ici on cherche la valeur exacte (LookAt:=xlWhole)'
    Set Trouve = PlageDeRecherchededonnees.Cells.Find(what:=Valeur_reCherchee, LookAt:=xlWhole)
    
    'traitement de l'erreur possible : Si on ne trouve rien :
    If Trouve Is Nothing Then
    
        'ici, traitement pour le cas où la valeur n'est pas trouvée
        AdresseTrouvee = Valeur_reCherchee & " n'est pas présent dans " & PlageDeRecherchededonnees.Address
    
    Else
        'ici, traitement pour le cas où la valeur est trouvée
        AdressereTrouvee = Trouve.Address
        numLigne2 = Trouve.Row

    'Active la fenêtre Retour_pret'
    Windows("Retour_pret.xlsm").Activate
    'active la feuille "Bon_pret"'
    Sheets("Retour_Pret").Select
    'Sélectionne la cellule B2 = date'
    Range("B2").Select
    'Copie la cellule A1'
    Range("B2").Copy
    'Active la fenêtre Pret'
    Windows("Tampon.xlsm").Activate
    'active la feuille "pret"'
    Sheets("Pret").Select
    'Sélectionne la ligne du CMS et la colonne L'
    Range("M" & Trouve.Row).Select
    'Sélectionne la cellule et la colle'
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Active la fenêtre Retour_pret'
    Windows("Retour_pret.xlsm").Activate
    'active la feuille "Retour_pret"'
    Sheets("Retour_pret").Select
    'Sélectionne la cellule C8'
    Range("C8").Select
    'Copie la cellule C8 = type de retour'
    Range("C8").Copy
    'Active la fenêtre Pret'
    Windows("Tampon.xlsm").Activate
    'active la feuille "pret"'
    Sheets("Pret").Select
    'Sélectionne la ligne du CMS et la colonne M'
    Range("N" & Trouve.Row).Select
    'Sélectionne la cellule et la colle'
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Active la fenêtre Pret'
    Windows("Tampon.xlsm").Activate
    'active la feuille "pret"'
    Sheets("Pret").Select
    'Effacer le filtre'
    Selection.AutoFilter
    
    '____________________________________________________________________________________________________________________'
                    'Supprimer l'onglet "Doublon" sans message'
    'Active la fenêtre Pret'
    Windows("Tampon.xlsm").Activate
    'active la feuille "pret"'
    Sheets("Pret").Select
    'Sélectionne la cellule A1'
    Range("A1").Select
    'Effacer le filtre'
    Selection.AutoFilter
    'Sélectionne l'onglet "Doublon"'
    Sheets("Doublon").Select
    'Supprime l'onglet "Doublon" sans message'
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Doublon").Delete
    Application.DisplayAlerts = True
    'Active la feuille "pret"'
    Sheets("Pret").Select

    Unload Numéro_de_prêt
        End If
      End If
     End If

End Sub

Private Sub CommandButton2_Click()

    '____________________________________________________________________________________________________________________'
                    'Supprimer l'onglet "Doublon" sans message'
    'Active la fenêtre Pret'
    Windows("Pret.xlsm").Activate
    'active la feuille "pret"'
    Sheets("Pret").Select
    'Sélectionne la cellule A1'
    Range("A1").Select
    'Effacer le filtre'
    Selection.AutoFilter
    'Sélectionne l'onglet "Doublon"'
    Sheets("Doublon").Select
    'Supprime l'onglet "Doublon" sans message'
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Doublon").Delete
    Application.DisplayAlerts = True
    'Active la feuille "pret"'
    Sheets("Pret").Select

    'Active la fenêtre Pret'
    Windows("Retour_Pret.xlsm").Activate
    'active la feuille "pret"'
    Sheets("Retour_Pret").Select

    Unload Numéro_de_prêt
    Exit Sub
End Sub
