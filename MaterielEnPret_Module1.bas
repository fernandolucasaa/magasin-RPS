Attribute VB_Name = "Module1"
'Function FichOuvert(F As String) As Boolean
'
'    On Error Resume Next
'    FichOuvert = Not Workbooks(F) Is Nothing
'End Function
'
'
'Sub macro()
'    Dim derlignes As Integer
'    Dim Valeur_Cherchee As String, AdresseTrouvee As String
'    Dim PlageDeRecherchededonnees As Range
'    Dim Valeur_reCherchee As String, AdressereTrouvee As String
'
'    Application.ScreenUpdating = False
'
'    '____________________________________________________________________________________________________________________'
'                'Test si le fichier est ouvert ou non'
'    Dim Fichier_Piece As String
'    Dim chemin_piece As String
'    Fichier_Piece = "Tampon.xlsm"
'
'
'    'Worksheets("Matériel En Prêt").Unprotect userinterfaceonly:=True, Password:="spr"
'
'
'    If FichOuvert(Fichier_Piece) Then
'    'Active la fenêtre '
'    Windows("Tampon.xlsm").Activate
'
'    Else
'    'Ouverture du fichier '
'    Workbooks.Open Filename:=chemin_piece & "T:\MSP\Boite_aux_lettres\Magasin\Tampon.xlsm"
'    Sheets("Pret").Select
'    End If
'
'
'    If Workbooks("Tampon.xlsm").ReadOnly = True Then
'      MsgBox ("Attention le fichier est en lecture seule, merci de fermer le fichier sur le poste concerné")
'      Exit Sub
'    Else
'
'    End If
'
'
'    'Active la fenêtre Retour_pret'
'    Windows("Tampon.xlsm").Activate
'    'Activer l'onglet "Retour_pret"'
'    Sheets("Pret").Select
'    ActiveSheet.Unprotect "spr"
'    derlignes = ActiveSheet.UsedRange.Rows.Count
'
'    For N = 2 To derlignes
'                'Active la fenêtre Retour_pret'
'                Windows("Tampon.xlsm").Activate
'                'Activer l'onglet "Retour_pret"'
'                Sheets("Pret").Select
'        If Range("A" & N).Select <> "" Then
'
'                'Récupération de la valeur de la cellule dans la feuille'
'                Valeur_Cherchee = Range("A" & N).Value
'                'Active la fenêtre Pret'
'                Windows("Matériel En Prêt .xlsm").Activate
'                'Active la feuille "Pet"'
'                Sheets("Pret").Select
'                'Sélectionne la plage de données dans laquelle on cherche la valeur dans la colonne 3 dans la feuille '
'                Set PlageDeRecherche = ActiveSheet.Columns(1)
'                'méthode find, ici on cherche la valeur exacte (LookAt:=xlWhole)
'                Set Trouve = PlageDeRecherche.Cells.Find(what:=Valeur_Cherchee, LookAt:=xlWhole)
'
'    'traitement de l'erreur possible : Si on ne trouve rien :
'            If Trouve Is Nothing Then
'                'ici, traitement pour le cas où la valeur n'est pas trouvée'
'                AdresseTrouvee = Valeur_Cherchee & " n'est pas présent dans " & PlageDeRecherche.Address
'
'                    'Active la fenêtre Retour_pret'
'                    Windows("Tampon.xlsm").Activate
'                    'Activer l'onglet "Retour_pret"'
'                    Sheets("Pret").Select
'                    Range("A" & N).EntireRow.Select
'                    Selection.Copy
'                    Windows("Matériel En Prêt .xlsm").Activate
'                    'Activer l'onglet "Retour_pret"'
'                    Sheets("Pret").Select
'                    Rows("2:2").Select
'                    Selection.Insert Shift:=xlDown
'                    Range("A2").Select
'
'            Else
'                'Remplissage de la date de retour de prêt'
'                'ici, traitement pour le cas où la valeur est trouvée'
'                AdresseTrouvee = Trouve.Address
'
'                'Active la fenêtre Retour_pret'
'                Windows("Tampon.xlsm").Activate
'                'Activer l'onglet "Retour_pret"'
'                Sheets("Pret").Select
'
'                If Range("M" & N).Value <> "" Then
'                numLigne2 = Trouve.Row
'                'Active la fenêtre Retour_pret'
'                Windows("Tampon.xlsm").Activate
'                'Activer l'onglet "Retour_pret"'
'                Sheets("Pret").Select
'                'Copie la cellule '
'                Range("M" & N).Copy
'                'Active la fenêtre Pret'
'                Windows("Matériel En Prêt .xlsm").Activate
'                'Activer l'onglet "pret"'
'                Sheets("Pret").Select
'                'Sélectionne la ligne du CMS et la colonne L'
'                Range("M" & Trouve.Row).Select
'                'Sélectionne la cellule et la colle'
'                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                    :=False, Transpose:=False
'
'                'Active la fenêtre Retour_pret'
'                Windows("Tampon.xlsm").Activate
'                'Activer l'onglet "Retour_pret"'
'                Sheets("Pret").Select
'                'Copie la cellule '
'                Range("N" & N).Copy
'                'Active la fenêtre Pret'
'                Windows("Matériel En Prêt .xlsm").Activate
'                'Activer l'onglet "pret"'
'                Sheets("Pret").Select
'                'Sélectionne la ligne du CMS et la colonne L'
'                Range("N" & Trouve.Row).Select
'                'Sélectionne la cellule et la colle'
'                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                    :=False, Transpose:=False
'
'                Else
'                End If
'                End If
'        Else
'
'        End If
'
'    Next
'
'    Application.ScreenUpdating = True
'
'    Dim derlig, X As Integer
'
'    derlig = ActiveSheet.UsedRange.Rows.Count
'
'    'Active la fenêtre Retour_pret'
'    Windows("Tampon.xlsm").Activate
'    'Activer l'onglet "Retour_pret"'
'    Sheets("Pret").Select
'    ActiveSheet.Range("$A$1:$AA$12").AutoFilter Field:=13, Criteria1:="<>"
'    Range("A1").Select
'    X = Range("A1:A" & derlig).SpecialCells(xlCellTypeVisible).Count
'
'        If X = 1 Then
'    ActiveSheet.ShowAllData
'    Range("A1").Select
'        Else
'    ActiveSheet.UsedRange.Select
'    Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1, Selection.Columns.Count).Select
'    Selection.Delete Shift:=xlUp
'    ActiveSheet.ShowAllData
'    Range("A1").Select
'        End If
'
'    Workbooks("Tampon.xlsm").Close SaveChanges:=True
'    Workbooks("Matériel En Prêt .xlsm").Close SaveChanges:=True
'
'
'End Sub
'
Function FichOuvert(F As String) As Boolean

On Error Resume Next
FichOuvert = Not Workbooks(F) Is Nothing

End Function

'Mettre à jour le fichier en utilisant le fichier "Tampon". Le code realise les fonctions suivantes
' - si les prets du "Tampon" ne sont pas ajoutes, on les ajoute
' - si les prets sont deja ajoutes, on copie les date de retour
' - on supprime (filtre) les pret deja retournes dans le fichier "Tampon"
' - on filtre les pret du fichier "Materiel En Pret" pour montrer que les fichiers pas retournes
Sub MAJ()

'Augmenter la vitesse de calcul
Application.ScreenUpdating = False
Application.DisplayStatusBar = False

'Calculer le temps de calcul
Dim start As Double
Dim seconds As Double

start = Timer

'Variables pour les fichiers
Dim materielEnPret As String, tampon As String

materielEnPret = "Matériel En Prêt .xlsm"
tampon = "Tampon.xlsm"

'Vérifier si le fichier est ouvert et disponible pour faire des modifications
If Not (FichOuvert(tampon)) Then
    Workbooks.Open Filename:="T:\MSP\Boite_aux_lettres\Magasin\" & tampon
End If
    
If Workbooks(tampon).ReadOnly = True Then
    MsgBox ("Attention le fichier est en lecture seule, merci de fermer le fichier sur le poste concerné")
    Exit Sub
End If

'Variables pour les feuiles
Dim tampon_Pret As Worksheet, materielEnPret_Pret As Worksheet

Set materielEnPret_Pret = Workbooks(materielEnPret).Worksheets("Pret")
Set tampon_Pret = Workbooks(tampon).Worksheets("Pret")

'Retirer la protection du fichier
tampon_Pret.Unprotect "spr"

Dim derlignes As Integer
derlignes = tampon_Pret.UsedRange.Rows.Count

Dim x As Integer, y As Integer, z As Integer, w As Integer
x = 2 'position où ajouter une nouvelle ligne
y = 0 'compteur des lignes actualisées
z = 0 'prets deja actualises
w = 0 'prets problematiques

'Retirer le filtre de "Materiel En Pret"
Workbooks(materielEnPret).Worksheets("Pret").Activate
If ActiveSheet.FilterMode Then
    ActiveSheet.ShowAllData
End If

Dim pret As String
Dim dateRetour As String

'Remplir le fichier "Matériel En Pret"
For N = 2 To derlignes

    pret = tampon_Pret.Range("A" & N)
    
    If pret <> "" Then
    
        Set PlageDeRecherche = materielEnPret_Pret.Columns(1)
        Set Trouve = PlageDeRecherche.Cells.Find(what:=pret, LookAt:=xlWhole)
        
        If Trouve Is Nothing Then   'Si pret n'est pas encore ajouté, on copie et ajoute une nouvelle ligne
                   
            Application.DisplayAlerts = False 'quand on copie, on copie des formules aussi et des references
            
            tampon_Pret.Range("A" & N).EntireRow.Copy
            materielEnPret_Pret.Rows(x).Insert Shift:=xlDown
            
            x = x + 1
            
            Application.DisplayAlerts = True
        
        Else 'Si pret déjà ajouté, on copie la date de retour et apres on suprime le pret dans le fichier "Tampon"

            AdresseTrouvee = Trouve.Address
            dateRetour = tampon_Pret.Range("M" & N).Value
            
            Dim dateRetour_materiel As String
            dateRetour_materiel = materielEnPret_Pret.Range("M" & Trouve.Row).Value

            'Si on a la date de retour dans le "Tampon" mais pas dans le "Materiel En Pret"
            If dateRetour <> "" And materielEnPret_Pret.Range("M" & Trouve.Row).Value = "" Then

                tampon_Pret.Range("M" & N).Copy
                materielEnPret_Pret.Range("M" & Trouve.Row).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                materielEnPret_Pret.Range("N" & Trouve.Row).Value = tampon_Pret.Range("N" & N).Value    'Type Retour

                y = y + 1 'prets actualises

            'Si la date de retour dans le "Tampon" et "Materiel En Pret" sont egals (soit eguals a vide, soit la meme date)
            ElseIf dateRetour = materielEnPret_Pret.Range("M" & Trouve.Row).Value Then
                
                'rien a faire
                z = z + 1

            Else
            
                w = w + 1
                'MsgBox "Verifiez pret " & materielEnPret_Pret.Range("A" & Trouve.Row).Value & " , (" & tampon_Pret.Range("A" & N) & ")"

            End If
            
        End If
        
    End If

Next

'Filtrer, suprimer les prets déjà retournés du fichier "Tampon"
tampon_Pret.Range("A1").AutoFilter Field:=13, Criteria1:="" 'Field : date de retour, Criteria1 : montrer que les pret qui ont la cellule vide

'Protéger et fermer
tampon_Pret.Protect "spr", UserInterfaceOnly:=True, AllowFiltering:=True
Workbooks(tampon).Close SaveChanges:=True

'Mettre le filtre
Workbooks(materielEnPret).Worksheets("Pret").Activate
If Not (ActiveSheet.FilterMode) Then
   ActiveSheet.Range("M1").AutoFilter Field:=13, Criteria1:=""
End If

'Sauvegarder
Workbooks(materielEnPret).Activate
Range("A2").Select
Workbooks(materielEnPret).Save

'Restaurer
Application.ScreenUpdating = True
Application.DisplayStatusBar = True

seconds = Round(Timer - start, 2)
MsgBox "Temps d'execution : " & seconds & " secondes. " & Chr(13) & (derlignes - 1) & " prets verifies" & Chr(13) & (x - 2) & " prets ajoutes" & Chr(13) & y & " prets actualises" & Chr(13) & z & " prets deja actualises" & Chr(13) & w & " lignes problematiques (doublees)."

End Sub

Sub verifier()

'Variables pour les fichiers
Dim materielEnPret As String
materielEnPret = "Matériel En Prêt .xlsm"

'Variables pour les feuiles
Dim tampon_Pret As Worksheet
Set materielEnPret_Pret = Workbooks(materielEnPret).Worksheets("Pret")

Dim verifier As String, flag As Integer
flag = 0

If MsgBox("Voulez-vous verifiez si il y a des lignes doublées ?", vbYesNo, "RPS") = vbYes Then

    Dim i As Integer, derlignes2 As Integer, valeur As String
    derlignes2 = materielEnPret_Pret.UsedRange.Rows.Count
    
    For M = 2 To derlignes2
    
        valeur = materielEnPret_Pret.Range("A" & M).Value
        If valeur <> "" Then
            i = Application.WorksheetFunction.CountIf(Range("A2:A3000"), valeur)
            If i > 1 Then
                flag = 1
                verifier = verifier & valeur & " "
            End If
        End If
        
    Next

End If

If flag = 0 Then
    MsgBox "Fin d'execution. Pas de prets doublés."
Else
    MsgBox "Fin d'execution. Prets doublés :" & Chr(13) & verifier
End If

End Sub

