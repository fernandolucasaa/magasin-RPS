Attribute VB_Name = "Module1"
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

