Attribute VB_Name = "Module2"
'Executer tous les fois que le code est modifie et sauvagarder le fichier pour enrigestrer le filtre
Sub Filtre()
Attribute Filtre.VB_ProcData.VB_Invoke_Func = " \n14"

If MsgBox("Si le filtre ne marche pas, appuyer sur oui.", vbYesNo, "RPS") = vbYes Then

    Worksheets("resultat").Protect userinterfaceonly:=True, Password:="spr"
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
    Range("A1").Select
    
End If

End Sub
