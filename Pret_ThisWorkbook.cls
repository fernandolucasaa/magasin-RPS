VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Activate()

Application.DisplayFullScreen = True 'Masquer le ruban

End Sub

'Empecher la fermeture du fichier
Private Sub Workbook_BeforeClose(Cancel As Boolean)

MsgBox "Ne pas fermer ce fichier"
Cancel = True

End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)

Application.DisplayAlerts = False
Sh.Delete
Application.DisplayAlerts = True
'Application.DisplayAlerts = False
'Sheets(Sheets.Count).Delete
'Application.DisplayAlerts = True

End Sub

Private Sub Workbook_Open()
    
Application.DisplayFullScreen = True

Worksheets("pret").Protect userinterfaceonly:=True, Password:="spr", AllowFiltering:=True
Worksheets("pret").ScrollArea = "A1"

'
'    With Application.CommandBars("Standard").Controls("Enre&gistrer").Enabled = False
'    End With
'    With Application.CommandBars("Worksheet Menu Bar").Controls("Fichier").Enabled = False
'    End With
'
'    Application.OnKey "^s", ""

   ' Application.EnableEvents = False    '?
    

End Sub
