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

'Application.DisplayFullScreen = True

End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)

Application.DisplayAlerts = False
Sheets(Sheets.Count).Delete
Application.DisplayAlerts = True

End Sub

Private Sub Workbook_Open()

Worksheets("Retour_pret").Protect userinterfaceonly:=True, Password:="spr"
Worksheets("Retour_pret").ScrollArea = "A1:E16"

'With Application.CommandBars("Standard").Controls("Enregistrer").Enabled = False
'End With

'With Application.CommandBars("Worksheet Menu Bar").Controls("Fichier").Enabled = False
'End With

Application.OnKey "^s", "" 'Desactive le raccourci clavier Ctrl+S

End Sub

