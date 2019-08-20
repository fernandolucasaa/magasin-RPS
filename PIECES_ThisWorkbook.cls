VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

'Protéger
Worksheets("resultat").Protect userinterfaceonly:=True, Password:="spr", AllowFiltering:=True, AllowUsingPivotTables:=True
Worksheets("resultat").ScrollArea = "A1:F20000"

'Creer les filtres de la première ligne
'Rows("1:1").Select
'Selection.AutoFilter
'Range("A1").Select
'Workbooks("PIECES.xlsm").Save

'
'Application.EnableEvents = False
Application.DisplayFullScreen = True

End Sub





