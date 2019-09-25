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
Windows("Tampon.xlsm").Activate
ActiveSheet.Protect "spr", userinterfaceonly:=True, AllowFiltering:=True, AllowUsingPivotTables:=True

'Creer les filtres de la première ligne
'Rows("1:1").AutoFilter
'Workbooks("Tampon.xlsm").Save
'Worksheets("Pret").ScrollArea = "A1:N200"

Windows("Tampon.xlsm").Activate
Range("A2").Select

'Workbooks("pret.xlsm").Close SaveChanges:=True
'Application.EnableEvents = False
Application.DisplayFullScreen = True

End Sub
