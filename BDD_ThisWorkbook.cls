VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open() 'Event

'UserInterfaceOnly = True permet d'exécuter des macros et protége seulement les modifications opéréees par l'utilisateur
Worksheets("BDD").Protect userinterfaceonly:=True, Password:="spr"

End Sub

