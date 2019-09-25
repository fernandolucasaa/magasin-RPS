VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'Code to move to next cell when Enter are pressed
Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

Dim cbo As OLEObject
Dim ws As Worksheet

Set ws = ActiveSheet
Set cbo = ws.OLEObjects("ComboBox1")

cbo.Activate
'open the drop down list automatically
Me.ComboBox1.DropDown

Select Case KeyCode
    Case 13 'Enter
      ActiveCell.Offset(1, 0).Activate
    Case Else
        'do nothing
End Select

End Sub

