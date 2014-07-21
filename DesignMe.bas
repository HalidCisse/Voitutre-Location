Attribute VB_Name = "Designer"
Option Explicit
Declare Sub InitCommonControls Lib "comctl32.dll" ()
Global m_cFlatten() As cFlatControl
Global m_iCount As Long

Function InitMe()
  InitCommonControls
End Function

 Function Design(FormD As Form)
Dim ctl As Control
Dim bDoIt As Boolean
Dim i As Long

   For Each ctl In FormD.Controls
      bDoIt = False
      If TypeOf ctl Is ComboBox Then
         bDoIt = True
      ElseIf TypeOf ctl Is TextBox Then
         'ctl.Text = ctl.Name & ", vbAccelerator"
         bDoIt = True
      ElseIf TypeOf ctl Is PictureBox Then
         bDoIt = True
      End If
      If (bDoIt) Then
         m_iCount = m_iCount + 1
         ReDim Preserve m_cFlatten(1 To m_iCount) As cFlatControl
         Set m_cFlatten(m_iCount) = New cFlatControl
         m_cFlatten(m_iCount).Attach ctl
      End If
     ' If TypeOf ctl Is ComboBox Then
         'For i = 1 To 20
         '   ctl.AddItem ctl.Name & ",Test Item " & i
         'Next i
         'ctl.ListIndex = 0
      'End If
   Next ctl
End Function
