<h2>Create a Class Module called "clsBar"</h2>

<pre>
Option Explicit
 
'Popup objects
Private cmdBar As CommandBar
Private WithEvents cmdCopyButton As CommandBarButton
Private WithEvents cmdPasteButton As CommandBarButton
 
'Useform to use
Private fmUserform As Object
 
'Control array of textbox
Private colControls As Collection
 
'Textbox Control
Private WithEvents tbControl As MSForms.TextBox
'Adds all the textbox in the userform to use the popup bar
Sub Initialize(ByVal UF As Object)
   Dim Ctl As MSForms.Control
   Dim cBar As clsBar
   For Each Ctl In UF.Controls
      If TypeName(Ctl) = "TextBox" Then
       
         'Check if we have initialized the control array
        If colControls Is Nothing Then
            Set colControls = New Collection
            Set fmUserform = UF
            'Create the popup
           CreateBar
         End If
          
         'Create a new instance of this class for each textbox
        Set cBar = New clsBar
         cBar.AssignControl Ctl, cmdBar
         'Add it to the control array
        colControls.Add cBar
      End If
   Next Ctl
End Sub
  
Private Sub Class_Terminate()
   'Delete the commandbar when the class is destroyed
  On Error Resume Next
   cmdBar.Delete
End Sub
 
'Click event of the copy button
Private Sub cmdCopyButton_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
   fmUserform.ActiveControl.Copy
   CancelDefault = True
End Sub
 
'Click event of the paste button
Private Sub cmdPasteButton_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
   fmUserform.ActiveControl.Paste
   CancelDefault = True
End Sub
 
'Right click event of each textbox
Private Sub tbControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
      ByVal X As Single, ByVal Y As Single)
       
   If Button = 2 And Shift = 0 Then
      'Display the popup
     cmdBar.ShowPopup
   End If
End Sub
 
Private Sub CreateBar()
   Set cmdBar = Application.CommandBars.Add(, msoBarPopup, False, True)
   'We’ll use the builtin Copy and Paste controls
  Set cmdCopyButton = cmdBar.Controls.Add(ID:=19)
   Set cmdPasteButton = cmdBar.Controls.Add(ID:=22)
End Sub
 
'Assigns the Textbox and the CommandBar to this instance of the class
Sub AssignControl(TB As MSForms.TextBox, Bar As CommandBar)
   Set tbControl = TB
   Set cmdBar = Bar
End Sub
</pre>

<h2>Add Into Userform</h2>

<pre>
Dim cBar As clsBar
 
Private Sub UserForm_Initialize()
   Set cBar = New clsBar
   cBar.Initialize Me
End Sub
</pre>
