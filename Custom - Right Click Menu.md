[Original Post (here)](https://msdn.microsoft.com/en-us/library/office/gg469862(v=office.14).aspx)
[Followup SO (here)](https://stackoverflow.com/questions/25127530/right-click-to-copy-not-functioning-properly)

<h2>Add as a Private Module</h2>
<pre>
Sub AddToCellMenu()
    Dim ContextMenu As CommandBar
    Dim MySubMenu As CommandBarControl

    ' Delete the controls first to avoid duplicates.
    Call DeleteFromCellMenu

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Add one built-in button(Save = 3) to the Cell context menu.
    ContextMenu.Controls.Add Type:=msoControlButton, ID:=3, before:=1

    ' Add one custom button to the Cell context menu.
    With ContextMenu.Controls.Add(Type:=msoControlButton, before:=2)
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "ToggleCaseMacro"
        .FaceId = 59
        .Caption = "Toggle Case Upper/Lower/Proper"
        .Tag = "My_Cell_Control_Tag"
    End With

    ' Add a custom submenu with three buttons.
    Set MySubMenu = ContextMenu.Controls.Add(Type:=msoControlPopup, before:=3)

    With MySubMenu
        .Caption = "Case Menu"
        .Tag = "My_Cell_Control_Tag"

        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "UpperMacro"
            .FaceId = 100
            .Caption = "Upper Case"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "LowerMacro"
            .FaceId = 91
            .Caption = "Lower Case"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "ProperMacro"
            .FaceId = 95
            .Caption = "Proper Case"
        End With
    End With

    ' Add a separator to the Cell context menu.
    ContextMenu.Controls(4).BeginGroup = True
End Sub

Sub DeleteFromCellMenu()
    Dim ContextMenu As CommandBar
    Dim ctrl As CommandBarControl

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Delete the custom controls with the Tag : My_Cell_Control_Tag.
    For Each ctrl In ContextMenu.Controls
        If ctrl.Tag = "My_Cell_Control_Tag" Then
            ctrl.Delete
        End If
    Next ctrl

    ' Delete the custom built-in Save button.
    On Error Resume Next
    ContextMenu.FindControl(ID:=3).Delete
    On Error GoTo 0
End Sub

Sub ToggleCaseMacro()
    Dim CaseRange As Range
    Dim CalcMode As Long
    Dim cell As Range

    On Error Resume Next
    Set CaseRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    For Each cell In CaseRange.Cells
        Select Case cell.Value
        Case UCase(cell.Value): cell.Value = LCase(cell.Value)
        Case LCase(cell.Value): cell.Value = StrConv(cell.Value, vbProperCase)
        Case Else: cell.Value = UCase(cell.Value)
        End Select
    Next cell

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub

Sub UpperMacro()
    Dim CaseRange As Range
    Dim CalcMode As Long
    Dim cell As Range

    On Error Resume Next
    Set CaseRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    For Each cell In CaseRange.Cells
        cell.Value = UCase(cell.Value)
    Next cell

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub

Sub LowerMacro()
    Dim CaseRange As Range
    Dim CalcMode As Long
    Dim cell As Range

    On Error Resume Next
    Set CaseRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    For Each cell In CaseRange.Cells
        cell.Value = LCase(cell.Value)
    Next cell

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub

Sub ProperMacro()
    Dim CaseRange As Range
    Dim CalcMode As Long
    Dim cell As Range

    On Error Resume Next
    Set CaseRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    For Each cell In CaseRange.Cells
        cell.Value = StrConv(cell.Value, vbProperCase)
    Next cell

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub
</pre>

<h2>Add to ThisWorkbook</h2>
<pre>
Private Sub Workbook_Activate()
    Call AddToCellMenu
End Sub

Private Sub Workbook_Deactivate()
    Call DeleteFromCellMenu
End Sub
</pre>
