Public Sub Sequence_Predecessor()
    Sequence_Clean
    Sequence "predecessor"
    Sequence_Show
End Sub
Public Sub Sequence_Successor()
    Sequence_Clean
    Sequence "successor"
    Sequence_Show
End Sub
Public Sub Sequence_Both()
    Sequence_Clean
    Sequence "predecessor"
    Sequence "successor"
    Sequence_Show
End Sub
Private Sub Sequence(logic As String)
'Sequences the path

    On Error GoTo EmergencyExit
    Dim T As Task

    If logic = "predecessor" Then
        For Each T In ActiveSelection.Tasks
            SequencePredecessors T
        Next T
    ElseIf logic = "successor" Then
        For Each T In ActiveSelection.Tasks
            TraceSuccessors T
        Next T
    End If
Exit Sub
EmergencyExit:
    HandlingErrors
End Sub
Private Sub SequencePredecessors(T As Task)
    Dim T2 As Task
    T.Flag19 = True

    For Each T2 In T.PredecessorTasks
        If T2.Flag19 = False Then
            SequencePredecessors T2
        End If
    Next T2
End Sub
Private Sub TraceSuccessors(T As Task)
    Dim T2 As Task
    T.Flag19 = True

    For Each T2 In T.SuccessorTasks
        If T2.Flag19 = False Then
            TraceSuccessors T2
        End If
    Next T2
End Sub
Private Sub Sequence_Clean()
    Dim T As Task
    For Each T In ActiveProject.Tasks
        If T.Flag19 = True Then T.Flag19 = False
    Next T
End Sub
Private Sub Sequence_Show()
    FilterEdit Name:="Tracing Filter", TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Flag19", test:="equals", Value:="Yes", ShowInMenu:=False, _
        ShowSummaryTasks:=True
    FilterApply Name:="Tracing Filter"
End Sub
Private Sub HandlingErrors()
    Select Case Err.Number
        Case 91
            MsgBox "The first selected row does not have a task name.", vbCritical
        Case 424
            MsgBox "The selected task is missing a task name.", vbCritical
        Case 1100
            MsgBox "This view and table combination doesn't have Outlines available. Try going to " & _
                        "View >> Data Group: Outline. If Outline is grayed out, try clicking on the task name." & _
                        vbNewLine & vbNewLine & "This error usually happens when the timeline or details pane is selected.", _
                    vbCritical, "Oops! Outline is Unavailable"
        Case 1101
            MsgBox "Make sure you are in the Gantt view sitting on a task to have this work." & vbNewLine & vbNewLine & _
                "Error#" & Str(Err.Number) & " - " & Err.Description, vbCritical, "Invalid View"
        Case Else
            MsgBox "Error#" & Str(Err.Number) & " - " & Err.Description & vbNewLine _
                    & "Line: " & Erl & vbNewLine _
                    , vbCritical
    End Select
End Sub