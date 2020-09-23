Attribute VB_Name = "mdlRamUsage"
Public Function RamUsage(Optional strProcess As String = "") As Double
    If strProcess = "" Then strProcess = UCase(App.EXEName) & ".EXE" 'Will count the current application as the process if no arguments given
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & strProcess & "'")
    For Each objProcess In colProcessList
           RamUsage = objProcess.workingSetSize / 1024
    Next
End Function
Public Function PFUsage(Optional strProcess As String = "") As Double
    If strProcess = "" Then strProcess = UCase(App.EXEName) & ".EXE" 'Will count the current application as the process if no arguments given
    
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & strProcess & "'")
    For Each objProcess In colProcessList
           PFUsage = objProcess.PageFileUsage / 1024
    Next
End Function
Public Function ListAllProcesses(txtBox As TextBox)
    With txtBox
        .Text = ""
        Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
        Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process")
            .Text = .Text & "Exename   " & vbTab & "RAM Usage" & vbTab & "PageFile Usage" & vbNewLine
            .Text = .Text & "---------------   " & vbTab & "--------------------" & vbTab & "-------------------------" & vbNewLine
        For Each objProcess In colProcessList
            .Text = .Text & TabIt(objProcess.name) & vbTab & TabIt(FormatUsage(objProcess.workingSetSize / 1024) & " Kb", 8) & vbTab & FormatUsage(objProcess.PageFileUsage / 1024) & "Kb" & vbNewLine
            .SelStart = Len(.Text)
            .SelLength = 0
        Next
    End With
End Function
Public Function FormatUsage(tUsage As Double)
    If Int(tUsage) = tUsage Then
        If tUsage = 0 Then
            FormatUsage = 0
        Else
            FormatUsage = Format(tUsage, "###,###")
        End If
    Else
        FormatUsage = Format(tUsage, "###,###.#")
    End If
End Function

Public Function TabIt(sTab As String, Optional tMax As Integer = 7) As String
    If Len(sTab) <= tMax Then
        TabIt = sTab & vbTab
    ElseIf Len(sTab) > 15 Then
        TabIt = Left$(sTab, 15)
    Else
        TabIt = sTab
    End If
End Function
