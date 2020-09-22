Attribute VB_Name = "finalfuncitons"
Dim com1
Dim com2
Dim tempport_scan
Dim tempproxy_scan
Dim templogin_data
Dim port_scan As Boolean
Dim proxy_scan As Boolean
Public login_data As Boolean
Public nul As Integer
Public Sub runstatus(msg As String)
running.Caption = msg
End Sub
Public Sub RUNTASK(task, ipad As String, operators)
'display the wait screen
running.Visible = True
TimeOut 1
runstatus "Clearing Scan Buffer... Please wait"

If mainform.Winsock1.Count > 1 Then
Do Until mainform.Winsock1.Count = 1
tempdel = Int(mainform.Winsock1.UBound)
Unload mainform.Winsock1(tempdel)
Loop
End If

runstatus "Scanning Address Type"
''is domain?
com1 = ipad Like "*.*"

''is ipo?
com2 = ipad Like "*.*.*"

runstatus "Parsing Operators"
''parse the operators
tempport_scan = Left(operators, 1)
If tempport_scan = 1 Then port_scan = True
If tempport_scan = 0 Then port_scan = False
'
tempproxy_scan = Right(Left(operators, 2), 1)
If tempproxy_scan = 1 Then proxy_scan = True
If tempproxy_scan = 0 Then proxy_scan = False
'
templogin_data = Right(operators, 1)
If templogin_data = 1 Then login_data = True
If templogin_data = 0 Then login_data = False

runstatus "Getting Ready to run"



ipiP = ipad
iporT = addtask.startport
If login_data = True Then
runstatus "Loging Scanner Data"
Open App.Path & "\Scanner Log.txt" For Append As #1
Write #1, "______________________________________________________"
Write #1, "Begin scans of: " & ipiP & "  " & "Portscan = " & port_scan & "   " & "Proxyscan = " & proxy_scan
Write #1, "Date<>Time  " & Date & "<>" & Time
ioi = 0
runstatus "Wating for results..."
runstatus "Recording Results in: " & App.Path & "\Scanner Log.txt"
Close #1
End If

If port_scan = True Then
For iporT = addtask.startport To addtask.stopport

Call cleanup

Load mainform.Winsock1(iporT)
If mainform.Winsock1(iporT).State <> sckClosed Then mainform.Winsock1(iporT + 1).Close
mainform.Winsock1(iporT).Connect ipiP, iporT
runstatus "Scan: " & iporT & ":" & mainform.Winsock1.Count & "  >  " & mainform.Winsock1(iporT).State
Next iporT
End If

If proxy_scan = True Then
runstatus "Looking for Proxy"
Load mainform.Winsock1(iporT)
If mainform.Winsock1(iporT).State <> sckClosed Then mainform.Winsock1(iporT + 1).Close
mainform.Winsock1(iporT).Connect ipiP, 8080
End If


runstatus "Finishing Ops"
mainform.Go = True
running.Visible = False
mainform.addbutt.Visible = True
mainform.Visible = True
End Sub

Sub portsc()



End Sub
Public Sub TimeOut(Duration As Double)
    ' standard timeout sub, causes a short pause in the code
    Dim StartTime As Double, x As Integer
    StartTime = Timer
    Do While Timer - StartTime < Duration
        x = DoEvents()
    Loop
End Sub

Public Sub cleanup()
'
End Sub
Public Sub out(message As Integer, outpt)
Write #message, output
End Sub
