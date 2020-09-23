VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWebAIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Internal declaration
Private WithEvents myIE As InternetExplorer
Attribute myIE.VB_VarHelpID = -1
Public Event OnIEQuit(AppState As Boolean)

Public Sub DisableWindow(criteria As String)
    Dim objShellWins    As New SHDocVw.ShellWindows
    Dim objExplorer     As SHDocVw.InternetExplorer
    Dim I As Integer, St, J As Integer
    
    St = Split(criteria, ",")
    I = UBound(St)
    On Error Resume Next
    For Each objExplorer In objShellWins
                DoEvents
                If TypeOf objExplorer.document Is HTMLDocument Then
                        Set myIE = objExplorer
                        For J = 0 To I
                                If InStr(LCase(myIE.LocationURL), LCase(St(J))) > 0 Then
                                        'If InStr(LCase(frmNBAss.Text1.Text), LCase(myIE.LocationURL)) = 0 Then
                                                frmNBAss.Caption = "IE Block   :  found " & myIE.LocationName
                                                Open App.Path & "\ieblock.html" For Output As #1
                                                Print #1, "<HTML>"
                                                Print #1, "<HEAD>"
                                                Print #1, "<TITLE> Web ต้องห้าม </TITLE>"
                                                Print #1, "</HEAD>"
                                                Print #1, "<BODY>"
                                                Print #1, "<H1>ท่านกำลังเข้าสู่ web ต้องห้าม <H2><BR>"
                                                Print #1, myIE.LocationName & "<BR>"
                                                Print #1, "ระบบไม่สามารถให้ท่านเข้าใช้งาน web ดังกล่าวได้ เพราะอาจเป็นอันตรายต่อเครื่องคอมพิวเตอร์"
                                                Print #1, "ที่สำคัญคือผู้ปกครองของท่านห้ามไว้"
                                                Print #1, "</BODY>"
                                                Print #1, "</HTML>"
                                                Close #1
                                                myIE.Navigate2 App.Path & "\ieblock.html"
                                                frmNBAss.Text1.Text = frmNBAss.Text1.Text & "[ " & St(J) & " ] " & myIE.LocationURL & vbCrLf
                                                'myIE.Quit
                                        'End If
                                End If
                        Next J
                End If
    Next
    Screen.MousePointer = vbDefault
End Sub
