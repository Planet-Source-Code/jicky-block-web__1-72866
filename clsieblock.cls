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
                                                Print #1, "<TITLE> Web µéÍ§ËéÒÁ </TITLE>"
                                                Print #1, "</HEAD>"
                                                Print #1, "<BODY>"
                                                Print #1, "<H1>·èÒ¹¡ÓÅÑ§à¢éÒÊÙè web µéÍ§ËéÒÁ <H2><BR>"
                                                Print #1, myIE.LocationName & "<BR>"
                                                Print #1, "ÃÐººäÁèÊÒÁÒÃ¶ãËé·èÒ¹à¢éÒãªé§Ò¹ web ´Ñ§¡ÅèÒÇä´é à¾ÃÒÐÍÒ¨à»ç¹ÍÑ¹µÃÒÂµèÍà¤Ã×èÍ§¤ÍÁ¾ÔÇàµÍÃì"
                                                Print #1, "·ÕèÊÓ¤Ñ­¤×Í¼Ùé»¡¤ÃÍ§¢Í§·èÒ¹ËéÒÁäÇé"
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
