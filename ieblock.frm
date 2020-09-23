VERSION 5.00
Begin VB.Form frmIEblock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IE Block"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "ieblock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   3540
      Top             =   2100
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Text            =   "Game,Sex,xxx,nude,fuck"
      Top             =   120
      Width           =   7275
   End
   Begin VB.TextBox Text1 
      Height          =   4155
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   480
      Width           =   8235
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   855
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   8175
      TabIndex        =   1
      Top             =   4920
      Width           =   8235
   End
End
Attribute VB_Name = "frmIEblock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
        Dim St As String
        
        If Command1.Caption = "Start" Then
                Timer1.Interval = 10000
                Command1.Caption = "Stop"
                Text2.Enabled = False
        Else
                St = InputBox("¡ÃØ³ÒãÊèÃËÑÊ¼èÒ¹")
                If St = "iloveujicky" Then
                        Timer1.Interval = 0
                        Text2.Enabled = True
                        Command1.Caption = "Start"
                End If
        End If
End Sub
Private Sub disable_taskman()
        Dim a
        a = Shell("REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System /v DisableTaskMgr /t REG_DWORD /d 1 /f", vbHide)
End Sub
 
Private Sub enable_taskman()
        Dim a
        a = Shell("REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System /v DisableTaskMgr /t REG_DWORD /d 0 /f", vbHide)
End Sub

Private Sub Form_Load()
        Dim St As String
        Me.Hide
        If Dir(App.Path & "\lastcriteria.txt") <> "" Then
                Open App.Path & "\lastcriteria.txt" For Input As #1
                If Not EOF(1) Then
                        Input #1, St
                        Text2.Text = St
                End If
                Close #1
        End If
        disable_taskman
        Command1_Click
End Sub

Private Sub Form_Initialize()
'This gets Loaded when your form starts
try.cbSize = Len(try)
try.hwnd = Me.hwnd
try.uId = vbNull
try.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
try.uCallBackMessage = WM_MOUSEMOVE

'To Change the Icon Displayed in the systray
'Change the Forms Icon
'This uses whatever Icon the Form Displays
try.hIcon = Me.Icon

'Tool Tip
try.szTip = "»éÍ§¡Ñ¹à´ç¡´×éÍáÍºàÅè¹à¹ç·ä§" & vbNullChar

Call Shell_NotifyIcon(NIM_ADD, try)
Call Shell_NotifyIcon(NIM_MODIFY, try)

'If u just want the systay icon to appear at start Hide the Form
'Me.Hide
End Sub

'Right Click and Dbl Click to launch an event

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Select Case X
        Case 7725:    'Dbl Left Click
            Me.Show
        End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If Command1.Caption = "Stop" Then
                'Caption = "IE Block  »Ô´â»Ãá¡ÃÁäÁèä´éËÃÍ¡äÍéË¹Ù"
                Me.Hide
                Cancel = True
        Else
                If Text1.Text <> "" Then
                        Open App.Path & "\log" & Format(Date, "MMddyyyy") & ".txt" For Append As #1
                        Print #1, Text1.Text
                        Close #1
                End If
                Open App.Path & "\lastcriteria.txt" For Output As #1
                Print #1, Text2.Text
                Close #1
                enable_taskman
                End
        End If
End Sub

Private Sub Timer1_Timer()
        Dim ie As New clsWebAIA
        Caption = "IE Block   :  Scanning..."
        ie.DisableWindow Text2.Text
        Set ie = Nothing
        Caption = "IE  Block"
End Sub
