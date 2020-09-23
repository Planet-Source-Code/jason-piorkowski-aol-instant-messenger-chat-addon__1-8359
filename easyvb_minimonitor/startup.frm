VERSION 5.00
Object = "{DBDB8EEA-6091-11D3-A321-C01F4AC10000}#1.8#0"; "CHATMON.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Mini-Monitor"
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer colorZ 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   1200
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      MaxLength       =   200
      TabIndex        =   0
      Text            =   "please wait . . ."
      Top             =   2160
      Width           =   6735
   End
   Begin VB.Timer docker 
      Interval        =   10
      Left            =   2640
      Top             =   1200
   End
   Begin VB.Timer Stype 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2520
      Top             =   360
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4440
      Top             =   600
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1680
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3720
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4320
      Top             =   1080
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1935
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3413
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"startup.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer About 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   1560
   End
   Begin FlaimChat.Flaim Flaim 
      Left            =   0
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "| Norm View"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mini-Monitor v1.5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "| On Win"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.Label value 
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      ToolTipText     =   "Value Command, Numbers that call certain functions."
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Max"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "| Sounds On |"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Timer()
Flaim.Enabled = False
Text1.Text = "About Mini Monitor . . . "
Timer1.Enabled = True
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
   Text1.SelStart = Len(Text1)
Text2.Text = "Please Wait..."
Text2.Enabled = False
Text1.Text = Text1.Text + "M"
Pause 0.1
Text1.Text = Text1.Text + "i"
Pause 0.1
Text1.Text = Text1.Text + "n"
Pause 0.1
Text1.Text = Text1.Text + "i"
Pause 0.1
Text1.Text = Text1.Text + "-"
Pause 0.1
Text1.Text = Text1.Text + "M"
Pause 0.1
Text1.Text = Text1.Text + "o"
Pause 0.1
Text1.Text = Text1.Text + "n"
Pause 0.1
Text1.Text = Text1.Text + "i"
Pause 0.1
Text1.Text = Text1.Text + "t"
Pause 0.1
Text1.Text = Text1.Text + "o"
Pause 0.1
Text1.Text = Text1.Text + "r"
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "v"
Pause 0.1
Text1.Text = Text1.Text + "1"
Pause 0.1
Text1.Text = Text1.Text + "."
Pause 0.1
Text1.Text = Text1.Text + "5"
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "-"
Pause 0.1
Timer5.Enabled = True
Text1.Text = Text1.Text + " "
Text1.Text = Text1.Text + "A"
Text1.Text = Text1.Text + "d"
Text1.Text = Text1.Text + "v"
Text1.Text = Text1.Text + "a"
Text1.Text = Text1.Text + "n"
Text1.Text = Text1.Text + "c"
Text1.Text = Text1.Text + "e"
Text1.Text = Text1.Text + "d"
Text1.Text = Text1.Text + " "
Text1.Text = Text1.Text + "E"
Text1.Text = Text1.Text + "d"
Text1.Text = Text1.Text + "i"
Text1.Text = Text1.Text + "t"
Text1.Text = Text1.Text + "i"
Text1.Text = Text1.Text + "o"
Text1.Text = Text1.Text + "n"
Timer1.Enabled = False
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
Timer2.Enabled = False
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
Pause 1
Timer1.Enabled = True
Text1.Text = Text1.Text + "P"
Pause 0.1
Text1.Text = Text1.Text + "r"
Pause 0.1
Text1.Text = Text1.Text + "o"
Pause 0.1
Text1.Text = Text1.Text + "g"
Pause 0.1
Text1.Text = Text1.Text + "r"
Pause 0.1
Text1.Text = Text1.Text + "a"
Pause 0.1
Text1.Text = Text1.Text + "m"
Pause 0.1
Text1.Text = Text1.Text + "m"
Pause 0.1
Text1.Text = Text1.Text + "e"
Pause 0.1
Text1.Text = Text1.Text + "r"
Pause 0.1
Text1.Text = Text1.Text + ":"
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "J"
Pause 0.1
Text1.Text = Text1.Text + "a"
Pause 0.1
Text1.Text = Text1.Text + "s"
Pause 0.1
Text1.Text = Text1.Text + "o"
Pause 0.1
Text1.Text = Text1.Text + "n"
Timer1.Enabled = False
Pause 0.5
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Timer1.Enabled = True
Text1.Text = Text1.Text + "w"
Pause 0.1
Text1.Text = Text1.Text + "w"
Pause 0.1
Text1.Text = Text1.Text + "w"
Pause 0.1
Text1.Text = Text1.Text + "."
Pause 0.1
Text1.Text = Text1.Text + "E"
Pause 0.1
Text1.Text = Text1.Text + "a"
Pause 0.1
Text1.Text = Text1.Text + "s"
Pause 0.1
Text1.Text = Text1.Text + "y"
Pause 0.1
Text1.Text = Text1.Text + "V"
Pause 0.1
Text1.Text = Text1.Text + "B"
Pause 0.1
Text1.Text = Text1.Text + "."
Pause 0.1
Text1.Text = Text1.Text + "c"
Pause 0.1
Text1.Text = Text1.Text + "o"
Pause 0.1
Text1.Text = Text1.Text + "m"
Timer1.Enabled = False
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
Pause 1
Timer1.Enabled = True
Text1.Text = Text1.Text + "R"
Pause 0.1
Text1.Text = Text1.Text + "e"
Pause 0.1
Text1.Text = Text1.Text + "l"
Pause 0.1
Text1.Text = Text1.Text + "e"
Pause 0.1
Text1.Text = Text1.Text + "a"
Pause 0.1
Text1.Text = Text1.Text + "s"
Pause 0.1
Text1.Text = Text1.Text + "e"
Pause 0.1
Text1.Text = Text1.Text + ":"
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "A"
Pause 0.1
Text1.Text = Text1.Text + "p"
Pause 0.1
Text1.Text = Text1.Text + "r"
Pause 0.1
Text1.Text = Text1.Text + "i"
Pause 0.1
Text1.Text = Text1.Text + "l"
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "2"
Pause 0.1
Text1.Text = Text1.Text + "0"
Pause 0.1
Text1.Text = Text1.Text + ","
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "2"
Pause 0.1
Text1.Text = Text1.Text + "0"
Pause 0.1
Text1.Text = Text1.Text + "0"
Pause 0.1
Text1.Text = Text1.Text + "0"
Pause 0.1
Text1.Text = Text1.Text + "."
Timer1.Enabled = False
Text1.Text = Text1.Text & vbCrLf

Text1.Text = Text1.Text & vbCrLf
Pause 1.5
Timer1.Enabled = True
Text1.Text = Text1.Text + "("
Pause 0.1
Text1.Text = Text1.Text + "c"
Pause 0.1
Text1.Text = Text1.Text + ")"
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "2"
Pause 0.1
Text1.Text = Text1.Text + "0"
Pause 0.1
Text1.Text = Text1.Text + "0"
Pause 0.1
Text1.Text = Text1.Text + "0"
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "T"
Pause 0.1
Text1.Text = Text1.Text + "h"
Pause 0.1
Text1.Text = Text1.Text + "e"
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "E"
Pause 0.1
Text1.Text = Text1.Text + "a"
Pause 0.1
Text1.Text = Text1.Text + "s"
Pause 0.1
Text1.Text = Text1.Text + "y"
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "N"
Pause 0.1
Text1.Text = Text1.Text + "e"
Pause 0.1
Text1.Text = Text1.Text + "t"
Pause 0.1
Text1.Text = Text1.Text + "w"
Pause 0.1
Text1.Text = Text1.Text + "o"
Pause 0.1
Text1.Text = Text1.Text + "r"
Pause 0.1
Text1.Text = Text1.Text + "k"
Flaim.Enabled = True
Timer1.Enabled = False
About.Enabled = False
Text2.Text = ""
Text2.Enabled = True
End Sub

Private Sub colorZ_Timer()
location = Text1.Find("/join NAME")
Text1.SelStart = location
Text1.SelLength = "10"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
       
location = Text1.Find("/clear")
Text1.SelStart = location
Text1.SelLength = "6"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
   
location = Text1.Find("/hide")
Text1.SelStart = location
Text1.SelLength = "5"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
      
location = Text1.Find("/clear")
Text1.SelStart = location
Text1.SelLength = "6"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
   
location = Text1.Find("/nav URL")
Text1.SelStart = location
Text1.SelLength = "8"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
   
location = Text1.Find("/about")
Text1.SelStart = location
Text1.SelLength = "6"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
   
location = Text1.Find("/chatroom")
Text1.SelStart = location
Text1.SelLength = "9"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("/textsound")
Text1.SelStart = location
Text1.SelLength = "10"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
   
location = Text1.Find("/show")
Text1.SelStart = location
Text1.SelLength = "5"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("/version")
Text1.SelStart = location
Text1.SelLength = "8"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("/exit")
Text1.SelStart = location
Text1.SelLength = "5"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("/close")
Text1.SelStart = location
Text1.SelLength = "6"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

'Start After commands text...

location = Text1.Find("- Information on Mini-Monitor.")
Text1.SelStart = location
Text1.SelLength = "30"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Shows Chatroom Options.")
Text1.SelStart = location
Text1.SelLength = "25"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Clears The Chatroom.")
Text1.SelStart = location
Text1.SelLength = "22"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Hides The Chatroom.")
Text1.SelStart = location
Text1.SelLength = "21"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Joins new room.")
Text1.SelStart = location
Text1.SelLength = "18"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Navigate To URL.")
Text1.SelStart = location
Text1.SelLength = "18"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Toggles Text Sound.")
Text1.SelStart = location
Text1.SelLength = "21"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Shows The Chatroom.")
Text1.SelStart = location
Text1.SelLength = "21"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Your Mini Monitor Version.")
Text1.SelStart = location
Text1.SelLength = "38"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Exits Mini-Monitor & Room.")
Text1.SelStart = location
Text1.SelLength = "28"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Exits Mini-Monitor Only.")
Text1.SelStart = location
Text1.SelLength = "26"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
colorZ.Enabled = False
Flaim.Enabled = True
End Sub

Private Sub docker_Timer()
    If Me.Left < 300 Then Me.Left = 0
    If Me.Top < 300 Then Me.Top = 0
    If Val(Me.Width + Me.Left) > Val(Screen.Width - 300) Then _
    Me.Left = Val(Screen.Width - Me.Width)
    If Val(Me.Height + Me.Top) > Val(Screen.Height - 300) Then _
    Me.Top = Val(Screen.Height - Me.Height)
End Sub

Private Sub Flaim_ChatLastLine(Who As String, What As String)
If Label2.Caption = "| Sounds On |" Then
Stype.Enabled = True
   For a = Len(Who$) To 19
Add$ = Add$ & " "
Next
Text1.Text = Text1.Text & vbCrLf & Who$ & Add$ & ":   " & What$
   Text1.SelStart = Len(Text1)
   Else
      For a = Len(Who$) To 19
Add$ = Add$ & " "
Next
Text1.Text = Text1.Text & vbCrLf & Who$ & Add$ & ":   " & What$
   Text1.SelStart = Len(Text1)
   End If
End Sub

Private Sub Form_DblClick()
If Form1.Height = 250 Then
Label4.Enabled = True
Form1.Height = 2475
Text1.Visible = True
Text2.Visible = True
Else
Form1.Height = 250
Label4.Enabled = False
Form1.Top = 0
Form1.Left = 0
Text1.Visible = False
Text2.Visible = False
End If
End Sub

Private Sub Form_Load()
Text1.SelFontSize = 7
Text1.BackColor = vbBlack
Text1.SelColor = vbGreen
Flaim.SetTopmost Me.hwnd
Label1.Caption = "Mini-Monitor " & App.Major & "." & App.Minor
Text1.Text = "Initializing..."
Me.Top = 0
Me.Left = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H00C0C0C0"
Label3.ForeColor = "&H00C0C0C0"
Label4.ForeColor = "&H00C0C0C0"
Label5.ForeColor = "&H00C0C0C0"
End Sub

Private Sub Form_Resize()
On Error GoTo thispart
Label4.Left = Form1.Width - Label4.Width
Label3.Left = Form1.Width - Label4.Width - Label3.Width
Text2.Height = 315
Text1.Top = Form1.ScaleTop + 250
Text1.Height = Form1.ScaleHeight - Text2.Height - 250
Text1.Width = Form1.ScaleWidth - Text1.Left
Text2.Top = Form1.ScaleHeight - Text2.Height
Text2.Width = Form1.ScaleWidth - Text2.Left
thispart:
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Aim_Chat_Send_Link("http://www.easyvb.com/minimonitor", "Mini-Monitor</a>" & "<i> Version " & App.Major & "." & App.Minor & "." & App.Revision & " Advanced Edition UN-Loaded (beta)</i></i>")
End
End Sub

Private Sub Label1_DblClick()
If Form1.Height = 250 Then
Label4.Enabled = True
Form1.Height = 2475
Text1.Visible = True
Text2.Visible = True
Else
Form1.Height = 250
Label4.Enabled = False
Form1.Top = 0
Form1.Left = 0
Text1.Visible = False
Text2.Visible = False
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H00C0C0C0"
Label3.ForeColor = "&H00C0C0C0"
Label4.ForeColor = "&H00C0C0C0"
Label5.ForeColor = "&H00C0C0C0"
Label6.ForeColor = "&H00C0C0C0"
End Sub

Private Sub Label2_Click()
If Label2.Caption = "| Sounds On |" Then
Label2.Caption = "| Sounds Off |"
Else
Label2.Caption = "| Sounds On |"
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&HFFFFFF"
End Sub

Private Sub Label3_Click()
If Form1.Height = 250 Then
Label4.Enabled = True
Form1.Height = 2475
Text1.Visible = True
Text2.Visible = True
Else
Form1.Height = 250
Label4.Enabled = False
Form1.Top = 0
Form1.Left = 0
Text1.Visible = False
Text2.Visible = False
End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = "&HFFFFFF"
End Sub

Private Sub Label4_Click()
If Label4.Caption = "Max" Then
Label3.Enabled = False
docker.Enabled = False
Form1.WindowState = 2
Label4.Caption = "Min"
   Text1.SelStart = Len(Text1)
Else
Label3.Enabled = True
Form1.WindowState = "0"
Label4.Caption = "Max"
docker.Enabled = True
Text1.SelStart = Len(Text1)
End If

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = "&HFFFFFF"
End Sub

Private Sub Label5_Click()
If Label5.Caption = "| On Win" Then
Label5.Caption = "| Off Win"
Flaim.SetNotTopmost Me.hwnd
Else
Label5.Caption = "| On Win"
Flaim.SetTopmost Me.hwnd
End If
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = "&HFFFFFF"
End Sub

Private Sub Scom_Timer()
Call Play_WAVfile("com.wav")
Scom.Enabled = False
End Sub

Private Sub Label6_Click()
If Text1.SelFontSize = 7 Then
Label6.Caption = "| Norm View"
Form1.Width = "8900"
Form1.Height = "3600"
Text1.Text = ""
Text1.SelFontSize = 10
Text1.Text = Text1.Text + "Normal Mode Enabled"
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf

Else
Form1.Width = "6735"
Form1.Height = "2475"
Label6.Caption = "| Mini View"
Text1.Text = ""
Text1.SelFontSize = 7
Text1.Text = Text1.Text + "Mini Mode Enabled"
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf

End If
End Sub

Private Sub RefColor_Timer()
location = Text1.Find("/join NAME")
Text1.SelStart = location
Text1.SelLength = "10"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
       
location = Text1.Find("/clear")
Text1.SelStart = location
Text1.SelLength = "6"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
   
location = Text1.Find("/hide")
Text1.SelStart = location
Text1.SelLength = "5"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
      
location = Text1.Find("/clear")
Text1.SelStart = location
Text1.SelLength = "6"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
   
location = Text1.Find("/nav URL")
Text1.SelStart = location
Text1.SelLength = "8"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
   
location = Text1.Find("/about")
Text1.SelStart = location
Text1.SelLength = "6"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
   
location = Text1.Find("/chatroom")
Text1.SelStart = location
Text1.SelLength = "9"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("/textsound")
Text1.SelStart = location
Text1.SelLength = "10"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
   
location = Text1.Find("/show")
Text1.SelStart = location
Text1.SelLength = "5"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("/version")
Text1.SelStart = location
Text1.SelLength = "8"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("/exit")
Text1.SelStart = location
Text1.SelLength = "5"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("/close")
Text1.SelStart = location
Text1.SelLength = "6"
Text1.SelColor = vbRed

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

'Start After commands text...

location = Text1.Find("- Information on Mini-Monitor.")
Text1.SelStart = location
Text1.SelLength = "30"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Shows Chatroom Options.")
Text1.SelStart = location
Text1.SelLength = "25"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Clears The Chatroom.")
Text1.SelStart = location
Text1.SelLength = "22"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Hides The Chatroom.")
Text1.SelStart = location
Text1.SelLength = "21"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Joins new room.")
Text1.SelStart = location
Text1.SelLength = "18"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Navigate To URL.")
Text1.SelStart = location
Text1.SelLength = "18"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Toggles Text Sound.")
Text1.SelStart = location
Text1.SelLength = "21"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Shows The Chatroom.")
Text1.SelStart = location
Text1.SelLength = "21"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Your Mini Monitor Version.")
Text1.SelStart = location
Text1.SelLength = "38"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Exits Mini-Monitor & Room.")
Text1.SelStart = location
Text1.SelLength = "28"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0

location = Text1.Find("- Exits Mini-Monitor Only.")
Text1.SelStart = location
Text1.SelLength = "26"
Text1.SelColor = vbWhite

Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
Flaim.Enabled = True
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = "&HFFFFFF"
End Sub

Private Sub Stype_Timer()
Call Play_WAVfile("click.wav")
Stype.Enabled = False
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H00C0C0C0"
Label3.ForeColor = "&H00C0C0C0"
Label4.ForeColor = "&H00C0C0C0"
Label5.ForeColor = "&H00C0C0C0"
Label6.ForeColor = "&H00C0C0C0"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If value.Caption = "0000" Then
Call Play_WAVfile("menu.wav")
Else
End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim location As Long
   'If Left$(Text2, 1) = "/" Then
'a = InStr(1, Text2, " ")
'"/help" = Left$(Text2, a - 1)
'End If
   If KeyCode = 13 And Text2.Text = "/help" Then
   Text1.Text = ""
   Text2.Text = ""
   Flaim.Enabled = False
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
Stype.Enabled = True
Text1.Text = Text1.Text + "     /about - Information on Mini-Monitor."
Text1.Text = Text1.Text & vbCrLf
   Text1.SelStart = Len(Text1)
'Stype.Enabled = True
'Text1.Text = Text1.Text + "/away reason - Tells room your away."
'Text1.Text = Text1.Text & vbCrLf
'   Text1.SelStart = Len(Text1)
Pause 0.1
Text1.Text = Text1.Text + "     /chatroom - Shows Chatroom Options."
Text1.Text = Text1.Text & vbCrLf
   Text1.SelStart = Len(Text1)
Text1.Text = Text1.Text + "     /clear - Clears The Chatroom."
Text1.Text = Text1.Text & vbCrLf
   Text1.SelStart = Len(Text1)
   Pause 0.1
'Text1.Text = Text1.Text + "     /fontsize # - Changes Font Size Of Chatroom."
'Text1.Text = Text1.Text & vbCrLf
'   Text1.SelStart = Len(Text1)
'Pause 0.1
Text1.Text = Text1.Text + "     /hide - Hides The Chatroom."
Text1.Text = Text1.Text & vbCrLf
   Text1.SelStart = Len(Text1)
Text1.Text = Text1.Text + "     /join Name - Joins new room."
Text1.Text = Text1.Text & vbCrLf
   Text1.SelStart = Len(Text1)
Pause 0.1
Text1.Text = Text1.Text + "     /nav URL - Navigate To URL."
Text1.Text = Text1.Text & vbCrLf
   Text1.SelStart = Len(Text1)
Pause 0.1
Text1.Text = Text1.Text + "     /textsound - Toggles Text Sound."
Text1.Text = Text1.Text & vbCrLf
   Text1.SelStart = Len(Text1)
Text1.Text = Text1.Text + "     /show - Shows The Chatroom."
Text1.Text = Text1.Text & vbCrLf
   Text1.SelStart = Len(Text1)
Text1.Text = Text1.Text + "     /version - Your Mini Monitor Version."
Text1.Text = Text1.Text & vbCrLf
   Text1.SelStart = Len(Text1)
Pause 0.1
Text1.Text = Text1.Text + "     /exit - Exits Mini-Monitor & Room."
Text1.Text = Text1.Text & vbCrLf
   Text1.SelStart = Len(Text1)
Pause 0.1
Text1.Text = Text1.Text + "     /close - Exits Mini-Monitor Only."
Text1.Text = Text1.Text & vbCrLf
   Text1.SelStart = Len(Text1)
               
colorZ.Enabled = True


   ElseIf KeyCode = 13 And Text2.Text = "/clear" Then
   Flaim.ChatClear
   Text1.Text = ""
       Text2.Text = ""
   ElseIf KeyCode = 13 And Text2.Text = "/exit" Then
Unload Me
Flaim.CloseAllChats
   ElseIf KeyCode = 13 And Text2.Text = "/close" Then
Call Chat_Show
Unload Me
   ElseIf KeyCode = 13 And Text2.Text = "/show" Then
Call Chat_Show
Text2.Text = ""
   ElseIf KeyCode = 13 And Text2.Text = "/hide" Then
Call Chat_Hide
Text2.Text = ""
 ElseIf KeyCode = 13 And Text2.Text = "/chatroom" Then
 Text1.Text = Text1.Text & vbCrLf
 Text1.Text = Text1.Text & vbCrLf
 Text2.Text = ""
Stype.Enabled = True
Text1.Text = Text1.Text + "     /num - Shows Number Of People In Room."
Text1.Text = Text1.Text & vbCrLf
   Text1.SelStart = Len(Text1)
   ElseIf KeyCode = 13 And Text2.Text = "/about" Then
About.Enabled = True
    Text2.Text = ""
       ElseIf KeyCode = 13 And Text2.Text = "/version" Then
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text + "Mini-Monitor " & "Version " & App.Major & "." & App.Minor & "." & App.Revision & " Advanced Edition"
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
      Text1.SelStart = Len(Text1)
    Text2.Text = ""
       ElseIf KeyCode = 13 And Text2.Text = "/num" Then
 Text1.Text = Text1.Text & vbCrLf
  Text1.Text = Text1.Text & vbCrLf
  Text1.Text = Text1.Text + "There are " & Flaim.CountPeople & " Chatting In " & Flaim.RoomName
  Text1.Text = Text1.Text & vbCrLf
   Text1.Text = Text1.Text & vbCrLf
      Text1.SelStart = Len(Text1)
    Text2.Text = ""
   ElseIf KeyCode = 13 And Text2.Text = "/textsound" And value.Caption = "0000" Then
value.Caption = "2180"
    Text2.Text = ""
   ElseIf KeyCode = 13 And Text2.Text = "/textsound" And value.Caption = "2180" Then
value.Caption = "0000"
    Text2.Text = ""
ElseIf KeyCode = 13 And InStr(1, Text2, "/join") Then
ThePlace = InStr(1, Text2, "/join")
ToReplace = Left(Text2, ThePlace + 5)
rooM = Replace(Text2, ToReplace, "")
Flaim.CloseAllChats
Call Aim_Chat_Invite_Quick("" & rooM)
    Text2.Text = ""
ElseIf KeyCode = 13 And InStr(1, Text2, "/nav") Then
ThePlace = InStr(1, Text2, "/nav")
ToReplace = Left(Text2, ThePlace + 4)
NaV = Replace(Text2, ToReplace, "")
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text + "Now Navigating To " & NaV
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
      Text1.SelStart = Len(Text1)
If Label4.Caption = "Min" Then
Label3.Enabled = True
Form1.WindowState = "0"
Label4.Caption = "Max"
docker.Enabled = True
Text1.SelStart = Len(Text1)
Call Navagate_2_A_Site("" & NaV)
    Text2.Text = ""
    Else
    Call Navagate_2_A_Site("" & NaV)
    Text2.Text = ""
    End If
'ElseIf KeyCode = 13 And InStr(1, Text2, "/fontsize") Then
'ThePlace = InStr(1, Text2, "/fontsize")
'ToReplace = Left(Text2, ThePlace + 9)
'Size = Replace(Text2, ToReplace, "")
'On Error Resume Next
'Text1.FontSize = Size
'Text1.Text = Text1.Text & vbCrLf
'Text1.Text = Text1.Text & vbCrLf
'Text1.Text = Text1.Text + "Your Font Size Is Now " & Size
'Text1.Text = Text1.Text & vbCrLf
'Text1.Text = Text1.Text & vbCrLf
'    Text2.Text = ""
'       Text1.SelStart = Len(Text1)
   ElseIf KeyCode = 13 Then
       Flaim.ChatSend Text2.Text
    Text2.Text = ""
    Text1.SelStart = Len(Text1)
    End If
    
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H00C0C0C0"
Label3.ForeColor = "&H00C0C0C0"
Label4.ForeColor = "&H00C0C0C0"
Label5.ForeColor = "&H00C0C0C0"
Label6.ForeColor = "&H00C0C0C0"
End Sub

Private Sub Timer1_Timer()
Call Play_WAVfile("beep.wav")
End Sub

Private Sub Timer2_Timer()
Timer1.Enabled = True
Text1.Text = "M"
Pause 0.1
Text1.Text = Text1.Text + "i"
Pause 0.1
Text1.Text = Text1.Text + "n"
Pause 0.1
Text1.Text = Text1.Text + "i"
Pause 0.1
Text1.Text = Text1.Text + "-"
Pause 0.1
Text1.Text = Text1.Text + "M"
Pause 0.1
Text1.Text = Text1.Text + "o"
Pause 0.1
Text1.Text = Text1.Text + "n"
Pause 0.1
Text1.Text = Text1.Text + "i"
Pause 0.1
Text1.Text = Text1.Text + "t"
Pause 0.1
Text1.Text = Text1.Text + "o"
Pause 0.1
Text1.Text = Text1.Text + "r"
Pause 0.1
Text1.Text = Text1.Text + " "
Timer1.Enabled = False
Timer5.Enabled = True
Pause 0.1
Text1.Text = Text1.Text & "Version " & App.Major & "." & App.Minor & "." & App.Revision & " Advanced Edition"
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
Timer5.Enabled = True
Timer2.Enabled = False
Pause 2
Timer1.Enabled = True
Text1.Text = Text1.Text + "G"
Pause 0.1
Text1.Text = Text1.Text + "a"
Pause 0.1
Text1.Text = Text1.Text + "t"
Pause 0.1
Text1.Text = Text1.Text + "h"
Pause 0.1
Text1.Text = Text1.Text + "e"
Pause 0.1
Text1.Text = Text1.Text + "r"
Pause 0.1
Text1.Text = Text1.Text + "i"
Pause 0.1
Text1.Text = Text1.Text + "n"
Pause 0.1
Text1.Text = Text1.Text + "g"
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "I"
Pause 0.1
Text1.Text = Text1.Text + "n"
Pause 0.1
Text1.Text = Text1.Text + "f"
Pause 0.1
Text1.Text = Text1.Text + "o"
Pause 0.1
Text1.Text = Text1.Text + "r"
Pause 0.1
Text1.Text = Text1.Text + "m"
Pause 0.1
Text1.Text = Text1.Text + "a"
Pause 0.1
Text1.Text = Text1.Text + "t"
Pause 0.1
Text1.Text = Text1.Text + "i"
Pause 0.1
Text1.Text = Text1.Text + "o"
Pause 0.1
Text1.Text = Text1.Text + "n"
Pause 0.1
Text1.Text = Text1.Text + "."
Pause 0.1
Text1.Text = Text1.Text + "."
Pause 0.1
Text1.Text = Text1.Text + "."
Timer1.Enabled = False
Pause 0.5
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
Timer4.Enabled = True
Text1.Text = Text1.Text + "I"
Text1.Text = Text1.Text + "n"
Text1.Text = Text1.Text + "t"
Text1.Text = Text1.Text + "e"
Text1.Text = Text1.Text + "r"
Text1.Text = Text1.Text + "n"
Text1.Text = Text1.Text + "e"
Text1.Text = Text1.Text + "t"
Text1.Text = Text1.Text + " "
Text1.Text = Text1.Text + "P"
Text1.Text = Text1.Text + "r"
Text1.Text = Text1.Text + "o"
Text1.Text = Text1.Text + "t"
Text1.Text = Text1.Text + "o"
Text1.Text = Text1.Text + "c"
Text1.Text = Text1.Text + "o"
Text1.Text = Text1.Text + "l"
Text1.Text = Text1.Text + ":      "
Pause 1.5
Timer5.Enabled = True
Text1.Text = Text1.Text & Get_IP

Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
Pause 1
Timer1.Enabled = True
Text1.Text = Text1.Text + "T"
Pause 0.1
Text1.Text = Text1.Text + "y"
Pause 0.1
Text1.Text = Text1.Text + "p"
Pause 0.1
Text1.Text = Text1.Text + "e"
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "/"
Pause 0.1
Text1.Text = Text1.Text + "h"
Pause 0.1
Text1.Text = Text1.Text + "e"
Pause 0.1
Text1.Text = Text1.Text + "l"
Pause 0.1
Text1.Text = Text1.Text + "p"
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "F"
Pause 0.1
Text1.Text = Text1.Text + "o"
Pause 0.1
Text1.Text = Text1.Text + "r"
Pause 0.1
Text1.Text = Text1.Text + " "
Pause 0.1
Text1.Text = Text1.Text + "C"
Pause 0.1
Text1.Text = Text1.Text + "o"
Pause 0.1
Text1.Text = Text1.Text + "m"
Pause 0.1
Text1.Text = Text1.Text + "m"
Pause 0.1
Text1.Text = Text1.Text + "a"
Pause 0.1
Text1.Text = Text1.Text + "n"
Pause 0.1
Text1.Text = Text1.Text + "d"
Pause 0.1
Text1.Text = Text1.Text + "s"
Pause 0.1
Text1.Text = Text1.Text + "."
Timer1.Enabled = False
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
Text1.Text = Text1.Text & vbCrLf
Call Chat_Hide
Flaim.Enabled = True
Flaim.ChatClear
Call Aim_Chat_Send_Link("http://www.easyvb.com/minimonitor", "Mini-Monitor</a>" & "<i> Version " & App.Major & "." & App.Minor & "." & App.Revision & " Advanced Edition Loaded (beta)</i></i>")
Text2.Text = ""
Text2.Enabled = True

End Sub

Private Sub Timer3_Timer()
Call Play_WAVfile("bounce.wav")
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Call Play_WAVfile("bounce.wav")
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
Call Play_WAVfile("bzz.wav")
Timer5.Enabled = False
End Sub
