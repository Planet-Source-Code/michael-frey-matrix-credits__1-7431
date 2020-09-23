VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   645
      Top             =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   1620
      Left            =   780
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   4995
      Width           =   5715
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3825
      Top             =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      Height          =   300
      Left            =   7035
      TabIndex        =   1
      Top             =   30
      Width           =   270
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   330
      Left            =   2730
      TabIndex        =   0
      Top             =   315
      Width           =   1800
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   420
      Left            =   495
      TabIndex        =   6
      Top             =   -210
      Width           =   6285
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   420
      Left            =   750
      TabIndex        =   4
      Top             =   2565
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   420
      Left            =   750
      TabIndex        =   3
      Top             =   2610
      Width           =   2970
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   450
      Left            =   60
      TabIndex        =   2
      Top             =   2580
      Width           =   7245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Hyro = "¿œŸ©¢ª£¥§©‡þ§¿Þ‡ç€ª"
Public Z

Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please enter names to show"
Exit Sub
End If
Me.Height = 4815
Command1.Visible = False
Text1.Text = Text1.Text + Chr(13) + Chr(10)
DoLabel4
Label1.Caption = ""
Label2.Caption = Hyro
Label3.Caption = Hyro
Timer2.Enabled = True
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Randomize
Label4.Top = -500
Label2.Left = -2000
Label3.Left = -2000
End Sub

Private Sub Timer1_Timer()
Label2.Left = Label2.Left + 200
Label3.Left = Label3.Left + 200
If Label2.Left = 2000 Then
a = InStr(1, Text1.Text, (Chr(13) + Chr(10)), 0)
If a = 0 Then a = Len(Text1.Text)
Label1.Caption = Mid(Text1, 1, a)
Text1.Text = Mid(Text1.Text, a + 2, Len(Text1.Text))
End If
If Label2.Left > 8000 Then
Timer1.Enabled = False
Label2.Left = -3000
Label3.Left = -3000
a = InStr(1, Text1.Text, (Chr(13) + Chr(10)), 0)
If a = 0 Then
Me.Height = 6780
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Command1.Visible = True
Else
Timer2.Enabled = True
DoLabel4
Timer1.Enabled = True
End If
End If
End Sub

Private Sub Timer2_Timer()
Label4.Top = Label4.Top + 120
Label4.Caption = ""
For x = 1 To Z
q = Int(Rnd * 13)
Label4.Caption = Label4.Caption + Mid(Hyro, q + 1, 1)
Next x
If Label4.Top > 5000 Then
Label4.Top = -500
Timer2.Enabled = False
End If
End Sub

Private Sub DoLabel4()
a = InStr(1, Text1.Text, (Chr(13) + Chr(10)), 0)
Z = Len(Mid(Text1, 1, a))
End Sub
