VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   309
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   451
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1320
      Top             =   3840
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pattern screensaver: CREATED BY JOHANNES.B 2000    JB_Rulez_54@hotmail.com"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.OptionButton Option2 
         Caption         =   "Lines"
         Height          =   255
         Left            =   3480
         TabIndex        =   23
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Circles"
         Height          =   255
         Left            =   2520
         TabIndex        =   22
         Top             =   1320
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   35
         Left            =   2400
         Max             =   255
         Min             =   1
         TabIndex        =   20
         Top             =   2040
         Value           =   255
         Width           =   2295
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "3"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "5"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   17
         Text            =   "5"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3480
         MaxLength       =   1
         TabIndex        =   14
         Text            =   "1"
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "5"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "500"
         Top             =   1200
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "RND Forecolor"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "RND Backcolor"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "100"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "100"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   2640
         Width           =   4455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Brightnes"
         Height          =   255
         Left            =   2400
         TabIndex        =   21
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "From"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "DrawWidth:"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Delay:"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "RND Size:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "RND Height:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "RND Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim JB1, JB2, JB3 As Integer
Dim ST As Integer
Dim A, B As Integer
Dim WA As Integer
Dim W As Integer
Private Sub Command1_Click()
ST = 1
Frame1.Visible = False
End Sub

Private Sub Command2_Click()
MsgBox "Please vote!"
End
End Sub

Private Sub Form_Click()
ST = 0
Frame1.Visible = True
End Sub

Private Sub Text1_Change()
On Error GoTo err
If Text1.Text <= 0 Then
Text1.Text = 1
End If
Exit Sub
err:
Text1.Text = 1
Exit Sub
End Sub


Private Sub Text2_Change()
On Error GoTo err
If Text2.Text <= 0 Then
Text2.Text = 1
End If
Exit Sub
err:
Text2.Text = 1
Exit Sub
End Sub


Private Sub Text3_Change()
On Error GoTo err
If Text3.Text <= 0 Then
Text3.Text = 1
End If
Exit Sub
err:
Text3.Text = 1
Exit Sub
End Sub


Private Sub Text4_Change()
On Error GoTo err
If Text4.Text <= 2 Then
Text4.Text = 2
End If
Exit Sub
err:
Text4.Text = 1
Exit Sub
End Sub


Private Sub Text5_Change()
Form1.DrawWidth = Text5.Text
On Error GoTo err
If Text5.Text <= 0 Then
Text5.Text = 1
End If
Exit Sub
err:
Text5.Text = 1
Exit Sub
End Sub


Private Sub Text6_Change()
On Error GoTo err
If Text6.Text <= 0 Then
Text6.Text = 1
End If
Exit Sub
err:
Text6.Text = 1
Exit Sub
End Sub

Private Sub Text7_Change()
On Error GoTo err
If Text7.Text <= 0 Then
Text7.Text = 1
End If
Exit Sub
err:
Text7.Text = 1
Exit Sub
End Sub


Private Sub Text8_Change()
On Error GoTo err
If Text8.Text <= 0 Then
Text8.Text = 1
End If
Exit Sub
err:
Text8.Text = 1
Exit Sub
End Sub


Private Sub Timer1_Timer()
On Error Resume Next

If ST = 1 Then
1:
If W = 1 Then
'Delay
WA = WA + 1

If WA = Text4.Text Then W = 0: WA = 0
Exit Sub
End If

Form1.Cls

'RND Values
JB1 = Int((Text1.Text - Text6.Text + 1) * Rnd + Text6.Text)
JB2 = Int((Text2.Text - Text7.Text + 1) * Rnd + Text7.Text)
JB3 = Int((Text3.Text - Text8.Text + 1) * Rnd + Text8.Text)

' Reset
A = -JB3
B = -JB3

'Set colors
If Check1.Value = 1 Then Form1.BackColor = RGB(Rnd * HScroll1.Value, Rnd * HScroll1.Value, Rnd * HScroll1.Value)
If Check2.Value = 1 Then Form1.ForeColor = RGB(Rnd * HScroll1.Value, Rnd * HScroll1.Value, Rnd * HScroll1.Value)

Do
A = A + JB1



If A > Form1.ScaleWidth + JB3 Then
A = -JB3
B = B + JB2
End If

If Option2.Value = True Then
Form1.Line (A, B)-(B, A)
End If
If Option1.Value = True Then
Form1.Circle (A, B), JB3
End If


Loop Until B > Form1.ScaleHeight + JB3
'Delay = true
W = 1
'New
GoTo 1

End If
End Sub


