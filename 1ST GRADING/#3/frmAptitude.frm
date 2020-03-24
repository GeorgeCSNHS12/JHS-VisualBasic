VERSION 5.00
Begin VB.Form frmAptitude 
   BackColor       =   &H00808000&
   Caption         =   "Aptitude"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdErase 
      BackColor       =   &H00FFFF80&
      Caption         =   "Erase"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdEvaluate 
      BackColor       =   &H00FFFF80&
      Caption         =   "Evaluate"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtVar2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtVar1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblEval2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Evaluation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label lblEval1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Evaluation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label lblVar2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Part 2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblVar1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Part 1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Enter Your Scores in Part I and Part II"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmAptitude"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdErase_Click()
txtVar1.Text = ""
txtVar2.Text = ""
lblEval1.Caption = "Evaluation"
lblEval2.Caption = "Evaluation"
End Sub

Private Sub cmdEvaluate_Click()
Dim Var1 As Single
Dim Var2 As Single
Dim Eval1 As Single
Dim Eval2 As Single

Var1 = Val(txtVar1.Text)
Var2 = Val(txtVar2.Text)

If Var1 >= 0 And Var1 <= 20 Then
    lblEval1.Caption = "Retake in 30 Days"
End If

If Var2 >= 0 And Var2 <= 10 Then
    lblEval2.Caption = "Retake in 30 Days"
End If


If Var1 >= 21 And Var1 <= 35 Then
    lblEval1.Caption = "Take Practical Test"
End If

If Var2 >= 11 And Var2 <= 25 Then
    lblEval2.Caption = "Take Practical Test"
End If


If Var1 >= 36 And Var1 <= 50 Then
    lblEval1.Caption = "Certification Completed"
End If

If Var2 >= 26 And Var2 <= 50 Then
    lblEval2.Caption = "Certification Completed"
End If


If Var1 > 50 Or Var1 < 0 Then
    lblEval1.Caption = "Invalid"
End If

If Var2 > 50 Or Var2 < 0 Then
    lblEval2.Caption = "Invalid"
End If
End Sub
