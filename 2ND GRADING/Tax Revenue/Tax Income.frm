VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H005C4334&
   Caption         =   "Tax Income"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7785
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdErase 
      BackColor       =   &H00FFFFC0&
      Cancel          =   -1  'True
      Caption         =   "Erase"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Calculate"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtTax 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2760
      TabIndex        =   4
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtIncome 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      BackColor       =   &H005C4334&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   2280
      Width           =   7695
   End
   Begin VB.Label lblTax 
      Alignment       =   2  'Center
      BackColor       =   &H005C4334&
      Caption         =   "Your Total Tax"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblIncome 
      Alignment       =   2  'Center
      BackColor       =   &H005C4334&
      Caption         =   "Your Income?"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H005C4334&
      Caption         =   "Bureau of Internal Revenue"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
Dim Income As Single
Dim Tax As Single

Income = Val(txtIncome.Text)

Select Case Income
    Case Is < "10000"
        txtTax.Text = 0
    Case Is >= "10000"
        If Income >= 10000 And Income < 50000 Then
            txtTax.Text = (Income * 0.5) - 5000
        End If
        If Income >= 50000 And Income <= 100000 Then
            txtTax.Text = Income * 0.5
        End If
        If Income > 100000 Then
            txtTax.Text = (Income - 100000) * 0.2 + (Income * 0.5)
        End If
        
    Case Else
        lblError.Caption = "Invalid Input!!!"
    End Select
        
End Sub

Private Sub cmdErase_Click()
txtIncome.Text = ""
txtTax.Text = ""
lblError.Caption = ""
End Sub
