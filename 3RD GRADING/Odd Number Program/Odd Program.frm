VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404000&
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7185
   ForeColor       =   &H00404000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDefault 
      BackColor       =   &H00FFFFC0&
      Cancel          =   -1  'True
      Caption         =   "Default"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdErase 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Erase"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdCompute 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Compute"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtSum 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtSecond 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtFirst 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblOdd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "How many odd numbers:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "The sum is"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Please input two boundary numbers"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCompute_Click()
Dim First As Integer
Dim Second As Integer
Dim Sum As Integer
Dim Sum2 As Integer
Dim odd As Integer

First = Val(txtFirst.Text)
Second = Val(txtSecond.Text)

Do While First <= Second
    If First Mod 2 = 1 Then
        Sum2 = Sum2 + First
        First = First + 1
        txtSum.Text = Sum2
        odd = odd + 1
    Else
        First = First + 1
    End If
    lblOdd.Caption = odd
Loop
cmdErase.SetFocus
End Sub

Private Sub cmdDefault_Click()
    txtFirst.Text = "7"
    txtSecond.Text = "34"
    cmdCompute.SetFocus
End Sub

Private Sub cmdErase_Click()
txtFirst.Text = ""
txtSecond.Text = ""
txtSum.Text = ""
lblOdd.Caption = "How many odd numbers:"
txtFirst.SetFocus
End Sub

Private Sub cmdInfo_Click()
    MsgBox "This Program is used to compute all odd numbers inside a two given numbers.   Click the Default button to display 7 and 34 in the text box", vbInformation, "Guide"
    txtFirst.SetFocus
End Sub

