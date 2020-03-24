VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Water Bills"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Delete"
      Height          =   315
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdCompute 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Compute"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox txtWaterBill 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtGallon 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox txtWaterRate 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "The Categories"
      ForeColor       =   &H8000000B&
      Height          =   1215
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   7575
      Begin VB.CommandButton cmdIndustriall 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Industrial Use"
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdCommerciall 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Commercial Use"
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdHomee 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Home Use"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Click Your Category Here"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   3360
      Width           =   7575
   End
   Begin VB.Label lblWaterBill 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Your Water Bill"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblGallon 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "How Many Gallons?"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label lblWaterRate 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Your Water Rate"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Koica Water District Company"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   7335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Santiago Water District Company"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCompute_Click()
Dim WaterRate As String
Dim Gallon As Single
Dim WaterBill As Single

Gallon = Val(txtGallon.Text)

Select Case txtWaterRate.Text

    Case "Home Use"
        txtWaterBill.Text = (Gallon * 2) + 250
        
    Case "Commercial Use"
        If Gallon <= 4000000 Then
        txtWaterBill.Text = 5000
        End If
        
        If Gallon > 4000000 Then
        txtWaterBill.Text = 5000 + (Gallon - 4000000) * 2
        End If
        
    Case "Industrial Use"
        If Gallon <= 4000000 Then
        txtWaterBill.Text = 5000
        End If
        
        If Gallon > 4000000 And Gallon <= 10000000 Then
        txtWaterBill.Text = 10000
        End If
        
        If Gallon > 10000000 Then
        txtWaterBill.Text = 15000
        End If
        
    Case Else
        lblError.Caption = "Invalid Input!!!"
        
    End Select
    
End Sub

Private Sub cmdDelete_Click()
txtWaterRate.Text = ""
txtGallon.Text = ""
txtWaterBill.Text = ""
lblError.Caption = ""
End Sub

Private Sub cmdHomee_Click()
txtWaterRate.Text = "Home Use"
End Sub

Private Sub cmdCommerciall_Click()
txtWaterRate.Text = "Commercial Use"
End Sub

Private Sub cmdIndustriall_Click()
txtWaterRate.Text = "Industrial Use"
End Sub

