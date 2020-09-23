VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "FeedBack"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1620
      TabIndex        =   9
      Top             =   2340
      Width           =   3015
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Submit"
         Height          =   195
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   2955
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   330
      TabIndex        =   0
      Text            =   "Your Company's User Feedback Form"
      Top             =   60
      Width           =   5655
   End
   Begin Project1.SwooshBox SwooshBox1 
      Height          =   285
      Left            =   3105
      TabIndex        =   1
      Top             =   660
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   503
      ForeColor       =   12582912
   End
   Begin Project1.SwooshBox SwooshBox2 
      Height          =   285
      Left            =   3105
      TabIndex        =   2
      Top             =   1020
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   503
      ForeColor       =   12582912
   End
   Begin Project1.SwooshBox SwooshBox3 
      Height          =   285
      Left            =   3105
      TabIndex        =   3
      Top             =   1380
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   503
      ForeColor       =   12582912
   End
   Begin Project1.SwooshBox SwooshBox4 
      Height          =   285
      Left            =   3105
      TabIndex        =   4
      Top             =   1740
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   503
      ForeColor       =   12582912
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Index           =   0
      Left            =   2475
      TabIndex        =   8
      Top             =   660
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      Height          =   195
      Index           =   1
      Left            =   2610
      TabIndex        =   7
      Top             =   1740
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      Height          =   195
      Index           =   2
      Left            =   1830
      TabIndex        =   6
      Top             =   1380
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail Address"
      Height          =   195
      Index           =   3
      Left            =   1890
      TabIndex        =   5
      Top             =   1020
      Width           =   1050
   End
   Begin VB.Line Line3 
      X1              =   1620
      X2              =   1620
      Y1              =   540
      Y2              =   2220
   End
   Begin VB.Line Line4 
      X1              =   1620
      X2              =   4680
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   6360
      X2              =   6360
      Y1              =   600
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   6360
      X2              =   6360
      Y1              =   5340
      Y2              =   600
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    MsgBox SwooshBox1.StartWidth & vbCrLf & SwooshBox1.EndWidth & vbCrLf & SwooshBox1.Width
    
End Sub

