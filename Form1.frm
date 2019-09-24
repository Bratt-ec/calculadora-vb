VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00369FD8&
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BACP_Salir 
      BackColor       =   &H00369FD8&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton BACP_Pi 
      BackColor       =   &H00369FD8&
      Caption         =   "Pi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton BACP_Ptc 
      BackColor       =   &H00369FD8&
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton BACP_SQr 
      BackColor       =   &H00369FD8&
      Caption         =   "SQR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton BACP_DIV 
      BackColor       =   &H00369FD8&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton BACP_MULT 
      BackColor       =   &H00369FD8&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton BACP_RES 
      BackColor       =   &H00369FD8&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton BACP_SUM 
      BackColor       =   &H00369FD8&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton BACP_IGUAL 
      BackColor       =   &H00369FD8&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton BACP_BORRAR 
      BackColor       =   &H00369FD8&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton BACP_0 
      BackColor       =   &H00369FD8&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton BACP_9 
      BackColor       =   &H00369FD8&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton BACP_8 
      BackColor       =   &H00369FD8&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton BACP_7 
      BackColor       =   &H00369FD8&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton BACP_6 
      BackColor       =   &H00369FD8&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton BACP_5 
      BackColor       =   &H00369FD8&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton BACP_4 
      BackColor       =   &H00369FD8&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton BACP_3 
      BackColor       =   &H00369FD8&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton BACP_2 
      BackColor       =   &H00369FD8&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton BACP_1 
      BackColor       =   &H00369FD8&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00369FD8&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00369FD8&
      Caption         =   "By: Bryan Castro"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   0
      TabIndex        =   22
      Top             =   5040
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00369FD8&
      Caption         =   "CALCULADORA 2.1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim producto1, caracter, a, b, c, d, e, f, g, h, i, j, concatene

Private Sub BACP_0_Click()
caracter = Text1.Text
concatene = concatene & j
Text1.Text = Val(concatene)
End Sub

Private Sub BACP_1_Click()
'caracter = Text1.Text'
concatene = concatene & a
Text1.Text = Val(concatene)
End Sub

Private Sub BACP_2_Click()
'caracter = Text1.Text'
concatene = concatene & b
Text1.Text = Val(concatene)
End Sub

Private Sub BACP_3_Click()
'caracter = Text1.Text'
concatene = concatene & c
Text1.Text = Val(concatene)
End Sub

Private Sub BACP_4_Click()
'caracter = Text1.Text'
concatene = concatene & d
Text1.Text = Val(concatene)
End Sub

Private Sub BACP_5_Click()
'caracter = Text1.Text'
concatene = concatene & e
Text1.Text = Val(concatene)
End Sub

Private Sub BACP_6_Click()
'caracter = Text1.Text'
concatene = concatene & f
Text1.Text = Val(concatene)
End Sub

Private Sub BACP_7_Click()
'caracter = Text1.Text'
concatene = concatene & g
Text1.Text = Val(concatene)
End Sub

Private Sub BACP_8_Click()
'caracter = Text1.Text'
concatene = concatene & h
Text1.Text = Val(concatene)
End Sub

Private Sub BACP_9_Click()
'caracter = Text1.Text'
concatene = concatene & i
Text1.Text = Val(concatene)
End Sub

Private Sub BACP_BORRAR_Click()
Text1.Text = " "
caracter = " "
concatene = " "
producto1 = " "
End Sub

Private Sub BACP_Dec_Click()
'caracter = Text1.Text'
concatene = concatene & k
Text1.Text = Val(concatene)
End Sub

Private Sub BACP_DIV_Click()
Text1.Text = " "
Text1.Text = "/"
producto1 = concatene
concatene = ""
caracter = Text1.Text
End Sub

Private Sub BACP_IGUAL_Click()
If LTrim(caracter) = "/" Then
Text1.Text = CDbl(producto1) / CDbl(concatene)
End If
If LTrim(caracter) = "*" Then
Text1.Text = CDbl(producto1) * CDbl(concatene)
End If
If LTrim(caracter) = "-" Then
Text1.Text = CDbl(producto1) - CDbl(concatene)
End If
If LTrim(caracter) = "+" Then
Text1.Text = CDbl(producto1) + CDbl(concatene)
End If
If LTrim(caracter) = "sqr" Then
Text1.Text = Sqr(CDbl(concatene))
End If
If LTrim(caracter) = "^" Then
Text1.Text = CDbl(producto1) ^ CDbl(concatene)
End If
End Sub

Private Sub BACP_MULT_Click()
Text1.Text = " "
Text1.Text = "*"
producto1 = concatene
concatene = ""
caracter = Text1.Text
End Sub

Private Sub BACP_Pi_Click()
Text1.Text = ""
Text1.Text = "Pi"
producto1 = concatene
concatene = 3.14
caracter = Text1.Text
End Sub

Private Sub BACP_Ptc_Click()
Text1.Text = ""
Text1.Text = "^"
producto1 = concatene
concatene = ""
caracter = Text1.Text
End Sub

Private Sub BACP_RES_Click()
Text1.Text = " "
Text1.Text = "-"
producto1 = concatene
concatene = ""
caracter = Text1.Text
End Sub

Private Sub BACP_Salir_Click()
Unload Me
End
End Sub

Private Sub BACP_SQr_Click()
Text1.Text = ""
Text1.Text = "sqr"
producto1 = concatene
concatene = ""
caracter = Text1.Text
End Sub

Private Sub BACP_SUM_Click()
Text1.Text = " "
Text1.Text = "+"
producto1 = concatene
concatene = ""
caracter = Text1.Text
End Sub

Private Sub Command3_Click()
'caracter = Text1.Text'
concatene = concatene & k
Text1.Text = Val(concatene)
End Sub

Private Sub Form_Load()
a = "1"
b = "2"
c = "3"
d = "4"
e = "5"
f = "6"
g = "7"
h = "8"
i = "9"
j = "0"
Text1.Text = "0"
Text1.Enabled = False
End Sub

