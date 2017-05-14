VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculator Production by Mahdi Abdollahi"
   ClientHeight    =   5115
   ClientLeft      =   2715
   ClientTop       =   2370
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   10995
   Begin VB.CommandButton Command32 
      Caption         =   " ÷—» N  « N"
      Height          =   495
      Left            =   8040
      TabIndex        =   38
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command31 
      Caption         =   "„Ã„Ê⁄ N  « N"
      Height          =   495
      Left            =   7200
      TabIndex        =   37
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Ã–—"
      Height          =   495
      Left            =   6360
      TabIndex        =   36
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command29 
      Caption         =   "%"
      Height          =   495
      Left            =   5520
      TabIndex        =   35
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H000000FF&
      Caption         =   "0"
      Height          =   495
      Left            =   8880
      TabIndex        =   27
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command27 
      Caption         =   "1"
      Height          =   495
      Left            =   1320
      TabIndex        =   26
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command26 
      Caption         =   "2"
      Height          =   495
      Left            =   2160
      TabIndex        =   25
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command25 
      Caption         =   "9"
      Height          =   495
      Left            =   8040
      TabIndex        =   24
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command24 
      Caption         =   "8"
      Height          =   495
      Left            =   7200
      TabIndex        =   23
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command23 
      Caption         =   "7"
      Height          =   495
      Left            =   6360
      TabIndex        =   22
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command22 
      Caption         =   "6"
      Height          =   495
      Left            =   5520
      TabIndex        =   21
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command21 
      Caption         =   "5"
      Height          =   495
      Left            =   4680
      TabIndex        =   20
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command20 
      Caption         =   "4"
      Height          =   495
      Left            =   3840
      TabIndex        =   19
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command19 
      Caption         =   "3"
      Height          =   495
      Left            =   3000
      TabIndex        =   18
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      Caption         =   "0"
      Height          =   495
      Left            =   8880
      TabIndex        =   17
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "C"
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      Caption         =   "C"
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H8000000D&
      Caption         =   "^"
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "/"
      Height          =   495
      Left            =   3840
      TabIndex        =   13
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "*"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "---"
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "+"
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "3"
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "4"
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "5"
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "6"
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "7"
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "8"
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "9"
      Height          =   495
      Left            =   8040
      TabIndex        =   3
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "C"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   1440
      TabIndex        =   34
      Top             =   360
      Width           =   8415
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   1440
      TabIndex        =   33
      Top             =   1800
      Width           =   8415
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   1440
      TabIndex        =   32
      Top             =   1080
      Width           =   8415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label 2"
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label 1"
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label 2"
      Height          =   255
      Left            =   10080
      TabIndex        =   29
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label 1"
      Height          =   255
      Left            =   10080
      TabIndex        =   28
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, d, e, f, g, h As Single
Private Sub Command1_Click()
Label1.Caption = Label1 + "1"
End Sub

Private Sub Command10_Click()
Label1.Caption = Label1 + "3"

End Sub

Private Sub Command11_Click()
If Label3.Caption = a Then

Label3.Caption = ""
End If

a = Val(Label1)
b = Val(Label2)
c = Val(Label3)
d = a + b
Label3.Caption = d

End Sub

Private Sub Command12_Click()
a = Val(Label1)
b = Val(Label2)
c = Val(Label3)
d = a - b
Label3.Caption = d
End Sub

Private Sub Command13_Click()
a = Val(Label1)
b = Val(Label2)
c = Val(Label3)
d = a * b
Label3.Caption = d
End Sub

Private Sub Command14_Click()
a = Val(Label1)
b = Val(Label2)
c = Val(Label3)
d = a / b
Label3.Caption = d
End Sub

Private Sub Command15_Click()
a = Val(Label1)
b = Val(Label2)
c = Val(Label3)
d = a ^ b
Label3.Caption = d
End Sub

Private Sub Command16_Click()
Label2.Caption = ""
End Sub

Private Sub Command17_Click()
Label3.Caption = ""
End Sub

Private Sub Command18_Click()
Label1.Caption = Label1 + "0"

End Sub

Private Sub Command19_Click()
Label2.Caption = Label2 + "3"

End Sub

Private Sub Command2_Click()
Label1.Caption = ""
End Sub

Private Sub Command20_Click()
Label2.Caption = Label2 + "4"
End Sub

Private Sub Command21_Click()
Label2.Caption = Label2 + "5"
End Sub

Private Sub Command22_Click()
Label2.Caption = Label2 + "6"
End Sub

Private Sub Command23_Click()
Label2.Caption = Label2 + "7"
End Sub

Private Sub Command24_Click()
Label2.Caption = Label2 + "8"
End Sub

Private Sub Command25_Click()
Label2.Caption = Label2 + "9"
End Sub

Private Sub Command26_Click()
Label2.Caption = Label2 + "2"
End Sub

Private Sub Command27_Click()
Label2.Caption = Label2 + "1"
End Sub

Private Sub Command28_Click()
Label2.Caption = Label2 + "0"
End Sub

Private Sub Command29_Click()
a = Val(Label1)
b = Val(Label2)
c = Val(Label3)
d = a / 100 * b
Label3.Caption = d
End Sub

Private Sub Command3_Click()
Label1.Caption = Label1 + "2"

End Sub

Private Sub Command30_Click()
a = Val(Label1)
b = Val(Label2)
c = Val(Label3)
d = a ^ 0.5
Label3.Caption = d
End Sub

Private Sub Command31_Click()
a = Val(Label1)
b = Val(Label2)
e = Val(Label3)
c = 0
For d = a To b
c = c + d
Next d
Label3.Caption = c

End Sub

Private Sub Command32_Click()
a = Val(Label1)
b = Val(Label2)
e = Val(Label3)
c = 1
For d = a To b
c = c * d
Next d
Label3.Caption = c
End Sub

Private Sub Command4_Click()
Label1.Caption = Label1 + "9"

End Sub

Private Sub Command5_Click()
Label1.Caption = Label1 + "8"

End Sub

Private Sub Command6_Click()
Label1.Caption = Label1 + "7"

End Sub

Private Sub Command7_Click()
Label1.Caption = Label1 + "6"

End Sub

Private Sub Command8_Click()
Label1.Caption = Label1 + "5"

End Sub

Private Sub Command9_Click()
Label1.Caption = Label1 + "4"

End Sub

Private Sub Form_Load()
Form1.BackColor = vbWhite
End Sub
