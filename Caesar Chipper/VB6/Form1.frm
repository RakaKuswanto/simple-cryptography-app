VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Caesar Chipper"
   ClientHeight    =   3210
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Help"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear Text"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.HScrollBar scrl 
      Height          =   255
      Left            =   120
      Max             =   26
      Min             =   1
      TabIndex        =   6
      Top             =   1200
      Value           =   1
      Width           =   4215
   End
   Begin VB.TextBox txthasil 
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Text            =   "hasil akan muncul disini"
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox txt 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Text            =   "ketik sesuatu disini"
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Get &PlainText"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Get Caesar Chipper"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Raka Kuswanto @ decenzo Lab"
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
txthasil.Text = ""
getCaesarText LCase(txt.Text), scrl.Value
End Sub

Private Sub Command2_Click()
txt.Text = ""
txthasil.Text = ""
End Sub

Private Sub Command3_Click()
DecryptCaesarText LCase(txt.Text), scrl.Value
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
MsgBox "1.Ketik apa yang akan anda ubah ke bentuk chipper di textbox paling atas" & vbCrLf & _
        "2.Hasil Akan Muncul Di textbox yang bawah" & vbCrLf & _
        "3.Untuk membalikan kembali, silahkan copy paste data di textbox hasil ke textbox yang paling atas", vbOKOnly, "Caesar Chipper"
End Sub

Private Sub Form_Load()
'sesar True, 2, 2
'getIdxFromStr ("a")
'MsgBox Mid("raka", 2, 1)
'MsgBox getCaesarText("xyz", 1)
End Sub

Private Sub scrl_Change()
Me.Caption = "Caesar Chipper - Digeser " & scrl.Value
End Sub
