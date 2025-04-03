VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   2436
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   6768
   LinkTopic       =   "Form1"
   ScaleHeight     =   2436
   ScaleWidth      =   6768
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   252
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   1332
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   2172
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2172
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Test As New clsProGrX_Global

Private Sub Command1_Click()


Text2.Text = Test.fxStringCifrado(Text1.Text)

End Sub
