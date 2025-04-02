VERSION 5.00
Begin VB.Form frmCR_AbonosConector 
   BorderStyle     =   0  'None
   Caption         =   "Abonos Conector!"
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Abonos Conector!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmCR_AbonosConector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
If GLOBALES.SysPlanPagos = 1 Then
   Call sbFormsCall("frmCajas_Crd_AbonosCtP", 0, , , True, Me)
Else
   Call sbFormsCall("frmCajas_Crd_AbonosStP", 0, , , True, Me)
End If

Me.Hide
End Sub

