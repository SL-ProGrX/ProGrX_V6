VERSION 5.00
Begin VB.Form frmSeguros_Informes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seguros: Informes"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboAseguradora 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmSeguros_Informes.frx":0000
      Left            =   2760
      List            =   "frmSeguros_Informes.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Aseguradora"
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmSeguros_Informes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Private Sub Form_Load()
Dim strSQL As String

strSQL = "select cod_aseguradora + ' - ' + nombre as 'ItmX' from seguros_Aseguradoras"
Call sbLlenaCbo(cboAseguradora, strSQL, False)
End Sub
