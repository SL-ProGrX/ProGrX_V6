VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmInvUnidadesConv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversión de Unidades"
   ClientHeight    =   5928
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7512
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5928
   ScaleWidth      =   7512
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5172
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   7092
      _Version        =   524288
      _ExtentX        =   12510
      _ExtentY        =   9123
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   487
      ScrollBars      =   2
      SpreadDesigner  =   "frmInvUnidadesConv.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Unidad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmInvUnidadesConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset


vGrid.AppearanceStyle = fxGridStyle

strSQL = "select rtrim(cod_unidad) + ' - ' + descripcion  as DescX from pv_unidades"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cbo.AddItem rs!descx
 rs.MoveNext
Loop

If rs.RecordCount > 0 Then
  rs.MoveFirst
  cbo.Text = rs!descx
End If

rs.Close

vGrid.MaxCols = 3
vGrid.MaxRows = 1

End Sub
