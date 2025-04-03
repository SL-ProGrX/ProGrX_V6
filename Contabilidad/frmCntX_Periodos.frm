VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#18.6#0"; "Codejock.Controls.v18.6.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#18.6#0"; "Codejock.ShortcutBar.v18.6.0.ocx"
Begin VB.Form frmCntX_Periodos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccione un Periodo"
   ClientHeight    =   6756
   ClientLeft      =   48
   ClientTop       =   216
   ClientWidth     =   6168
   HelpContextID   =   2005
   Icon            =   "frmCntX_Periodos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6756
   ScaleWidth      =   6168
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5172
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5892
      _Version        =   1179654
      _ExtentX        =   10393
      _ExtentY        =   9123
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.RadioButton rbPeriodo 
      Height          =   252
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   2052
      _Version        =   1179654
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Pendientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton rbPeriodo 
      Height          =   252
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Width           =   2052
      _Version        =   1179654
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cerrados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   852
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6252
      _Version        =   1179654
      _ExtentX        =   11028
      _ExtentY        =   1503
      _StockProps     =   14
      Caption         =   "Seleccione el periodo a visualizar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
End
Attribute VB_Name = "frmCntX_Periodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
vModulo = 20

With lsw.ColumnHeaders
  .Clear
  .Add , , "Año", 1000
  .Add , , "Mes", 1000, vbCenter
  .Add , , "Periodo", 3500
End With

Call sbBuscar
End Sub

Private Sub sbBuscar()
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListViewItem, strResultado As String

Me.MousePointer = vbHourglass

Select Case True
  Case rbPeriodo.Item(0).Value 'Pendientes
    strSQL = "select * from CntX_Periodos where cod_contabilidad = " _
           & gCntX_Parametros.CodigoConta & " and estado = 'P'" _
           & " order by anio,mes"
           
  Case rbPeriodo.Item(1).Value 'Cerrados
    strSQL = "select * from CntX_Periodos where cod_contabilidad = " _
           & gCntX_Parametros.CodigoConta & " and estado = 'C'" _
           & " order by anio desc,mes desc"
End Select

Call OpenRecordSet(rs, strSQL, 0)
lsw.ListItems.Clear
Do While Not rs.EOF
  
  
    Select Case rs!Mes
    Case 1
        strResultado = "ENERO DE " & rs!Anio
    Case 2
        strResultado = "FEBRERO DE " & rs!Anio
    Case 3
        strResultado = "MARZO DE " & rs!Anio
    Case 4
        strResultado = "ABRIL DE " & rs!Anio
    Case 5
        strResultado = "MAYO DE " & rs!Anio
    Case 6
        strResultado = "JUNIO DE " & rs!Anio
    Case 7
        strResultado = "JULIO DE " & rs!Anio
    Case 8
        strResultado = "AGOSTO DE " & rs!Anio
    Case 9
        strResultado = "SEPTIEMBRE DE " & rs!Anio
    Case 10
        strResultado = "OCTUBRE DE " & rs!Anio
    Case 11
        strResultado = "NOVIEMBRE DE " & rs!Anio
    Case 12
        strResultado = "DICIEMBRE DE " & rs!Anio
  End Select

  Set itmX = lsw.ListItems.Add(, , rs!Anio)
    itmX.SubItems(1) = rs!Mes
    itmX.SubItems(2) = strResultado
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
  gCntX_Parametros.PeriodoAnio = Item.Text
  gCntX_Parametros.PeriodoMes = Item.SubItems(1)
  Unload Me
End Sub

Private Sub lsw_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  gCntX_Parametros.PeriodoAnio = lsw.SelectedItem
  gCntX_Parametros.PeriodoMes = lsw.SelectedItem.SubItems(1)
  Unload Me
End If
End Sub

Private Sub rbPeriodo_Click(Index As Integer)
Call sbBuscar
End Sub
