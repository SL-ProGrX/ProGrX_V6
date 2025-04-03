VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.ShortcutBar.v22.0.0.ocx"
Begin VB.Form frmActivos_Periodos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Periodos"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmActivos_Periodos.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5412
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   6132
      _Version        =   1441792
      _ExtentX        =   10816
      _ExtentY        =   9546
      _StockProps     =   77
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
      Sorted          =   -1  'True
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.RadioButton rbPeriodo 
      Height          =   252
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   2052
      _Version        =   1441792
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Pendientes"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton rbPeriodo 
      Height          =   252
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Width           =   2052
      _Version        =   1441792
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cerrados"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   852
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6492
      _Version        =   1441792
      _ExtentX        =   11451
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
Attribute VB_Name = "frmActivos_Periodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkPeriodos_Click()
Call sbBuscar
End Sub

Private Sub Form_Load()

With lsw.ColumnHeaders
   .Clear
   .Add , , "Año", 1000
   .Add , , "Mes", 900, vbCenter
   .Add , , "Corte", 1400
   .Add , , "Periodo", 3000
End With

Call sbBuscar


End Sub

Private Sub sbBuscar()
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListViewItem, strResultado As String

Me.MousePointer = vbHourglass


Select Case True
  Case rbPeriodo.Item(0).Value  'Pendientes
    strSQL = "select *, dbo.fxActivos_FechaAnioMesToDatetime(Anio,Mes) as 'PeriodoCorte'" _
           & " from Activos_Periodos where estado = 'P'" _
           & " order by anio,mes"

  Case rbPeriodo.Item(1).Value  'Cancelados
    strSQL = "select *, dbo.fxActivos_FechaAnioMesToDatetime(Anio,Mes) as 'PeriodoCorte'" _
           & " from Activos_Periodos where estado = 'C'" _
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
  itmX.SubItems(2) = Format(rs!PeriodoCorte, "yyyy/mm/dd")
  itmX.SubItems(3) = strResultado
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
  gActivos.Anio = Item.Text
  gActivos.Mes = Item.SubItems(1)
  gActivos.Periodo = Item.SubItems(2)
  UnLoad Me
End Sub


Private Sub rbPeriodo_Click(Index As Integer)
Call sbBuscar
End Sub
