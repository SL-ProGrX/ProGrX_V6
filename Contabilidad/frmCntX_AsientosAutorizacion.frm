VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmCntX_AsientosAutorizacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Autorización de Asientos Foráneos"
   ClientHeight    =   7272
   ClientLeft      =   -12
   ClientTop       =   348
   ClientWidth     =   10788
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7272
   ScaleWidth      =   10788
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5892
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   10572
      _Version        =   1245187
      _ExtentX        =   18648
      _ExtentY        =   10393
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   200
      Left            =   240
      TabIndex        =   5
      Top             =   1060
      Width           =   200
      _Version        =   1245187
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Transparent     =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   492
      Left            =   7680
      TabIndex        =   1
      Top             =   240
      Width           =   1452
      _Version        =   1245187
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCntX_AsientosAutorizacion.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdAutorizar 
      Height          =   492
      Left            =   9120
      TabIndex        =   2
      Top             =   240
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Autorizar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCntX_AsientosAutorizacion.frx":0A1E
   End
   Begin XtremeSuiteControls.FlatEdit txtCAsiento 
      Height          =   312
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   852
      _Version        =   1245187
      _ExtentX        =   1503
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDAsiento 
      Height          =   312
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   3972
      _Version        =   1245187
      _ExtentX        =   7006
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   10572
      _Version        =   1245187
      _ExtentX        =   18648
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Asientos pendientes de Autorización"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Asiento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1428
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmCntX_AsientosAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub chkTodos_Click()
Dim i As Integer


For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub

Private Sub cmdAutorizar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer


Me.MousePointer = vbHourglass

strSQL = ""
For i = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(i).Checked Then
      
   strSQL = strSQL & Space(10) & "update Cntx_Asientos set user_autoriza = '" & glogon.Usuario _
          & "',fecha_autoriza = getdate()" _
          & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
          & " and tipo_asiento = '" & lsw.ListItems.Item(i).SubItems(1) _
          & "' and num_asiento = '" & lsw.ListItems.Item(i).Text & "'"
   If Len(strSQL) > 20000 Then
       Call ConectionExecute(strSQL, 0)
       strSQL = ""
   End If
 End If
Next i
   'Aplica Lote Final
   If Len(strSQL) > 0 Then
       Call ConectionExecute(strSQL, 0)
   End If

Me.MousePointer = vbDefault

MsgBox "Asientos Foráneos Autorizados Satisfactorimente...", vbInformation
Call sbBuscar

End Sub

Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "select A.Num_asiento,A.Tipo_Asiento,A.Descripcion,A.Fecha_Asiento" _
       & ",sum(isnull(D.monto_Debito,0)) as Debitos,sum(isnull(D.monto_credito,0)) as creditos" _
       & " from Cntx_Asientos A left join Cntx_Asientos_Detalle D on A.cod_contabilidad = D.cod_contabilidad" _
       & " and A.tipo_asiento = D.tipo_asiento and A.num_asiento = D.num_asiento" _
       & " where A.tipo_asiento = '" & txtCAsiento & "' and A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and A.modulo <> 20 and A.fecha_autoriza is null and A.anio = " & gCntX_Parametros.PeriodoAnio _
       & " and A.mes = " & gCntX_Parametros.PeriodoMes _
       & " group by A.Num_asiento,A.Tipo_Asiento,A.Descripcion,A.Fecha_Asiento"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Num_Asiento)
     itmX.SubItems(1) = rs!Tipo_Asiento
     itmX.SubItems(2) = rs!Descripcion
     itmX.SubItems(3) = Format(rs!fecha_asiento, "dd/mm/yyyy")
     itmX.SubItems(4) = Format(rs!Debitos, "Standard")
     itmX.SubItems(5) = Format(rs!Creditos, "Standard")
     itmX.Checked = chkTodos.Value
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 20
End Sub

Private Sub Form_Load()

vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


With lsw.ColumnHeaders
    .Clear
    .Add , , "No. Asiento", 2500
    .Add , , "Tipo", 900, vbCenter
    .Add , , "Descripción", 4500
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Débitos", 2100, vbRightJustify
    .Add , , "Crébitos", 2100, vbRightJustify
End With


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub txtCAsiento_Change()
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select descripcion from CntX_Tipos_Asientos where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " and tipo_asiento = '" _
       & txtCAsiento.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  txtDAsiento = rs!Descripcion
End If
rs.Close

lsw.ListItems.Clear

End Sub

Private Sub txtCAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "Tipo_Asiento"
    gBusquedas.Orden = "Tipo_Asiento"
    gBusquedas.Consulta = "select Tipo_Asiento,descripcion from CntX_Tipos_Asientos"
    gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
    frmBusquedas.Show vbModal
    txtCAsiento = gBusquedas.Resultado
End If
End Sub


