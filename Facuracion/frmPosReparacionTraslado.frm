VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPosReparacionTraslado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Traslado a Taller"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlb 
      Height          =   810
      Left            =   5040
      TabIndex        =   2
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1429
      ButtonWidth     =   1693
      ButtonHeight    =   1429
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Trasladar"
            Key             =   "trasladar"
            Object.ToolTipText     =   "Trasladar Articulos a Taller"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Boleta"
            Key             =   "boleta"
            Object.ToolTipText     =   "Boleta de Envio"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmPosReparacionTraslado.frx":0000
      Left            =   2760
      List            =   "frmPosReparacionTraslado.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Estado Actual del proveedor"
      Top             =   120
      Width           =   6255
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1920
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosReparacionTraslado.frx":0022
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosReparacionTraslado.frx":6884
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2892
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   11412
      _Version        =   524288
      _ExtentX        =   20130
      _ExtentY        =   5101
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
      MaxCols         =   484
      ScrollBars      =   2
      SpreadDesigner  =   "frmPosReparacionTraslado.frx":D0E6
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   11400
      X2              =   0
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmPosReparacionTraslado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean



Private Sub cbo_Click()
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "select D.linea,D.cod_producto,P.descripcion,nserie,cod_factura,detalle,boleta_traslado" _
       & " from pos_reparacion_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
       & " and D.cod_orden = '" & GLOBALES.gTag & "' and D.cod_proveedor = " & cbo.ItemData(cbo.ListIndex)
Call sbCargaGrid(vGrid, 7, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1

Exit Sub

vError:
  vGrid.MaxRows = 0
End Sub



Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

Me.Caption = "SR. Traslado a Taller [Boleta : " & GLOBALES.gTag & "]"

vModulo = 33
vGrid.AppearanceStyle = fxGridStyle


vPaso = True

strSQL = "select cod_proveedor,descripcion from cxp_proveedores" _
       & " where cod_proveedor in(select cod_proveedor from pos_reparacion_detalle" _
       & " where cod_orden = '" & GLOBALES.gTag & "')"
Call OpenRecordSet(rs, strSQL)
cbo.Clear
Do While Not rs.EOF
 cbo.AddItem rs!Descripcion
 cbo.ItemData(cbo.NewIndex) = rs!cod_proveedor
 
 If rs.AbsolutePosition = 1 Then cbo.Text = rs!Descripcion

 rs.MoveNext
Loop
rs.Close

vPaso = False
Call cbo_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbTrasladar()
Dim vExiste As Boolean, i As Integer

'Cuenta Cuantos Registros Pendientes existen
'de lo contrario no procesa nada

On Error GoTo vError

Me.MousePointer = vbHourglass

vExiste = False

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.col = vGrid.MaxCols
  
  If vGrid.Text = "" Then
    vExiste = True
    Exit For
  End If
Next i

If vExiste Then
    'Procesa Todas las Lineas Pendientes del proveedor actual para esta boleta
    Call sbPosSRAfectaInv(GLOBALES.gTag, "T", 0, cbo.ItemData(cbo.ListIndex), "", True)
    Me.MousePointer = vbDefault
    MsgBox "Registros Trasladados Satisfactoriamente...", vbInformation
    Call sbBoleta
    Call cbo_Click
Else
    Me.MousePointer = vbDefault
    MsgBox "No hay Registros Pendientes para su Traslado...", vbInformation
End If

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbBoleta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vSQL As String, vOrden As String

vSQL = ""
vOrden = ""

strSQL = "select max(BOLETA_TRASLADO) as Boleta" _
       & " From pos_reparacion_detalle" _
       & " where cod_orden = '" & GLOBALES.gTag & "' and BOLETA_TRASLADO is not null" _
       & " and cod_proveedor = " & cbo.ItemData(cbo.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  
  vSQL = "{POS_REPARACION_DETALLE.BOLETA_TRASLADO} = '" & rs!Boleta & "'"
  Call sbPosReportesSR("BoletaTraslado", "SERVICIO DE REPARACION", "TRASLADO A TALLER", vSQL, vOrden)

End If
rs.Close

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "trasladar"
    Call sbTrasladar
  Case "boleta"
    Call sbBoleta
End Select

End Sub
