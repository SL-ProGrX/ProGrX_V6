VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCprCompraDev 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolucion de Mercaderia comprada"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   11175
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
      Height          =   312
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   480
      Width           =   4572
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   25
      Top             =   4920
      Width           =   6135
      Begin VB.TextBox txtGenFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   27
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtGenUser 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Generado /  Fecha"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   600
         TabIndex        =   29
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Generado / Usuario"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   600
         TabIndex        =   28
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.TextBox txtRegistro 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtOrden 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtCompra 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4440
      TabIndex        =   20
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtProveedor 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   19
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   960
      Width           =   4572
   End
   Begin VB.TextBox txtProvCod 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   5925
      Width           =   1812
   End
   Begin VB.TextBox txtImpuestos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5565
      Width           =   1812
   End
   Begin VB.TextBox txtDescuento 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5235
      Width           =   1812
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4875
      Width           =   1812
   End
   Begin VB.TextBox txtDevolucion 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtNotas 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1680
      Width           =   8535
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repBoleta"
                  Text            =   "Boleta "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repListadoGeneral"
                  Text            =   "Listado General"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repSep1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repAntiguedadSaldos"
                  Text            =   "Antiguedad Saldos"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repPagosPendientes"
                  Text            =   "Pagos Pendientes"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgAux01 
      Left            =   5880
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCprCompraDev.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCprCompraDev.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCprCompraDev.frx":0BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCprCompraDev.frx":0F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCprCompraDev.frx":1238
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2292
      Left            =   0
      TabIndex        =   31
      Top             =   2400
      Width           =   10932
      _Version        =   524288
      _ExtentX        =   19283
      _ExtentY        =   4043
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
      MaxCols         =   488
      ScrollBars      =   2
      SpreadDesigner  =   "frmCprCompraDev.frx":1554
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   10
      Left            =   8280
      TabIndex        =   24
      Top             =   1320
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "# Compra/Orden"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   3120
      TabIndex        =   21
      Top             =   1320
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   972
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   10080
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   9
      Left            =   7200
      TabIndex        =   17
      Top             =   5928
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "(+) Impuestos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   8
      Left            =   7200
      TabIndex        =   16
      Top             =   5568
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "(-) Descuento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Index           =   7
      Left            =   7200
      TabIndex        =   15
      Top             =   5232
      Width           =   1812
   End
   Begin VB.Label Label1 
      Caption         =   "Sub Total"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   7200
      TabIndex        =   14
      Top             =   4872
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "No. Dev:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   8280
      TabIndex        =   12
      Top             =   960
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Factura"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   972
   End
End
Attribute VB_Name = "frmCprCompraDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vDevolucion As String
Dim vMascara As String

Private Sub Form_Activate()
vModulo = 35
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

 vModulo = 35
 vMascara = "0000000000"
 
 vGrid.AppearanceStyle = fxGridStyle
 
 Call sbCprCboCargosPer(cbo)
 
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
Dim i As Integer

vCodigo = ""
txtCodigo = ""
vDevolucion = ""
txtDevolucion = ""

txtFecha = Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss")
txtNotas = ""

txtOrden = ""
txtCompra = ""


vGrid.MaxRows = 1
vGrid.MaxCols = 7
For i = 1 To vGrid.MaxCols
  vGrid.col = i
  vGrid.Text = ""
Next

txtSubTotal = 0
txtDescuento = 0
txtImpuestos = 0
txtTotal = 0

txtCodigo.Enabled = True
txtDevolucion.Enabled = True

txtProvCod = ""
txtProveedor = ""


End Sub




Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      
      txtCodigo.Enabled = True
      txtCodigo.SetFocus
      txtDevolucion.Enabled = False
      
      vGrid.Enabled = True
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      vGrid.Enabled = False
      vGrid.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsultaFac(vCodigo)
      End If

    Case "CONSULTAR"
'       gBusquedas.Columna = "descripcion"
'       gBusquedas.Orden = "descripcion"
'       gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
'       frmBusquedas.Show vbModal
'       txtCodigo.SetFocus
'       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
'       txtNombre.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Function sbDesActProdCancelados()
Dim lng As Long

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 3
 If vGrid.Text <> "" Then
   vGrid.CellTag = CCur(vGrid.Text)
   If CCur(vGrid.Text) <= 0 Then
      vGrid.col = 1
      vGrid.Text = ""
   End If
 End If
Next lng

End Function


Private Sub sbConsultaFac(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select E.*,P.descripcion as Proveedor" _
       & " from cpr_compras E inner join cxp_Proveedores P on E.cod_proveedor = P.cod_proveedor" _
       & " where E.cod_factura = '" & xCodigo & "' and E.cod_proveedor = " & txtProvCod
       
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
'  Call sbToolBar(tlb, "edicion")
  vEdita = False 'False
  Call sbLimpiaPantalla
  
  vCodigo = rs!cod_Factura
  txtCodigo = rs!cod_Factura
  
  txtProvCod = rs!cod_Proveedor
  txtProveedor = rs!Proveedor
  
  txtOrden = rs!cod_orden
  txtCompra = rs!cod_compra
  
  txtRegistro = rs!fecha
  
 'Tengo que conservar visibles aquellos que ya fueron despachados para conservar el consecutivo
 'de la linea del detalle, indica la bodega por defecto
  strSQL = "select D.cod_producto,P.descripcion,(D.cantidad - isnull(D.cantidad_devuelta,0)) as Cantidad" _
         & ",D.cod_bodega,D.precio,D.imp_ventas,(((D.cantidad - isnull(D.cantidad_devuelta,0)) * D.precio)" _
         & " * ((D.imp_ventas / 100) + 1)) as Total" _
         & " from cpr_compras_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " where D.cod_factura = '" & rs!cod_Factura & "' and D.cod_proveedor = " & rs!cod_Proveedor _
         & " order by D.Linea"
  
  Call sbCargaGrid(vGrid, 7, strSQL)
  Call sbDesActProdCancelados
  
  Call sbCalculaTotales
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select D.*,P.descripcion as Proveedor, rtrim(C.cod_cargo) + ' - ' + rtrim(C.descripcion) as CargoX" _
       & " from cpr_compras_dev D inner join cxp_Proveedores P on D.cod_proveedor = P.cod_proveedor" _
       & " inner join cxp_cargos C on D.cod_cargo = C.cod_cargo" _
       & " where D.cod_compra_dev = '" & Format(xCodigo, vMascara) & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!cod_Factura
  txtCodigo = rs!cod_Factura
  
  txtDevolucion = rs!cod_compra_dev
  vDevolucion = rs!cod_compra_dev
  
  txtProvCod = rs!cod_Proveedor
  txtProveedor = rs!Proveedor
  
  cbo.Text = rs!cargoX
  
  txtFecha = Format(rs!fecha, "yyyy/mm/dd hh:mm:ss")
  txtNotas = rs!Notas & ""
  
  txtImpuestos = Format(rs!imp_ventas, "Standard")
  txtGenFecha = rs!genera_fecha & ""
  txtGenUser = rs!genera_user & ""
  
  strSQL = "select D.cod_producto,P.descripcion,D.cantidad,D.cod_bodega,D.precio,D.imp_ventas," _
         & "(D.cantidad * D.precio) + (D.cantidad * D.precio * (D.imp_ventas / 100)) as Total" _
         & " from cpr_compra_devDet D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " where D.cod_compra_dev = '" & rs!cod_compra_dev & "' order by D.Linea"
  Call sbCargaGrid(vGrid, 7, strSQL)
  
  strSQL = "select cod_orden,cod_compra,fecha from cpr_compras where cod_factura = '" & rs!cod_Factura _
         & "' and cod_proveedor = " & rs!cod_Proveedor
  rs.Close
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then
     txtOrden = rs!cod_orden & ""
     txtCompra = rs!cod_compra & ""
     txtRegistro = rs!fecha & ""
  End If
  
  Call sbCalculaTotales
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer
Dim vMensaje As String


On Error GoTo vError

vMensaje = ""
fxValida = True

'Verifica Bodegas y Articulos
vMensaje = fxInvVerificaLineaDetalle(vGrid, 3, "S", 1, 4)

'Verifica Periodo
If Not fxInvPeriodos(fxFechaServidor) Then vMensaje = vMensaje & vbCrLf & " - El Periodo del Movimiento no es válido ..."

'If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Proveedor no es válido ..."
'Verifiqua que exista la factura y que no se encuentre anulada

If IsNumeric(txtProvCod) Then
    strSQL = "select estado from cpr_compras where cod_factura = '" & txtCodigo _
           & "' and cod_proveedor = " & txtProvCod & " and estado in('P','D')"
    Call OpenRecordSet(rs, strSQL)
    If rs.EOF And rs.BOF Then
       vMensaje = vMensaje & vbCrLf & " - No se encontró registro de la factura, o se encuentra Anulada, verifique..."
    End If
    rs.Close
Else
   vMensaje = vMensaje & vbCrLf & " - El codigo del Proveedor no es válido, verifique..."
End If

'Verifica que las cantidades de las devoluciones no sean mayores al original pendiente
For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.col = 1
  If vGrid.Text <> "" Then
    vGrid.col = 3
    If CCur(vGrid.Text) > CCur(vGrid.CellTag) Then
     vMensaje = vMensaje & vbCrLf & " - Las Cantidad devoluciones en la Linea " & i & ", es mayor al remanente..."
    End If
  End If
Next i

If CCur(txtTotal) = 0 Then vMensaje = vMensaje & vbCrLf & " - El total de la mercadería es cero, verifique..."


vError:

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Function fxConsecDev() As String
Dim strSQL As String, rs As New ADODB.Recordset

'Consecutivo de la Orden
strSQL = "select isnull(max(cod_compra_dev),0) + 1 as Ultimo from cpr_compras_dev"
Call OpenRecordSet(rs, strSQL)
  fxConsecDev = Format(rs!ultimo, vMascara)
rs.Close

End Function


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer, vFecha As Date
Dim curCantidad As Currency, vCodPro As String, vCodBodega As String
Dim curPrecio As Currency, curImpVentas As Currency, curImpConsumo As Currency

On Error GoTo vError

If vEdita Then
   MsgBox "No se puede editar una Devolución Guardada...", vbInformation
   Exit Sub
Else
   
   vFecha = fxFechaServidor
   vDevolucion = fxConsecDev
   txtDevolucion = vDevolucion
   
   txtCodigo = vCodigo
   vCodigo = txtCodigo
   
   strSQL = "insert cpr_compras_dev(cod_compra_dev,cod_factura,cod_proveedor,fecha,sub_total,descuento,imp_ventas" _
          & ",imp_consumo,total,notas,asiento_estado,genera_user,genera_fecha,cod_cargo) values('" & vDevolucion _
          & "','" & vCodigo & "'," & txtProvCod & ",dbo.MyGetdate()," _
          & CCur(txtSubTotal) & "," & CCur(txtDescuento) & "," & CCur(txtImpuestos) & ",0," _
          & CCur(txtTotal) & ",'" & txtNotas & "','P','" & glogon.Usuario & "',dbo.MyGetdate(),'" & fxCodigoCbo(cbo) & "')"
   Call ConectionExecute(strSQL)

  Call Bitacora("Registra", "Devolucion Fact Compra.: " & vCodigo & " Dev: " & vDevolucion)

  txtDevolucion.Enabled = True

End If

'Guardar Detalle de la Orden
strSQL = "delete cpr_compra_devDet where cod_compra_dev = '" & vDevolucion & "'"
Call ConectionExecute(strSQL)

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  
  vGrid.col = 3
  curCantidad = CCur(IIf((vGrid.Text = ""), 0, vGrid.Text))
  
  vGrid.col = 1
  
  If vGrid.Text <> "" And curCantidad > 0 Then
    
    vGrid.col = 1
    vCodPro = Trim(vGrid.Text)
    strSQL = "insert cpr_compra_devDet(linea,cod_compra_dev,cod_producto,cantidad,cod_bodega" _
           & ",precio,imp_ventas,imp_consumo) values(" & i & ",'" & vDevolucion & "','" _
           & vGrid.Text & "'," & curCantidad & ",'"
    vGrid.col = 4
    vCodBodega = Trim(vGrid.Text)
    strSQL = strSQL & vGrid.Text & "',"
    vGrid.col = 5
    curPrecio = CCur(vGrid.Text)
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.col = 6
    curImpVentas = CCur(vGrid.Text)
    curImpConsumo = 0
    strSQL = strSQL & CCur(vGrid.Text) & ",0)"
    Call ConectionExecute(strSQL)
  
    'Actualizar Aqui el Inventario y la Factura
    vGrid.col = 3
    strSQL = "update cpr_compras_detalle set cantidad_devuelta = isnull(cantidad_devuelta,0) + " _
           & CCur(vGrid.Text) & " where linea = " & i & " and cod_factura = '" & vCodigo _
           & "' and cod_proveedor = " & txtProvCod
    Call ConectionExecute(strSQL)
    
    Call sbInvInventario(vCodPro, curCantidad, vCodBodega, vDevolucion, "Compra.Dev", vFecha _
            , curPrecio, curImpConsumo, curImpVentas, "S")
    
  End If
Next i

'Crear Cargo Flotante por el Monto de la Devolucion
   strSQL = "select isnull(max(ID),0) as ultimo from cxp_cargosper where cod_proveedor = " & txtProvCod
   Call OpenRecordSet(rs, strSQL)
     i = rs!ultimo + 1
   rs.Close
   
   strSQL = "insert cxp_cargosper(id,cod_proveedor,cod_cargo,tipo,valor,vence,saldo,concepto,detalle,recaudado)" _
          & " values(" & i & "," & txtProvCod & ",'" & fxCodigoCbo(cbo) & "','M'," & CCur(txtTotal) _
          & ",'" & Format(fxFechaServidor, "yyyy/mm/dd") & "'," & CCur(txtTotal) & ",'DEVOLUCION MERCADERIA - FACTURA DE COMPRA','" _
          & "FACTURA : " & txtCodigo & vbCrLf & "USUARIO :" & glogon.Usuario & "',0)"
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Cargo Adicional a Prov:" & txtProvCod & " Sec: " & i)
   
   strSQL = "update cxp_proveedores set saldo = isnull(saldo,0) - " & CCur(txtTotal.Text) _
          & " where cod_proveedor = " & txtProvCod
   Call ConectionExecute(strSQL)
   
'Fin de la Actualizacion de la CxP

strSQL = "update cpr_compras set estado = 'D' where cod_compra = '" & Format(txtCompra, vMascara) & "'"
Call ConectionExecute(strSQL)


Call sbToolBar(tlb, "activo")
Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
   'no se pueden Ejecutar Borrados en Ordenes
'  strSQL = "delete cxp_proveedores where cod_proveedor = " & vCodigo
'  Call ConectionExecute(strSQL)

'  Call Bitacora("Elimina", "ER ESPECIAL : " & vCodigo & " EMP: " & vParametros.CodigoEmpresa)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_factura"
  gBusquedas.Orden = "cod_factura"
  gBusquedas.Consulta = "select E.cod_factura,P.descripcion as Proveedor,E.total" _
            & " from cpr_compras E inner join cxp_Proveedores P on E.cod_proveedor = P.cod_proveedor"
  gBusquedas.Filtro = " and E.cod_proveedor = " & txtProvCod
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsultaFac(gBusquedas.Resultado)
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" Then Call sbConsultaFac(txtCodigo)
End Sub


Private Sub txtCompra_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtCompra = Format(txtCompra, vMascara)
End If
End Sub

Private Sub txtDescuento_GotFocus()
On Error GoTo vError
txtDescuento = CCur(txtDescuento)
Exit Sub
vError:
  MsgBox "Información del Descuento no es válida...", vbCritical
End Sub

Private Sub txtDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTotal.SetFocus
End Sub

Private Sub txtDescuento_LostFocus()
On Error GoTo vError

txtDescuento = Format(CCur(txtDescuento), "Standard")
txtTotal = Format(CCur(txtSubTotal) + CCur(txtImpuestos) - CCur(txtDescuento), "Standard")

Exit Sub
vError:
  MsgBox "Información del Descuento no es válida...", vbCritical
End Sub

Private Sub txtDevolucion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Mascara = vMascara
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "D.cod_compra_dev"
  gBusquedas.Orden = "D.cod_compra_dev"
  gBusquedas.Consulta = "select D.cod_compra_dev,P.descripcion as Proveedor,D.cod_factura,D.notas,D.fecha" _
          & " from cpr_compras_dev D inner join cxp_proveedores P on D.cod_proveedor = P.cod_proveedor"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtDevolucion = gBusquedas.Resultado
  If txtDevolucion <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtDevolucion_LostFocus()
If txtDevolucion <> "" Then Call sbConsulta(txtDevolucion)
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
vError:
End Sub


Private Sub sbCalculaTotales()
Dim curSubTotal As Currency, curIV As Currency
Dim curTmpPrecio As Currency, curTmpIV As Currency, curTmpCant As Currency
Dim i As Integer, lng As Long

'**********************************************    OJO
'Revisar esta formula por la situacion del descuento, si es antes o despues del
'impuesto de ventas, por ahora está despues del impuesto
curSubTotal = 0
curIV = 0

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 3
 If vGrid.Text <> "" Then
   curTmpCant = CCur(vGrid.Text)
   If curTmpCant > 0 Then
    vGrid.col = 5
    curTmpPrecio = CCur(vGrid.Text)
    vGrid.col = 6
    curTmpIV = CCur(vGrid.Text) / 100

    curSubTotal = curSubTotal + (curTmpCant * curTmpPrecio)
    curIV = curIV + ((curTmpCant * curTmpPrecio) * curTmpIV)
    
    vGrid.col = 7
    vGrid.Text = CStr((curTmpCant * curTmpPrecio) * (curTmpIV + 1))
    
   End If
 End If
Next lng

txtSubTotal = Format(curSubTotal, "Standard")
txtImpuestos = Format(curIV, "Standard")
txtTotal = Format(curSubTotal + curIV - CCur(txtDescuento), "Standard")

End Sub



Private Sub txtProvCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProveedor.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtProvCod = gBusquedas.Resultado
  txtProveedor = gBusquedas.Resultado2
End If

End Sub

Private Sub txtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtProvCod = gBusquedas.Resultado
  txtProveedor = gBusquedas.Resultado2
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(7) As Variant, x As Integer

'Abrir Nueva Linea
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    Call sbCalculaTotales
  End If
End If

'Consular Articulo
'If vGrid.ActiveCol = 1 And KeyCode = vbKeyF4 Then
'   frmBusquedaArticulos.Show vbModal
'   vGrid.Row = vGrid.ActiveRow
'   vGrid.Col = 1
'   vGrid.Text = gBusquedas.Resultado
'End If

'Consular Bodegas
If vGrid.ActiveCol = 4 And KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_bodega"
   gBusquedas.Orden = "cod_bodega"
   gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
   gBusquedas.Filtro = " and permite_salidas = 1"
   frmBusquedas.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 4
   vGrid.Text = gBusquedas.Resultado
End If


''Borrar una linea
'If KeyCode = vbKeyDelete Then
'  vGrid.Row = vGrid.ActiveRow
'  vGrid.Col = 7
'  For lng = vGrid.ActiveRow To vGrid.MaxRows
'     vGrid.Row = lng + 1
'     For x = 1 To 7
'        vGrid.Col = x
'        vTemp(x) = vGrid.Text
'     Next x
'
'     vGrid.Row = lng
'     For x = 1 To 7
'       vGrid.Col = x
'       vGrid.Text = vTemp(x)
'     Next x
'  Next lng
'  vGrid.MaxRows = vGrid.MaxRows - 1
'  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
'  Call sbCalculaTotales
'End If


End Sub



Private Sub vGrid_KeyPress(KeyAscii As Integer)
Dim curCantidad As Currency, curPrecio As Currency, curIV As Currency

On Error GoTo vError
'Calcula Total
Select Case vGrid.ActiveCol
  Case 3, 5, 6
    vGrid.Row = vGrid.ActiveRow
    vGrid.col = 3
    curCantidad = CCur(vGrid.Text)
    vGrid.col = 5
    curPrecio = CCur(vGrid.Text)
    vGrid.col = 6
    curIV = CCur(vGrid.Text)
    vGrid.col = 7
    vGrid.Text = (curPrecio * curCantidad) + ((curPrecio * curCantidad) * (curIV / 100))
   Call sbCalculaTotales
  Case Else 'No Aplica
End Select
vError:
End Sub








