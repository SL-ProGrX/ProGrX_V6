VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInvTranRequisiciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requisiciones"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10455
   Begin VB.CheckBox chkPlantilla 
      Appearance      =   0  'Flat
      Caption         =   "Plantilla ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   480
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5355
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Solicitado Por"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3246
            MinWidth        =   3246
            Object.ToolTipText     =   "Fecha Solicitud"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Autorizado Por"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3246
            MinWidth        =   3246
            Object.ToolTipText     =   "Fecha Autorizacion"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cboCausa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   960
      Width           =   4695
   End
   Begin VB.TextBox txtNotas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1320
      Width           =   7935
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4992
      Width           =   1575
   End
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtDocumento 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8280
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRequisiciones.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRequisiciones.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRequisiciones.frx":0BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRequisiciones.frx":0F0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   408
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10452
      _ExtentX        =   18441
      _ExtentY        =   714
      _CBWidth        =   10455
      _CBHeight       =   405
      _Version        =   "6.7.9816"
      Child1          =   "tlb"
      MinHeight1      =   270
      Width1          =   4065
      NewRow1         =   0   'False
      Child2          =   "tlbProcesos"
      MinHeight2      =   315
      Width2          =   2445
      NewRow2         =   0   'False
      MinHeight3      =   360
      Width3          =   1995
      NewRow3         =   0   'False
      Begin MSComctlLib.Toolbar tlbProcesos 
         Height          =   312
         Left            =   4224
         TabIndex        =   19
         Top             =   48
         Width           =   2292
         _ExtentX        =   4048
         _ExtentY        =   556
         ButtonWidth     =   1820
         ButtonHeight    =   550
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Plantillas"
               Key             =   "Plantilla"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Procesar"
               Key             =   "Procesar"
               Object.ToolTipText     =   "Ejecuta la Transaccion"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   264
         Left            =   132
         TabIndex        =   18
         Top             =   72
         Width           =   3912
         _ExtentX        =   6906
         _ExtentY        =   476
         ButtonWidth     =   487
         ButtonHeight    =   466
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
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repBoleta"
                     Text            =   "Boleta "
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repListadoGeneral"
                     Text            =   "Listado General"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   40
         Width           =   2100
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1179649
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2772
      Left            =   0
      TabIndex        =   21
      Top             =   2040
      Width           =   10332
      _Version        =   524288
      _ExtentX        =   18225
      _ExtentY        =   4890
      _StockProps     =   64
      ArrowsExitEditMode=   -1  'True
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
      SpreadDesigner  =   "frmInvRequisiciones.frx":1228
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Causa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   10
      Left            =   1560
      TabIndex        =   13
      Top             =   960
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   1560
      TabIndex        =   12
      Top             =   1320
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Sub Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   5760
      TabIndex        =   11
      Top             =   4992
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "# Boleta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   7200
      TabIndex        =   9
      Top             =   480
      Width           =   732
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   10320
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "# Documento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   7200
      TabIndex        =   8
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label lblCantidad 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   2520
      TabIndex        =   7
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label lblLineas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   2295
   End
End
Attribute VB_Name = "frmInvTranRequisiciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vScroll As Boolean

Private Sub cboCausa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
'    strSQL = "select Top 1 Boleta from pv_invTransac" _
'           & " where Tipo = '" & vTipo & "'"
'
'    If FlatScrollBar.Value = 1 Then
'       strSQL = strSQL & " and boleta > '" & Format(txtCodigo, vMascara) & "' order by Boleta asc"
'    Else
'       strSQL = strSQL & " and boleta < '" & Format(txtCodigo, vMascara) & "' order by Boleta desc"
'    End If
'
'    Call OpenRecordSet(rs, strSQL)
'    If Not rs.EOF And Not rs.BOF Then
'      Call sbConsulta(rs!Boleta)
'    End If
'    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Activate()
vModulo = 32
End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 32
 vGrid.AppearanceStyle = fxGridStyle

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True

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

vCodigo = 0
txtCodigo = ""

txtDocumento = ""
txtFecha = Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss")

txtNotas = ""

txtEstado = ""
txtEstado.Tag = "S"

Call sbInvESCombo("R", cboCausa)

vGrid.MaxRows = 1
vGrid.MaxCols = 5
For i = 1 To vGrid.MaxCols
  vGrid.col = i
  vGrid.Text = ""
Next

txtSubTotal = 0

txtCodigo.Enabled = True

With StatusBarX.Panels
  .Item(1).Text = ""
  .Item(2).Text = ""
  .Item(3).Text = ""
  .Item(4).Text = ""
End With

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      cboCausa.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNotas.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
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

Private Sub sbConsulta(lngCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select X.*,(rtrim(C.cod_entsal) + ' - ' + C.descripcion) as Causa" _
       & " from pv_requisiciones X inner join pv_entrada_salida C on X.cod_entsal = C.cod_entsal" _
       & " where X.cod_requisicion = " & lngCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!cod_requisicion
  txtCodigo = rs!cod_requisicion
  
  txtDocumento = rs!Documento
  
  cboCausa.Text = Trim(rs!Causa)
  
  txtFecha = Format(rs!genera_fecha, "yyyy/mm/dd hh:mm:ss")
  
  txtNotas = rs!notas & ""
  
  Select Case rs!Estado
    Case "S" 'Solicitada
        txtEstado = "Solicitada"
    Case "A" 'Autorizada
        txtEstado = "Autorizada"
    Case "R" 'Rechazada
        txtEstado = "Rechazada"
  End Select
  txtEstado.Tag = rs!Estado
    
  With StatusBarX.Panels
    .Item(1) = rs!genera_user
    .Item(2) = rs!genera_fecha
    .Item(3) = rs!Autoriza_user & ""
    .Item(4) = rs!Autoriza_Fecha & ""
  End With
    

  strSQL = "select D.cod_producto,P.descripcion,D.cantidad,P.costo_regular,(D.cantidad * P.costo_regular) as Total" _
         & ",isnull(D.despacho,0) as Despacho" _
         & " from pv_requi_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " where D.cod_requisicion = " & rs!cod_requisicion _
         & " order by D.Linea"
  
  Call sbCargaGridLocal(vGrid, 5, strSQL)
  
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

Public Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String, Optional vBorra As Boolean = True)
Dim rs As New ADODB.Recordset, i As Integer

If vBorra Then
    vGrid.MaxCols = vGridMaxCol
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.col = i
     vGrid.Text = ""
    Next i
End If

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    vGrid.Text = CStr(rs.Fields(i - 1).Value)
  Next i
  
  vGrid.col = 3
  vGrid.TextTip = TextTipFixed
  vGrid.CellNote = "Unidades Despachadas : " & rs!despacho & vbCrLf & " Unidades Pendientes : " _
                 & (rs!cantidad - rs!despacho)
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

End Sub



Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

On Error GoTo vError

vMensaje = fxInvVerificaLineaDetalle(vGrid, 3, "R", 1)

If Not fxInvPeriodos(fxFechaServidor) Then vMensaje = vMensaje & vbCrLf & " - El periodo en el que desea realizar el movimiento se encuentra cerrado ..."

vError:

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, i As Integer
Dim rs As New ADODB.Recordset
Dim vFecha As Date, curCantidad As Currency

On Error GoTo vError


If txtEstado.Tag <> "S" Then
  MsgBox "Esta requisicion no esta solicitada, No se puede Modificar...", vbExclamation
  Exit Sub
End If

vFecha = fxFechaServidor

If vEdita Then
    strSQL = "update pv_requisiciones set cod_entsal = '" & fxCodigoCbo(cboCausa) _
           & "',genera_fecha = '" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "',documento = '" _
           & txtDocumento & "', notas = '" & txtNotas & "',genera_user = '" & glogon.Usuario _
           & "',plantilla = " & chkPlantilla.Value _
           & " where cod_requisicion = " & vCodigo
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Requisicion Cod: " & vCodigo)

Else
    
    strSQL = "select isnull(max(cod_requisicion),0)+1 as Ultimo from pv_requisiciones"
    Call OpenRecordSet(rs, strSQL)
      vCodigo = rs!ultimo
    rs.Close
    txtCodigo = vCodigo
    
    strSQL = "insert pv_requisiciones(cod_requisicion,cod_entsal,genera_fecha,documento,notas" _
           & ",genera_user,estado,plantilla)" _
           & " values(" & vCodigo & ",'" & fxCodigoCbo(cboCausa) & "','" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") _
           & "','" & txtDocumento & "','" & txtNotas & "','" & glogon.Usuario & "','S'," & chkPlantilla.Value & ")"
    
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Registra", "Requisicion Cod: " & vCodigo)

End If

txtCodigo.Enabled = True

'Guardar Detalle de la requisicion
strSQL = "delete pv_requi_detalle where cod_requisicion = " & vCodigo
Call ConectionExecute(strSQL)

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  
  vGrid.col = 3
  curCantidad = CCur(IIf((vGrid.Text = ""), 0, vGrid.Text))
  
  vGrid.col = 1
  
  If vGrid.Text <> "" And curCantidad > 0 Then
    
    vGrid.col = 1
    strSQL = "insert pv_requi_detalle(linea,cod_requisicion,cod_producto,cantidad,despacho)" _
           & " values(" & i & "," & vCodigo & ",'" & vGrid.Text & "'," & curCantidad & ",0)"
    Call ConectionExecute(strSQL)
  
  End If
Next i

'*********************************** fin

txtEstado.Tag = "S"
txtEstado = "Solicitada"

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
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim i As Integer, vSQL As String

vSQL = ""

Select Case UCase(ButtonMenu.Key)
  Case "REPBOLETA"
     
     i = MsgBox("Desea visualizar solo el Traslado Actual", vbYesNo)
     If i = vbYes Then vSQL = "{PV_TRASLADOS.cod_requisicion} = " & txtCodigo
     
     Call sbInvReportes("BoletaTraslado", "Boleta de Traslados", "", vSQL)

  Case "REPLISTADOGENERAL"
     Call sbInvReportes("TrasladoListado", "TRASLADOS", "Listado General", vSQL)

End Select


End Sub

Private Sub tlbProcesos_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

gBusquedas.Columna = "cod_requisicion"
gBusquedas.Orden = "cod_requisicion"
gBusquedas.Consulta = "select cod_requisicion,genera_user,genera_fecha,documento,notas" _
          & " from pv_requisiciones"
gBusquedas.Filtro = " and plantilla = 1"
gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""
frmBusquedas.Show vbModal

If gBusquedas.Resultado <> "" Then
  Call sbLimpiaPantalla
  Call sbConsulta(CLng(gBusquedas.Resultado))
  txtCodigo = ""
  chkPlantilla.Value = vbUnchecked
  vEdita = False
  txtCodigo.Enabled = False
  cboCausa.SetFocus
  Call sbToolBar(tlb, "edicion")
End If


Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCausa.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_requisicion"
  gBusquedas.Orden = "cod_requisicion"
  gBusquedas.Consulta = "select cod_requisicion,documento,notas from pv_requisiciones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
End Sub

Private Sub sbCalculaTotales()
Dim curSubTotal As Currency
Dim curTmpPrecio As Currency, curTmpCant As Currency
Dim i As Integer, lng As Long
Dim iLineas As Integer, curCantidad As Currency


'**********************************************    OJO
'Revisar esta formula por la situacion del descuento, si es antes o despues del
'impuesto de ventas, por ahora está despues del impuesto

curSubTotal = 0

iLineas = 0
curCantidad = 0

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 3
 If vGrid.Text <> "" Then
    curTmpCant = CCur(vGrid.Text)
    vGrid.col = 4
    curTmpPrecio = CCur(vGrid.Text)

    curSubTotal = curSubTotal + (curTmpCant * curTmpPrecio)
    curCantidad = curCantidad + curTmpCant
    iLineas = iLineas + 1
 End If
Next lng

txtSubTotal = Format(curSubTotal, "Standard")

lblLineas.Caption = "Líneas   : " & iLineas
lblCantidad.Caption = "Cantidad : " & Format(curCantidad, "Standard")

End Sub

Private Sub sbConsultaArticulo(fila As Long, Columna As Integer, vCriterio As String)
Dim strSQL As String, rs As New ADODB.Recordset, vPaso As Boolean

'Busquedas
'1. Por Codigo del Articulo
'2. Por Codigo de Barras
'3. Por Codigo del Fabricante
vPaso = False

strSQL = "select cod_producto,descripcion,costo_regular,impuesto_ventas from pv_productos" _
       & " where cod_producto = '" & vCriterio & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then vPaso = True

If Not vPaso Then
  rs.Close
  strSQL = "select cod_producto,descripcion,costo_regular,impuesto_ventas from pv_productos" _
         & " where cod_barras = '" & vCriterio & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then vPaso = True
End If

If Not vPaso Then
  rs.Close
  strSQL = "select cod_producto,descripcion,costo_regular,impuesto_ventas from pv_productos" _
         & " where cod_fabricante = '" & vCriterio & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then vPaso = True
End If

If Not vPaso Then
  MsgBox "No se encontró el Articulo en la Base de Datos...", vbExclamation
Else
  vGrid.Row = fila
  vGrid.col = 1
  vGrid.Text = rs!cod_producto
  vGrid.col = 2
  vGrid.Text = rs!Descripcion
  vGrid.col = 4
  vGrid.Text = CStr(rs!costo_regular)
  vGrid.col = 3
  vGrid.Text = 1
End If
rs.Close


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(8) As Variant, x As Integer

'Abrir Nueva Linea
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    Call sbCalculaTotales
  End If
End If

'Consulta Articulo
If vGrid.ActiveCol = 1 And KeyCode = vbKeyReturn Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  Call sbConsultaArticulo(vGrid.ActiveRow, vGrid.ActiveCol, vGrid.Text)
End If

'Consular Articulo
If vGrid.ActiveCol = 1 And KeyCode = vbKeyF4 Then
   frmBusquedaArticulos.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 1
   vGrid.Text = gBusquedas.Resultado
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = vGrid.MaxCols
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To vGrid.MaxCols
        vGrid.col = x
        vTemp(x) = vGrid.Text
     Next x

     vGrid.Row = lng
     For x = 1 To vGrid.MaxCols
       vGrid.col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  Call sbCalculaTotales
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


End Sub


Private Sub vGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim curCantidad As Currency, curPrecio As Currency

On Error GoTo vError
'Calcula Total
Select Case vGrid.ActiveCol
  Case 3, 4
    vGrid.Row = vGrid.ActiveRow
    vGrid.col = 3
    curCantidad = CCur(vGrid.Text)
    vGrid.col = 4
    curPrecio = CCur(vGrid.Text)
    vGrid.col = 5
    vGrid.Text = (curPrecio * curCantidad)
   Call sbCalculaTotales
  Case Else 'No Aplica
End Select
vError:

End Sub
