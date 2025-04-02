VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCntX_AsientosInv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asientos de Ajuste para Inventarios Periodicos"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   11325
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCredito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtDebito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   7770
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtDiferencia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtPeriodo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8430
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "(F4) Descripción del Periodo"
      Top             =   480
      Width           =   2865
   End
   Begin VB.TextBox txtAnio 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7890
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   4
      ToolTipText     =   "(F4) Año del Periodo"
      Top             =   480
      Width           =   525
   End
   Begin VB.TextBox txtMes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7560
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   3
      ToolTipText     =   "(F4) Mes del Periodo"
      Top             =   480
      Width           =   315
   End
   Begin VB.TextBox txtNAsiento 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   990
      MaxLength       =   15
      TabIndex        =   2
      ToolTipText     =   "(F4) Número del Asiento"
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   990
      MaxLength       =   60
      TabIndex        =   1
      Top             =   840
      Width           =   10305
   End
   Begin VB.TextBox txtNotas 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   675
      Left            =   990
      MaxLength       =   60
      TabIndex        =   0
      Top             =   1200
      Width           =   10305
   End
   Begin MSComCtl2.DTPicker dtpAsientoFecha 
      Height          =   315
      Left            =   4980
      TabIndex        =   6
      Top             =   480
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   160628739
      CurrentDate     =   36212
   End
   Begin MSComctlLib.ImageList ImageListMenu 
      Left            =   10680
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_AsientosInv.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_AsientosInv.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_AsientosInv.frx":0D2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11325
      _ExtentX        =   19976
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
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3372
      Left            =   0
      TabIndex        =   13
      Top             =   2040
      Width           =   11292
      _Version        =   524288
      _ExtentX        =   19918
      _ExtentY        =   5948
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   491
      ScrollBars      =   2
      SpreadDesigner  =   "frmCntX_AsientosInv.frx":113E
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   0
   End
   Begin VB.Label lblAsientoEstado 
      Caption         =   "Estado del Asiento."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   5640
      Width           =   3795
   End
   Begin VB.Label lsblr 
      Caption         =   "Diferencia:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4110
      TabIndex        =   18
      Top             =   5670
      Width           =   915
   End
   Begin VB.Label Label5 
      Caption         =   "Totales:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6990
      TabIndex        =   17
      Top             =   5670
      Width           =   645
   End
   Begin VB.Label Label15 
      Caption         =   "Fecha Asiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3780
      TabIndex        =   11
      Top             =   480
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6600
      TabIndex        =   10
      Top             =   480
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "Nº Asiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   9
      Top             =   480
      Width           =   825
   End
   Begin VB.Label Label4 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   8
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   705
   End
End
Attribute VB_Name = "frmCntX_AsientosInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type xUltimos
   Anio As Long
   Mes As Integer
   NumAsiento As String
   Detalle As String
   Documento As String
   fecha As Date
End Type
Dim vEdita As Boolean, vBusca As Integer, vUltimos As xUltimos
Dim vScroll As Boolean


Private Sub sbRefrescaInformacion()
Dim strResultado As String
On Error GoTo vError
txtAnio = Val(txtAnio)
dtpAsientoFecha = CDate(txtAnio & "/" & txtMes & "/01")
  Select Case Val(txtMes)
    Case 1
        strResultado = "ENERO DEL " & txtAnio
    Case 2
        strResultado = "FEBRERO DEL " & txtAnio
    Case 3
        strResultado = "MARZO DEL " & txtAnio
    Case 4
        strResultado = "ABRIL DEL " & txtAnio
    Case 5
        strResultado = "MAYO DEL " & txtAnio
    Case 6
        strResultado = "JUNIO DEL " & txtAnio
    Case 7
        strResultado = "JULIO DEL " & txtAnio
    Case 8
        strResultado = "AGOSTO DEL " & txtAnio
    Case 9
        strResultado = "SETIEMBRE DEL " & txtAnio
    Case 10
        strResultado = "OCTUBRE DEL " & txtAnio
    Case 11
        strResultado = "NOVIEMBRE DEL " & txtAnio
    Case 12
        strResultado = "DICIEMBRE DEL " & txtAnio
  End Select

  txtPeriodo = strResultado

Exit Sub

vError:
End Sub

Private Sub dtpAsientoFecha_Change()
txtAnio = Year(dtpAsientoFecha.Value)
txtMes = Month(dtpAsientoFecha.Value)
End Sub

Private Sub dtpAsientoFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End Sub

Private Sub sbLimpiezaParcial(iCodigo As Integer)
vGrid.MaxRows = 0
vGrid.MaxRows = 1

txtDescripcion = ""

Select Case iCodigo
  Case 1 'Cambia el Tipo de Asiento
    txtNAsiento = ""
  Case 2 'Cambia el periodo
    txtNAsiento = ""
   
End Select

End Sub


Private Sub Form_Activate()
vModulo = 20
End Sub

Private Sub Form_Load()

vModulo = 20

Set Me.Icon = frmContenedor.Icon

vGrid.AppearanceStyle = fxGridStyle


vEdita = True
Call sbToolBarIconos(tlb)
Call sbToolBar(tlb, "nuevo")

vUltimos.Anio = gCntX_Parametros.PeriodoAnio
vUltimos.Mes = gCntX_Parametros.PeriodoMes
  
vEdita = False
Call sbLimpiaPantalla
Call sbToolBar(tlb, "edicion")
 
Call Formularios(Me)
Call RefrescaTags(Me)
 
End Sub

Private Function fxVerificaAsiento() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
Dim vMensaje As String, lng As Long

'Verificar Periodo
'Fecha del Asiento vrs Periodo
'Numero de Asiento
'CntX_Cuentas (En el Detalle)

fxVerificaAsiento = True
vMensaje = ""

'strSQL = "select isnull(count(*),0) as existe from CntX_Periodos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
'       & " and anio = " & txtAnio & " and mes = " & txtMes & " and estado = 'P'"
'Call OpenRecordSet(rsX, strSQL, 0)
'  If rsX!existe = 0 Then vMensaje = vMensaje & vbCrLf & "- Periodo Indicado se encuentra Cerrado o No se ha creado..."
'rsX.Close


If Month(dtpAsientoFecha) <> txtMes Then vMensaje = vMensaje & vbCrLf & "- El Mes del Periodo no se encuentra en la fecha del Asiento..."

If Year(dtpAsientoFecha) <> txtAnio Then vMensaje = vMensaje & vbCrLf & "- El Año del Periodo no se encuentra en la fecha del Asiento..."

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 1
 If vGrid.Text <> "" Then
   vGrid.col = 2
   If vGrid.Text = "" Then
      vGrid.col = 1
      vMensaje = vMensaje & vbCrLf & "- Cuenta " & vGrid.Text & " No Existe"
   End If
 End If
Next lng

If Len(vMensaje) > 0 Then
  fxVerificaAsiento = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbLimpiaPantalla()
vBusca = 1
txtAnio = vUltimos.Anio
txtMes = vUltimos.Mes
txtPeriodo = ""
txtCredito = 0
txtDebito = 0
txtDescripcion = ""
txtDiferencia = 0
txtNAsiento = ""
vGrid.MaxRows = 0
vGrid.MaxRows = 1
vGrid.MaxCols = 7

Call sbRefrescaInformacion

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      Call sbToolBar(tlb, "edicion")
    
    Case "MODIFICAR", "EDITAR"
        vEdita = True
        txtDescripcion.SetFocus
        Call sbToolBar(tlb, "edicion")
    
    Case "BORRAR"
        Call sbBorrar
      
    Case "GUARDAR", "SALVAR"
      Call sbGuardar
    
    Case "DESHACER"
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
    
    Case "CONSULTAR"
       Select Case vBusca
         Case 3 'Numero de Asiento
            gBusquedas.Columna = "Num_Asiento"
            gBusquedas.Orden = "Num_Asiento"
            gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                              & " and anio = " & txtAnio & " and mes = " & txtMes
            gBusquedas.Consulta = "select Num_asiento,descripcion from CntX_Inv_Asientos"
            frmBusquedas.Show vbModal
            txtNAsiento = gBusquedas.Resultado
            Call sbConsultaAsiento(txtNAsiento)
            
         Case 4 'Descripcion del numero de Asiento
            gBusquedas.Columna = "Descripcion"
            gBusquedas.Orden = "Descripcion"
            gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                              & " and anio = " & txtAnio & " and mes = " & txtMes
            gBusquedas.Consulta = "select Num_asiento,descripcion from CntX_Inv_Asientos"
            frmBusquedas.Show vbModal
            txtNAsiento = gBusquedas.Resultado
            Call sbConsultaAsiento(txtNAsiento)
       End Select

    Case "REPORTES"
      
'      strSQL = "{Cntx_Asientos.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
'             & " AND {Cntx_Asientos.TIPO_ASIENTO} = '" & txtCAsiento & "' AND " _
'             & " {Cntx_Asientos.NUM_ASIENTO} = '" & txtNAsiento & "'"
'
'      Call sbCntX_Reportes("ASIENTO", strSQL)
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
    
    Case "CERRAR"
      Unload Me
End Select

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    If i = 1 Then
        vGrid.Text = fxCntX_CuentaFormato(True, CStr(rs.Fields(i - 1).Value))
    Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value)
    End If
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop

rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub sbConsultaAsiento(strNumero As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from CntX_Inv_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and num_asiento = '" & strNumero & "'"

Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
 
  vUltimos.Anio = rs!Anio
  vUltimos.Mes = rs!Mes
  vUltimos.NumAsiento = rs!Num_Asiento
  vUltimos.fecha = rs!fecha_asiento
  
  'llenar datos en pantalla
  
  txtAnio = vUltimos.Anio
  txtMes = vUltimos.Mes
  Call sbRefrescaInformacion
  
  txtDescripcion = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
  txtNotas = IIf(IsNull(rs!Notas), "", rs!Notas)

  dtpAsientoFecha.Value = vUltimos.fecha
  txtNAsiento = vUltimos.NumAsiento
  
  lblAsientoEstado.Caption = "Este Asiento se encuentra en Cola"
  
  
strSQL = "select A.cod_cuenta,B.descripcion,documento,detalle,monto_debito,monto_credito,num_linea" _
       & " from CntX_Inv_Asientos_detalle A inner join CntX_Cuentas B on A.cod_cuenta = B.cod_cuenta" _
       & " and A.cod_contabilidad = B.cod_contabilidad" _
       & " where A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and num_asiento = '" & vUltimos.NumAsiento & "'" _
       & " order by num_linea"
   
  Call sbCargaGridLocal(vGrid, 7, strSQL)
 
  Call sbSumaDebitosCreditos

End If

rs.Close
Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbGuardar()
Dim strSQL As String, lng As Long

On Error GoTo vError

If fxVerificaAsiento Then
    If vEdita Then
      
      strSQL = "update CntX_Inv_Asientos set descripcion = '" & UCase(txtDescripcion) _
             & "',fecha_asiento = '" & Format(dtpAsientoFecha.Value, "yyyy/mm/dd") _
             & "',notas = '" & Trim(txtNotas) _
             & "' where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
             & " and num_asiento = '" & txtNAsiento & "'"
      Call ConectionExecute(strSQL, 0)
     
      strSQL = "delete CntX_Inv_Asientos_detalle where cod_contabilidad = " _
             & gCntX_Parametros.CodigoConta & " and num_asiento = '" & txtNAsiento & "'"
      Call ConectionExecute(strSQL, 0)
    
      For lng = 1 To vGrid.MaxRows
        vGrid.Row = lng
        vGrid.col = 1
        If vGrid.Text <> "" Then
            strSQL = "insert into CntX_Inv_Asientos_detalle(num_asiento,cod_contabilidad" _
                   & ",num_linea,cod_cuenta,documento,detalle,monto_debito,monto_credito" _
                   & ") values('" & txtNAsiento & "'," _
                   & gCntX_Parametros.CodigoConta & "," & lng & ",'"
            vGrid.Row = lng
            vGrid.col = 1
            strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "','"
            vGrid.col = 3
            strSQL = strSQL & vGrid.Text & "','"
            vGrid.col = 4
            strSQL = strSQL & vGrid.Text & "',"
            vGrid.col = 5
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ","
            vGrid.col = 6
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ")"

            Call ConectionExecute(strSQL, 0)
              
         End If 'vgrid.Text <> ""
       
       Next lng
    
      Call Bitacora("Modifica", "Asiento INV.PER. : " & txtNAsiento & " Conta." & gCntX_Parametros.CodigoConta)
    
    
    Else 'Inserta
       strSQL = "insert into CntX_Inv_Asientos(cod_contabilidad,num_asiento,anio,mes" _
              & ",fecha_asiento,descripcion,notas) values(" _
              & gCntX_Parametros.CodigoConta & ",'" & txtNAsiento & "'," & txtAnio _
              & "," & txtMes & ",'" & Format(dtpAsientoFecha.Value, "yyyy/mm/dd") & "','" & UCase(txtDescripcion) _
              & "','" & Trim(txtNotas) & "')"
       Call ConectionExecute(strSQL, 0)
       
       For lng = 1 To vGrid.MaxRows
        vGrid.Row = lng
        vGrid.col = 1
        If vGrid.Text <> "" Then
            strSQL = "insert into CntX_Inv_Asientos_detalle(num_asiento,cod_contabilidad,cod_unidad" _
                   & ",num_linea,cod_cuenta,documento,detalle,monto_debito,monto_credito" _
                   & ") values('" & txtNAsiento & "'," _
                   & gCntX_Parametros.CodigoConta & ",'OC'," & lng & ",'"
            vGrid.Row = lng
            vGrid.col = 1
            strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "','"
            vGrid.col = 3
            strSQL = strSQL & vGrid.Text & "','"
            vGrid.col = 4
            strSQL = strSQL & vGrid.Text & "',"
            vGrid.col = 5
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ","
            vGrid.col = 6
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ")" _
          
            Call ConectionExecute(strSQL, 0)
              
         End If 'vgrid.Text <> ""
       
       Next lng
       
       
        Call Bitacora("Registra", "Asiento INV.PER. : " & txtNAsiento & " Conta." & gCntX_Parametros.CodigoConta)
        
    End If 'Si Inserta o Actualiza

        Call sbToolBar(tlb, "activo")
        Call sbConsultaAsiento(txtNAsiento)
        
        vEdita = True
        
        MsgBox "Información guardada satisfactoriamente...", vbInformation


End If 'Verificacion del Asiento

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String
On Error GoTo vError
i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
  strSQL = "delete CntX_Inv_Asientos_detalle where cod_contabilidad = " _
         & gCntX_Parametros.CodigoConta & " and num_asiento = '" & txtNAsiento & "'"
  Call ConectionExecute(strSQL, 0)
  
  strSQL = "delete CntX_Inv_Asientos where cod_contabilidad = " _
         & gCntX_Parametros.CodigoConta & " and num_asiento = '" & txtNAsiento & "'"
  Call ConectionExecute(strSQL, 0)
  

  Call Bitacora("Elimina", "Asiento INV.PER. : " & txtNAsiento & " Conta." _
                  & gCntX_Parametros.CodigoConta)

  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtAnio_Change()
 Call sbRefrescaInformacion
End Sub

Private Sub txtDescripcion_GotFocus()
vBusca = 4
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
End Sub

Private Sub txtMes_Change()
 Call sbRefrescaInformacion
End Sub

Private Sub txtNAsiento_GotFocus()
vBusca = 3
End Sub

Private Sub txtNAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 dtpAsientoFecha.SetFocus
 Call sbConsultaAsiento(txtNAsiento)
End If
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
End Sub


Private Function fxVerificaCuenta(strCuenta As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select isnull(count(*),0) as Existe from CntX_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_cuenta = '" & strCuenta & "' and acepta_movimientos = 1"

Call OpenRecordSet(rsX, strSQL, 0)
fxVerificaCuenta = IIf((rsX!Existe = 0), False, True)
rsX.Close
End Function

Private Sub sbSumaDebitosCreditos()
Dim x As Long, TC As Currency
  
  'Por Ahora Así
  TC = 1
  
  txtDebito = 0
  txtCredito = 0
  For x = 1 To vGrid.MaxRows
      vGrid.Row = x
      vGrid.col = 5
      txtDebito = CCur(txtDebito) + (CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) * TC)
      vGrid.col = 6
      txtCredito = CCur(txtCredito) + (CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) * TC)
  Next x
  txtDiferencia = txtDebito - txtCredito
  txtDebito = Format(txtDebito, "Standard")
  txtCredito = Format(txtCredito, "Standard")
  txtDiferencia = Format(txtDiferencia, "Standard")

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(8) As Variant, x As Integer
'MsgBox "Columna : " & vGrid.Col
'MsgBox "Columna Activa: " & vGrid.ActiveCol
'MsgBox "Fila : " & vGrid.Row
'MsgBox "Fila Activa: " & vGrid.ActiveRow

If KeyCode = vbKeyDelete Then
  
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 7
  If vGrid.Text <> "" Then 'Existe en la Base de datos
    'Preguntar y si la respuesta es afirmativa eliminar de la Base de datos
  
  
  End If
  
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To 8
        vGrid.col = x
        vTemp(x) = vGrid.Text
     Next x
     
     vGrid.Row = lng
     For x = 1 To 8
       vGrid.col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  
  Call sbSumaDebitosCreditos
  
  
End If

If KeyCode = vbKeyF4 And vGrid.ActiveCol = 1 Then
'  gBusquedas.Columna = "cod_cuenta"
'  gBusquedas.Orden = "cod_cuenta"
'  gBusquedas.Filtro = " and acepta_movimientos = 'S' and cod_contabilidad = " & gCntX_Parametros.CodigoConta
'  gBusquedas.Consulta = "select cod_cuenta, descripcion from CntX_Cuentas"
'  frmBusquedas.Show vbModal
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
'  vGrid.Text = gBusquedas.Resultado
End If

If (KeyCode = 13 Or KeyCode = vbKeyTab) Then
    vGrid.col = vGrid.ActiveCol
    vGrid.Row = vGrid.ActiveRow
    
    Select Case vGrid.ActiveCol
      Case 1 'Cuenta
        vGrid.Text = fxCntX_CuentaFormato(True, vGrid.Text)
        i = fxCntX_CuentaFormato(False, vGrid.Text)
        If fxVerificaCuenta(CStr(i)) Then
          vGrid.col = 2
          vGrid.Text = fxCntX_Cuenta("D", CStr(i))
          'Ubicar Tipo de Cambio Aqui
'          vGrid.col = 8
'          vGrid.Text = fxCntX_TipoCambio(i)
        Else
          MsgBox "Cuenta no es válida : " & vbCrLf & " - No Existe o No Acepta Movimientos" _
                 & vbCrLf & " - VERIFIQUE O MODIFIQUE EN EL CATALAGO DE CntX_Cuentas", vbCritical
        End If
        
      Case 3
        vUltimos.Documento = vGrid.Text
      
      Case 4
        vUltimos.Detalle = vGrid.Text
        
      Case 5 'Debe
        If Val(vGrid.Text) > 0 Then
            vGrid.col = vGrid.ActiveCol + 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
            
        End If
      
      Case 6 'Haber
        If Val(vGrid.Text) > 0 Then
            vGrid.col = vGrid.ActiveCol - 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
        
        End If
      
      Case 7 'Nueva linea
        If vGrid.MaxRows = vGrid.Row Then
            vGrid.MaxRows = vGrid.MaxRows + 1
            vGrid.Row = vGrid.MaxRows
            vGrid.col = 3
            vGrid.Text = vUltimos.Documento
            vGrid.col = 4
            vGrid.Text = vUltimos.Detalle
        End If
    End Select
End If

If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
    vGrid.col = 3
    vGrid.Text = vUltimos.Documento
    vGrid.col = 4
    vGrid.Text = vUltimos.Detalle
End If


End Sub



