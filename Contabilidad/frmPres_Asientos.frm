VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPres_Asientos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asientos de Presupuesto / Traslado de Partidas"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMes 
      Height          =   315
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   21
      ToolTipText     =   "(F4) Mes del Periodo"
      Top             =   1200
      Width           =   435
   End
   Begin VB.TextBox txtAnio 
      Height          =   315
      Left            =   3450
      MaxLength       =   4
      TabIndex        =   20
      ToolTipText     =   "(F4) Año del Periodo"
      Top             =   1200
      Width           =   645
   End
   Begin VB.TextBox txtPeriodo 
      Height          =   315
      Left            =   4110
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "(F4) Descripción del Periodo"
      Top             =   1200
      Width           =   5985
   End
   Begin VB.TextBox txtMesInclusion 
      Height          =   315
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   18
      ToolTipText     =   "(F4) Mes del Periodo"
      Top             =   1560
      Width           =   435
   End
   Begin VB.TextBox txtAnioInclusion 
      Height          =   315
      Left            =   3450
      MaxLength       =   4
      TabIndex        =   17
      ToolTipText     =   "(F4) Año del Periodo"
      Top             =   1560
      Width           =   645
   End
   Begin VB.TextBox txtPeriodoInclusion 
      Height          =   315
      Left            =   4110
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "(F4) Descripción del Periodo"
      Top             =   1560
      Width           =   5985
   End
   Begin VB.TextBox txtCredito 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtDebito 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5970
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtDiferencia 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Width           =   8895
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   8760
      TabIndex        =   3
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   141950979
      CurrentDate     =   37278
   End
   Begin VB.TextBox txtNAsiento 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageListMenu 
      Left            =   9480
      Top             =   -120
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
            Picture         =   "frmPres_Asientos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPres_Asientos.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPres_Asientos.frx":0D2C
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
      Width           =   10140
      _ExtentX        =   17886
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
      Height          =   2535
      Left            =   0
      TabIndex        =   22
      Top             =   2040
      Width           =   10095
      _Version        =   524288
      _ExtentX        =   17806
      _ExtentY        =   4471
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
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmPres_Asientos.frx":113E
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   0
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periodos de Inclusión"
      Height          =   315
      Index           =   5
      Left            =   1200
      TabIndex        =   15
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periodos de Exclusión"
      Height          =   315
      Index           =   4
      Left            =   1200
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Periodos"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblAsientoEstado 
      Caption         =   "Estado del Asiento."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   0
      TabIndex        =   11
      Top             =   5100
      Width           =   2115
   End
   Begin VB.Label lsblr 
      Caption         =   "Diferencia:"
      Height          =   255
      Left            =   2550
      TabIndex        =   10
      Top             =   5070
      Width           =   795
   End
   Begin VB.Label Label5 
      Caption         =   "Totales:"
      Height          =   255
      Left            =   5070
      TabIndex        =   9
      Top             =   5070
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Num. Asiento"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmPres_Asientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type xUltimos
   NumAsiento As String
   Detalle    As String
   Fecha      As Date
   AplAsiento As Integer
End Type
Dim vEdita As Boolean, vBusca As Integer, vUltimos As xUltimos

Private Sub dtpFecha_Change()
txtAnio = Year(dtpFecha.Value)
txtMes = Month(dtpFecha.Value)

txtAnioInclusion = txtAnio
txtMesInclusion = txtMes

End Sub

Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Set Me.Icon = frmContenedor.Icon
vPaso = False

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
  
 dtpFecha = Format(fxFechaServidor, "yyyy/mm/dd")
 
 vEdita = False
 Call sbLimpiaPantalla
 
'  Call Formularios(Me)
 ' Call RefrescaTags(Me)
 
End Sub

Private Function fxVerificaAsiento() As Boolean
Dim rs As New ADODB.Recordset, strSQL As String
Dim vMensaje As String, lng As Long

'Verificar que el Perido de Exclusion no este cerrado
'Verificar que todas las cuentas sean presupuestarias

fxVerificaAsiento = True
vMensaje = ""


strSQL = "select coalesce(count(*),0) as Existe from periodos where COD_CONTABILIDAD = " _
       & vCodEmpresa & " and Anio = " & txtAnio & " and mes = " & txtMes _
       & " and estado = 'P'"
Call OpenRecordSet(rs, strSQL, 0)
If rs!Existe = 0 Then
   vMensaje = vMensaje & vbCrLf & "El Periodo de Exclusion se encuentra Cerrado..."
End If
rs.Close

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.Col = 1
 If vGrid.Text <> "" Then
   vGrid.Col = 2
   If vGrid.Text = "" Then
      vGrid.Col = 1
      vMensaje = vMensaje & vbCrLf & "- Cuenta " & vGrid.Text & " No Existe"
   End If
 End If
Next lng


If CCur(txtDiferencia) <> 0 Then vMensaje = vMensaje & vbCrLf & "- El asiento no se encuentra Balanceado, por motivos" _
            & " de Seguridad no se permite Guardar Dicha información en asientos de CNTX_CONSOLIDA_DEFINICION"

If Len(vMensaje) > 0 Then
  fxVerificaAsiento = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbLimpiaPantalla()
vBusca = 1
txtCredito = 0
txtDebito = 0
txtDescripcion = ""
txtDiferencia = 0
txtNAsiento = ""
vGrid.MaxRows = 0
vGrid.MaxRows = 1
vGrid.MaxCols = 6
End Sub


Private Sub sbReporte(strSQL As String)

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "ContaExpress"
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(2) = "Usuario='" & glogon.Usuario & "'"
 .Formulas(3) = "Mascara='" & vParametros.MascaraCod & "'"
 .Formulas(4) = "SubTitulo='Periodo : " & txtPeriodo & "'"
 .Connect = glogon.ConectRPT
 
 .ReportFileName = App.Path & "\PreAsientos.rpt"

 .SelectionFormula = strSQL
 .PrintReport
  
End With

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtNAsiento.SetFocus
      Call sbToolBar(tlb, "edicion")
    
    Case "MODIFICAR", "EDITAR"
      If vUltimos.AplAsiento = 0 Then
        vEdita = True
        txtDescripcion.SetFocus
        Call sbToolBar(tlb, "edicion")
      Else
        MsgBox "Este Asiento ya Fue Aplicado (No se puede Modificar/Borrar)", vbInformation
      End If
    
    Case "BORRAR"
      If vUltimos.AplAsiento = 0 Then
        Call sbBorrar
      Else
        MsgBox "Este Asiento ya Fue Aplicado (No se puede Modificar/Borrar)", vbInformation
      End If
      
    Case "GUARDAR", "SALVAR"
      Call sbGuardar
    
    Case "DESHACER"
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
    
    Case "CONSULTAR"
      Select Case vBusca
       Case 3
            gBusquedas.Columna = "descripcion"
            gBusquedas.Orden = "descripcion"
            gBusquedas.Consulta = "select cod_asiento,descripcion,fecha from pre_asientos"
            gBusquedas.Filtro = " and COD_CONTABILIDAD = " & vParametros.CodigoEmpresa
            frmBusquedas.Show vbModal
            txtNAsiento = gBusquedas.Resultado
            txtNAsiento.SetFocus
       Case 4
            gBusquedas.Columna = "Cod_asiento"
            gBusquedas.Orden = "Cod_asiento"
            gBusquedas.Consulta = "select cod_asiento,descripcion,fecha from pre_asientos"
            gBusquedas.Filtro = " and COD_CONTABILIDAD = " & vParametros.CodigoEmpresa
            frmBusquedas.Show vbModal
            txtNAsiento = gBusquedas.Resultado
            txtNAsiento.SetFocus
       End Select
    
    Case "REPORTES"
      
      strSQL = "{PRE_ASIENTOS.COD_CONTABILIDAD} = " & vParametros.CodigoEmpresa _
             & " AND {PRE_ASIENTOS.COD_ASIENTO} = '" & txtNAsiento & "'"
      
      Call sbReporte(strSQL)
    
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
    vGrid.Col = i
    If i = 1 Then
        vGrid.Text = fxFormatoCuenta(True, CStr(rs.Fields(i - 1).Value))
        vGrid.Col = 4
        vGrid.TextTip = TextTipFixed
        vGrid.CellNote = fxNotaPrePeriodo(CStr(rs.Fields(i - 1).Value), txtAnio, txtMes)

        vGrid.Col = 5
        vGrid.TextTip = TextTipFixed
        vGrid.CellNote = fxNotaPrePeriodo(CStr(rs.Fields(i - 1).Value), txtAnioInclusion, txtMesInclusion)
        
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

strSQL = "select * from pre_asientos where COD_CONTABILIDAD = " & vParametros.CodigoEmpresa _
       & " and cod_asiento = '" & strNumero & "'"
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
 
  vUltimos.NumAsiento = rs!cod_asiento
  vUltimos.Fecha = rs!Fecha
  vUltimos.AplAsiento = 1
  
  'llenar datos en pantalla
  
  txtAnio = rs!eAnio
  txtMes = rs!eMes
  
  txtAnioInclusion = rs!iAnio
  txtMesInclusion = rs!iMes
  
  txtDescripcion = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
  dtpFecha.Value = vUltimos.Fecha
  txtNAsiento = vUltimos.NumAsiento
  
  If vUltimos.AplAsiento = 0 Then
    lblAsientoEstado.Caption = "Este Asiento se Encuentra Pendiente"
  Else
    lblAsientoEstado.Caption = "Este Asiento se Encuentra Mayorizado"
  End If
  
strSQL = "select A.cod_cuenta,B.descripcion,detalle,debitos,creditos,linea" _
       & " from pre_asientos_detalle A inner join cuentas B on A.cod_cuenta = B.cod_cuenta" _
       & " And A.COD_CONTABILIDAD = B.COD_CONTABILIDAD" _
       & " where A.COD_CONTABILIDAD = " & vParametros.CodigoEmpresa & " and cod_asiento = '" _
       & vUltimos.NumAsiento & "' order by linea"
   
  Call sbCargaGridLocal(vGrid, 6, strSQL)
 
  Call sbSumaDebitosCreditos

  If vUltimos.AplAsiento = 1 Then
    vGrid.Lock = True
  Else
    vGrid.Lock = False
  End If

End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub

Private Sub sbGuardar()
Dim strSQL As String, lng As Long

On Error GoTo vError

If fxVerificaAsiento Then
      
    If vEdita Then
      
      strSQL = "update pre_asientos set descripcion = '" & UCase(txtDescripcion) _
             & "',fecha = '" & Format(dtpFecha.Value, "yyyy/mm/dd") _
             & "',IAnio = " & txtAnioInclusion & ",IMes = " & txtMesInclusion _
             & ",EAnio = " & txtAnio & ",EMes = " & txtMes _
             & " where COD_CONTABILIDAD = " & vParametros.CodigoEmpresa _
             & " and cod_asiento = '" & txtNAsiento & "'"
      Call ConectionExecute(strSQL, 0)
     
      strSQL = "delete pre_asientos_detalle where COD_CONTABILIDAD = " _
             & vParametros.CodigoEmpresa & " and cod_asiento = '" & txtNAsiento & "'"
      Call ConectionExecute(strSQL, 0)
    
      For lng = 1 To vGrid.MaxRows
        vGrid.Row = lng
        vGrid.Col = 1
        If vGrid.Text <> "" Then
            strSQL = "insert into pre_asientos_detalle(cod_asiento,COD_CONTABILIDAD" _
                   & ",linea,cod_cuenta,detalle,debitos,creditos" _
                   & ") values('" & txtNAsiento & "'," & vParametros.CodigoEmpresa _
                   & "," & lng & ",'"
            vGrid.Row = lng
            vGrid.Col = 1
            strSQL = strSQL & fxFormatoCuenta(False, vGrid.Text) & "','"
            vGrid.Col = 3
            strSQL = strSQL & vGrid.Text & "',"
            vGrid.Col = 4
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ","
            vGrid.Col = 5
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ")" _
          
            Call ConectionExecute(strSQL, 0)
              
         End If 'vgrid.Text <> ""
       
       Next lng
    
      Call Bitacora("Modifica", "Asiento Presuesto: " & txtNAsiento & " Empresa :" & vParametros.CodigoEmpresa)
    
    
    Else 'Inserta
       strSQL = "insert into pre_asientos(COD_CONTABILIDAD,cod_asiento" _
              & ",fecha,descripcion,fecha_sistema,user_crea,IAnio,IMes,EAnio,EMes)" _
              & " values(" & vParametros.CodigoEmpresa & ",'" & txtNAsiento & "','" _
              & Format(dtpFecha.Value, "yyyy/mm/dd") & "','" & UCase(txtDescripcion) _
              & "',getdate(),'" & glogon.Usuario & "'," & txtAnioInclusion _
              & "," & txtMesInclusion & "," & txtAnio & "," & txtMes & ")"
       Call ConectionExecute(strSQL, 0)
       
      For lng = 1 To vGrid.MaxRows
        vGrid.Row = lng
        vGrid.Col = 1
        If vGrid.Text <> "" Then
            strSQL = "insert into pre_asientos_detalle(cod_asiento,COD_CONTABILIDAD" _
                   & ",linea,cod_cuenta,detalle,debitos,creditos" _
                   & ") values('" & txtNAsiento & "'," & vParametros.CodigoEmpresa _
                   & "," & lng & ",'"
            vGrid.Row = lng
            vGrid.Col = 1
            strSQL = strSQL & fxFormatoCuenta(False, vGrid.Text) & "','"
            vGrid.Col = 3
            strSQL = strSQL & vGrid.Text & "',"
            vGrid.Col = 4
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ","
            vGrid.Col = 5
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ")" _
          
            Call ConectionExecute(strSQL, 0)
              
         End If 'vgrid.Text <> ""
       
       Next lng
       
       
       Call Bitacora("Registra", "Asiento Presupuesto: " & txtNAsiento & " Empresa : " & vParametros.CodigoEmpresa)
        
       'Mayorizar el Asiento de Presupuesto
        Call sbMayorizar(txtNAsiento, dtpFecha.Value)
        
    End If 'Si Inserta o Actualiza

        Call sbToolBar(tlb, "activo")
        Call sbConsultaAsiento(txtNAsiento)
        
        vEdita = True
        
        MsgBox "Información guardada satisfactoriamente...", vbInformation


End If 'Verificacion del Asiento

' Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox Err.Description, vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String
On Error GoTo vError
i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
  strSQL = "delete pre_asientos_detalle where COD_CONTABILIDAD = " _
         & vParametros.CodigoEmpresa & " and cod_asiento = '" & txtNAsiento & "'"
  Call ConectionExecute(strSQL, 0)
  
  strSQL = "delete pre_asientos where COD_CONTABILIDAD = " _
         & vParametros.CodigoEmpresa & " and cod_asiento = '" & txtNAsiento & "'"
  Call ConectionExecute(strSQL, 0)
  

  Call Bitacora("Elimina", "Asiento Presupuesto: " & txtNAsiento & " Empresa :" _
                  & vParametros.CodigoEmpresa)

  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub sbRefrescaInformacion(vAnio As Long, vMes As Integer, Obx As Object)
Dim strResultado As String

On Error GoTo vError
  
  Select Case vMes
    Case 1
        strResultado = "ENERO DEL " & vAnio
    Case 2
        strResultado = "FEBRERO DEL " & vAnio
    Case 3
        strResultado = "MARZO DEL " & vAnio
    Case 4
        strResultado = "ABRIL DEL " & vAnio
    Case 5
        strResultado = "MAYO DEL " & vAnio
    Case 6
        strResultado = "JUNIO DEL " & vAnio
    Case 7
        strResultado = "JULIO DEL " & vAnio
    Case 8
        strResultado = "AGOSTO DEL " & vAnio
    Case 9
        strResultado = "SETIEMBRE DEL " & vAnio
    Case 10
        strResultado = "OCTUBRE DEL " & vAnio
    Case 11
        strResultado = "NOVIEMBRE DEL " & vAnio
    Case 12
        strResultado = "DICIEMBRE DEL " & vAnio
  End Select

  Obx.Text = strResultado

Exit Sub

vError:
End Sub



Private Sub txtAnio_Change()
On Error Resume Next
Call sbRefrescaInformacion(txtAnio, txtMes, txtPeriodo)
End Sub

Private Sub txtAnioInclusion_Change()
On Error Resume Next
Call sbRefrescaInformacion(txtAnioInclusion, txtMesInclusion, txtPeriodoInclusion)
End Sub

Private Sub txtDescripcion_GotFocus()
vBusca = 4
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
End Sub

Private Sub txtMes_Change()
On Error Resume Next
Call sbRefrescaInformacion(txtAnio, txtMes, txtPeriodo)
End Sub

Private Sub txtMesInclusion_Change()
On Error Resume Next
Call sbRefrescaInformacion(txtAnioInclusion, txtMesInclusion, txtPeriodoInclusion)
End Sub

Private Sub txtNAsiento_GotFocus()
vBusca = 3
End Sub

Private Sub txtNAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 dtpFecha.SetFocus
 Call sbConsultaAsiento(txtNAsiento)
End If
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
End Sub


Private Function fxVerificaCuenta(strCuenta As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select coalesce(count(*),0) as Existe from cuentas where COD_CONTABILIDAD = " & vParametros.CodigoEmpresa _
       & " and cod_cuenta = '" & strCuenta & "' and presupuesto = 'S'"
Call OpenRecordSet(rsX, strSQL, 0)
fxVerificaCuenta = IIf((rsX!Existe = 0), False, True)
rsX.Close
End Function

Private Sub sbSumaDebitosCreditos()
Dim x As Long
  txtDebito = 0
  txtCredito = 0
  For x = 1 To vGrid.MaxRows
      vGrid.Row = x
      vGrid.Col = 4
      txtDebito = CCur(txtDebito) + CCur(IIf(vGrid.Text = "", 0, vGrid.Text))
      vGrid.Col = 5
      txtCredito = CCur(txtCredito) + CCur(IIf(vGrid.Text = "", 0, vGrid.Text))
  Next x
  txtDiferencia = txtDebito - txtCredito
  txtDebito = Format(txtDebito, "Standard")
  txtCredito = Format(txtCredito, "Standard")
  txtDiferencia = Format(txtDiferencia, "Standard")

End Sub


Private Function fxNotaPrePeriodo(vCuenta As String, vAnio As Integer, vMes As Integer) As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCadena As String

vCadena = ""

strSQL = "select P.*,C.descripcion,(coalesce(M.saldo_inicial,0) + coalesce(M.total_Debitos,0)" _
       & " + coalesce(M.total_creditos,0)) as Real" _
       & " from presupuesto P inner join Cuentas C on P.cod_cuenta = C.cod_cuenta and P.COD_CONTABILIDAD = C.COD_CONTABILIDAD" _
       & " left join Movimiento_Cuentas M on P.cod_cuenta = M.cod_cuenta" _
       & " and P.anio = M.anio and P.mes = M.mes and P.COD_CONTABILIDAD = M.COD_CONTABILIDAD" _
       & " where P.COD_CONTABILIDAD = " & vCodEmpresa & " and P.cod_cuenta = '" & vCuenta _
       & "' and P.anio = " & vAnio & " and P.Mes = " & vMes
Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
 vCadena = "PERIODO : " & vAnio & "-" & Format(vMes, "00") & vbCrLf _
         & "NO SE HA DEFINIDO PARTIDA"
Else
 vCadena = "PERIODO : " & vAnio & "-" & Format(vMes, "00") & vbCrLf _
         & "ORIGINAL: " & Format(rs!presu_original, "Standard") & vbCrLf _
         & "(+) Ajustes: " & Format(rs!ajuste_positivo, "Standard") & vbCrLf _
         & "(-) Ajustes: " & Format(rs!ajuste_negativo, "Standard") & vbCrLf _
         & "ACTUAL: " & Format(rs!presu_actual, "Standard") & vbCrLf _
         & "REAL: " & Format(rs!Real, "Standard") & vbCrLf _
         & "DIF: " & Format(rs!presu_actual - rs!Real, "Standard")

End If
rs.Close

fxNotaPrePeriodo = vCadena

End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(6) As Variant, x As Integer
Dim strSQL As String, rs As New ADODB.Recordset

'MsgBox "Columna : " & vGrid.Col
'MsgBox "Columna Activa: " & vGrid.ActiveCol
'MsgBox "Fila : " & vGrid.Row
'MsgBox "Fila Activa: " & vGrid.ActiveRow

If KeyCode = vbKeyDelete Then
  
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 6
  If vGrid.Text <> "" Then 'Existe en la Base de datos
    'Preguntar y si la respuesta es afirmativa eliminar de la Base de datos
  End If
  
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To 6
        vGrid.Col = x
        vTemp(x) = vGrid.Text
     Next x
     
     vGrid.Row = lng
     For x = 1 To 6
       vGrid.Col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  
  Call sbSumaDebitosCreditos
  
End If


If KeyCode = vbKeyF4 And vGrid.ActiveCol = 1 Then
  gBusquedas.Columna = "cod_cuenta"
  gBusquedas.Orden = "cod_cuenta"
  gBusquedas.Filtro = " and presupuesto = 'S' and COD_CONTABILIDAD = " & vParametros.CodigoEmpresa
  gBusquedas.Consulta = "select cod_cuenta, descripcion from cuentas"
  
  frmBusquedas.Show vbModal
  
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow

  vGrid.Text = gBusquedas.Resultado
End If

If (KeyCode = 13 Or KeyCode = vbKeyTab) Then
    vGrid.Col = vGrid.ActiveCol
    vGrid.Row = vGrid.ActiveRow
    
    Select Case vGrid.ActiveCol
      Case 1 'Cuenta
        vGrid.Text = fxFormatoCuenta(True, vGrid.Text)
        i = fxFormatoCuenta(False, vGrid.Text)
        If fxVerificaCuenta(CStr(i)) Then
          vGrid.Col = 2
          vGrid.Text = fxCuenta("D", CStr(i))
        Else
          MsgBox "Cuenta no es válida : " & vbCrLf & " - No Existe o No esta Asignada al Presupuesto" _
                 & vbCrLf & " - VERIFIQUE O MODIFIQUE EN EL CATALAGO DE CUENTAS", vbCritical
        End If
        
        vGrid.Col = 4
        vGrid.TextTip = TextTipFixed
        vGrid.CellNote = fxNotaPrePeriodo(CStr(i), txtAnio, txtMes)

        vGrid.Col = 5
        vGrid.TextTip = TextTipFixed
        vGrid.CellNote = fxNotaPrePeriodo(CStr(i), txtAnioInclusion, txtMesInclusion)
        
      Case 3
        vUltimos.Detalle = vGrid.Text
        
      Case 4 'Debe
        If Val(vGrid.Text) > 0 Then
            vGrid.Col = vGrid.ActiveCol + 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
            
        End If
      
      Case 5 'Haber
        If Val(vGrid.Text) > 0 Then
            vGrid.Col = vGrid.ActiveCol - 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
        
        End If
      
      Case 6 'Nueva linea
        If vGrid.MaxRows = vGrid.Row Then
            vGrid.MaxRows = vGrid.MaxRows + 1
            vGrid.Row = vGrid.MaxRows
            vGrid.Col = 3
            vGrid.Text = vUltimos.Detalle
        End If
    End Select
End If


If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
    vGrid.Col = 3
    vGrid.Text = vUltimos.Detalle
End If


End Sub



Private Function fxExisteMovimiento(lngAnio As Long, iMes As Integer, strCuenta As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select coalesce(count(*),0) as existe from presupuesto where COD_CONTABILIDAD = " _
       & vParametros.CodigoEmpresa & " and anio = " & lngAnio & " and mes = " & iMes _
       & " and cod_cuenta = '" & strCuenta & "'"

Call OpenRecordSet(rsX, strSQL, 0)
fxExisteMovimiento = IIf((rsX!Existe = 0), False, True)
rsX.Close

End Function



Private Sub sbGuardaMovimiento(lngAnio As Long, iMes As Integer, vCuentaActual As String _
                                , curDebe As Currency, curHaber As Currency)
Dim strSQL As String

If fxExisteMovimiento(lngAnio, iMes, vCuentaActual) Then
 
 strSQL = "update presupuesto set ajuste_negativo = ajuste_negativo + " & curDebe _
        & ", Ajuste_positivo = Ajuste_Positivo + " & curHaber _
        & ", presu_actual = presu_actual + " & curHaber & " - " & curDebe _
        & " where COD_CONTABILIDAD = " & vParametros.CodigoEmpresa & " and anio = " _
        & lngAnio & " and mes = " & iMes & " and cod_cuenta = '" _
        & vCuentaActual & "'"
 Call ConectionExecute(strSQL, 0)
 
Else
  strSQL = "insert into presupuesto(anio,mes,COD_CONTABILIDAD,cod_cuenta" _
         & ",presu_original,ajuste_negativo,ajuste_positivo,presu_actual)" _
         & " values(" & lngAnio & "," & iMes & "," & vParametros.CodigoEmpresa _
         & ",'" & vCuentaActual & "',0," & curDebe & "," & curHaber & "," _
         & IIf((curDebe > 0), (curDebe * -1), curHaber) & ")"
 Call ConectionExecute(strSQL, 0)

End If
End Sub

Private Sub sbMayorizar(vNumeroAsiento As String, vFecha As Date)
Dim strSQL As String, rs As New ADODB.Recordset

Screen.MousePointer = vbHourglass

strSQL = "select A.*,D.cod_cuenta,D.debitos,D.creditos" _
       & " from pre_Asientos A inner join pre_asientos_detalle D on A.COD_CONTABILIDAD = D.COD_CONTABILIDAD" _
       & " and A.cod_asiento = D.cod_asiento" _
       & " where A.COD_CONTABILIDAD = " & vParametros.CodigoEmpresa _
       & " and A.cod_asiento = '" & vNumeroAsiento & "'"
Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
  
  'Guarda Inclusion
  If rs!creditos > 0 Then
      Call sbGuardaMovimiento(rs!iAnio, rs!iMes, rs!COD_Cuenta, 0, rs!creditos)
  End If
  
  'Guarda Exclusion
  If rs!debitos > 0 Then
      Call sbGuardaMovimiento(rs!eAnio, rs!eMes, rs!COD_Cuenta, rs!debitos, 0)
  End If
  
  
  rs.MoveNext
Loop

rs.Close

Screen.MousePointer = vbDefault

End Sub

