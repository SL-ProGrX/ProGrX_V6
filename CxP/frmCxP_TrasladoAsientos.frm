VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCxPTrasladoAsientos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CxP: Traslado de Asientos a Contabilidad"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   9285
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   135
      Left            =   120
      TabIndex        =   16
      Top             =   8160
      Width           =   9015
      _Version        =   1441793
      _ExtentX        =   15901
      _ExtentY        =   238
      _StockProps     =   93
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkBalanceados 
      Height          =   372
      Left            =   2760
      TabIndex        =   1
      Top             =   1320
      Width           =   2652
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Solo Asientos Balanceados"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   612
      Left            =   6120
      TabIndex        =   2
      Top             =   1320
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Buscar"
      BackColor       =   -2147483633
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
      Picture         =   "frmCxP_TrasladoAsientos.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnReActivar 
      Height          =   612
      Left            =   7560
      TabIndex        =   3
      Top             =   1320
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Re Activar"
      BackColor       =   -2147483633
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
      Picture         =   "frmCxP_TrasladoAsientos.frx":0A1E
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   960
      TabIndex        =   4
      Top             =   1320
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.CheckBox chkReActivar 
      Height          =   372
      Left            =   2760
      TabIndex        =   6
      Top             =   1680
      Width           =   2652
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Re Activar Automáticamente"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6012
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   9012
      _Version        =   1441793
      _ExtentX        =   15896
      _ExtentY        =   10604
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Pendientes"
      Item(0).ControlCount=   5
      Item(0).Control(0)=   "cmdTrasladar"
      Item(0).Control(1)=   "Label1(2)"
      Item(0).Control(2)=   "lblEstatus"
      Item(0).Control(3)=   "chkDocumentos"
      Item(0).Control(4)=   "lswDocumentos"
      Item(1).Caption =   "Desbalanceados"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5412
         Left            =   -69760
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   8652
         _Version        =   524288
         _ExtentX        =   15261
         _ExtentY        =   9546
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   7
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmCxP_TrasladoAsientos.frx":1392
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton cmdTrasladar 
         Height          =   612
         Left            =   7320
         TabIndex        =   11
         Top             =   5160
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Trasladar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmCxP_TrasladoAsientos.frx":1A98
      End
      Begin XtremeSuiteControls.ListView lswDocumentos 
         Height          =   4212
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   8772
         _Version        =   1441793
         _ExtentX        =   15473
         _ExtentY        =   7429
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
         Checkboxes      =   -1  'True
         MultiSelect     =   -1  'True
         HideSelection   =   0   'False
         View            =   3
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkDocumentos 
         Height          =   372
         Left            =   7800
         TabIndex        =   13
         Top             =   480
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Todos"
         BackColor       =   -2147483633
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
         Value           =   1
         Alignment       =   1
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos Pendientes de traslado...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   3372
      End
      Begin VB.Label lblEstatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   792
         Left            =   240
         TabIndex        =   14
         Top             =   5160
         Width           =   6972
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Traslado de Asientos a Contabilidad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   7572
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmCxPTrasladoAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mAsientosDiarios As Boolean

Public Function fxValidaPeriodoAsiento(vFecha As Date) As Boolean
Dim strSQL As String, rsX As New ADODB.Recordset

strSQL = "select * from CntX_periodos where anio = " & Year(vFecha) & " and mes = " & Month(vFecha) _
        & " and estado = 'P' and cod_contabilidad = " & GLOBALES.gEnlace

Call OpenRecordSet(rsX, strSQL, 0)

If rsX.EOF And rsX.BOF Then
 fxValidaPeriodoAsiento = False
Else
 fxValidaPeriodoAsiento = True
End If
rsX.Close

End Function

Public Function fxUltimaLineaAsiento(pTipoAsiento As String, pNumAsiento As String, vFecha As Date) As Integer
Dim strSQL As String, rsX As New ADODB.Recordset

strSQL = "Select isnull(max(num_linea),0) as Linea from CntX_asientos_detalle" _
       & " where num_asiento = '" & pNumAsiento & "' and Tipo_asiento = '" & pTipoAsiento & "'" _
       & " and cod_contabilidad = " & GLOBALES.gEnlace

Call OpenRecordSet(rsX, strSQL, 0)
    fxUltimaLineaAsiento = IIf(IsNull(rsX!Linea), 0, rsX!Linea)
rsX.Close

End Function

Public Function fxVerificaExistenciaAsiento(pTipoAsiento As String, pNumAsiento As String, vFecha As Date) As Boolean
Dim strSQL As String, rsX As New ADODB.Recordset

strSQL = "Select num_asiento from CntX_Asientos where anio = " & Year(vFecha) & " and mes = " & Month(vFecha) _
        & " and tipo_asiento = '" & pTipoAsiento & "' and num_asiento = '" & pNumAsiento _
        & "' and cod_contabilidad = " & GLOBALES.gEnlace
Call OpenRecordSet(rsX, strSQL, 0)

If rsX.EOF And rsX.BOF Then
 fxVerificaExistenciaAsiento = False
Else
 fxVerificaExistenciaAsiento = True
End If
rsX.Close

End Function


Private Sub sbReActivar(Optional pAutomatico As Integer = 0)
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass


'Revisa Asientos Trasladados que no estan en la contabilidad
strSQL = "exec  spSys_Asiento_Revisa_Traslado '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00','" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59','CxP'"
Call ConectionExecute(strSQL)

If glogon.error Then Exit Sub

Me.MousePointer = vbDefault

If pAutomatico = 0 Then
    MsgBox "Revisión de Documentos realizada satisfactoriamente!", vbInformation
    Call sbBuscar
End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnBuscar_Click()

If chkReActivar.Value = vbChecked Then
    Call sbReActivar(1)
End If

Call sbBuscar
End Sub

Private Sub btnReActivar_Click()
Call sbReActivar(0)
End Sub

Private Sub chkDocumentos_Click()
Dim i As Integer

For i = 1 To lswDocumentos.ListItems.Count
  lswDocumentos.ListItems.Item(i).Checked = chkDocumentos.Value
Next i

End Sub

Private Sub cmdTrasladar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vTipoAsiento As String, vTipoDiario As Boolean, vTipoDoc As String
Dim vMascara As String

vTipoDiario = IIf((fxCxPParametro("04") = "S"), True, False)
vMascara = ""


'Limpia Casos en Cero

strSQL = "delete CXP_PAGOPROVCARGOS where MONTO = 0"
Call ConectionExecute(strSQL)

With lswDocumentos.ListItems
  For i = 1 To .Count
   If .Item(i).Checked Then
     
      Select Case .Item(i).Key
        Case "0x1" 'Facturas Registradas
           vTipoAsiento = fxCxPParametro("05")
           vTipoDoc = "FT"
        Case "0x2" 'Facturas Anuladas
           vTipoAsiento = fxCxPParametro("06")
           vTipoDoc = "FA"
        Case "0x3" 'Cargos Flotante Monto
           vTipoAsiento = fxCxPParametro("07")
           vTipoDoc = "CM"
        Case "0x4" 'Cargos Flotante Porcentual
           vTipoAsiento = fxCxPParametro("07")
           vTipoDoc = "CP"
        Case "0x5" 'Cargos Directos
           vTipoAsiento = fxCxPParametro("07")
           vTipoDoc = "CD"
        Case "0x6" 'Cargos de Anticipos de Facturas Canceladas vía Cargos/Retencion
           vTipoAsiento = fxCxPParametro("07")
           vTipoDoc = "CA"
      End Select
           
        Call sbAsientoIndividual(vTipoDoc, vTipoAsiento, vMascara)
           
'    If vTipoDiario Then
'        Call sbAsientoTipoDiario(vTipoDoc, vTipoAsiento)
'     Else
'        Call sbAsientoIndividual(vTipoDoc, vTipoAsiento, vMascara)
'    End If
      
      
   End If
  Next i
End With

Call Bitacora("Aplica", "Traslada Asientos: " & Format(dtpInicio.Value, "dd/mm/yyyy") & " - " & Format(dtpCorte.Value, "dd/mm/yyyy"))
Call sbBuscar


MsgBox "Se realizó el Traslado de Asientos a Contabilidad...!", vbInformation
Me.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
 vModulo = 30
End Sub

Private Sub Form_Load()

vModulo = 30


Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

mAsientosDiarios = IIf((fxCxPParametro("04") = "S"), True, False)

With lswDocumentos.ColumnHeaders
    .Clear
    .Add , , "Tipo Transacción", 3200
    .Add , , "Pendientes", 1400, vbCenter
    .Add , , "Bloqueados", 1400, vbCenter
End With


End Sub

Private Sub sbAsientoTipoDiario(pTipoDoc As String, pTipoAsiento As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vNumAsiento As String, intLinea As Long, vConcepto As String


Dim DH As String
Dim rsTmp As New ADODB.Recordset, vFecha As Date


lblEstatus.Caption = "Cargando Información..."
lblEstatus.Refresh
PrgBar.Value = 1

On Error GoTo vError


'Inicia Transaccion
'glogon.Conection.BeginTrans
'
''Sacar los Documentos de Inicio y Corte
'strSQL = "select year(Registro_Fecha) as Anio,month(Registro_Fecha) as Mes,day(Registro_Fecha) as Dia,isnull(min(COD_TRANSACCION),0) as Inicio, isnull(max(COD_TRANSACCION),0) as Corte" _
'       & " from SIF_TRANSACCIONES where traspaso = 'P'" _
'       & " and TIPO_DOCUMENTO in('" & pTipoDoc & "') and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
'       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
'       & " group by year(Registro_Fecha),month(Registro_Fecha),day(Registro_Fecha)"
'
'
'rs.CursorLocation = adUseServer
'Call OpenRecordSet(rs, strSQL)
'
'PrgBar.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
'lblEstatus.Caption = "Procesando Asientos..." & pTipoDoc
'lblEstatus.Refresh
'
'Do While Not rs.EOF
' vFecha = rs!Anio & "/" & rs!Mes & "/" & rs!Dia
'
' If fxValidaPeriodoAsiento(vFecha) Then  'Verificar el Periodo Abierto en contabilidad
'
'    lblEstatus.Caption = "Procesando Asientos..." & pTipoDoc & "[" & rs!inicio & "-" & rs!corte & "]"
'    lblEstatus.Refresh
'
'    vNumAsiento = pTipoDoc & ".a." & rs!Anio & ".m." & rs!Mes & ".d." & rs!Dia
'
'
'    'Crea el Maestro de Asiento
'    strSQL = "insert CntX_asientos(cod_contabilidad,Tipo_Asiento,num_asiento,anio,mes,fecha_asiento,descripcion,balanceado,modulo)" _
'           & " values(" & GLOBALES.gEnlace & ",'" & pTipoAsiento & "','" & vNumAsiento & "'," & rs!Anio & "," & rs!Mes _
'           & ",'" & Format(vFecha, "yyyy/mm/dd") & "','ASIENTO DE DIARIO','S'," & vModulo & ")"
'    Call ConectionExecute(strSQL)
'
'    intLinea = 1
'
'    'Detalle del Asiento de diario
'
'    strSQL = "select Tra.Referencia_01,Tra.Referencia_02,Tra.Documento,Asi.*,Con.Descripcion as 'ConceptoDesc'" _
'           & " from SIF_TRANSACCIONES Tra inner join SIF_TRANSACCIONES_ASIENTO Asi on Tra.COD_TRANSACCION =  Asi.COD_TRANSACCION and Tra.TIPO_DOCUMENTO = Asi.TIPO_DOCUMENTO" _
'           & "  inner join SIF_Conceptos Con on Tra.Cod_Concepto = Con.cod_Concepto" _
'           & " where Tra.TIPO_DOCUMENTO = '" & pTipoDoc & "' and Tra.COD_TRANSACCION between '" & rs!inicio & "' and '" & rs!corte & "'"
'    Call OpenRecordSet(rsTmp, strSQL, 0)
'    Do While Not rsTmp.EOF
'
'        vConcepto = Trim(rsTmp!ConceptoDesc) & ".Ref." & Trim(rsTmp!Referencia_01) & "." & Trim(rsTmp!Referencia_02)
'
'        If UCase(rsTmp!Tipo_Movimiento) = "C" Then
'          strSQL = "insert CntX_Asientos_detalle(cod_contabilidad,Tipo_Asiento,num_asiento,num_linea,cod_cuenta,monto_debito" _
'                 & ",monto_credito,detalle,documento,cod_unidad,cod_centro_Costo,cod_divisa,Tipo_cambio)" _
'                 & " values(" & rsTmp!cod_Contabilidad & ",'" & pTipoAsiento & "','" & vNumAsiento & "'," & intLinea & ",'" & Trim(rsTmp!cod_cuenta) _
'                 & "',0," & rsTmp!Monto & ",'" & vConcepto _
'                 & "','" & rsTmp!Tipo_Documento & "." & Format(rsTmp!Cod_Transaccion, "0000000000") & "','" & rsTmp!cod_unidad & "','" & rsTmp!cod_centro_Costo & "','" & rsTmp!cod_divisa & "',1)"
'
'        Else
'          strSQL = "insert CntX_Asientos_detalle(cod_contabilidad,Tipo_Asiento,num_asiento,num_linea,cod_cuenta,monto_debito" _
'                 & ",monto_credito,detalle,documento,cod_unidad,cod_centro_Costo,cod_divisa,Tipo_cambio)" _
'                 & " values(" & rsTmp!cod_Contabilidad & ",'" & pTipoAsiento & "','" & vNumAsiento & "'," & intLinea & ",'" & Trim(rsTmp!cod_cuenta) _
'                 & "'," & rsTmp!Monto & ",0,'" & vConcepto & "','" & rsTmp!Tipo_Documento & "." & Format(rsTmp!Cod_Transaccion, "0000000000") _
'                 & "','" & rsTmp!cod_unidad & "','" & rsTmp!cod_centro_Costo & "','" & rsTmp!cod_divisa & "',1)"
'        End If
'
'        Call ConectionExecute(strSQL)
'        intLinea = intLinea + 1
'
'      rsTmp.MoveNext
'    Loop
'    rsTmp.Close
'
'
'    'Actualiza Tabla de SIF_TRANSACCIONES
'    strSQL = "Update SIF_TRANSACCIONES set Traspaso = 'G', TRASPASO_FECHA = dbo.MyGetdate()" _
'            & ",TRASPASO_USUARIO = '" & glogon.Usuario & "' where COD_TRANSACCION between '" & rs!inicio _
'            & "' and '" & rs!corte & "' and TIPO_DOCUMENTO = '" & pTipoDoc & "'"
'
'    Call ConectionExecute(strSQL)
'
'
' Else 'Verificacion del periodo
'
'   MsgBox "Existen asientos que no pueden ser trasladados porque el periodo fué cerrado...", vbInformation
'
'
' End If 'Verificacion del periodo
'
' If PrgBar.Value < PrgBar.Max Then PrgBar.Value = PrgBar.Value + 1
' rs.MoveNext
'Loop
'rs.Close
'
''Cierra Transaccion
'glogon.Conection.CommitTrans


lblEstatus.Caption = ""
lblEstatus.Refresh
PrgBar.Value = 1


Exit Sub

vError:
    lblEstatus.Caption = ""
    lblEstatus.Refresh
    PrgBar.Value = 1
    Me.MousePointer = vbDefault
    glogon.Conection.RollbackTrans
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub


Private Sub sbAsientoIndividual(pTipoDoc As String, pTipoAsiento As String, Optional pMascara As String = "")
Dim strSQL As String, rs As New ADODB.Recordset
Dim vNumAsiento As String, vUnidad As String

On Error GoTo vError

Me.MousePointer = vbHourglass

lblEstatus.Caption = "Cargando Información..."
lblEstatus.Refresh
PrgBar.Value = 1

pTipoAsiento = Trim(pTipoAsiento)
pMascara = Trim(pMascara)
vUnidad = Trim(fxCxPParametro("01"))

Select Case pTipoDoc
  Case "FT" 'Factura Registrada
        strSQL = "select Ft.*,Ft.cod_factura as 'Cod_Transaccion', Ft.Fecha as 'Registro_Fecha'" _
               & ",isnull(Ft.Creacion_User,'') + char(13) + char(10) + Prov.Descripcion + ' -> ' + Ft.Notas as 'AsientoNotas'" _
               & ",convert(varchar(10),Ft.COD_PROVEEDOR) + '..' + Prov.Descripcion as 'Referencia'" _
               & ",'Factura No.:' + Ft.cod_factura + '.. Prov:' + convert(varchar(10),Ft.COD_PROVEEDOR)  as 'AsientoDesc'" _
               & " from CxP_Facturas Ft inner join CxP_Proveedores Prov on Ft.cod_Proveedor = Prov.cod_Proveedor" _
               & " where Ft.Asiento_Generado = 'P' and Ft.Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
               
        If chkBalanceados.Value = vbChecked Then
            strSQL = strSQL & " and dbo.fxCxP_AsientoBalanceado('factura',Ft.COD_PROVEEDOR, Ft.COD_FACTURA) = 1"
        End If
               
       strSQL = strSQL & " order by Creacion_Fecha"
  
  Case "FA" 'Factura Anulada
        strSQL = "select Ft.*,Ft.cod_factura as 'Cod_Transaccion', Ft.anula_fecha as 'Registro_Fecha'" _
               & ",isnull(Ft.Creacion_User,'') + char(13) + char(10) + Prov.Descripcion + ' -> ' + Ft.Notas as 'AsientoNotas'" _
               & ",convert(varchar(10),Ft.COD_PROVEEDOR) + '..' + Prov.Descripcion as 'Referencia'" _
               & ",'Factura Anulada No.:' + Ft.cod_factura + '.. Prov:' + convert(varchar(10),Ft.COD_PROVEEDOR)  as 'AsientoDesc'" _
               & " from CxP_Facturas Ft inner join CxP_Proveedores Prov on Ft.cod_Proveedor = Prov.cod_Proveedor" _
               & " where Ft.Estado = 'A' and Ft.anula_asiento_fecha is null and Ft.anula_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  
        If chkBalanceados.Value = vbChecked Then
            strSQL = strSQL & " and dbo.fxCxP_AsientoBalanceado('factura',Ft.COD_PROVEEDOR, Ft.COD_FACTURA) = 1"
        End If
               
       strSQL = strSQL & " order by anula_fecha"
  
  Case "CM" 'Cargos Flotante base Monto
          strSQL = "select C.*, convert(varchar(10),[ID]) as 'Cod_Transaccion',C.concepto as 'Descripcion',C.detalle as 'AsientoNotas'" _
               & ", T.descripcion as 'Cargo', T.cod_cuenta as 'CtaCargo',P.cod_cuenta as 'CtaProveedor'" _
               & ", T.descripcion + '.. ID: ' + convert(varchar(10),C.[ID]) + '.. Prov: ' + convert(varchar(10),C.COD_PROVEEDOR)  as 'AsientoDesc'" _
               & ", convert(varchar(10),P.COD_PROVEEDOR) + '..' + P.Descripcion as 'Referencia'" _
               & " from cxP_cargosPer C inner join cxp_cargos T on C.cod_Cargo = T.cod_Cargo  " _
               & " inner join cxp_proveedores P on C.cod_Proveedor = P.cod_Proveedor" _
               & " where C.Tipo = 'M' and C.concepto not in('*** PAGO ANTICIPADO ***')" _
               & " and C.REGISTRO_FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") _
               & " 23:59:59' and C.Asiento_Fecha is null  order by C.REGISTRO_FECHA"
               
  Case "CP" 'Cargos Flotante Base Porcentual
          strSQL = "select Car.*, convert(varchar(10),Car.[ID]) + '.' + rtrim(Car.cod_Factura) as 'Cod_Transaccion'" _
               & ",Per.concepto as 'Descripcion'" _
               & ", convert(varchar(10),P.COD_PROVEEDOR) + '..' + P.Descripcion as 'Referencia'" _
               & ",'Cargo de Anticipo/Fact.Cancelada vía Ret. Prov:' + P.descripcion + '  Fact.:' + Car.cod_Factura + ' No.Pago: ' + convert(varchar(30), Pg.NPago)  as 'AsientoNotas'" _
               & ",T.descripcion + '.. ID: ' + convert(varchar(10),Car.[ID]) + '.. Prov: ' + convert(varchar(10),Car.COD_PROVEEDOR)   as 'AsientoDesc'" _
               & ",T.descripcion as 'Cargo', T.cod_cuenta as 'CtaCargo',P.cod_cuenta as 'CtaProveedor'" _
               & " from cxp_PagoProvCargos Car inner join CXP_CARGOSPER Per on Car.COD_PROVEEDOR = Per.COD_PROVEEDOR and Car.ID = Per.ID" _
               & " inner join cxp_Cargos T on Car.cod_Cargo = T.cod_Cargo" _
               & " inner join cxp_proveedores P on Car.cod_Proveedor = P.cod_Proveedor" _
               & " where Per.TIPO = 'P' and Car.TIPO_PROCESO = 'F' and Car.Asiento_Fecha is null" _
               & " and Car.REGISTRO_FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'  order by Car.REGISTRO_FECHA"
                          
  Case "CD" 'Cargos Directos de la Factura
          strSQL = "select Car.*, convert(varchar(10),Car.[IDX_Consec]) + '.' + rtrim(Car.cod_Factura) as 'Cod_Transaccion'" _
               & ",'Cargo de Anticipo/Fact.Cancelada vía Ret. Prov:' + P.descripcion + '  Fact.:' + Car.cod_Factura + ' No.Pago: ' + convert(varchar(30), Car.NPago)  as 'AsientoNotas'" _
               & ",T.descripcion + '.. ID: ' + convert(varchar(10),Car.[ID]) + '.. Prov: ' + + convert(varchar(10),Car.COD_PROVEEDOR)   as 'AsientoDesc'" _
               & ",T.descripcion as 'Detalle', T.cod_cuenta as 'CtaCargo',P.cod_cuenta as 'CtaProveedor'" _
               & ", convert(varchar(10),P.COD_PROVEEDOR) + '..' + P.Descripcion as 'Referencia'" _
               & " from cxp_PagoProvCargos Car  inner join cxp_Cargos T on Car.cod_Cargo = T.cod_Cargo" _
               & " inner join cxp_proveedores P on Car.cod_Proveedor = P.cod_Proveedor" _
               & " where Car.TIPO_PROCESO = 'D' and Car.Asiento_Fecha is null and isnull(Car.ID,0) = 0" _
               & " and Car.REGISTRO_FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'  order by Car.REGISTRO_FECHA"
  
  Case "CA" 'Cargo de Anticipo de Factura Cancelada vía Cargo/Retención
        strSQL = "select Car.*, convert(varchar(10),Car.[ID]) + '.' + rtrim(Car.cod_Factura) as 'Cod_Transaccion'" _
               & ",'Cargo de Anticipo/Fact.Cancelada vía Ret. Prov:' + P.descripcion + '  Fact.:' + Car.cod_Factura + ' No.Pago: ' + convert(varchar(30), Pg.NPago)  as 'AsientoNotas'" _
               & ",T.descripcion + '.. ID: ' + convert(varchar(10),Car.[ID]) + '.. Prov: ' + + convert(varchar(10),Car.COD_PROVEEDOR)   as 'AsientoDesc'" _
               & ",T.descripcion as 'Detalle', T.cod_cuenta as 'CtaCargo',P.cod_cuenta as 'CtaProveedor'" _
               & ", convert(varchar(10),P.COD_PROVEEDOR) + '..' + P.Descripcion as 'Referencia'" _
               & " from CXP_PAGOPROV Pg inner join cxp_PagoProvCargos Car on Pg.COD_PROVEEDOR = Car.COD_PROVEEDOR" _
               & " and Pg.COD_FACTURA = Car.COD_FACTURA and Pg.NPAGO = Car.NPAGO and isnull(Car.ID,0) > 0" _
               & "  inner join cxp_proveedores P on Car.cod_Proveedor = P.cod_Proveedor" _
               & "  inner join cxp_Cargos T on Car.cod_Cargo = T.cod_Cargo" _
               & "  inner join CXP_ANTICIPOS At on PG.COD_PROVEEDOR = At.COD_PROVEEDOR and At.ID_CARGO = Car.ID" _
               & " where Pg.TIPO_CANCELACION = 'C'" _
               & "  and Car.ASIENTO_FECHA is null and isnull(Car.ID,0) > 0" _
               & "  and Car.REGISTRO_FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'  order by Car.REGISTRO_FECHA"
               

  
End Select
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
lblEstatus.Caption = "Procesando Asientos...(" & pTipoDoc & "->" & pTipoAsiento & ")"
lblEstatus.Refresh

Do While Not rs.EOF
 If fxValidaPeriodoAsiento(rs!Registro_Fecha) Then 'Verificar el Periodo Abierto en contabilidad
   
   If pMascara <> "" Then
       vNumAsiento = pTipoDoc & "." & Format(rs!cod_proveedor, "00") & "." & Format(rs!Cod_Transaccion, pMascara)
   Else
       vNumAsiento = pTipoDoc & "." & Format(rs!cod_proveedor, "00") & "." & rs!Cod_Transaccion
   End If
   
   strSQL = "insert CntX_Asientos(cod_contabilidad,Tipo_Asiento,Num_Asiento,Anio,Mes,Fecha_Asiento,descripcion,balanceado,modulo, notas, referencia)" _
          & " values(" & GLOBALES.gEnlace & ",'" & pTipoAsiento & "','" & vNumAsiento & "'," & Year(rs!Registro_Fecha) & "," & Month(rs!Registro_Fecha) _
          & ",'" & Format(rs!Registro_Fecha, "yyyy/mm/dd") & "','" & Mid(Trim(rs!AsientoDesc), 1, 60) & "','S'," & vModulo _
          & ",'" & fxSysCleanTxtInject(Mid(rs!AsientoNotas, 1, 100)) & "','" & Mid(rs!Referencia, 1, 200) & "')"
    
    'Crea Detalle y Actualizar el estado de la transaccion
    Select Case pTipoDoc
      Case "FT" 'Factura Registrada
            
            strSQL = strSQL & Space(10) & "insert into CntX_Asientos_detalle(cod_contabilidad,TIPO_ASIENTO,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
                   & ",detalle,documento,cod_unidad,cod_divisa,TIPO_Cambio,cod_centro_costo)" _
                   & " (select Asi.COD_CONTABILIDAD,'" & pTipoAsiento & "','" & vNumAsiento & "',Asi.LINEA, Asi.COD_CUENTA" _
                   & " , case when Asi.DebeHaber in('D') then Asi.MONTO else 0 end" _
                   & " , case when Asi.DebeHaber not in('D') then Asi.MONTO else 0 end" _
                   & ",'Prov.' + convert(varchar(10),Tra.cod_proveedor) + '.Fact.' + RTRIM(Tra.cod_factura)" _
                   & ",isnull(Tra.Cod_Factura,''), Asi.COD_UNIDAD,Asi.COD_DIVISA,Asi.Tipo_Cambio, Asi.COD_CENTRO_COSTO" _
                   & " from cxp_facturas Tra inner join cxp_facturas_detalle Asi on Tra.cod_proveedor = Asi.cod_proveedor" _
                   & " and Tra.cod_factura = Asi.cod_factura" _
                   & " where Tra.cod_proveedor = '" & rs!cod_proveedor & "' AND Tra.cod_factura = '" & rs!cod_Factura & "')"
            
            strSQL = strSQL & Space(10) & "update cxp_facturas set asiento_fecha = dbo.MyGetdate(), asiento_generado = 'G'" _
                   & " where cod_proveedor = " & rs!cod_proveedor & " and cod_factura = '" & rs!cod_Factura & "'"
            Call ConectionExecute(strSQL)
      
      
      Case "FA" 'Factura Anulada
            
            strSQL = strSQL & Space(10) & "insert into CntX_Asientos_detalle(cod_contabilidad,TIPO_ASIENTO,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
                   & ",detalle,documento,cod_unidad,cod_divisa,TIPO_Cambio,cod_centro_costo)" _
                   & " (select Asi.COD_CONTABILIDAD,'" & pTipoAsiento & "','" & vNumAsiento & "',Asi.LINEA, Asi.COD_CUENTA" _
                   & " , case when Asi.DebeHaber in('D') then 0 else Asi.MONTO end" _
                   & " , case when Asi.DebeHaber not in('D') then 0 else Asi.MONTO end" _
                   & ",'Prov.' + convert(varchar(10),Tra.cod_proveedor) + '.Fact.' + RTRIM(Tra.cod_factura)" _
                   & ",isnull(Tra.Cod_Factura,''), Asi.COD_UNIDAD,Asi.COD_DIVISA,Asi.Tipo_Cambio, Asi.COD_CENTRO_COSTO" _
                   & " from cxp_facturas Tra inner join cxp_facturas_detalle Asi on Tra.cod_proveedor = Asi.cod_proveedor" _
                   & " and Tra.cod_factura = Asi.cod_factura" _
                   & " where Tra.cod_proveedor = '" & rs!cod_proveedor & "' AND Tra.cod_factura = '" & rs!cod_Factura & "')"
            
            strSQL = strSQL & Space(10) & "update cxp_facturas set anula_asiento_fecha = dbo.MyGetdate()" _
                   & " where cod_proveedor = " & rs!cod_proveedor & " and cod_factura = '" & rs!cod_Factura & "'"
            Call ConectionExecute(strSQL)
      
      Case "CM" 'Cargos Flotante Base Monto
          
          strSQL = strSQL & Space(10) & "insert into CntX_Asientos_detalle(cod_contabilidad,TIPO_ASIENTO,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
                   & ",detalle,documento,cod_unidad,cod_divisa,TIPO_Cambio,cod_centro_costo)" _
                   & " values(" & GLOBALES.gEnlace & ",'" & pTipoAsiento & "','" & vNumAsiento & "',1,'" & rs!CtaProveedor & "'," _
                   & rs!Valor & ",0,'" & Mid(rs!Detalle, 1, 100) & "','" & Format(rs!cod_proveedor, "00") & "." & rs!Cod_Transaccion & "." & Trim(rs!COD_CARGO) _
                   & "','" & vUnidad & "','" & rs!cod_Divisa & "'," & rs!TIPO_CAMBIO & ",'')"
            
          strSQL = strSQL & Space(10) & "insert into CntX_Asientos_detalle(cod_contabilidad,TIPO_ASIENTO,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
                   & ",detalle,documento,cod_unidad,cod_divisa,TIPO_Cambio,cod_centro_costo)" _
                   & " values(" & GLOBALES.gEnlace & ",'" & pTipoAsiento & "','" & vNumAsiento & "',2,'" & rs!CtaCargo & "',0," _
                   & rs!Valor & ",'" & Mid(rs!Detalle, 1, 100) & "','" & Format(rs!cod_proveedor, "00") & "." & rs!Cod_Transaccion & "." & Trim(rs!COD_CARGO) _
                   & "','" & vUnidad & "','" & rs!cod_Divisa & "'," & rs!TIPO_CAMBIO & ",'')"
      
            strSQL = strSQL & Space(10) & "update cxP_cargosPer set asiento_fecha = dbo.MyGetdate(), asiento_usuario = '" & glogon.Usuario & "'" _
                   & " where cod_proveedor = " & rs!cod_proveedor & " and [ID] = " & rs!Id & " and cod_Cargo = '" & rs!COD_CARGO & "'"
            Call ConectionExecute(strSQL)
      
      Case "CP", "CD", "CA"
          'Cargos Flotante Base Porcentual
          'Cargos Directos
          'Cargos Flotante de Anticipos con Facturas Canceladas vía Cargos/Retención
          strSQL = strSQL & Space(10) & "insert into CntX_Asientos_detalle(cod_contabilidad,TIPO_ASIENTO,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
                   & ",detalle,documento,cod_unidad,cod_divisa,TIPO_Cambio,cod_centro_costo)" _
                   & " values(" & GLOBALES.gEnlace & ",'" & pTipoAsiento & "','" & vNumAsiento & "',1,'" & rs!CtaProveedor & "'," _
                   & rs!Monto & ",0,'" & Mid(rs!Detalle, 1, 100) & "','" & Format(rs!cod_proveedor, "00") & "." & rs!Cod_Transaccion _
                   & "','" & vUnidad & "','" & rs!cod_Divisa & "'," & rs!TIPO_CAMBIO & ",'')"
      
          strSQL = strSQL & Space(10) & "insert into CntX_Asientos_detalle(cod_contabilidad,TIPO_ASIENTO,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
                   & ",detalle,documento,cod_unidad,cod_divisa,TIPO_Cambio,cod_centro_costo)" _
                   & " values(" & GLOBALES.gEnlace & ",'" & pTipoAsiento & "','" & vNumAsiento & "',2,'" & rs!CtaCargo & "',0," _
                   & rs!Monto & ",'" & Mid(rs!Detalle, 1, 100) & "','" & Format(rs!cod_proveedor, "00") & "." & rs!Cod_Transaccion _
                   & "','" & vUnidad & "','" & rs!cod_Divisa & "'," & rs!TIPO_CAMBIO & ",'')"
      
            strSQL = strSQL & Space(10) & "update cxp_PagoProvCargos set asiento_fecha = dbo.MyGetdate(), asiento_usuario = '" & glogon.Usuario & "'" _
                   & " where cod_proveedor = " & rs!cod_proveedor & " and IDX_Consec = " & rs!IdX_Consec
            Call ConectionExecute(strSQL)
      
      
    End Select
 
 
 Else
    MsgBox "Existen asientos que no pueden ser trasladados porque el periodo fué cerrado...", vbInformation
 End If 'Periodo
 
 If PrgBar.Value < PrgBar.Max Then PrgBar.Value = PrgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

lblEstatus.Caption = ""
lblEstatus.Refresh
PrgBar.Value = 1


Exit Sub

vError:
    lblEstatus.Caption = ""
    lblEstatus.Refresh
    PrgBar.Value = 1
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 End Sub



Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

lswDocumentos.ListItems.Clear

'Facturas Registradas
strSQL = "select count(*) as 'Casos' from CxP_Facturas " _
       & " where Asiento_Generado = 'P' and Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

If chkBalanceados.Value = vbChecked Then
    strSQL = strSQL & " and dbo.fxCxP_AsientoBalanceado('factura',COD_PROVEEDOR, COD_FACTURA) = 1"
End If
       
Call OpenRecordSet(rs, strSQL)
If rs!Casos > 0 Then
  Set itmX = lswDocumentos.ListItems.Add(, "0x1", "Facturas Registradas")
      itmX.SubItems(1) = rs!Casos
End If
rs.Close
  
'Facturas Anuladas
strSQL = "select count(*) as 'Casos' from CxP_Facturas " _
       & " where Estado = 'A' and anula_asiento_fecha is null and anula_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

If chkBalanceados.Value = vbChecked Then
    strSQL = strSQL & " and dbo.fxCxP_AsientoBalanceado('factura',COD_PROVEEDOR, COD_FACTURA) = 1"
End If

Call OpenRecordSet(rs, strSQL)
If rs!Casos > 0 Then
  Set itmX = lswDocumentos.ListItems.Add(, "0x2", "Facturas Anuladas")
      itmX.SubItems(1) = rs!Casos
End If
rs.Close
 
  
'Cargos Flotante base Monto
strSQL = "select count(*) as 'Casos'" _
     & " from cxP_cargosPer C inner join cxp_cargos T on C.cod_Cargo = T.cod_Cargo  " _
     & " inner join cxp_proveedores P on C.cod_Proveedor = P.cod_Proveedor" _
     & " where C.Tipo = 'M' and C.concepto not in('*** PAGO ANTICIPADO ***')" _
     & " and C.REGISTRO_FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
     & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59' and C.Asiento_Fecha is null"
Call OpenRecordSet(rs, strSQL)
If rs!Casos > 0 Then
  Set itmX = lswDocumentos.ListItems.Add(, "0x3", "Cargos Flotante Base Monto")
      itmX.SubItems(1) = rs!Casos
End If
rs.Close

'Cargos Flotante Base Porcentual
strSQL = "select count(*) as 'Casos'" _
     & " from cxp_PagoProvCargos Car inner join CXP_CARGOSPER Per on Car.COD_PROVEEDOR = Per.COD_PROVEEDOR and Car.ID = Per.ID" _
     & " inner join cxp_Cargos T on Car.cod_Cargo = T.cod_Cargo" _
     & " inner join cxp_proveedores P on Car.cod_Proveedor = P.cod_Proveedor" _
     & " where Per.TIPO = 'P' and Car.TIPO_PROCESO = 'F' and Car.Asiento_Fecha is null" _
     & " and Car.REGISTRO_FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
     & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
Call OpenRecordSet(rs, strSQL)
If rs!Casos > 0 Then
    Set itmX = lswDocumentos.ListItems.Add(, "0x4", "Cargos Flotante Base Porcentual")
        itmX.SubItems(1) = rs!Casos
End If
rs.Close

'Cargos Directos de la Factura
strSQL = "select count(*) as 'Casos'" _
     & " from cxp_PagoProvCargos Car  inner join cxp_Cargos T on Car.cod_Cargo = T.cod_Cargo" _
     & " inner join cxp_proveedores P on Car.cod_Proveedor = P.cod_Proveedor" _
     & " where Car.TIPO_PROCESO = 'D' and Car.Asiento_Fecha is null and isnull(Car.ID,0) = 0" _
     & " and Car.REGISTRO_FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
     & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59' and Car.Asiento_Fecha is null"
Call OpenRecordSet(rs, strSQL)
If rs!Casos > 0 Then
    Set itmX = lswDocumentos.ListItems.Add(, "0x5", "Cargos Directos")
        itmX.SubItems(1) = rs!Casos
End If
rs.Close



'Cargos Flotantes de Anticipos con Cobro en Factura Cancelada con Retencion o Cargo
strSQL = "select count(*) as 'Casos'" _
     & " from CXP_PAGOPROV Pg inner join cxp_PagoProvCargos Car on Pg.COD_PROVEEDOR = Car.COD_PROVEEDOR and Pg.COD_FACTURA = Car.COD_FACTURA and Pg.NPAGO = Car.NPAGO" _
     & "  inner join CXP_ANTICIPOS At on PG.COD_PROVEEDOR = At.COD_PROVEEDOR and At.ID_CARGO = Car.ID" _
     & " where Pg.TIPO_CANCELACION = 'C'" _
     & "  and Car.ASIENTO_FECHA is null and isnull(Car.ID,0) > 0" _
     & "  and Car.REGISTRO_FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
     & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
Call OpenRecordSet(rs, strSQL)
If rs!Casos > 0 Then
    Set itmX = lswDocumentos.ListItems.Add(, "0x6", "Cobro de Anticipos con Facturas Retenidas/Ajustadas")
        itmX.SubItems(1) = rs!Casos
End If
rs.Close



'Carga Transacciones con Asientos desbalanceados
strSQL = "select 'Factura' as Tipo,cod_factura as 'Transacccion', creacion_fecha as 'Fecha'" _
       & ", Creacion_User as 'Usuario', Total as 'Monto','Proveedor.: ' + convert(varchar(30), cod_proveedor) as 'Referencia', Notas " _
       & " from CxP_Facturas " _
       & " where Asiento_Generado = 'P' and Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
       & " and dbo.fxCxP_AsientoBalanceado('factura',COD_PROVEEDOR, COD_FACTURA) = 0" _
       & " UNION " _
       & "select 'Factura' as Tipo,cod_factura as 'Transacccion', Anula_fecha as 'Fecha'" _
       & ", Anula_User as 'Usuario', Total as 'Monto','Proveedor.: ' + convert(varchar(30), cod_proveedor) as 'Referencia', Notas " _
       & " from CxP_Facturas " _
       & " where Estado = 'A' and anula_asiento_fecha is null and anula_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
       & " and dbo.fxCxP_AsientoBalanceado('factura',COD_PROVEEDOR, COD_FACTURA) = 0" _
       & " order by FECHA"
Call sbCargaGrid(vGrid, 7, strSQL, True)

Me.MousePointer = vbDefault

Call chkDocumentos_Click

End Sub
