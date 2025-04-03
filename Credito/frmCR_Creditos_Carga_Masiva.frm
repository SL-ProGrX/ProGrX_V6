VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCR_Creditos_Carga_Masiva 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creditos: Carga Masiva "
   ClientHeight    =   7740
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11970
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   372
      Index           =   0
      Left            =   8640
      TabIndex        =   12
      Top             =   7080
      Width           =   1212
      _Version        =   1310723
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   14
      Picture         =   "frmCR_Creditos_Carga_Masiva.frx":0000
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.ComboBox cboLinea 
      Height          =   312
      Left            =   3000
      TabIndex        =   0
      Top             =   1440
      Width           =   6852
      _Version        =   1310723
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   372
      Left            =   9960
      TabIndex        =   1
      Top             =   960
      Width           =   492
      _Version        =   1310723
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmCR_Creditos_Carga_Masiva.frx":0727
   End
   Begin XtremeSuiteControls.PushButton btnCargar 
      Height          =   372
      Left            =   10440
      TabIndex        =   2
      Top             =   960
      Width           =   492
      _Version        =   1310723
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmCR_Creditos_Carga_Masiva.frx":0E27
   End
   Begin XtremeSuiteControls.PushButton btnInfo 
      Height          =   372
      Left            =   10920
      TabIndex        =   3
      Top             =   960
      Width           =   492
      _Version        =   1310723
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmCR_Creditos_Carga_Masiva.frx":1540
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   372
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   6852
      _Version        =   1310723
      _ExtentX        =   12086
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboPrideduc 
      Height          =   312
      Left            =   7920
      TabIndex        =   5
      Top             =   2160
      Width           =   1932
      _Version        =   1310723
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4212
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   11772
      _Version        =   524288
      _ExtentX        =   20764
      _ExtentY        =   7429
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
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_Creditos_Carga_Masiva.frx":1C59
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   312
      Left            =   1440
      TabIndex        =   11
      Top             =   7080
      Width           =   1932
      _Version        =   1310723
      _ExtentX        =   3408
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   372
      Index           =   1
      Left            =   9840
      TabIndex        =   13
      Top             =   7080
      Width           =   1212
      _Version        =   1310723
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   14
      Picture         =   "frmCR_Creditos_Carga_Masiva.frx":2330
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.ComboBox cboConfirma 
      Height          =   312
      Left            =   3000
      TabIndex        =   14
      Top             =   1800
      Width           =   6852
      _Version        =   1310723
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   372
      Left            =   2640
      TabIndex        =   16
      Top             =   240
      Width           =   4332
      _Version        =   1310723
      _ExtentX        =   7641
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Carga de Créditos en Lote"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirma"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   1920
      TabIndex        =   15
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   372
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   7080
      Width           =   852
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   1920
      TabIndex        =   8
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Primer deducción"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   6240
      TabIndex        =   6
      Top             =   2160
      Width           =   1572
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12732
   End
End
Attribute VB_Name = "frmCR_Creditos_Carga_Masiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mAseguradoraId As String

Private Sub sbLimpia()
    vGrid.MaxRows = 0
    txtMonto.Text = 0
'    txtComision.Text = 0
'    txtNeto.Text = 0

End Sub


Private Sub btnAplicar_Click(Index As Integer)
Select Case Index
  Case 0 'aplicar
    If vGrid.MaxRows = 0 Or CCur(txtMonto.Text) = 0 Then
       MsgBox "No existen casos para procesar...[verifique!]", vbExclamation
       Exit Sub
    End If
    Call sbProcesar
  
  Case 1 'cancelar
    txtArchivo.Text = ""
    Call sbLimpia
    
End Select
End Sub

Private Sub btnBuscar_Click()
txtArchivo.Text = ""

With frmContenedor.CD
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]"
        .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If

        If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
            'Ok
        Else
            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
            Exit Sub
        End If

 txtArchivo.Text = .FileName

End With

End Sub

Private Sub btnCargar_Click()
    Call sbCargaArchivo
End Sub

Private Sub btnInfo_Click()
  MsgBox "Archivo de Carga: Microsoft Excel" & vbCrLf _
        & " - Columnas: CEDULA, NOMBRE, MONTO, TASA, PLAZO, DOCUMENTO" & vbCrLf _
        & " - Nombre de la Hoja: IMPORT" _
    , vbInformation, "Información del Archivo de Carga"
End Sub


Private Sub cboLinea_Click()
If vPaso Or cboLinea.ListCount = 0 Then Exit Sub
 Call sbLimpia
End Sub


Private Sub cboPrideduc_Click()
If vPaso Or cboPrideduc.ListCount = 0 Then Exit Sub
 Call sbLimpia
End Sub

Function fxFechaProcesoSiguiente(lngFecha As Long) As Long
Dim strMes As String, strAnio As String, strFecha As String
Dim iMes As Integer, iAnio As Integer
strFecha = Trim(CStr(lngFecha))
     strAnio = Mid(strFecha, 1, 4)
     strMes = Mid(strFecha, 5, 2)
     iAnio = CInt(strAnio)
     iMes = CInt(strMes)
     If CInt(strMes) = 12 Then
         iAnio = iAnio + 1
         strAnio = Trim(str(iAnio))
         strMes = "01"
     Else
       Select Case iMes
       Case 1, 2, 3, 4, 5, 6, 7, 8
         iMes = iMes + 1
         strMes = "0" & Trim(str(iMes))
       Case 9, 10, 11
         iMes = iMes + 1
         strMes = Trim(str(iMes))
       End Select
     End If
     fxFechaProcesoSiguiente = CLng(Trim(strAnio) & Trim(strMes))
End Function

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
Dim vProceso As Long

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


vPaso = True


cboLinea.Clear
cboConfirma.Clear

strSQL = "select rtrim(codigo) as 'IdX' , rtrim(descripcion) + '  ['  + rtrim(codigo) + ']' as 'ItmX'" _
       & " from catalogo where retencion = 'N' and activo = 1" _
       & " and codigo not in(select codigo_ase from fnd_planes) order by descripcion"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 cboLinea.AddItem rs!itmX & ""
 cboLinea.ItemData(cboLinea.ListCount - 1) = CStr(rs!IdX)
 
 cboConfirma.AddItem rs!itmX & ""
 cboConfirma.ItemData(cboConfirma.ListCount - 1) = CStr(rs!IdX)
 
 
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboLinea.Text = rs!itmX & ""
End If
rs.Close


txtArchivo.Text = ""

vGrid.MaxCols = 7
vGrid.MaxRows = 0

vProceso = GLOBALES.glngFechaCR
cboPrideduc.AddItem vProceso

For i = 1 To 6
  vProceso = fxFechaProcesoSiguiente(vProceso)
  cboPrideduc.AddItem vProceso
Next i
cboPrideduc.Text = GLOBALES.glngFechaCR

vPaso = False

End Sub

Private Sub sbCargaArchivo()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset
Dim strCadena As String, curMonto As Currency, curComision As Currency, iLinea As Long

Dim pCliente As String, pProceso As Long, pComision As Currency, pAseguradora As String
Dim pCedula As String, pNombre As String, pDocumento As String
Dim pMonto As Currency, pPlazo As Integer, pTasa As Currency, pCuota As Currency


If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

If cboLinea.ListCount <= 0 Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

curMonto = 0
curComision = 0
iLinea = 0

pProceso = cboPrideduc.Text
pAseguradora = "N/A"
pCliente = cboLinea.ItemData(cboLinea.ListIndex)

'La llave es solo codigo y proceso
'& " and cod_aseguradora = '" & pAseguradora & "'"

strSQL = "delete CRD_CREDITOS_CARGADO_H where codigo = '" & pCliente _
       & "' and PROCESO = " & pProceso
       
Call ConectionExecute(strSQL)

strSQL = "" 'Inicializa Bloque



Set rsExcel = Excel_Load(txtArchivo.Text, "Import")
iLinea = 0

Do While Not rsExcel.EOF

    iLinea = iLinea + 1
    
    pCedula = Trim(CStr(rsExcel!Cedula & ""))
    
    
     If pCedula <> "" Then
                
            pNombre = Trim(CStr(rsExcel!Nombre))
            pMonto = CCur(IIf(IsNull(rsExcel!Monto), 0, rsExcel!Monto))
            pDocumento = CCur(IIf(IsNull(rsExcel!Documento), 0, rsExcel!Documento))
           
            
            
            curMonto = curMonto + pMonto
            
            pPlazo = rsExcel!Plazo
            pTasa = rsExcel!Tasa
            pCuota = 0
                
                strSQL = strSQL & Space(10) & "Insert CRD_CREDITOS_CARGADO_H(LINEA,CODIGO,cod_aseguradora,PROCESO,CEDULA,MONTO,NOMBRE,TIPO" _
                        & ", PLAZO, TASA, CUOTA, COMISION, DOCUMENTO)" _
                        & " VALUES(" & iLinea & ",'" & pCliente & "','" & pAseguradora & "'," & pProceso & ",'" & pCedula & "'," & pMonto & ",'" & pNombre _
                        & "','D'," & pPlazo & "," & pCuota & "," & pTasa & "," & pComision & ", '" & pDocumento & "')"
     End If
  
     
     If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
     End If
  
  rsExcel.MoveNext
Loop



'Procesa Lote Final
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If


'Procesa Revisión de la Carga de Datos
curMonto = 0
curComision = 0

strSQL = "exec spCrd_Creditos_Cargado_Revisado '" & pCliente & "','" & pAseguradora & "'," & pProceso
Call OpenRecordSet(rs, strSQL)

With vGrid
    .MaxRows = 0
    Do While Not rs.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .col = 1
        .Text = rs!Cedula
        .col = 2
        .Text = rs!Nombre
        .col = 3
        .Text = CStr(rs!Monto)
        .col = 4
        .Text = CStr(rs!Plazo)
        .col = 5
        .Text = CStr(rs!Tasa)
        .col = 6
        .Text = CStr(rs!Cuota)
        .col = 7
        .Text = CStr(rs!Documento)
        
        curMonto = curMonto + rs!Monto
'        curComision = curComision + rs!Comision
        
        rs.MoveNext
    Loop
    rs.Close
End With


'Totales
txtMonto.Text = Format(curMonto, "Standard")
'txtComision.Text = Format(curComision, "Standard")
'txtNeto.Text = Format(curMonto - curComision, "Standard")

Me.MousePointer = vbDefault

MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0
    txtMonto.Text = 0
'    txtComision.Text = 0
'    txtNeto.Text = 0
End Sub

Private Function fxMaestroTesoreria(vTipoDocumento As String, vBanco As Integer, vMonto As Currency, vCodigo As String _
                              , vBeneficiario As String, vOP As Long, vDetalle1 As String, vReferencia As Long _
                              , vDetalle2 As String, vCuenta As String, vFecha As Date, vUnidad As String, vConcepto As String) As Long                                  'Regresa el NSOLICITUD
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngSol As Long

strSQL = "insert Tes_Transacciones(cod_concepto,cod_unidad,id_banco,tipo,tipo_Beneficiario,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
       & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,referencia,op,genera,actualiza,user_solicita,autoriza,user_autoriza,fecha_autorizacion)" _
       & " values('" & vConcepto & "','" & vUnidad & "'," & vBanco & ",'" & vTipoDocumento & "',5,'" & vCodigo & "','" & vBeneficiario & "'," & vMonto _
       & ",'" & Format(vFecha, "yyyy/mm/dd") & "','P','P','Pol','C','" & vCuenta _
       & "','" & vDetalle1 & "','" & vDetalle2 & "'," & vReferencia & "," & vOP & ",'S','S','" & glogon.Usuario & "'"
       
If UCase(vTipoDocumento) = "CK" Then
   strSQL = strSQL & ",'S','" & glogon.Usuario & "',dbo.MyGetdate())"
Else
   strSQL = strSQL & ",'N',null,null)"
End If
Call ConectionExecute(strSQL)

strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones"
Call OpenRecordSet(rsX, strSQL, 0)
 strSQL = "select * from Tes_Transacciones where nsolicitud = " & rsX!solicitud
rsX.Close

lngSol = 0

Call OpenRecordSet(rsX, strSQL, 0)
If Trim(rsX!Codigo) = Trim(vCodigo) Then lngSol = rsX!NSolicitud
rsX.Close

If lngSol = 0 Then
  strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones where codigo ='" & vCodigo _
         & "'"
  rsX.CursorLocation = adUseServer
  Call OpenRecordSet(rsX, strSQL, 0)
  lngSol = rsX!solicitud
  rsX.Close
End If

fxMaestroTesoreria = lngSol

End Function



Private Sub sbCreaDetalle(vSolicitud As Long, vCtaConta As String, vMonto As Currency, vDH As String, vLinea As Integer, vUnidad As String)
Dim strSQL As String

strSQL = "insert Tes_Trans_Asiento(nsolicitud,cuenta_contable,monto,debehaber,linea,cod_unidad) values(" _
       & vSolicitud & ",'" & Trim(vCtaConta) & "'," & vMonto & ",'" & vDH _
       & "'," & vLinea & ",'" & vUnidad & "')"
Call ConectionExecute(strSQL)

End Sub

Private Function fxCtaBanco(pBanco As Integer) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select CTACONTA from Tes_Bancos where id_banco =" & pBanco
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
 fxCtaBanco = ""
Else
 fxCtaBanco = rsX!ctaConta
End If
rsX.Close
End Function


Private Function fxCtaPuente(pCodigo As String) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select CtaPuente from Catalogo where codigo  ='" & pCodigo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
     fxCtaPuente = ""
Else
     fxCtaPuente = rsX!CtaPuente
End If

rsX.Close

End Function



Private Sub sbProcesar()
Dim strSQL As String, pCedula As String, i As Long
Dim pClienteId As String, pAseguradora As String, pProceso As Long
Dim pTesoreriaId As Long, vFecha As Date, pConfirma As String

Dim pCuenta As String, pUnidad As String, pConcepto As String, pTipo As String

On Error GoTo vError



pClienteId = cboLinea.ItemData(cboLinea.ListIndex)
pConfirma = cboConfirma.ItemData(cboConfirma.ListIndex)

If pClienteId <> pConfirma Then
   MsgBox "La confirmación de la línea/cliente ha fallado, revise!", vbExclamation
   Exit Sub
End If

pAseguradora = "N/A"
pProceso = cboPrideduc.Text

vFecha = fxFechaServidor


Me.MousePointer = vbHourglass


'Procesa Lote
strSQL = "exec spCrd_Creditos_Cargado_Procesa '" & cboLinea.ItemData(cboLinea.ListIndex) _
       & "'," & cboPrideduc.Text & ",'" & pAseguradora _
       & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

txtArchivo.Text = ""
Call sbLimpia

Me.MousePointer = vbDefault

MsgBox "Cargado en Lote de Créditos realizado satisfactoriamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxRevisaInst(pCedula As String) As String
Dim Resultado As String


Resultado = "Ok"



fxRevisaInst = Resultado
End Function

