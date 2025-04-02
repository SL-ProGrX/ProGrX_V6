VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmFNDTrasladoPatrimonio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traslado de Fondo a Cuentas de Patrimonio / Custodias"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   Icon            =   "frmFNDTrasladoPatrimonio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   9240
   Begin XtremeSuiteControls.CheckBox chkContratos 
      Height          =   216
      Left            =   840
      TabIndex        =   6
      Top             =   1500
      Width           =   216
      _Version        =   1310723
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      ForeColor       =   16777215
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      Appearance      =   2
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDTrasladoPatrimonio.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   105
      Left            =   0
      TabIndex        =   0
      Top             =   7740
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4692
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   8892
      _Version        =   524288
      _ExtentX        =   15684
      _ExtentY        =   8276
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
      SpreadDesigner  =   "frmFNDTrasladoPatrimonio.frx":0413
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8400
      TabIndex        =   4
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   612
      Left            =   7560
      TabIndex        =   5
      Top             =   6960
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "&Aplicar"
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
      Picture         =   "frmFNDTrasladoPatrimonio.frx":0B3C
   End
   Begin XtremeSuiteControls.PushButton btnConsulta 
      Height          =   372
      Left            =   7080
      TabIndex        =   7
      Top             =   840
      Width           =   1212
      _Version        =   1310723
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "&Buscar"
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
      Picture         =   "frmFNDTrasladoPatrimonio.frx":1314
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   312
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   6492
      _Version        =   1310723
      _ExtentX        =   11456
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1800
      TabIndex        =   10
      Top             =   480
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3120
      TabIndex        =   11
      Top             =   480
      Width           =   5172
      _Version        =   1310723
      _ExtentX        =   9123
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDestino 
      Height          =   312
      Left            =   1800
      TabIndex        =   14
      Top             =   840
      Width           =   5172
      _Version        =   1310723
      _ExtentX        =   9123
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   0
      Left            =   600
      TabIndex        =   13
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   1
      Left            =   600
      TabIndex        =   12
      Top             =   480
      Width           =   1332
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   372
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   9252
      _Version        =   1310723
      _ExtentX        =   16319
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Seleccione los Contratos a Trasladar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.44
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   2
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   1212
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   6120
      Width           =   7335
   End
End
Attribute VB_Name = "frmFNDTrasladoPatrimonio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean

Private Sub btnConsulta_Click()
 Call sbCargaContratos
End Sub

Private Sub cboOperadora_Click()
txtCodigo_LostFocus
End Sub

Private Sub cboOperadora_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtCodigo.SetFocus
End Sub


Private Sub chkContratos_Click()
Dim i As Integer

For i = 1 To vGrid.MaxRows
 vGrid.Row = i
 vGrid.col = 1
 vGrid.Value = chkContratos.Value
Next i

End Sub


Private Sub sbDocumentoLocal(lngRecibo As Long, vOperadora As Long, vPlan As String, vMonto As Currency)
Dim strSQL As String, rs As New ADODB.Recordset, strLinea(4) As String
Dim strCliente As String, vCuenta As String
Dim curMonto As Currency, i As Integer, vTipo As String
Dim vCuentaFND As String, vConcepto As String
Dim vDivisa  As String


vConcepto = "TRANSFERENCIA DEL FONDO:" & vPlan & " OP:" & vOperadora

vAseDocDetalle = ""
vAseDocDeposito = ""

vDivisa = fxFndDivisa(vOperadora, vPlan)
vCuentaFND = fxgFNDCuentaPlan(vOperadora, vPlan)


'Cuenta de Devoluciones Sistema ASE/SIF
strSQL = "select cta_devoluciones from par_afah"
Call OpenRecordSet(rs, strSQL)
  vCuenta = Trim(rs!cta_devoluciones)
rs.Close

strCliente = "APLICACION GENERAL"

strLinea(1) = "MONTO TRASLADO :" & Format(vMonto, "Standard")
strLinea(2) = "DESTINO        :" & UCase(txtDestino.Text)

strLinea(3) = "PLAN :" & txtDescripcion.Text
strLinea(4) = ""


If GLOBALES.SysDocVersion = 1 Then
    strSQL = "insert fnd_documentos(tipo,id_documento,cod_operadora,cliente,concepto,fecha," _
            & "monto,usuario,detalle1,detalle2,detalle3,detalle4,detalle,dp)" _
            & " values('NC'," & lngRecibo & "," & vOperadora & ",'" & strCliente & "','" & vConcepto _
            & "',dbo.MyGetdate()," _
            & vMonto & ",'" & Trim(glogon.Usuario) _
            & "','" & Mid(strLinea(1), 1, 40) & "','" & Mid(strLinea(2), 1, 40) & "','" & Mid(strLinea(3), 1, 4) & "','" _
            & Mid(strLinea(4), 1, 40) & "','" & vAseDocDetalle & "','" & vAseDocDeposito & "')"
    Call ConectionExecute(strSQL)
    
    strSQL = "insert fnd_asientos(Cod_Operadora,Tipo,Id_documento,Fnd_Cuenta,Fnd_Monto," _
           & "Fnd_Debehaber) Values(" & vOperadora & ",'NC'," _
           & lngRecibo & ",'" & vCuentaFND & "'," & vMonto & ",'D')"
    Call ConectionExecute(strSQL)
    
    strSQL = "insert fnd_asientos(Cod_Operadora,Tipo,Id_documento,Fnd_Cuenta,Fnd_Monto," _
           & "Fnd_Debehaber) Values(" & vOperadora & ",'NC'," _
           & lngRecibo & ",'" & vCuenta & "'," & vMonto & ",'C')"
    Call ConectionExecute(strSQL)
    
Else

    strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,CLIENTE_IDENTIFICACION,CLIENTE_NOMBRE" _
           & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
           & ",linea1,linea2,linea3,linea4,detalle,documento)" _
           & " values('" & lngRecibo & "','FTRA',dbo.MyGetdate(),'" & glogon.Usuario & "','" _
           & "','" & Trim(strCliente) & "','FND003'," & vMonto & ",'P','" & vOperadora _
           & "','" & vPlan & "','','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
           & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
           & vAseDocDetalle & "','" & vAseDocDeposito & "')"
    Call ConectionExecute(strSQL)
        
    strSQL = "exec spSIFDocsAsiento 'FTRA','" & lngRecibo & "'," & vMonto & "" _
           & ",'D','" & vDivisa & "',1," & GLOBALES.gEnlace & ",'" & GLOBALES.gOficinaUnidad & "','','" & vCuentaFND _
            & "','" & vOperadora & "','" & vPlan & "','" & vAseDocDeposito & "'"
    Call ConectionExecute(strSQL)
    
    strSQL = "exec spSIFDocsAsiento 'FTRA','" & lngRecibo & "'," & vMonto & "" _
           & ",'C','" & vDivisa & "',1," & GLOBALES.gEnlace & ",'" & GLOBALES.gOficinaUnidad & "','','" & vCuenta _
            & "','" & vOperadora & "','" & vPlan & "','" & vAseDocDeposito & "'"
    Call ConectionExecute(strSQL)
    
End If




End Sub

Private Function fxExisteAhorro(vCedula As String) As Boolean
Dim vSQL As String, rsX As New ADODB.Recordset

On Error GoTo vError


vSQL = "select isnull(count(*),0) as Existe from ahorro_consolidado" _
     & " where cedula = '" & vCedula & "'"
rsX.Open vSQL, glogon.Conectado, adOpenStatic
fxExisteAhorro = IIf((rsX!Existe > 0), True, False)
rsX.Close

Exit Function

vError:
fxExisteAhorro = True

End Function


Private Function fxgFNDDocumentoConASE(vTipo As String) As Long
Dim rsX As New ADODB.Recordset, vSQL As String, strCampo As String

'glogon.Conection.BeginTrans
If GLOBALES.SysDocVersion = 1 Then
    Select Case UCase(vTipo)
      Case "RE"
        strCampo = "CS_RECIBO"
      Case "DP"
        strCampo = "CS_DEPOSITO"
      Case "ND"
        strCampo = "CS_NOTA_DEBITO"
      Case "NC"
        strCampo = "CS_NOTA_CREDITO"
        
    End Select
    
    vSQL = "select " & strCampo & " as Consecutivo from ase_consecutivos"
    rsX.Open vSQL, glogon.Conection, adOpenStatic
     fxgFNDDocumentoConASE = rsX!Consecutivo
    rsX.Close
    vSQL = "update ase_consecutivos set " & strCampo & " = " & strCampo & "+ 1"
    glogon.Conection.Execute vSQL
    'glogon.Conection.CommitTrans
Else
 'Control de Documentos version 2
    vSQL = "exec spSIFDocsConsecutivo '" & vTipo & "'"
    rsX.Open vSQL, glogon.Conection, adOpenStatic
     fxgFNDDocumentoConASE = rsX!Consecutivo
    rsX.Close
End If
End Function


Private Sub CmdAplicar_Click()
Dim vOperadora As Long, vPlan As String, vCedula As String, vContrato As Long, vDestino As String
Dim vFecha As Date, curMonto As Currency, curTotal As Currency, vPaso As Boolean, vEstado As String

Dim i As Integer, vNC As Long, vTipo As String, vNC_Pat As Long, vTipoDoc As String
Dim strSQL As String, rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
Dim vDivisa As String, vConcepto As String



vFecha = fxFechaServidor

vOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
vPlan = Trim(txtCodigo.Text)

vDivisa = fxFndDivisa(vOperadora, vPlan)

'Pasos:
' 1. Aplicar Nota de Credito, e Indicar Estado Cancelados
' 1.1 Verificar El Estado de La Persona, para Ver si Aplica el Traslado
'     de Fondos
' 2. Trasladar con Nota de Credito, al Fondo de Patrimonio
' 3. Imprimir Nota de Credito de la Operadora
' 4. Imprimir Nota de Credito del Sistema  (Patrimonio)


'Traslados Autorizados
' Socios = Todos
' Ex-Asociado Interno = Custodias
' Ex-Asociado ExEmpleado = Ninguno
' No Socio = Ninguno

vConcepto = "FND003"
'Consecutivo de la Nota de Credito
If GLOBALES.SysDocVersion = 1 Then
    vTipoDoc = "7"
    vNC = fxgFNDDocumentoConsecutivo("NOTA CREDITO", vOperadora)
Else
    vTipoDoc = "FTRA"
    vNC = fxDocumentoConsecutivo(vTipoDoc)
End If



curTotal = 0
curMonto = 0
For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.col = 1
    vEstado = vGrid.CellTag
    vGrid.col = 2
    vContrato = vGrid.Text
    vGrid.col = 3
    vCedula = vGrid.Text
    vGrid.col = 6
    curMonto = CCur(vGrid.Text)
    
    curTotal = curTotal + curMonto
    
    vGrid.col = 1
    
    If vGrid.Value = vbChecked Then
        'Inserta Anulacion del Total del Contrato en el Detalle
        strSQL = "insert fnd_contratos_detalle(cod_operadora,cod_plan,cod_contrato,monto,fecha" _
               & ",fecha_proceso,tcon,ncon,usuario,cod_concepto,cod_caja) values(" & vOperadora & ",'" & vPlan & "'," & vContrato _
               & "," & curMonto * -1 & ",dbo.MyGetdate()," & Year(vFecha) & Format(Month(vFecha), "00") & ",'" & vTipoDoc & "','" & vNC _
               & "','" & glogon.Usuario & "','" & vConcepto & "','')"
        Call ConectionExecute(strSQL)
        
        'Actualiza el Consolidado del Contrato a CERO.
        strSQL = "update fnd_contratos set aportes = 0, rendimiento = 0 where cod_operadora = " & vOperadora _
               & " and cod_plan = '" & vPlan & "' and cod_contrato = " & vContrato
        Call ConectionExecute(strSQL)
    End If
    
Next i


'Crea Nota de Credito en el Fondo
Call sbDocumentoLocal(vNC, vOperadora, vPlan, curTotal)


'FIN - PROCEDIMIENTO EN EL FONDO
'*******************************************************************************************

'REALIZAR LOS ABONOS en Patrimonio


strSQL = "select S.cedula,S.estadoactual,abs(D.monto) as 'Monto',isnull(A.cedula,'NE') as 'Existe' " _
       & " from Socios S inner join fnd_contratos C on S.cedula = C.cedula" _
       & " inner join fnd_contratos_detalle D on C.cod_operadora = D.cod_operadora and C.cod_plan = D.cod_Plan and C.cod_contrato = D.cod_Contrato" _
       & " left join Ahorro_Consolidado A on S.cedula = A.cedula" _
       & " where D.Tcon = '" & vTipoDoc & "' and D.ncon = '" & vNC & "'"
Call OpenRecordSet(rs, strSQL)



'Sacar el Consecutivo de la nota de Credito en Patrimonio (No cambiar el orden por el TipoDoc)
If GLOBALES.SysDocVersion = 1 Then
    vTipoDoc = "7"
    vNC_Pat = fxDocumentoConsecutivo("NC")
Else
    vTipoDoc = "NC"
    vNC_Pat = fxDocumentoConsecutivo("NC")
End If


Do While Not rs.EOF
  Select Case Trim(txtDestino.Tag)
     Case "O" 'Aporte Obrero
        vTipo = "O"
        If rs!Existe <> "NE" Then
          strSQL = "update ahorro_consolidado set ahorro = ahorro + " & rs!Monto _
                 & " where cedula = '" & Trim(rs!Cedula) & "'"
        Else
          strSQL = "insert ahorro_consolidado(cedula,ahorro,aporte,capitaliza,custodia, extra)" _
                 & " values('" & Trim(rs!Cedula) & "'," & rs!Monto & ",0,0,0,0)"
        End If
     
     Case "P" 'Aporte Patronal
        If rs!EstadoActual = "S" Then
            vTipo = "P"
            If rs!Existe <> "NE" Then
              strSQL = "update ahorro_consolidado set aporte = aporte + " & rs!Monto _
                     & " where cedula = '" & Trim(rs!Cedula) & "'"
            Else
              strSQL = "insert ahorro_consolidado(cedula,ahorro,aporte,capitaliza,custodia,extra)" _
                     & " values('" & Trim(rs!Cedula) & "',0," & rs!Monto & ",0,0,0)"
            End If
        End If
        
        If rs!EstadoActual = "A" Then
            vTipo = "X"
            If rs!Existe <> "NE" Then
              strSQL = "update ahorro_consolidado set custodia = isnull(custodia,0) + " & rs!Monto _
                     & " where cedula = '" & Trim(rs!Cedula) & "'"
            Else
              strSQL = "insert ahorro_consolidado(cedula,ahorro,aporte,capitaliza,custodia,extra)" _
                     & " values('" & Trim(rs!Cedula) & "',0,0,0," & rs!Monto & ",0)"
            End If
        End If
     
     Case "C" 'Capitalizacion Excedentes
        vTipo = "C"
        If rs!Existe <> "NE" Then
          strSQL = "update ahorro_consolidado set capitaliza = capitaliza + " & rs!Monto _
                 & " where cedula = '" & Trim(rs!Cedula) & "'"
        Else
          strSQL = "insert ahorro_consolidado(cedula,ahorro,aporte,capitaliza,custodia,extra)" _
                 & " values('" & Trim(rs!Cedula) & "',0,0," & rs!Monto & ",0,0)"
        End If
  
  End Select
  'Guarda Consolidado
  Call ConectionExecute(strSQL)
  
  'Guarda detalle
  strSQL = "insert ahorro_detallado(cedula,tipo,monto,fecha,fechaproc,estado,numcom,Tcon,Ncon,usuario,cod_Caja,cod_concepto) values('" _
         & Trim(rs!Cedula) & "','" & vTipo & "'," & rs!Monto & ",dbo.MyGetdate()," & Year(vFecha) & Format(Month(vFecha), "00") _
         & ",'A','NC-" & vNC_Pat & "','" & vTipoDoc & "','" & vNC_Pat & "','" & glogon.Usuario & "','','" & vConcepto & "')"
  Call ConectionExecute(strSQL)

 rs.MoveNext
Loop
rs.Close


'********************************************************************************
'Crea NOTA CREDITO MAESTRO

curMonto = 0
   
   'Documentos 2
    strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
           & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,cod_oficina" _
           & ",linea1,linea2,linea3,detalle,documento)" _
           & " values('" & vNC_Pat & "','" & vTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "',''" _
           & ",'APLICACION GENERAL','" & vConcepto & "',0,'P','" & vOperadora & "','" & vPlan & "','" & GLOBALES.gOficinaTitular _
           & "','Op.:" & cboOperadora.Text & "','Plan :" & vPlan & "','" & txtDescripcion & "','" _
           & txtDestino.Text & "','')"
                
    Call ConectionExecute(strSQL)
    
    
    strSQL = "select sum(A.monto) as Monto,A.tipo,S.estadoactual" _
           & " from ahorro_detallado A inner join socios S on A.cedula = S.cedula" _
           & " where A.tcon = '" & vTipoDoc & "' and A.Ncon = '" & vNC_Pat & "' group by A.tipo,S.estadoactual"
    Call OpenRecordSet(rs, strSQL)
    
    strSQL = "select cta_custodia,cta_obrero,cta_patronal,cta_capitaliza,cta_devoluciones" _
           & " from par_afah"
    Call OpenRecordSet(rsTmp, strSQL, 0)
    
    Do While Not rs.EOF
      curMonto = curMonto + rs!Monto
      
      Select Case txtDestino.Tag
         Case "O" 'Aporte Obrero
            
            strSQL = "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNC_Pat & "'," & rs!Monto & "" _
                   & ",'C','" & vDivisa & "',1," & GLOBALES.gEnlace & ",'" & GLOBALES.gOficinaUnidad & "','','" & rsTmp!cta_obrero _
                   & "','" & vOperadora & "','" & vPlan & "',''"
            Call ConectionExecute(strSQL)
         Case "P" 'Aporte Patronal
            
            If rs!EstadoActual = "S" Then
                strSQL = "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNC_Pat & "'," & rs!Monto & "" _
                       & ",'C','" & vDivisa & "',1," & GLOBALES.gEnlace & ",'" & GLOBALES.gOficinaUnidad & "','','" & rsTmp!cta_patronal _
                       & "','" & vOperadora & "','" & vPlan & "',''"
                Call ConectionExecute(strSQL)
            
            Else
                strSQL = "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNC_Pat & "'," & rs!Monto & "" _
                       & ",'C','" & vDivisa & "',1," & GLOBALES.gEnlace & ",'" & GLOBALES.gOficinaUnidad & "','','" & rsTmp!cta_custodia _
                       & "','" & vOperadora & "','" & vPlan & "',''"
               Call ConectionExecute(strSQL)
            End If
         
         Case "C" 'Capitalizacion Excedentes
         
            strSQL = "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNC_Pat & "'," & rs!Monto & "" _
                   & ",'C','" & vDivisa & "',1," & GLOBALES.gEnlace & ",'" & GLOBALES.gOficinaUnidad & "','','" & rsTmp!cta_capitaliza _
                   & "','" & vOperadora & "','" & vPlan & "',''"
            Call ConectionExecute(strSQL)
      
      End Select
      
      rs.MoveNext
    Loop
    
    'CIERRA ASIENTO
    strSQL = "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNC_Pat & "'," & curMonto & "" _
           & ",'D','" & vDivisa & "',1," & GLOBALES.gEnlace & ",'" & GLOBALES.gOficinaUnidad & "','','" & rsTmp!cta_devoluciones _
           & "','" & vOperadora & "','" & vPlan & "',''"
    Call ConectionExecute(strSQL)
    
    rsTmp.Close
    rs.Close

 

Call sbCargaContratos

Me.MousePointer = vbDefault

MsgBox "Transferencias de Fondos Realizada Satisfactoriamente..." & vbCrLf & vbCrLf _
       & " --> Nota de Credito en el Fondo: " & vNC & vbCrLf & "--> Nota de Crédito en Patrimonio: " & vNC_Pat, vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 cod_Plan from fnd_Planes" _
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_Plan > '" & txtCodigo.Text & "' and Patrimonio_Enlace = 1 order by cod_Plan asc"
    Else
       strSQL = strSQL & " where cod_Plan < '" & txtCodigo.Text & "' and Patrimonio_Enlace = 1 order by cod_Plan desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsultaPlan(rs!cod_Plan)
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 18 'Fondo de Inversion

strSQL = "select rtrim(descripcion) as 'ItmX',cod_operadora as 'IdX' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True
 
vGrid.MaxCols = 6
vGrid.MaxRows = 0

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
 Call sbCargaContratos
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Consulta = "select cod_plan as 'Plan',descripcion from fnd_planes"
   gBusquedas.Filtro = " And cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
                     & " and Patrimonio_Enlace = 1"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      Call sbConsultaPlan(txtCodigo.Text)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
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
    vGrid.Text = CStr(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1)))
  Next i
  vGrid.col = 1
  vGrid.CellTag = rs!EstadoActual
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

  vGrid.MaxRows = vGrid.MaxRows - 1
End Sub


Private Sub sbCargaContratos()
Dim strSQL As String


Me.MousePointer = vbHourglass

strSQL = "select " & chkContratos.Value & ",C.COD_CONTRATO,C.CEDULA,S.NOMBRE,E.descripcion as 'EstadoPersona',(C.APORTES + C.RENDIMIENTO) AS MONTO,S.EstadoActual" _
       & " from fnd_contratos C " _
       & " inner join Fnd_Planes P on C.cod_OPERADORA = P.COD_OPERADORA AND C.COD_PLAN = P.COD_PLAN" _
       & " inner join Socios S on C.cedula = S.cedula" _
       & " inner join AHORRO_CONSOLIDADO A on S.cedula = A.cedula AND A.COD_DIVISA = P.COD_MONEDA" _
       & " inner join AFI_ESTADOS_PERSONA E on S.estadoActual = E.cod_Estado" _
       & " where C.cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & " and C.cod_plan = '" & txtCodigo.Text & "' and C.estado = 'A'" _
       & " and (C.aportes + C.rendimiento) > 0"
Select Case txtDestino.Text
   Case "Aporte Obrero"
        strSQL = strSQL & " and S.estadoActual = 'S'"
   Case "Aporte Patronal"
        strSQL = strSQL & " and S.estadoActual in('A','S')"
   Case "Capitalización"
        strSQL = strSQL & " and S.estadoActual = 'S'"
End Select
strSQL = strSQL & " ORDER BY C.cedula"

Call sbCargaGridLocal(vGrid, 6, strSQL, True)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbConsultaPlan(vPlan As String)
Dim strSQL As String, rs As New ADODB.Recordset

txtCodigo.Text = vPlan
txtDescripcion.Text = ""
txtDestino.Text = ""
vGrid.MaxRows = 0

strSQL = "Select Descripcion, Patrimonio_Tipo, case when Patrimonio_Tipo = 'P' then 'Aporte Patronal'  " _
       & " when Patrimonio_Tipo = 'O' then 'Aporte Obrero' when Patrimonio_Tipo = 'C' then 'Capitalización' else 'No Identificado' end as 'Patrimonio'" _
       & " from Fnd_Planes where Cod_Operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & " And Cod_Plan = '" & Trim(txtCodigo.Text) & "' and Patrimonio_Enlace = 1"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtDescripcion.Text = Trim(rs!Descripcion)
   txtDestino.Text = Trim(rs!Patrimonio)
   txtDestino.Tag = Trim(rs!Patrimonio_Tipo)
End If
rs.Close

End Sub

Private Sub txtCodigo_LostFocus()
 Call sbConsultaPlan(txtCodigo.Text)
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDestino.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   gBusquedas.Filtro = " And Cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
                     = " and Patrimonio_Enlace = 1"
   frmBusquedas.Show vbModal
   
   If Trim(gBusquedas.Resultado) <> "" Then
        Call sbConsultaPlan(gBusquedas.Resultado)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub


