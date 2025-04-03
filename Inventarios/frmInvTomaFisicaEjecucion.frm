VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmInvTomaFisicaEjecucion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ejecutar Toma Fisica"
   ClientHeight    =   2592
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   10776
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2592
   ScaleWidth      =   10776
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton cmdProcesar 
      Height          =   732
      Left            =   8040
      TabIndex        =   0
      Top             =   1560
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Procesar"
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
      Appearance      =   14
      Picture         =   "frmInvTomaFisicaEjecucion.frx":0000
   End
   Begin XtremeSuiteControls.ComboBox cboEntrada 
      Height          =   312
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   5892
      _Version        =   1245187
      _ExtentX        =   10393
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboSalida 
      Height          =   312
      Left            =   1440
      TabIndex        =   3
      Top             =   1920
      Width           =   5892
      _Version        =   1245187
      _ExtentX        =   10393
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Procesar Entradas y Salidas de la Toma Física"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   3120
      TabIndex        =   5
      Top             =   360
      Width           =   6852
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1092
      _Version        =   1245187
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Salidas: "
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1092
      _Version        =   1245187
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Entradas: "
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmInvTomaFisicaEjecucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxTransacConsec(vTipo As String) As Long
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(max(boleta),0) as Ultimo from pv_InvTransac where tipo = '" & vTipo & "'"
Call OpenRecordSet(rs, strSQL)
  fxTransacConsec = rs!ultimo + 1
rs.Close

End Function

Private Sub cmdProcesar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngCodEntrada As Long, lngCodSalida As Long
Dim vCausaEntrada As String, vCausaSalida  As String
Dim vTomaFisica As String, lngTomaFisica As Long, vMascara As String
Dim rsTmp As New ADODB.Recordset, vNotas As String, i As Integer, vFecha As Date
Dim curMontoEntradas As Currency, curMontoSalidas As Currency

'Procedimiento
'1. Identificar Si hay que ejecutar Entradas y Salida
'2. Crear Ordenes de Entrada/Salida segun sea el caso (Mostrar como ejecutadas y Descargadas)
'3. Crear Entradas y Salidas (Afectar Inventarios)
'4. Las Entradas no deben de generar CxP al proveedor.
'5. Actualizar datos de la Toma Fisica (Info. de Ejecución)

On Error GoTo vError

Me.MousePointer = vbHourglass

'Inicializa Variables
vMascara = "0000000000"
lngTomaFisica = GLOBALES.gTag
vTomaFisica = "TF-" & Format(lngTomaFisica, vMascara)
lngCodEntrada = 0
lngCodSalida = 0
vCausaEntrada = cboEntrada.ItemData(cboEntrada.ListIndex)
vCausaSalida = cboSalida.ItemData(cboSalida.ListIndex)
curMontoEntradas = 0
curMontoSalidas = 0


vNotas = "PROCESO AUTOMATICO DE AJUSTE POR TOMA FISICA # " & lngTomaFisica

'Verifica que la toma fisica no haya sido ejecutada
strSQL = "select isnull(count(*),0) as Existe from pv_invTomaFisica" _
       & " where consecutivo = " & lngTomaFisica & " and Estado = 'S'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  Me.MousePointer = vbDefault
  MsgBox "La Toma Física Actual ya ha sido Ejecutada Anteriormente... (Verifique)", vbCritical
  Exit Sub
End If
rs.Close


'Verifica que el periodo se encuentre Abierto
strSQL = "select fecha_corte from pv_invTomaFisica where consecutivo = " & lngTomaFisica
Call OpenRecordSet(rs, strSQL)
If Not fxInvPeriodos(rs!fecha_corte) Then
  Me.MousePointer = vbDefault
  MsgBox "El periodo, en el que se encuentra el corte de la Toma Fisica se encuentra Cerrado...", vbCritical
  Exit Sub
Else
  vFecha = rs!fecha_corte
End If
rs.Close
'Revisa si tiene que Aplicar Entradas y Salidas, de lo contrario queda en Cero cada Variable

' ENTRADAS *************************************************************************
strSQL = "select isnull(count(*),0) as Existe from pv_invTF_detalle" _
       & " where consecutivo = " & lngTomaFisica & " and Existencia_fisica > Existencia_logica"
Call OpenRecordSet(rsTmp, strSQL, 0)
If rsTmp!Existe > 0 Then
    
    lngCodEntrada = fxTransacConsec("E")

    strSQL = "INSERT PV_INVTRANSAC(BOLETA,TIPO,FECHA,COD_ENTSAL,DOCUMENTO,ESTADO,PLANTILLA,NOTAS" _
           & ",FECHA_SISTEMA,GENERA_USER,GENERA_FECHA,AUTORIZA_USER,AUTORIZA_FECHA,PROCESA_USER" _
           & ",PROCESA_FECHA,TOTAL) values('" & Format(lngCodEntrada, vMascara) & "','E','" & Format(vFecha, "yyyy/mm/dd") _
           & "','" & vCausaEntrada & "','" & vTomaFisica & "','P',0,'" & vNotas & "',dbo.MyGetdate(),'" & glogon.Usuario _
           & "',dbo.MyGetdate(),'" & glogon.Usuario & "',dbo.MyGetdate(),'" & glogon.Usuario & "',dbo.MyGetdate(),0)"
    Call ConectionExecute(strSQL)
    
    strSQL = "INSERT INTO PV_INVTRADET(LINEA,BOLETA,TIPO,COD_BODEGA,COD_PRODUCTO,COD_BODEGA_DESTINO,CANTIDAD,PRECIO,DESPACHO)" _
           & "( select ROW_NUMBER() OVER (ORDER BY T.Consecutivo),'" & Format(lngCodEntrada, vMascara) & "','E'" _
           & ",T.COD_BODEGA, T.COD_PRODUCTO, T.COD_BODEGA, (T.existencia_fisica - T.Existencia_logica)" _
           & ",P.Costo_Regular,0" _
           & " from pv_invTF_detalle T inner join pv_productos P on T.cod_producto = P.cod_producto" _
           & " where T.consecutivo = " & lngTomaFisica & " and T.existencia_fisica > T.existencia_logica)"
    Call ConectionExecute(strSQL)

             
    strSQL = "select sum((T.existencia_fisica - T.Existencia_logica) * P.costo_regular) as 'Total'" _
           & " from pv_invTF_detalle T inner join pv_productos P on T.cod_producto = P.cod_producto" _
           & " where T.consecutivo = " & lngTomaFisica & " and T.existencia_fisica > T.existencia_logica"
    Call OpenRecordSet(rs, strSQL, 0)
        curMontoEntradas = rs!Total
    rs.Close
     
    'Actualiza Monto de la Entrada
    strSQL = "update PV_INVTRANSAC set Total = " & curMontoEntradas _
           & " where tipo = 'E' and boleta = '" & Format(lngCodEntrada, vMascara) & "'"
    Call ConectionExecute(strSQL)

End If
rsTmp.Close


' SALIDAS *************************************************************************
strSQL = "select isnull(count(*),0) as Existe from pv_invTF_detalle" _
       & " where consecutivo = " & lngTomaFisica & " and existencia_logica > existencia_fisica"

Call OpenRecordSet(rsTmp, strSQL, 0)
If rsTmp!Existe > 0 Then

    lngCodSalida = fxTransacConsec("S")

    strSQL = "INSERT PV_INVTRANSAC(BOLETA,TIPO,FECHA,COD_ENTSAL,DOCUMENTO,ESTADO,PLANTILLA,NOTAS" _
           & ",FECHA_SISTEMA,GENERA_USER,GENERA_FECHA,AUTORIZA_USER,AUTORIZA_FECHA,PROCESA_USER" _
           & ",PROCESA_FECHA,TOTAL) values('" & Format(lngCodSalida, vMascara) & "','S','" & Format(vFecha, "yyyy/mm/dd") _
           & "','" & vCausaSalida & "','" & vTomaFisica & "','P',0,'" & vNotas & "',dbo.MyGetdate(),'" & glogon.Usuario _
           & "',dbo.MyGetdate(),'" & glogon.Usuario & "',dbo.MyGetdate(),'" & glogon.Usuario & "',dbo.MyGetdate(),0)"
    Call ConectionExecute(strSQL)


    strSQL = "INSERT INTO PV_INVTRADET(LINEA,BOLETA,TIPO,COD_BODEGA,COD_PRODUCTO,COD_BODEGA_DESTINO,CANTIDAD,PRECIO,DESPACHO)" _
           & "( select ROW_NUMBER() OVER (ORDER BY T.Consecutivo),'" & Format(lngCodSalida, vMascara) & "','S'" _
           & ",T.COD_BODEGA, T.COD_PRODUCTO, T.COD_BODEGA, (T.Existencia_logica - T.existencia_fisica)" _
           & ",P.Costo_Regular,0" _
           & " from pv_invTF_detalle T inner join pv_productos P on T.cod_producto = P.cod_producto" _
           & " where T.consecutivo = " & lngTomaFisica & " and T.existencia_logica > T.existencia_fisica )"
    Call ConectionExecute(strSQL)

             
    strSQL = "select sum((T.existencia_logica - T.existencia_fisica) * P.costo_regular) as 'Total'" _
           & " from pv_invTF_detalle T inner join pv_productos P on T.cod_producto = P.cod_producto" _
           & " where T.consecutivo = " & lngTomaFisica & " and T.existencia_logica > T.existencia_fisica"
    Call OpenRecordSet(rs, strSQL, 0)
        curMontoSalidas = rs!Total
    rs.Close
    
    'Actualiza monto de la Salida
    strSQL = "update PV_INVTRANSAC set Total = " & curMontoSalidas _
           & " where tipo = 'S' and boleta = '" & Format(lngCodSalida, vMascara) & "'"
    Call ConectionExecute(strSQL)
    
    
End If
rsTmp.Close


'Procesa Inventario
strSQL = "exec spINVTranProcesa 'E','" & Format(lngCodEntrada, vMascara) & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

strSQL = "exec spINVTranProcesa 'S','" & Format(lngCodSalida, vMascara) & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


'Actualiza datos de la toma fisica (Proceso)
strSQL = "update pv_InvTomaFisica set fecha_aplica = dbo.MyGetdate(),user_aplica = '" & glogon.Usuario _
       & "',causa_entrada = '" & vCausaEntrada & "',causa_salida = '" & vCausaSalida _
       & "',cod_entradaG = " & lngCodEntrada _
       & ",cod_salidaG = " & lngCodSalida & ",estado = 'P'" _
       & " where consecutivo = " & lngTomaFisica
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Toma Física Realizada Satisfactoriamente...", vbInformation

Unload Me

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

Call sbInvESCombo("E", cboEntrada)
Call sbInvESCombo("S", cboSalida)

End Sub
