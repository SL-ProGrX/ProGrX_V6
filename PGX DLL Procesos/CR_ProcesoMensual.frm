VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCR_ProcesoMensual 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proceso Mensual : Aplicación de Abonos"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3855
   HelpContextID   =   3023
   Icon            =   "CR_ProcesoMensual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optProcesoMensual 
      Caption         =   "Detalla las Deducciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "De CCSS para ASECCSS"
      Top             =   1560
      Width           =   3375
   End
   Begin VB.OptionButton optProcesoMensual 
      Caption         =   "[ Traslados al Fondo de Ahorros ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Reporte de los morosos del periodo"
      Top             =   3360
      Width           =   3495
   End
   Begin VB.OptionButton optProcesoMensual 
      Caption         =   "[ Recalcula Saldo del Mes ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Reporte de los morosos del periodo"
      Top             =   3000
      Width           =   3495
   End
   Begin VB.OptionButton optProcesoMensual 
      Caption         =   "Fecha del Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Fecha Maestra del Proceso Mensual"
      Top             =   480
      Width           =   3495
   End
   Begin VB.OptionButton optProcesoMensual 
      Caption         =   "Genera Deducciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "De ASECCSS Para CCSS"
      Top             =   840
      Width           =   3015
   End
   Begin VB.OptionButton optProcesoMensual 
      Caption         =   "Carga Deducciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "De CCSS para ASECCSS"
      Top             =   1200
      Width           =   3375
   End
   Begin VB.OptionButton optProcesoMensual 
      Caption         =   "Aplica Abonos y Genera Asiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Aplica los abonos reportados a los préstamos"
      Top             =   1920
      Width           =   3375
   End
   Begin VB.OptionButton optProcesoMensual 
      Caption         =   "Recalcula Interes Moratorio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Reporte de los morosos del periodo"
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Frame fraFechaProceso 
      Caption         =   "Cambia Fecha Proceso"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   3855
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "CR_ProcesoMensual.frx":000C
         Left            =   600
         List            =   "CR_ProcesoMensual.frx":0037
         TabIndex        =   11
         ToolTipText     =   "Mes a procesar"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Left            =   2520
         TabIndex        =   10
         ToolTipText     =   "Año a procesar"
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgCambiaFecha 
         Height          =   375
         Left            =   3240
         Picture         =   "CR_ProcesoMensual.frx":009F
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "MES"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "AÑO"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
   End
   Begin MSComctlLib.ProgressBar prgProcesoMensual 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   4800
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      ButtonWidth     =   2037
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ejecutar "
            Key             =   "ejecutar"
            Object.ToolTipText     =   "Ejecuta la opción marcada "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ayuda"
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda sobre el proceso mensual"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Cerrar"
            Key             =   "cerrar"
            Object.ToolTipText     =   "Cierra la venta"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.OptionButton optProcesoMensual 
      Caption         =   "Lista de Inconsistencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Casos extraños encontrados en la aplicación de ahorros"
      Top             =   2280
      Width           =   2895
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   3240
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
            Picture         =   "CR_ProcesoMensual.frx":03A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CR_ProcesoMensual.frx":0C85
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CR_ProcesoMensual.frx":0FA1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgDesgloza 
      Height          =   345
      Left            =   3600
      Picture         =   "CR_ProcesoMensual.frx":12BD
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   225
   End
   Begin VB.Image imgRepInconsistencias 
      Height          =   255
      Left            =   3240
      Picture         =   "CR_ProcesoMensual.frx":15C7
      Stretch         =   -1  'True
      ToolTipText     =   "Reporte de Inconsistencias"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image imgGeneraArchivo 
      Height          =   255
      Left            =   3240
      Picture         =   "CR_ProcesoMensual.frx":18D1
      Stretch         =   -1  'True
      ToolTipText     =   "Vuelve A Generar Archivo"
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00808080&
      Caption         =   "Estatus..."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4560
      Width           =   3855
   End
   Begin VB.Image imgRecalculaSaldosMorosos 
      Height          =   345
      Left            =   3600
      Picture         =   "CR_ProcesoMensual.frx":1BDB
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   225
   End
   Begin VB.Image imgListaInconsistencias 
      Height          =   345
      Left            =   3600
      Picture         =   "CR_ProcesoMensual.frx":1EE5
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   225
   End
   Begin VB.Image imgAplica 
      Height          =   345
      Left            =   3600
      Picture         =   "CR_ProcesoMensual.frx":21EF
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   225
   End
   Begin VB.Image imgCarga 
      Height          =   345
      Left            =   3600
      Picture         =   "CR_ProcesoMensual.frx":24F9
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   225
   End
   Begin VB.Image imgGenera 
      Height          =   345
      Left            =   3600
      Picture         =   "CR_ProcesoMensual.frx":2803
      Stretch         =   -1  'True
      Top             =   720
      Width           =   225
   End
   Begin VB.Image imgFechaProceso 
      Height          =   345
      Left            =   3600
      Picture         =   "CR_ProcesoMensual.frx":2B0D
      Stretch         =   -1  'True
      Top             =   360
      Width           =   225
   End
End
Attribute VB_Name = "frmCR_ProcesoMensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type PRM_CR
  Id_solicitud    As Long      'Numero de Operación
  Codigo          As String    'Código de Préstamo
  Cedula          As String    'Cédula de la Persona
  Abono           As Currency  'Abono que se Recibe
  Cuota           As Currency  'Cuota de la Operación
  Tipo            As String    'Tipo A = Código Atraso, C = Código Normal
  Inconsistencia  As Integer   '0 - > 10
  abIntC          As Currency  'Desglose Abono a Intereses Corrientes de la Cuota
  abIntM          As Currency  'Desglose Abono a Intereses Moratorios de la Cuota
  abAmortiza      As Currency  'Desglose Abono a Principal de la Cuota
  ORG_ID_Sol      As Long      'Número de Operacion Original (Por si refunde a Otra)
  ORG_Codigo      As String    'Código Original (Por si refunde a otra OP)
  ID_Referencia   As Long      'Identity sea de Morosidad u Operacion (Para Agilizar la aplicacion)
  IND_APL         As Integer   'Indica si se debe de APlicar u Omitir el procesamiento del registro
  Saldo           As Currency  'Saldo de la Operacion ** Solo en TIPO = C
  Interes         As Integer   'Tasa de Interes
  MorIntMor       As Currency  'Genera registro Moroso con la Dif. de Ser Menor el Abono
  MorIntCor       As Currency  'Genera registro Moroso con la Dif. de Ser Menor el Abono
  MorAmortiza     As Currency  'Genera registro Moroso con la Dif. de Ser Menor el Abono
  MorFechaProceso As Long      'Genera registro Moroso con la Dif. de Ser Menor el Abono
  AmortizaOtro    As String    'Parametro para Refundir otra Operación o NO
  Opex            As Integer   'Indicador de Ex Socio
  Retencion       As Boolean   'Indica si la deducion tiene o no tratamiento como retencion
End Type
Dim vCaso As PRM_CR
Dim mFechaSistema As String

Private Function fxCedula(Id_solicitud As Long)
Dim rs As New ADODB.Recordset, str As String

fxCedula = ""
str = "select cedula from reg_creditos where id_solicitud=" & Trim(Id_solicitud)
With rs
    .Open str, glogon.Conection, adOpenStatic
    If Not .EOF Then
     fxCedula = !Cedula
    End If
    .Close
End With

End Function

Private Sub sbActualizaSaldos()
Dim strSQL As String, rs As New ADODB.Recordset


'Corrige Refundiciones sin referencias
strSQL = "delete refundiciones where id_solicitud = 0"
glogon.Conection.Execute strSQL


'Actualiza el Saldo del Mes para los casos de liquidacion
'del socios durante el corte de envio y recepción de los abonos

strSQL = "select max(tmp_fecha) as FechaInicio,getdate() as FechaCorte from asientos_tmp where tmp_tipo = 'PRM'"
rs.Open strSQL, glogon.Conection, adOpenStatic
strSQL = "update reg_creditos set saldo_mes = saldo" _
       & " where estado = 'A' and Saldo > 0 and cedula in(select cedula from liquidacion" _
       & " where fecliq between '" & Format(IIf(IsNull(rs!fechainicio), fxFechaServidor, rs!fechainicio), "yyyy/mm/dd") _
       & "' and '" & Format(rs!fechacorte, "yyyy/mm/dd") & "')"
glogon.Conection.Execute strSQL
rs.Close

strSQL = "update reg_creditos set saldo = saldo -saldo, amortiza = amortiza + saldo" _
        & ",estado = 'C' Where Saldo > 0 And Saldo < 1"
glogon.Conection.Execute strSQL

strSQL = "update reg_creditos set saldo = 0, amortiza = amortiza + saldo" _
        & ",estado = 'C'  where saldo between -1 and -0.0009"
glogon.Conection.Execute strSQL

strSQL = "update reg_creditos set estado = 'C' where saldo <= 0 and estado = 'A'" _
        & " and proceso = 'N'"
glogon.Conection.Execute strSQL

strSQL = "update reg_creditos set interesv = int where interesv < 0 and estado = 'A' and saldo > 0"
glogon.Conection.Execute strSQL

strSQL = "Select R.id_solicitud from Vista_Morosidad V inner join Reg_creditos R on V.id_solicitud = R.id_solicitud" _
       & " and R.estado = 'C'"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 strSQL = "Delete From Morosidad where id_solicitud=" & rs!Id_solicitud & " and Estado='A'"
 glogon.Conection.Execute strSQL
 rs.MoveNext
Loop
rs.Close


'Retenciones
strSQL = "update reg_creditos set estado = 'C' where ((cuota * plazo) - amortiza) " _
       & "between -1 and -0.0009 and estado = 'A' and codigo in(select codigo from " _
       & "catalogo where retencion = 'S' or poliza = 'S')"
glogon.Conection.Execute strSQL

strSQL = "update reg_creditos set estado = 'C' where ((cuota * plazo) - amortiza) " _
       & "between 0 and 1 and estado = 'A' and codigo in(select codigo from " _
       & "catalogo where retencion = 'S' or poliza = 'S')"
glogon.Conection.Execute strSQL


End Sub


Private Function fxCodigo(Id_solicitud As Long)
Dim rs As New ADODB.Recordset, str As String

fxCodigo = ""
str = "select codigo from reg_creditos where id_solicitud=" & Trim(Id_solicitud)
With rs
    .Open str, glogon.Conection, adOpenStatic
    If Not .EOF Then
     fxCodigo = !Codigo
    End If
    .Close
End With

End Function

Private Sub EnviaCuotasNormales()
Dim strSQL As String

'HACER ESTO EN EL SERVIDOR

lblStatus.Caption = "Procesando Abonos Ordinarios..."
lblStatus.Refresh

'Inserta Casos donde la cuota es menor o igual al saldo del mes
strSQL = "insert into CuotasEnviadas(id_solicitud,codigo,fecpro,cedula,cuota,morosidad)" _
        & " select id_solicitud,codigo," & GLOBALES.glngFechaCR & ",cedula,cuota,0 from reg_creditos" _
        & " where estado='A' and proceso='N' and SALDO_MES > 0  and saldo > 0 and prideduc<=" _
        & GLOBALES.glngFechaCR & " AND IND_DEDUCE_PLANILLA = 'S' and cuota <= (saldo_mes + (saldo_mes * interesv /1200))"
glogon.Conection.Execute strSQL

'Inserta Casos donde la cuota es mayor al saldo del mes
'Hay que enviar el saldo del mes + los intereses de ese saldo
strSQL = "insert into CuotasEnviadas(id_solicitud,codigo,fecpro,cedula,cuota,morosidad)" _
        & " select id_solicitud,codigo," & GLOBALES.glngFechaCR & ",cedula," _
        & "(saldo_mes + (saldo_mes * interesv / 1200)) as cuotaX,0 from reg_creditos" _
        & " where estado='A' and proceso='N' and SALDO_MES > 0  and saldo > 0 and prideduc<=" _
        & GLOBALES.glngFechaCR & " AND IND_DEDUCE_PLANILLA = 'S' and cuota > (saldo_mes + (saldo_mes * interesv /1200))"
glogon.Conection.Execute strSQL

lblStatus.Caption = "Actualizando Contador Cuotas Planillas..."
lblStatus.Refresh
strSQL = "update reg_creditos set cuotas_planilla= cuotas_planilla + 1 " _
       & "where estado='A' and proceso='N' and SALDO_MES > 0 " _
       & "and prideduc<=" & GLOBALES.glngFechaCR & " AND IND_DEDUCE_PLANILLA = 'S'"
glogon.Conection.Execute strSQL

End Sub

Private Sub EnviaMorosidadMasVieja()
Dim strId_SolicitudAnt As String, strSQL As String, strId_Moro As String
Dim curCuota As String, strCuota As String, strCedula As String, strCu As String
Dim strCodigoA As String, strCed As String, strCedu As String
Dim ilongi As Integer, i As Integer, rec As New ADODB.Recordset, strId_solicitud As String
Dim rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
Dim lngXPBN As Long

On Error GoTo vError
      
strSQL = "select A.id_solicitud from morosidad A INNER JOIN reg_Creditos B" _
   & " ON A.id_solicitud = B.id_solicitud where A.estado='A' and A.ESTADOI <> 'J'" _
   & " and B.IND_DEDUCE_PLANILLA = 'S' and B.estado = 'A'" _
   & " group by A.id_solicitud"
      
            
lblStatus.Caption = "Procesando Cuotas Morosas..."
lblStatus.Refresh
      
With rec
    .CursorLocation = adUseServer
    .Open strSQL, glogon.Conection, adOpenStatic
    
    prgProcesoMensual.Value = 1
    prgProcesoMensual.Max = .RecordCount + 1
    
    Do While Not .EOF
    
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '''''''''ENVIO LA CUOTA MAS VIEJA DE MOROSIDAD CON CODIGOA
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ''SELECCIONE EL MAXIMO DE ID_MORO PUES SE CREA DESPUES DE APLICACION
       ''DE ABONOS EL REGISTRO MOROSO CON LA CUOTA RESPECTIVA
       ''EL MINIMO PORQUE ES LA CUOTA MOROSA MAS VIEJA DE LAS QUE
       ''EXISTAN LA QUE SE ENVIA.
    
    
              strSQL = "select max(id_moro) as mid_moro from MOROSIDAD" _
                      & " where fechap=" & "(select min(fechap) from morosidad " _
                      & " where id_solicitud=" & rec!Id_solicitud _
                      & " and estado='A' AND ESTADOI <> 'J'" _
                      & " group by id_solicitud) and id_solicitud=" & !Id_solicitud
              
                  rs.Open strSQL, glogon.Conection, adOpenStatic
                  If Not rs.EOF And Not rs.BOF And IsNull(rs!mid_moro) = False Then
                          strSQL = "select R.id_solicitud,R.cedula,(M.intc+M.intm+M.amortiza) as Cuota,C.codigoa" _
                                 & " from morosidad M inner join Reg_Creditos R on M.id_solicitud = R.id_solicitud" _
                                 & " inner join Catalogo C on M.codigo = C.codigo" _
                                 & " Where M.id_moro = " & rs!mid_moro
                          rsTmp.Open strSQL, glogon.Conection, adOpenStatic
                        
                          strSQL = "insert into CuotasEnviadas(id_solicitud,codigo,fecpro,cedula," _
                                 & "cuota,morosidad)values(" & rsTmp!Id_solicitud & ",'" _
                                 & rsTmp!codigoa & "'," & GLOBALES.glngFechaCR & ",'" & Trim(rsTmp!Cedula) _
                                 & "'," & rsTmp!Cuota & ",1)"
                          glogon.Conection.Execute strSQL
                          rsTmp.Close
                  End If
                  rs.Close
             
             .MoveNext
             If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
             lblStatus.Caption = "Creando Arc.Mora. Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
             lblStatus.Refresh
    
    Loop
    .Close
    End With

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub AbonosAplicados(FechaProceso As Long)
Dim rs As New ADODB.Recordset, lngCuenta As Long, strSQL As String

'Limpia Tabla
glogon.Conection.Execute "delete aplicacioncr"

prgProcesoMensual.Min = 1
prgProcesoMensual.Value = 1

lblStatus.Caption = "Cargando Códigos ..."
lblStatus.Refresh

If lngCuenta = 0 Then

 'Carga catalogo en la tabla APLICACIONCR
 rs.Open "select codigo from catalogo", glogon.Conection, adOpenStatic
 prgProcesoMensual.Max = rs.RecordCount + 1
 
 Do While Not rs.EOF
  strSQL = "insert aplicacioncr(apl_codigo,apl_fechap) values('" & rs!Codigo _
    & "'," & FechaProceso & ")"
  glogon.Conection.Execute strSQL
  If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
  
  rs.MoveNext
 
 Loop
  rs.Close
  
  Call LLenaAbonosParaAsientos(FechaProceso)
  Call AsientoProcesoMensualCreditos(FechaProceso)
Else
 
 'Todavia no se sabe se van a borrar los datos a futuro
End If

prgProcesoMensual.Value = 1
lblStatus.Caption = "Estatus..."
lblStatus.Refresh

End Sub


Private Sub LLenaAbonosParaAsientos(FP As Long)
Dim rsAbonos As New ADODB.Recordset, strSQL As String

prgProcesoMensual.Value = 1
lblStatus.Caption = "Cargando Abonos Aplicados..."
lblStatus.Refresh

With rsAbonos
 
strSQL = "select codigo,opex,sum(intc) as Intc, sum(intm) as intm, sum(amortiza) as Amortiza" _
       & " From prm_creditos" & " Where fecha_proceso = " & FP & " And ind_paso = 1" _
       & " group by codigo,opex"
 .Open strSQL, glogon.Conection, adOpenStatic
 prgProcesoMensual.Value = .RecordCount + 1
 
 Do While Not .EOF
  strSQL = "update aplicacioncr set "
  Select Case !Opex
   Case 0
     strSQL = strSQL + "apl_nintc = apl_nintc + " & IIf(IsNull(!intc), 0, !intc)
     strSQL = strSQL + ",apl_nintm = apl_nintm + " & IIf(IsNull(!intm), 0, !intm)
     strSQL = strSQL + ",apl_namortiza = apl_namortiza + " & IIf(IsNull(!amortiza), 0, !amortiza)
   Case 1
     strSQL = strSQL + "apl_ointc = apl_ointc + " & IIf(IsNull(!intc), 0, !intc)
     strSQL = strSQL + ",apl_ointm = apl_ointm + " & IIf(IsNull(!intm), 0, !intm)
     strSQL = strSQL + ",apl_oamortiza = apl_oamortiza + " & IIf(IsNull(!amortiza), 0, !amortiza)
  End Select
  strSQL = strSQL + " where apl_codigo = '" & !Codigo & "' and apl_fechap = " & GLOBALES.glngFechaCR
  glogon.Conection.Execute strSQL
  
  If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
  lblStatus.Caption = "Car.Abonos.Apl # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
  lblStatus.Refresh
  .MoveNext
 Loop
 .Close
End With

End Sub


Private Sub AsientoProcesoMensualCreditos(FP As Long)
Dim rsX As New ADODB.Recordset, strSQL As String
Dim rs As New ADODB.Recordset, curTotal As Currency
Dim vFecha As Date

curTotal = 0

vFecha = fxFechaServidor

prgProcesoMensual.Value = 1
lblStatus.Caption = "Generando Asiento del Proceso..."
lblStatus.Refresh

With rsX
 
 strSQL = "select A.*,C.*" _
        & " from aplicacioncr A inner join catalogo C on A.apl_codigo = C.codigo" _
        & " where apl_fechap = " & FP
 
 .Open strSQL, glogon.Conection, adOpenStatic
 prgProcesoMensual.Max = .RecordCount + 1
 
 Do While Not .EOF
   
   If !APL_NINTC > 0 Then
     strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber" _
          & ",tmp_fecha,tmp_estado_asiento) values('PRM','" & glogon.Usuario & "','ABONOS - " & !apl_codigo & "','" _
          & !ctanintc & "'," & !APL_NINTC & ",'H','" _
          & Format(vFecha, "yyyy/mm/dd") & "','P')"
     glogon.Conection.Execute strSQL
     curTotal = curTotal + !APL_NINTC
   End If
  
   If !APL_NINTM > 0 Then
     strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber" _
          & ",tmp_fecha,tmp_estado_asiento) values('PRM','" & glogon.Usuario & "','ABONOS - " & !apl_codigo & "','" _
          & !ctanintm & "'," & !APL_NINTM & ",'H','" _
          & Format(vFecha, "yyyy/mm/dd") & "','P')"
     glogon.Conection.Execute strSQL
     curTotal = curTotal + !APL_NINTM
   End If
  
   If !APL_NAMORTIZA > 0 Then
     strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber" _
          & ",tmp_fecha,tmp_estado_asiento) values('PRM','" & glogon.Usuario & "','ABONOS - " & !apl_codigo & "','" _
          & !ctanamort & "'," & !APL_NAMORTIZA & ",'H','" _
          & Format(vFecha, "yyyy/mm/dd") & "','P')"
     glogon.Conection.Execute strSQL
     curTotal = curTotal + !APL_NAMORTIZA
   End If

   If !APL_OINTC > 0 Then
     strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber" _
          & ",tmp_fecha,tmp_estado_asiento) values('PRM','" & glogon.Usuario & "','ABONOS - " & !apl_codigo & "','" _
          & !ctaointc & "'," & !APL_OINTC & ",'H','" _
          & Format(vFecha, "yyyy/mm/dd") & "','P')"
     glogon.Conection.Execute strSQL
     curTotal = curTotal + !APL_OINTC
   End If
  
   If !APL_OINTM > 0 Then
     strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber" _
          & ",tmp_fecha,tmp_estado_asiento) values('PRM','" & glogon.Usuario & "','ABONOS - " & !apl_codigo & "','" _
          & !ctaointm & "'," & !APL_OINTM & ",'H','" _
          & Format(vFecha, "yyyy/mm/dd") & "','P')"
     glogon.Conection.Execute strSQL
     curTotal = curTotal + !APL_OINTM
   End If
  
   If !APL_OAMORTIZA > 0 Then
     strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber" _
          & ",tmp_fecha,tmp_estado_asiento) values('PRM','" & glogon.Usuario & "','ABONOS - " & !apl_codigo & "','" _
          & !ctaoamort & "'," & !APL_OAMORTIZA & ",'H','" _
          & Format(vFecha, "yyyy/mm/dd") & "','P')"
     glogon.Conection.Execute strSQL
     curTotal = curTotal + !APL_OAMORTIZA
   End If
   
   
   If !APL_JINTC > 0 Then
     strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber" _
          & ",tmp_fecha,tmp_estado_asiento) values('PRM','" & glogon.Usuario & "','ABONOS - " & !apl_codigo & "','" _
          & !ctacintc & "'," & !apl_cintc & ",'H','" _
          & Format(vFecha, "yyyy/mm/dd") & "','P')"
     glogon.Conection.Execute strSQL
     curTotal = curTotal + !APL_JINTC
   End If
  
   If !APL_JINTM > 0 Then
     strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber" _
          & ",tmp_fecha,tmp_estado_asiento) values('PRM','" & glogon.Usuario & "','ABONOS - " & !apl_codigo & "','" _
          & !ctacintm & "'," & !APL_JINTM & ",'H','" _
          & Format(vFecha, "yyyy/mm/dd") & "','P')"
     glogon.Conection.Execute strSQL
     curTotal = curTotal + !APL_JINTM
   End If
  
   If !APL_JAMORTIZA > 0 Then
     strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber" _
          & ",tmp_fecha,tmp_estado_asiento) values('PRM','" & glogon.Usuario & "','ABONOS - " & !apl_codigo & "','" _
          & !ctacamort & "'," & !APL_JAMORTIZA & ",'H','" _
          & Format(vFecha, "yyyy/mm/dd") & "','P')"
     glogon.Conection.Execute strSQL
     curTotal = curTotal + !APL_JAMORTIZA
   End If
   
  
  If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
  
  .MoveNext
 Loop
 .Close
 
 'Insertar la Cuenta de Devoluciones e Inconsistencias Aqui, Sacar el Monto de Estadas
 strSQL = "select coalesce(sum(Abono),0) as DevInco" _
        & " from prm_creditos where ind_paso = 0 and fecha_proceso = " & FP
 rs.Open strSQL, glogon.Conection, adOpenStatic
 
 .Open "select * from par_ahcr", glogon.Conection, adOpenStatic
 
 strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber" _
        & ",tmp_fecha,tmp_estado_asiento) values('PRM','" & glogon.Usuario & "','ABONOS - DEV','" _
        & !cr_cta_Incon & "'," & rs!DevInco & ",'H','" _
        & Format(vFecha, "yyyy/mm/dd") & "','P')"
 glogon.Conection.Execute strSQL
 
 'Insertar AQUI LA CUENTA POR COBRAR A LA CCSS, Aplicaciones + Inconsistencias y Devoluciones
 
     strSQL = "insert asientos_tmp(tmp_tipo,tmp_usuario,tmp_caso,tmp_cuenta,tmp_monto,tmp_debehaber" _
          & ",tmp_fecha,tmp_estado_asiento) values('PRM','" & glogon.Usuario & "','ABONOS - CXC','" _
          & !cr_cta_ccpatrono & "'," & curTotal + rs!DevInco & ",'D','" _
          & Format(vFecha, "yyyy/mm/dd") & "','P')"
 .Close
 
 rs.Close
  
 glogon.Conection.Execute strSQL

End With

End Sub

Private Function fxTipoDeCaso(IdX As Long) As String
Dim rsX As New ADODB.Recordset

rsX.CursorLocation = adUseServer
rsX.Open "select opex,proceso from reg_creditos where id_solicitud = " & IdX, glogon.Conection, adOpenStatic
If rsX!Opex = 1 Then
 fxTipoDeCaso = "O"
Else
 fxTipoDeCaso = "N"
End If

If rsX!proceso = "J" Then
 fxTipoDeCaso = "J"
End If
rsX.Close

End Function

Private Sub GeneraMorosidad()
Dim strSQL As String, curIntcds As Currency, curAmortizads As Currency
Dim rs As New ADODB.Recordset, curCuota As Currency, vInt As Currency


prgProcesoMensual.Value = 1
strSQL = " select coalesce(Count(*),0) as Total from Reg_creditos where estado='A' and Proceso='N' and Saldo_Mes > 0" _
       & " and fecult<" & GLOBALES.glngFechaCR & " and prideduc<=" & GLOBALES.glngFechaCR
'Con esto aplican las polizas de vida y otros convenios
' & " and interesv>0"

rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF Then
  prgProcesoMensual.Max = rs!total + 2
End If
rs.Close

strSQL = "select * from Reg_creditos where estado='A' and Proceso='N' and Saldo_Mes > 0" _
    & " and fecult<" & GLOBALES.glngFechaCR & " and prideduc<=" & GLOBALES.glngFechaCR

rs.Open strSQL, glogon.Conection, adOpenForwardOnly
 Do While Not rs.EOF
    
    vInt = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
    If vInt < 0 Then vInt = 0
    
    If rs!Cuota >= rs!Saldo_Mes Then
       curCuota = rs!Saldo_Mes
       curIntcds = (vInt / (12 * 100)) * IIf(IsNull(rs!Saldo_Mes), rs!Saldo, rs!Saldo_Mes)
       curCuota = curCuota + curIntcds
       curAmortizads = curCuota - curIntcds
    Else
       curCuota = rs!Cuota
       curIntcds = (vInt / (12 * 100)) * IIf(IsNull(rs!Saldo_Mes), rs!Saldo, rs!Saldo_Mes)
       curAmortizads = curCuota - curIntcds
    End If
    
    strSQL = "insert into morosidad(codigo,id_solicitud,fechap,intc," _
        & "intm,amortiza,estado,fecap,estadoi,fecult,cuota_morosa) values('" _
        & Trim(rs!Codigo) & "'," & rs!Id_solicitud & "," & GLOBALES.glngFechaCR & "," & curIntcds & "," _
        & "0," & curAmortizads & "," & "'A'," & GLOBALES.glngFechaCR & "," _
        & "'A','" & Format(mFechaSistema, "yyyy/mm/dd") & "'," & curCuota & ")"
    glogon.Conection.Execute strSQL
    
    If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
    lblStatus.Caption = "Gen. Mor. Regs # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
    lblStatus.Refresh
    rs.MoveNext
 Loop
rs.Close
prgProcesoMensual.Value = 1

End Sub

Private Sub sbReporteDetalleDeducciones()
On Error GoTo CapturaError

Me.MousePointer = vbHourglass


With frmContenedor.crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "PROCESO MENSUAL - ABONOS CARGADOS"
     
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "DD/MM/YYYY") & "'"
    .Formulas(3) = "subtitulo='FECHA PROCESO :" & Format(GLOBALES.glngFechaCR, "####-##") _
                 & " USUARIO : " & glogon.Usuario & "'"
    .ReportFileName = GLOBALES.gReportes + "\credito\reportes\Prm_Creditos_Carga.rpt"
    .SelectionFormula = "{PRM_CREDITOS.FECHA_PROCESO}=" & GLOBALES.glngFechaCR
    .PrintReport
End With

Me.MousePointer = vbDefault


Exit Sub
CapturaError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub ReporteGeneracion()

Me.MousePointer = vbHourglass

With frmContenedor.crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "PROCESO MENSUAL - CUOTAS ENVIADAS A DEDUCCION"
 
 .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fecha='" & Format(GLOBALES.glngFechaCR, "####-##") & "'"
 .Formulas(3) = "titulo='CUOTAS ENVIADAS A DEDUCCION'"
 .Formulas(4) = "usuario='" & glogon.Usuario & "'"

 .ReportFileName = GLOBALES.gReportes + "\credito\reportes\crCuotasEnviadas.rpt"
 .SelectionFormula = "{CUOTASENVIADAS.FECPRO}=" & GLOBALES.glngFechaCR
 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub ReporteGeneracionUNI()

Me.MousePointer = vbHourglass

With frmContenedor.crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "PROCESO MENSUAL - CUOTAS ENVIADAS A DEDUCCION"
 
 .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fecha='" & Format(GLOBALES.glngFechaCR, "####-##") & "'"
 .Formulas(3) = "titulo='CUOTAS UNIFICADAS ENVIADAS A DEDUCCION'"
 .Formulas(4) = "usuario='" & glogon.Usuario & "'"

 .ReportFileName = GLOBALES.gReportes & "\credito\reportes\crPrmPlanilla.rpt"
 .SelectionFormula = "{PRM_PLANILLAS.FECPRO}=" & GLOBALES.glngFechaCR
 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub



Private Sub ReporteAplicacion()
On Error GoTo CapturaError

Me.MousePointer = vbHourglass

With frmContenedor.crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "PROCESO MENSUAL - ABONOS APLICADOS"
     
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(GLOBALES.glngFechaCR, "####-##") & "'"
    .Formulas(3) = "titulo='ABONOS APLICADOS'"
    .Formulas(4) = "usuario='" & glogon.Usuario & "'"
    .ReportFileName = GLOBALES.gReportes + "\credito\reportes\crAbonosAplicados.rpt"
    
    .SelectionFormula = "{APLICACIONCR.APL_FECHAP}=" & GLOBALES.glngFechaCR
    
    .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub
CapturaError:
 MsgBox Err.Description, vbCritical
 
End Sub

Private Sub CalculaSaldoMes(Optional vSoloRetenciones As Boolean = False)
Dim rs As New ADODB.Recordset, strSQL As String
On Error GoTo CapturaError

prgProcesoMensual.Min = 1

If Not vSoloRetenciones Then
    
    'Actualiza el Saldo Del Mes para los creditos (cartera)
    strSQL = "select R.id_solicitud,(R.saldo - coalesce(V.amortiza,0)) as SaldoMes" _
           & " from reg_creditos R left join Vista_Morosidad V on R.id_solicitud = V.id_solicitud" _
           & " inner join Catalogo C on R.codigo = C.codigo" _
           & " Where R.estado = 'A' And R.Saldo > 0 and C.retencion = 'N' and C.poliza = 'N'" _
           & " and (R.saldo - coalesce(V.amortiza,0)) <> R.Saldo_Mes"
    rs.CursorLocation = adUseServer
    rs.Open strSQL, glogon.Conection, adOpenStatic
    
    prgProcesoMensual.Max = rs.RecordCount + 1
    
    
    lblStatus.Caption = "SALDO MES I - Cargando... Casos Irregulares"
    lblStatus.Refresh
    
     Do While Not rs.EOF
        If rs!saldomes > 0 Then
           strSQL = "update reg_creditos set saldo_mes = " & rs!saldomes _
                  & " where id_solicitud = " & rs!Id_solicitud
        Else
           strSQL = "update reg_creditos set saldo_mes = 0 " _
                  & " where id_solicitud = " & rs!Id_solicitud
        End If
        glogon.Conection.Execute strSQL
    
        If prgProcesoMensual.Value < prgProcesoMensual.Max Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
        
        lblStatus.Caption = "Actu.Saldo Mes # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max _
                    & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
        lblStatus.Refresh
        
        rs.MoveNext
     
     Loop
     rs.Close
End If

'Actualiza el Saldo Del Mes para los Retenciones
strSQL = "select R.id_solicitud,(((R.cuota * R.plazo) - R.amortiza) - coalesce(V.amortiza,0)) as SaldoMes" _
       & " from reg_creditos R left join Vista_Morosidad V on R.id_solicitud = V.id_solicitud" _
       & " inner join Catalogo C on R.codigo = C.codigo" _
       & " Where R.plazo < 999 and R.estado = 'A' And R.Saldo > 0 and (C.retencion = 'S' or C.poliza = 'S')" _
       & " and (((R.cuota * R.plazo) - R.amortiza) - coalesce(V.amortiza,0)) <> R.Saldo_Mes"
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic

prgProcesoMensual.Max = rs.RecordCount + 1


lblStatus.Caption = "SALDO MES II - Cargando... Casos Irregulares"
lblStatus.Refresh

 Do While Not rs.EOF
       If rs!saldomes > 0 Then
        strSQL = "update reg_creditos set saldo_mes = " & rs!saldomes _
               & " where id_solicitud = " & rs!Id_solicitud
       Else
        strSQL = "update reg_creditos set saldo_mes = 0 " _
               & " where id_solicitud = " & rs!Id_solicitud
       End If
       
       glogon.Conection.Execute strSQL

    If prgProcesoMensual.Value < prgProcesoMensual.Max Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
    
    lblStatus.Caption = "Actu.Saldo Mes # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max _
                & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
    lblStatus.Refresh
    
    rs.MoveNext
 
 Loop
 rs.Close


prgProcesoMensual = 1
lblStatus.Caption = ""

Exit Sub

CapturaError:
 MsgBox Err.Description, vbCritical

End Sub


Private Function fxExisteMorosidad(strNs As String) As Boolean
Dim str As String, rec As New ADODB.Recordset

fxExisteMorosidad = False
str = "select id_moro from MOROSIDAD where id_solicitud=" & strNs & " and estado='A'"
With rec
 .Open str, glogon.Conection, adOpenStatic
 If Not .EOF And Not IsNull(!id_moro) Then
    fxExisteMorosidad = True
 End If
 .Close
End With

End Function

Private Sub ReporteInconsistencias(lngFecha As Long)
    
On Error GoTo CapturaError
 
Me.MousePointer = vbHourglass
 
  With frmContenedor.crt
   .Reset
   .WindowShowGroupTree = True
   .WindowShowPrintSetupBtn = True
   .WindowShowRefreshBtn = True
   .WindowShowSearchBtn = True
   .WindowState = crptMaximized
   .WindowTitle = "PROCESO MENSUAL - INCONSISTENCIAS"
   .ReportFileName = GLOBALES.gReportes & "\credito\reportes\Prm_Creditos_Inconsistencias.rpt"
   .Formulas(0) = "subtitulo='FECHA PROCESO :" & Format(lngFecha, "####-##") _
                 & " USUARIO : " & glogon.Usuario & "'"
   .Formulas(1) = "fecha='" & Format(fxFechaServidor, "DD/MM/YYYY") & "'"
   .Formulas(2) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
   .Formulas(3) = "institucion='CAJA COSTARRICENSE DE SEGURO SOCIAL'"
   .SelectionFormula = "{PRM_CREDITOS.FECHA_PROCESO} = " & lngFecha _
           & " and {PRM_CREDITOS.INCONSISTENCIA} > 0 AND ABS({PRM_CREDITOS.ABONO} - {PRM_CREDITOS.CUOTA}) > 1 "
   .PrintReport
 End With
 
Me.MousePointer = vbDefault
 
Exit Sub

CapturaError:
  MsgBox Err.Description, vbCritical
End Sub

Sub RecalculaCuotaEnMora()
Dim strSQL As String, rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
Dim curIntM As Currency, iMeses As Integer, curDiferencia As Currency

On Error GoTo CapturaError

'Recalculo el interes Moratorio de las cuotas en Mora

Me.MousePointer = vbHourglass

prgProcesoMensual.Min = 1
lblStatus.Caption = "Actualizando Int. Moratorio..."
lblStatus.Refresh


strSQL = "select coalesce(count(*),0) as Total from morosidad where " _
       & "estado='A' AND ESTADOI <> 'J'"
rs.Open strSQL, glogon.Conection, adOpenStatic
prgProcesoMensual.Max = rs!total + 1
rs.Close

strSQL = "select B.id_solicitud,B.codigo,B.montoapr,B.opex,A.* " _
       & "from MOROSIDAD A INNER JOIN REG_CREDITOS B " _
       & "ON A.id_solicitud = B.id_solicitud WHERE A.estado='A' " _
       & "and A.estadoi <> 'J'"
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic

Do While Not rs.EOF
    
    iMeses = Meses(GLOBALES.glngFechaCR, rs!fecap) + 1
   
    strSQL = "select * from RANGOS where codigo='" & rs!Codigo & "'" _
           & " and " & rs!montoapr & " between de and hasta"
    
    curIntM = 0
    
    With rsTmp
     .Open strSQL, glogon.Conection, adOpenStatic
     If Not .EOF And .RecordCount >= 1 Then
      If rs!Opex = 0 Then 'socios
        curIntM = rs!amortiza * (!intm_soc / (12 * 100)) * iMeses
      Else  'exsocios
        curIntM = rs!amortiza * (!intm_nsoc / (12 * 100)) * iMeses
      End If
     End If
     .Close
    End With  'rstmp
    
    strSQL = "update MOROSIDAD set intm=" & curIntM & " where id_moro=" & rs!id_moro
    glogon.Conection.Execute strSQL

If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
lblStatus.Caption = "Act.Int.Mora. Reg # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
lblStatus.Refresh

rs.MoveNext
Loop

rs.Close

'Elimina Inconsistencias

lblStatus.Caption = "Eliminando Inconsistencias Moratorias"
lblStatus.Refresh

strSQL = "update morosidad set intc = 0 where estado = 'A' and intc < 0"
glogon.Conection.Execute strSQL
strSQL = "update morosidad set intm = 0 where estado = 'A' and intm < 0"
glogon.Conection.Execute strSQL
strSQL = "update morosidad set amortiza = 0 where estado = 'A' and amortiza < 0"
glogon.Conection.Execute strSQL

lblStatus.Caption = "Eliminando Mora del Fondo ExtraOrd."
lblStatus.Refresh
'Elimina la mora de los codigos usados en el Modulo del FND
strSQL = "delete morosidad where estado = 'A' and codigo in(" _
       & "select codigo_ase as Codigo from fnd_planes)"
glogon.Conection.Execute strSQL


'Ajustando Mora para Evitar Saldos del Mes Negativos
lblStatus.Caption = "Ajustando Amortizaciones en Mora"
lblStatus.Refresh

strSQL = "select R.id_solicitud,R.saldo,V.amortiza,C.retencion,C.poliza,V.cuota" _
       & " from reg_creditos R inner join vista_morosidad V on R.id_solicitud = V.id_solicitud" _
       & " inner join catalogo C on R.codigo = C.codigo" _
       & " where R.saldo < V.amortiza and C.retencion = 'N' and C.poliza = 'N'"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 If rs!Saldo = 0 Then
   'Borra la mora
   strSQL = "delete morosidad where estado = 'A' and id_solicitud = " & rs!Id_solicitud
   glogon.Conection.Execute strSQL
 Else
   curDiferencia = rs!amortiza - rs!Saldo
   strSQL = "select id_moro,amortiza from morosidad where estado = 'A'" _
          & " and id_solicitud = " & rs!Id_solicitud & " order by amortiza desc"
   rsTmp.Open strSQL, glogon.Conection, adOpenStatic
   Do While Not rsTmp.EOF
    If curDiferencia > 0 Then
     
     If rsTmp!amortiza >= curDiferencia Then
       strSQL = "update morosidad set amortiza = amortiza - " & curDiferencia _
              & " where id_moro = " & rsTmp!id_moro
       glogon.Conection.Execute strSQL
       curDiferencia = 0
      Else
       strSQL = "update morosidad set amortiza = 0 where id_moro = " & rsTmp!id_moro
       glogon.Conection.Execute strSQL
       curDiferencia = curDiferencia - rsTmp!amortiza
      End If
      
     End If 'Diferencia > 0
     rsTmp.MoveNext
   Loop
   rsTmp.Close
   
   If curDiferencia > 0 Then
     MsgBox "Iconsistencia en Morosidad vrs Saldos, Revisar Manualmente la Operacion " _
           & rs!Id_solicitud, vbExclamation
   End If
 
 End If
 rs.MoveNext
Loop
rs.Close

'Actualiza Seguimiento del Proceso Mensual
glogon.Conection.Execute "update par_ahcr set cr_rec = 1"

Call Bitacora("Aplica", "PRM-CREDITO Recalculo Cuentas en Mora")

Call EstadoActualProceso

Me.MousePointer = vbDefault
prgProcesoMensual.Value = 1

MsgBox "Actualización de Intereses Moratorios Realizado ...", vbInformation

lblStatus.Caption = "Estatus..."
lblStatus.Refresh

Exit Sub

CapturaError:
 prgProcesoMensual.Value = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub sbCuotaOrdinaria(rs As ADODB.Recordset)
Dim strSQL As String, lngFechaUltMov As Long, rsX As New ADODB.Recordset
Dim strEstado As String, curIntC As Currency, curAmortiza As Currency
Dim curSaldoMes As Currency

strSQL = "select fecult,saldo,estado,prideduc,saldo_mes,plazo,montoapr,cuota,amortiza" _
       & " from reg_creditos where id_solicitud = " & rs!referencia

rsX.CursorLocation = adUseServer
rsX.Open strSQL, glogon.Conection, adOpenStatic

If rsX.EOF And rsX.BOF Then Exit Sub


If IIf(IsNull(rsX!fecult), rsX!prideduc, rsX!fecult) < GLOBALES.glngFechaCR Then
  lngFechaUltMov = GLOBALES.glngFechaCR
Else
  lngFechaUltMov = fxFechaProcesoSiguiente(IIf(IsNull(rsX!fecult), rsX!prideduc, rsX!fecult))
End If


'Ajusta Fecha de Corte Segun
Select Case rs!Inconsistencia
  Case 2 'Refundiciones
    lngFechaUltMov = GLOBALES.glngFechaCR
  Case 5 'Adelantos de Cuotas Extraordinaria
    lngFechaUltMov = rsX!fecult
End Select

If rs!Retencion = "S" Then
    If (rsX!Cuota * rsX!Plazo) > (rsX!amortiza + rs!amortiza) Then
      strEstado = "A"
    Else
      strEstado = "C"
    End If
Else
    If rsX!Saldo > rs!amortiza Then
      strEstado = "A"
    Else
      strEstado = "C"
    End If
End If

strSQL = "Insert creditos_dt(id_solicitud,codigo,cuota,abono,intcp,amortiza,fechas," _
       & "fechap,tcon,ncon) values(" & rs!Id_solicitud & ",'" & rs!Codigo & "'," _
       & rs!Cuota & "," & rs!Abono & "," & rs!intc & "," _
       & rs!amortiza & ",'" & mFechaSistema & "'," & lngFechaUltMov & ",1," _
       & GLOBALES.glngFechaCR & ")"
glogon.Conection.Execute strSQL
 
'Verificar si hay que enviar la diferencia a morosidad

'Esto es por si la mora que se genera es mayor que el saldo corregirla desde aqui
curAmortiza = 0
curSaldoMes = rsX!Saldo_Mes - rs!amortiza
curAmortiza = IIf((curSaldoMes < rs!MorAmortiza), curSaldoMes, rs!MorAmortiza)

'Genera Mora, pero si la Amortizacion es mayor a 10 colones

If rs!Inconsistencia = 1 And curAmortiza > 10 Then
  strSQL = "insert morosidad(id_solicitud,codigo,fechap,intc,intm,amortiza," _
         & "estado,fecap,estadoi,fecult,cuota_morosa) values(" & rs!Id_solicitud _
         & ",'" & rs!Codigo & "'," & GLOBALES.glngFechaCR & "," & rs!MorIntCor _
         & ",0," & curAmortiza & ",'A'," & GLOBALES.glngFechaCR & ",'A','" _
         & mFechaSistema & "'," & rs!Cuota & ")"
  glogon.Conection.Execute strSQL
End If


If rs!Retencion = "S" Then
    strSQL = "update reg_creditos set amortiza = amortiza + " & rs!amortiza & ",interesc = interesc + " _
            & rs!intc & ",estado = '" & strEstado & "',fecult = " & lngFechaUltMov _
            & " where id_solicitud = " & rs!Id_solicitud
Else
    strSQL = "update reg_creditos set saldo = saldo - " & rs!amortiza _
            & ",amortiza = amortiza + " & rs!amortiza & ",interesc = interesc + " _
            & rs!intc & ",estado = '" & strEstado & "',fecult = " & lngFechaUltMov _
            & ",saldo_mes = saldo_mes - " & rs!amortiza + curAmortiza _
            & " where id_solicitud = " & rs!Id_solicitud
End If
glogon.Conection.Execute strSQL

rsX.Close
 
End Sub


Private Sub sbCuotaMorosa(rs As ADODB.Recordset)
Dim strSQL As String, rsX As New ADODB.Recordset
Dim strEstado As String

rsX.CursorLocation = adUseServer
rsX.Open "select fecult,saldo,estado,saldo_mes,cuota,plazo,amortiza" _
       & " from reg_creditos where id_solicitud = " & rs!Id_solicitud, glogon.Conection, adOpenStatic

If rs!Retencion = "S" Then
    If (rsX!Cuota * rsX!Plazo) > (rsX!amortiza + rs!amortiza) Then
      strEstado = "A"
    Else
      strEstado = "C"
    End If
Else
    If rsX!Saldo > rs!amortiza Then
      strEstado = "A"
    Else
      strEstado = "C"
    End If
End If

strSQL = "update morosidad set abintc = " & rs!intc & ",abintm = " & rs!intm _
       & ",abamortiza = " & rs!amortiza & ",estado = 'C',TCON = 1,NCON = " & GLOBALES.glngFechaCR _
       & ",fecult = '" & mFechaSistema & "' Where id_moro = " & rs!referencia
glogon.Conection.Execute strSQL

'Insertar registro con diferencia
If (rs!MorIntMor + rs!MorAmortiza + rs!MorIntCor) > 0 Then
  strSQL = "insert morosidad(id_solicitud,codigo,fechap,intc,intm,amortiza," _
         & "estado,fecap,estadoi,fecult,cuota_morosa) values(" & rs!Id_solicitud _
         & ",'" & rs!Codigo & "'," & rs!MORFECHAP & "," & rs!MorIntCor _
         & "," & rs!MorIntMor & "," & rs!MorAmortiza & ",'A'," & GLOBALES.glngFechaCR & ",'A','" _
         & mFechaSistema & "'," & rs!Cuota & ")"
  glogon.Conection.Execute strSQL
End If


If rs!Retencion = "S" Then
    strSQL = "update reg_creditos set amortiza = amortiza + " & rs!amortiza _
            & ",interesc = interesc + " & rs!intc + rs!intm & ",estado = '" & strEstado & "'" _
            & " where id_solicitud = " & rs!Id_solicitud
Else
    strSQL = "update reg_creditos set saldo = saldo - " & rs!amortiza _
            & ",amortiza = amortiza + " & rs!amortiza & ",interesc = interesc + " _
            & rs!intc + rs!intm & ",estado = '" & strEstado & "'" _
            & ",saldo_mes = saldo_mes - " & rs!amortiza _
            & " where id_solicitud = " & rs!Id_solicitud
End If

glogon.Conection.Execute strSQL

End Sub


Private Sub AplicaAbonos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset, vFecha As Date

'VARIABLE DE MODULO CON LA FECHA DEL SISTEMA, PARA EVITAR CONSTANTES ACCESOS
'A LA BASE DE DATOS POR CONCEPTO DE FECHA
mFechaSistema = Format(fxFechaServidor, "yyyy/mm/dd")
vFecha = fxFechaServidor


On Error GoTo CapturaError

Me.MousePointer = vbHourglass

prgProcesoMensual.Min = 1
prgProcesoMensual.Value = 1

lblStatus = "Aplicando abonos  "
lblStatus.Refresh

strSQL = "select * from prm_creditos" _
       & " where fecha_proceso =" & GLOBALES.glngFechaCR _
       & " and id_aplicacion = 1 and ind_paso = 0"
rs.Open strSQL, glogon.Conection, adOpenStatic

prgProcesoMensual.Max = rs.RecordCount + 2

Do While Not rs.EOF
  Select Case UCase(Trim(rs!Tipo))
    Case "C"
      Call sbCuotaOrdinaria(rs)
    Case "A"
      Call sbCuotaMorosa(rs)
  End Select
  'INDICA QUE EL REGISTRO YA SE PROCESO
  strSQL = "UPDATE PRM_CREDITOS SET IND_PASO = 1 WHERE PRM_ID = " & rs!PRM_ID
  glogon.Conection.Execute strSQL
  If prgProcesoMensual.Value < prgProcesoMensual.Max Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
  lblStatus.Caption = "Aplicando Registro # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
  lblStatus.Refresh
  rs.MoveNext
Loop
rs.Close

lblStatus = "Generando morosidades para creditos sin abono"
lblStatus.Refresh

'Genera Morosidad
Call GeneraMorosidad

'Asiento
Call AbonosAplicados(GLOBALES.glngFechaCR)

strSQL = "select * from par_ahcr"
rs.Open strSQL, glogon.Conection, adOpenStatic

'Pregunta si quiere Aplicar Inconsistencias
If rs!cr_Apl_Incon = 1 Then
   Call sbActualizaSaldos
   Call CalculaSaldoMes
   Call sbCrAplicaIncon(rs!cr_cta_Incon, GLOBALES.glngFechaCR)
End If
rs.Close

strSQL = "select * from par_ahcr"
rs.Open strSQL, glogon.Conection, adOpenStatic

'Pregunta si quiere Enviar las Devoluciones de Socios al Fondo
If rs!cr_FndSoc = 1 Then
   
   lblStatus.Caption = "Registrando Dev Socios en el Fondo..."
   lblStatus.Refresh
   
   strSQL = "select P.Cedula,S.nombre,coalesce(sum(P.Abono),0) as Abono" _
          & " from prm_creditos P inner Join Socios S ON P.cedula = S.cedula" _
          & " where P.fecha_proceso =" & GLOBALES.glngFechaCR & " and S.EstadoActual = 'S'" _
          & " and P.ind_paso = 0 group by P.cedula,S.nombre"
   rs2.Open strSQL, glogon.Conection, adOpenStatic
   
   prgProcesoMensual.Value = 1
   prgProcesoMensual.Max = rs2.RecordCount + 2
   
   Do While Not rs2.EOF
     'Solo Mayores a un colon
     If rs2!Abono > 1 Then

        Call sbFNDMaestro(rs!cr_fndsoc_op, rs!cr_FndSoc_Plan, rs2!Cedula, rs2!Abono)
        
        strSQL = "delete prm_creditos where Fecha_Proceso = " & GLOBALES.glngFechaCR _
               & " and ind_paso = 0 and cedula = '" & Trim(rs2!Cedula) & "'"
        glogon.Conection.Execute strSQL
        
        'Insertar Detalle en el Fondo del Proceso Mensual Para Reporteria
        Call sbFNDRegistroPRM(rs!cr_fndsoc_op, rs!cr_FndSoc_Plan, rs2!Cedula, rs2!Abono, vFecha, "C")
     
     End If
     prgProcesoMensual.Value = prgProcesoMensual.Value + 1
     rs2.MoveNext
   Loop
   rs2.Close
   
   'Realiza el Asiento Contra la Cuenta de Inconsistencias Antes Cargada
   Call sbFNDAsiento(GLOBALES.glngFechaCR, rs!cr_fndsoc_op, rs!cr_FndSoc_Plan, rs!cr_cta_Incon)

End If


'Pregunta si quiere Enviar las Devoluciones de Ex-Socios al Fondo
If rs!cr_FndexSoc = 1 Then
   
   lblStatus.Caption = "Registrando Dev Ex-Socios en el Fondo..."
   lblStatus.Refresh
   
   strSQL = "select P.Cedula,S.nombre,coalesce(sum(P.Abono),0) as Abono" _
          & " from prm_creditos P inner Join Socios S on P.cedula = S.cedula" _
          & " where P.fecha_proceso =" & GLOBALES.glngFechaCR & " and S.EstadoActual <> 'S'" _
          & " and P.ind_paso = 0 group by P.cedula,S.nombre"
   rs2.Open strSQL, glogon.Conection, adOpenStatic
   
   prgProcesoMensual.Value = 1
   prgProcesoMensual.Max = rs2.RecordCount + 2
   
   Do While Not rs2.EOF
     'Solo Mayores a un colon
     If rs2!Abono > 1 Then
     
        Call sbFNDMaestro(rs!cr_fndExSoc_op, rs!cr_FndExSoc_Plan, rs2!Cedula, rs2!Abono)
        
        strSQL = "delete prm_creditos where Fecha_Proceso = " & GLOBALES.glngFechaCR _
               & " and ind_paso = 0 and cedula = '" & Trim(rs2!Cedula) & "'"
        glogon.Conection.Execute strSQL
        
        'Insertar Detalle en el Fondo del Proceso Mensual Para Reporteria
        Call sbFNDRegistroPRM(rs!cr_fndExSoc_op, rs!cr_FndExSoc_Plan, rs2!Cedula, rs2!Abono, vFecha, "C")
     
     End If
     prgProcesoMensual.Value = prgProcesoMensual.Value + 1
     rs2.MoveNext
   Loop
   rs2.Close
   
   'Realiza el Asiento Contra la Cuenta de Inconsistencias Antes Cargada
   Call sbFNDAsiento(GLOBALES.glngFechaCR, rs!cr_fndExSoc_op, rs!cr_FndExSoc_Plan, rs!cr_cta_Incon)

End If

rs.Close


glogon.Conection.Execute "update par_ahcr set cr_apl = 1"

Call Bitacora("Aplica", "PRM-CREDITO Aplica Abonos")

Call EstadoActualProceso

MsgBox "Información Aplicada ...", vbInformation

lblStatus = "Generando reporte de abonos aplicados"
lblStatus.Refresh

'LIMPIA REGISTROS CON CERO MOVIMIENTOS
glogon.Conection.Execute "DELETE APLICACIONCR" _
            & " Where(APL_NINTC + APL_NINTM + APL_NAMORTIZA + APL_OINTC + APL_OINTM + " _
            & "APL_OAMORTIZA + APL_JINTC + APL_JINTM + APL_JAMORTIZA) = 0"
Call ReporteAplicacion

'Carga Traslados al Fondo
Call sbRepFondo(GLOBALES.glngFechaCR)


lblStatus.Caption = "Estado..."

prgProcesoMensual.Value = 1
Me.MousePointer = vbDefault

Exit Sub

CapturaError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical
  Resume

End Sub


Private Sub sbVerificaCodigos()
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select codigo,refunde,retencion,poliza from catalogo where codigo='" & vCaso.ORG_Codigo & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs.EOF And rs.BOF Then
  vCaso.Tipo = "A"
  rs.Close
  strSQL = "select codigo,refunde,retencion,poliza from catalogo where codigoa='" & vCaso.ORG_Codigo & "'"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  If rs.EOF And rs.BOF Then
    vCaso.Inconsistencia = 6
  Else
    vCaso.Codigo = Trim(rs!Codigo)
    vCaso.AmortizaOtro = rs!refunde
    If rs!Retencion = "S" Or rs!poliza = "S" Then
      vCaso.Retencion = True
      vCaso.AmortizaOtro = "N"
    End If
  
  End If
  
Else
  
  vCaso.Tipo = "C"
  vCaso.Codigo = Trim(rs!Codigo)
  vCaso.AmortizaOtro = rs!refunde
  If rs!Retencion = "S" Or rs!poliza = "S" Then
     vCaso.Retencion = True
     vCaso.AmortizaOtro = "N"
   End If

End If
rs.Close
End Sub

Private Sub sbCalcularRefundicion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curInteres As Currency, curAmortiza As Currency
Dim curMonto As Currency

'Busca: Refunde con otra operacion una con el mismo codigo
strSQL = "select * from refundiciones where id_solicitudr = " & vCaso.Id_solicitud _
       & " and codigo = '" & vCaso.Codigo & "'"
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
  vCaso.abAmortiza = vCaso.Abono
  vCaso.Inconsistencia = 11
Else
  vCaso.ORG_ID_Sol = rs!Id_solicitud
  'Calcula Intereses y Amortizacion para la refundicion
  curMonto = rs!Monto
  strSQL = "select codigo,interesv from reg_creditos where id_solicitud = " & rs!Id_solicitud
  rs.Close
  rs.CursorLocation = adUseServer
  rs.Open strSQL, glogon.Conection, adOpenStatic
  
  vCaso.Interes = rs!interesv
  vCaso.Interes = IIf((vCaso.Interes < 0), 0, vCaso.Interes)
  vCaso.Saldo = curMonto
  
  vCaso.abIntC = vCaso.Saldo * (vCaso.Interes / 1200)

  If vCaso.abIntC > vCaso.Abono Then
   vCaso.abIntC = vCaso.Abono
  Else
   vCaso.abAmortiza = vCaso.Abono - vCaso.abIntC
  End If
End If
rs.Close
End Sub

Private Sub sbVerificaRegistroCreditos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean, rs2 As New ADODB.Recordset
Dim curMonto As Currency, curAbonoMora As Currency

Select Case vCaso.Tipo
 Case "C"
   strSQL = "select codigo,opex,cuota,id_solicitud,saldo_mes,saldo,interesv,plazo,int,fecult,prideduc " _
          & "from reg_creditos where cedula ='" & vCaso.Cedula & "'" _
          & " and codigo ='" & vCaso.Codigo & "' and estado = 'A' and proceso = 'N'"
   rs.CursorLocation = adUseServer
   rs.Open strSQL, glogon.Conection, adOpenStatic
   
   If rs.EOF And rs.BOF Then
    'No encontró la operación
    'Buscar anterior o si no algun prestamo o refundicion, si es que debe hacerlo
     vCaso.Inconsistencia = 3
     rs.Close
     
     
     strSQL = "select id_solicitud,opex from reg_creditos where cedula = '" & vCaso.Cedula & "'" _
            & " and estado = 'A' and proceso = 'N' AND SALDO > 0"
     rs.CursorLocation = adUseServer
     rs.Open strSQL, glogon.Conection, adOpenStatic
     
     vPaso = False
     
     If rs.EOF And rs.BOF Then
       vCaso.Inconsistencia = 8
     
     Else 'EOF, BOF segunda Busqueda
      
      'El estado opex, es para el prestamo en donde va a caer el abono
      'y no el original, ya que este ya fue cancelado supuestamente en una
      'Formalizacion por medio de refundicion
      
      vCaso.Opex = rs!Opex
      
      Do While Not rs.EOF
        strSQL = "select * from refundiciones where id_solicitudr = " & rs!Id_solicitud _
               & " and codigo = '" & vCaso.Codigo & "'"
        rs2.CursorLocation = adUseServer
        rs2.Open strSQL, glogon.Conection, adOpenStatic
        If Not rs2.EOF And Not rs2.BOF Then
          vPaso = True
          vCaso.ORG_ID_Sol = rs2!Id_solicitud
          vCaso.ID_Referencia = rs2!Id_solicitudr
          vCaso.Id_solicitud = rs2!Id_solicitudr
          
          'Calcula Intereses y Amortizacion para la refundicion
          curMonto = rs2!Monto
          strSQL = "select interesv,opex,int from reg_creditos where id_solicitud = " & rs2!Id_solicitud
          rs2.Close
          rs2.CursorLocation = adUseServer
          rs2.Open strSQL, glogon.Conection, adOpenStatic
          
          vCaso.Interes = IIf(IsNull(rs2!interesv), rs2!Int, rs2!interesv)
          vCaso.Interes = IIf((vCaso.Interes < 0), 0, vCaso.Interes)
          vCaso.Saldo = curMonto
          vCaso.abIntC = vCaso.Saldo * (vCaso.Interes / 1200)
          
          If vCaso.abIntC > vCaso.Abono Then
           vCaso.abIntC = vCaso.Abono
           vCaso.abAmortiza = 0
          Else
           vCaso.abAmortiza = vCaso.Abono - vCaso.abIntC
          End If
          
        End If
        rs2.Close
       
      rs.MoveNext
     Loop
     
'
'     rs.Close
'
     End If 'EOF, BOF segunda Busqueda de otros prestamos
     
     If vPaso = False Then
     'No deberia de pasar porque la unica forma de amortiza a otra operacion deberia
     'de ser porque esta fue refundida por otra
       vCaso.abAmortiza = vCaso.Abono
       vCaso.Inconsistencia = 11
     End If

   
   Else 'EOF - Si la encontro
   
   
    'Si de Primeras encuentra la cedula y codigo activos verificar que no sea
    'una nueva operacion con los mismos parametros
    
    If rs!prideduc > GLOBALES.glngFechaCR Then 'Es una refundicion
       
       'El opex, se determina en el caso del prestamo activo y no del estado de las refundiciones
       'ya que el abono se aplica a este ultimo
       
       vCaso.Opex = rs!Opex
       
       vCaso.Id_solicitud = rs!Id_solicitud
       vCaso.ID_Referencia = rs!Id_solicitud
       vCaso.ORG_ID_Sol = rs!Id_solicitud
       vCaso.Inconsistencia = 3
       
       Call sbCalcularRefundicion
    
    Else 'Primer Deduccion, es el registro original de deduccion
    
      vCaso.Opex = rs!Opex
      vCaso.Id_solicitud = rs!Id_solicitud
      vCaso.ID_Referencia = rs!Id_solicitud
      vCaso.ORG_ID_Sol = rs!Id_solicitud
      
      vCaso.Interes = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
      vCaso.Interes = IIf((vCaso.Interes < 0), 0, vCaso.Interes)
      vCaso.Saldo = IIf(IsNull(rs!Saldo_Mes), rs!Saldo, rs!Saldo_Mes)
      vCaso.abIntC = vCaso.Saldo * (vCaso.Interes / 1200)
      
      'Saca nueva cuota para saldos menores que la cuota
      If (vCaso.Saldo + vCaso.abIntC) >= rs!Cuota Then
         vCaso.Cuota = rs!Cuota
      Else
         vCaso.Cuota = (vCaso.Saldo + vCaso.abIntC)
      End If
      
      If vCaso.abIntC > vCaso.Abono Then
        vCaso.MorIntCor = vCaso.abIntC - vCaso.Abono
        vCaso.MorAmortiza = IIf((vCaso.Cuota - vCaso.abIntC > vCaso.Saldo), vCaso.Saldo, vCaso.Cuota - vCaso.abIntC)
        vCaso.abIntC = vCaso.Abono
      Else
        vCaso.abAmortiza = vCaso.Abono - vCaso.abIntC
      End If
      
      If (vCaso.Cuota - vCaso.Abono) < -1 Then
       vCaso.Inconsistencia = 2
       vCaso.MorFechaProceso = GLOBALES.glngFechaCR
      End If
      
      If (vCaso.Abono - vCaso.Cuota) < -1 And IIf(IsNull(rs!fecult), _
            fxFechaProcesoAnterior(GLOBALES.glngFechaCR), rs!fecult) _
            < GLOBALES.glngFechaCR Then vCaso.Inconsistencia = 1
      If Abs(vCaso.Abono - vCaso.Cuota) < 1 Then vCaso.Inconsistencia = 0
      
    End If 'Primer Deduccion
   
   End If 'EOF and BOF
   rs.Close

 
 Case "A"
 
   strSQL = "select max(id_solicitud) as id_solicitud from reg_creditos where cedula ='" & vCaso.Cedula _
          & "' and codigo = '" & vCaso.Codigo & "' and estado ='A' and proceso ='N'"
   
   rs.CursorLocation = adUseServer
   rs.Open strSQL, glogon.Conection, adOpenStatic
   If (rs.EOF And rs.BOF) Or IsNull(rs!Id_solicitud) Then
     
     vCaso.Inconsistencia = 8
   
   Else
     strSQL = "select id_solicitud,opex from reg_creditos where id_solicitud = " & rs!Id_solicitud
     rs.Close
     rs.CursorLocation = adUseServer
     rs.Open strSQL, glogon.Conection, adOpenStatic
     
     vCaso.Opex = rs!Opex
     'Selecciona la cuota mas vieja
     strSQL = "select max(id_moro) as mid_moro from MOROSIDAD" _
          & " where fechap=" & "(select min(fechap) from morosidad " _
          & " where id_solicitud=" & rs!Id_solicitud _
          & " and estado='A' AND ESTADOI <> 'J') AND ID_SOLICITUD = " & rs!Id_solicitud
'          & " group by id_solicitud"
     rs2.CursorLocation = adUseServer
     rs2.Open strSQL, glogon.Conection, adOpenStatic
          
     If Not rs2.EOF And Not rs2.BOF And Not IsNull(rs2!mid_moro) Then
     
     strSQL = "select * from morosidad where id_moro =" & rs2!mid_moro
     
     rs2.Close
           
     rs2.CursorLocation = adUseServer
     rs2.Open strSQL, glogon.Conection, adOpenStatic
     'Repartir Abono
     'Para todos los casos
     vCaso.ID_Referencia = rs2!id_moro
     vCaso.Id_solicitud = rs2!Id_solicitud
     vCaso.ORG_ID_Sol = rs2!Id_solicitud
     vCaso.Cuota = rs2!intc + rs2!intm + rs2!amortiza
     '1
     If vCaso.Abono - (rs2!intc + rs2!intm + rs2!amortiza) < -1 Then
       vCaso.Inconsistencia = 5
       'Generar informacion para la cuota restante (Morosidad)
       curAbonoMora = vCaso.Abono
       If rs2!intc > curAbonoMora Then
         vCaso.abIntC = curAbonoMora
         curAbonoMora = 0
       Else
         vCaso.abIntC = rs2!intc
         curAbonoMora = curAbonoMora - rs2!intc
       End If
       
       If rs2!intm > curAbonoMora Then
         vCaso.abIntM = curAbonoMora
         curAbonoMora = 0
       Else
         vCaso.abIntM = rs2!intm
         curAbonoMora = curAbonoMora - rs2!intm
       End If
       
       If rs2!amortiza > curAbonoMora Then
         vCaso.abAmortiza = curAbonoMora
         curAbonoMora = 0
       Else
         vCaso.abAmortiza = rs2!amortiza
         curAbonoMora = curAbonoMora - rs2!amortiza
       End If
       
       
       vCaso.MorIntCor = rs2!intc - vCaso.abIntC
       vCaso.MorIntMor = rs2!intm - vCaso.abIntM
       vCaso.MorAmortiza = rs2!amortiza - vCaso.abAmortiza
       vCaso.MorFechaProceso = rs2!fechap
       
     End If
     
     '2
     If (rs2!intc + rs2!intm + rs2!amortiza) - vCaso.Abono < -1 Then
       vCaso.Inconsistencia = 4
       vCaso.abIntC = rs2!intc
       vCaso.abIntM = rs2!intm
       vCaso.abAmortiza = vCaso.Abono - (rs2!intc + rs2!intm)
     End If
     
     '3
     If Abs(rs2!intc + rs2!intm + rs2!amortiza - vCaso.Abono) < 1 Then
       vCaso.Inconsistencia = 0
       vCaso.abIntC = rs2!intc
       vCaso.abAmortiza = rs2!amortiza
       vCaso.abIntM = rs2!intm
     End If
            
            
     Else
        vCaso.Inconsistencia = 7
            
     End If ' Not rs2.EOF And Not rs2.BOF
            
     rs2.Close
   
   End If 'EOF MAX()
   rs.Close
End Select
End Sub


Private Sub sbCrGuardaDesgloce()
Dim strSQL As String

'Indica cuales inconsistencias son aplicables y cuales no
'Inconsistencia = 1 (Abono Menor), 2 (Refundicion), 3 (Devolucion)

Select Case vCaso.Inconsistencia
 Case 0, 1, 2, 4, 5
   vCaso.IND_APL = 1
 Case 3, 16
   vCaso.IND_APL = 0
End Select

'insertar aqui el registro
With vCaso
strSQL = "insert prm_creditos(id_solicitud,codigo,cedula,abono,cuota,tipo,inconsistencia," _
        & "intc,intm,amortiza,id_solicitud_original,codigo_original,id_aplicacion,fecha_proceso," _
        & "ind_paso,saldo,interes,morintmor,morintcor,moramortiza,morfechap,refunde,referencia,opex,fecha,retencion)" _
        & " values(" & .Id_solicitud & ",'" & .Codigo & "','" & .Cedula & "'," & .Abono & "," & .Cuota & ",'" & .Tipo & "'," & .Inconsistencia _
        & "," & .abIntC & "," & .abIntM & "," & .abAmortiza & "," & .ORG_ID_Sol & ",'" & .ORG_Codigo & "'," & .IND_APL & "," & GLOBALES.glngFechaCR _
        & ",0," & .Saldo & "," & .Interes & "," & .MorIntMor & "," & .MorIntCor & "," & .MorAmortiza & "," & .MorFechaProceso _
        & ",'" & .AmortizaOtro & "'," & .ID_Referencia & "," & .Opex & ",'" & Format(mFechaSistema, "yyyy/mm/dd") _
        & "','" & IIf(vCaso.Retencion, "S", "N") & "')"
End With
glogon.Conection.Execute strSQL

End Sub

Private Function fxCrDegloceRefundicion(vOperacion As Long, vMonto As Currency, vDias As Integer) As Currency
Dim strSQL As String, rs As New ADODB.Recordset


'Calcula los intereses transcurridos de los prestamos refundidos y
'se los abona al nuevo credito. (los dias son del primero hasta el dia
'de la formalizacion del nuevo credito. Calendario 360 dias.

strSQL = "select R.id_solicitud,R.codigo,R.interesv,X.monto,R.cuota" _
       & " from reg_creditos R inner join refundiciones X on R.id_solicitud = X.id_solicitud" _
       & " where X.id_solicitudr = " & vOperacion
rs.Open strSQL, glogon.Conection, adOpenStatic

Do While Not rs.EOF
  If vMonto > 0 Then
    vCaso.Cuota = rs!Cuota
    vCaso.ORG_ID_Sol = rs!Id_solicitud
    vCaso.ORG_Codigo = rs!Codigo
    vCaso.Saldo = rs!Monto
    vCaso.Interes = rs!interesv
    vCaso.AmortizaOtro = "S"
    
    If vMonto >= (vCaso.Saldo * vCaso.Interes * vDias) / 36000 Then
       vCaso.abIntC = CCur(Format(((vCaso.Saldo * vCaso.Interes * vDias) / 36000), "Standard"))
       vMonto = vMonto - vCaso.abIntC
    Else
       vCaso.abIntC = vMonto
       vMonto = 0
    End If
    
    If vMonto >= vCaso.Cuota - ((vCaso.Saldo * vCaso.Interes * vDias) / 36000) Then
       vCaso.abAmortiza = vCaso.Cuota - ((vCaso.Saldo * vCaso.Interes * vDias) / 36000)
       vMonto = vMonto - vCaso.abAmortiza
    Else
       vCaso.abAmortiza = vMonto
       vMonto = 0
    End If
    
    'Elimina Sobretes con decimales
    If vMonto < 1 Then
      vCaso.abAmortiza = vCaso.abAmortiza + vMonto
      vMonto = 0
    End If
    
    If vCaso.abAmortiza < 0 Then vCaso.abAmortiza = 0
    
    vCaso.Abono = vCaso.abAmortiza + vCaso.abIntC + vCaso.abIntM
    'Insertar Registro Aqui
    Call sbCrGuardaDesgloce
  
  End If
  
  rs.MoveNext

Loop
rs.Close


fxCrDegloceRefundicion = vMonto


End Function

Private Function fxDesgloceMora(vOperacion As Long, vMonto As Currency) As Currency
Dim strSQL As String, rs As New ADODB.Recordset

'strSQL = "select * from morosidad where id_solicitud = " & vOperacion _
'       & " and fechap in(select min(fechap) from morosidad" _
'       & " where estado = 'A' and id_solicitud = " & vOperacion & ")"

strSQL = "select * from morosidad where id_moro in(select min(id_moro) from morosidad" _
       & " where estado = 'A' and id_solicitud = " & vOperacion & ")"
rs.Open strSQL, glogon.Conection, adOpenStatic

'En buena teoria no deberia regresar BOF, ya que fue evaluado con anterioridad

vCaso.Cuota = rs!intc + rs!intm + rs!amortiza
vCaso.ID_Referencia = rs!id_moro

If rs!intm > vMonto Then
  vCaso.abIntM = vMonto
  vMonto = 0
Else
  vCaso.abIntM = rs!intm
  vMonto = vMonto - rs!intm
End If

If rs!intc > vMonto Then
  vCaso.abIntC = vMonto
  vMonto = 0
Else
  vCaso.abIntC = rs!intc
  vMonto = vMonto - rs!intc
End If

If rs!amortiza > vMonto Then
  vCaso.abAmortiza = vMonto
  vMonto = 0
Else
  vCaso.abAmortiza = rs!amortiza
  vMonto = vMonto - rs!amortiza
End If

vCaso.MorIntCor = rs!intc - vCaso.abIntC
vCaso.MorIntMor = rs!intm - vCaso.abIntM
vCaso.MorAmortiza = rs!amortiza - vCaso.abAmortiza
vCaso.MorFechaProceso = rs!fechap
  
If vCaso.MorAmortiza + vCaso.MorIntCor + vCaso.MorIntMor > 0 Then
  vCaso.Inconsistencia = 1
End If

rs.Close

fxDesgloceMora = vMonto

End Function


Private Sub sbDesgloceUnificacionNew()
Dim strSQL As String, rs As New ADODB.Recordset, curMonto As Currency
Dim rsTmp As New ADODB.Recordset, rsTmp2 As New ADODB.Recordset

lblStatus = "Limpiando y Actualizando ..."
lblStatus.Refresh

strSQL = "delete from prm_creditos where fecha_proceso=" & GLOBALES.glngFechaCR
      ' & " and cod_institucion = " & GLOBALES.gInstitucion
glogon.Conection.Execute strSQL 'Borra El Historial en el Proceso de Creditos

Call sbActualizaSaldos   'Actualiza Estado de los creditos y corrige inconsistencias menores
Call CalculaSaldoMes(True)  'Actualiza Saldo del Mes para Retenciones

mFechaSistema = fxFechaServidor

strSQL = "select P.cedula,coalesce(sum(P.monto),0) as Monto,coalesce(sum(V.cuota),0) as Cuota" _
       & " from prm_cargado P left join vPRM_MORA V on P.cedula = V.cedula" _
       & " where P.fecha_proceso = " & GLOBALES.glngFechaCR _
       & " and P.tipo = 3 and P.cod_institucion = 1 group by P.cedula"
rs.Open strSQL, glogon.Conection, adOpenStatic

lblStatus.Caption = "Desglozando Creditos..."
lblStatus.Refresh

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 2

Do While Not rs.EOF
   
    curMonto = rs!Monto
    
    
    If rs!Cuota > 0 Then
        strSQL = "select R.id_solicitud,R.codigo,R.interesv,R.saldo,R.estado," _
               & "C.poliza,C.retencion,C.prioridad,Coalesce(V.intc,0) as IntC," _
               & "Coalesce(V.intm,0) as IntM,Coalesce(V.Amortiza,0) as Amortiza,R.saldo_mes" _
               & ",C.refunde,R.opex,R.cuota,R.FecUlt" _
               & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
               & " left join Vista_morosidad V on R.id_solicitud = V.id_solicitud" _
               & " where R.cedula = '" & Trim(rs!Cedula) & "' and R.prideduc <= " & GLOBALES.glngFechaCR _
               & " and R.proceso = 'N' and R.estado = 'A' and R.saldo > 0" _
               & " and R.ind_deduce_Planilla = 'S'" _
               & " order by C.prioridad asc"
    Else
        strSQL = "select R.id_solicitud,R.codigo,R.interesv,R.saldo,R.estado," _
               & "C.poliza,C.retencion,C.prioridad,0 as IntC,0 as IntM,0 as Amortiza" _
               & ",R.saldo_mes,C.refunde,R.opex,R.cuota,R.FecUlt" _
               & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
               & " where R.cedula = '" & Trim(rs!Cedula) & "' and R.prideduc <= " & GLOBALES.glngFechaCR _
               & " and R.proceso = 'N' and R.estado = 'A' and R.saldo > 0" _
               & " and R.ind_deduce_Planilla = 'S'" _
               & " order by C.prioridad asc"
    End If
    
    rsTmp.Open strSQL, glogon.Conection, adOpenStatic
    
    vCaso.Cedula = Trim(rs!Cedula)
'*******************************************************************************************
'PASO 1 : Revisa y Aplica Primero morosidad de Todas las Operaciones Atrasadas con Prioridad
'*******************************************************************************************
    If rs!Cuota > 0 Then 'Esta Moroso
    
        Do While Not rsTmp.EOF
          If (rsTmp!intc + rsTmp!intm + rsTmp!amortiza) > 0 Then
           If curMonto > 0 Then
                
                vCaso.ORG_ID_Sol = rsTmp!Id_solicitud
                vCaso.ORG_Codigo = rsTmp!Codigo
                vCaso.Codigo = rsTmp!Codigo
                vCaso.Id_solicitud = rsTmp!Id_solicitud
                vCaso.Saldo = rsTmp!Saldo_Mes
                vCaso.Interes = rsTmp!interesv
                vCaso.AmortizaOtro = rsTmp!refunde
                vCaso.Opex = rsTmp!Opex
                
                vCaso.Inconsistencia = 0
                vCaso.IND_APL = 0
                
                If rsTmp!Retencion = "S" Or rsTmp!poliza = "S" Then
                    vCaso.Retencion = True
                Else
                    vCaso.Retencion = False
                End If
               
                 vCaso.Tipo = "A"
                 vCaso.ID_Referencia = 0
                 vCaso.abAmortiza = 0
                 vCaso.abIntC = 0
                 vCaso.abIntM = 0
                 vCaso.Abono = 0
                 vCaso.Cuota = 0
            
                 vCaso.MorIntMor = 0
                 vCaso.MorIntCor = 0
                 vCaso.MorAmortiza = 0
                 vCaso.MorFechaProceso = 0
               
                 curMonto = fxDesgloceMora(rsTmp!Id_solicitud, curMonto)
               
                 vCaso.Abono = vCaso.abAmortiza + vCaso.abIntC + vCaso.abIntM
                 Call sbCrGuardaDesgloce            'Insertar Aqui Registro
             
                 'Actualiza Saldo para El Abono Ordinario
                 vCaso.Saldo = rsTmp!Saldo - vCaso.abAmortiza
           
            End If 'CurMonto > 0
          End If 'Atrasado > 0
          rsTmp.MoveNext
        Loop
        
        'Regresa el Cursor para Aplicar Paso 2
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
        End If
    
    
    End If 'Cuota > 0 -> Esta Moroso

'***********************************************************
'PASO 2 : APLICA ABONOS ORDINARIOS
'***********************************************************
    
    Do While Not rsTmp.EOF
        'inicializa
        vCaso.ORG_ID_Sol = rsTmp!Id_solicitud
        vCaso.ORG_Codigo = rsTmp!Codigo
        vCaso.Codigo = rsTmp!Codigo
        vCaso.Id_solicitud = rsTmp!Id_solicitud
        vCaso.Saldo = rsTmp!Saldo_Mes
        vCaso.Interes = rsTmp!interesv
        vCaso.AmortizaOtro = rsTmp!refunde
        vCaso.Opex = rsTmp!Opex
        
        
        vCaso.Inconsistencia = 0
        vCaso.IND_APL = 0
        
        If rsTmp!Retencion = "S" Or rsTmp!poliza = "S" Then
            vCaso.Retencion = True
        Else
            vCaso.Retencion = False
        End If
        
        vCaso.Tipo = "C"
        vCaso.ID_Referencia = 0
        vCaso.abAmortiza = 0
        vCaso.abIntC = 0
        vCaso.abIntM = 0
        vCaso.Abono = 0
        vCaso.Cuota = 0
    
        vCaso.MorIntMor = 0
        vCaso.MorIntCor = 0
        vCaso.MorAmortiza = 0
        vCaso.MorFechaProceso = 0
      
        If curMonto > 0 Then
            
            vCaso.Saldo = rsTmp!Saldo
            
            'Identifica la Cuota
            If (vCaso.Saldo + ((vCaso.Saldo * vCaso.Interes) / 1200)) > rsTmp!Cuota Then
                vCaso.Cuota = rsTmp!Cuota
            Else
                vCaso.Cuota = vCaso.Saldo + ((vCaso.Saldo * vCaso.Interes) / 1200)
            End If
            
            vCaso.ID_Referencia = rsTmp!Id_solicitud
            
            
            'Ver Aqui Nueva Inconsistencias / TRATAMIENTO PARA ADELANTOS DE CUOTAS
            'SI ES MAYOR/IGUAL QUE LA CUOTA -> ABONO ORDINARIO / SI NO ABONO EXTRAORDINARIO
            'inconsistencia numero 4 y 5/ MOD:2003/08/19
            
            'Verifica que no sea un adelanto de cuota / de lo contrario lo procesa como tal
            If rsTmp!fecult >= GLOBALES.glngFechaCR Then
                    If (vCaso.Cuota - curMonto) < 1 Then
                       'Si el Abono es Mayor Igual que la Cuota, realiza Abono Ordinario
                       'y avanza la fecha de pago
                       
                       vCaso.Inconsistencia = 4
                       vCaso.abIntC = CCur(Format(((vCaso.Saldo * vCaso.Interes) / 1200), "Standard"))
                       vCaso.abAmortiza = vCaso.Cuota - vCaso.abIntC
                       curMonto = curMonto - (vCaso.abAmortiza + vCaso.abIntC)
                    
                    Else
                       'Si el Abono es Menor que la Cuota, entonces lo Amortiza todo
                       'Abono ExtraOrdinario
                       vCaso.Inconsistencia = 5
                       vCaso.abAmortiza = curMonto
                       curMonto = 0
                    End If
               
            Else
              'Proceso Anterior / Cuotas Normales
                    If curMonto >= (vCaso.Saldo * vCaso.Interes) / 1200 Then
                       vCaso.abIntC = CCur(Format(((vCaso.Saldo * vCaso.Interes) / 1200), "Standard"))
                       curMonto = curMonto - vCaso.abIntC
                    Else
                       vCaso.abIntC = curMonto
                       curMonto = 0
                       vCaso.MorFechaProceso = GLOBALES.glngFechaCR
                       vCaso.MorIntCor = ((vCaso.Saldo * vCaso.Interes) / 1200) - vCaso.abIntC
                       vCaso.Inconsistencia = 1
                    End If
                
                    If curMonto >= vCaso.Cuota - ((vCaso.Saldo * vCaso.Interes) / 1200) Then
                       vCaso.abAmortiza = vCaso.Cuota - ((vCaso.Saldo * vCaso.Interes) / 1200)
                       curMonto = curMonto - vCaso.abAmortiza
                    Else
                       vCaso.abAmortiza = curMonto
                       curMonto = 0
                       vCaso.MorFechaProceso = GLOBALES.glngFechaCR
                       vCaso.MorAmortiza = (vCaso.Cuota - ((vCaso.Saldo * vCaso.Interes) / 1200)) - vCaso.abAmortiza
                       vCaso.Inconsistencia = 1
                    End If
            End If 'Adelanto o Abono Normal
            
            vCaso.Abono = vCaso.abAmortiza + vCaso.abIntC + vCaso.abIntM
            
            Call sbCrGuardaDesgloce             'Insertar Registro Aqui

         End If 'CurMonto > 0
      
         rsTmp.MoveNext
    
    Loop
    rsTmp.Close
    
'***********************************************************
'PASO 3 : Sobrantes e Inconsistencias
'***********************************************************
    
    'Si queda un Monto, ver Refundiciones y Aplicar y si Aun queda, Enviar con Inconsistencia
    'De devolucion
    
    If curMonto > 0 Then
        strSQL = "select R.id_solicitud,R.codigo,R.opex,R.fechaforp" _
               & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
               & " where R.cedula = '" & Trim(rs!Cedula) & "' and R.prideduc > " & GLOBALES.glngFechaCR _
               & " and R.proceso = 'N' and R.estado = 'A' and R.saldo > 0 " _
               & " and C.retencion = 'N' and C.poliza = 'N' order by C.prioridad asc"
        rsTmp.Open strSQL, glogon.Conection, adOpenStatic
        'inicializa
        Do While Not rsTmp.EOF
            vCaso.Codigo = rsTmp!Codigo
            vCaso.Id_solicitud = rsTmp!Id_solicitud
            vCaso.Opex = rsTmp!Opex
            vCaso.Retencion = False
            vCaso.Inconsistencia = 2 'Refundicion
            vCaso.IND_APL = 0
            
            'Aplicar al Ordinario
            vCaso.Tipo = "C"
            vCaso.ID_Referencia = rsTmp!Id_solicitud
            vCaso.abAmortiza = 0
            vCaso.abIntC = 0
            vCaso.abIntM = 0
            vCaso.Abono = 0
            vCaso.Cuota = 0
            
            vCaso.MorIntMor = 0
            vCaso.MorIntCor = 0
            vCaso.MorAmortiza = 0
            vCaso.MorFechaProceso = 0
            
            curMonto = fxCrDegloceRefundicion(rsTmp!Id_solicitud, curMonto, Day(rsTmp!fechaforp))
            
          rsTmp.MoveNext
        Loop
        rsTmp.Close
    End If 'Refundiciones
    
    
    If curMonto > 0 Then
      vCaso.Inconsistencia = 3 'Devolucion
      vCaso.Abono = curMonto
      vCaso.Codigo = ""
      vCaso.Id_solicitud = 0
      vCaso.Opex = 0
      vCaso.Retencion = False
      vCaso.IND_APL = 0
    
      'Aplicar al Ordinario
      vCaso.Tipo = "C"
      vCaso.ID_Referencia = 0
      vCaso.abAmortiza = 0
      vCaso.abIntC = 0
      vCaso.abIntM = 0
      vCaso.Cuota = 0
    
      vCaso.MorIntMor = 0
      vCaso.MorIntCor = 0
      vCaso.MorAmortiza = 0
      vCaso.MorFechaProceso = 0
      
      vCaso.ORG_ID_Sol = 0
      vCaso.ORG_Codigo = ""
      vCaso.Saldo = 0
      vCaso.Interes = 0
      vCaso.AmortizaOtro = "S"
      
      Call sbCrGuardaDesgloce
      
    End If
    
  rs.MoveNext
    
  If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
  lblStatus.Caption = "Cargando..Registro # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
  lblStatus.Refresh
Loop

lblStatus.Caption = ""
Me.MousePointer = vbDefault
prgProcesoMensual.Value = 1

glogon.Conection.Execute "update par_ahcr set cr_des = 1"

Call Bitacora("Aplica", "PRM-CREDITO Detalla las Deducciones")

Call sbReporteDetalleDeducciones

Call EstadoActualProceso


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub


Private Sub sbReporteCargado(vFecha As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

With frmContenedor.crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "PROCESO MENSUAL - CARGADO DE INFORMACION"
     
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(vFecha, "####-##") & "'"
    .Formulas(3) = "usuario='" & glogon.Usuario & "'"
    .Formulas(4) = "institucion='Caja Costarricense Seguro Social'"
    .ReportFileName = GLOBALES.gReportes & "\credito\reportes\PRMCargados.rpt"
    
    .SelectionFormula = "{PRM_CARGADO.FECHA_PROCESO} = " & vFecha _
              & " AND {PRM_CARGADO.COD_INSTITUCION} = 1"
    
    .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 MsgBox Err.Description, vbCritical
 
End Sub



Private Sub sbCargaDeduccionesNew()
Dim fn, strCadena As String, lng As Long
Dim strMonto As String, strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
fn = FreeFile

prgProcesoMensual.Min = 1
prgProcesoMensual.Max = 2

lblStatus = "Limpiando y Actualizando ..."
lblStatus.Refresh

strSQL = "delete from prm_creditos where fecha_proceso=" & GLOBALES.glngFechaCR
glogon.Conection.Execute strSQL

With frmContenedor.dlg
 .DialogTitle = "Localice archivo de deducciones de abonos a cartera"
 .Filter = "34586."
 .InitDir = "C:\"
 .ShowOpen
End With

If frmContenedor.dlg.FileName = "" Then
 MsgBox "Seleccione el Archivo de Deducciones del Proceso " & Format(GLOBALES.glngFechaCR, "####-##"), vbInformation
 Exit Sub
End If

Me.MousePointer = vbHourglass

Open frmContenedor.dlg.FileName For Input As #fn   ' Lee el archivo.
 Do While Not EOF(fn)
   Input #fn, strCadena
   prgProcesoMensual.Max = prgProcesoMensual.Max + 1
 Loop
Close #fn

lblStatus = "Cargando ..."
lblStatus.Refresh

mFechaSistema = fxFechaServidor

Open frmContenedor.dlg.FileName For Input As #fn   'Lee el Archivo y lo compara
Do While Not EOF(fn)
    Input #fn, strCadena
    
    If Trim(strCadena) <> "" Then
        strMonto = Format(Mid(strCadena, 28, 13), "###########")
        strMonto = LTrim(RTrim(strMonto))
        If Len(strMonto) > 2 Then
         strMonto = Mid(strMonto, 1, Len(strMonto) - 2) & "." & Mid(strMonto, Len(strMonto) - 1, Len(strMonto))
        Else
         strMonto = "0" & "." & strMonto
        End If
        strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto) values(1" _
               & ",1," & GLOBALES.glngFechaCR & ",3,'" & Trim(Format(Mid(strCadena, 1, 11), "###########")) _
               & "'," & CCur(strMonto) & ")"
        glogon.Conection.Execute strSQL
    End If
      
    If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
    lblStatus.Caption = "Cargando..Registro # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
    lblStatus.Refresh
Loop

lblStatus.Caption = ""
Me.MousePointer = vbDefault
prgProcesoMensual.Value = 1

Close #fn

lblStatus.Caption = "Estado..."

glogon.Conection.Execute "update par_ahcr set cr_car = 1"

Call Bitacora("Aplica", "PRM-CREDITO Carga Deducciones")

Call sbReporteCargado(GLOBALES.glngFechaCR)

Call EstadoActualProceso

MsgBox "Información Cargada ...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub



Sub GeneraDeducciones()
Dim rs As New ADODB.Recordset, strSQL As String
Dim iRespuesta As Integer, lngFechaAnterior As Long

On Error GoTo CapturaError

'Actualiza Estado de los creditos y corrige inconsistencias menores
lblStatus.Caption = "Corrigiendo Inconsistencias Menores..."
lblStatus.Refresh
Call sbActualizaSaldos

'CALCULA EL CAMPO SALDO MES PARA TODOS LOS REGISTROS ACTIVOS CON PROCESO NORMAL Y SALDO >0
lblStatus = "Calculando Saldo del Mes..."
lblStatus.Refresh

With rs
  .Source = "select coalesce(cr_saldo,0) as cr_saldo from par_ahcr"
  .Open , glogon.Conection, adOpenStatic
  If !cr_saldo = 0 Then
    .Close
    Call CalculaSaldoMes
    glogon.Conection.Execute "update par_ahcr set cr_saldo = 1"
  Else
     iRespuesta = MsgBox("El Cálculo del Saldo Mensual, ya Fue Realizado, Desea Volverlo a Cálcular", vbYesNo)
     If iRespuesta = vbYes Then
        .Close
        Call CalculaSaldoMes
     Else
        .Close
     End If
  End If
End With

lblStatus = "Limpiando Información Anterior..."
lblStatus.Refresh

'Conserva Tres Meses
lngFechaAnterior = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
lngFechaAnterior = fxFechaProcesoAnterior(lngFechaAnterior)

strSQL = "delete from CuotasEnviadas where fecpro<=" & lngFechaAnterior _
       & " or fecpro = " & GLOBALES.glngFechaCR
glogon.Conection.Execute strSQL

Call EnviaCuotasNormales
Call EnviaMorosidadMasVieja
   
'Crea Archivo de Texto Aqui
Call imgGeneraArchivo_Click

Call Bitacora("Aplica", "PRM-CREDITO Genera Deducciones")

glogon.Conection.Execute "update par_ahcr set cr_gen = 1"
Call EstadoActualProceso

lblStatus.Caption = "Generando reporte de cuotas a enviarse"
 
'Call ReporteGeneracion
Call ReporteGeneracionUNI
lblStatus.Caption = "Estado..."
 
Exit Sub
CapturaError:
  MsgBox Err.Description, vbCritical
End Sub

Sub EstadoActualProceso()
Dim rs As New ADODB.Recordset

rs.Open "select * from par_ahcr", glogon.Conection, adOpenStatic
    
  tlbPrincipal.Buttons.Item(1).Enabled = True
  tlbPrincipal.Buttons.Item(2).Enabled = True
  tlbPrincipal.Buttons.Item(3).Enabled = True
  
  cboMes.Text = fxConvierteMES(Val(Mid(GLOBALES.glngFechaCR, 5, 2)))
  txtAno.Text = Mid(GLOBALES.glngFechaCR, 1, 4)

If rs.EOF And rs.BOF Then
  MsgBox "NO EXISTEN PARAMETROS DEL PROCESO - !! DEBE CREARLOS ANTES DE ENTRAR AQUI !! -"
  rs.Close
  Unload Me
Else

'  optProcesoMensual - Apunta al siguiente

   ImgFechaProceso.Visible = False
   imgGenera.Visible = False
   imgCarga.Visible = False
   imgDesgloza.Visible = False
   imgAplica.Visible = False
   imgListainconsistencias.Visible = False
   imgRecalculaSaldosMorosos.Visible = False
   imgGeneraArchivo.Visible = False
    
    If rs!cr_gen = 1 Then
       optProcesoMensual(2).Value = True
       imgGenera.Visible = True
       ImgFechaProceso.Visible = True
       imgGeneraArchivo.Visible = True
    End If
    
    If rs!cr_car = 1 Then
       optProcesoMensual(3).Value = True
       imgCarga.Visible = True
    End If
    
    If rs!cr_des = 1 Then
       optProcesoMensual(8).Value = True
       imgDesgloza.Visible = True
    End If
    
    
    If rs!cr_apl = 1 Then
       optProcesoMensual(4).Value = True
       imgAplica.Visible = True
    End If
    
    If rs!cr_incon = 1 Then
       optProcesoMensual(5).Value = True
       imgListainconsistencias.Visible = True
    End If
    
    If rs!cr_rec = 1 Then
       optProcesoMensual(0).Value = True
       imgRecalculaSaldosMorosos.Visible = True
       ImgFechaProceso.Visible = False
    End If

  rs.Close
  
  Call RefrescaTags(Me)

End If

End Sub

Function fxValidaPaso() As Boolean
fxValidaPaso = True

Select Case True
  Case optProcesoMensual(0).Value 'Fecha Proceso Mensual
   If imgRecalculaSaldosMorosos.Visible = False Then
     fxValidaPaso = False
     MsgBox "No Procede.. No ha seguido todo el orden - Se procesa solo Reporteria de Existir "
   End If
  
  Case optProcesoMensual(1).Value 'Genera Deducciones
   If imgGenera.Visible = True Then
     fxValidaPaso = False
     MsgBox "No Procede.. No ha seguido todo el orden - Se procesa solo Reporteria de Existir "
   End If
  
  Case optProcesoMensual(2).Value 'Carga Deducciones
   If imgCarga.Visible = True Or imgGenera.Visible = False Then
     fxValidaPaso = False
     MsgBox "No Procede.. No ha seguido todo el orden - Se procesa solo Reporteria de Existir "
   End If
  
  Case optProcesoMensual(8).Value 'Detalla Deducciones
   If imgAplica.Visible = True Or imgCarga.Visible = False Then
     fxValidaPaso = False
     MsgBox "No Procede.. No ha seguido todo el orden - Se procesa solo Reporteria de Existir "
   End If
   
  Case optProcesoMensual(3).Value 'Aplica
   If imgAplica.Visible = True Or imgDesgloza.Visible = False Then
     fxValidaPaso = False
     MsgBox "No Procede.. No ha seguido todo el orden - Se procesa solo Reporteria de Existir "
   End If
  
  Case optProcesoMensual(4).Value 'Lista de Inconsistencias
   If imgAplica.Visible = False Or imgListainconsistencias.Visible = True Then
     fxValidaPaso = False
     MsgBox "No Procede.. No ha seguido todo el orden - Se procesa solo Reporteria de Existir "
   End If
  
  Case optProcesoMensual(5).Value 'Recalcula Cuotas a Morosos
   If imgAplica.Visible = False Or imgRecalculaSaldosMorosos.Visible = True Then
     fxValidaPaso = False
     MsgBox "No Procede.. No ha seguido todo el orden - Se procesa solo Reporteria de Existir "
   End If
End Select
    
End Function


Private Sub Form_Load()
vModulo = 3
glogon.Conection.CommandTimeout = 360
Call Formularios(Me)
Call EstadoActualProceso

End Sub

Sub FechaProcesoSiguiente()
Dim Mes As Integer

Mes = fxConvierteMES(cboMes.Text)

If Mes = 12 Then
 txtAno.Text = Val(txtAno) + 1
 Mes = 0
End If
 Mes = Mes + 1
 cboMes.Text = fxConvierteMES(Mes)
End Sub

Private Sub imgAplica_Click()
Call ReporteAplicacion
End Sub

Private Sub imgCambiaFecha_Click()
Dim strSQL As String, FechaProceso As String
Dim iMes As Integer, lngAnio As Long, vFecha As Date

On Error GoTo CapturaError

If fraFechaProceso.Enabled = False Then
   Exit Sub
End If

imgCambiaFecha.BorderStyle = 1

'Cambia la Fecha de Calculo
iMes = fxConvierteMES(cboMes.Text)
lngAnio = txtAno

FechaProceso = lngAnio & Format(iMes, "00")

If iMes = 12 Then
   iMes = 1
   lngAnio = lngAnio + 1
Else
   iMes = iMes + 1
End If

vFecha = lngAnio & "/" & Format(iMes, "00") & "/01"
vFecha = DateAdd("d", -1, vFecha)

strSQL = "update par_ahcr set cr_fec = " & FechaProceso _
       & ",cr_gen = 0, cr_car = 0, cr_apl = 0, cr_incon = 0,cr_saldo =0, cr_rec = 0" _
       & ",cr_fecha_calculo = '" & Format(vFecha, "yyyy/mm/dd") & "'"
glogon.Conection.Execute strSQL

Call Bitacora("Aplica", "PRM-CREDITO Cambia Fecha Proceso")

Call CargaParametros 'Vuelve a cargar los parametros iniciales GENERALES

Call EstadoActualProceso

MsgBox "La fecha de proceso fue cambiada a : " & GLOBALES.glngFechaCR
imgCambiaFecha.BorderStyle = 0
fraFechaProceso.Enabled = False

Exit Sub

CapturaError:
 Call ProcedimientoErrores(Me.Name, Err)

End Sub


Private Sub imgCarga_Click()
Call sbReporteCargado(GLOBALES.glngFechaCR)
End Sub

Private Sub imgDesgloza_Click()
Call sbReporteDetalleDeducciones
End Sub

Private Sub imgGenera_Click()
Call ReporteGeneracionUNI
End Sub

Private Sub imgGeneraArchivo_Click()
Dim rs As New ADODB.Recordset, strFproAnt As String
Dim strSQL As String, strRuta As String
Dim strOld As String, strNue As String, i As Integer
Dim fnFile, iRespuesta As Integer, strCadena As String
Dim strfile As String, strCodigo As String

On Error GoTo CapturaError

fnFile = FreeFile
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

strfile = ""
strfile = Dir("C:\CCSS", vbDirectory)
If Not (UCase(strfile) = "CCSS") Then 'El directorio no existe
   MkDir "C:\CCSS"
End If

strfile = ""
strfile = Dir("C:\CCSS\ABONOS", vbDirectory)
If Not (UCase(strfile) = "ABONOS") Then 'El directorio no existe
   MkDir "C:\CCSS\ABONOS"
End If

strRuta = "C:\CCSS\ABONOS\" & GLOBALES.glngFechaCR

strOld = strRuta & "\ARC34586.TXT"
strfile = Dir(strOld, vbArchive)
If strfile = "ARC34586.TXT" Then 'El archivo existe
 Kill "c:\CCSS\ABONOS\" & GLOBALES.glngFechaCR & "\" & "ARC34586.txt"
End If

strfile = ""
strfile = Dir(strRuta, vbDirectory)
If Not (strfile = CStr(GLOBALES.glngFechaCR)) Then  'El directorio no existe
  MkDir strRuta
End If

Open strOld For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
lblStatus.Refresh

strSQL = "select * from cuotasenviadas where fecpro = " & GLOBALES.glngFechaCR _
       & " order by cedula"
       
'Cuota Unificada 2003/03/20
strSQL = "select * from prm_planillas where fecPro = " & GLOBALES.glngFechaCR _
       & " order by cedula"
rs.Open strSQL, glogon.Conection, adOpenStatic

prgProcesoMensual.Max = rs.RecordCount
prgProcesoMensual.Value = 1

i = 0

Do While Not rs.EOF
 strCodigo = rs!Codigo

 Do While Len(strCodigo) < 10
   strCodigo = " " & strCodigo
 Loop
 
 strCadena = Format(Mid(Trim(rs!Cedula), 1, 11), "00000000000")
 strCadena = strCadena & "9999062" & Format((rs!Cuota * 100), "0000000000000") & "0000000099999999"
 strCadena = strCadena & strCodigo & "00000000000000000000000000100000000"
 
 Print #fnFile, strCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 
 If i = 21 Then
  i = 1
  lblStatus.Refresh
 Else
  i = i + 1
 End If

 rs.MoveNext
Loop
rs.Close
Close #fnFile
 
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & strOld, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

CapturaError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical

End Sub

Private Sub imgListaInconsistencias_Click()
  Call ReporteInconsistencias(GLOBALES.glngFechaCR)
End Sub

Private Sub imgRepInconsistencias_Click()
Dim vFecha As Long

On Error GoTo vError
vFecha = CLng(InputBox("Especifique la fecha de proceso " & vbCrLf _
        & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Inconsistencias"))
Call ReporteInconsistencias(vFecha)

Exit Sub

vError:
 MsgBox Err.Description, vbCritical
 
End Sub


Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim vFecha As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DESCRIPCION
'Funcion principal que controla los botones del toolbar y ademas llama a las funciones
'que ejecutan cada uno de los procesos
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  Select Case Button.Key
    Case "ejecutar"
       Select Case True
         Case optProcesoMensual(0).Value 'Fecha Proceso Mensual
          If fxValidaPaso Then
            fraFechaProceso.Enabled = True
            imgCambiaFecha.Enabled = True
            Call FechaProcesoSiguiente
          End If
         Case optProcesoMensual(1).Value 'Genera Deducciones
          If fxValidaPaso Then
            Call GeneraDeducciones
          Else
            Call ReporteGeneracion
          End If
          
         Case optProcesoMensual(2).Value 'Carga Deducciones
          If fxValidaPaso Then
           Call sbCargaDeduccionesNew
          Else
           Call sbReporteCargado(GLOBALES.glngFechaCR)
          End If
          
         Case optProcesoMensual(8).Value 'Detalla las Deducciones
          If fxValidaPaso Then
           Call sbDesgloceUnificacionNew
          Else
           Call sbReporteDetalleDeducciones
          End If
          
          
         Case optProcesoMensual(3).Value 'Aplica
          If fxValidaPaso Then
            Call AplicaAbonos
          Else
            Call ReporteAplicacion
          End If
          
         Case optProcesoMensual(4).Value 'Lista de Inconsistencias
          If fxValidaPaso Then
            Call Bitacora("Aplica", "PRM-CREDITO Reporte Inconsistencias")
            glogon.Conection.Execute "update par_ahcr set cr_incon = 1"
            Call ReporteInconsistencias(GLOBALES.glngFechaCR)
            Call EstadoActualProceso
          Else
            Call ReporteInconsistencias(GLOBALES.glngFechaCR)
          End If
         
         Case optProcesoMensual(5).Value 'Recalculo de Cuotas Morosas
          If fxValidaPaso Then
            Call RecalculaCuotaEnMora
          End If
       
         Case optProcesoMensual(6).Value 'Recalculo Saldo del Mes es Opcional
           Me.MousePointer = vbHourglass
           Call CalculaSaldoMes
           Me.MousePointer = vbDefault
       
         Case optProcesoMensual(7).Value 'Traslados al Fondo de Ahorros Opcional
           
            vFecha = CLng(InputBox("Especifique la fecha de proceso " & vbCrLf _
                   & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Traslados al Fondo"))
                   Call sbRepFondo(vFecha)
       End Select
    Case "cerrar"
       Unload Me
  End Select

End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
 Call Valida(KeyAscii)
End Sub
