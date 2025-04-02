VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCC_PE_PlanillaDirecta 
   Caption         =   "Proyectos Especialiados: Planilla Directa"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   10848
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   10848
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBarX 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   20
      Top             =   7545
      Width           =   10845
      _ExtentX        =   19135
      _ExtentY        =   233
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.ComboBox cboPlanillaNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtNumDoc 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtNoExisten 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox txtCasos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Monto"
      Top             =   7200
      Width           =   1575
   End
   Begin VB.ComboBox cboProceso 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtArchivo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   6855
   End
   Begin VB.ComboBox cboInstitucion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
   Begin MSComctlLib.Toolbar tlbX 
      Height          =   528
      Left            =   9480
      TabIndex        =   3
      Top             =   960
      Width           =   456
      _ExtentX        =   804
      _ExtentY        =   931
      ButtonWidth     =   487
      ButtonHeight    =   466
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "buscar"
            Object.ToolTipText     =   "Buscar archivos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cargar"
            Object.ToolTipText     =   "Cargar información"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10080
      Top             =   240
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_PE_PlanillaDirecta.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_PE_PlanillaDirecta.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_PE_PlanillaDirecta.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_PE_PlanillaDirecta.frx":13926
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_PE_PlanillaDirecta.frx":1A188
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbProceso 
      Height          =   336
      Left            =   6600
      TabIndex        =   7
      Top             =   7080
      Width           =   3804
      _ExtentX        =   6710
      _ExtentY        =   593
      ButtonWidth     =   1778
      ButtonHeight    =   550
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar"
            Key             =   "Aplicar"
            Object.ToolTipText     =   "Procesar Información"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bitácora"
            Key             =   "Bitacora"
            Object.ToolTipText     =   "Ver Bitácora de Aplicacones"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Operación"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3975
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   10455
      _Version        =   524288
      _ExtentX        =   18441
      _ExtentY        =   7011
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   13
      ScrollBars      =   2
      SpreadDesigner  =   "frmCC_PE_PlanillaDirecta.frx":1A2A3
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label2 
      Caption         =   "No. Planilla"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   19
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "No. Doc. Apl."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   16
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "EDICION UNICA PARA SAVA!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   252
      Left            =   2520
      TabIndex        =   15
      Top             =   120
      Width           =   5172
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No existen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   14
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Casos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   13
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   7200
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   10680
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label2 
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1200
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   10
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   1200
      TabIndex        =   9
      Top             =   360
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmCC_PE_PlanillaDirecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mInstitucion As Long, mOperadora As Long, mPlan As String
Dim mCodigoDeduc As String, mCuentaCxC As String, mContrato As Long
Dim lCodigoDeduc As String, vPaso As Boolean

Private Sub sbLimpia()
If vPaso Then Exit Sub

    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtCasos.Text = 0
    txtNoExisten.Text = 0

    txtArchivo.Text = ""
    
    
    txtNumDoc.Text = Format(cboProceso.Text, "####-##") & "." & Format(cboInstitucion.ItemData(cboInstitucion.ListIndex), "000") & ".PD." & cboPlanillaNo.Text
    
End Sub


Private Sub cboPlanillaNo_Click()
 Call sbLimpia
End Sub


Private Sub cboInstitucion_Click()
 Call sbLimpia
 
 mInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
 Call sbDocumentosInstitucion
 
End Sub


Private Sub cboProceso_Click()
 Call sbLimpia
End Sub

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()
Dim strSQL As String, i As Integer
Dim vProceso As Currency

vModulo = 1

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

txtArchivo.Text = ""

vGrid.MaxCols = 13
vGrid.MaxRows = 0

vPaso = True

strSQL = "select cod_institucion as IdX,descripcion as ItmX from instituciones where activa = 1"
Call sbLlenaCbo(cboInstitucion, strSQL, False, True)

vProceso = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
vProceso = fxFechaProcesoAnterior(vProceso)
cboProceso.AddItem vProceso

For i = 1 To 6
  vProceso = fxFechaProcesoSiguiente(vProceso)
  cboProceso.AddItem vProceso
Next i
cboProceso.Text = GLOBALES.glngFechaCR

cboPlanillaNo.AddItem "1"
cboPlanillaNo.AddItem "2"
cboPlanillaNo.AddItem "3"
cboPlanillaNo.Text = "1"

vPaso = False

Call sbLimpia

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbCargaDeducciones(vTipo As Integer)
Dim strCadena As String, curMonto As Currency
Dim fn As Long, lCasos As Long
Dim strMonto  As String
Dim strCedula As String
Dim strNombre As String
Dim i As Integer

On Error GoTo vError
vGrid.MaxRows = 0

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

If cboInstitucion.ListCount <= 0 Then
    MsgBox "No existe ninguna Institución, no se puede procesar el archivo...", vbCritical
    Exit Sub
End If


Me.MousePointer = vbHourglass

vGrid.MaxRows = 0
curMonto = 0
lCasos = 0 'Total

Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset
Dim vOperacion As Long, vSaldo As Currency

 txtNoExisten.Text = 0
 txtCasos.Text = 0
 txtMonto.Text = 0
 
 Set rsExcel = Excel_Load(txtArchivo.Text, "Import")

 
 With rsExcel
 
     Do While Not .EOF
     
      
     
       If Trim(!Cliente) <> "" Then
             vGrid.MaxRows = vGrid.MaxRows + 1
             vGrid.Row = vGrid.MaxRows
             vGrid.col = 1 'Operacion Origen
             vGrid.Text = !Operacion
             
             If fxNombre(!Cliente) = "" Then
                 txtNoExisten.Text = txtNoExisten + 1
                 vGrid.CellTag = "-1"
             Else
                'Busqueda de los datos de la Operacion
                strSQL = "select R.id_Solicitud, R.Codigo, R.Saldo , R.Cedula, S.Nombre" _
                       & " from reg_Creditos R inner join Socios S on R.cedula = S.cedula" _
                       & " where R.cedula = '" & !Cliente & "' and R.Estado = 'A'" _
                       & " and S.cod_Institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
                       & " and R.documento_referido = '" & !Operacion & "'"
                
                Call OpenRecordSet(rs, strSQL)
                If rs.EOF And rs.BOF Then
                    txtNoExisten.Text = txtNoExisten + 1
                    vGrid.CellTag = "-1"
                
                Else
                    vGrid.col = 2  'Operacion SIF
                    vGrid.Text = CStr(rs!id_solicitud)
                    
                    vGrid.col = 12 'Línea SIF
                    vGrid.Text = rs!Codigo
                    vGrid.col = 13 'Saldo SIF
                    vGrid.Text = Format(rs!Saldo - !Principal, "Standard")
                End If
                rs.Close
             End If
             
             
             vGrid.col = 3 'Cédula del Cliente
             vGrid.Text = !Cliente
             vGrid.col = 4 'Nombre del Cliente
             vGrid.Text = !Nombre
             vGrid.col = 5 'Fecha Pago Real
             vGrid.Text = !fecha
             vGrid.col = 6 'No. Recibo
             vGrid.Text = !Recibo
             vGrid.col = 7 'Abono
             vGrid.Text = Format(!Abono, "Standard")
             vGrid.col = 8 'Int. Cor.
             vGrid.Text = Format(!Int_Cor, "Standard")
             vGrid.col = 9 'Int. Mor.
             vGrid.Text = Format(!Int_Mor, "Standard")
             vGrid.col = 10 'Principal
             vGrid.Text = Format(!Principal, "Standard")
             vGrid.col = 11 'Saldo Referencia
             vGrid.Text = Format(!Saldo, "Standard")
             
             

             
             
             
             curMonto = curMonto + !Abono
             txtCasos = txtCasos + 1
             txtCasos.Refresh
        End If
       .MoveNext
     Loop
     .Close
 End With


'Totales
txtMonto.Text = Format(curMonto, "Standard")

Me.MousePointer = vbDefault

MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia
End Sub

Private Sub sbProcesar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lng As Long, vCodigo As String
Dim vFecha As Date, vProceso As Long
Dim vProcesados As Long
Dim vTipoDoc As String, vNumDoc As String, vConcepto As String


Dim pCedula As String, pNombre As String, pMovimiento As String, pMonto As Currency
Dim pIntCor As Currency, pIntMor As Currency, pPrincipal As Currency
Dim pSaldo As Currency, pOperacion As Integer, pOrigenp As String, pOrigenSaldo As Currency, pFechaPago As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor
vProceso = cboProceso.Text

vProcesados = 0

vNumDoc = txtNumDoc.Text
vConcepto = "CRD007"
vTipoDoc = "PLA"
 
'Verifica que no se haya aplicado anteriormente
strSQL = "select count(*) as 'Existe' from sif_Transacciones where Tipo_Documento = '" & vTipoDoc _
       & "' and cod_Transaccion = '" & vNumDoc & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
  Me.MousePointer = vbDefault
  MsgBox "Esta Planilla Especial ya fue aplicada anteriormente...Verifique!", vbExclamation
  Exit Sub
End If
rs.Close


With vGrid
    ProgressBarX.Max = .MaxRows + 1
    For lng = 1 To .MaxRows
    
       .Row = lng
       .col = 1
       strSQL = "insert PRM_DIR_MODO_01(LINEA, NUM_DOC, COD_INSTITUCION, OPERACION_ORIGEN, CEDULA, NOMBRE, ID_SOLICITUD, CODIGO" _
              & " , FECHA_PAGO_REAL, FECHA_APLICACION, NUM_COMPROBANTE, ABONO, INT_COR, INT_MOR, CARGOS" _
              & " , PRINCIPAL, SALDO_REFERENCIA, APLICADO, APLICADO_ID_SEQ, INCONSISTENCIA)" _
              & " VALUES(" & lng & ",'" & vNumDoc & "'," & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ",'" & .Text _
              & "','"
       .col = 3 'Cedula
       strSQL = strSQL & .Text & "','"
       .col = 4 'Nombre
       strSQL = strSQL & .Text & "',"
              
       .col = 2 'Operacion SIF
       If .Text = "" Then
          strSQL = strSQL & "0,'"
       Else
          strSQL = strSQL & .Text & ",'"
       End If
       .col = 12 'Codigo
       strSQL = strSQL & .Text & "','"

       .col = 5 'Fecha de Referencia (Pago)
       strSQL = strSQL & Format(.Text, "yyyy/mm/dd") & "',Null,'"
       
       .col = 6 'Documento de Referencia (Pago)
       strSQL = strSQL & .Text & "',"
       
       .col = 7 'Abono
       strSQL = strSQL & CCur(.Text) & ","
       
       .col = 8 'Int Cor
       strSQL = strSQL & CCur(.Text) & ","
       
       .col = 9 'Int Mor
       strSQL = strSQL & CCur(.Text) & ",0,"
       
       .col = 10 'Principal
       strSQL = strSQL & CCur(.Text) & ","
       
       .col = 11 'Saldo de Referencia en el Origen
       strSQL = strSQL & CCur(.Text) & ",0,0,'"


        .col = 1 'Verifica Inconsistencia
        If .CellTag = "-1" Then
            strSQL = strSQL & "No Existe Relacion')"
        Else
            strSQL = strSQL & "')"
        End If
    
     Call ConectionExecute(strSQL)

    ProgressBarX.Value = lng
    Next lng
End With


'Procesa Casos
strSQL = "select * from PRM_DIR_MODO_01" _
       & " where NUM_DOC = '" & vNumDoc & "' and isnull(Aplicado,0) = 0" _
       & " and inconsistencia = ''" _
       & " order by OPERACION_ORIGEN, FECHA_PAGO_REAL, NUM_COMPROBANTE"
Call OpenRecordSet(rs, strSQL)
ProgressBarX.Value = 1
ProgressBarX.Max = rs.RecordCount + 2

Do While Not rs.EOF
  
'create proc spCRDPlanillaDirectaModo1(@Linea int, @Operacion int, @Proceso int, @Concepto varchar(10), @Usuario varchar(30)
'        , @TipoCom varchar(10), @NumCom varchar(30), @Abono dec(16,2), @IntCor dec(16,2), @IntMor dec(16,2)
'        , @Principal dec(16,2), @Cargo dec(16,2)
'        , @SaldoRef dec(16,2), @NReciboRef varchar(30), @FechaPagoRef datetime, @OperacionRef varchar(30)
'        , @FechaChr datetime, @Caja varchar(10) = '', @ReCalculaCta smallint = 0, @Actualiza smallint = 1)
'
  strSQL = "exec spCRDPlanillaDirectaModo1 " & rs!Linea & "," & rs!id_solicitud & "," & vProceso & ",'CRD007','" & glogon.Usuario _
         & "','" & vTipoDoc & "','" & vNumDoc & "'," & rs!Abono & "," & rs!Int_Cor & "," & rs!Int_Mor _
         & "," & rs!Principal & ",0," _
         & rs!Saldo_Referencia & ",'" & rs!Num_Comprobante & "','" & Format(rs!Fecha_PAgo_Real, "yyyy/mm/dd") & "','" & rs!Operacion_Origen _
         & "','" & Format(vFecha, "yyyy/mm/dd") & "','',0,1"
  Call ConectionExecute(strSQL)
    
  ProgressBarX.Value = ProgressBarX.Value + 1
  rs.MoveNext
Loop
rs.Close

'Realiza el Asiento + Comprobante de Aplicacion
''create proc spCRDPlanillaDirectaModo1Asiento(@TipoDoc varchar(10), @NumDoc varchar(30), @Fecha datetime, @Usuario varchar(30), @Institucion int, @Proceso int)
strSQL = "exec spCRDPlanillaDirectaModo1Asiento '" & vTipoDoc & "','" & vNumDoc & "','" & Format(vFecha, "yyyy/mm/dd") _
       & "','" & glogon.Usuario & "'," & cboInstitucion.ItemData(cboInstitucion.ListIndex) & "," & vProceso
Call ConectionExecute(strSQL)
  
  
Me.MousePointer = vbDefault
MsgBox "Proceso Aplicado Satisfactoriamente... Registros Procesados :" & vGrid.MaxRows

Call sbLimpia

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia


End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - 600
vGrid.Height = Me.Height - (vGrid.top + 1500)

imgBanner.Width = Me.Width

tlbProceso.top = vGrid.top + vGrid.Height + 350
txtCasos.top = tlbProceso.top
txtMonto.top = tlbProceso.top
txtNoExisten.top = tlbProceso.top

Label2.Item(0).top = tlbProceso.top
Label2.Item(1).top = tlbProceso.top - Label2.Item(1).Height + 20
Label2.Item(2).top = Label2.Item(1).top
End Sub

Private Sub tlbProceso_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Aplicar"
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen deducciones cargadas...[verifique!]", vbExclamation
       Exit Sub
    End If
   
    Call sbProcesar
  
  Case "cancelar"
    vGrid.MaxRows = 0
    txtArchivo.Text = ""
  Case "Bitacora"
'    frmFNDPlanillaBitacora.Show
End Select

End Sub

Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)
        
Select Case Button.Key
  
  Case "buscar"
        txtArchivo.Text = ""
       Call sbBuscaArchivo(1)

  
  Case "cargar"
       Call sbCargaDeducciones(1)

End Select

End Sub


Private Sub sbBuscaArchivo(vTipo As Integer)


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


Private Sub sbDocumentosInstitucion()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select cta_credito from instituciones where cod_institucion  = " & mInstitucion & ""
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    mCuentaCxC = rs!cta_credito
Else
    mCuentaCxC = ""
End If
rs.Close

End Sub


Public Sub sbBitacoraPlanilla(pTransaccion As String, pInstitucion As Long, pProceso As Long _
                , pGestion As String, pMonto As Currency, pPlan As String, Optional pDocumento As String = "")
'Dim strSQL As String, rs As New ADODB.Recordset
'
'
'strSQL = "select isnull(max(id_seq),0) + 1 as Consecutivo from fnd_prm_bitacora" _
'       & " where cod_institucion = " & pInstitucion & " and cod_plan  = '" & pPlan & "' and proceso = " & pProceso
'
'
'Call OpenRecordSet(rs, strSQL)
'    strSQL = "insert fnd_prm_bitacora(id_seq,cod_institucion,proceso,cod_plan,gestion,transaccion,documento,usuario,fecha,casos,monto) values(" _
'           & rs!consecutivo & "," & pInstitucion & "," & pProceso & ",'" & pPlan & "','" & pGestion & "','" & pTransaccion _
'           & "','" & pDocumento & "','" & glogon.Usuario & "',dbo.MyGetdate(),'" & txtCasos.Text & "'," & CCur(pMonto) & ")"
'rs.Close
'
'
'Call ConectionExecute(strSQL)

End Sub


Private Function fxAplicada() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "select isnull(count(id_seq),0) as Cantidad from CRD_OPERACION_TRANSAC" _
       & " where Tcon = 'PLA' and ncon = '" & txtNumDoc.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Cantidad > 0 Then
   fxAplicada = True
Else
   fxAplicada = False
End If
rs.Close

End Function

