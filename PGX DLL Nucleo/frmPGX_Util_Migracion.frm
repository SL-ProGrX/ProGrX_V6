VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmPGX_Util_Migracion 
   Caption         =   "Utilitario: Migrar Datos"
   ClientHeight    =   7764
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   11052
   LinkTopic       =   "Form1"
   ScaleHeight     =   7764
   ScaleWidth      =   11052
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optX 
      Caption         =   "Créditos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   5160
      TabIndex        =   15
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton optX 
      Caption         =   "Fondos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   3960
      TabIndex        =   14
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton optX 
      Caption         =   "Patrimonio"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2760
      TabIndex        =   13
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton optX 
      Caption         =   "Personas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   11
      Top             =   1080
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Monto"
      Top             =   7200
      Width           =   1575
   End
   Begin VB.TextBox txtCasos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Height          =   315
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox txtNoExisten 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Height          =   315
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   7200
      Width           =   975
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
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
   Begin VB.Data DaoControl 
      Caption         =   "DaoControl"
      Connect         =   "Excel 8.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   2100
   End
   Begin MSComDlg.CommonDialog Cmd 
      Left            =   10320
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10440
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
            Picture         =   "frmPGX_Util_Migracion.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_Util_Migracion.frx":6862
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_Util_Migracion.frx":D0C4
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_Util_Migracion.frx":13926
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_Util_Migracion.frx":1A188
            Key             =   "IMG5"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbX 
      Height          =   528
      Left            =   8520
      TabIndex        =   1
      Top             =   120
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
            ImageKey        =   "IMG1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cargar"
            Object.ToolTipText     =   "Cargar información"
            ImageKey        =   "IMG4"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBarX 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   7635
      Width           =   11055
      _ExtentX        =   19495
      _ExtentY        =   233
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.Toolbar tlbProceso 
      Height          =   336
      Left            =   8400
      TabIndex        =   4
      Top             =   7200
      Width           =   2484
      _ExtentX        =   4382
      _ExtentY        =   593
      ButtonWidth     =   1778
      ButtonHeight    =   550
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar"
            Key             =   "Aplicar"
            Object.ToolTipText     =   "Procesar Información en ProGrX"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Operación"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5295
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   10815
      _Version        =   524288
      _ExtentX        =   19076
      _ExtentY        =   9340
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
      MaxCols         =   12
      SpreadDesigner  =   "frmPGX_Util_Migracion.frx":1A2A3
      VScrollSpecialType=   2
      AppearanceStyle =   1
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
      Left            =   120
      TabIndex        =   10
      Top             =   7200
      Width           =   855
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
      Left            =   2520
      TabIndex        =   9
      Top             =   6960
      Width           =   975
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
      Left            =   3480
      TabIndex        =   8
      Top             =   6960
      Width           =   975
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmPGX_Util_Migracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbLimpia()
If vPaso Then Exit Sub

    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtCasos.Text = 0
    txtNoExisten.Text = 0

    txtArchivo.Text = ""
End Sub

Private Sub Form_Activate()
vModulo = 10
End Sub

Private Sub Form_Load()

vModulo = 10

vGrid.AppearanceStyle = fxGridStyle

txtArchivo.Text = ""

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


Me.MousePointer = vbHourglass

vGrid.MaxRows = 0
curMonto = 0
lCasos = 0 'Total



 DaoControl.Connect = "Excel 8.0;"
 DaoControl.DatabaseName = txtArchivo.Text
 DaoControl.RecordSource = "Import$"
 DaoControl.Refresh
 
Dim strSQL As String, rs As New ADODB.Recordset
Dim vOperacion As Long, vSaldo As Currency

 txtNoExisten.Text = 0
 txtCasos.Text = 0
 txtMonto.Text = 0
 
 With DaoControl.Recordset
 
     Do While Not .EOF
             vGrid.MaxRows = vGrid.MaxRows + 1
             vGrid.Row = vGrid.MaxRows
             vGrid.Col = 1
             vGrid.Text = !Codigo
             vGrid.Col = 2
             vGrid.Text = !Operacion & ""
             vGrid.Col = 3
             vGrid.Text = !Cedula
             vGrid.Col = 4
             vGrid.Text = !Nombre & ""
             vGrid.Col = 5
             vGrid.Text = !Formaliza & ""
             vGrid.Col = 6
             vGrid.Text = CStr(!PriDeduc)
             vGrid.Col = 7
             vGrid.Text = CStr(!FecUlt)
             vGrid.Col = 8
             vGrid.Text = CStr(!Monto)
             vGrid.Col = 9
             vGrid.Text = CStr(!Plazo)
             vGrid.Col = 10
             vGrid.Text = CStr(!Tasa)
             vGrid.Col = 11
             vGrid.Text = CStr(!Cuota)
             vGrid.Col = 12
             vGrid.Text = CStr(!Saldo)
             
             curMonto = curMonto + !Saldo
             txtCasos = txtCasos + 1
             txtCasos.Refresh
        .MoveNext
     Loop
 End With
 DaoControl.Recordset.Close


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
Dim vProcesados As Long, lng As Long

Dim pCedula As String, pNombre As String, pLinea As String
Dim pMonto As Currency, pTasa As Currency, pPlazo As Integer, pCuota As Currency, pSaldo As Currency
Dim pOperacion As Long, pPriDeduc As Long, pFecUlt As Long, pFormaliza As Date, pReferencia As String

Dim vFecha As Date, vComite As Integer, vGarantia As String, vObservaciones As String, vDestino As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor
vComite = 1
vGarantia = "N"
vObservaciones = "Datos Migrados"
vDestino = ""


With vGrid
    ProgressBarX.Max = .MaxRows + 1
    For lng = 1 To .MaxRows
    
       .Row = lng
       
       .Col = 1
       pLinea = Trim(.Text)
       .Col = 2
       pReferencia = Trim(.Text)
       .Col = 3
       pCedula = Trim(.Text)
       .Col = 4
       pNombre = Trim(.Text)
       .Col = 5
       pFormaliza = Trim(.Text)
       .Col = 6
       pPriDeduc = Trim(.Text)
       .Col = 7
       pFecUlt = Trim(.Text)
       .Col = 8
       pMonto = CCur(.Text)
       .Col = 9
       pPlazo = CInt(.Text)
       .Col = 10
       pTasa = CCur(.Text)
       .Col = 11
       pCuota = CCur(.Text)
       .Col = 12
       pSaldo = CCur(.Text)
       
       
    
    'Insertar la operacion
    strSQL = "insert reg_creditos(codigo,id_comite,cedula,montosol,montoapr,monto_girado" _
           & ",saldo,amortiza,interesc,saldo_mes,cuota,int,interesv,plazo,userrec,userres" _
           & ",userfor,usertesoreria,tesoreria,fechasol,fechares,fechaforp,fechaforf" _
           & ",fecha_calculo_int,garantia,primer_cuota,tdocumento,ndocumento,pagare" _
           & ",firma_deudor,premio,observacion,estado,prideduc,fecult,estadosol,documento_referido" _
           & ",cod_destino)" _
           & " values('" & UCase(pLinea) & "'," & vComite & ",'" _
           & Trim(pCedula) & "'," & pMonto & "," & pMonto & "," & pMonto & "," & pSaldo & "," & pMonto - pSaldo & ",0," _
           & pSaldo & "," & pCuota & "," & pTasa & "," & pTasa & "," & pPlazo & ",'" & glogon.Usuario & "','" & glogon.Usuario _
           & "','" & glogon.Usuario & "'," & "'" & glogon.Usuario & "','" & Format(pFormaliza, "yyyy/mm/dd") & "','" _
           & Format(pFormaliza, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(pFormaliza, "yyyy/mm/dd") & "','" _
           & Format(pFormaliza, "yyyy/mm/dd") & "','" & Format(pFormaliza, "yyyy/mm/dd") & "','" & vGarantia & "'" _
           & ",'N','OT','" & pReferencia & "',0,1,0,'" & vObservaciones & "','A'," & pPriDeduc _
           & "," & pFecUlt & ",'F','" & pReferencia & "',Null)"
      
     Call ConectionExecute(strSQL)
     'pOperacion = fxUltimaOperacion(pCedula)
     
     If GLOBALES.SysPlanPagos = 1 Then
        strSQL = "exec spCrdPlanPagos " & pOperacion
        Call ConectionExecute(strSQL)
     End If
     
    ProgressBarX.Value = lng
    Next lng
End With


  
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
vGrid.Height = Me.Height - (vGrid.Top + 1500)


tlbProceso.Top = vGrid.Top + vGrid.Height + 350
txtCasos.Top = tlbProceso.Top
txtMonto.Top = tlbProceso.Top
txtNoExisten.Top = tlbProceso.Top

End Sub

Private Sub tlbProceso_ButtonClick(ByVal Button As MSComctlLib.Button)


If UCase(glogon.Usuario) <> "PEDRO" Then
   MsgBox "Esta Opción es restringida para usuarios administradores", vbExclamation
   Exit Sub
End If

Select Case Button.Key
  Case "Aplicar"
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen casos a Procesar, verifique!", vbExclamation
       Exit Sub
    End If
   
    Call sbProcesar
  
  Case "cancelar"
    vGrid.MaxRows = 0
    txtArchivo.Text = ""

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


With Cmd
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL 97-2003]..."
        .Filter = "*.xls"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If

        If UCase(Right(.FileName, 3)) <> "XLS" Then
            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
            Exit Sub
        End If
        txtArchivo.Text = .FileName

End With

End Sub



