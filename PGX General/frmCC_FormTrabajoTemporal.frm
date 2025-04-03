VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmCC_FormTrabajoTemporal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario para Trabajo temporal"
   ClientHeight    =   7644
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   10704
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7644
   ScaleWidth      =   10704
   Begin VB.TextBox txtContratos 
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox txtSocios 
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
      TabIndex        =   10
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox txtCasos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   9
      Top             =   7080
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
      TabIndex        =   8
      ToolTipText     =   "Monto"
      Top             =   7080
      Width           =   1575
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
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2100
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
      TabIndex        =   6
      Top             =   2040
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
      TabIndex        =   5
      Top             =   1200
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
      Height          =   330
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   6855
   End
   Begin VB.ComboBox cboOperadora 
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   3
      Top             =   120
      Width           =   6855
   End
   Begin VB.ComboBox cboPlan 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   6855
   End
   Begin VB.CheckBox chkExcel 
      Alignment       =   1  'Right Justify
      Caption         =   "Forzar el Formato del Archivo a Excel!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      TabIndex        =   1
      Top             =   2040
      Width           =   3255
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   168
      Left            =   0
      TabIndex        =   0
      Top             =   7476
      Width           =   10704
      _ExtentX        =   18881
      _ExtentY        =   296
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.Toolbar tlbX 
      Height          =   528
      Left            =   9480
      TabIndex        =   7
      Top             =   1200
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
   Begin MSComDlg.CommonDialog Cmd 
      Left            =   9960
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10080
      Top             =   120
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
            Picture         =   "frmCC_FormTrabajoTemporal.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_FormTrabajoTemporal.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_FormTrabajoTemporal.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_FormTrabajoTemporal.frx":13926
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_FormTrabajoTemporal.frx":1A188
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbProceso 
      Height          =   330
      Left            =   6600
      TabIndex        =   12
      Top             =   6960
      Width           =   3810
      _ExtentX        =   6710
      _ExtentY        =   572
      ButtonWidth     =   1931
      ButtonHeight    =   582
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
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3975
      Left            =   120
      TabIndex        =   13
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
      MaxCols         =   5
      ScrollBars      =   2
      SpreadDesigner  =   "frmCC_FormTrabajoTemporal.frx":1A2A3
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sin Cont."
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
      Index           =   3
      Left            =   4560
      TabIndex        =   22
      Top             =   6840
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
      Left            =   3600
      TabIndex        =   21
      Top             =   6840
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
      TabIndex        =   20
      Top             =   6840
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   6480
      X2              =   6480
      Y1              =   6720
      Y2              =   7560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   11280
      Y1              =   6720
      Y2              =   6720
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
      TabIndex        =   19
      Top             =   7080
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   10680
      Y1              =   3000
      Y2              =   3000
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
      TabIndex        =   18
      Top             =   2040
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
      TabIndex        =   17
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   16
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Operadora"
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
      Index           =   0
      Left            =   1200
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Plan"
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
      Index           =   3
      Left            =   1200
      TabIndex        =   14
      Top             =   840
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   576
      Left            =   240
      Picture         =   "frmCC_FormTrabajoTemporal.frx":1BFC6
      Top             =   120
      Width           =   576
   End
End
Attribute VB_Name = "frmCC_FormTrabajoTemporal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mInstitucion As Long, mOperadora As Long, mPlan As String
Dim mCodigoDeduc As String, mCuentaCxC As String, mContrato As Long
Dim lCodigoDeduc As String, vPaso As Boolean

Private Sub sbLimpia()
    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtCasos.Text = 0
    txtSocios.Text = 0
    txtContratos.Text = 0
    txtArchivo.Text = ""
End Sub


Private Sub cboCliente_Change()
 Call sbLimpia
End Sub




Private Sub cboInstitucion_Click()
 Call sbLimpia
 
 mInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
 Call sbDocumentosInstitucion
 
End Sub

Private Sub cboOperadora_Click()
Dim strSQL As String

mOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)

strSQL = "select rtrim(cod_plan) + ' - ' + descripcion as ItmX from fnd_planes where deduce_independiente = 1 and cod_operadora = " & mOperadora
vPaso = True

Call sbLlenaCbo(cboPlan, strSQL, False, False)

vPaso = False

Call cboPlan_Click

End Sub

Private Sub cboPlan_Click()

If vPaso Then Exit Sub

mPlan = SIFGlobal.fxCodText(cboPlan.Text)
lCodigoDeduc = fxCodigoDeduccion
mCodigoDeduc = CStr(lCodigoDeduc)
 
End Sub

Private Sub cboProceso_Change()
 Call sbLimpia
End Sub

Private Sub cboTipo_Change()
 Call sbLimpia
End Sub

Private Sub chkExcel_Click()
 Call sbLimpia
End Sub

Private Sub Form_Activate()
vModulo = 18
End Sub

Private Sub Form_Load()
Dim strSQL As String, i As Integer
Dim vProceso As Long

vModulo = 18
Me.Icon = Me.Picture
vGrid.AppearanceStyle = fxGridStyle


strSQL = "select cod_institucion as IdX,descripcion as ItmX from instituciones where activa = 1"
Call sbLlenaCbo(cboInstitucion, strSQL, False, True)


strSQL = "select cod_operadora as IdX, descripcion as ItmX from FND_Operadoras"
Call sbLlenaCbo(cboOperadora, strSQL, False, True)

txtArchivo.Text = ""

vGrid.MaxCols = 7
vGrid.MaxRows = 0

vProceso = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
vProceso = fxFechaProcesoAnterior(vProceso)
cboProceso.AddItem vProceso

For i = 1 To 6
  vProceso = fxFechaProcesoSiguiente(vProceso)
  cboProceso.AddItem vProceso
Next i
cboProceso.Text = GLOBALES.glngFechaCR

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

If cboOperadora.ListCount <= 0 Then
    MsgBox "No existe ninguna Operadora, no se puede procesar el archivo...", vbCritical
    Exit Sub
End If

If cboInstitucion.ListCount <= 0 Then
    MsgBox "No existe ninguna Institución, no se puede procesar el archivo...", vbCritical
    Exit Sub
End If
If cboPlan.ListCount <= 0 Then
   MsgBox "No existe ningun plan, no se puede procesar el archivo...", vbCritical
   Exit Sub
End If


Me.MousePointer = vbHourglass

vGrid.MaxRows = 0
curMonto = 0
lCasos = 0 'Total

If vTipo = 1 Then 'Archivo de excel

        DaoControl.Connect = "Excel 8.0;"
        DaoControl.DatabaseName = txtArchivo.Text
        DaoControl.RecordSource = "SIF$"
        DaoControl.Refresh
        
        If LCase(DaoControl.Recordset.Fields(0).Name) <> "cedula" Then
           MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
                 "Los campos son Cedula, Nombre, Fondos"
           Exit Sub
        End If
        
        If LCase(DaoControl.Recordset.Fields(1).Name) <> "nombre" Then
           MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
                 "Los campos son Cedula, Nombre, Fondos"
           Exit Sub
        End If
        
        If LCase(DaoControl.Recordset.Fields(2).Name) <> "fondos" Then
           MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
                 "Los campos son Cedula, Nombre, Fondos"
           Exit Sub
        End If
        
        
        With vGrid
        
            Do While Not DaoControl.Recordset.EOF
              If Trim(DaoControl.Recordset!Cedula) <> "" Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .col = 1
                    .Text = DaoControl.Recordset!Cedula
                    .col = 2
                    .Text = DaoControl.Recordset!Nombre
                    .col = 3
                    If fxNombre(DaoControl.Recordset!Cedula) = "" Then
                        .Value = 1
                        txtSocios.Text = txtSocios + 1
                    Else
                        .Value = 0
                    End If
                    .col = 4
                    If fxExisteContrato(DaoControl.Recordset!Cedula) Then
                        .Value = 0
                        .CellTag = mContrato
                    Else
                        .Value = 1
                        txtContratos = txtContratos + 1
                        txtContratos.Refresh
                    End If
                    .col = 5
                    .Text = Format(DaoControl.Recordset!fondos, "Standard")
                    curMonto = curMonto + DaoControl.Recordset!fondos
                    txtCasos = txtCasos + 1
                    txtCasos.Refresh
               End If
              DaoControl.Recordset.MoveNext
            Loop
        End With
        DaoControl.Recordset.Close


Else 'Archivo Texto
    fn = FreeFile
    Open txtArchivo.Text For Input As #fn    ' Lee el archivo.
    Do While Not EOF(fn)
        Input #fn, strCadena
        If Mid(strCadena, 12, 6) = lCodigoDeduc Then
            strNombre = ""
            strCedula = ""
            'monto del archivo
            strMonto = Format(Mid(strCadena, 28, 13), "###########")
            strMonto = LTrim(RTrim(strMonto))
            If Len(strMonto) > 2 Then
                strMonto = Mid(strMonto, 1, Len(strMonto) - 2) & "." & Mid(strMonto, Len(strMonto) - 1, Len(strMonto))
            Else
                strMonto = "0" & "." & strMonto
            End If
            
            curMonto = curMonto + strMonto
            If Len(strCadena) > "54" Then
               strCedula = Trim(Format(Mid(strCadena, 1, 11), "###########"))
               strNombre = Trim(Mid(strCadena, Len(strCadena) - 31, 30))
            End If
            With vGrid
                If Len(strCadena) > "54" Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .col = 1
                    .Text = strCedula
                    .col = 2
                    .Text = strNombre
                    .col = 3
                    If fxNombre(strCedula) = "" Then
                        .Value = 1
                        txtSocios.Text = txtSocios + 1
                    Else
                        .Value = 0
                    End If
                    .col = 4
                    If fxExisteContrato(strCedula) Then
                        .Value = 0
                        txtContratos = txtContratos + 1
                        txtContratos.Refresh
                    Else
                        .Value = 1
                    End If
                    .col = 5
                    .Text = Format(strMonto, "Standard")
               End If
            End With
            txtCasos = txtCasos + 1
            txtCasos.Refresh
        End If
    Loop
    Close #fn
        
End If 'end if tipo archivo


'Totales
txtMonto.Text = Format(curMonto, "Standard")
Me.MousePointer = vbDefault
MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical
    Call sbLimpia
End Sub

Private Sub sbProcesar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lng As Long, vCodigo As String
Dim vFecha As Date, vProceso As Long
Dim pCedula As String, pNombre As String, pMovimiento As String, pMonto As Currency
Dim vProcesados As Long, vGarantia As String, vComite As Integer, vPlazo As Integer
Dim vTipoDoc As String, vNumDoc As String, vConcepto As String


On Error GoTo vError

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor
vProceso = cboProceso.Text


vProcesados = 0

vNumDoc = vProceso & "." & Format(mInstitucion, "00") & ".Rnd.Pend"
vConcepto = "FND004"

vTipoDoc = "PLA"

 
With vGrid
    
    For lng = 1 To .MaxRows
    
       .Row = lng
       .col = 1
       pCedula = Trim(.Text)
       .col = 2
       pNombre = Trim(.Text)
       .col = 5
       pMonto = CCur(.Text)
       .col = 3
       If .Value = 1 Then
            If GLOBALES.SysASEVersion Then
                strSQL = "insert socios(id_promotor,cedula,cod_institucion,up,ut,cod_profesion" _
                       & ",cod_sector,FechaIngreso,EstadoActual,Nombre,TIPO_ID) values(" _
                       & "1,'" & pCedula & "'," & mInstitucion & ",'','',1,1,'" & Format(vFecha, "yyyy/mm/dd") & "','N','" & pNombre & "',1)"
            
            Else
                strSQL = "insert socios(id_promotor,cedula,cod_institucion,cod_departamento,cod_seccion,cod_profesion" _
                       & ",cod_sector,FechaIngreso,EstadoActual,Nombre,TIPO_ID) values(" _
                       & "1,'" & pCedula & "'," & mInstitucion & ",'','',1,1,'" & Format(vFecha, "yyyy/mm/dd") & "','N','" & pNombre & "',1)"
            End If
            Call ConectionExecute(strSQL)
            
            strSQL = "insert ahorro_consolidado(cedula,ahorro,aporte) values('" & pCedula & "',0,0)"
            Call ConectionExecute(strSQL)
       End If
               
               
       .col = 4
       'valida si el monto es mayor que cero
       If pMonto > 0 Then
            If .Value = 1 Then
                strSQL = "exec spPrmFondosPlanilla " & mInstitucion & "," & vProceso & "," & mOperadora _
                       & ",'" & Trim(mPlan) & "','" & Trim(pCedula) & "'," & pMonto & ",'" & vNumDoc & "','" & Trim(mCuentaCxC) _
                       & "','R','" & Format(vFecha, "yyyy/mm/dd") & "'"
                Call ConectionExecute(strSQL)
            Else
                strSQL = "update fnd_contratos set Rendimiento = Rendimiento + " & CCur(pMonto) & " where cod_contrato = " & .CellTag & " and cod_plan = '" & mPlan & "'"
                Call ConectionExecute(strSQL)
           
               strSQL = "Insert fnd_contratos_detalle(Cod_operadora, Cod_plan, Cod_Contrato, Fecha, Monto, Fecha_Proceso" _
                      & ", Tcon, Ncon, Fecha_Acredita,cod_concepto,usuario,cod_caja)" _
                      & " Values(" & mOperadora & ", '" & mPlan & "', " & .CellTag & ", dbo.MyGetdate(), " & pMonto & ", " & vProceso & " ,'" & vTipoDoc & "'" _
                      & " ,'" & vNumDoc & "' , '" & Format(vFecha, "yyyymmdd") & "','" & vConcepto & "','" & glogon.Usuario & "','') "
                Call ConectionExecute(strSQL)
            End If
       End If
        
    Next lng


End With

 Call sbFndAsiento(vProceso, mOperadora, mPlan, mCuentaCxC, vNumDoc)
  
Me.MousePointer = vbDefault
MsgBox "Proceso Aplicado Satisfactoriamente... Registros Procesados :" & vGrid.MaxRows

Call sbLimpia

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical
    Resume
    Call sbLimpia


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
    frmFNDPlanillaBitacora.Show
End Select

End Sub

Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "select planilla from instituciones" _
       & " where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
       
Call OpenRecordSet(rs, strSQL)
        
Select Case Button.Key
  
  Case "buscar"
        txtArchivo.Text = ""
    If chkExcel.Value = vbChecked Then
       Call sbBuscaArchivo(1)
    Else
        Select Case Trim(rs!planilla)
            Case "00", "03"
                Call sbBuscaArchivo(1)
            Case Else
                Call sbBuscaArchivo(2)
        End Select
    End If
  
  Case "cargar"
    If chkExcel.Value = vbChecked Then
       Call sbCargaDeducciones(1)
    Else
         Select Case Trim(rs!planilla)
            Case "00", "03"
               Call sbCargaDeducciones(1)
            Case Else
               Call sbCargaDeducciones(2)
        End Select
    End If
End Select

rs.Close

End Sub


Private Sub sbBuscaArchivo(vTipo As Integer)


With Cmd
    If vTipo = 1 Or chkExcel.Value = vbChecked Then
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
    Else
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo de Deducciones [Texto]..."
        .Filter = "*.txt"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If
        If UCase(Right(.FileName, 3)) = "XLS" Then
            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
            Exit Sub
        End If
        
        'If UCase(Right(.FileName, 3)) <> "TXT" Or UCase(Right(.FileName, 3)) <> "DAT" Then
         '   MsgBox "La Extensión del Archivo no es válido...", vbExclamation
         '   Exit Sub
        'End If

        txtArchivo.Text = .FileName

End If
End With

End Sub






Private Function fxExisteContrato(vCedula As String) As Boolean

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select cod_contrato from fnd_contratos where cedula = '" & vCedula & "' And cod_operadora = " & mOperadora & "" _
         & " and cod_plan = '" & mPlan & "' and estado ='A'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF Then
    fxExisteContrato = False
Else
    fxExisteContrato = True
    mContrato = rs!cod_contrato
End If
rs.Close
End Function


Private Sub sbDocumentosInstitucion()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select cta_Fondos from instituciones where cod_institucion  = " & mInstitucion & ""
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    mCuentaCxC = rs!cta_fondos
Else
    mCuentaCxC = ""
End If
rs.Close

End Sub

Private Function fxCodigoDeduccion() As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select CODIGO_DEDUC from fnd_planes where cod_plan = '" & mPlan & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
  fxCodigoDeduccion = Trim(rs!CODIGO_DEDUC)
Else
  fxCodigoDeduccion = ""
End If

End Function





Private Sub sbFndAsiento(vProceso As Long, vOperadora As Long, vPlan As String _
        , vCuentaPlanilla As String, Optional vComprobante As String = "")
Dim strSQL As String '


strSQL = "exec spFndPlanillaDirectaAsiento " & vProceso & "," & mInstitucion & "," & vOperadora & ",'" & vPlan _
       & "','" & Trim(vCuentaPlanilla) & "','" & vComprobante & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

End Sub




