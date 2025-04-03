VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCR_ReportesRevision 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5328
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   9504
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5328
   ScaleWidth      =   9504
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboFBase 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   480
      Width           =   3735
   End
   Begin VB.ComboBox cboOmisiones 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   3360
      Width           =   3735
   End
   Begin VB.ComboBox cboUsuarioRevisa 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   3000
      Width           =   3735
   End
   Begin VB.ComboBox cboEtiqueta 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2640
      Width           =   3735
   End
   Begin VB.ComboBox cboInstitucion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3720
      Width           =   3735
   End
   Begin VB.ComboBox cboComite 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2280
      Width           =   3735
   End
   Begin VB.ComboBox cboGarantia 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   4560
      MaxLength       =   4
      TabIndex        =   13
      Top             =   4440
      Width           =   855
   End
   Begin VB.CheckBox chkLineas 
      Appearance      =   0  'Flat
      Caption         =   "Todas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5640
      TabIndex        =   12
      Top             =   4140
      Width           =   1095
   End
   Begin VB.ComboBox cboUsuarios 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1560
      Width           =   3735
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ComboBox cboOficina 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
   End
   Begin MSComctlLib.TreeView ArbolExp 
      Height          =   4200
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4215
      _ExtentX        =   7430
      _ExtentY        =   7408
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgArbol"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   6240
      TabIndex        =   4
      Top             =   840
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   194904067
      CurrentDate     =   36278
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   8040
      TabIndex        =   5
      Top             =   840
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   194904067
      CurrentDate     =   36278
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   312
      Left            =   8280
      TabIndex        =   29
      Top             =   4920
      Width           =   1092
      _ExtentX        =   1926
      _ExtentY        =   550
      ButtonWidth     =   1693
      ButtonHeight    =   550
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgArbol"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte"
            Key             =   "Reporte"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgArbol 
      Left            =   120
      Top             =   4680
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesRevision.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesRevision.frx":6862
            Key             =   "imgCRD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesRevision.frx":6980
            Key             =   "imgCBR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesRevision.frx":6AAA
            Key             =   "imgSGT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesRevision.frx":6BD0
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesRevision.frx":6CDE
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesRevision.frx":6DEB
            Key             =   "imgEspecial"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesRevision.frx":6F04
            Key             =   "imgRetenciones"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesRevision.frx":7032
            Key             =   "imgSeguridad"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesRevision.frx":713F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Base"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   8
      Left            =   4560
      TabIndex        =   32
      Top             =   480
      Width           =   1452
   End
   Begin VB.Label lblReporte 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   30
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   4560
      TabIndex        =   28
      Top             =   3720
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Omisión"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   4560
      TabIndex        =   27
      Top             =   3360
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Us Revisa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   4560
      TabIndex        =   25
      Top             =   3000
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   4560
      TabIndex        =   23
      Top             =   2640
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   12
      Left            =   4560
      TabIndex        =   21
      Top             =   4125
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comité"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   17
      Left            =   4560
      TabIndex        =   20
      Top             =   2280
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Garantía"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   16
      Left            =   4560
      TabIndex        =   19
      Top             =   1920
      Width           =   1332
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5400
      TabIndex        =   18
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   14
      Left            =   4560
      TabIndex        =   17
      Top             =   1560
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Reporte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   10
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   4560
      TabIndex        =   9
      Top             =   840
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inicio"
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
      Height          =   315
      Index           =   5
      Left            =   5640
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corte"
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
      Height          =   315
      Index           =   6
      Left            =   7560
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   26
      Left            =   4560
      TabIndex        =   6
      Top             =   1200
      Width           =   1092
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes de Revisiones Analistas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmCR_ReportesRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()

vModulo = 3

Call Formularios(Me)
Call RefrescaTags(Me)

Call sbInicializa

End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
chkLineas.Value = vbChecked

cboTipo.Clear
cboTipo.AddItem "Detalle"
cboTipo.AddItem "Resumen"
cboTipo.Text = "Detalle"

cboFBase.Clear
cboFBase.AddItem "Formalización"
cboFBase.AddItem "Etiqueta"
cboFBase.Text = "Etiqueta"

strSQL = "select cod_grupo + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  crd_grupos"
Call sbLlenaCbo(cboUsuarios, strSQL)

strSQL = "select cod_grupo + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  crd_grupos"
Call sbLlenaCbo(cboUsuarioRevisa, strSQL)

strSQL = "select rtrim(Garantia) + ' - ' + rtrim(descripcion) as Itmx" _
       & " from crd_garantia_tipos"
Call sbLlenaCbo(cboGarantia, strSQL, True, False)

strSQL = "select rtrim(cod_oficina) + ' - ' + rtrim(descripcion) as Itmx" _
       & " from SIF_Oficinas order by cod_oficina"
Call sbLlenaCbo(cboOficina, strSQL, True, False)

strSQL = "select rtrim(descripcion) as Itmx, id_comite as Idx" _
       & " from comites"
Call sbLlenaCbo(cboComite, strSQL, True, True)

strSQL = "select rtrim(descripcion) as Itmx, cod_institucion as Idx" _
       & " from instituciones"
Call sbLlenaCbo(cboInstitucion, strSQL, True, True)

Call sbCargarComboEtiquetas

strSQL = "select cast(ID_ERROR as varchar(15)) + ' - ' + rtrim(descripcion) as ItmX , ID_ERROR as Idx" _
         & " from CRD_ANALISIS_ERRORES"
Call sbLlenaCbo(cboOmisiones, strSQL, True, True)

Call sbRefrescaArbol

Me.MousePointer = vbDefault

End Sub

Private Sub sbCargarComboEtiquetas()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

    cboEtiqueta.AddItem ("TODOS")
    
    strSQL = "SELECT TAG_CODIGO as llave,DESCRIPCION as describe FROM CRD_TAGS "
'       & "' order by TAG_CODIGO"
    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
      cboEtiqueta.AddItem Trim(rs!llave) & " - " & Trim(rs!describe)
      rs.MoveNext
    Loop
    rs.Close
    
    cboEtiqueta.Text = "TODOS"

    
    Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call sbReporte
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  lblDescripcion.Caption = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCodigo_LostFocus()
 If Len(Trim(txtCodigo)) > 0 Then lblDescripcion.Caption = fxDescribeCodigo(Trim(txtCodigo))
End Sub

Function fxDescribeCodigo(strCodigo As String) As String
Dim rsX As New ADODB.Recordset

rsX.Open "select descripcion from catalogo where codigo = '" & Trim(strCodigo) & "'", glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
 MsgBox "No se encontró codigo - " & strCodigo, vbCritical
Else
 fxDescribeCodigo = IIf(IsNull(rsX!Descripcion), "", rsX!Descripcion)
End If
rsX.Close
End Function

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo vError

  lblReporte.Caption = Node.Parent.Text & " " & Node.Text
  lblReporte.Tag = Node.Key

vError:
End Sub


Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean, Optional xkey As String = "N")
Dim nodx As Node, vKey As String
On Error Resume Next

Set nodx = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodx.Text = vTexto
    nodx.Tag = nodx.Index
    nodx.Image = vImagen
    If xkey = "N" Then
        nodx.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodx.Key = xkey
    End If
    
vKey = nodx.Key

If vExpand Then
    Set nodx = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodx.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodx.Tag = nodx.Index
End If
    
End Sub


Private Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim rs As New ADODB.Recordset, strSQL As String
Dim vPadre As String


With ArbolExp
  .Nodes.Clear
  Set vNode = .Nodes.Add(, , "Etiquetas", "Etiquetas", "imgRoot")
  Set vNode = .Nodes.Add(, , "Omisiones", "Omisiones", "imgRoot")
  Call sbCreaNodos("Etiquetas", "General", "imgCRD", False, "etiquetas.general")
  Call sbCreaNodos("Etiquetas", "Por Usuario", "imgCRD", False, "etiquetas.usuario")
  Call sbCreaNodos("Etiquetas", "Por Usuario Resumen", "imgCRD", False, "etiquetas.resumen")
  Call sbCreaNodos("Etiquetas", "Por Usuario/Garantía", "imgCRD", False, "etiquetas.garantia")
  Call sbCreaNodos("Omisiones", "General", "imgCRD", False, "omisiones.general")
  Call sbCreaNodos("Omisiones", "Por Usuario", "imgCRD", False, "omisiones.usuario")
  Call sbCreaNodos("Omisiones", "Por Usuario Resumen", "imgCRD", False, "omisiones.resumen")
  Call sbCreaNodos("Omisiones", "Por Usuario/Garantía", "imgCRD", False, "omisiones.garantia")

  .Nodes(1).Expanded = True
  .Nodes(2).Expanded = True
End With

End Sub



Private Sub sbReporte()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String, vTemp As String
Dim strSQL As String, rs As New ADODB.Recordset, vReporte As String

On Error GoTo vError

If lblReporte.Tag = "" Then
    MsgBox "Seleccione el reporte que desea imprimir"
    Exit Sub
End If

Me.MousePointer = vbHourglass

vFiltro = ""
strSQL = ""


With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Créditos"
 
    .Connect = glogon.ConectRPT
    
    If lblReporte.Tag = "omisiones.general" Or lblReporte.Tag = "omisiones.usuario" Or lblReporte.Tag = "omisiones.resumen" Then
        strSQL = strSQL & "{REG_CREDITOS.FECHAFORP}"
        vSubTitulo = "Operaciones formalizadas entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
    Else
        Select Case Mid(cboFBase.Text, 1, 1)
          Case "E"
            strSQL = strSQL & "{CRD_OPERACION_TAGS.REGISTRO_FECHA}"
            vSubTitulo = "Registro de etiquetas entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
          Case "F"
            strSQL = strSQL & "{REG_CREDITOS.FECHAFORP}"
            vSubTitulo = "Operaciones formalizadas entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
        End Select
    End If
    strSQL = strSQL & " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ") to date(" _
       & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"

    If cboOficina.Text <> "TODOS" Then
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & "{REG_CREDITOS.COD_OFICINA_F} = '" & fxCodigoCbo(cboOficina) & "'"
      vFiltro = vFiltro & " / Oficina : " & cboOficina.Text
    End If
    
    If cboGarantia.Text <> "TODOS" Then
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & "{REG_CREDITOS.GARANTIA} = '" & fxCodigoCbo(cboGarantia) & "'"
      vFiltro = vFiltro & " / Garantía : " & cboGarantia.Text
    End If
 
    If cboUsuarios.Text <> "TODOS" Then
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{CRD_GRPUSERS.COD_GRUPO} = '" & fxCodigoCbo(cboUsuarios) & "'"
          vFiltro = vFiltro & "/ Usuarios : " & cboUsuarios.Text
    End If

    If chkLineas.Value = vbUnchecked Then
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & "{REG_CREDITOS.Codigo} = '" & Trim(txtCodigo) & "'"
      vFiltro = vFiltro & "/ LINEA : " & UCase(txtCodigo)
    End If
    
    If lblReporte.Tag = "etiquetas.general" Or lblReporte.Tag = "etiquetas.usuario" Or lblReporte.Tag = "etiquetas.resumen" Then
    
        If cboEtiqueta.Text <> "TODOS" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{CRD_OPERACION_TAGS.TAG_CODIGO} = '" & SIFGlobal.fxCodText(cboEtiqueta) & "'"
          vFiltro = vFiltro & "/ Etiqueta : " & cboEtiqueta.Text
        End If
        
        If cboUsuarioRevisa.Text <> "TODOS" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{CRD_GRPUSERS_1.COD_GRUPO} = '" & fxCodigoCbo(cboUsuarioRevisa) & "'"
          vFiltro = vFiltro & "/ Us.Etiq : " & cboUsuarioRevisa.Text
        End If
        
    End If
    
    If cboOmisiones.Text <> "TODOS" Then
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{CRD_ANALISIS_ERRORESREG.ID_ERROR} = " & cboOmisiones.ItemData(cboOmisiones.ListIndex) & ""
        vFiltro = vFiltro & "/ Omición : " & cboOmisiones.Text
    End If
 
    If cboComite.Text <> "TODOS" Then
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & "{REG_CREDITOS.id_comite} = " & cboComite.ItemData(cboComite.ListIndex) & ""
      vFiltro = vFiltro & "/ COMITE : " & cboComite.Text
    End If
 
    If cboInstitucion.Text <> "TODOS" Then
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & "{SOCIOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ""
      vFiltro = vFiltro & "/ INSTITUCION : " & cboInstitucion.Text
    End If
 
    Select Case Trim(lblReporte.Tag)
    Case "etiquetas.general"
        vTitulo = "Etiquetas en Créditos"
        vReporte = "Credito_EtiquetasAnalistas.rpt"
    Case "etiquetas.usuario"
        vTitulo = "Etiquetas por Usuario"
        vReporte = "Credito_EtiquetasUsuarios.rpt"
    Case "etiquetas.resumen"
        vTitulo = "Etiquetas por Usuario Resumen"
        vReporte = "Credito_EtiquetasUsuariosResumen.rpt"
    Case "etiquetas.garantia"
        vTitulo = "Etiquetas por Usuario / Garantía"
        vReporte = "Credito_EtiquetasUsuariosGarantia.rpt"
    Case "omisiones.general"
        vTitulo = "Omisiones en Revisiones de Créditos"
        vReporte = "Credito_OmisionesAnalistas.rpt"
    Case "omisiones.usuario"
        vTitulo = "Omisiones en Revisiones por Usuario"
        vReporte = "Credito_OmisionesUsuarios.rpt"
    Case "omisiones.resumen"
        vTitulo = "Omisiones en Revisiones por Usuario Resumen"
        vReporte = "Credito_OmisionesUsuariosResumen.rpt"
    Case "omisiones.garantia"
        vTitulo = "Omisiones por Usuario / Garantía"
        vReporte = "Credito_OmisionesGarantia.rpt"
    End Select
 
    .Formulas(0) = "fxFecha='Fecha: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(3) = "fxTitulo='" & vTitulo & "'"
    .Formulas(4) = "fxSubTitulo='" & Mid(vSubTitulo, 1, 250) & "'"
    .Formulas(5) = "fxFiltros='" & Mid(vFiltro, 1, 250) & "'"
    
    .ReportFileName = SIFGlobal.fxPathReportes(vReporte)
 
    .SelectionFormula = strSQL
    
    .PrintReport

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




