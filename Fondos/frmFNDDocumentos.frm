VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFNDDocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cola de Documentos - FND"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   HelpContextID   =   9006
   Icon            =   "frmFNDDocumentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   8175
   Begin VB.ComboBox cboOperadora 
      Height          =   315
      ItemData        =   "frmFNDDocumentos.frx":030A
      Left            =   840
      List            =   "frmFNDDocumentos.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   480
      Width           =   2295
   End
   Begin VB.Frame fraTraspaso 
      Caption         =   "Pase de Asientos a Contabilidad"
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   2040
      TabIndex        =   37
      Top             =   3120
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CheckBox chkAS_Depositos 
         Appearance      =   0  'Flat
         Caption         =   "Depósitos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkAS_Recibos 
         Appearance      =   0  'Flat
         Caption         =   "Recibos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkAS_ND 
         Appearance      =   0  'Flat
         Caption         =   "Notas de Débito"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   47
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkAS_NC 
         Appearance      =   0  'Flat
         Caption         =   "Notas de Crédito"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   46
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdAS_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3240
         TabIndex        =   41
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdAS_Aceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2280
         TabIndex        =   40
         Top             =   1440
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   135
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   4200
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblEstatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame fraReportes 
      Caption         =   "Reportes de Control "
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   3120
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CheckBox chkUsuarioActual 
         Caption         =   "Solo Usuario Actual"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CheckBox chkPorUsuario 
         Caption         =   "Agrupado por Usuario"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Width           =   1935
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Traspasos Generados"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3360
         TabIndex        =   32
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   2280
         TabIndex        =   31
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox chkFechaTraspaso 
         Alignment       =   1  'Right Justify
         Caption         =   "Fec.Base Traspaso"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2280
         TabIndex        =   30
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CheckBox chkFechaEmision 
         Alignment       =   1  'Right Justify
         Caption         =   "Fec.Base Emisión"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2280
         TabIndex        =   29
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkTodasLasFechas 
         Alignment       =   1  'Right Justify
         Caption         =   "Todas las Fechas "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   28
         Top             =   1080
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   300
         Left            =   3000
         TabIndex        =   26
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   111345667
         CurrentDate     =   36462
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Pendientes Traspaso"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "General"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   300
         Left            =   3000
         TabIndex        =   27
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   111345667
         CurrentDate     =   36462
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   4320
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         BorderWidth     =   2
         X1              =   120
         X2              =   4320
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label8 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   25
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      ItemData        =   "frmFNDDocumentos.frx":030E
      Left            =   3960
      List            =   "frmFNDDocumentos.frx":0310
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   480
      Width           =   1695
   End
   Begin MSComctlLib.Toolbar tlbDocumentos 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   635
      ButtonWidth     =   2196
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Reportes"
            Key             =   "reportes"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Traspaso"
            Key             =   "traspaso"
            Object.ToolTipText     =   "Pasa Asientos de Recibos a Contabilidad"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox txtDocumento 
      Height          =   315
      Left            =   6480
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin MSComctlLib.ListView lswAsiento 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   2940
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cuenta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Debe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Haber"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   8175
      Begin VB.TextBox txtConcepto 
         Height          =   315
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   20
         ToolTipText     =   "Concepto del Recibo"
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txtFechaTraspasa 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3960
         TabIndex        =   18
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtUS_Traspasa 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtUS_Genera 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6360
         TabIndex        =   13
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtFechaGenera 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3960
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6360
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtBeneficiario 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   4455
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6720
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFNDDocumentos.frx":0312
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFNDDocumentos.frx":0632
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFNDDocumentos.frx":0952
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFNDDocumentos.frx":0C72
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblUS 
         Caption         =   "Concepto"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha.Tra"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   19
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblUS 
         Caption         =   "US.Traspasa"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblUS 
         Caption         =   "US.Genera"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha.Gen"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Monto"
         Height          =   255
         Left            =   5640
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Estado"
         Height          =   255
         Left            =   5640
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Operadora"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   0
      TabIndex        =   45
      Top             =   480
      Width           =   855
   End
   Begin VB.Image imgReImpresion 
      Height          =   255
      Left            =   7800
      Picture         =   "frmFNDDocumentos.frx":0F8E
      Stretch         =   -1  'True
      ToolTipText     =   "Presione Aqui para Reimprimir el Doc."
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   3120
      TabIndex        =   35
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asiento del Documento"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   8175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Doc #"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   5640
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmFNDDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbLimpiaDatos()
 txtBeneficiario = ""
 txtEstado = ""
 txtFechaGenera = ""
 txtFechaTraspasa = ""
 txtConcepto = ""
 txtMonto = ""
 txtUS_Traspasa = ""
 txtUS_Genera = ""
 lswAsiento.ListItems.Clear
End Sub

Private Sub cmdAS_Aceptar_Click()
Dim iRespuesta As Integer

iRespuesta = MsgBox("Esta seguro de realizar el traspaso a contabilidad", vbYesNo)

If iRespuesta = vbYes Then
   Call sbGeneraAsientos
End If 'Respuesta

End Sub

Private Sub cmdAS_Cancelar_Click()
fraTraspaso.Visible = False
End Sub

Private Sub chkTodasLasFechas_Click()
If chkTodasLasFechas.Value = vbChecked Then
 dtpDesde.Enabled = False
 dtpHasta.Enabled = False
 chkFechaEmision.Value = 0
 chkFechaEmision.Enabled = False
 chkFechaTraspaso.Value = 0
 chkFechaTraspaso.Enabled = False
Else
 dtpDesde.Enabled = True
 dtpHasta.Enabled = True
 chkFechaEmision.Enabled = True
 chkFechaTraspaso.Enabled = True
End If
End Sub

Private Sub cmdCancelar_Click()
 fraReportes.Visible = False
End Sub

Private Function fxFechaReportes(vTipo As Integer) As String
If vTipo = 1 Then
 fxFechaReportes = Year(dtpDesde.Value) & "," & Month(dtpDesde.Value) & "," & Day(dtpDesde.Value)
Else
 fxFechaReportes = Year(dtpHasta.Value) & "," & Month(dtpHasta.Value) & "," & Day(dtpHasta.Value)
End If
End Function


Private Sub cmdImprimir_Click()
Dim vTipo As String, vOperadora As Long

If (chkTodasLasFechas.Value + chkFechaEmision.Value _
    + chkFechaTraspaso.Value) = 0 Then
  MsgBox "No se ha especificado ninguna fecha como parámetro de busqueda...", vbInformation
  Exit Sub
End If

Me.MousePointer = vbHourglass

vTipo = fxTipoDocumento(cboTipo.Text)
vOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)

With frmContenedor.Crt
    .Reset
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes de Control de Documentos del Fondo"
   
    .Connect = glogon.ConectRPT
   
   If Me.chkPorUsuario.Value = vbUnchecked Then
    .ReportFileName = SIFGlobal.fxPathReportes("Fondos_DocumentosControl.rpt")
   Else
    .ReportFileName = SIFGlobal.fxPathReportes("Fondos_DocumentosControlUsr.rpt")
   End If

    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "usuario='" & glogon.Usuario & "'"
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
     
    'Agregar
    .Formulas(5) = "operadora='" & cboOperadora.Text & "'"

  Select Case True
   Case optReportes(0).Value  'Reporte General
     .Formulas(3) = "SUBTITULO='REPORTE GENERAL - " & UCase(cboTipo.Text) & "'"
        If chkFechaEmision.Value = vbChecked Then
          .SelectionFormula = "{FND_DOCUMENTOS.FECHA} in Date(" & fxFechaReportes(1) & ")" _
                    & " To Date(" & fxFechaReportes(0) & ")" _
                    & " AND {FND_DOCUMENTOS.TIPO} = '" & vTipo & "' AND {FND_DOCUMENTOS.COD_OPERADORA} = " & vOperadora
          .Formulas(4) = "fecha_emision = 'Fecha Emisión entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
        End If
        If chkFechaTraspaso.Value = vbChecked Then
          .SelectionFormula = "{FND_DOCUMENTOS.FECHA_TRASPASO} >= Date(" & fxFechaReportes(1) & ")" _
                    & " AND {FND_DOCUMENTOS.FECHA_TRASPASO} <= Date(" & fxFechaReportes(0) & ")" _
                    & " AND {FND_DOCUMENTOS.TIPO} = '" & vTipo & "' AND {FND_DOCUMENTOS.COD_OPERADORA} = " & vOperadora
          .Formulas(4) = "fecha_traspaso = 'Fecha Traspaso entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
        End If

'   Case optReportes(1).Value  'Emitidos
'     .SelectionFormula = "{FND_DOCUMENTOS.TIPO} = '" & vTipo & "' AND {FND_DOCUMENTOS.COD_OPERADORA} = " & vOperadora
'     .Formulas(3) = "SUBTITULO='" & UCase(cboTipo.Text) & " - EMITIDOS'"
   Case optReportes(3).Value  'Pendientes de Traspaso
     .SelectionFormula = "isnull({FND_DOCUMENTOS.TRASPASO}) = true" _
                       & " AND {FND_DOCUMENTOS.TIPO} = '" & vTipo & "'" _
                       & " AND {FND_DOCUMENTOS.COD_OPERADORA} = " & vOperadora
     .Formulas(3) = "SUBTITULO='" & UCase(cboTipo.Text) & " - PENDIENTES TRSP.'"
   Case optReportes(4).Value  'Traspasos Generados
     .SelectionFormula = "isnull({FND_DOCUMENTOS.TRASPASO}) = FALSE" _
                       & " AND {FND_DOCUMENTOS.TIPO} = '" & vTipo & "'" _
                       & " AND {FND_DOCUMENTOS.COD_OPERADORA} = " & vOperadora
     .Formulas(3) = "SUBTITULO='" & UCase(cboTipo.Text) & " - TRSP. GENERADOS'"
  End Select

  If chkTodasLasFechas.Value = vbUnchecked And Not optReportes(0).Value Then
    If chkFechaEmision.Value = vbChecked Then
      .SelectionFormula = .SelectionFormula & " AND {FND_DOCUMENTOS.FECHA} in Date(" & fxFechaReportes(1) & ")" _
                & " to Date(" & fxFechaReportes(0) & ")"
      .Formulas(4) = "fecha_emision = 'Fecha Emisión entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
    End If
    If chkFechaTraspaso.Value = vbChecked Then
      .SelectionFormula = .SelectionFormula & " AND {FND_DOCUMENTOS.FECHA_TRASPASO} >= Date(" & fxFechaReportes(1) & ")" _
                & " AND {FND_DOCUMENTOS.FECHA_TRASPASO} <= Date(" & fxFechaReportes(0) & ")"
      .Formulas(4) = "fecha_traspaso = 'Fecha Traspaso entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
    End If

   End If

   If chkUsuarioActual.Value = vbChecked Then
     .SelectionFormula = .SelectionFormula & " AND {FND_DOCUMENTOS.USUARIO} = '" _
                       & glogon.Usuario & "'"
   End If


   .PrintReport

End With

Me.MousePointer = vbDefault

End Sub


Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
 dtpDesde.Value = fxFechaServidor
 dtpHasta.Value = dtpDesde
 
' vModulo = 10 'Cuentas Corrientes
 vModulo = 18 'Fondo de Inversion

 cboTipo.AddItem "Recibo"
 cboTipo.AddItem "Nota Credito"
 cboTipo.AddItem "Nota Debito"
 cboTipo.AddItem "Depositos"
 
 cboTipo.Text = "Recibo"

 Call sbgFNDCargaCombos(cboOperadora, "Operadoras")

 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
End Sub

Sub sbGeneraAsientos()
Dim rs As New ADODB.Recordset, strSQL As String
Dim intLinea As Integer, DH As String, strDocumentos As String
Dim rs2 As New ADODB.Recordset, vTipoAsiento As String
Dim vFecha As Date, vDetalle As String, vNumAsiento As String

Me.MousePointer = vbHourglass
Me.fraTraspaso.Visible = True

lblEstatus.Caption = "Cargando Información..."
lblEstatus.Refresh
prgBar.Value = 1

On Error GoTo CapturaError

strDocumentos = ""
vFecha = fxFechaServidor


If chkAS_Depositos.Value = vbChecked Then strDocumentos = "'DP'"

If chkAS_NC.Value = vbChecked Then
  If Len(strDocumentos) > 1 Then
    strDocumentos = strDocumentos & ",'NC'"
  Else
    strDocumentos = "'NC'"
  End If
End If

If chkAS_ND.Value = vbChecked Then
  If Len(strDocumentos) > 1 Then
    strDocumentos = strDocumentos & ",'ND'"
  Else
    strDocumentos = "'ND'"
  End If
End If

If chkAS_Recibos.Value = vbChecked Then
  If Len(strDocumentos) > 1 Then
    strDocumentos = strDocumentos & ",'RE'"
  Else
    strDocumentos = "'RE'"
  End If
End If


If Len(strDocumentos) = 0 Then
  Me.MousePointer = vbDefault
  Exit Sub
End If


strSQL = "select * from fnd_documentos where fecha_traspaso is null" _
       & " and tipo in(" & strDocumentos & ") and cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & " order by tipo,id_documento"
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

prgBar.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
lblEstatus.Caption = "Procesando Asientos..."
lblEstatus.Refresh

Do While Not rs.EOF
  
 If fxgCntPeriodoValida(rs!fecha) Then 'Verificar el Periodo Abierto en contabilidad
    'Crea Maestro
   vTipoAsiento = fxgFNDTipoAsientoDoc(rs!Tipo)
   vNumAsiento = "FND" & Format(rs!Cod_Operadora, "00") & "-" & Format(rs!id_documento, "00000000")
   
   strSQL = "insert cntX_asientos(cod_contabilidad,tipo_asiento,num_asiento,anio,mes,fecha_asiento,descripcion,balanceado,modulo)" _
          & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','" & vNumAsiento & "'," & Year(rs!fecha) & "," & Month(rs!fecha) _
          & ",'" & Format(rs!fecha, "yyyy/mm/dd") & "','" & rs!Concepto & "','S'," & vModulo & ")"
   Call ConectionExecute(strSQL)

    'Crea Detalle
    intLinea = 1
    strSQL = "select * from fnd_asientos where id_documento = " & rs!id_documento _
             & " and tipo = '" & rs!Tipo & "' and cod_operadora = " & rs!Cod_Operadora
             
    rs2.CursorLocation = adUseServer
    rs2.Open strSQL, glogon.Conection, adOpenStatic
    Do While Not rs2.EOF
        If UCase(rs2!fnd_DEBEHABER) = "H" Then  'dc - dh
          DH = "C"
        Else
          DH = rs2!fnd_DEBEHABER
        End If
        'Ahora se pone en el detalle de la cuenta el numero de deposito y luego
        'Lo que alcance del concepto
        vDetalle = rs!Concepto
        vDetalle = Mid(vDetalle, 1, 59)

        If DH = "C" Then 'Acredita
            strSQL = "insert cntx_asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
                & ",detalle,documento,cod_unidad,cod_divisa,tipo_cambio,cod_centro_costo)" _
                & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','" & vNumAsiento & "'," & intLinea & "," & Trim(rs2!fnd_cuenta) _
                & ",0," & rs2!fnd_monto & ",'" & vDetalle & "','" & Format(rs2!id_documento, "00000000") _
                & "','OC','COL',1,'')"
        Else 'Debita
            strSQL = "insert cntx_asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
                & ",detalle,documento,cod_unidad,cod_divisa,tipo_cambio,cod_centro_costo)" _
                & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','" & vNumAsiento & "'," & intLinea & "," & Trim(rs2!fnd_cuenta) _
                & "," & rs2!fnd_monto & ",0,'" & vDetalle & "','" & Format(rs2!id_documento, "00000000") _
                & "','OC','COL',1,'')"
        End If
        
        Call ConectionExecute(strSQL)
        intLinea = intLinea + 1
        rs2.MoveNext
    Loop
    rs2.Close

    'Actualizar el estado del recibo
    strSQL = "Update fnd_documentos set FECHA_TRASPASO = '" & Format(vFecha, "yyyy/mm/dd") _
            & "',us_traspaso = '" & glogon.Usuario & "' where id_documento = " & rs!id_documento _
            & " and tipo = '" & rs!Tipo & "' and cod_operadora = " & rs!Cod_Operadora
    Call ConectionExecute(strSQL)
 
 Else
  
  MsgBox "Existen asientos que no pueden ser trasladados porque el periodo fué cerrado...", vbInformation
 
 End If 'Periodo

 If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
 
 rs.MoveNext

Loop
rs.Close

lblEstatus.Caption = ""
lblEstatus.Refresh
prgBar.Value = 1

'Call Bitacora("Aplica", "Asientos del Control de Documentos FND")

MsgBox "Se realizó el pase de asientos a contabilidad ", vbInformation
Me.MousePointer = vbDefault
Me.fraTraspaso.Visible = False

Exit Sub

CapturaError:
    lblEstatus.Caption = ""
    lblEstatus.Refresh
    prgBar.Value = 1
    Me.MousePointer = vbDefault
    Me.fraTraspaso.Visible = False
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub imgReImpresion_Click()
Dim vRecibo As Long, vTipoDoc As String, vOperadora As Long

On Error GoTo vError


vOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
vTipoDoc = fxTipoDocumento(cboTipo.Text)
vRecibo = txtDocumento

Call sbImprimeDocumento(vRecibo, vTipoDoc, vOperadora)
'
'
'Dim strSQL As String, x As New clsImpresoras
'Dim vDriver, vTipo As String
'Dim vFlat As Boolean, rs As New ADODB.Recordset
'Dim vEmpresa As String, vCedJur As String
'
'On Error GoTo vError:
'
'vFlat = False
'
'
''Enlace con el control de documentos del sistema ASE, para Politica General
''Despues hay que descenlazar.
'
'strSQL = "select cs_utilizar_reciboFlat as Flat from ase_consecutivos"
'Call OpenRecordSet(rs, strSQL)
'  vFlat = IIf((rs!Flat = "S"), True, False)
'rs.Close
'
'strSQL = "select nombre,cedula_juridica from sif_empresa"
'Call OpenRecordSet(rs, strSQL)
' vEmpresa = UCase(rs!Nombre & "")
' vCedJur = Trim(rs!cedula_juridica & "")
'rs.Close
'
'
'
'vTipo = fxTipoDocumento(cboTipo.Text)
'
'With frmContenedor.Crt
'   .Reset
'
'   .Connect = glogon.ConectRPT
'
'   If vTipo = "RE" Then
'     x.TipoImpresora = Recibos
'     x.Reset
'     .PrinterDriver = x.Controlador
'     .PrinterName = x.Nombre
'     .PrinterPort = x.Puerto
'
'     .PrinterSelect
'
'     .Destination = crptToPrinter
'
'      If vFlat Then
'         .Formulas(0) = "fxEmpresa = '" & vEmpresa & "'"
'         .Formulas(1) = "fxCedJur = '" & vCedJur & "'"
'         .ReportFileName = SIFGlobal.fxPathReportes("Fondos_DocumentoFlat.rpt")
'      Else
'         .ReportFileName = SIFGlobal.fxPathReportes("Fondos_DocumentoCls.rpt")
'      End If
'
'     .SelectionFormula = "{FND_DOCUMENTOS.ID_DOCUMENTO} = " & Trim(txtDocumento) _
'                     & " AND {FND_DOCUMENTOS.TIPO} = '" & vTipo & "'" _
'                     & " AND {FND_DOCUMENTOS.COD_OPERADORA} = " & cboOperadora.ItemData(cboOperadora.ListIndex)
'
'   Else
'     .WindowShowPrintSetupBtn = True
'     .WindowShowRefreshBtn = True
'     .WindowShowSearchBtn = True
'     .WindowState = crptMaximized
'     .WindowTitle = "Reportes de Control de Documentos"
'     .ReportFileName = SIFGlobal.fxPathReportes("Fondos_DocumentoBoleta.rpt")
'
'     .SelectionFormula = "{FND_DOCUMENTOS.ID_DOCUMENTO} = " & Trim(txtDocumento) _
'                     & " AND {FND_DOCUMENTOS.TIPO} = '" & vTipo & "'" _
'                     & " AND {FND_DOCUMENTOS.COD_OPERADORA} = " & cboOperadora.ItemData(cboOperadora.ListIndex)
'
'
''                     & " AND {CUENTAS.cod_contabilidad} = " & GLOBALES.gEnlace
''
'   End If
'   .PrintReport
'End With

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlbDocumentos_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim iRespuesta As Integer
Select Case Button.Key
  Case "traspaso"
    fraTraspaso.Visible = True
  Case "reportes"
    fraReportes.Visible = True
End Select
End Sub


Function fxTipoDocumento(vTipo As String) As String

    Select Case UCase(Trim(vTipo))
       Case "RECIBO"
            fxTipoDocumento = "RE"
       Case "DEPOSITO", "DEPOSITOS"
            fxTipoDocumento = "DP"
       Case "NOTA CREDITO"
            fxTipoDocumento = "NC"
       Case "NOTA DEBITO"
            fxTipoDocumento = "ND"
       Case Else
            fxTipoDocumento = vTipo
    End Select

End Function





Public Sub sbImprimeDocumento(pDocumentoId As Long, vTipo As String, ByVal vOperadora As Long)
Dim strSQL As String, x As New clsImpresoras
Dim vFlat As Boolean, rs As New ADODB.Recordset
Dim vEmpresa As String, vCedJur As String

On Error GoTo vError

vFlat = False


'Enlace con el control de documentos del sistema ASE, para Politica General
'Despues hay que descenlazar.

strSQL = "select cs_utilizar_reciboFlat as Flat from ase_consecutivos"
Call OpenRecordSet(rs, strSQL)
  vFlat = IIf((rs!Flat = "S"), True, False)
rs.Close

strSQL = "select nombre,cedula_juridica from sif_empresa"
Call OpenRecordSet(rs, strSQL)
 vEmpresa = UCase(rs!Nombre & "")
 vCedJur = Trim(rs!cedula_juridica & "")
rs.Close

x.TipoImpresora = Recibos
 
With frmContenedor.Crt
  .Reset
  
  .Connect = glogon.ConectRPT
  
 If vTipo = "RE" Or vTipo = "FRE" Then
    x.TipoImpresora = Recibos
    x.Reset
    .PrinterDriver = x.Controlador
    .PrinterName = x.Nombre
    .PrinterPort = x.Puerto

    .PrinterSelect

    .Destination = crptToPrinter

    If vFlat Then
        .Formulas(0) = "fxEmpresa = '" & vEmpresa & "'"
        .Formulas(1) = "fxCedJur = '" & vCedJur & "'"

            .ReportFileName = SIFGlobal.fxPathReportes("Fondos_DocumentoFlat.rpt")

    Else
        .Formulas(0) = "fxEmpresa = '" & vEmpresa & "'"
        .Formulas(1) = "fxCedJur = '" & vCedJur & "'"
            .ReportFileName = SIFGlobal.fxPathReportes("Fondos_DocumentoCls.rpt")
    End If

 Else
       .ReportFileName = SIFGlobal.fxPathReportes("Fondos_DocumentoBoleta.rpt")
    vFlat = False
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes de Control de Documentos"
 End If

   
    .SelectionFormula = "{FND_DOCUMENTOS.ID_DOCUMENTO} = " & pDocumentoId _
                      & " AND {FND_DOCUMENTOS.TIPO} = '" & vTipo & "'" _
                      & " AND {FND_DOCUMENTOS.COD_OPERADORA} = " & vOperadora

  
                    
 If Not vFlat And vTipo <> "RE" Then
    .SubreportToChange = "sbAsiento"
    .StoredProcParam(0) = vTipo
    .StoredProcParam(1) = pDocumentoId
    .StoredProcParam(2) = 1
 End If
 
 .PrintReport

End With

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCargaDocumento(vOperadora As Long, vTipo As String, lngDocumento As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListItem, curDebe As Currency, curHaber As Currency
Dim strTipo As String

curDebe = 0
curHaber = 0

strTipo = fxTipoDocumento(vTipo)

On Error Resume Next

strSQL = "select * from FND_Documentos where id_Documento = " & lngDocumento _
        & " and tipo = '" & strTipo & "' and cod_operadora = " & vOperadora
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

If rs.EOF And rs.BOF Then
  rs.Close
  MsgBox "No se encontró Documento", vbCritical
  Exit Sub
End If
 txtBeneficiario = IIf(IsNull(rs!Cliente), "", rs!Cliente)
 
 txtEstado = "Activo"
 txtFechaGenera = rs!fecha
 txtFechaTraspasa = IIf(IsNull(rs!fecha_traspaso), "", rs!fecha_traspaso)
 txtConcepto = rs!Concepto
 txtMonto = Format(rs!Monto, "###,###,###,##0.00")
 txtUS_Traspasa = IIf(IsNull(rs!us_traspaso), "", rs!us_traspaso)
 txtUS_Genera = rs!Usuario
rs.Close

strSQL = "select * from fnd_asientos where id_documento=" & lngDocumento _
        & " and tipo = '" & strTipo & "' and cod_operadora = " & vOperadora
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

With lswAsiento
   .ListItems.Clear
   Do While Not rs.EOF
     Set itmX = .ListItems.Add(.ListItems.Count + 1, , Format(rs!fnd_cuenta, GLOBALES.gstrMascara))
       itmX.SubItems(1) = fxgCntCuentaDesc(Trim(rs!fnd_cuenta))
       If rs!fnd_DEBEHABER = "D" Then
          itmX.SubItems(2) = Format(rs!fnd_monto, "###,###,###,##0.00")
          curDebe = curDebe + rs!fnd_monto
       Else
          itmX.SubItems(3) = Format(rs!fnd_monto, "###,###,###,##0.00")
          curHaber = curHaber + rs!fnd_monto
       End If
     rs.MoveNext
    Loop
    rs.Close
     Set itmX = .ListItems.Add(.ListItems.Count + 1, , "")
      itmX.SubItems(2) = "---------------------------"
      itmX.SubItems(3) = "---------------------------"
    
     Set itmX = .ListItems.Add(.ListItems.Count + 1, , "TOTALES")
      itmX.SubItems(2) = Format(curDebe, "###,###,###,##0.00")
      itmX.SubItems(3) = Format(curHaber, "###,###,###,##0.00")
End With
End Sub


Private Sub txtDocumento_Change()
 Call sbLimpiaDatos
End Sub

Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
 Call sbCargaDocumento(cboOperadora.ItemData(cboOperadora.ListIndex), cboTipo.Text, txtDocumento)
End If
End Sub

