VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCC_DocumentosOLD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cola de Documentos"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   HelpContextID   =   9006
   Icon            =   "CC_Recibos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   8175
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraConsultaDP 
      Caption         =   "Consulta de Documentos con # Depósitos xxx"
      ForeColor       =   &H00800000&
      Height          =   4695
      Left            =   3720
      TabIndex        =   86
      Top             =   960
      Visible         =   0   'False
      Width           =   6615
      Begin MSComctlLib.ListView lswDP 
         Height          =   3855
         Left            =   120
         TabIndex        =   89
         Top             =   720
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   6800
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Documento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Concepto"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cliente"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Deposito"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtNumDp 
         Height          =   315
         Left            =   1200
         TabIndex        =   88
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "# Depósito"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CheckBox chkConsultaDP 
      Caption         =   "Depósitos"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame fraConfiguracion 
      Caption         =   "Configuración"
      ForeColor       =   &H8000000D&
      Height          =   3135
      Left            =   120
      TabIndex        =   46
      Top             =   720
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox txtID_RE 
         Height          =   315
         Left            =   4200
         TabIndex        =   61
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtID_NC 
         Height          =   315
         Left            =   4200
         TabIndex        =   64
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtID_ND 
         Height          =   315
         Left            =   4200
         TabIndex        =   67
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtID_DP 
         Height          =   315
         Left            =   4200
         TabIndex        =   70
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtTA_DP 
         Height          =   315
         Left            =   3240
         TabIndex        =   69
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtTA_ND 
         Height          =   315
         Left            =   3240
         TabIndex        =   66
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtTA_NC 
         Height          =   315
         Left            =   3240
         TabIndex        =   63
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtTA_RE 
         Height          =   315
         Left            =   3240
         TabIndex        =   60
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox chkUtilizaRecibo 
         Alignment       =   1  'Right Justify
         Caption         =   "Utilizar Documentos"
         Height          =   255
         Left            =   3360
         TabIndex        =   48
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdCL 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4320
         TabIndex        =   73
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton cmdGuardaConfiguracion 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   3360
         TabIndex        =   72
         Top             =   2640
         Width           =   855
      End
      Begin MSMask.MaskEdBox medNC 
         Height          =   315
         Left            =   1440
         TabIndex        =   62
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medRE 
         Height          =   315
         Left            =   1440
         TabIndex        =   59
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medND 
         Height          =   315
         Left            =   1440
         TabIndex        =   65
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medDP 
         Height          =   315
         Left            =   1440
         TabIndex        =   68
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Consec."
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   6
         Left            =   4200
         TabIndex        =   71
         Top             =   600
         Width           =   975
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   5160
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   120
         X2              =   5160
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label9 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Depósitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   58
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota Débito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   57
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   56
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Recibos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   55
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta Contable"
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   0
         Left            =   1440
         TabIndex        =   54
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Asiento"
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   1
         Left            =   3240
         TabIndex        =   47
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame fraTraspaso 
      Caption         =   "Pase de Asientos a Contabilidad"
      ForeColor       =   &H8000000D&
      Height          =   1935
      Left            =   120
      TabIndex        =   51
      Top             =   3720
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CheckBox chkAsientoResumen 
         Caption         =   "Crear Asiento Resumen"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton cmdAS_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3240
         TabIndex        =   79
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdAS_Aceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2280
         TabIndex        =   78
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chkAS_NC 
         Caption         =   "Notas de Crédito"
         Height          =   255
         Left            =   1800
         TabIndex        =   77
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkAS_ND 
         Caption         =   "Notas de Débito"
         Height          =   255
         Left            =   1800
         TabIndex        =   76
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkAS_Recibos 
         Caption         =   "Recibos"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkAS_Depositos 
         Caption         =   "Depósitos"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   720
         Width           =   1695
      End
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   135
         Left            =   120
         TabIndex        =   52
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   4200
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   120
         X2              =   4200
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblEstatus 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "..."
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame fraReportes 
      Caption         =   "Reportes de Control "
      ForeColor       =   &H8000000D&
      Height          =   2895
      Left            =   840
      TabIndex        =   29
      Top             =   840
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CheckBox chkUsuarioActual 
         Caption         =   "Solo Usuario Actual"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CheckBox chkPorUsuario 
         Caption         =   "Agrupado por Usuario"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Traspasos Generados"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   44
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3360
         TabIndex        =   43
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   2280
         TabIndex        =   42
         Top             =   2400
         Width           =   975
      End
      Begin VB.CheckBox chkFechaTraspaso 
         Alignment       =   1  'Right Justify
         Caption         =   "Fec.Base Traspaso"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2280
         TabIndex        =   41
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CheckBox chkFechaEmision 
         Alignment       =   1  'Right Justify
         Caption         =   "Fec.Base Emisión"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2280
         TabIndex        =   40
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CheckBox chkFechaAnulacion 
         Alignment       =   1  'Right Justify
         Caption         =   "Fec.Base Anulación"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2280
         TabIndex        =   39
         Top             =   1320
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
         TabIndex        =   38
         Top             =   1080
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   300
         Left            =   3000
         TabIndex        =   36
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60948483
         CurrentDate     =   36462
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Pendientes Traspaso"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Anulados"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "Emitidos"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optReportes 
         Caption         =   "General"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   300
         Left            =   3000
         TabIndex        =   37
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60948483
         CurrentDate     =   36462
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   4320
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         BorderWidth     =   2
         X1              =   120
         X2              =   4320
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label8 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   35
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   34
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      ItemData        =   "CC_Recibos.frx":030A
      Left            =   840
      List            =   "CC_Recibos.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   480
      Width           =   1935
   End
   Begin MSComctlLib.Toolbar tlbRecibos 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   635
      ButtonWidth     =   2646
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Reportes"
            Key             =   "reportes"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Configuración"
            Key             =   "configuracion"
            Object.ToolTipText     =   "Configuracion de los Documentos ASE"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anular"
            Key             =   "anular"
            Object.ToolTipText     =   "Anula Recibo Actual"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Left            =   3600
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin MSComctlLib.ListView lswAsiento 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   4200
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
      BorderStyle     =   1
      Appearance      =   1
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
      Height          =   3015
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   8175
      Begin VB.TextBox txtDetalle 
         Height          =   615
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   80
         ToolTipText     =   "Detalle de la nota"
         Top             =   2280
         Width           =   6975
      End
      Begin VB.TextBox txtConcepto 
         Height          =   615
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   28
         ToolTipText     =   "Concepto del Recibo"
         Top             =   1680
         Width           =   6975
      End
      Begin VB.TextBox txtPago 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6480
         TabIndex        =   27
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtFechaTraspasa 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3960
         TabIndex        =   24
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtFechaAnula 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6480
         TabIndex        =   22
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtUS_Traspasa 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtUS_Anula 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6480
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtUS_Genera 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtTipo 
         Height          =   315
         Left            =   3960
         TabIndex        =   15
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtEstado 
         Height          =   315
         Left            =   6480
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtFechaGenera 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3960
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtBeneficiario 
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   4455
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7200
         Top             =   1560
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
               Picture         =   "CC_Recibos.frx":030E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CC_Recibos.frx":062E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CC_Recibos.frx":094E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CC_Recibos.frx":0C6E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblUS 
         Caption         =   "Detalle"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   82
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblUS 
         Caption         =   "Concepto"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   81
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Pago"
         Height          =   255
         Left            =   5640
         TabIndex        =   26
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Fec.Tra"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   25
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Fec.Anu"
         Height          =   255
         Index           =   2
         Left            =   5760
         TabIndex        =   23
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblUS 
         Caption         =   "US.Traspasa"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblUS 
         Caption         =   "US.Anula"
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblUS 
         Caption         =   "US.Genera"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Fec.Gen"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   9
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Monto"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
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
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   5520
      TabIndex        =   90
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1245185
   End
   Begin VB.Image imgReImpresion 
      Height          =   255
      Left            =   6120
      Picture         =   "CC_Recibos.frx":0F8A
      Stretch         =   -1  'True
      ToolTipText     =   "Presione Aqui para Reimprimir el Doc."
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   49
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ASIENTO - DOCUMENTO"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   8175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Doc #"
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   0
      Left            =   2760
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmCC_DocumentosOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean

Private Sub sbLimpiaDatos()
 txtBeneficiario = ""
 txtEstado = ""
 txtFechaAnula = ""
 txtFechaGenera = ""
 txtFechaTraspasa = ""
 txtConcepto = ""
 txtMonto = ""
 txtPago = ""
 txtTipo = ""
 txtUS_Anula = ""
 txtUS_Traspasa = ""
 txtUS_Genera = ""
 txtDetalle = ""
 lswAsiento.ListItems.Clear
End Sub

Private Sub chkConsultaDP_Click()
If chkConsultaDP.Value = vbChecked Then
  fraConsultaDP.Visible = True
  txtNumDp = ""
  lswDP.ListItems.Clear
Else
  fraConsultaDP.Visible = False
End If
End Sub

Private Sub cmdAS_Aceptar_Click()
Dim iRespuesta As Integer
iRespuesta = MsgBox("Esta seguro de realizar el traspaso a contabilidad", vbYesNo)
If iRespuesta = vbYes Then
 If chkAsientoResumen.Value = vbChecked Then
    Me.MousePointer = vbHourglass
        If chkAS_Depositos.Value = vbChecked Then Call sbGeneraAsientosResumen("DP")
        If chkAS_NC.Value = vbChecked Then Call sbGeneraAsientosResumen("NC")
        If chkAS_ND.Value = vbChecked Then Call sbGeneraAsientosResumen("ND")
        If chkAS_Recibos.Value = vbChecked Then Call sbGeneraAsientosResumen("RE")
    MsgBox "Se realizó el pase de asientos a contabilidad ", vbInformation
    Me.MousePointer = vbDefault
    Me.fraTraspaso.Visible = False

 Else
     Call sbGeneraAsientos
 End If
End If 'Respuesta
End Sub

Private Sub cmdAS_Cancelar_Click()
fraTraspaso.Visible = False
End Sub

Private Sub chkTodasLasFechas_Click()
If chkTodasLasFechas.Value = vbChecked Then
 dtpDesde.Enabled = False
 dtpHasta.Enabled = False
 chkFechaAnulacion.Value = 0
 chkFechaAnulacion.Enabled = False
 chkFechaEmision.Value = 0
 chkFechaEmision.Enabled = False
 chkFechaTraspaso.Value = 0
 chkFechaTraspaso.Enabled = False
Else
 dtpDesde.Enabled = True
 dtpHasta.Enabled = True
 chkFechaAnulacion.Enabled = True
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


Private Sub cmdCL_Click()
 Me.fraConfiguracion.Visible = False
End Sub

Private Function fxValidaTipoAsiento(vTipo As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
strSQL = "select coalesce(count(*),0) as Existe from tipos_asientos " _
      & " where TIPO_ASIENTO = '" & vTipo & "' and cod_empresa = " & GLOBALES.gEnlace
rsX.Open strSQL, glogon.Conection, adOpenStatic
fxValidaTipoAsiento = IIf((rsX!existe = 1), True, False)
rsX.Close
End Function


Private Function fxValidaCuenta(vCuenta As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
strSQL = "select coalesce(count(*),0) as Existe from cuentas " _
      & " where cod_cuenta = '" & vCuenta & "' and acepta_movimientos ='S' and cod_empresa = " & GLOBALES.gEnlace
rsX.Open strSQL, glogon.Conection, adOpenStatic
fxValidaCuenta = IIf((rsX!existe = 1), True, False)
rsX.Close
End Function


Private Sub cmdGuardaConfiguracion_Click()
Dim rs As New ADODB.Recordset, strSQL As String
Dim vPasa As Boolean
On Error GoTo vError

'Validacion
vPasa = fxValidaTipoAsiento(txtTA_RE)
If vPasa Then vPasa = fxValidaTipoAsiento(txtTA_NC)
If vPasa Then vPasa = fxValidaTipoAsiento(txtTA_ND)
If vPasa Then vPasa = fxValidaTipoAsiento(txtTA_DP)
If vPasa Then vPasa = fxValidaCuenta(medRE)
If vPasa Then vPasa = fxValidaCuenta(medNC)
If vPasa Then vPasa = fxValidaCuenta(medND)
If vPasa Then vPasa = fxValidaCuenta(medDP)

If Not vPasa Then
  MsgBox "Datos inválidos verifiquelos...", vbCritical
  Exit Sub
End If

'Guardar Configuracion
strSQL = "update ase_consecutivos set " _
       & "cs_nota_credito = " & txtID_NC _
       & ",cs_nota_debito = " & txtID_ND _
       & ",cs_deposito = " & txtID_DP _
       & ",cs_recibo = " & txtID_RE _
       & ",cs_utilizar_recibo = '" & IIf((chkUtilizaRecibo.Value = 1), "S", "N") _
       & "',cs_nc_cuenta = '" & medNC _
       & "',cs_nd_cuenta = '" & medND _
       & "',cs_dp_cuenta = '" & medDP _
       & "',cs_re_cuenta = '" & medRE _
       & "',cs_nc_asiento = '" & txtTA_NC _
       & "',cs_nd_asiento = '" & txtTA_ND _
       & "',cs_dp_asiento = '" & txtTA_DP _
       & "',cs_re_asiento = '" & txtTA_RE & "'"
glogon.Conection.Execute strSQL


MsgBox "Parámetros Guardados Satisfactoriamente", vbInformation


Exit Sub
vError:
 MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdImprimir_Click()
Dim vTipo As String
If (chkTodasLasFechas.Value + chkFechaAnulacion.Value + chkFechaEmision.Value _
    + chkFechaTraspaso.Value) = 0 Then
  MsgBox "No se ha especificado ninguna fecha como parámetro de busqueda...", vbInformation
  Exit Sub
End If

Me.MousePointer = vbHourglass

vTipo = fxTipoASEDoc(cboTipo.Text)

With frmContenedor.Crt
    .Reset
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes de Control de Documentos"
   If Me.chkPorUsuario.Value = vbUnchecked Then
    .ReportFileName = GLOBALES.gReportes + "\estados\ControlDocumentos.rpt"
   Else
    .ReportFileName = GLOBALES.gReportes + "\estados\ControlDocumentosUSER.rpt"
   End If

    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "usuario='" & glogon.Usuario & "'"
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
  
  Select Case True
   Case optReportes(0).Value  'Reporte General
     .Formulas(3) = "SUBTITULO='REPORTE GENERAL - " & UCase(cboTipo.Text) & "'"
       If chkFechaAnulacion.Value = vbChecked Then
          .SelectionFormula = "{ASE_DOCUMENTOS.FECHA_ANULACION} >= Date(" & fxFechaReportes(1) & ")" _
                    & " AND {ASE_DOCUMENTOS.FECHA_ANULACION} <= Date(" & fxFechaReportes(0) & ")" _
                    & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
          .Formulas(4) = "fecha_anulacion = 'Fecha Anulación entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
        
        End If
        If chkFechaEmision.Value = vbChecked Then
          .SelectionFormula = "{ASE_DOCUMENTOS.FECHA} in Date(" & fxFechaReportes(1) & ")" _
                    & " To Date(" & fxFechaReportes(0) & ")" _
                    & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
          .Formulas(4) = "fecha_emision = 'Fecha Emisión entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
        End If
        If chkFechaTraspaso.Value = vbChecked Then
          .SelectionFormula = "{ASE_DOCUMENTOS.FECHA_TRASPASO} >= Date(" & fxFechaReportes(1) & ")" _
                    & " AND {ASE_DOCUMENTOS.FECHA_TRASPASO} <= Date(" & fxFechaReportes(0) & ")" _
                    & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
          .Formulas(4) = "fecha_traspaso = 'Fecha Traspaso entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
        End If
        
   Case optReportes(1).Value  'Emitidos
     .SelectionFormula = "{ASE_DOCUMENTOS.ESTADO} = 'I'" _
                       & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
     .Formulas(3) = "SUBTITULO='" & UCase(cboTipo.Text) & " - EMITIDOS'"
   Case optReportes(2).Value  'Anulados
     .SelectionFormula = "{ASE_DOCUMENTOS.ESTADO} = 'A'" _
                       & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
     .Formulas(3) = "SUBTITULO='" & UCase(cboTipo.Text) & " - ANULADOS'"
   Case optReportes(3).Value  'Pendientes de Traspaso
     .SelectionFormula = "{ASE_DOCUMENTOS.TRASPASO} = 'P'" _
                       & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
     .Formulas(3) = "SUBTITULO='" & UCase(cboTipo.Text) & " - PENDIENTES TRSP.'"
   Case optReportes(4).Value  'Traspasos Generados
     .SelectionFormula = "{ASE_DOCUMENTOS.GENERADOS} = 'G'" _
                       & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
     .Formulas(3) = "SUBTITULO='" & UCase(cboTipo.Text) & " - TRSP. GENERADOS'"
  End Select

  If chkTodasLasFechas.Value = vbUnchecked And Not optReportes(0).Value Then
    If chkFechaAnulacion.Value = vbChecked Then
      .SelectionFormula = .SelectionFormula & " AND {ASE_DOCUMENTOS.FECHA_ANULACION} >= Date(" & fxFechaReportes(1) & ")" _
                & " AND {ASE_DOCUMENTOS.FECHA_ANULACION} <= Date(" & fxFechaReportes(0) & ")" _
                & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
      .Formulas(4) = "fecha_anulacion = 'Fecha Anulación entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
    End If
    If chkFechaEmision.Value = vbChecked Then
      .SelectionFormula = .SelectionFormula & " AND {ASE_DOCUMENTOS.FECHA} in Date(" & fxFechaReportes(1) & ")" _
                & " to Date(" & fxFechaReportes(0) & ")"
'                & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
      .Formulas(4) = "fecha_emision = 'Fecha Emisión entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
    End If
    If chkFechaTraspaso.Value = vbChecked Then
      .SelectionFormula = .SelectionFormula & " AND {ASE_DOCUMENTOS.FECHA_TRASPASO} >= Date(" & fxFechaReportes(1) & ")" _
                & " AND {ASE_DOCUMENTOS.FECHA_TRASPASO} <= Date(" & fxFechaReportes(0) & ")" _
                & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
      .Formulas(4) = "fecha_traspaso = 'Fecha Traspaso entre " & Format(dtpDesde.Value, "dd/mm/yyyy") & " y " & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"
    End If
   
   End If
   
   If chkUsuarioActual.Value = vbChecked Then
     .SelectionFormula = .SelectionFormula & " AND {ASE_DOCUMENTOS.USUARIO} = '" _
                       & glogon.Usuario & "'"
   End If
   
   
   .PrintReport
   
End With

Me.MousePointer = vbDefault

End Sub

Sub CreaDetalleAsiento(strTipo As String, strCaso As String, strCuenta As String, vFecha As Date, curMonto As Currency, DH As String, intLinea As Integer)
Dim strSQL As String, strNumero_Asiento As String

strNumero_Asiento = strTipo & "-" & Format(Mid(Trim(strCaso), 1, 5), "00000#") & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")


If UCase(DH) <> "D" Then 'dc - dh
 strSQL = "insert asientos_detalle(cod_empresa,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_dedito,monto_credito,detalle,documento)" _
        & " values(" & GLOBALES.gEnlace & ",'AS','" & strNumero_Asiento & "'," & intLinea & "," & Trim(strCuenta) & ",0," _
        & curMonto & ",'" & strCaso & "','" & strTipo & "')"
Else
 strSQL = "insert asientos_detalle(cod_empresa,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_dedito,monto_credito,detalle,documento)" _
        & " values(" & GLOBALES.gEnlace & ",'AS','" & strNumero_Asiento & "'," & intLinea & "," & Trim(strCuenta) & "," & curMonto & ",0,'" _
        & strCaso & "','" & strTipo & "')"
End If

glogon.Conection.Execute strSQL

End Sub

Private Sub sbBuscaConfiguracion(vObj As Object, vTipo As String)

If UCase(vTipo) = "C" Then
    gBusquedas.Columna = "cod_cuenta"
    gBusquedas.Consulta = "select cod_cuenta,descripcion from cuentas"
    gBusquedas.Orden = "cod_cuenta"
    gBusquedas.Filtro = " and acepta_movimientos = 'S' and cod_empresa = " & GLOBALES.gEnlace
Else
    gBusquedas.Columna = "tipo_asiento"
    gBusquedas.Consulta = "select tipo_asiento,descripcion from tipos_asientos"
    gBusquedas.Orden = "tipo_asiento"
    gBusquedas.Filtro = " and cod_empresa = " & GLOBALES.gEnlace
End If

frmBusquedas.Show vbModal
vObj.Text = gBusquedas.Resultado
SendKeys "{TAB}"

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 tipo,id_documento from ase_documentos" _
           & " where Tipo = '" & fxTipoASEDoc(cboTipo.Text) & "'"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and id_documento > " & txtDocumento & " order by id_documento asc"
    Else
       strSQL = strSQL & " and id_documento < " & txtDocumento & " order by id_documento desc"
    End If
    
    rs.Open strSQL, glogon.Conection, adOpenStatic
    If Not rs.EOF And Not rs.BOF Then
      txtDocumento = rs!id_documento
      Call txtDocumento_KeyPress(vbKeyReturn)
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
  MsgBox "Consulte a Su Administrador de Base de Datos, sobre Transacciones con TOP y Record Count", vbInformation


End Sub


Private Sub Form_Load()
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 dtpDesde.Value = fxFechaServidor
 dtpHasta.Value = dtpDesde.Value
 vModulo = 10 'Cuentas Corrientes
 cboTipo.AddItem "Recibo"
 cboTipo.AddItem "Nota Credito"
 cboTipo.AddItem "Nota Debito"
 cboTipo.AddItem "Depositos"
 
 cboTipo.Text = "Recibo"

 
 Call Formularios(Me)
 Call RefrescaTags(Me)
End Sub

Sub sbAsientoUnoAUno(rs As ADODB.Recordset)
Dim rs2 As New ADODB.Recordset, DH As String
Dim strSQL As String, intLinea As Integer, vCuenta As String
Dim vFecha As Date

On Error GoTo vError
 strSQL = "select CS_RE_CUENTA from ase_consecutivos"
 rs2.Open strSQL, glogon.Conection, adOpenStatic
 vCuenta = rs2!cs_re_cuenta
 rs2.Close

 vFecha = fxFechaServidor

 If fxValidaPeriodoAsiento(rs!Fecha) Then 'Verificar el Periodo Abierto en contabilidad
   'Crea Maestro
   strSQL = "insert asientos(cod_empresa,tipo_asiento,num_asiento,anio,mes,fecha_asiento,descripcion,balanceado)" _
          & " values(" & GLOBALES.gEnlace & ",'" & fxTipoAsientoDoc(rs!Tipo) & "','A-" & Format(rs!id_documento, "00000000") & "'," & Year(vFecha) _
          & "," & Month(vFecha) & ",'" & Format(vFecha, "yyyy/mm/dd") & "','" & rs!concepto & "','S')"
   glogon.Conection.Execute strSQL
    
    
    'Crea Detalle
    strSQL = "insert asientos_detalle(cod_empresa,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito,detalle,documento)" _
          & " values(" & GLOBALES.gEnlace & ",'" & fxTipoAsientoDoc(rs!Tipo) & "','A-" & Format(rs!id_documento, "00000000") & "',1," & Trim(vCuenta) & ",1," _
          & "0,'" & rs!concepto & "','" & Format(rs!id_documento, "00000000") & "')"
    glogon.Conection.Execute strSQL

    strSQL = "insert asientos_detalle(cod_empresa,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito,detalle,documento)" _
          & " values(" & GLOBALES.gEnlace & ",'" & fxTipoAsientoDoc(rs!Tipo) & "','A-" & Format(rs!id_documento, "00000000") & "',2," & Trim(vCuenta) & ",0," _
          & "1,'" & rs!concepto & "','" & Format(rs!id_documento, "00000000") & "')"
    glogon.Conection.Execute strSQL

    
    'Borra el Asiento del Auxiliar ASE
     strSQL = "DELETE ASE_ASIENTOS where tipo = '" & rs!Tipo & "' and id_documento = " & rs!id_documento
     glogon.Conection.Execute strSQL
    
    'Crea Nuevo Asiento en el Auxiliar ASE
     strSQL = "Insert ASE_ASIENTOS(ID_DOCUMENTO, Tipo, RECAS_CUENTA, RECAS_MONTO,RECAS_DEBEHABER)" _
            & " VALUES(" & rs!id_documento & ",'" & rs!Tipo & "','" & vCuenta & "',1,'D')"
     glogon.Conection.Execute strSQL
     
     strSQL = "Insert ASE_ASIENTOS(ID_DOCUMENTO, Tipo, RECAS_CUENTA, RECAS_MONTO,RECAS_DEBEHABER)" _
            & " VALUES(" & rs!id_documento & ",'" & rs!Tipo & "','" & vCuenta & "',1,'H')"
     glogon.Conection.Execute strSQL
    
 Else
  MsgBox "Existen asientos que no pueden ser trasladados porque el periodo fué cerrado..."
 End If 'Periodo

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub


Sub sbAsientoReversion(rs As ADODB.Recordset)
Dim rs2 As New ADODB.Recordset, DH As String
Dim strSQL As String, intLinea As Integer

On Error GoTo vError

 If fxValidaPeriodoAsiento(rs!Fecha) Then 'Verificar el Periodo Abierto en contabilidad
   'Crea Maestro
   strSQL = "insert asientos(cod_empresa,tipo_asiento,num_asiento,anio,mes,fecha_asiento,descripcion,balanceado)" _
          & " values(" & GLOBALES.gEnlace & ",'" & fxTipoAsientoDoc(rs!Tipo) & "','A-" & Format(rs!id_documento, "00000000") & "'," & Year(fxFechaServidor) & "," & Month(fxFechaServidor) _
          & ",'" & Format(fxFechaServidor, "yyyy/mm/dd") & "','" & rs!concepto & "','S')"
   glogon.Conection.Execute strSQL
    
    
    
    'Crea Detalle
    intLinea = 1
    rs2.CursorLocation = adUseServer
    rs2.Open "select * from ase_asientos where id_documento = " & rs!id_documento _
             & " and tipo = '" & rs!Tipo & "'", glogon.Conection, adOpenStatic
    Do While rs2.EOF = False
        If UCase(rs2!RECAS_DEBEHABER) = "H" Then  'dc - dh
          DH = "C"
        Else
          DH = rs2!RECAS_DEBEHABER
        End If
        
       'LE DOY VUELTA PARA REVERSAR
        
        If DH = "D" Then 'Acreditar
            strSQL = "insert asientos_detalle(cod_empresa,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito,detalle,documento)" _
                   & " values(" & GLOBALES.gEnlace & ",'" & fxTipoAsientoDoc(rs!Tipo) & "','A-" & Format(rs2!id_documento, "00000000") & "'," & intLinea _
                   & "," & Trim(rs2!RECAS_CUENTA) & ",0," & rs2!RECAS_MONTO & ",'" & rs!concepto & "','" & Format(rs2!id_documento, "00000000") & "')"
        
        Else 'Debitar
            strSQL = "insert asientos_detalle(cod_empresa,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito,detalle,documento)" _
                   & " values(" & GLOBALES.gEnlace & ",'" & fxTipoAsientoDoc(rs!Tipo) & "','A-" & Format(rs2!id_documento, "00000000") & "'," & intLinea _
                   & "," & Trim(rs2!RECAS_CUENTA) & "," & rs2!RECAS_MONTO & ",0,'" & rs!concepto & "','" & Format(rs2!id_documento, "00000000") & "')"
        End If
        
        glogon.Conection.Execute strSQL
        intLinea = intLinea + 1
        rs2.MoveNext
    Loop
    rs2.Close

 Else
  MsgBox "Existen asientos que no pueden ser trasladados porque el periodo fué cerrado..."
 End If 'Periodo

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Sub sbAnulaDocumento()
'Verifica el traspaso y Genera ASiento de Reversion si se encuentra
'en el mismo mes de emision, tambien tiene que verificar el estado del
'recibo, ademas actualiza el estado,fechaanulacion y el usuario
Dim rs As New ADODB.Recordset, strSQL As String, rs3 As New ADODB.Recordset
Dim strInforma As String, rs2 As New ADODB.Recordset, vTipoDoc As String
Dim mFechaSistema As Date, iRespuesta As Integer

On Error GoTo vError

vTipoDoc = fxTipoASEDoc(cboTipo.Text)

If vTipoDoc <> "RE" Then
  MsgBox "Este Tipo de Documento no se puede Anular, debe recurrir a otro método", vbCritical
  Exit Sub
End If

mFechaSistema = fxFechaServidor

strSQL = "Select * from ase_documentos where id_documento = " & txtDocumento _
       & " and tipo = '" & vTipoDoc & "'"
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF Then
 If rs!Estado = "I" And Format(rs!Fecha, "yyyy/mm/dd") = Format(mFechaSistema, "yyyy/mm/dd") Then
   strSQL = "update ase_documentos set estado = 'A',fecha_anulacion = '" _
          & Format(mFechaSistema, "yyyy/mm/dd") & "',us_anula = '" _
          & glogon.Usuario & "' where id_documento = " & txtDocumento _
          & " and tipo = '" & vTipoDoc & "'"
   glogon.Conection.Execute strSQL
   
   If rs!traspaso = "G" Then
     'Ya se generó a Contabilidad, verificar si fue durante el mes
     'y hacer asiento de reversión, de lo contrario informar al usuario
     If Year(rs!Fecha) = Year(mFechaSistema) _
           And Month(rs!Fecha) = Month(mFechaSistema) Then
       Call sbAsientoReversion(rs)
       strInforma = " - Asiento de Reversión Registrado en contabilidad..."
     Else
        strInforma = "Documento se emitió en un mes anterior, se tiene que reportar a contabilidad" _
               & " para corrección contable..."
     End If
   
   Else 'Traspaso
       
       Call sbAsientoUnoAUno(rs)
   
   End If
 
     'Anular Movimiento en REG_CREDITOS,CREDITOS_DT Y MOROSIDAD
     'Siempre y Cuando no sea porque el comprobante se duplico, de lo contrario
     'No hace falta reversar el movimiento, ya que solo se ejecuto uno vez
       
   iRespuesta = MsgBox("Desea Reversar movimientos en el Estado de Cuenta", vbYesNo)
   If iRespuesta = vbYes Then
     
      rs3.CursorLocation = adUseServer
      Select Case Mid(txtConcepto, 1, 7)
         Case "ABONO O" 'Abono ordinario
            strSQL = "select * from creditos_dt where estado = 'A' and tcon = " _
                   & fxTipoASENumero(vTipoDoc) & " and ncon =" & txtDocumento
            rs3.Open strSQL, glogon.Conection, adOpenStatic
            If Not rs3.EOF And Not rs3.BOF Then
               strSQL = "delete creditos_dt where consec = " & rs3!consec
               glogon.Conection.Execute strSQL
               strSQL = "update reg_creditos set " _
                      & "estado = 'A'," _
                      & "saldo = saldo + " & rs3!amortiza & "," _
                      & "saldo_mes = saldo_mes + " & rs3!amortiza & "," _
                      & "amortiza = amortiza - " & rs3!amortiza & "," _
                      & "interesc = interesc - " & rs3!intcp _
                      & " where id_solicitud = " & rs3!Id_solicitud
               glogon.Conection.Execute strSQL
            End If
            rs3.Close
            
         Case "ABONO E" 'Abono Extraordinario
         
            strSQL = "select * from creditos_dt where estado = 'A' and tcon = " _
                   & fxTipoASENumero(vTipoDoc) & " and ncon =" & txtDocumento
            rs3.Open strSQL, glogon.Conection, adOpenStatic
            If Not rs3.EOF And Not rs3.BOF Then
               strSQL = "delete creditos_dt where consec = " & rs3!consec
               glogon.Conection.Execute strSQL
               rs2.Open "select fecult from reg_Creditos where id_solicitud = " & rs3!Id_solicitud, glogon.Conection, adOpenStatic
               strSQL = "update reg_creditos set " _
                      & "estado = 'A',fecult = " & fxFechaProcesoAnterior(rs2!fecult) & "," _
                      & "saldo = saldo + " & rs3!amortiza & "," _
                      & "amortiza = amortiza - " & rs3!amortiza _
                      & " where id_solicitud = " & rs3!Id_solicitud
               glogon.Conection.Execute strSQL
               rs2.Close
            End If
            rs3.Close
         
         Case "ABONO A" 'morosidad
         
            strSQL = "select coalesce(sum(abintc),0) as intc, coalesce(sum(abintm),0) as intm, " _
                   & "coalesce(sum(abamortiza),0) as amortiza,id_solicitud  from morosidad " _
                   & "where estado = 'C' and tcon = " _
                   & fxTipoASENumero(vTipoDoc) & " and ncon =" & txtDocumento _
                   & " group by id_solicitud"

            rs3.Open strSQL, glogon.Conection, adOpenStatic
            If Not rs3.EOF And Not rs3.BOF Then
               strSQL = "update morosidad set estado = 'A',abintc=0,abintm=0" _
                      & ",abamortiza = 0 where tcon = " _
                      & fxTipoASENumero(vTipoDoc) & " and ncon =" & txtDocumento
               glogon.Conection.Execute strSQL
               strSQL = "update reg_creditos set " _
                      & "estado = 'A'," _
                      & "saldo = saldo + " & rs3!amortiza & "," _
                      & "saldo_mes = saldo_mes + " & rs3!amortiza & "," _
                      & "amortiza = amortiza - " & rs3!amortiza & "," _
                      & "interesc = interesc - " & rs3!intc + rs3!intm _
                      & " where id_solicitud = " & rs3!Id_solicitud
               glogon.Conection.Execute strSQL
            End If
            rs3.Close
         
         
         
         Case "ARREGLO" 'arreglo de pago
                   
            strSQL = "select id_solicitud,sum(abintc) + sum(abintm) as interes from morosidad " _
                    & "where estado = 'C' and tcon = " _
                    & fxTipoASENumero(vTipoDoc) & " and ncon =" & txtDocumento _
                    & " group by id_solicitud"
            rs3.Open strSQL, glogon.Conection, adOpenStatic
            If Not rs3.EOF And Not rs3.BOF Then
               strSQL = "delete morosidad where estado = 'A' and id_solicitud = " & rs3!Id_solicitud
               glogon.Conection.Execute strSQL
               
               strSQL = "update morosidad set estado = 'A',abintc=0,abintm=0" _
                      & ",abamortiza = 0 where tcon = " _
                      & fxTipoASENumero(vTipoDoc) & " and ncon =" & txtDocumento
               glogon.Conection.Execute strSQL
               
               
               strSQL = "update reg_creditos set " _
                      & "estado = 'A'," _
                      & "interesc = interesc - " & rs3!Interes _
                      & " where id_solicitud = " & rs3!Id_solicitud
               glogon.Conection.Execute strSQL
            End If
            rs3.Close
      
      End Select
      
    End If 'Reversar movimientos en el estado de cuenta
    
      MsgBox "- Documento Anulado " & vbCrLf & strInforma, vbInformation
      Call sbCargaDocumento(cboTipo.Text, txtDocumento)
      Call Bitacora("Anula", cboTipo.Text & " #" & txtDocumento)
 
 Else
   MsgBox "Este Documento ya fue anulado o se emitió un día diferente al actual...", vbInformation
 End If

End If

rs.Close

Exit Sub
vError:
MsgBox Err.Description, vbCritical


End Sub



Sub sbGeneraAsientosResumen(vTipoDocumento As String)
Dim rs As New ADODB.Recordset, strSQL As String
Dim intLinea As Integer, DH As String, vTipoAsiento As String
Dim rsTmp As New ADODB.Recordset, vFecha As Date
Dim lngInicio As Long, lngCorte As Long

lblEstatus.Caption = "Cargando Información..."
lblEstatus.Refresh
prgBar.Value = 1

On Error GoTo vError

vFecha = fxFechaServidor

'Determina el tipo de Asiento para Contabilidad

vTipoAsiento = fxTipoAsientoDoc(vTipoDocumento)

'Inicia Transaccion
glogon.Conection.BeginTrans

'Sacar los Documentos de Inicio y Corte

strSQL = "select fecha,coalesce(min(id_documento),0) as Inicio, coalesce(max(id_documento),0) as Corte" _
       & " from ase_documentos where estado = 'I' and traspaso = 'P'" _
       & " and tipo = '" & vTipoAsiento & "'" _
       & " group by fecha"
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
lblEstatus.Caption = "Procesando Asientos..." & vTipoDocumento
lblEstatus.Refresh

Do While Not rs.EOF
 
 If fxValidaPeriodoAsiento(rs!Fecha) Then 'Verificar el Periodo Abierto en contabilidad
    
    
    lblEstatus.Caption = "Procesando Asientos..." & vTipoDocumento & "[" & rs!inicio & "-" & rs!corte & "]"
    lblEstatus.Refresh
    
    
    'Crea el Maestro de Asiento
    strSQL = "insert asientos(cod_empresa,tipo_asiento,num_asiento,anio,mes,fecha_asiento,descripcion,balanceado)" _
           & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','SIF" & rs!inicio & "-" & rs!corte & "'," & Year(rs!Fecha) & "," & Month(rs!Fecha) _
           & ",'" & Format(rs!Fecha, "yyyy/mm/dd") & "','RESUMEN DE DOCUMENTOS','S')"
    glogon.Conection.Execute strSQL
        
    intLinea = 1
    
    'Crea El Detalle del Asiento Resumen
    strSQL = "select recas_cuenta as cuenta,sum(recas_monto) as Monto,recas_debehaber as DH" _
           & " from ase_asientos where id_documento between " & rs!inicio & " and " & rs!corte _
           & " and tipo = '" & vTipoDocumento & "' group by recas_cuenta,recas_debehaber"
    rsTmp.Open strSQL, glogon.Conection, adOpenStatic
    Do While Not rsTmp.EOF
        If UCase(rsTmp!DH) = "H" Then  'dc - dh
          DH = "C"
          strSQL = "insert asientos_detalle(cod_empresa,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito,detalle,documento)" _
                 & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','SIF" & rs!inicio & "-" & rs!corte & "'," & intLinea & ",'" & Trim(rsTmp!Cuenta) _
                 & "',0," & rsTmp!Monto & ",'RESUMEN DE DOCUMENTOS SIF','" & rs!inicio & "-" & rs!corte & "')"
        
        Else
          DH = rsTmp!DH
          strSQL = "insert asientos_detalle(cod_empresa,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito,detalle,documento)" _
                 & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','SIF" & rs!inicio & "-" & rs!corte & "'," & intLinea & ",'" & Trim(rsTmp!Cuenta) _
                 & "'," & rsTmp!Monto & ",0,'RESUMEN DE DOCUMENTOS SIF','" & rs!inicio & "-" & rs!corte & "')"
        End If
        
        glogon.Conection.Execute strSQL
        intLinea = intLinea + 1
      
      rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    
    'Actualiza Tabla de ASE_DOCUMENTOS
    strSQL = "Update ase_documentos set traspaso = 'G', FECHA_TRASPASO = '" & Format(vFecha, "yyyy/mm/dd") _
            & "',us_traspaso = '" & glogon.Usuario & "' where id_documento between " & rs!inicio _
            & " and " & rs!corte & " and tipo = '" & vTipoDocumento & "'"
    glogon.Conection.Execute strSQL
    
    
 Else 'Verificacion del periodo
  
   MsgBox "Existen asientos que no pueden ser trasladados porque el periodo fué cerrado...", vbInformation
  
  
 End If 'Verificacion del periodo
 
 If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

'Cierra Transaccion
glogon.Conection.CommitTrans


lblEstatus.Caption = ""
lblEstatus.Refresh
prgBar.Value = 1

Call Bitacora("Aplica", "Asientos del Control de Documentos ASE:" & vTipoDocumento)

Exit Sub

vError:
    lblEstatus.Caption = ""
    lblEstatus.Refresh
    prgBar.Value = 1
    Me.MousePointer = vbDefault
    glogon.Conection.RollbackTrans
    MsgBox Err.Description, vbCritical
End Sub


Sub sbGeneraAsientos()
Dim rs As New ADODB.Recordset, strSQL As String
Dim intLinea As Integer, DH As String, strDocumentos As String
Dim rs2 As New ADODB.Recordset, vTipoAsiento As String
Dim vFecha As Date, vDetalle As String

Me.MousePointer = vbHourglass
Me.fraTraspaso.Visible = True

lblEstatus.Caption = "Cargando Información..."
lblEstatus.Refresh
prgBar.Value = 1

On Error GoTo vError

strDocumentos = ""
vFecha = fxFechaServidor


If chkAS_Depositos.Value = vbChecked Then strDocumentos = "'DP'"

If chkAS_NC.Value = vbChecked Then
  If Len(strDocumentos) > 1 Then
    strDocumentos = strDocumentos + ",'NC'"
  Else
    strDocumentos = "'NC'"
  End If
End If

If chkAS_ND.Value = vbChecked Then
  If Len(strDocumentos) > 1 Then
    strDocumentos = strDocumentos + ",'ND'"
  Else
    strDocumentos = "'ND'"
  End If
End If

If chkAS_Recibos.Value = vbChecked Then
  If Len(strDocumentos) > 1 Then
    strDocumentos = strDocumentos + ",'RE'"
  Else
    strDocumentos = "'RE'"
  End If
End If


If Len(strDocumentos) = 0 Then
  Me.MousePointer = vbDefault
  Exit Sub
End If


strSQL = "select * from ase_documentos where estado = 'I' and traspaso = 'P'" _
       & " and tipo in(" & strDocumentos & ")"
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
lblEstatus.Caption = "Procesando Asientos..."
lblEstatus.Refresh

Do While Not rs.EOF
 If fxValidaPeriodoAsiento(rs!Fecha) Then 'Verificar el Periodo Abierto en contabilidad
    'Crea Maestro
   vTipoAsiento = fxTipoAsientoDoc(rs!Tipo)
   strSQL = "insert asientos(cod_empresa,tipo_asiento,num_asiento,anio,mes,fecha_asiento,descripcion,balanceado)" _
          & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','SIF" & Format(rs!id_documento, "00000000") & "'," & Year(rs!Fecha) & "," & Month(rs!Fecha) _
          & ",'" & Format(rs!Fecha, "yyyy/mm/dd") & "','" & rs!concepto & "','S')"
   glogon.Conection.Execute strSQL
    
    'Crea Detalle
    intLinea = 1
    rs2.CursorLocation = adUseServer
    rs2.Open "select * from ase_asientos where id_documento = " & rs!id_documento _
             & " and tipo = '" & rs!Tipo & "'", glogon.Conection, adOpenStatic
    Do While Not rs2.EOF
        If UCase(rs2!RECAS_DEBEHABER) = "H" Then  'dc - dh
          DH = "C"
        Else
          DH = rs2!RECAS_DEBEHABER
        End If
        'Ahora se pone en el detalle de la cuenta el numero de deposito y luego
        'Lo que alcance del concepto
        vDetalle = ""
        If IsNull(rs!dp) Then
          vDetalle = rs!concepto
        Else
          If Trim(rs!dp) = "" Then
              vDetalle = rs!concepto
          Else
            vDetalle = "DP." & Trim(rs!dp) & " - " & rs!concepto
          End If
        End If
        vDetalle = Mid(vDetalle, 1, 59)
        
        If DH = "C" Then 'Acredita
            strSQL = "insert asientos_detalle(cod_empresa,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito,detalle,documento)" _
                & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','SIF" & Format(rs2!id_documento, "00000000") & "'," & intLinea & "," & Trim(rs2!RECAS_CUENTA) _
                & ",0," & rs2!RECAS_MONTO & ",'" & vDetalle & "','" & Format(rs2!id_documento, "00000000") & "')"
        Else 'Debita
            strSQL = "insert asientos_detalle(cod_empresa,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito,detalle,documento)" _
                & " values(" & GLOBALES.gEnlace & ",'" & vTipoAsiento & "','SIF" & Format(rs2!id_documento, "00000000") & "'," & intLinea & "," & Trim(rs2!RECAS_CUENTA) _
                & "," & rs2!RECAS_MONTO & ",0,'" & vDetalle & "','" & Format(rs2!id_documento, "00000000") & "')"
        End If
        If Len(Trim(rs2!RECAS_CUENTA)) > 0 Then
          glogon.Conection.Execute strSQL
          intLinea = intLinea + 1
        End If
        rs2.MoveNext
    Loop
    rs2.Close
    
    'Actualizar el estado del recibo
    strSQL = "Update ase_documentos set traspaso = 'G', FECHA_TRASPASO = '" & Format(vFecha, "yyyy/mm/dd") _
            & "',us_traspaso = '" & glogon.Usuario & "' where id_documento = " & rs!id_documento _
            & " and tipo = '" & rs!Tipo & "'"
    glogon.Conection.Execute strSQL
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

Call Bitacora("Aplica", "Asientos del Control de Documentos ASE")

MsgBox "Se realizó el pase de asientos a contabilidad ", vbInformation
Me.MousePointer = vbDefault
Me.fraTraspaso.Visible = False

Exit Sub

vError:
    lblEstatus.Caption = ""
    lblEstatus.Refresh
    prgBar.Value = 1
    Me.MousePointer = vbDefault
    Me.fraTraspaso.Visible = False
    MsgBox Err.Description, vbCritical


End Sub

Private Sub sbConfiguracion()
Dim rs As New ADODB.Recordset, strSQL As String

fraConfiguracion.Visible = True
On Error Resume Next
medND.Format = GLOBALES.gstrMascara
medNC.Format = GLOBALES.gstrMascara
medDP.Format = GLOBALES.gstrMascara
medRE.Format = GLOBALES.gstrMascara


rs.Open "select * from ASE_CONSECUTIVOS", glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
  'INSERTAR
  strSQL = "insert ase_consecutivos(cs_nota_credito,cs_nota_debito,cs_deposito," _
         & "cs_recibo,cs_utilizar_recibo,cs_nd_cuenta,cs_nc_cuenta,cs_dp_cuenta" _
         & ",cs_re_cuenta,cs_nc_asiento,cs_nd_asiento,cs_dp_asiento,cs_re_asiento) " _
         & "values(1,1,1,1,'N','','','','','','','','')"
  glogon.Conection.Execute strSQL
End If
rs.Close

rs.Open "select * from ASE_CONSECUTIVOS", glogon.Conection, adOpenStatic
medND.Text = Trim(rs!cs_nd_cuenta)
medNC.Text = Trim(rs!cs_nc_cuenta)
medDP.Text = Trim(rs!cs_dp_cuenta)
medRE.Text = Trim(rs!cs_re_cuenta)
txtTA_ND = Trim(rs!cs_nd_asiento)
txtTA_NC = Trim(rs!cs_nc_asiento)
txtTA_DP = Trim(rs!cs_dp_asiento)
txtTA_RE = Trim(rs!cs_re_asiento)
txtID_NC = rs!cs_nota_credito
txtID_ND = rs!cs_nota_debito
txtID_DP = rs!cs_deposito
txtID_RE = rs!cs_recibo
chkUtilizaRecibo.Value = IIf((UCase(rs!cs_utilizar_recibo) = "S"), 1, 0)
rs.Close


End Sub


Private Sub imgReImpresion_Click()
Dim strSQL As String, X As New clsImpresoras
Dim vDriver, vTipo As String

Me.MousePointer = vbHourglass

vTipo = fxTipoASEDoc(cboTipo.Text)

With frmContenedor.Crt
   .Reset
   If vTipo = "RE" Then
     X.TipoImpresora = Recibos
     X.Reset
     .PrinterDriver = X.Controlador
     .PrinterName = X.Nombre
     .PrinterPort = X.Puerto
     .Destination = crptToPrinter
     .ReportFileName = GLOBALES.gReportes & "\Estados\Documento.rpt"
     .SelectionFormula = "{ASE_DOCUMENTOS.ID_DOCUMENTO} = " & Trim(txtDocumento) _
                       & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
   
   Else
     .WindowShowPrintSetupBtn = True
     .WindowShowRefreshBtn = True
     .WindowShowSearchBtn = True
     .WindowState = crptMaximized
     .WindowTitle = "Reportes de Control de Documentos"
     .ReportFileName = GLOBALES.gReportes & "\Estados\DocumentoNotas.rpt"
     .SelectionFormula = "{ASE_DOCUMENTOS.ID_DOCUMENTO} = " & Trim(txtDocumento) _
                       & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "' and {CUENTAS.COD_EMPRESA} = " & GLOBALES.gEnlace
   
   End If
   .PrintReport
End With

strSQL = "Update ASE_DOCUMENTOS set Estado='I' Where Id_DOCUMENTO=" & Trim(txtDocumento) _
         & " AND TIPO = '" & vTipo & "'"
glogon.Conection.Execute strSQL


Me.MousePointer = vbDefault

End Sub

Private Sub medDP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(medDP, "C")
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub medNC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(medNC, "C")
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub



Private Sub medND_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(medND, "C")
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub medRE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(medRE, "C")
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub


Private Sub tlbRecibos_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim iRespuesta As Integer
Select Case Button.Key
  Case "anular"
     iRespuesta = MsgBox("Esta seguro de que desea anular el " & cboTipo.Text & " #" & txtDocumento, vbYesNo)
     If iRespuesta = vbYes Then
      Call sbAnulaDocumento
     End If
  Case "traspaso"
      fraTraspaso.Visible = True
  Case "configuracion"
     Call sbConfiguracion
  Case "reportes"
    fraReportes.Visible = True
End Select
End Sub

Private Sub tlbRecibos_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim strSQL As String, X As New clsImpresoras
Dim vDriver, vTipo As String

vTipo = fxTipoASEDoc(cboTipo.Text)

Select Case ButtonMenu.Key
 Case "repReImpresion"
                           
       With frmContenedor.Crt
          .Reset
          If vTipo = "RE" Then
            X.TipoImpresora = Recibos
            X.Reset
            .PrinterDriver = X.Controlador
            .PrinterName = X.Nombre
            .PrinterPort = X.Puerto
            .Destination = crptToPrinter
            .ReportFileName = GLOBALES.gReportes & "\Estados\Documento.rpt"
          Else
            .WindowShowPrintSetupBtn = True
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            .WindowState = crptMaximized
            .WindowTitle = "Reportes de Control de Documentos"
            .ReportFileName = GLOBALES.gReportes & "\Estados\DocumentoNotas.rpt"
          End If
          .SelectionFormula = "{ASE_DOCUMENTOS.ID_DOCUMENTO} = " & Trim(txtDocumento) _
                            & " AND {ASE_DOCUMENTOS.TIPO} = '" & vTipo & "'"
          .PrintReport
       End With
       
       strSQL = "Update ASE_DOCUMENTOS set Estado='I' Where Id_DOCUMENTO=" & Trim(txtDocumento) _
                & " AND TIPO = '" & vTipo & "'"
       glogon.Conection.Execute strSQL
       
 Case "repRecibosFecha"
   fraReportes.Visible = True
End Select
End Sub

Private Sub sbCargaDocumento(vTipo As String, lngDocumento As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListItem, curDebe As Currency, curHaber As Currency
Dim strTipo As String

curDebe = 0
curHaber = 0

strTipo = fxTipoASEDoc(vTipo)

On Error Resume Next

strSQL = "select * from ase_Documentos where id_Documento = " & lngDocumento _
        & " and tipo = '" & strTipo & "'"
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs.EOF And rs.BOF Then
  rs.Close
  MsgBox "No se encontró Documento", vbCritical
  Exit Sub
End If
 txtBeneficiario = IIf(IsNull(rs!cliente), "", rs!cliente)
 Select Case rs!Estado
   Case "I" 'Impreso
      txtEstado = "Impreso"
   Case "P" 'Pendiente
      txtEstado = "Pendiente"
   Case "A" 'Anulado
      txtEstado = "Anulado"
 End Select
 txtFechaAnula = IIf(IsNull(rs!fecha_anulacion), "", rs!fecha_anulacion)
 txtFechaGenera = rs!Fecha
 txtFechaTraspasa = IIf(IsNull(rs!fecha_traspaso), "", rs!fecha_traspaso)
 txtConcepto = rs!concepto
 txtMonto = Format(rs!Monto, "###,###,###,##0.00")
 Select Case rs!tipo_pago
   Case "E"
     txtPago = "Efectivo"
   Case "M"
     txtPago = "Mixto"
   Case "C"
     txtPago = "Cheque"
   Case "D"
     txtPago = "Depósito"
 End Select

 txtTipo = rs!tipo_pago
 txtDetalle = IIf(IsNull(rs!Detalle), "", rs!Detalle)
 txtUS_Anula = IIf(IsNull(rs!us_anula), "", rs!us_anula)
 txtUS_Traspasa = IIf(IsNull(rs!us_traspaso), "", rs!us_traspaso)
 txtUS_Genera = rs!Usuario
rs.Close

strSQL = "select * from ase_asientos where id_documento=" & lngDocumento _
        & " and tipo = '" & strTipo & "'"
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic

lswAsiento.ListItems.Clear
With lswAsiento
Do While Not rs.EOF
 Set itmX = .ListItems.Add(.ListItems.Count + 1, , Format(rs!RECAS_CUENTA, GLOBALES.gstrMascara))
   itmX.SubItems(1) = fxDescribeCuenta(Trim(rs!RECAS_CUENTA))
   If rs!RECAS_DEBEHABER = "D" Then
      itmX.SubItems(2) = Format(rs!RECAS_MONTO, "###,###,###,##0.00")
      curDebe = curDebe + rs!RECAS_MONTO
   Else
      itmX.SubItems(3) = Format(rs!RECAS_MONTO, "###,###,###,##0.00")
      curHaber = curHaber + rs!RECAS_MONTO
   End If
 rs.MoveNext
Loop
 Set itmX = .ListItems.Add(.ListItems.Count + 1, , "")
  itmX.SubItems(2) = "---------------------------"
  itmX.SubItems(3) = "---------------------------"
 
 Set itmX = .ListItems.Add(.ListItems.Count + 1, , "TOTALES")
  itmX.SubItems(2) = Format(curDebe, "###,###,###,##0.00")
  itmX.SubItems(3) = Format(curHaber, "###,###,###,##0.00")
End With
rs.Close
End Sub


Private Sub txtDocumento_Change()
 Call sbLimpiaDatos
End Sub

Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
 Call sbCargaDocumento(cboTipo.Text, txtDocumento)
End If
End Sub

Private Sub txtID_DP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtID_NC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtID_ND_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtID_RE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub


Private Sub txtNumDp_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 'Nada
Else
 Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "select * from ase_documentos where dp = '" & txtNumDp & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
lswDP.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lswDP.ListItems.Add(, , rs!Tipo)
   itmX.SubItems(1) = rs!id_documento
   itmX.SubItems(2) = rs!concepto & ""
   itmX.SubItems(3) = rs!cliente & ""
   itmX.SubItems(4) = Format(rs!Fecha, "dd/mm/yyyy")
   itmX.SubItems(5) = rs!dp & ""
 rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

End Sub

Private Sub txtTA_DP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(txtTA_DP, "T")
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtTA_NC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(txtTA_NC, "T")
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtTA_ND_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(txtTA_ND, "T")
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtTA_RE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBuscaConfiguracion(txtTA_RE, "T")
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
