VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmTES_ReversaTransFerencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reversión De transferencias"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   8490
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   14420
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Aplica Reversión"
      TabPicture(0)   =   "frmTES_ReversaTransFerencia.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line2(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(37)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line2(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblEnd2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblEnd1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblSolicitudes"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line2(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "prgBar"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ImageList1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "tlbBuscar"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lsw"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtObservaciones"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtNdocumento"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtMonto"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtCasos"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdAplicar"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtContraseña"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cbo"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Consulta Reversión"
      TabPicture(1)   =   "frmTES_ReversaTransFerencia.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(4)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(5)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Line2(3)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(7)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Line2(4)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "tblBusqueda"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lswDetalle"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lswConsulta"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "dtpCorte"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "dtpInicio"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cboBancos"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtMontoDet"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtcasosDet"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cmdImprime"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      Begin VB.CommandButton cmdImprime 
         Caption         =   "&Reporte"
         Height          =   375
         Left            =   -68880
         TabIndex        =   33
         Top             =   7680
         Width           =   1335
      End
      Begin VB.TextBox txtcasosDet 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -73920
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   7680
         Width           =   855
      End
      Begin VB.TextBox txtMontoDet 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -72120
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   7680
         Width           =   1935
      End
      Begin VB.ComboBox cboBancos 
         Height          =   315
         Left            =   -73440
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   600
         Width           =   6015
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   6375
      End
      Begin VB.TextBox txtContraseña 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   6780
         Width           =   1695
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Reversar"
         Height          =   375
         Left            =   6360
         TabIndex        =   5
         Top             =   7140
         Width           =   1335
      End
      Begin VB.TextBox txtCasos 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   4740
         Width           =   855
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   4740
         Width           =   1935
      End
      Begin VB.TextBox txtNdocumento 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   795
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   5580
         Width           =   6330
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   3015
         Left            =   120
         TabIndex        =   8
         Top             =   1620
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5318
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "# Solicitud"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Beneficiario"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbBuscar 
         Height          =   570
         Left            =   3720
         TabIndex        =   9
         Top             =   840
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   1005
         ButtonWidth     =   1376
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Buscar"
               Key             =   "Buscar"
               Object.ToolTipText     =   "Busca Solicitudes Pendientes"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6840
         Top             =   4860
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTES_ReversaTransFerencia.frx":0038
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTES_ReversaTransFerencia.frx":0352
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   150
         Left            =   120
         TabIndex        =   10
         Top             =   7665
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   265
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   -73440
         TabIndex        =   22
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   137297923
         CurrentDate     =   37321
      End
      Begin MSComCtl2.DTPicker dtpCorte 
         Height          =   315
         Left            =   -72000
         TabIndex        =   23
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   137297923
         CurrentDate     =   37321
      End
      Begin MSComctlLib.ListView lswConsulta 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   25
         Top             =   1920
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         Enabled         =   0   'False
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "# Reversioón"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Autorizado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Generado X"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Observaciones"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lswDetalle 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   28
         Top             =   4560
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5318
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         Enabled         =   0   'False
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "# Solicitud"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Beneficiario"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.Toolbar tblBusqueda 
         Height          =   570
         Left            =   -70560
         TabIndex        =   34
         Top             =   960
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   1005
         ButtonWidth     =   1376
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Buscar"
               Key             =   "Buscar"
               Object.ToolTipText     =   "Busca Solicitudes Pendientes"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   -74760
         X2              =   -67080
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalle Reversión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   32
         Top             =   4200
         Width           =   3015
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   -74760
         X2              =   -67080
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reversión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   31
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "# Casos"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -74640
         TabIndex        =   30
         Top             =   7680
         Width           =   735
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -72960
         TabIndex        =   29
         Top             =   7680
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Fechas Reversión"
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   24
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Banco"
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Banco"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Contraseña"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   6780
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   120
         X2              =   7560
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label lblSolicitudes 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solicitudes Pertenecientes al la transferencia"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1380
         Width           =   7575
      End
      Begin VB.Label lblEnd1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "# Casos"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   4740
         Width           =   735
      End
      Begin VB.Label lblEnd2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1920
         TabIndex        =   14
         Top             =   4740
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   120
         X2              =   7680
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   7800
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reversión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   37
         Left            =   120
         TabIndex        =   13
         Top             =   5220
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   5580
         Width           =   1575
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   120
         X2              =   7800
         Y1              =   6660
         Y2              =   6660
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Autoriza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   6420
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmTES_ReversaTransFerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vpaso As Boolean, fFechaEmision As Date, vCodigoImprime As Integer
Private Sub cbo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vpaso Then Exit Sub

If cbo.ListCount = 0 Then
   cbo.AddItem " "
   cbo.ItemData(cbo.NewIndex) = 0
   cbo.Text = " "
End If

strSQL = "Select max(cast(documento_base as int))as documento from cheques where id_banco = " & cbo.ItemData(cbo.ListIndex) & ""
rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF Or Not rs.BOF Then
  If Not IsNull(rs!Documento) Then
    txtNdocumento = rs!Documento
  Else
    txtNdocumento = Empty
  End If
End If

txtCasos = 0
txtMonto = 0
rs.Close
lsw.ListItems.Clear
vpaso = False
End Sub

Private Sub cmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, iConsecutivo As Integer

On Error GoTo vError


If Trim(txtContraseña) = "" Then
   MsgBox "No se puede Autorizar" & vbCrLf & "Suministre La Contraseña De Autorización", vbExclamation, "Faltan Datos"
   Call Form_Load
   Exit Sub
End If

If Trim(txtObservaciones) = "" Then
   MsgBox "Porfavor digite observaciones para la reversión"
   Call Form_Load
   Exit Sub
End If

If Format(fFechaEmision, "dd/mm/yyyy") <> Format(fxFechaServidor, "dd/mm/yyyy") Then
   MsgBox "Esta intentando reversar una transferencia de un dia pasado " & fFechaEmision
   Call Form_Load
   Exit Sub
End If


'If UCase(txtUSuarioAutoriza) = UCase(glogon.Usuario) Then
'   MsgBox "No puede ser autorizado por el usuarui actual"
'   Call Form_Load
'   Exit Sub
'End If



Me.MousePointer = vbHourglass

strSQL = "Select * From Tes_Autorizaciones Where Clave='" _
       & fxCifrado(Trim(txtContraseña)) & "' and nombre = '" & glogon.Usuario _
       & "' and estado = 'A'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
     MsgBox "No se puede Reversar", vbExclamation, "Contraseña Incorrecta, o no Existe Nivel de Autorización"
     rs.Close
     Me.MousePointer = vbDefault
     txtCasos = 0
     txtMonto = 0
     txtContraseña = Empty
     
     Exit Sub
End If
rs.Close
glogon.Conection.BeginTrans
   
PrgBar.Max = lsw.ListItems.Count + 1
PrgBar.Value = 1

PrgBar.Visible = True

strSQL = "update cheques set estado = 'P', ndocumento = '' " _
                & " where documento_base = '" & txtNdocumento & "'  and id_banco = " & cbo.ItemData(cbo.ListIndex) & ""
 glogon.Conection.Execute strSQL
 
iConsecutivo = fxConsecutivo
strSQL = "insert into tes_te_reversion(id_reversion,autorizado,user_genera,fecha_genera,observaciones,id_banco,documento) " _
        & "values(" & iConsecutivo & ",'" & UCase(glogon.Usuario) & "','" & glogon.Usuario & "','" & Format(fxFechaServidor, "yyyymmdd") & "','" & txtObservaciones & "'," & cbo.ItemData(cbo.ListIndex) & ",'" & txtNdocumento & "')"
glogon.Conection.Execute strSQL

Call Bitacora("Aplica", "Reversion Transferencia = " & txtNdocumento & " COD.BAN:" _
            & cbo.ItemData(cbo.ListIndex))
        
        
        
For i = 1 To lsw.ListItems.Count
lsw.SelectedItem = lsw.ListItems(i)
    strSQL = "insert into tes_te_reversion_det(nsolicitud,id_reversion,cedula,nombre,monto,cta_ahorros)" _
            & "values('" & lsw.SelectedItem.Text & "'," & iConsecutivo & ",'" & lsw.SelectedItem.SubItems(1) & "'," _
            & "'" & lsw.SelectedItem.SubItems(2) & "'," & CCur(lsw.SelectedItem.SubItems(3)) & ",'" & lsw.SelectedItem.SubItems(4) & "')"
   
    glogon.Conection.Execute strSQL
   
   PrgBar.Value = PrgBar.Value + 1
   
Next i
   

PrgBar.Visible = False
glogon.Conection.CommitTrans



Call SbImprimeReporte(iConsecutivo)



Call Form_Load

Me.MousePointer = vbDefault


Exit Sub

vError:
  glogon.Conection.RollbackTrans
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical
  

End Sub

Private Sub cmdImprime_Click()
Call SbImprimeReporte(vCodigoImprime)
lswConsulta.Enabled = False
lswDetalle.Enabled = False
lswConsulta.ListItems.Clear
lswDetalle.ListItems.Clear
txtcasosDet = 0
txtMontoDet = 0
End Sub

Private Sub Form_Load()


lsw.ListItems.Clear
txtNdocumento = Empty
txtMonto = Empty
txtCasos = Empty
vpaso = True
SSTab1.Tab = 0
Call sbgCargaCboBancoGestion(cbo, glogon.Usuario, "Autoriza")
vpaso = False
'Call cbo_Click
Me.Icon = MDIMenu.Icon
Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

Private Sub lswConsulta_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, curMonto As Currency, itmX As ListItem

On Error GoTo vError

    Me.MousePointer = vbHourglass
    
    lswDetalle.ListItems.Clear
    i = 0
    curMonto = 0
    vCodigoImprime = lswConsulta.SelectedItem.Text
    
    strSQL = "select * from tes_te_reversion_det where id_reversion = " & lswConsulta.SelectedItem.Text & " "
    
    rs.Open strSQL, glogon.Conection, adOpenForwardOnly
    If Not rs.BOF Or Not rs.EOF Then
      lswDetalle.Enabled = True
    Else
      lswDetalle.Enabled = False
    End If
    Do While Not rs.EOF
     Set itmX = lswDetalle.ListItems.Add(, , rs!NSolicitud)
         itmX.SubItems(1) = rs!Cedula
         itmX.SubItems(2) = rs!Nombre
         itmX.SubItems(3) = Format(rs!Monto, "Standard")
         itmX.SubItems(4) = rs!cta_ahorros
         'itmX.Checked = chkMarcas.Value
         curMonto = curMonto + rs!Monto
         i = i + 1
      
         
     rs.MoveNext
    Loop
    
    rs.Close

    txtcasosDet = Format(i, "###,###,###,##0")
    txtMontoDet = Format(curMonto, "Standard")
    
    Me.MousePointer = vbDefault
    Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
   vpaso = True
   Call sbgCargaCboBancoGestion(cboBancos, glogon.Usuario, "Autoriza")
   vpaso = False
  lswConsulta.ListItems.Clear
  lswDetalle.ListItems.Clear
  dtpCorte = Format(fxFechaServidor, "dd/mm/yyyy")
  dtpInicio = dtpCorte
End If
End Sub

Private Sub tblBusqueda_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

    If vpaso Then Exit Sub
    
    If cbo.ListCount = 0 Then
        cbo.AddItem " "
        cbo.ItemData(cbo.NewIndex) = 0
        cbo.Text = " "
    End If
    
    strSQL = "select * from tes_te_reversion where id_banco = " & cboBancos.ItemData(cboBancos.ListIndex) & " and " _
    & " fecha_genera between '" & Format(dtpInicio, "yyyymmdd") & "' and '" & Format(dtpCorte, "yyyymmdd") & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    If Not rs.EOF Or Not rs.BOF Then
        lswConsulta.Enabled = True
    Else
        lswConsulta.Enabled = False
    End If
    
    Do While Not rs.EOF
        Set itmX = lswConsulta.ListItems.Add(, , rs!id_reversion)
        itmX.SubItems(1) = rs!autorizado
        itmX.SubItems(2) = rs!user_genera
        itmX.SubItems(3) = Format(rs!fecha_genera, "dd/mm/yyyy")
        itmX.SubItems(4) = rs!observaciones
        'itmX.Checked = chkMarcas.Value
        
        rs.MoveNext
    Loop
    
    rs.Close
    
    vpaso = False

End Sub

Private Sub tlbBuscar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, curMonto As Currency, itmX As ListItem

    On Error GoTo vError
    
    
    If txtNdocumento = Empty Then
        MsgBox "No ha seleccionado la transferencia a Reversar"
    End If
    
    
    Me.MousePointer = vbHourglass
    
    lsw.ListItems.Clear
    i = 0
    curMonto = 0
    
    strSQL = "select fecha_emision,nsolicitud,codigo,beneficiario,monto,cta_ahorros from cheques" _
    & " where documento_base = '" & txtNdocumento & "' and id_banco = " & cbo.ItemData(cbo.ListIndex) & " "
    
    rs.Open strSQL, glogon.Conection, adOpenForwardOnly
    Do While Not rs.EOF
        Set itmX = lsw.ListItems.Add(, , rs!NSolicitud)
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = rs!beneficiario
        itmX.SubItems(3) = Format(rs!Monto, "Standard")
        itmX.SubItems(4) = rs!cta_ahorros
        'itmX.Checked = chkMarcas.Value
        curMonto = curMonto + rs!Monto
        i = i + 1
        
        
        rs.MoveNext
    Loop
    If rs.EOF Then
       rs.MoveFirst
       fFechaEmision = rs!fecha_emision
    End If
    
    rs.Close
    
    
    
    txtCasos = Format(i, "###,###,###,##0")
    txtMonto = Format(curMonto, "Standard")
    
    Me.MousePointer = vbDefault
    Exit Sub
    
vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical
    
End Sub


Private Function fxConsecutivo() As Long
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select coalesce(Max(id_reversion),0) as consecutivo from tes_te_reversion"

rs.Open strSQL, glogon.Conection, adOpenStatic

  fxConsecutivo = rs!consecutivo + 1
  
rs.Close

End Function
Private Sub SbImprimeReporte(vconsec As Integer)
Dim strSQL As String

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Tesorería"
    
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
  '  .Formulas(2) = "Banco='" & cbo.Text & "'"
    .Formulas(3) = "Usuario='" & glogon.Usuario & "'"

    .ReportFileName = App.Path & "\Reportes\Tesreversiontransferencia.rpt"
    
    strSQL = "{tes_te_reversion.id_reversion} = " & vconsec & ""
    .SelectionFormula = strSQL
    .PrintReport

End With
End Sub
