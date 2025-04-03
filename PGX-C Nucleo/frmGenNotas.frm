VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmGenNotas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Movimientos y Notas"
   ClientHeight    =   5040
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   7512
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7512
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab ssTab 
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   7335
      _ExtentX        =   12933
      _ExtentY        =   7641
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Detalle"
      TabPicture(0)   =   "frmGenNotas.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lsw"
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(2)=   "txtConcepto"
      Tab(0).Control(3)=   "txtUsuario"
      Tab(0).Control(4)=   "txtOrigen"
      Tab(0).Control(5)=   "txtFecha"
      Tab(0).Control(6)=   "txtModulo"
      Tab(0).Control(7)=   "Label4(3)"
      Tab(0).Control(8)=   "Label4(2)"
      Tab(0).Control(9)=   "Label4(1)"
      Tab(0).Control(10)=   "Label4(0)"
      Tab(0).Control(11)=   "Label3"
      Tab(0).Control(12)=   "Label2"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Traspaso"
      TabPicture(1)   =   "frmGenNotas.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Line1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkNC"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkND"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkAsientos"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chkResumenDiario"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdTrasladar"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdTrasladar 
         Caption         =   "&Trasladar"
         Height          =   375
         Left            =   6000
         TabIndex        =   21
         Top             =   2640
         Width           =   975
      End
      Begin VB.CheckBox chkResumenDiario 
         Alignment       =   1  'Right Justify
         Caption         =   "Asiento Tipo Resumen Diario"
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   1800
         Width           =   3135
      End
      Begin VB.CheckBox chkAsientos 
         Caption         =   "Asientos de Diario"
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   1440
         Width           =   3615
      End
      Begin VB.CheckBox chkND 
         Caption         =   "Notas de Débito"
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   1080
         Width           =   3615
      End
      Begin VB.CheckBox chkNC 
         Caption         =   "Notas de Crédito"
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   720
         Width           =   3615
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   16
         Top             =   1980
         Width           =   7095
         _ExtentX        =   12510
         _ExtentY        =   3831
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Débitos"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Créditos"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1620
         Width           =   6015
      End
      Begin VB.TextBox txtConcepto 
         Height          =   315
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1260
         Width           =   6015
      End
      Begin VB.TextBox txtUsuario 
         Height          =   315
         Left            =   -71040
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   900
         Width           =   3255
      End
      Begin VB.TextBox txtOrigen 
         Height          =   315
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Left            =   -71040
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   540
         Width           =   3255
      End
      Begin VB.TextBox txtModulo 
         Height          =   315
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   540
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   6960
         X2              =   240
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   240
         X2              =   6960
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label4 
         Caption         =   "Entidad"
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   14
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Concepto"
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   12
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario"
         Height          =   255
         Index           =   1
         Left            =   -72120
         TabIndex        =   10
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Origen"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   8
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   -72120
         TabIndex        =   6
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Módulo"
         Height          =   255
         Left            =   -74880
         TabIndex        =   4
         Top             =   540
         Width           =   1095
      End
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      ItemData        =   "frmGenNotas.frx":0038
      Left            =   2640
      List            =   "frmGenNotas.frx":0045
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtDocumento 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image imgImprimir 
      Height          =   255
      Left            =   4320
      Picture         =   "frmGenNotas.frx":007E
      Stretch         =   -1  'True
      ToolTipText     =   "ReImprimir Documento"
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Documento"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmGenNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
cbo.Text = "Asientos Diarios"
End Sub

