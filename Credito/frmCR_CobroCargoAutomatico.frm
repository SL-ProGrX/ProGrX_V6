VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmCR_CobroCargoAutomatico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobro: Cargo Automático (Tarjetas)"
   ClientHeight    =   7560
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   12144
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   12144
   Begin TabDlg.SSTab ssTab 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   11895
      _ExtentX        =   20976
      _ExtentY        =   11028
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lote"
      TabPicture(0)   =   "frmCR_CobroCargoAutomatico.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line4(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lsw"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtComprobante"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtUsuario"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtFecha"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtNotas"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboFechaCuota"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Carga"
      TabPicture(1)   =   "frmCR_CobroCargoAutomatico.frx":06B4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Aplicación"
      TabPicture(2)   =   "frmCR_CobroCargoAutomatico.frx":0E76
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Consulta"
      TabPicture(3)   =   "frmCR_CobroCargoAutomatico.frx":1839
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.ComboBox cboFechaCuota 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtNotas 
         Alignment       =   1  'Right Justify
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
         Height          =   2115
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox txtFecha 
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
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtUsuario 
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
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txtComprobante 
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
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   1440
         Width           =   2655
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   4815
         Left            =   4560
         TabIndex        =   3
         Top             =   600
         Width           =   7095
         _ExtentX        =   12510
         _ExtentY        =   8488
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Lote"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Comprobante"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   5009
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Notas"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Cuota"
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
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Notas"
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
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
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
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario"
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
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Comprobante"
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
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   240
         X2              =   11640
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lotes Pendientes:"
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
         Index           =   0
         Left            =   4560
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lote:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Image imgBanner 
      Height          =   1335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12255
   End
End
Attribute VB_Name = "frmCR_CobroCargoAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

