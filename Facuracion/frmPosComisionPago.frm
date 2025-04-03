VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPosComisionPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de Comisiones"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   8460
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   7440
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   315
      Left            =   6960
      TabIndex        =   4
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Frame fraComision 
      Caption         =   "Cargando y Calculando Información [Espere...]"
      Height          =   1335
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   5295
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Reporte"
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "&Detalle"
      Height          =   315
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   184352771
      CurrentDate     =   37321
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   184352771
      CurrentDate     =   37321
   End
   Begin MSComCtl2.DTPicker dtpReporte 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   5160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   184352771
      CurrentDate     =   37169
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   4575
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Agente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Ventas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Comisión"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lbl1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ventas Realizadas entre  "
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Generación"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   1335
   End
End
Attribute VB_Name = "frmPosComisionPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

