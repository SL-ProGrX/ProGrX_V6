VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCR_ConveniosCargosCxP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rebajos vía (Cargos de CxP)"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCR_ConveniosCargosCxP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   13140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDisponible 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   6360
      Width           =   1695
   End
   Begin XtremeSuiteControls.GroupBox GroupBox_Registro 
      Height          =   4452
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   12852
      _Version        =   1310723
      _ExtentX        =   22669
      _ExtentY        =   7853
      _StockProps     =   79
      Caption         =   "Nuevos Rebajos:"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.GroupBox GroupBox_Pendientes 
         Height          =   4092
         Left            =   6720
         TabIndex        =   38
         Top             =   480
         Width           =   6120
         _Version        =   1310723
         _ExtentX        =   10795
         _ExtentY        =   7218
         _StockProps     =   79
         Caption         =   "                    Cargos pendientes de Cobro:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin MSComctlLib.ListView lsw 
            Height          =   3372
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "Presione Doble Click para agregar"
            Top             =   480
            Width           =   5892
            _ExtentX        =   10398
            _ExtentY        =   5953
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
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
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No. Trans."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Frac.Pend."
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Cargo"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Saldo"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Descripción"
               Object.Width           =   4304
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Documento"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Detalle"
               Object.Width           =   4304
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Inicio Cobro"
               Object.Width           =   4304
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Cod/Cargo"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "No. Pagos"
               Object.Width           =   2540
            EndProperty
         End
         Begin XtremeSuiteControls.PushButton btnPendientes 
            Height          =   372
            Left            =   120
            TabIndex        =   40
            Top             =   0
            Width           =   492
            _Version        =   1310723
            _ExtentX        =   868
            _ExtentY        =   656
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FlatStyle       =   -1  'True
            Appearance      =   2
            Picture         =   "frmCR_ConveniosCargosCxP.frx":000C
         End
      End
      Begin VB.TextBox txtCargoCobroInicio 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   960
         Width           =   1692
      End
      Begin VB.TextBox txtNFraccion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   960
         Width           =   1212
      End
      Begin VB.TextBox txtNTransac 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   960
         Width           =   1692
      End
      Begin XtremeSuiteControls.PushButton btnNuevo 
         Height          =   612
         Left            =   3840
         TabIndex        =   25
         Top             =   3720
         Width           =   1332
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Nuevo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmCR_ConveniosCargosCxP.frx":0803
      End
      Begin VB.TextBox txtDetalle 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   672
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   2640
         Width           =   4812
      End
      Begin VB.TextBox txtDocumento 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   23
         Top             =   1920
         Width           =   1932
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   22
         Top             =   2280
         Width           =   1932
      End
      Begin VB.TextBox txtCargoDesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1440
         Width           =   3612
      End
      Begin VB.TextBox txtCargoCod 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   20
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1440
         Width           =   972
      End
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   612
         Left            =   5160
         TabIndex        =   26
         Top             =   3720
         Width           =   1332
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmCR_ConveniosCargosCxP.frx":0FBC
      End
      Begin XtremeSuiteControls.PushButton btnCerrar 
         Height          =   372
         Left            =   12360
         TabIndex        =   34
         Top             =   0
         Width           =   492
         _Version        =   1310723
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         Appearance      =   2
         Picture         =   "frmCR_ConveniosCargosCxP.frx":16C1
      End
      Begin VB.Label Label2 
         Caption         =   "Cobra a partir:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   7
         Left            =   4800
         TabIndex        =   33
         Top             =   720
         Width           =   1572
      End
      Begin VB.Label Label2 
         Caption         =   "No. Fracción:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   6
         Left            =   3600
         TabIndex        =   32
         Top             =   720
         Width           =   1572
      End
      Begin VB.Label Label2 
         Caption         =   "No. Transacción:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   1920
         TabIndex        =   31
         Top             =   720
         Width           =   1572
      End
      Begin VB.Label Label2 
         Caption         =   "Cargo Pendiente:"
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
         Left            =   360
         TabIndex        =   27
         Top             =   960
         Width           =   1452
      End
      Begin VB.Label Label2 
         Caption         =   "Monto:"
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
         Index           =   3
         Left            =   360
         TabIndex        =   19
         Top             =   2280
         Width           =   1452
      End
      Begin VB.Label Label2 
         Caption         =   "Detalle:"
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
         Left            =   360
         TabIndex        =   18
         Top             =   2640
         Width           =   1452
      End
      Begin VB.Label Label2 
         Caption         =   "Documento:"
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
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Top             =   1920
         Width           =   1452
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Cargo:"
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
         Left            =   360
         TabIndex        =   16
         Top             =   1440
         Width           =   1452
      End
   End
   Begin VB.TextBox txtVencido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtCxP_Flotante 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox txtEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox txtOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin XtremeSuiteControls.GroupBox GroupBox_Lista 
      Height          =   4572
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   12852
      _Version        =   1310723
      _ExtentX        =   22669
      _ExtentY        =   8064
      _StockProps     =   79
      Caption         =   "Rebajos registrados a la orden de liquidación:"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3972
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   12732
         _Version        =   524288
         _ExtentX        =   22458
         _ExtentY        =   7006
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   10
         RowHeaderDisplay=   2
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_ConveniosCargosCxP.frx":1E8E
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnAgregar 
         Height          =   372
         Left            =   12360
         TabIndex        =   35
         Top             =   0
         Width           =   492
         _Version        =   1310723
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         Appearance      =   2
         Picture         =   "frmCR_ConveniosCargosCxP.frx":2D61
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Disponible:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   8760
      TabIndex        =   37
      Top             =   6360
      Width           =   2172
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vencido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   3
      Left            =   10680
      TabIndex        =   12
      Top             =   240
      Width           =   1452
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Flotante"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   2
      Left            =   9240
      TabIndex        =   10
      Top             =   240
      Width           =   1452
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   7800
      TabIndex        =   7
      Top             =   240
      Width           =   1452
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total de Rebajos:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   8760
      TabIndex        =   5
      Top             =   6000
      Width           =   2172
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Convenio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   852
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No. Orden"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   20
      Left            =   6360
      TabIndex        =   3
      Top             =   240
      Width           =   1452
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmCR_ConveniosCargosCxP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String, vOrden As Long
Dim vCodCargo As String, vMonto As Double
Dim vEstado As String, vPaso As Boolean, vDisponible As Currency




Private Sub sbConsulta_EnCobro()
Dim strSQL As String, rs  As New ADODB.Recordset, itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass


GroupBox_Lista.Visible = False
GroupBox_Registro.Visible = True

lsw.ListItems.Clear

strSQL = "exec spConvenios_Rebajos_Consulta_Pendientes '" & vCodigo & "',Null," & vOrden
'@Convenio varchar(10), @Corte datetime = null, @Orden int = 0)
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!ID_TRANSAC)
      itmX.SubItems(1) = rs!FRACCIONES_PENDIENTES
      itmX.SubItems(2) = Format(rs!cargo, "Standard")
      itmX.SubItems(3) = Format(rs!SALDO, "Standard")
      itmX.SubItems(4) = rs!CargoDesc
      itmX.SubItems(5) = rs!Documento
      itmX.SubItems(6) = rs!Detalle
      itmX.SubItems(7) = Format(rs!COBRO_INICIO_FECHA, "dd/mm/yyyy")
      itmX.SubItems(8) = rs!cod_cargo
      itmX.SubItems(9) = rs!Numero_Pagos + 1


  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub btnAgregar_Click()

If vEstado = "C" Then
  MsgBox "No se puede modificar una orden en estado CERRADA!!!"
  Exit Sub
End If

Call btnNuevo_Click
End Sub

Private Sub btnCerrar_Click()
   Call sbConsultaCargos(vCodigo)
End Sub

Private Sub btnGuardar_Click()
Dim strSQL As String, i As Integer
Dim pTransaccion As Long

On Error GoTo vError
 
If vEstado = "C" Then
  MsgBox "No se puede modificar una orden en estado CERRADA!!!"
  Exit Sub
End If
 
i = vbYes

If IsNumeric(txtNTransac.Text) Then
  pTransaccion = txtNTransac.Text
Else
  pTransaccion = 0
End If

If txtCargoCod.Text = "" Then
   MsgBox "Indique un Código de Cargo válido!", vbExclamation
   Exit Sub
End If

If Not IsNumeric(txtMonto.Text) Then
   MsgBox "El monto del cargo no es válido!", vbExclamation
   Exit Sub
Else
  If CCur(txtMonto.Text) <= 0 Then
        MsgBox "El monto del cargo no es válido!", vbExclamation
        Exit Sub
  End If
End If

If CCur(txtMonto.Text) > (vDisponible - CCur(txtTotal.Text)) Then
      MsgBox "El monto del cargo EXCEDE el disponible!", vbExclamation
      Exit Sub
End If


strSQL = "exec spConvenios_Orden_Cargos_CxP_Registro '" & vCodigo & "'," & vOrden & ",0," & pTransaccion _
       & ",'" & txtCargoCod.Text & "'," & CCur(txtMonto.Text) & ",'" & Mid(txtDocumento.Text, 1, 30) _
       & "','" & Mid(txtDetalle.Text, 1, 500) & "','" & glogon.Usuario & "','A'"

Call ConectionExecute(strSQL)
If glogon.error Then Exit Sub

txtTotal.Text = Format(CCur(txtTotal.Text) + CCur(txtMonto.Text), "Standard")
txtDisponible.Text = Format(vDisponible - CCur(txtTotal.Text), "Standard")

i = MsgBox("Cargo registrado satisfactoriamente. Desea seguir agregando cargos?", vbYesNo)
If i = vbYes Then
    Call btnNuevo_Click
    If pTransaccion > 0 Then
        txtVencido.Text = Format(fxRebajoVencido(vCodigo, vOrden), "Standard")
    End If
Else
    Call sbConsultaCargos(vCodigo)
End If

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub



Private Sub btnNuevo_Click()

txtNTransac.Text = ""
txtNFraccion.Text = ""
txtCargoCobroInicio.Text = ""

txtCargoCod.Text = ""
txtCargoDesc.Text = ""
txtMonto.Text = "0.00"
txtDocumento.Text = ""
txtDetalle.Text = ""


Call sbConsulta_EnCobro

End Sub

Private Sub btnPendientes_Click()
Dim pLeft As Long, pWidth As Long
Dim pWLsw As Long

If GroupBox_Pendientes.Width = 6120 Then
    pWidth = 12852
    pLeft = 0
    pWLsw = 12732
Else
    pWidth = 6120
    pLeft = 6720
    pWLsw = 5892
End If

GroupBox_Pendientes.Width = pWidth
GroupBox_Pendientes.Left = pLeft
lsw.Width = pWLsw


End Sub

Private Sub Form_Activate()
  vModulo = 16
End Sub

Private Sub Form_Load()
 vModulo = 16
 
 Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 
 vGrid.MaxRows = 0
 vGrid.MaxCols = 10
 txtTotal.Text = 0
 
 vCodigo = GLOBALES.gTag
 vOrden = GLOBALES.gTag2
 
 GroupBox_Registro.Top = GroupBox_Lista.Top
 GroupBox_Registro.Left = GroupBox_Lista.Left
 
 GroupBox_Registro.Height = GroupBox_Lista.Height
 GroupBox_Registro.Width = GroupBox_Lista.Width
 
 GroupBox_Lista.Visible = True
 GroupBox_Registro.Visible = False
 
 Call sbConsultaConvenio(vCodigo, vOrden)
 Call sbConsultaCargos(vCodigo)
 
End Sub

Private Sub sbCalculaTotales()
Dim curTotal As Currency
Dim i As Integer

On Error GoTo vError

curTotal = 0

With vGrid
    For i = 1 To .MaxRows
     .Row = i
     .Col = 10
     curTotal = curTotal + CCur(.Text)
    Next i
End With

txtTotal.Text = Format(curTotal, "Standard")
txtDisponible.Text = Format(vDisponible - curTotal, "Standard")


Exit Sub

vError:

End Sub

Private Function fxRebajoVencido(ByVal pConvenio As String, pOrden As Long) As Currency
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As Currency

On Error GoTo vError

vResultado = 0

   strSQL = "select dbo.fxConvenio_CxP_Rebajos_Vencidos('" & pConvenio & "'," & pOrden & ") AS 'Rebajo_Vencido'"
   Call OpenRecordSet(rs, strSQL)

   If Not rs.EOF Then
     vResultado = rs!Rebajo_Vencido
   End If
   rs.Close


fxRebajoVencido = vResultado


Exit Function

vError:
    fxRebajoVencido = vResultado

End Function


Private Sub sbConsultaConvenio(ByVal pConvenio As String, pOrden As Long)
Dim strSQL As String, rs As New ADODB.Recordset
  
On Error GoTo vError
'- IVA_REFERENCIA_CRD
   strSQL = "select O.COD_CONVENIO,C.DESCRIPCION,O.COD_ORDEN,O.ESTADO, TOTAL_PAGAR + TOTAL_REBAJOS_CXP as 'Disponible'" _
          & ",dbo.fxConvenio_CxP_CargosFlotantes('" & pConvenio & "') AS 'CxP_Flotante'" _
          & ",dbo.fxConvenio_CxP_Rebajos_Vencidos('" & pConvenio & "'," & pOrden & ") AS 'Rebajo_Vencido'" _
          & " from CRD_CONVENIOS_ORDENES O" _
          & "  inner join CRD_CONVENIOS C on O.COD_CONVENIO = C.COD_CONVENIO" _
          & " where O.COD_CONVENIO = '" & pConvenio & "' AND O.COD_ORDEN = " & pOrden
   Call OpenRecordSet(rs, strSQL)

   If Not rs.EOF Then
      txtCodigo.Text = rs!COD_CONVENIO
      txtDescripcion.Text = rs!Descripcion
      txtOrden.Text = rs!cod_orden
      vEstado = rs!estado
      
      Select Case vEstado
        Case "A"
          txtEstado.Text = "Abierta"
        Case "C"
          txtEstado.Text = "Cerrada"
        Case "T"
          txtEstado.Text = "Tramitada"
      End Select
   End If
    
    vDisponible = rs!Disponible - rs!CxP_flotante
   
    txtCxP_Flotante.Text = Format(rs!CxP_flotante, "Standard")
    txtVencido.Text = Format(rs!Rebajo_Vencido, "Standard")
    txtDisponible.Text = Format(vDisponible, "Standard")
   rs.Close

Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbConsultaCargos(vConvenio As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMonto As Currency
On Error GoTo vError

vMonto = 0


vPaso = True

vGrid.MaxRows = 0

GroupBox_Lista.Visible = True
GroupBox_Registro.Visible = False

  strSQL = "exec spConvenios_Orden_Cargos_CxP '" & vConvenio & "'," & vOrden & ",1"
  Call OpenRecordSet(rs, strSQL)
    
  With vGrid
   Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     
     .Col = 3
     .Text = CStr(rs!Num_Linea)
     
     .Col = 4
     .Text = rs!cod_cargo
     
     .Col = 5
     .Text = rs!Descripcion
     
     .Col = 6
     .Text = CStr(rs!ID_TRANSAC)
     
     .Col = 7
     .Text = CStr(rs!ID_FRACCION)
    
     .Col = 8
     .Text = rs!Documento & ""
     
     .Col = 9
     .Text = rs!Detalle & ""
     
     .Col = 10
     .Text = IIf(IsNull(rs!Monto), 0, Format(rs!Monto, "Standard"))
     vMonto = vMonto + rs!Monto
     
     .RowHeight(.Row) = .MaxTextRowHeight(.Row)
     
     rs.MoveNext
   Loop
   rs.Close

  End With
    
 txtVencido.Text = Format(fxRebajoVencido(vCodigo, vOrden), "Standard")
 Call sbCalculaTotales


vPaso = False

Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub






Private Sub lsw_DblClick()

If lsw.ListItems.Count = 0 Then Exit Sub

With lsw.SelectedItem

Call btnPendientes_Click

txtNTransac.Text = .Text
txtNFraccion.Text = .ListSubItems.Item(9).Text
txtCargoCobroInicio.Text = .ListSubItems.Item(7).Text

txtCargoCod.Text = .ListSubItems.Item(8).Text
txtCargoDesc.Text = .ListSubItems.Item(4).Text
txtMonto.Text = .ListSubItems.Item(2).Text
txtDocumento.Text = .ListSubItems.Item(5).Text
txtDetalle.Text = .ListSubItems.Item(6).Text


txtMonto.SetFocus

End With


End Sub

Private Sub txtCargoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCargoDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Ca.COD_CARGO"
  gBusquedas.Orden = "Ca.COD_CARGO"
  gBusquedas.Consulta = "select Ca.COD_CARGO, Ca.DESCRIPCION" _
        & " from CRD_CONVENIOS_CARGOS_CXP Cc inner join CXP_CARGOS Ca on Cc.COD_CARGO = Ca.COD_CARGO"
  gBusquedas.Filtro = " AND Cc.COD_CONVENIO = '" & txtCodigo.Text & "' and Ca.ACTIVO = 1"
  frmBusquedas.Show vbModal
  txtCargoCod = gBusquedas.Resultado
  txtCargoDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCargoCod_LostFocus()
txtCargoDesc = fxSIFCCodigos("D", txtCargoCod, "CargosProv")
End Sub

Private Sub txtCargoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Ca.DESCRIPCION"
  gBusquedas.Orden = "Ca.DESCRIPCION"
  gBusquedas.Consulta = "select Ca.COD_CARGO, Ca.DESCRIPCION" _
        & " from CRD_CONVENIOS_CARGOS_CXP Cc inner join CXP_CARGOS Ca on Cc.COD_CARGO = Ca.COD_CARGO"
  gBusquedas.Filtro = " AND Cc.COD_CONVENIO = '" & txtCodigo.Text & "' and Ca.ACTIVO = 1"
  frmBusquedas.Show vbModal
  txtCargoCod = gBusquedas.Resultado
  txtCargoDesc = gBusquedas.Resultado2
End If

End Sub


Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub txtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then btnGuardar.SetFocus
End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError
 txtMonto = CCur(txtMonto)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDetalle.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
 txtMonto = Format(CCur(txtMonto), "Standard")
vError:
End Sub


Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vLinea As Long, vTransaccion As Long

On Error GoTo vError
  
  
If vPaso Then Exit Sub
 
If vEstado = "C" Then
  MsgBox "No se puede modificar una orden en estado CERRADA!!!"
  Exit Sub
End If
 
 
 With vGrid
     .Row = Row

     .Col = 3
     vLinea = .Text
     .Col = 6
     vTransaccion = .Text

     Select Case Col
       Case 1 'Borrado
            strSQL = "exec spConvenios_Orden_Cargos_CxP_Registro '" & vCodigo & "'," & vOrden & "," & vLinea _
                   & ",0,'',0,'','','" & glogon.Usuario & "','B'"
            
            Call ConectionExecute(strSQL)
            If Not glogon.error Then
                'Remover Fila
                vGrid.DeleteRows vGrid.ActiveRow, 1
                vGrid.MaxRows = vGrid.MaxRows - 1
                vGrid.Row = vGrid.ActiveRow
                
                'Actualizar Rebajos Vencidos
                If vTransaccion > 0 Then
                    txtVencido.Text = Format(fxRebajoVencido(vCodigo, vOrden), "Standard")
                End If
            End If
            
            
            Call sbCalculaTotales
        
       Case 2 'Nuevo
            Call btnNuevo_Click
     
     End Select
 End With
  
  
Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCodCargo As String, vMonto As Double, vLinea As Long
Dim vDocumento As String, vDetalle As String

On Error GoTo vError
  
 
If vEstado = "C" Then
  MsgBox "No se puede modificar una orden en estado CERRADA!!!"
  Exit Sub
End If
 
 With vGrid
   If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
          
     .Row = .ActiveRow
     
     .Col = 3
     vLinea = .Text
     
     .Col = 4
     vCodCargo = .Text

     .Col = 8
     vDocumento = Trim(.Text)
     .Col = 9
     vDetalle = Trim(.Text)
     .Col = 10
     vMonto = .Text
     
     'Codigo de Transaccion
     .Col = 6
     
     If vMonto > 0 Then
        strSQL = "exec spConvenios_Orden_Cargos_CxP_Registro '" & vCodigo & "'," & vOrden & "," & vLinea & "," & .Text _
               & ",'" & vCodCargo & "'," & CCur(vMonto) & ",'" & Mid(vDocumento, 1, 30) _
               & "','" & Mid(vDetalle, 1, 500) & "','" & glogon.Usuario & "','E'"
        Call OpenRecordSet(rs, strSQL)
        If Not glogon.error Then
                .Col = 10
                .Text = Format(rs!Monto, "Standard")
        End If
        rs.Close
     End If
     
     Call sbCalculaTotales
     
  End If
 End With
  
  
Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxValidaExiste(ByVal vCodCargos As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
       
fxValidaExiste = True

strSQL = " Select COD_CARGO" _
       & " from CRD_CONVENIOS_DT_CARGOS_CXP" _
       & " where COD_ORDEN = " & vOrden & " and COD_CONVENIO = '" & vCodigo & "' and COD_CARGO = '" & vCodCargos & "'"
Call OpenRecordSet(rs, strSQL)
     
If rs.EOF Then
   fxValidaExiste = False
End If
  
Exit Function
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



