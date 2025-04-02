VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Begin VB.Form frmAF_CD_Comites 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Comites"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13830
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAF_CD_Comites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   13830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcPlan 
      Height          =   6975
      Left            =   9000
      TabIndex        =   60
      Top             =   1800
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8493
      _ExtentY        =   12303
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   1
      Item(0).Caption =   "Plan de Trabajo"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "btnMensaje(0)"
      Item(0).Control(1)=   "gridMensajes"
      Item(0).Control(2)=   "btnMensaje(1)"
      Begin XtremeSuiteControls.PushButton btnMensaje 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   61
         Top             =   360
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Plan de Trabajo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAF_CD_Comites.frx":0A02
      End
      Begin FPSpreadADO.fpSpread gridMensajes 
         Height          =   6000
         Left            =   120
         TabIndex        =   62
         Top             =   840
         Width           =   4440
         _Version        =   524288
         _ExtentX        =   7832
         _ExtentY        =   10583
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
         MaxCols         =   1
         ScrollBars      =   0
         SpreadDesigner  =   "frmAF_CD_Comites.frx":1026
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnMensaje 
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   63
         Top             =   360
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_CD_Comites.frx":1539
      End
   End
   Begin XtremeSuiteControls.CheckBox chkAsociaUnidad 
      Height          =   255
      Left            =   3480
      TabIndex        =   46
      Top             =   480
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Asociar Unidad"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Frame fraBusquedaGeneral 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   8535
      Begin XtremeSuiteControls.ListView lswSelecBusqueda 
         Height          =   2295
         Left            =   1320
         TabIndex        =   76
         Top             =   600
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   4048
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCod_Comite_Ingreso 
         Height          =   330
         Left            =   1320
         TabIndex        =   67
         Top             =   240
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDesc_Comite_Ingreso 
         Height          =   330
         Left            =   2400
         TabIndex        =   68
         Top             =   240
         Width           =   6135
         _Version        =   1441793
         _ExtentX        =   10821
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtBusqueda 
         Height          =   330
         Left            =   1320
         TabIndex        =   69
         Top             =   240
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label6 
         Caption         =   "Presione doble click para agregar en la lista"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9480
      TabIndex        =   0
      Top             =   840
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10800
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Comites.frx":1C39
            Key             =   ""
            Object.Tag             =   "Nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Comites.frx":264B
            Key             =   ""
            Object.Tag             =   "Aplicar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Comites.frx":305D
            Key             =   ""
            Object.Tag             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Comites.frx":3A6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Comites.frx":4481
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Comites.frx":4E93
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Comites.frx":58A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Comites.frx":62B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_Comites.frx":6CC9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Unidades"
      TabPicture(0)   =   "frmAF_CD_Comites.frx":76DB
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ShortcutCaption1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lswComites"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Actividades"
      TabPicture(1)   =   "frmAF_CD_Comites.frx":80ED
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lswActividades"
      Tab(1).Control(1)=   "ShortcutCaption1(1)"
      Tab(1).Control(2)=   "Label8(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Ejecutivo(s)"
      TabPicture(2)   =   "frmAF_CD_Comites.frx":8AFF
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lswEjecutivo"
      Tab(2).Control(1)=   "ShortcutCaption1(2)"
      Tab(2).Control(2)=   "Label8(2)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Miembros"
      TabPicture(3)   =   "frmAF_CD_Comites.frx":9511
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraMiembros"
      Tab(3).Control(1)=   "chkActivos"
      Tab(3).Control(2)=   "tlbNuevo"
      Tab(3).Control(3)=   "lswMiembrosComite"
      Tab(3).Control(4)=   "Label3"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Liquidaciones"
      TabPicture(4)   =   "frmAF_CD_Comites.frx":9F23
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ShortcutCaption1(3)"
      Tab(4).Control(1)=   "ListView1"
      Tab(4).Control(2)=   "lswLiquidaciones"
      Tab(4).ControlCount=   3
      Begin XtremeSuiteControls.ListView lswMiembrosComite 
         Height          =   5175
         Left            =   -74400
         TabIndex        =   72
         Top             =   1440
         Width           =   7935
         _Version        =   1441793
         _ExtentX        =   13996
         _ExtentY        =   9128
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswComites 
         Height          =   3135
         Left            =   1440
         TabIndex        =   64
         Top             =   3720
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   5530
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswActividades 
         Height          =   3135
         Left            =   -73560
         TabIndex        =   65
         Top             =   3720
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   5530
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswEjecutivo 
         Height          =   3135
         Left            =   -73560
         TabIndex        =   66
         Top             =   3720
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   5530
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswLiquidaciones 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   70
         Top             =   480
         Width           =   8175
         _Version        =   1441793
         _ExtentX        =   14420
         _ExtentY        =   5741
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin VB.Frame fraMiembros 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6615
         Left            =   -75000
         TabIndex        =   17
         Top             =   360
         Width           =   8775
         Begin XtremeSuiteControls.PushButton btnMiembros 
            Height          =   375
            Index           =   2
            Left            =   8040
            TabIndex        =   75
            ToolTipText     =   "Cerrar"
            Top             =   120
            Width           =   495
            _Version        =   1441793
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAF_CD_Comites.frx":9F3F
         End
         Begin XtremeSuiteControls.PushButton btnMiembros 
            Height          =   375
            Index           =   1
            Left            =   6840
            TabIndex        =   74
            Top             =   120
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Guardar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAF_CD_Comites.frx":A655
         End
         Begin XtremeSuiteControls.PushButton btnMiembros 
            Height          =   375
            Index           =   0
            Left            =   5640
            TabIndex        =   73
            Top             =   120
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Nuevo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAF_CD_Comites.frx":AD86
         End
         Begin TabDlg.SSTab tabGeneral 
            Height          =   6015
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Visible         =   0   'False
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   10610
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            BackColor       =   -2147483644
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "General"
            TabPicture(0)   =   "frmAF_CD_Comites.frx":B3B8
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label15"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label16"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Line1"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label19"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "lblMail"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label14"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label18"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "lblCelular"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "lblTelefono"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Label12"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "Label4"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "Line2"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "Label5"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "Label7"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "Label10"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "Label13"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "Label20"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "lblUT"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "Label21"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "cboRango"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "cboPuesto"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "chkDesembolso"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "chkActivo"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "txtNotas"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).Control(24)=   "txtCorreoJefe"
            Tab(0).Control(24).Enabled=   0   'False
            Tab(0).Control(25)=   "txtNombreJefe"
            Tab(0).Control(25).Enabled=   0   'False
            Tab(0).Control(26)=   "txtCelularJefe"
            Tab(0).Control(26).Enabled=   0   'False
            Tab(0).Control(27)=   "txtTelefonoJefe"
            Tab(0).Control(27).Enabled=   0   'False
            Tab(0).Control(28)=   "dtpFechaEleccion"
            Tab(0).Control(28).Enabled=   0   'False
            Tab(0).Control(29)=   "fraBuscaMiembros"
            Tab(0).Control(29).Enabled=   0   'False
            Tab(0).ControlCount=   30
            TabCaption(1)   =   "Historial"
            TabPicture(1)   =   "frmAF_CD_Comites.frx":BDCA
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lswMiembros_H"
            Tab(1).ControlCount=   1
            Begin XtremeSuiteControls.ListView lswMiembros_H 
               Height          =   5415
               Left            =   -74880
               TabIndex        =   71
               Top             =   480
               Width           =   8175
               _Version        =   1441793
               _ExtentX        =   14420
               _ExtentY        =   9551
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               View            =   3
               FullRowSelect   =   -1  'True
               Appearance      =   17
               UseVisualStyle  =   0   'False
            End
            Begin VB.Frame fraBuscaMiembros 
               Height          =   735
               Left            =   120
               TabIndex        =   19
               Top             =   480
               Visible         =   0   'False
               Width           =   8175
               Begin VB.TextBox txtCedula 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   120
                  TabIndex        =   20
                  ToolTipText     =   "Presione (F4) para seleccionar miembros"
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.Label lblNombre 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   21
                  Top             =   240
                  Width           =   6495
               End
            End
            Begin XtremeSuiteControls.DateTimePicker dtpFechaEleccion 
               Height          =   330
               Left            =   1440
               TabIndex        =   44
               Top             =   1680
               Width           =   1815
               _Version        =   1441793
               _ExtentX        =   3201
               _ExtentY        =   582
               _StockProps     =   68
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   3
            End
            Begin XtremeSuiteControls.FlatEdit txtTelefonoJefe 
               Height          =   330
               Left            =   1440
               TabIndex        =   50
               Top             =   3600
               Width           =   1575
               _Version        =   1441793
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCelularJefe 
               Height          =   330
               Left            =   4680
               TabIndex        =   51
               Top             =   3600
               Width           =   1575
               _Version        =   1441793
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtNombreJefe 
               Height          =   330
               Left            =   3000
               TabIndex        =   49
               Top             =   3120
               Width           =   4815
               _Version        =   1441793
               _ExtentX        =   8493
               _ExtentY        =   582
               _StockProps     =   77
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCorreoJefe 
               Height          =   330
               Left            =   1440
               TabIndex        =   52
               Top             =   4080
               Width           =   6375
               _Version        =   1441793
               _ExtentX        =   11245
               _ExtentY        =   582
               _StockProps     =   77
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtNotas 
               Height          =   570
               Left            =   1440
               TabIndex        =   53
               Top             =   2160
               Width           =   6375
               _Version        =   1441793
               _ExtentX        =   11245
               _ExtentY        =   1005
               _StockProps     =   77
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MultiLine       =   -1  'True
               ScrollBars      =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.CheckBox chkActivo 
               Height          =   255
               Left            =   6360
               TabIndex        =   54
               Top             =   1440
               Width           =   1335
               _Version        =   1441793
               _ExtentX        =   2355
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Activo "
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               Value           =   1
               Alignment       =   1
            End
            Begin XtremeSuiteControls.CheckBox chkDesembolso 
               Height          =   255
               Left            =   6360
               TabIndex        =   55
               Top             =   1800
               Width           =   1335
               _Version        =   1441793
               _ExtentX        =   2355
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Desembolso"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               Alignment       =   1
            End
            Begin XtremeSuiteControls.ComboBox cboPuesto 
               Height          =   330
               Left            =   1440
               TabIndex        =   57
               Top             =   1320
               Width           =   1815
               _Version        =   1441793
               _ExtentX        =   3201
               _ExtentY        =   582
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox cboRango 
               Height          =   330
               Left            =   1440
               TabIndex        =   58
               Top             =   3120
               Width           =   1575
               _Version        =   1441793
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox1"
            End
            Begin VB.Label Label21 
               Caption         =   "Fecha Elecc."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label lblUT 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   6000
               TabIndex        =   37
               Top             =   4920
               Width           =   1815
            End
            Begin VB.Label Label20 
               Caption         =   "U.T."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5280
               TabIndex        =   36
               Top             =   4920
               Width           =   375
            End
            Begin VB.Label Label13 
               Caption         =   "E-Mail"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   4080
               Width           =   615
            End
            Begin VB.Label Label10 
               Caption         =   "Celular"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   34
               Top             =   3600
               Width           =   855
            End
            Begin VB.Label Label7 
               Caption         =   "Teléfono"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   3600
               Width           =   975
            End
            Begin VB.Label Label5 
               Caption         =   "Nombre"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   3120
               Width           =   975
            End
            Begin VB.Line Line2 
               BorderColor     =   &H80000000&
               X1              =   1440
               X2              =   8280
               Y1              =   2880
               Y2              =   2880
            End
            Begin VB.Label Label4 
               Caption         =   "Datos Jefatura"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   2760
               Width           =   1695
            End
            Begin VB.Label Label12 
               Caption         =   "Celular"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3000
               TabIndex        =   30
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label lblTelefono 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1440
               TabIndex        =   29
               Top             =   4920
               Width           =   1395
            End
            Begin VB.Label lblCelular 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   3720
               TabIndex        =   28
               Top             =   4920
               Width           =   1395
            End
            Begin VB.Label Label18 
               Caption         =   "Teléfono"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   4920
               Width           =   1215
            End
            Begin VB.Label Label14 
               Caption         =   "E Mail"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   5280
               Width           =   975
            End
            Begin VB.Label lblMail 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1440
               TabIndex        =   25
               Top             =   5280
               Width           =   6375
            End
            Begin VB.Label Label19 
               Caption         =   "Referencia"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   4560
               Width           =   855
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000000&
               X1              =   960
               X2              =   8280
               Y1              =   4680
               Y2              =   4680
            End
            Begin VB.Label Label16 
               Caption         =   "Notas"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label Label15 
               Caption         =   "Puesto"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   1320
               Width           =   615
            End
         End
      End
      Begin VB.CheckBox chkActivos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Activos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -68760
         TabIndex        =   7
         Top             =   1080
         Value           =   1  'Checked
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2445
         Left            =   -74760
         TabIndex        =   38
         Top             =   4200
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4313
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Op"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Plan de Trabajo"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Monto"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Solicitud"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Liquidación"
            Object.Width           =   1940
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbNuevo 
         Height          =   360
         Left            =   -67320
         TabIndex        =   39
         Top             =   960
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Nuevo"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   3
         Left            =   -74760
         TabIndex        =   56
         Top             =   3840
         Width           =   8175
         _Version        =   1441793
         _ExtentX        =   14420
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Histórico"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   2
         Left            =   -73560
         TabIndex        =   43
         Top             =   3360
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Seleccione los Ejecutivos Autorizados para este Comité"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   1
         Left            =   -73560
         TabIndex        =   42
         Top             =   3360
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Seleccione las Actividades Autorizadas para este comité"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   41
         Top             =   3360
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Seleccione las Unidades Autorizadas para este comité"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
      Begin VB.Label Label8 
         Caption         =   "Presione doble click para borrar de la lista"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   -74640
         TabIndex        =   16
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Presione doble click para borrar de la lista"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   -74640
         TabIndex        =   15
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Presione doble click para borrar de la lista"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Miembros del Comité"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74280
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
      End
   End
   Begin ComCtl3.CoolBar CoolBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13830
      _ExtentX        =   24395
      _ExtentY        =   688
      _CBWidth        =   13830
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlb"
      MinHeight1      =   330
      Width1          =   3795
      NewRow1         =   0   'False
      MinHeight2      =   330
      Width2          =   2805
      NewRow2         =   0   'False
      Child3          =   "tlbBuscar"
      MinHeight3      =   330
      Width3          =   2235
      NewRow3         =   0   'False
      Begin VB.TextBox txtUnidadRelacionada 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   30
         Width           =   2535
      End
      Begin MSComctlLib.Toolbar tlbBuscar 
         Height          =   330
         Left            =   6825
         TabIndex        =   10
         ToolTipText     =   "Buscar UP Relacionada con el Comité"
         Top             =   30
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   582
         ButtonWidth     =   1640
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Buscar"
               Key             =   "Buscar1"
               ImageIndex      =   1
            EndProperty
         EndProperty
         Begin VB.TextBox txtConsultaComite 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   0
            Width           =   990
         End
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   9
         Top             =   30
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "borrar"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "consultar"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
   End
   Begin XtremeSuiteControls.CheckBox chkComiteActivo 
      Height          =   255
      Left            =   7200
      TabIndex        =   47
      Top             =   480
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Activo ?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtComite 
      Height          =   330
      Left            =   2280
      TabIndex        =   45
      Top             =   840
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcionComite 
      Height          =   330
      Left            =   3360
      TabIndex        =   48
      Top             =   840
      Width           =   5775
      _Version        =   1441793
      _ExtentX        =   10186
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboDirectores 
      Height          =   330
      Left            =   2280
      TabIndex        =   59
      Top             =   1200
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Director"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Comité"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1080
      TabIndex        =   1
      Top             =   825
      Width           =   945
   End
End
Attribute VB_Name = "frmAF_CD_Comites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim itm As ListItem


Dim vLista As Boolean
Dim vValidacion As Boolean
Dim vScroll As Boolean
Dim vVerificaComite As Boolean

Private Sub sbVerificaComite()
 
 vVerificaComite = False
 strSQL = "select U.cod_comite,C.descripcion,U.codigo_up from afi_cd_comites C inner join afi_cd_comites_unidades U " _
          & " on U.cod_comite = C.cod_comite " _
          & " where U.codigo_up = '" & txtComite.Text & "'"
          Call OpenRecordSet(rs, strSQL)
           
        If Not rs.EOF Then
         If rs!cod_comite <> rs!codigo_up Then
           MsgBox "Esta unidad ya pertenece al comite <" & rs!cod_comite & " " & rs!Descripcion & ">", vbInformation, "Información"
           vVerificaComite = True
         End If
        End If
     rs.Close

End Sub

Private Sub sbGuardaMiembros()

Dim Id_puesto As Integer
Dim Puesto_Asignado As Boolean
Dim vConsecutivo As Integer

vConsecutivo = fxConsecutivo("linea", "Comites_nombramientos_h")

  If txtCedula.Text = Empty Or lblNombre.Caption = Empty Or txtNombreJefe.Text = Empty Then
   MsgBox "Falta información", vbInformation, "Información"
   Exit Sub
  End If

  
 Select Case True
     Case cboPuesto.Text = "PRESIDENTE"
       Id_puesto = 1
     Case cboPuesto.Text = "VICEPRESIDENTE"
       Id_puesto = 2
     Case cboPuesto.Text = "SECRETARIO"
       Id_puesto = 3
     Case cboPuesto.Text = "TESORERO"
       Id_puesto = 4
     Case cboPuesto.Text = "FISCAL"
       Id_puesto = 5
     Case cboPuesto.Text = "VOCAL"
       Id_puesto = 6
     Case cboPuesto.Text = "VOCAL2"
       Id_puesto = 7
     Case cboPuesto.Text = "DELEGADO"
       Id_puesto = 8
 End Select
 
' 'Verifica que solo exista un miembro por puesto
' strSQL = "Select count(COD_PUESTO) from afi_cd_nombramientos " _
'        & "where COD_PUESTO = '" & Id_puesto & "' and cod_comite = '" & txtComite.Text & "'"
'
' call OpenRecordSet(rs, strsql)
'
'  If rs.Fields(0) > 0 Then
'       MsgBox " Solo se debe de asignar un miembro por puesto"
'       Exit Sub
'  End If
'
'  rs.Close
  
  strSQL = "select cedula from afi_cd_nombramientos " _
         & " where cedula ='" & txtCedula.Text & "' and cod_comite = '" & txtComite.Text & "'"
  Call OpenRecordSet(rs, strSQL)
 
 
 If rs.EOF Then
             
            'Ingresa datos en AFI_CD_NOMBRAMIENTOS
    strSQL = "insert into afi_cd_nombramientos (cedula,COD_PUESTO,APL_DESEMBOLSOS,notas,cod_comite,activo," _
              & " registro_fecha,registro_usuario,NOMBRE_JEFE,TELEFONO_JEFE,CELULAR_JEFE,CORREO_JEFE,RANGO_JEFE,FECHA_ELECCION)" _
              & " values('" & Trim(txtCedula.Text) & "','" & Id_puesto & "','" & chkDesembolso.Value & "'," _
              & " '" & txtNotas.Text & "','" & txtComite.Text & "','" & chkActivo.Value & "','" & Format(fxFechaServidor, "yyyymmdd") & "'," _
              & " '" & glogon.Usuario & "','" & txtNombreJefe.Text & "','" & txtTelefonoJefe.Text & "','" & txtCelularJefe.Text & "'," _
              & " '" & txtCorreoJefe.Text & "','" & cboRango.Text & "','" & Format(dtpFechaEleccion.Value, "yyyymmdd") & "')"
    
    Call ConectionExecute(strSQL)
    Call Bitacora("Ingresa", "Comite:" & txtComite.Text & " Nombramientos:" & txtCedula.Text & " Puesto: " & cboPuesto.Text & "")
            
            'Ingresa datos en AFI_CD_NOMBRAMIENTOS_H
    strSQL = "insert into afi_cd_nombramientos_h (cod_comite,cedula,cod_puesto,linea,APL_DESEMBOLSOS,registro_fecha," _
              & " registro_usuario,activo,FECHA_ELECCION)" _
              & " values('" & txtComite.Text & " ','" & txtCedula.Text & "','" & Id_puesto & "','" & vConsecutivo & "'," _
              & " '" & chkDesembolso.Value & "','" & Format(fxFechaServidor, "yyyymmdd") & "','" & glogon.Usuario & "'," _
              & " '" & chkActivo.Value & "','" & Format(dtpFechaEleccion.Value, "yyyymmdd") & "')"
              Call ConectionExecute(strSQL)
      Call Bitacora("Ingresa Historia", "Comite:" & txtComite.Text & " Nombramientos:" & txtCedula.Text & " Puesto: " & cboPuesto.Text & "")
 Else
     
            'Actualizar datos en AFI_CD_NOMBRAMIENTOS
    strSQL = "update afi_cd_nombramientos set APL_DESEMBOLSOS = '" & chkDesembolso.Value & "'" _
           & ", registro_usuario ='" & glogon.Usuario & "',notas='" & txtNotas.Text & "'" _
           & ", activo='" & chkActivo.Value & "',cod_puesto = '" & Id_puesto & "' " _
           & ", NOMBRE_JEFE = '" & txtNombreJefe.Text & "',TELEFONO_JEFE = '" & txtTelefonoJefe.Text & "'" _
           & ", CELULAR_JEFE = '" & txtCelularJefe.Text & "',CORREO_JEFE = '" & txtCorreoJefe.Text & "'" _
           & ", RANGO_JEFE= '" & cboRango.Text & "', FECHA_ELECCION='" & Format(dtpFechaEleccion.Value, "yyyymmdd") & "'" _
           & " where cod_comite ='" & txtComite.Text & "' and " _
           & " cedula='" & txtCedula.Text & "' "
    Call ConectionExecute(strSQL)
            
            'Ingresa datos en AFI_CD_NOMBRAMIENTOS_H
    strSQL = "insert into afi_cd_nombramientos_h (cod_comite,cedula,cod_puesto,linea,APL_DESEMBOLSOS,registro_fecha" _
           & ", registro_usuario,activo,FECHA_ELECCION)" _
           & " values('" & txtComite.Text & " ','" & txtCedula.Text & "','" & Id_puesto & "','" & vConsecutivo & "'" _
           & ", '" & chkDesembolso.Value & "','" & Format(fxFechaServidor, "yyyymmdd") & "','" & glogon.Usuario & "'" _
           & ", '" & chkActivo.Value & "','" & Format(dtpFechaEleccion.Value, "yyyymmdd") & "')"
    Call ConectionExecute(strSQL)
    
  
  End If
 rs.Close

 MsgBox "Miembro(s) Aplicado(s) al comité", vbInformation, "Información"


End Sub

Private Sub sbHistorialMiembros(vCedula As String)
Dim vPuesto As String

lswMiembros_H.ListItems.Clear

strSQL = "select S.cedula,S.nombre,H.cod_puesto,H.registro_fecha,H.registro_usuario, " _
         & "H.apl_desembolsos,H.activo,H.cod_comite,H.Nombre_Jefe, H.Telefono_Jefe, " _
         & "H.Celular_Jefe,H.Correo_Jefe,H.Rango_Jefe " _
         & "from afi_cd_nombramientos H INNER join socios S on H.cedula = S.cedula " _
         & "where H.COD_COMITE = '" & txtComite & "' order by H.cod_puesto"
Call OpenRecordSet(rs, strSQL)
    
Do While Not rs.EOF
   Select Case True
     Case rs!cod_puesto = 1
           vPuesto = "PRESIDENTE"
     Case rs!cod_puesto = 2
           vPuesto = "VICEPRESIDENTE"
     Case rs!cod_puesto = 3
           vPuesto = "SECRETARIO"
     Case rs!cod_puesto = 4
           vPuesto = "TESORERO"
     Case rs!cod_puesto = 5
           vPuesto = "FISCAL"
     Case rs!cod_puesto = 6
           vPuesto = "VOCAL"
     Case rs!cod_puesto = 7
           vPuesto = "VOCAL2"
   End Select
   
   
Set itmX = lswMiembros_H.ListItems.Add(, , Trim(IIf(IsNull(rs!Cedula), "", rs!Cedula)))
    itmX.SubItems(1) = Trim(IIf(IsNull(rs!Nombre), "", rs!Nombre))
    itmX.SubItems(2) = Trim(IIf(IsNull(vPuesto), "", vPuesto))
    itmX.SubItems(3) = Trim(IIf(IsNull(rs!Registro_Fecha), "", rs!Registro_Fecha))
    itmX.SubItems(4) = Trim(IIf(IsNull(rs!Registro_Usuario), "", rs!Registro_Usuario))
    
    If rs!apl_desembolsos = 1 Then
        itmX.SubItems(5) = "DESEMBOLSO"
    Else
        itmX.SubItems(5) = "SINDESEMBOLSO"
    End If
    
    If rs!activo = 1 Then
      itmX.SubItems(6) = "ACTIVO"
    Else
      itmX.SubItems(6) = "INACTIVO"
    End If
    
    itmX.SubItems(7) = Trim(IIf(IsNull(rs!cod_comite), "", rs!cod_comite))

 rs.MoveNext
Loop
rs.Close
 
End Sub

Private Sub sbIngresoActividades()

'Ingreso en la tabla de afi_cd_actividades
  If lswActividades.ListItems.Count > 0 Then
      strSQL = "select count(*) as 'Existe' from afi_cd_comites_actividades" _
             & " where cod_comite = '" & txtComite.Text & "' and cod_actividad = '" & lswSelecBusqueda.SelectedItem.Text & "'"
      Call OpenRecordSet(rs, strSQL)

         If rs!Existe = 0 Then
            strSQL = "insert afi_cd_comites_actividades (cod_comite,cod_actividad) values('" & txtComite.Text & "'," _
                     & "'" & lswSelecBusqueda.SelectedItem.Text & "' ) "
            Call ConectionExecute(strSQL)
            Call Bitacora("Ingresa", "Actividades:" & lswSelecBusqueda.SelectedItem.Text & " Comite: " & txtComite.Text & "")
         
         End If
       rs.Close
  End If

End Sub

Private Sub sbIngresoComites()

On Error GoTo vError
 
strSQL = "select count(*) as 'Existe' from afi_cd_comites where cod_comite = '" & txtComite.Text & "'"
Call OpenRecordSet(rs, strSQL)
 
If rs!Existe = 0 Then
    'Ingreso en la tabla de afi_cd_Comites
    strSQL = "Insert afi_cd_comites(cod_comite, cod_director, Descripcion, activo, registro_usuario, registro_fecha) " _
             & "values ('" & txtComite.Text & "','" & cboDirectores.ItemData(cboDirectores.ListIndex) & "','" & txtDescripcionComite.Text & "'," _
             & "'" & chkComiteActivo.Value & "','" & glogon.Usuario & "','" & Format(fxFechaServidor, "yyyymmdd") & "')"
    Call ConectionExecute(strSQL)
End If
rs.Close

If lswComites.ListItems.Count > 0 Then
    strSQL = "select count(*) as 'Existe' from afi_cd_comites_unidades" _
           & " where cod_comite = '" & txtComite.Text & "' and codigo_up = '" & lswSelecBusqueda.SelectedItem.Text & "'"
    Call OpenRecordSet(rs, strSQL)
    
       If rs!Existe = 0 Then
          strSQL = "insert afi_cd_comites_unidades (cod_comite,codigo_up) values('" & txtComite.Text _
                 & "', '" & lswSelecBusqueda.SelectedItem.Text & "')"
          Call ConectionExecute(strSQL)
          Call Bitacora("Ingresa", "Comite:" & lswComites.SelectedItem.Text & " Comite: " & txtComite.Text & "")
       End If
     rs.Close
End If


Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbIngresoEjecutivos()

'Ingreso en la tabla de afi_cd_ejecutivo
  If lswEjecutivo.ListItems.Count > 0 Then
     
       strSQL = "select count(*) as 'Existe' from afi_cd_comites_ejecutivo" _
              & " where cod_comite = '" & txtComite.Text & "' and id_promotor = '" & lswSelecBusqueda.SelectedItem.Text & "'"
        Call OpenRecordSet(rs, strSQL)
     
         If rs!Existe = 0 Then
            strSQL = "insert afi_cd_comites_ejecutivo (cod_comite,id_promotor) values('" & txtComite.Text _
                   & "', '" & lswSelecBusqueda.SelectedItem.Text & "' ) "
            Call ConectionExecute(strSQL)
            Call Bitacora("Ingresa", "Ejecutivo:" & lswSelecBusqueda.SelectedItem.Text & " Comite: " & txtComite.Text & "")
         End If
       rs.Close
  
  End If

End Sub

Private Sub sbLimpia()
Dim vControl As Control
Dim x As Integer, y As Integer

SSTab1.Tab = 0

SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False

For Each vControl In Me
  If TypeOf vControl Is TextBox Then
     vControl.Text = ""
  End If
Next

For Each vControl In Me
  If TypeOf vControl Is ListView Then
     vControl.ListItems.Clear
  End If
Next

For Each vControl In Me
  If TypeOf vControl Is CheckBox Then
     vControl.Value = 0
  End If
Next

chkAsociaUnidad.Visible = True
chkAsociaUnidad.Value = 0
chkComiteActivo.Value = 1

End Sub

Private Sub sbLLamaAsociado()

strSQL = "select cedula,nombre from socios where cedula = '" & txtCedula.Text & "' and estadoactual = 'S'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
   txtCedula.Text = rs!Cedula
   lblNombre.Caption = rs!Nombre
End If
rs.Close

End Sub

Private Sub sbMiembrosActivosComite(vComite As String)

'Llama Miembros

strSQL = "select N.cedula, S.Nombre, N.cod_puesto, N.notas, N.apl_desembolsos, N.activo" _
         & " from afi_cd_comites C" _
         & "   left join afi_cd_nombramientos N on C.cod_comite = N.cod_comite" _
         & "   left join socios S on N.cedula = S.cedula" _
         & " where C.cod_comite = '" & vComite & "' and N.activo = " & chkActivos.Value
Call OpenRecordSet(rs, strSQL)


Do While Not rs.EOF
'Llama Ejecutivos
 Set itmX = lswMiembrosComite.ListItems.Add(, , IIf(IsNull(rs!Cedula), "", rs!Cedula))
     itmX.SubItems(1) = Trim(IIf(IsNull(rs!Nombre), "", rs!Nombre))
     Select Case True
      Case rs!cod_puesto = 1
       itmX.SubItems(2) = "PRESIDENTE"
      Case rs!cod_puesto = 2
       itmX.SubItems(2) = "VICEPRESIDENTE"
      Case rs!cod_puesto = 3
       itmX.SubItems(2) = "SECRETARIO"
      Case rs!cod_puesto = 4
       itmX.SubItems(2) = "TESORERO"
      Case rs!cod_puesto = 5
       itmX.SubItems(2) = "FISCAL"
      Case rs!cod_puesto = 6
       itmX.SubItems(2) = "VOCAL"
      Case rs!cod_puesto = 7
       itmX.SubItems(2) = "VOCAL2"
     End Select
     
     Select Case True
       Case rs!activo = 1
         itmX.SubItems(3) = "SI"
       Case Else
         itmX.SubItems(3) = "NO"
     End Select
     
     Select Case True
       Case rs!apl_desembolsos = 1
         itmX.SubItems(4) = "SI"
       Case Else
         itmX.SubItems(4) = "NO"
     End Select
     itmX.SubItems(5) = Trim(IIf(IsNull(rs!NOTAS), "", rs!NOTAS))
  rs.MoveNext
Loop
rs.Close

End Sub

Sub sbModificarComite()
If txtComite.Text = Empty And txtDescripcionComite.Text = Empty Then
 MsgBox "No hay ningun comité para procesar", vbInformation, "Información"
 Exit Sub
End If
 'Actualizar afi_cd_Comites
      strSQL = "update afi_cd_comites set cod_director = '" & cboDirectores.ItemData(cboDirectores.ListIndex) & "'," _
               & "descripcion = '" & txtDescripcionComite.Text & "',activo = '" & chkComiteActivo.Value & "'," _
               & "modifica_usuario = '" & glogon.Usuario & "',modifica_fecha = '" & Format(fxFechaServidor, "yyyymmdd") & "'" _
               & "where cod_comite = '" & txtComite.Text & "'"
               Call ConectionExecute(strSQL)
               MsgBox "Comite Modificado", vbInformation, "Información"

End Sub


Private Function fxValida() As Boolean

'Dim vMensaje As String
'
'vMensaje = ""
'fxValida = True
'
''If CboPuesto.Text = "" Then vMensaje = vMensaje & vbCrLf & "No se definio los puestos"
'If txtComites.Text = "" Then vMensaje = vMensaje & vbCrLf & "No se definio el comite"
'If TxtCedula.Text = "" Then vMensaje = vMensaje & vbCrLf & "No se especifico una cédula"
'
'
'If Len(vMensaje) > 0 Then
'  fxValida = False
'  MsgBox vMensaje, vbCritical
'End If

End Function

Function FxNomComite(vUnidad As String)
   Dim rs As New ADODB.Recordset
   Dim strSQL As String
  
   strSQL = "select cod_comite,descripcion from afi_cd_comites where cod_comite = '" & vUnidad & "'"
            Call OpenRecordSet(rs, strSQL)
   If rs.EOF Then
      FxNomComite = "No existe unidad definida en Comites y Delegados"
   Else
      FxNomComite = rs!Descripcion
   End If

End Function

Private Function fxConsecutivo(vCampo As String, vTabla As String) As Integer

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select coalesce(Max( " & vCampo & " ),0) as Consecutivo from " & vTabla & ""

Call OpenRecordSet(rs, strSQL)
  fxConsecutivo = rs!Consecutivo + 1
rs.Close

End Function

'
'Public Function fxCodigoCbo(cbo As ComboBox) As String
'
'Dim i As Integer, vPaso As Boolean
'Dim x As Integer
'
'If cbo.ListCount = 0 Then
'  fxCodigoCbo = ""
'  Exit Function
'End If
'
'vPaso = True
'i = 1
'x = Len(cbo.Text)
'Do While vPaso
'  If Mid(cbo.Text, i, 1) = "-" Then
'    vPaso = False
'    i = i - 1
'  Else
'    i = i + 1
'  End If
'  If i = x Then Exit Do
'Loop
'fxCodigoCbo = Trim(Mid(cbo.Text, 1, i))
'
'End Function



Private Sub sbLlamaDirectores()


cboDirectores.Clear

strSQL = "select cod_director, nombre from afi_cd_directores"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
    cboDirectores.AddItem Trim(rs!Nombre)
    cboDirectores.ItemData(cboDirectores.NewIndex) = rs!cod_director
rs.MoveNext
Loop

If rs.RecordCount > 0 Then
 rs.MoveFirst
 cboDirectores.Text = Trim(rs!Nombre)
End If
rs.Close

End Sub

Sub sbCargaComite()


Dim x As Integer
Dim A As Integer
Dim B As Integer
Dim ver As Integer
Dim vPuesto As Integer
Dim vEstadoLiqu As String


vVerificaComite = False
lswComites.ListItems.Clear
lswActividades.ListItems.Clear
lswEjecutivo.ListItems.Clear
lswMiembrosComite.ListItems.Clear
lswLiquidaciones.ListItems.Clear

txtUnidadRelacionada.Text = Empty
'Llama a comite seleccionado

strSQL = "select E.codigo_up, U.descripcion as 'UP', C.descripcion, C.cod_director, C.activo, D.nombre" _
       & " from afi_cd_comites C" _
       & "   left join afi_cd_comites_unidades E on C.cod_comite = E.cod_comite" _
       & "   left join uprogramatica U on E.codigo_up = U.codigo" _
       & "   left join afi_cd_directores D on C.cod_director = D.cod_director " _
       & " where C.cod_comite = '" & txtComite.Text & "'" _
       & " group by E.codigo_up, U.descripcion,C.descripcion, C.cod_director, C.activo, D.nombre"
Call OpenRecordSet(rs, strSQL)

     If Not rs.EOF Then
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = True
        SSTab1.TabEnabled(4) = True
        
        chkComiteActivo.Value = rs!activo
        txtDescripcionComite = rs!Descripcion
        cboDirectores.Text = rs!Nombre
        chkAsociaUnidad.Visible = False
        
        lswComites.Checkboxes = True
        While Not rs.EOF
            'LLama comites agrupados
            
             Set itmX = lswComites.ListItems.Add(, , IIf(IsNull(rs!codigo_up), "", rs!codigo_up))
             itmX.Checked = True
             itmX.SubItems(1) = Trim(IIf(IsNull(rs!up), "", rs!up))
          rs.MoveNext
         Wend
   Else
     rs.Close
       strSQL = "select Top 1 codigo,descripcion from uprogramatica where codigo = '" & txtComite.Text & "'"
        Call OpenRecordSet(rs, strSQL)
         If Not rs.EOF Then
             txtComite.Text = rs!Codigo
             txtDescripcionComite.Text = rs!Descripcion
             chkAsociaUnidad.Visible = True
         End If
   End If
rs.Close

strSQL = "select C.cod_comite,C.descripcion as NombreComite,U.codigo,U.descripcion as UpNombre " _
         & " from afi_cd_comites C inner join uprogramatica U on C.cod_comite = U.codigo where C.cod_comite = '" & txtComite.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
     txtUnidadRelacionada.Text = Trim(rs!upnombre)
End If
rs.Close

strSQL = "select A.cod_actividad,T.descripcion from afi_cd_comites C left join afi_cd_comites_actividades A " _
         & "on C.cod_comite = A.cod_comite left join afi_cd_actividades T on A.cod_actividad = T.cod_actividad " _
         & " where C.cod_comite = '" & txtComite.Text & "'"
Call OpenRecordSet(rs, strSQL)
            
While Not rs.EOF
        'Llama acividades
         Set itmX = lswActividades.ListItems.Add(, , rs!Cod_Actividad & "")
         itmX.SubItems(1) = rs!Descripcion & ""
         rs.MoveNext
Wend
rs.Close
            
'Llama Ejecutivos
strSQL = "select V.id_promotor,P.nombre" _
       & " from afi_cd_comites C left join afi_cd_comites_ejecutivo V on C.cod_comite = V.cod_comite" _
       & "  left join promotores P on V.id_promotor = P.id_promotor " _
       & " where C.cod_comite = '" & txtComite.Text & "'"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
    Set itmX = lswEjecutivo.ListItems.Add(, , rs!id_promotor & "")
        itmX.SubItems(1) = rs!Nombre & ""
    rs.MoveNext
Loop
rs.Close

txtDescripcionComite.SetFocus


'Llama los miembros activos del comite
Call sbMiembrosActivosComite(txtComite.Text)

'Carga los datos de la liquidaciones pendientes
strSQL = "select A.noperacion,C.notas,sum(A.monto)as Monto,C.estado,C.tesoreria_nsolicitud,C.liquida_fecha " _
         & " from afi_cd_cuentas C inner join  afi_cd_cuentas_actividades A on C.noperacion = A.noperacion " _
         & " where C.cod_comite = '" & txtComite.Text & "' and estado ='T'" _
         & " group by C.notas,A.noperacion,C.estado,C.tesoreria_nsolicitud,C.liquida_fecha"
Call OpenRecordSet(rs, strSQL)


Do While Not rs.EOF
      Set itmX = lswLiquidaciones.ListItems.Add(, , Trim(rs!Noperacion))
      itmX.SubItems(1) = rs!NOTAS & ""
      itmX.SubItems(2) = Format(rs!Monto, "Standard")
      itmX.SubItems(3) = rs!TESORERIA_NSOLICITUD & ""
      itmX.SubItems(4) = Format(rs!LIQUIDA_FECHA, "dd/mm/yyyy")
  rs.MoveNext
Loop
rs.Close

'Carga los datos de los mensajes
Call sbCargarMensajes

End Sub

Private Sub sbCargarMensajes()

Me.MousePointer = vbHourglass

 
strSQL = "Select MENSAJE,FECHA,VENCIMIENTO,USUARIO" _
     & " From AFI_CD_COMITES_MENSAJES" _
     & " where COD_COMITE = '" & txtComite.Text & "' and VENCIMIENTO >= getdate()" _
     & " order by FECHA desc"
Call OpenRecordSet(rs, strSQL)

With gridMensajes

.MaxRows = 0

Do While Not rs.EOF
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
    .AutoSize = True
    .Col = 1
    .Text = rs!Mensaje
    
    .Col = 1
    .CellNote = ">> Datos de Registro <<" & vbCrLf _
              & "Fecha   : " & Format(rs!Fecha, "dd/mm/yyyy") & vbCrLf _
              & "Vence   : " & rs!Vencimiento & vbCrLf _
              & "Usuario : " & rs!Usuario & ""
     
   rs.MoveNext
Loop
rs.Close
    
End With
    
Me.MousePointer = vbDefault
    
End Sub

Sub sbCargaElementos()

Select Case True

  Case SSTab1.Tab = 0
     txtCod_Comite_Ingreso.Text = ""
     txtDesc_Comite_Ingreso.Text = ""
     fraBusquedaGeneral.Visible = False
  Case SSTab1.Tab = 1
     fraBusquedaGeneral.Visible = True
  Case SSTab1.Tab = 2
     fraBusquedaGeneral.Visible = True
  Case SSTab1.Tab = 3
     fraBusquedaGeneral.Visible = False
     tlbNuevo.Visible = True
     lswMiembrosComite.Visible = True

  Case SSTab1.Tab = 4
     fraBusquedaGeneral.Visible = False
End Select
End Sub

Sub sbCargaListView(View_Uno As Object, View_Segundo As Object)

Dim itmX As ListViewItem
Dim A As Integer, B As Integer
    
For A = 1 To View_Uno.ListItems.Count
     
  If View_Uno.ListItems.Item(A).Selected = True Then
         
         For B = 1 To View_Segundo.ListItems.Count
           If View_Segundo.ListItems.Item(B) = View_Uno.ListItems.Item(A) Then
              MsgBox "Ya se encuentra asignado " & View_Uno.ListItems(A).SubItems(1) & "  en el listado", vbInformation, "Información"
              vValidacion = True
              Exit Sub
           End If
         Next B
         
  Set itmX = View_Segundo.ListItems.Add(, , View_Uno.ListItems.Item(A))
      itmX.SubItems(1) = View_Uno.ListItems(A).SubItems(1)
  End If

Next A

End Sub

Private Sub btnMensaje_Click(Index As Integer)
Select Case Index
  Case 0 '"Agregar"
     GLOBALES.gTag = txtComite.Text
     Call sbFormsCall("frmAF_CD_Plan", 1, , , False, Me)
  
  Case 1 '"Actualizar"
     Call sbCargarMensajes
 End Select
 
End Sub

Private Sub btnMiembros_Click(Index As Integer)
Select Case Index
   Case 0 '"NUEVO"
        fraMiembros.Visible = True
        fraBuscaMiembros.Visible = True
        tabGeneral.Tab = 0
        tabGeneral.Visible = True
        
'        tlbMiembros.Visible = True
        
        txtCedula.Text = Empty
        lblNombre.Caption = Empty
        chkDesembolso.Value = 0
        chkActivo.Value = 1
        txtNotas.Text = Empty
        lblTelefono.Caption = Empty
        lblCelular.Caption = Empty
        lblMail.Caption = Empty
        lblUT.Caption = Empty
        lswMiembros_H.ListItems.Clear
'        tlbMiembros.Visible = True
        
    Case 1 '"GUARDAR", "SALVAR"
      Call sbGuardaMiembros
      lswMiembrosComite.ListItems.Clear
      Call sbDatosMiembro(txtCedula.Text)
      Call sbHistorialMiembros(txtCedula.Text)
      Call sbMiembrosActivosComite(txtComite.Text)
      
    Case 2 '"CERRAR", "SALIR"
       fraMiembros.Visible = False
       fraBuscaMiembros.Visible = False
       tabGeneral.Visible = False

End Select

End Sub

Private Sub CboPuesto_Click()
'Dim IDpuesto As Integer
'
'IDpuesto = fxCodigoCbo(CboPuesto)
'
'strSQL = "select cod_puesto from afi_cd_nombramientos where cod_puesto = " & IDpuesto & " and activo = 1 " _
'         & "and cod_comite = '" & txtComite.Text & "' group by cod_puesto "
'         Call OpenRecordSet(rs, strSQL)
'         If Not rs.EOF Then
'            MsgBox "Ya este puesto esta seleccionado, debe de seleccionar otro", vbInformation, "Información"
'         End If
'rs.Close

End Sub

Private Sub sbDatosMiembro(vCedula As String)

Dim vPuesto As Integer
Dim vTipo As Integer

'Datos del Asociado perteneciente a la unidad
  
strSQL = "select S.cedula, S.nombre, UT.UT_DESCRIPCION, T.numero, T.tipo, af_email, N.cod_puesto, N.notas" _
    & ", N.activo,N.apl_desembolsos, N.Nombre_Jefe, N.Telefono_Jefe, N.Celular_Jefe, N.Correo_Jefe, N.Rango_Jefe, N.FECHA_ELECCION " _
    & " from afi_cd_nombramientos N right join socios S on N.cedula = S.cedula " _
    & " inner join UTRABAJO UT on UT.UT_CODIGO = S.UT " _
    & " left join telefonos T on S.cedula = T.cedula " _
    & " where S.cedula = '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
   txtCedula.Text = IIf(IsNull(rs!Cedula), 0, rs!Cedula)
   lblMail.Caption = IIf(IsNull(rs!AF_Email), "@", rs!AF_Email)
   lblNombre.Caption = IIf(IsNull(rs!Nombre), "", rs!Nombre)
   chkActivo.Value = IIf(IsNull(rs!activo), 0, rs!activo)
   chkDesembolso.Value = IIf(IsNull(rs!apl_desembolsos), 0, rs!apl_desembolsos)
   txtNotas.Text = IIf(IsNull(rs!NOTAS), "", rs!NOTAS)
   lblUT.Caption = IIf(IsNull(rs!UT_Descripcion), 0, rs!UT_Descripcion)
   txtNombreJefe.Text = IIf(IsNull(rs!Nombre_Jefe), "", rs!Nombre_Jefe)
   txtTelefonoJefe.Text = IIf(IsNull(rs!Telefono_Jefe), "", rs!Telefono_Jefe)
   txtCelularJefe.Text = IIf(IsNull(rs!Celular_Jefe), "", rs!Celular_Jefe)
   txtCorreoJefe.Text = IIf(IsNull(rs!Correo_Jefe), "", rs!Correo_Jefe)
   cboRango.Text = IIf(IsNull(rs!Rango_Jefe), "SD", rs!Rango_Jefe)
   dtpFechaEleccion.Value = Format(IIf(IsNull(rs!FECHA_ELECCION), fxFechaServidor, rs!FECHA_ELECCION), "dd/mm/yyyy")

vPuesto = IIf(IsNull(rs!cod_puesto), 0, rs!cod_puesto)
  If vPuesto <> 0 Then
   Select Case True
     Case rs!cod_puesto = 1
        cboPuesto.Text = "PRESIDENTE"
     Case rs!cod_puesto = 2
        cboPuesto.Text = "VICEPRESIDENTE"
     Case rs!cod_puesto = 3
        cboPuesto.Text = "SECRETARIO"
     Case rs!cod_puesto = 4
        cboPuesto.Text = "TESORERO"
     Case rs!cod_puesto = 5
        cboPuesto.Text = "FISCAL"
     Case rs!cod_puesto = 6
        cboPuesto.Text = "VOCAL"
     Case rs!cod_puesto = 7
        cboPuesto.Text = "VOCAL2"
     Case rs!cod_puesto = 8
        cboPuesto.Text = "DELEGADO"
  End Select
  End If
End If


Do While Not rs.EOF
    vTipo = IIf(IsNull(rs!Tipo), 0, rs!Tipo)
    
    If vTipo <> 0 Then
      Select Case True
         Case rs!Tipo = 1
              lblTelefono.Caption = rs!Numero
         Case rs!Tipo = 3
              lblCelular.Caption = rs!Numero
      End Select
    End If
    
    rs.MoveNext
Loop
   
rs.Close

End Sub

Sub sbVerificacion_Complementos()

strSQL = "select top 1 cod_director from afi_cd_directores"
          Call OpenRecordSet(rs, strSQL)
           If rs.EOF Then
             MsgBox "Debe de ingresar los directores antes de procesar la información de los comites", vbInformation, "Información"
           End If
       rs.Close
strSQL = "select top 1 cod_actividad from afi_cd_actividades"
          Call OpenRecordSet(rs, strSQL)
           If rs.EOF Then
             MsgBox "Debe de ingresar las actividades antes de procesar la información de los comites", vbInformation, "Información"
           End If
       rs.Close
End Sub

Private Sub chkActivos_Click()
 lswMiembrosComite.ListItems.Clear
 Call sbMiembrosActivosComite(txtComite.Text)
End Sub

Private Sub chkAsociaUnidad_Click()

If chkAsociaUnidad.Value = 0 Then
 chkAsociaUnidad.Visible = False
Else
 chkAsociaUnidad.Visible = True
 txtDescripcionComite.Text = Empty
 txtDescripcionComite.SetFocus
End If

End Sub

Private Sub cmdAgregar_Click()
If txtCedula.Text = Empty Or lblNombre.Caption = Empty Then
  MsgBox "No se encuentran datos para ingresar", vbInformation, "Información"
  Exit Sub
End If

strSQL = "select cod_comite from afi_cd_nombramientos where cedula = '" & txtCedula.Text & "' and activo = 1"
         Call OpenRecordSet(rs, strSQL)
     If Not rs.EOF Then
         MsgBox "Este asociado es miembro activo del comite " & rs!cod_comite & "", vbInformation, "Información"
          txtCedula.Text = Empty
          lblNombre.Caption = Empty
          txtCedula.SetFocus
         Exit Sub
     End If
rs.Close


txtCedula.Text = Empty
lblNombre.Caption = Empty
txtCedula.SetFocus

End Sub

Private Sub FlatScrollBar_Change()

Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
fraMiembros.Visible = False


If vScroll Then
    strSQL = "select Top 1 cod_comite from afi_cd_comites"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_comite > '" & txtComite.Text & "' order by cod_comite asc"
    Else
       strSQL = strSQL & " where cod_comite < '" & txtComite.Text & "' order by cod_comite desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    
    
    
    If Not rs.EOF And Not rs.BOF Then
      txtComite.Text = rs!cod_comite
      Call txtComite_KeyDown(vbKeyReturn, 0)
    End If

End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
   vModulo = 40
End Sub

Private Sub Form_Load()
 vModulo = 40
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbToolBarIconos(tlb)
 
 With lswComites.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripcion", lswComites.Width - 1400
 End With
 
 With lswActividades.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripcion", lswComites.Width - 1400
 End With
 
 With lswEjecutivo.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripcion", lswComites.Width - 1400
 End With
 
 With lswSelecBusqueda.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripcion", lswComites.Width - 1400
 End With
 lswSelecBusqueda.Checkboxes = True
 
With lswLiquidaciones.ColumnHeaders
    .Clear
    .Add , , "N.Operación", 1200
    .Add , , "Notas", 2100
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Bancos Id", 1200, vbCenter
    .Add , , "F.Liquida", 1800
End With
 
With lswMiembrosComite.ColumnHeaders
    .Clear
    .Add , , "Cédula", 1200
    .Add , , "Nombre", 3100
    .Add , , "Puesto", 2500
    .Add , , "Activo?", 1200, vbCenter
    .Add , , "Desembolsa?", 1200, vbCenter
    .Add , , "Notas", 2500
End With
 
 
With lswMiembros_H.ColumnHeaders
    .Clear
    .Add , , "Cédula", 1200
    .Add , , "Nombre", 3100
    .Add , , "Puesto", 2500
    
    .Add , , "R.Fecha", 2800
    .Add , , "R.Usuario", 2500, vbCenter
    
    .Add , , "Desembolsa?", 1200, vbCenter
    .Add , , "Activo?", 1200, vbCenter
    .Add , , "Comité", 2500
End With
 
 
 
' tlbMiembros.ImageList = frmContenedor.imgToolbarIcons01
' tlbMiembros.Buttons.Item(1).Image = 1
' tlbMiembros.Buttons.Item(2).Image = 7
' tlbMiembros.Buttons.Item(4).Image = 5

' tlbMiembros.Visible = True
 
 Call sbToolBar(tlb, "nuevo")
 
 Call sbVerificacion_Complementos
 chkAsociaUnidad.Visible = False
 Call sbCargaElementos
 
strSQL = "select cod_director as 'IdX', rtrim(Nombre) as 'Itmx'from afi_cd_directores"
Call sbCbo_Llena_New(cboDirectores, strSQL, False, True)

 fraBusquedaGeneral.Visible = True
 fraBusquedaGeneral.Top = 2160
 fraBusquedaGeneral.Left = 360
 
 SSTab1.Tab = 0
 
 SSTab1.TabEnabled(1) = False
 SSTab1.TabEnabled(2) = False
 SSTab1.TabEnabled(3) = False
 SSTab1.TabEnabled(4) = False
 
 vLista = True

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 cboPuesto.AddItem ("PRESIDENTE")
 cboPuesto.AddItem ("VICEPRESIDENTE")
 cboPuesto.AddItem ("SECRETARIO")
 cboPuesto.AddItem ("TESORERO")
 cboPuesto.AddItem ("FISCAL")
 cboPuesto.AddItem ("VOCAL")
 cboPuesto.AddItem ("VOCAL2")
 cboPuesto.AddItem ("DELEGADO")
 cboPuesto.Text = "PRESIDENTE"

 cboRango.AddItem ("Lic.")
 cboRango.AddItem ("Msc.")
 cboRango.AddItem ("Ing.")
 cboRango.AddItem ("Dr.")
 cboRango.AddItem ("Bach.")
 cboRango.AddItem ("Sr.")
 cboRango.AddItem ("Sra.")
 cboRango.AddItem ("Srta.")
 cboRango.AddItem ("SD")
 cboRango.Text = "Sr."

 gridMensajes.MaxCols = 1
 gridMensajes.MaxRows = 0
 gridMensajes.TextTip = TextTipFixed
 gridMensajes.TextTipDelay = 1000
 
 dtpFechaEleccion.Value = fxFechaServidor
 
 If GLOBALES.gTag <> Empty Then
   txtComite.Text = GLOBALES.gTag
 End If
 
End Sub

Private Sub lswActividades_DblClick()

'Borrar afi_cd_actividades

If lswActividades.ListItems.Count > 0 Then
      strSQL = "delete afi_cd_comites_actividades" _
             & " where cod_actividad = '" & lswActividades.SelectedItem.Text & "' and cod_comite = '" & txtComite.Text & "' "
      Call ConectionExecute(strSQL)
      Call Bitacora("Borra", "Actividades:" & lswActividades.SelectedItem.Text & " Comite: " & txtComite.Text & "")
                 
End If
lswActividades.ListItems.Clear

strSQL = "select U.cod_actividad,A.descripcion" _
       & " from afi_cd_comites C" _
       & "  left join afi_cd_comites_actividades U on C.cod_comite = U.cod_comite" _
       & "  left join afi_cd_actividades A on A.cod_actividad = U.cod_actividad " _
       & " where C.cod_comite ='" & txtComite.Text & "' and U.cod_actividad is not null"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    Set itmX = lswActividades.ListItems.Add(, , rs!Cod_Actividad & "")
                  itmX.SubItems(1) = rs!Descripcion & ""
    rs.MoveNext
Loop
rs.Close

End Sub


Private Sub lswComites_DblClick()

'Borrar afi_cd_comites_unidades
If lswComites.ListItems.Count > 0 Then
      strSQL = "delete afi_cd_comites_unidades where codigo_up = '" & lswComites.SelectedItem.Text & "' " _
               & " and cod_comite = '" & txtComite.Text & "' "
      Call ConectionExecute(strSQL)
      Call Bitacora("Borra", "Comite:" & txtComite.Text & " Unidad: " & lswComites.SelectedItem.Text & "")
End If

lswComites.ListItems.Clear

strSQL = "select U.codigo_up,P.descripcion" _
       & "  from afi_cd_comites C left join afi_cd_comites_unidades U " _
       & "on C.cod_comite = U.cod_comite left join uprogramatica P on P.codigo = U.codigo_up " _
       & "where C.cod_comite ='" & txtComite.Text & "' and U.codigo_up is not null"
Call OpenRecordSet(rs, strSQL)
    
Do While Not rs.EOF
    Set itmX = lswComites.ListItems.Add(, , rs!codigo_up & "")
        itmX.SubItems(1) = rs!Descripcion & ""
    rs.MoveNext
Loop
rs.Close

End Sub


Private Sub lswEjecutivo_DblClick()

'Borrar afi_cd_comites_ejecutivo

If lswEjecutivo.ListItems.Count > 0 Then
      strSQL = "delete afi_cd_comites_ejecutivo where id_promotor = '" & lswEjecutivo.SelectedItem.Text & "' " _
               & " and cod_comite = '" & txtComite.Text & "' "
               
      Call ConectionExecute(strSQL)
      Call Bitacora("Borra", "Ejecutivo:" & lswEjecutivo.SelectedItem.Text & " Unidad: " & txtComite.Text & "")
               
End If

lswEjecutivo.ListItems.Clear

strSQL = "select E.id_promotor,P.nombre" _
       & " from afi_cd_comites C" _
       & " left join afi_cd_comites_ejecutivo E on C.cod_comite = E.cod_comite" _
       & " left join promotores P on E.id_promotor = P.id_promotor " _
       & " where C.cod_comite ='" & txtComite.Text & "' and E.id_promotor is not null"
Call OpenRecordSet(rs, strSQL)
    
Do While Not rs.EOF
    Set itmX = lswEjecutivo.ListItems.Add(, , rs!id_promotor & "")
        itmX.SubItems(1) = rs!Nombre & "2"
    rs.MoveNext
Loop
rs.Close
 
End Sub

Private Sub LswMiembros_Click()
'Call sbHistorial

End Sub


Private Sub lswLiquidaciones_DblClick()
  GLOBALES.gTag = txtComite.Text
  Call sbFormsCall("frmAF_CD_Liquidaciones")

End Sub


Private Sub lswMiembrosComite_DblClick()
    
  fraMiembros.Visible = True
  fraBuscaMiembros.Visible = True
  tabGeneral.Tab = 0
  tabGeneral.Visible = True
'  tlbMiembros.Visible = True
  txtCedula.Text = Empty
  lblNombre.Caption = Empty
  chkDesembolso.Value = 0
  chkActivo.Value = 1
  txtNotas.Text = Empty
  lblTelefono.Caption = Empty
  lblCelular.Caption = Empty
  lblMail.Caption = Empty
  lblUT.Caption = Empty
  lswMiembros_H.ListItems.Clear
'  tlbMiembros.Visible = True
  tlbNuevo.Visible = True
  dtpFechaEleccion.Visible = True
  
  If lswMiembrosComite.ListItems.Count > 0 Then
  Call sbDatosMiembro(lswMiembrosComite.SelectedItem.Text)
  Call sbHistorialMiembros(lswMiembrosComite.SelectedItem.Text)
 End If

lswMiembrosComite.Visible = False


End Sub


Private Sub lswMiembrosComite_KeyDown(KeyCode As Integer, Shift As Integer)

Dim S As Integer

If KeyCode = vbKeyDelete Then
 If lswMiembrosComite.SelectedItem.Selected = True Then
  S = MsgBox("Desea borrar este miembro del comité", vbInformation + vbYesNo, "Información")
    
    If S = vbYes Then
       strSQL = "delete afi_cd_nombramientos where cedula = '" & lswMiembrosComite.SelectedItem.Text & "'"
      Call ConectionExecute(strSQL)
                 
                 MsgBox "Miembro eliminado", vbInformation, "Información"
                 lswMiembrosComite.ListItems.Clear
'    Call Bitacora("Borra", "Miembro:" & lswMiembrosComite.SelectedItem.Text & " Comite: " & txtComite.Text & "")
    Call sbMiembrosActivosComite(txtComite.Text)
    End If
 End If
End If

End Sub

Private Sub lswSelecBusqueda_DblClick()

If txtComite.Text = Empty Or txtDescripcionComite.Text = Empty Then
  MsgBox "No se puede procesar la información necesita ingresar incomité", vbInformation, "Información"
  txtComite.SetFocus
  Exit Sub
End If
If lswSelecBusqueda.ListItems.Count = 0 Then
  MsgBox "Listado Vacio", vbInformation, "Información"
  Exit Sub
End If


strSQL = "select U.cod_comite,C.descripcion" _
       & " from afi_cd_comites C left join afi_cd_comites_unidades U " _
       & "on U.cod_comite = C.cod_comite where U.codigo_up = '" & lswSelecBusqueda.SelectedItem.Text & "'"
Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF Then
        MsgBox "La Unidad que seleccionó ya pertenece al comité  " & rs!cod_comite & "  " & rs!Descripcion & " ", vbInformation, "Información"
        rs.Close
        Exit Sub
   End If
rs.Close
vValidacion = False


Select Case True
 Case SSTab1.Tab = 0
  Call sbCargaListView(lswSelecBusqueda, lswComites)
  If vValidacion = True Then Exit Sub
  Call sbIngresoComites
  SSTab1.TabEnabled(1) = True
 
 Case SSTab1.Tab = 1
  Call sbCargaListView(lswSelecBusqueda, lswActividades)
  If vValidacion = True Then Exit Sub
  Call sbIngresoActividades
  SSTab1.TabEnabled(2) = True
 
 Case SSTab1.Tab = 2
  Call sbCargaListView(lswSelecBusqueda, lswEjecutivo)
  If vValidacion = True Then Exit Sub
  Call sbIngresoEjecutivos
  SSTab1.TabEnabled(3) = True
End Select


End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
  txtBusqueda.Text = ""
  lswSelecBusqueda.ListItems.Clear
  Call sbCargaElementos
  
  If SSTab1.Tab = 0 Then
    txtCod_Comite_Ingreso.Visible = True
    txtDesc_Comite_Ingreso.Visible = True
    txtBusqueda.Visible = False
  Else
    txtCod_Comite_Ingreso.Visible = False
    txtDesc_Comite_Ingreso.Visible = False
    txtBusqueda.Visible = True
  End If
  fraBusquedaGeneral.Refresh
   
  If SSTab1.Tab = 3 Then
   tabGeneral.Tab = 0
   fraBusquedaGeneral.Visible = False
   fraMiembros.Visible = False
   fraMiembros.Left = 360
   fraMiembros.Top = 450
   
    Else
   fraBusquedaGeneral.Visible = True
   fraMiembros.Visible = True
   fraBuscaMiembros.Visible = False
   tabGeneral.Visible = False
  End If
  
  If SSTab1.Tab = 4 Then
    fraBusquedaGeneral.Visible = False
  End If
  
End Sub

Private Sub tabGeneral_Click(PreviousTab As Integer)
If tabGeneral.Tab = 1 Then
  fraBuscaMiembros.Visible = False
Else
  fraBuscaMiembros.Visible = True
End If
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      Call sbLimpia
    Case "GUARDAR", "SALVAR"
       Call sbModificarComite
       Call sbToolBar(tlb, "nuevo")
    Case "DESHACER"
       Call sbLimpia
       Call sbToolBar(tlb, "nuevo")
    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub tlbBuscar_Click()

 If txtConsultaComite.Text = Empty Then
    MsgBox "No se tienen datos"
    Exit Sub
 Else
   strSQL = "select cod_comite from afi_cd_comites_unidades where codigo_up = '" & txtConsultaComite.Text & "'"
            Call OpenRecordSet(rs, strSQL)
           If Not rs.EOF Then
              txtComite.Text = rs!cod_comite
           End If
   rs.Close
   Call txtComite_KeyDown(vbKeyReturn, 0)
 End If

End Sub



Private Sub tlbNuevo_ButtonClick(ByVal Button As MSComctlLib.Button)
  fraMiembros.Visible = True
  fraBuscaMiembros.Visible = True
  tabGeneral.Tab = 0
  tabGeneral.Visible = True
'  tlbMiembros.Visible = True
  txtCedula.Text = Empty
  lblNombre.Caption = Empty
  chkDesembolso.Value = 0
  chkActivo.Value = 1
  txtNotas.Text = Empty
  lblTelefono.Caption = Empty
  lblCelular.Caption = Empty
  lblMail.Caption = Empty
  lblUT.Caption = Empty
  lswMiembros_H.ListItems.Clear
'  tlbMiembros.Visible = True
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

Select Case KeyAscii
      'Case 48 To 57, 8
      Case 13
       lswSelecBusqueda.ListItems.Clear
       Select Case True
          Case SSTab1.Tab = 1
                strSQL = "select cod_actividad,descripcion from afi_cd_actividades " _
                         & "where descripcion like '" & "%" & txtBusqueda.Text & "%" & "'"
                                 Call OpenRecordSet(rs, strSQL)
                           While Not rs.EOF
                              Set itmX = lswSelecBusqueda.ListItems.Add(, , rs!Cod_Actividad)
                                  itmX.Checked = True
                                  itmX.SubItems(1) = Trim(rs!Descripcion)
                              rs.MoveNext
                           Wend
                       rs.Close
                
          Case SSTab1.Tab = 2
                strSQL = "select id_promotor,nombre from promotores where tipo = 'P' " _
                         & "and nombre like '" & "%" & txtBusqueda.Text & "%" & "'"
                         Call OpenRecordSet(rs, strSQL)
                     While Not rs.EOF
                        Set itmX = lswSelecBusqueda.ListItems.Add(, , rs!id_promotor)
                            itmX.Checked = True
                            itmX.SubItems(1) = Trim(rs!Nombre)
                        rs.MoveNext
                     Wend
                     
                 rs.Close
                 Case Else
                           KeyAscii = 0
                End Select
           Case SSTab1.Tab = 3
  End Select
End Sub


Private Sub TxtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
       txtCedula.Text = Empty
       lblNombre.Caption = Empty
       chkDesembolso.Value = 0
       chkActivo.Value = 1
       txtNotas.Text = Empty
       lblTelefono.Caption = Empty
       lblCelular.Caption = Empty
       lblMail.Caption = Empty
       lswMiembros_H.ListItems.Clear
       txtNombreJefe.Text = Empty
       txtTelefonoJefe.Text = Empty
       txtCelularJefe.Text = Empty
       txtCorreoJefe.Text = Empty
       cboRango.Text = "SD"
       
       gBusquedas.Columna = "nombre"
       gBusquedas.Orden = "cedula"
       gBusquedas.Consulta = "select Cedula, Nombre from socios"
       gBusquedas.Filtro = " and estadoactual = 'S'"
       frmBusquedas.Show vbModal
       txtCedula.SetFocus
       txtCedula.Text = gBusquedas.Resultado
       lblNombre.Caption = gBusquedas.Resultado2
       Call TxtCedula_KeyPress(vbKeyReturn)
End If

End Sub


Private Sub TxtCedula_KeyPress(KeyAscii As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

Select Case KeyAscii
  Case 48 To 57, 8
  Case 13
   If txtCedula.Text = "" Then
        MsgBox "No se defino la cedula del asociado", vbInformation + vbOKOnly, "Información"
     Else
        Call sbDatosMiembro(txtCedula.Text)

     End If
 Case Else
   KeyAscii = 0
End Select
End Sub


Private Sub txtCod_Comite_Ingreso_KeyPress(KeyAscii As Integer)
lswSelecBusqueda.ListItems.Clear
Select Case KeyAscii
      Case 48 To 57, 8
      Case 13
        If txtCod_Comite_Ingreso.Text = Empty Then Exit Sub
         strSQL = "select codigo,descripcion from uprogramatica where codigo = '" & txtCod_Comite_Ingreso.Text & "'"
                Call OpenRecordSet(rs, strSQL)
           If Not rs.EOF Then
                txtDesc_Comite_Ingreso.Text = rs!Descripcion
                Set itm = lswSelecBusqueda.ListItems.Add(, , txtCod_Comite_Ingreso.Text)
                itm.SubItems(1) = txtDesc_Comite_Ingreso.Text
           End If
              rs.Close
      Case Else
       KeyAscii = 0
    End Select
  
End Sub

Private Sub txtComite_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
       Call sbVerificaComite
       If vVerificaComite = True Then Exit Sub
       Call sbCargaComite
       Call sbToolBar(tlb, "edicion")
End Select
End Sub

Private Sub txtDesc_Comite_Ingreso_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    'Case 48 To 57, 8
     Case 13
        txtCod_Comite_Ingreso.Text = ""
        lswSelecBusqueda.ListItems.Clear
        strSQL = "select descripcion,codigo from uprogramatica " _
                         & "where descripcion like '" & "%" & txtDesc_Comite_Ingreso.Text & "%" & "'"
                           Call OpenRecordSet(rs, strSQL)
                           While Not rs.EOF
                              Set itmX = lswSelecBusqueda.ListItems.Add(, , rs!Codigo)
                                  itmX.SubItems(1) = Trim(rs!Descripcion)
                              rs.MoveNext
                           Wend
                 rs.Close
'      Case Else
'       KeyAscii = 0
 End Select
    
End Sub

Private Sub txtDesc_Comite_Ingreso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    'gBusquedas.Filtro = "and tipo = 'P'"
    gBusquedas.Consulta = "select descripcion,codigo from uprogramatica"
    frmBusquedas.Show vbModal
    txtDesc_Comite_Ingreso.Text = gBusquedas.Resultado
    txtCod_Comite_Ingreso.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtDescripcionComite_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    'gBusquedas.Filtro = "and tipo = 'P'"
    gBusquedas.Consulta = "select descripcion,cod_comite from afi_cd_comites"
    frmBusquedas.Show vbModal
    txtDescripcionComite.Text = gBusquedas.Resultado
    txtComite.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtDescripcionComite_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
