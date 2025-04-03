VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_APA_CortesGarantias 
   AutoRedraw      =   -1  'True
   Caption         =   "Administración de Pagarés - Cortes Garantías"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14070
   Icon            =   "frmCR_APA_CortesGarantias.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   14070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMensajeDB 
      Height          =   1335
      Left            =   3000
      TabIndex        =   46
      Top             =   8040
      Width           =   5655
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Consultado la Base de Datos, este proceso puede tardar varios minutos, no debe cerrar el sistema"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   960
         TabIndex        =   47
         Top             =   360
         Width           =   3735
      End
      Begin VB.Image Image2 
         Height          =   660
         Left            =   4800
         Picture         =   "frmCR_APA_CortesGarantias.frx":6852
         Top             =   360
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   630
         Left            =   240
         Picture         =   "frmCR_APA_CortesGarantias.frx":6E2D
         Top             =   360
         Width           =   585
      End
   End
   Begin TabDlg.SSTab SSTabFiltros 
      Height          =   6375
      Left            =   120
      TabIndex        =   33
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Filtros"
      TabPicture(0)   =   "frmCR_APA_CortesGarantias.frx":72FC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lswGarantias"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraFiltros"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Otros"
      TabPicture(1)   =   "frmCR_APA_CortesGarantias.frx":7318
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label27"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label30"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label33"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtMoraFiltros"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtSaldoFiltros"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtRecursos"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtLinea"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtDestino"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin VB.TextBox txtDestino 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   75
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txtLinea 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   73
         Top             =   3840
         Width           =   2295
      End
      Begin VB.TextBox txtRecursos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   71
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Frame fraFiltros 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   -74880
         TabIndex        =   42
         Top             =   3360
         Width           =   2655
         Begin MSComctlLib.ImageCombo cboLineaCredito 
            Height          =   345
            Left            =   0
            TabIndex        =   79
            Top             =   2280
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   609
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFecFiltrosDesde 
            Height          =   330
            Left            =   0
            TabIndex        =   80
            Top             =   360
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
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
         Begin XtremeSuiteControls.DateTimePicker dtpFecFiltrosHasta 
            Height          =   330
            Left            =   1320
            TabIndex        =   81
            Top             =   360
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
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
         Begin XtremeSuiteControls.ComboBox cboEstadoFiltro 
            Height          =   330
            Left            =   0
            TabIndex        =   83
            Top             =   960
            Width           =   2535
            _Version        =   1441793
            _ExtentX        =   4471
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
         Begin XtremeSuiteControls.ComboBox cboCategoria 
            Height          =   330
            Left            =   0
            TabIndex        =   84
            Top             =   1560
            Width           =   2535
            _Version        =   1441793
            _ExtentX        =   4471
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
         Begin VB.Label Label34 
            Caption         =   "Línea Crédito"
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
            Left            =   0
            TabIndex        =   78
            Top             =   1920
            Width           =   1332
         End
         Begin VB.Label Label4 
            Caption         =   "Clasificación"
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
            Left            =   0
            TabIndex        =   45
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Estado"
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
            Left            =   0
            TabIndex        =   44
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label23 
            Caption         =   "Fecha Formalización"
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
            Left            =   0
            TabIndex        =   43
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.TextBox txtSaldoFiltros 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   37
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtMoraFiltros 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   36
         Top             =   960
         Width           =   2295
      End
      Begin MSComctlLib.ListView lswGarantias 
         Height          =   2265
         Left            =   -74880
         TabIndex        =   34
         Top             =   960
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   3995
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   0
      End
      Begin VB.Label Label33 
         Caption         =   "Línea"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label30 
         Caption         =   "Recursos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Saldo >="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Mora >="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Garantías"
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
         Left            =   -74400
         TabIndex        =   35
         Top             =   600
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   360
         Left            =   -74880
         Picture         =   "frmCR_APA_CortesGarantias.frx":7334
         Top             =   480
         Width           =   360
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6375
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Histórico"
      TabPicture(0)   =   "frmCR_APA_CortesGarantias.frx":77F1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tlbCortes"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "vGridCortes"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "+ Cortes"
      TabPicture(1)   =   "frmCR_APA_CortesGarantias.frx":780D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label11"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label12"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label13"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label14"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label15"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label16"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label17"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label18"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Line10"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Line11"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label19"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label20"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Line15"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "dtpFecha_Corte"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "tlbDatosCortes"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtNotaCorte"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtSaldo_Operacion"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtCierre_Fecha"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtRegistro_Fecha"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtSaldo_Responsabilidad"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtCierre_Usuario"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtRegistro_Usuario"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtEstadoCorte"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "Detalle"
      TabPicture(2)   =   "frmCR_APA_CortesGarantias.frx":7829
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Line19"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "tblDetalle"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "vGridDetalle"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "FraTotalesDetalle"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Inclusiones"
      TabPicture(3)   =   "frmCR_APA_CortesGarantias.frx":7845
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Line17"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "vGridInclusiones"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "tblInclusiones"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "FraTotalesInclusiones"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.Frame FraTotalesInclusiones 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   -74880
         TabIndex        =   59
         Top             =   5280
         Width           =   10692
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Responsabilidad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   3480
            TabIndex        =   68
            Top             =   0
            Width           =   1452
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo Op"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   120
            TabIndex        =   65
            Top             =   0
            Width           =   1284
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Garantias"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   1272
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Diferencia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   3480
            TabIndex        =   62
            Top             =   360
            Width           =   1452
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Marcado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   7320
            TabIndex        =   60
            Top             =   0
            Width           =   1152
         End
         Begin VB.Label lblTotal_GarantiasInc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   1320
            TabIndex        =   67
            Top             =   360
            Width           =   1968
         End
         Begin VB.Label lblTotal_OperacionInc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   1360
            TabIndex        =   66
            Top             =   0
            Width           =   1932
         End
         Begin VB.Label lblTotal_ResponsabilidadInc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   4920
            TabIndex        =   69
            Top             =   0
            Width           =   2076
         End
         Begin VB.Label lblTotal_DiferenciaInc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   4920
            TabIndex        =   63
            Top             =   360
            Width           =   2088
         End
         Begin VB.Label lblTotal_MarcadoInc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   8400
            TabIndex        =   61
            Top             =   0
            Width           =   2052
         End
      End
      Begin VB.Frame FraTotalesDetalle 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   -74880
         TabIndex        =   48
         Top             =   5280
         Width           =   10692
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Marcado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   7200
            TabIndex        =   58
            Top             =   0
            Width           =   1392
         End
         Begin VB.Label lblTotal_MarcadoDet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   8520
            TabIndex        =   57
            Top             =   0
            Width           =   2088
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Diferencia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   3240
            TabIndex        =   56
            Top             =   360
            Width           =   1812
         End
         Begin VB.Label lblTotal_DiferenciaDet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   5040
            TabIndex        =   55
            Top             =   360
            Width           =   1968
         End
         Begin VB.Label lblTotalCorte 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Garantias"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   1032
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo Op"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   120
            TabIndex        =   53
            Top             =   0
            Width           =   1044
         End
         Begin VB.Label lblTotal_OperacionDet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   1080
            TabIndex        =   52
            Top             =   0
            Width           =   1932
         End
         Begin VB.Label lblTotal_GarantiasDet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   1080
            TabIndex        =   51
            Top             =   360
            Width           =   1956
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Responsabilidad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   3240
            TabIndex        =   50
            Top             =   0
            Width           =   1812
         End
         Begin VB.Label lblTotal_ResponsabilidadDet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   5040
            TabIndex        =   49
            Top             =   0
            Width           =   1956
         End
      End
      Begin MSComctlLib.Toolbar tblInclusiones 
         Height          =   330
         Left            =   -74760
         TabIndex        =   30
         Top             =   600
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   556
         ButtonWidth     =   1693
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Incluir"
               Key             =   "Incluir"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Filtrar"
               Key             =   "Filtrar"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Marcar"
               Key             =   "Marcar"
               ImageIndex      =   7
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Todo"
                     Text            =   "Todo"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "MarcarHSaldo"
                     Text            =   "Hasta Responsabilidad"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "limpiar"
                     Text            =   "Limpiar"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtEstadoCorte 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox txtRegistro_Usuario 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69840
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtCierre_Usuario 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69840
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox txtSaldo_Responsabilidad 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69840
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox txtRegistro_Fecha 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtCierre_Fecha 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox txtSaldo_Operacion 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox txtNotaCorte 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -73560
         TabIndex        =   8
         Top             =   1560
         Width           =   5415
      End
      Begin FPSpreadADO.fpSpread vGridDetalle 
         Height          =   3852
         Left            =   -74880
         TabIndex        =   7
         Top             =   1200
         Width           =   10572
         _Version        =   524288
         _ExtentX        =   18648
         _ExtentY        =   6795
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   491
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_APA_CortesGarantias.frx":7861
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbDatosCortes 
         Height          =   330
         Left            =   -74760
         TabIndex        =   26
         Top             =   600
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   556
         ButtonWidth     =   1826
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Guardar"
               Key             =   "Guardar"
               Object.ToolTipText     =   "Agregar Acreedor"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar Acreedor"
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGridCortes 
         Height          =   4812
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   10452
         _Version        =   524288
         _ExtentX        =   18436
         _ExtentY        =   8488
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   493
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_APA_CortesGarantias.frx":82CA
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tblDetalle 
         Height          =   330
         Left            =   -74760
         TabIndex        =   31
         Top             =   600
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   556
         ButtonWidth     =   2090
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Excluir"
               Key             =   "Excluir"
               ImageIndex      =   9
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ExCategoria"
                     Text            =   "Exclusión Categoría"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ExAcreedor"
                     Text            =   "Exclusión Acreedor"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Filtrar"
               Key             =   "Filtrar"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Todos"
               Key             =   "Todos"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Actualizar"
               Key             =   "Actualizar"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGridInclusiones 
         Height          =   3852
         Left            =   -74880
         TabIndex        =   40
         Top             =   1200
         Width           =   10452
         _Version        =   524288
         _ExtentX        =   18436
         _ExtentY        =   6795
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   490
         MaxRows         =   501
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_APA_CortesGarantias.frx":8A96
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbCortes 
         Height          =   312
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   7548
         _ExtentX        =   13309
         _ExtentY        =   556
         ButtonWidth     =   1931
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Nuevo"
               Key             =   "Nuevo"
               Object.ToolTipText     =   "Agregar Acreedor"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar Acreedor"
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Editar"
               Key             =   "Editar"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ver"
               Key             =   "Ver"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalle"
               Key             =   "Detalle"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cerrar"
               Key             =   "Cerrar"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "Imprimir"
               ImageIndex      =   16
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "FormatoBP"
                     Text            =   "Formato BP"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "FormatoBCR"
                     Text            =   "Formato BCR"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "FormatoGeneral"
                     Text            =   "Formato General"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Inclusiones"
                     Text            =   "Inv Inclusiones"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin VB.Line Line1 
            X1              =   240
            X2              =   5400
            Y1              =   360
            Y2              =   240
         End
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha_Corte 
         Height          =   330
         Left            =   -73560
         TabIndex        =   82
         Top             =   1200
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
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
      Begin VB.Line Line19 
         BorderColor     =   &H80000004&
         X1              =   -74880
         X2              =   -68040
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line17 
         BorderColor     =   &H80000004&
         X1              =   -74880
         X2              =   -68040
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74760
         X2              =   -68160
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line14 
         BorderColor     =   &H80000004&
         X1              =   120
         X2              =   6960
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label20 
         Caption         =   "Saldos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   28
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Cierre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   27
         Top             =   4080
         Width           =   975
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74760
         X2              =   -68040
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74760
         X2              =   -68040
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   25
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Responsabilidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   -71280
         TabIndex        =   22
         Top             =   5160
         Width           =   1332
      End
      Begin VB.Label Label15 
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71280
         TabIndex        =   21
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71280
         TabIndex        =   20
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   18
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Registro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   3360
         Width           =   975
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74760
         X2              =   -68040
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label10 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha Corte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Frame fraPrincipal 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12972
      Begin MSComctlLib.ImageCombo cboOperaciones 
         Height          =   345
         Left            =   5400
         TabIndex        =   76
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   609
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin VB.TextBox txtCod_Acreedor 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNombreAcreedor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   3375
      End
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   312
         Left            =   8280
         TabIndex        =   3
         Top             =   480
         Width           =   1068
         _ExtentX        =   1879
         _ExtentY        =   556
         ButtonWidth     =   1640
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cargar"
               Key             =   "Cargar"
               Object.ToolTipText     =   "Agregar Acreedor"
               ImageIndex      =   1
            EndProperty
         EndProperty
         Begin VB.Line Line7 
            X1              =   120
            X2              =   5280
            Y1              =   360
            Y2              =   240
         End
      End
      Begin VB.Label txtSaldoOperacion 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   9600
         TabIndex        =   77
         Top             =   480
         Width           =   2652
      End
      Begin VB.Label Label25 
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   32
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Acreedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":94F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":FD56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":165B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":1CE1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":2367C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":29EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":30740
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":36FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":3D804
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":44066
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":4A8C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":5112A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":5798C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":5E1EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":64A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":67EA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_CortesGarantias.frx":6E709
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCR_APA_CortesGarantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mCod_Acreedor As String
Dim mOperacion As String
Dim mFechaCorte As String
Dim vEdita As Boolean
Dim vCambios As Boolean
Dim n As Integer
Dim strSQL As String
Dim i As Integer
Private rsLocal As New ADODB.Recordset
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private m_intLastHeight As Integer
Private m_intLastWidth As Integer

Private Sub sbLlenarListaGarantias()

On Error GoTo error

    Dim vItem As MSComctlLib.ListItem
    Dim vLvw As MSComctlLib.ListView
    Dim vKey As String
    Dim rs As New ADODB.Recordset
    
    
    Me.lswGarantias.ColumnHeaders.Clear
    Me.lswGarantias.ListItems.Clear
    
    Set vLvw = Me.lswGarantias
    vLvw.ColumnHeaders.Add , , "Garantía", 2100
    
    strSQL = "select GARANTIA, DESCRIPCION from CRD_GARANTIA_TIPOS " & _
             " order by DESCRIPCION "
    Call OpenRecordSet(rs, strSQL)

    While Not rs.EOF
        
        vKey = Trim(rs.Fields("GARANTIA")) & "(GA)"
        
        Set vItem = lswGarantias.ListItems.Add(, vKey, Trim(rs.Fields!Descripcion))
        
        rs.MoveNext
    Wend


    Exit Sub
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub sbCargarDatos()

    Call sbLlenarListaGarantias
    Call sbCargarCombo

End Sub


Private Sub sbCargarCombo()

    dtpFecFiltrosDesde.Value = fxFechaServidor()
    dtpFecFiltrosHasta.Value = dtpFecFiltrosDesde.Value
    dtpFecha_Corte.Value = dtpFecFiltrosDesde.Value
    
    
    cboEstadoFiltro.Clear
    cboEstadoFiltro.AddItem "Activas"
    cboEstadoFiltro.AddItem "Canceladas"
    cboEstadoFiltro.AddItem "Nulas"
    cboEstadoFiltro.AddItem "Activas y Canceladas"
    cboEstadoFiltro.AddItem "**Todos**"
    cboEstadoFiltro.Text = "**Todos**"
    
    Call sbCargaComboCategoria
    Call sbCargaComboLinea
    

End Sub

Sub sbCargaComboLinea()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError
  
  strSQL = "select CODIGO,DESCRIPCION " _
         & "from CATALOGO"
         
  Call OpenRecordSet(rs, strSQL)
  
  Do While Not rs.EOF
    cboLineaCredito.ComboItems.Add , rs.Fields("CODIGO") & "(id)", UCase(Trim(rs.Fields("DESCRIPCION")))
    rs.MoveNext
  Loop
  
  rs.Close
  
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Sub sbCargaComboCategoria()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError
    cboCategoria.Clear
    
    cboCategoria.AddItem "**Todas**"
    
    strSQL = "select COD_MORA from CBR_CLASIFICACION_MORA"
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboCategoria.AddItem IIf(IsNull(rs!COD_MORA), "", Trim(rs!COD_MORA))
      rs.MoveNext
    Loop
    rs.Close
    
    cboCategoria.Text = "**Todas**"
    
    Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
 
End Sub

Private Sub dtpCorte_Change()
    vEdita = True
End Sub

Private Sub sbCargaOperaciones()
    Dim strSQL As String
    Dim rs As New ADODB.Recordset

    strSQL = "select OPERACION from dbo.CRD_APA_OPERACIONES " _
           & "where COD_ACREEDOR='" & txtCod_Acreedor.Text & "' order by OPERACION"

    Call OpenRecordSet(rs, strSQL)
    cboOperaciones.ComboItems.Clear
    Do While Not rs.EOF
        cboOperaciones.ComboItems.Add , , rs.Fields(0)
        rs.MoveNext
    Loop

    rs.Close
End Sub




Private Sub cboLineaCredito_Click()
    cboLineaCredito.ToolTipText = cboLineaCredito.SelectedItem
End Sub

Private Sub dtpFecha_Corte_Change()
    vCambios = True
End Sub

Private Sub Form_Activate()
    vModulo = 14 'Modulo de Credito
End Sub

Private Sub sbNombreAcreedor()
    Dim strSQL As String, rs As New ADODB.Recordset
    
    If txtCod_Acreedor.Text <> Empty Then
        strSQL = "select DESCRIPCION from CRD_APA_ACREEDORES where COD_ACREEDOR = " & pc(Trim(txtCod_Acreedor))
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            txtNombreAcreedor = rs.Fields(0)
        Else
            txtNombreAcreedor = Empty
            MsgBox "No existe un acreedor con ese código"
            txtCod_Acreedor.SetFocus
        End If
    End If
    
End Sub

Private Sub Form_Load()

    If GLOBALES.gEnlace = 0 Then
        Call sbgCntParametros
    End If
    
    '' Carga nombre de la ternimal
    If Len(glogon.Maquina) = 0 Then
        Call sbMaquina
    End If

    ssTab.Tab = 0
    SSTabFiltros.Tab = 0
    Call ssTab_Click(0)
    fraMensajeDB.Visible = False
    vGridInclusiones.MaxRows = 0
    vGridInclusiones.MaxCols = 14
    
    m_intLastHeight = Me.Height
    m_intLastWidth = Me.Width
    
    Call sbCargarDatos
    Call sbCargarListaCortes
    
    
End Sub

Private Sub Form_Resize()
'' Procedimiento para posicionar los controles al max y minimizar la pantalla
On Error GoTo vError
    
    
        fraPrincipal.Width = Me.Width - 400
        
        ssTab.Width = Me.Width - SSTabFiltros.Width - 400
        ssTab.Height = Me.Height - 2000
        SSTabFiltros.Height = ssTab.Height
    
        vGridInclusiones.Width = ssTab.Width - 180
        vGridInclusiones.Height = ssTab.Height - 2400
    
        vGridDetalle.Width = vGridInclusiones.Width
        vGridDetalle.Height = vGridInclusiones.Height
    
        vGridCortes.Width = vGridInclusiones.Width
        vGridCortes.Height = vGridInclusiones.Height
    
        FraTotalesDetalle.Top = vGridDetalle.Top + vGridDetalle.Height + 200
        FraTotalesDetalle.Width = vGridDetalle.Width

        FraTotalesInclusiones.Top = FraTotalesDetalle.Top
        FraTotalesInclusiones.Width = FraTotalesDetalle.Width
        
        lswGarantias.Height = (SSTabFiltros.Height / 2) - 500
        fraFiltros.Top = lswGarantias.Top + lswGarantias.Height + 50
    

    
    Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

    
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
    Select Case ssTab.Tab
    Case 0
        ssTab.TabEnabled(0) = True
        ssTab.TabEnabled(1) = False
        ssTab.TabEnabled(2) = False
        ssTab.TabEnabled(3) = False
        Call sbCargarListaCortes
    Case 1
        ssTab.TabEnabled(0) = True
        ssTab.TabEnabled(1) = True
        ssTab.TabEnabled(2) = False
        ssTab.TabEnabled(3) = False
    Case 2, 3
        ssTab.TabEnabled(0) = True
        ssTab.TabEnabled(1) = False
        ssTab.TabEnabled(2) = True
        ssTab.TabEnabled(3) = True
    
    End Select
End Sub

Private Sub tblDetalle_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case UCase(Button.Key)
        Case "FILTRAR"
            Call sbCargarListaDetalle(True)
            
        Case "TODOS"
            Call sbCargarListaDetalle(False)
            
        Case "ACTUALIZAR"
            If fxCorteEstadoActivo = False Then
                MsgBox "Solo se pueden actualizar cortes activos"
                Exit Sub
            End If
            If MsgBox("Está seguro que sea actualizar el corte seleccionado ", vbExclamation + vbYesNo) = vbNo Then
                Exit Sub
            End If
            Call sbMostrarMensajeDB(True)
            DoEvents
            Call sbActualizarGarantiasCorte
            
            Call Bitacora("ACTUALIZAR", "APA Corte:" & mFechaCorte & " Operación:" & Trim(mOperacion) & " Acreedor:" & Trim(mCod_Acreedor))
            
            Call sbCargarListaDetalle(False)
            Call sbMostrarMensajeDB(False)
        Case "EXCLUIR"

    End Select
End Sub

Private Sub tblDetalle_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim OperacionExc As String
    Dim Tipo As Integer
    Select Case UCase(ButtonMenu.Key)
        Case "EXCATEGORIA"
            If MsgBox("Está seguro que sea excluir por categoría las garantías seleccionadas", vbExclamation + vbYesNo) = vbNo Then
                Exit Sub
            End If
            Tipo = 1
        Case "EXACREEDOR"
            Tipo = 2
            If MsgBox("Está seguro que sea excluir por solicitud del acreedor las garantías seleccionadas", vbExclamation + vbYesNo) = vbNo Then
                Exit Sub
            End If
    End Select
    
    If fxCorteEstadoActivo = False Then
        MsgBox "Solo se pueden excluir garantías en cortes activos"
        Exit Sub
    End If

    Call sbExcluirGarantiasMarcadas(Tipo)
    Call sbCargarListaDetalle(False)
    
End Sub

Private Sub sbExcluirGarantiasMarcadas(ByVal Tipo As Integer)
    Dim IdSolicitud As String

On Error GoTo vError
    Me.MousePointer = vbHourglass
    vGridDetalle.Row = 1
    vGridDetalle.Col = 1
    For i = 1 To vGridDetalle.MaxRows
        vGridDetalle.Row = i
        
        If vGridDetalle.Value = 1 Then
        
                vGridDetalle.Col = 2
                IdSolicitud = vGridDetalle.Value
                
                If IdSolicitud <> Empty Then
                
                    Call sbExcluirGarantia(Trim(IdSolicitud), Tipo)
    
                    Call Bitacora("EXCLUIR", "APA Garantía:" & Trim(IdSolicitud) & " Fecha Corte:" & mFechaCorte & " Operación:" & Trim(mOperacion) & " Acreedor:" & Trim(mCod_Acreedor))
                    
                End If
                
                vGridDetalle.Col = 1
        End If
    Next i
    Me.MousePointer = vbDefault
    MsgBox "Información guardada satisfactoriamente...", vbInformation
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub tblInclusiones_ButtonClick(ByVal Button As MSComctlLib.Button)
     
    Select Case UCase(Button.Key)
        Case "INCLUIR"
            If fxCorteEstadoActivo = False Then
                MsgBox "Solo se pueden incluir en cortes activos"
                Exit Sub
            End If
            Call sbIncluirGarantiasMarcadas
            Call sbCargarListaInclusiones
            Call sbCargarListaDetalle(False)
        Case "FILTRAR"
            If fxValidaFechasFiltro = False Then
                Exit Sub
            End If
            Call sbMostrarMensajeDB(True)
            DoEvents
            Call sbCargarListaInclusiones
            Call sbMostrarMensajeDB(False)
        Case "LIMPIAR"
            Call sbGridLimpiarMarcas(vGridInclusiones)
            
        Case "MARCAR"
            
            
    End Select
End Sub

Private Sub sbIncluirGarantiasMarcadas()
'' Procedimiento para Ingresar las Garantias marcadas en la lista de inclusiones

Dim IdSolicitud As String, Tasa As Double, Plazo As Integer, Cuota As Double
Dim Saldo As Double, Categoria As String, Mora_Intereses As Double
Dim Mora_Principal As Double, Mora_Cuotas As Integer

On Error GoTo vError

    vGridInclusiones.Row = 1
    vGridInclusiones.Col = 1
    For i = 1 To vGridInclusiones.MaxRows
        vGridInclusiones.Row = i
        
        If vGridInclusiones.Value = 1 Then
        
                vGridInclusiones.Col = 2
                IdSolicitud = vGridInclusiones.Value
                
                If IdSolicitud <> Empty Then
                
                    vGridInclusiones.Col = 10
                    Tasa = vGridInclusiones.Value
                    vGridInclusiones.Col = 11
                    Plazo = vGridInclusiones.Value
                    vGridInclusiones.Col = 4
                    Cuota = vGridInclusiones.Value
                    vGridInclusiones.Col = 5
                    Saldo = vGridInclusiones.Value
                    vGridInclusiones.Col = 6
                    Categoria = vGridInclusiones.Value
                    vGridInclusiones.Col = 12
                    Mora_Cuotas = vGridInclusiones.Value
                    vGridInclusiones.Col = 13
                    Mora_Intereses = vGridInclusiones.Value
                    vGridInclusiones.Col = 14
                    Mora_Principal = vGridInclusiones.Value
                    
                    Call sbAgregarOperacionCorte(IdSolicitud, Tasa, Plazo, _
                    Cuota, Saldo, Categoria, Mora_Intereses, Mora_Principal, _
                    Mora_Cuotas)
                    
                    Call Bitacora("REGISTRA", "APA Garantía:" & Trim(IdSolicitud) & "Corte:" & mFechaCorte & " Operación:" & Trim(mOperacion) & " Acreedor:" & Trim(mCod_Acreedor))
                    
                End If
                
                vGridInclusiones.Col = 1
        End If
    Next i
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValidaFechasFiltro() As Boolean
    'Valida Fechas de Filtros
    Dim FechaDesde As Date, FechaHasta As Date
    
    fxValidaFechasFiltro = True
    FechaDesde = dtpFecFiltrosDesde.Value
    FechaHasta = dtpFecFiltrosHasta.Value
    If FechaDesde > FechaHasta Then
        MsgBox "La fecha de incio de los filtros no puede ser mayor a la fecha fin"
        fxValidaFechasFiltro = False
    End If
    If DateDiff("d", FechaDesde, FechaHasta) > 5 Then
        If MsgBox("El rango de fechas es mayor a 5 días, este proceso va a durar varios minutos, ¿Desea continuar? ", vbExclamation + vbYesNo) = vbNo Then
            fxValidaFechasFiltro = False
        End If
    End If
    
End Function

Private Sub tblInclusiones_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case UCase(ButtonMenu.Key)
        Case "LIMPIAR"
            Call sbGridLimpiarMarcas(vGridInclusiones)
        Case "TODO"
            Call sbGridMarcarTodo(vGridInclusiones)
        Case "MARCARHSALDO"
          If CDbl(lblTotal_DiferenciaInc) < 0 Then
            Call sbGridMarcarHastaSaldo(vGridInclusiones, CDbl(Abs(lblTotal_DiferenciaInc.Caption)))
          End If
    End Select
End Sub

Private Sub tlbCortes_ButtonClick(ByVal Button As MSComctlLib.Button)
    If mOperacion = Empty Then
        MsgBox "Debe seleccionar la operación para agregar el corte"
        Exit Sub
    End If
    If mCod_Acreedor = Empty Then Exit Sub
    
    Select Case UCase(Button.Key)
        Case "NUEVO"
            
            If mCod_Acreedor = Empty Or mOperacion = Empty Then
                MsgBox "Debe seleccionar la operación"
                Exit Sub
            End If
        
            vEdita = False
            vCambios = False
            ssTab.Tab = 1
            Call sbLimpiarControlesCortes
            txtNotaCorte.SetFocus
            sbHabilitarDatosCortes ("NUEVO")
            
        Case "DETALLE"
        
            Call sbCargarCorteSeleccionado
            If mFechaCorte = Empty Then
                MsgBox "Debe seleccionar el corte que desea editar"
                Exit Sub
            End If
            vCambios = False
            ssTab.Tab = 2
            Call sbCargarListaDetalle(False)
            
        Case "VER"
            
            Call sbCargarCorteSeleccionado
            If mFechaCorte = Empty Then
                MsgBox "Debe seleccionar el corte que desea editar"
                Exit Sub
            End If
            Call sbConsultaCorte(mOperacion, mCod_Acreedor, mFechaCorte)
            ssTab.Tab = 1
            sbHabilitarDatosCortes ("Ver")
            
        Case "EDITAR"
        
            vCambios = False
            vEdita = True
            Call sbCargarCorteSeleccionado
            If mFechaCorte = Empty Then
                MsgBox "Debe seleccionar el corte que desea editar"
                Exit Sub
            End If
            If fxCorteEstadoActivo = False Then
                MsgBox "Solo se puede editar cortes en estado activo"
                Exit Sub
            End If
            Call sbConsultaCorte(mOperacion, mCod_Acreedor, mFechaCorte)
            ssTab.Tab = 1
            sbHabilitarDatosCortes ("Editar")
            
        Case "CERRAR"
        
            Call sbCargarCorteSeleccionado
            If mFechaCorte = Empty Then
                MsgBox "Debe seleccionar el corte que desea editar"
                Exit Sub
            End If
            If fxCorteEstadoActivo = False Then
                MsgBox "Solo se puede cerrar cortes en estado activo"
                Exit Sub
            End If
            '' Valida que diferencia no sea menor a cero
            If fxGridValorMarcado(vGridCortes, 6) < 0 Then
                MsgBox "No se puede cerrar cortes con faltante"
                Exit Sub
            End If
            If MsgBox("Está seguro que sea cerrar el corte seleccionado ", vbExclamation + vbYesNo) = vbNo Then
                Exit Sub
            End If
            Call sbCerrarCorte
            
            Call Bitacora("CERRAR", "APA Fecha Corte:" & mFechaCorte & " Operación:" & Trim(mOperacion) & " Acreedor:" & Trim(mCod_Acreedor))
            
            Call sbCargarListaCortes
            
        Case "IMPRIMIR"

            
    End Select
End Sub

Private Sub sbCargarCorteSeleccionado()
    mFechaCorte = Empty
    mFechaCorte = Format(fxGridValorMarcado(vGridCortes, 2), "yyyymmdd")
End Sub

Private Sub sbHabilitarDatosCortes(ByVal Modo As String)
    Select Case UCase(Modo)
    Case "VER"
        dtpFecha_Corte.Enabled = False
        txtNotaCorte.Locked = True
        tlbDatosCortes.Enabled = False
    Case "NUEVO"
        dtpFecha_Corte.Enabled = True
        txtNotaCorte.Locked = False
        tlbDatosCortes.Enabled = True
    Case "EDITAR"
        dtpFecha_Corte.Enabled = False
        txtNotaCorte.Locked = False
        tlbDatosCortes.Enabled = True
    End Select
End Sub


Private Sub tlbCortes_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Call sbCargarCorteSeleccionado
    If mFechaCorte = Empty Then
        MsgBox "Debe seleccionar el corte que desea imprimir"
        Exit Sub
    End If
    
    Select Case UCase(ButtonMenu.Key)
        Case "FORMATOBP"
            Call sbImprimir(0)
        Case "FORMATOBCR"
            Call sbImprimir(1)
        Case "FORMATOGENERAL"
            Call sbImprimir(2)
        Case "INCLUSIONES"
            Call sbImprimir(3)
    End Select
End Sub

Private Sub tlbDatosCortes_ButtonClick(ByVal Button As MSComctlLib.Button)
   
    Select Case UCase(Button.Key)
        Case "GUARDAR"
            If vCambios = True Then
                If fxValidaCortes Then
                    Call sbGuardarCorte
                Else
                    Exit Sub
                End If
            End If
            Call sbCargarListaCortes
            Call sbLimpiarControlesCortes
            ssTab.Tab = 0
    End Select
   
End Sub

Private Sub sbGuardarCorte()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    If vEdita = False Then
        Call sbAgregarCorte
        Call Bitacora("REGISTRA", "APA Fecha Corte:" & Format(dtpFecha_Corte.Value, "dd/mm/yyyy") & " Operación:" & Trim(mOperacion) & " Acreedor:" & Trim(mCod_Acreedor))
    Else
        Call sbEditarCorte
        Call Bitacora("MODIFICA", "APA Fecha Corte:" & Format(dtpFecha_Corte.Value, "dd/mm/yyyy") & " Operación:" & Trim(mOperacion) & " Acreedor:" & Trim(mCod_Acreedor))
    End If
    'Call sbToolBar(tlbPrincipal, "activo")
    Call RefrescaTags(Me)
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub sbAgregarCorte()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    Me.MousePointer = vbHourglass
    
    strSQL = "exec spCRDAPAGARANTIASCORTES_A " & pcc(mCod_Acreedor) _
                                            & pcc(mOperacion) _
                                            & pcc(Format(dtpFecha_Corte.Value, "yyyymmdd")) _
                                            & pcc(Format(fxFechaServidor, "yyyymmdd hh:mm:ss")) _
                                            & pcc(glogon.Usuario) _
                                            & pc(txtNotaCorte)
                                                
    Call OpenRecordSet(rs, strSQL)
    Me.MousePointer = vbDefault
    MsgBox "Información guardada satisfactoriamente...", vbInformation
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox "Ocurrió un error en visual basic al agregar la información ingresada. Error " & Err.Description
End Sub


Private Sub sbActualizarGarantiasCorte()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    Me.MousePointer = vbHourglass
    
    strSQL = "exec spCRDAPAGARANTIAS_H_Actualiza " & pcc(mCod_Acreedor) _
                                            & pcc(mOperacion) _
                                            & pc(mFechaCorte)
                                                
    Call OpenRecordSet(rs, strSQL)
    Me.MousePointer = vbDefault
    MsgBox "El corte fue actualizado satisfactoriamente...", vbInformation
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox "Ocurrió un error en visual basic al actualizar la información del corte. Error " & Err.Description
End Sub


Private Sub sbCerrarCorte()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

    Me.MousePointer = vbHourglass
    
    strSQL = "exec spCRDAPAGARANTIASCORTES_CERRAR " & pcc(mCod_Acreedor) & _
                                                pcc(mOperacion) & _
                                                pcc(mFechaCorte) & _
                                                pcc(Format(fxFechaServidor, "yyyymmdd hh:mm:ss")) & _
                                                pc(glogon.Usuario)
                                                
    Call OpenRecordSet(rs, strSQL)
    Me.MousePointer = vbDefault
    MsgBox "El corte fue cerrado satisfactoriamente...", vbInformation
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox "Ocurrió un error en visual basic al cerrar el corte. Error " & Err.Description
End Sub


Private Sub sbEditarCorte()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    Me.MousePointer = vbHourglass
    
    strSQL = "update CRD_APA_GARANTIAS_CORTES set NOTAS = " & pc(Trim(txtNotaCorte)) & _
                " where COD_ACREEDOR = " & pc(mCod_Acreedor) & _
                " and OPERACION = " & pc(mOperacion) & _
                " and FECHA_CORTE = " & pc(mFechaCorte)
                                                
    Call OpenRecordSet(rs, strSQL)
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox "Ocurrió un error en visual basic al editar la información del corte. Error " & Err.Description
End Sub

Private Sub sbExcluirGarantia(ByVal Id_Solicitud As String, ByVal Tipo As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

    strSQL = "exec spCRDAPAGARANTIAS_H_Excluir " & pcc(mCod_Acreedor) _
                                            & pcc(mOperacion) _
                                            & pcc(mFechaCorte) _
                                            & Trim(Id_Solicitud) & "," _
                                            & Tipo
                                                
    Call OpenRecordSet(rs, strSQL)
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox "Ocurrió un error en visual basic al excluir la garantía seleccionada. Error " & Err.Description
End Sub

Private Sub sbAgregarOperacionCorte(ByVal IdSolicitud As String, _
                                    ByVal Tasa As Double, _
                                    ByVal Plazo As Integer, _
                                    ByVal Cuota As Double, _
                                    ByVal Saldo As Double, _
                                    ByVal Categoria As String, _
                                    ByVal Mora_Intereses As Double, _
                                    ByVal Mora_Pricipal As Double, _
                                    ByVal Mora_Cuotas As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

    Me.MousePointer = vbHourglass
    
    strSQL = "exec spCRDAPAGARANTIAS_H_A " & pcc(mCod_Acreedor) _
                                                & pcc(cboOperaciones.Text) _
                                                & pcc(mFechaCorte) _
                                                & pcc(IdSolicitud) _
                                                & Tasa & "," _
                                                & Plazo & "," _
                                                & Cuota & "," _
                                                & Saldo & "," _
                                                & pc(Categoria) & "," _
                                                & Mora_Intereses & "," _
                                                & Mora_Pricipal & "," _
                                                & Mora_Cuotas
                                                
    Call OpenRecordSet(rs, strSQL)
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox "Ocurrió un error en visual basic al agregar la operacion " & mCod_Acreedor & " Error " & Err.Description
End Sub


Private Function fxValidaCortes() As Boolean
    Dim strSQL As String, rs As New ADODB.Recordset
    Dim vMensaje As String
    
    vMensaje = ""
    fxValidaCortes = True
    
        If vEdita = False Then

            'Verifica que exista ningún corte en esa fecha
            strSQL = "select isnull(count(*),0) as Existe from CRD_APA_GARANTIAS_CORTES" _
                   & " where COD_ACREEDOR = '" & Trim(mCod_Acreedor) & "'" _
                   & " and OPERACION = '" & Trim(mOperacion) & "'" _
                   & " and FECHA_CORTE = '" & Format(dtpFecha_Corte.Value, "yyyymmdd") & "'"
            Call OpenRecordSet(rs, strSQL)
            If rs!Existe > 0 Then
               vMensaje = vMensaje & vbCrLf & " Ya Existe un corte para la fecha seleccionada "
            End If
            rs.Close
            
            'Verifica que exista ningún corte Abierto
            strSQL = "select isnull(count(*),0) as Existe from CRD_APA_GARANTIAS_CORTES" _
                   & " where COD_ACREEDOR = '" & Trim(mCod_Acreedor) & "'" _
                   & " and OPERACION = '" & Trim(mOperacion) & "'" _
                   & " and ESTADO = 'A'"
            Call OpenRecordSet(rs, strSQL)
            If rs!Existe > 0 Then
               vMensaje = vMensaje & vbCrLf & "Existe un corte abierto en esta operación"
            End If
            rs.Close
            
        End If
    
    If Len(vMensaje) > 0 Then
      fxValidaCortes = False
      MsgBox vMensaje, vbCritical
    End If

End Function

Private Function fxCorteEstadoActivo() As Boolean
    'Consulta si el estado del corte es activo
    Dim strSQL As String, rs As New ADODB.Recordset
    Dim vMensaje As String
    
    fxCorteEstadoActivo = False

    strSQL = "select ESTADO from CRD_APA_GARANTIAS_CORTES" _
                & " where COD_ACREEDOR = " & pc(Trim(mCod_Acreedor)) _
                & " and OPERACION = " & pc(Trim(mOperacion)) _
                & " and FECHA_CORTE = " & pc(Trim(mFechaCorte))
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        If rs.Fields(0) = "A" Then
            fxCorteEstadoActivo = True
        Else
            fxCorteEstadoActivo = False
        End If
    Else
        fxCorteEstadoActivo = False
    End If
    rs.Close

End Function


Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
        Case "CARGAR"
            Call sbCargarDatosOperacion
    End Select
    
End Sub

Private Sub txtCod_Acreedor_Change()
 vCambios = True
End Sub

Private Sub txtCod_Acreedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCod_Acreedor.SetFocus
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "Cod_Acreedor"
        gBusquedas.Orden = "Cod_Acreedor"
        gBusquedas.Filtro = ""
        gBusquedas.Consulta = "select cod_acreedor,descripcion from crd_apa_acreedores"
        frmBusquedas.Show vbModal
        txtCod_Acreedor = gBusquedas.Resultado
        txtNombreAcreedor = gBusquedas.Resultado2
        cboOperaciones.SetFocus
        
        Call sbCargaOperaciones
    
    End If
End Sub

Private Sub txtCod_Acreedor_LostFocus()
   Call sbNombreAcreedor
   Call sbCargaOperaciones
End Sub

Private Sub txtDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyF4 Then
       txtDestino.Text = Empty
       txtDestino.Locked = True
       Exit Sub
    Else
        gBusquedas.Columna = "cod_destino"
        gBusquedas.Orden = "cod_destino"
        gBusquedas.Filtro = ""
        gBusquedas.Consulta = "select cod_destino,descripcion" _
                            & " from catalogo_destinos"
        frmBusquedas.Show vbModal
        txtDestino.Tag = gBusquedas.Resultado
        txtDestino = gBusquedas.Resultado2
    
    End If
End Sub


Private Sub txtLinea_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyF4 Then
       txtDestino.Text = Empty
       txtDestino.Locked = True
       Exit Sub
    Else
        gBusquedas.Columna = "CODIGO"
        gBusquedas.Orden = "CODIGO"
        gBusquedas.Filtro = ""
        gBusquedas.Consulta = "select CODIGO,DESCRIPCION" _
                            & " from  CATALOGO"
        frmBusquedas.Show vbModal
        txtLinea.Tag = gBusquedas.Resultado
        txtLinea = gBusquedas.Resultado2
    
    End If
End Sub

Private Sub txtNombreAcreedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombreAcreedor.SetFocus
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "Cod_Acreedor"
        gBusquedas.Orden = "Cod_Acreedor"
        gBusquedas.Filtro = ""
        gBusquedas.Consulta = "select cod_acreedor,descripcion from crd_apa_acreedores"
        frmBusquedas.Show vbModal
        txtCod_Acreedor = gBusquedas.Resultado
        txtNombreAcreedor = gBusquedas.Resultado2
        cboOperaciones.SetFocus
        
        Call sbCargaOperaciones
    
    End If
End Sub

Private Sub txtNotaCorte_Change()
    vCambios = True
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboOperaciones.SetFocus
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "COD_ACREEDOR"
        gBusquedas.Orden = "COD_ACREEDOR"
        gBusquedas.Filtro = ""
        gBusquedas.Consulta = "select OPERACION" _
                            & " from CRD_APA_OPERACIONES "
                            
        frmBusquedas.Show vbModal
        cboOperaciones.Text = gBusquedas.Resultado
        
        Call sbCargarDatosOperacion

    
    End If
End Sub

Private Sub sbCargarDatosOperacion()

    If Len(cboOperaciones.Text) > 0 Then
        
        Dim strSQL As String, rs As New ADODB.Recordset

       strSQL = "select OP.SALDO, AC.DESCRIPCION, AC.COD_ACREEDOR " _
                & "from CRD_APA_OPERACIONES OP " _
                & " inner join CRD_APA_ACREEDORES AC on AC.COD_ACREEDOR = OP.COD_ACREEDOR" _
                & " where OP.OPERACION = '" & Trim(cboOperaciones.Text) & "'"
                            
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            
            txtSaldoOperacion = Format(rs!Saldo, "Standard")
            txtNombreAcreedor = Trim(rs!Descripcion)
            mCod_Acreedor = Trim(rs!Cod_Acreedor)
            mOperacion = Trim(cboOperaciones.Text)
        Else
            sbLimpiarListas
        End If
        
        Call sbCargarListaCortes
    Else
        Call sbLimpiarListas
    End If
End Sub

Private Sub sbLimpiarListas()
    txtNombreAcreedor.Text = Empty
    txtSaldoOperacion = Empty
    mCod_Acreedor = Empty
    mOperacion = Empty
    vGridInclusiones.MaxRows = 0

End Sub


Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Call sbCargarDatosOperacion
    End If
    
End Sub

Private Sub sbCargarListaCortes()
    Dim strSQL As String
    
    'Consulta la lista de las Cortes
    strSQL = "select FECHA_CORTE, SALDO_OPERACION, " & _
        " dbo.fxCRDAPASaldoCorteGarantias(COD_ACREEDOR,OPERACION,FECHA_CORTE), " & _
        " dbo.fxCRDAPASaldoCorteResponsabilidad(COD_ACREEDOR,OPERACION,FECHA_CORTE), " & _
        " dbo.fxCRDAPASaldoCorteDiferencia(COD_ACREEDOR,OPERACION,FECHA_CORTE), " & _
        " case ESTADO when 'A' then 'Abierto' when 'C' then 'Cerrado' else '' end as ESTADO , REGISTRO_FECHA" & _
        " from CRD_APA_GARANTIAS_CORTES where COD_ACREEDOR = '" & Trim(mCod_Acreedor) & "'" & _
        " and OPERACION =  '" & Trim(cboOperaciones.Text) & "'" & _
        " order by FECHA_CORTE desc"
        
        
    Call sbCargaGridCheckIni(vGridCortes, 7, strSQL)
    vGridCortes.MaxRows = vGridCortes.MaxRows - 1
End Sub

Private Sub sbCargarListaInclusiones()
    Dim strSQL As String, garantiasSeleccionadas As String
    Me.MousePointer = vbHourglass
        
    'Consulta la lista de las Cortes
    strSQL = "select ID_SOLICITUD, MONTOAPR, CUOTA, SALDO, CLASIFICACION" & _
        ",CODIGO, FECHAFORP, GARANTIA, INT, PLAZO, MORA_CUOTAS, MORA_INTERESES, MORA_PRINCIPAL" & _
        ",COD_DESTINO,COD_GRUPO from dbo.vAPACreditosInclusiones where FECHAFORP between '" & Format(dtpFecFiltrosDesde.Value, "yyyymmdd") & "'" & _
        " and '" & Format(dtpFecFiltrosHasta.Value, "yyyymmdd") & "'"
        
    If cboCategoria.Text <> "**Todas**" Then
        strSQL = strSQL & " and CLASIFICACION = '" & Trim(cboCategoria.Text) & "'"
    End If
    
    If cboLineaCredito.SelText <> "" Then
        strSQL = strSQL & " and CODIGO = '" & Trim(DeCodificaPrimaryKey(cboLineaCredito.SelectedItem.Key, 1, "(id)")) & "'"
    End If
    
    garantiasSeleccionadas = fxGarantiasSeleccionadas
    If garantiasSeleccionadas <> Empty Then
        strSQL = strSQL & " and GARANTIA in (" & garantiasSeleccionadas & ")"
    End If
    
    If txtSaldoFiltros.Text <> Empty Then
       If IsNumeric(txtSaldoFiltros) = True Then
            strSQL = strSQL & " and SALDO >= " & Trim(txtSaldoFiltros)
       End If
    End If
    
    If txtMoraFiltros.Text <> Empty Then
       If IsNumeric(txtMoraFiltros) = True Then
            strSQL = strSQL & " and MORA_INTERESES+MORA_PRINCIPAL >= " & Trim(txtMoraFiltros)
       End If
    End If
               
    If txtDestino.Text <> Empty Then
       strSQL = strSQL & " and COD_DESTINO = " & Trim(txtDestino.Tag)
       
    End If
    
    If txtRecursos.Text <> Empty Then
       strSQL = strSQL & " and COD_GRUPO = " & Trim(txtRecursos.Tag)
    End If
    
    If txtLinea.Text <> Empty Then
       strSQL = strSQL & " and CODIGO >= " & Trim(txtLinea.Tag)
    End If
               
    Call sbCargaGridCheckIni(vGridInclusiones, 13, strSQL)
    vGridInclusiones.MaxRows = vGridInclusiones.MaxRows - 1
    
    Me.MousePointer = vbDefault
End Sub

Private Sub sbCargarListaDetalle(ByVal Filtrar As Boolean)
    Dim strSQL As String, garantiasSeleccionadas As String
    Me.MousePointer = vbHourglass
    
        
    'Consulta la lista de detalle de Cortes
    strSQL = "select  G.ID_SOLICITUD_1,R.MONTOAPR,G.CUOTA,G.SALDO,G.CATEGORIA,R.CODIGO,R.FECHAFORP," & _
        "R.GARANTIA,case when G.ESTADO = 'I' then 'Inclusión'when G.ESTADO = 'A' then 'Activa' " & _
        " when G.ESTADO = 'E' then 'Exclusión' when G.ESTADO = 'ECV' then 'Ex Canc y Ven' when G.ESTADO = 'EC' then 'Ex Categoria' " & _
        " when G.ESTADO = 'EA' then 'Ex Acreedor' end as ESTADO,G.TASA,G.PLAZO,G.MORA_CUOTAS,G.MORA_INTERESES,G.MORA_PRINCIPAL " & _
        " from CRD_APA_GARANTIAS_H G inner join REG_CREDITOS R on R.ID_SOLICITUD = G.ID_SOLICITUD_1" & _
        " where G.COD_ACREEDOR = " & pc(mCod_Acreedor) & _
        " and G.OPERACION = " & pc(mOperacion) & _
        " and G.FECHA_CORTE = " & pc(mFechaCorte)
        
    If Filtrar = True Then
        If cboCategoria.Text <> "**Todas**" Then
            strSQL = strSQL & " and G.CATEGORIA = '" & Trim(cboCategoria.Text) & "'"
        End If
    
        garantiasSeleccionadas = fxGarantiasSeleccionadas
        If garantiasSeleccionadas <> Empty Then
            strSQL = strSQL & " and R.GARANTIA in (" & garantiasSeleccionadas & ")"
        End If
        
        strSQL = strSQL & " and R.FECHAFORP between '" & Format(dtpFecFiltrosDesde.Value, "yyyymmdd") & "'" & _
                    " and '" & Format(dtpFecFiltrosHasta.Value, "yyyymmdd") & "'"
    
    End If
               
    Call sbCargaGridCheckIni(vGridDetalle, 14, strSQL)
    vGridDetalle.MaxRows = vGridDetalle.MaxRows - 1
    
    'Consulta el total del corte
    Call sbTotalesCorte
    
    Me.MousePointer = vbDefault
End Sub

Private Sub sbTotalesCorte()
    Dim strSQL As String, rs As New ADODB.Recordset
    Dim Saldo_Corte As Double
    Me.MousePointer = vbHourglass
    
    'Consulta Saldo del la operación al corte
    strSQL = "select  isnull(SALDO_OPERACION,0) as Saldo, " & _
        " dbo.fxCRDAPASaldoCorteGarantias(COD_ACREEDOR,OPERACION,FECHA_CORTE) as SaldoCorteGarantias , " & _
        " dbo.fxCRDAPASaldoCorteResponsabilidad(COD_ACREEDOR,OPERACION,FECHA_CORTE) as SaldoCorteResponsabilidad, " & _
        " dbo.fxCRDAPASaldoCorteDiferencia(COD_ACREEDOR,OPERACION,FECHA_CORTE) as SaldoCorteDiferencia " & _
        " from CRD_APA_GARANTIAS_CORTES " & _
        " where COD_ACREEDOR = " & pc(mCod_Acreedor) & _
        " and OPERACION = " & pc(mOperacion) & _
        " and FECHA_CORTE = " & pc(mFechaCorte)
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        lblTotal_OperacionDet.Caption = Format(IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0)), "Standard")
        lblTotal_OperacionInc.Caption = Format(IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0)), "Standard")
        
        lblTotal_GarantiasDet.Caption = Format(IIf(IsNull(rs.Fields(1)), 0, rs.Fields(1)), "Standard")
        lblTotal_GarantiasInc.Caption = Format(IIf(IsNull(rs.Fields(1)), 0, rs.Fields(1)), "Standard")
        
        lblTotal_ResponsabilidadDet.Caption = Format(IIf(IsNull(rs.Fields(2)), 0, rs.Fields(2)), "Standard")
        lblTotal_ResponsabilidadInc.Caption = Format(IIf(IsNull(rs.Fields(2)), 0, rs.Fields(2)), "Standard")
        
        lblTotal_DiferenciaDet.Caption = Format(IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3)), "Standard")
        lblTotal_DiferenciaInc.Caption = Format(IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3)), "Standard")
        
    End If
    
    lblTotal_MarcadoDet = Format(fxGridSumaMarcado(vGridDetalle, 5), "Standard")
    lblTotal_MarcadoInc = Format(fxGridSumaMarcado(vGridDetalle, 5), "Standard")
    lblTotal_GarantiasInc = Format(fxGridSumaMarcado(vGridDetalle, 5), "Standard")
    
    Me.MousePointer = vbDefault
End Sub

Private Sub sbLimpiarControlesCortes()
    dtpFecha_Corte.Value = fxFechaServidor()
    txtNotaCorte.Text = Empty
    txtEstadoCorte.Text = Empty
    txtRegistro_Fecha.Text = Empty
    txtRegistro_Usuario.Text = Empty
    txtCierre_Fecha.Text = Empty
    txtCierre_Usuario.Text = Empty
    txtSaldo_Operacion.Text = Empty
    txtSaldo_Responsabilidad.Text = Empty
End Sub

Private Sub txtOperacion_LostFocus()
    Call sbCargarDatosOperacion
End Sub

Private Sub txtRecursos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyF4 Then
       txtDestino.Text = Empty
       txtDestino.Locked = True
       Exit Sub
    Else
        gBusquedas.Columna = "COD_GRUPO"
        gBusquedas.Orden = "COD_GRUPO"
        gBusquedas.Filtro = ""
        gBusquedas.Consulta = "select COD_GRUPO,DESCRIPCION" _
                            & " from CATALOGO_GRUPOS"
        frmBusquedas.Show vbModal
        txtRecursos.Tag = gBusquedas.Resultado
        txtRecursos = gBusquedas.Resultado2
    
    End If
End Sub

Private Sub vGridCortes_Click(ByVal Col As Long, ByVal Row As Long)
    Call sbGridMarcarSoloUno(vGridCortes, Row)
End Sub

Public Function fxGarantiasSeleccionadas() As String
'' Función busca las garantías seleccionadas en la lista de garantias
    On Error GoTo vError
    Dim i As Long
        fxGarantiasSeleccionadas = Empty
        For i = lswGarantias.ListItems.Count To 1 Step -1
            If lswGarantias.ListItems(i).Checked Then
                If fxGarantiasSeleccionadas = Empty Then
                    fxGarantiasSeleccionadas = "'" & Trim(DeCodificaPrimaryKey(lswGarantias.ListItems(i).Key, 1, "(GA)")) & "'"
                Else
                    fxGarantiasSeleccionadas = fxGarantiasSeleccionadas & ",'" & Trim(DeCodificaPrimaryKey(lswGarantias.ListItems(i).Key, 1, "(GA)")) & "'"
                End If
            End If
        Next i
        Exit Function
vError:
        MsgBox fxSys_Error_Handler(Err.Description)

End Function

Private Sub sbMostrarMensajeDB(ByVal Mostrar As Boolean)
    If Mostrar = True Then
        fraMensajeDB.Top = Me.Height / 2 - fraMensajeDB.Height / 2
        fraMensajeDB.Left = Me.Width / 2 - fraMensajeDB.Width / 2
        fraMensajeDB.Visible = True
    Else
        fraMensajeDB.Visible = False
    End If
End Sub

Private Sub vGridDetalle_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
   lblTotal_MarcadoDet = Format(fxGridSumaMarcado(vGridDetalle, 5), "Standard")
End Sub

Private Sub vGridInclusiones_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    lblTotal_MarcadoInc = Format(fxGridSumaMarcado(vGridInclusiones, 5), "Standard")
    lblTotal_GarantiasInc = Format(fxGridSumaMarcado(vGridDetalle, 5), "Standard")
End Sub

Private Sub sbConsultaCorte(Operacion As String, Acreedor As String, Fecha_Corte As String)
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer


On Error GoTo vError

    Me.MousePointer = vbHourglass
    
    strSQL = "SELECT COD_ACREEDOR,OPERACION,FECHA_CORTE as FECHA_CORTE,REGISTRO_FECHA,REGISTRO_USUARIO,CIERRE_FECHA,CIERRE_USUARIO," _
              & "ESTADO,NOTAS,SALDO_OPERACION,SALDO_RESPONSABILIDAD from CRD_APA_GARANTIAS_CORTES " _
              & " where COD_ACREEDOR = '" & Trim(Acreedor) & "'" _
              & " and OPERACION = '" & Trim(Operacion) & "'" _
              & " and FECHA_CORTE = '" & Trim(Fecha_Corte) & "'"
              
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.BOF And Not rs.EOF Then
        dtpFecha_Corte = IIf(IsNull(rs!Fecha_Corte), Empty, rs!Fecha_Corte)
        txtNotaCorte.Text = IIf(IsNull(rs!Notas), Empty, rs!Notas)
        Select Case IIf(IsNull(rs!Estado), Empty, rs!Estado)
            Case "A"
                txtEstadoCorte.Text = "Abierto"
            Case "C"
                txtEstadoCorte.Text = "Cerrado"
        End Select
        txtRegistro_Fecha = Format(IIf(IsNull(rs!REGISTRO_FECHA), Empty, rs!REGISTRO_FECHA), "dd/mm/yyyy hh:mm")
        txtRegistro_Usuario = IIf(IsNull(rs!REGISTRO_USUARIO), Empty, rs!REGISTRO_USUARIO)
        txtCierre_Fecha = Format(IIf(IsNull(rs!CIERRE_FECHA), Empty, rs!CIERRE_FECHA), "dd/mm/yyyy hh:mm")
        txtCierre_Usuario = IIf(IsNull(rs!CIERRE_USUARIO), Empty, rs!CIERRE_USUARIO)
        txtSaldo_Operacion = Format(IIf(IsNull(rs!SALDO_OPERACION), Empty, rs!SALDO_OPERACION), "Standard")
        txtSaldo_Responsabilidad = Format(IIf(IsNull(rs!SALDO_RESPONSABILIDAD), Empty, rs!SALDO_RESPONSABILIDAD), "Standard")
        
    Else
      
      MsgBox "No se encontró registro verifique...", vbInformation
    
    End If
    
    rs.Close
    Me.MousePointer = vbDefault
    Call RefrescaTags(Me)

    Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbImprimir(ByVal ReporteNumero As Integer)

    Dim vNombreCia As String, vCategoria As String
    Dim strSQL As String
    Dim Reporte As String
    Dim vTitulo As String
    
    Me.MousePointer = vbHourglass
    
    strSQL = "{vAPAGarantias.COD_ACREEDOR} = " & pc(Trim(mCod_Acreedor)) _
                & " and {vAPAGarantias.OPERACION} = " & pc(Trim(mOperacion)) _
                & " and CDate({vAPAGarantias.FECHA_CORTE}) = Date(" & Format(Format(fxGridValorMarcado(vGridCortes, 2), "dd-mm-yyyy"), "yyyy,mm,dd") & ")"
                
    vNombreCia = GLOBALES.gstrNombreEmpresa
    
    vTitulo = "Operación...: " & mOperacion
    
    Select Case ReporteNumero
    Case 0 ' Reporte Corte Formato Banco Popular
        
'        vTitulo = "Detalle Pagarés-Garantía del Contrato de Mandato Irrevocable No. 27-08 entre el Banco Popular y de Desarrollo Comunal y ASECCSS"
        vCategoria = "Categoría AA"
        Reporte = SIFGlobal.fxPathReportes("Acreedores_BancoFormatoBP.rpt")
    
    Case 1 ' Reporte Corte Formato Banco de Costa Rica
        
        vTitulo = ""
        vCategoria = ""
        
        Reporte = SIFGlobal.fxPathReportes("Acreedores_BancoFormatoBCR.rpt")
    
    Case 2 ' Reporte Corte Formato General
        
        vTitulo = ""
        vCategoria = ""
        Reporte = SIFGlobal.fxPathReportes("Acreedores_BancoFormatoGeneral.rpt")
   
    Case 3 ' Reporte Inventario Microfilmado
                
        strSQL = strSQL & " and {vAPAGarantias.ESTADO} = 'I' "
        vTitulo = ""
        vCategoria = ""
        Reporte = SIFGlobal.fxPathReportes("Acreedores_Consulta_Inv_Operaciones.rpt")
    
    End Select
    
    With frmContenedor.Crt
    
      .Reset
      .WindowShowPrintSetupBtn = True
      .WindowShowRefreshBtn = True
      .WindowShowSearchBtn = True
      .WindowState = crptMaximized
      .WindowTitle = "Reportes Administración Pagarés"
      .Connect = glogon.ConectRPT
      .ReportFileName = Reporte
      .Formulas(0) = "fxNombreCia=" & pc(vNombreCia)
      .Formulas(1) = "fxTitulo1=" & pc(vTitulo)
      .Formulas(2) = "fxCategoria='" & pc(vCategoria)
      .SelectionFormula = strSQL
      .PrintReport
      
    End With
    
    Select Case ReporteNumero
    Case 0 ' Reporte Corte Formato Banco Popular
        
        vTitulo = ""
        vCategoria = ""
        Reporte = SIFGlobal.fxPathReportes("Acreedores_BancoTotalesBP.rpt")
    
    Case 1 ' Reporte Corte Formato Banco de Costa Rica
        
        vTitulo = ""
        vCategoria = ""
        Reporte = SIFGlobal.fxPathReportes("Acreedores_BancoTotalesBCR.rpt")
    
    Case 2 ' Reporte Corte Formato General
        
        vTitulo = ""
        vCategoria = ""
        Reporte = SIFGlobal.fxPathReportes("Acreedores_BancoTotalesGeneral.rpt")
    
    End Select
    
    If ReporteNumero <> 3 Then
    
        With frmContenedor.Crt
            .Reset
            .Connect = Empty
            .WindowShowGroupTree = True
            .WindowShowPrintSetupBtn = True
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            .WindowTitle = "Reportes Administración Pagarés"
            '.Destination = crptToPrinter
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Connect = glogon.ConectRPT
            .ReportFileName = Reporte
            .Formulas(0) = "fxNombreCia=" & pc(vNombreCia)
            .StoredProcParam(0) = Trim(mCod_Acreedor)
            .StoredProcParam(1) = Trim(mOperacion)
            .StoredProcParam(2) = Format(fxGridValorMarcado(vGridCortes, 2), "yyyy-mm-dd 00:00:00.000")
            .PrintReport
        End With
    End If
    
    Me.MousePointer = vbDefault

End Sub


