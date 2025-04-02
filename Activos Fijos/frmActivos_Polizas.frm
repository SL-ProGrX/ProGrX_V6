VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.ShortcutBar.v22.0.0.ocx"
Begin VB.Form frmActivos_Polizas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro y Asignación de Polizas"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   5652
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9732
      _ExtentX        =   17171
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pólizas"
      TabPicture(0)   =   "frmActivos_Polizas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Image1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line1(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblEstado"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCodigo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtMonto"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtNumPoliza"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtDocumento"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtDescripcion"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cbo"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtObservacion"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "dtpVence"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "dtpInicia"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Asignación"
      TabPicture(1)   =   "frmActivos_Polizas.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btnBuscar"
      Tab(1).Control(1)=   "chkLsw"
      Tab(1).Control(2)=   "lsw"
      Tab(1).Control(3)=   "cboTipo"
      Tab(1).Control(4)=   "txtActivoPlaca"
      Tab(1).Control(5)=   "txtActivoDesc"
      Tab(1).Control(6)=   "txtAsPoliza"
      Tab(1).Control(7)=   "txtAsPolDesc"
      Tab(1).Control(8)=   "Label1(11)"
      Tab(1).Control(9)=   "scTitulo"
      Tab(1).Control(10)=   "Label1(10)"
      Tab(1).Control(11)=   "Label1(9)"
      Tab(1).ControlCount=   12
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3612
         Left            =   -74880
         TabIndex        =   16
         Top             =   1920
         Width           =   9492
         _Version        =   1441792
         _ExtentX        =   16743
         _ExtentY        =   6371
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   310
         Left            =   -66480
         TabIndex        =   32
         Top             =   1200
         Width           =   372
         _Version        =   1441792
         _ExtentX        =   656
         _ExtentY        =   547
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkLsw 
         Height          =   200
         Left            =   -74760
         TabIndex        =   18
         Top             =   1680
         Width           =   200
         _Version        =   1441792
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicia 
         Height          =   312
         Left            =   5160
         TabIndex        =   14
         Top             =   2160
         Width           =   1332
         _Version        =   1441792
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.DateTimePicker dtpVence 
         Height          =   312
         Left            =   7800
         TabIndex        =   15
         Top             =   2160
         Width           =   1332
         _Version        =   1441792
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   -73200
         TabIndex        =   19
         Top             =   840
         Width           =   6612
         _Version        =   1441792
         _ExtentX        =   11668
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtActivoPlaca 
         Height          =   315
         Left            =   -73200
         TabIndex        =   21
         Top             =   1200
         Width           =   1572
         _Version        =   1441792
         _ExtentX        =   2773
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtActivoDesc 
         Height          =   312
         Left            =   -71640
         TabIndex        =   22
         Top             =   1200
         Width           =   5052
         _Version        =   1441792
         _ExtentX        =   8911
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtAsPoliza 
         Height          =   312
         Left            =   -73200
         TabIndex        =   23
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   480
         Width           =   1572
         _Version        =   1441792
         _ExtentX        =   2773
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtAsPolDesc 
         Height          =   312
         Left            =   -71640
         TabIndex        =   24
         Top             =   480
         Width           =   5052
         _Version        =   1441792
         _ExtentX        =   8911
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtObservacion 
         Height          =   1992
         Left            =   1800
         TabIndex        =   25
         Top             =   3480
         Width           =   7452
         _Version        =   1441792
         _ExtentX        =   13144
         _ExtentY        =   3514
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   1800
         TabIndex        =   26
         Top             =   1800
         Width           =   7332
         _Version        =   1441792
         _ExtentX        =   12938
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   1800
         TabIndex        =   27
         Top             =   1440
         Width           =   7332
         _Version        =   1441792
         _ExtentX        =   12933
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   312
         Left            =   1800
         TabIndex        =   28
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2640
         Width           =   2172
         _Version        =   1441792
         _ExtentX        =   3831
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNumPoliza 
         Height          =   315
         Left            =   1800
         TabIndex        =   30
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2160
         Width           =   2172
         _Version        =   1441792
         _ExtentX        =   3831
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   1800
         TabIndex        =   29
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3000
         Width           =   2172
         _Version        =   1441792
         _ExtentX        =   3831
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   432
         Left            =   1800
         TabIndex        =   31
         Top             =   600
         Width           =   2172
         _Version        =   1441792
         _ExtentX        =   3831
         _ExtentY        =   762
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Placa/Nombre"
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
         Height          =   252
         Index           =   11
         Left            =   -74760
         TabIndex        =   20
         Top             =   1200
         Width           =   1452
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   372
         Left            =   -74880
         TabIndex        =   17
         Top             =   1560
         Width           =   9492
         _Version        =   1441792
         _ExtentX        =   16743
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Seleccione los Activos que tiene cobertura con esta póliza"
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
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "  xxx"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   312
         Left            =   4080
         TabIndex        =   13
         Top             =   600
         Width           =   4932
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   276
         X2              =   9120
         Y1              =   1236
         Y2              =   1236
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmActivos_Polizas.frx":0038
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Activo"
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
         Height          =   252
         Index           =   10
         Left            =   -74760
         TabIndex        =   12
         Top             =   840
         Width           =   852
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Póliza"
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
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Monto"
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
         Left            =   240
         TabIndex        =   10
         Top             =   3000
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "Observacion"
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
         Left            =   240
         TabIndex        =   9
         Top             =   3480
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
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
         Index           =   6
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "Vence"
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
         Index           =   5
         Left            =   6960
         TabIndex        =   7
         Top             =   2160
         Width           =   612
      End
      Begin VB.Label Label1 
         Caption         =   "Inicia"
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
         Left            =   4320
         TabIndex        =   6
         Top             =   2160
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
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
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
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
         Index           =   2
         Left            =   1080
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Número"
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
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
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
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   1452
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
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
            Key             =   "reportes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmActivos_Polizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vPaso As Boolean

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNumPoliza.SetFocus
End Sub


Private Sub btnBuscar_Click()
Call sbPolizas_Asignacion
End Sub

Private Sub cboTipo_Click()
If Not vPaso Then Exit Sub
Call sbPolizas_Asignacion
End Sub

Private Sub chkLsw_Click()
Dim strSQL As String, x As Byte
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

For i = 1 To lsw.ListItems.Count
  If chkLsw.Value = vbChecked Then
     If Not lsw.ListItems.Item(i).Checked Then
        strSQL = "insert Activos_polizas_asigna(cod_poliza,num_placa) values(" _
               & txtAsPoliza & ",'" & lsw.ListItems.Item(i).Text & "')"
     End If
  Else
     If lsw.ListItems.Item(i).Checked Then
        strSQL = "delete Activos_polizas_asigna where cod_poliza = " _
               & txtAsPoliza & " and num_placa = '" & lsw.ListItems.Item(i).Text & "')"
     End If
  End If
  Call ConectionExecute(strSQL)
  lsw.ListItems.Item(i).Checked = chkLsw.Value
Next i

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub dtpInicia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpVence.SetFocus
End Sub

Private Sub dtpVence_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 36

 

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()
Dim strSQL As String

vCodigo = ""
txtCodigo = ""

ssTab.Tab = 0

txtDescripcion = ""
lblEstado.Caption = ""

strSQL = "select rtrim(tipo_poliza) as 'Idx', rtrim(descripcion) as 'ItmX'" _
       & " from Activos_polizas_tipos order by tipo_poliza"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

txtMonto = 0
txtNumPoliza = ""
txtDocumento = ""
txtObservacion = ""
dtpInicia.Value = fxFechaServidor
dtpVence.Value = dtpInicia

End Sub



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert Activos_polizas_asg(cod_poliza,num_placa,registro_fecha,registro_usuario) values('" & txtAsPoliza _
          & "','" & Item.Text & "',getdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete Activos_polizas_asg where cod_poliza = '" & txtAsPoliza _
          & "' and num_placa = '" & Item.Text & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String

If ssTab.Tab = 1 Then
  txtAsPoliza = ""
  txtAsPolDesc = ""
  lsw.ListItems.Clear
  vPaso = False
    strSQL = "select rtrim(tipo_activo) as 'IdX',  rtrim(descripcion) as 'ItmX'" _
       & " from Activos_tipo_activo order by tipo_activo"
    Call sbCbo_Llena_New(cboTipo, strSQL, False, True)
  vPaso = True
End If
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtDescripcion.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtDescripcion.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select cod_poliza,descripcion from Activos_polizas"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select P.*,rtrim(T.descripcion) as 'TipoDesc',getdate() as FechaX" _
       & " from Activos_polizas_tipos T inner join Activos_polizas P on T.tipo_poliza = P.tipo_poliza" _
       & " where P.cod_poliza = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!cod_poliza
  txtCodigo = rs!cod_poliza
 
  txtDescripcion = rs!Descripcion
  txtObservacion = rs!observacion
  dtpInicia.Value = rs!fecha_inicio
  dtpVence.Value = rs!fecha_vence
  
  Call sbCboAsignaDato(cbo, rs!TipoDesc, True, rs!Tipo_Poliza)
  
  txtMonto = Format(rs!monto, "Standard")
  txtNumPoliza = rs!num_poliza
  txtDocumento = rs!Documento
  
  If rs!fecha_vence < rs!fechaX Then
    lblEstado.Caption = "    Poliza Vencida"
    lblEstado.ForeColor = vbRed
  Else
    lblEstado.Caption = "    Poliza Activa"
    lblEstado.ForeColor = vbGrayText
  End If

  
  
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

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion de la poliza no es válido ..."
If dtpVence.Value < dtpInicia.Value Then vMensaje = vMensaje & vbCrLf & " - La fecha de vencimiento no puede ser menor a la inicial ..."
If Not IsNumeric(txtMonto) Then vMensaje = vMensaje & vbCrLf & " - Monto no es válido ..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  strSQL = "update Activos_polizas set descripcion = '" & txtDescripcion.Text _
         & "',observacion = '" & txtObservacion & "',monto = " & CCur(txtMonto) _
         & ",fecha_sistema = getdate(),fecha_inicio = '" & Format(dtpInicia.Value, "yyyy/mm/dd") _
         & "',fecha_vence = '" & Format(dtpVence.Value, "yyyy/mm/dd") & "',num_poliza = '" _
         & txtNumPoliza & "',documento = '" & txtDocumento & "',tipo_poliza = '" _
         & cbo.ItemData(cbo.ListIndex) _
         & "', modifica_fecha = getdate(), modifica_usuario = '" & glogon.Usuario _
         & "' where cod_poliza = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Póliza: " & vCodigo)

Else
   
   vCodigo = txtCodigo.Text
   
   strSQL = "insert Activos_polizas(cod_poliza,tipo_poliza,descripcion,observacion,fecha_sistema,fecha_inicio,fecha_vence" _
          & ",monto,num_poliza,documento,registro_fecha,registro_usuario) values('" & vCodigo & "','" _
          & cbo.ItemData(cbo.ListIndex) & "','" & txtDescripcion.Text _
          & "','" & txtObservacion & "',getdate(),'" & Format(dtpInicia.Value, "yyyy/mm/dd") _
          & "','" & Format(dtpVence.Value, "yyyy/mm/dd") & "'," & CCur(txtMonto) & ",'" _
          & txtNumPoliza & "','" & txtDocumento & "',getdate(),'" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Póliza: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbConsulta(txtCodigo.Text)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete Activos_polizas where cod_poliza = " & vCodigo
  Call ConectionExecute(strSQL)
  
'  Call sbBitacora("Elimina", "Tipo Activo : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbPolizas_Asignacion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError


If txtAsPoliza = "" Then Exit Sub

Me.MousePointer = vbHourglass


lsw.ListItems.Clear
lsw.Checkboxes = True
With lsw.ColumnHeaders
    .Clear
    .Add , , "Placa", 1800
    .Add , , "Nombre", 5700
    .Add , , "Estado", 1800, vbCenter
End With

strSQL = "select A.num_placa,A.nombre,A.estado,P.cod_poliza" _
       & " from Activos_Principal A left join Activos_polizas_asg P" _
       & " on A.num_placa = P.num_placa and P.cod_poliza = '" & txtAsPoliza & "'" _
       & " where A.tipo_activo = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
       
       
If Len(txtActivoPlaca.Text) > 0 Then
    strSQL = strSQL & " and A.Num_Placa like '%" & txtActivoPlaca.Text & "%'"
End If

If Len(txtActivoDesc.Text) > 0 Then
    strSQL = strSQL & " and A.nombre like '%" & txtActivoDesc.Text & "%'"
End If
       
strSQL = strSQL & " order by A.num_placa"

vPaso = True

    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
     Set itmX = lsw.ListItems.Add(, , rs!num_placa)
         itmX.SubItems(1) = rs!Nombre
         itmX.SubItems(2) = IIf((rs!Estado = "A"), "VIGENTE", "RETIRADO")
         itmX.Checked = IIf(IsNull(rs!cod_poliza), False, True)
     rs.MoveNext
    Loop
    rs.Close

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub txtActivoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  Call sbPolizas_Asignacion
End If
End Sub

Private Sub txtActivoPlaca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   Call sbPolizas_Asignacion
End If
End Sub

Private Sub txtAsPolDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "P.cod_poliza"
  gBusquedas.Orden = "P.cod_poliza"
  gBusquedas.Consulta = "select P.cod_poliza,P.descripcion,T.descripcion as Tipo" _
                      & " from Activos_polizas P inner join Activos_polizas_Tipos T on P.tipo_poliza = T.tipo_poliza"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtAsPoliza = gBusquedas.Resultado
  txtAsPolDesc = gBusquedas.Resultado2
  Call sbPolizas_Asignacion
End If
End Sub

Private Sub txtAsPoliza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "P.cod_poliza"
  gBusquedas.Orden = "P.cod_poliza"
  gBusquedas.Consulta = "select P.cod_poliza,P.descripcion,T.descripcion as Tipo" _
                      & " from Activos_polizas P inner join Activos_polizas_Tipos T on P.tipo_poliza = T.tipo_poliza"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtAsPoliza = gBusquedas.Resultado
  txtAsPolDesc = gBusquedas.Resultado2
  Call sbPolizas_Asignacion
End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" Then Call sbConsulta(txtCodigo)
  txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_poliza,descripcion from Activos_polizas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_poliza,descripcion from Activos_polizas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If
End Sub


Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError
 txtMonto = CCur(txtMonto)
vError:
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
 txtMonto = Format(CCur(txtMonto), "Standard")
vError:
End Sub


Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservacion.SetFocus
End Sub

Private Sub txtNumPoliza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInicia.SetFocus
End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End Sub
