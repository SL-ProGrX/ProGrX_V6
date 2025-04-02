VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmFNDConsultaContratos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Consulta de Contratos"
   ClientHeight    =   7725
   ClientLeft      =   2355
   ClientTop       =   2430
   ClientWidth     =   11115
   Icon            =   "frmFNDConsultaContratos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   11115
   Begin XtremeSuiteControls.TabControl ssTab 
      Height          =   6732
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   10932
      _Version        =   1310723
      _ExtentX        =   19283
      _ExtentY        =   11874
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
      ItemCount       =   3
      Item(0).Caption =   "Contratos"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "GroupBox2"
      Item(0).Control(1)=   "GroupBox1"
      Item(0).Control(2)=   "lsw"
      Item(0).Control(3)=   "cmdHistorico"
      Item(0).Control(4)=   "cmdEstadoCuenta"
      Item(0).Control(5)=   "lblContrato"
      Item(1).Caption =   "Liquidaciones"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "fraEntrega"
      Item(1).Control(1)=   "cmdBoleta"
      Item(1).Control(2)=   "lblBoleta"
      Item(1).Control(3)=   "lswRet"
      Item(1).Control(4)=   "btnReversion"
      Item(1).Control(5)=   "scConsulta"
      Item(2).Caption =   "Movimientos"
      Item(2).ControlCount=   11
      Item(2).Control(0)=   "chkTodas"
      Item(2).Control(1)=   "txtPlan"
      Item(2).Control(2)=   "txtContrato"
      Item(2).Control(3)=   "vGrid"
      Item(2).Control(4)=   "dtpDesde"
      Item(2).Control(5)=   "dtpHasta"
      Item(2).Control(6)=   "btnBuscar"
      Item(2).Control(7)=   "Label8"
      Item(2).Control(8)=   "Label7"
      Item(2).Control(9)=   "Label6"
      Item(2).Control(10)=   "Label5"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3132
         Left            =   1920
         TabIndex        =   19
         Top             =   480
         Width           =   8892
         _Version        =   1310723
         _ExtentX        =   15684
         _ExtentY        =   5524
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
         Appearance      =   16
         Sorted          =   -1  'True
      End
      Begin XtremeSuiteControls.ListView lswRet 
         Height          =   4812
         Left            =   -69880
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   10692
         _Version        =   1310723
         _ExtentX        =   18860
         _ExtentY        =   8488
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
         Appearance      =   16
         Sorted          =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkTodas 
         Height          =   252
         Left            =   -64600
         TabIndex        =   30
         Top             =   720
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1310723
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "&Todas"
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.GroupBox fraEntrega 
         Height          =   2292
         Left            =   -67720
         TabIndex        =   22
         Top             =   1560
         Visible         =   0   'False
         Width           =   6492
         _Version        =   1310723
         _ExtentX        =   11451
         _ExtentY        =   4043
         _StockProps     =   79
         Caption         =   "Marcar como Entregado el Retiro?"
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
         Appearance      =   16
         BorderStyle     =   1
         Begin VB.TextBox txtEntregaMonto 
            Alignment       =   1  'Right Justify
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
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox txtEntregaBoleta 
            Alignment       =   1  'Right Justify
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
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   480
            Width           =   1935
         End
         Begin MSComctlLib.Toolbar tlbProceso 
            Height          =   312
            Left            =   3720
            TabIndex        =   27
            Top             =   1440
            Width           =   2616
            _ExtentX        =   4604
            _ExtentY        =   556
            ButtonWidth     =   1931
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Aplicar"
                  Key             =   "aplicar"
                  Object.ToolTipText     =   "Aplicar Archivo"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Cancelar"
                  Key             =   "cancelar"
                  Object.ToolTipText     =   "cancelar operacion"
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageListX 
            Left            =   240
            Top             =   240
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
                  Picture         =   "frmFNDConsultaContratos.frx":030A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmFNDConsultaContratos.frx":6B6C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label Label4 
            Caption         =   "Monto del Retiro"
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
            Index           =   1
            Left            =   1440
            TabIndex        =   26
            Top             =   840
            Width           =   2292
         End
         Begin VB.Label Label4 
            Caption         =   "Id. Boleta de Retiro"
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
            Index           =   0
            Left            =   1440
            TabIndex        =   25
            Top             =   480
            Width           =   2292
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5292
         Left            =   -69760
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   10572
         _Version        =   524288
         _ExtentX        =   18648
         _ExtentY        =   9335
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
         SpreadDesigner  =   "frmFNDConsultaContratos.frx":D3CE
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   495
         Left            =   -61000
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1310723
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Buscar"
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
         Picture         =   "frmFNDConsultaContratos.frx":DCDD
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   3732
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   1812
         _Version        =   1310723
         _ExtentX        =   3196
         _ExtentY        =   6583
         _StockProps     =   79
         ForeColor       =   8421504
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton opt 
            Height          =   492
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1572
            _Version        =   1310723
            _ExtentX        =   2773
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Activo"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   17
            Checked         =   -1  'True
            Picture         =   "frmFNDConsultaContratos.frx":E3DD
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton opt 
            Height          =   492
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   1572
            _Version        =   1310723
            _ExtentX        =   2773
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Liquidado"
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
            FlatStyle       =   -1  'True
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmFNDConsultaContratos.frx":EB4F
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton opt 
            Height          =   492
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   1560
            Width           =   1572
            _Version        =   1310723
            _ExtentX        =   2773
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Inactivo"
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
            FlatStyle       =   -1  'True
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmFNDConsultaContratos.frx":F2C1
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton opt 
            Height          =   492
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   2160
            Width           =   1572
            _Version        =   1310723
            _ExtentX        =   2773
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Bloqueado"
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
            FlatStyle       =   -1  'True
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmFNDConsultaContratos.frx":FA32
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton opt 
            Height          =   492
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   2760
            Width           =   1572
            _Version        =   1310723
            _ExtentX        =   2773
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Cerrados"
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
            FlatStyle       =   -1  'True
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmFNDConsultaContratos.frx":101A4
            ImageAlignment  =   0
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1932
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   10692
         _Version        =   1310723
         _ExtentX        =   18860
         _ExtentY        =   3408
         _StockProps     =   79
         Caption         =   "Sub Cuentas Relacionadas..: "
         ForeColor       =   8421504
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lswSubCuentas 
            Height          =   1572
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   10572
            _Version        =   1310723
            _ExtentX        =   18648
            _ExtentY        =   2773
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
            Appearance      =   16
            Sorted          =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton cmdHistorico 
         Height          =   495
         Left            =   6600
         TabIndex        =   16
         Top             =   3720
         Width           =   1935
         _Version        =   1310723
         _ExtentX        =   3413
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Histórico"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmFNDConsultaContratos.frx":10916
      End
      Begin XtremeSuiteControls.PushButton cmdEstadoCuenta 
         Height          =   495
         Left            =   8520
         TabIndex        =   17
         Top             =   3720
         Width           =   2295
         _Version        =   1310723
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Estado de Cuenta"
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
         Picture         =   "frmFNDConsultaContratos.frx":11016
      End
      Begin XtremeSuiteControls.FlatEdit txtPlan 
         Height          =   312
         Left            =   -69640
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1310723
         _ExtentX        =   1926
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtContrato 
         Height          =   312
         Left            =   -68560
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1310723
         _ExtentX        =   1926
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   312
         Left            =   -67480
         TabIndex        =   31
         Top             =   720
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   550
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHasta 
         Height          =   312
         Left            =   -66160
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   550
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton btnReversion 
         Height          =   612
         Left            =   -61840
         TabIndex        =   33
         Top             =   5760
         Visible         =   0   'False
         Width           =   2652
         _Version        =   1310723
         _ExtentX        =   4678
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reversa Retiro/Liquidación"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmFNDConsultaContratos.frx":1171D
      End
      Begin XtremeSuiteControls.PushButton cmdBoleta 
         Height          =   612
         Left            =   -63640
         TabIndex        =   34
         Top             =   5760
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1310723
         _ExtentX        =   3196
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Boleta"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmFNDConsultaContratos.frx":120AA
      End
      Begin XtremeShortcutBar.ShortcutCaption scConsulta 
         Height          =   375
         Left            =   -69880
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   10695
         _Version        =   1310723
         _ExtentX        =   18865
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Listado de Retiros/Liquidaciones Registradas (Doble Click / Entrega)"
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
      Begin XtremeSuiteControls.Label lblContrato 
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   3840
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "[Contrato]"
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
         Alignment       =   2
      End
      Begin VB.Label lblBoleta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "[Boleta]"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   -66640
         TabIndex        =   8
         Top             =   5880
         Visible         =   0   'False
         Width           =   2772
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Contrato"
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
         Left            =   -68560
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Plan"
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
         Left            =   -69640
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Inicio"
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
         Height          =   300
         Left            =   -67480
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Corte"
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
         Height          =   300
         Left            =   -66160
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDConsultaContratos.frx":12866
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4200
      TabIndex        =   35
      Top             =   240
      Width           =   6852
      _Version        =   1310723
      _ExtentX        =   12086
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2040
      TabIndex        =   36
      Top             =   240
      Width           =   2172
      _Version        =   1310723
      _ExtentX        =   3831
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11412
   End
End
Attribute VB_Name = "frmFNDConsultaContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vRA_Access As Boolean


Private Sub btnBuscar_Click()
Call sbConsultaMovimientos
End Sub

Private Sub btnReversion_Click()
Dim strSQL As String

Me.MousePointer = vbHourglass

On Error GoTo vError

If lblBoleta.Tag = "" Then Exit Sub

strSQL = "exec spFndReversaLiq " & lblBoleta.Tag & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

If glogon.error Then
  Me.MousePointer = vbDefault
  Exit Sub
End If

Call Bitacora("Aplica", "Reversión de la Liquidación No.: " & lblBoleta.Tag)
Call sbConsultaLiquidaciones

Me.MousePointer = vbDefault
MsgBox "Reversión realizada satisfactoriamente!", vbInformation


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdBoleta_Click()

Me.MousePointer = vbHourglass

On Error GoTo vError

If lblBoleta.Tag = "" Then Exit Sub

With frmContenedor.Crt
  .Reset
  .WindowShowGroupTree = True
  .WindowShowPrintSetupBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowState = crptMaximized
  .WindowTitle = "Reportes del Módulo de Fondos"
    
  .Connect = glogon.ConectRPT
  
  .ReportFileName = SIFGlobal.fxPathReportes("Fondos_LiquidacionBoleta.rpt")
  
  .Formulas(0) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
  .Formulas(1) = "fxCodigoBarras= '*" & lblBoleta.Tag & "*'"
  
  .SelectionFormula = "{FND_LIQUIDACION.CONSEC} =" & lblBoleta.Tag

  .SubreportToChange = "sbAsiento"
  
  .StoredProcParam(0) = "FLIQ"
  .StoredProcParam(1) = lblBoleta.Tag
  .StoredProcParam(2) = 1


  .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdEstadoCuenta_Click()
If Trim(txtCedula) = "" Then Exit Sub

With frmContenedor.Crt
  .Reset
  .WindowShowPrintSetupBtn = True
  .WindowShowExportBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowShowZoomCtl = True
  .WindowTitle = "Fondos de Ahorros e Inversiones"
  .WindowState = crptMaximized
  
  .Connect = glogon.ConectRPT
  
  .ReportFileName = SIFGlobal.fxPathReportes("Fondos_EstadoConsolidado.rpt")
  .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
  .Formulas(1) = "Usuario='" & Trim(glogon.Usuario) & "'"
  .Formulas(2) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
  .Formulas(3) = "SubTitulo=' Reporte al " & Format(fxFechaServidor, "yyyy/mm/dd") & "'"
  .SelectionFormula = "{FND_CONTRATOS.CEDULA} ='" & Trim(txtCedula) & "'"
  .PrintReport
End With

End Sub

Private Sub cmdHistorico_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

With frmContenedor.Crt
  .Reset
  .WindowShowPrintSetupBtn = True
  .WindowShowExportBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowShowZoomCtl = True
  .WindowTitle = "Fondos de Ahorros e Inversiones"
  .WindowState = crptMaximized
  
  .Connect = glogon.ConectRPT
  
  .ReportFileName = SIFGlobal.fxPathReportes("Fondos_EstadoDetalladoH.rpt")
  .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
  .Formulas(1) = "Usuario='" & Trim(glogon.Usuario) & "'"
  .Formulas(2) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
  .Formulas(3) = "SubTitulo='" & Format(fxFechaServidor, "yyyy/mm/dd") & "'"
  
  .SelectionFormula = "{FND_CONTRATOS.COD_CONTRATO} = " _
                    & lblContrato.Caption & " and {FND_CONTRATOS.COD_PLAN} = '" & lblContrato.Tag _
                    & "' and {FND_CONTRATOS.COD_OPERADORA} = " & lsw.SelectedItem
  
  .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Public Sub sbConsultaExterna(pCedula As String)

txtCedula.Text = pCedula
Call txtCedula_LostFocus

End Sub

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
vModulo = 18 'Fondo de Inversion


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

ssTab.Item(0).Selected = True
ssTab.Item(2).Enabled = False

dtpHasta.Value = Format(fxFechaServidor, "dd/mm/yyyy")
dtpDesde.Value = Format(DateAdd("m", -2, dtpHasta), "dd/mm/yyyy")
ssTab.Item(2).Enabled = False

With lswRet.ColumnHeaders
    .Clear
    .Add , , "Plan", 1000, vbCenter
    .Add , , "Descripción", 2500
    .Add , , "Contrato", 1200, vbCenter
    .Add , , "No.Boleta", 1200, vbCenter
    .Add , , "Fecha", 1800
    .Add , , "Usuario", 1800
    .Add , , "Monto", 1500, vbRightJustify
    .Add , , "Estado", 1500, vbCenter
    .Add , , "Tesoreria Id", 2100, vbCenter
    .Add , , "Tes. Fecha", 2100
    .Add , , "Tes. Usuario", 2100
    .Add , , "Entrega/Fecha", 2100
    .Add , , "Entrega/Usuario", 2100
    
End With



With lsw.ColumnHeaders
  .Clear
  .Add , , "Operadora Id", 0
  .Add , , "Operadora", 1600
  .Add , , "Plan", 1200, vbCenter
  .Add , , "Descripción", 2200
  .Add , , "No.Contrato", 1400, vbCenter
  .Add , , "Estado", 1200, vbCenter
  .Add , , "Fecha", 1200, vbCenter
  .Add , , "Mensualidad", 1200, vbRightJustify
  .Add , , "Plazo", 900, vbCenter
  .Add , , "Renueva?", 900, vbCenter
  .Add , , "Inc.Anual?", 1000, vbCenter
  .Add , , "Inc.Tipo?", 1000, vbCenter
  .Add , , "Aportes", 1600, vbRightJustify
  .Add , , "Rendimiento", 1600, vbRightJustify
  .Add , , "Total", 1600, vbRightJustify
  .Add , , "En Tránsito", 1600, vbRightJustify
  .Add , , "Op.Reten.", 1200, vbCenter
End With


With lswSubCuentas.ColumnHeaders

  .Clear
  .Add , , "Plan Id", 1200, vbCenter
  .Add , , "Contrato", 1200, vbCenter
  .Add , , "Sub Id", 1200, vbCenter
  .Add , , "Identificación", 1400
  .Add , , "Nombre", 3200
  .Add , , "Mensualidad", 1200, vbRightJustify
  .Add , , "Estado?", 1100, vbCenter
  .Add , , "Aportes", 1600, vbRightJustify
  .Add , , "Rendimiento", 1600, vbRightJustify
  .Add , , "Total", 1600, vbRightJustify

End With

Me.Height = 8172
Me.Width = 11520


Call Formularios(Me)
Call RefrescaTags(Me)



End Sub

Private Sub sbConsultaLiquidaciones()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

txtNombre.SetFocus

ssTab.Item(1).Selected = True

fraEntrega.Visible = False

lswRet.ListItems.Clear

'Carga Boleta
strSQL = "select C.cod_plan,P.descripcion,C.cod_contrato,L.consec,L.fecha,L.usuario,L.aportes_liq+L.rendi_liq as 'Monto'" _
       & ",L.traspaso_tesoreria,L.Traspaso_usuario,L.Solicitud_Tesoreria,isnull(L.Estado,'P') as 'Estado'" _
       & " from fnd_contratos C inner join fnd_liquidacion L on C.cod_operadora = L.cod_operadora" _
       & " and C.cod_plan = L.cod_plan and C.cod_Contrato = L.cod_contrato" _
       & " inner join fnd_planes P on C.cod_plan = P.cod_plan and P.cod_operadora = C.cod_operadora" _
       & " Where C.cedula = '" & txtCedula.Text & "' order by L.consec desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswRet.ListItems.Add(, , rs!cod_Plan)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!COD_CONTRATO
     itmX.SubItems(3) = rs!consec
     itmX.SubItems(4) = Format(rs!fecha, "dd/mm/yyyy")
     itmX.SubItems(5) = rs!Usuario
     itmX.SubItems(6) = Format(rs!Monto, "Standard")
     If rs!Estado = "P" Then
        itmX.SubItems(7) = "Procesada"
     Else
        itmX.SubItems(7) = "Reversada"
     End If
     itmX.SubItems(8) = rs!Solicitud_Tesoreria & ""
     itmX.SubItems(9) = Format(rs!traspaso_tesoreria & "", "dd/mm/yyyy")
     itmX.SubItems(10) = rs!Traspaso_usuario & ""

 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub sbConsultaContrato()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotal(2) As Currency

If txtCedula = "" Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass


ssTab.Item(0).Selected = True

lsw.ListItems.Clear
lswRet.ListItems.Clear

lswSubCuentas.ListItems.Clear
lblContrato.Tag = ""
lblContrato.Caption = ">> Contrato <<"

lblBoleta.Tag = ""
lblBoleta.Caption = ">> Boleta <<"

curTotal(0) = 0
curTotal(1) = 0
curTotal(2) = 0

txtNombre.Text = fxNombre(txtCedula)

strSQL = "Select S.Nombre,O.Descripcion,P.Descripcion as DPlan,F.Cod_Operadora" _
       & ",F.Cod_plan,F.cod_Contrato,F.Estado,F.Liq_Fecha" _
       & ",F.Fecha_Inicio,F.Monto,F.Plazo,F.Renueva,F.Inc_Anual,F.Inc_Tipo,F.Aportes" _
       & ",F.Rendimiento,F.Operacion,F.Monto_Transito" _
       & " From Socios S" _
       & " inner join Fnd_Contratos F on S.Cedula = F.Cedula" _
       & " inner join Fnd_operadoras O on F.cod_operadora = O.cod_operadora" _
       & " inner join Fnd_planes P on F.Cod_plan = P.Cod_plan" _
       & " Where S.cedula='" & Trim(txtCedula) _
       & "'  AND dbo.fxFndColaboradorVisualiza(F.COD_OPERADORA, F.COD_PLAN, F.cedula,S.EstadoActual, '" & glogon.Usuario & "') = 1"

Select Case True
  Case opt.Item(0).Checked  'Activos
      strSQL = strSQL & " and F.estado = 'A' order by F.Fecha_Inicio desc,F.cod_Plan,F.cod_Contrato"
  Case opt.Item(1).Checked 'Liquidados
      strSQL = strSQL & " and F.estado = 'L' order by F.Liq_Fecha desc"
  Case opt.Item(2).Checked 'inactivos
      strSQL = strSQL & " and F.estado = 'I' order by F.Fecha_Inicio desc,F.cod_Plan,F.cod_Contrato"
  Case opt.Item(3).Checked 'Bloqueados
      strSQL = strSQL & " and F.estado = 'B' order by F.Fecha_Inicio desc,F.cod_Plan,F.cod_Contrato"
  Case opt.Item(4).Checked 'Cerrados
      strSQL = strSQL & " and F.estado = 'C' order by F.Fecha_Inicio desc,F.cod_Plan,F.cod_Contrato"

End Select

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!cod_Operadora)
       itmX.SubItems(1) = Trim(rs!Descripcion)
       itmX.SubItems(2) = Trim(rs!cod_Plan)
       itmX.SubItems(3) = Trim(rs!DPlan)
       itmX.SubItems(4) = rs!COD_CONTRATO
       itmX.SubItems(5) = fxFndEstadoContrato(rs!Estado)
       If rs!Estado = "A" Then
           itmX.SubItems(6) = Format(rs!Fecha_Inicio, "dd/mm/yyyy")
       Else
           itmX.SubItems(6) = Format(rs!Liq_Fecha, "dd/mm/yyyy")
       End If

       itmX.SubItems(7) = Format(rs!Monto, "Standard")
       itmX.SubItems(8) = rs!Plazo
       itmX.SubItems(9) = IIf(rs!Renueva = "S", "SI", "NO")
       itmX.SubItems(10) = Format(rs!Inc_Anual, "Standard")
       itmX.SubItems(11) = IIf(rs!inc_tipo = "P", "Porcentaje", "Monto")
       itmX.SubItems(12) = Format(rs!aportes, "Standard")
       itmX.SubItems(13) = Format(rs!rendimiento, "Standard")
       itmX.SubItems(14) = Format(rs!aportes + rs!rendimiento, "Standard")
       itmX.SubItems(15) = Format(rs!Monto_Transito, "Standard")
       
       itmX.SubItems(16) = IIf(IsNull(rs!Operacion), "", rs!Operacion)
    
    curTotal(0) = curTotal(0) + rs!Monto
    curTotal(1) = curTotal(1) + rs!aportes
    curTotal(2) = curTotal(2) + rs!rendimiento
   
   rs.MoveNext
Loop
rs.Close

Set itmX = lsw.ListItems.Add(, , "")
    itmX.SubItems(7) = "_________"
    itmX.SubItems(12) = "_________"
    itmX.SubItems(13) = "_________"
    itmX.SubItems(14) = "_________"

Set itmX = lsw.ListItems.Add(, , "")
    itmX.SubItems(7) = Format(curTotal(0), "Standard")
    itmX.SubItems(12) = Format(curTotal(1), "Standard")
    itmX.SubItems(13) = Format(curTotal(2), "Standard")
    itmX.SubItems(14) = Format(curTotal(1) + curTotal(2), "Standard")



Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Resize()
On Error Resume Next


imgBanner.Width = Me.Width

ssTab.Width = Me.Width - 150
ssTab.Height = Me.Height - (ssTab.top + 250)

lsw.Height = ssTab.Height - (lsw.top + GroupBox1.Height + 1200)
lsw.Width = ssTab.Width - (lsw.Left + 300)


lblContrato.top = lsw.Height + (lsw.top + 150)
cmdHistorico.top = lblContrato.top
cmdEstadoCuenta.top = lblContrato.top

GroupBox1.top = lsw.top + lsw.Height + 800
GroupBox1.Width = ssTab.Width - 250
lswSubCuentas.Width = GroupBox1.Width - 250

scConsulta.Width = ssTab.Width - 250
lswRet.Width = scConsulta.Width
lswRet.Height = ssTab.Height - 2000

lblBoleta.top = lswRet.Height + lswRet.top + 150
cmdBoleta.top = lblBoleta.top
btnReversion.top = lblBoleta.top


vGrid.Width = ssTab.Width - 350
vGrid.Height = ssTab.Height - (vGrid.top + 200)

End Sub

Private Sub lsw_DblClick()

On Error GoTo vError

If lsw.ListItems.Count = 0 Then Exit Sub
If lsw.SelectedItem = "" Then Exit Sub

gFondos.Cedula = txtCedula
gFondos.Operadora = lsw.SelectedItem
gFondos.Plan = lsw.SelectedItem.SubItems(2)
gFondos.Contrato = lsw.SelectedItem.SubItems(4)
gFondos.SubCuenta = 0

frmFNDConsultaDetalle.Show vbModal

vError:
End Sub



Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotal(2) As Currency
 
If lsw.ListItems.Count = 0 Then Exit Sub
If Item = "" Then Exit Sub


Me.MousePointer = vbHourglass

On Error GoTo vError

lblContrato.Tag = Item.SubItems(2)
lblContrato.Caption = Item.SubItems(4)

curTotal(0) = 0
curTotal(1) = 0
curTotal(2) = 0

strSQL = "select * from fnd_subCuentas where cod_operadora = " & Item.Text _
       & " and cod_plan = '" & lblContrato.Tag & "' and cod_contrato = " & lblContrato.Caption
Call OpenRecordSet(rs, strSQL)
lswSubCuentas.ListItems.Clear

Do While Not rs.EOF
   Set itmX = lswSubCuentas.ListItems.Add(, , rs!cod_Plan)
       itmX.SubItems(1) = rs!COD_CONTRATO
       itmX.SubItems(2) = rs!IdX
       itmX.SubItems(3) = rs!Cedula
       itmX.SubItems(4) = rs!Nombre
       itmX.SubItems(5) = Format(rs!Cuota, "Standard")
       itmX.SubItems(6) = IIf(rs!Estado = "A", "Activo", "Liquidado")
       itmX.SubItems(7) = Format(rs!aportes, "Standard")
       itmX.SubItems(8) = Format(rs!rendimiento, "Standard")
       itmX.SubItems(9) = Format(rs!aportes + rs!rendimiento, "Standard")
       itmX.Tag = rs!cod_Operadora

    curTotal(0) = curTotal(0) + rs!Cuota
    curTotal(1) = curTotal(1) + rs!aportes
    curTotal(2) = curTotal(2) + rs!rendimiento
    rs.MoveNext
Loop
rs.Close

Set itmX = lswSubCuentas.ListItems.Add(, , "")
    itmX.SubItems(5) = "_________"
    itmX.SubItems(7) = "_________"
    itmX.SubItems(8) = "_________"
    itmX.SubItems(9) = "_________"

Set itmX = lswSubCuentas.ListItems.Add(, , "")
    itmX.SubItems(5) = Format(curTotal(0), "Standard")
    itmX.SubItems(7) = Format(curTotal(1), "Standard")
    itmX.SubItems(8) = Format(curTotal(2), "Standard")
    itmX.SubItems(9) = Format(curTotal(1) + curTotal(2), "Standard")



Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswRet_Click()

If lswRet.ListItems.Count > 0 Then
   lblBoleta.Tag = lswRet.SelectedItem.SubItems(3)
   lblBoleta.Caption = "Boleta : " & lswRet.SelectedItem.SubItems(3)
End If

End Sub

Private Sub lswRet_DblClick()

If lswRet.ListItems.Count > 0 Then
        
   If Len(Trim(lswRet.SelectedItem.SubItems(7))) > 0 Then Exit Sub
   
   fraEntrega.Visible = True
   txtEntregaBoleta.Text = lswRet.SelectedItem.SubItems(3)
   txtEntregaMonto.Text = lswRet.SelectedItem.SubItems(6)
End If

End Sub

Private Sub lswRet_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

lblBoleta.Caption = "Boleta No." & Item.SubItems(3)
lblBoleta.Tag = Item.SubItems(3)

If btnReversion.Tag = "1" Then
    If Mid(Item.SubItems(7), 1, 1) = "P" Then
       btnReversion.Enabled = True
    Else
       btnReversion.Enabled = False
    End If
End If

End Sub

Private Sub lswSubCuentas_DblClick()
If lswSubCuentas.ListItems.Count = 0 Then Exit Sub

If lswSubCuentas.SelectedItem = "" Then Exit Sub

gFondos.Cedula = txtCedula
gFondos.Operadora = lswSubCuentas.SelectedItem.Tag
gFondos.Plan = lswSubCuentas.SelectedItem
gFondos.Contrato = lswSubCuentas.SelectedItem.SubItems(1)
gFondos.SubCuenta = lswSubCuentas.SelectedItem.SubItems(2)

frmFNDConsultaDetalle.Show vbModal

End Sub

Private Sub opt_Click(Index As Integer)
Dim i As Integer

For i = 0 To opt.Count - 1
  If i = Index Then
     opt.Item(i).Checked = True
  Else
     opt.Item(i).Checked = False
  End If
Next i

Call sbConsultaContrato


End Sub


Private Sub ssTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Select Case Item.Index
 Case 0 'Consulta Contrato
   Call sbConsultaContrato
 Case 1 'Retiros
   Call sbConsultaLiquidaciones
 Case 2 '
   txtContrato = ""
   txtPlan = ""
   Call sbConsultaMovimientos
End Select

Call Form_Resize

End Sub

Private Sub tlbProceso_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case Button.Key
  Case "aplicar"
     strSQL = "update fnd_liquidacion set traspaso_Tesoreria = dbo.MyGetdate(), traspaso_usuario = '" & glogon.Usuario _
             & "', solicitud_tesoreria = 0 where consec = " & txtEntregaBoleta.Text
     Call ConectionExecute(strSQL)
     
     MsgBox "Retiro Entregado Satisfactoriamente...!", vbInformation
     Call sbConsultaLiquidaciones
   
  Case "cancelar"
    'Nada
End Select

    fraEntrega.Visible = False


End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Identificación"
   gBusquedas.Col2Name = "Id Alterno"
   gBusquedas.Col3Name = "Nombre"
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "Select Cedula,CedulaR, Nombre From Socios"
   
   frmBusquedas.Show vbModal
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCedula.Text = Trim(gBusquedas.Resultado)
      txtNombre.Text = Trim(gBusquedas.Resultado3)
   End If
   txtCedula_LostFocus
   
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

If KeyCode = vbKeyReturn Then txtCedula_LostFocus
End Sub

Private Sub txtCedula_LostFocus()
If Trim(txtCedula.Text) = "" Then
   txtNombre.Text = ""
   ssTab.Item(2).Enabled = False
   lsw.ListItems.Clear
Else
    
    'Valida Acceso a Expediente
    vRA_Access = fxSys_RA_Consulta(Trim(txtCedula.Text), glogon.Usuario)
     
    If Not vRA_Access Then
        MsgBox "Esta persona se encuentra con -> Expediente Restringido <- Requiere de Autorización para Consultar!", vbExclamation
        txtCedula.Text = ""
        txtNombre.Text = ""
        Exit Sub
    End If
    
    
   Call sbConsultaContrato
   Call sbOperadora
   ssTab.Item(2).Enabled = True
End If
End Sub



Private Sub txtContrato_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtpDesde.SetFocus
End Sub

Private Sub txtContrato_LostFocus()
Call sbConsultaMovimientos
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCedula = Trim(gBusquedas.Resultado)
      txtNombre = Trim(gBusquedas.Resultado2)
   End If
   txtCedula_LostFocus
   
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub



Private Sub sbConsultaMovimientos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

Me.MousePointer = vbHourglass

On Error GoTo vError

 
 strSQL = "Select D.cod_fnd_Detalle,D.Monto,D.Fecha_Proceso,D.Fecha,isnull(Doc.descripcion,''),D.nCon,D.Fecha_Acredita,D.cod_contrato, D.Cod_plan,P.descripcion " _
          & " from fnd_contratos_detalle D  inner join  fnd_planes P on D.cod_plan = P.cod_plan " _
          & " inner join fnd_contratos C on D.cod_plan = C.cod_plan and D.cod_contrato = C.cod_contrato " _
          & " left join SIF_Documentos Doc on D.Tcon = Doc.Tipo_Documento" _
          & " where D.cod_operadora = " & gFondos.Operadora & " and C.cedula = '" & txtCedula.Text & "'"
          
 If Trim(txtPlan.Text) <> "" Then strSQL = strSQL & " And D.cod_plan='" & txtPlan.Text & "'"
 If Trim(txtContrato.Text) <> "" Then strSQL = strSQL & " And D.Cod_Contrato='" & txtContrato.Text & "'"
 If chkTodas.Value = vbUnchecked Then
    strSQL = strSQL & " And  D.Fecha  between '" & Format(dtpDesde, "yyyy/mm/dd") _
           & " 00:00:00' and  '" & Format(dtpHasta, "yyyy/mm/dd") & " 23:59:59'"
 End If
 strSQL = strSQL & " order by D.Fecha desc"
  
 Call sbCargaGridwOrder(vGrid, 10, strSQL, True)


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub txtPlan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"

   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"

   gBusquedas.Filtro = " And Cod_operadora=" & Trim(gFondos.Operadora)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_Planes"
   frmBusquedas.Show vbModal

   txtPlan = gBusquedas.Resultado
   txtPlan.SetFocus
   gBusquedas.Resultado = ""
End If
If KeyCode = vbKeyReturn Then txtContrato.SetFocus
  

End Sub

Private Sub sbOperadora()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select cod_operadora from fnd_contratos where cedula = '" & txtCedula & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
   gFondos.Operadora = rs!cod_Operadora
Else
   gFondos.Operadora = 0
End If
rs.Close
End Sub

Private Sub txtPlan_LostFocus()
Call sbConsultaMovimientos
End Sub
