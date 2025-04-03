VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCR_SeguimientoDesembolsos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "xx"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10440
   HelpContextID   =   3017
   Icon            =   "frmCR_SeguimientoDesembolsos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   10215
      _Version        =   1572864
      _ExtentX        =   18018
      _ExtentY        =   3413
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
      View            =   3
      FullRowSelect   =   -1  'True
      FlatScrollBar   =   -1  'True
      Appearance      =   16
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   10080
      Top             =   360
   End
   Begin XtremeSuiteControls.GroupBox fra 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   10215
      _Version        =   1572864
      _ExtentX        =   18018
      _ExtentY        =   7858
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2175
         Left            =   0
         TabIndex        =   24
         Top             =   2520
         Width           =   10095
         _Version        =   1572864
         _ExtentX        =   17806
         _ExtentY        =   3836
         _StockProps     =   79
         Caption         =   "Datos de la Cuenta Destino"
         ForeColor       =   4210752
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtIdentificación 
            Height          =   330
            Left            =   5280
            TabIndex        =   25
            Top             =   360
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3831
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboCuenta 
            Height          =   315
            Left            =   5280
            TabIndex        =   26
            Top             =   720
            Width           =   4215
            _Version        =   1572864
            _ExtentX        =   7435
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.PushButton btnCuenta 
            Height          =   315
            Left            =   9600
            TabIndex        =   27
            Top             =   720
            Width           =   375
            _Version        =   1572864
            _ExtentX        =   656
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "..."
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
            Appearance      =   17
         End
         Begin XtremeSuiteControls.ComboBox cboTipoId 
            Height          =   330
            Left            =   1680
            TabIndex        =   30
            Top             =   360
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboDivisa 
            Height          =   330
            Left            =   1680
            TabIndex        =   32
            Top             =   720
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtEntidad 
            Height          =   330
            Left            =   1680
            TabIndex        =   35
            Top             =   1080
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3831
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCorreo 
            Height          =   315
            Left            =   5280
            TabIndex        =   37
            Top             =   1080
            Width           =   4215
            _Version        =   1572864
            _ExtentX        =   7435
            _ExtentY        =   556
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDetalle 
            Height          =   435
            Left            =   1680
            TabIndex        =   39
            Top             =   1440
            Width           =   7815
            _Version        =   1572864
            _ExtentX        =   13785
            _ExtentY        =   767
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle"
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
            Index           =   10
            Left            =   360
            TabIndex        =   38
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Correo"
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
            Index           =   9
            Left            =   3960
            TabIndex        =   36
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Entidad"
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
            Index           =   8
            Left            =   360
            TabIndex        =   34
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
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
            Index           =   7
            Left            =   360
            TabIndex        =   33
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Id"
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
            Left            =   360
            TabIndex        =   31
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Identificación"
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
            Left            =   3960
            TabIndex        =   29
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta"
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
            Index           =   14
            Left            =   3960
            TabIndex        =   28
            Top             =   720
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.CheckBox chkDesembolso 
         Height          =   252
         Left            =   6960
         TabIndex        =   13
         Top             =   1680
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Desembolsa?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtConcepto 
         Height          =   312
         Left            =   1680
         TabIndex        =   6
         Top             =   600
         Width           =   7812
         _Version        =   1572864
         _ExtentX        =   13779
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   312
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   7812
         _Version        =   1572864
         _ExtentX        =   13785
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
         Height          =   312
         Left            =   3840
         TabIndex        =   9
         Top             =   1320
         Width           =   5652
         _Version        =   1572864
         _ExtentX        =   9970
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDifiere 
         Height          =   312
         Left            =   5280
         TabIndex        =   12
         Top             =   1680
         Width           =   1332
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtCuenta 
         Height          =   312
         Left            =   1680
         TabIndex        =   8
         Top             =   1320
         Width           =   2172
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   1680
         TabIndex        =   10
         Top             =   1680
         Width           =   2172
         _Version        =   1572864
         _ExtentX        =   3831
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtReferencia 
         Height          =   312
         Left            =   1680
         TabIndex        =   21
         Top             =   2040
         Width           =   2172
         _Version        =   1572864
         _ExtentX        =   3831
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
         Height          =   330
         Left            =   5280
         TabIndex        =   22
         Top             =   2040
         Width           =   4215
         _Version        =   1572864
         _ExtentX        =   7435
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Giro"
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
         Index           =   6
         Left            =   3960
         TabIndex        =   23
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
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
         TabIndex        =   20
         Top             =   2040
         Width           =   972
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   18
         Top             =   120
         Width           =   10335
         _Version        =   1572864
         _ExtentX        =   18230
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Indique el desembolso o Rebajo que desea registrar:"
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
         VisualTheme     =   6
      End
      Begin VB.Image imgBusqueda_Rapida 
         Height          =   264
         Index           =   0
         Left            =   9600
         Picture         =   "frmCR_SeguimientoDesembolsos.frx":030A
         Stretch         =   -1  'True
         ToolTipText     =   "Busqueda Rápida"
         Top             =   1320
         Width           =   264
      End
      Begin VB.Image imgBusqueda_Rapida 
         Height          =   252
         Index           =   1
         Left            =   9600
         Picture         =   "frmCR_SeguimientoDesembolsos.frx":0D18
         Stretch         =   -1  'True
         ToolTipText     =   "Busqueda Rápida"
         Top             =   600
         Width           =   252
      End
      Begin VB.Label lblDifiere 
         BackStyle       =   0  'Transparent
         Caption         =   "Diferir hasta?"
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
         Left            =   3960
         TabIndex        =   11
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
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
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Index           =   4
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   972
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta "
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
         TabIndex        =   3
         Top             =   1320
         Width           =   1572
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
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
         TabIndex        =   2
         Top             =   960
         Width           =   1212
      End
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   330
      Left            =   6360
      TabIndex        =   17
      Top             =   480
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertar"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
            Object.ToolTipText     =   "Borra el registro en pantalla de la base de datos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
            Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
            Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cerrar"
            Object.ToolTipText     =   "Cierra esta ventana"
         EndProperty
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption scRegistrado 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Width           =   10215
      _Version        =   1572864
      _ExtentX        =   18013
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Desembolsos / Rebajos (Registrados)"
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
      VisualTheme     =   6
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Disponible:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   7200
      TabIndex        =   15
      Top             =   84
      Width           =   1092
   End
   Begin VB.Label lblDisponible 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   8040
      TabIndex        =   14
      Top             =   84
      Width           =   2052
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10572
      _Version        =   1572864
      _ExtentX        =   18648
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Desembolsos y Rebajos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
End
Attribute VB_Name = "frmCR_SeguimientoDesembolsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngID_Desembolso As Long
Dim vEdita As Integer, mFecha As Date
Dim curPrimerCuota As Currency, curPoliza As Currency, curInteres As Currency

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean


Private Sub btnCuenta_Click()
Dim strSQL As String

On Error GoTo vError


GLOBALES.gTag = Trim(txtIdentificación.Text)
GLOBALES.gTag2 = "CRD"

frmCC_Cuentas_Bancarias.Show vbModal

txtIdentificación_LostFocus

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

Me.Caption = "Operación : " & Operacion.Operacion

curPrimerCuota = 0
curPoliza = 0
curInteres = 0


cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.AddItem fxTipoDocumento("ND")
cboTipoDocumento.Text = fxTipoDocumento("TE")

vPaso = True
strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
       & " order by Tipo_Id"
    Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)
vPaso = False

strSQL = "select COD_DIVISA as 'IdX', DESCRIPCION as 'ItmX'   From vSys_Divisas"
Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)

cboBanco.Clear

With lsw.ColumnHeaders
    .Clear
    .Add , , "Concepto", 3500
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Cuenta", 1600, vbCenter
    .Add , , "Retiene?", 1000, vbCenter
    .Add , , "Modifica?", 1000, vbCenter
    .Add , , "Difiere?", 1000, vbCenter
    .Add , , "Corte", 1800, vbCenter
    .Add , , "Cuenta Bancaria", 3400
    .Add , , "[Id Cta Bancaria]", 1200, vbCenter
    
    .Add , , "Referencia", 1200, vbCenter
    .Add , , "Identificación", 1200, vbCenter
    
    .Add , , "Emite", 900, vbCenter
    .Add , , "IBAN/Interna", 2000, vbCenter
    
    
End With

Call sbToolBarIconos(tlbPrincipal, False)

With tlbPrincipal
    .Buttons(1).Enabled = True
    .Buttons(2).Enabled = False
    .Buttons(3).Enabled = False
    .Buttons(4).Enabled = False
    .Buttons(5).Enabled = False
End With
fra.Enabled = False

End Sub

Sub LimpiaDatos()
 txtConcepto = ""
 txtCuentaDesc.Text = ""
 txtCuenta.Text = ""
 txtMonto.Text = ""
 lngID_Desembolso = 0
 
 txtReferencia.Text = ""
 txtIdentificación.Text = ""
 
 cboCuenta.Clear
 cboTipoDocumento.Text = fxTipoDocumento("ND")
 fra.Enabled = False
 
 txtConcepto.Locked = True
 imgBusqueda_Rapida.Item(0).Enabled = False
 txtCuenta.Enabled = False
 
End Sub

Private Sub sbCargaDesembolsos()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem
Dim curTotal As Currency

curTotal = 0

strSQL = "select D.*,isnull(D.cod_Banco,0) as 'Banco',  rtrim(isnull(B.descripcion,'')) as 'BancoDesc'" _
       & " from Desembolsos D left join Tes_Bancos B on D.cod_Banco = B.id_Banco" _
       & " where D.id_solicitud= " & Operacion.Operacion

Call OpenRecordSet(rs, strSQL)
lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!CONCEPTO)
      itmX.SubItems(1) = Format(rs!Monto, "Standard")
      itmX.SubItems(2) = fxCntX_CuentaFormato(True, rs!Cuenta_Conta, 0)
      itmX.SubItems(3) = IIf(rs!retener = 1, "Si", "No")
      itmX.SubItems(4) = IIf(rs!Modifica = 1, "Si", "No")
      itmX.SubItems(5) = IIf(rs!DIFERIDO_APLICA = 1, "Si", "No")
      itmX.SubItems(6) = rs!DIFERIDO_CORTE
      
      itmX.SubItems(7) = Trim(rs!BancoDesc & "")
      itmX.SubItems(8) = rs!Banco
      
      itmX.SubItems(9) = rs!Referencia & ""
      itmX.SubItems(10) = rs!Identificacion & ""
      
      itmX.SubItems(11) = rs!TDOCUMENTO & ""
      itmX.SubItems(12) = rs!CTA_BANCO & ""
      
      
      itmX.Tag = rs!id_desembolso
      
      curTotal = curTotal + rs!Monto
 rs.MoveNext
Loop
rs.Close

scRegistrado.Caption = "Desembolsos / Rebajos (Registrado..:" & Format(curTotal, "Standard") & ")"

End Sub



Private Sub imgBusqueda_Rapida_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


Select Case Index
 Case 0
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuenta.Text = gCuenta
    txtCuenta.Text = fxCntX_CuentaFormato(True, txtCuenta.Text, 0)
    txtCuentaDesc.Text = fxgCntCuentaDesc(gCuenta)
    
 Case 1

    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "select cod_condeb as 'Código',Descripcion from concepto_desemb"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Filtro = " and Activo = 1"
    frmBusquedas.Show vbModal
    
    txtConcepto = gBusquedas.Resultado2
    
    strSQL = "select retiene,modifica,cod_cuenta,difiere,dbo.MyGetdate() as 'difiere_fecha' from concepto_desemb where cod_condeb = " & gBusquedas.Resultado
    Call OpenRecordSet(rs, strSQL)
    
    txtConcepto.Tag = rs!Retiene
    
    
    txtCuentaDesc.Tag = rs!Modifica
    txtCuenta.Text = fxCntX_CuentaFormato(True, rs!cod_cuenta, 0)
    
    txtCuentaDesc.Text = fxgCntCuentaDesc(rs!cod_cuenta)
    
    mFecha = rs!difiere_fecha
    
    If rs!Retiene = 1 Then
       chkDesembolso.Value = vbUnchecked
       cboBanco.Enabled = False
       
       cboTipoDocumento.Text = fxTipoDocumento("ND")
       cboTipoDocumento.Enabled = False
    Else
       chkDesembolso.Value = vbChecked
       cboBanco.Enabled = True
       
       cboTipoDocumento.Text = fxTipoDocumento("TE")
       cboTipoDocumento.Enabled = True
    End If
    
    If rs!Difiere = 1 Then
       dtpDifiere.Value = rs!difiere_fecha
       dtpDifiere.Visible = True
       lblDifiere.Visible = True
    Else
       dtpDifiere.Visible = False
       lblDifiere.Visible = False
    End If
    
    
    If rs!Modifica = 1 Then
       txtConcepto.Locked = False
       txtCuenta.Enabled = True
       imgBusqueda_Rapida.Item(0).Enabled = True
     Else
       txtConcepto.Locked = True
       txtCuenta.Enabled = False
       imgBusqueda_Rapida.Item(0).Enabled = False
    End If
    
    rs.Close

End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

 lngID_Desembolso = Item.Tag
 txtConcepto.Text = Item.Text
 txtMonto.Text = Item.SubItems(1)
 txtCuenta.Text = Item.SubItems(2)
 txtCuentaDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, Item.SubItems(2), 0))
 
 txtConcepto.Tag = IIf(Item.SubItems(3) = "Si", 1, 0)
 txtCuentaDesc.Tag = IIf(Item.SubItems(4) = "Si", 1, 0)
    
 txtReferencia.Text = Item.SubItems(9)
 txtIdentificación.Text = Item.SubItems(10)
 
 cboTipoDocumento.Text = fxTipoDocumento(Item.SubItems(11))
 
 cboCuenta.Clear
 Call sbCboAsignaDato(cboCuenta, Item.SubItems(12), True, Item.SubItems(12))
 
    
    
 'Retiene
 If IIf(Item.SubItems(3) = "Si", 1, 0) = 1 Then
    chkDesembolso.Value = vbUnchecked
    cboBanco.Enabled = False
    
    cboTipoDocumento.Enabled = False
    
 Else
    chkDesembolso.Value = vbChecked
    cboBanco.Enabled = True
    Call sbCboAsignaDato(cboBanco, Trim(Item.SubItems(7)), True, Item.SubItems(8))
 
    cboTipoDocumento.Enabled = True
 
 End If
    
 'Modifica
 If IIf(Item.SubItems(4) = "Si", 1, 0) = 1 Then
    txtConcepto.Locked = False
    txtCuenta.Enabled = True
    imgBusqueda_Rapida.Item(0).Enabled = True
 Else
    txtConcepto.Locked = True
    txtCuenta.Enabled = False
    imgBusqueda_Rapida.Item(0).Enabled = False
 End If
 

 If IIf(Item.SubItems(5) = "Si", 1, 0) = 1 Then
    lblDifiere.Visible = True
    dtpDifiere.Visible = True
    
    dtpDifiere.Value = Item.SubItems(6)

 Else
    lblDifiere.Visible = False
    dtpDifiere.Visible = False
 End If


With tlbPrincipal
   .Buttons(1).Enabled = False
   .Buttons(2).Enabled = True
   .Buttons(3).Enabled = True
   .Buttons(4).Enabled = False
   .Buttons(5).Enabled = False
End With


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

curPrimerCuota = 0
curPoliza = 0
curInteres = 0

cboBanco.Clear
mFecha = fxFechaServidor


strSQL = "exec spCrd_SGT_Bancos_Desembolso '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  MsgBox "No existen Bancos [Creados/Asignados], verifique en Tesoreria...", vbCritical

Else
 Do While Not rs.EOF
    cboBanco.AddItem IIf(IsNull(rs!Descripcion), "SIN DESCRIPCION", rs!Descripcion)
    cboBanco.ItemData(cboBanco.ListCount - 1) = CStr(rs!Id_Banco)
   
   rs.MoveNext
 Loop
 rs.MoveFirst
 Call sbCboAsignaDato(cboBanco, IIf(IsNull(rs!Descripcion), "SIN DESCRIPCION", rs!Descripcion), True, rs!Id_Banco)
End If
rs.Close


strSQL = "select R.Primer_Cuota,R.Garantia,R.montoapr,R.cuota,R.int,C.convenio,R.cod_destino" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
       & " where R.id_solicitud =" & Operacion.Operacion
Call OpenRecordSet(rs, strSQL)

If fxCobraTasaFormaliza(rs!cod_destino & "") Then
  curInteres = fxInteresesHastaFormalizar(Operacion.FechaDesembolso, , Operacion.PriDeduc, Operacion.DiaPago)
End If
  
  
  If rs!PRIMER_CUOTA = "S" Then
    curPrimerCuota = rs!Cuota
   If curInteres > 0 Then 'Convenios no cobran
      curInteres = fxInteresesDiasPrimerCuota(Operacion.FechaDesembolso, rs!montoapr, rs!Int)
   End If
  End If
  
  If rs!Garantia <> "H" And rs!Convenio = "N" Then curPoliza = fxCuotaPolizaVida(rs!montoapr)
rs.Close


 lblDisponible.Caption = Format(Operacion.MontoAprobado - (fxMontoEnGeneral(Operacion.Operacion) _
                       + curInteres + curPrimerCuota + curPoliza) _
                       , "Standard")

 

 Call sbCargaDesembolsos

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 txtCuenta.Text = fxCntX_CuentaFormato(False, txtCuenta.Text, 0)
 txtCuentaDesc.Text = fxgCntCuentaDesc(txtCuenta.Text)
 txtCuenta.Text = fxCntX_CuentaFormato(True, txtCuenta.Text, 0)
 txtConcepto.SetFocus
End If
End Sub

Private Sub txtIdentificación_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

'(@Identificacion varchar(30), @BancoId int, @DivisaCheck smallint = 0)"
strSQL = "exec spSys_Cuentas_Bancarias '" & txtIdentificación & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call OpenRecordSet(rs, strSQL)

cboCuenta.Clear
Do While Not rs.EOF
  cboCuenta.AddItem rs!IdX
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub
vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError
    txtMonto.Text = CCur(txtMonto.Text)
vError:

End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
On Error GoTo vError
  If KeyAscii = vbKeyReturn Then txtCuenta.SetFocus
vError:
End Sub

Private Function fxVerificaDatosDesembolsos() As Boolean
Dim vMensaje As String

fxVerificaDatosDesembolsos = True
vMensaje = ""
If Len(Trim(txtConcepto)) = 0 Then vMensaje = vMensaje & vbCrLf & "- El concepto no es válido"
If IsNumeric(txtMonto.Text) Then
   If CCur(txtMonto.Text) > CCur(lblDisponible.Caption) Then vMensaje = vMensaje & vbCrLf & "- El monto a desembolsar es mayor al disponible del préstamo"
End If

If Not fxgCntCuentaValida(txtCuenta.Text) Then vMensaje = vMensaje & vbCrLf & "- La cuenta contable no es válida"


If lblDifiere.Visible And dtpDifiere.Value < mFecha Then
    vMensaje = vMensaje & vbCrLf & "- El corte para diferir no puede ser menor a hoy!"
End If

If Len(vMensaje) > 0 Then
  fxVerificaDatosDesembolsos = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardaDesembolso()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipo As String, vBanco As Integer
Dim vDifiere As Integer

On Error GoTo vError
 
 vTipo = fxTipoDocumento(cboTipoDocumento.Text)
 vBanco = cboBanco.ItemData(cboBanco.ListIndex)
 
 If chkDesembolso.Value = vbUnchecked Then
    vBanco = 0
    vTipo = "ND"
 End If
 
 If lblDifiere.Visible Then
   vDifiere = 1
 Else
   vDifiere = 0
   dtpDifiere.Value = mFecha
 End If

 If vEdita = 0 Then
   strSQL = "insert desembolsos(ID_SOLICITUD,CODIGO,CONCEPTO,MONTO,CUENTA_CONTA," _
          & "TDOCUMENTO,DEPOSITAR,COD_BANCO,RETENER,MODIFICA,DIFERIDO_APLICA,DIFERIDO_CORTE" _
          & ", REFERENCIA, IDENTIFICACION, CTA_BANCO) values(" _
          & Operacion.Operacion & ",'" & Operacion.Codigo & "','" & UCase(txtConcepto) _
          & "'," & CCur(txtMonto.Text) & ",'" & fxCntX_CuentaFormato(False, Trim(txtCuenta.Text), 0) & "'" _
          & ",'" & vTipo & "',0," & vBanco & "," & IIf((chkDesembolso.Value = vbChecked), 0, 1) _
          & "," & txtCuentaDesc.Tag & "," & vDifiere & ",'" & Format(dtpDifiere.Value, "yyyy/mm/dd") _
          & "','" & txtReferencia.Text & "','" & Trim(txtIdentificación.Text) & "','" & Trim(cboCuenta.Text) & "')"
 Else
   strSQL = "update desembolsos set concepto = '" & Trim(UCase(txtConcepto)) _
          & "',monto=" & CCur(txtMonto.Text) & ",cuenta_conta = '" & fxCntX_CuentaFormato(False, Trim(txtCuenta.Text), 0) _
          & "',retener = " & IIf((chkDesembolso.Value = vbChecked), 0, 1) & ",modifica = " & txtCuentaDesc.Tag _
          & ",tdocumento = '" & vTipo & "',cod_banco = " & vBanco & ", CTA_BANCO = '" & Trim(cboCuenta.Text) & "'" _
          & ",DIFERIDO_APLICA = " & vDifiere & ",DIFERIDO_CORTE = ' " & Format(dtpDifiere.Value, "yyyy/mm/dd") _
          & "', REFERENCIA = '" & txtReferencia.Text & "', IDENTIFICACION = '" & txtIdentificación.Text _
          & "' where id_desembolso = " & lngID_Desembolso
 End If
 Call ConectionExecute(strSQL)
 
 MsgBox "Desembolso Guardado Satisfactoriamente...", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim iRespuesta As Integer, strSQL As String

Select Case Button.Key
  Case "insertar", "nuevo"
   vEdita = 0
   Call LimpiaDatos
    With tlbPrincipal
       .Buttons(1).Enabled = False
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = True
       .Buttons(5).Enabled = True
    End With
    fra.Enabled = True
    txtConcepto.SetFocus
    Call imgBusqueda_Rapida_Click(1)
  
  Case "editar", "modificar"
   vEdita = 1
    With tlbPrincipal
       .Buttons(1).Enabled = False
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = True
       .Buttons(5).Enabled = True
    End With
    fra.Enabled = True
    txtConcepto.SetFocus
  
  Case "borrar"
   strSQL = "delete desembolsos where id_desembolso=" & lngID_Desembolso
   If lngID_Desembolso > 0 Then
    iRespuesta = MsgBox("Esta seguro que desea eliminar este desembolso", vbYesNo)
    If iRespuesta = vbYes Then
      Call ConectionExecute(strSQL)
      lblDisponible.Caption = Format(Operacion.MontoAprobado - (fxMontoEnGeneral(Operacion.Operacion) _
                            + curInteres + curPrimerCuota + curPoliza) _
                            , "Standard")
      Call sbCargaDesembolsos
      Call LimpiaDatos
    Else
      Call LimpiaDatos
    End If
    With tlbPrincipal
       .Buttons(1).Enabled = True
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = False
       .Buttons(5).Enabled = False
     End With
    
   End If
  
  Case "salvar", "guardar"
    If fxVerificaDatosDesembolsos Then
      Call sbGuardaDesembolso
      lblDisponible.Caption = Format(Operacion.MontoAprobado - (fxMontoEnGeneral(Operacion.Operacion) _
                            + curInteres + curPrimerCuota + curPoliza) _
                            , "Standard")
      Call sbCargaDesembolsos
      With tlbPrincipal
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        .Buttons(4).Enabled = False
        .Buttons(5).Enabled = False
      End With
      Call LimpiaDatos
    Else
      MsgBox "Información Ingresada es Incorrecta por favor verifique...", vbInformation
    End If
  
  Case "deshacer"
    Call LimpiaDatos
    With tlbPrincipal
       .Buttons(1).Enabled = True
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = False
       .Buttons(5).Enabled = False
    End With
  
  Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
        
  Case "salir", "cerrar"
    Unload Me
End Select
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  txtConcepto = UCase(txtConcepto)
  txtMonto.SetFocus
End If
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
    txtMonto.Text = Format(CCur(txtMonto.Text), "Standard")
vError:
End Sub
