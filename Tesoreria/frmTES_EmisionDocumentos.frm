VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmTES_EmisionDocumentos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Emisión de Documentos"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11295
   HelpContextID   =   1007
   Icon            =   "frmTES_EmisionDocumentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3372
      Left            =   120
      TabIndex        =   25
      Top             =   2916
      Width           =   11052
      _Version        =   1572864
      _ExtentX        =   19494
      _ExtentY        =   5948
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
      Appearance      =   16
      UseVisualStyle  =   0   'False
      Sorted          =   -1  'True
   End
   Begin VB.Frame fraCuentaPuenta 
      BorderStyle     =   0  'None
      Caption         =   "Solicitudes : Registradas en la Cuenta de Bancos (Puente)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   7212
      Left            =   0
      TabIndex        =   2
      Top             =   8040
      Visible         =   0   'False
      Width           =   11412
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5532
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   10932
         _Version        =   524288
         _ExtentX        =   19283
         _ExtentY        =   9758
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         MaxRows         =   1000000
         ScrollBars      =   0
         SpreadDesigner  =   "frmTES_EmisionDocumentos.frx":030A
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboCtaPuente 
         Height          =   312
         Left            =   2520
         TabIndex        =   26
         Top             =   360
         Width           =   6972
         _Version        =   1572864
         _ExtentX        =   12303
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
      Begin XtremeSuiteControls.PushButton btnPuenteAccion 
         Height          =   432
         Index           =   0
         Left            =   6240
         TabIndex        =   27
         Top             =   6600
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
         _ExtentY        =   762
         _StockProps     =   79
         Caption         =   "&Mover a Cuenta Actual"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmTES_EmisionDocumentos.frx":0A3B
      End
      Begin XtremeSuiteControls.PushButton btnPuenteAccion 
         Height          =   432
         Index           =   1
         Left            =   8760
         TabIndex        =   28
         Top             =   6600
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   762
         _StockProps     =   79
         Caption         =   "&Cerrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmTES_EmisionDocumentos.frx":1240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Solicitudes registradas en la Cuenta Bancaria [Puente]. Seleccione las que desee incorporar a su cuenta actual e indique mover."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   480
         TabIndex        =   5
         Top             =   6600
         Width           =   5412
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Puente"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   1692
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1212
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   5292
      _Version        =   1572864
      _ExtentX        =   9334
      _ExtentY        =   2138
      _StockProps     =   79
      Caption         =   "Casos a emitir:"
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
      Begin XtremeSuiteControls.RadioButton optGeneraPor 
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Solicitudes"
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   2280
         TabIndex        =   29
         Top             =   720
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   3600
         TabIndex        =   30
         Top             =   720
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
      Begin XtremeSuiteControls.FlatEdit txtGeneraNumeroDe 
         Height          =   312
         Left            =   2280
         TabIndex        =   37
         Top             =   360
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.FlatEdit txtGeneraNumeroHasta 
         Height          =   312
         Left            =   3600
         TabIndex        =   38
         Top             =   360
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.RadioButton optGeneraPor 
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   40
         Top             =   720
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fechas"
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   0
         Left            =   2280
         TabIndex        =   20
         Top             =   120
         Width           =   1332
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   2
         Left            =   3600
         TabIndex        =   19
         Top             =   120
         Width           =   1212
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1212
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   11052
      _Version        =   1572864
      _ExtentX        =   19494
      _ExtentY        =   2138
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdGenera 
         Height          =   672
         Left            =   9720
         TabIndex        =   9
         Top             =   360
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   1182
         _StockProps     =   79
         Caption         =   "&Genera"
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
         Picture         =   "frmTES_EmisionDocumentos.frx":1A0D
      End
      Begin XtremeSuiteControls.PushButton cmdPrevista 
         Height          =   672
         Left            =   8040
         TabIndex        =   10
         Top             =   360
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   1182
         _StockProps     =   79
         Caption         =   "&Prevista"
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
         Picture         =   "frmTES_EmisionDocumentos.frx":2212
      End
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   672
         Left            =   6480
         TabIndex        =   11
         Top             =   360
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   1185
         _StockProps     =   79
         Caption         =   "&Informe de Solicitudes"
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
         Picture         =   "frmTES_EmisionDocumentos.frx":28D9
      End
      Begin XtremeSuiteControls.PushButton cmdPuente 
         Height          =   672
         Left            =   5160
         TabIndex        =   12
         Top             =   360
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   1182
         _StockProps     =   79
         Caption         =   "&Cuenta Puente"
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
         Picture         =   "frmTES_EmisionDocumentos.frx":3095
      End
      Begin XtremeSuiteControls.PushButton cmdCuentaVerifica 
         Height          =   672
         Left            =   3840
         TabIndex        =   13
         Top             =   360
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   1182
         _StockProps     =   79
         Caption         =   "&Verifica Cuentas"
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
         Picture         =   "frmTES_EmisionDocumentos.frx":3857
      End
      Begin XtremeSuiteControls.FlatEdit txtCasos 
         Height          =   312
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   612
         _Version        =   1572864
         _ExtentX        =   1080
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
         Text            =   "0"
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   840
         TabIndex        =   36
         Top             =   600
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label lblEnd2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   972
      End
      Begin VB.Label lblEnd1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Casos"
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   612
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10680
      Top             =   120
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
            Picture         =   "frmTES_EmisionDocumentos.frx":41CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_EmisionDocumentos.frx":42D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   600
   End
   Begin XtremeSuiteControls.GroupBox GroupBox_CK 
      Height          =   1284
      Left            =   6120
      TabIndex        =   14
      Top             =   1200
      Width           =   4812
      _Version        =   1572864
      _ExtentX        =   8488
      _ExtentY        =   2265
      _StockProps     =   79
      Caption         =   "Control de Tiraje:"
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
      Begin XtremeSuiteControls.FlatEdit txtDocInicial 
         Height          =   312
         Left            =   3120
         TabIndex        =   32
         Top             =   240
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
         Text            =   "0"
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtVerificacion 
         Height          =   312
         Left            =   3120
         TabIndex        =   33
         Top             =   600
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
         Text            =   "0"
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCantidadSolicitudes 
         Height          =   312
         Left            =   3120
         TabIndex        =   34
         Top             =   960
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
         Text            =   "0"
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label lblCheque 
         Caption         =   "Documento Inicial"
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
         Left            =   840
         TabIndex        =   17
         Top             =   300
         Width           =   2172
      End
      Begin VB.Label lblSolicitudesGenerar 
         Caption         =   "Solicitudes a Generar"
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
         Left            =   840
         TabIndex        =   16
         Top             =   1020
         Width           =   2172
      End
      Begin VB.Label lblVerifica 
         Caption         =   "Verificar generación cada"
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
         Left            =   840
         TabIndex        =   15
         Top             =   660
         Width           =   2172
      End
   End
   Begin XtremeSuiteControls.ComboBox cboDoc 
      Height          =   312
      Left            =   5400
      TabIndex        =   21
      Top             =   360
      Width           =   5292
      _Version        =   1572864
      _ExtentX        =   9340
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Width           =   5172
      _Version        =   1572864
      _ExtentX        =   9128
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
   Begin XtremeSuiteControls.ComboBox cboFormato 
      Height          =   312
      Left            =   5400
      TabIndex        =   23
      Top             =   720
      Width           =   3492
      _Version        =   1572864
      _ExtentX        =   6165
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
   Begin XtremeSuiteControls.ComboBox cboPlan 
      Height          =   312
      Left            =   8880
      TabIndex        =   41
      Top             =   720
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.PushButton btnPlanes 
      Height          =   315
      Left            =   10800
      TabIndex        =   42
      Top             =   720
      Width           =   372
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   556
      _StockProps     =   79
      BackColor       =   -2147483643
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmTES_EmisionDocumentos.frx":4409
   End
   Begin XtremeShortcutBar.ShortcutCaption lblX 
      Height          =   372
      Left            =   120
      TabIndex        =   31
      Top             =   2520
      Width           =   11052
      _Version        =   1572864
      _ExtentX        =   19494
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Solicitudes a Generar"
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Formato TE:"
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
      Index           =   1
      Left            =   3960
      TabIndex        =   24
      Top             =   720
      Width           =   2172
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta Bancaria..:"
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
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Documento..:"
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
      Height          =   315
      Index           =   4
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmTES_EmisionDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim mSolInicio As String, mSolCorte As String, mFechaInicio As String, mFechaCorte As String

''Procedimiento para crear el nuevo archivo del BCR, Banca Empresarial
Private Sub sbTeBCT_Enlace()
Dim rs As New ADODB.Recordset, strArchivo As String
Dim strPath As String, strCadena As String, i As Integer
Dim vCuentaBanco As String, vFecha As Date, vPaso As Boolean
Dim vConArchivo As String, vCedulaReg As String, vNumNegocio As String, vRazon As String
Dim strSQL As String, fn, vTesKeyCh As String
Dim strCedJuridica As String
Dim iLineInicio As Integer 'variable para la linea con la que inicia el detalle de las transferencias
Dim strLinea As String

On Error GoTo vError



vPaso = False
vFecha = fxFechaServidor


i = 1

fn = FreeFile

vRazon = ""
vNumNegocio = ""
vCedulaReg = ""

    
     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "+")
     
     gTesGlobal.BancoNombre = cbo.Text
     iLineInicio = 1
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     strArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & ".txt"
     
     Open strArchivo For Output As #1

     '******************************
     ' DETALLE DE LA TRANSFERENCIA *
     '******************************

    strSQL = "exec spTES_BCT_Enlace " & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc _
            & "'," & gTesGlobal.BancoConsec & "," & txtCantidadSolicitudes _
            & "," & mSolInicio & "," & mSolCorte & "," & mFechaInicio & "," & mFechaCorte
    Call OpenRecordSet(rs, strSQL)
    
    Do While Not rs.EOF
        strCadena = rs!Linea
        Print #1, strCadena
     
     rs.MoveNext
    Loop
    rs.Close
   
     Close #1   ' Close file.
     
     Me.Hide
         
     frmTES_Transferencias.lblBanco = cbo.Text
     frmTES_Transferencias.lblArchivo = strArchivo
     frmTES_Transferencias.Show vbModal
     
     Me.Show
Exit Sub

vError:
       
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
   If vPaso Then
      gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "-")
   End If
End Sub






''Procedimiento para crear el nuevo archivo del BCR, Banca Empresarial
Private Sub sbTeBCR_Empresarial()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite la Transferencia en formato Empresarial para el BCR. Genera archivo de
'               texto en la direccion "C:\NombreEmpresa\Banco\Fecha\ConsecutivoDeposito.txt"
'               y finalmente despliega el formulario de control de transferencias.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset, strArchivo As String
Dim strPath As String, strCadena As String, i As Integer
Dim vCuentaBanco As String, vFecha As Date, vPaso As Boolean
Dim vConArchivo As String, vCedulaReg As String, vNumNegocio As String, vRazon As String
Dim strSQL As String, fn, vTesKeyCh As String
Dim strCedJuridica As String
Dim iLineInicio As Integer 'variable para la linea con la que inicia el detalle de las transferencias
Dim strLinea As String

On Error GoTo vError



vPaso = False
vFecha = fxFechaServidor

' Linea 1 -> Numero de Negocio (Registrado en el Sistema del BCR, SCIC)
' Linea 2 -> Cedula de Registro
' Linea 3 -> Razon o Detalle del Pago
i = 1

fn = FreeFile

vRazon = ""
vNumNegocio = ""
vCedulaReg = ""

strSQL = "select  REPLACE(cedula_juridica,'-','') as 'Cedula_Juridica',NOMBRE" _
       & " From SIF_EMPRESA"
Call OpenRecordSet(rs, strSQL, 0)
    vNumNegocio = Trim(rs!cedula_juridica)
    vCedulaReg = Trim(rs!cedula_juridica)
    vRazon = "TRANSFERENCIAS " & rs!Nombre
rs.Close

'strArchivo = SIFGlobal.DirectorioDeResultados & "\Configuracion\BCRFormat.ini"
'If Dir(strArchivo, vbArchive) = "" Then
'  strArchivo = App.Path & "\BCRFormat.ini"
'End If
'
'Open strArchivo For Input As #fn
' Do While Not EOF(fn)
'   Input #fn, strSQL
'   Select Case i
'     Case 1
'       vNumNegocio = strSQL
'     Case 2
'       vCedulaReg = strSQL
'     Case 3
'       vRazon = strSQL
'   End Select
'   i = i + 1
' Loop
'Close #fn   ' Close file.

    
     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "+")
     
     gTesGlobal.BancoNombre = cbo.Text
     vPaso = True
     i = 1
    strSQL = "select documento_base,count(*) From Tes_Transacciones" _
         & " where id_banco = " & gTesGlobal.BancoID & " and fecha_emision = '" _
         & Format(vFecha, "yyyy/mm/dd") & "' and estado = 'T' group by documento_base"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     i = i + 1
     rs.MoveNext
    Loop
    rs.Close
    vConArchivo = Format(i, "000")
     
     
     strSQL = "select dbo.fxTesCantidadTEDiarias('" & Format(vFecha, "yyyy/mm/dd") & "' ," & gTesGlobal.BancoID & ") as 'Cantidad'"
     Call OpenRecordSet(rs, strSQL)
         iLineInicio = rs!Cantidad
     rs.Close
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     strArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & ".txt"
     
     Open strArchivo For Output As #1

     'REGISTRO DE CONTROL
     i = 1
    
     strCadena = "000"                                                              'Estado 3
     strCadena = strCadena & SIFGlobal.fxStringRelleno(vCedulaReg, "I", "0", 12)    'Cedula Juridica 12
     strCadena = strCadena & vConArchivo                                            'Consecutivo Archivo 3
     strCadena = strCadena & Format(vFecha, "ddmmyyyy")                             'Fecha Aplicacion 8
     strCadena = strCadena & "000000000000"                                         'Cedula de Registro 12
     strCadena = strCadena & "000000000000"                                         '12 TestKey  no se genera, se rellena con ceros
     strCadena = strCadena & "000000"                                               '6 Hora Estado Se rellena con ceros
     strCadena = strCadena & Space(6)                                               'filler 6 espacios en blanco
     strCadena = strCadena & "TLB"                                                  'Tipo de archivo
     strCadena = strCadena & Space(128)                                             'filler 128 espacios en blanco
     strCadena = strCadena & "D"                                                    'Tipo de movinento Debido
    
     Print #1, strCadena
     
  
    'DEBITOS
    strSQL = "exec spTES_BCR_Empresarial 2," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
           & "'," & gTesGlobal.BancoConsec & "," & txtCantidadSolicitudes _
           & "," & mSolInicio & "," & mSolCorte & "," & mFechaInicio & "," & mFechaCorte

    Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
        Print #1, rs!Linea2 'Debitos
        rs.MoveNext
     Loop
   rs.Close
   
    'CREDITOS
    strSQL = "exec spTES_BCR_Empresarial 3," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
           & "'," & gTesGlobal.BancoConsec & "," & txtCantidadSolicitudes _
           & "," & mSolInicio & "," & mSolCorte & "," & mFechaInicio & "," & mFechaCorte
    Call OpenRecordSet(rs, strSQL)
 
     Do While Not rs.EOF
        Print #1, rs!Linea3 'Creditos
        rs.MoveNext
     Loop
     rs.Close
     
     Close #1   ' Close file.
     
     Me.Hide
         
     frmTES_Transferencias.lblBanco = cbo.Text
     frmTES_Transferencias.lblArchivo = strArchivo
     frmTES_Transferencias.Show vbModal
     
     Me.Show
Exit Sub

vError:
       
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
   If vPaso Then
      gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "-")
   End If

End Sub


Private Sub sbTeBNCR_Sinpe()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite la Transferencia en formato Empresarial para el BCR. Genera archivo de
'               texto en la direccion "C:\NombreEmpresa\Banco\Fecha\ConsecutivoDeposito.txt"
'               y finalmente despliega el formulario de control de transferencias.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset, strArchivo As String
Dim strPath As String, strCadena As String, i As Integer
Dim vCuentaBanco As String, vFecha As Date, vPaso As Boolean
Dim strSQL As String, fn, vTesKeyCh As String
Dim strLinea As String

On Error GoTo vError

vPaso = False
vFecha = fxFechaServidor

' Linea 1 -> Numero de Negocio (Registrado en el Sistema del BCR, SCIC)
' Linea 2 -> Cedula de Registro
' Linea 3 -> Razon o Detalle del Pago
i = 1

fn = FreeFile

    
     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "+")
     
     gTesGlobal.BancoNombre = cbo.Text
     vPaso = True
     
     
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     strArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & ".tef"
     
     Open strArchivo For Output As #1

     'REGISTRO DE CONTROL
     i = 1
    

    'ENCABEZADO: LINEA 1
    strSQL = "exec spTES_BNCR_SINPE 1," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc _
           & "'," & gTesGlobal.BancoConsec & "," & txtCantidadSolicitudes _
           & "," & mSolInicio & "," & mSolCorte & "," & mFechaInicio & "," & mFechaCorte
    Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
        Print #1, rs!Linea1 'Encabezado
        rs.MoveNext
     Loop
     rs.Close
     
  
    'DEBITOS
    strSQL = "exec spTES_BNCR_SINPE 2," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc _
           & "'," & gTesGlobal.BancoConsec & "," & txtCantidadSolicitudes _
           & "," & mSolInicio & "," & mSolCorte & "," & mFechaInicio & "," & mFechaCorte
    Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
        Print #1, rs!Linea2 'Debitos
        rs.MoveNext
     Loop
     rs.Close
   
    'CREDITOS
    strSQL = "exec spTES_BNCR_SINPE 3," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc _
           & "'," & gTesGlobal.BancoConsec & "," & txtCantidadSolicitudes _
           & "," & mSolInicio & "," & mSolCorte & "," & mFechaInicio & "," & mFechaCorte
    Call OpenRecordSet(rs, strSQL)
 
     Do While Not rs.EOF
        Print #1, rs!Linea3 'Creditos
        rs.MoveNext
     Loop
     rs.Close
     
     Close #1   ' Close file.
     
     Me.Hide
         
     frmTES_Transferencias.lblBanco = cbo.Text
     frmTES_Transferencias.lblArchivo = strArchivo
     frmTES_Transferencias.Show vbModal
     
     Me.Show
Exit Sub

vError:
       
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   If vPaso Then
      gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "-")
   End If

End Sub

Private Sub sbTeFormatoEstandar(pFormato As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Generacion con Formatos Estandares de Transferencias Bancarias
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset, strArchivo As String
Dim strPath As String, strCadena As String, i As Integer
Dim vCuentaBanco As String, vFecha As Date, vPaso As Boolean
Dim vConArchivo As String, vCedulaReg As String, vNumNegocio As String, vRazon As String
Dim strSQL As String, fn, vTesKeyCh As String
Dim strCedJuridica As String
Dim iLineInicio As Integer 'variable para la linea con la que inicia el detalle de las transferencias
Dim strLinea As String

On Error GoTo vError



vPaso = False
vFecha = fxFechaServidor

i = 1

fn = FreeFile

vRazon = ""
vNumNegocio = ""
vCedulaReg = ""

strSQL = "select  REPLACE(cedula_juridica,'-','') as 'Cedula_Juridica',NOMBRE" _
       & " From SIF_EMPRESA"
Call OpenRecordSet(rs, strSQL, 0)
    vNumNegocio = Trim(rs!cedula_juridica)
    vCedulaReg = Trim(rs!cedula_juridica)
    vRazon = "TRANSFERENCIAS " & rs!Nombre
rs.Close


Dim vExtension As String, vProcedimiento As String

strSQL = "select Procedimiento,Extension from vTes_Formatos where cod_formato = '" & pFormato & "'"
Call OpenRecordSet(rs, strSQL, 0)
    vExtension = Trim(rs!Extension)
    vProcedimiento = Trim(rs!Procedimiento)
rs.Close


    
     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "+", cboPlan.ItemData(cboPlan.ListIndex))
     
     gTesGlobal.BancoPlan = cboPlan.ItemData(cboPlan.ListIndex)
     
     gTesGlobal.BancoNombre = cbo.Text
     vPaso = True
     i = 1
    strSQL = "select documento_base,count(*) From Tes_Transacciones" _
         & " where id_banco = " & gTesGlobal.BancoID & " and  convert(varchar,fecha_emision ,106) = '" _
         & Format(vFecha, "yyyy/mm/dd") & "' and estado = 'T' group by documento_base"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     i = i + 1
     rs.MoveNext
    Loop
    rs.Close
    vConArchivo = Format(i, "000")
     
     
     strSQL = "select dbo.fxTesCantidadTEDiarias('" & Format(vFecha, "yyyy/mm/dd") & "' ," & gTesGlobal.BancoID & ") as 'Cantidad'"
     Call OpenRecordSet(rs, strSQL)
         iLineInicio = rs!Cantidad
     rs.Close
     
     
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     strArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & "." & vExtension
     
     Open strArchivo For Output As #1

'     'REGISTRO DE CONTROL
'     i = 1
'
'     strCadena = "000"                                                              'Estado 3
'     strCadena = strCadena & SIFGlobal.fxStringRelleno(vCedulaReg, "I", "0", 12)    'Cedula Juridica 12
'     strCadena = strCadena & vConArchivo                                            'Consecutivo Archivo 3
'     strCadena = strCadena & Format(vFecha, "ddmmyyyy")                             'Fecha Aplicacion 8
'     strCadena = strCadena & "000000000000"                                         'Cedula de Registro 12
'     strCadena = strCadena & "000000000000"                                         '12 Filler con 0
'     strCadena = strCadena & "000000"                                               '6 Hora Estado Se rellena con ceros
'     strCadena = strCadena & SIFGlobal.fxStringRelleno("", "D", "0", 138)           '138 Filler con 0
'
'     Print #1, strCadena
     
     'LINEA CONTROL
     strSQL = "exec " & vProcedimiento & " 1," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
            & "'," & gTesGlobal.BancoConsec & "," & txtCantidadSolicitudes _
            & "," & mSolInicio & "," & mSolCorte & "," & mFechaInicio & "," & mFechaCorte
            
     If cboPlan.ItemData(cboPlan.ListIndex) <> "-sp-" Then
        strSQL = strSQL & ",'" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
     End If
            
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
        If rs!Linea1 <> "" Then
            Print #1, rs!Linea1 'Linea Control
        End If
        rs.MoveNext
     Loop
     rs.Close
  
  
  
     'DEBITOS
     strSQL = "exec " & vProcedimiento & " 2," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
            & "'," & gTesGlobal.BancoConsec & "," & txtCantidadSolicitudes _
            & "," & mSolInicio & "," & mSolCorte & "," & mFechaInicio & "," & mFechaCorte
     
     If cboPlan.ItemData(cboPlan.ListIndex) <> "-sp-" Then
        strSQL = strSQL & ",'" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
     End If
     
     Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
           If rs!Linea2 <> "" Then
                Print #1, rs!Linea2 'Debitos
           End If
           rs.MoveNext
        Loop
        rs.Close
     
     'CREDITOS
     strSQL = "exec " & vProcedimiento & " 3," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
            & "'," & gTesGlobal.BancoConsec & "," & txtCantidadSolicitudes _
            & "," & mSolInicio & "," & mSolCorte & "," & mFechaInicio & "," & mFechaCorte
     
     If cboPlan.ItemData(cboPlan.ListIndex) <> "-sp-" Then
        strSQL = strSQL & ",'" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
     End If
     
     Call OpenRecordSet(rs, strSQL)
 
     Do While Not rs.EOF
        If rs!Linea3 <> "" Then
            Print #1, rs!Linea3 'Creditos
        End If
        rs.MoveNext
     Loop
     rs.Close
     
     Close #1   ' Close file.
     
     Me.Hide
         
     frmTES_Transferencias.lblBanco = cbo.Text
     frmTES_Transferencias.lblArchivo = strArchivo
     frmTES_Transferencias.Show vbModal
     
     Me.Show
     
     
     
Exit Sub

vError:
       
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
   If vPaso Then
      gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "-", gTesGlobal.BancoPlan)
   End If

End Sub





''Procedimiento para crear el nuevo archivo del BCR, Banca Empresarial
Private Sub sbTeBCR_Comercial()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite la Transferencia en formato Empresarial para el BCR. Genera archivo de
'               texto en la direccion "C:\NombreEmpresa\Banco\Fecha\ConsecutivoDeposito.txt"
'               y finalmente despliega el formulario de control de transferencias.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset, strArchivo As String
Dim strPath As String, strCadena As String, i As Integer
Dim vCuentaBanco As String, vFecha As Date, vPaso As Boolean
Dim vConArchivo As String, vCedulaReg As String, vNumNegocio As String, vRazon As String
Dim strSQL As String, fn, vTesKeyCh As String
Dim strCedJuridica As String
Dim iLineInicio As Integer 'variable para la linea con la que inicia el detalle de las transferencias
Dim strLinea As String

On Error GoTo vError



vPaso = False
vFecha = fxFechaServidor

' Linea 1 -> Numero de Negocio (Registrado en el Sistema del BCR, SCIC)
' Linea 2 -> Cedula de Registro
' Linea 3 -> Razon o Detalle del Pago
i = 1

fn = FreeFile

vRazon = ""
vNumNegocio = ""
vCedulaReg = ""

strSQL = "select  REPLACE(cedula_juridica,'-','') as 'Cedula_Juridica',NOMBRE" _
       & " From SIF_EMPRESA"
Call OpenRecordSet(rs, strSQL, 0)
    vNumNegocio = Trim(rs!cedula_juridica)
    vCedulaReg = Trim(rs!cedula_juridica)
    vRazon = "TRANSFERENCIAS " & rs!Nombre
rs.Close

    
     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "+")
     
     gTesGlobal.BancoNombre = cbo.Text
     vPaso = True
     i = 1
    strSQL = "select documento_base,count(*) From Tes_Transacciones" _
         & " where id_banco = " & gTesGlobal.BancoID & " and fecha_emision = '" _
         & Format(vFecha, "yyyy/mm/dd") & "' and estado = 'T' group by documento_base"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     i = i + 1
     rs.MoveNext
    Loop
    rs.Close
    vConArchivo = Format(i, "000")
     
     
     strSQL = "select dbo.fxTesCantidadTEDiarias('" & Format(vFecha, "yyyy/mm/dd") & "' ," & gTesGlobal.BancoID & ") as 'Cantidad'"
     Call OpenRecordSet(rs, strSQL)
         iLineInicio = rs!Cantidad
     rs.Close
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     strArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & ".txt"
     
     Open strArchivo For Output As #1

     'REGISTRO DE CONTROL
     i = 1
    
     strCadena = "000"                                                              'Estado 3
     strCadena = strCadena & SIFGlobal.fxStringRelleno(vCedulaReg, "I", "0", 12)    'Cedula Juridica 12
     strCadena = strCadena & vConArchivo                                            'Consecutivo Archivo 3
     strCadena = strCadena & Format(vFecha, "ddmmyyyy")                             'Fecha Aplicacion 8
     strCadena = strCadena & "000000000000"                                         'Cedula de Registro 12
     strCadena = strCadena & "000000000000"                                         '12 Filler con 0
     strCadena = strCadena & "000000"                                               '6 Hora Estado Se rellena con ceros
     strCadena = strCadena & SIFGlobal.fxStringRelleno("", "D", "0", 138)           '138 Filler con 0
   
     Print #1, strCadena
     
     
    'DEBITOS
    strSQL = "exec spTES_BCR_Comercial 2," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
           & "'," & gTesGlobal.BancoConsec & "," & txtCantidadSolicitudes _
           & "," & mSolInicio & "," & mSolCorte & "," & mFechaInicio & "," & mFechaCorte
    Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
        Print #1, rs!Linea2 'Debitos
        rs.MoveNext
     Loop
   rs.Close
   
    'CREDITOS
    strSQL = "exec spTES_BCR_Comercial 3," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
           & "'," & gTesGlobal.BancoConsec & "," & txtCantidadSolicitudes _
           & "," & mSolInicio & "," & mSolCorte & "," & mFechaInicio & "," & mFechaCorte
    Call OpenRecordSet(rs, strSQL)
 
     Do While Not rs.EOF
        Print #1, rs!Linea3 'Creditos
        rs.MoveNext
     Loop
     rs.Close
     
     Close #1   ' Close file.
     
     Me.Hide
         
     frmTES_Transferencias.lblBanco = cbo.Text
     frmTES_Transferencias.lblArchivo = strArchivo
     frmTES_Transferencias.Show vbModal
     
     Me.Show
Exit Sub

vError:
       
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
   If vPaso Then
      gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "-")
   End If

End Sub


Private Sub sbTeBCR_Planilla(strSQL As String, Optional vTestKey As Long = 0, Optional vMontoTotal As Currency = 0)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite la Transferencia en formato para el Banco Nacional. Genera archivo de
'               texto en la direccion "C:\NombreEmpresa\Banco\Fecha\ConsecutivoDeposito.txt"
'               y finalmente despliega el formulario de control de transferencias.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset, strArchivo As String
Dim strPath As String, strCadena As String, i As Integer
Dim vCuentaBanco As String, vFecha As Date, vPaso As Boolean
Dim vConArchivo As String, vCedulaReg As String, vNumNegocio As String, vRazon As String
Dim ySQL As String, fn, vTesKeyCh As String

On Error GoTo vError

'Se Asume lo siguiente:
' 1. Todas las cuentas son: xxx,xxxxxxx,x
' Ejemplo : 91300127060
' Del ejemplo anterior: 913 es el numero de oficina donde se realizo la apertura
'                   0012706 es el numero de cuenta de la persona
'                         0 es el digito verificador
'
' 2. Dado lo anterior el sistema procesará todo el numero completo, y realizará
'    los desgloces necesarios, por lo que la información que se suministre tendrá
'    que estar completa

gstrQuery = strSQL

vPaso = False
vFecha = fxFechaServidor

'Leer Archivo de Texto : BCRFormat.ini
' Linea 1 -> Numero de Negocio (Registrado en el Sistema del BCR, SCIC)
' Linea 2 -> Cedula de Registro
' Linea 3 -> Razon o Detalle del Pago
i = 1

fn = FreeFile

vRazon = ""
vNumNegocio = ""
vCedulaReg = ""

strArchivo = SIFGlobal.DirectorioDeResultados & "\Configuracion\BCRFormat.ini"
If Dir(strArchivo, vbArchive) = "" Then
  strArchivo = App.Path & "\BCRFormat.ini"
End If

Open strArchivo For Input As #fn
 Do While Not EOF(fn)
   Input #fn, ySQL
   Select Case i
     Case 1
       vNumNegocio = ySQL
     Case 2
       vCedulaReg = ySQL
     Case 3
       vRazon = ySQL
   End Select
   i = i + 1
 Loop
Close #fn   ' Close file.


For i = Len(vRazon) To 30
  vRazon = vRazon & " "
Next i

'Calcular el Numero de Archivo , Numero de la Transferencia en el Dia
i = 1
ySQL = "select documento_base,count(*) From Tes_Transacciones" _
     & " where id_banco = " & cbo.ItemData(cbo.ListIndex) & " and fecha_emision = '" _
     & Format(vFecha, "yyyy/mm/dd") & "' and estado = 'T' group by documento_base"
rs.Open ySQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 i = i + 1
 rs.MoveNext
Loop
rs.Close
vConArchivo = Format(i, "000")
    
    
'Crear y Sacar la cuenta de Tes_Bancos, se Asume que esta cuenta tiene el digito verificador
ySQL = "select Cta from Tes_Bancos where id_Banco = " & cbo.ItemData(cbo.ListIndex)
rs.Open ySQL, glogon.Conection, adOpenStatic
 'Se indica la oficina 001 de apertura por Omision
 vCuentaBanco = "001" & Format(Trim(rs!Cta), "00000000")
rs.Close
    
Dim xTesKey As Long
    
'Calcular TestKey Complementario (de la primera Linea)
ySQL = "select dbo.fxTESBCRTestkey('" & vCuentaBanco & "'," & vMontoTotal & ") as TestKey"
rs.Open ySQL, glogon.Conection, adOpenStatic
            
If vTestKey + rs!TestKey > 2147483468 Then
        vTestKey = 2147483468
Else
        vTestKey = vTestKey + rs!TestKey
End If
rs.Close

'Validando Largo del TestKey  = 12
vTesKeyCh = Trim(CStr(vTestKey))
If Len(vTesKeyCh) > 12 Then
  vTestKey = Right(vTesKeyCh, 12)
End If
    
    
     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "+")
     gTesGlobal.BancoNombre = cbo.Text
     vPaso = True
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     strArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & ".BCR"
     
     Open strArchivo For Output As #1

     '*****************************************
     'ENCABEZADO DEL FORMATO DE TRANSFERENCIA *
     '*****************************************

     strCadena = "000" 'Estado
     strCadena = strCadena & vNumNegocio            '12 char
     strCadena = strCadena & vConArchivo            '3 char
     strCadena = strCadena & "000000"               '6 Filler
     strCadena = strCadena & vCedulaReg             '12 char
     strCadena = strCadena & Format(vTestKey, "000000000000") '12 TestKey ** Generarlo **
     strCadena = strCadena & "000000"               '6 Hora
     strCadena = strCadena & Format(Day(vFecha), "00") & Format(Month(vFecha), "00") & Format(Year(vFecha), "0000")
     strCadena = strCadena & Space(21)              'filler 21 char
     strCadena = strCadena & "Y"                    'Señal de Y2k
    
     Print #1, strCadena
     
     
     
     '******************************
     ' DETALLE DE LA TRANSFERENCIA *
     '******************************
     
     'Linea 1 es la de Debito cuenta Bancaria
     
     i = 1
          
        strCadena = "000"                       'Estado Relleno con Ceros
        strCadena = strCadena & "1"             'Concepto 1 = Cuenta Corriente / 2 Cuenta Ahorro
        strCadena = strCadena & "00000"         'Filler 5
        strCadena = strCadena & Mid(Trim(vCuentaBanco), 1, 11) 'Oficina -> 3c, Cuenta -> 7 + 1 Digito verificador
        strCadena = strCadena & "1"             'Moneda  1 = Colones, 2 = Dolares
        strCadena = strCadena & "4"             '2 -> Credito, 4 -> Debito
        strCadena = strCadena & "0000"          'Codigo de Causa
        strCadena = strCadena & Format(gTesGlobal.BancoConsec, "0000") & Format(i, "0000") 'Numero de Documento 8
        strCadena = strCadena & Format((vMontoTotal * 100), "000000000000") '12 Sin Decimales
        strCadena = strCadena & Format(Day(vFecha), "00") & Format(Month(vFecha), "00") & Format(Year(vFecha), "0000")
        strCadena = strCadena & "0"             'Filler 1
        strCadena = strCadena & vRazon          'Razon de Transferencia (Detalle) 30
     
        Print #1, strCadena
     
     
     Call OpenRecordSet(rs, strSQL)
     
     
     Do While Not rs.EOF
        i = i + 1
      
        strCadena = "000"                       'Estado Relleno con Ceros
        strCadena = strCadena & "2"             'Concepto 1 = Cuenta Corriente / 2 Cuenta Ahorro
        strCadena = strCadena & "00000"         'Filler 5
        strCadena = strCadena & Mid(Trim(rs!Cta_Ahorros), 1, 11) 'Oficina -> 3c, Cuenta -> 7 + 1 Digito verificador
        strCadena = strCadena & "1"             'Moneda  1 = Colones, 2 = Dolares
        strCadena = strCadena & "2"             '2 -> Credito, 4 -> Debito
        strCadena = strCadena & "0000"          'Codigo de Causa
        strCadena = strCadena & Format(gTesGlobal.BancoConsec, "0000") & Format(i, "0000") 'Numero de Documento 8
        strCadena = strCadena & Format((rs!Monto * 100), "000000000000") '12 Sin Decimales
        strCadena = strCadena & Format(Day(vFecha), "00") & Format(Month(vFecha), "00") & Format(Year(vFecha), "0000")
        strCadena = strCadena & "0"             'Filler 1
        strCadena = strCadena & vRazon          'Razon de Transferencia (Detalle) 30
        
        Print #1, strCadena
        
        rs.MoveNext
     Loop
     rs.Close
     
     Close #1   ' Close file.
     
     Me.Hide
         
     frmTES_Transferencias.lblBanco = cbo.Text
     frmTES_Transferencias.lblArchivo = strArchivo
     frmTES_Transferencias.Show vbModal
     
     Me.Show
Exit Sub

vError:
       
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   Resume
   
   If vPaso Then
      gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "-")
   End If
   
End Sub



Private Sub sbTeBancoNacional(strSQL As String, curPlanilla As Currency)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite la Transferencia en formato para el Banco Nacional. Genera archivo de
'               texto en la direccion "C:\NombreEmpresa\Banco\Fecha\ConsecutivoDeposito.txt"
'               y finalmente despliega el formulario de control de transferencias.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset
Dim strArchivo As String, strCedula As String, strMonto As String
Dim strPath As String, strCadena As String, i As Integer
Dim curMonto1 As Currency, curMonto2 As Currency
Dim curCuentas As Currency, vFecha As Date, vPaso As Boolean
Dim vConcepto As String, vCuentaEmpresa As String, vCuenta As String


Dim vNumCliente As String, pSQL As String


On Error GoTo vError

'Se Asume lo siguiente:
' 1. Todas las cuentas son: x[1,2]00,01,XXX,XXXXXX,X
' Ejemplo : 100010318052919
' Del ejemplo anterior: 100 es el numero de cuenta corriente, puede ser tambien 200 que es de ahorros
'                        01 es el tipo de moneda (solo colones se aceptaran)
'                       031 es el numero de la apertura de la cuenta (Sucursal)
'                    805291 es el numero de cuenta de la persona
'                         9 es el digito verificador
'
' 2. Dado lo anterior el sistema procesará todo el numero completo, y realizará
'    los desgloces necesarios, por lo que la información que se suministre tendrá
'    que estar completa

gstrQuery = strSQL

vPaso = False
vFecha = Format(fxFechaServidor, "dd/mm/yyyy")
curMonto1 = IIf(IsNull(curPlanilla), 0, curPlanilla)
strMonto = CStr(Format(IIf(IsNull(curPlanilla), 0, curPlanilla), "0000000000.00"))
strMonto = Mid(strMonto, 1, Len(strMonto) - 3) & Mid(strMonto, Len(strMonto) - 1, 2)

vConcepto = SIFGlobal.fxStringRelleno("TF " & gPortal.Empresa_Name, "I", " ", 30)

pSQL = "select Cta,codigo_Cliente from tes_Bancos" _
       & " Where id_Banco = " & cbo.ItemData(cbo.ListIndex)
Call OpenRecordSet(rs, pSQL)
 vCuentaEmpresa = fxDepuraString(rs!Cta, "-")
 vNumCliente = Trim(rs!Codigo_Cliente & "")
rs.Close

vNumCliente = SIFGlobal.fxStringRelleno(vNumCliente, "I", "0", 6)


     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Bancos y Consecutivo
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "+")
     gTesGlobal.BancoNombre = cbo.Text
     vPaso = True
     
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     Open strArchivo & "\" & gTesGlobal.BancoConsec & ".ENV" For Output As #1

     '*****************************************
     'ENCABEZADO DEL FORMATO DE TRANSFERENCIA *
     '*****************************************

     strCadena = "1"
     strCadena = strCadena & vNumCliente
     strCadena = strCadena & Format(Day(vFecha), "00") & Format(Month(vFecha), "00") & Format(Year(vFecha), "0000")
     strCadena = strCadena & Format(gTesGlobal.BancoID, "000000000000")
     strCadena = strCadena & "1" & "0000"
     strCadena = strCadena & strMonto
     strCadena = strCadena & "000000000000000000000000"
     
     Print #1, strCadena
     
     
     
     '******************************
     ' DETALLE DE LA TRANSFERENCIA *
     '******************************
     
     Call OpenRecordSet(rs, strSQL)
     
     i = 0
     
     Do While Not rs.EOF
        i = i + 1
        
    ' x[1,2]00,01,XXX,XXXXXX,X
        vCuenta = Replace(Trim(rs!Cta_Ahorros), "-", "")
        
        strCadena = "3" 'Credito
        strCadena = strCadena & Mid(vCuenta, 6, 3) & Mid(vCuenta, 1, 3) & "01"  '  "20001"        '  "000" 'Oficina de Apertura
        strCadena = strCadena & Right(vCuenta, 7) ' & Mid(Trim(rs!Cta_Ahorros), 9, 7)      ' trim(RS!cta_ahorros)
                                'Incluye: 100 o 200 -> Tipo de Cuenta (Corriente-Ahorros)
                                '          01 -> Tipo Moneda (Colones)
                                '      000000 -> Cuenta de la Persona
                                '           0 -> Digito Verificador

        'Suma las Cuentas para Registro de Totales
        ' - Solo tiene que la cuenta de la persona
            
            curCuentas = curCuentas + CCur(Mid(Right(Trim(vCuenta), 7), 1, 6))   'Sin Verificador
            curMonto2 = curMonto2 + rs!Monto
        
'''        strCadena = "3" 'Credito
'''        strCadena = strCadena & Mid(Trim(rs!Cta_Ahorros), 1, 3) & "20001"     '  "000" 'Oficina de Apertura
'''        strCadena = strCadena & Right(Trim(rs!Cta_Ahorros), 7) ' & Mid(Trim(rs!Cta_Ahorros), 9, 7)      ' trim(RS!cta_ahorros)
'''                                'Incluye: 100 o 200 -> Tipo de Cuenta (Corriente-Ahorros)
'''                                '          01 -> Tipo Moneda (Colones)
'''                                '      000000 -> Cuenta de la Persona
'''                                '           0 -> Digito Verificador
'''
'''        'Suma las Cuentas para Registro de Totales
'''        ' - Solo tiene que la cuenta de la persona
'''
'''            curCuentas = curCuentas + CCur(Mid(Right(Trim(rs!Cta_Ahorros), 7), 1, 6))   'Sin Verificador
'''            curMonto2 = curMonto2 + rs!Monto
        
        'Fin del Calculo de las cuentas y del Monto de acreditaciones
        
        strCadena = strCadena & Format(i, "00000000") '8d Numero Comprobante (Consecutivo Interno)
                
        strMonto = CStr(Format(rs!Monto, "0000000000.00")) '12d Monto sin el punto decimal
        strMonto = Mid(strMonto, 1, Len(strMonto) - 3) & Mid(strMonto, Len(strMonto) - 1, 2)
        
        strCadena = strCadena & strMonto                       'corte
        strCadena = strCadena & SIFGlobal.fxStringRelleno(vConcepto, "D", " ", 30)  '30d Concepto de Pago
        strCadena = strCadena & "00" 'Fin de Linea
                    
        Print #1, strCadena
        
        rs.MoveNext
     Loop
     rs.Close
     
     '*********************************************************
     'CREA ULTIMA LINEA DE DETALLE CON EL DEBITO A LA EMPRESA *
     '*********************************************************
     strCadena = "2" & Mid(Trim(vCuentaEmpresa), 1, 3)   'Movimiento de Debito, y 000 Sucursal de Apertura
     strCadena = strCadena & "10001" 'Cuenta Corriente y Moneda en Colones
     strCadena = strCadena & Right(Trim(vCuentaEmpresa), 7) 'Cuenta de la Empresa  + Digito Verificador
     strCadena = strCadena & Format(i + 1, "00000000") 'Numero Comprobante
        
      strMonto = CStr(Format(curMonto2, "0000000000.00")) '12d Monto sin el punto decimal
      strMonto = Mid(strMonto, 1, Len(strMonto) - 3) & Mid(strMonto, Len(strMonto) - 1, 2)
        
     strCadena = strCadena & strMonto 'Total de los Creditos para Debitar a esta cuenta
     strCadena = strCadena & vConcepto '30d Concepto de Pago
     strCadena = strCadena & "00" 'Fin de Linea
    
     Print #1, strCadena
     curCuentas = curCuentas + CCur(Mid(Right(Trim(vCuentaEmpresa), 7), 1, 6))   'Sin Verificador

     '**************************************************
     'REGISTRO DE CONTROL DEL ARCHIVO DE TRANSFERENCIA *
     '**************************************************
     
     strCadena = "4" 'Codigo de Control de registro
     strMonto = CStr(Format(curMonto1 + curMonto2, "0000000000000.00")) 'Suma Debitos y Creditos de la Transferencia
     strMonto = Mid(strMonto, 1, Len(strMonto) - 3) & Mid(strMonto, Len(strMonto) - 1, 2)
     strCadena = strCadena & strMonto
     
     strMonto = CStr(Format(curCuentas, "0000000000")) 'Sumatoria de Cuentas
     strCadena = strCadena & strMonto
     strCadena = strCadena & "0000000000"
     strCadena = strCadena & "000000000000"
     strCadena = strCadena & "000000000000"
     strCadena = strCadena & "00000000"
     
     Print #1, strCadena
     
     Close #1   ' Close file.
     
     Me.Hide
         
     frmTES_Transferencias.lblBanco = cbo.Text
     frmTES_Transferencias.lblArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & ".ENV"
     frmTES_Transferencias.Show vbModal
     
Exit Sub

vError:
       
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

   If vPaso Then
      gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "-")
   End If
   
End Sub


Private Sub sbTeBancoPopular(strSQL As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite la Transferencia en formato para el Banco Popular. Genera archivo de
'               texto en la direccion "C:\NombreEmpresa\Banco\Fecha\ConsecutivoDeposito.txt"
'               y finalmente despliega el formulario de control de transferencias.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim recCheques As New ADODB.Recordset
Dim strArchivo As String, strNombre As String
Dim strMonto As String, strCuenta As String
Dim strPath As String, strCadena As String
Dim strSelf As String, strProducto As String
Dim strEstado As String, strTipo As String
Dim strFecha As String, intI As Integer
Dim vFecha As Date, vPaso As Boolean


On Error GoTo vError

vFecha = Format(fxFechaServidor, "dd/mm/yyyy")

'Cada Global, para ser utilizada en el modulo de Ejecucion de la Transferencia
gstrQuery = strSQL
vPaso = False

With recCheques
     .Open strSQL, glogon.Conection, adOpenStatic
   
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir Trim(strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir Trim(strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir Trim(strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir Trim(strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir Trim(strArchivo)
        Else
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(Trim(strArchivo), vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir Trim(strArchivo)
           End If
        End If
     End If
     
     ChDir Trim(strArchivo)
          
     'Inicializa Variables Globales y Consecutivos
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "+")
     gTesGlobal.BancoNombre = cbo.Text
     vPaso = True
     
     Open strArchivo & "\" & gTesGlobal.BancoConsec & ".txt" For Output As #1
     
     Do While Not .EOF
        
        
        Select Case Len(Trim(!Codigo))
           Case 8
               strCadena = "0" & Mid(Trim(!Codigo), 1, 1) & "0" & Mid(Trim(!Codigo), 2, 7)
           Case 9
               strCadena = "0" & Trim(!Codigo)
           Case Is < 8
               strCadena = Format(!Codigo, "0000000000")
           Case Is > 10
               strCadena = Mid(Trim(!Codigo), 1, 4) & "0" & Mid(Trim(!Codigo), 6, 5)
           Case Else
               strCadena = Trim(!Codigo)
        End Select
        
        strNombre = Trim(!Beneficiario)
                
        If Len(strNombre) > 30 Then
         strNombre = Mid(strNombre, 1, 30)
        Else
         Do Until Len(strNombre) = 30
           strNombre = strNombre & " "
         Loop
        End If
        
        strCadena = strCadena & strNombre
        
        strCuenta = IIf(IsNull(!Cta_Ahorros), "0", Trim(!Cta_Ahorros))
        
        If Len(strCuenta) > 13 Then
           strCuenta = Mid(strCuenta, 1, 13)
        Else
         Do Until Len(strCuenta) = 13
            strCuenta = "0" & strCuenta
         Loop
        End If
        
        strCadena = strCadena & strCuenta
        
        strSelf = " "
        strCadena = strCadena & strSelf
        
        strMonto = CStr(Format(!Monto, "000000000.00"))
        strMonto = Mid(strMonto, 1, Len(strMonto) - 3) & Mid(strMonto, Len(strMonto) - 1, 2)
        
        strCadena = strCadena & strMonto
                
        strFecha = Format(Day(vFecha), "00")
        strFecha = strFecha & Format(Month(vFecha), "00")
        strFecha = strFecha & Format(Year(vFecha), "0000")
        
        strCadena = strCadena & strFecha
        
        strTipo = "A"
        strCadena = strCadena & strTipo
        
        strProducto = "06"
        strCadena = strCadena & strProducto
        
        strEstado = "P"
        strCadena = strCadena & strEstado
        
        strCadena = strCadena & strFecha
        strCadena = strCadena & strMonto
                
        Print #1, strCadena
 
        .MoveNext
     Loop
     Close #1   ' Close file.
     .Close
         
         
     Me.Hide
         
     frmTES_Transferencias.lblBanco = cbo.Text
     frmTES_Transferencias.lblArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & ".txt"
     frmTES_Transferencias.Show vbModal
     

End With
     
Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   If vPaso Then
      gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "-")
   End If
     
     
End Sub



Private Sub sbBuscaDocumentos()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select isnull(count(*),0) as Total,isnull(Min(nsolicitud),0) as Minimo" _
       & ",isnull(Max(nsolicitud),0) as Maximo from Tes_Transacciones" _
       & " Where Estado='P' And Tipo='" & cboDoc.ItemData(cboDoc.ListIndex) _
       & "' And ID_Banco = " & cbo.ItemData(cbo.ListIndex)
Call OpenRecordSet(rs, strSQL)
  
If rs!Total > 0 Then
   txtCantidadSolicitudes = rs!Total
   txtGeneraNumeroDe = rs!Minimo
   txtGeneraNumeroHasta = rs!Maximo
Else
   txtCantidadSolicitudes = 0
   txtGeneraNumeroDe = ""
   txtGeneraNumeroHasta = ""
End If
rs.Close

txtDocInicial = fxTesTipoDocConsec(cbo.ItemData(cbo.ListIndex), cboDoc.ItemData(cboDoc.ListIndex), "/", cboPlan.ItemData(cboPlan.ListIndex))

If Trim(fxTesTipoDocExtraeDato(cbo.ItemData(cbo.ListIndex), cboDoc.ItemData(cboDoc.ListIndex), "mod_consec")) = "1" Then
   txtDocInicial.Locked = False
Else
   txtDocInicial.Locked = True
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbEmitirDocumentos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vConsecutivo As String, vVerificacion As Integer
Dim vComprobante As String, vAutoConsec As Boolean, itmX As ListViewItem
Dim lngConteo As Long, i As Integer, x As New clsImpresoras
Dim vBanco As Integer, vTipo As String, vFecha As Date
Dim vFirmaDesde As Currency, vFirmaHasta As Currency, vFirmas As Boolean, vFirmaAutorizada As Boolean
Dim strDec As String, curMonto As Currency, vFormatoTe As String
Dim vSQLx As String, vLugarEmision As String

lsw.ListItems.Clear

vBanco = cbo.ItemData(cbo.ListIndex)
strChequesFirmas = ""
strChequesSinFirmas = ""
        
Call sbCargaArchivosEspeciales(vBanco)

     
mSolInicio = "Null"
mSolCorte = "Null"
mFechaInicio = "Null"
mFechaCorte = "Null"

Select Case True
   Case optGeneraPor(0).Value
        strSQL = strSQL & " And NSolicitud Between " & Trim(txtGeneraNumeroDe) & " And " & Trim(txtGeneraNumeroHasta)
        mSolInicio = Trim(txtGeneraNumeroDe)
        mSolCorte = Trim(txtGeneraNumeroHasta)
   Case optGeneraPor(1).Value
        mFechaInicio = "'" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'"
        mFechaCorte = "'" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End Select


vTipo = cboDoc.ItemData(cboDoc.ListIndex)
vFecha = fxFechaServidor
vConsecutivo = 0
vFirmas = False

vFormatoTe = cboFormato.ItemData(cboFormato.ListIndex)


strSQL = "select doc_auto,comprobante from tes_banco_docs" _
       & " where id_banco = " & vBanco _
       & " and tipo = '" & vTipo & "'"
Call OpenRecordSet(rs, strSQL)
  vComprobante = rs!comprobante
  vAutoConsec = IIf((rs!doc_auto = 1), True, False)
rs.Close

strSQL = "select firmas_desde,firmas_hasta,formato_transferencia,Lugar_Emision  from Tes_Bancos where id_banco = " & vBanco
Call OpenRecordSet(rs, strSQL)
    vFirmaDesde = rs!firmas_desde
    vFirmaHasta = rs!firmas_hasta
    vLugarEmision = Trim(rs!Lugar_Emision & "")
rs.Close

strSQL = "select isnull(count(*),0) as Existe from TES_BANCO_FIRMASAUT where id_Banco = " & vBanco _
       & " and usuario = '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
    vFirmas = IIf((rs!Existe = 0), False, True)
rs.Close


strSQL = "Select TOP " & txtCantidadSolicitudes & " * From Tes_Transacciones Where Estado = 'P' And Tipo = '" _
       & vTipo & "' And ID_Banco=" & vBanco & " And Autoriza='S' and fecha_hold is null"
Select Case True
   Case optGeneraPor(0).Value
        strSQL = strSQL & " And NSolicitud Between " & Trim(txtGeneraNumeroDe) _
               & " And " & Trim(txtGeneraNumeroHasta)
   Case optGeneraPor(1).Value
        strSQL = strSQL & " And Fecha_Solicitud Between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' And '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End Select
strSQL = strSQL & " Order by Nsolicitud"


Select Case vComprobante
  Case "01", "02", "03" 'CK formula continua /CK Bloque / Registro Doc
  
        If vAutoConsec Then
            'Revisa que el Consecutivo, Sea Modificable o No, si lo es inicializar por el indicado por el usuario
            If txtDocInicial.Locked Then
                vConsecutivo = fxTesTipoDocConsec(vBanco, vTipo, "/")
            Else
               If vConsecutivo = 0 Then
                  vConsecutivo = txtDocInicial
                  vSQLx = "update tes_banco_docs set consecutivo = " & vConsecutivo _
                         & " where id_banco = " & vBanco _
                         & " and tipo = '" & vTipo & "'"
                  Call ConectionExecute(vSQLx)
               Else
                  vConsecutivo = fxTesTipoDocConsec(vBanco, vTipo, "+")
               End If
            End If
            
        End If
        
        lngConteo = 0
        
        lsw.ListItems.Clear
        
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          
          lsw.Tag = lsw.ListItems.Count 'Ultimo Item
          
          For i = 1 To CInt(txtVerificacion)
        
                'Indica que el documento esta autorizado para que se utilice firma electronica
                If IsNull(rs!FIRMAS_AUTORIZA_FECHA) Then
                    vFirmaAutorizada = False
                Else
                    vFirmaAutorizada = True
                End If
                
                strSQL = "Update Tes_Transacciones Set Estado='I',Fecha_Emision='" & Format(vFecha, "yyyy/mm/dd") _
                       & "',Ubicacion_Actual='T',FECHA_TRASLADO='" & Format(vFecha, "yyyy/mm/dd") & "',User_Genera = '" _
                       & glogon.Usuario & "'"
                If vAutoConsec Then
                   strSQL = strSQL & ",NDocumento='" & vConsecutivo & "'"
                End If
                
                strSQL = strSQL & " where NSolicitud=" & rs!NSolicitud
                Call ConectionExecute(strSQL)
                
                Call sbTesBancosAfectacion(rs!NSolicitud, "E")
                Call sbTesBitacoraEspecial(rs!NSolicitud, "10", "")
                Call Bitacora("Genera", "Genero Solicitud  " & rs!NSolicitud)
                                    
                'Actualiza Cuentas Corrientes
                Call sbTESActualizaCC(rs!Codigo, rs!Tipo, CStr(vConsecutivo), rs!Id_Banco, IIf(IsNull(rs!Op), 0, rs!Op), rs!Modulo, rs!submodulo, IIf(IsNull(rs!Referencia), 0, rs!Referencia))
        
                Set itmX = lsw.ListItems.Add(, , rs!NSolicitud)
                    itmX.SubItems(1) = rs!Beneficiario
                    If vAutoConsec Then
                        itmX.SubItems(2) = vConsecutivo
                    Else
                        itmX.SubItems(2) = rs!nDocumento & ""
                    End If
                    itmX.SubItems(3) = rs!Monto
                    itmX.SubItems(4) = vFecha
        
                If vAutoConsec Then
                   vConsecutivo = fxTesTipoDocConsec(vBanco, vTipo, "+")
                End If
               
               'Imprime Comprobante
               Select Case vComprobante
                  Case "01" 'Cheques Formula Continua
                  
                    With frmContenedor.Crt
                        .Reset
                        
                        x.TipoImpresora = Cheques
                        x.Reset
     
                        .PrinterDriver = x.Controlador
                        .PrinterName = x.Nombre
                        .PrinterPort = x.Puerto
                        
                        .Connect = glogon.ConectRPT
                        
                        If vLugarEmision <> "" Then
                           vLugarEmision = vLugarEmision & ", "
                        End If
                        
                        .Formulas(0) = "Fecha='" & vLugarEmision & Day(vFecha) & " DE " & fxTesMesDescripcion(vFecha) & " DE " & Year(vFecha) & "'"
                        .Formulas(1) = "Año='" & Year(vFecha) & "'"
                        
                        '*******Codigo Nuevo para Monto en Letras 2003/03/21
                        strDec = Format(rs!Monto, "##################.00")
                        strDec = Trim(strDec)
                        strDec = Mid(strDec, Len(strDec) - 1, 2)
                        
                        curMonto = Mid(Format(rs!Monto, "#################0.00"), 1, Len(Format(rs!Monto, "#################0.00")) - 3)
                        .Formulas(2) = "Letras='**" & Trim(UCase(Conversion(CStr(curMonto))))
                        
                        If Trim(strDec) <> "00" Then
                           .Formulas(2) = .Formulas(2) & UCase(" Con " & Trim(strDec) & "/100 " & fxDescDivisa(rs!cod_Divisa) & "**'")
                        Else
                           .Formulas(2) = .Formulas(2) & " " & UCase(fxDescDivisa(rs!cod_Divisa)) & "**'"
                        End If
                        '********** Fin de la Modificacion del Monto en Letras
                                                
                                                
                        'Si utiliza firmas, preguntar por el rango en montos
                        If vFirmas Then
                            If (rs!Monto >= vFirmaDesde And rs!Monto <= vFirmaHasta) Or vFirmaAutorizada Then
                                 .ReportFileName = SIFGlobal.fxPathReportes(strChequesFirmas) 'Reporte con Firmas
                            Else
                                If vFirmaAutorizada Then
                                 .ReportFileName = SIFGlobal.fxPathReportes(strChequesFirmas) 'Reporte con Firmas
                                Else
                                   .ReportFileName = SIFGlobal.fxPathReportes(strChequesSinFirmas) 'Reporte sin Firmas
                                End If
                            End If
                        Else
                                   .ReportFileName = SIFGlobal.fxPathReportes(strChequesSinFirmas) 'Reporte sin Firmas
                        End If
                        
                        .SelectionFormula = "{CHEQUES.NSOLICITUD}=" & rs!NSolicitud
                        .Destination = crptToPrinter
                        .PrintReport
               
                    End With
                  
                  Case "02", "03" 'Cheques Block / Boleta de Transaccion
                        With frmContenedor.Crt
                            .Reset
                            .Connect = glogon.ConectRPT
                            .ReportFileName = SIFGlobal.fxPathReportes("Banking_BoletaRegistro.rpt")
                            .SelectionFormula = "{CHEQUES.NSOLICITUD} = " & rs!NSolicitud
                            
                            
                            .SubreportToChange = "sbDetalle"
                        
                            .StoredProcParam(0) = rs!NSolicitud
                            
                            
                            .Destination = crptToPrinter
                            .PrintReport
                        
                        
                        End With
               End Select
            
              'Codigo de Salida de los dos Ciclos
              rs.MoveNext
              If rs.EOF Then
                Exit For
              End If
            Next i
        
            Me.Hide
            frmTES_Genera.Show vbModal 'Genera mensaje
            Me.Show
            
            DoEvents
            
            If Not gblnContinua Then Exit Do
        Loop
        rs.Close
        
        
  
  Case "04" 'Transferencias Electrónicas

     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
'     gTesGlobal.BancoID = vBanco
'     gTesGlobal.BancoTDoc = vTipo
'     gTesGlobal.BancoConsec = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "+")
'     gTesGlobal.BancoNombre = cbo.Text
    
    gTesGlobal.BancoPlan = cboPlan.ItemData(cboPlan.ListIndex)
    
    Select Case vFormatoTe
      Case "A" 'A - BNCR. Internet Banking
            
            vSQLx = strSQL
            
            strSQL = "(Select TOP " & txtCantidadSolicitudes & " nsolicitud From Tes_Transacciones Where Estado = 'P' And Tipo = '" _
                   & vTipo & "' And ID_Banco=" & vBanco & " And Autoriza='S' and fecha_hold is null"
            Select Case True
               Case optGeneraPor(0).Value
                    strSQL = strSQL & " And NSolicitud Between " & Trim(txtGeneraNumeroDe) _
                           & " And " & Trim(txtGeneraNumeroHasta)
               Case optGeneraPor(1).Value
                    strSQL = strSQL & " And Fecha_Solicitud Between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                           & " 00:00:00' And '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
            End Select
            strSQL = strSQL & " Order by Nsolicitud)"
            
            strSQL = "select sum(monto) as PLx from Tes_Transacciones where nsolicitud in" & strSQL
            Call OpenRecordSet(rs, strSQL)
             Call sbTeBancoNacional(vSQLx, rs!plx)
            rs.Close
           
      Case "B" 'B - Banco Popular
           Call sbTeBancoPopular(strSQL)
           
      Case "C" 'C - BCR. Planilla Empresarial
           
            vSQLx = strSQL
            
            strSQL = "(Select TOP " & txtCantidadSolicitudes & " nsolicitud From Tes_Transacciones Where Estado = 'P' And Tipo = '" _
                   & vTipo & "' And ID_Banco=" & vBanco & " And Autoriza='S' and fecha_hold is null"
            Select Case True
               Case optGeneraPor(0).Value
                    strSQL = strSQL & " And NSolicitud Between " & Trim(txtGeneraNumeroDe) _
                           & " And " & Trim(txtGeneraNumeroHasta)
               Case optGeneraPor(1).Value
                    strSQL = strSQL & " And Fecha_Solicitud Between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                           & " 00:00:00' And '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
            
            End Select
            strSQL = strSQL & " Order by Nsolicitud)"
            
            strSQL = "select sum(dbo.fxTESBCRTestkey(cta_ahorros,monto)) as TestKeyX, sum(Monto) as Monto" _
                   & " from Tes_Transacciones where nsolicitud in" & strSQL
                   
            Dim xTestKey As Long
            Call OpenRecordSet(rs, strSQL)
            
            If rs!TestKeyX > 2147483468 Then
                    xTestKey = 2147483468
            Else
                    xTestKey = rs!TestKeyX
            End If
             Call sbTeBCR_Planilla(vSQLx, xTestKey, rs!Monto)
            rs.Close
    
      Case "D" 'D - BCR. Empresas
            gstrQuery = strSQL
            Call sbTeBCR_Empresarial
    
      Case "E" 'E - BCT. Enlace
            gstrQuery = strSQL
            Call sbTeBCT_Enlace
    
      Case "F" 'F - BCR. Comercial
            gstrQuery = strSQL
            Call sbTeBCR_Comercial
    
      Case "G" 'G - BN Formato SINPE
            gstrQuery = strSQL
            Call sbTeBNCR_Sinpe
      
      
      Case "DV1", "DV2"
            gstrQuery = strSQL
            Call sbTeFormatoEstandar(vFormatoTe)
      
      Case "S"
    
    
      Case Else
            gstrQuery = strSQL
            Call sbTeFormatoEstandar(vFormatoTe)
    End Select

End Select

Call sbBuscaDocumentos
lsw.ListItems.Clear

End Sub

Private Sub sbCargaLsw()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Despliega en pantalla las solicitudes pendientes que estan autorizadas y
'               que estan dentro del rango de parametros suministrado por el usuario.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, lngConsec As Long, lngConsecInt As Long
Dim vFecha As Date, i As Integer, curMonto As Currency

lsw.Visible = True


If Trim(txtCantidadSolicitudes) = "0" Then
   lsw.ListItems.Clear
   Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

vFecha = Format(fxFechaServidor, "dd/mm/yyyy")

lngConsec = txtDocInicial
lngConsecInt = 0

If cboDoc.ItemData(cboDoc.ListIndex) = "TE" Then
    lngConsecInt = fxTesTipoDocConsecInterno(cbo.ItemData(cbo.ListIndex), cboDoc.ItemData(cboDoc.ListIndex), "/", cboPlan.ItemData(cboPlan.ListIndex))
End If

strSQL = "Select TOP " & txtCantidadSolicitudes & " *, dbo.fxTes_Cuentas_Bancarias_Pass(id_Banco,Cta_Ahorros) as 'Pass'" _
       & " From Tes_Transacciones Where Estado='P' And Tipo = '" & cboDoc.ItemData(cboDoc.ListIndex) _
       & "' And Id_Banco=" & cbo.ItemData(cbo.ListIndex) & " And Autoriza = 'S' and fecha_hold is null"

Select Case True
   Case optGeneraPor(0).Value
        strSQL = strSQL & " And NSolicitud Between " & Trim(txtGeneraNumeroDe) & " And " & Trim(txtGeneraNumeroHasta)
   Case optGeneraPor(1).Value
        strSQL = strSQL & " And Fecha_Solicitud Between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' And '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End Select

strSQL = strSQL & " Order by NSolicitud"

lsw.ListItems.Clear

i = 0
curMonto = 0

lsw.Tag = "S"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!NSolicitud)
        itmX.SubItems(1) = rs!Beneficiario
        
        If cboDoc.ItemData(cboDoc.ListIndex) = "TE" Then
            itmX.SubItems(2) = Format(lngConsec, "000") & "-" & lngConsecInt
        Else
            itmX.SubItems(2) = lngConsec
        End If
        
        itmX.SubItems(3) = Format(rs!Monto, "Standard")
        itmX.SubItems(4) = vFecha
        itmX.SubItems(5) = Trim(rs!Cta_Ahorros & "")
        itmX.SubItems(6) = IIf(IsNull(rs!FIRMAS_AUTORIZA_FECHA), "No", "Sí")
        
    If rs!Pass = 0 And rs!Tipo = "TE" Then
          itmX.ForeColor = vbRed
          itmX.Bold = True
          lsw.Tag = "N"
    End If
        
    i = i + 1
    curMonto = curMonto + rs!Monto
    
    
    If cboDoc.ItemData(cboDoc.ListIndex) = "TE" Then
        lngConsecInt = lngConsecInt + 1
    Else
        lngConsec = lngConsec + 1
    End If
    
    rs.MoveNext
Loop
rs.Close

txtCasos = Format(i, "###,###,###,##0")
txtMonto = Format(curMonto, "Standard")

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Function fxVerificaDatos(Optional vPrevista As Boolean = False) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String, vTipo As String

On Error GoTo vError

vMensaje = ""
vTipo = cboDoc.ItemData(cboDoc.ListIndex)



Select Case fxTesTipoDocGestion(cbo.ItemData(cbo.ListIndex), vTipo)
    Case "CK"
        If Trim(txtDocInicial) = "" Then vMensaje = vMensaje & vbCrLf & " - Suministre el Consecutivo para el Cheque Inicial"
        If Trim(txtVerificacion) = "" Then vMensaje = vMensaje & vbCrLf & " - Suministre el Intervalo de Verificación"
        
        If Trim(txtVerificacion) <> "" And IsNumeric(txtVerificacion) Then
            If CInt(txtVerificacion) = 0 Then vMensaje = vMensaje & vbCrLf & " - Suministre un Intervalo de Verificación mayor a Cero"
        End If
        
        If Not vPrevista Then
            strSQL = "Select ndocumento From Tes_Transacciones Where id_Banco = " & cbo.ItemData(cbo.ListIndex) _
                   & " And ndocumento between '" & Val(txtDocInicial) & "' And '" _
                   & Val(txtDocInicial) + Val((lsw.ListItems.Count) - 1) & "' and Tipo = '" & vTipo & "'"
                   
                   
            Call OpenRecordSet(rs, strSQL)
            Do While Not rs.EOF
              If Val(rs!nDocumento) >= Val(txtDocInicial) _
                    And Val(rs!nDocumento) <= Val((Me.lsw.ListItems.Count) - 1) Then
                
                vMensaje = vMensaje & vbCrLf & "Ya existe un Documento asignado [" & Trim(rs!nDocumento) & "]" _
                        & vbCrLf & " dentro del rango suministrado"
              End If
              rs.MoveNext
            Loop
            rs.Close
        End If
    
    Case "TE"
        'Nada
       
End Select


Select Case True
  Case optGeneraPor(0).Value
    If Trim(txtGeneraNumeroDe) = "" Then vMensaje = vMensaje & vbCrLf & " - Indique el Número de Solicitud Inicial"
    If Trim(txtGeneraNumeroHasta) = "" Then vMensaje = vMensaje & vbCrLf & " - Indique el Número de Solicitud Corte"
    If CLng(txtGeneraNumeroDe) > CLng(txtGeneraNumeroHasta) Then vMensaje = vMensaje & vbCrLf & " - El Número de Solicitud Inicial no debe ser mayor a la de Corte"
  
  Case optGeneraPor(1).Value
   If dtpInicio.Value > dtpCorte.Value Then vMensaje = vMensaje & vbCrLf & " - Error en Rango de Fechas"

End Select

If Len(vMensaje) = 0 Then
   fxVerificaDatos = True
Else
   fxVerificaDatos = False
   MsgBox vMensaje, vbExclamation
End If

Exit Function

vError:
   fxVerificaDatos = False
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   

End Function

Private Sub btnPlanes_Click()

GLOBALES.gTag = cbo.ItemData(cbo.ListIndex)

If cboPlan.ItemData(cboPlan.ListIndex) <> "-sp-" Then
    GLOBALES.gTag2 = cboPlan.ItemData(cboPlan.ListIndex)
Else
    GLOBALES.gTag2 = ""
End If

Call sbFormsCall("frmTES_TE_Planes", vbModal, , , False, Me)

Call cboDoc_Click

End Sub

Private Sub btnPuenteAccion_Click(Index As Integer)
Select Case Index
  Case 0 'Mover
    Call sbCtaPuenteAplica
  
  Case 1 'Cerrar
    fraCuentaPuenta.Visible = False
    lsw.Visible = True

    Call optGeneraPor_Click(0)
    Call cboDoc_Click
    Call cmdPrevista_Click
End Select
End Sub

Private Sub cbo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vPaso Then Exit Sub

If cbo.ListCount = 0 Then
   cbo.AddItem " "
    If TypeOf cbo Is XtremeSuiteControls.ComboBox Then
        cbo.ItemData(cbo.ListCount - 1) = CStr(0)
    Else
        cbo.ItemData(cbo.NewIndex) = 0
    End If
      cbo.Text = " "
End If

vPaso = True
    Call sbTesTiposDocsCargaCboAcceso(cboDoc, glogon.Usuario, cbo.ItemData(cbo.ListIndex), "G")
vPaso = False

strSQL = "exec spTes_Formatos_Bancos " & cbo.ItemData(cbo.ListIndex)
Call sbCbo_Llena_New(cboFormato, strSQL, False, True)

strSQL = "select Bp.COD_PLAN as 'IdX', Bp.COD_PLAN as 'ItmX'" _
       & " from TES_BANCOS B inner join TES_BANCO_PLANES_TE Bp on B.ID_BANCO = Bp.ID_BANCO" _
       & " Where B.ID_BANCO = " & cbo.ItemData(cbo.ListIndex) & " And B.UTILIZA_PLAN = 1" _
       & " order by Bp.COD_PLAN  asc"
Call sbCbo_Llena_New(cboPlan, strSQL, False, True)
If cboPlan.ListCount = 0 Then
   cboPlan.AddItem "Sin Plan"
   cboPlan.ItemData(cboPlan.ListCount - 1) = "-sp-"
   cboPlan.Text = "Sin Plan"
End If

Call cboDoc_Click

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    

End Sub


Private Sub cboCtaPuente_Click()
Dim strSQL As String

If vPaso Then Exit Sub
If cboCtaPuente.ListCount <= 0 Then Exit Sub

vGrid.MaxRows = 0
vGrid.MaxRows = 7

Me.MousePointer = vbHourglass

strSQL = "select 0,nsolicitud,codigo,beneficiario,monto,tipo,cta_Ahorros" _
       & " from Tes_Transacciones where id_banco = " & cboCtaPuente.ItemData(cboCtaPuente.ListIndex) _
       & " and  ESTADO = 'P' and Tipo = '" & cboDoc.ItemData(cboDoc.ListIndex) & "'"
Call sbCargaGrid(vGrid, 7, strSQL)

vGrid.MaxRows = vGrid.MaxRows - 1
Me.MousePointer = vbDefault

End Sub

Private Sub cboDoc_Click()

If vPaso Then Exit Sub

'If Mid(cboDoc.Text, 1, 2) = "CK" Then
'    GroupBox_CK.Visible = True
'Else
'    GroupBox_CK.Visible = False
'End If


txtDocInicial = fxTesTipoDocConsec(cbo.ItemData(cbo.ListIndex), cboDoc.ItemData(cboDoc.ListIndex), "/", cboPlan.ItemData(cboPlan.ListIndex))

txtCasos = 0
txtMonto = 0

lsw.ListItems.Clear

Timer1.Interval = 1

End Sub

Private Sub cboPlan_Click()
If vPaso Then Exit Sub

Call cboDoc_Click

End Sub

Private Sub cmdCuentaVerifica_Click()
Call sbRevisaCuentas
End Sub

Private Sub cmdGenera_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Verifica que no falten datos para efectuar la generacion de las solicitudes
'               pendientes.
'REFERENCIAS:   sbGenerarDocumento - (Verifica que no falten datos para efectuar la generacion
'               de las solicitudes pendientes que estan autorizadas)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo vError

Me.MousePointer = vbHourglass

Call cmdPrevista_Click

If fxVerificaDatos(False) Then
  If lsw.Tag = "S" Then
    Call sbEmitirDocumentos
    Call sbCargaLsw
  Else
    Me.MousePointer = vbDefault
    MsgBox "Existen Solicitudes con Restricciones en las Cuentas de Ahorros, verifique (Marcas en Rojo)...", vbExclamation
  End If
End If

Me.MousePointer = vbDefault

Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub cmdPrevista_Click()

On Error GoTo vError

If fxVerificaDatos(True) Then
    Call sbCargaLsw
End If

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdPuente_Click()
 fraCuentaPuenta.top = 0
 fraCuentaPuenta.Height = Me.Height
 
 fraCuentaPuenta.Visible = True
 lsw.Visible = False
 Call cboCtaPuente_Click
 
End Sub

Private Sub cmdReporte_Click()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Banking"
    
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "Banco='" & cbo.Text & "'"

    .ReportFileName = SIFGlobal.fxPathReportes("Banking_AutorizaPendientes.rpt")
    
    strSQL = "{CHEQUES.ESTADO} = 'P' And {CHEQUES.AUTORIZA} = 'N' And ISNULL({CHEQUES.FECHA_HOLD})" _
           & " And {CHEQUES.TIPO}='" & cboDoc.ItemData(cboDoc.ListIndex) _
           & "' And {CHEQUES.ID_BANCO} = " & cbo.ItemData(cbo.ListIndex) _
           & " And {CHEQUES.NSOLICITUD} in " & txtGeneraNumeroDe & " to " & txtGeneraNumeroHasta _
           & " And {CHEQUES.FECHA_SOLICITUD} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
           & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    .SelectionFormula = strSQL
    .PrintReport

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
Dim strSQL As String
On Error GoTo vError

vModulo = 9

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

vPaso = True
 Call sbTesBancoCargaCboAccesoGestion(cbo, glogon.Usuario, "Genera")
vPaso = False
 
'Carga Cuenta Puente
strSQL = "select B.id_Banco as 'IdX' ,rtrim(B.descripcion) as 'ItmX'" _
       & " from Tes_Bancos B inner join tes_Banco_ASG A on B.id_Banco = A.id_Banco" _
       & " and A.nombre ='" & glogon.Usuario & "' Where B.estado = 'A' and B.puente  = 1"
Call sbCbo_Llena_New(cboCtaPuente, strSQL, False, True)
       
 
dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value


txtVerificacion = 5


With lsw.ColumnHeaders
  .Clear
  .Add , , "No. Solicitud", 1200
  .Add , , "Beneficiario", 3200
  .Add , , "No.Documento", 1400, vbCenter
  .Add , , "Monto", 1400, vbRightJustify
  .Add , , "Fecha", 1400, vbCenter
  .Add , , "Cuenta", 2400
  .Add , , "Firmas", 1000, vbCenter
End With

Call cbo_Click

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbRevisaCuentas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spTes_Cuentas_Revisa " & cbo.ItemData(cbo.ListIndex)
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Call cmdPrevista_Click

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Call cmdPrevista_Click

End Sub




'Private Sub lsw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo vError
'
'    lsw.SortKey = ColumnHeader.Index - 1
'
'    If (lsw.SortOrder = lvwAscending) Then
'        lsw.SortOrder = lvwDescending
'    Else
'        lsw.SortOrder = lvwAscending
'    End If
'
'    lsw.Sorted = True
'    Exit Sub
'
'vError:
'   MsgBox "Ocurrió un error al ordenar los datos de la columna seleccionada.", vbCritical
'
'End Sub

Private Sub optGeneraPor_Click(Index As Integer)
Select Case Index
  Case 0
    txtGeneraNumeroDe.Enabled = True
    txtGeneraNumeroHasta.Enabled = True
    txtGeneraNumeroDe.SetFocus

    dtpInicio.Enabled = False
    dtpCorte.Enabled = False
    
   
  Case 2
    dtpInicio.Enabled = True
    dtpCorte.Enabled = True
    dtpInicio.SetFocus

    txtGeneraNumeroDe.Enabled = False
    txtGeneraNumeroHasta.Enabled = False

End Select

End Sub


Private Sub Timer1_Timer()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Despliega el # de solicitudes pendientes por generar.
'REFERENCIAS:   sbBuscaDocumentos - (Despliega el # de solicitudes pendientes por generar,
'               asi como el # de Solicitud tanto de la primera como de la ultima)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo vError

Timer1.Interval = 0

Me.MousePointer = vbHourglass
 Call sbBuscaDocumentos
Me.MousePointer = vbDefault

vError:


End Sub

Private Function fxCuentaBanco(vBanco As Integer) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select ctaconta as Cuenta from Tes_Bancos where id_banco = " & vBanco
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
 fxCuentaBanco = Trim(rs!Cuenta)
Else
 fxCuentaBanco = ""
End If
rs.Close

End Function


Private Sub sbCtaPuenteAplica()
Dim strSQL As String, i As Long

On Error GoTo vError

If cboCtaPuente.ItemData(cboCtaPuente.ListIndex) = cbo.ItemData(cbo.ListIndex) Then
  MsgBox "La cuenta Bancaria puente es igual a la cuenta bancaria actual...!", vbExclamation
  Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = ""

For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.Col = 1
    If vGrid.Value = vbChecked Then
        vGrid.Col = 2
        strSQL = strSQL & Space(10) & "exec spTes_Traslados_Cuenta_Puente " & vGrid.Text & "," & cbo.ItemData(cbo.ListIndex) & ",'" & glogon.Usuario & "'"
        
        If Len(strSQL) > 20000 Then
           Call ConectionExecute(strSQL)
           strSQL = ""
        End If
    End If
Next i

'Lote Final
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If


Me.MousePointer = vbDefault

fraCuentaPuenta.Visible = False
lsw.Visible = True

Call optGeneraPor_Click(0)
Call cboDoc_Click
Call cmdPrevista_Click

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Call cboCtaPuente_Click

End Sub




Private Sub txtCantidadSolicitudes_LostFocus()

If Trim(txtCantidadSolicitudes) = "" Then
   txtCantidadSolicitudes = "0"
End If

End Sub

Private Sub txtDocInicial_KeyPress(KeyAscii As Integer)
KeyAscii = Validacion(KeyAscii)

If KeyAscii = vbKeyReturn Then
   If cboDoc.ItemData(cboDoc.ListIndex) <> "TE" Then
      txtCantidadSolicitudes.SetFocus
   End If
End If
End Sub


Private Sub txtGeneraNumeroDe_KeyPress(KeyAscii As Integer)
KeyAscii = Validacion(KeyAscii)

If KeyAscii = vbKeyReturn Then
   txtGeneraNumeroHasta.SetFocus
End If
End Sub


Private Sub txtGeneraNumeroHasta_KeyPress(KeyAscii As Integer)
KeyAscii = Validacion(KeyAscii)

If KeyAscii = vbKeyReturn Then
   cmdPrevista.SetFocus
End If
End Sub


Private Sub txtVerificacion_KeyPress(KeyAscii As Integer)
KeyAscii = Validacion(KeyAscii)

If KeyAscii = vbKeyReturn Then
   optGeneraPor(0).SetFocus
End If
End Sub


