VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCC_EstadoCuenta 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estados de Cuenta"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   HelpContextID   =   9007
   Icon            =   "frmCC_EstadoCuenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   8745
   Begin XtremeSuiteControls.CheckBox chkSegmentos 
      Height          =   252
      Left            =   5640
      TabIndex        =   14
      Top             =   1800
      Width           =   3012
      _Version        =   1441793
      _ExtentX        =   5313
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Utilizar Filtros por Segmentos "
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   16
      Alignment       =   1
   End
   Begin XtremeSuiteControls.RadioButton rbSalida 
      Height          =   252
      Index           =   0
      Left            =   3240
      TabIndex        =   29
      Top             =   1320
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Pantalla"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   2292
      Left            =   240
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   8412
      _Version        =   1441793
      _ExtentX        =   14838
      _ExtentY        =   4043
      _StockProps     =   79
      Caption         =   "Filtros:"
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   2280
         TabIndex        =   17
         Top             =   720
         Width           =   6012
         _Version        =   1441793
         _ExtentX        =   10610
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
      Begin XtremeSuiteControls.ComboBox cboDept 
         Height          =   312
         Left            =   2280
         TabIndex        =   20
         Top             =   1080
         Width           =   6012
         _Version        =   1441793
         _ExtentX        =   10610
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
      Begin XtremeSuiteControls.ComboBox cboSeccion 
         Height          =   312
         Left            =   2280
         TabIndex        =   21
         Top             =   1440
         Width           =   6012
         _Version        =   1441793
         _ExtentX        =   10610
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
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   2280
         TabIndex        =   28
         Top             =   360
         Width           =   2532
         _Version        =   1441793
         _ExtentX        =   4471
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
      Begin XtremeSuiteControls.CheckBox chkSinCorreo 
         Height          =   252
         Left            =   2280
         TabIndex        =   34
         Top             =   1920
         Width           =   5892
         _Version        =   1441793
         _ExtentX        =   10393
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Imprimir únicamente las personas que no tienen Email registrado   "
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   3
         Left            =   720
         TabIndex        =   33
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Seccion"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Departatamento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   18
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Institución"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
   End
   Begin XtremeSuiteControls.PushButton cmdReporteExcedentes 
      Height          =   732
      Left            =   6600
      TabIndex        =   11
      Top             =   5040
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Excedentes"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCC_EstadoCuenta.frx":000C
   End
   Begin VB.Frame fraASE 
      BorderStyle     =   0  'None
      Height          =   1452
      Left            =   9240
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox lblDescripcionCentro 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Presione (F4) para Consultar"
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtCentro 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         ToolTipText     =   "Presione (F4) Para Consultar"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtUnidad 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         ToolTipText     =   "Presione (F4) Para Consultar"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox lblDescripcion 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Presione (F4) para Consultar"
         Top             =   360
         Width           =   4335
      End
      Begin VB.CheckBox chkUnidad 
         Caption         =   "Por Unidad programática"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   -120
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "C. Trabajo"
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
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   852
      End
      Begin VB.Label Label2 
         Caption         =   "Código"
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
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton cmdReporte 
      Height          =   732
      Left            =   6600
      TabIndex        =   12
      Top             =   4200
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Estado de Cuenta"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCC_EstadoCuenta.frx":09DA
   End
   Begin XtremeSuiteControls.PushButton cmdConstanciaCrd 
      Height          =   732
      Left            =   6600
      TabIndex        =   13
      Top             =   5880
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Constancia de Deudas"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCC_EstadoCuenta.frx":1196
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3240
      TabIndex        =   23
      Top             =   480
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   4920
      TabIndex        =   24
      Top             =   5880
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
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
   Begin XtremeSuiteControls.ComboBox cboCorte 
      Height          =   312
      Left            =   4920
      TabIndex        =   27
      Top             =   4320
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2566
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
   Begin XtremeSuiteControls.RadioButton rbSalida 
      Height          =   252
      Index           =   1
      Left            =   5160
      TabIndex        =   30
      Top             =   1320
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "E-mail"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton rbSalida 
      Height          =   252
      Index           =   2
      Left            =   7080
      TabIndex        =   31
      Top             =   1320
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Impresora"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtEmail 
      Height          =   312
      Left            =   3240
      TabIndex        =   32
      Top             =   840
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1560
      TabIndex        =   22
      Top             =   480
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
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
   Begin XtremeSuiteControls.ComboBox cboPeriodo 
      Height          =   330
      Left            =   3480
      TabIndex        =   37
      Top             =   5280
      Width           =   2895
      _Version        =   1441793
      _ExtentX        =   5106
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   1800
      TabIndex        =   36
      Top             =   840
      Width           =   1212
   End
   Begin XtremeShortcutBar.ShortcutCaption lblEstado 
      Height          =   372
      Left            =   0
      TabIndex        =   35
      Top             =   6960
      Width           =   9132
      _Version        =   1441793
      _ExtentX        =   16108
      _ExtentY        =   656
      _StockProps     =   14
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
   Begin VB.Label lblCorte 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Corte para el Estado de Cuenta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   26
      Top             =   4320
      Width           =   4452
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo de los Excedentes"
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
      Index           =   1
      Left            =   2760
      TabIndex        =   25
      Top             =   5040
      Width           =   3495
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Corte Intereses / Día Pago para constancia de deudas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   5880
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1212
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmCC_EstadoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vSalida As Integer
Dim vRA_Access As Boolean


Private Sub cbo_Click()
Dim strSQL As String

If vPaso Or cbo.ListCount <= 0 Then Exit Sub

Dim pInstitucion As Long

If cbo.Text = "TODOS" Then
    pInstitucion = 0
Else
    pInstitucion = cbo.ItemData(cbo.ListIndex)
End If


vPaso = True
    strSQL = "select rtrim(cod_departamento) as  'IdX', rtrim(descripcion) as 'ItmX'" _
           & " from AFdepartamentos where cod_institucion = " & pInstitucion
    Call sbCbo_Llena_New(cboDept, strSQL, True, True)
vPaso = False

End Sub



Private Sub cboDept_Click()
Dim strSQL As String

If vPaso Or cboDept.ListCount <= 0 Then Exit Sub

Dim pInstitucion As Long

If cbo.Text = "TODOS" Then
    pInstitucion = 0
Else
    pInstitucion = cbo.ItemData(cbo.ListIndex)
End If

vPaso = True
    strSQL = "select rtrim(cod_seccion) as 'IdX',rtrim(descripcion) as 'ItmX'" _
           & " from AFSecciones where cod_institucion = " & pInstitucion _
           & " and cod_departamento = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
    Call sbCbo_Llena_New(cboSeccion, strSQL, True, True)
vPaso = False

End Sub

Private Sub sbEstado_Masivo_Email()
Dim strSQL As String, rs As New ADODB.Recordset

Dim pEstado As String, pInstitucion As String
Dim pDept As String, pSeccion As String


On Error GoTo vError

Me.MousePointer = vbHourglass
'@Institucion int = Null, @Departamento varchar(10) = Null
'                                     , @Seccion varchar(10) = Null, @EstadoPersona varchar(10) = Null
'                                     , @Usuario varchar(30) )
                                     
strSQL = "exec spSys_Estados_Cuenta_Email "

If cbo.Text = "TODOS" Then
    strSQL = strSQL & "Null,"
    
    pInstitucion = "T"
Else
    strSQL = strSQL & cbo.ItemData(cbo.ListIndex) & ","
    pInstitucion = cbo.ItemData(cbo.ListIndex)
End If

If cboDept.Text = "TODOS" Then
    strSQL = strSQL & "Null,"
    pDept = "T"
Else
    strSQL = strSQL & "'" & cboDept.ItemData(cboDept.ListIndex) & "',"
    pDept = cboDept.ItemData(cboDept.ListIndex)
End If

If cboSeccion.Text = "TODOS" Then
    strSQL = strSQL & "Null,"
    pSeccion = "T"
Else
    strSQL = strSQL & "'" & cboSeccion.ItemData(cboSeccion.ListIndex) & "',"
    pSeccion = cboSeccion.ItemData(cboSeccion.ListIndex)
End If

If cboEstado.Text = "TODOS" Then
    strSQL = strSQL & "Null,"
Else
    strSQL = strSQL & "'" & cboEstado.ItemData(cboEstado.ListIndex) & "',"
End If
pEstado = cboEstado.Text

strSQL = strSQL & "'" & glogon.Usuario & "', '" & Format(cboCorte.ItemData(cboCorte.ListIndex), "YYYY-MM-DD") & " 23:59:00'"

lblEstado.Caption = "Procesando Estados de Cuenta [Espere]"

Call OpenRecordSet(rs, strSQL)
If Not glogon.error Then
    Me.MousePointer = vbDefault
    MsgBox "Estados de Cuentas, Notificados vía Email. Satifactoriamente! ", vbInformation
End If
rs.Close

lblEstado.Caption = ""

Call Bitacora("Aplica", "EC Masivo: [I: " & pInstitucion & ", D: " _
            & pDept & ", S: " & pSeccion _
            & "] Estado: " & pEstado)

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    lblEstado.Caption = ""
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub sbReporteEC_ASE()
Dim strRuta As String, strSQL As String, rs As New ADODB.Recordset
Dim recCreditos As New ADODB.Recordset, vFecha As Date


vFecha = fxFechaServidor



If chkUnidad.Value = vbChecked Then
   If Trim(txtUnidad) = "" Then
       MsgBox "Especifique La Unidad Programatica", vbExclamation
       txtUnidad.SetFocus
       Exit Sub
   End If
   
Else
   If Trim(txtCedula) = "" Then
      MsgBox "Especifique La Cédula de identidad de la Persona", vbExclamation
      txtCedula.SetFocus
      Exit Sub
   End If
End If


Me.MousePointer = vbHourglass

On Error GoTo vError


With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = False
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "Reportes Módulo de Cuentas Corrientes"
     .Formulas(0) = "Fecha='REPORTE AL  " & Format(vFecha, "dd/mm/yyyy") & "'"
     .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
     .Connect = glogon.ConectRPT
End With


If chkUnidad.Value = vbChecked Then
 With frmContenedor.Crt
    .Formulas(14) = "SistemaFecha = 'Fecha/Hora : " & vFecha & "'"
    .Formulas(15) = "SistemaUsuario = 'Usuario : " & glogon.Usuario & "'"
        
        
    .ReportFileName = SIFGlobal.fxPathReportes("Sys_EstadoCuenta02Unidad.rpt")
    If cboEstado.Text = "TODOS" Then
        If txtCentro = "" Then
          .SelectionFormula = "{VISTA_SOCIOAHORROS.UP} ='" & Trim(txtUnidad) & "'"
        Else
          .SelectionFormula = "{VISTA_SOCIOAHORROS.UP} ='" & Trim(txtUnidad) & "' and {VISTA_SOCIOAHORROS.UT} ='" & Trim(txtCentro) & "'"
        End If
    Else
        
        If txtCentro = "" Then
          .SelectionFormula = "{VISTA_SOCIOAHORROS.ESTADOACTUAL} = '" & cboEstado.ItemData(cboEstado.ListIndex) & "' AND {VISTA_SOCIOAHORROS.UP} ='" & Trim(txtUnidad) & "'"
        Else
          .SelectionFormula = "{VISTA_SOCIOAHORROS.ESTADOACTUAL} = '" & cboEstado.ItemData(cboEstado.ListIndex) & "' AND {VISTA_SOCIOAHORROS.UP} ='" & Trim(txtUnidad) & "' and {VISTA_SOCIOAHORROS.UT} ='" & Trim(txtCentro) & "'"
        End If
    End If
    
    .SubreportToChange = "CreditosActivos"
    .SelectionFormula = "{vEC_Credito.cedula} = {?Pm-VISTA_SOCIOAHORROS.CEDULA} And {vEC_Credito.VISIBLE_EC} = 1"
      
    .SubreportToChange = "Fianzas"
    .SelectionFormula = "{FIADORES.CEDULAF} = {?Pm-VISTA_SOCIOAHORROS.CEDULA} And {vEC_Credito.VISIBLE_EC} = 1"
   
    Call Bitacora("Imprime", "Estado Cuenta x Unidad : " & txtUnidad)
 
   If rbSalida.Item(2).Value Then .Destination = crptToPrinter
   .PrintReport
 
 End With

Else
      
 strSQL = "Select isnull(count(*),0) as Existe From Reg_Creditos" _
                    & " Where Cedula='" & Trim(txtCedula) & "' And Estado='A'"
 Call OpenRecordSet(recCreditos, strSQL)
      
 With frmContenedor.Crt
  
          strSQL = "exec spVoxAhorros '" & txtCedula & "'"
          Call OpenRecordSet(rs, strSQL)
            .Formulas(2) = "fxDSOMonto = " & rs!Disponible
            .Formulas(3) = "fxDSOSaldo = " & rs!Saldos
            .Formulas(4) = "fxDSOPlazo = " & rs!Plazo
            .Formulas(5) = "fxDSOCuota = " & CCur(fxCalcula_Cuota(rs!Disponible - rs!Saldos, rs!Plazo, rs!Tasa))
          rs.Close
        
          strSQL = "exec spVoxFiduciario '" & txtCedula & "'"
          Call OpenRecordSet(rs, strSQL)
            .Formulas(6) = "fxDPEMonto = " & rs!Disponible
            .Formulas(7) = "fxDPESaldo = " & rs!Saldos
            .Formulas(8) = "fxDPEPlazo = " & rs!Plazo
            .Formulas(9) = "fxDPECuota = " & CCur(fxCalcula_Cuota(rs!Disponible - rs!Saldos, rs!Plazo, rs!Tasa))
          rs.Close
        
        strSQL = "exec spVoxExcedenteCredito '" & txtCedula & "'"
        Call OpenRecordSet(rs, strSQL)
            .Formulas(10) = "fxEXCMonto = " & rs!Base
            .Formulas(11) = "fxEXCSaldo = " & rs!Saldos
            .Formulas(12) = "fxEXCPlazo = '---'"
            .Formulas(13) = "fxEXCCuota = '---'"
        rs.Close
        
        .Formulas(14) = "SistemaFecha = 'Fecha/Hora : " & vFecha & "'"
        .Formulas(15) = "SistemaUsuario = 'Usuario : " & glogon.Usuario & "'"
        
        
          .ReportFileName = SIFGlobal.fxPathReportes("Sys_EstadoCuenta02ASE.rpt")
          .SelectionFormula = "{VISTA_SOCIOAHORROS.CEDULA} ='" & Trim(txtCedula) & "'"
  
            .SubreportToChange = "CreditosActivos"
            .SelectionFormula = "{vEC_Credito.cedula} = {?Pm-VISTA_SOCIOAHORROS.CEDULA} And {vEC_Credito.VISIBLE_EC} = 1"
              
            .SubreportToChange = "Fianzas"
            .SelectionFormula = "{FIADORES.CEDULAF} = {?Pm-VISTA_SOCIOAHORROS.CEDULA} And {vEC_Credito.VISIBLE_EC} = 1"
  
  
  Call Bitacora("Imprime", "Estado Cuenta Ced." & txtCedula)
  
  
  If rbSalida.Item(2).Value Then .Destination = crptToPrinter
  .PrintReport
 End With
      
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbReporteEC_SYS()

If chkSegmentos.Value = vbUnchecked Then
  Select Case vSalida
    Case 0 'Pantalla
      Call sbEstadoCuenta(txtCedula, 0, cboCorte.ItemData(cboCorte.ListIndex))

    Case 2 'Impresora
      Call sbEstadoCuenta(txtCedula, 2, cboCorte.ItemData(cboCorte.ListIndex))
  
  End Select

  Exit Sub
End If


'Masivo: Pantalla e Impresora
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFecha As Date, vFechaCorte As Date, vCorte As String
Dim vPatrimonio As Integer, vFondos As Integer, vCreditos As Integer, vFianzas As Integer

vFecha = fxFechaServidor
vCorte = cboCorte.ItemData(cboCorte.ListIndex)

If vCorte <> "HOY" Then
    vFechaCorte = vCorte
    If Abs(DateDiff("d", vFecha, vFechaCorte)) = 0 Then
        vCorte = "HOY"
    End If
End If



strSQL = "select * from sif_empresa"
Call OpenRecordSet(rs, strSQL)
  vPatrimonio = rs!ec_visible_patrimonio
  vFondos = rs!ec_visible_fondos
  vCreditos = rs!ec_visible_creditos
  vFianzas = rs!ec_visible_fianzas
rs.Close



strSQL = ""
If vCorte = "HOY" Then
         
        If cboEstado.Text <> "TODOS" Then
                strSQL = "{vSIF_EC_Principal.ESTADOACTUAL} = '" & cboEstado.ItemData(cboEstado.ListIndex) & "'"
        End If
        
        If cbo.Text <> "TODOS" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{vSIF_EC_Principal.COD_INSTITUCION} = " & cbo.ItemData(cbo.ListIndex)
        End If
        
        If Not (cboDept.Text = "TODOS" Or cboDept.Text = "") Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{vSIF_EC_Principal.COD_DEPARTAMENTO} = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
        End If
        
        If Not (cboSeccion.Text = "TODOS" Or cboSeccion.Text = "") Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{vSIF_EC_Principal.COD_SECCION} = '" & cboSeccion.ItemData(cboSeccion.ListIndex) & "'"
        End If


        If chkSinCorreo.Value = xtpChecked Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{vSIF_EC_Principal.AF_EMAIL} = ''"
        End If

Else
       strSQL = "YEAR({vSYS_EC_Corte_Principal.CORTE}) = " & Year(vFechaCorte) _
              & " AND MONTH({vSYS_EC_Corte_Principal.CORTE}) = " & Month(vFechaCorte)
        
        
        If cboEstado.Text <> "TODOS" Then
           strSQL = strSQL & " AND {vSYS_EC_Corte_Principal.ESTADOACTUAL} = '" & cboEstado.ItemData(cboEstado.ListIndex) & "'"
        End If
        
        If cbo.Text <> "TODOS" Then
          strSQL = strSQL & " AND {vSYS_EC_Corte_Principal.COD_INSTITUCION} = " & cbo.ItemData(cbo.ListIndex)
        End If
        
        If Not (cboDept.Text = "TODOS" Or cboDept.Text = "") Then
          strSQL = strSQL & " AND {vSYS_EC_Corte_Principal.COD_DEPARTAMENTO} = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
        End If
        
        If Not (cboSeccion.Text = "TODOS" Or cboSeccion.Text = "") Then
          strSQL = strSQL & " AND {vSYS_EC_Corte_Principal.COD_SECCION} = '" & cboSeccion.ItemData(cboSeccion.ListIndex) & "'"
        End If

        If chkSinCorreo.Value = xtpChecked Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{vSYS_EC_Corte_Principal.AF_EMAIL} = ''"
        End If

End If


'Estado de Cuenta por Segmentos
Call sbEstadoCuentaInst(vSalida, strSQL, vCorte, vFechaCorte _
                , vPatrimonio, vFondos, vCreditos, vFianzas)


End Sub





Private Sub chkSegmentos_Click()
If chkSegmentos.Value = xtpChecked Then
    gbFiltros.Visible = True
Else
    gbFiltros.Visible = False
End If
End Sub

Private Sub cmdConstanciaCrd_Click()

Call sbCrdConstancia(txtCedula.Text, dtpCorte.Value, vSalida)

End Sub

Private Sub cmdReporte_Click()

If rbSalida(1).Value Then
    If chkSegmentos.Value = xtpChecked Then
        Call sbEstado_Masivo_Email
    Else
        If txtEmail.Text <> "" Then
          Call sbEstadoCuenta_Email_Corte(txtCedula.Text, txtEmail.Text, cboCorte.ItemData(cboCorte.ListIndex))
        Else
          MsgBox "La persona no cuenta con un correo registrado, verifique!", vbExclamation
        End If
    End If
   
Else
        
    Call sbReporteEC_SYS
'    If Not GLOBALES.SysASEVersion Then
'      Call sbReporteEC_SYS
'    Else
'      Call sbReporteEC_ASE
'    End If
End If

End Sub

Private Sub sbBusqueda(Index As Integer)

gBusquedas.Convertir = "N"
gBusquedas.Col1Name = "Identificación Colilla"
gBusquedas.Col2Name = "Identificación Real"
gBusquedas.Col3Name = "Nombre"
gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"

Select Case Index
   
   Case 0

        gBusquedas.Columna = "cedula"
        gBusquedas.Orden = "cedula"
   
   Case 2
        gBusquedas.Columna = "nombre"
        gBusquedas.Orden = "nombre"

End Select

    frmBusquedas.Show vbModal
    
    txtCedula.Text = Trim(gBusquedas.Resultado)
    txtNombre.Text = Trim(gBusquedas.Resultado3)


End Sub


Private Sub cmdReporteExcedentes_Click()
Dim strRuta As String

Me.MousePointer = vbHourglass

On Error GoTo vError

If chkSegmentos.Value = vbUnchecked Then
     Call sbEstadoExcedentes(txtCedula.Text, cboPeriodo.ItemData(cboPeriodo.ListIndex))
     Call Bitacora("Imprime", "Estado Excedentes Ced." & txtCedula & " Periodo: " & cboPeriodo.Text)

    Me.MousePointer = vbDefault
    Exit Sub
End If


With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = True
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "Reportes de Cuentas Corrientes"
     
     .Connect = glogon.ConectRPT
     
     If GLOBALES.SysASEVersion Then
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_EstadoExcedentesUnidad.rpt")
     Else
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_EstadoExcedentes.rpt")
     End If
     
     .Formulas(0) = "SistemaFecha = 'Fecha/Hora : " & fxFechaServidor & "'"
     .Formulas(1) = "SistemaUsuario = 'Usuario : " & glogon.Usuario & "'"
     
     
     .SelectionFormula = "{EXC_CIERRE.ID_PERIODO} = " & cboPeriodo.ItemData(cboPeriodo.ListIndex)
     
     
     If cbo.Text <> "TODOS" Then
         .SelectionFormula = .SelectionFormula & " AND {SOCIOS.COD_INSTITUCION} = " & cbo.ItemData(cbo.ListIndex)
     End If
     
     
     
     If GLOBALES.SysASEVersion And chkUnidad.Value = vbChecked Then
       If Len(txtUnidad.Text) > 0 Then
         .SelectionFormula = .SelectionFormula & " AND {SOCIOS.UP} = '" & txtUnidad.Text & "'"
       End If
      
       If Len(txtCentro.Text) > 0 Then
         .SelectionFormula = .SelectionFormula & " AND {SOCIOS.UT} = '" & txtCentro.Text & "'"
       End If
     End If 'GLOBALES.SysASEVersion And chkUnidad.Value = vbChecked
     
'     .SubreportToChange = "Exc_Carga"
'     .SelectionFormula = "{EXC_CARGA.ID_PERIODO} = {?Pm-EXC_CIERRE.ID_PERIODO} AND {EXC_CARGA.CEDULA} = {?Pm-EXC_CIERRE.CEDULA}"
    
    
     If vSalida = 2 Then .Destination = crptToPrinter
     .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMes As Integer, vFecha As Date

On Error GoTo vError

vModulo = 10

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

If GLOBALES.SysASEVersion Then
   fraASE.Visible = True
Else
   fraASE.Visible = False
End If

vPaso = True
    strSQL = "select Idx, ItmX  From vExc_Periodos where ESTADO = 'C' order by IdX desc"
    Call sbCbo_Llena_New(cboPeriodo, strSQL, False, True)
vPaso = False


vPaso = True
    strSQL = "select cod_institucion as 'Idx',rtrim(descripcion) as 'ItmX' from instituciones where Activa = 1"
    Call sbCbo_Llena_New(cbo, strSQL, True, True)
vPaso = False


vPaso = True
    strSQL = "select rtrim(cod_estado) as 'IdX', rtrim(descripcion) as 'ItmX' from AFI_Estados_Persona where ACTIVO = 1"
    Call sbCbo_Llena_New(cboEstado, strSQL, True, True)
vPaso = False

strSQL = "exec spSys_Periodos_Cierre_Consulta"
Call sbCbo_Llena_New(cboCorte, strSQL, False, False)


Call cbo_Click
Call cboDept_Click

vFecha = fxFechaServidor
dtpCorte.Value = vFecha


'strSQL = "select isnull(max(periodo_de),0) as PeriodoI from excedentes_parcierre"
'Call OpenRecordSet(rs, strSQL)
'    txtPeriodoDe = rs!periodoi
'rs.Close
'
'If txtPeriodoDe = 0 Then
'
'    vMes = Month(vFecha)
'
'    If vMes > 9 Then
'      txtPeriodoDe = Year(vFecha)
'    Else
'      txtPeriodoDe = Year(vFecha) + 1
'    End If
'End If

Call chkSegmentos_Click
Call rbSalida_Click(0)

vError:


End Sub


Private Sub lblDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdReporte.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select codigo,descripcion from uprogramatica"
    gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtUnidad.Text = gBusquedas.Resultado
  lblDescripcion.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub lblDescripcionCentro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdReporte.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "ut_descripcion"
    gBusquedas.Orden = "ut_descripcion"
    gBusquedas.Consulta = "select ut_codigo,ut_descripcion from utrabajo"
    gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCentro.Text = gBusquedas.Resultado
  lblDescripcionCentro = gBusquedas.Resultado2
End If
End Sub

Private Sub rbSalida_Click(Index As Integer)

vSalida = Index

Select Case Index
    Case 0, 2
        chkSinCorreo.Visible = True
        lblCorte.Visible = True
        cboCorte.Visible = True
    Case 1 'Email
        chkSinCorreo.Visible = False
        lblCorte.Visible = True
        cboCorte.Visible = True
End Select

End Sub

Private Sub txtCedula_Change()
 txtNombre = fxNombre(txtCedula)
End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(0)
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
KeyAscii = (Validacion(KeyAscii))

If KeyAscii = vbKeyReturn Then
   cmdReporte.SetFocus
End If
End Sub


Private Sub txtCedula_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error Resume Next


txtNombre.Text = ""
txtEmail.Text = ""

If Trim(txtCedula.Text) <> "" Then
 
    'Valida Acceso a Expediente
    vRA_Access = fxSys_RA_Consulta(Trim(txtCedula.Text), glogon.Usuario)
     
    If Not vRA_Access Then
        MsgBox "Esta persona se encuentra con -> Expediente Restringido <- Requiere de Autorización para Consultar!", vbExclamation
        txtCedula.Text = ""
        txtNombre.Text = ""
        Exit Sub
    End If

 
 strSQL = "Select nombre, af_Email from Socios Where Cedula ='" & Trim(txtCedula.Text) & "'"
 Call OpenRecordSet(rs, strSQL)
 
 If rs.EOF And rs.BOF Then
    MsgBox "No se encontró registro", vbExclamation
    txtCedula.Text = ""
    txtCedula.SetFocus
 Else
    txtNombre.Text = Trim(rs!Nombre & "")
    txtEmail.Text = Trim(rs!AF_Email & "")
 End If
 
 rs.Close

End If

End Sub


Private Sub txtCentro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then lblDescripcionCentro.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "Ut_codigo"
    gBusquedas.Orden = "Ut_codigo"
    gBusquedas.Consulta = "select Ut_codigo,Ut_descripcion from utrabajo"
    gBusquedas.Filtro = " and UT_CODIGO in(select UT from SOCIOS where UP = '" & txtUnidad.Text & "' group by UT)"
    
  frmBusquedas.Show vbModal
  txtCentro.Text = gBusquedas.Resultado
  lblDescripcionCentro.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCentro_LostFocus()
 lblDescripcionCentro.Text = fxCentroTabajo(txtCentro)
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(2)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   cmdReporte.SetFocus
End If
End Sub

Private Sub txtUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then lblDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "codigo"
    gBusquedas.Orden = "codigo"
    gBusquedas.Consulta = "select codigo,descripcion from uprogramatica"
    gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtUnidad.Text = gBusquedas.Resultado
  lblDescripcion.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtUnidad_LostFocus()
 lblDescripcion.Text = fxgAFIDepartamento(cbo.ItemData(cbo.ListIndex), txtUnidad.Text)

End Sub


Private Function fxCentroTabajo(vCodigo As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select ut_descripcion from utrabajo where ut_codigo = '" & vCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    fxCentroTabajo = rs!UT_Descripcion
Else
  fxCentroTabajo = "No existe este Centro de Trabajo"
End If
rs.Close

End Function
