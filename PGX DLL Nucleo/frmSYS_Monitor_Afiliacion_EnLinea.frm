VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmSYS_Monitor_Afiliacion_EnLinea 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Afiliación en Línea"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   18000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox gbCaso 
      Height          =   10335
      Left            =   12120
      TabIndex        =   25
      Top             =   0
      Width           =   5775
      _Version        =   1310723
      _ExtentX        =   10186
      _ExtentY        =   18230
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkPoliza 
         Height          =   375
         Left            =   3720
         TabIndex        =   70
         Top             =   9000
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Suscribe Póliza  "
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   315
         Left            =   1800
         TabIndex        =   26
         Top             =   480
         Width           =   2055
         _Version        =   1310723
         _ExtentX        =   3619
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIdAlterno 
         Height          =   315
         Left            =   1800
         TabIndex        =   28
         Top             =   840
         Width           =   2055
         _Version        =   1310723
         _ExtentX        =   3619
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail_01 
         Height          =   315
         Left            =   120
         TabIndex        =   47
         Top             =   4800
         Width           =   5415
         _Version        =   1310723
         _ExtentX        =   9551
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.FlatEdit txtEmail_02 
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   5400
         Width           =   5415
         _Version        =   1310723
         _ExtentX        =   9551
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.FlatEdit txtDireccion 
         Height          =   555
         Left            =   120
         TabIndex        =   51
         Top             =   6600
         Width           =   5415
         _Version        =   1310723
         _ExtentX        =   9551
         _ExtentY        =   979
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnResolucion 
         Height          =   615
         Left            =   3240
         TabIndex        =   57
         Top             =   9480
         Width           =   2295
         _Version        =   1310723
         _ExtentX        =   4043
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Aplicar Resolución"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmSYS_Monitor_Afiliacion_EnLinea.frx":0000
         ImageAlignment  =   0
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.ComboBox cboResolucion 
         Height          =   330
         Left            =   120
         TabIndex        =   58
         Top             =   9600
         Width           =   2895
         _Version        =   1310723
         _ExtentX        =   5106
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido1 
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido2 
         Height          =   315
         Left            =   1920
         TabIndex        =   33
         Top             =   2160
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   315
         Left            =   3720
         TabIndex        =   35
         Top             =   2160
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEstadoCivil 
         Height          =   315
         Left            =   120
         TabIndex        =   61
         Top             =   3000
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtGenero 
         Height          =   315
         Left            =   1920
         TabIndex        =   62
         Top             =   3000
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelMovil 
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   3960
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelHab 
         Height          =   315
         Left            =   1920
         TabIndex        =   43
         Top             =   3960
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.FlatEdit txtTelTrabajo 
         Height          =   315
         Left            =   3720
         TabIndex        =   45
         Top             =   3960
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtProvincia 
         Height          =   315
         Left            =   120
         TabIndex        =   53
         Top             =   6240
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCanton 
         Height          =   315
         Left            =   1920
         TabIndex        =   54
         Top             =   6240
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDistrito 
         Height          =   315
         Left            =   3720
         TabIndex        =   55
         Top             =   6240
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEmpresa 
         Height          =   315
         Left            =   120
         TabIndex        =   64
         Top             =   7560
         Width           =   5415
         _Version        =   1310723
         _ExtentX        =   9551
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.FlatEdit txtFechaNac 
         Height          =   315
         Left            =   1800
         TabIndex        =   65
         Top             =   1200
         Width           =   2055
         _Version        =   1310723
         _ExtentX        =   3625
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTramite 
         Height          =   510
         Left            =   3960
         TabIndex        =   66
         Top             =   480
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
         _ExtentY        =   900
         _StockProps     =   77
         ForeColor       =   0
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   510
         Left            =   3960
         TabIndex        =   67
         Top             =   1080
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
         _ExtentY        =   900
         _StockProps     =   77
         ForeColor       =   0
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNacionalidad 
         Height          =   315
         Left            =   3720
         TabIndex        =   63
         Top             =   3000
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFechaIngresoEmpresa 
         Height          =   315
         Left            =   3720
         TabIndex        =   69
         Top             =   7920
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtReferencia 
         Height          =   315
         Left            =   120
         TabIndex        =   71
         Top             =   8640
         Width           =   5415
         _Version        =   1310723
         _ExtentX        =   9551
         _ExtentY        =   556
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
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Persona Referida"
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
         Left            =   120
         TabIndex        =   72
         Top             =   8400
         Width           =   2295
      End
      Begin XtremeSuiteControls.Label Label2x 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   68
         Top             =   7920
         Width           =   3135
         _Version        =   1310723
         _ExtentX        =   5530
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha Ingreso a la Empresa:"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblResolucion 
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   9360
         Width           =   1935
         _Version        =   1310723
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Resolución"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2x 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   56
         Top             =   7320
         Width           =   1335
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Empresa"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
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
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   6000
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Email No.1"
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
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Email No.2"
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
         Left            =   120
         TabIndex        =   49
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. Trabajo"
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
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   46
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. Habitación"
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
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   44
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. Móvil"
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   42
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nacionalidad"
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
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   40
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nacimiento"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Genero"
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
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   38
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Civil"
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
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
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
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   36
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido No.2"
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
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   34
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido No.1"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Width           =   1575
      End
      Begin XtremeShortcutBar.ShortcutCaption scCaso 
         Height          =   372
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   10812
         _Version        =   1310723
         _ExtentX        =   19071
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Selecciones un Caso!"
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
      End
      Begin XtremeSuiteControls.Label Label2x 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1335
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Id Empleado"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2x 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1335
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Identificación"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   3000
      Top             =   120
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1926
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Buscar"
      BackColor       =   -2147483633
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
      Picture         =   "frmSYS_Monitor_Afiliacion_EnLinea.frx":0805
      TextImageRelation=   1
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   735
      Left            =   1680
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1926
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Exportar"
      BackColor       =   -2147483633
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
      Picture         =   "frmSYS_Monitor_Afiliacion_EnLinea.frx":1223
      TextImageRelation=   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.ComboBox cboFecha 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
      _Version        =   1310723
      _ExtentX        =   2990
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
      _Version        =   1310723
      _ExtentX        =   2990
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txt_IdAlterno 
      Height          =   330
      Left            =   1080
      TabIndex        =   8
      Top             =   840
      Width           =   1695
      _Version        =   1310723
      _ExtentX        =   2984
      _ExtentY        =   593
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
   Begin XtremeSuiteControls.FlatEdit txt_Cedula 
      Height          =   330
      Left            =   1080
      TabIndex        =   9
      Top             =   360
      Width           =   1695
      _Version        =   1310723
      _ExtentX        =   2984
      _ExtentY        =   593
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
   Begin XtremeSuiteControls.FlatEdit txtEA_Casos 
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   5400
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   556
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtED_Casos 
      Height          =   315
      Left            =   1320
      TabIndex        =   11
      Top             =   6000
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   556
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtEP_Casos 
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Top             =   6600
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   556
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtES_Casos 
      Height          =   315
      Left            =   1320
      TabIndex        =   13
      Top             =   7200
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   556
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   9975
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   8775
      _Version        =   524288
      _ExtentX        =   15478
      _ExtentY        =   17595
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
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
      MaxCols         =   27
      SpreadDesigner  =   "frmSYS_Monitor_Afiliacion_EnLinea.frx":1A28
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txt_Nombre 
      Height          =   330
      Left            =   1080
      TabIndex        =   59
      Top             =   1320
      Width           =   1695
      _Version        =   1310723
      _ExtentX        =   2984
      _ExtentY        =   593
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
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
      Left            =   240
      TabIndex        =   60
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas:"
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
      Left            =   240
      TabIndex        =   24
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado:"
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
      Index           =   3
      Left            =   240
      TabIndex        =   23
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   480
      TabIndex        =   22
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   480
      TabIndex        =   21
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Id Alterno:"
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
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula:"
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
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Resumen de Estado de Solicitudes:"
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
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Aprobadas:"
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
      Index           =   7
      Left            =   360
      TabIndex        =   17
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Denegadas:"
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
      Index           =   10
      Left            =   360
      TabIndex        =   16
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pendientes:"
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
      Index           =   11
      Left            =   360
      TabIndex        =   15
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitadas:"
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
      Index           =   12
      Left            =   360
      TabIndex        =   14
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image imgMainBanner 
      Height          =   9990
      Left            =   0
      Picture         =   "frmSYS_Monitor_Afiliacion_EnLinea.frx":2673
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3210
   End
End
Attribute VB_Name = "frmSYS_Monitor_Afiliacion_EnLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mFecUltMovUpdate As Integer
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnExportar_Click()
Call sbExportar
End Sub



Private Sub sbExportar()
Dim vHeaders As vGridHeaders

    
  vHeaders.Columnas = 26
  vHeaders.Headers(1) = "Tramite No."
  vHeaders.Headers(2) = "Estado"
  vHeaders.Headers(3) = "Identificacion"
  vHeaders.Headers(4) = "Id Alterno"
  vHeaders.Headers(5) = "Apellido No.1"
  vHeaders.Headers(6) = "Apellido No.2"
  vHeaders.Headers(7) = "Nombre"
  vHeaders.Headers(8) = "Fecha Nac."
  vHeaders.Headers(19) = "Estado Civil"
  vHeaders.Headers(10) = "Género"
  vHeaders.Headers(11) = "Nacionalidad"
  vHeaders.Headers(12) = "Empresa"
  vHeaders.Headers(13) = "Tel.Móvil"
  vHeaders.Headers(14) = "Tel.Habitación"
  vHeaders.Headers(15) = "Tel.Trabajo"
  vHeaders.Headers(16) = "Email No.1"
  vHeaders.Headers(17) = "Email No.2"
  vHeaders.Headers(18) = "Provincia"
  vHeaders.Headers(19) = "Cantón"
  vHeaders.Headers(20) = "Distrito"
  vHeaders.Headers(21) = "Dirección"
  vHeaders.Headers(22) = "Fecha Registro"
  vHeaders.Headers(23) = "Fecha Gestión"
  vHeaders.Headers(24) = "Usuario Gestiona"
  vHeaders.Headers(25) = "Póliza?"
  vHeaders.Headers(26) = "Deducc.Planilla"
  
    
    Call sbSIFGridExportar(vGrid, vHeaders, "Afiliacion_EnLiena_Consulta")


End Sub

Private Sub sbFiltro_Aplica(ByRef pSQL As String)
Dim pWhere As Boolean

pWhere = False

If cboEstado.Text <> "Todos" Then
   pSQL = pSQL & " Where Estado = '" & Mid(cboEstado.Text, 1, 1) & "'"
   pWhere = True
End If


Select Case cboFecha.Text
  Case "Registro"
      If pWhere Then
            pSQL = pSQL & " and Registro_Fecha between  '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                    & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
      Else
            pSQL = pSQL & " Where Registro_Fecha between  '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                    & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

            pWhere = True
      End If

  Case "Resolución"
      If pWhere Then
            pSQL = pSQL & " and RESUELTO_FECHA between  '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                    & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
      Else
            pSQL = pSQL & " Where RESUELTO_FECHA between  '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                    & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

            pWhere = True
      End If

End Select



If Trim(txt_Cedula.Text) <> "" Then
   If pWhere Then
        pSQL = pSQL & " and Cedula = '" & txt_Cedula.Text & "'"
   Else
        pSQL = pSQL & " Where Cedula = '" & txt_Cedula.Text & "'"

        pWhere = True

   End If
End If


If Trim(txt_IdAlterno.Text) <> "" Then
   If pWhere Then
        pSQL = pSQL & " and ID_COLILLA = '" & txt_IdAlterno.Text & "'"
   Else
        pSQL = pSQL & " Where ID_COLILLA = '" & txt_IdAlterno.Text & "'"

        pWhere = True

   End If
End If


If Trim(txt_Nombre.Text) <> "" Then
   If pWhere Then
        pSQL = pSQL & " and (APELLIDO_1 + ' ' + APELLIDO_2 + ' ' + NOMBRE_1 + ' ' + NOMBRE_2) like '%" & txt_IdAlterno.Text & "%'"
   Else
        pSQL = pSQL & " Where (APELLIDO_1 + ' ' + APELLIDO_2 + ' ' + NOMBRE_1 + ' ' + NOMBRE_2) like '%" & txt_IdAlterno.Text & "%'"

        pWhere = True

   End If
End If

End Sub

Private Sub sbBuscar()
Dim vCadena As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select '', SOLICITUD_ID, Estado_Desc, CEDULA, ID_COLILLA, APELLIDO_1, APELLIDO_2, NOMBRE_1 + ' ' + NOMBRE_2 AS 'NOMBRE'" _
       & ", CONVERT(VARCHAR(10), FECHA_NACIMIENTO, 23) AS 'FECHA_NAC', EstadoCivil_Desc, Sexo_Desc, Nacionalidad_Desc" _
       & ", Institucion_Desc, TEL_MOVIL, TEL_HABITACION , TEL_TRABAJO, EMAIL_01, EMAIL_02" _
       & ", Provincia_Desc, Canton_Desc, Distrito_Desc, DIRECCION" _
       & ", REGISTRO_FECHA, RESUELTO_FECHA, RESUELTO_USUARIO, I_POLIZA_VIDA_FAMILIAR, I_AUTORIZACION_DEDUC" _
       & " From vAFI_Afiliacion_EnLinea"
Call sbFiltro_Aplica(strSQL)

strSQL = strSQL & " order by SOLICITUD_ID"

vPaso = True

Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL, True)
If vGrid.MaxRows > 1 Then
    vGrid.MaxRows = vGrid.MaxRows - 1
End If

vPaso = False

' Proceso Informacion de Resumen
txtEP_Casos.Text = Format(0, "###,###0")

txtES_Casos.Text = Format(0, "###,###0")

txtEA_Casos.Text = Format(0, "###,###0")

txtED_Casos.Text = Format(0, "###,###0")


strSQL = "exec spAFI_Afiliacion_EnLinea_Resumen '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
       & ",'" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Select Case rs!ESTADO
    Case "P"
        txtEP_Casos.Text = Format(rs!Casos, "###,###0")

    Case "S"
        txtES_Casos.Text = Format(rs!Casos, "###,###0")
    
    Case "A"
        txtEA_Casos.Text = Format(rs!Casos, "###,###0")
    
    Case "D"
        txtED_Casos.Text = Format(rs!Casos, "###,###0")

 End Select
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbLimpia()

scCaso.Tag = 0
scCaso.Caption = "Seleccione un Caso!"

txtTramite.Text = ""
txtEstado.Text = ""
txtEstado.Tag = ""

txtCedula.Text = ""
txtIdAlterno.Text = ""
txtApellido1.Text = ""
txtApellido2.Text = ""
txtNombre.Text = ""
txtFechaNac.Text = ""
txtEstadoCivil.Text = ""
txtGenero.Text = ""
txtNacionalidad.Text = ""
txtTelMovil.Text = ""
txtTelHab.Text = ""
txtTelTrabajo.Text = ""
txtEmail_01.Text = ""
txtEmail_02.Text = ""
txtProvincia.Text = ""
txtCanton.Text = ""
txtDistrito.Text = ""
txtDireccion.Text = ""
txtEmpresa.Text = ""
txtFechaIngresoEmpresa.Text = ""

Call sbResolucion_Load

End Sub

Private Sub sbCaso_Consulta(pCaso As Long)
Dim vCadena As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbLimpia

strSQL = "select *, APELLIDO_1+ ' ' +  APELLIDO_2+ ' ' +  NOMBRE_1 + ' ' + NOMBRE_2 as 'NOMBRE_COMPLETO' " _
       & " From vAFI_Afiliacion_EnLinea " _
       & " Where SOLICITUD_ID = " & pCaso
Call OpenRecordSet(rs, strSQL)
'I_POlIZA_VIDA_FAMILIAR, REFERENCIA
scCaso.Tag = rs!SOLICITUD_ID

scCaso.Caption = "Id: " & rs!SOLICITUD_ID & "  " & rs!Nombre_Completo & ""

txtTramite.Text = rs!SOLICITUD_ID
txtEstado.Text = rs!Estado_Desc
txtEstado.Tag = rs!ESTADO


txtCedula.Text = rs!CEDULA
txtIdAlterno.Text = rs!ID_COLILLA
txtApellido1.Text = rs!APELLIDO_1
txtApellido2.Text = rs!APELLIDO_2
txtNombre.Text = rs!Nombre_1
txtFechaNac.Text = Format(rs!FECHA_NACIMIENTO, "dd/MM/yyyy")
txtEstadoCivil.Text = rs!EstadoCivil_Desc
txtGenero.Text = rs!Sexo_Desc
txtNacionalidad.Text = rs!Nacionalidad_Desc
txtTelMovil.Text = rs!TEL_MOVIL
txtTelHab.Text = rs!TEL_HABITACION
txtTelTrabajo.Text = rs!TEL_TRABAJO
txtEmail_01.Text = rs!EMAIL_01
txtEmail_02.Text = rs!EMAIL_02
txtProvincia.Text = rs!Provincia_Desc
txtCanton.Text = rs!Canton_Desc
txtDistrito.Text = rs!Distrito_Desc
txtDireccion.Text = rs!DIRECCION
txtEmpresa.Text = rs!Institucion_Desc

txtFechaIngresoEmpresa.Text = Format(rs!FECHA_INGRESO_LABORAL, "dd/MM/yyyy")

chkPoliza.Value = rs!I_POLIZA_VIDA_FAMILIAR

rs.Close
Me.MousePointer = vbDefault


Call sbResolucion_Load

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnResolucion_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spAFI_Afiliacion_EnLinea_Resolucion " & scCaso.Tag & ",'" & Mid(cboResolucion.Text, 1, 1) _
            & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

If Mid(cboResolucion.Text, 1, 1) = "A" Then
   MsgBox "La Solicitud de Afiliación fue aprobada y procesada!", vbInformation
End If

Call sbCaso_Consulta(scCaso.Tag)


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub cboFecha_Click()
If vPaso Then Exit Sub

If cboFecha.Text = "Todas" Then
   dtpInicio.Enabled = False
   dtpCorte.Enabled = False
Else
   dtpInicio.Enabled = True
   dtpCorte.Enabled = True
End If

End Sub



Private Sub txt_Cedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Cédula Colilla"
    gBusquedas.Col2Name = "Cédula Real"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    txt_Cedula.Text = Trim(gBusquedas.Resultado)
End If

End Sub

Private Sub Form_Activate()
vModulo = 1

End Sub

Private Sub Form_Load()

vModulo = 1
vGrid.AppearanceStyle = fxGridStyle



End Sub



Private Sub Form_Resize()
Dim pHeight As Long, pWidth As Long

On Error Resume Next


pHeight = 9975
pWidth = 18120


If Me.Height < pHeight Then
   Me.Height = pHeight
End If

If Me.Width < pWidth Then
   Me.Width = pWidth
End If

imgMainBanner.Height = Me.Height
gbCaso.Height = Me.Height

vGrid.Width = Me.Width - (vGrid.Left + gbCaso.Width + 160)
vGrid.Height = Me.Height - (vGrid.Top + 500)

gbCaso.Left = vGrid.Left + vGrid.Width + 60


End Sub



Private Sub sbResolucion_Load()

cboResolucion.Text = txtEstado.Text

Select Case Mid(txtEstado.Text, 1, 1)
    Case "R", "P", "S"
        cboResolucion.Visible = True
        btnResolucion.Visible = True
        
    Case "A", "D", ""
        cboResolucion.Visible = False
        btnResolucion.Visible = False
End Select

End Sub


Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

On Error GoTo vError

vPaso = True

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -15, dtpCorte.Value)

cboFecha.Clear
cboFecha.AddItem "Registro"
cboFecha.AddItem "Resolución"
cboFecha.AddItem "Todas"
cboFecha.Text = "Registro"

cboEstado.Clear
cboEstado.AddItem "Todos"
cboEstado.AddItem "Solicitada"
cboEstado.AddItem "Pendiente"
cboEstado.AddItem "Aprobada"
cboEstado.AddItem "Denegada"

cboEstado.Text = "Solicitada"


cboResolucion.Clear
cboResolucion.AddItem "Solicitada"
cboResolucion.AddItem "Pendiente"
cboResolucion.AddItem "Aprobada"
cboResolucion.AddItem "Denegada"

cboResolucion.Text = "Solicitada"

vPaso = False

Call cboFecha_Click

Call Formularios(Me)
Call RefrescaTags(Me)

Call sbLimpia
Call sbBuscar

Exit Sub

vError:

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

Dim vTramite As Long

On Error GoTo vError

vGrid.Row = Row
vGrid.Col = 2

vTramite = vGrid.Text

scCaso.Tag = ""
scCaso.Caption = "Indique un Caso!"

vGrid.Col = 3
scCaso.Tag = vTramite
scCaso.Caption = "No. " & vTramite & " ¦ " & vGrid.Text

Call sbCaso_Consulta(vTramite)

Exit Sub

vError:

End Sub




