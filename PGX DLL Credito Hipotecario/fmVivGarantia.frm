VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmVivGarantia 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información de Garantía"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   10470
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7335
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   10215
      _Version        =   1572864
      _ExtentX        =   18018
      _ExtentY        =   12938
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
      Color           =   128
      ItemCount       =   4
      Item(0).Caption =   "General"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lvwGarantias"
      Item(1).Caption =   "Garantía"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "GroupBox1(0)"
      Item(1).Control(1)=   "GroupBox1(1)"
      Item(1).Control(2)=   "tcAux"
      Item(2).Caption =   "Derechos"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "GroupBox2"
      Item(3).Caption =   "Historial del Trámite"
      Item(3).ControlCount=   54
      Item(3).Control(0)=   "Label3(22)"
      Item(3).Control(1)=   "Label3(23)"
      Item(3).Control(2)=   "Label3(24)"
      Item(3).Control(3)=   "lblFechaRegistro"
      Item(3).Control(4)=   "lblEstadoActual"
      Item(3).Control(5)=   "lblUsuarioRegistro"
      Item(3).Control(6)=   "ShortcutCaption1(0)"
      Item(3).Control(7)=   "ShortcutCaption1(1)"
      Item(3).Control(8)=   "Label3(25)"
      Item(3).Control(9)=   "Label3(26)"
      Item(3).Control(10)=   "Label3(27)"
      Item(3).Control(11)=   "Label3(28)"
      Item(3).Control(12)=   "Label3(29)"
      Item(3).Control(13)=   "Label3(30)"
      Item(3).Control(14)=   "Label3(31)"
      Item(3).Control(15)=   "Label3(32)"
      Item(3).Control(16)=   "Label3(33)"
      Item(3).Control(17)=   "Label3(34)"
      Item(3).Control(18)=   "Label3(35)"
      Item(3).Control(19)=   "Label3(36)"
      Item(3).Control(20)=   "Label3(37)"
      Item(3).Control(21)=   "Label3(38)"
      Item(3).Control(22)=   "Label3(39)"
      Item(3).Control(23)=   "Label3(40)"
      Item(3).Control(24)=   "Label3(41)"
      Item(3).Control(25)=   "Label3(42)"
      Item(3).Control(26)=   "Label3(43)"
      Item(3).Control(27)=   "Label3(44)"
      Item(3).Control(28)=   "Label3(45)"
      Item(3).Control(29)=   "Label3(46)"
      Item(3).Control(30)=   "Label3(47)"
      Item(3).Control(31)=   "Label3(48)"
      Item(3).Control(32)=   "Label3(49)"
      Item(3).Control(33)=   "Label3(50)"
      Item(3).Control(34)=   "lblNombreIng"
      Item(3).Control(35)=   "lblNombreAbog"
      Item(3).Control(36)=   "lblEstadoIng"
      Item(3).Control(37)=   "lblEstadoAbog"
      Item(3).Control(38)=   "lblAsignacionFecha"
      Item(3).Control(39)=   "lblAsignacionFechaAbog"
      Item(3).Control(40)=   "lblAsignacionUsuario"
      Item(3).Control(41)=   "lblAsignacionUsuarioAbog"
      Item(3).Control(42)=   "lblEntregaFecha"
      Item(3).Control(43)=   "lblEntregaFechaAbog"
      Item(3).Control(44)=   "lblEntregaUsuario"
      Item(3).Control(45)=   "lblEntregaUsuarioAbog"
      Item(3).Control(46)=   "lblRecepcionFecha"
      Item(3).Control(47)=   "lblFirmasFecha"
      Item(3).Control(48)=   "lblRecepcionUsuario"
      Item(3).Control(49)=   "lblFirmasUsuario"
      Item(3).Control(50)=   "lblRegistroFecha"
      Item(3).Control(51)=   "lblRegistroFechaAbog"
      Item(3).Control(52)=   "lblRegistroUsuario"
      Item(3).Control(53)=   "lblRegistroUsuarioAbog"
      Begin XtremeSuiteControls.ListView lvwGarantias 
         Height          =   6735
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   9975
         _Version        =   1572864
         _ExtentX        =   17595
         _ExtentY        =   11880
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
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   6975
         Left            =   -70000
         TabIndex        =   73
         Top             =   360
         Visible         =   0   'False
         Width           =   10215
         _Version        =   1572864
         _ExtentX        =   18018
         _ExtentY        =   12303
         _StockProps     =   79
         Caption         =   "Información de Dueños"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lvwListaDuenos 
            Height          =   3735
            Left            =   960
            TabIndex        =   74
            Top             =   3120
            Width           =   8415
            _Version        =   1572864
            _ExtentX        =   14843
            _ExtentY        =   6588
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
         End
         Begin XtremeSuiteControls.ComboBox cboProvinciasDuenos 
            Height          =   330
            Left            =   960
            TabIndex        =   75
            Top             =   1320
            Width           =   2655
            _Version        =   1572864
            _ExtentX        =   4683
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
         End
         Begin XtremeSuiteControls.ComboBox cboCantonesDueno 
            Height          =   330
            Left            =   3720
            TabIndex        =   76
            Top             =   1320
            Width           =   2775
            _Version        =   1572864
            _ExtentX        =   4895
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
         End
         Begin XtremeSuiteControls.ComboBox cboDistritosDuenos 
            Height          =   330
            Left            =   6600
            TabIndex        =   77
            Top             =   1320
            Width           =   2775
            _Version        =   1572864
            _ExtentX        =   4895
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
         End
         Begin XtremeSuiteControls.FlatEdit txtDireccionDueno 
            Height          =   735
            Left            =   960
            TabIndex        =   78
            Top             =   1680
            Width           =   8415
            _Version        =   1572864
            _ExtentX        =   14843
            _ExtentY        =   1296
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
         Begin XtremeSuiteControls.FlatEdit txtCedulaDueno 
            Height          =   315
            Left            =   960
            TabIndex        =   79
            Top             =   600
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3619
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
         Begin XtremeSuiteControls.FlatEdit txtNombreDueno 
            Height          =   315
            Left            =   3000
            TabIndex        =   80
            Top             =   600
            Width           =   6375
            _Version        =   1572864
            _ExtentX        =   11245
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   2
            Left            =   960
            TabIndex        =   137
            Top             =   2640
            Width           =   8415
            _Version        =   1572864
            _ExtentX        =   14843
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Lista de Dueños Registrados:"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   21
            Left            =   6600
            TabIndex        =   85
            Top             =   1080
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Distrito"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   20
            Left            =   3720
            TabIndex        =   84
            Top             =   1080
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cantón"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   19
            Left            =   960
            TabIndex        =   83
            Top             =   1080
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Provincia"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   18
            Left            =   3000
            TabIndex        =   82
            Top             =   360
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Nombre"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   17
            Left            =   960
            TabIndex        =   81
            Top             =   360
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cédula"
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
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   3255
         Left            =   -70000
         TabIndex        =   43
         Top             =   4080
         Visible         =   0   'False
         Width           =   10215
         _Version        =   1572864
         _ExtentX        =   18018
         _ExtentY        =   5741
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Color           =   32
         ItemCount       =   2
         Item(0).Caption =   "Anotaciones"
         Item(0).ControlCount=   7
         Item(0).Control(0)=   "txtObservaciones"
         Item(0).Control(1)=   "Label3(14)"
         Item(0).Control(2)=   "btnHonorariosDetalle"
         Item(0).Control(3)=   "optObservacion(0)"
         Item(0).Control(4)=   "optObservacion(1)"
         Item(0).Control(5)=   "optObservacion(2)"
         Item(0).Control(6)=   "txtMontoNoGravable"
         Item(1).Caption =   "Avalúo Posterior"
         Item(1).ControlCount=   21
         Item(1).Control(0)=   "Label10(17)"
         Item(1).Control(1)=   "Label10(16)"
         Item(1).Control(2)=   "Label10(15)"
         Item(1).Control(3)=   "Label10(13)"
         Item(1).Control(4)=   "Label10(12)"
         Item(1).Control(5)=   "txtTotal"
         Item(1).Control(6)=   "txtValorConstruccion"
         Item(1).Control(7)=   "txtValorTerreno"
         Item(1).Control(8)=   "dtpFechaInspeccion"
         Item(1).Control(9)=   "txtViaticos"
         Item(1).Control(10)=   "txtAvaluo_Notas"
         Item(1).Control(11)=   "Label10(14)"
         Item(1).Control(12)=   "optPersonal"
         Item(1).Control(13)=   "optComercial"
         Item(1).Control(14)=   "Label10(18)"
         Item(1).Control(15)=   "txtIdIngeniero"
         Item(1).Control(16)=   "txtNombreIngeniero"
         Item(1).Control(17)=   "Label3(15)"
         Item(1).Control(18)=   "txtIdAbogado"
         Item(1).Control(19)=   "txtNombreAbogado"
         Item(1).Control(20)=   "Label3(16)"
         Begin XtremeSuiteControls.FlatEdit txtNombreAbogado 
            Height          =   315
            Left            =   -66640
            TabIndex        =   71
            Top             =   2880
            Visible         =   0   'False
            Width           =   6735
            _Version        =   1572864
            _ExtentX        =   11880
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.RadioButton optObservacion 
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   48
            Top             =   480
            Width           =   3135
            _Version        =   1572864
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Anotaciones sobre la finca"
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
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtObservaciones 
            Height          =   1215
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   9975
            _Version        =   1572864
            _ExtentX        =   17595
            _ExtentY        =   2143
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
         Begin XtremeSuiteControls.PushButton btnHonorariosDetalle 
            Height          =   435
            Left            =   7440
            TabIndex        =   47
            Top             =   2640
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Detalle de Honorarios"
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
            Picture         =   "fmVivGarantia.frx":0000
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.RadioButton optObservacion 
            Height          =   375
            Index           =   1
            Left            =   3840
            TabIndex        =   49
            Top             =   480
            Width           =   3135
            _Version        =   1572864
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Gravámenes"
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
         End
         Begin XtremeSuiteControls.RadioButton optObservacion 
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   50
            Top             =   480
            Width           =   3015
            _Version        =   1572864
            _ExtentX        =   5318
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Notas del gravámen"
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
         End
         Begin XtremeSuiteControls.FlatEdit txtMontoNoGravable 
            Height          =   315
            Left            =   7440
            TabIndex        =   46
            Top             =   2280
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTotal 
            Height          =   315
            Left            =   -67960
            TabIndex        =   57
            Top             =   2040
            Visible         =   0   'False
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
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
            Text            =   "0.00"
            BackColor       =   16777152
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtValorConstruccion 
            Height          =   315
            Left            =   -67960
            TabIndex        =   58
            Top             =   1680
            Visible         =   0   'False
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
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
            Text            =   "0.00"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtValorTerreno 
            Height          =   315
            Left            =   -67960
            TabIndex        =   59
            Top             =   1320
            Visible         =   0   'False
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
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
            Text            =   "0.00"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaInspeccion 
            Height          =   315
            Left            =   -67960
            TabIndex        =   60
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtViaticos 
            Height          =   315
            Left            =   -67960
            TabIndex        =   61
            Top             =   840
            Visible         =   0   'False
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
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
            Text            =   "0.00"
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAvaluo_Notas 
            Height          =   975
            Left            =   -65440
            TabIndex        =   62
            Top             =   840
            Visible         =   0   'False
            Width           =   5535
            _Version        =   1572864
            _ExtentX        =   9763
            _ExtentY        =   1720
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
         Begin XtremeSuiteControls.RadioButton optPersonal 
            Height          =   255
            Left            =   -63760
            TabIndex        =   64
            Top             =   2040
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Personal"
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
         End
         Begin XtremeSuiteControls.RadioButton optComercial 
            Height          =   255
            Left            =   -62080
            TabIndex        =   65
            Top             =   2040
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Comercial"
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
         End
         Begin XtremeSuiteControls.FlatEdit txtIdIngeniero 
            Height          =   315
            Left            =   -67960
            TabIndex        =   67
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   2520
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   556
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNombreIngeniero 
            Height          =   315
            Left            =   -66640
            TabIndex        =   68
            Top             =   2520
            Visible         =   0   'False
            Width           =   6735
            _Version        =   1572864
            _ExtentX        =   11880
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtIdAbogado 
            Height          =   315
            Left            =   -67960
            TabIndex        =   70
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   2880
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   556
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   16
            Left            =   -69880
            TabIndex        =   72
            Top             =   2880
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Abogado"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   15
            Left            =   -69880
            TabIndex        =   69
            Top             =   2520
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Ingeniero"
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
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   18
            Left            =   -65440
            TabIndex        =   66
            Top             =   2040
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo de Póliza"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   14
            Left            =   -65440
            TabIndex        =   63
            Top             =   480
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Observaciones:"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   12
            Left            =   -69880
            TabIndex        =   56
            Top             =   840
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Viáticos"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   13
            Left            =   -69880
            TabIndex        =   55
            Top             =   480
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha de Inspección"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   15
            Left            =   -69880
            TabIndex        =   54
            Top             =   1320
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Valor del terreno"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   16
            Left            =   -69880
            TabIndex        =   53
            Top             =   1680
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Valor contrucción"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   17
            Left            =   -69880
            TabIndex        =   52
            Top             =   2040
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Valor total inmueble"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   14
            Left            =   5280
            TabIndex        =   45
            Top             =   2280
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Monto No Gravable:"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   3495
         Index           =   0
         Left            =   -69880
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1572864
         _ExtentX        =   8705
         _ExtentY        =   6165
         _StockProps     =   79
         Caption         =   "Información de la Propiedad"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkCoberturaPrimerGrado 
            Height          =   255
            Left            =   480
            TabIndex        =   27
            Top             =   2280
            Width           =   3855
            _Version        =   1572864
            _ExtentX        =   6800
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Aplica cobertura en primer grado"
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
            Appearance      =   17
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton btnGradoHipoteca 
            Height          =   315
            Left            =   4440
            TabIndex        =   26
            Top             =   1800
            Width           =   375
            _Version        =   1572864
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.FlatEdit txtNumeroFinca 
            Height          =   315
            Left            =   2160
            TabIndex        =   21
            Top             =   360
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtTipoDerecho 
            Height          =   315
            Left            =   2160
            TabIndex        =   22
            Top             =   720
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtNumPlanoCatastro 
            Height          =   315
            Left            =   2160
            TabIndex        =   23
            Top             =   1080
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtAreaFinca 
            Height          =   315
            Left            =   2160
            TabIndex        =   24
            Top             =   1440
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.ComboBox cboGradoHipoteca 
            Height          =   330
            Left            =   2160
            TabIndex        =   25
            Top             =   1800
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
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
         End
         Begin XtremeSuiteControls.CheckBox chkRegistraCalAvaluo 
            Height          =   255
            Left            =   480
            TabIndex        =   28
            Top             =   2520
            Width           =   3855
            _Version        =   1572864
            _ExtentX        =   6800
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Registar Cálculo de Avalúo"
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
            Appearance      =   17
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkRegistraCalHonorarios 
            Height          =   255
            Left            =   480
            TabIndex        =   29
            Top             =   2760
            Width           =   3855
            _Version        =   1572864
            _ExtentX        =   6800
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Registar Cálculo de Honorarios"
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
            Appearance      =   17
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkDetalleHonorarios 
            Height          =   255
            Left            =   480
            TabIndex        =   30
            Top             =   3120
            Width           =   3855
            _Version        =   1572864
            _ExtentX        =   6800
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Registar detalle manual"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   20
            Top             =   1800
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Grado Hipoteca"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Área (m2)"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "No. plano catastro"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo Derecho"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "No. Finca"
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
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   3495
         Index           =   1
         Left            =   -64840
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1572864
         _ExtentX        =   8705
         _ExtentY        =   6165
         _StockProps     =   79
         Caption         =   "Dirección de la Propiedad"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboUbicacionProvincia 
            Height          =   315
            Left            =   1200
            TabIndex        =   37
            Top             =   360
            Width           =   3495
            _Version        =   1572864
            _ExtentX        =   6165
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
         End
         Begin XtremeSuiteControls.ComboBox cboUbicacionCanton 
            Height          =   315
            Left            =   1200
            TabIndex        =   38
            Top             =   720
            Width           =   3495
            _Version        =   1572864
            _ExtentX        =   6165
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
         End
         Begin XtremeSuiteControls.ComboBox cboUbicacionDistrito 
            Height          =   315
            Left            =   1200
            TabIndex        =   39
            Top             =   1080
            Width           =   3495
            _Version        =   1572864
            _ExtentX        =   6165
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
         End
         Begin XtremeSuiteControls.FlatEdit txtDireccion 
            Height          =   1095
            Left            =   1200
            TabIndex        =   40
            Top             =   1440
            Width           =   3495
            _Version        =   1572864
            _ExtentX        =   6165
            _ExtentY        =   1931
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
         Begin XtremeSuiteControls.ComboBox cboZonas 
            Height          =   315
            Left            =   1200
            TabIndex        =   41
            Top             =   2640
            Width           =   3495
            _Version        =   1572864
            _ExtentX        =   6165
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
         End
         Begin XtremeSuiteControls.ComboBox cboTipo_Poliza 
            Height          =   315
            Left            =   1200
            TabIndex        =   42
            Top             =   3000
            Width           =   3495
            _Version        =   1572864
            _ExtentX        =   6165
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
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   36
            Top             =   3000
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Póliza"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   35
            Top             =   2640
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Zona"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Dirección"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   33
            Top             =   1080
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Distrito"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cantón"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Provincia"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   50
         Left            =   -63760
         TabIndex        =   7
         Top             =   6840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   49
         Left            =   -63760
         TabIndex        =   12
         Top             =   6480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   48
         Left            =   -63760
         TabIndex        =   51
         Top             =   5640
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   47
         Left            =   -63760
         TabIndex        =   136
         Top             =   5280
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   46
         Left            =   -63760
         TabIndex        =   135
         Top             =   4560
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   45
         Left            =   -63760
         TabIndex        =   134
         Top             =   4200
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   44
         Left            =   -63760
         TabIndex        =   133
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   43
         Left            =   -63760
         TabIndex        =   132
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   42
         Left            =   -64600
         TabIndex        =   131
         Top             =   6120
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Inscripción garantía"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   41
         Left            =   -64600
         TabIndex        =   130
         Top             =   4920
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Registro de Firmas"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   40
         Left            =   -64600
         TabIndex        =   129
         Top             =   3840
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Entrega documentos"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   39
         Left            =   -64600
         TabIndex        =   128
         Top             =   2760
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Asignación"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   38
         Left            =   -64600
         TabIndex        =   127
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estado del Proceso"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   37
         Left            =   -69040
         TabIndex        =   126
         Top             =   6840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   36
         Left            =   -69040
         TabIndex        =   125
         Top             =   6480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   35
         Left            =   -69040
         TabIndex        =   124
         Top             =   5640
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   34
         Left            =   -69040
         TabIndex        =   123
         Top             =   5280
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   33
         Left            =   -69040
         TabIndex        =   122
         Top             =   4560
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   32
         Left            =   -69040
         TabIndex        =   121
         Top             =   4200
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   31
         Left            =   -69040
         TabIndex        =   120
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   30
         Left            =   -69040
         TabIndex        =   119
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   29
         Left            =   -69880
         TabIndex        =   118
         Top             =   6120
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Registro Avalúo"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   28
         Left            =   -69880
         TabIndex        =   117
         Top             =   4920
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Recepción documentos"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   27
         Left            =   -69880
         TabIndex        =   116
         Top             =   3840
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Entrega documentos"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   26
         Left            =   -69880
         TabIndex        =   115
         Top             =   2760
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Asignación"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   25
         Left            =   -69880
         TabIndex        =   114
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estado del Proceso"
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
      Begin XtremeSuiteControls.Label lblRegistroUsuarioAbog 
         Height          =   315
         Left            =   -62320
         TabIndex        =   113
         Top             =   6840
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblRegistroUsuario 
         Height          =   315
         Left            =   -67600
         TabIndex        =   112
         Top             =   6840
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblRegistroFechaAbog 
         Height          =   315
         Left            =   -62320
         TabIndex        =   111
         Top             =   6480
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblRegistroFecha 
         Height          =   315
         Left            =   -67600
         TabIndex        =   110
         Top             =   6480
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblFirmasUsuario 
         Height          =   315
         Left            =   -62320
         TabIndex        =   109
         Top             =   5640
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblRecepcionUsuario 
         Height          =   315
         Left            =   -67600
         TabIndex        =   108
         Top             =   5640
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblFirmasFecha 
         Height          =   315
         Left            =   -62320
         TabIndex        =   107
         Top             =   5280
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblRecepcionFecha 
         Height          =   315
         Left            =   -67600
         TabIndex        =   106
         Top             =   5280
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblEntregaUsuarioAbog 
         Height          =   315
         Left            =   -62320
         TabIndex        =   105
         Top             =   4560
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblEntregaUsuario 
         Height          =   315
         Left            =   -67600
         TabIndex        =   104
         Top             =   4560
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblEntregaFechaAbog 
         Height          =   315
         Left            =   -62320
         TabIndex        =   103
         Top             =   4200
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblEntregaFecha 
         Height          =   315
         Left            =   -67600
         TabIndex        =   102
         Top             =   4200
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblAsignacionUsuarioAbog 
         Height          =   315
         Left            =   -62320
         TabIndex        =   101
         Top             =   3480
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblAsignacionUsuario 
         Height          =   315
         Left            =   -67600
         TabIndex        =   100
         Top             =   3480
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblAsignacionFechaAbog 
         Height          =   315
         Left            =   -62320
         TabIndex        =   99
         Top             =   3120
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblAsignacionFecha 
         Height          =   315
         Left            =   -67600
         TabIndex        =   98
         Top             =   3120
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblEstadoAbog 
         Height          =   315
         Left            =   -62320
         TabIndex        =   97
         Top             =   2280
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblEstadoIng 
         Height          =   315
         Left            =   -67600
         TabIndex        =   96
         Top             =   2280
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblNombreAbog 
         Height          =   315
         Left            =   -64600
         TabIndex        =   95
         Top             =   1920
         Visible         =   0   'False
         Width           =   4695
         _Version        =   1572864
         _ExtentX        =   8281
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblNombreIng 
         Height          =   315
         Left            =   -69880
         TabIndex        =   94
         Top             =   1920
         Visible         =   0   'False
         Width           =   4695
         _Version        =   1572864
         _ExtentX        =   8281
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   1
         Left            =   -64960
         TabIndex        =   93
         Top             =   1440
         Visible         =   0   'False
         Width           =   5055
         _Version        =   1572864
         _ExtentX        =   8916
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Abogado Asignado"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   0
         Left            =   -69880
         TabIndex        =   92
         Top             =   1440
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1572864
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Ingeniero Asignado"
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
      Begin XtremeSuiteControls.Label lblUsuarioRegistro 
         Height          =   315
         Left            =   -66520
         TabIndex        =   91
         Top             =   840
         Visible         =   0   'False
         Width           =   3255
         _Version        =   1572864
         _ExtentX        =   5741
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblEstadoActual 
         Height          =   315
         Left            =   -63160
         TabIndex        =   90
         Top             =   840
         Visible         =   0   'False
         Width           =   3255
         _Version        =   1572864
         _ExtentX        =   5741
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblFechaRegistro 
         Height          =   315
         Left            =   -69880
         TabIndex        =   89
         Top             =   840
         Visible         =   0   'False
         Width           =   3255
         _Version        =   1572864
         _ExtentX        =   5741
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   24
         Left            =   -66520
         TabIndex        =   88
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   23
         Left            =   -63040
         TabIndex        =   87
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estado actual"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   22
         Left            =   -69880
         TabIndex        =   86
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha de registro"
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
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   10470
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlbPrincipal"
      MinHeight1      =   330
      Width1          =   4260
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   2520
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   330
         Left            =   4455
         TabIndex        =   2
         Top             =   30
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   582
         ButtonWidth     =   2249
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Avaluo"
               Key             =   "avaluo"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Historial"
               Key             =   "historial"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Coberturas"
               Key             =   "Coberturas"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
               Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
               Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "borrar"
               Object.ToolTipText     =   "Borrar la información en la base de datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
               Object.ToolTipText     =   "Guarda la información en la base de datos"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
               Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "imprimir"
               Object.ToolTipText     =   "Imprimir ficha de garantía"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9480
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmVivGarantia.frx":0720
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmVivGarantia.frx":170E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmVivGarantia.frx":17FBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmVivGarantia.frx":2D12E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmVivGarantia.frx":33990
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmVivGarantia.frx":3A1F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmVivGarantia.frx":40A54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtExpediente 
      Height          =   315
      Left            =   6840
      TabIndex        =   5
      Top             =   480
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   4080
      TabIndex        =   6
      Top             =   840
      Width           =   4935
      _Version        =   1572864
      _ExtentX        =   8705
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   15
      Top             =   480
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Expediente"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   14
      Top             =   840
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Identificación"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   480
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Operacion"
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
Attribute VB_Name = "frmVivGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vEditar As Boolean, vPaso As Boolean

Public Item_Lista_Seleccionado As XtremeSuiteControls.ListViewItem
Private m_cargarCantones As Boolean
Private m_cargarGradoHipoteca As Boolean
Private m_cambioDatos As Boolean
Public m_NumeroOperacion As String

Private vItem As XtremeSuiteControls.ListViewItem
Private m_IdGarantia As Long
Private vCantonMascara As String, vDistritoMascara As String

Private m_EstadoSol As String

Public Sub sbListaDuenos()

Dim vKey As String

On Error GoTo vError

If m_IdGarantia = -1 Then Exit Sub

lvwListaDuenos.ColumnHeaders.Clear
lvwListaDuenos.ListItems.Clear

lvwListaDuenos.ColumnHeaders.Add , , "Cédula", 2000
lvwListaDuenos.ColumnHeaders.Add , , "Nombre", 2000
lvwListaDuenos.ColumnHeaders.Add , , "Provincia", 2000
lvwListaDuenos.ColumnHeaders.Add , , "Canton", 2000
lvwListaDuenos.ColumnHeaders.Add , , "Distrito", 2000
lvwListaDuenos.ColumnHeaders.Add , , "Usuario Registro", 2000
lvwListaDuenos.ColumnHeaders.Add , , "Fecha Registro", 2000

If ObjConsultar.fxTraerListaDuenosxGarantia(m_IdGarantia) Then
    With glogon.Recordset
        While Not .EOF
            vKey = "(VV)" & .Fields!IdGarantia _
                    & "(Ig)" & Trim(.Fields!cedula) & "(Cd)"
            
        Set vItem = lvwListaDuenos.ListItems.Add(, vKey, .Fields!cedula)
            vItem.SubItems(1) = Trim(.Fields!Nombre)
            vItem.SubItems(2) = Trim(.Fields!DescProvincia)
            vItem.SubItems(3) = Trim(.Fields!DescCanton)
            vItem.SubItems(4) = IIf(IsNull(.Fields!DescDistrito), "", Trim(.Fields!DescDistrito))
            vItem.SubItems(5) = Trim(.Fields!RegistroUsuario)
            vItem.SubItems(6) = Format(.Fields!RegistroFecha, "dd-mm-yyyy")
            .MoveNext
        Wend
    End With
End If

  Exit Sub
  
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Public Sub sbListaGarantias()

On Error GoTo vError
Dim vKey As String

lvwGarantias.ColumnHeaders.Clear
lvwGarantias.ListItems.Clear

lvwGarantias.ColumnHeaders.Add , , "Número Finca", 2000
lvwGarantias.ColumnHeaders.Add , , "Plano catastro", 2000
lvwGarantias.ColumnHeaders.Add , , "Tipo Derecho", 2000
lvwGarantias.ColumnHeaders.Add , , "Grado Hiporteca", 2000
lvwGarantias.ColumnHeaders.Add , , "Área Finca", 2000
lvwGarantias.ColumnHeaders.Add , , "Usuario Registro", 2000
lvwGarantias.ColumnHeaders.Add , , "Fecha Registro", 2000



If ObjConsultar.fxTraerGarantiasxOperacion(txtOperacion.Text, txtExpediente.Text) Then
While Not glogon.Recordset.EOF
    vKey = "(VV)" & Trim(txtOperacion.Text) _
    & "(Op)" & Trim(glogon.Recordset("NumeroFinca")) & "(Nf)"
           
    Set vItem = lvwGarantias.ListItems.Add(, vKey, glogon.Recordset("NumeroFinca"))
        vItem.SubItems(1) = Trim(glogon.Recordset.Fields!NumPlanoCatastro)
        vItem.SubItems(2) = Trim(glogon.Recordset.Fields!TipoDerecho)
        vItem.SubItems(3) = Trim(glogon.Recordset.Fields!DescGradoHipoteca)
        vItem.SubItems(4) = Trim(glogon.Recordset.Fields!AreaFinca)
        vItem.SubItems(5) = Trim(glogon.Recordset.Fields!RegistroUsuario)
        vItem.SubItems(6) = Format(glogon.Recordset.Fields!RegistroFecha, "dd-mm-yyyy")
        
        vItem.Tag = glogon.Recordset.Fields!IdGarantia
        
       glogon.Recordset.MoveNext
    Wend
End If

  Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Public Function sbCantonesxProvincia(cbo As XtremeSuiteControls.ComboBox, ByVal pProvincia As String) As Boolean

On Error GoTo vError
cbo.Clear
If ObjConsultar.fxTraerCantonesxProvincia(pProvincia) Then
    Do While Not glogon.Recordset.EOF
        cbo.AddItem Trim(glogon.Recordset!Descripcion)
        cbo.ItemData(cbo.ListCount - 1) = CStr(glogon.Recordset!Canton)
 
        glogon.Recordset.MoveNext
    Loop
    glogon.Recordset.MoveFirst
    cbo.Text = Trim(glogon.Recordset!Descripcion)
    
    m_cargarCantones = True
End If
  
Exit Function

vError:
  
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Function

Public Function sbCargarZonasxDistribPoli(cbo As ComboBox, ByVal pProvincia As Integer, ByVal pCanton As String) As Boolean

On Error GoTo vError
cbo.Clear
If ObjConsultar.fxZonaAsigna_TraerxDistPol(pProvincia, pCanton) Then
    Do While Not glogon.Recordset.EOF
        cbo.AddItem Trim(glogon.Recordset!Zona)
        cbo.ItemData(cbo.ListCount - 1) = CStr(glogon.Recordset!idZona)
        glogon.Recordset.MoveNext
    Loop
    glogon.Recordset.MoveFirst
    cbo.Text = Trim(glogon.Recordset!Zona)
'Else
'  MsgBox "No existen " & vTipo & " una zona asignada para el cantón seleccionado", vbCritical
End If
  
Exit Function

vError:
  
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Function

Public Function sbDistritosxCantones(ByRef cbo As XtremeSuiteControls.ComboBox, ByVal pProvincia As String, ByVal pCanton As String) As Boolean

On Error GoTo vError
cbo.Clear
If ObjConsultar.fxTraerDistritosxCanton(pProvincia, pCanton) Then
    Do While Not glogon.Recordset.EOF
        cbo.AddItem Trim(glogon.Recordset!NombreDistrito)
        cbo.ItemData(cbo.ListCount - 1) = CStr(glogon.Recordset!Distrito)
        
        glogon.Recordset.MoveNext
    Loop
    glogon.Recordset.MoveFirst
    cbo.Text = Trim(glogon.Recordset!NombreDistrito)
    
'Else
'  MsgBox "No existen " & vTipo & " Creadas...(Debe Crearlos)", vbCritical
End If
  
Exit Function

vError:
  
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Function


Private Sub sbCargaGradoHiporteca()

Dim i As Integer
cboGradoHipoteca.Clear

cboGradoHipoteca.AddItem "Primer Grado"
cboGradoHipoteca.AddItem "Segundo Grado"
cboGradoHipoteca.AddItem "Tercer Grado"

For i = 0 To 2
  cboGradoHipoteca.ListIndex = i
  cboGradoHipoteca.ItemData(cboGradoHipoteca.ListIndex) = i + 1
Next i
m_cargarGradoHipoteca = True
cboGradoHipoteca.Text = "Primer Grado"

End Sub

Private Sub sbCargaTipoPoliza()

Dim i As Integer
cboTipo_Poliza.Clear

cboTipo_Poliza.AddItem "Comercial"
cboTipo_Poliza.AddItem "Personal"

cboTipo_Poliza.Text = "Comercial"

End Sub

Private Sub sbHabilitaTab(ByVal pTab As Integer)
Select Case pTab
    Case 1 'Inicial para consulta
        
    
        tcMain.Item(0).Enabled = True
        tcMain.Item(1).Enabled = False
        tcMain.Item(2).Enabled = False
        
        tcMain.Item(0).Selected = True
        
    Case 2 'Hablitia cuando es nuevo
        tcMain.Item(0).Enabled = False
        tcMain.Item(1).Enabled = True
        tcMain.Item(2).Enabled = False
        
        tcMain.Item(1).Selected = True
    
    Case 3 'Hablitia todos
        tcMain.Item(0).Enabled = True
        tcMain.Item(1).Enabled = True
        tcMain.Item(2).Enabled = True
        
End Select


End Sub


'------Operaciones Garantia--------------------------------------------
Private Sub sbLigarDatosGarantia()

Dim vrstemp As ADODB.Recordset

On Error GoTo vError

If Not (glogon.Recordset.EOF) Then
    Set vrstemp = glogon.Recordset

    With vrstemp
        m_IdGarantia = !IdGarantia
        txtNumeroFinca.Text = !NumeroFinca
        txtTipoDerecho.Text = !TipoDerecho
        txtNumPlanoCatastro.Text = !NumPlanoCatastro
        chkCoberturaPrimerGrado.Enabled = True
        m_cargarGradoHipoteca = False
        
        Select Case !GradoHipoteca
            Case "P"
                cboGradoHipoteca.ListIndex = 0
                chkCoberturaPrimerGrado.Enabled = False
                chkCoberturaPrimerGrado.Value = 1
                btnGradoHipoteca.Visible = False
                
                 
                
            Case "S"
                cboGradoHipoteca.ListIndex = 1
                btnGradoHipoteca.Visible = True
            Case "T"
                cboGradoHipoteca.ListIndex = 2
                btnGradoHipoteca.Visible = True
        End Select
        
        If Not IsNull(!CoberturaPrimerGrado) Then
            chkCoberturaPrimerGrado.Value = !CoberturaPrimerGrado
        End If
        
        If Not IsNull(!RegistraCalAvaluo) Then
            chkRegistraCalAvaluo.Value = !RegistraCalAvaluo
        End If
        
        If Not IsNull(!RegistraCalHonorarios) Then
            chkRegistraCalHonorarios.Value = !RegistraCalHonorarios
        End If
        
        If Not IsNull(!RegistraCalHonorariosDT) Then
            chkDetalleHonorarios.Value = !RegistraCalHonorariosDT
        End If
        
        txtAreaFinca.Text = CStr(!AreaFinca)
        
        m_cargarGradoHipoteca = True
        m_cargarCantones = False

        If Not IsNull(!DescProvincia) Then
            cboUbicacionProvincia.Text = !DescProvincia
        End If
        
        If Not IsNull(!DescCanton) Then
            If cboUbicacionCanton.ListCount > 0 Then
                cboUbicacionCanton.Text = Trim(!DescCanton)
            End If
        End If
        
        If Not IsNull(!DescDistrito) Then
            If cboUbicacionDistrito.ListCount > 0 Then
                cboUbicacionDistrito.Text = Trim(!DescDistrito)
            End If
        
        End If
        
        If Not IsNull(!DescZona) Then
            If cboZonas.ListCount > 0 Then
                m_cargarCantones = False
                cboZonas.Text = Trim(!DescZona)
                m_cargarCantones = True
            End If
        End If
        
        txtDireccion.Text = IIf(IsNull(Trim(!Direccion)), "", Trim(!Direccion))
        gObservacion(0) = IIf(IsNull(Trim(!AnotacionesFinca)), "", Trim(!AnotacionesFinca))
        gObservacion(1) = IIf(IsNull(Trim(!Gravamenes)), "", Trim(!Gravamenes))
        gObservacion(2) = IIf(IsNull(Trim(!AnotacionesGravamen)), "", Trim(!AnotacionesGravamen))
        optObservacion(0).Value = True
        txtObservaciones.Text = gObservacion(0)
    
    txtMontoNoGravable.Text = IIf(IsNull(!MontoNoGravable), 0, Format(!MontoNoGravable, "Standard"))
    
     
    'Solo muestre el boton de consulta del avaluo si la garantia si ya se le registro el avaluo
    ' y no este asignada a un abogado
    
    If Not IsNull(!ValorTerreno) Then
        tlbAux.Buttons.Item(1).Visible = True
        
        chkRegistraCalAvaluo.Enabled = False
    Else
        tlbAux.Buttons.Item(1).Visible = False
        chkRegistraCalAvaluo.Enabled = True
    End If
    If !Estado = "R" Then
        chkRegistraCalHonorarios.Enabled = False
        chkDetalleHonorarios.Enabled = False
    Else
        chkRegistraCalHonorarios.Enabled = True
        chkDetalleHonorarios.Enabled = False
    End If
    
    If !Tipo_Poliza = "P" Then
        cboTipo_Poliza.Text = "Personal"
    Else
        cboTipo_Poliza.Text = "Comercial"
    End If
        
    
    .Close
    End With

    m_cambioDatos = False
End If

Exit Sub
vError:
    ObjMensajes.deError ("Ocurrió un error en Visual Basic al mostrar la información solicitada. Error:" & Err.Description)

End Sub

Private Sub sbLigarDatosDuenos()
On Error GoTo vError
Dim vrstemp As ADODB.Recordset

If Not (glogon.Recordset.EOF) Then
Set vrstemp = glogon.Recordset
    With vrstemp
    m_IdGarantia = .Fields!IdGarantia
    txtCedulaDueno.Text = .Fields!cedula
    txtNombreDueno.Text = .Fields!Nombre
    txtDireccionDueno.Text = IIf(IsNull(.Fields!Direccion), "", Trim(.Fields!Direccion))
    
    m_cargarCantones = False
    cboProvinciasDuenos.Text = Trim(.Fields!DescProvincia)
    
    Call sbCantonesxProvincia(cboCantonesDueno, cboProvinciasDuenos.ItemData(cboProvinciasDuenos.ListIndex))
    If cboCantonesDueno.ListCount > 0 Then
        cboCantonesDueno.Text = Trim(.Fields!DescCanton)
    End If
     m_cargarCantones = True
    If Not IsNull(.Fields!DescDistrito) Then
        If cboDistritosDuenos.ListCount > 0 Then
            cboDistritosDuenos.Text = Trim(.Fields!DescDistrito)
        End If
    End If
    .Close
    End With
    m_cambioDatos = False
End If
salir:
Exit Sub
vError:
ObjMensajes.deError ("Ocurrió un error en Visual Basic al mostrar la información solicitada. Error:" & Err.Description)
End Sub


Private Function fxValidaDatosGarantia(ByVal pEditar As Boolean) As Boolean

Dim vGradoHipoteca As String
Dim vUbicacionDistrito As String
Dim vZona As Integer

On Error GoTo error

fxValidaDatosGarantia = False

ReDim gParametros(0 To 24)

If ObjConsultar.fxEstadoOperacion(txtOperacion.Text) = "F" Then
    Me.MousePointer = vbDefault
    MsgBox ("No es posible realizar movimientos para un número de operación en estado FORMALIZADA.")
    Exit Function
End If

If (Len(Trim(txtNumeroFinca.Text)) = 0) Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe de ingresar un número de finca válido.")
    txtNumeroFinca.SetFocus
    Exit Function
End If
If (Len(Trim(txtTipoDerecho.Text)) = 0) Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe de ingresar un tipo de derecho.")
    txtTipoDerecho.SetFocus
    Exit Function
End If
If (Len(Trim(txtNumPlanoCatastro.Text)) = 0) Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe de ingresar un número de plano catastro.")
    txtNumPlanoCatastro.SetFocus
    Exit Function
End If

If (Len(Trim(txtAreaFinca.Text)) = 0) Or Not IsNumeric(txtAreaFinca.Text) Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe de ingresar el área en metro cuadrados. (Solo números)")
    txtAreaFinca.SetFocus
    Exit Function
End If

If cboUbicacionProvincia.ListCount = 0 Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe seleccionar una provincia.")
    cboUbicacionProvincia.SetFocus
    Exit Function
End If
If cboUbicacionCanton.ListCount = 0 Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe seleccionar un cantón.")
    cboUbicacionCanton.SetFocus
    Exit Function
End If

If cboZonas.ListCount = 0 Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe seleccionar una zona.")
    cboZonas.SetFocus
    Exit Function
End If

vGradoHipoteca = ""
Select Case cboGradoHipoteca.ItemData(cboGradoHipoteca.ListIndex)
    Case 1 '"Primer Grado"
        vGradoHipoteca = "P"
    Case 2 '"Segundo Grado"
        vGradoHipoteca = "S"
    Case 3 '"Tercer Grado"
        vGradoHipoteca = "T"
End Select

'' Verifica al modificar si tiene detalle de garantía
If pEditar = True Then

    If ObjConsultar.fxValidaDetalleGarantia(vGradoHipoteca, m_IdGarantia) = False Then
        Me.MousePointer = vbDefault
        MsgBox ("Antes de modificar el grado de la garantía, debe revisar el grado de las garantías en el detalle de acredores")
        Exit Function
    End If

End If



vZona = -1
vUbicacionDistrito = "-1"
If cboUbicacionDistrito.ListCount > 0 Then
    vUbicacionDistrito = cboUbicacionDistrito.ItemData(cboUbicacionDistrito.ListIndex)
End If

If cboZonas.ListCount > 0 Then
    vZona = cboZonas.ItemData(cboZonas.ListIndex)
End If

gParametros(1) = IIf((vZona = -1), ObjNull.NullInt, vZona)
gParametros(2) = cboUbicacionProvincia.ItemData(cboUbicacionProvincia.ListIndex)
gParametros(3) = cboUbicacionCanton.ItemData(cboUbicacionCanton.ListIndex)
gParametros(4) = cboUbicacionDistrito.ItemData(cboUbicacionDistrito.ListIndex)
gParametros(5) = Trim(txtNumeroFinca.Text)
gParametros(6) = Trim(txtTipoDerecho.Text)
gParametros(7) = Trim(txtNumPlanoCatastro.Text)
gParametros(8) = vGradoHipoteca
gParametros(9) = Trim(txtAreaFinca.Text)
gParametros(10) = "S" 'T= Tramitado,R= Recida, S = Solicitada
gParametros(11) = IIf((Len(Trim(txtDireccion.Text)) = 0), ObjNull.SetNull, Trim(txtDireccion.Text))
gParametros(12) = IIf((Len(Trim(gObservacion(0))) = 0), ObjNull.SetNull, Trim(gObservacion(0)))
gParametros(13) = IIf((Len(Trim(gObservacion(1))) = 0), ObjNull.SetNull, Trim(gObservacion(1)))
gParametros(14) = IIf((Len(Trim(gObservacion(2))) = 0), ObjNull.SetNull, Trim(gObservacion(2)))
gParametros(15) = ObjNull.SetNull  'pObservacionAvaluo
gParametros(16) = glogon.Usuario
gParametros(17) = "1900/01/01"
gParametros(18) = txtOperacion.Text
gParametros(19) = chkCoberturaPrimerGrado.Value
gParametros(20) = chkRegistraCalAvaluo.Value 'RegistraCalAvaluo
gParametros(21) = chkRegistraCalHonorarios.Value 'RegistraCalHonorarios
gParametros(22) = chkDetalleHonorarios.Value ' RegistraCalHonorariosDT
gParametros(23) = Mid(cboTipo_Poliza.Text, 1, 1)
gParametros(24) = txtExpediente.Text

fxValidaDatosGarantia = True
salir:
    Exit Function
error:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbModificarGarantia()
On Error GoTo error

Me.MousePointer = vbHourglass

If m_IdGarantia <> -1 Then

 If Not fxValidaDatosGarantia(True) Then Exit Sub
 
 m_IdGarantia = ObjAgregar.fxViviendaGarantia(m_IdGarantia, gParametros(1), gParametros(2), gParametros(3), gParametros(4), _
                                                    gParametros(5), gParametros(6), gParametros(7), gParametros(8), gParametros(9), _
                                                    gParametros(10), gParametros(11), gParametros(12), gParametros(13), gParametros(14), _
                                                    gParametros(15), gParametros(16), gParametros(17), gParametros(18), gParametros(19), _
                                                    gParametros(20), gParametros(21), gParametros(22), gParametros(23), gParametros(24))
                                      
 If m_IdGarantia <> -1 Then
    m_cambioDatos = False
    
    Call Bitacora("MODIFICA", "Garantía vivienda: " & m_IdGarantia)
    
    MsgBox "Información fue actualizada correctamente.", vbInformation
    Call sbClearControles(Me)
    Call sbHabilitaTab(1)
    Call sbListaGarantias
    Call sbToolBar(Me.tlbPrincipal, "nuevo")
 End If

End If

salir:
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    Call ObjMensajes.deError("Ocurrió un error en visual basic al modificar la información ingresada. Error " & Err.Description)
End Sub

Private Sub sbAgregarGarantia()
Dim vIdGarantia As Integer

On Error GoTo vError
Me.MousePointer = vbHourglass

' Valida los casos con avaluo posterior a formalización

If tcAux.Item(1).Enabled = True Then
    If fxValidaDatosAvaluo = False Then Exit Sub
End If

If fxValidaDatosGarantia(False) = False Then Exit Sub

gParametros(0) = m_IdGarantia

m_IdGarantia = ObjAgregar.fxViviendaGarantia(gParametros(0), gParametros(1), gParametros(2), gParametros(3), gParametros(4), _
                                             gParametros(5), gParametros(6), gParametros(7), gParametros(8), gParametros(9), _
                                             gParametros(10), gParametros(11), gParametros(12), gParametros(13), gParametros(14), gParametros(15), _
                                             gParametros(16), gParametros(17), gParametros(18), gParametros(19), _
                                             gParametros(20), gParametros(21), gParametros(22), gParametros(23), gParametros(24))
                                     
gParametros(0) = m_IdGarantia

If m_IdGarantia <> -1 Then
    m_cambioDatos = False
    
    If Mid(cboGradoHipoteca.Text, 1, 1) <> "P" Then
        Call sbMostrarPantallaDetalle
    End If
    
    Call Bitacora("REGISTRA", "Garantía Hipotecaria: " & m_IdGarantia)
    
    ' Inserta avaluo posterior si aplica
    If tcAux.Item(1).Enabled Then
        Call sbAgregarAvaluo(m_IdGarantia)
        tcAux.Item(1).Enabled = False
    End If
    
    MsgBox "Información fue registrada corretamente.", vbInformation
    
    Call sbClearControles(Me)
    Call sbHabilitaTab(1)
    Call sbListaGarantias
    Call sbToolBar(Me.tlbPrincipal, "nuevo")
    

    
    
End If

salir:
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Resume
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbGuardar()

Select Case True

    Case tcMain.Item(1).Selected
        If m_cambioDatos = False Then Exit Sub
        
        If Not vEditar Then 'Nuevo
            Call sbAgregarGarantia
           
        Else 'Modificar
            Call sbModificarGarantia
        End If
    Case tcMain.Item(2).Selected
        If m_cambioDatos = False Then Exit Sub
        If Not vEditar Then 'Nuevo
            Call sbAgregarDueno
        Else 'Modificar
            Call sbModificarDuenos
        End If
        
End Select
  
End Sub
'---------------------Guardar informacion de duenos segun numero de garantia--------------------
Private Function fxValidaDatosDueno() As Boolean

Dim vMensaje As String
Dim vDistrito As String

On Error GoTo vError


fxValidaDatosDueno = False

ReDim gParametros(0 To 8)

vMensaje = ""

If ObjConsultar.fxEstadoOperacion(txtOperacion.Text) = "F" Then
    Me.MousePointer = vbDefault
    MsgBox ("No es posible realizar movimientos para un número de operación en estado FORMALIZADA.")
    Exit Function
End If

If (Len(Trim(txtCedulaDueno.Text)) = 0) Then vMensaje = vMensaje & " - Debe de ingresar un número de cédula válido." & vbCrLf
If (Len(Trim(txtNombreDueno.Text)) = 0) Then vMensaje = vMensaje & " - Debe de ingresar un nombre de dueño." & vbCrLf
If cboProvinciasDuenos.ListCount = 0 Then vMensaje = vMensaje & " - Debe seleccionar una provincia." & vbCrLf
If cboCantonesDueno.ListCount = 0 Then vMensaje = vMensaje & " - Debe seleccionar un cantón." & vbCrLf

vDistrito = "-1"
If cboDistritosDuenos.ListCount > 0 Then
    vDistrito = cboDistritosDuenos.ItemData(cboDistritosDuenos.ListIndex)
End If

If Len(vMensaje) = 0 Then
  fxValidaDatosDueno = True
Else
  fxValidaDatosDueno = False
  Call ObjMensajes.deDatos("-1", vMensaje)
  Me.MousePointer = vbDefault
  Exit Function
End If

gParametros(0) = Trim(txtCedulaDueno.Text)
gParametros(1) = m_IdGarantia
gParametros(2) = cboProvinciasDuenos.ItemData(cboProvinciasDuenos.ListIndex)
gParametros(3) = cboCantonesDueno.ItemData(cboCantonesDueno.ListIndex)
gParametros(4) = cboDistritosDuenos.ItemData(cboDistritosDuenos.ListIndex)
gParametros(5) = Trim(txtNombreDueno.Text)
gParametros(6) = IIf((Len(txtDireccionDueno.Text) = 0), ObjNull.NullString, Trim(txtDireccionDueno.Text))
gParametros(7) = glogon.Usuario
gParametros(8) = "1900/01/01"

fxValidaDatosDueno = True

salir:
    Exit Function
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbModificarDuenos()
On Error GoTo error

Me.MousePointer = vbHourglass
If m_IdGarantia <> -1 Then
 If fxValidaDatosDueno = False Then Exit Sub
 
 If ObjAgregar.fxDerechosGarantia(1, gParametros(0), gParametros(1), gParametros(2), gParametros(3), _
                                                    gParametros(4), gParametros(5), gParametros(6), gParametros(7), gParametros(8)) Then

    m_cambioDatos = False
    
    Call Bitacora("MODIFICA", "Dueño garantía vivienda Gar: " & gParametros(1) & " Dueño:" & gParametros(0))
    
    MsgBox "Información fue actualizada correctamente.", vbInformation
    Call sbLimpiaDatosDuenos
    Call sbListaDuenos
    Call sbToolBar(Me.tlbPrincipal, "nuevo")
    
 End If

End If
salir:
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    Call ObjMensajes.deError("Ocurrió un error en visual basic al modificar la información ingresada. Error " & Err.Description)
End Sub

Private Sub sbAgregarDueno()
On Error GoTo vError

Me.MousePointer = vbHourglass
If fxValidaDatosDueno() = False Then Exit Sub
If ObjAgregar.fxDerechosGarantia(-1, gParametros(0), gParametros(1), gParametros(2), gParametros(3), gParametros(4), _
                                                   gParametros(5), gParametros(6), gParametros(7), gParametros(8)) Then
    m_cambioDatos = False
    
    Call Bitacora("REGISTRA", "Dueño garantía vivienda Gar: " & gParametros(1) & " Dueño:" & gParametros(0))
    
    MsgBox "Información fue registrada corretamente.", vbInformation
    Call sbLimpiaDatosDuenos
    Call sbListaDuenos
    Call sbToolBar(Me.tlbPrincipal, "nuevo")
End If

salir:
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbLimpiaDatosDuenos(Optional ByVal ptodos As Boolean = True)

If ptodos Then
    txtCedulaDueno.Text = Empty
    txtCedulaDueno.SetFocus
End If
    cboProvinciasDuenos.ListIndex = 0
    txtNombreDueno.Text = Empty
    txtDireccionDueno.Text = Empty
    
End Sub


Private Sub btnGradoHipoteca_Click()
    If m_cambioDatos = True Then
        MsgBox "Debe de guardar primero los cambios para ingresar a esta opción"
        Exit Sub
    End If
    
    Call sbMostrarPantallaDetalle
End Sub

Private Sub btnHonorariosDetalle_Click()


GLOBALES.gTag = m_IdGarantia

Call sbFormsCall("frmVivConsultaHonorariosDetalle", 1, , , False, Me)


End Sub

Private Sub cboCantonesDueno_Click()
Dim strSQL As String

If vPaso Then Exit Sub

    strSQL = "select Distrito as Idx, rtrim(Descripcion) as ItmX from Distritos" _
            & " where provincia = '" & cboProvinciasDuenos.ItemData(cboProvinciasDuenos.ListIndex) _
            & "' and Canton = '" & cboCantonesDueno.ItemData(cboCantonesDueno.ListIndex) _
            & "' order by descripcion"
    Call sbCbo_Llena_New(cboDistritosDuenos, strSQL, False, True)

'Agrega Distrito En Limpio, ya que este dato es opcional
cboDistritosDuenos.AddItem " "
cboDistritosDuenos.Text = " "

End Sub

Private Sub cboDistritosDuenos_Click()
m_cambioDatos = True
End Sub


Private Sub cboTipo_Poliza_Click()
    m_cambioDatos = True
End Sub

Private Sub chkDetalleHonorarios_Click()
m_cambioDatos = True
If chkDetalleHonorarios.Value Then
    chkRegistraCalHonorarios.Value = 1
    'chkDetalleHonorarios.Value = 1
End If
End Sub

Private Sub chkRegistraCalAvaluo_Click()
m_cambioDatos = True
End Sub

Private Sub chkRegistraCalHonorarios_Click()
m_cambioDatos = True
'If chkRegistraCalHonorarios.Value Then
'    chkDetalleHonorarios.Enabled = True
'Else
'    chkDetalleHonorarios.Enabled = False
'    chkDetalleHonorarios.Value = 0
'End If
End Sub



Private Sub Form_Activate()
vModulo = 3 'Modulo de Credito
End Sub

Private Sub Form_Unload(Cancel As Integer)
gOperacion = 0
End Sub





Private Sub lvwGarantias_DblClick()
Dim vTemp As String
If lvwGarantias.ListItems.Count > 0 Then
' Call sbCargarDistribPolitica(1)
    
    Call sbHabilitaTab(3)
    
    tcMain.Item(1).Selected = True
    tcAux.Item(0).Selected = True
    
    
    If Item_Lista_Seleccionado Is Nothing Then Exit Sub
    
    
    txtNumeroFinca.Text = Item_Lista_Seleccionado.Tag
    Call sbTraerGarantia(0)
    
    
'    txtOperacion.Text = fxDeCodePK(Item_Lista_Seleccionado.Key, 5, "(Op)")
'    txtNumeroFinca.Text = fxDeCodePK(Item_Lista_Seleccionado.Key, gPosIni, "(Nf)")
'
'    If Len(txtNumeroFinca.Text) > 0 And Len(txtOperacion.Text) > 0 Then
'        Call sbTraerGarantia(1)
'    End If

End If

End Sub

Private Sub lvwGarantias_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Set Item_Lista_Seleccionado = Item
End Sub

Private Sub lvwListaDuenos_DblClick()
Dim vTemp As String

If lvwListaDuenos.ListItems.Count > 0 Then
    vTemp = fxDeCodePK(Item_Lista_Seleccionado.Key, 5, "(Ig)")
    txtCedulaDueno.Text = fxDeCodePK(Item_Lista_Seleccionado.Key, gPosIni, "(Cd)")
    Call txtCedulaDueno_LostFocus
    txtNombreDueno.SetFocus
End If

End Sub

Private Sub lvwListaDuenos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Set Item_Lista_Seleccionado = Item
End Sub

Private Sub optObservacion_Click(Index As Integer)

Select Case Index
 Case 0
    txtObservaciones.Text = gObservacion(0)
    txtObservaciones.SetFocus
    
 Case 1
    txtObservaciones.Text = gObservacion(1)
    txtObservaciones.SetFocus
    
 Case 2
    txtObservaciones.Text = gObservacion(2)
    txtObservaciones.SetFocus
    
End Select

End Sub

Private Sub SbInicializaVentana()

If Not vEditar Then
    
    Select Case True
        Case tcMain.Item(0).Selected, tcMain.Item(1).Selected
            Call sbHabilitaTab(2)
            Call sbClearControles(Me)
            Call sbCargaGradoHiporteca
            
           ' Call sbCargarDistribPolitica(1)
            
            btnGradoHipoteca.Visible = False
            chkCoberturaPrimerGrado.Enabled = False
            chkCoberturaPrimerGrado.Value = 1
            chkRegistraCalAvaluo.Enabled = True
            chkRegistraCalAvaluo.Value = 1
            chkRegistraCalHonorarios.Value = 1
'            chkDetalleHonorarios.Value = 0
        Case tcMain.Item(2).Selected
            Call sbLimpiaDatosDuenos
           ' Call sbCargarDistribPolitica(2)
            Call sbListaDuenos
            txtCedulaDueno.SetFocus
            
    End Select
    m_cambioDatos = False
End If

End Sub

Private Sub sbClearControles(vforma As Object, Optional ByVal ptodos As Boolean = True)
Dim vControl As Control
If ptodos Then
    For Each vControl In vforma
      If TypeOf vControl Is TextBox Then
         vControl.Text = ""
      End If
    Next
Else
    For Each vControl In vforma
      If TypeOf vControl Is TextBox Then
        If vControl.Name <> "txtNumeroFinca" Then
            vControl.Text = ""
        End If
        cboGradoHipoteca.Text = "Primer Grado"
        
      End If
    Next
End If
m_IdGarantia = -1
gObservacion(0) = ""
gObservacion(1) = ""
gObservacion(2) = ""

tcAux.Item(0).Selected = True

End Sub

Private Sub sbTraerGarantia(ByVal vTab As Integer)

Select Case vTab
    Case 0
        Call sbToolBar(Me.tlbPrincipal, "edicion")
        If ObjConsultar.fxTraerGarantiasxId(txtNumeroFinca.Text) Then
            Call sbLigarDatosGarantia
            Call sbToolBar(tlbPrincipal, "Activo")
        End If
    Case 1
        Call sbToolBar(Me.tlbPrincipal, "edicion")
        If ObjConsultar.fxTraerGarantiasxNumeroFinca(txtNumeroFinca.Text, txtOperacion.Text) Then
            Call sbLigarDatosGarantia
            Call sbToolBar(tlbPrincipal, "Activo")
        Else
            '' Para cargar datos de garantia si ha sido digitada en otro crédito
            If ObjConsultar.fxTraerGarantiasSoloxNumeroFinca(txtNumeroFinca.Text) Then
               Call sbLigarDatosGarantia
               Call sbToolBar(tlbPrincipal, "Activo")
            Else
                Call sbClearControles(Me, False)
            End If
        End If
        
    Case 2
        
        Call sbToolBar(Me.tlbPrincipal, "edicion")
        If ObjConsultar.fxTraerDuenoGarantia(m_IdGarantia, txtCedulaDueno.Text) Then
            Call sbLigarDatosDuenos
            Call sbToolBar(tlbPrincipal, "Activo")
        Else
            Call sbLimpiaDatosDuenos(False)
        End If
End Select
    
End Sub





Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
m_cambioDatos = False

Select Case Item.Index
    Case 0
        Call sbListaGarantias
    Case 1
        txtNumeroFinca.SetFocus
        m_cambioDatos = False
    Case 2
        
        Call sbListaDuenos
        txtCedulaDueno.SetFocus
        m_cambioDatos = False
    Case "3"
    Call sbLimipiarInfoTramite
    Call sbTraerInfoTramite
        
End Select

End Sub

Private Sub txtIdAbogado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "IdContacto"
        gBusquedas.Orden = "Nombre"
        gBusquedas.Filtro = " and TipoProfesional = 'A' "
        gBusquedas.Consulta = "select IdContacto,Nombre " _
                            & " from ViviendaContactos "
        frmBusquedas.Show vbModal
        txtIdAbogado = gBusquedas.Resultado
        txtNombreAbogado = gBusquedas.Resultado2
    
    End If
End Sub

Private Sub txtIdIngeniero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "IdContacto"
        gBusquedas.Orden = "Nombre"
        gBusquedas.Filtro = " and TipoProfesional = 'I' "
        gBusquedas.Consulta = "select IdContacto,Nombre " _
                            & " from ViviendaContactos "
        frmBusquedas.Show vbModal
        txtIdIngeniero = gBusquedas.Resultado
        txtNombreIngeniero = gBusquedas.Resultado2
    
    End If
End Sub

Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo vError

Select Case Button.Key

    Case "avaluo"
        
        frmVivRegistroAvaluo.vNumOperacion = txtOperacion.Text
        frmVivRegistroAvaluo.vIdGarantia = m_IdGarantia
               
        GLOBALES.gTag = txtOperacion.Text
        GLOBALES.gTag2 = m_IdGarantia
        GLOBALES.gTag3 = 0
 
 
        Call sbFormsCall("frmVivRegistroAvaluo", vbModal, , , False, Me)
        
    
    Case "historial"
        MsgBox "Pendiente", vbExclamation
        
    Case "Coberturas"
        gOperacion = txtOperacion.Text
        
        Call sbFormsCall("frmVivCoberturas", vbModal, , , False, Me)
End Select

Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
    
End Sub


Private Sub sbMostrarPantallaDetalle()
    GLOBALES.gTag = cboGradoHipoteca.Text
    GLOBALES.gTag2 = m_IdGarantia
    GLOBALES.gTag3 = "-1"
    Call sbSIFForms("frmVivDetalleGarantia", 1, , , False)

End Sub


Private Sub txtCedulaDueno_Change()
m_cambioDatos = True
End Sub

Private Sub txtCedulaDueno_LostFocus()
If Len(txtCedulaDueno.Text) > 0 And m_IdGarantia <> -1 Then

Call sbTraerGarantia(2)

glogon.strSQL = "select nombre from socios where cedula = '" & txtCedulaDueno.Text & "'"

If execSql(glogon.strSQL, True) Then
    txtNombreDueno.Text = glogon.Recordset.Fields!Nombre
End If



End If

End Sub

Private Sub txtDireccionDueno_Change()
m_cambioDatos = True
End Sub

Private Sub txtNombreAbogado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "Nombre"
        gBusquedas.Orden = "Nombre"
        gBusquedas.Filtro = " and TipoProfesional = 'A' "
        gBusquedas.Consulta = "select IdContacto,Nombre " _
                            & " from ViviendaContactos "
        frmBusquedas.Show vbModal
        txtIdAbogado = gBusquedas.Resultado
        txtNombreAbogado = gBusquedas.Resultado2
    
    End If
End Sub

Private Sub txtNombreDueno_Change()
m_cambioDatos = True
End Sub

Private Sub txtNombreIngeniero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "Nombre"
        gBusquedas.Orden = "Nombre"
        gBusquedas.Filtro = " and TipoProfesional = 'I' "
        gBusquedas.Consulta = "select IdContacto,Nombre " _
                            & " from ViviendaContactos "
        frmBusquedas.Show vbModal
        txtIdIngeniero = gBusquedas.Resultado
        txtNombreIngeniero = gBusquedas.Resultado2
    
    End If
End Sub

Private Sub TxtObservaciones_Change()
m_cambioDatos = True
If optObservacion.Item(0) = True Then 'AnotacionesFinca
   gObservacion(0) = txtObservaciones.Text
ElseIf optObservacion.Item(1) = True Then 'Gravamenes
   gObservacion(1) = txtObservaciones.Text
ElseIf optObservacion.Item(2) = True Then 'AnotacionesGravamen
   gObservacion(2) = txtObservaciones.Text
End If
End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo vError
    
Select Case Button.Key
        
    Case "nuevo"
        vEditar = False
        Call sbToolBar(Me.tlbPrincipal, "edicion")
        Call SbInicializaVentana
        Call sbControlaAvaluoPosterior
        
    Case "editar"
        vEditar = True
        Call sbToolBar(Me.tlbPrincipal, "edicion")
        Call sbHabilitaTab(3)
        
    Case "borrar"
        Call sbBorrar
        
    Case "guardar"
         Call sbGuardar
        
    Case "deshacer"
        vEditar = False
        Call sbToolBar(Me.tlbPrincipal, "nuevo")
        Call sbHabilitaTab(1)
    Case "imprimir"
        Call sbImprimirRptGarantia
        
    Case ""
End Select
    
Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
    
End Sub

Private Sub sbImprimirRptGarantia()
On Error GoTo error

    If m_IdGarantia = 0 Then
        Exit Sub
    End If

    With frmContenedor.Crt
        .Reset
        .Connect = Empty
        
        .WindowShowGroupTree = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowTitle = "Garantía de Créditos Hipotecarios"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Connect = glogon.ConectRPT
        
        .ReportFileName = SIFGlobal.fxPathReportes("Credito_Hipotecario_Garantia.rpt")
        
        .Formulas(0) = "Empresa= '" & GLOBALES.gstrNombreEmpresa & "'"
        .StoredProcParam(0) = m_IdGarantia

        .PrintReport
    End With
    frmContenedor.Crt.StoredProcParam(0) = Empty
    Exit Sub
error:
MsgBox ("Ocurrió un error al imprimir reporte de ficha de garantía. Error Nativo: " & Err.Description)
End Sub


Private Sub sbBorrar()

On Error GoTo vError

Select Case True
    Case tcMain.Item(2).Selected  '2
        If vEditar = False And Len(txtCedulaDueno.Text) > 0 Then
            If ObjConsultar.fxEstadoOperacion(txtOperacion.Text) = "F" Then
                Me.MousePointer = vbDefault
                MsgBox ("No es posible realizar movimientos para un número de operación en estado FORMALIZADA.")
                Exit Sub
            End If
            If ObjBorrar.fxDerechoDeGarantia(m_IdGarantia, txtCedulaDueno.Text) Then
                    
                Call Bitacora("BORRA", "Garantía vivienda: " & m_IdGarantia)
                
                Call ObjMensajes.deDatos("06")
                Call sbLimpiaDatosDuenos
                Call sbListaDuenos
                Call sbToolBar(tlbPrincipal, "nuevo")
            End If
        End If
    
End Select
 
 Exit Sub
 
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub cboUbicacionProvincia_Click()
Dim strSQL As String

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboUbicacionProvincia.ItemData(cboUbicacionProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboUbicacionCanton, strSQL, False, True)
vPaso = False

Call cboUbicacionCanton_Click

End Sub

Private Sub cboZonas_Click()
m_cambioDatos = True
End Sub

Private Sub cboUbicacionCanton_Click()
Dim strSQL As String

If vPaso Then Exit Sub

    strSQL = "select Distrito as Idx, rtrim(Descripcion) as ItmX from Distritos" _
            & " where provincia = '" & cboUbicacionProvincia.ItemData(cboUbicacionProvincia.ListIndex) _
            & "' and Canton = '" & cboUbicacionCanton.ItemData(cboUbicacionCanton.ListIndex) _
            & "' order by descripcion"
    Call sbCbo_Llena_New(cboUbicacionDistrito, strSQL, False, True)

'Agrega Distrito En Limpio, ya que este dato es opcional
cboUbicacionDistrito.AddItem " "
cboUbicacionDistrito.Text = " "
End Sub

Private Sub cboUbicacionDistrito_Click()
m_cambioDatos = True


End Sub

Private Sub cboGradoHipoteca_Click()
m_cambioDatos = True
If m_cargarGradoHipoteca = False Then Exit Sub
btnGradoHipoteca.Visible = False
If cboGradoHipoteca.ItemData(cboGradoHipoteca.ListIndex) = 1 Then
    chkCoberturaPrimerGrado.Enabled = False
    chkCoberturaPrimerGrado.Value = 1
Else
    If m_IdGarantia <> -1 Then 'Nuevo
        btnGradoHipoteca.Visible = True
    End If

    chkCoberturaPrimerGrado.Enabled = True
    chkCoberturaPrimerGrado.Value = 0
End If

End Sub

Private Sub txtAreaFinca_Change()
m_cambioDatos = True

End Sub

Private Sub txtDireccion_Change()
m_cambioDatos = True
End Sub
Private Sub txtNumeroFinca_LostFocus()
'No consulte la finca así
'If Len(txtNumeroFinca.Text) > 0 And Len(txtOperacion.Text) > 0 Then
'    Call sbTraerGarantia(1)
'End If
End Sub
Private Sub txtNumeroFinca_Change()
m_cambioDatos = True
End Sub
 
Private Sub txtNumPlanoCatastro_Change()
m_cambioDatos = True
End Sub
Private Sub txtTipoDerecho_Change()
m_cambioDatos = True
End Sub
Private Sub cboProvinciasDuenos_Click()
Dim strSQL As String

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvinciasDuenos.ItemData(cboProvinciasDuenos.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboUbicacionCanton, strSQL, False, True)
vPaso = False

Call cboCantonesDueno_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.ActiveControl.Name = "txtDireccion" Then Exit Sub
If Me.ActiveControl.Name = "txtObservaciones" Then Exit Sub
If Me.ActiveControl.Name = "txtDireccionDueno" Then Exit Sub



If (KeyCode = vbKeyReturn) Then
    Call gsbPulsarTecla(vbKeyTab)
    ElseIf KeyCode = vbKeyF4 Then
        Call sbBusqueda(Me.ActiveControl.Name)
End If

End Sub

Private Sub sbBusqueda(ByVal Control As String)

gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""
gBusquedas.Convertir = "N"

Select Case Control
   
    Case "txtCedulaDueno"
        gBusquedas.Consulta = "select Cedula,Nombre from socios"
        gBusquedas.Orden = "cedula"
        gBusquedas.Columna = "cedula"
        frmBusquedas.Show vbModal
        txtCedulaDueno.Text = gBusquedas.Resultado
        If Len(Trim(txtCedulaDueno.Text)) > 0 Then
          Call txtCedulaDueno_LostFocus
        End If
    Case "txtNombreDueno"
        gBusquedas.Consulta = "select Cedula,Nombre from socios"
        gBusquedas.Orden = "nombre"
        gBusquedas.Columna = "nombre"
        frmBusquedas.Show vbModal
        txtCedulaDueno.Text = gBusquedas.Resultado
        If Len(Trim(txtCedulaDueno.Text)) > 0 Then
          Call txtCedulaDueno_LostFocus
        End If
End Select

End Sub
Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 3 'Modulo de Credito


'Inicializa Barra

vEditar = False
Call sbToolBarIconos(tlbPrincipal, False)
Call sbToolBar(tlbPrincipal, "nuevo")
'Inicializa Seguridad
Call Formularios(Me)
Call RefrescaTags(Me)

txtOperacion.Text = gOperacion
txtExpediente.Text = gExpediente

m_cargarGradoHipoteca = False
m_cargarCantones = False



Call sbHabilitaTab(1)

'Call sbCargarDistribPolitica(1)

vPaso = True
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboUbicacionProvincia, strSQL, False, True)
    Call sbCbo_Llena_New(cboProvinciasDuenos, strSQL, False, True)
    
    strSQL = "select idZona as 'IdX', rtrim(Descripcion) as 'ItmX' from ViviendaZonas"
    Call sbCbo_Llena_New(cboZonas, strSQL, False, True)
    
vPaso = False



txtOperacion.Text = CStr(gOperacion)
txtExpediente.Text = gExpediente

Call sbCargaGradoHiporteca
Call sbTraerInformacionOperacion
Call sbListaGarantias
Call sbCargaTipoPoliza

'Call sbCargarDistribPolitica(2)
'Call cboUbicacionProvincia_Click
'Call cboProvinciasDuenos_Click

tcAux.Item(1).Enabled = False

End Sub

Private Sub sbLimipiarInfoTramite()

On Error GoTo vError
    
    lblFechaRegistro.Caption = Empty
    lblUsuarioRegistro.Caption = Empty
    lblEstadoActual.Caption = Empty
    
    'Infomacion del proceso Ingenieros
    
    lblNombreIng.Caption = Empty
    lblEstadoIng.Caption = Empty
    lblAsignacionFecha.Caption = Empty
    lblAsignacionUsuario.Caption = Empty
    lblEntregaFecha.Caption = Empty
    lblEntregaUsuario.Caption = Empty
    lblRecepcionFecha.Caption = Empty
    lblRecepcionUsuario.Caption = Empty
    lblRegistroFecha.Caption = Empty
    lblRegistroUsuario.Caption = Empty
 'Infomacion del proceso Abogados
    lblNombreAbog.Caption = Empty
    lblEstadoAbog.Caption = Empty
    lblAsignacionFechaAbog.Caption = Empty
    lblAsignacionUsuarioAbog.Caption = Empty
    lblEntregaFechaAbog.Caption = Empty
    lblEntregaUsuarioAbog.Caption = Empty
    lblFirmasFecha.Caption = Empty
    lblFirmasUsuario.Caption = Empty
    lblRegistroFechaAbog.Caption = Empty
    lblRegistroUsuarioAbog.Caption = Empty
    

Exit Sub
vError:
    Call ObjMensajes.deError("Ocurrió un error en visual basic al consultar la información según número de operación. Error " & Err.Description)
End Sub
Private Sub sbTraerInfoTramite()
On Error GoTo vError
 
If ObjConsultar.fxTraerInfoTramite(m_IdGarantia, "I") Then
    With glogon.Recordset.Fields
    lblFechaRegistro.Caption = ""
    If Not IsNull(!RegistroFecha) Then
        lblFechaRegistro.Caption = Format(!RegistroFecha, "dd/mm/yyyy hh:mm AMPM")
    End If
    
    lblUsuarioRegistro.Caption = !RegistroUsuario
    lblEstadoActual.Caption = Trim(!GEstado)
    
    'Infomacion del proceso Ingenieros
    
    lblNombreIng.Caption = !Nombre & ""

    lblEstadoIng.Caption = !EstadoProf & ""
    
    lblAsignacionFecha.Caption = ""
    If Not IsNull(!AsignacionFecha) Then
        lblAsignacionFecha.Caption = Format(!AsignacionFecha, "dd/mm/yyyy hh:mm AMPM")
    End If
    lblAsignacionUsuario.Caption = !AsignacionUsuario & ""
    
    lblEntregaFecha.Caption = ""
    If Not IsNull(!EntregaFecha) Then
       lblEntregaFecha.Caption = Format(!EntregaFecha, "dd/mm/yyyy hh:mm AMPM")
    End If
    lblEntregaUsuario.Caption = !EntregaUsuario & ""
     
    lblRecepcionFecha.Caption = ""
    If Not IsNull(!RecepcionFecha) Then
       lblRecepcionFecha.Caption = Format(!RecepcionFecha, "dd/mm/yyyy hh:mm AMPM")
    End If
    lblRecepcionUsuario.Caption = !RecepcionUsuario & ""
    
    lblRegistroFecha.Caption = ""
    If Not IsNull(!RegistroFechaProf) Then
       lblRegistroFecha.Caption = Format(!RegistroFechaProf, "dd/mm/yyyy hh:mm AMPM")
    End If
    lblRegistroUsuario.Caption = !RegistroUsuarioProf & ""
    
    End With
    
End If

 'Infomacion del proceso Abogados
If ObjConsultar.fxTraerInfoTramite(m_IdGarantia, "A") Then
    With glogon.Recordset.Fields
    
    lblNombreAbog.Caption = !Nombre & ""
    lblEstadoAbog.Caption = !EstadoProf & ""
    
    lblAsignacionFechaAbog.Caption = ""
    If Not IsNull(!AsignacionFecha) Then
       lblAsignacionFechaAbog.Caption = Format(!AsignacionFecha, "dd/mm/yyyy hh:mm AMPM")
    End If
    lblAsignacionUsuarioAbog.Caption = !AsignacionUsuario & ""
    
    lblEntregaFechaAbog.Caption = ""
    If Not IsNull(!EntregaFecha) Then
       lblEntregaFechaAbog.Caption = Format(!EntregaFecha, "dd/mm/yyyy hh:mm AMPM")
    End If
    lblEntregaUsuarioAbog.Caption = !EntregaUsuario & ""
    
    lblFirmasFecha.Caption = ""
    If Not IsNull(!firmasFecha) Then
       lblFirmasFecha.Caption = Format(!firmasFecha, "dd/mm/yyyy hh:mm AMPM")
    End If
    lblFirmasUsuario.Caption = !FirmasUsuario & ""
    
    lblRegistroFechaAbog.Caption = ""
    If Not IsNull(!RegistroFecha) Then
       lblRegistroFechaAbog.Caption = Format(!RegistroFechaProf, "dd/mm/yyyy hh:mm AMPM")
    End If
    lblRegistroUsuarioAbog.Caption = !RegistroUsuarioProf & ""
    End With
    
End If
salir:
    Exit Sub
vError:
    Call ObjMensajes.deError("Ocurrió un error en visual basic al consultar la información según número de operación. Error " & Err.Description)
End Sub

Private Sub sbTraerInformacionOperacion()
On Error GoTo vError

If txtOperacion.Text <> "0" Then
    glogon.strSQL = "SELECT S.CEDULA, S.NOMBRE, R.ID_SOLICITUD,R.CODIGO,  C.DESCRIPCION as 'DescLinea'" _
                    & ", ISNULL(P.COD_PREANALISIS, '') AS 'Expediente', R.ESTADOSOL" _
                    & " FROM REG_CREDITOS R" _
                    & "        INNER JOIN SOCIOS S   ON R.CEDULA = S.CEDULA" _
                    & "        INNER JOIN CATALOGO C ON R.CODIGO = C.CODIGO" _
                    & "    LEFT OUTER JOIN CRD_PREA_PREANALISIS P ON R.ID_SOLICITUD = P.ID_SOLICITUD" _
                    & " Where R.ID_SOLICITUD = " & txtOperacion.Text
Else
    glogon.strSQL = "SELECT S.CEDULA, S.NOMBRE, P.ID_SOLICITUD, P.COD_LINEA AS 'CODIGO',  C.DESCRIPCION as 'DescLinea'" _
                  & ", ISNULL(P.COD_PREANALISIS, '') AS 'Expediente', ISNULL(R.ESTADOSOL, 'P') AS 'ESTADOSOL'" _
                  & "  FROM CRD_PREA_PREANALISIS P" _
                  & "  INNER JOIN SOCIOS S   ON P.CEDULA = S.CEDULA" _
                  & "  INNER JOIN CATALOGO C ON P.COD_LINEA = C.CODIGO" _
                  & "   LEFT JOIN REG_CREDITOS R ON p.ID_SOLICITUD = R.ID_SOLICITUD" _
                  & "  Where P.COD_PREANALISIS  = '" & txtExpediente.Text & "'"
End If
                
Call OpenRecordSet(glogon.Recordset, glogon.strSQL)
With glogon.Recordset


 txtCedula.Text = Trim(!cedula)
 txtNombre.Text = Trim(!Nombre)
 txtExpediente.Text = Trim(!Expediente)
 m_EstadoSol = Trim(!ESTADOSOL)


End With

salir:
    Exit Sub
vError:
    Call ObjMensajes.deError("Ocurrió un error en visual basic al consultar la información según número de operación. Error " & Err.Description)
End Sub

Private Sub sbControlaAvaluoPosterior()
On Error GoTo vError
    tcAux.Item(1).Enabled = False
    If m_EstadoSol = "F" Then
        If ObjConsultar.fxTraerNumGarantiasOperacion(txtOperacion.Text) Then
            If (glogon.Recordset.Fields!cantidad) = 0 Then
            
                tcAux.Item(1).Enabled = True
                
                 
                
                txtViaticos.Text = Format(0, "Standard")
                txtValorConstruccion.Text = Format(0, "Standard")
                txtValorTerreno.Text = Format(0, "Standard")
                txtTotal.Text = Format(0, "Standard")
                dtpFechaInspeccion.Value = Format(fxFechaServidor, "DD/MM/YYYY") ' ObjConsultar.fxFechaServer
            End If
        End If
    End If

salir:
    Exit Sub
vError:
    Call ObjMensajes.deError("Ocurrió un error en visual basic al consultar la información según número de operación. Error " & Err.Description)
End Sub


Private Sub sbAgregarAvaluo(ByVal pIdGarantia As Long)
On Error GoTo vError

'If m_cambioDatos = False Then Exit Sub
If fxValidaDatosAvaluo(pIdGarantia) = False Then Exit Sub
'If (MsgBox("¿Desea guardar la información digitada.?", vbQuestion + vbYesNo) = vbNo) Then Exit Sub
Me.MousePointer = vbHourglass
If ObjAgregar.fxRegistroAvaluoPosterior(gParametros(1), gParametros(2), gParametros(3), gParametros(4), _
                                                   gParametros(5), gParametros(6), gParametros(7), gParametros(8), gParametros(9), gParametros(10), gParametros(11)) Then
    
Call Bitacora("APLICA", "Registro avaluo Garantia Vivienda: " & gParametros(1) & " Contacto: " & gParametros(2))
    
End If

salir:
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxValidaDatosAvaluo(Optional pIdGarantia As Long) As Boolean
On Error GoTo vError

fxValidaDatosAvaluo = False

If txtIdIngeniero.Text = Empty Then
        Me.MousePointer = vbDefault
        MsgBox ("Información de avaluo, el ingeniero no puede estar en blanco")
        Exit Function
End If

If txtIdAbogado.Text = Empty Then
        Me.MousePointer = vbDefault
        MsgBox ("Información de avaluo, el abogado no puede estar en blanco")
        Exit Function
End If

' Valida Exista el Ingeniero
If ObjConsultar.fxTraerExisteContacto(Trim(txtIdIngeniero), "I") Then
    If glogon.Recordset.Fields!cantidad = 0 Then
        Me.MousePointer = vbDefault
        MsgBox ("Información de avaluo, el ingeniero no existe")
        Exit Function
    End If
End If

' Valida Exista el Ingeniero
If ObjConsultar.fxTraerExisteContacto(Trim(txtIdAbogado), "A") Then
    If glogon.Recordset.Fields!cantidad = 0 Then
        Me.MousePointer = vbDefault
        MsgBox ("Información de avaluo, el abogado no existe")
        Exit Function
    End If
End If

ReDim gParametros(1 To 11)

' No aplica en avaluo posterior
'If fxValidaRegistroAvaluo(vIdGarantia) Then
'   Me.MousePointer = vbDefault
'    Msgbox ("La información de avaluo no puede ser modificada, ya fue registrado")
'    Exit Function
'End If

gParametros(1) = pIdGarantia
gParametros(2) = txtIdIngeniero
gParametros(3) = Format(dtpFechaInspeccion.Value, "yyyy/mm/dd")

If Not IsNumeric(txtValorTerreno.Text) Then
    gParametros(4) = 0
Else
    gParametros(4) = CCur(txtValorTerreno.Text)
End If
If Not IsNumeric(txtValorConstruccion.Text) Then
    gParametros(5) = 0
Else
    gParametros(5) = CCur(txtValorConstruccion.Text)
End If

gParametros(6) = IIf((Len(txtObservaciones.Text) = 0), ObjNull.NullString, Trim(txtObservaciones.Text))
gParametros(7) = glogon.Usuario
gParametros(8) = "1900/01/01"

If Not IsNumeric(txtViaticos.Text) Then
    gParametros(9) = 0
Else
    gParametros(9) = CCur(txtViaticos.Text)
End If

If optPersonal.Value = True Then
    gParametros(10) = "P"
Else
    gParametros(10) = "C"
End If

gParametros(11) = txtIdAbogado.Text

fxValidaDatosAvaluo = True

salir:
    Exit Function
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Function fxValidaRegistroAvaluo(ByVal pIdGarantia As Long) As Boolean

On Error GoTo vError

fxValidaRegistroAvaluo = False
                
glogon.strSQL = "SELECT G.ValorConstruccion, G.ValorTerreno" & _
                " FROM  ViviendaGarantia AS G" & _
                " where G.IdGarantia = " & pIdGarantia
          
                       
If execSql(glogon.strSQL, True) Then
    fxValidaRegistroAvaluo = IIf(IsNull(glogon.Recordset.Fields!ValorConstruccion), False, True)
End If
Exit Function

vError:
MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Function


Private Sub txtValorConstruccion_LostFocus()
    If Val(txtValorConstruccion.Text) = 0 Then
        txtValorConstruccion.Text = Format(0, "Standard")
    Else
        txtValorConstruccion.Text = Format(txtValorConstruccion.Text, "Standard")
    End If
End Sub

Private Sub txtValorTerreno_LostFocus()
    If Val(txtValorTerreno.Text) = 0 Then
        txtValorTerreno.Text = Format(0, "Standard")
    Else
        txtValorTerreno.Text = Format(txtValorTerreno.Text, "Standard")
    End If
End Sub

Private Sub txtViaticos_LostFocus()
    If Val(txtViaticos.Text) = 0 Then
        txtViaticos.Text = Format(0, "Standard")
    Else
        txtViaticos.Text = Format(txtViaticos.Text, "Standard")
    End If
End Sub

