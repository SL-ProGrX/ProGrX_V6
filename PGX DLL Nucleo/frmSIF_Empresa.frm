VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmSIF_Empresa 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Definición de la Empresa"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10305
   Icon            =   "frmSIF_Empresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9840
      Top             =   120
   End
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   615
      Left            =   8760
      TabIndex        =   0
      Top             =   7080
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Guardar"
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
      TextAlignment   =   1
      Appearance      =   14
      Picture         =   "frmSIF_Empresa.frx":3482
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   10095
      _Version        =   1441793
      _ExtentX        =   17806
      _ExtentY        =   10186
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
      ItemCount       =   8
      SelectedItem    =   7
      Item(0).Caption =   "Empresa"
      Item(0).ControlCount=   15
      Item(0).Control(0)=   "GroupBox3"
      Item(0).Control(1)=   "txtEmail"
      Item(0).Control(2)=   "txtAptoPostal"
      Item(0).Control(3)=   "txtFax"
      Item(0).Control(4)=   "txtTelefono"
      Item(0).Control(5)=   "txtCedJur"
      Item(0).Control(6)=   "txtNombre"
      Item(0).Control(7)=   "txtSitioWeb"
      Item(0).Control(8)=   "Label2(0)"
      Item(0).Control(9)=   "Label2(1)"
      Item(0).Control(10)=   "Label2(2)"
      Item(0).Control(11)=   "Label2(3)"
      Item(0).Control(12)=   "Label2(4)"
      Item(0).Control(13)=   "Label2(5)"
      Item(0).Control(14)=   "Label2(8)"
      Item(1).Caption =   "Pagaré"
      Item(1).ControlCount=   17
      Item(1).Control(0)=   "txtPagDomicilio"
      Item(1).Control(1)=   "txtPagCedJur"
      Item(1).Control(2)=   "txtPagNomCorto"
      Item(1).Control(3)=   "txtPagNomLargo"
      Item(1).Control(4)=   "txtPag_Seccion2"
      Item(1).Control(5)=   "Label2(9)"
      Item(1).Control(6)=   "Label2(10)"
      Item(1).Control(7)=   "Label2(11)"
      Item(1).Control(8)=   "Label2(12)"
      Item(1).Control(9)=   "txtPag_Seccion1"
      Item(1).Control(10)=   "gbPagaré"
      Item(1).Control(11)=   "GroupBox4"
      Item(1).Control(12)=   "txtRepresentante"
      Item(1).Control(13)=   "Label2(26)"
      Item(1).Control(14)=   "txtRepresentanteId"
      Item(1).Control(15)=   "Label2(27)"
      Item(1).Control(16)=   "txtRepresentanteCalidades"
      Item(2).Caption =   "Estado de Cuenta"
      Item(2).ControlCount=   13
      Item(2).Control(0)=   "chkEstadoCuenta"
      Item(2).Control(1)=   "txtEC_PiePagina"
      Item(2).Control(2)=   "chkVisiblePatrimonio"
      Item(2).Control(3)=   "chkVisibleFondos"
      Item(2).Control(4)=   "chkVisibleCreditos"
      Item(2).Control(5)=   "chkVisibleFianzas"
      Item(2).Control(6)=   "txtEC_Encabezado"
      Item(2).Control(7)=   "Label2(13)"
      Item(2).Control(8)=   "Label2(14)"
      Item(2).Control(9)=   "txtLIQ_LeyendaPago"
      Item(2).Control(10)=   "Label2(28)"
      Item(2).Control(11)=   "chkVisibleExcedentes"
      Item(2).Control(12)=   "chkVisibleDisponibles"
      Item(3).Caption =   "Misión/Visión"
      Item(3).ControlCount=   6
      Item(3).Control(0)=   "txtVision"
      Item(3).Control(1)=   "txtMision"
      Item(3).Control(2)=   "txtSlogan"
      Item(3).Control(3)=   "Label2(15)"
      Item(3).Control(4)=   "Label2(16)"
      Item(3).Control(5)=   "Label2(17)"
      Item(4).Caption =   "Logos"
      Item(4).ControlCount=   10
      Item(4).Control(0)=   "txtImagenLogo"
      Item(4).Control(1)=   "optImagenes(0)"
      Item(4).Control(2)=   "txtImagenFondo"
      Item(4).Control(3)=   "optImagenes(1)"
      Item(4).Control(4)=   "picImagen"
      Item(4).Control(5)=   "Label2(18)"
      Item(4).Control(6)=   "btnImagenes(0)"
      Item(4).Control(7)=   "btnImagenes(1)"
      Item(4).Control(8)=   "btnImagenes(2)"
      Item(4).Control(9)=   "btnImagenes(3)"
      Item(5).Caption =   "Consentimiento"
      Item(5).ControlCount=   4
      Item(5).Control(0)=   "txtConsentimiento_Texto"
      Item(5).Control(1)=   "txtConsentimiento_Titulo"
      Item(5).Control(2)=   "Label2(19)"
      Item(5).Control(3)=   "Label2(20)"
      Item(6).Caption =   "Constancias"
      Item(6).ControlCount=   9
      Item(6).Control(0)=   "txtConstanciaPAT"
      Item(6).Control(1)=   "txtConstanciaCRD"
      Item(6).Control(2)=   "txtConstanciaCRDPie"
      Item(6).Control(3)=   "txtConstanciaPATPie"
      Item(6).Control(4)=   "chkFechaVinculación"
      Item(6).Control(5)=   "Label2(21)"
      Item(6).Control(6)=   "Label2(22)"
      Item(6).Control(7)=   "Label2(23)"
      Item(6).Control(8)=   "Label2(24)"
      Item(7).Caption =   "Bloqueo de Fechas"
      Item(7).ControlCount=   3
      Item(7).Control(0)=   "GroupBox2"
      Item(7).Control(1)=   "GroupBox1"
      Item(7).Control(2)=   "Label1"
      Begin XtremeSuiteControls.FlatEdit txtRepresentanteCalidades 
         Height          =   1155
         Left            =   -67600
         TabIndex        =   88
         Top             =   2640
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11668
         _ExtentY        =   2037
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.PushButton btnImagenes 
         Height          =   312
         Index           =   0
         Left            =   -61120
         TabIndex        =   61
         Top             =   720
         Visible         =   0   'False
         Width           =   312
         _Version        =   1441793
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Picture         =   "frmSIF_Empresa.frx":3B79
      End
      Begin XtremeSuiteControls.CheckBox chkEstadoCuenta 
         Height          =   255
         Left            =   -63520
         TabIndex        =   45
         Top             =   3960
         Visible         =   0   'False
         Width           =   5895
         _Version        =   1441793
         _ExtentX        =   10393
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Utilizar Estado de Cuenta Comercial"
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.GroupBox gbPagaré 
         Height          =   1695
         Left            =   -69880
         TabIndex        =   37
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   2984
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin XtremeSuiteControls.RadioButton rbPagareSeccion 
            Height          =   372
            Index           =   0
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Sección 1   "
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
            Value           =   -1  'True
            Alignment       =   1
         End
         Begin XtremeSuiteControls.RadioButton rbPagareSeccion 
            Height          =   372
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   600
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Sección 2   "
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
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1335
         Left            =   -69760
         TabIndex        =   2
         Top             =   4440
         Visible         =   0   'False
         Width           =   9615
         _Version        =   1441793
         _ExtentX        =   16954
         _ExtentY        =   2350
         _StockProps     =   79
         Caption         =   "Contabilidad de enlace:"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cbo 
            Height          =   312
            Left            =   2280
            TabIndex        =   3
            Top             =   360
            Width           =   6612
            _Version        =   1441793
            _ExtentX        =   11668
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtCtaDesc 
            Height          =   312
            Left            =   4560
            TabIndex        =   25
            Top             =   720
            Width           =   4332
            _Version        =   1441793
            _ExtentX        =   7641
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtCtaCod 
            Height          =   312
            Left            =   2280
            TabIndex        =   26
            Top             =   720
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   7
            Left            =   600
            TabIndex        =   17
            Top             =   720
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cuenta Default"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   6
            Left            =   600
            TabIndex        =   16
            Top             =   360
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Contabilidad"
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
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1215
         Left            =   720
         TabIndex        =   4
         Top             =   3480
         Width           =   8535
         _Version        =   1441793
         _ExtentX        =   15049
         _ExtentY        =   2138
         _StockProps     =   79
         Caption         =   "Des Bloqueo de Fecha"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton cmdDesbloquear 
            Height          =   612
            Left            =   4320
            TabIndex        =   5
            Top             =   360
            Width           =   2892
            _Version        =   1441793
            _ExtentX        =   5101
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Desbloquear"
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
            TextAlignment   =   1
            Appearance      =   16
            Picture         =   "frmSIF_Empresa.frx":4279
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.Label lblFechaBloqueada 
            Height          =   612
            Left            =   0
            TabIndex        =   79
            Top             =   360
            Width           =   3852
            _Version        =   1441793
            _ExtentX        =   6794
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "..."
            ForeColor       =   16777215
            BackColor       =   16761024
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
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1095
         Left            =   720
         TabIndex        =   6
         Top             =   1920
         Width           =   8535
         _Version        =   1441793
         _ExtentX        =   15049
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Bloqueo de Fecha:"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton cmdBloquearFecha 
            Height          =   612
            Left            =   4320
            TabIndex        =   7
            Top             =   480
            Width           =   2892
            _Version        =   1441793
            _ExtentX        =   5101
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Bloqueo de Fecha en Auxiliares"
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
            TextAlignment   =   1
            Appearance      =   16
            Picture         =   "frmSIF_Empresa.frx":4C06
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.DateTimePicker dtpBloqueo 
            Height          =   312
            Left            =   2520
            TabIndex        =   8
            Top             =   480
            Width           =   1332
            _Version        =   1441793
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   25
            Left            =   1800
            TabIndex        =   78
            Top             =   240
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3619
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Fecha a Bloquear:"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtCedJur 
         Height          =   315
         Left            =   -67480
         TabIndex        =   19
         Top             =   1440
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4043
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   315
         Left            =   -67480
         TabIndex        =   20
         Top             =   2520
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11663
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtSitioWeb 
         Height          =   315
         Left            =   -67480
         TabIndex        =   21
         Top             =   2880
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11663
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtAptoPostal 
         Height          =   315
         Left            =   -63160
         TabIndex        =   22
         Top             =   1440
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4043
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtFax 
         Height          =   315
         Left            =   -63160
         TabIndex        =   23
         Top             =   1800
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4043
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono 
         Height          =   315
         Left            =   -67480
         TabIndex        =   24
         Top             =   1800
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4043
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPagNomLargo 
         Height          =   312
         Left            =   -67600
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1441793
         _ExtentX        =   11663
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPagNomCorto 
         Height          =   312
         Left            =   -67600
         TabIndex        =   32
         Top             =   960
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1441793
         _ExtentX        =   11663
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPagCedJur 
         Height          =   312
         Left            =   -67600
         TabIndex        =   33
         Top             =   1320
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1441793
         _ExtentX        =   11663
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPagDomicilio 
         Height          =   795
         Left            =   -67600
         TabIndex        =   34
         Top             =   1680
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11668
         _ExtentY        =   1402
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPag_Seccion2 
         Height          =   1755
         Left            =   -67600
         TabIndex        =   35
         Top             =   3960
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11663
         _ExtentY        =   3090
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPag_Seccion1 
         Height          =   1755
         Left            =   -67600
         TabIndex        =   36
         Top             =   3960
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11663
         _ExtentY        =   3090
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtEC_Encabezado 
         Height          =   915
         Left            =   -67840
         TabIndex        =   41
         Top             =   600
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11668
         _ExtentY        =   1614
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtEC_PiePagina 
         Height          =   915
         Left            =   -67840
         TabIndex        =   42
         Top             =   1680
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11668
         _ExtentY        =   1614
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkVisiblePatrimonio 
         Height          =   255
         Left            =   -67840
         TabIndex        =   46
         Top             =   2880
         Visible         =   0   'False
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Visualizar Sección de Patrimonio"
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkVisibleFondos 
         Height          =   255
         Left            =   -67840
         TabIndex        =   47
         Top             =   3240
         Visible         =   0   'False
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Visualizar Sección de Fondos"
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkVisibleCreditos 
         Height          =   255
         Left            =   -67840
         TabIndex        =   48
         Top             =   3600
         Visible         =   0   'False
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Visualizar Sección de Créditos y Retenciones"
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkVisibleFianzas 
         Height          =   255
         Left            =   -67840
         TabIndex        =   49
         Top             =   3960
         Visible         =   0   'False
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Visualizar Sección de Fianzas"
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtMision 
         Height          =   1152
         Left            =   -67720
         TabIndex        =   50
         Top             =   600
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1441793
         _ExtentX        =   11663
         _ExtentY        =   2032
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtVision 
         Height          =   1155
         Left            =   -67720
         TabIndex        =   51
         Top             =   1920
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11663
         _ExtentY        =   2032
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtSlogan 
         Height          =   1155
         Left            =   -67720
         TabIndex        =   54
         Top             =   3240
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11663
         _ExtentY        =   2032
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtImagenLogo 
         Height          =   312
         Left            =   -67840
         TabIndex        =   56
         Top             =   720
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1441793
         _ExtentX        =   11663
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtImagenFondo 
         Height          =   312
         Left            =   -67840
         TabIndex        =   57
         Top             =   1320
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1441793
         _ExtentX        =   11663
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton optImagenes 
         Height          =   492
         Index           =   0
         Left            =   -70000
         TabIndex        =   58
         Top             =   600
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Logo   "
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
         Value           =   -1  'True
         Alignment       =   1
      End
      Begin XtremeSuiteControls.RadioButton optImagenes 
         Height          =   492
         Index           =   1
         Left            =   -70000
         TabIndex        =   59
         Top             =   1200
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Fondo   "
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton btnImagenes 
         Height          =   312
         Index           =   1
         Left            =   -60760
         TabIndex        =   62
         Top             =   720
         Visible         =   0   'False
         Width           =   312
         _Version        =   1441793
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Picture         =   "frmSIF_Empresa.frx":52EE
      End
      Begin XtremeSuiteControls.PushButton btnImagenes 
         Height          =   312
         Index           =   2
         Left            =   -61120
         TabIndex        =   63
         Top             =   1320
         Visible         =   0   'False
         Width           =   312
         _Version        =   1441793
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Picture         =   "frmSIF_Empresa.frx":5A1F
      End
      Begin XtremeSuiteControls.PushButton btnImagenes 
         Height          =   312
         Index           =   3
         Left            =   -60760
         TabIndex        =   64
         Top             =   1320
         Visible         =   0   'False
         Width           =   312
         _Version        =   1441793
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Picture         =   "frmSIF_Empresa.frx":611F
      End
      Begin XtremeSuiteControls.FlatEdit txtConsentimiento_Titulo 
         Height          =   912
         Left            =   -69280
         TabIndex        =   65
         Top             =   840
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1441793
         _ExtentX        =   16108
         _ExtentY        =   1609
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtConsentimiento_Texto 
         Height          =   3315
         Left            =   -69280
         TabIndex        =   66
         Top             =   2280
         Visible         =   0   'False
         Width           =   9135
         _Version        =   1441793
         _ExtentX        =   16113
         _ExtentY        =   5847
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtConstanciaCRD 
         Height          =   675
         Left            =   -69520
         TabIndex        =   69
         Top             =   840
         Visible         =   0   'False
         Width           =   9135
         _Version        =   1441793
         _ExtentX        =   16108
         _ExtentY        =   1185
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtConstanciaCRDPie 
         Height          =   675
         Left            =   -69520
         TabIndex        =   72
         Top             =   2040
         Visible         =   0   'False
         Width           =   9135
         _Version        =   1441793
         _ExtentX        =   16108
         _ExtentY        =   1185
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtConstanciaPAT 
         Height          =   675
         Left            =   -69520
         TabIndex        =   73
         Top             =   3360
         Visible         =   0   'False
         Width           =   9135
         _Version        =   1441793
         _ExtentX        =   16108
         _ExtentY        =   1185
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtConstanciaPATPie 
         Height          =   675
         Left            =   -69520
         TabIndex        =   76
         Top             =   4680
         Visible         =   0   'False
         Width           =   9135
         _Version        =   1441793
         _ExtentX        =   16108
         _ExtentY        =   1185
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkFechaVinculación 
         Height          =   255
         Left            =   -65800
         TabIndex        =   77
         Top             =   480
         Visible         =   0   'False
         Width           =   5415
         _Version        =   1441793
         _ExtentX        =   9546
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Incluir la fecha de registro en la empresa?   "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         TextAlignment   =   1
         Appearance      =   2
         Alignment       =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   1215
         Left            =   -69880
         TabIndex        =   81
         Top             =   2520
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   2143
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin XtremeSuiteControls.RadioButton rbRepresentante 
            Height          =   372
            Index           =   0
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Representante"
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
            Value           =   -1  'True
            Alignment       =   1
         End
         Begin XtremeSuiteControls.RadioButton rbRepresentante 
            Height          =   372
            Index           =   1
            Left            =   120
            TabIndex        =   83
            Top             =   600
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Calidades  "
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
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtRepresentante 
         Height          =   315
         Left            =   -67600
         TabIndex        =   84
         Top             =   2880
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11663
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtRepresentanteId 
         Height          =   315
         Left            =   -67600
         TabIndex        =   86
         Top             =   3480
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11663
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   555
         Left            =   -67480
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11668
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtLIQ_LeyendaPago 
         Height          =   915
         Left            =   -67840
         TabIndex        =   89
         Top             =   4440
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1441793
         _ExtentX        =   11668
         _ExtentY        =   1614
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkVisibleExcedentes 
         Height          =   255
         Left            =   -63520
         TabIndex        =   91
         Top             =   2880
         Visible         =   0   'False
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Visualizar Sección de Excedentes"
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkVisibleDisponibles 
         Height          =   255
         Left            =   -63520
         TabIndex        =   92
         Top             =   3240
         Visible         =   0   'False
         Width           =   5895
         _Version        =   1441793
         _ExtentX        =   10393
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Visualizar Sección de Disponibles"
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   1215
         Index           =   28
         Left            =   -69520
         TabIndex        =   90
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Leyenda para la Boleta de Liquidación. Sección de Pago"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   27
         Left            =   -67600
         TabIndex        =   87
         Top             =   3240
         Visible         =   0   'False
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4895
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   26
         Left            =   -67600
         TabIndex        =   85
         Top             =   2640
         Visible         =   0   'False
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nombre del Representante Legal"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   972
         Left            =   600
         TabIndex        =   80
         Top             =   600
         Width           =   8772
         _Version        =   1441793
         _ExtentX        =   15473
         _ExtentY        =   1714
         _StockProps     =   79
         Caption         =   $"frmSIF_Empresa.frx":6850
         ForeColor       =   128
         BackColor       =   12640511
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   135
         Index           =   24
         Left            =   -69880
         TabIndex        =   75
         Top             =   4440
         Visible         =   0   'False
         Width           =   5295
         _Version        =   1441793
         _ExtentX        =   9334
         _ExtentY        =   233
         _StockProps     =   79
         Caption         =   "Notas al Cierre para Constancias de Aportes.:"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   135
         Index           =   23
         Left            =   -69880
         TabIndex        =   74
         Top             =   3120
         Visible         =   0   'False
         Width           =   5295
         _Version        =   1441793
         _ExtentX        =   9334
         _ExtentY        =   233
         _StockProps     =   79
         Caption         =   "Encabezado para Constancias de Aportes.:"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   135
         Index           =   22
         Left            =   -69880
         TabIndex        =   71
         Top             =   1800
         Visible         =   0   'False
         Width           =   5295
         _Version        =   1441793
         _ExtentX        =   9334
         _ExtentY        =   233
         _StockProps     =   79
         Caption         =   "Notas al Cierre para Constancias de Crédito.:"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   135
         Index           =   21
         Left            =   -69880
         TabIndex        =   70
         Top             =   600
         Visible         =   0   'False
         Width           =   5295
         _Version        =   1441793
         _ExtentX        =   9334
         _ExtentY        =   233
         _StockProps     =   79
         Caption         =   "Encabezado para Constancias de Créditos.:"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   20
         Left            =   -69760
         TabIndex        =   68
         Top             =   1920
         Visible         =   0   'False
         Width           =   3732
         _Version        =   1441793
         _ExtentX        =   6583
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuerpo del documento de consentimiento: "
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   19
         Left            =   -69760
         TabIndex        =   67
         Top             =   480
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Título del documento:"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   18
         Left            =   -66880
         TabIndex        =   60
         Top             =   1800
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Vista Preliminar.:"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   17
         Left            =   -69160
         TabIndex        =   55
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Slogan"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   16
         Left            =   -69160
         TabIndex        =   53
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Visión"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   15
         Left            =   -69160
         TabIndex        =   52
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Misión"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   14
         Left            =   -69280
         TabIndex        =   44
         Top             =   1680
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Pie de Página"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   13
         Left            =   -69280
         TabIndex        =   43
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Encabezado"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   12
         Left            =   -69280
         TabIndex        =   30
         Top             =   1680
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Domicilio"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   11
         Left            =   -69280
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ced.Jur. Letras"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   10
         Left            =   -69280
         TabIndex        =   28
         Top             =   960
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Nombre Corto"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   9
         Left            =   -69280
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Nombre Largo"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   8
         Left            =   -64600
         TabIndex        =   18
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tel. Fax."
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   -68560
         TabIndex        =   15
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Sitio Web"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   4
         Left            =   -68560
         TabIndex        =   14
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "E-Mail"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   3
         Left            =   -64600
         TabIndex        =   13
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Apt.Postal"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   -68560
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Teléfono"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   -68560
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ced.Jur."
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   0
         Left            =   -68560
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
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
      Begin VB.Image picImagen 
         Appearance      =   0  'Flat
         Height          =   2652
         Left            =   -64960
         Stretch         =   -1  'True
         Top             =   1800
         Visible         =   0   'False
         Width           =   3732
      End
   End
   Begin XtremeSuiteControls.Label lblName 
      Height          =   372
      Left            =   3480
      TabIndex        =   40
      Top             =   360
      Width           =   6612
      _Version        =   1441793
      _ExtentX        =   11663
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Empresa"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
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
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmSIF_Empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean

Private Sub btnImagenes_Click(Index As Integer)
Dim strSQL As String

On Error GoTo vError

Select Case Index
  Case 0, 2 'Buscar
        frmContenedor.CD.ShowOpen
        frmContenedor.CD.DialogTitle = "Buscar Imagen..."
        frmContenedor.CD.InitDir = "C:\"
        
        'Indica el mismo index del toolbar que el tipo de imagen seleccionado para coincidir
        optImagenes.Item(Index).Value = True
        
        Select Case Index
          Case 0 'Logo
            txtImagenLogo.Text = frmContenedor.CD.FileName
          Case 2 'Fondo de Pantalla
            txtImagenFondo.Text = frmContenedor.CD.FileName
        End Select
        
        picImagen.Picture = LoadPicture(frmContenedor.CD.FileName)
  
  Case 1, 3 'guardar"
    strSQL = "select * from SIF_Empresa"
    
    Select Case True
      Case optImagenes.Item(0).Value 'Logo
        
        If Trim(txtImagenLogo.Text) <> "" Then
           If Not fxImagen_Guardar(strSQL, "Logo", txtImagenLogo.Text) Then
              MsgBox "Imagen del Logo -> no guardada!!!", vbExclamation
           End If
        End If
        
        
      Case optImagenes.Item(1).Value 'Fondo
        If Trim(txtImagenFondo.Text) <> "" Then
           If Not fxImagen_Guardar(strSQL, "Fondo_Pantalla", txtImagenFondo.Text) Then
              MsgBox "Imagen del Fondo de Pantalla -> no guardada!!!", vbExclamation
           End If
        End If
    End Select
  
End Select 'Index

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdBloquearFecha_Click()
Dim strSQL As String, i As Integer

On Error GoTo vError


i = MsgBox("Esta seguro que desea BLOQUEAR la fecha de los auxiliares", vbYesNo)
If i = vbNo Then Exit Sub
  
  
strSQL = "exec spSys_BloqueoFechaAuxiliar '" & Format(dtpBloqueo.Value, "yyyy/mm/dd") & " 22:00:00','B','" & glogon.Usuario & "'"
glogon.Conection.Execute strSQL

Call Bitacora("Aplica", "Bloquea Fecha Auxiliar: " & Format(dtpBloqueo.Value, "yyyy/mm/dd"))

MsgBox "Bloqueo de Fecha de Auxiliar establecida: " & Format(dtpBloqueo.Value, "dd/mm/yyyy"), vbInformation


Call Form_Load

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub

Private Sub cmdDesbloquear_Click()
Dim strSQL As String, i As Integer

On Error GoTo vError

i = MsgBox("Esta seguro que desea DES-Bloquear la fecha de los auxiliares y utilizar la actual", vbYesNo)
If i = vbNo Then Exit Sub

strSQL = "exec spSys_BloqueoFechaAuxiliar '" & Format(dtpBloqueo.Value, "yyyy/mm/dd") & " 22:00:00','D', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "DES-Bloqueo Fecha Auxiliar")

MsgBox "DES-Bloqueo de Fecha de auxiliares realizado satisfactoriamente. Ahora utiliza fecha actual!", vbInformation


Call Form_Load

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdGuardar_Click()
Dim strSQL As String

On Error GoTo vError


If vEdita Then
   strSQL = "update sif_empresa set nombre = '" & Trim(txtNombre) & "',email = '" & Trim(txtEmail) & "', apto_postal = '" _
          & Trim(txtAptoPostal) & "',cedula_juridica = '" & Trim(txtCedJur) & "',telefonoemp = '" & Trim(txtTelefono) & "',fax = '" _
          & Trim(txtFax) & "',EstadoCuenta = '" & IIf((chkEstadoCuenta.Value = vbChecked), "C", "S") _
          & "',PAG_NOMLARGO = '" & Trim(txtPagNomLargo) & "',pag_NomCorto = '" & Trim(txtPagNomCorto) _
          & "',PAG_CedJurLE = '" & Trim(txtPagCedJur) & "',PAG_Domicilio = '" & Trim(txtPagDomicilio) _
          & "',cod_empresa_enlace = " & cbo.ItemData(cbo.ListIndex) & ",ec_Nota01 = '" & Trim(txtEC_Encabezado) _
          & "',ec_Nota02 = '" & Trim(txtEC_PiePagina) & "', ec_Visible_Patrimonio = " & chkVisiblePatrimonio.Value & ", EC_VISIBLE_EXCEDENTES = " & chkVisibleExcedentes.Value _
          & ",ec_visible_creditos = " & chkVisibleCreditos.Value & ",ec_visible_fondos = " & chkVisibleFondos.Value & ", EC_VISIBLE_DISPONIBLE = " & chkVisibleDisponibles.Value _
          & ",ec_visible_fianzas = " & chkVisibleFianzas.Value & ",Sitio_Web = '" & txtSitioWeb.Text _
          & "',Mision = '" & txtMision.Text & "', Vision = '" & txtVision.Text & "', Slogan = '" & txtSlogan.Text _
          & "',Pag_Seccion_01 = '" & txtPag_Seccion1.Text & "',Pag_Seccion_02 = '" & txtPag_Seccion2.Text & "'" _
          & ",Consentimiento_Contacto_Titulo = '" & txtConsentimiento_Titulo.Text _
          & "', Consentimiento_Contacto_Texto = '" & txtConsentimiento_Texto.Text & "', LIQ_BOLETA_PIE = '" & txtLIQ_LeyendaPago.Text & "'" _
          & ",CONSTANCIA_CRD_ENCABEZADO = '" & txtConstanciaCRD.Text & "',CONSTANCIA_PAT_ENCABEZADO = '" & txtConstanciaPAT.Text & "'" _
          & ",CONSTANCIA_CRD_PIE = '" & txtConstanciaCRDPie.Text & "', CONSTANCIA_PAT_PIE = '" & txtConstanciaPATPie.Text _
          & "', CONSTANCIA_FECHA_VINCULACION = " & chkFechaVinculación.Value _
          & ", COD_CUENTA_NO_CFG = '" & fxgCntCuentaFormato(False, txtCtaCod.Text, 0) & "', REPRESENTANTE_LEGAL = '" & Trim(txtRepresentante.Text) _
          & "', REPRESENTANTE_ID = '" & Trim(txtRepresentanteId.Text) & "', REPRESENTANTE_CALIDADES = '" & Trim(txtRepresentanteCalidades.Text) & "'"


Else
   strSQL = "insert into sif_empresa(nombre,cedula_juridica,email," _
          & "telefonoemp,fax,apto_postal,cod_empresa_enlace,EstadoCuenta,pag_nomLargo,pag_nomCorto" _
          & ",pag_CedJurLe,pag_Domicilio,EC_Nota01,EC_Nota02,Ec_visible_creditos,Ec_visible_patrimonio, EC_VISIBLE_EXCEDENTES" _
          & ",Ec_visible_fondos,Ec_visible_fianzas, EC_VISIBLE_DISPONIBLE, Sitio_Web,Mision,Vision,Slogan,Pag_Seccion_01,Pag_Seccion_02" _
          & ",Consentimiento_Contacto_Titulo,Consentimiento_Contacto_Texto,CONSTANCIA_PAT_ENCABEZADO" _
          & ",CONSTANCIA_CRD_ENCABEZADO,CONSTANCIA_CRD_PIE, CONSTANCIA_PAT_PIE,  CONSTANCIA_FECHA_VINCULACION" _
          & ", REPRESENTANTE_LEGAL, REPRESENTANTE_ID, REPRESENTANTE_CALIDADES, LIQ_BOLETA_PIE) values('" _
          & txtNombre & "','" & txtCedJur & "','" & txtEmail & "','" _
          & txtTelefono & "','" & txtFax & "','" & txtAptoPostal _
          & "'," & cbo.ItemData(cbo.ListIndex) & ",'" _
          & IIf((chkEstadoCuenta.Value = vbChecked), "C", "S") & "','" & txtPagNomLargo & "','" & txtPagNomCorto _
          & "','" & txtPagCedJur & "','" & txtPagDomicilio & "','" & txtEC_Encabezado & "','" & txtEC_PiePagina _
          & "'," & chkVisibleCreditos.Value & "," & chkVisiblePatrimonio.Value & ", " & chkVisibleExcedentes.Value & ", " & chkVisibleFondos.Value _
          & "," & chkVisibleFianzas.Value & ", " & chkVisibleDisponibles.Value & ",'" & txtSitioWeb.Text & "','" & txtMision.Text & "','" _
          & txtVision.Text & "','" & txtSlogan.Text & "','" & txtPag_Seccion1.Text & "','" & txtPag_Seccion2.Text _
          & "','" & txtConsentimiento_Titulo.Text & "','" & txtConsentimiento_Texto.Text _
          & "','" & txtConstanciaPAT.Text & "','" & txtConstanciaCRD.Text & "','" & txtConstanciaCRDPie.Text _
          & "','" & txtConstanciaPATPie.Text & "'," & chkFechaVinculación.Value _
          & ", '" & Trim(txtRepresentante.Text) & "','" & Trim(txtRepresentanteId.Text) & "','" & Trim(txtRepresentanteCalidades.Text) _
          & "', '" & txtLIQ_LeyendaPago.Text & "')"
End If

Call ConectionExecute(strSQL)
GLOBALES.gEnlace = cbo.ItemData(cbo.ListIndex)

MsgBox "Infomación Guardada Satisfactoriamente...", vbInformation

UnLoad Me
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

On Error GoTo vError

'Carga Combo de Contabilidades
strSQL = "select COD_CONTABILIDAD as 'Idx', rtrim(NOMBRE) as 'ItmX'  from CNTX_CONTABILIDADES"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

If GLOBALES.gEnlace = 0 Then
   
   vEdita = False
      
   
Else 'Editar
  vEdita = True
  strSQL = "select Ns.*,getdate() as 'Fecha_Actual'" _
         & ", isnull(Cta.Cod_Cuenta_Mask,'') as 'Cta_Codigo', isnull(Cta.Descripcion,'') as 'Cta_Desc'" _
         & " from sif_empresa Ns left join vCNTX_CUENTAS_LOCAL Cta on Ns.COD_EMPRESA_ENLACE = Cta.COD_CONTABILIDAD " _
         & " and Ns.COD_CUENTA_NO_CFG = Cta.cod_Cuenta "
  Call OpenRecordSet(rs, strSQL)
  
  txtNombre.Text = rs!Nombre & ""
  txtCedJur.Text = Trim(rs!cedula_juridica & "")
  txtTelefono.Text = Trim(rs!telefonoemp & "")
  txtFax.Text = Trim(rs!fax & "")
  txtAptoPostal.Text = Trim(rs!apto_postal & "")
  txtEmail.Text = Trim(rs!Email & "")
  txtSitioWeb.Text = Trim(rs!Sitio_Web & "")
  
  
  txtPagNomLargo.Text = Trim(rs!pag_nomlargo & "")
  txtPagNomCorto.Text = Trim(rs!pag_nomCorto & "")
  txtPagCedJur.Text = Trim(rs!pag_cedJurLe & "")
  txtPagDomicilio.Text = Trim(rs!pag_domicilio & "")
  txtPag_Seccion1.Text = Trim(rs!Pag_Seccion_01 & "")
  txtPag_Seccion2.Text = Trim(rs!Pag_Seccion_02 & "")
  
  txtRepresentante.Text = Trim(rs!REPRESENTANTE_LEGAL & "")
  txtRepresentanteId.Text = Trim(rs!REPRESENTANTE_ID & "")
  txtRepresentanteCalidades.Text = Trim(rs!REPRESENTANTE_CALIDADES & "")
  
  txtCtaCod.Text = rs!Cta_Codigo
  txtCtaDesc.Text = rs!Cta_Desc
  
  chkEstadoCuenta.Value = IIf((rs!estadoCuenta = "S"), vbUnchecked, vbChecked)
    
  chkVisibleCreditos.Value = rs!ec_visible_creditos
  chkVisiblePatrimonio.Value = rs!ec_visible_patrimonio
  chkVisibleFondos.Value = rs!ec_visible_fondos
  chkVisibleFianzas.Value = rs!ec_visible_fianzas
  chkVisibleExcedentes.Value = rs!EC_VISIBLE_EXCEDENTES
  chkVisibleDisponibles.Value = rs!EC_VISIBLE_DISPONIBLE
  
  txtEC_Encabezado.Text = rs!EC_Nota01 & ""
  txtEC_PiePagina.Text = rs!EC_Nota02 & ""
  txtLIQ_LeyendaPago.Text = rs!LIQ_BOLETA_PIE & ""
  
  txtMision.Text = rs!Mision & ""
  txtVision.Text = rs!Vision & ""
  txtSlogan.Text = rs!Slogan & ""

  txtConsentimiento_Titulo.Text = rs!Consentimiento_Contacto_Titulo & ""
  txtConsentimiento_Texto.Text = rs!Consentimiento_Contacto_Texto & ""
  
  txtConstanciaCRD.Text = rs!CONSTANCIA_CRD_ENCABEZADO & ""
  txtConstanciaCRDPie.Text = rs!CONSTANCIA_CRD_PIE & ""
  
  chkFechaVinculación.Value = rs!CONSTANCIA_FECHA_VINCULACION
      
  txtConstanciaPAT.Text = rs!CONSTANCIA_PAT_ENCABEZADO & ""
  txtConstanciaPATPie.Text = rs!CONSTANCIA_PAT_PIE & ""
  
  
  
  If IsNull(rs!Fecha_Congela) Then
     lblFechaBloqueada.Caption = "No Existe Ningún bloqueo de fecha establecido!"
     dtpBloqueo.Value = rs!Fecha_Actual
     lblFechaBloqueada.Tag = 0
  Else
     lblFechaBloqueada.Caption = "BLOQUEO DE FECHA ESTABLECIDO EN: " & Format(rs!Fecha_Congela, "dd/mm/yyyy")
     dtpBloqueo.Value = rs!Fecha_Congela
     lblFechaBloqueada.Tag = 1
  End If
  
  
   strSQL = "select COD_CONTABILIDAD, rtrim(NOMBRE) as 'Nombre'  from CNTX_CONTABILIDADES" _
          & " where COD_CONTABILIDAD in(select COD_EMPRESA_ENLACE from SIF_EMPRESA)"
  
  rs.Close
  
  Call OpenRecordSet(rs, strSQL)
    Call sbCboAsignaDato(cbo, rs!Nombre, True, rs!Cod_Contabilidad)
  rs.Close

  Call rbPagareSeccion_Click(0)
  Call rbRepresentante_Click(0)
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub Form_Load()
vModulo = 10


Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

tcMain.Item(0).Selected = True


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub optImagenes_Click(Index As Integer)
Call sbLeeImagen
End Sub


Private Sub rbPagareSeccion_Click(Index As Integer)
Select Case True
  Case rbPagareSeccion.Item(0).Value    'Seccion 1
    txtPag_Seccion1.Visible = True
    txtPag_Seccion2.Visible = False
  Case rbPagareSeccion.Item(1).Value 'Seccion 2
    txtPag_Seccion1.Visible = False
    txtPag_Seccion2.Visible = True
End Select
End Sub

Private Sub sbLeeImagen()
Dim strSQL As String

strSQL = "select Logo, Fondo_Pantalla from SIF_Empresa"

Select Case True
   Case optImagenes.Item(0).Value 'Logo
     Set picImagen.Picture = fxImagen_Leer(strSQL, "Logo")
   Case optImagenes.Item(1).Value 'Fondo de Pantalla
     Set picImagen.Picture = fxImagen_Leer(strSQL, "Fondo_Pantalla")
End Select

End Sub


Private Sub rbRepresentante_Click(Index As Integer)
Select Case True
  Case rbRepresentante.Item(0).Value
    txtRepresentanteCalidades.Visible = False
  Case rbRepresentante.Item(1).Value
    txtRepresentanteCalidades.Visible = True

End Select
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
lblName.Caption = Item.Caption

On Error GoTo vError

cmdGuardar.Visible = True

Select Case Item.Index
    Case 0 'Empresa
        txtNombre.SetFocus
    
    Case 1 'Pagaré
        Call rbPagareSeccion_Click(0)
        Call rbRepresentante_Click(0)
        
        txtPagNomLargo.SetFocus
    
    Case 2 'Estado de Cuenta
        txtEC_Encabezado.SetFocus
        
    Case 3 'Mision/Vision
        txtMision.SetFocus
    
    Case 4 'Logos
        txtImagenLogo.SetFocus
        cmdGuardar.Visible = False
        Call sbLeeImagen
        
    Case 5 'Consentimiento
        txtConsentimiento_Titulo.SetFocus
        
    Case 6 'Constancias
        txtConstanciaCRD.SetFocus
        
    Case 7 'Fechas
        cmdGuardar.Visible = False

End Select


Exit Sub

vError:


End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub

Private Sub txtCtaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaCod.Text = gCuenta
   txtCtaDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCtaCod.Text = fxgCntCuentaFormato(True, txtCtaCod, 0)
End If

End Sub
