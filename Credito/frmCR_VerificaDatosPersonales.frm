VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCR_VerificaDatosPersonales 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualización de Información Personal"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11970
   HelpContextID   =   3029
   Icon            =   "frmCR_VerificaDatosPersonales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lswCRM 
      Height          =   3975
      Left            =   8160
      TabIndex        =   56
      Top             =   2400
      Width           =   3735
      _Version        =   1572864
      _ExtentX        =   6588
      _ExtentY        =   7011
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
      Checkboxes      =   -1  'True
      View            =   2
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.TabControl tcCRM 
      Height          =   405
      Left            =   8160
      TabIndex        =   57
      Top             =   2040
      Width           =   3735
      _Version        =   1572864
      _ExtentX        =   6588
      _ExtentY        =   714
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
      Item(0).Caption =   "Bienes"
      Item(0).ControlCount=   0
      Item(1).Caption =   "Canales"
      Item(1).ControlCount=   0
      Item(2).Caption =   "Gustos"
      Item(2).ControlCount=   0
      Item(3).Caption =   "Escolaridad"
      Item(3).ControlCount=   0
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4332
      Left            =   0
      TabIndex        =   14
      Top             =   2040
      Width           =   8172
      _Version        =   1572864
      _ExtentX        =   14414
      _ExtentY        =   7641
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
      Item(0).Caption =   "Contacto"
      Item(0).ControlCount=   21
      Item(0).Control(0)=   "cboSexo"
      Item(0).Control(1)=   "cboEstadoCivil"
      Item(0).Control(2)=   "Label1(0)"
      Item(0).Control(3)=   "Label14"
      Item(0).Control(4)=   "cboProvincia"
      Item(0).Control(5)=   "cboCanton"
      Item(0).Control(6)=   "cboDistrito"
      Item(0).Control(7)=   "cboNacionalidad"
      Item(0).Control(8)=   "txtEmail"
      Item(0).Control(9)=   "txtEmail_02"
      Item(0).Control(10)=   "txtDireccion"
      Item(0).Control(11)=   "txtNotificaciones"
      Item(0).Control(12)=   "Label18(2)"
      Item(0).Control(13)=   "Label(25)"
      Item(0).Control(14)=   "Label7"
      Item(0).Control(15)=   "Label10(9)"
      Item(0).Control(16)=   "Label11"
      Item(0).Control(17)=   "Label10(0)"
      Item(0).Control(18)=   "Label18(0)"
      Item(0).Control(19)=   "txtAptoPostal"
      Item(0).Control(20)=   "dtpFecha"
      Item(1).Caption =   "Cónyuge"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "gbConyugue"
      Item(1).Control(1)=   "gbAlbacea"
      Item(2).Caption =   "Nombramientos"
      Item(2).ControlCount=   7
      Item(2).Control(0)=   "lswNombramiento"
      Item(2).Control(1)=   "dtpNombramiento"
      Item(2).Control(2)=   "txtAniosSerivicio"
      Item(2).Control(3)=   "Label18(1)"
      Item(2).Control(4)=   "Label18(3)"
      Item(2).Control(5)=   "cboEstadoLaboral"
      Item(2).Control(6)=   "btnNombramiento"
      Begin XtremeSuiteControls.ListView lswNombramiento 
         Height          =   2892
         Left            =   -69880
         TabIndex        =   36
         Top             =   1320
         Visible         =   0   'False
         Width           =   7932
         _Version        =   1572864
         _ExtentX        =   13991
         _ExtentY        =   5101
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
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbConyugue 
         Height          =   1692
         Left            =   -69880
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   7812
         _Version        =   1572864
         _ExtentX        =   13779
         _ExtentY        =   2984
         _StockProps     =   79
         Caption         =   "Datos del Cónyuge"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtConyugeNombre 
            Height          =   312
            Left            =   2640
            TabIndex        =   49
            Top             =   600
            Width           =   5172
            _Version        =   1572864
            _ExtentX        =   9123
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtConyugeCedula 
            Height          =   312
            Left            =   840
            TabIndex        =   51
            Top             =   600
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtConyugeTelTrabajo 
            Height          =   312
            Left            =   2640
            TabIndex        =   53
            Top             =   960
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtConyugeTelCelular 
            Height          =   312
            Left            =   2640
            TabIndex        =   54
            Top             =   1320
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtConyugeTelTrabajoExt 
            Height          =   312
            Left            =   5400
            TabIndex        =   55
            Top             =   960
            Width           =   1212
            _Version        =   1572864
            _ExtentX        =   2138
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label 
            Caption         =   "Trabajo"
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
            Index           =   17
            Left            =   1680
            TabIndex        =   48
            Top             =   996
            Width           =   1032
         End
         Begin VB.Label Label 
            Caption         =   "Extensión"
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
            Index           =   18
            Left            =   4440
            TabIndex        =   47
            Top             =   996
            Width           =   792
         End
         Begin VB.Label Label 
            Caption         =   "Móvil"
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
            Index           =   19
            Left            =   1680
            TabIndex        =   46
            Top             =   1320
            Width           =   1032
         End
         Begin VB.Label Label 
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
            Height          =   315
            Index           =   15
            Left            =   840
            TabIndex        =   45
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label Label 
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
            Height          =   312
            Index           =   16
            Left            =   2640
            TabIndex        =   44
            Top             =   360
            Width           =   1008
         End
      End
      Begin XtremeSuiteControls.ComboBox cboSexo 
         Height          =   312
         Left            =   2040
         TabIndex        =   15
         Top             =   480
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2355
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoCivil 
         Height          =   312
         Left            =   5880
         TabIndex        =   16
         Top             =   480
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3625
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   312
         Left            =   2040
         TabIndex        =   17
         Top             =   840
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
      Begin XtremeSuiteControls.ComboBox cboProvincia 
         Height          =   312
         Left            =   2040
         TabIndex        =   20
         Top             =   2280
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2990
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboCanton 
         Height          =   312
         Left            =   3840
         TabIndex        =   21
         Top             =   2280
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3413
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboDistrito 
         Height          =   312
         Left            =   5880
         TabIndex        =   22
         Top             =   2280
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3625
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboNacionalidad 
         Height          =   312
         Left            =   5880
         TabIndex        =   23
         Top             =   840
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3625
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   312
         Left            =   2040
         TabIndex        =   24
         Top             =   1200
         Width           =   5892
         _Version        =   1572864
         _ExtentX        =   10393
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail_02 
         Height          =   312
         Left            =   2040
         TabIndex        =   25
         Top             =   1560
         Width           =   5892
         _Version        =   1572864
         _ExtentX        =   10393
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAptoPostal 
         Height          =   312
         Left            =   2040
         TabIndex        =   26
         Top             =   1920
         Width           =   5892
         _Version        =   1572864
         _ExtentX        =   10393
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDireccion 
         Height          =   672
         Left            =   2040
         TabIndex        =   27
         Top             =   2640
         Width           =   5892
         _Version        =   1572864
         _ExtentX        =   10393
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotificaciones 
         Height          =   672
         Left            =   2040
         TabIndex        =   28
         Top             =   3360
         Width           =   5892
         _Version        =   1572864
         _ExtentX        =   10393
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpNombramiento 
         Height          =   312
         Left            =   -67960
         TabIndex        =   37
         Top             =   840
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.FlatEdit txtAniosSerivicio 
         Height          =   312
         Left            =   -66520
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   852
         _Version        =   1572864
         _ExtentX        =   1503
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
      Begin XtremeSuiteControls.GroupBox gbAlbacea 
         Height          =   1692
         Left            =   -69880
         TabIndex        =   41
         Top             =   2400
         Visible         =   0   'False
         Width           =   7812
         _Version        =   1572864
         _ExtentX        =   13779
         _ExtentY        =   2984
         _StockProps     =   79
         Caption         =   "Datos del Albacea"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtAlbaceaNombre 
            Height          =   312
            Left            =   2640
            TabIndex        =   50
            Top             =   720
            Width           =   5172
            _Version        =   1572864
            _ExtentX        =   9123
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAlbaceaCedula 
            Height          =   312
            Left            =   840
            TabIndex        =   52
            Top             =   720
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label 
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
            Height          =   312
            Index           =   1
            Left            =   2640
            TabIndex        =   43
            Top             =   480
            Width           =   1008
         End
         Begin VB.Label Label 
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
            Height          =   312
            Index           =   0
            Left            =   840
            TabIndex        =   42
            Top             =   480
            Width           =   1692
         End
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoLaboral 
         Height          =   312
         Left            =   -67960
         TabIndex        =   59
         Top             =   480
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4048
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnNombramiento 
         Height          =   492
         Left            =   -65200
         TabIndex        =   61
         Top             =   720
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Agregar"
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
         Picture         =   "frmCR_VerificaDatosPersonales.frx":030A
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   -69760
         TabIndex        =   60
         Top             =   480
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "A partir de:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   -69760
         TabIndex        =   38
         Top             =   840
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado Civil:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   4080
         TabIndex        =   35
         Top             =   480
         Width           =   1692
      End
      Begin VB.Label Label10 
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
         Height          =   252
         Index           =   0
         Left            =   600
         TabIndex        =   34
         Top             =   1200
         Width           =   1092
      End
      Begin VB.Label Label11 
         Caption         =   "Apto. Postal"
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
         Left            =   600
         TabIndex        =   33
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label10 
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
         Height          =   252
         Index           =   9
         Left            =   600
         TabIndex        =   32
         Top             =   1560
         Width           =   1092
      End
      Begin VB.Label Label7 
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
         Left            =   600
         TabIndex        =   31
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label 
         Caption         =   "Notificaciones"
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
         Index           =   25
         Left            =   600
         TabIndex        =   30
         Top             =   3360
         Width           =   1365
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Nacionalidad:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   4080
         TabIndex        =   29
         Top             =   840
         Width           =   1692
      End
      Begin VB.Label Label14 
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
         Height          =   252
         Left            =   600
         TabIndex        =   19
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
   End
   Begin XtremeSuiteControls.PushButton cmdTelefonos 
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   6480
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Telefonos"
      BackColor       =   16777215
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
      Picture         =   "frmCR_VerificaDatosPersonales.frx":0A2A
      ImageAlignment  =   4
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   9
      Top             =   7200
      Width           =   11964
      _ExtentX        =   21114
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   9596
            MinWidth        =   9596
            Object.ToolTipText     =   "Ultima Modificación"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   1482
            MinWidth        =   1482
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Bevel           =   0
            Object.Width           =   1482
            MinWidth        =   1482
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton cmdGrabar 
      Height          =   495
      Left            =   6600
      TabIndex        =   13
      Top             =   6480
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Guardar"
      BackColor       =   16777215
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
      Picture         =   "frmCR_VerificaDatosPersonales.frx":0B06
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBeneficiarios 
      Height          =   495
      Left            =   3840
      TabIndex        =   58
      Top             =   6480
      Width           =   2295
      _Version        =   1572864
      _ExtentX        =   4048
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Beneficiarios y Otros Contactos"
      BackColor       =   16777215
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
      Picture         =   "frmCR_VerificaDatosPersonales.frx":1237
      ImageAlignment  =   4
   End
   Begin VB.Label lblUnidadEtiqueta 
      BackStyle       =   0  'Transparent
      Caption         =   "Sección"
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
      Index           =   1
      Left            =   1320
      TabIndex        =   11
      Top             =   1560
      Width           =   1452
   End
   Begin VB.Label lblUnidadEtiqueta 
      BackStyle       =   0  'Transparent
      Caption         =   "Departamento"
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
      Index           =   0
      Left            =   1320
      TabIndex        =   10
      Top             =   1200
      Width           =   1452
   End
   Begin VB.Label lblInstCod 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   312
      Left            =   2880
      TabIndex        =   7
      Top             =   840
      Width           =   810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Institución"
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
      Index           =   3
      Left            =   1320
      TabIndex        =   6
      Top             =   840
      Width           =   1332
   End
   Begin VB.Label lblSeccion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   312
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label lblSeccionDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   312
      Left            =   3720
      TabIndex        =   4
      Top             =   1560
      Width           =   5292
   End
   Begin VB.Label lblDepartamento 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   312
      Left            =   2880
      TabIndex        =   3
      Top             =   1200
      Width           =   810
   End
   Begin VB.Label lblDepartamentoDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   312
      Left            =   3720
      TabIndex        =   2
      Top             =   1200
      Width           =   5292
   End
   Begin VB.Label lblCedula 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblInstDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   312
      Left            =   3720
      TabIndex        =   8
      Top             =   840
      Width           =   5292
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "[NOMBRE]"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   552
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11532
   End
   Begin VB.Image imgBanner 
      Height          =   2052
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12252
   End
End
Attribute VB_Name = "frmCR_VerificaDatosPersonales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset, strSQL As String
Public vPaso As Boolean, vFechaActual As Date

Private Sub sbConsulta()

On Error Resume Next

vFechaActual = fxFechaServidor

'Limpia Datos
lblNombre.Caption = ""
lblCedula.Caption = ""

txtEmail_02.Text = ""
txtEmail = ""
txtAptoPostal = ""
txtNotificaciones.Text = ""
txtDireccion.Text = ""

dtpFecha.Value = vFechaActual

tcMain.Item(0).Selected = True

strSQL = "exec spAFI_Persona_Consulta '" & GLOBALES.gCedulaActual & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    lblNombre.Caption = rs!Nombre
    lblCedula.Caption = GLOBALES.gCedulaActual
    
    Call sbCboAsignaDato(cboProvincia, rs!ProvinciaDesc & "")  'Se activa el Click ->    Call cboProvincia_Click
    Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")   'Se activa el Click
    Call sbCboAsignaDato(cboDistrito, rs!DistritoDesc & "")
        
    
   Call sbCboAsignaDato(cboEstadoCivil, rs!EstadoCivilDesc & "", True, rs!EstadoCivil)  'Se activa el Click ->    Call cboProvincia_Click
   If Not IsNull(rs!Nacionalidad) Then
       vPaso = True
       Call sbCboAsignaDato(cboNacionalidad, Trim(rs!Nacionalidad), True, rs!Cod_Nacionalidad)
       vPaso = False
   End If
     
    
    
    txtDireccion = Trim(rs!direccion & "")
    
    dtpFecha.Value = IIf(IsNull(rs!fecha_nac), Date, rs!fecha_nac)
    cboSexo.Text = IIf(IsNull(rs!sexo), "Masculino", IIf((rs!sexo = "F"), "Femenino", "Masculino"))
    
    txtAptoPostal.Text = Trim(rs!apto & "")
    txtEmail.Text = Trim(rs!AF_Email & "")
    txtEmail_02.Text = Trim(rs!Email_02 & "")
     
    
    StatusBarX.Panels(1).Text = "Ultima Actualización : "
    If IsNull(rs!ActualizaFecha) Then
      StatusBarX.Panels(1).Text = StatusBarX.Panels(1).Text & "No hay / Usuario : No hay"
    Else
      StatusBarX.Panels(1).Text = StatusBarX.Panels(1).Text & Format(rs!ActualizaFecha, "dd/mm/yyyy") & " / Usuario : " & rs!ActualizaUser
    End If
    
   Call sbCboAsignaDato(cboEstadoLaboral, rs!EstadoLaboralDesc, True, rs!estadoLaboral)
    
    If GLOBALES.SysASEVersion Then
        lblSeccion.Caption = rs!UT & ""
        lblSeccionDesc.Caption = rs!UT_Desc & ""
        
        lblDepartamento.Caption = rs!UP & ""
        lblDepartamentoDesc.Caption = rs!UP_Desc & ""
        
        
        
    Else
        lblSeccion.Caption = Trim(rs!cod_seccion & "")
        lblDepartamento.Caption = Trim(rs!cod_departamento & "")
    
        lblSeccionDesc.Caption = rs!SeccionDesc & ""
        lblDepartamentoDesc.Caption = rs!DepartamentoDesc & ""
    End If
    
    lblInstCod.Caption = Trim(rs!cod_institucion & "")
    lblInstDesc.Caption = rs!InstitucionDesc & ""
    
   dtpNombramiento.Value = IIf(IsNull(rs!nombramiento_fecha), rs!FechaIngreso, rs!nombramiento_fecha)
   
   txtAniosSerivicio.Text = Trim(rs!AnioServicio)
   
   txtConyugeCedula.Text = Trim(rs!conyuge_cedula & "")
   txtConyugeNombre.Text = Trim(rs!conyuge_nombre & "")
   txtConyugeTelCelular.Text = Trim(rs!conyuge_TelCell & "")
   txtConyugeTelTrabajo.Text = Trim(rs!conyuge_TelTra & "")
   txtConyugeTelTrabajoExt.Text = Trim(rs!conyuge_TelTraExt & "")
   
   txtAlbaceaCedula.Text = Trim(rs!albacea_Cedula & "")
   txtAlbaceaNombre.Text = Trim(rs!albacea_nombre & "")
   
   txtNotificaciones.Text = Trim(rs!Notificaciones & "")
    
   'Carga los datos para el CRM
   tcCRM.Item(0).Selected = True
   Call sbCRM_Consulta(0)
   
End If
rs.Close

End Sub


Private Sub btnBeneficiarios_Click()
 Call sbFormsCall("frmAF_Beneficiarios", 1, , , False, Me)

End Sub

Private Sub btnNombramiento_Click()
Dim strSQL As String

On Error GoTo vError

'Revisa Nombramiento / Variaciones para Registrarlas en el Histórico
'Nuevo Modelo: 2020/02/26 {PBN}
strSQL = "exec spAFI_Persona_Nombramientos_Add '" & Trim(GLOBALES.gCedulaActual) & "','" & cboEstadoLaboral.ItemData(cboEstadoLaboral.ListIndex) _
        & "','" & Format(dtpNombramiento.Value, "yyyy/mm/dd") & "','" & glogon.Usuario & "', 'A'"
Call ConectionExecute(strSQL)

MsgBox "Nombramiento registrado satisfactoriamente!", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboCanton_Click()
Dim strSQL As String

If vPaso Then Exit Sub

    strSQL = "select Distrito as Idx, rtrim(Descripcion) as ItmX from Distritos" _
            & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
            & "' and Canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
            & "' order by descripcion"
    Call sbCbo_Llena_New(cboDistrito, strSQL, False, True)

'Agrega Distrito En Limpio, ya que este dato es opcional
cboDistrito.AddItem " "
cboDistrito.Text = " "
End Sub

Private Sub cboCanton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDistrito.SetFocus
End Sub


Private Sub cboProvincia_Click()
Dim strSQL As String

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboCanton, strSQL, False, True)
vPaso = False

Call cboCanton_Click


End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub


Private Function fxValida()
Dim i As Integer
Dim vMensaje As String

vMensaje = ""
 

If Not fxEmail_Valida(txtEmail.Text) Then
    vMensaje = vMensaje & " - El Email principal no es válido!" & vbCrLf
End If

If Len(Trim(txtEmail_02.Text)) > 0 Then
    If Not fxEmail_Valida(txtEmail_02.Text) Then
        vMensaje = vMensaje & " - El Email secundario no es válido!" & vbCrLf
    End If
End If


If Trim(cboProvincia.Text) = "" Then vMensaje = vMensaje & " - No se especificó la Provincia" & vbCrLf
If Trim(cboCanton.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Cantón" & vbCrLf
If Trim(cboDistrito.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Distrito en la dirección" & vbCrLf

'If Trim(txtDireccion) = "" Then vMensaje = vMensaje & " - No se especificó la Dirección" & vbCrLf
If Not fxDireccion_Valida(Trim(txtDireccion), "-,#,*") Then vMensaje = vMensaje & " - No se especificó la Dirección" & vbCrLf

If Len(vMensaje) = 0 Then
  fxValida = True
Else
  fxValida = False
  MsgBox vMensaje, vbExclamation

End If

End Function

Private Sub cmdGrabar_Click()
Dim strSQL As String, bol As Boolean

On Error GoTo vError

If Not fxValida Then
  Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "update socios set provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "', Canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
       & "',distrito = '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "',estadocivil='" & cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex) _
       & "',direccion='" & txtDireccion.Text & "',fecha_nac='" & Format(dtpFecha, "yyyy/mm/dd") _
       & "',apto = '" & Trim(txtAptoPostal) & "',af_email = '" & Trim(txtEmail) & "',sexo = '" & IIf((cboSexo.Text = "Femenino"), "F", "M") _
       & "',EstadoLaboral = '" & cboEstadoLaboral.ItemData(cboEstadoLaboral.ListIndex) & "',ActualizaFecha = dbo.MyGetdate(), ActualizaUser = '" & glogon.Usuario _
       & "',Nombramiento_Fecha = '" & Format(dtpNombramiento.Value, "yyyy/mm/dd") & "',Conyuge_Cedula = '" & txtConyugeCedula.Text _
       & "', Conyuge_Nombre = '" & txtConyugeNombre.Text & "',Conyuge_TelCell = '" & txtConyugeTelCelular.Text _
       & "',Conyuge_TelTra = '" & txtConyugeTelTrabajo.Text & "',Conyuge_TelTraExt = '" & txtConyugeTelTrabajoExt.Text _
       & "',Notificaciones = '" & txtNotificaciones.Text & "',Albacea_cedula = '" & txtAlbaceaCedula.Text & "',Albacea_nombre = '" _
       & txtAlbaceaNombre.Text & "', Email_02 = '" & Trim(txtEmail_02.Text) & "'" _
       & " Where Cedula = '" & GLOBALES.gCedulaActual & "'"

'Revisa Nombramiento / Variaciones para Registrarlas en el Histórico
'Nuevo Modelo: 2020/02/26 {PBN}
strSQL = strSQL & Space(10) & "exec spAFI_Persona_Nombramientos_Add '" & Trim(GLOBALES.gCedulaActual) & "','" & cboEstadoLaboral.ItemData(cboEstadoLaboral.ListIndex) _
        & "','" & Format(dtpNombramiento.Value, "yyyy/mm/dd") & "','" & glogon.Usuario & "', 'A'"


'Direcciones Adicionales
'Nuevo Modelo: 2020/02/26 {PBN}
strSQL = strSQL & Space(10) & "exec spAFI_Persona_Direccion_Add '" & Trim(GLOBALES.gCedulaActual) & "','" & cboProvincia.ItemData(cboProvincia.ListIndex) _
       & "','" & cboCanton.ItemData(cboCanton.ListIndex) & "','" & cboDistrito.ItemData(cboDistrito.ListIndex) _
       & "','" & Trim(txtDireccion.Text) & "','" & Trim(txtEmail.Text) & "','" & Trim(txtEmail_02.Text) _
       & "','','','" & glogon.Usuario & "', 'A'"

'Aplica
Call ConectionExecute(strSQL)


Call Bitacora("Modifica", "Informacion de la Persona con cedula=" & Trim(GLOBALES.gCedulaActual))

Me.MousePointer = vbDefault

MsgBox "Información Actualizada Satisfactoriamente...", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cmdTelefonos_Click()
 Call sbFormsCall("frmAF_Telefonos", 1, , , False, Me)
End Sub

Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAptoPostal.SetFocus
End Sub

Private Sub sbCRM_Consulta(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

lswCRM.ColumnHeaders.Clear
lswCRM.ListItems.Clear
lswCRM.Checkboxes = False


Select Case Index
  Case 0 'Bienes
  
    lswCRM.ListItems.Clear
    lswCRM.ColumnHeaders.Clear
    lswCRM.ColumnHeaders.Add 1, , "Bienes", 3400
'    lswCRM.ColumnHeaders.Add 2, , "Usuario", 2100
'    lswCRM.ColumnHeaders.Add 3, , "Fecha", 2500
    lswCRM.Checkboxes = True
    
    strSQL = "exec spAFI_Persona_Bienes_Consulta '" & GLOBALES.gCedulaActual & "',1"
    Call OpenRecordSet(rs, strSQL)
    
    vPaso = True
    With lswCRM.ListItems
       .Clear
       Do While Not rs.EOF
        Set itmX = .Add(, , rs!Descripcion)
'            itmX.SubItems(1) = rs!registro_usuario & ""
'            itmX.SubItems(2) = rs!registro_fecha & ""
            itmX.Tag = rs!Bien_Tipo
            
            itmX.Checked = IIf((rs!asignado = 1), True, False)
            
        rs.MoveNext
       Loop
       rs.Close
    End With
    vPaso = False
    
    
  Case 1 'Canal de Comunicacion
  
    lswCRM.ListItems.Clear
    lswCRM.ColumnHeaders.Clear
    lswCRM.ColumnHeaders.Add 1, , "Tipo de Canal", 3400
'    lswCRM.ColumnHeaders.Add 2, , "Usuario", 2100
'    lswCRM.ColumnHeaders.Add 3, , "Fecha", 2500
    lswCRM.Checkboxes = True
    
    strSQL = "exec spAFI_Persona_Canales_Consulta '" & GLOBALES.gCedulaActual & "',1"
    Call OpenRecordSet(rs, strSQL)
    
    vPaso = True
    With lswCRM.ListItems
       .Clear
       Do While Not rs.EOF
        Set itmX = .Add(, , rs!Descripcion)
'            itmX.SubItems(1) = rs!registro_usuario & ""
'            itmX.SubItems(2) = rs!registro_fecha & ""
            itmX.Tag = rs!Canal_Tipo
            
            itmX.Checked = IIf((rs!asignado = 1), True, False)
            
        rs.MoveNext
       Loop
       rs.Close
    End With
    vPaso = False

  
  Case 2 'Gustos y Preferencias
  
    lswCRM.ListItems.Clear
    lswCRM.ColumnHeaders.Clear
    lswCRM.ColumnHeaders.Add 1, , "Gustos y Preferencias", 3400
'    lswCRM.ColumnHeaders.Add 2, , "Usuario", 2100
'    lswCRM.ColumnHeaders.Add 3, , "Fecha", 2500
    lswCRM.Checkboxes = True
  
    strSQL = "exec spAFI_Persona_Preferencias_Consulta '" & GLOBALES.gCedulaActual & "',1"
    Call OpenRecordSet(rs, strSQL)
    
    vPaso = True
    With lswCRM.ListItems
       .Clear
       Do While Not rs.EOF
        Set itmX = .Add(, , rs!Descripcion)
'            itmX.SubItems(1) = rs!registro_usuario & ""
'            itmX.SubItems(2) = rs!registro_fecha & ""
            itmX.Tag = rs!cod_preferencia
            
            itmX.Checked = IIf((rs!asignado = 1), True, False)
            
        rs.MoveNext
       Loop
       rs.Close
    End With
    vPaso = False
  
  
  
  Case 3 'Escolaridad
  
    lswCRM.ListItems.Clear
    lswCRM.ColumnHeaders.Clear
    lswCRM.ColumnHeaders.Add 1, , "Nivel de Escolaridad", 3400
'    lswCRM.ColumnHeaders.Add 2, , "Usuario", 2100
'    lswCRM.ColumnHeaders.Add 3, , "Fecha", 2500
    lswCRM.Checkboxes = True
    
    strSQL = "exec spAFI_Persona_Escolaridad_Consulta '" & GLOBALES.gCedulaActual & "',1"
    Call OpenRecordSet(rs, strSQL)
    
    vPaso = True
    With lswCRM.ListItems
       .Clear
       Do While Not rs.EOF
        Set itmX = .Add(, , rs!Descripcion)
'            itmX.SubItems(1) = rs!registro_usuario & ""
'            itmX.SubItems(2) = rs!registro_fecha & ""
            itmX.Tag = rs!Escolaridad_Tipo
            
            itmX.Checked = IIf((rs!asignado = 1), True, False)
            
        rs.MoveNext
       Loop
       rs.Close
    End With
    vPaso = False
  
    
End Select

lswCRM.HideColumnHeaders = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
'CAMBIA AL MODULO DE AFILIACION
'PARA REGISTRAR EN BITACORAS, EL MOVIMIENTO EN EL MODULO CORRECTO
vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

If GLOBALES.SysASEVersion Then
  lblUnidadEtiqueta(0).Caption = "U. Programatica"
  lblUnidadEtiqueta(1).Caption = "U. Trabajo"
Else
  lblUnidadEtiqueta(0).Caption = "Departamento"
  lblUnidadEtiqueta(1).Caption = "Sección"
End If


vPaso = True
'Call sbCargaCbo(cboProvincia, "provincias")
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)
vPaso = False

cboSexo.Clear
cboSexo.AddItem "Masculino"
cboSexo.AddItem "Femenino"
cboSexo.Text = "Masculino"

strSQL = "select cod_nacionalidad as 'IdX', Descripcion as 'ItmX' from sys_nacionalidades" _
       & " where Activo = 1" _
       & " order by Omision desc, Descripcion asc"
Call sbCbo_Llena_New(cboNacionalidad, strSQL, False, True)

strSQL = "select Estado_Civil as 'IdX', Descripcion as 'ItmX' from SYS_ESTADO_CIVIL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboEstadoCivil, strSQL, False, True)

strSQL = "select Estado_Laboral as 'IdX', Descripcion as 'ItmX' from AFI_ESTADO_LABORAL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboEstadoLaboral, strSQL, False, True)

If GLOBALES.gCedulaActual <> "" Then
 Call sbConsulta
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Establece el Modulo de Credito para el sistema de seguridad y Bitacoras
vModulo = 3
End Sub




Private Sub sbNombramientos_Consulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

lswNombramiento.ListItems.Clear
lswNombramiento.ColumnHeaders.Clear
lswNombramiento.ColumnHeaders.Add 1, , "Estado", 1200
lswNombramiento.ColumnHeaders.Add 2, , "A Partir", 1500
lswNombramiento.ColumnHeaders.Add 3, , "Fecha", 2500
lswNombramiento.ColumnHeaders.Add 4, , "Usuario", 2500
        

strSQL = "exec spAFI_Persona_Nombramientos_Consulta '" & GLOBALES.gCedulaActual & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswNombramiento.ListItems.Add(, , rs!EstadoLaboralDesc)
      itmX.SubItems(1) = Format(rs!fecha, "dd/mm/yyyy")
      itmX.SubItems(2) = rs!Registro_Fecha
      itmX.SubItems(3) = rs!Registro_Usuario
  rs.MoveNext
Loop
rs.Close
  
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswCRM_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

Select Case tcCRM.Selected.Index
  Case 0 'Bienes"
    strSQL = "exec spAFI_Persona_Bienes_Registra '" & GLOBALES.gCedulaActual & "','" & Item.Tag & "','" _
           & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
  
  Case 1 'Canales
    strSQL = "exec spAFI_Persona_Canales_Registra '" & GLOBALES.gCedulaActual & "','" & Item.Tag & "','" _
           & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
  
  Case 2 'Preferencias"
    strSQL = "exec spAFI_Persona_Preferencias_Registra '" & GLOBALES.gCedulaActual & "','" & Item.Tag & "','" _
           & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
  
  
  Case 3 'Escolaridad"
    strSQL = "exec spAFI_Persona_Escolaridad_Registra '" & GLOBALES.gCedulaActual & "','" & Item.Tag & "','" _
           & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
  
  Case Else
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub tcCRM_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Call sbCRM_Consulta(Item.Index)
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 2 Then
    Call sbNombramientos_Consulta
End If
End Sub

Private Sub txtAlbaceaCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAlbaceaNombre.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtAlbaceaCedula = gBusquedas.Resultado
  txtAlbaceaNombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtAlbaceaNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeCedula.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtAlbaceaCedula = gBusquedas.Resultado
  txtAlbaceaNombre = gBusquedas.Resultado2
End If
End Sub

Private Sub txtAptoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub


Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstadoCivil.SetFocus
End Sub

Private Sub txtDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub txtConyugeCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeNombre.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtConyugeCedula = gBusquedas.Resultado
  txtConyugeNombre = gBusquedas.Resultado2
End If

End Sub



Private Sub txtConyugeNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeTelTrabajo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtConyugeCedula = gBusquedas.Resultado
  txtConyugeNombre = gBusquedas.Resultado2
End If

End Sub


Private Sub txtConyugeTelTrabajo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeTelTrabajoExt.SetFocus
End Sub

Private Sub txtConyugeTelTrabajoExt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeTelCelular.SetFocus
End Sub

