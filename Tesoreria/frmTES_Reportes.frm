VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmTES_Reportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   HelpContextID   =   1011
   Icon            =   "frmTES_Reportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   8895
   WhatsThisHelp   =   -1  'True
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7095
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   8655
      _Version        =   1441793
      _ExtentX        =   15266
      _ExtentY        =   12515
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
      ItemCount       =   2
      Item(0).Caption =   "Informes"
      Item(0).ControlCount=   33
      Item(0).Control(0)=   "cmdImprime"
      Item(0).Control(1)=   "cbo"
      Item(0).Control(2)=   "cboTipo"
      Item(0).Control(3)=   "cboEstado"
      Item(0).Control(4)=   "cboFecha"
      Item(0).Control(5)=   "cboBusqueda"
      Item(0).Control(6)=   "cboUsuario"
      Item(0).Control(7)=   "cboConcepto"
      Item(0).Control(8)=   "cboUnidad"
      Item(0).Control(9)=   "dtpInicio"
      Item(0).Control(10)=   "dtpCorte"
      Item(0).Control(11)=   "txtToken"
      Item(0).Control(12)=   "txtUsuario"
      Item(0).Control(13)=   "txtDesde"
      Item(0).Control(14)=   "txtHasta"
      Item(0).Control(15)=   "Label2(9)"
      Item(0).Control(16)=   "label1(0)"
      Item(0).Control(17)=   "Label2(0)"
      Item(0).Control(18)=   "label1(1)"
      Item(0).Control(19)=   "label1(5)"
      Item(0).Control(20)=   "Label2(1)"
      Item(0).Control(21)=   "Label3(1)"
      Item(0).Control(22)=   "Label3(0)"
      Item(0).Control(23)=   "Label2(2)"
      Item(0).Control(24)=   "Label2(3)"
      Item(0).Control(25)=   "Label2(4)"
      Item(0).Control(26)=   "chkDocumentos"
      Item(0).Control(27)=   "chkConcepto"
      Item(0).Control(28)=   "chkUnidad"
      Item(0).Control(29)=   "chkTokenAgrupado"
      Item(0).Control(30)=   "chkIncluirDetalle"
      Item(0).Control(31)=   "chkRef"
      Item(0).Control(32)=   "chkModoProtegido"
      Item(1).Caption =   "Cubos"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "fraCargaDatos"
      Begin XtremeSuiteControls.PushButton cmdImprime 
         Height          =   795
         Left            =   6840
         TabIndex        =   2
         Top             =   6120
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   1402
         _StockProps     =   79
         Caption         =   "&Reporte"
         BackColor       =   -2147483633
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
         Picture         =   "frmTES_Reportes.frx":6852
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   5172
         _Version        =   1441793
         _ExtentX        =   9128
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
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   1920
         TabIndex        =   4
         Top             =   840
         Width           =   5172
         _Version        =   1441793
         _ExtentX        =   9128
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
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   1920
         TabIndex        =   5
         Top             =   1560
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.ComboBox cboFecha 
         Height          =   312
         Left            =   1920
         TabIndex        =   6
         Top             =   1920
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.ComboBox cboBusqueda 
         Height          =   312
         Left            =   1920
         TabIndex        =   7
         Top             =   2280
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.ComboBox cboUsuario 
         Height          =   312
         Left            =   1920
         TabIndex        =   8
         Top             =   4440
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.ComboBox cboConcepto 
         Height          =   312
         Left            =   1920
         TabIndex        =   9
         Top             =   4800
         Width           =   4572
         _Version        =   1441793
         _ExtentX        =   8070
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
      Begin XtremeSuiteControls.ComboBox cboUnidad 
         Height          =   312
         Left            =   1920
         TabIndex        =   10
         Top             =   5160
         Width           =   4572
         _Version        =   1441793
         _ExtentX        =   8070
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   4080
         TabIndex        =   11
         Top             =   1920
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   5520
         TabIndex        =   12
         Top             =   1920
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
      Begin XtremeSuiteControls.FlatEdit txtToken 
         Height          =   312
         Left            =   1920
         TabIndex        =   13
         Top             =   3480
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   312
         Left            =   4080
         TabIndex        =   14
         Top             =   4440
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDesde 
         Height          =   312
         Left            =   1920
         TabIndex        =   15
         Top             =   2640
         Width           =   5052
         _Version        =   1441793
         _ExtentX        =   8911
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtHasta 
         Height          =   312
         Left            =   1920
         TabIndex        =   16
         Top             =   3000
         Width           =   5052
         _Version        =   1441793
         _ExtentX        =   8911
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkDocumentos 
         Height          =   255
         Left            =   7200
         TabIndex        =   28
         Top             =   840
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Todos"
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
      End
      Begin XtremeSuiteControls.CheckBox chkConcepto 
         Height          =   372
         Left            =   6720
         TabIndex        =   29
         Top             =   4800
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Todos"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkUnidad 
         Height          =   252
         Left            =   6720
         TabIndex        =   30
         Top             =   5160
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkTokenAgrupado 
         Height          =   372
         Left            =   1920
         TabIndex        =   31
         Top             =   3840
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "&Agrupado por Token"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkIncluirDetalle 
         Height          =   372
         Left            =   3600
         TabIndex        =   32
         Top             =   5880
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Incluir Detalle del Documento"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkRef 
         Height          =   252
         Left            =   3600
         TabIndex        =   33
         Top             =   6240
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Incluir Referencias"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.GroupBox fraCargaDatos 
         Height          =   4332
         Left            =   -69760
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   7932
         _Version        =   1441793
         _ExtentX        =   13991
         _ExtentY        =   7641
         _StockProps     =   79
         Caption         =   "Cargar Datos para Analisis"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.RadioButton OptX 
            Height          =   252
            Index           =   0
            Left            =   1320
            TabIndex        =   35
            Top             =   1680
            Width           =   2892
            _Version        =   1441793
            _ExtentX        =   5101
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Base: Transaccional"
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
            Appearance      =   16
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaInicioCubo 
            Height          =   312
            Left            =   5760
            TabIndex        =   36
            Top             =   1680
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
         Begin XtremeSuiteControls.DateTimePicker dtpFechaCorteCubo 
            Height          =   312
            Left            =   5760
            TabIndex        =   37
            Top             =   2040
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
         Begin XtremeSuiteControls.RadioButton OptX 
            Height          =   252
            Index           =   1
            Left            =   1320
            TabIndex        =   38
            Top             =   2040
            Width           =   2892
            _Version        =   1441793
            _ExtentX        =   5101
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Base: Contable"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton btnCubo_Genera 
            Height          =   675
            Left            =   5760
            TabIndex        =   39
            Top             =   3600
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   1191
            _StockProps     =   79
            Caption         =   "&Generar"
            BackColor       =   -2147483633
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
            Picture         =   "frmTES_Reportes.frx":700E
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "(Corte)"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   8
            Left            =   7200
            TabIndex        =   44
            Top             =   2040
            Width           =   972
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "(Inicio)"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   7
            Left            =   7200
            TabIndex        =   43
            Top             =   1680
            Width           =   972
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Emisión"
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
            Index           =   6
            Left            =   4200
            TabIndex        =   42
            Top             =   1680
            Width           =   1692
         End
         Begin VB.Label Label2 
            Caption         =   "Proceso para cargar información para Analisis de Tesorería, por rango de fechas.  (Cubos)"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Index           =   5
            Left            =   1320
            TabIndex        =   41
            Top             =   960
            Width           =   6612
         End
         Begin VB.Image Image3 
            Height          =   630
            Left            =   360
            Picture         =   "frmTES_Reportes.frx":7813
            Top             =   720
            Width           =   585
         End
         Begin VB.Label lblStatus 
            Caption         =   "Este proceso puede tardar varios minutos, espere el mensaje de proceso concluido."
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
            Height          =   972
            Left            =   360
            TabIndex        =   40
            Top             =   3720
            Width           =   5172
         End
      End
      Begin XtremeSuiteControls.CheckBox chkModoProtegido 
         Height          =   375
         Left            =   3600
         TabIndex        =   45
         Top             =   6480
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Visualizar TF en Modo Protegido"
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
         Alignment       =   1
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad"
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
         Index           =   4
         Left            =   240
         TabIndex        =   27
         Top             =   5196
         Width           =   1692
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   4836
         Width           =   1692
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario Base"
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
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   4476
         Width           =   1692
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   0
         Left            =   840
         TabIndex        =   24
         Top             =   2676
         Width           =   972
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   1
         Left            =   840
         TabIndex        =   23
         Top             =   3036
         Width           =   972
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Base"
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
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   1956
         Width           =   1692
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta (Bancos)"
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
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   516
         Width           =   1812
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
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
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   2316
         Width           =   1692
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   1596
         Width           =   1812
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Documento"
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
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   876
         Width           =   1812
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Token"
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
         Index           =   9
         Left            =   240
         TabIndex        =   17
         Top             =   3516
         Width           =   1692
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Bancos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   10
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   3972
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmTES_Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, gDocumento As String

Private Sub btnCubo_Genera_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case True
  Case OptX.Item(0).Value  'Base Transaccional
        strSQL = "exec spTesAnalisisCubo '" & Format(dtpFechaInicioCubo.Value, "yyyy/mm/dd") & "','" & Format(dtpFechaCorteCubo.Value, "yyyy/mm/dd") & "'"
        vMensaje = "Tesoreria"
  Case OptX.Item(1).Value 'Base Contable
        strSQL = "exec spTesAnalisisContableCubo '" & Format(dtpFechaInicioCubo.Value, "yyyy/mm/dd") & "','" & Format(dtpFechaCorteCubo.Value, "yyyy/mm/dd") & "'"
        vMensaje = "TesoreriaConta"
End Select
Call ConectionExecute(strSQL)

lblStatus.Caption = "Proceso Concluido con éxito, la información puede ser utilizada desde la base de datos de análisis, cubo: " & vMensaje

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub cbo_Click()
If vPaso Then Exit Sub

On Error GoTo vError

If cbo.ListCount = 0 Then
   cbo.AddItem " "
   cbo.ItemData(cbo.NewIndex) = 0
   cbo.Text = " "
End If

vPaso = True
Call sbTesTiposDocsCargaCbo(cboTipo, cbo.ItemData(cbo.ListIndex))
vPaso = False

If gDocumento <> "" Then
  cboTipo.Text = gDocumento
End If

vError:

End Sub

Private Sub cboBusqueda_Click()
If Trim(cboBusqueda) = "----------------" Then
   cboBusqueda.Text = "Todos"
End If

If Mid(cboBusqueda.Text, 1, 1) = "T" Then
  txtDesde.Enabled = False
Else
  txtDesde.Enabled = True
End If

txtHasta.Enabled = txtDesde.Enabled

End Sub

Private Sub cboEstado_Click()
If Trim(cboEstado) = "----------------" Then
   cboEstado = "Solicitados"
End If

If Trim(cboEstado) = "Todos" Then
   cboEstado = "Solicitados"
End If
End Sub

Private Sub cboFecha_Click()

If Mid(cboFecha.Text, 1, 1) = "T" Then
  dtpInicio.Enabled = False
Else
  dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub


Private Sub cboTipo_Click()
On Error GoTo vError

If vPaso Then Exit Sub
gDocumento = cboTipo.Text

vError:
End Sub

Private Sub cboUsuario_Click()
If Mid(cboUsuario.Text, 1, 1) = "T" Then
  txtUsuario.Enabled = False
Else
  txtUsuario.Enabled = True
End If
End Sub

Private Sub chkConcepto_Click()
If chkConcepto.Value = vbChecked Then
   cboConcepto.Enabled = False
Else
   cboConcepto.Enabled = True
End If
End Sub

Private Sub chkDocumentos_Click()

If chkDocumentos.Value = vbChecked Then
  cboTipo.Enabled = False
Else
  cboTipo.Enabled = True
End If

End Sub


Private Sub chkUnidad_Click()

If chkUnidad.Value = vbChecked Then
   cboUnidad.Enabled = False
Else
   cboUnidad.Enabled = True
End If

End Sub

Private Sub cmdImprime_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite los reportes de tesoreria en base a los diferentes parametros
'               suministrados por el usuario.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, strBusqueda As String
Dim strEstado As String, blnDesde As Boolean, blnHasta As Boolean

On Error GoTo vError

Me.MousePointer = vbHourglass

blnDesde = IIf(txtDesde = "", False, True)
blnHasta = IIf(txtHasta = "", False, True)

If blnDesde = False And blnHasta = False Then
   'No Corrobora nada
Else
   If blnDesde = True And blnHasta = True Then
     
    
     Select Case Trim(cboBusqueda)
        Case "Por Número de Caso"
              strBusqueda = "{TES_TRANSACCIONES.NSOLICITUD} in " & Trim(txtDesde)
              strBusqueda = strBusqueda & " to " & Trim(txtHasta)
        Case "Por Beneficiario"
              strBusqueda = "{TES_TRANSACCIONES.BENEFICIARIO} in '" & Trim(txtDesde)
              strBusqueda = strBusqueda & "' to '" & Trim(txtHasta) & "'"
        Case "Por Número de Documento"
              strBusqueda = "{TES_TRANSACCIONES.NDOCUMENTO} in '" & Trim(txtDesde)
              strBusqueda = strBusqueda & "' to '" & Trim(txtHasta) & "'"
        Case "Por Número de Referencia"
              strBusqueda = "{TES_TRANSACCIONES.OP} in " & Trim(txtDesde)
              strBusqueda = strBusqueda & " to " & Trim(txtHasta)
        End Select
   Else
     MsgBox "Falta algún Parametro de Busqueda", vbExclamation, "Faltan Datos"
     Me.MousePointer = vbDefault
     Exit Sub
   End If

End If

strSQL = ""

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Tesorería"
    
    .Formulas(0) = "Fecha='" & Format(fxFechaServidor, ("dd/mm/yyyy")) & "'"
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    
    If chkDocumentos.Value = xtpChecked Then
    .Formulas(2) = "TipoEstado='TODOS'"
    Else
    .Formulas(2) = "TipoEstado='" & fxTesTiposDocDescribe(cboTipo.ItemData(cboTipo.ListIndex)) & "'"
    End If
    .Formulas(3) = "Modulo='" & UCase(cboEstado.Text) & "'"
    .Formulas(4) = "Banco='" & Trim(cbo.Text) & "'"
    .Formulas(8) = "Token='" & Trim(txtToken.Text) & "'"
    .Formulas(5) = "fxImprimeDetalle=" & chkIncluirDetalle.Value
    .Formulas(6) = "fxImprimeRef=" & chkRef.Value
     
    
    .Connect = glogon.ConectRPT
     
    Select Case Trim(cboEstado.Text)
      Case "Solicitados", "Emitidos"
           If chkTokenAgrupado.Value = vbChecked Then
               .ReportFileName = SIFGlobal.fxPathReportes("Banking_ListadoGeneralToken.rpt")
           Else
               .ReportFileName = SIFGlobal.fxPathReportes("Banking_ListadoGeneral.rpt")
           End If
           
           Select Case Mid(cboEstado.Text, 1, 1)
                Case "S"
                    strEstado = "{TES_TRANSACCIONES.ESTADO} = 'P'"
                Case "E"
                    strEstado = "({TES_TRANSACCIONES.ESTADO} = 'I' or {TES_TRANSACCIONES.ESTADO} = 'T' or {TES_TRANSACCIONES.ESTADO} = 'E')"
            End Select
     
      Case "Anulados"
           .ReportFileName = SIFGlobal.fxPathReportes("Banking_ListadoGeneralAnulados.rpt")
           strEstado = "{TES_TRANSACCIONES.ESTADO} =  'A'"
         
    End Select
    
    
    
    If Trim(cboEstado.Text) <> "General" Then
       If Len(strSQL) > 0 Then strSQL = strSQL & " AND " & strEstado
'       strSQL = strSQL & "{TES_TRANSACCIONES.ESTADO} " & strEstado
    End If
    
    
    If chkDocumentos.Value = vbUnchecked Then
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
       strSQL = strSQL & "{TES_TRANSACCIONES.TIPO} ='" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
    End If
        
    Select Case Mid(cboFecha.Text, 1, 1)
      Case "S" 'Solicitud
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "({TES_TRANSACCIONES.FECHA_SOLICITUD} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
               & ") To Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & "))"
      Case "E" 'Emision
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "({TES_TRANSACCIONES.FECHA_EMISION} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
               & ") To Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & "))"
      Case "A" 'Anulacion
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "({TES_TRANSACCIONES.FECHA_ANULA} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
               & ") To Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & "))"
      Case "T" 'Todas
       .Formulas(7) = "Rango='TODAS LAS FECHAS'"
      
    End Select
        
    If Mid(cboFecha.Text, 1, 1) <> "T" Then
       .Formulas(7) = "Rango='Reporte del  " & Format(dtpInicio.Value, "dd/mm/yyyy") & "  al  " & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
    End If
    
    
    Select Case Mid(cboUsuario.Text, 1, 1)
      Case "S" 'Solicitud
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{TES_TRANSACCIONES.USER_SOLICITA} = '" & txtUsuario & "'"
      Case "E" 'Emision
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{TES_TRANSACCIONES.USER_GENERA} = '" & txtUsuario & "'"
      Case "A" 'Anulacion
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{TES_TRANSACCIONES.USER_ANULA} = '" & txtUsuario & "'"
      Case "T" 'Todas
    End Select
    
    
   If chkUnidad.Value = vbUnchecked Then
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{TES_TRANSACCIONES.COD_UNIDAD} = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'"
   End If
   
   If chkConcepto.Value = vbUnchecked Then
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{TES_TRANSACCIONES.COD_CONCEPTO} = '" & cboConcepto.ItemData(cboConcepto.ListIndex) & "'"
   End If
    
   If blnDesde = True And blnHasta = True Then
     If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & strBusqueda
   End If
    
   If Trim(txtToken.Text) <> "" Then
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{TES_TRANSACCIONES.ID_TOKEN} = '" & txtToken.Text & "'"
   End If
    
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
    strSQL = strSQL & "{TES_TRANSACCIONES.ID_BANCO}=" & cbo.ItemData(cbo.ListIndex)
    
    strSQL = strSQL & "And ISNULL({TES_TRANSACCIONES.FECHA_HOLD})"
    
    
    If chkModoProtegido.Value = xtpUnchecked Then
                strSQL = strSQL & " AND ISNULL({TES_TRANSACCIONES.MODO_PROTEGIDO})"
    End If
    
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
 vModulo = 9
 
Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

gDocumento = ""
tcMain.Item(0).Selected = True

vPaso = True
 Call sbTesBancoCargaCboAccesoGeneral(cbo)
vPaso = False

Call cbo_Click
 
 cboFecha.Clear
 cboFecha.AddItem "Solicitud"
 cboFecha.AddItem "Emisión"
 cboFecha.AddItem "Anulación"
 cboFecha.AddItem "Todas"
 cboFecha.Text = "Solicitud"
 
 
 cboUsuario.Clear
 cboUsuario.AddItem "Solicita"
 cboUsuario.AddItem "Emite"
 cboUsuario.AddItem "Anula"
 cboUsuario.AddItem "Todos"
 cboUsuario.Text = "Todos"
  
 Call cboUsuario_Click
  
 Call sbTESCombos(cboEstado, "estado")
 Call sbTESCombos(cboBusqueda, "busqueda")
 cboBusqueda.Text = "Todos"
 
 Call sbTesUnidadesCargaCboGeneral(cboUnidad)
 Call sbTesConceptosCargaCboGeneral(cboConcepto)
 
 dtpInicio.Value = fxFechaServidor
 dtpCorte.Value = dtpInicio.Value
  
 dtpFechaInicioCubo.Value = dtpInicio.Value
 dtpFechaCorteCubo.Value = dtpInicio.Value
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Dim strSQL As String, rs As New ADODB.Recordset
 
 strSQL = "select count(*) as 'Existe' from TES_AUTORIZACIONES" _
        & " Where ESTADO = 'A' and NOMBRE = '" & glogon.Usuario & "'"
 Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 1 Then
    chkModoProtegido.Enabled = True
 Else
    chkModoProtegido.Enabled = False
 End If
 
End Sub



Private Sub txtDesde_KeyPress(KeyAscii As Integer)
On Error GoTo vError

If Trim(cboBusqueda) <> "Por Beneficiario" Then
   KeyAscii = Validacion(KeyAscii)
End If

If KeyAscii = vbKeyReturn Then
    txtHasta.SetFocus
End If

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub txtHasta_KeyPress(KeyAscii As Integer)
If Trim(cboBusqueda) <> "Por Beneficiario" Then
   KeyAscii = Validacion(KeyAscii)
End If

If KeyAscii = vbKeyReturn Then
  If dtpInicio.Enabled = True Then
     dtpInicio.SetFocus
  Else
    cmdImprime.SetFocus
  End If
End If
End Sub

Private Sub txtToken_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = ""
  gBusquedas.Columna = "id_token"
  gBusquedas.Orden = "id_token"
  gBusquedas.Consulta = "select id_token,registro_fecha,estado from Tes_tokens"
  gBusquedas.Filtro = ""
  gBusquedas.Orden = "registro_fecha desc"
  frmBusquedas.Show vbModal
  txtToken.Text = gBusquedas.Resultado
End If
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""

If KeyCode = vbKeyF4 Then
 gBusquedas.Columna = "nombre"
 gBusquedas.Orden = "nombre"
 gBusquedas.Consulta = "select Nombre as Usuario,Descripcion from usuarios"
 gBusquedas.Filtro = ""
 frmBusquedas.Show vbModal
 txtUsuario = gBusquedas.Resultado
End If


End Sub
