VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmAF_CRRenuncia 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planteamiento de Renuncias"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10320
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   10320
   Begin XtremeSuiteControls.TabControl tcRenuncia 
      Height          =   7935
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   10095
      _Version        =   1572864
      _ExtentX        =   17806
      _ExtentY        =   13996
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
      Item(0).Caption =   "Renuncia"
      Item(0).ControlCount=   4
      Item(0).Control(0)=   "tcMain"
      Item(0).Control(1)=   "btnSiguiente"
      Item(0).Control(2)=   "btnAnterior"
      Item(0).Control(3)=   "cmdNuevo"
      Item(1).Caption =   "Control"
      Item(1).ControlCount=   16
      Item(1).Control(0)=   "FlatEdit1"
      Item(1).Control(1)=   "txtRegUser"
      Item(1).Control(2)=   "FlatEdit3"
      Item(1).Control(3)=   "txtResUser"
      Item(1).Control(4)=   "FlatEdit5"
      Item(1).Control(5)=   "txtEstadoControl"
      Item(1).Control(6)=   "FlatEdit7"
      Item(1).Control(7)=   "txtRegFecha"
      Item(1).Control(8)=   "FlatEdit9"
      Item(1).Control(9)=   "txtResFecha"
      Item(1).Control(10)=   "FlatEdit11"
      Item(1).Control(11)=   "txtVence"
      Item(1).Control(12)=   "lsw"
      Item(1).Control(13)=   "FlatEdit2"
      Item(1).Control(14)=   "txtCasoId"
      Item(1).Control(15)=   "gbResolucion"
      Item(2).Caption =   "Histórico"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "lswHistorico"
      Begin XtremeSuiteControls.ListView lswHistorico 
         Height          =   7215
         Left            =   -70000
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   10095
         _Version        =   1572864
         _ExtentX        =   17806
         _ExtentY        =   12726
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
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   2655
         Left            =   -69760
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   9615
         _Version        =   1572864
         _ExtentX        =   16960
         _ExtentY        =   4683
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbResolucion 
         Height          =   2655
         Left            =   -69760
         TabIndex        =   20
         Top             =   4680
         Visible         =   0   'False
         Width           =   9375
         _Version        =   1572864
         _ExtentX        =   16531
         _ExtentY        =   4678
         _StockProps     =   79
         Caption         =   "Resolución: "
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
         Begin XtremeSuiteControls.ComboBox cboGestion 
            Height          =   312
            Left            =   1080
            TabIndex        =   21
            Top             =   480
            Width           =   5532
            _Version        =   1572864
            _ExtentX        =   9763
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
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboResolucion 
            Height          =   312
            Left            =   6600
            TabIndex        =   22
            Top             =   480
            Width           =   2772
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
         Begin XtremeSuiteControls.FlatEdit txtGestionNota 
            Height          =   792
            Left            =   1080
            TabIndex        =   23
            Top             =   1080
            Width           =   8292
            _Version        =   1572864
            _ExtentX        =   14626
            _ExtentY        =   1397
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
         Begin XtremeSuiteControls.PushButton btnGestion 
            Height          =   495
            Left            =   7920
            TabIndex        =   27
            Top             =   2040
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Gestión"
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
            Picture         =   "frmAF_CRRenuncia.frx":0000
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   5
            Left            =   1080
            TabIndex        =   26
            Top             =   840
            Width           =   1092
            _Version        =   1572864
            _ExtentX        =   1926
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Notas: "
            BackColor       =   -2147483633
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   6
            Left            =   6600
            TabIndex        =   25
            Top             =   240
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Resolución:"
            BackColor       =   -2147483633
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   7
            Left            =   1080
            TabIndex        =   24
            Top             =   240
            Width           =   1092
            _Version        =   1572864
            _ExtentX        =   1926
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Gestión: "
            BackColor       =   -2147483633
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
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtCasoId 
         Height          =   555
         Left            =   -69160
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   979
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
      Begin XtremeSuiteControls.FlatEdit txtVence 
         Height          =   315
         Left            =   -62200
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.FlatEdit txtResFecha 
         Height          =   315
         Left            =   -62200
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.FlatEdit txtRegFecha 
         Height          =   315
         Left            =   -62200
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.FlatEdit txtEstadoControl 
         Height          =   315
         Left            =   -66160
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.FlatEdit txtResUser 
         Height          =   315
         Left            =   -66160
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.FlatEdit txtRegUser 
         Height          =   315
         Left            =   -66160
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.FlatEdit FlatEdit1 
         Height          =   315
         Left            =   -67480
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2350
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
         Text            =   "Registro.: "
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit3 
         Height          =   315
         Left            =   -67480
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2350
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
         Text            =   "Resuelto.: "
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit5 
         Height          =   315
         Left            =   -67480
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2350
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
         Text            =   "Estado.: "
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit7 
         Height          =   315
         Left            =   -63640
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2561
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
         Text            =   "Fecha.: "
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit9 
         Height          =   315
         Left            =   -63640
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2561
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
         Text            =   "Fecha.: "
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit11 
         Height          =   315
         Left            =   -63640
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2561
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
         Text            =   "Vencimiento.: "
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit2 
         Height          =   555
         Left            =   -69760
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   615
         _Version        =   1572864
         _ExtentX        =   1080
         _ExtentY        =   974
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
         Text            =   "Id.:"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.TabControl tcMain 
         Height          =   8055
         Left            =   0
         TabIndex        =   28
         Top             =   360
         Width           =   10095
         _Version        =   1572864
         _ExtentX        =   17806
         _ExtentY        =   14208
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
         ItemCount       =   5
         Item(0).Caption =   "General"
         Item(0).ControlCount=   26
         Item(0).Control(0)=   "GroupBox1(4)"
         Item(0).Control(1)=   "GroupBox1(1)"
         Item(0).Control(2)=   "GroupBox1(0)"
         Item(0).Control(3)=   "lswMotivos"
         Item(0).Control(4)=   "txtPromotorCod"
         Item(0).Control(5)=   "txtCodigo"
         Item(0).Control(6)=   "txtEstado"
         Item(0).Control(7)=   "txtVencimiento"
         Item(0).Control(8)=   "cboTipo"
         Item(0).Control(9)=   "txtPromotorDesc"
         Item(0).Control(10)=   "txtNotas"
         Item(0).Control(11)=   "btnBoleta"
         Item(0).Control(12)=   "Label2(8)"
         Item(0).Control(13)=   "Label2(4)"
         Item(0).Control(14)=   "Label2(3)"
         Item(0).Control(15)=   "Label2(2)"
         Item(0).Control(16)=   "chkMortalidad"
         Item(0).Control(17)=   "chkReingreso"
         Item(0).Control(18)=   "chkVolver"
         Item(0).Control(19)=   "ListView1"
         Item(0).Control(20)=   "Label2(1)"
         Item(0).Control(21)=   "Label2(0)"
         Item(0).Control(22)=   "cboCausa"
         Item(0).Control(23)=   "chkTasaAjuste"
         Item(0).Control(24)=   "chkAltPlanilla"
         Item(0).Control(25)=   "btnActualizaDatos"
         Item(1).Caption =   "Patrimonio"
         Item(1).ControlCount=   31
         Item(1).Control(0)=   "txtRetenerMonto"
         Item(1).Control(1)=   "Label4(4)"
         Item(1).Control(2)=   "lblTotalNeto(0)"
         Item(1).Control(3)=   "Label4(2)"
         Item(1).Control(4)=   "Label4(1)"
         Item(1).Control(5)=   "lblTotalBruto"
         Item(1).Control(6)=   "Label4(0)"
         Item(1).Control(7)=   "Label3(0)"
         Item(1).Control(8)=   "lblAporteExtra"
         Item(1).Control(9)=   "lblCapitalizacion"
         Item(1).Control(10)=   "lblFCI"
         Item(1).Control(11)=   "lblAportePatronal"
         Item(1).Control(12)=   "lblAporteObrero"
         Item(1).Control(13)=   "chkAplObrero"
         Item(1).Control(14)=   "chkAplPatronal"
         Item(1).Control(15)=   "chkAplCapGen"
         Item(1).Control(16)=   "chkAplCapExtra"
         Item(1).Control(17)=   "lblRenta"
         Item(1).Control(18)=   "Label3(1)"
         Item(1).Control(19)=   "chkAplExcedente"
         Item(1).Control(20)=   "Label3(2)"
         Item(1).Control(21)=   "lblExcedenteRenta"
         Item(1).Control(22)=   "lblExcedente"
         Item(1).Control(23)=   "lblCustodia"
         Item(1).Control(24)=   "txtDivisa"
         Item(1).Control(25)=   "scTitulos(0)"
         Item(1).Control(26)=   "txtTipoCambio"
         Item(1).Control(27)=   "Label4(8)"
         Item(1).Control(28)=   "txtDivisaLocal"
         Item(1).Control(29)=   "Label5"
         Item(1).Control(30)=   "lswRenta"
         Item(2).Caption =   "Planes de Ahorros"
         Item(2).ControlCount=   10
         Item(2).Control(0)=   "lswPlanes"
         Item(2).Control(1)=   "Label4(5)"
         Item(2).Control(2)=   "lblTotalNeto(1)"
         Item(2).Control(3)=   "lblTotalNeto(2)"
         Item(2).Control(4)=   "Label4(3)"
         Item(2).Control(5)=   "Label4(6)"
         Item(2).Control(6)=   "Label4(7)"
         Item(2).Control(7)=   "txtFndRendGravado"
         Item(2).Control(8)=   "txtFndRendLiquidar"
         Item(2).Control(9)=   "scTitulos(1)"
         Item(3).Caption =   "Abonos"
         Item(3).ControlCount=   8
         Item(3).Control(0)=   "cmdDistribucionAuto"
         Item(3).Control(1)=   "fraAbono"
         Item(3).Control(2)=   "lblLsw"
         Item(3).Control(3)=   "Label2(9)"
         Item(3).Control(4)=   "lswAbonos"
         Item(3).Control(5)=   "chkArregloPago"
         Item(3).Control(6)=   "Label15"
         Item(3).Control(7)=   "txtSinpeNegativo"
         Item(4).Caption =   "Resumen"
         Item(4).ControlCount=   3
         Item(4).Control(0)=   "cmdGuardar"
         Item(4).Control(1)=   "txtSumario"
         Item(4).Control(2)=   "lblSumario"
         Begin XtremeSuiteControls.ListView lswRenta 
            Height          =   1215
            Left            =   -69040
            TabIndex        =   152
            Top             =   5880
            Visible         =   0   'False
            Width           =   6495
            _Version        =   1572864
            _ExtentX        =   11456
            _ExtentY        =   2143
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
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ListView lswAbonos 
            Height          =   5415
            Left            =   -69880
            TabIndex        =   141
            Top             =   1080
            Visible         =   0   'False
            Width           =   9855
            _Version        =   1572864
            _ExtentX        =   17383
            _ExtentY        =   9551
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
            View            =   3
            GridLines       =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ListView ListView1 
            Height          =   4695
            Left            =   -69880
            TabIndex        =   63
            Top             =   960
            Width           =   9855
            _Version        =   1572864
            _ExtentX        =   17383
            _ExtentY        =   8281
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
            View            =   3
            GridLines       =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.ListView lswPlanes 
            Height          =   5055
            Left            =   -69880
            TabIndex        =   62
            Top             =   840
            Visible         =   0   'False
            Width           =   9855
            _Version        =   1572864
            _ExtentX        =   17383
            _ExtentY        =   8916
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
            Checkboxes      =   -1  'True
            View            =   3
            GridLines       =   -1  'True
            FullRowSelect   =   -1  'True
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin VB.TextBox txtRetenerMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   -66520
            MultiLine       =   -1  'True
            TabIndex        =   61
            Text            =   "frmAF_CRRenuncia.frx":0727
            Top             =   4080
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Frame fraAbono 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4815
            Left            =   -69880
            TabIndex        =   29
            Top             =   1320
            Visible         =   0   'False
            Width           =   9855
            Begin XtremeSuiteControls.GroupBox GroupBox2 
               Height          =   975
               Left            =   0
               TabIndex        =   30
               Top             =   3720
               Width           =   9735
               _Version        =   1572864
               _ExtentX        =   17171
               _ExtentY        =   1720
               _StockProps     =   79
               ForeColor       =   16711680
               BackColor       =   16777215
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               BorderStyle     =   1
               Begin XtremeSuiteControls.PushButton cmdMAceptar 
                  Height          =   375
                  Left            =   6360
                  TabIndex        =   31
                  Top             =   360
                  Width           =   1335
                  _Version        =   1572864
                  _ExtentX        =   2350
                  _ExtentY        =   656
                  _StockProps     =   79
                  Caption         =   "Aceptar"
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
                  UseVisualStyle  =   -1  'True
                  Appearance      =   17
               End
               Begin XtremeSuiteControls.PushButton cmdMCancelar 
                  Height          =   375
                  Left            =   7680
                  TabIndex        =   32
                  Top             =   360
                  Width           =   1335
                  _Version        =   1572864
                  _ExtentX        =   2350
                  _ExtentY        =   656
                  _StockProps     =   79
                  Caption         =   "Cancelar"
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
                  UseVisualStyle  =   -1  'True
                  Appearance      =   17
               End
               Begin VB.Label lblDisponible 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "0"
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
                  Height          =   312
                  Left            =   1920
                  TabIndex        =   34
                  Top             =   360
                  Width           =   1812
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Disponible"
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
                  Left            =   240
                  TabIndex        =   33
                  Top             =   360
                  Width           =   1332
               End
            End
            Begin XtremeSuiteControls.FlatEdit txtMAbono 
               Height          =   315
               Left            =   5880
               TabIndex        =   150
               ToolTipText     =   "Tipo de Cambio"
               Top             =   3360
               Width           =   1815
               _Version        =   1572864
               _ExtentX        =   3201
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
               Text            =   "0"
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeShortcutBar.ShortcutCaption scTitulos 
               Height          =   375
               Index           =   2
               Left            =   0
               TabIndex        =   139
               Top             =   0
               Width           =   9855
               _Version        =   1572864
               _ExtentX        =   17383
               _ExtentY        =   661
               _StockProps     =   14
               Caption         =   "Aplicación manual del Abono"
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
               SubItemCaption  =   -1  'True
               Alignment       =   1
               ForeColor       =   16711680
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Mora/Vencido Total:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   240
               TabIndex        =   60
               Top             =   1920
               Width           =   1215
            End
            Begin VB.Label lblMMoraTotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   1560
               TabIndex        =   59
               Top             =   2040
               Width           =   1935
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Cargos"
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
               Left            =   4440
               TabIndex        =   58
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lblMCargos 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   5880
               TabIndex        =   57
               Top             =   1320
               Width           =   1815
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Pólizas"
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
               Left            =   4440
               TabIndex        =   56
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lblMPolizas 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   5880
               TabIndex        =   55
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label lblMLineaDesc 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   240
               TabIndex        =   54
               Top             =   3360
               Width           =   3495
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Abono"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   1
               Left            =   4440
               TabIndex        =   53
               Top             =   3360
               Width           =   1335
            End
            Begin VB.Label lblMMoraIntMor 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   5880
               TabIndex        =   52
               Top             =   2040
               Width           =   1815
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Int.Mor."
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
               Left            =   4440
               TabIndex        =   51
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label lblMMorIntCor 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   5880
               TabIndex        =   50
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Int.Cor."
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
               Left            =   4440
               TabIndex        =   49
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label lblMMorPrincipal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   5880
               TabIndex        =   48
               Top             =   2400
               Width           =   1815
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Principal"
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
               Left            =   4440
               TabIndex        =   47
               Top             =   2400
               Width           =   1095
            End
            Begin VB.Label lblMSaldo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   5880
               TabIndex        =   46
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo"
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
               Left            =   4440
               TabIndex        =   45
               Top             =   600
               Width           =   975
            End
            Begin VB.Label lblMTotalDeuda 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
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
               Height          =   315
               Left            =   5880
               TabIndex        =   44
               Top             =   3000
               Width           =   1815
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Deuda"
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
               Index           =   0
               Left            =   4440
               TabIndex        =   43
               Top             =   3000
               Width           =   1335
            End
            Begin VB.Label lblMGarantia 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   240
               TabIndex        =   42
               Top             =   3000
               Width           =   3495
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Garantía / Línea"
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
               Left            =   240
               TabIndex        =   41
               Top             =   2760
               Width           =   1455
            End
            Begin VB.Label lblMTipo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   1560
               TabIndex        =   40
               Top             =   1320
               Width           =   1935
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Left            =   240
               TabIndex        =   39
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lblMCodigo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   1560
               TabIndex        =   38
               Top             =   960
               Width           =   1935
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Línea"
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
               Left            =   240
               TabIndex        =   37
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lblMOperacion 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   1560
               TabIndex        =   36
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Operación"
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
               Left            =   240
               TabIndex        =   35
               Top             =   600
               Width           =   975
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   1095
            Index           =   4
            Left            =   240
            TabIndex        =   64
            Top             =   4560
            Width           =   9615
            _Version        =   1572864
            _ExtentX        =   16960
            _ExtentY        =   1931
            _StockProps     =   79
            Caption         =   "Desembolso"
            ForeColor       =   8421504
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
            Begin XtremeSuiteControls.ComboBox cboBanco 
               Height          =   330
               Left            =   1680
               TabIndex        =   65
               Top             =   360
               Width           =   5175
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
            Begin XtremeSuiteControls.ComboBox cboCuenta 
               Height          =   330
               Left            =   1680
               TabIndex        =   66
               Top             =   720
               Width           =   5175
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
            Begin XtremeSuiteControls.ComboBox cboTipoDoc 
               Height          =   315
               Left            =   7920
               TabIndex        =   67
               Top             =   360
               Width           =   1575
               _Version        =   1572864
               _ExtentX        =   2778
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
            Begin XtremeSuiteControls.DateTimePicker dtpPago 
               Height          =   315
               Left            =   7920
               TabIndex        =   68
               Top             =   720
               Width           =   1575
               _Version        =   1572864
               _ExtentX        =   2773
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
            Begin VB.Label Label16 
               Caption         =   "F.Pago"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   7200
               TabIndex        =   72
               Top             =   750
               Width           =   1305
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
               Index           =   7
               Left            =   600
               TabIndex        =   71
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label1 
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
               Height          =   255
               Index           =   6
               Left            =   600
               TabIndex        =   70
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "T.Doc."
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
               Index           =   5
               Left            =   7200
               TabIndex        =   69
               Top             =   360
               Width           =   615
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   1575
            Index           =   1
            Left            =   240
            TabIndex        =   73
            Top             =   6000
            Width           =   3015
            _Version        =   1572864
            _ExtentX        =   5318
            _ExtentY        =   2778
            _StockProps     =   79
            Caption         =   "Datos Actuales "
            ForeColor       =   8421504
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
            Begin XtremeSuiteControls.Label lblBoleta 
               Height          =   315
               Left            =   1200
               TabIndex        =   155
               Top             =   1080
               Width           =   1695
               _Version        =   1572864
               _ExtentX        =   2990
               _ExtentY        =   556
               _StockProps     =   79
               ForeColor       =   16711680
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
               Alignment       =   2
               WordWrap        =   -1  'True
            End
            Begin XtremeSuiteControls.Label lblIngreso 
               Height          =   315
               Left            =   1200
               TabIndex        =   154
               Top             =   720
               Width           =   1695
               _Version        =   1572864
               _ExtentX        =   2990
               _ExtentY        =   556
               _StockProps     =   79
               ForeColor       =   16711680
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
               Alignment       =   2
               WordWrap        =   -1  'True
            End
            Begin XtremeSuiteControls.Label lblEstadoActual 
               Height          =   315
               Left            =   1200
               TabIndex        =   153
               Top             =   360
               Width           =   1695
               _Version        =   1572864
               _ExtentX        =   2990
               _ExtentY        =   556
               _StockProps     =   79
               ForeColor       =   16711680
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
               Alignment       =   2
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               Caption         =   "Boleta"
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
               Index           =   11
               Left            =   120
               TabIndex        =   76
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Ingreso"
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
               Left            =   120
               TabIndex        =   75
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label1 
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
               Height          =   252
               Index           =   9
               Left            =   120
               TabIndex        =   74
               Top             =   360
               Width           =   612
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   1575
            Index           =   0
            Left            =   3360
            TabIndex        =   77
            Top             =   6000
            Width           =   6495
            _Version        =   1572864
            _ExtentX        =   11456
            _ExtentY        =   2778
            _StockProps     =   79
            Caption         =   "Datos de la Acción de Personal "
            ForeColor       =   8421504
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
            Begin XtremeSuiteControls.DateTimePicker dtpAc_fecha 
               Height          =   315
               Left            =   3840
               TabIndex        =   78
               Top             =   360
               Width           =   1575
               _Version        =   1572864
               _ExtentX        =   2773
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
            Begin XtremeSuiteControls.FlatEdit txtAc_Boleta 
               Height          =   315
               Left            =   3840
               TabIndex        =   79
               Top             =   1080
               Width           =   1575
               _Version        =   1572864
               _ExtentX        =   2778
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
            Begin XtremeSuiteControls.ComboBox cboAC_Tipo 
               Height          =   315
               Left            =   3840
               TabIndex        =   142
               Top             =   720
               Width           =   1575
               _Version        =   1572864
               _ExtentX        =   2778
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
            Begin XtremeSuiteControls.Label lblAcFecha 
               Height          =   315
               Left            =   3840
               TabIndex        =   159
               Top             =   360
               Width           =   1575
               _Version        =   1572864
               _ExtentX        =   2778
               _ExtentY        =   556
               _StockProps     =   79
               ForeColor       =   16711680
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
               Alignment       =   2
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblAc_Boleta 
               Alignment       =   1  'Right Justify
               Caption         =   "Boleta"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   3000
               TabIndex        =   157
               Top             =   1080
               Width           =   660
            End
            Begin VB.Label lblAC_Tipo 
               Alignment       =   1  'Right Justify
               Caption         =   "Tipo Acción"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   2520
               TabIndex        =   156
               Top             =   720
               Width           =   1140
            End
            Begin VB.Image imgFechaAccion 
               Height          =   240
               Left            =   5640
               Picture         =   "frmAF_CRRenuncia.frx":0729
               Top             =   375
               Width           =   240
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "Rige a partir de"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   2280
               TabIndex        =   80
               Top             =   390
               Width           =   1425
            End
         End
         Begin XtremeSuiteControls.CheckBox chkAplObrero 
            Height          =   252
            Left            =   -69520
            TabIndex        =   81
            Top             =   1080
            Visible         =   0   'False
            Width           =   2532
            _Version        =   1572864
            _ExtentX        =   4466
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplica Aporte Obrero"
            ForeColor       =   0
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
            Enabled         =   0   'False
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.CheckBox chkAplPatronal 
            Height          =   252
            Left            =   -69520
            TabIndex        =   82
            Top             =   1440
            Visible         =   0   'False
            Width           =   2532
            _Version        =   1572864
            _ExtentX        =   4466
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplica Aporte Patronal"
            ForeColor       =   0
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
            Enabled         =   0   'False
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkAplCapGen 
            Height          =   252
            Left            =   -69520
            TabIndex        =   83
            Top             =   2280
            Visible         =   0   'False
            Width           =   2532
            _Version        =   1572864
            _ExtentX        =   4466
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplica Capitalización"
            ForeColor       =   0
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
            Enabled         =   0   'False
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkAplCapExtra 
            Height          =   252
            Left            =   -69520
            TabIndex        =   84
            Top             =   3120
            Visible         =   0   'False
            Width           =   2532
            _Version        =   1572864
            _ExtentX        =   4466
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplica Ahorro Extraordinario"
            ForeColor       =   0
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
            Enabled         =   0   'False
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton cmdDistribucionAuto 
            Height          =   495
            Left            =   -66520
            TabIndex        =   85
            Top             =   405
            Visible         =   0   'False
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Distribución Automática"
            ForeColor       =   0
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAF_CRRenuncia.frx":0846
         End
         Begin XtremeSuiteControls.FlatEdit txtFndRendLiquidar 
            Height          =   315
            Left            =   -67480
            TabIndex        =   86
            Top             =   6120
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
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
            Text            =   "0"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFndRendGravado 
            Height          =   315
            Left            =   -67480
            TabIndex        =   87
            Top             =   6480
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
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
            Text            =   "0"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkAplExcedente 
            Height          =   252
            Left            =   -69520
            TabIndex        =   88
            Top             =   2640
            Visible         =   0   'False
            Width           =   2892
            _Version        =   1572864
            _ExtentX        =   5101
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplica Excedente del Periodo"
            ForeColor       =   0
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtDivisa 
            Height          =   312
            Left            =   -64840
            TabIndex        =   89
            ToolTipText     =   "Divisa Origen"
            Top             =   5040
            Visible         =   0   'False
            Width           =   612
            _Version        =   1572864
            _ExtentX        =   1080
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
            Text            =   "COL"
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTipoCambio 
            Height          =   312
            Left            =   -66520
            TabIndex        =   90
            ToolTipText     =   "Tipo de Cambio"
            Top             =   5040
            Visible         =   0   'False
            Width           =   1692
            _Version        =   1572864
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
            Text            =   "0"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDivisaLocal 
            Height          =   312
            Left            =   -64840
            TabIndex        =   91
            ToolTipText     =   "Divisa Convertida"
            Top             =   4680
            Visible         =   0   'False
            Width           =   612
            _Version        =   1572864
            _ExtentX        =   1080
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
            Text            =   "COL"
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ListView lswMotivos 
            Height          =   1695
            Left            =   1920
            TabIndex        =   92
            Top             =   1320
            Width           =   3495
            _Version        =   1572864
            _ExtentX        =   6165
            _ExtentY        =   2990
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
            View            =   3
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPromotorCod 
            Height          =   315
            Left            =   1920
            TabIndex        =   93
            Top             =   3480
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2984
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCodigo 
            Height          =   330
            Left            =   1920
            TabIndex        =   94
            Top             =   480
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.FlatEdit txtEstado 
            Height          =   330
            Left            =   3600
            TabIndex        =   95
            Top             =   480
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.FlatEdit txtVencimiento 
            Height          =   330
            Left            =   5880
            TabIndex        =   96
            Top             =   480
            Width           =   2775
            _Version        =   1572864
            _ExtentX        =   4895
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.ComboBox cboCausa 
            Height          =   330
            Left            =   3600
            TabIndex        =   97
            Top             =   840
            Width           =   6135
            _Version        =   1572864
            _ExtentX        =   10821
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
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboTipo 
            Height          =   330
            Left            =   1920
            TabIndex        =   98
            Top             =   840
            Width           =   1695
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
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtPromotorDesc 
            Height          =   315
            Left            =   3600
            TabIndex        =   99
            Top             =   3480
            Width           =   6135
            _Version        =   1572864
            _ExtentX        =   10821
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
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   675
            Left            =   1920
            TabIndex        =   100
            Top             =   3840
            Width           =   7815
            _Version        =   1572864
            _ExtentX        =   13785
            _ExtentY        =   1191
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
         Begin XtremeSuiteControls.PushButton btnBoleta 
            Height          =   315
            Left            =   8760
            TabIndex        =   101
            Top             =   480
            Width           =   735
            _Version        =   1572864
            _ExtentX        =   1291
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Boleta"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
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
         Begin XtremeSuiteControls.CheckBox chkMortalidad 
            Height          =   255
            Left            =   5640
            TabIndex        =   102
            Top             =   1320
            Width           =   5655
            _Version        =   1572864
            _ExtentX        =   9975
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Renuncia por Mortalidad"
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
         Begin XtremeSuiteControls.CheckBox chkReingreso 
            Height          =   255
            Left            =   5640
            TabIndex        =   103
            Top             =   2760
            Width           =   3255
            _Version        =   1572864
            _ExtentX        =   5741
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Aplica para Re-Ingreso Automático"
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
         Begin XtremeSuiteControls.CheckBox chkVolver 
            Height          =   255
            Left            =   5640
            TabIndex        =   104
            Top             =   2400
            Width           =   5655
            _Version        =   1572864
            _ExtentX        =   9975
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Esta dispuesto a volver afiliarse a futuro ?"
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
         Begin XtremeSuiteControls.CheckBox chkTasaAjuste 
            Height          =   255
            Left            =   5640
            TabIndex        =   143
            Top             =   2040
            Width           =   5655
            _Version        =   1572864
            _ExtentX        =   9975
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Aplicar Aumento de Tasas"
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
         Begin XtremeSuiteControls.PushButton cmdGuardar 
            Height          =   495
            Left            =   -61600
            TabIndex        =   147
            Top             =   6600
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Guardar"
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
            Picture         =   "frmAF_CRRenuncia.frx":0F5F
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.FlatEdit txtSumario 
            Height          =   5775
            Left            =   -70000
            TabIndex        =   148
            Top             =   720
            Visible         =   0   'False
            Width           =   10095
            _Version        =   1572864
            _ExtentX        =   17806
            _ExtentY        =   10186
            _StockProps     =   77
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   2
         End
         Begin XtremeSuiteControls.CheckBox chkAltPlanilla 
            Height          =   255
            Left            =   5640
            TabIndex        =   158
            Top             =   1680
            Width           =   4455
            _Version        =   1572864
            _ExtentX        =   7853
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cancela Créditos Pendientes por Medio de Planilla ?"
            ForeColor       =   0
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
         Begin XtremeSuiteControls.CheckBox chkArregloPago 
            Height          =   255
            Left            =   -64360
            TabIndex        =   160
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Arreglo Pago?"
            ForeColor       =   16711680
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
         End
         Begin XtremeSuiteControls.FlatEdit txtSinpeNegativo 
            Height          =   330
            Left            =   -67360
            TabIndex        =   162
            Top             =   6720
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   582
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
            Text            =   "0"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnActualizaDatos 
            Height          =   315
            Left            =   6360
            TabIndex        =   164
            Top             =   3120
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Actualiza Afiliación"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
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
         Begin XtremeSuiteControls.Label Label15 
            Height          =   330
            Left            =   -69760
            TabIndex        =   161
            Top             =   6720
            Visible         =   0   'False
            Width           =   2415
            _Version        =   1572864
            _ExtentX        =   4260
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Fondo de Sinpe Negativo?"
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
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   255
            Left            =   -69040
            TabIndex        =   151
            Top             =   5520
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Detalle de Renta"
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
         Begin XtremeShortcutBar.ShortcutCaption lblSumario 
            Height          =   375
            Left            =   -70000
            TabIndex        =   149
            Top             =   360
            Visible         =   0   'False
            Width           =   10095
            _Version        =   1572864
            _ExtentX        =   17806
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "..:: Sumario ::.."
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
            SubItemCaption  =   -1  'True
            Alignment       =   1
            ForeColor       =   16711680
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   510
            Index           =   9
            Left            =   -69880
            TabIndex        =   140
            Top             =   360
            Visible         =   0   'False
            Width           =   3495
            _Version        =   1572864
            _ExtentX        =   6165
            _ExtentY        =   900
            _StockProps     =   79
            Caption         =   "Indique o Modifique con Doble Click el Abono a Cada Operación"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   138
            Top             =   480
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Código"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   510
            Index           =   1
            Left            =   -69880
            TabIndex        =   137
            Top             =   360
            Width           =   3975
            _Version        =   1572864
            _ExtentX        =   7011
            _ExtentY        =   900
            _StockProps     =   79
            Caption         =   "Indique o Modifique con Doble Click el Abono a Cada Operación"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblLsw 
            Height          =   495
            Left            =   -63040
            TabIndex        =   136
            Top             =   360
            Visible         =   0   'False
            Width           =   2895
            _Version        =   1572864
            _ExtentX        =   5106
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "..."
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
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
         Begin VB.Label Label4 
            Caption         =   "Tipo de Cambio/ Divisa Origen"
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
            Left            =   -69040
            TabIndex        =   135
            Top             =   5040
            Visible         =   0   'False
            Width           =   2532
         End
         Begin XtremeShortcutBar.ShortcutCaption scTitulos 
            Height          =   375
            Index           =   1
            Left            =   -70000
            TabIndex        =   134
            Top             =   360
            Visible         =   0   'False
            Width           =   10215
            _Version        =   1572864
            _ExtentX        =   18018
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Marque los Planes de Ahorros a Liquidar"
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
            SubItemCaption  =   -1  'True
            Alignment       =   1
            ForeColor       =   16711680
         End
         Begin XtremeShortcutBar.ShortcutCaption scTitulos 
            Height          =   375
            Index           =   0
            Left            =   -70000
            TabIndex        =   133
            Top             =   360
            Visible         =   0   'False
            Width           =   10095
            _Version        =   1572864
            _ExtentX        =   17806
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Indique los Aportes a Utilizar en Abonos a Deudas"
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
            SubItemCaption  =   -1  'True
            Alignment       =   1
            ForeColor       =   16711680
         End
         Begin VB.Label lblCustodia 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   312
            Left            =   -66520
            TabIndex        =   132
            Top             =   1800
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Label lblExcedente 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   312
            Left            =   -66520
            TabIndex        =   131
            Top             =   2640
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Label lblExcedenteRenta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   312
            Left            =   -64840
            TabIndex        =   130
            Top             =   2640
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "I.R."
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
            Height          =   312
            Index           =   2
            Left            =   -63160
            TabIndex        =   129
            ToolTipText     =   "Impuesto de Renta"
            Top             =   2640
            Visible         =   0   'False
            Width           =   492
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Rendimiento Gravado:"
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
            Left            =   -70000
            TabIndex        =   128
            Top             =   6480
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Rendimiento a Liquidar:"
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
            Left            =   -70000
            TabIndex        =   127
            Top             =   6120
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Impuesto Renta s/Rend:"
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
            Index           =   3
            Left            =   -65080
            TabIndex        =   126
            Top             =   6120
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label lblTotalNeto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   315
            Index           =   2
            Left            =   -62200
            TabIndex        =   125
            Top             =   6120
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "I.R."
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
            Height          =   315
            Index           =   1
            Left            =   -63160
            TabIndex        =   124
            ToolTipText     =   "Impuesto de Renta"
            Top             =   2280
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblRenta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   312
            Left            =   -64840
            TabIndex        =   123
            Top             =   2280
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Label lblTotalNeto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   315
            Index           =   1
            Left            =   -62200
            TabIndex        =   122
            Top             =   6480
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Neto Disponible:"
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
            Index           =   5
            Left            =   -65080
            TabIndex        =   121
            Top             =   6480
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label lblAporteObrero 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   312
            Left            =   -66520
            TabIndex        =   120
            Top             =   1080
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Label lblAportePatronal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   312
            Left            =   -66520
            TabIndex        =   119
            Top             =   1440
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Label lblFCI 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   312
            Left            =   -64840
            TabIndex        =   118
            Top             =   1440
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Label lblCapitalizacion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   312
            Left            =   -66520
            TabIndex        =   117
            Top             =   2280
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Label lblAporteExtra 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   312
            Left            =   -66520
            TabIndex        =   116
            Top             =   3120
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "F.C.I."
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
            Height          =   312
            Index           =   0
            Left            =   -63160
            TabIndex        =   115
            ToolTipText     =   "Fondo de Capitalización Individual"
            Top             =   1440
            Visible         =   0   'False
            Width           =   492
         End
         Begin VB.Label Label4 
            Caption         =   "Total Bruto Disponible"
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
            Left            =   -69040
            TabIndex        =   114
            Top             =   3720
            Visible         =   0   'False
            Width           =   2412
         End
         Begin VB.Label lblTotalBruto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   312
            Left            =   -66520
            TabIndex        =   113
            Top             =   3720
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Label Label4 
            Caption         =   "Retener Monto por la Suma de"
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
            Left            =   -69040
            TabIndex        =   112
            Top             =   4080
            Visible         =   0   'False
            Width           =   2412
         End
         Begin VB.Label Label4 
            Caption         =   "Total Neto Disponible"
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
            Index           =   2
            Left            =   -69040
            TabIndex        =   111
            Top             =   4680
            Visible         =   0   'False
            Width           =   2412
         End
         Begin VB.Label lblTotalNeto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   312
            Index           =   0
            Left            =   -66520
            TabIndex        =   110
            Top             =   4680
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Label Label4 
            Caption         =   "Impuesto Renta [Pendiente] sobre Capitalización + Adelanto Excedentes"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   612
            Index           =   4
            Left            =   -64720
            TabIndex        =   109
            Top             =   4080
            Visible         =   0   'False
            Width           =   3252
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   375
            Index           =   8
            Left            =   360
            TabIndex        =   108
            Top             =   1800
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1926
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Motivos Específicos"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   107
            Top             =   3840
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Notas"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   106
            Top             =   3480
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1926
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Ejecutivo"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   105
            Top             =   840
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Tipo y Causa"
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
      Begin XtremeSuiteControls.PushButton btnSiguiente 
         Height          =   330
         Left            =   8760
         TabIndex        =   144
         Top             =   0
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Siguiente"
         ForeColor       =   0
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_CRRenuncia.frx":1690
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton btnAnterior 
         Height          =   330
         Left            =   7440
         TabIndex        =   145
         Top             =   0
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Anterior"
         ForeColor       =   0
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_CRRenuncia.frx":1F4B
      End
      Begin XtremeSuiteControls.PushButton cmdNuevo 
         Height          =   330
         Left            =   6360
         TabIndex        =   146
         Top             =   0
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1931
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Nuevo"
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
         Picture         =   "frmAF_CRRenuncia.frx":2809
         ImageAlignment  =   4
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   4080
      TabIndex        =   0
      Top             =   360
      Width           =   4935
      _Version        =   1572864
      _ExtentX        =   8700
      _ExtentY        =   582
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
      Height          =   330
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.PushButton btnAdjuntos 
      Height          =   330
      Left            =   9000
      TabIndex        =   163
      ToolTipText     =   "Adjuntar Documentos"
      Top             =   360
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Adjuntos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Picture         =   "frmAF_CRRenuncia.frx":2E3B
   End
   Begin XtremeSuiteControls.Label lblCbrStatus 
      Height          =   255
      Left            =   2400
      TabIndex        =   165
      Top             =   720
      Visible         =   0   'False
      Width           =   6615
      _Version        =   1572864
      _ExtentX        =   11668
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Esta Persona tiene operaciones en Cobro Judicial"
      ForeColor       =   16777215
      BackColor       =   8421631
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Id.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.Image imgBanner 
      Height          =   1005
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_CRRenuncia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCt As String, vUp As String, fcVencimiento As String
Dim vCodigo As Long, vEstado As String, bConsulta As Boolean
Dim vPaso As Boolean, vModoSIF As Boolean, mFechaSistema As Date

Dim mTotalLiq As Currency, mTotalRetenido As Currency, mTotalNeto As Currency

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub btnActualizaDatos_Click()
Dim frm As Form

On Error GoTo vError

If txtCedula.Text <> "" Then
    Call sbFormsCall("frmAF_Principal", , , , , Me, True)
    Call sbFormActivo("frmAF_Principal", frm)
    
  Me.MousePointer = vbHourglass
    
    Sleep 2000
    frm.sbConsultaExterna (txtCedula.Text)
  Me.MousePointer = vbDefault

End If

Exit Sub

vError:

End Sub

Private Sub btnAdjuntos_Click()
If txtCedula.Text <> "" Then
 gGA.Modulo = "CL_03"
 gGA.Llave_01 = txtCedula.Text
 gGA.Llave_02 = ""
 gGA.Llave_03 = ""
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End If
End Sub

Private Sub btnBoleta_Click()
If Not IsNumeric(vCodigo) Then Exit Sub

Call sbRenunciaBoleta(vCodigo)

End Sub

Private Sub btnGestion_Click()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spAFI_Renuncia_CambioEstado " & vCodigo & ", '" & Mid(cboResolucion, 1, 1) & "', '" & cboGestion.ItemData(cboGestion.ListIndex) _
       & "', '" & txtGestionNota.Text & "', '" & glogon.Usuario & "', '" & glogon.Maquina & "', '" & glogon.AppVersion & "'"
Call ConectionExecute(strSQL)


txtGestionNota.Text = ""

Me.MousePointer = vbDefault
MsgBox "Información Registrada satisfactoriamente", vbInformation

Call sbConsulta(vCodigo)

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description)
End Sub


Private Function fxSegGestion(vRenuncia As Long, vCodg As String) As Integer
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select isnull(Max(id),0) as consecutivo from afi_cr_seguimiento" _
        & " where cod_renuncia = " & vRenuncia & " and cod_gestion = '" & vCodg & "'"
        
Call OpenRecordSet(rs, strSQL)
  fxSegGestion = rs!Consecutivo + 1
rs.Close

End Function




Private Sub cboCausa_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pCausa As Long
If vPaso Then Exit Sub
If cboCausa.ListCount = 0 Then Exit Sub

chkMortalidad.Value = vbUnchecked
chkAltPlanilla.Value = vbChecked

chkMortalidad.Enabled = False
chkAltPlanilla.Enabled = False

chkTasaAjuste.Enabled = True
chkTasaAjuste.Value = xtpChecked


chkVolver.Enabled = True

pCausa = cboCausa.ItemData(cboCausa.ListIndex)

strSQL = "select mortalidad, liq_alterna, Tipo_Apl, AJUSTE_TASAS from causas_renuncias where id_causa = " & pCausa
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then

  If rs!Mortalidad = 1 Then
     chkMortalidad.Enabled = True
     chkMortalidad.Value = vbChecked
'     cboTipoDoc.Clear
'     cboTipoDoc.AddItem "Fondo"
'     cboTipoDoc.ItemData(cboTipoDoc.ListCount - 1) = "FD"
'     cboTipoDoc.Text = "Fondo"
     
     chkTasaAjuste.Value = xtpUnchecked
     chkTasaAjuste.Enabled = False
     
     chkVolver.Value = xtpUnchecked
     chkVolver.Enabled = False
  
  End If
  
  If rs!liq_Alterna = 1 Then
     chkAltPlanilla.Value = vbChecked
     chkAltPlanilla.Enabled = True
  End If
  
  'Cambia a Opcion de Arreglo de Pago
  If rs!AJUSTE_TASAS = 1 Then
     chkArregloPago.Value = xtpChecked
     chkArregloPago.Enabled = True
  Else
     chkArregloPago.Value = xtpUnchecked
     chkArregloPago.Enabled = False
  End If
  
  
  If rs!Tipo_Apl = "P" Then
     chkReingreso.Value = xtpUnchecked
     chkReingreso.Enabled = False
  
     chkVolver.Value = xtpUnchecked
     chkVolver.Enabled = False
  
  Else
     chkReingreso.Enabled = True
  End If
  
End If
rs.Close

'Actualiza Combos de Emision
strSQL = "exec spAFI_Renuncia_Emite_TDoc " & cboBanco.ItemData(cboBanco.ListIndex) & ", " & chkMortalidad.Value _
       & ", '" & txtCedula.Text & "', '" & IIf((Mid(cboTipo.Text, 1, 2) = "01"), "A", "P") & "', " & pCausa
Call sbCbo_Llena_New(cboTipoDoc, strSQL, False, True)

End Sub

Private Sub txtCuentaAhorros_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call btnSiguiente_Click

End Sub

Private Sub cboTipo_Click()

Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

If Mid(cboTipo.Text, 1, 2) = "01" Then
    dtpAc_fecha.Visible = False
    txtAc_Boleta.Visible = False
    cboAc_Tipo.Visible = False
Else
    dtpAc_fecha.Visible = True
    txtAc_Boleta.Visible = True
    cboAc_Tipo.Visible = True
End If

lblAc_Boleta.Visible = txtAc_Boleta.Visible
lblAC_Tipo.Visible = cboAc_Tipo.Visible


vPaso = True
    'Carga Causas
    strSQL = "select id_Causa as 'IdX', Descripcion as 'ItmX'" _
           & " from causas_renuncias WHERE ACTIVO = 1" _
           & " and Tipo_Apl in('A', '" & IIf((Mid(cboTipo.Text, 1, 2) = "01"), "I", "P") & "')"
    Call sbCbo_Llena_New(cboCausa, strSQL, False, True)
vPaso = False
Call cboCausa_Click

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub cboTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboBanco.SetFocus
End Sub


Private Sub chkMortalidad_Click()
'If chkMortalidad Then
'    cboTipo.Locked = True
'Else
'   cboTipo.Locked = False
'End If
End Sub


Private Function fxReIngreso_Valida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
Dim vMensaje As String

vMensaje = ""


'--------------------------------------------------------------------------------------------------------------------
strSQL = "select dbo.fxAFI_Afiliacion_Valida_Beneficiarios(S.Cedula) as 'Beneficiarios', getdate() as 'Fecha'" _
       & ", dbo.fxAFI_Afiliacion_Valida_Telefonos(S.Cedula) as 'Telefonos'" _
       & ", dbo.fxAFI_Afiliacion_Valida_Beneficiarios_MenoresSinAlbacea(S.Cedula) as 'MenoresSinAlbacea'" _
       & ", S.AF_Email, S.Email_02, isnull(S.SALARIO_DIVISA,'COL') as 'SalarioDivisa', isnull(S.SALARIO_MONTO,0) as 'SalarioDevengado'" _
       & ", isnull(S.I_BENEFICIARIOS,0) as 'I_Beneficiario', isnull(S.FECHA_NAC, getdate()) as 'FNac', isnull(S.FECHA_VEN_CED, getdate()) as 'FCed'" _
       & ", S.Id_promotor, S.Cod_Nacionalidad, S.Cod_Pais_Nac, isnull(S.EstadoLaboral, '') as 'EstadoLab', isnull(S.ACTIVIDADES, 0) as 'C_Actividad'" _
       & ", isnull(S.cod_departamento,'') as 'Dept', isnull(S.UP,'') as 'UP', isnull(S.COD_CARGO,'') as 'Puesto', isnull(S.NIVEL_ACADEMICO,'0') as 'Nivel'" _
       & ", isnull(S.cod_profesion,0) as 'Profesion', isnull(S.Albacea_cedula,'') as 'Albacea'" _
       & ", isnull(S.Provincia,'') as 'Provincia', isnull(S.Canton,'') as 'Canton', isnull(S.Distrito,'') as 'Distrito', isnull(S.Direccion,'') as 'Direccion'" _
       & " from socios S where S.Cedula = '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then

      If rs!FNac > DateAdd("yyyy", -17, rs!fecha) Then
        vMensaje = vMensaje & " - Verifique la Fecha de nacimiento, la persona es menor de edad!" & vbCrLf
      End If
    
      If rs!FCed <= DateAdd("d", 20, rs!fecha) Then
        vMensaje = vMensaje & " - Verifique la fecha de Vencimiento del documento de Identidad, está pronta a vencer!" & vbCrLf
      End If

    If rs!Profesion = 0 Then vMensaje = vMensaje & " - Profesion no es válida!" & vbCrLf
    If rs!EstadoLab = "" Then vMensaje = vMensaje & " - No se especificó el Estado Laboral" & vbCrLf
    If Trim(rs!nivel) = "0" Then vMensaje = vMensaje & " - No se especificó el nivel academico" & vbCrLf
    If rs!C_Actividad = "0" Then vMensaje = vMensaje & " - No se especificó Actividad (Oficina Cumplimiento)" & vbCrLf
    
 
    
    If GLOBALES.SysASEVersion Then
        If Len(rs!Dept) + Len(rs!UP) = 0 Then
           vMensaje = vMensaje & " - No se especificó el Departamento o Unidad Programatica" & vbCrLf
        End If
    End If
    
    If Len(rs!PUESTO) = 0 Then vMensaje = vMensaje & " - Tienen que indicar el Puesto que desempeña" & vbCrLf
    
    If rs!Beneficiarios = 0 And rs!I_Beneficiario = 1 Then
          vMensaje = vMensaje & " - No se han registrado los Beneficiarios o Estan incompletos!" & vbCrLf
    End If
    
    If Not fxEmail_Valida(rs!AF_Email) Then
        vMensaje = vMensaje & " - El Email principal no es válido!" & vbCrLf
    End If
    
    If Len(Trim(rs!Email_02 & "")) > 0 Then
        If Not fxEmail_Valida(rs!Email_02 & "") Then
            vMensaje = vMensaje & " - El Email secundario no es válido!" & vbCrLf
        End If
    End If
    
    If rs!SalarioDivisa = "COL" Then
        If rs!SalarioDevengado < 100000 Or rs!SalarioDevengado > 10000000 Then vMensaje = vMensaje & " - Salario Devengado no es válido" & vbCrLf
    Else
        If rs!SalarioDevengado < 200 Or rs!SalarioDevengado > 20000 Then vMensaje = vMensaje & " - Salario Devengado no es válido" & vbCrLf
    End If
    
    If rs!Provincia = "" Then vMensaje = vMensaje & " - No se especificó la Provincia" & vbCrLf
    If rs!Canton = "" Then vMensaje = vMensaje & " - No se especificó el Cantón" & vbCrLf
    If rs!distrito = "" Then vMensaje = vMensaje & " - No se especificó el Distrito en la dirección" & vbCrLf
    
    'If Trim(txtDireccion) = "" Then vMensaje = vMensaje & " - No se especificó la Dirección" & vbCrLf
    If Not fxDireccion_Valida(Trim(rs!direccion), "-,#,*") Then vMensaje = vMensaje & " - No se especificó la Dirección" & vbCrLf
    
    
    If rs!MenoresSinAlbacea > 0 And rs!I_Beneficiario = 1 And Len(rs!Albacea) <= 5 Then
          vMensaje = vMensaje & " - Existen Beneficiarios Menores de Edad y no se han indicado lo(s) Albacea(s)!" & vbCrLf
    End If
    
    If rs!Telefonos = 0 Then
          vMensaje = vMensaje & " - No se han registrado los telefonos de contacto!" & vbCrLf
    End If

End If

If Len(vMensaje) = 0 Then
  fxReIngreso_Valida = True
Else
  fxReIngreso_Valida = False
  MsgBox vMensaje, vbExclamation

End If

End Function

Private Sub chkReingreso_Click()

If chkReingreso.Value = xtpChecked Then
    chkVolver.Value = xtpChecked
    chkVolver.Enabled = False
    
    If Not fxReIngreso_Valida() Then
        chkReingreso.Value = xtpUnchecked
    End If

Else
    chkVolver.Value = xtpUnchecked
    chkVolver.Enabled = True
End If

End Sub


Private Sub sbBoleta_Afiliacion()

On Error GoTo vError

Me.MousePointer = vbHourglass


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes Módulo de Personas"
 
 .Connect = glogon.ConectRPT
 
  .ReportFileName = SIFGlobal.fxPathReportes("Personas_Boleta_Afiliacion.rpt")
  
  .StoredProcParam(0) = Trim(txtCedula.Text)
  .StoredProcParam(1) = 0

  .SubreportToChange = "sbBeneficiarios"
  .StoredProcParam(0) = Trim(txtCedula.Text)
   
 .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:

    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cmdGuardar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

Dim vAC_Fecha As String, vAC_TipoBoleta As String

On Error GoTo vError

If Not fxVerificaDatos Then Exit Sub
    
    
    
If lblAC_Tipo.Visible Then
    vAC_Fecha = "'" & Format(dtpAc_fecha.Value, "yyyy-mm-dd") & "'"
    vAC_TipoBoleta = cboAc_Tipo.ItemData(cboAc_Tipo.ListIndex)
Else
    vAC_Fecha = "Null"
    vAC_TipoBoleta = "Null"
End If
    
  
strSQL = "exec spAFI_Renuncia_Liquidacion_Guarda " & vCodigo & ", '" & Trim(txtCedula.Text) & "', " & cboCausa.ItemData(cboCausa.ListIndex) _
       & ", " & txtPromotorCod.Text & ", " & chkMortalidad.Value & ", " & chkReingreso.Value & ", " & chkAltPlanilla.Value & ", " & chkVolver.Value _
       & ", " & chkTasaAjuste.Value & ", " & chkAplObrero.Value & ", " & chkAplPatronal.Value & ", " & chkAplCapGen.Value _
       & ", " & chkAplCapExtra.Value & ", 0, '" & IIf((Mid(cboTipo.Text, 1, 2) = "01"), "A", "P") & "', '" & glogon.Usuario _
       & "', '" & Mid(txtNotas.Text, 1, 1000) & "', '" & GLOBALES.gOficinaTitular & "', '" & cboTipoDoc.ItemData(cboTipoDoc.ListIndex) & "', " & cboBanco.ItemData(cboBanco.ListIndex) _
       & ", '" & cboCuenta.ItemData(cboCuenta.ListIndex) & "', Null, " & mTotalNeto & ", " & mTotalLiq & ", " & mTotalRetenido _
       & ", " & vAC_Fecha & ", '" & txtAc_Boleta.Text & "', '" & glogon.Maquina & "', '" & glogon.AppVersion & "', " & vAC_TipoBoleta

Call OpenRecordSet(rs, strSQL)
vCodigo = rs!RenunciaId
txtCodigo.Text = vCodigo

'Registra Detalle de Sub Motivos
strSQL = ""
With lswMotivos.ListItems
For i = 1 To .Count
  If .Item(i).Checked Then
     strSQL = strSQL & Space(10) & "exec spAFI_CR_Motivos_Registra " & txtCodigo.Text & ",'" & .Item(i).Tag & "','A','" & glogon.Usuario & "'"
  End If
Next i
End With

If Len(strSQL) > 0 Then
         Call ConectionExecute(strSQL)
End If


'Registrar Planes de Ahorros
strSQL = ""
With lswPlanes.ListItems
  For i = 1 To .Count
        strSQL = strSQL & Space(10) & "insert into AFI_CR_RENUNCIAS_PLANES(COD_RENUNCIA, COD_CONTRATO, COD_OPERADORA, COD_PLAN, DISPONIBLE, MULTA" _
               & ", REND_PENDIENTE, LIQ_FND, APORTES, RENDIMIENTOS, COD_DIVISA, TIPO_CAMBIO, MARCADA)" _
               & " values(" & vCodigo & "," & CLng(.Item(i).Text) & "," & CLng(.Item(i).Tag) & ",'" & .Item(i).SubItems(1) _
               & "'," & CCur(.Item(i).SubItems(2)) & "," & CCur(.Item(i).SubItems(6)) & "," & CCur(.Item(i).SubItems(5)) _
               & ",0," & CCur(.Item(i).SubItems(3)) & "," & CCur(.Item(i).SubItems(4)) & ",'" & Trim(.Item(i).SubItems(12)) _
               & "'," & CCur(.Item(i).SubItems(13)) & ", " & IIf(.Item(i).Checked, 1, 0) & ")"
  Next i
End With

If Len(strSQL) > 0 Then
    'Registra Lista a Procesar
    Call ConectionExecute(strSQL)
End If


'Registrar Abonos
strSQL = ""
With lswAbonos.ListItems
  For i = 1 To .Count
  
    strSQL = strSQL & Space(10) & "insert AFI_CR_RENUNCIAS_ABONOS(COD_RENUNCIA, ID_SOLICITUD, CODIGO, ABONO" _
           & ", SALDO, CARGOS, MORA_INTC, MORA_INTM, MORA_PRIN, COD_DIVISA, TIPO_CAMBIO, TIPO, GARANTIA, MARCADO)" _
           & " values(" & vCodigo & ", " & .Item(i).Text & ", '" & .Item(i).SubItems(1) & "', " & CCur(.Item(i).SubItems(11)) _
           & ", " & CCur(.Item(i).SubItems(5)) & ", " & CCur(.Item(i).SubItems(9)) + CCur(.Item(i).SubItems(10)) _
           & ", " & CCur(.Item(i).SubItems(6)) & ", " & CCur(.Item(i).SubItems(7)) & ", " & CCur(.Item(i).SubItems(8)) _
           & ", '" & Trim(.Item(i).SubItems(12)) & "', " & CCur(.Item(i).SubItems(13)) _
           & ", '" & .Item(i).SubItems(3) & "', '" & .Item(i).SubItems(4) & "', 1)"
  Next i

If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
End If

End With


'Finaliza
tcRenuncia.Item(1).Enabled = True

MsgBox "Información Registrada Satisfactoriamente", vbInformation

Call sbRenunciaBoleta(vCodigo)

If chkReingreso.Value = xtpChecked Then
    Call sbBoleta_Afiliacion
End If

Call sbConsulta(vCodigo)
  

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub cmdNuevo_Click()
    bConsulta = False
    Call sbInicializa
End Sub

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 1


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lswMotivos.ColumnHeaders
    .Clear
    .Add , , "Descripción", 3500
End With

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Fecha", 1800
    .Add , , "Usuario", 1800
    .Add , , "Estado", 1400
    .Add , , "Notas", 3270
End With


With lswRenta.ColumnHeaders
    .Clear
    .Add , , "Rng. Inicio", 1800, vbRightJustify
    .Add , , "Rng. Corte", 1800, vbRightJustify
    .Add , , "Renta", 1800, vbRightJustify
    .Add , , " [ % ]", 1000, vbRightJustify

End With

With lswHistorico.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Causa", 3000
    .Add , , "Identificación", 1400
    .Add , , "Tipo", 1100, vbCenter
    .Add , , "Estado", 1200, vbCenter
    .Add , , "Vencimiento", 1800, vbCenter
    .Add , , "Notas", 3000
    .Add , , "Ejecutivo", 1100
    .Add , , "Reingreso?", 1100, vbCenter
    .Add , , "Ejecutivo.Id", 10
    .Add , , "Nombre", 3000
End With



 strSQL = "select rtrim(cod_gestion) as 'IdX' , rtrim(descripcion) as 'ItmX' from afi_cr_gestiones"
 Call sbCbo_Llena_New(cboGestion, strSQL, False, True)
 
  cboResolucion.AddItem "Transito"
  cboResolucion.AddItem "Rescatada"
  cboResolucion.AddItem "Perdida"
  cboResolucion.Text = "Transito"

tcRenuncia.Item(0).Selected = True

tcMain.Item(0).Selected = True

With lswPlanes.ColumnHeaders
   .Clear
   .Add , , "No. Contrato", 1440
   .Add , , "Plan", 1140, vbCenter
   .Add , , "Disponible", 1440, vbRightJustify
   .Add , , "Aportes", 1440, vbRightJustify
   .Add , , "Rendimientos", 1440, vbRightJustify
   .Add , , "Rend.Pend.", 1440, vbRightJustify
   .Add , , "(-) Multas", 1440, vbRightJustify
   .Add , , "Operadora", 1100, vbCenter
   .Add , , "Plan Desc.", 3440
   .Add , , "ISR Monto", 2440, vbRightJustify
   .Add , , "ISR Porc", 1440, vbRightJustify
   .Add , , "ISR Apl?", 1440, vbCenter
   .Add , , "Divisa", 1440, vbCenter
   .Add , , "T.C.", 1440, vbRightJustify
End With

With lswAbonos.ColumnHeaders
   .Add , , "Operación", 1200
   .Add , , "Código", 1000, vbCenter
   .Add , , "Descripción", 2000
   .Add , , "Tipo", 1000, vbCenter
   .Add , , "Garantia", 1040, vbCenter
   .Add , , "Saldo", 1440, vbRightJustify
   .Add , , "Int.Cor.", 1140, vbRightJustify
   .Add , , "Int.Mor.", 1140, vbRightJustify
   .Add , , "Principal", 1240, vbRightJustify
   .Add , , "Cargos", 1140, vbRightJustify
   .Add , , "Pólizas", 1140, vbRightJustify
   .Add , , "Abono", 1440, vbRightJustify
   .Add , , "Divisa", 1440, vbCenter
   .Add , , "T.C.", 1440, vbRightJustify
End With


vPaso = True
    'Carga Cuentas Bancarias Autorizadas
    strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
    Call sbCbo_Llena_New(cboBanco, strSQL, False, True)
vPaso = False




'Carga Tipos de Acciones (Documentos)
strSQL = "select Id_Documento as 'IdX', Descripcion as 'ItmX' from AFI_CR_RENUNCIAS_TIPO_DOCUMENTO"
Call sbCbo_Llena_New(cboAc_Tipo, strSQL, False, True)

'Activa Opcion de Arreglos de Pago a Usuarios Autorizados
strSQL = "select dbo.fxAFI_Renuncia_Arreglo_Pago('" & glogon.Usuario & "') as 'Acceso'"
Call OpenRecordSet(rs, strSQL)
    chkArregloPago.Value = rs!Acceso
    If rs!Acceso = 0 Then
        chkArregloPago.Enabled = False
    End If
rs.Close


'Modo de Sistema (ASE o SIF)
'vModoSIF = fxModoSIF
vModoSIF = True

mFechaSistema = fxFechaServidor

dtpAc_fecha.Value = Format(mFechaSistema, "dd/mm/yyyy")
dtpPago.Value = dtpAc_fecha.Value


Call Formularios(Me)
'Call sbLimpiaDatos
Call RefrescaTags(Me)


Call cmdNuevo_Click

bConsulta = True


End Sub


Private Sub sbControl_Consulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, strEstado As String


Me.MousePointer = vbHourglass
    
  txtCasoId.Text = CStr(vCodigo)
    
  lsw.ListItems.Clear
  
  strSQL = "select * from afi_cr_renuncias where cod_renuncia = " & vCodigo & ""
  Call OpenRecordSet(rs, strSQL)
  
  gbResolucion.Enabled = False
  
If rs.RecordCount > 0 Then
     
     txtRegUser.Text = IIf(Not IsNull(rs!registro_user), rs!registro_user, "")
     txtRegFecha.Text = IIf(Not IsNull(rs!Registro_Fecha), Format(rs!Registro_Fecha, "mm/dd/yyyy"), "")
     txtResFecha.Text = IIf(Not IsNull(rs!resuelto_fecha), Format(rs!resuelto_fecha, "mm/dd/yyyy"), "")
     txtResUser.Text = IIf(Not IsNull(rs!resuelto_user), rs!resuelto_user, "")
     Select Case rs!Estado
         Case "T"
           txtEstadoControl.Text = "Transito"
           gbResolucion.Enabled = True
         Case "P"
           txtEstadoControl.Text = "Perdida"
        Case "R"
           txtEstadoControl.Text = "Rescatada"
        Case "V"
           txtEstadoControl.Text = "Vencida"
     End Select
     
     txtVencimiento.Text = IIf(Not IsNull(rs!Vencimiento), Format(rs!Vencimiento, "mm/dd/yyyy"), "")
     
     cboResolucion.Text = txtEstadoControl.Text
     
End If
rs.Close

'Lista
strSQL = "select * from afi_cr_seguimiento where cod_renuncia = " & vCodigo & ""
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!Id)
   
   itmX.SubItems(1) = IIf(Not IsNull(rs!fecha), Format(rs!fecha, "mm/dd/yyyy"), "")
   itmX.SubItems(2) = IIf(Not IsNull(rs!Usuario), rs!Usuario, "")
   itmX.SubItems(4) = IIf(Not IsNull(rs!Notas), rs!Notas, "")
   Select Case rs!Estado
         Case "T"
           itmX.SubItems(3) = "Transito"
         Case "P"
           itmX.SubItems(3) = "Perdida"
        Case "R"
           itmX.SubItems(3) = "Rescatada"
        Case "V"
           itmX.SubItems(3) = "Vencida"
   End Select
   
    rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub lswHistorico_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

Call sbConsulta(Item.Text)

End Sub


Private Sub lswMotivos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub
If txtCodigo.Text = "" Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spAFI_CR_Motivos_Registra " & txtCodigo.Text & ",'" & Item.Tag & "','" _
       & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
  
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub tcRenuncia_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 1 'Control
        Call sbControl_Consulta
    Case 2 'Historico
        Call sbHistorico_Consulta
End Select

End Sub

Private Sub txtAc_Boleta_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TempCedula As String

If KeyCode = vbKeyReturn Then
  TempCedula = txtCedula.Text
  cmdNuevo_Click
  txtCedula.Text = TempCedula
  txtNombre.Text = fxNombre(txtCedula.Text)
  Call sbCargaDatos
'  cboCausa.SetFocus
End If

If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Identificación"
   gBusquedas.Col2Name = "Nombre"
   gBusquedas.Col3Name = "Id Alterno"

   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Consulta = "select Cedula,Nombre,CedulaR from socios"
   gBusquedas.Convertir = "N"
   gBusquedas.Filtro = " and estadoactual in('S','A')"
   frmBusquedas.Show vbModal
    
   txtCedula.Text = Trim(gBusquedas.Resultado)
   txtNombre.Text = gBusquedas.Resultado2
   
   Call sbCargaDatos
End If


End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And IsNumeric(txtCodigo.Text) Then
  Call sbConsulta(txtCodigo.Text)
End If
End Sub

Private Sub sbInicializa()

vCodigo = Empty

tcRenuncia.Item(0).Selected = True
tcMain.Item(0).Selected = True

txtCodigo.Text = ""
txtVencimiento.Text = ""
txtEstado.Text = ""

txtCedula.Text = ""
txtNombre.Text = ""

txtPromotorCod.Text = ""
txtPromotorDesc.Text = ""

lswMotivos.ListItems.Clear

chkReingreso.Value = xtpUnchecked
chkVolver.Value = xtpUnchecked
chkMortalidad.Value = xtpUnchecked
chkTasaAjuste.Value = xtpChecked

cmdGuardar.Enabled = True

txtNotas.Text = ""

tcRenuncia.Item(1).Enabled = False

If cboCausa.Locked Then cboCausa.Locked = False

Call sbConsulta_Motivos(0)

cboCuenta.Clear

On Error Resume Next
If txtCedula.Enabled Then
    txtCedula.SetFocus
End If

Call RefrescaTags(Me)

End Sub


Private Function fxRenuncia() As Long
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select isnull(Max(cod_renuncia),0) as consecutivo from afi_cr_renuncias"
        
Call OpenRecordSet(rs, strSQL)
  fxRenuncia = rs!Consecutivo + 1
rs.Close

End Function


Private Sub sbUniProgra()
Dim strSQL As String, rs As New ADODB.Recordset

If GLOBALES.SysASEVersion Then
    strSQL = "select UP,CT from socios where cedula = '" & txtCedula & "'"
Else
    strSQL = "select cod_departamento as 'UP',cod_departamento as 'CT' from socios where cedula = '" & txtCedula & "'"
End If

Call OpenRecordSet(rs, strSQL)

If rs.RecordCount > 0 Then
    vCt = rs!CT
    vUp = rs!UP
End If

End Sub


Private Function fxVencimiento() As Date
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select *,dbo.MyGetdate() as 'FechaServidor' from afi_cr_parametros"

Call OpenRecordSet(rs, strSQL)

If rs!tipo_vencimiento = "F" Then
    fxVencimiento = rs!fecha_limite
Else
    fxVencimiento = DateAdd("d", rs!dias_vence, rs!FechaServidor)
End If

End Function

Private Sub sbHistorico_Consulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError
 
Me.MousePointer = vbHourglass

strSQL = "Select R.*,rTrim(C.Descripcion) as 'CausaX',S.nombre" _
       & ",isnull(P.id_promotor,0) as 'Id_Promotor',isnull(P.nombre,'AFILIACION UNIVERSAL') as PromotorX" _
       & " from afi_cr_renuncias R inner join causas_renuncias C on R.id_causa = C.id_causa" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " left join Promotores P on R.id_Promotor = P.id_Promotor" _
       & " where R.cedula = '" & txtCedula.Text & "'"


Call OpenRecordSet(rs, strSQL, 0)

lswHistorico.ListItems.Clear

Do While Not rs.EOF
 
 Set itmX = lswHistorico.ListItems.Add(, , CStr(rs!Cod_Renuncia))
     itmX.SubItems(1) = rs!CausaX
     itmX.SubItems(2) = rs!Cedula
  
  Select Case rs!Tipo
    Case "P"
       itmX.SubItems(3) = "PATRONAL"
    Case "A"
       itmX.SubItems(3) = "ASOCIACION"
  End Select
  
  Select Case rs!Estado
   Case "T"
    itmX.SubItems(4) = "Transito"
   Case "R"
    itmX.SubItems(4) = "Rescatada"
   Case "P"
    itmX.SubItems(4) = "Perdida"
   Case "V"
    itmX.SubItems(4) = "Vencida"
  End Select
  
  itmX.SubItems(5) = rs!Vencimiento
  itmX.SubItems(6) = rs!Notas
  itmX.SubItems(7) = rs!PromotorX
  itmX.SubItems(8) = rs!Aplica_Reingreso
  itmX.SubItems(9) = rs!ID_PROMOTOR
  itmX.SubItems(10) = rs!Nombre
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Public Sub sbConsulta_Externa(pRenuncia As Long)

Call sbConsulta(pRenuncia)

End Sub

Public Sub sbConsulta_Externa_Cedula(pCedula As String)


txtCedula.Text = pCedula
Call txtCedula_KeyDown(vbKeyReturn, 0)

End Sub



Private Sub sbConsulta(pRenuncia As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError
 
Me.MousePointer = vbHourglass

'strSQL = "Select R.*,rTrim(C.Descripcion) as 'CausaX',S.nombre" _
'       & ",isnull(P.id_promotor,0) as 'Id_Promotor',isnull(P.nombre,'AFILIACION UNIVERSAL') as PromotorX" _
'       & " from afi_cr_renuncias R inner join causas_renuncias C on R.id_causa = C.id_causa" _
'       & " inner join Socios S on R.cedula = S.cedula" _
'       & " left join Promotores P on R.id_Promotor = P.id_Promotor" _
'       & " where R.cod_renuncia = " & pRenuncia


strSQL = "select *, dbo.fxSys_Cuentas_Mask(cuenta) as 'Cuenta_Desc' from vAFI_Renuncias Where cod_renuncia = " & pRenuncia

Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
 
 tcRenuncia.Item(0).Selected = True
 
 vCodigo = rs!Cod_Renuncia
 txtCodigo.Text = rs!Cod_Renuncia
 txtNombre.Text = rs!Nombre
 txtCedula.Text = rs!Cedula
 txtVencimiento.Text = rs!Vencimiento
 
 cboTipo.Clear
 Select Case rs!EstadoActual
    Case "S"
       cboTipo.AddItem "01 - Ren.Asociación"
       cboTipo.AddItem "02 - Ren.Patronal"
    Case "A"
       cboTipo.AddItem "02 - Ren.Patronal"
 End Select
 
 Select Case rs!Tipo
    Case "P"
       cboTipo.AddItem "02 - Ren.Patronal"
       cboTipo.Text = "02 - Ren.Patronal"
    Case "A"
       cboTipo.AddItem "01 - Ren.Asociación"
       cboTipo.Text = "01 - Ren.Asociación"
  End Select
  
  
 Call sbCboAsignaDato(cboCausa, Trim(rs!Causa_Desc), True, rs!Id_Causa)
  
  
  Select Case rs!Estado
   Case "T"
    txtEstado.Text = "Transito"
   Case "R"
    txtEstado.Text = "Rescatada"
   Case "P"
    txtEstado.Text = "Perdida"
   Case "V"
    txtEstado.Text = "Vencida"
  End Select
  
  txtPromotorDesc.Text = rs!Ejecutivo_Desc
  txtPromotorCod.Text = rs!ID_PROMOTOR
  
  chkReingreso.Value = rs!Aplica_Reingreso
  chkMortalidad.Value = rs!Mortalidad
  
  chkVolver.Value = rs!Volver
  chkTasaAjuste.Value = rs!Aumenta_Puntos
  txtNotas.Text = rs!Notas
  
  
 Call sbCboAsignaDato(cboBanco, Trim(rs!Banco_Desc), True, rs!Id_Banco)
  
 Call sbCboAsignaDato(cboCuenta, Trim(rs!Cuenta_Desc), True, rs!Cuenta)
 Call sbCboAsignaDato(cboTipoDoc, Trim(rs!Tipo_Documento_Desc), True, rs!Tipo_documento)
  
  
 txtAc_Boleta.Text = rs!Boleta & ""
 dtpAc_fecha.Value = IIf(IsNull(rs!Ac_Fecha), Now, rs!Ac_Fecha)
  
  
   
  tcRenuncia.Item(1).Enabled = True
End If
rs.Close


'Consulta Motivos
Call sbConsulta_Motivos(pRenuncia)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbConsulta_Motivos(pRenuncia As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError
 
Me.MousePointer = vbHourglass


'Consulta Motivos
lswMotivos.ListItems.Clear
With lswMotivos.ColumnHeaders
    .Clear
    .Add , , "Descripción", 6200
End With

  
strSQL = "exec spAFI_CR_Motivos_Consulta " & pRenuncia & ",1"
Call OpenRecordSet(rs, strSQL)

vPaso = True
With lswMotivos.ListItems
   .Clear
   Do While Not rs.EOF
    Set itmX = .Add(, , rs!Descripcion)
        itmX.Tag = rs!Cod_Motivo
        
        itmX.Checked = IIf((rs!asignado = 1), True, False)
        
    rs.MoveNext
   Loop
   rs.Close
End With
vPaso = False


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub sbRenunciaBoleta(vCodRenuncia As Long)
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass


 With frmContenedor.Crt
     .Reset

     .WindowShowPrintSetupBtn = True
     .WindowShowRefreshBtn = True
     .WindowShowSearchBtn = True
     .WindowState = crptMaximized
     .WindowTitle = "Reportes del Módulo de Personas"
     .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
     .Formulas(5) = "fxCodigoBarras = '*" & txtCodigo.Text & "*'"
     
     .Connect = glogon.ConectRPT
     
     .ReportFileName = SIFGlobal.fxPathReportes("Personas_CrBoletaRenuncias.rpt")
      strSQL = "{vAFI_Renuncia_Boleta.cod_renuncia} = " & vCodRenuncia
     
     .SelectionFormula = strSQL
     .PrintReport
 End With
 
Me.MousePointer = vbDefault
 
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub




Private Function fxPromotor(pPromotor As Long) As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As String

Me.MousePointer = vbHourglass

strSQL = "select nombre from promotores where id_promotor = " & pPromotor & ""
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    vResultado = rs!Nombre
Else
    vResultado = ""
End If
rs.Close

Me.MousePointer = vbDefault

fxPromotor = vResultado

End Function


Private Sub txtPromotorCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPromotorDesc.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "S"
  gBusquedas.Columna = "id_Promotor"
  gBusquedas.Orden = "id_Promotor"
  gBusquedas.Consulta = "select id_Promotor,Nombre from Promotores"
  gBusquedas.Filtro = " and Estado = 1 and tipo <> 'C'"
  frmBusquedas.Show vbModal
  txtPromotorCod.Text = gBusquedas.Resultado
  txtPromotorDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtPromotorCod_LostFocus()

If IsNumeric(txtPromotorCod) Then
   txtPromotorDesc.Text = fxPromotor(txtPromotorCod.Text)
Else
   txtPromotorDesc.Text = ""
End If

End Sub

Private Sub txtPromotorDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkMortalidad.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "Nombre"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Consulta = "select id_Promotor,Nombre from Promotores"
  gBusquedas.Filtro = " and Estado = 1 and tipo <> 'C'"
  frmBusquedas.Show vbModal
  txtPromotorCod.Text = gBusquedas.Resultado
  txtPromotorDesc.Text = gBusquedas.Resultado2
End If
End Sub


'-----------------------------------------------------------------------------------------------------------
' Codigo de Liquidacion



Private Function fxVerificaDatos() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

vMensaje = ""

If txtCedula = "" Or txtNombre = "" Or txtPromotorCod.Text = "" Or Not IsNumeric(txtPromotorCod.Text) Then
   vMensaje = vMensaje & vbCrLf & " - Datos Erroneos, verifique!..."
End If

'Valida que no exista otra renuncia en transito
If txtCodigo.Text = "" Then
     strSQL = "select dbo.fxAFI_Renuncia_Activa('" & txtCedula.Text & "') as 'Resultado'"
Else
    If IsNumeric(txtCodigo.Text) Then
        strSQL = "select dbo.fxAFI_Renuncia_Activa_Otra('" & txtCedula.Text & "', " & txtCodigo.Text & ") as 'Resultado'"
    Else
        strSQL = "select dbo.fxAFI_Renuncia_Activa_Otra('" & txtCedula.Text & "', 0) as 'Resultado'"
    End If
End If

Call OpenRecordSet(rs, strSQL)
If rs!Resultado = 1 Then
   vMensaje = vMensaje & vbCrLf & " - Esta Persona ya tiene una renuncia en tramite o la actual ya fue liquidada, verifique!..."
End If


strSQL = "select isnull(count(*),0) as Existe from socios where cedula = '" & txtCedula & "'"
Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then vMensaje = vMensaje & " - La persona especificada no existe Registrada..." & vbCrLf
rs.Close

If cboTipoDoc.ItemData(cboTipoDoc.ListIndex) = "TE" And (cboCuenta.ListCount = 0 Or cboCuenta.Text = "") Then
    vMensaje = vMensaje & " - No se ha indicado una cuenta bancaria para realizar la transferencia a la persona..." & vbCrLf

End If

If cboCausa.ListCount = 0 Then
    vMensaje = vMensaje & " - No se ha indicado una causa de reuncia..." & vbCrLf
End If

If Mid(cboTipo, 1, 2) = "03" Or Mid(cboTipo, 1, 2) = "" Then vMensaje = vMensaje & " - El proceso siguiente no aplica..." & vbCrLf


If lblAc_Boleta.Visible Then
  If Len(txtAc_Boleta.Text) = 0 Then
    vMensaje = vMensaje & " - Especifique el Número de Boleta de Accion de Personal..." & vbCrLf
  End If
End If

If txtCasoId.Text <> "" Then
   If Mid(txtEstado.Text, 1, 1) = "P" Then
    vMensaje = vMensaje & " - Esta renuncia se encuntra Perdida, no puede ser modificada!" & vbCrLf
   End If
End If

If Len(vMensaje) > 0 Then
 MsgBox vMensaje, vbCritical
 fxVerificaDatos = False
Else
 fxVerificaDatos = True
End If

End Function


Private Sub sbAportesTotales()
Dim curTotalBruto As Currency, curRenta As Currency

curTotalBruto = 0

If chkAplObrero.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblAporteObrero.Caption)
If chkAplPatronal.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblAportePatronal.Caption)
If chkAplPatronal.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblCustodia.Caption)

If chkAplPatronal.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblFCI.Caption)
If chkAplCapGen.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblCapitalizacion.Caption)
If chkAplCapExtra.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblAporteExtra.Caption)
If chkAplExcedente.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblExcedente.Caption)


lblTotalBruto.Caption = Format(curTotalBruto, "Standard")

curRenta = CCur(lblRenta.Caption)

If chkAplExcedente.Value = xtpChecked Then
    curRenta = curRenta + CCur(lblExcedenteRenta.Caption)
End If

txtRetenerMonto = Format(curRenta, "Standard")

If IsNumeric(txtRetenerMonto) Then
  lblTotalNeto(0).Caption = Format(curTotalBruto - CCur(txtRetenerMonto), "Standard")
Else
  lblTotalNeto(0).Caption = Format(curTotalBruto, "Standard")
End If

End Sub



Private Sub btnAnterior_Click()
Dim i As Integer

If tcMain.SelectedItem > 0 Then
 tcMain.Item(tcMain.SelectedItem - 1).Selected = True
 
 For i = 0 To tcMain.ItemCount - 1
   tcMain.Item(i).Enabled = False
 Next i
 'Preguntar si desea limpiar los datos
 i = MsgBox("Desea Limpiar Los Datos Anteriores...", vbYesNo)
 If i = vbYes Then
   Call sbLimpiaDatos
   Call sbCargaDatos
 End If
End If

tcMain.Item(tcMain.SelectedItem).Enabled = True
        
End Sub


Private Sub btnSiguiente_Click()
Dim i As Integer
           

If tcMain.SelectedItem = 0 Then
    If fxVerificaDatos Then
       tcMain.Item(tcMain.SelectedItem + 1).Selected = True
       Call sbLimpiaDatos
       Call sbCargaDatos
    End If
Else
    If tcMain.SelectedItem < 4 Then
       tcMain.Item(tcMain.SelectedItem + 1).Selected = True
      Call sbLimpiaDatos
      Call sbCargaDatos
    End If
End If

End Sub

Private Sub cboBanco_Click()

If vPaso Or cboBanco.ListCount = 0 Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & txtCedula.Text & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)


strSQL = "exec spAFI_Renuncia_Emite_TDoc " & cboBanco.ItemData(cboBanco.ListIndex) & ", " & chkMortalidad.Value
Call sbCbo_Llena_New(cboTipoDoc, strSQL, False, True)

vError:

End Sub

Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCuenta.SetFocus
End Sub


Private Sub chkAplCapExtra_Click()
Call sbAportesTotales
End Sub

Private Sub chkAplCapGen_Click()
Call sbAportesTotales
End Sub

Private Sub chkAplObrero_Click()
Call sbAportesTotales
End Sub

Private Sub chkAplPatronal_Click()
Call sbAportesTotales
End Sub




Private Sub cmdDistribucionAuto_Click()
'Dim curDisponible As Currency, i As Integer
'Dim curAbono As Currency

Me.MousePointer = vbHourglass

Call sbCargaDatos
'
''Inicializa
'curDisponible = lblTotalNeto(1).Caption
'
'For i = 1 To lswAbonos.ListItems.Count
'  lswAbonos.ListItems(i).SubItems(11) = 0
'Next i
'
'
''Distruye Primero a Morosidad
'With lswAbonos.ListItems
' For i = 1 To .Count
'
'   curAbono = 0
'
'   'Polizas
'   If curDisponible > CCur(.Item(i).SubItems(10)) Then
'     curAbono = curAbono + CCur(.Item(i).SubItems(10))
'     curDisponible = curDisponible - CCur(.Item(i).SubItems(10))
'   Else
'     curAbono = curAbono + curDisponible
'     curDisponible = 0
'   End If
'
'   'Cargos
'   If curDisponible > CCur(.Item(i).SubItems(9)) Then
'     curAbono = curAbono + CCur(.Item(i).SubItems(9))
'     curDisponible = curDisponible - CCur(.Item(i).SubItems(9))
'   Else
'     curAbono = curAbono + curDisponible
'     curDisponible = 0
'   End If
'
'
'   'MoraIntCor
'   If curDisponible > CCur(.Item(i).SubItems(6)) Then
'     curAbono = curAbono + CCur(.Item(i).SubItems(6))
'     curDisponible = curDisponible - CCur(.Item(i).SubItems(6))
'   Else
'     curAbono = curAbono + curDisponible
'     curDisponible = 0
'   End If
'
'   'MoraIntMor
'   If curDisponible > CCur(.Item(i).SubItems(7)) Then
'     curAbono = curAbono + CCur(.Item(i).SubItems(7))
'     curDisponible = curDisponible - CCur(.Item(i).SubItems(7))
'   Else
'     curAbono = curAbono + curDisponible
'     curDisponible = 0
'   End If
'
'
'   'Principal atrasado
'    If curDisponible > CCur(.Item(i).SubItems(8)) Then
'      curAbono = curAbono + CCur(.Item(i).SubItems(8))
'      curDisponible = curDisponible - CCur(.Item(i).SubItems(8))
'    Else
'      curAbono = curAbono + curDisponible
'      curDisponible = 0
'    End If
'
'   .Item(i).SubItems(11) = Format(curAbono, "Standard")
'
'   lblDisponible.Caption = Format(curDisponible, "Standard")
'
' Next i
'End With
'
'
'
'
'
'
''Distribuye el Restante
'With lswAbonos.ListItems
' For i = 1 To .Count
'
'   curAbono = 0
'
'
'   'Saldo o principal atrasado
'    If curDisponible > (CCur(.Item(i).SubItems(5)) - CCur(.Item(i).SubItems(8))) Then
'      curAbono = curAbono + (CCur(.Item(i).SubItems(5)) - CCur(.Item(i).SubItems(8)))
'      curDisponible = curDisponible - (CCur(.Item(i).SubItems(5)) - CCur(.Item(i).SubItems(8)))
'    Else
'      curAbono = curAbono + curDisponible
'      curDisponible = 0
'    End If
'
'   .Item(i).SubItems(11) = Format(curAbono + CCur(.Item(i).SubItems(11)), "Standard")
'
'   lblDisponible.Caption = Format(curDisponible, "Standard")
'
' Next i
'End With
'
Me.MousePointer = vbDefault
MsgBox "Distribución Automática Aplicada...", vbInformation

End Sub

Private Sub cmdMAceptar_Click()
Dim i As Integer

If CCur(txtMAbono) > CCur(lblMTotalDeuda.Caption) Then
  MsgBox "El monto del Abono es Mayor que el Total Adeudado...", vbExclamation
  Exit Sub
End If

If CCur(txtMAbono) > CCur(lblDisponible.Caption) Then
  MsgBox "El monto del Abono es Mayor que el Disponible de Aplicación...", vbExclamation
  Exit Sub
End If

'Pasar Dato del Abono
lswAbonos.SelectedItem.SubItems(11) = txtMAbono
lblDisponible.Caption = Format(CCur(lblDisponible.Caption) - CCur(txtMAbono), "Standard")

fraAbono.Visible = False
lswAbonos.Visible = True

End Sub

Private Sub cmdMCancelar_Click()
lblDisponible.Caption = Format(CCur(lblDisponible.Caption) - CCur(lswAbonos.SelectedItem.SubItems(11)), "Standard")

fraAbono.Visible = False
lswAbonos.Visible = True

End Sub


Private Sub sbLimpiaDatos()

tcMain.Item(0).Enabled = False
tcMain.Item(1).Enabled = False
tcMain.Item(2).Enabled = False
tcMain.Item(3).Enabled = False
tcMain.Item(4).Enabled = False


Select Case tcMain.SelectedItem
  Case 0 'Renuncia
    tcMain.Item(0).Enabled = True
    txtCedula = ""
    txtNombre = ""
    cboTipo.Clear
    cboTipoDoc.Text = "Cheque"
    cboCuenta.Clear
    
    lblEstadoActual.Caption = ""
    lblIngreso.Caption = ""
    lblBoleta.Caption = ""
    
    txtAc_Boleta.Text = ""
    dtpAc_fecha.Value = mFechaSistema
    dtpAc_fecha.Visible = False
    lblAcFecha.Caption = ""
    
     
  
  Case 1 'Aportes
    tcMain.Item(1).Enabled = True
    
    lblAporteObrero.Caption = 0
    lblAportePatronal.Caption = 0
    lblCustodia.Caption = 0
    lblFCI.Caption = 0
    lblCapitalizacion.Caption = 0
    lblAporteExtra.Caption = 0
    lblRenta.Caption = 0
    lblExcedenteRenta.Caption = 0
    lblExcedente.Caption = 0
    lblExcedente.Tag = 0
     
    lblTotalBruto.Caption = 0
    txtRetenerMonto = 0
    lblTotalNeto(0).Caption = 0
     
    chkAplObrero.Value = xtpUnchecked
    chkAplPatronal.Value = xtpUnchecked
    chkAplCapGen.Value = xtpUnchecked
    chkAplCapExtra.Value = xtpUnchecked
    chkAplExcedente.Value = xtpUnchecked
     
    chkAplObrero.Enabled = False
    chkAplPatronal.Enabled = False
    chkAplCapGen.Enabled = False
    chkAplCapExtra.Enabled = False
    chkAplExcedente.Enabled = False
     
  Case 2 'Planes de Ahorro
    tcMain.Item(2).Enabled = True
    
    lswPlanes.ListItems.Clear
    lblTotalNeto.Item(1).Caption = lblTotalNeto.Item(0).Caption
         
     
  Case 3 'Abonos
    tcMain.Item(4).Enabled = True
    
    lswAbonos.ListItems.Clear
'    lblLsw.Caption = ""
     
    lblDisponible.Caption = 0
     
    lblMOperacion.Caption = ""
    lblMCodigo.Caption = ""
    lblMTipo.Caption = ""
    lblMGarantia.Caption = ""
    lblMLineaDesc.Caption = ""
    lblMTotalDeuda.Caption = ""
     
    lblMSaldo.Caption = ""
    lblMMorPrincipal.Caption = ""
    lblMMorIntCor.Caption = ""
    lblMMoraIntMor.Caption = ""
    lblMCargos.Caption = ""
    lblMPolizas.Caption = ""
    txtMAbono = 0
    
  
  Case 4 'Sumario
    tcMain.Item(4).Enabled = True
    
    txtSumario = ""


End Select

End Sub

'Mantiene por Compatibilidad
Private Function fxFCI(vCedula As String) As Currency
 fxFCI = 0
End Function

Private Sub sbCargaDatos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, rsTmp As New ADODB.Recordset
Dim curTotalLiq As Currency, curTotalPrestamos As Currency
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case tcMain.SelectedItem
  Case 0 'Renuncia
    strSQL = "select S.cedula,S.nombre,S.fechaingreso,S.estadoactual,0 as Boleta,isnull(E.descripcion,'') as 'EstadoPersona'" _
           & ", dbo.fxAFI_Renuncia_Activa(S.Cedula) as 'Valida'" _
           & ", dbo.fxCBR_Cobro_Judicial_Indica(S.Cedula) as 'CbrJud'" _
           & " from socios S inner join AFI_ESTADOS_PERSONA E on S.estadoActual = E.cod_estado" _
           & " where S.cedula = '" & txtCedula & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
       txtCedula.Text = rs!Cedula
       txtNombre.Text = rs!Nombre & ""
       cboTipo.Clear
       
       If rs!CbrJud = 1 Then
        lblCbrStatus.Visible = True
       Else
        lblCbrStatus.Visible = False
       End If
       
       lblEstadoActual.Caption = rs!EstadoPersona
       lblEstadoActual.Tag = rs!EstadoActual
       Select Case UCase(rs!EstadoActual)
          Case "S"
            cboTipo.AddItem "01 - Ren.Asociación"
            cboTipo.AddItem "02 - Ren.Patronal"
            cboTipo.Text = "01 - Ren.Asociación"
          Case "A"
            cboTipo.AddItem "02 - Ren.Patronal"
            cboTipo.Text = "02 - Ren.Patronal"
          Case "P"
            cboTipo.AddItem "03 - No Aplica"
            cboTipo.Text = "03 - No Aplica"
          Case "N"
            cboTipo.AddItem "03 - No Aplica"
            cboTipo.Text = "03 - No Aplica"
       End Select
       
       lblIngreso.Caption = Format(IIf(IsNull(rs!FechaIngreso), Date, rs!FechaIngreso), "yyyy/mm/dd")
       lblBoleta.Caption = IIf(IsNull(rs!Boleta), 0, rs!Boleta)
    
        Me.MousePointer = vbDefault
        If rs!Valida = 1 Then
           MsgBox " - Esta Persona ya tiene una renuncia en tramite, verifique!...", vbExclamation
        End If
    
    
    Else
       MsgBox "No Se encontró ningun registro de la Persona, verifique...", vbInformation
    End If
    rs.Close
        
    Call cboBanco_Click
    Call cboTipo_Click
    
  Case 1 'Aportes
    
    strSQL = "exec spAFI_Liq_Consulta_Patrimonio '" & txtCedula & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF And Not rs.BOF Then
        lblAporteObrero.Caption = Format(rs!ahorro, "Standard")
        lblAportePatronal.Caption = Format(rs!Aporte, "Standard")
        
        lblCustodia.Caption = Format(rs!Custodia, "Standard")
        
        lblCapitalizacion.Caption = Format(rs!capitaliza, "Standard")
        lblAporteExtra.Caption = Format(rs!Extra, "Standard")
        
        
        lblRenta.Caption = Format(rs!Renta, "Standard")
        
        lblExcedente.Caption = Format(rs!Excedente, "Standard")
        lblExcedenteRenta.Caption = Format(rs!EXC_RENTA, "Standard")
        lblExcedente.Tag = rs!EXC_APLICA
        
        txtRetenerMonto = Format(rs!Renta, "Standard")
        
        txtDivisa.Text = rs!cod_Divisa
        
        txtDivisaLocal.Text = rs!divisa_local
        txtTipoCambio.Text = Format(rs!TIPO_CAMBIO, "########0.0000")
        
    End If
    rs.Close
    lblFCI.Caption = "0"

    chkAplPatronal.Value = xtpUnchecked
    chkAplExcedente.Value = xtpChecked
    
    
    'Detalla Renta
    strSQL = "exec spExc_Renta_Detallada " & CCur(lblCapitalizacion.Caption)
    Call OpenRecordSet(rs, strSQL)
    
    lswRenta.ListItems.Clear
    Do While Not rs.EOF
     Set itmX = lswRenta.ListItems.Add(, , Format(rs!Desde, "Standard"))
         itmX.SubItems(1) = Format(rs!Hasta, "Standard")
         itmX.SubItems(2) = Format(rs!Renta, "Standard")
         itmX.SubItems(3) = Format(rs!Porcentaje, "Standard")
     rs.MoveNext
    Loop
    rs.Close
    
    
    Select Case Mid(cboTipo, 1, 2)
      Case "01" 'Renuncia a la Asociación (Liquidacion Parcial)
        chkAplObrero.Enabled = True
        chkAplCapGen.Enabled = True
        chkAplCapExtra.Enabled = True
        
        chkAplPatronal.Enabled = False
        chkAplExcedente.Enabled = False
        
      Case "02" 'Renuncia Patronal (Liquidacion Total)
        chkAplObrero.Enabled = True
        chkAplPatronal.Enabled = True
        chkAplCapGen.Enabled = True
        chkAplCapExtra.Enabled = True
        
        chkAplExcedente.Enabled = False
        If chkAplExcedente.Tag = "1" Then
            chkAplExcedente.Enabled = True
        End If
        
    End Select
     
    'Aplica Marcas por Default
        chkAplObrero.Value = vbChecked
        chkAplCapGen.Value = vbChecked
        chkAplCapExtra.Value = vbChecked
        
        
     
    Call sbAportesTotales
     
    'Bloquea Opciones de Check
        chkAplObrero.Enabled = False
        chkAplPatronal.Enabled = False
        chkAplCapGen.Enabled = False
        chkAplCapExtra.Enabled = False
     
     '********** SE DESBLOQUEA ESTA OPCION PORQUE HAY UN COMITE SI DECIDE LA CAUSA
     'DE MUERTE, POR TANTO NO SE SABE Y SE DEJA A CRITERIO DEL USUARIO
     '**********
'    'Si la Causa es por Muerte no se le aplica nada a las deudas
'    If chkMortalidad.Value = vbChecked Then
'        chkAplObrero.Enabled = False
'        chkAplPatronal.Enabled = False
'        chkAplCapGen.Enabled = False
'        chkAplCapExtra.Enabled = False
'    End If
     
  Case 2 'Planes de Ahorros
  
     vPaso = True
     
    Select Case Mid(cboTipo, 1, 2)
      Case "01" 'Renuncia a la Asociación (Liquidacion Parcial)
         strSQL = "exec spAfiLiquidaListaPlanes '" & txtCedula.Text & "','A'"
      
      Case "02" 'Renuncia Patronal (Liquidacion Total)
         strSQL = "exec spAfiLiquidaListaPlanes '" & txtCedula.Text & "','P'"
    End Select
     
     
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       Set itmX = lswPlanes.ListItems.Add(, , rs!COD_CONTRATO)
           itmX.Tag = rs!COD_OPERADORA
           itmX.SubItems(1) = rs!COD_PLAN
           itmX.SubItems(2) = Format(rs!Aportes + rs!Rendimiento + rs!RendPendiente - rs!Multa, "Standard")
           itmX.SubItems(3) = Format(rs!Aportes, "Standard")
           itmX.SubItems(4) = Format(rs!Rendimiento, "Standard")
           itmX.SubItems(5) = Format(rs!RendPendiente, "Standard")
           itmX.SubItems(6) = Format(rs!Multa, "Standard")
           itmX.SubItems(7) = rs!operadoraX
           itmX.SubItems(8) = rs!PlanX
           itmX.SubItems(9) = 0
           itmX.SubItems(10) = 0
           itmX.SubItems(11) = IIf((rs!RENTA_GLOBAL = 1), "Sí", "No")
           itmX.SubItems(12) = rs!cod_Divisa
           itmX.SubItems(13) = rs!TIPO_CAMBIO
           
           
       rs.MoveNext
     Loop
     rs.Close
     vPaso = False
     
     txtFndRendGravado.Text = "0"
     txtFndRendLiquidar.Text = "0"
     lblTotalNeto.Item(2).Caption = "0" 'ISR Monto
     lblTotalNeto.Item(1).Caption = Format(CCur(lblTotalNeto.Item(0).Caption), "Standard")


    'Marcada por Default
    For i = 1 To lswPlanes.ListItems.Count
        lswPlanes.ListItems.Item(i).Checked = True
    Next i
    
  
  Case 3 'Abonos
    'op,cod,tipo,garantia,saldo,moraintc,moraintm,moraprin,abono
    lblDisponible.Caption = lblTotalNeto(1).Caption
    lswAbonos.ListItems.Clear
    
    Dim curAbono As Currency
    
    curAbono = 0
    
    strSQL = "exec spAfi_Liquidacion_CreditosPersona '" & txtCedula.Text & "', " & CCur(lblDisponible.Caption)
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
      
         Set itmX = lswAbonos.ListItems.Add(, , rs!ID_SOLICITUD)
             itmX.SubItems(1) = rs!Codigo
             itmX.SubItems(2) = rs!Descripcion
             itmX.SubItems(3) = rs!Detalle
             itmX.SubItems(4) = rs!GarantiaX
             itmX.SubItems(5) = Format(rs!Saldo, "Standard")
             itmX.SubItems(6) = Format(rs!IntC, "Standard")
             itmX.SubItems(7) = Format(rs!IntM, "Standard")
             itmX.SubItems(8) = Format(rs!Amortiza, "Standard")
             itmX.SubItems(9) = Format(rs!Cargos, "Standard")
             itmX.SubItems(10) = Format(rs!Polizas, "Standard")
             itmX.SubItems(11) = Format(rs!Abono, "Standard")
             itmX.SubItems(12) = rs!cod_Divisa
             itmX.SubItems(13) = rs!TIPO_CAMBIO
      
            curAbono = curAbono + rs!Abono
      
      rs.MoveNext
    Loop
    rs.Close
    
    lblDisponible.Caption = Format(CCur(lblTotalNeto(1).Caption) - curAbono, "Standard")

    
'    'Aplica por Default Abono Auto
'    Call cmdDistribucionAuto_Click
    
'    chkArregloPago.Value = xtpChecked
    
    'Verifica SINPE Negativo
    
    strSQL = "exec spFnd_Sinpe_Negativo '" & txtCedula.Text & "'"
    Call OpenRecordSet(rs, strSQL, 0)
      txtSinpeNegativo.Text = Format(rs!Monto_Negativo, "Standard")
    rs.Close
    
    tcMain.Item(4).Enabled = False
    
  Case 4 'Sumario
    
    
'    If chkMortalidad.Value = vbChecked Then
'      txtNotas = "RENUNCIA POR MORTALIDAD, SE DEBE GIRAR MONTO ESPECIFICADO A BENEFICIARIOS" _
'                     & ", SE ADJUNTA BOLETA CON EL ASIENTO DE LA LIQUIDACION"
'    Else
'      txtNotas = ""
'    End If
    
    curTotalLiq = 0
    curTotalPrestamos = 0
    
    txtSumario = "-----------------------------------------------------" & vbCrLf
    txtSumario = txtSumario & "PROCESAR LA LIQUIDACION PARA: " & txtNombre & vbCrLf
    txtSumario = txtSumario & "-----------------------------------------------------" & vbCrLf
    txtSumario = txtSumario & "RENUNCIA >>>>>" & vbCrLf
    txtSumario = txtSumario & "TIPO: " & UCase(cboTipo.Text) & vbCrLf
    txtSumario = txtSumario & "CAUSA: " & cboCausa.Text & vbCrLf
    txtSumario = txtSumario & "MORTALIDAD: " & IIf(chkMortalidad.Value = vbChecked, "SI", "NO") & vbCrLf & vbCrLf
    txtSumario = txtSumario & "DEPOSITOS >>>>>" & vbCrLf
    txtSumario = txtSumario & "TIPO: " & UCase(cboTipoDoc.Text) & vbCrLf
    txtSumario = txtSumario & "BANCO: " & cboBanco.Text & vbCrLf
    txtSumario = txtSumario & "CUENTA: " & cboCuenta.Text & vbCrLf & vbCrLf
    
    txtSumario = txtSumario & "APLICACION DE APORTES >>>>>" & vbCrLf
    
    If chkAplObrero.Value = vbChecked Then
      txtSumario = txtSumario & "[x] APORTE OBRERO : " & lblAporteObrero.Caption & vbCrLf
    Else
      txtSumario = txtSumario & "[ ] APORTE OBRERO : " & lblAporteObrero.Caption & vbCrLf
    End If

    If chkAplPatronal.Value = vbChecked Then
      txtSumario = txtSumario & "[x] APORTE PATRONAL : " & lblAportePatronal.Caption & vbCrLf
      txtSumario = txtSumario & "[x] APORTE CUSTODIA : " & lblCustodia.Caption & vbCrLf
    Else
      txtSumario = txtSumario & "[ ] APORTE PATRONAL : " & lblAportePatronal.Caption & vbCrLf
      txtSumario = txtSumario & "[ ] APORTE CUSTODIA : " & lblCustodia.Caption & vbCrLf
    End If

    If chkAplCapGen.Value = vbChecked Then
      txtSumario = txtSumario & "[x] CAPITALIZACION : " & lblCapitalizacion.Caption & vbCrLf
    Else
      txtSumario = txtSumario & "[ ] CAPITALIZACION : " & lblCapitalizacion.Caption & vbCrLf
    End If


    If chkAplExcedente.Value = vbChecked Then
      txtSumario = txtSumario & "[x] EXCEDENTE PERIODO : " & lblExcedente.Caption & vbCrLf
    Else
      txtSumario = txtSumario & "[ ] EXCEDENTE PERIODO : " & lblExcedente.Caption & vbCrLf
    End If


    If chkAplCapExtra.Value = vbChecked Then
      txtSumario = txtSumario & "[x] AHORRO EXTRAORDINARIO : " & lblAporteExtra.Caption & vbCrLf
    Else
      txtSumario = txtSumario & "[ ] AHORRO EXTRAORDINARIO : " & lblAporteExtra.Caption & vbCrLf & vbCrLf
    End If
    
    Select Case Mid(cboTipo.Text, 1, 2)
      Case "01" 'Liq.Interna
        curTotalLiq = CCur(lblAporteExtra.Caption) + CCur(lblAporteObrero.Caption) _
                    + CCur(lblCapitalizacion.Caption)
      Case "02" 'Liq.Total
        curTotalLiq = CCur(lblAporteExtra.Caption) + CCur(lblAporteObrero.Caption) _
                    + CCur(lblCapitalizacion.Caption) _
                    + CCur(lblAportePatronal.Caption) + CCur(lblCustodia.Caption)
        
        If lblExcedente.Tag = "1" Then
          curTotalLiq = curTotalLiq + CCur(lblExcedente.Caption)
        End If
        
    End Select
    
    txtSumario = txtSumario & vbCrLf & "PLANES DE AHORRO A LIQUIDAR: " & vbCrLf

    With lswPlanes.ListItems
        For i = 1 To .Count
          If .Item(i).Checked Then
              curTotalLiq = curTotalLiq + .Item(i).SubItems(2)
              txtSumario = txtSumario & "CONTRATO: " & .Item(i) & " >> PLAN: " & .Item(i).SubItems(1) & " >> MONTO: " & .Item(i).SubItems(2) & vbCrLf
          End If
        Next i
    End With
    
    txtSumario = txtSumario & vbCrLf & "APLICACIONES A CREDITOS: " & vbCrLf
    
    Dim pSaldoRes As Currency, pAbono As Currency
    
    With lswAbonos.ListItems
        For i = 1 To .Count
          If .Item(i).SubItems(11) > 0 Then
              pAbono = CCur(.Item(i).SubItems(11))
              pSaldoRes = CCur(.Item(i).SubItems(5))
              
              pAbono = pAbono - (CCur(.Item(i).SubItems(6)) + CCur(.Item(i).SubItems(7)) + CCur(.Item(i).SubItems(9)) _
                    + CCur(.Item(i).SubItems(10)))
              
              If pAbono > 0 Then
                pSaldoRes = pSaldoRes - pAbono
              End If
              
              txtSumario = txtSumario & "OPERACION: " & .Item(i).Text & " >> CODIGO: " & .Item(i).SubItems(1) & " >> GARANTIA: " & .Item(i).SubItems(4) & vbCrLf _
                    & " >> TOTAL ABONO    : " & .Item(i).SubItems(11) & vbCrLf _
                    & " >> ABONO PRINCIPAL: " & Format(pAbono, "Standard") _
                    & " >> NUEVO SALDO: " & Format(pSaldoRes, "Standard") & vbCrLf & vbCrLf
          End If
        Next i
    End With
    
    
    curTotalLiq = curTotalLiq - CCur(txtRetenerMonto)
    curTotalPrestamos = CCur(lblTotalNeto(1).Caption) - CCur(lblDisponible.Caption)

    
    mTotalLiq = curTotalLiq
    mTotalRetenido = CCur(txtRetenerMonto.Text)
    mTotalNeto = curTotalLiq - curTotalPrestamos
    
    txtSumario = txtSumario & vbCrLf & vbCrLf & "TOTALES >>>>>" & vbCrLf
    txtSumario = txtSumario & "TOTAL A LIQUIDAR : " & Format(curTotalLiq, "Standard") & vbCrLf
    txtSumario = txtSumario & "TOTAL APLICADO A PRESTAMOS : " & Format(curTotalPrestamos, "Standard") & vbCrLf
    txtSumario = txtSumario & "TOTAL RETENIDO : " & txtRetenerMonto & vbCrLf
    txtSumario = txtSumario & "TOTAL A GIRAR : " & Format(curTotalLiq - curTotalPrestamos, "Standard") & vbCrLf
    
    
End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Function fxModoSIF() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim Resultado As Boolean

On Error GoTo vError:

Resultado = True

'strSQL = "select Top 1 Fecha from renuncias"
'Call OpenRecordSet(rs, strSQL)
'    Resultado = False
'rs.Close

fxModoSIF = Resultado

Exit Function

vError:
 fxModoSIF = Resultado
 
End Function



Private Sub lswAbonos_DblClick()

'Solo los que tienen opcion de Arreglos de Pago, pueden hacer abonos manuales
If Not chkArregloPago.Enabled Then Exit Sub
If chkArregloPago.Value = xtpUnchecked Then Exit Sub

fraAbono.Left = lswAbonos.Left
fraAbono.top = lswAbonos.top
fraAbono.Visible = True
lswAbonos.Visible = False

With lswAbonos.SelectedItem
    lblMOperacion.Caption = .Text
    lblMCodigo.Caption = .SubItems(1)
    lblMLineaDesc.Caption = .SubItems(2)
    lblMTipo.Caption = .SubItems(3)
    lblMGarantia.Caption = .SubItems(4)
    
    lblMSaldo.Caption = .SubItems(5)
    lblMMorIntCor.Caption = .SubItems(6)
    lblMMoraIntMor.Caption = .SubItems(7)
    lblMMorPrincipal.Caption = .SubItems(8)
    lblMCargos.Caption = .SubItems(9)
    lblMPolizas.Caption = .SubItems(10)
    
    lblMTotalDeuda.Caption = Format(CCur(.SubItems(5)) + CCur(.SubItems(6)) + CCur(.SubItems(7)) + CCur(.SubItems(9)) + CCur(.SubItems(10)), "Standard")
    
    lblMMoraTotal.Caption = Format(CCur(.SubItems(8)) + CCur(.SubItems(6)) + CCur(.SubItems(7)) + CCur(.SubItems(9)) + CCur(.SubItems(10)), "Standard")
    
    txtMAbono = .SubItems(11)
    
    lblDisponible.Caption = Format(CCur(lblDisponible.Caption) + CCur(.SubItems(11)), "Standard")
End With


End Sub

Private Sub lswAbonos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim curAbono As Currency, curDeuda As Currency
Dim curMora As Currency

On Error GoTo vError

curAbono = CCur(lswAbonos.SelectedItem.SubItems(11))
curDeuda = CCur(lswAbonos.SelectedItem.SubItems(5)) + CCur(lswAbonos.SelectedItem.SubItems(6)) + CCur(lswAbonos.SelectedItem.SubItems(7) + CCur(lswAbonos.SelectedItem.SubItems(9)) + CCur(lswAbonos.SelectedItem.SubItems(10)))
curMora = CCur(lswAbonos.SelectedItem.SubItems(8)) + CCur(lswAbonos.SelectedItem.SubItems(6)) + CCur(lswAbonos.SelectedItem.SubItems(7) + CCur(lswAbonos.SelectedItem.SubItems(9)) + CCur(lswAbonos.SelectedItem.SubItems(10)))

lblLsw.Caption = "La Operación : " & lswAbonos.SelectedItem & vbCrLf
If curDeuda > curAbono Then
    If curMora > curAbono Then
      lblLsw.Caption = lblLsw.Caption & " -- Queda con Morosidad"
    Else
      lblLsw.Caption = lblLsw.Caption & " -- Queda al día"
    End If
Else
  lblLsw.Caption = lblLsw.Caption & " -- Queda Cancelada"
End If

Exit Sub

vError:
  lblLsw.Caption = "..."

End Sub



Private Sub lswPlanes_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

'Si No marca Reingreso, debe bloquear la opcion para que no desmarque
If chkReingreso.Value = xtpUnchecked Then
   Item.Checked = True
End If


Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass



lblTotalNeto.Item(1).Caption = CCur(lblTotalNeto.Item(0).Caption)
txtFndRendLiquidar.Text = "0"

With lswPlanes.ListItems
    For i = 1 To .Count
     
         If .Item(i).Checked Then
           lblTotalNeto.Item(1).Caption = Format(CCur(lblTotalNeto.Item(1).Caption) + CCur(.Item(i).SubItems(2)), "Standard")
           If Mid(.Item(i).SubItems(11), 1, 1) = "S" Then
               txtFndRendLiquidar.Text = CCur(txtFndRendLiquidar.Text) + CCur(.Item(i).SubItems(4)) + CCur(.Item(i).SubItems(5))
           End If
        End If
    
    
    Next i

End With

'Consulta Renta Global
   strSQL = "exec spFnd_Renta_Global '" & txtCedula & "', '" & Format(mFechaSistema, "yyyy/mm/dd hh:mm") _
       & "'," & CCur(txtFndRendLiquidar.Text)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   Item.SubItems(10) = Format(rs!RG_Porcentaje, "Standard")
  
   txtFndRendGravado.Text = Format(rs!Retiro_Gravable, "Standard")
   lblTotalNeto.Item(2).Caption = Format(rs!ISR_MONTO, "Standard")
End If
rs.Close
       

lblTotalNeto.Item(1).Caption = Format(CCur(lblTotalNeto.Item(1).Caption) - CCur(lblTotalNeto.Item(2).Caption), "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 
End Sub

