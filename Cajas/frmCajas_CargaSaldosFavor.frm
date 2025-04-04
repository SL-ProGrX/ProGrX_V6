VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCajas_CargaSaldosFavor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Administraci�n de Saldos a Favor"
   ClientHeight    =   9375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18465
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   18465
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   960
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmCajas_CargaSaldosFavor.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7815
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   17655
      _Version        =   1572864
      _ExtentX        =   31141
      _ExtentY        =   13785
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
      SelectedItem    =   1
      Item(0).Caption =   "Saldos en Cajas"
      Item(0).ControlCount=   27
      Item(0).Control(0)=   "chkFechas"
      Item(0).Control(1)=   "dtpRegistroInicio"
      Item(0).Control(2)=   "dtpRegistroCorte"
      Item(0).Control(3)=   "chkSF_Saldos"
      Item(0).Control(4)=   "chkSA_Marcas"
      Item(0).Control(5)=   "txtUsuario"
      Item(0).Control(6)=   "txtCedula"
      Item(0).Control(7)=   "txtNombre"
      Item(0).Control(8)=   "txtNumDoc"
      Item(0).Control(9)=   "cboTipoSaldo"
      Item(0).Control(10)=   "cboTipoLiquidacion"
      Item(0).Control(11)=   "Label2(18)"
      Item(0).Control(12)=   "Label2(5)"
      Item(0).Control(13)=   "Label2(4)"
      Item(0).Control(14)=   "Label2(3)"
      Item(0).Control(15)=   "Label2(2)"
      Item(0).Control(16)=   "Label2(1)"
      Item(0).Control(17)=   "Label2(0)"
      Item(0).Control(18)=   "vGrid"
      Item(0).Control(19)=   "chkI_Recursos"
      Item(0).Control(20)=   "Label2(10)"
      Item(0).Control(21)=   "Label2(11)"
      Item(0).Control(22)=   "cboSC_EntidadPagadora"
      Item(0).Control(23)=   "cboSC_OrigenRecursos"
      Item(0).Control(24)=   "txtSC_Mnt_Inicio"
      Item(0).Control(25)=   "txtSC_Mnt_Hasta"
      Item(0).Control(26)=   "Label2(12)"
      Item(1).Caption =   "Identificaci�n de Dep�sitos"
      Item(1).ControlCount=   14
      Item(1).Control(0)=   "vGridId"
      Item(1).Control(1)=   "dtpId_Inicio"
      Item(1).Control(2)=   "dtpId_Corte"
      Item(1).Control(3)=   "cboBanco"
      Item(1).Control(4)=   "txtId_NumDoc"
      Item(1).Control(5)=   "Label2(16)"
      Item(1).Control(6)=   "Label2(7)"
      Item(1).Control(7)=   "Label2(6)"
      Item(1).Control(8)=   "fraIdentifica"
      Item(1).Control(9)=   "FlatEdit1"
      Item(1).Control(10)=   "Label2(8)"
      Item(1).Control(11)=   "Label2(9)"
      Item(1).Control(12)=   "txtMnt_Inicio"
      Item(1).Control(13)=   "txtMnt_Hasta"
      Item(2).Caption =   "Carga Masiva de Casos Identificados"
      Item(2).ControlCount=   12
      Item(2).Control(0)=   "fraCargaIdentTotales"
      Item(2).Control(1)=   "vGridCarga"
      Item(2).Control(2)=   "cboCargaFormaPago"
      Item(2).Control(3)=   "cboCargaCuentaBanco"
      Item(2).Control(4)=   "chkCargaDepositos"
      Item(2).Control(5)=   "txtArchivo"
      Item(2).Control(6)=   "Label1(2)"
      Item(2).Control(7)=   "Label2(19)"
      Item(2).Control(8)=   "Label1(1)"
      Item(2).Control(9)=   "btnArchivo(0)"
      Item(2).Control(10)=   "btnArchivo(1)"
      Item(2).Control(11)=   "btnArchivo(2)"
      Begin XtremeSuiteControls.GroupBox fraIdentifica 
         Height          =   6735
         Left            =   2040
         TabIndex        =   55
         Top             =   1080
         Width           =   8415
         _Version        =   1572864
         _ExtentX        =   14843
         _ExtentY        =   11880
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Begin XtremeSuiteControls.ListView lsw 
            Height          =   2295
            Left            =   0
            TabIndex        =   98
            Top             =   480
            Width           =   8415
            _Version        =   1572864
            _ExtentX        =   14843
            _ExtentY        =   4048
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
            Appearance      =   21
            ShowBorder      =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtId_NSolicitud 
            Height          =   315
            Left            =   1560
            TabIndex        =   63
            Top             =   960
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   556
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtId_NumDocId 
            Height          =   315
            Left            =   5880
            TabIndex        =   64
            Top             =   960
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   556
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtId_Fecha 
            Height          =   315
            Left            =   5880
            TabIndex        =   67
            Top             =   2280
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   556
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtId_Cedula 
            Height          =   315
            Left            =   1560
            TabIndex        =   73
            Top             =   3720
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
         Begin XtremeSuiteControls.FlatEdit txtId_Nombre 
            Height          =   315
            Left            =   1560
            TabIndex        =   72
            Top             =   4080
            Width           =   6495
            _Version        =   1572864
            _ExtentX        =   11456
            _ExtentY        =   556
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
            BackColor       =   16777215
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtId_Banco 
            Height          =   315
            Left            =   1560
            TabIndex        =   65
            Top             =   1440
            Width           =   6495
            _Version        =   1572864
            _ExtentX        =   11456
            _ExtentY        =   556
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
            BackColor       =   16777215
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtId_Descripcion 
            Height          =   435
            Left            =   1560
            TabIndex        =   68
            Top             =   1800
            Width           =   6495
            _Version        =   1572864
            _ExtentX        =   11456
            _ExtentY        =   767
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
            BackColor       =   16777215
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtId_Monto 
            Height          =   315
            Left            =   5880
            TabIndex        =   66
            Top             =   2880
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   556
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
            BackColor       =   16777215
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnIdentifica 
            Height          =   495
            Index           =   0
            Left            =   5640
            TabIndex        =   74
            Top             =   5880
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            Picture         =   "frmCajas_CargaSaldosFavor.frx":0700
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.PushButton btnIdentifica 
            Height          =   495
            Index           =   1
            Left            =   6840
            TabIndex        =   75
            Top             =   5880
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Cancelar"
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
            Picture         =   "frmCajas_CargaSaldosFavor.frx":0E27
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.ComboBox cboOrigenRecursos 
            Height          =   330
            Left            =   3000
            TabIndex        =   76
            Top             =   5400
            Width           =   5055
            _Version        =   1572864
            _ExtentX        =   8916
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
         Begin XtremeSuiteControls.ComboBox cboPagadores 
            Height          =   330
            Left            =   3000
            TabIndex        =   78
            Top             =   5040
            Width           =   5055
            _Version        =   1572864
            _ExtentX        =   8916
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
         Begin XtremeSuiteControls.PushButton btnAdjuntos 
            Height          =   330
            Left            =   6840
            TabIndex        =   81
            ToolTipText     =   "Adjuntar Documentos"
            Top             =   4680
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Adjuntos"
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
            Picture         =   "frmCajas_CargaSaldosFavor.frx":153D
         End
         Begin XtremeSuiteControls.FlatEdit txtDepositoId 
            Height          =   315
            Left            =   1560
            TabIndex        =   92
            Top             =   480
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   556
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   12
            Left            =   360
            TabIndex        =   93
            Top             =   480
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Id Dep�sito:"
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
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   11
            Left            =   360
            TabIndex        =   80
            Top             =   4560
            Width           =   3615
            _Version        =   1572864
            _ExtentX        =   6376
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Identificaci�n de Recursos:"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   375
            Index           =   10
            Left            =   1560
            TabIndex        =   79
            Top             =   4920
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Pagadores"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   375
            Index           =   9
            Left            =   1560
            TabIndex        =   77
            Top             =   5280
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Origen Recursos"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   71
            Top             =   4080
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Nombre:"
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
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   70
            Top             =   3720
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Identificaci�n:"
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
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   69
            Top             =   3240
            Width           =   3615
            _Version        =   1572864
            _ExtentX        =   6376
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Identificaci�n del Caso:"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   5
            Left            =   4560
            TabIndex        =   62
            Top             =   2280
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha:"
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
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   4
            Left            =   4680
            TabIndex        =   61
            Top             =   2880
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Monto:"
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
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   3
            Left            =   4560
            TabIndex        =   60
            Top             =   960
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "No. Documento:"
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
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   59
            Top             =   1920
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Descripci�n:"
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
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   58
            Top             =   1440
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cuenta:"
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
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   57
            Top             =   960
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "No. Solicitud:"
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
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   9495
            _Version        =   1572864
            _ExtentX        =   16748
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Identificaci�n del Propietario del Dep�sito"
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
      End
      Begin XtremeSuiteControls.GroupBox fraCargaIdentTotales 
         Height          =   1095
         Left            =   -70000
         TabIndex        =   44
         Top             =   6720
         Visible         =   0   'False
         Width           =   11895
         _Version        =   1572864
         _ExtentX        =   20981
         _ExtentY        =   1931
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnProceso 
            Height          =   495
            Index           =   0
            Left            =   8760
            TabIndex        =   45
            Top             =   360
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            Picture         =   "frmCajas_CargaSaldosFavor.frx":15C6
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.PushButton btnProceso 
            Height          =   495
            Index           =   1
            Left            =   9960
            TabIndex        =   46
            Top             =   360
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Cancelar"
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
            Picture         =   "frmCajas_CargaSaldosFavor.frx":1CED
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   315
            Left            =   1320
            TabIndex        =   48
            Top             =   480
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCasos 
            Height          =   315
            Left            =   3000
            TabIndex        =   49
            Top             =   480
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSocios 
            Height          =   315
            Left            =   4080
            TabIndex        =   50
            Top             =   480
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtContratos 
            Height          =   315
            Left            =   5160
            TabIndex        =   51
            Top             =   480
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   3
            Left            =   5160
            TabIndex        =   54
            Top             =   240
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Ident?"
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
            Index           =   2
            Left            =   4080
            TabIndex        =   53
            Top             =   240
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Existe?"
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
            Index           =   1
            Left            =   3000
            TabIndex        =   52
            Top             =   240
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Casos"
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
            Index           =   0
            Left            =   360
            TabIndex        =   47
            Top             =   480
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Totales"
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
      End
      Begin XtremeSuiteControls.CheckBox chkFechas 
         Height          =   375
         Left            =   -67960
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Todas las Fechas"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpRegistroInicio 
         Height          =   315
         Left            =   -69040
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
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
      Begin XtremeSuiteControls.DateTimePicker dtpRegistroCorte 
         Height          =   315
         Left            =   -67720
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
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
      Begin XtremeSuiteControls.CheckBox chkSF_Saldos 
         Height          =   375
         Left            =   -63760
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Saldos en cero"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkSA_Marcas 
         Height          =   375
         Left            =   -69280
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Marcar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   315
         Left            =   -69040
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   315
         Left            =   -64720
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   315
         Left            =   -64720
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtNumDoc 
         Height          =   315
         Left            =   -61240
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4466
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
      Begin XtremeSuiteControls.ComboBox cboTipoSaldo 
         Height          =   315
         Left            =   -61240
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
         _Version        =   1572864
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
      Begin XtremeSuiteControls.ComboBox cboTipoLiquidacion 
         Height          =   315
         Left            =   -61240
         TabIndex        =   13
         Top             =   1200
         Visible         =   0   'False
         Width           =   2535
         _Version        =   1572864
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5895
         Left            =   -70000
         TabIndex        =   21
         Top             =   1680
         Visible         =   0   'False
         Width           =   14415
         _Version        =   524288
         _ExtentX        =   25426
         _ExtentY        =   10398
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
         MaxCols         =   23
         SpreadDesigner  =   "frmCajas_CargaSaldosFavor.frx":2403
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridId 
         Height          =   6375
         Left            =   0
         TabIndex        =   22
         Top             =   1200
         Width           =   11655
         _Version        =   524288
         _ExtentX        =   20558
         _ExtentY        =   11245
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
         MaxCols         =   17
         SpreadDesigner  =   "frmCajas_CargaSaldosFavor.frx":34BA
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpId_Inicio 
         Height          =   315
         Left            =   960
         TabIndex        =   23
         Top             =   480
         Width           =   1335
         _Version        =   1572864
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
      Begin XtremeSuiteControls.DateTimePicker dtpId_Corte 
         Height          =   315
         Left            =   2280
         TabIndex        =   24
         Top             =   480
         Width           =   1335
         _Version        =   1572864
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
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   330
         Left            =   5280
         TabIndex        =   25
         Top             =   480
         Width           =   7095
         _Version        =   1572864
         _ExtentX        =   12515
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
      Begin XtremeSuiteControls.FlatEdit txtId_NumDoc 
         Height          =   315
         Left            =   960
         TabIndex        =   26
         Top             =   840
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
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
      Begin FPSpreadADO.fpSpread vGridCarga 
         Height          =   4695
         Left            =   -70000
         TabIndex        =   30
         Top             =   2160
         Visible         =   0   'False
         Width           =   11895
         _Version        =   524288
         _ExtentX        =   20976
         _ExtentY        =   8276
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
         MaxCols         =   9
         SpreadDesigner  =   "frmCajas_CargaSaldosFavor.frx":3FE2
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboCargaFormaPago 
         Height          =   315
         Left            =   -68440
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
         _Version        =   1572864
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
      Begin XtremeSuiteControls.ComboBox cboCargaCuentaBanco 
         Height          =   315
         Left            =   -68440
         TabIndex        =   32
         Top             =   960
         Visible         =   0   'False
         Width           =   8175
         _Version        =   1572864
         _ExtentX        =   14420
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
      Begin XtremeSuiteControls.CheckBox chkCargaDepositos 
         Height          =   375
         Left            =   -65680
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4043
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Sincronizar con Control de Dep�sitos"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtArchivo 
         Height          =   675
         Left            =   -68440
         TabIndex        =   34
         Top             =   1320
         Visible         =   0   'False
         Width           =   8175
         _Version        =   1572864
         _ExtentX        =   14414
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
         Alignment       =   2
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   375
         Index           =   0
         Left            =   -60040
         TabIndex        =   41
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmCajas_CargaSaldosFavor.frx":47B7
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   375
         Index           =   1
         Left            =   -59560
         TabIndex        =   42
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmCajas_CargaSaldosFavor.frx":4EB7
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   375
         Index           =   2
         Left            =   -59080
         TabIndex        =   43
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmCajas_CargaSaldosFavor.frx":55D0
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit1 
         Height          =   315
         Left            =   5280
         TabIndex        =   82
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMnt_Inicio 
         Height          =   315
         Left            =   8760
         TabIndex        =   85
         Top             =   840
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
      Begin XtremeSuiteControls.FlatEdit txtMnt_Hasta 
         Height          =   315
         Left            =   10560
         TabIndex        =   86
         Top             =   840
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
      Begin XtremeSuiteControls.CheckBox chkI_Recursos 
         Height          =   375
         Left            =   -66040
         TabIndex        =   87
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1572864
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Casos con Identificaci�n de Recursos"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboSC_EntidadPagadora 
         Height          =   330
         Left            =   -56920
         TabIndex        =   88
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1572864
         _ExtentX        =   5106
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
      Begin XtremeSuiteControls.ComboBox cboSC_OrigenRecursos 
         Height          =   330
         Left            =   -56920
         TabIndex        =   90
         Top             =   840
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1572864
         _ExtentX        =   5106
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
      Begin XtremeSuiteControls.FlatEdit txtSC_Mnt_Inicio 
         Height          =   315
         Left            =   -56920
         TabIndex        =   94
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
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
      Begin XtremeSuiteControls.FlatEdit txtSC_Mnt_Hasta 
         Height          =   315
         Left            =   -55480
         TabIndex        =   95
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
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
      Begin VB.Label Label2 
         Caption         =   "Montos.: "
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
         Index           =   12
         Left            =   -58480
         TabIndex        =   96
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Origen Recursos:"
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
         Left            =   -58480
         TabIndex        =   91
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Entidad Pagadora:"
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
         Left            =   -58480
         TabIndex        =   89
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Montos.: "
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
         Left            =   7800
         TabIndex        =   84
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "C�dula Origen.: "
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
         Left            =   3840
         TabIndex        =   83
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
         DataField       =   "Banco"
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
         Index           =   1
         Left            =   -69760
         TabIndex        =   37
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Doc.:"
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
         Index           =   19
         Left            =   -69760
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo"
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
         Index           =   2
         Left            =   -69760
         TabIndex        =   35
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha .:"
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
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "No. Doc.:"
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
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta .:"
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
         Index           =   16
         Left            =   3840
         TabIndex        =   27
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario.:"
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
         Left            =   -69880
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha .:"
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
         Left            =   -69880
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Identificaci�n.:"
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
         Left            =   -66160
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre.:"
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
         Left            =   -65920
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Doc.:"
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
         Index           =   4
         Left            =   -62320
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Doc.:"
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
         Left            =   -62320
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Liq.:"
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
         Index           =   18
         Left            =   -62320
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarC 
      Height          =   150
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   15015
      _Version        =   1572864
      _ExtentX        =   26485
      _ExtentY        =   265
      _StockProps     =   93
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   39
      Top             =   960
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Liquidar"
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
      Picture         =   "frmCajas_CargaSaldosFavor.frx":5CE9
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   40
      Top             =   960
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Configurar"
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
      Picture         =   "frmCajas_CargaSaldosFavor.frx":65BA
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   97
      Top             =   960
      Width           =   2655
      _Version        =   1572864
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Reenv�o de Correo a Cobros"
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
      Picture         =   "frmCajas_CargaSaldosFavor.frx":6CAD
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   99
      Top             =   960
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Identificar"
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
      Picture         =   "frmCajas_CargaSaldosFavor.frx":6E18
      ImageAlignment  =   4
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Administraci�n de Saldos a Favor"
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
      Height          =   495
      Index           =   0
      Left            =   1880
      TabIndex        =   0
      Top             =   240
      Width           =   5505
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmCajas_CargaSaldosFavor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub sbIdentifica_Lista()

On Error GoTo vError

Dim i As Long, curTotal As Currency
Dim pDepositoId As Long, pTesoreriaId As Long, pDocumento As String, pMonto As Currency, pFecha As Date, pDescripcion As String, pCuenta As String

If vPaso Then Exit Sub

txtId_Cedula.Text = GLOBALES.gTag
txtId_Nombre.Text = GLOBALES.gTag2

lsw.ListItems.Clear
curTotal = 0

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id Deposito", 1200
    .Add , , "Id Tesoreria", 1200
    .Add , , "Documento", 1800
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Fecha", 1200, vbCenter
    .Add , , "Descripci�n", 2200
    .Add , , "Cuenta", 2200
End With

With vGridId
    
For i = 1 To .MaxRows
    .Row = i
    .Col = 1
    If .Value = vbChecked Then
        .Col = 2
        pDepositoId = .Text
        .Col = 3
        pTesoreriaId = IIf(IsNumeric(.Text), .Text, 0)
        .Col = 4
        pCuenta = .Text
        
        .Col = 6
        pDocumento = .Text
        .Col = 7
        pFecha = .Text
        .Col = 8
        pMonto = .Text
        .Col = 9
        pDescripcion = .Text
        
        curTotal = curTotal + pMonto
        
        Set itmX = lsw.ListItems.Add(, , pDepositoId)
            itmX.SubItems(1) = pTesoreriaId
            itmX.SubItems(2) = pDocumento
            itmX.SubItems(3) = Format(pMonto, "Standard")
            itmX.SubItems(4) = Format(pFecha, "yyyy-mm-dd")
            itmX.SubItems(5) = pDescripcion
            itmX.SubItems(6) = pCuenta
            
            .Col = 4
            itmX.Tag = .CellTag 'Id Banco
    End If
    
Next i

End With

txtId_Monto.Text = Format(curTotal, "Standard")

If lsw.ListItems.Count = 0 Then
    MsgBox "No se ha seleccionado ning�n caso!", vbExclamation
    Exit Sub
End If

fraIdentifica.Visible = True


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnAccion_Click(Index As Integer)
Select Case Index
  Case 0 'Buscar
  
     Select Case tcMain.SelectedItem
        Case 0 'Saldos Actuales
            vPaso = True
            Call sbConsultaSaldosAfavor
            vPaso = False
        Case 1 'Por Identificar
            vPaso = True
            Call sbConsultaDPTramite
            vPaso = False
     End Select
     
  Case 1 'Liquidar
     If cboTipoLiquidacion.ListCount > 0 Then
        If Mid(cboTipoLiquidacion.Text, 1, 1) <> "N" Then
           Call sbLiquidaSF
        Else
           MsgBox "No se ha indicado un metodo de liquidaci�n del Saldo a Favor!", vbExclamation
        End If
     Else
           MsgBox "No se ha indicado un metodo de liquidaci�n del Saldo a Favor!", vbExclamation
     End If
    
  Case 2 'Configura
       Call sbFormsCall("frmCajas_SaldosFavorLiquidaConfigura", 1, , , False, Me)
       
       
  Case 3 'Notifica Correo de Saldo a Favor Identificado!
    Call sbNotificaSaldoaFavor

  Case 4 'Identificar
    Call sbIdentifica_Lista
  
End Select

End Sub

Private Sub btnAdjuntos_Click()
If txtId_Cedula.Text <> "" Then
 gGA.Modulo = "CAJ"
 gGA.Llave_01 = txtId_Cedula.Text
 gGA.Llave_02 = ""
 gGA.Llave_03 = ""
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End If

End Sub

Private Sub btnArchivo_Click(Index As Integer)
Dim vMensaje As String
  
Select Case Index
  
  Case 0 'buscar
        txtArchivo.Text = ""
        Call sbArchivoBusca

  Case 1 'cargar
       Call sbArchivoCarga


  Case 2 'info
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: Import" & vbCrLf _
              & " 3. Columnas.: DOCUMENTO, FECHA, MONTO, DESCRIPCION, CEDULA"
     
     MsgBox vMensaje, vbInformation
     
     
End Select

End Sub

Private Sub btnIdentifica_Click(Index As Integer)
On Error GoTo vError

If Index = 1 Then
   fraIdentifica.Visible = False
   Exit Sub
End If

If txtId_Nombre.Text = "" Then
    MsgBox "No se ha especificado ning�n Id de Cliente v�lido", vbExclamation
    Exit Sub
End If

If lsw.ListItems.Count = 0 Then
    MsgBox "No se ha seleccionado ning�n caso!", vbExclamation
    Exit Sub
End If

Me.MousePointer = vbHourglass

Dim i As Long
Dim pDepositoId As Long, pBancoId As Long, pDocumento As String

With lsw.ListItems
    For i = 1 To .Count
      pDepositoId = .Item(i).Text
      pBancoId = .Item(i).Tag
      pDocumento = .Item(i).SubItems(2)
        
        strSQL = "exec spCajas_Identifica_TES_Depositos " & pBancoId & ",'" & pDocumento & "','" & txtId_Cedula.Text _
               & "', '" & txtId_Nombre.Text & "', '" & glogon.Usuario & "', '" & cboPagadores.ItemData(cboPagadores.ListIndex) _
               & "', '" & cboOrigenRecursos.ItemData(cboOrigenRecursos.ListIndex) & "', " & pDepositoId
        Call ConectionExecute(strSQL)

    Next i
    
End With

Me.MousePointer = vbDefault

fraIdentifica.Visible = False

MsgBox "caso identificado correctamente, notificaciones realizadas!", vbInformation

vPaso = True
    Call sbConsultaDPTramite
vPaso = False

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnProceso_Click(Index As Integer)
Select Case Index
  Case 0 'Aplicar
  
    If vGridCarga.MaxRows = 0 Then
       MsgBox "No existen registros cargados...[verifique!]", vbExclamation
       Exit Sub
    End If
   
    Call sbProcesar
  
  Case 1 'cancelar
    vGridCarga.MaxRows = 0
    txtArchivo.Text = ""

End Select

End Sub

Private Sub cboBanco_Click()
vGridId.MaxRows = 0
End Sub

Private Sub cboCargaFormaPago_Click()


If vPaso Then Exit Sub
If cboCargaFormaPago.ListCount = 0 Then Exit Sub



strSQL = "select Tipo from sif_formas_pago where COD_FORMA_PAGO = '" & cboCargaFormaPago.ItemData(cboCargaFormaPago.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Tipo = "B" Then
   chkCargaDepositos.Value = vbChecked
Else
   chkCargaDepositos.Value = vbUnchecked
End If
rs.Close

cboCargaCuentaBanco.Clear

strSQL = "SELECT Bn.ID_BANCO as 'Idx', '(' + rtrim(Bn.CTA) + ') ' + rtrim(Bn.DESCRIPCION) as 'ItmX'" _
       & " FROM SIF_FORMAS_PAGO_BANCOS_ASG Fp inner join TES_BANCOS Bn on Fp.ID_BANCO = Bn.ID_BANCO" _
       & " where Fp.COD_FORMA_PAGO = '" & cboCargaFormaPago.ItemData(cboCargaFormaPago.ListIndex) & "'"
Call sbCbo_Llena_New(cboCargaCuentaBanco, strSQL, False, True)

End Sub

Private Sub cboTipoSaldo_Click()
If vPaso Or cboTipoSaldo.ListCount = 0 Then Exit Sub


cboTipoLiquidacion.Clear

If cboTipoSaldo.Text <> "TODOS" Then
  strSQL = "select dbo.fxCajas_SaldoFavorTipoLiquidacion('" & cboTipoSaldo.ItemData(cboTipoSaldo.ListIndex) _
         & "','" & glogon.Usuario & "') as 'TipoLiquidacion'"
         
  strSQL = "exec spCajas_SaldoFavorTipoLiquidacion '" & glogon.Usuario & "', '" & cboTipoSaldo.ItemData(cboTipoSaldo.ListIndex) & "'"
  Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
    cboTipoLiquidacion.AddItem rs!Tipo
    rs.MoveNext
  Loop
  If Not rs.EOF And Not rs.BOF Then
    rs.MoveFirst
    cboTipoLiquidacion.Text = rs!Tipo
  End If
  rs.Close
End If


End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = vbChecked Then
   dtpRegistroInicio.Enabled = False
Else
   dtpRegistroInicio.Enabled = True
End If

dtpRegistroCorte.Enabled = dtpRegistroInicio.Enabled
  
End Sub


Private Sub chkSA_Marcas_Click()
Dim i As Long


For i = 1 To vGrid.MaxRows
   vGrid.Row = 1
   vGrid.Col = 1
   vGrid.Value = chkSA_Marcas.Value
Next i


End Sub

Private Sub Form_Activate()
 vModulo = 5

End Sub

Private Sub Form_Load()

vModulo = 5

'Carga las cuentas bancarias asiganadas a la forma de pago
vPaso = True


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

cboBanco.Clear

strSQL = "exec spCajas_DepositosCuentasBancariasAut 'DP'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboBanco.AddItem Trim(rs!Cta) & " - " & Trim(rs!Descripcion & "")
 cboBanco.ItemData(cboBanco.ListCount - 1) = CStr(rs!Id_Banco)
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
    cboBanco.Text = Trim(rs!Cta) & " - " & Trim(rs!Descripcion & "")
End If
rs.Close


strSQL = "select  rtrim(DOC_TIPO) as 'IdX', rtrim(DESCRIPCION) as 'itmX' from CAJAS_SALDOS_FAVOR_TIPOS" _
       & " WHERE ACTIVO = 1 ORDER BY DOC_TIPO"
Call sbCbo_Llena_New(cboTipoSaldo, strSQL, True, True)


strSQL = "select  rtrim(COD_FORMA_PAGO) as 'IdX', rtrim(DESCRIPCION) as 'itmX' from SIF_FORMAS_PAGO" _
       & " WHERE ACTIVA = 1 AND TIPO IN('B','T')  ORDER BY COD_FORMA_PAGO"
Call sbCbo_Llena_New(cboCargaFormaPago, strSQL, False, True)


'Identificacion de Recursos
strSQL = "select COD_ENTIDAD_PAGO as 'IdX', DESCRIPCION AS 'ItmX' from SIF_ENTIDADES_PAGO" _
       & " WHERE ACTIVA = 1 ORDER BY COD_ENTIDAD_PAGO"
Call sbCbo_Llena_New(cboPagadores, strSQL, False, True)

strSQL = "select COD_ORIGEN_RECURSOS as 'IdX', DESCRIPCION AS 'ItmX' from SIF_ORIGEN_RECURSOS" _
       & "  WHERE ACTIVA = 1 ORDER BY COD_ORIGEN_RECURSOS"
Call sbCbo_Llena_New(cboOrigenRecursos, strSQL, False, True)


'Consulta de Saldos en Cajas
strSQL = "select COD_ENTIDAD_PAGO as 'IdX', DESCRIPCION AS 'ItmX' from SIF_ENTIDADES_PAGO" _
       & " WHERE ACTIVA = 1 ORDER BY DESCRIPCION"
Call sbCbo_Llena_New(cboSC_EntidadPagadora, strSQL, True, True)

strSQL = "select COD_ORIGEN_RECURSOS as 'IdX', DESCRIPCION AS 'ItmX' from SIF_ORIGEN_RECURSOS" _
       & "  WHERE ACTIVA = 1 ORDER BY DESCRIPCION"
Call sbCbo_Llena_New(cboSC_OrigenRecursos, strSQL, True, True)


vPaso = False

Call cboCargaFormaPago_Click


vPaso = True
    vGrid.MaxCols = 23
    vGrid.MaxRows = 0
    
    vGridId.MaxCols = 17
    vGridId.MaxRows = 0

    vGridCarga.MaxRows = 0

vPaso = False

txtMnt_Inicio.Text = Format(0, "Standard")
txtMnt_Hasta.Text = Format(999999999999.99, "Standard")

txtSC_Mnt_Inicio.Text = Format(0, "Standard")
txtSC_Mnt_Hasta.Text = Format(999999999999.99, "Standard")


dtpId_Inicio.Value = fxFechaServidor
dtpId_Corte.Value = dtpId_Inicio.Value


dtpRegistroInicio.Value = dtpId_Inicio.Value
dtpRegistroCorte.Value = dtpId_Inicio.Value

Call chkFechas_Click


tcMain.Item(0).Selected = True

Call RefrescaTags(Me)
Call Formularios(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next


tcMain.Width = Me.Width - 400
tcMain.Height = Me.Height - (tcMain.top + 600)

vGrid.Width = tcMain.Width - 300
vGrid.Height = tcMain.Height - (vGrid.top + 200)


vGridId.Width = vGrid.Width
vGridId.Height = tcMain.Height - (vGridId.top + 200)

vGridCarga.Width = vGrid.Width
vGridCarga.Height = tcMain.Height - (vGridCarga.top + fraCargaIdentTotales.Height + 200)

fraCargaIdentTotales.top = vGridCarga.top + vGridCarga.Height + 20

fraCargaIdentTotales.Width = tcMain.Width

imgBanner.Width = Me.Width

ProgressBarC.Width = tcMain.Width

End Sub


Private Sub sbLiquidaSF()
Dim i As Long, vIdSaldoFavor As Long
Dim vMetodo As String, vMonto As Currency


Me.MousePointer = vbHourglass

On Error GoTo vError

vMetodo = Mid(cboTipoLiquidacion.Text, 1, 1)

ProgressBarC.Visible = True

With vGrid
 ProgressBarC.Max = .MaxRows
 
 For i = 1 To .MaxRows
    .Row = i
    .Col = 1
    If .Value = vbChecked Then
       .Col = 3
       vIdSaldoFavor = .Text
       
       .Col = 9
       vMonto = CCur(.Text)
       
       If vMonto > 0 Then
       
        Select Case vMetodo
            Case "T" 'Tesoreria
                     strSQL = "exec spCajas_SaldoFavorLiquidacionTesoreria " & vIdSaldoFavor & ",'" & glogon.Usuario & "'"
            Case "F" 'Fondos
                     strSQL = "exec spCajas_SaldoFavorLiquidacionFondos " & vIdSaldoFavor & ",'" & glogon.Usuario & "'"
            Case "E" 'Excluye
                     strSQL = "exec spCajas_SaldoFavorLiquidacionExclusion " & vIdSaldoFavor & ",'" & glogon.Usuario & "'"
            Case "C" 'Cajas
                     strSQL = "exec spCajas_SaldoFavorLiquidacionRC_Efectivo " & vIdSaldoFavor & ",'" & glogon.Usuario & "'"
            Case Else
                     strSQL = "exec spCajas_SaldoFavorLiquidacionTesoreria " & vIdSaldoFavor & ",'" & glogon.Usuario & "'"
        End Select
        Call ConectionExecute(strSQL)
                
        Call Bitacora("Aplica", "Liquidaci�n de Saldo a Favor: " & cboTipoLiquidacion.Text & " (id." & vIdSaldoFavor & ")")
       End If 'vMonto > 0
    
    End If
    ProgressBarC.Value = i
 Next i
End With
Me.MousePointer = vbDefault

MsgBox "Saldos a Favor liquidados Satisfactoriamente..!", vbInformation

'Refresca la Lista
vPaso = True
    Call sbConsultaSaldosAfavor
vPaso = False


ProgressBarC.Visible = False

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
ProgressBarC.Visible = False
Call sbConsultaSaldosAfavor

End Sub


Private Sub sbNotificaSaldoaFavor()
Dim i As Long, vIdSaldoFavor As Long

Dim iCasos As Long

Me.MousePointer = vbHourglass

On Error GoTo vError

iCasos = 0
ProgressBarC.Visible = True

With vGrid
 ProgressBarC.Max = .MaxRows
 
 For i = 1 To .MaxRows
    .Row = i
    .Col = 1
    If .Value = vbChecked Then
       .Col = 3
       vIdSaldoFavor = .Text
       'exec spCajasNotificaDepositos @Banco, @Documento, @Linea
        strSQL = "exec spCajasNotificaDepositos Null, Null, " & vIdSaldoFavor
        Call ConectionExecute(strSQL)
        
        iCasos = iCasos + 1
    
    End If
    ProgressBarC.Value = i
 Next i
End With
Me.MousePointer = vbDefault

If iCasos > 0 Then
    MsgBox "Se han notificado a Cobros los casos marcados..!", vbInformation
Else
    MsgBox "No se proces� ning�n caso! Verifique que existan casos seleccionados!", vbExclamation
End If


ProgressBarC.Visible = False

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbConsultaSaldosAfavor()
Dim i As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select *" _
       & " From vCajas_Saldos_Favor"

If chkSF_Saldos.Value = vbChecked Then
    strSQL = strSQL & " Where Saldo <= 0"
Else
    strSQL = strSQL & " Where Saldo > 0"
End If


If chkFechas.Value = vbUnchecked Then
    strSQL = strSQL & " and registro_fecha between '" & Format(dtpRegistroInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
       & Format(dtpRegistroCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End If

If Len(Trim(txtCedula.Text)) > 0 Then
    strSQL = strSQL & " and Cedula like '%" & txtCedula.Text & "%'"
End If

If Len(Trim(txtNombre.Text)) > 0 Then
    strSQL = strSQL & " and isnull(Nombre,'') like '%" & txtNombre.Text & "%'"
End If

If Trim(cboTipoSaldo.Text) <> "TODOS" Then
    strSQL = strSQL & " and Doc_Tipo = '" & cboTipoSaldo.ItemData(cboTipoSaldo.ListIndex) & "'"
End If

If Len(Trim(txtNumDoc.Text)) > 0 Then
    strSQL = strSQL & " and Doc_Numero like '%" & txtNumDoc.Text & "%'"
End If

If Len(Trim(txtUsuario.Text)) > 0 Then
    strSQL = strSQL & " and Registro_Usuario like '%" & txtUsuario.Text & "%'"
End If

If Trim(cboSC_EntidadPagadora.Text) <> "TODOS" Then
    strSQL = strSQL & " and COD_ENTIDAD_PAGO = '" & cboSC_EntidadPagadora.ItemData(cboSC_EntidadPagadora.ListIndex) & "'"
End If

If Trim(cboSC_OrigenRecursos.Text) <> "TODOS" Then
    strSQL = strSQL & " and COD_ORIGEN_RECURSOS = '" & cboSC_OrigenRecursos.ItemData(cboSC_OrigenRecursos.ListIndex) & "'"
End If

If chkI_Recursos.Value = xtpChecked Then
    strSQL = strSQL & " and COD_ORIGEN_RECURSOS is not null"
End If

strSQL = strSQL & " and Monto between " & CCur(txtSC_Mnt_Inicio.Text) & " and " & CCur(txtSC_Mnt_Hasta.Text)

strSQL = strSQL & " ORDER BY REGISTRO_FECHA DESC"

Call OpenRecordSet(rs, strSQL)

vGrid.MaxRows = 0


  Do While Not rs.EOF
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
         
    vGrid.Col = 1
    vGrid.Value = chkSA_Marcas.Value
    
    For i = 3 To vGrid.MaxCols
      vGrid.Col = i
      Select Case i
         Case 3 'Linea
            vGrid.Text = CStr(rs!Linea)
         Case 4 'Cedula
            vGrid.Text = rs!Cedula & ""
         Case 5 'Nombre
            vGrid.Text = rs!Nombre & ""
         Case 6 'Tipo Doc
            vGrid.Text = rs!DOC_TIPO
         Case 7 'Num Documento
            vGrid.Text = rs!Doc_Numero
         Case 8 'Monto
            vGrid.Text = Format(rs!Monto, "Standard")
         Case 9 'Saldo
            vGrid.Text = Format(rs!Saldo, "Standard")
         Case 10 'Divisa
            vGrid.Text = rs!cod_Divisa & ""
           
         
         
         Case 11 'Registro Fecha
            vGrid.Text = rs!REGISTRO_FECHA_FORMAT & ""
         Case 12 'Registro Usuario
            vGrid.Text = rs!Registro_Usuario & ""
      
      
         Case 13 'Liquida Fecha
            vGrid.Text = rs!Liq_Fecha & ""
         Case 14 'Liquida Usuario
            vGrid.Text = rs!LIQ_USUARIO & ""
         Case 15 'Liquida Monto
            vGrid.Text = rs!LIQ_MONTO & ""
         Case 16 'Liquida No. Tesoreria
            vGrid.Text = rs!LIQ_NSOLICITUD & ""
         Case 17 'Liquida Plan
            vGrid.Text = rs!LIQ_PLAN & ""
         Case 18 'Liquida Contra
            vGrid.Text = rs!LIQ_CONTRATO & ""
         Case 19 'Liquida Tipo Comprobante
            vGrid.Text = rs!LIQ_TIPO_DOC & ""
         Case 20 'Liquida nO. Comprobante
            vGrid.Text = rs!LIQ_NUM_DOC & ""
      
         Case 21 'Cuenta Bancaria Relacionada
            vGrid.Text = rs!BancoDesc & ""
            
         Case 22 'Entidad Pagadora
            vGrid.Text = rs!EntidadPagoDesc & ""
      
         Case 23 'Origen de Recursos
            vGrid.Text = rs!OrigenRecursoDesc & ""
      
      
      End Select
    Next i
     rs.MoveNext
   Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbConsultaDPTramite()
Dim i As Long

On Error GoTo vError

If cboBanco.ListCount = 0 Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "select *" _
        & " From vTes_Depositos_Tramite_Identifica" _
        & " Where ID_REQUERIDA = 1 And IDENTIFICADO = 0" _
        & " and  fecha between '" & Format(dtpId_Inicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpId_Corte.Value, "yyyy/mm/dd") & " 23:59:59'"


If Len(Trim(txtId_NumDoc.Text)) > 0 Then
    strSQL = strSQL & " and Documento like '%" & txtId_NumDoc.Text & "%'"
End If

strSQL = strSQL & " and Id_Banco = " & cboBanco.ItemData(cboBanco.ListIndex)

strSQL = strSQL & " and Monto between " & CCur(txtMnt_Inicio.Text) & " and " & CCur(txtMnt_Hasta.Text)


Call OpenRecordSet(rs, strSQL)

vGridId.MaxRows = 0

'DP_TRAMITE_ID
  Do While Not rs.EOF
    vGridId.MaxRows = vGridId.MaxRows + 1
    vGridId.Row = vGridId.MaxRows
         
    vGridId.Col = 1

    For i = 2 To vGridId.MaxCols
      vGridId.Col = i
      Select Case i
         Case 2 'Id Tramite
            vGridId.Text = CStr(rs!DP_TRAMITE_ID)
         
         Case 3 'Id
            vGridId.Text = CStr(rs!NSolicitud)
         Case 4 'Cuenta
            vGridId.Text = rs!BancoDesc & ""
            vGridId.CellTag = rs!Id_Banco
         Case 5 ' Tipo
            vGridId.Text = "DP"
         Case 6 'Num Documento
            vGridId.Text = rs!Documento
         Case 7 'Fecha del Documento
            vGridId.Text = Format(rs!fecha, "dd/mm/yyyy")
         Case 8 'Monto
            vGridId.Text = Format(rs!Monto, "Standard")
         Case 9 'Descripcion
            vGridId.Text = rs!Descripcion
         Case 10 'Registro Fecha
            vGridId.Text = rs!Registro_Fecha & ""
         Case 11 'Registro Usuario
            vGridId.Text = rs!Registro_Usuario & ""
            
         'Datos Sinpe
         Case 12
            vGridId.Text = rs!Id_Servicio & ""
         Case 13
            vGridId.Text = rs!Cuenta_Iban_Origen & ""
         Case 14
            vGridId.Text = rs!CEDULA_ORIGEN & ""
         Case 15
            vGridId.Text = rs!Nombre_Origen & ""
         Case 16
            vGridId.Text = rs!Telefono_Origen & ""
         Case 17
            vGridId.Text = rs!Banco_Origen & ""

            
      
      End Select
    Next i
     rs.MoveNext
   Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbComprobanteSF(pId As Long)
Dim x As New clsImpresoras
Dim vFlat As Boolean
Dim vEmpresa As String, vCedJur As String
Dim vArchivo As String

On Error GoTo vError

strSQL = "select nombre,cedula_juridica from sif_empresa"
Call OpenRecordSet(rs, strSQL)
 vEmpresa = UCase(rs!Nombre & "")
 vCedJur = Trim(rs!cedula_juridica & "")
rs.Close

With frmContenedor.Crt
   .Reset
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Cajas: Comprobante de Descargo de Saldos a Favor"
   
   .Connect = glogon.ConectRPT
   
    x.TipoImpresora = Recibos
    x.Reset
    .PrinterDriver = x.Controlador
    .PrinterName = x.Nombre
    .PrinterPort = x.Puerto
    
    .PrinterSelect
    
    .Destination = crptToWindow
     
    .Formulas(0) = "fxEmpresa = '" & vEmpresa & "'"
    .Formulas(1) = "fxCedJur = '" & vCedJur & "'"
    .Formulas(2) = "fxUsuario = '" & glogon.Usuario & "'"
    .Formulas(3) = "fxFecha = '" & fxFechaServidor & "'"
    
    .ReportFileName = SIFGlobal.fxPathReportes("Cajas_SF_Comprobante.rpt")
    
    .SelectionFormula = "{CAJAS_SALDO_FAVOR.LINEA} = " & pId
    
   .PrintReport
End With

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaLimpia()

    txtMonto.Text = 0
    txtCasos.Text = 0
    txtSocios.Text = 0
    txtContratos.Text = 0
    txtArchivo.Text = ""
    
    vGridCarga.MaxRows = 0
    
End Sub



Private Sub tlbCarga_ButtonClick(ByVal Button As MSComctlLib.Button)


End Sub

Private Sub sbArchivoBusca()

With frmContenedor.CD
    .InitDir = "C:\"
    .DialogTitle = "Localice Archivo de Dep�sitos del Banco  [Microsoft EXCEL]"
    .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
    .ShowOpen
    
    If .FileName = "" Then
        MsgBox "Archivo no v�lido...", vbExclamation
        Exit Sub
    End If
    
    If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
        'Ok
    Else
        MsgBox "La Extensi�n del Archivo no es v�lido...", vbExclamation
        Exit Sub
    End If

    txtArchivo.Text = .FileName
End With

End Sub


Private Sub sbArchivoCarga()
Dim vCtrlDPActivo As Boolean, rsTmp As New ADODB.Recordset

Dim i As Integer, iCampos As Integer, vExiste As Integer
Dim vFecha As Date, vDocumento As String, vMonto As Currency, vDescripcion As String
Dim vCedula As String, vNombre As String, vInconsistencia As String
Dim curMonto As Currency, lCasos As Long

On Error GoTo vError
vGrid.MaxRows = 0

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

If chkCargaDepositos.Value = vbChecked Then
    vCtrlDPActivo = True
Else
    vCtrlDPActivo = False
End If



If vCtrlDPActivo Then
    If cboCargaCuentaBanco.ListCount <= 0 Then
        MsgBox "No existe ninguna cuenta bancaria, no se puede procesar el archivo...", vbCritical
        Exit Sub
    End If
End If

Me.MousePointer = vbHourglass

vGrid.MaxRows = 0

curMonto = 0
txtCasos = 0 'Total


Set rs = Excel_Load(txtArchivo.Text, "Import")

'Verifica Estructura del Archivo

iCampos = 0
For i = 0 To rs.RecordCount - 1
   Select Case UCase(rs(i).Name)
      Case "DOCUMENTO", "FECHA", "MONTO", "DESCRIPCION", "CEDULA"
        iCampos = iCampos + 1
      Case Else
      
   End Select
Next i

If iCampos < 5 Then
   Me.MousePointer = vbDefault
   MsgBox "1. No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "2. Los campos son Documento, Fecha, Monto, Descripcion, CEDULA", vbExclamation
    rs.Close
 
   Exit Sub
End If


With vGridCarga



    Do While Not rs.EOF
        vDocumento = Trim(rs!Documento)
        vFecha = rs!fecha
        vMonto = rs!Monto
        vDescripcion = rs!Descripcion & ""
        vCedula = rs!Cedula & ""
        vNombre = Trim(fxNombre(vCedula))
        
      If vDocumento <> "" Then
            
          If vCtrlDPActivo Then
                strSQL = "select dbo.fxTes_DP_Cargado(" & cboBanco.ItemData(cboBanco.ListIndex) & ",'" & vDocumento & "',''," & vMonto & ") as Existe"
                Call OpenRecordSet(rsTmp, strSQL)
                  vExiste = rsTmp!Existe
                  If vExiste > 0 Then vExiste = 1
                  
                  Select Case rsTmp!Existe
                        Case 0 'Sin Inconsistencia
                          vInconsistencia = ""
                        Case 1 'Existe  / Identificado
                          vInconsistencia = "Existe  / Identificado"
                        Case 2 'Existe  / No Identificado
                          vInconsistencia = "Existe  / No Identificado"
                        Case 3 'Existe Registro pero a nombre de otra persona
                          vInconsistencia = "Existe Registro pero a nombre de otra persona"
                        Case 4 'Existe Registro con Monto Diferente
                          vInconsistencia = "Existe Registro con Monto Diferente"
                  End Select
                  
                rsTmp.Close
           End If
           
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1
            .Value = vbChecked
            
            .Col = 2
            .Value = vExiste
            
            .Col = 3
            .Text = vDocumento
            .Col = 4
            .Text = CStr(vMonto)
            .Col = 5
            .Text = vFecha
            .Col = 6
            .Text = vDescripcion
            .Col = 7
            .Text = vInconsistencia
            .Col = 8
            .Text = vCedula
            .Col = 9
            .Text = vNombre
            
            If vNombre = "" Then
               .Col = 7
               .Text = "No existe registro de la Persona"
            End If
            
            curMonto = curMonto + vMonto
            txtCasos = txtCasos + 1
            txtCasos.Refresh
       
       End If
       rs.MoveNext
    Loop
End With
        
'Totales
txtMonto.Text = Format(curMonto, "Standard")
Me.MousePointer = vbDefault


MsgBox "Informaci�n Cargada Satisfactoriamente", vbInformation

rs.Close

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   Call sbCargaLimpia
End Sub





Private Sub sbProcesar()
Dim vCtrlDPActivo As Boolean
Dim i As Long, vDescripcion As String, vCuenta As String, vInconsistencia As String, vCedula As String, vNombre As String
Dim vRequiereId As Integer, vDocumento As String, vMonto As Currency, vFecha As Date, vExiste As Integer
Dim vMensaje As Boolean, vCasos As Long, vBanco As Long

On Error GoTo vError

If chkCargaDepositos.Value = vbChecked Then
  vCtrlDPActivo = True
Else
  vCtrlDPActivo = False
End If


strSQL = "select cod_cuenta from sif_Formas_Pago where cod_Forma_Pago = '" & cboCargaFormaPago.ItemData(cboCargaFormaPago.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
  vCuenta = Trim(rs!cod_cuenta)
rs.Close


If Not fxgCntCuentaValida(vCuenta) Then
   MsgBox "La cuenta especificada para registro no es v�lida...verifique!", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass

vMensaje = False
vCasos = 0

With vGridCarga
    For i = 1 To .MaxRows

       .Row = i
       .Col = 1
       vRequiereId = .Value
       .Col = 2
       vExiste = .Value
       .Col = 3
       vDocumento = .Text
       .Col = 4
       vMonto = CCur(.Text)
       .Col = 5
       vFecha = Format(.Text, "yyyy/mm/dd")
       .Col = 6
       vDescripcion = .Text
       .Col = 7
       vInconsistencia = .Text
       .Col = 8
       vCedula = .Text
       .Col = 9
       vNombre = .Text
       
       
       If vCtrlDPActivo Then
           
            vBanco = cboCargaCuentaBanco.ItemData(cboCargaCuentaBanco.ListIndex)
            If vInconsistencia = "" Or vInconsistencia = "No existe registro de la Persona" Then
                strSQL = "insert TES_DEPOSITOS_TRAMITE(id_Banco,documento,nsolicitud,fecha,monto,descripcion,registro_fecha,registro_usuario " _
                       & ",id_requerida,identificado, cod_cuenta)" _
                       & " values(" & vBanco & ",'" & vDocumento & "',0,'" & Format(vFecha, "yyyy/mm/dd") _
                       & "'," & vMonto & ",'" & vDescripcion & "',dbo.MyGetdate(),'" & glogon.Usuario & "'," & vRequiereId & ",0,'" & vCuenta & "')"
                Call ConectionExecute(strSQL)
                
                vCasos = vCasos + 1
                                
                If vInconsistencia = "" Then
                    strSQL = "exec spCajas_Identifica_TES_Depositos " & vBanco & ",'" & vDocumento & "','" & vCedula _
                           & "','" & vNombre & "','" & glogon.Usuario & "'"
                    Call ConectionExecute(strSQL)
                End If
            Else
                strSQL = "insert TES_DEPOSITOS_TRAMITE_INCONSISTENCIAS(id_Banco,documento,fecha,monto,descripcion,registro_fecha,registro_usuario " _
                       & ",inconsistencia)" _
                       & " values(" & vBanco & ",'" & vDocumento & "','" & Format(vFecha, "yyyy/mm/dd") _
                       & "'," & vMonto & ",'" & vDescripcion & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & vInconsistencia & "')"
                Call ConectionExecute(strSQL)
                
                vMensaje = True
            End If
        
      Else
        'Carga Simple en Saldos a Favor
        If vInconsistencia <> "" Then
                strSQL = "exec spCajas_SaldoFavorCarga '" & cboCargaFormaPago.ItemData(cboCargaFormaPago.ListIndex) & "','" & vDocumento _
                       & "','" & vCedula & "','" & vNombre & "','" & glogon.Usuario & "'"
                Call ConectionExecute(strSQL)
                
                vCasos = vCasos + 1
        End If
      
      End If 'CtrlDPACtivo
       
    Next i
End With
Me.MousePointer = vbDefault

If vCasos = 0 Then
    MsgBox "No se procesaron casos *--Revisados--* para el control de dep�sitos!", vbExclamation
Else
    MsgBox "Carga realizada Satisfactoriamente... Registros Procesados :" & vCasos, vbInformation
End If

If vMensaje Then
    MsgBox "Se presentaron inconsistencias en la carga..Revise en el TAB de consulta de inconsistencias!", vbExclamation
End If


Call sbCargaLimpia

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbCargaLimpia

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

btnAccion(3).Visible = False
btnAccion(4).Visible = False

Select Case Item.Index
    Case 0
        btnAccion(3).Visible = True
    Case 1
        btnAccion(4).Visible = True
            
        fraIdentifica.Visible = False
     Case 2
End Select

End Sub

Private Sub txtId_Cedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtId_Nombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Cedula"
  gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
  gBusquedas.Orden = "Cedula"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtId_Cedula.Text = gBusquedas.Resultado
  txtId_Nombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtId_Cedula_LostFocus()
txtId_Nombre.Text = fxNombre(txtId_Cedula)
End Sub



Private Sub txtId_Nombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Nombre"
  gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtId_Cedula.Text = gBusquedas.Resultado
  txtId_Nombre.Text = gBusquedas.Resultado2
End If

End Sub



Private Sub txtMnt_Hasta_GotFocus()
On Error GoTo vError
  txtMnt_Hasta.Text = CCur(txtMnt_Hasta.Text)
vError:
End Sub

Private Sub txtMnt_Hasta_LostFocus()
On Error GoTo vError
  txtMnt_Hasta.Text = Format(CCur(txtMnt_Hasta.Text), "Standard")
vError:
End Sub

Private Sub txtMnt_Inicio_GotFocus()
On Error GoTo vError
  txtMnt_Inicio.Text = CCur(txtMnt_Inicio.Text)
vError:
End Sub

Private Sub txtMnt_Inicio_LostFocus()
On Error GoTo vError
  txtMnt_Inicio.Text = Format(CCur(txtMnt_Inicio.Text), "Standard")
vError:
End Sub


Private Sub txtSC_Mnt_Hasta_GotFocus()
On Error GoTo vError
  txtSC_Mnt_Hasta.Text = CCur(txtSC_Mnt_Hasta.Text)
vError:
End Sub

Private Sub txtSC_Mnt_Hasta_LostFocus()
On Error GoTo vError
  txtSC_Mnt_Hasta.Text = Format(CCur(txtSC_Mnt_Hasta.Text), "Standard")
vError:
End Sub

Private Sub txtSC_Mnt_Inicio_GotFocus()
On Error GoTo vError
  txtSC_Mnt_Inicio.Text = CCur(txtSC_Mnt_Inicio.Text)
vError:
End Sub

Private Sub txtSC_Mnt_Inicio_LostFocus()
On Error GoTo vError
  txtSC_Mnt_Inicio.Text = Format(CCur(txtSC_Mnt_Inicio.Text), "Standard")
vError:
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
 
If vPaso Or Col = 1 Then Exit Sub

 vGrid.Row = Row
 vGrid.Col = 3

  Call sbComprobanteSF(vGrid.Text)
  
End Sub

