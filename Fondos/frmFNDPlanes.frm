VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmFNDPlanes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Planes"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   Icon            =   "frmFNDPlanes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   9525
   Begin XtremeSuiteControls.CheckBox chkFiltra 
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   159
      Top             =   1200
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Auto Gestionables"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7575
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   9255
      _Version        =   1572864
      _ExtentX        =   16325
      _ExtentY        =   13361
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
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "tcAux"
      Item(1).Caption =   "Retiros"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "vGrid"
      Item(1).Control(1)=   "Label7(0)"
      Item(1).Control(2)=   "Label8(2)"
      Item(1).Control(3)=   "Label7(1)"
      Item(2).Caption =   "(+/-) Puntos"
      Item(2).ControlCount=   5
      Item(2).Control(0)=   "vaGrid"
      Item(2).Control(1)=   "Label7(2)"
      Item(2).Control(2)=   "scRegla"
      Item(2).Control(3)=   "gbRegla"
      Item(2).Control(4)=   "lswR"
      Item(3).Caption =   "Tasas Aplicadas"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "vhGrid"
      Item(3).Control(1)=   "Label8(4)"
      Item(4).Caption =   "Destinos"
      Item(4).ControlCount=   2
      Item(4).Control(0)=   "lswDestinos"
      Item(4).Control(1)=   "Label7(3)"
      Begin XtremeSuiteControls.ListView lswR 
         Height          =   2055
         Left            =   -70000
         TabIndex        =   144
         Top             =   360
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   3625
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
      Begin XtremeSuiteControls.GroupBox gbRegla 
         Height          =   2295
         Left            =   -70000
         TabIndex        =   145
         Top             =   2400
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   4048
         _StockProps     =   79
         Caption         =   "Regla:"
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
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnRegla 
            Height          =   375
            Index           =   0
            Left            =   1440
            TabIndex        =   152
            Top             =   1800
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Nueva"
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
            Picture         =   "frmFNDPlanes.frx":000C
         End
         Begin XtremeSuiteControls.ComboBox cboR_Tipo 
            Height          =   330
            Left            =   3960
            TabIndex        =   149
            Top             =   600
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.DateTimePicker dtpR_Fecha 
            Height          =   315
            Left            =   7560
            TabIndex        =   150
            Top             =   600
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2561
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
         Begin XtremeSuiteControls.FlatEdit txtR_Justifica 
            Height          =   675
            Left            =   1440
            TabIndex        =   151
            Top             =   960
            Width           =   7575
            _Version        =   1572864
            _ExtentX        =   13361
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
         Begin XtremeSuiteControls.PushButton btnRegla 
            Height          =   375
            Index           =   1
            Left            =   2640
            TabIndex        =   153
            Top             =   1800
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Guardar"
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
            Picture         =   "frmFNDPlanes.frx":063E
         End
         Begin XtremeSuiteControls.PushButton btnRegla_Activa 
            Height          =   375
            Index           =   2
            Left            =   5880
            TabIndex        =   154
            Top             =   1800
            Width           =   2535
            _Version        =   1572864
            _ExtentX        =   4471
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Activar Regla de Tasas"
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
            Picture         =   "frmFNDPlanes.frx":0D6F
         End
         Begin XtremeSuiteControls.FlatEdit txtR_Id 
            Height          =   315
            Left            =   1440
            TabIndex        =   156
            Top             =   240
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtR_Vigente 
            Height          =   315
            Left            =   1440
            TabIndex        =   157
            Top             =   600
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   158
            Top             =   600
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Vigente:"
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
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   155
            Top             =   240
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Id:"
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
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   148
            Top             =   960
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Justificación"
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
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   1
            Left            =   6120
            TabIndex        =   147
            Top             =   600
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha Referencia"
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
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   146
            Top             =   600
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo de Regla"
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
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   7095
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
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
         ItemCount       =   5
         Item(0).Caption =   "Principal"
         Item(0).ControlCount=   13
         Item(0).Control(0)=   "cboMoneda"
         Item(0).Control(1)=   "txtUltContrato"
         Item(0).Control(2)=   "txtNotas"
         Item(0).Control(3)=   "cboGrupo"
         Item(0).Control(4)=   "Label6(1)"
         Item(0).Control(5)=   "Label6(0)"
         Item(0).Control(6)=   "Label5(1)"
         Item(0).Control(7)=   "Label2(1)"
         Item(0).Control(8)=   "GroupBox1(0)"
         Item(0).Control(9)=   "GroupBox1(1)"
         Item(0).Control(10)=   "cboEstado"
         Item(0).Control(11)=   "Label2(3)"
         Item(0).Control(12)=   "cboTipoPlan"
         Item(1).Caption =   "Paso 2"
         Item(1).ControlCount=   11
         Item(1).Control(0)=   "chkSINPE"
         Item(1).Control(1)=   "chkRequiereBeneficiario"
         Item(1).Control(2)=   "chkControlaSaldo"
         Item(1).Control(3)=   "chkCDP"
         Item(1).Control(4)=   "chkCuentaMaestra"
         Item(1).Control(5)=   "GroupBox2(0)"
         Item(1).Control(6)=   "GroupBox2(1)"
         Item(1).Control(7)=   "GroupBox2(2)"
         Item(1).Control(8)=   "chkCDP_PagaCupones"
         Item(1).Control(9)=   "txtTasaMargenNegociacion"
         Item(1).Control(10)=   "Label11(5)"
         Item(2).Caption =   "Paso 3"
         Item(2).ControlCount=   8
         Item(2).Control(0)=   "chkPagoTercero"
         Item(2).Control(1)=   "chkVisibleEC"
         Item(2).Control(2)=   "chkLiqSocio"
         Item(2).Control(3)=   "GroupBox1(2)"
         Item(2).Control(4)=   "GroupBox1(3)"
         Item(2).Control(5)=   "chkLiqPlanesAhorros"
         Item(2).Control(6)=   "gbWebSite"
         Item(2).Control(7)=   "GroupBox3"
         Item(3).Caption =   "Plan Contable"
         Item(3).ControlCount=   21
         Item(3).Control(0)=   "txtCuentaImpuestos"
         Item(3).Control(1)=   "txtCuentaImpuestosDesc"
         Item(3).Control(2)=   "txtCuentaRend"
         Item(3).Control(3)=   "txtCuentaRendDesc"
         Item(3).Control(4)=   "txtCuentaIngRetirosDesc"
         Item(3).Control(5)=   "txtCuentaIngRetiros"
         Item(3).Control(6)=   "txtCuentaGstComisionDesc"
         Item(3).Control(7)=   "txtCuentaGstComision"
         Item(3).Control(8)=   "txtCuentaIngComisionDesc"
         Item(3).Control(9)=   "txtCuentaIngComision"
         Item(3).Control(10)=   "txtCuentaDesc"
         Item(3).Control(11)=   "txtCuentaCod"
         Item(3).Control(12)=   "txtCuentaGasDesc"
         Item(3).Control(13)=   "txtCuentaGasto"
         Item(3).Control(14)=   "Label5(9)"
         Item(3).Control(15)=   "Label5(6)"
         Item(3).Control(16)=   "Label5(5)"
         Item(3).Control(17)=   "Label5(4)"
         Item(3).Control(18)=   "Label5(3)"
         Item(3).Control(19)=   "Label5(2)"
         Item(3).Control(20)=   "Label5(0)"
         Item(4).Caption =   "Estados/Vencimientos"
         Item(4).ControlCount=   7
         Item(4).Control(0)=   "lswEstados"
         Item(4).Control(1)=   "ShortcutCaption1"
         Item(4).Control(2)=   "ShortcutCaption2(0)"
         Item(4).Control(3)=   "ShortcutCaption2(1)"
         Item(4).Control(4)=   "gbVencimientos"
         Item(4).Control(5)=   "lswVence_Plazos"
         Item(4).Control(6)=   "chkVence_Plazo_Sol"
         Begin XtremeSuiteControls.ListView lswEstados 
            Height          =   2655
            Left            =   -69880
            TabIndex        =   128
            Top             =   1080
            Visible         =   0   'False
            Width           =   4335
            _Version        =   1572864
            _ExtentX        =   7646
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
            Checkboxes      =   -1  'True
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.ListView lswVence_Plazos 
            Height          =   2655
            Left            =   -65320
            TabIndex        =   130
            Top             =   1080
            Visible         =   0   'False
            Width           =   4575
            _Version        =   1572864
            _ExtentX        =   8070
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
            Checkboxes      =   -1  'True
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.GroupBox GroupBox3 
            Height          =   2535
            Left            =   -65560
            TabIndex        =   172
            Top             =   1920
            Visible         =   0   'False
            Width           =   4695
            _Version        =   1572864
            _ExtentX        =   8281
            _ExtentY        =   4471
            _StockProps     =   79
            Caption         =   "Sinpe y Patrimonio"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            BorderStyle     =   1
            Begin XtremeSuiteControls.ComboBox cboTipoPatrimonio 
               Height          =   315
               Left            =   1320
               TabIndex        =   173
               Top             =   1440
               Width           =   1455
               _Version        =   1572864
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
            Begin XtremeSuiteControls.CheckBox chkEnlazaPatrimonio 
               Height          =   255
               Left            =   480
               TabIndex        =   174
               Top             =   1200
               Width           =   4215
               _Version        =   1572864
               _ExtentX        =   7429
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Enlazado a Patrimonio?"
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
            Begin XtremeSuiteControls.CheckBox chkPat_Unifica 
               Height          =   375
               Left            =   1320
               TabIndex        =   175
               Top             =   1800
               Width           =   2895
               _Version        =   1572864
               _ExtentX        =   5101
               _ExtentY        =   656
               _StockProps     =   79
               Caption         =   "Unificar en Estado de Cuenta en Patrimonio?"
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
            Begin XtremeSuiteControls.CheckBox chkSINPE_Mov 
               Height          =   255
               Left            =   480
               TabIndex        =   176
               Top             =   360
               Width           =   2895
               _Version        =   1572864
               _ExtentX        =   5101
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Permite Movimiento SINPE ?"
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
            Begin XtremeSuiteControls.ComboBox cboSinpeTipos 
               Height          =   315
               Left            =   1320
               TabIndex        =   177
               Top             =   720
               Width           =   1455
               _Version        =   1572864
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
         End
         Begin XtremeSuiteControls.GroupBox gbWebSite 
            Height          =   2535
            Left            =   -69760
            TabIndex        =   163
            Top             =   1920
            Visible         =   0   'False
            Width           =   4215
            _Version        =   1572864
            _ExtentX        =   7435
            _ExtentY        =   4471
            _StockProps     =   79
            Caption         =   "Auto Gestión"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            BorderStyle     =   1
            Begin XtremeSuiteControls.CheckBox chkWebSite 
               Height          =   375
               Left            =   480
               TabIndex        =   164
               Top             =   240
               Width           =   3135
               _Version        =   1572864
               _ExtentX        =   5530
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Disponible en el WebSite/App ?"
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
            Begin XtremeSuiteControls.CheckBox chkLiquidaWebSite 
               Height          =   255
               Left            =   480
               TabIndex        =   165
               Top             =   1800
               Width           =   2895
               _Version        =   1572864
               _ExtentX        =   5106
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Permite Liquidar"
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
            Begin XtremeSuiteControls.CheckBox chkVenceWeb 
               Height          =   255
               Left            =   3240
               TabIndex        =   166
               Top             =   600
               Width           =   2655
               _Version        =   1572864
               _ExtentX        =   4683
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Vence?"
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
            Begin XtremeSuiteControls.DateTimePicker dtpVenceWeb 
               Height          =   315
               Left            =   1320
               TabIndex        =   167
               Top             =   600
               Width           =   1455
               _Version        =   1572864
               _ExtentX        =   2561
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
            Begin XtremeSuiteControls.CheckBox chkRetiroParcial 
               Height          =   375
               Left            =   1320
               TabIndex        =   168
               Top             =   2040
               Width           =   2895
               _Version        =   1572864
               _ExtentX        =   5106
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Permite Retiros Parciales?"
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
            Begin XtremeSuiteControls.PushButton btnFechaCorte 
               Height          =   315
               Left            =   2775
               TabIndex        =   169
               ToolTipText     =   "Corrige la Fecha de Corte para Todos los Contratos bajo este Plan"
               Top             =   600
               Width           =   375
               _Version        =   1572864
               _ExtentX        =   656
               _ExtentY        =   547
               _StockProps     =   79
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
               FlatStyle       =   -1  'True
               Appearance      =   16
               Picture         =   "frmFNDPlanes.frx":1496
            End
            Begin XtremeSuiteControls.CheckBox chkWebCrea 
               Height          =   255
               Left            =   480
               TabIndex        =   170
               Top             =   1080
               Width           =   3135
               _Version        =   1572864
               _ExtentX        =   5530
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Crear Nuevos Contratos"
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
            Begin XtremeSuiteControls.CheckBox chkWebModifica 
               Height          =   255
               Left            =   480
               TabIndex        =   171
               Top             =   1440
               Width           =   3135
               _Version        =   1572864
               _ExtentX        =   5530
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Modifica Cuota "
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
         End
         Begin XtremeSuiteControls.CheckBox chkVence_Plazo_Sol 
            Height          =   210
            Left            =   -65080
            TabIndex        =   138
            Top             =   600
            Visible         =   0   'False
            Width           =   210
            _Version        =   1572864
            _ExtentX        =   370
            _ExtentY        =   370
            _StockProps     =   79
            BackColor       =   -2147483633
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.GroupBox gbVencimientos 
            Height          =   2895
            Left            =   -69880
            TabIndex        =   132
            Top             =   3960
            Visible         =   0   'False
            Width           =   9135
            _Version        =   1572864
            _ExtentX        =   16113
            _ExtentY        =   5106
            _StockProps     =   79
            Caption         =   "Acciones para el Vencimiento del Plan:"
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
            Begin XtremeSuiteControls.ComboBox cboVence_Accion 
               Height          =   330
               Left            =   1920
               TabIndex        =   133
               Top             =   600
               Width           =   2415
               _Version        =   1572864
               _ExtentX        =   4260
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
            Begin XtremeSuiteControls.FlatEdit txtVence_Plan 
               Height          =   315
               Left            =   1920
               TabIndex        =   136
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   1200
               Width           =   1215
               _Version        =   1572864
               _ExtentX        =   2138
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
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtVence_PlanDesc 
               Height          =   315
               Left            =   3120
               TabIndex        =   137
               Top             =   1200
               Width           =   5895
               _Version        =   1572864
               _ExtentX        =   10398
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
            Begin XtremeSuiteControls.CheckBox chkVence_Renueva 
               Height          =   210
               Left            =   4800
               TabIndex        =   139
               Top             =   600
               Width           =   3210
               _Version        =   1572864
               _ExtentX        =   5662
               _ExtentY        =   370
               _StockProps     =   79
               Caption         =   "Renueva al Vencimiento?"
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
               Appearance      =   21
            End
            Begin XtremeSuiteControls.CheckBox chkVence_AplTasaCntVencidos 
               Height          =   210
               Left            =   1920
               TabIndex        =   178
               Top             =   1800
               Width           =   5250
               _Version        =   1572864
               _ExtentX        =   9260
               _ExtentY        =   370
               _StockProps     =   79
               Caption         =   "Aplica Tasa de Rendimiento a Contratos Vencidos?"
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
               Appearance      =   21
            End
            Begin XtremeSuiteControls.CheckBox chkVence_ActivaControlVencimiento 
               Height          =   210
               Left            =   1920
               TabIndex        =   179
               Top             =   2160
               Width           =   5250
               _Version        =   1572864
               _ExtentX        =   9260
               _ExtentY        =   370
               _StockProps     =   79
               Caption         =   "Activar el Control de Vencimiento?"
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
               Appearance      =   21
            End
            Begin XtremeSuiteControls.Label lblAccion 
               Height          =   495
               Index           =   1
               Left            =   240
               TabIndex        =   135
               Top             =   1200
               Width           =   1335
               _Version        =   1572864
               _ExtentX        =   2355
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "Plan para Traslado:"
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
               WordWrap        =   -1  'True
            End
            Begin XtremeSuiteControls.Label lblAccion 
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   134
               Top             =   600
               Width           =   1455
               _Version        =   1572864
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Acción:"
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
               WordWrap        =   -1  'True
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   2055
            Index           =   0
            Left            =   -69880
            TabIndex        =   98
            Top             =   1920
            Visible         =   0   'False
            Width           =   8895
            _Version        =   1572864
            _ExtentX        =   15684
            _ExtentY        =   3619
            _StockProps     =   79
            Caption         =   "Rendimientos"
            ForeColor       =   4210752
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
            Begin XtremeSuiteControls.ComboBox cboBaseCalculo 
               Height          =   312
               Left            =   2760
               TabIndex        =   112
               Top             =   720
               Width           =   1092
               _Version        =   1572864
               _ExtentX        =   1931
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
            Begin XtremeSuiteControls.CheckBox chkRendimientos 
               Height          =   372
               Left            =   360
               TabIndex        =   113
               Top             =   240
               Width           =   3492
               _Version        =   1572864
               _ExtentX        =   6159
               _ExtentY        =   656
               _StockProps     =   79
               Caption         =   "Aplica cálculo de rendimientos ?"
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
               Value           =   1
            End
            Begin XtremeSuiteControls.CheckBox chkTasaFluctuante 
               Height          =   252
               Left            =   3960
               TabIndex        =   114
               Top             =   720
               Width           =   4332
               _Version        =   1572864
               _ExtentX        =   7641
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Solicita al Usuario (Tasa Fluctuante)"
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
            Begin XtremeSuiteControls.CheckBox chkCapiltalizaRendimiento 
               Height          =   252
               Left            =   3960
               TabIndex        =   115
               Top             =   960
               Width           =   4332
               _Version        =   1572864
               _ExtentX        =   7641
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Capitaliza Rendimientos?"
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
            Begin XtremeSuiteControls.CheckBox chkUtilizaTBP 
               Height          =   252
               Left            =   3960
               TabIndex        =   116
               Top             =   1320
               Width           =   4332
               _Version        =   1572864
               _ExtentX        =   7641
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Utiliza Tasa Básica Pasiva "
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
            Begin XtremeSuiteControls.CheckBox chkRenuevaTasa 
               Height          =   252
               Left            =   3960
               TabIndex        =   117
               Top             =   1680
               Width           =   4332
               _Version        =   1572864
               _ExtentX        =   7641
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Renueva tasa al vencimiento"
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
            Begin XtremeSuiteControls.FlatEdit txtTasaBase 
               Height          =   312
               Left            =   2760
               TabIndex        =   118
               Top             =   1320
               Width           =   1092
               _Version        =   1572864
               _ExtentX        =   1926
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtNuevaTasa 
               Height          =   312
               Left            =   2760
               TabIndex        =   119
               Top             =   1680
               Width           =   1092
               _Version        =   1572864
               _ExtentX        =   1926
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.CheckBox chkRendimientoAuto 
               Height          =   252
               Left            =   3960
               TabIndex        =   123
               Top             =   480
               Width           =   4692
               _Version        =   1572864
               _ExtentX        =   8276
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Aplica Rendimiento Automático al Cierre de mes?"
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
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               Caption         =   "Tasa al vencimiento"
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
               Left            =   960
               TabIndex        =   122
               Top             =   1680
               Width           =   1932
            End
            Begin VB.Label Label11 
               Caption         =   "Tasa Base"
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
               Left            =   960
               TabIndex        =   121
               Top             =   1320
               Width           =   1212
            End
            Begin VB.Label Label11 
               Caption         =   "Base de cálculo"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   960
               TabIndex        =   120
               Top             =   720
               Width           =   1695
            End
         End
         Begin XtremeSuiteControls.CheckBox chkCDP 
            Height          =   372
            Left            =   -69520
            TabIndex        =   55
            Top             =   360
            Visible         =   0   'False
            Width           =   4332
            _Version        =   1572864
            _ExtentX        =   7641
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Administrar como C.D.P. (Certificado a Plazo)"
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
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   1095
            Index           =   0
            Left            =   360
            TabIndex        =   26
            Top             =   3120
            Width           =   8295
            _Version        =   1572864
            _ExtentX        =   14626
            _ExtentY        =   1926
            _StockProps     =   79
            Caption         =   "Información de Aportación / Mínimos"
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
            Begin XtremeSuiteControls.ComboBox cboTipoAporte 
               Height          =   312
               Left            =   1440
               TabIndex        =   36
               Top             =   360
               Width           =   1452
               _Version        =   1572864
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
            Begin XtremeSuiteControls.ComboBox cboPlazo 
               Height          =   312
               Left            =   2040
               TabIndex        =   41
               Top             =   720
               Width           =   852
               _Version        =   1572864
               _ExtentX        =   1508
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
            Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
               Height          =   312
               Left            =   3600
               TabIndex        =   73
               Top             =   360
               Width           =   1332
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtMonto 
               Height          =   312
               Left            =   3600
               TabIndex        =   74
               Top             =   720
               Width           =   1332
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtInversion 
               Height          =   312
               Left            =   6000
               TabIndex        =   75
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPlazo 
               Height          =   330
               Left            =   1440
               TabIndex        =   76
               Top             =   720
               Width           =   612
               _Version        =   1572864
               _ExtentX        =   1080
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
            Begin VB.Label Label3 
               Caption         =   "(Porcentaje Referencia)"
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
               Index           =   14
               Left            =   5040
               TabIndex        =   54
               Top             =   360
               Width           =   2772
            End
            Begin VB.Label Label4 
               Caption         =   "[ % ]"
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
               Left            =   3000
               TabIndex        =   53
               Top             =   360
               Width           =   1092
            End
            Begin VB.Label Label4 
               Caption         =   "Inversion"
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
               Left            =   5040
               TabIndex        =   33
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label3 
               Caption         =   "Deducción"
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
               Index           =   13
               Left            =   480
               TabIndex        =   32
               Top             =   360
               Width           =   972
            End
            Begin VB.Label Label3 
               Caption         =   "Plazo"
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
               Left            =   480
               TabIndex        =   31
               Top             =   720
               Width           =   732
            End
            Begin VB.Label Label4 
               Caption         =   "Cuota"
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
               Left            =   3000
               TabIndex        =   30
               Top             =   720
               Width           =   1092
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   2175
            Index           =   1
            Left            =   360
            TabIndex        =   27
            Top             =   4560
            Width           =   8295
            _Version        =   1572864
            _ExtentX        =   14626
            _ExtentY        =   3831
            _StockProps     =   79
            Caption         =   "Método de Recaudación del Plan"
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
            Begin XtremeSuiteControls.PushButton btnConfig_Recaudo 
               Height          =   330
               Left            =   7560
               TabIndex        =   126
               ToolTipText     =   "Crea Automaticamente el Código para Recaudacion (Retención Asociada)"
               Top             =   360
               Width           =   372
               _Version        =   1572864
               _ExtentX        =   656
               _ExtentY        =   582
               _StockProps     =   79
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               Picture         =   "frmFNDPlanes.frx":1B89
            End
            Begin XtremeSuiteControls.CheckBox chkDeducirPlanilla 
               Height          =   372
               Left            =   1800
               TabIndex        =   60
               Top             =   720
               Width           =   4332
               _Version        =   1572864
               _ExtentX        =   7641
               _ExtentY        =   656
               _StockProps     =   79
               Caption         =   "Deducir este plan x planillas ?"
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
               Value           =   1
            End
            Begin XtremeSuiteControls.CheckBox chkGeneraMora 
               Height          =   252
               Left            =   2280
               TabIndex        =   61
               Top             =   1080
               Width           =   4332
               _Version        =   1572864
               _ExtentX        =   7641
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Genera cuotas atrasadas ante no pago ?"
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
            Begin XtremeSuiteControls.CheckBox chkDeducIndependiente 
               Height          =   252
               Left            =   2280
               TabIndex        =   62
               Top             =   1320
               Width           =   4332
               _Version        =   1572864
               _ExtentX        =   7641
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Se deduce de forma independiente?"
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
            Begin XtremeSuiteControls.FlatEdit txtAseCod 
               Height          =   312
               Left            =   1440
               TabIndex        =   68
               Top             =   360
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtAseDesc 
               Height          =   312
               Left            =   2640
               TabIndex        =   69
               Top             =   360
               Width           =   4812
               _Version        =   1572864
               _ExtentX        =   8488
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCodigoDeduc 
               Height          =   312
               Left            =   1440
               TabIndex        =   70
               Top             =   1680
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label9 
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
               Height          =   252
               Left            =   480
               TabIndex        =   35
               Top             =   360
               Width           =   852
            End
            Begin VB.Label Label3 
               Caption         =   "Código de Deducción (independiente)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   252
               Index           =   9
               Left            =   3240
               TabIndex        =   34
               Top             =   1680
               Width           =   3972
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   1335
            Index           =   2
            Left            =   -69760
            TabIndex        =   28
            Top             =   6000
            Visible         =   0   'False
            Width           =   8415
            _Version        =   1572864
            _ExtentX        =   14843
            _ExtentY        =   2355
            _StockProps     =   79
            Caption         =   "Otros"
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
            Begin XtremeSuiteControls.FlatEdit txtContratosPersona 
               Height          =   315
               Left            =   3360
               TabIndex        =   80
               Top             =   360
               Width           =   615
               _Version        =   1572864
               _ExtentX        =   1080
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtComVentaTasa 
               Height          =   315
               Left            =   3360
               TabIndex        =   81
               Top             =   720
               Width           =   615
               _Version        =   1572864
               _ExtentX        =   1080
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtComVentaMonto 
               Height          =   315
               Left            =   6720
               TabIndex        =   82
               Top             =   720
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label3 
               Caption         =   "No. Contratos Activos por Persona ?"
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
               Left            =   120
               TabIndex        =   51
               Top             =   360
               Width           =   3255
            End
            Begin VB.Label Label3 
               Caption         =   "Monto de Comisión s/Venta"
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
               Index           =   6
               Left            =   4680
               TabIndex        =   50
               Top             =   585
               Width           =   1575
            End
            Begin VB.Label Label3 
               Caption         =   "Comisión Venta s/Inversion C.D.P."
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
               Left            =   240
               TabIndex        =   49
               Top             =   720
               Width           =   3135
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "%"
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
               Left            =   3960
               TabIndex        =   48
               Top             =   720
               Width           =   255
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   1095
            Index           =   3
            Left            =   -69760
            TabIndex        =   29
            Top             =   4680
            Visible         =   0   'False
            Width           =   8415
            _Version        =   1572864
            _ExtentX        =   14838
            _ExtentY        =   1926
            _StockProps     =   79
            Caption         =   "Comisiones de Administración / Impuestos"
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
            Begin XtremeSuiteControls.FlatEdit txtTasaComAportes 
               Height          =   315
               Left            =   3360
               TabIndex        =   77
               Top             =   360
               Width           =   615
               _Version        =   1572864
               _ExtentX        =   1080
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtTasaComRend 
               Height          =   315
               Left            =   3360
               TabIndex        =   78
               Top             =   720
               Width           =   615
               _Version        =   1572864
               _ExtentX        =   1080
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtImpuestosRend 
               Height          =   312
               Left            =   7200
               TabIndex        =   79
               Top             =   360
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.CheckBox chkRentaGlobal 
               Height          =   495
               Left            =   4680
               TabIndex        =   97
               Top             =   720
               Width           =   3135
               _Version        =   1572864
               _ExtentX        =   5530
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "Utiliza Renta Global?"
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
               Alignment       =   1
            End
            Begin VB.Label Label3 
               Caption         =   "Comisión sobre Aportes"
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
               Left            =   600
               TabIndex        =   47
               Top             =   360
               Width           =   2655
            End
            Begin VB.Label Label3 
               Caption         =   "Comisión sobre Rendimientos"
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
               Left            =   600
               TabIndex        =   46
               Top             =   720
               Width           =   2775
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "%"
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
               Index           =   4
               Left            =   3960
               TabIndex        =   45
               Top             =   360
               Width           =   255
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "%"
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
               Index           =   5
               Left            =   3960
               TabIndex        =   44
               Top             =   720
               Width           =   255
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "%"
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
               Index           =   11
               Left            =   7800
               TabIndex        =   43
               Top             =   360
               Width           =   252
            End
            Begin VB.Label Label3 
               Caption         =   "Impuesto s/ Rendimientos"
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
               Index           =   12
               Left            =   4680
               TabIndex        =   42
               Top             =   360
               Width           =   2532
            End
         End
         Begin XtremeSuiteControls.ComboBox cboEstado 
            Height          =   312
            Left            =   1800
            TabIndex        =   37
            Top             =   2280
            Width           =   1452
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboGrupo 
            Height          =   312
            Left            =   1800
            TabIndex        =   38
            Top             =   840
            Width           =   6012
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboMoneda 
            Height          =   312
            Left            =   4440
            TabIndex        =   40
            Top             =   2280
            Width           =   3372
            _Version        =   1572864
            _ExtentX        =   5953
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
         Begin XtremeSuiteControls.CheckBox chkControlaSaldo 
            Height          =   255
            Left            =   -69520
            TabIndex        =   56
            Top             =   1560
            Visible         =   0   'False
            Width           =   4335
            _Version        =   1572864
            _ExtentX        =   7646
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "El Plan lleva control de acumulados ""Saldos"" "
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
         Begin XtremeSuiteControls.CheckBox chkSINPE 
            Height          =   375
            Left            =   -64960
            TabIndex        =   57
            Top             =   1080
            Visible         =   0   'False
            Width           =   4335
            _Version        =   1572864
            _ExtentX        =   7641
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Cuenta SINPE ? Código Interno: "
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
         Begin XtremeSuiteControls.CheckBox chkCuentaMaestra 
            Height          =   372
            Left            =   -64960
            TabIndex        =   58
            Top             =   360
            Visible         =   0   'False
            Width           =   4332
            _Version        =   1572864
            _ExtentX        =   7641
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Se administra como cuenta maestra ?"
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
         Begin XtremeSuiteControls.CheckBox chkRequiereBeneficiario 
            Height          =   372
            Left            =   -64960
            TabIndex        =   59
            Top             =   720
            Visible         =   0   'False
            Width           =   4212
            _Version        =   1572864
            _ExtentX        =   7429
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Requiere Beneficiario?"
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
         Begin XtremeSuiteControls.CheckBox chkLiqSocio 
            Height          =   375
            Left            =   -69280
            TabIndex        =   63
            Top             =   360
            Visible         =   0   'False
            Width           =   7335
            _Version        =   1572864
            _ExtentX        =   12933
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Puede liquidarse en la Liquidación General de Afiliaciones?"
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
         Begin XtremeSuiteControls.CheckBox chkPagoTercero 
            Height          =   255
            Left            =   -69280
            TabIndex        =   64
            Top             =   1080
            Visible         =   0   'False
            Width           =   6375
            _Version        =   1572864
            _ExtentX        =   11239
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite realizar Retiros/Liq. con Pago a Terceros ?"
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
         Begin XtremeSuiteControls.CheckBox chkVisibleEC 
            Height          =   255
            Left            =   -69280
            TabIndex        =   65
            Top             =   1440
            Visible         =   0   'False
            Width           =   4335
            _Version        =   1572864
            _ExtentX        =   7646
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Plan es visible en el Estado de Cuenta ?"
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
            Value           =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   1032
            Left            =   1800
            TabIndex        =   71
            Top             =   1200
            Width           =   6012
            _Version        =   1572864
            _ExtentX        =   10604
            _ExtentY        =   1820
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
         Begin XtremeSuiteControls.FlatEdit txtUltContrato 
            Height          =   312
            Left            =   4440
            TabIndex        =   72
            Top             =   480
            Width           =   3372
            _Version        =   1572864
            _ExtentX        =   5948
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaCod 
            Height          =   312
            Left            =   -68680
            TabIndex        =   83
            Top             =   960
            Visible         =   0   'False
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
            Height          =   312
            Left            =   -66880
            TabIndex        =   84
            Top             =   960
            Visible         =   0   'False
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaRend 
            Height          =   312
            Left            =   -68680
            TabIndex        =   85
            Top             =   1680
            Visible         =   0   'False
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaRendDesc 
            Height          =   312
            Left            =   -66880
            TabIndex        =   86
            Top             =   1680
            Visible         =   0   'False
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaGasto 
            Height          =   312
            Left            =   -68680
            TabIndex        =   87
            Top             =   2400
            Visible         =   0   'False
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaGasDesc 
            Height          =   312
            Left            =   -66880
            TabIndex        =   88
            Top             =   2400
            Visible         =   0   'False
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaIngComision 
            Height          =   312
            Left            =   -68680
            TabIndex        =   89
            Top             =   3120
            Visible         =   0   'False
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaIngComisionDesc 
            Height          =   312
            Left            =   -66880
            TabIndex        =   90
            Top             =   3120
            Visible         =   0   'False
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaIngRetiros 
            Height          =   312
            Left            =   -68680
            TabIndex        =   91
            Top             =   3840
            Visible         =   0   'False
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaIngRetirosDesc 
            Height          =   312
            Left            =   -66880
            TabIndex        =   92
            Top             =   3840
            Visible         =   0   'False
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaGstComision 
            Height          =   312
            Left            =   -68680
            TabIndex        =   93
            Top             =   4560
            Visible         =   0   'False
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaGstComisionDesc 
            Height          =   312
            Left            =   -66880
            TabIndex        =   94
            Top             =   4560
            Visible         =   0   'False
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaImpuestos 
            Height          =   312
            Left            =   -68680
            TabIndex        =   95
            Top             =   5280
            Visible         =   0   'False
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaImpuestosDesc 
            Height          =   312
            Left            =   -66880
            TabIndex        =   96
            Top             =   5280
            Visible         =   0   'False
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   1455
            Index           =   1
            Left            =   -69880
            TabIndex        =   99
            Top             =   4080
            Visible         =   0   'False
            Width           =   8895
            _Version        =   1572864
            _ExtentX        =   15684
            _ExtentY        =   2561
            _StockProps     =   79
            Caption         =   "Garantía Back to Back"
            ForeColor       =   4210752
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
            Begin XtremeSuiteControls.CheckBox chkGarantia 
               Height          =   252
               Left            =   480
               TabIndex        =   105
               Top             =   360
               Width           =   4932
               _Version        =   1572864
               _ExtentX        =   8700
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Aplica como Garantía de Créditos / Back to Back?"
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
               TextAlignment   =   4
               Appearance      =   16
               Value           =   1
            End
            Begin XtremeSuiteControls.CheckBox chkBacktoBackIntegra 
               Height          =   252
               Left            =   3960
               TabIndex        =   106
               Top             =   720
               Width           =   4332
               _Version        =   1572864
               _ExtentX        =   7641
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Integra al disponible con otros planes"
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
            Begin XtremeSuiteControls.FlatEdit txtBacktoBackDisponible 
               Height          =   312
               Left            =   2760
               TabIndex        =   107
               Top             =   720
               Width           =   1092
               _Version        =   1572864
               _ExtentX        =   1926
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtBacktoBackTasaAdCredito 
               Height          =   312
               Left            =   2760
               TabIndex        =   108
               Top             =   1080
               Width           =   1092
               _Version        =   1572864
               _ExtentX        =   1926
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label11 
               Caption         =   "Pts Adicionales"
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
               Left            =   960
               TabIndex        =   111
               Top             =   1080
               Width           =   1572
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               Caption         =   "[ + Crédito]"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   3
               Left            =   3960
               TabIndex        =   110
               Top             =   1080
               Width           =   3615
            End
            Begin VB.Label Label11 
               Caption         =   "% Disponible"
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
               Left            =   960
               TabIndex        =   109
               Top             =   720
               Width           =   1212
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   1455
            Index           =   2
            Left            =   -69880
            TabIndex        =   100
            Top             =   5760
            Visible         =   0   'False
            Width           =   8895
            _Version        =   1572864
            _ExtentX        =   15684
            _ExtentY        =   2561
            _StockProps     =   79
            Caption         =   "Cajas y Formas de Pago"
            ForeColor       =   4210752
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
            Begin XtremeSuiteControls.CheckBox chkMovCajas 
               Height          =   252
               Left            =   360
               TabIndex        =   101
               Top             =   360
               Width           =   4332
               _Version        =   1572864
               _ExtentX        =   7641
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Movimientos en Cajas ?"
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
               Value           =   1
            End
            Begin XtremeSuiteControls.CheckBox chkCajasRetiros 
               Height          =   255
               Left            =   360
               TabIndex        =   102
               Top             =   720
               Width           =   4455
               _Version        =   1572864
               _ExtentX        =   7858
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Forma de pago en Cajas de Cuentas Corrientes?"
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
            Begin XtremeSuiteControls.CheckBox chkMovEntreFondos 
               Height          =   252
               Left            =   4920
               TabIndex        =   103
               Top             =   360
               Width           =   4332
               _Version        =   1572864
               _ExtentX        =   7641
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Movimientos entre Fondos Propios?"
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
            Begin XtremeSuiteControls.CheckBox chkMovEntreFondosTerceros 
               Height          =   252
               Left            =   4920
               TabIndex        =   104
               Top             =   720
               Width           =   4692
               _Version        =   1572864
               _ExtentX        =   8276
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Movimientos entre Fondos con Terceros?"
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
            Begin XtremeSuiteControls.CheckBox chkFP_Servicios 
               Height          =   255
               Left            =   360
               TabIndex        =   124
               Top             =   1080
               Width           =   4455
               _Version        =   1572864
               _ExtentX        =   7858
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Forma de pago en Servicios AutoGestionables?"
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
            Begin XtremeSuiteControls.CheckBox chkFP_POS 
               Height          =   252
               Left            =   4920
               TabIndex        =   125
               Top             =   1080
               Width           =   4332
               _Version        =   1572864
               _ExtentX        =   7641
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Forma de pago en los POS?"
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
         End
         Begin XtremeSuiteControls.ComboBox cboTipoPlan 
            Height          =   330
            Left            =   1800
            TabIndex        =   141
            Top             =   480
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.CheckBox chkLiqPlanesAhorros 
            Height          =   255
            Left            =   -69280
            TabIndex        =   142
            Top             =   720
            Visible         =   0   'False
            Width           =   7335
            _Version        =   1572864
            _ExtentX        =   12938
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Puede liquidarse desde Planes de Ahorros?"
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
         Begin XtremeSuiteControls.CheckBox chkCDP_PagaCupones 
            Height          =   375
            Left            =   -69160
            TabIndex        =   180
            Top             =   720
            Visible         =   0   'False
            Width           =   3375
            _Version        =   1572864
            _ExtentX        =   5953
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Pago de Cupones ?"
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
         Begin XtremeSuiteControls.FlatEdit txtTasaMargenNegociacion 
            Height          =   315
            Left            =   -69160
            TabIndex        =   181
            Top             =   1200
            Visible         =   0   'False
            Width           =   735
            _Version        =   1572864
            _ExtentX        =   1296
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label11 
            Caption         =   "Margen de Negociación"
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
            Left            =   -68320
            TabIndex        =   182
            ToolTipText     =   "Tasa Preferencial: Puntos de Variación aplicables a la Tasa Oficial"
            Top             =   1200
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
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
            Height          =   315
            Index           =   3
            Left            =   840
            TabIndex        =   140
            Top             =   480
            Width           =   975
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   510
            Index           =   1
            Left            =   -65920
            TabIndex        =   131
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   900
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
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   510
            Index           =   0
            Left            =   -64720
            TabIndex        =   129
            Top             =   480
            Visible         =   0   'False
            Width           =   3975
            _Version        =   1572864
            _ExtentX        =   7011
            _ExtentY        =   900
            _StockProps     =   14
            Caption         =   "Solicita Plazo para el Vencimiento"
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
            Height          =   510
            Left            =   -70000
            TabIndex        =   127
            Top             =   480
            Visible         =   0   'False
            Width           =   4095
            _Version        =   1572864
            _ExtentX        =   7223
            _ExtentY        =   900
            _StockProps     =   14
            Caption         =   "Estados de la Persona Permitidos para este Plan"
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
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo"
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
            Index           =   1
            Left            =   840
            TabIndex        =   25
            Top             =   840
            Width           =   972
         End
         Begin VB.Label Label5 
            Caption         =   "Notas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   840
            TabIndex        =   24
            Top             =   1320
            Width           =   852
         End
         Begin VB.Label Label6 
            Caption         =   "Estado"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   840
            TabIndex        =   23
            Top             =   2280
            Width           =   612
         End
         Begin VB.Label Label6 
            Caption         =   "Divisa"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   3600
            TabIndex        =   22
            Top             =   2280
            Width           =   732
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta del Plan de ahorro o inversión"
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
            Left            =   -69640
            TabIndex        =   21
            Top             =   720
            Visible         =   0   'False
            Width           =   4212
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta de Gasto Financiero"
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
            Left            =   -69640
            TabIndex        =   20
            Top             =   2160
            Visible         =   0   'False
            Width           =   3132
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta de Ingreso x Comisión de Administración del Fondo"
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
            Left            =   -69640
            TabIndex        =   19
            Top             =   2880
            Visible         =   0   'False
            Width           =   7095
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta de Gasto x Comisión de venta de contratos"
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
            Left            =   -69640
            TabIndex        =   18
            Top             =   4320
            Visible         =   0   'False
            Width           =   4452
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta de Ingreso x Retiros Anticipados"
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
            Left            =   -69640
            TabIndex        =   17
            Top             =   3600
            Visible         =   0   'False
            Width           =   4452
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta de Registro de Rendimientos del Plan"
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
            Left            =   -69640
            TabIndex        =   16
            Top             =   1440
            Visible         =   0   'False
            Width           =   6495
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta de Impuestos s/Rendimientos"
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
            Left            =   -69640
            TabIndex        =   15
            Top             =   5040
            Visible         =   0   'False
            Width           =   4452
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5772
         Left            =   -67360
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   6612
         _Version        =   524288
         _ExtentX        =   11663
         _ExtentY        =   10181
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
         MaxCols         =   493
         ScrollBars      =   2
         SpreadDesigner  =   "frmFNDPlanes.frx":245A
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vaGrid 
         Height          =   2295
         Left            =   -66640
         TabIndex        =   10
         Top             =   5160
         Visible         =   0   'False
         Width           =   5415
         _Version        =   524288
         _ExtentX        =   9551
         _ExtentY        =   4048
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
         MaxCols         =   494
         ScrollBars      =   2
         SpreadDesigner  =   "frmFNDPlanes.frx":2AFA
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vhGrid 
         Height          =   6255
         Left            =   -69760
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   8895
         _Version        =   524288
         _ExtentX        =   15690
         _ExtentY        =   11033
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
         MaxCols         =   495
         ScrollBars      =   2
         SpreadDesigner  =   "frmFNDPlanes.frx":31DB
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.ListView lswDestinos 
         Height          =   6372
         Left            =   -67360
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   6372
         _ExtentX        =   11245
         _ExtentY        =   11245
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Destinos"
            Object.Width           =   7832
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption scRegla 
         Height          =   375
         Left            =   -70000
         TabIndex        =   143
         Top             =   4680
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Detalle de Tasas de la Regla:"
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
         SubItemCaption  =   -1  'True
         ForeColor       =   4210752
      End
      Begin VB.Label Label7 
         Caption         =   "Indicar los posibles usos o destinos que la persona puede darle a este plan."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2172
         Index           =   3
         Left            =   -69760
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Historial de Tasas Aplicadas para Cálculo de Rendimientos"
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
         Height          =   330
         Index           =   4
         Left            =   -69760
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   8892
      End
      Begin VB.Label Label7 
         Caption         =   "Indicar los puntos de aumento o disminución de Tasas de rendimientos, según el vencimiento del contrato."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Index           =   2
         Left            =   -69400
         TabIndex        =   11
         Top             =   5160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Indicar el rango de días antes del vencimiento original."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1452
         Index           =   1
         Left            =   -69640
         TabIndex        =   9
         Top             =   2880
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Indica los Porcentajes (%) de multas ante retiros anticipados"
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
         Height          =   330
         Index           =   2
         Left            =   -70000
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   9255
      End
      Begin VB.Label Label7 
         Caption         =   "Indicar el % de castigo sobe el principal cuando exista un retiro anticipado, al vencimiento del Plazo Original."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1692
         Index           =   0
         Left            =   -69640
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   1572
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Planes"
                  Text            =   "Listado de Planes"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Retiros"
                  Text            =   "Tabla de Retiros"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8040
      TabIndex        =   3
      Top             =   840
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   312
      Left            =   1680
      TabIndex        =   39
      Top             =   480
      Width           =   6132
      _Version        =   1572864
      _ExtentX        =   10821
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1680
      TabIndex        =   66
      Top             =   840
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   2880
      TabIndex        =   67
      Top             =   840
      Width           =   4932
      _Version        =   1572864
      _ExtentX        =   8700
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
   Begin XtremeSuiteControls.CheckBox chkFiltra 
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   160
      Top             =   1200
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Activos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFiltra 
      Height          =   255
      Index           =   2
      Left            =   6000
      TabIndex        =   162
      Top             =   1200
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Certificados a Plazo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar Planes:"
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
      Height          =   315
      Index           =   4
      Left            =   1440
      TabIndex        =   161
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Image imgCopia 
      Height          =   240
      Left            =   8040
      Picture         =   "frmFNDPlanes.frx":384F
      ToolTipText     =   "Copiar Configuración de un Plan a Otro"
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
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
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
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
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1092
   End
End
Attribute VB_Name = "frmFNDPlanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigoP As String, vTipoBusca As String
Dim vCodigo As Long, vGuardar As Boolean, vScroll As Boolean
Dim vPaso As Boolean, vSearch As Boolean


Private Sub sbLimpiaPantalla()

tcMain.Item(0).Selected = True

tcMain.Item(1).Enabled = False
tcMain.Item(2).Enabled = False
tcMain.Item(3).Enabled = False
tcMain.Item(4).Enabled = False


tcAux.Item(0).Selected = True
tcAux.Item(4).Enabled = False
 
 vTipoBusca = "D"
 vCodigo = 0
 vCodigoP = ""
 vEdita = False
 
 cboPlazo.Text = "Días"
 
 txtUltContrato = ""
 
 txtCodigo = ""
 txtDescripcion = ""
 txtNotas.Text = ""
 txtCodigoDeduc.Text = ""
 
 
 txtAseCod = ""
 txtAseDesc = ""
 chkDeducirPlanilla.Value = vbChecked
 chkGeneraMora.Value = vbChecked
 
 txtPlazo = "1"
 txtPorcentaje.Text = 0
 txtMonto.Text = 0
 txtInversion.Text = 0
  
 'Caracteristicas
 chkCDP.Value = vbUnchecked
 chkControlaSaldo.Value = vbChecked
  
 txtTasaMargenNegociacion.Text = "0.00"
  
 chkLiqSocio.Value = xtpChecked
 chkLiqPlanesAhorros.Value = xtpChecked
  
 chkPagoTercero.Value = xtpUnchecked
 chkVisibleEC.Value = xtpChecked
  
 chkCuentaMaestra.Value = vbUnchecked

  
 'Rendimientos
 chkRendimientos.Value = vbUnchecked
 cboBaseCalculo.Text = 365
 txtTasaBase = 0
 chkTasaFluctuante.Value = vbChecked
 chkCapiltalizaRendimiento.Value = vbChecked
 chkUtilizaTBP.Value = vbUnchecked
 chkRenuevaTasa.Value = vbUnchecked
 txtNuevaTasa.Enabled = False
 txtNuevaTasa.Text = 0
    
 'BackToBack
 chkGarantia.Value = vbUnchecked
 txtBacktoBackDisponible.Text = 0
 txtBacktoBackTasaAdCredito = 0
 chkBacktoBackIntegra.Value = vbUnchecked
 
 'SINPE / Mov. Entre Fondos
 chkSINPE.Value = vbUnchecked
 chkSINPE.Caption = "Cuenta SINPE? Código Interno: xx"
 chkMovEntreFondos.Value = vbUnchecked
 
 chkSINPE_Mov.Value = xtpChecked
 cboSinpeTipos.Text = "Todos"
  
  
 'Otros
 chkMovCajas.Value = vbChecked
 chkCajasRetiros.Value = vbUnchecked
 chkPagoTercero.Value = vbUnchecked
 
 txtContratosPersona.Text = 1
 txtComVentaMonto.Text = 0
 txtComVentaTasa.Text = 0
 txtTasaComAportes.Text = 0
 txtTasaComRend.Text = 0
 txtImpuestosRend.Text = 0
 
 'Web
 chkVenceWeb.Value = vbUnchecked
 chkLiquidaWebSite.Value = vbUnchecked
 chkWebSite.Value = vbUnchecked
 dtpVenceWeb.Value = Format(fxFechaServidor, "dd/mm/yyyy")
 
 chkWebCrea.Value = xtpUnchecked
 chkWebModifica.Value = xtpUnchecked
 
 chkRetiroParcial.Value = xtpChecked
 
 chkPat_Unifica.Value = xtpUnchecked
 
 
 'Plan Contable
 txtCuentaCod = ""
 txtCuentaDesc = ""
 txtCuentaRend = ""
 txtCuentaRendDesc = ""
 txtCuentaGasto = ""
 txtCuentaGasDesc = ""
 txtCuentaGstComision = ""
 txtCuentaGstComisionDesc = ""
 txtCuentaIngComision = ""
 txtCuentaIngComisionDesc = ""
 txtCuentaIngRetiros = ""
 txtCuentaIngRetirosDesc = ""
 txtCuentaImpuestos.Text = ""
 txtCuentaImpuestosDesc.Text = ""
  
  
 'Actualiza Bloqueos
 chkGarantia_Click
 chkRendimientos_Click
 chkUtilizaTBP_Click
 
 'Vencimientos
 cboVence_Accion.Text = "NA"
 
 chkVence_Renueva.Value = xtpUnchecked
 txtVence_Plan.Text = ""
 txtVence_PlanDesc.Text = ""
 
 chkVence_AplTasaCntVencidos.Value = xtpChecked
 chkVence_ActivaControlVencimiento.Value = xtpChecked
 
 
 
 
End Sub




Private Sub btnConfig_Recaudo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "exec spFnd_Plan_Recaudo_Config " & cboOperadora.ItemData(cboOperadora.ListIndex) & ",'" & txtCodigo.Text & "','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

If Not glogon.error Then
    txtAseCod.Text = rs!Codigo & ""
    txtAseDesc.Text = rs!Descripcion & ""
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnFechaCorte_Click()
Dim strSQL As String

On Error GoTo vError


If dtpVenceWeb.Value < fxFechaServidor Or chkCDP.Value = xtpChecked Then
    MsgBox "La Fecha de Corte es menor a la fecha actual o El Plan es un Certificado a Plazo, verifique!", vbExclamation
    Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "exec spFnd_Plan_FechaCorte_Update " & cboOperadora.ItemData(cboOperadora.ListIndex) & ", '" & txtCodigo.Text _
        & "','" & Format(dtpVenceWeb.Value, "yyyy-mm-dd") & "','" & glogon.Usuario & "'"

Call ConectionExecute(strSQL)

strSQL = "Fecha de Corte: " & Format(dtpVenceWeb.Value, "yyyy-mm-dd") & ", para el Plan: " & txtCodigo.Text
Call Bitacora("Aplica", strSQL)

Me.MousePointer = vbDefault

MsgBox "Todos los Contratos se actualizaron con esta fecha de corte: " _
       & Format(dtpVenceWeb.Value, "yyyy-mm-dd"), vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnRegla_Activa_Click(Index As Integer)

If txtR_Id.Text = "0" Or Not IsNumeric(txtR_Id.Text) Then Exit Sub

Call sbReglas_Activa

End Sub

Private Sub btnRegla_Click(Index As Integer)

Select Case Index
    Case 0 'Nueva
        Call sbReglas_New
    
    Case 1 'Guardar
        Call sbReglas_Update
End Select


End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboMoneda.SetFocus
End Sub

Private Sub cboGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub cboMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlazo.SetFocus
End Sub

Private Sub cboOperadora_Click()
'txtCodigo_LostFocus
End Sub

Private Sub cboOperadora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus
End Sub

Private Sub cboPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub




Private Sub chkDeducirPlanilla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkGeneraMora.SetFocus
End Sub


Private Sub chkEnlazaPatrimonio_Click()
If chkEnlazaPatrimonio Then
  cboTipoPatrimonio.Enabled = True
Else
  cboTipoPatrimonio.Enabled = False
End If
End Sub

Private Sub chkGarantia_Click()
If chkGarantia.Value = vbChecked Then
   txtBacktoBackDisponible.Enabled = True
   txtBacktoBackTasaAdCredito.Enabled = True
   chkBacktoBackIntegra.Enabled = True
Else
   txtBacktoBackDisponible.Enabled = False
   txtBacktoBackTasaAdCredito.Enabled = False
   chkBacktoBackIntegra.Enabled = False
End If
End Sub

Private Sub chkGeneraMora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 tcAux.Item(1).Selected = True

 chkCDP.SetFocus
End If
End Sub

Private Sub chkRendimientos_Click()

cboBaseCalculo.Enabled = False
txtTasaBase.Enabled = False
chkTasaFluctuante.Enabled = False
chkCapiltalizaRendimiento.Enabled = False
chkUtilizaTBP.Enabled = False

If chkRendimientos.Value = vbChecked Then
    cboBaseCalculo.Enabled = True
    txtTasaBase.Enabled = True
    chkTasaFluctuante.Enabled = True
    chkCapiltalizaRendimiento.Enabled = True
    chkUtilizaTBP.Enabled = True
End If

chkUtilizaTBP_Click

End Sub


Private Sub chkRenuevaTasa_Click()
If chkRenuevaTasa.Value Then
  txtNuevaTasa.Enabled = True
'  txtNuevaTasa.SetFocus
Else
  txtNuevaTasa.Enabled = False
  txtNuevaTasa.Text = 0
End If
End Sub

Private Sub chkUtilizaTBP_Click()
If chkUtilizaTBP.Value = vbChecked Then
   txtTasaBase.Enabled = False
Else
   txtTasaBase.Enabled = True
End If
End Sub

Private Sub chkVenceWeb_Click()
    If chkVenceWeb.Value = vbChecked Then
        dtpVenceWeb.Enabled = True
    Else
        dtpVenceWeb.Enabled = False
    End If
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_plan from fnd_planes" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
    
    If chkFiltra(0).Value = xtpChecked Then
        strSQL = strSQL & " And (WEB_CREAR = 1 or WEB_LIQUIDA = 1 or WEBSITE = 1)"
    End If
    
    If chkFiltra(1).Value = xtpChecked Then
        strSQL = strSQL & " And ESTADO = 'A'"
    End If
    
    If chkFiltra(2).Value = xtpChecked Then
        strSQL = strSQL & " And TIPO_CDP = 1"
    End If
    
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and cod_plan > '" & txtCodigo & "' order by cod_plan asc"
    Else
       strSQL = strSQL & " and cod_plan < '" & txtCodigo & "' order by cod_plan desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsultaPlan(rs!COD_PLAN)
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
Dim strSQL As String
 
vModulo = 18 'Fondo de Inversion
 
vGrid.AppearanceStyle = fxGridStyle
vhGrid.AppearanceStyle = vGrid.AppearanceStyle
vaGrid.AppearanceStyle = vGrid.AppearanceStyle

vSearch = False
 
vEdita = False
vTipoBusca = "C"
vGuardar = True

vPaso = False
 
cboSinpeTipos.AddItem "Todos"
cboSinpeTipos.ItemData(cboSinpeTipos.ListCount - 1) = "3"
cboSinpeTipos.AddItem "Débitos"
cboSinpeTipos.ItemData(cboSinpeTipos.ListCount - 1) = "1"
cboSinpeTipos.AddItem "Créditos"
cboSinpeTipos.ItemData(cboSinpeTipos.ListCount - 1) = "2"
cboSinpeTipos.AddItem "Ninguno"
cboSinpeTipos.ItemData(cboSinpeTipos.ListCount - 1) = "0"


cboR_Tipo.AddItem "Todos"
cboR_Tipo.AddItem "Apertura"
cboR_Tipo.Text = "Todos"

With lswR.ColumnHeaders
    .Clear
    .Add , , "Regla Id", 1100
    .Add , , "Tipo", 1200, vbCenter
    .Add , , "Vigente", 1200, vbCenter
    .Add , , "Apartir", 1200, vbCenter
    .Add , , "Justificación", 3200
    .Add , , "Reg.Usuario", 1200, vbCenter
    .Add , , "Reg.Fecha", 1200, vbCenter
    .Add , , "Mod.Usuario", 1200, vbCenter
    .Add , , "Mod.Fecha", 1200, vbCenter
    .Add , , "Act.Usuario", 1200, vbCenter
    .Add , , "Act.Fecha", 1200, vbCenter
End With
 
vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

strSQL = "select COD_TIPO_PLAN as 'IdX', DESCRIPCION as 'ItmX' from FND_PLANES_TIPO_PLAN"
Call sbCbo_Llena_New(cboTipoPlan, strSQL, False, True)

With cboVence_Accion
 .Clear
 
 .AddItem "Ninguna"
 .ItemData(.ListCount - 1) = "NA"
 
 .AddItem "Liquidación y Traslado de Fondos"
 .ItemData(.ListCount - 1) = "TF"
 
 .AddItem "Liquidación y Desembolso en Bancos"
 .ItemData(.ListCount - 1) = "DB"
 
 .Text = "Ninguna"
 
End With


With lswEstados.ColumnHeaders
  .Clear
  .Add , , "Estados", lswEstados.Width - 150
End With

With lswVence_Plazos.ColumnHeaders
  .Clear
  .Add , , "Plazo en Meses", lswVence_Plazos.Width - 250
End With



strSQL = "select rtrim(cod_Divisa) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
        & " from CntX_Divisas" _
        & " Where cod_Contabilidad = " & GLOBALES.gEnlace
Call sbCbo_Llena_New(cboMoneda, strSQL, False, True)

strSQL = "select rtrim(descripcion) as 'ItmX',cod_operadora as 'Idx' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)


strSQL = "select rtrim(cod_grupo) as 'IdX', rtrim(descripcion) as ItmX from fnd_grupos"
Call sbCbo_Llena_New(cboGrupo, strSQL, False, True)

cboBaseCalculo.Clear
cboBaseCalculo.AddItem "360"
cboBaseCalculo.AddItem "365"
cboBaseCalculo.Text = "360"

 cboTipoAporte.Clear
 cboTipoAporte.AddItem "Monto"
 cboTipoAporte.AddItem "Porcentaje"
 cboTipoAporte.Text = "Monto"

 cboPlazo.Clear
 cboPlazo.AddItem "Meses"
 cboPlazo.AddItem "Días"
 cboPlazo.Text = "Meses"
 
 cboEstado.Clear
 cboEstado.AddItem "Activo"
 cboEstado.AddItem "Inactivo"
 cboEstado.Text = "Activo"
 
 
 cboTipoPatrimonio.Clear
 cboTipoPatrimonio.AddItem "Obrero"
 cboTipoPatrimonio.AddItem "Patronal"
 cboTipoPatrimonio.AddItem "Capitalizado"
 cboTipoPatrimonio.AddItem "Excedente"
 cboTipoPatrimonio.Text = "Obrero"
 
 cboTipoPatrimonio.Enabled = False
 

Call sbToolBarIconos(tlb)
Call sbToolBar(tlb, "nuevo")

Call sbLimpiaPantalla

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Function fxValida() As Boolean
Dim rsX As New ADODB.Recordset, strSQLx As String
Dim vMensaje As String

vMensaje = ""

On Error GoTo vError

If chkCDP.Value = xtpChecked And CCur(txtInversion.Text) <= 0 Then
  vMensaje = vMensaje & vbCrLf & " - Para un CDP la inversión no puede ser igual o menor a CERO!"
End If

If chkCDP.Value = xtpUnchecked Then
  If Mid(cboTipoAporte.Text, 1, 1) = "M" And CCur(txtMonto.Text) <= 0 Then
      vMensaje = vMensaje & vbCrLf & " - El monto de recaudo no puede ser igual o menor a CERO!"
  End If
End If

If Trim(txtCodigo) = "" Then vMensaje = vMensaje & vbCrLf & " - No se indico el código del Plan"
If Trim(txtDescripcion) = "" Then vMensaje = vMensaje & vbCrLf & " - No se indico la descripción"

If Trim(txtCuentaCod) = "" Or Trim(txtCuentaGasto) = "" Then vMensaje = vMensaje & vbCrLf & " - Las cuentas contables no son válidas"
If Trim(txtCuentaIngComision) = "" Or Trim(txtCuentaIngRetiros) = "" Then vMensaje = vMensaje & vbCrLf & " - Las cuentas contables no son válidas"
If Trim(txtCuentaGstComision) = "" Or Trim(txtCuentaImpuestos) = "" Then vMensaje = vMensaje & vbCrLf & " - Las cuentas contables no son válidas"

If Len(vMensaje) = 0 Then
    If Not fxgCntCuentaValida(txtCuentaCod) Then vMensaje = vMensaje & vbCrLf & " - La cuenta del PLAN no es válida!"
    If Not fxgCntCuentaValida(txtCuentaGasto) Then vMensaje = vMensaje & vbCrLf & " - La cuenta del RENDIMIENTO no es válida!"
    If Not fxgCntCuentaValida(txtCuentaIngComision) Then vMensaje = vMensaje & vbCrLf & " - La cuenta para Ingresos por Comisión de Administración no es válida!"
    If Not fxgCntCuentaValida(txtCuentaIngRetiros) Then vMensaje = vMensaje & vbCrLf & " - La cuenta para Ingresos por Retiros Anticipados no es válida!"
    If Not fxgCntCuentaValida(txtCuentaGstComision) Then vMensaje = vMensaje & vbCrLf & " - La cuenta para Gasto por Comisiones no es válida!"
    If Not fxgCntCuentaValida(txtCuentaImpuestos) Then vMensaje = vMensaje & vbCrLf & " - La cuenta del Impuesto s/Rendimientos no es válida!"""
End If

If Not IsNumeric(txtPlazo) Then vMensaje = vMensaje & vbCrLf & " - El Plazo mínimo no es válido"
If Not IsNumeric(txtMonto) Then vMensaje = vMensaje & vbCrLf & " - La Cuota mínima no es válida"
If Not IsNumeric(txtInversion) Then vMensaje = vMensaje & vbCrLf & " - El Monto de Inversión mínimo no es válido"

If IsNumeric(txtPorcentaje.Text) Then
   If CCur(txtPorcentaje) > 100 Or CCur(txtPorcentaje) < 0 Then
        vMensaje = vMensaje & vbCrLf & " - El % de Referencia no es válido"
   End If
Else
        vMensaje = vMensaje & vbCrLf & " - El % de Referencia no es válido"
End If


If Trim(txtAseCod) = "" Then vMensaje = vMensaje & vbCrLf & " - El código de retención de planillas no es válido"

If Not IsNumeric(txtTasaBase) Then vMensaje = vMensaje & vbCrLf & " - La Tasa base de rendimiento no es válida"
If Not IsNumeric(txtBacktoBackDisponible) Then vMensaje = vMensaje & vbCrLf & " - El Porcentaje de Disponible para Garantía BackToBack no es válido"
If Not IsNumeric(txtBacktoBackTasaAdCredito) Then vMensaje = vMensaje & vbCrLf & " - Los Puntos Adicionales a la Tasa de BackToBack no es válida"

If IsNumeric(txtTasaBase) Then
  If CCur(txtTasaBase) > 100 Or CCur(txtTasaBase) < 0 Then vMensaje = vMensaje & vbCrLf & " - La Tasa base de rendimiento no es válida"
End If
If IsNumeric(txtBacktoBackDisponible) Then
  If CCur(txtBacktoBackDisponible) > 300 Or CCur(txtBacktoBackDisponible) < 0 Then vMensaje = vMensaje & vbCrLf & " - El Porcentaje de Disponible para Garantía BackToBack no es válido"
End If
If IsNumeric(txtBacktoBackTasaAdCredito) Then
  If CCur(txtBacktoBackTasaAdCredito) > 100 Or CCur(txtBacktoBackTasaAdCredito) < 0 Then vMensaje = vMensaje & vbCrLf & " - Los Puntos Adicionales a la Tasa de BackToBack no es válida"
End If


If Not IsNumeric(txtContratosPersona) Then vMensaje = vMensaje & vbCrLf & " - El dato de número de contratos activos x persona no es válida"
If Not IsNumeric(txtComVentaMonto) Then vMensaje = vMensaje & vbCrLf & " - El monto de comisión de venta x contrato vendido no es válida"

If IsNumeric(txtComVentaTasa) Then
  If CCur(txtComVentaTasa) > 100 Or CCur(txtComVentaTasa) < 0 Then vMensaje = vMensaje & vbCrLf & " - La Tasa de Comision de Ventas sobre la inversión no es válida"
End If

If IsNumeric(txtTasaComAportes) Then
  If CCur(txtTasaComAportes) > 100 Or CCur(txtTasaComAportes) < 0 Then vMensaje = vMensaje & vbCrLf & " - La Tasa de Comisión de Administración Sobre Aportes no es válida"
End If
If IsNumeric(txtTasaComRend) Then
  If CCur(txtTasaComRend) > 100 Or CCur(txtTasaComRend) < 0 Then vMensaje = vMensaje & vbCrLf & " - La Tasa de Comisión de Administración Sobre Rendimientos no es válida"
End If
If IsNumeric(txtImpuestosRend) Then
  If CCur(txtImpuestosRend) > 100 Or CCur(txtImpuestosRend) < 0 Then vMensaje = vMensaje & vbCrLf & " - El Porcentaje de Impuesto s/Rendimientos no es válido"
End If

'Revisar aqui cambios de parametros clave, que afecte la base actual
'Sacar Cantidad de Contratos Activos y validar cambio de moneda y base de calculo.
If Len(vMensaje) = 0 Then
    strSQLx = "select P.cod_operadora,P.cod_plan,P.cod_moneda,P.Base_Calculo,count(*) as Contratos" _
            & ",sum(C.Rendimiento) as Rendimiento" _
            & " from fnd_planes P inner join fnd_Contratos C on P.cod_plan = C.cod_plan" _
            & " and P.cod_operadora = C.cod_operadora and C.estado = 'A'" _
            & " where P.cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
            & " and P.cod_plan = '" & txtCodigo & "'" _
            & " group by P.cod_operadora,P.cod_plan,P.cod_moneda,P.Base_Calculo"
    rsX.Open strSQLx, glogon.Conection, adOpenStatic
    If Not rsX.EOF And Not rsX.BOF Then
      
      
    
       If (cboMoneda.ItemData(cboMoneda.ListIndex) <> Trim(rsX!Cod_Moneda)) And rsX!Contratos > 0 Then
            vMensaje = vMensaje & vbCrLf & " - No se puede cambiar la divisa base, ya que existen contratos activos..."
       End If
       
       If (CInt(cboBaseCalculo.Text) <> rsX!Base_Calculo) And rsX!Rendimiento > 0 Then
            vMensaje = vMensaje & vbCrLf & " - No se puede cambiar la base de cálculo de rendimientos, ya que existen contratos a los que se les cálculo con otra base..."
       End If
       
    End If
    rsX.Close
End If

If Len(vMensaje) = 0 Then
   fxValida = True
Else
   fxValida = False
   MsgBox vMensaje, vbExclamation
End If

Exit Function

vError:
   vMensaje = vMensaje & vbCrLf & " Error de Procesamiento!"
   fxValida = False
   MsgBox vMensaje, vbExclamation
 

End Function


Private Sub sbGuardar()
Dim strSQL As String

On Error GoTo vError

If Not fxValida Then Exit Sub



If vEdita Then
  strSQL = "Update FND_Planes Set Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
          & ",Cod_plan='" & UCase(Trim(txtCodigo)) & "',Descripcion='" & UCase(Trim(txtDescripcion)) _
          & "',TIPO_DEDUC = '" & Mid(cboTipoAporte.Text, 1, 1) & "', PORC_DEDUC = " & CCur(txtPorcentaje.Text) & ",plazo_tipo = '" & Mid(cboPlazo.Text, 1, 1) _
          & "',Plazo_Minimo = " & txtPlazo & ",Monto_minimo = " & CCur(txtMonto) & ",Inversion_minimo = " & CCur(txtInversion) _
          & ",Cuenta_Conta='" & fxgCntCuentaFormato(False, txtCuentaCod, 0) & "',Cuenta_Gasto='" & fxgCntCuentaFormato(False, txtCuentaGasto, 0) _
          & "',Codigo_Ase = '" & Trim(txtAseCod) & "',DEDUCIR_PLANILLA = " & chkDeducirPlanilla.Value & ",genera_mora = " & chkGeneraMora.Value _
          & ",cod_grupo = '" & cboGrupo.ItemData(cboGrupo.ListIndex) & "',cod_moneda = '" & cboMoneda.ItemData(cboMoneda.ListIndex) & "',estado = '" & Mid(cboEstado.Text, 1, 1) _
          & "',Sirve_Garantia = " & chkGarantia.Value & ", GARANTIA_PORC_DISP = " & CCur(txtBacktoBackDisponible) _
          & ",GARANTIA_INTEGRADA = " & chkBacktoBackIntegra.Value & ",GARANTIA_TASAAD = " & CCur(txtBacktoBackTasaAdCredito) _
          & ",CAPITALIZA_RENDIMIENTOS = " & chkCapiltalizaRendimiento.Value & ",calcula_rend = " & chkRendimientos.Value _
          & ",BASE_CALCULO = " & cboBaseCalculo.Text & ",TASA_BASE = " & CCur(txtTasaBase) & ",UTILIZA_TBP = " & chkUtilizaTBP.Value & ",WEB_LIQUIDA = " & chkLiquidaWebSite.Value
          
    If chkVenceWeb.Value = vbChecked Then
        strSQL = strSQL & ",WEB_VENCE = '" & Format(dtpVenceWeb.Value, "yyyymmdd") & "'"
    Else
        strSQL = strSQL & ",WEB_VENCE = NULL"
    End If
          
   strSQL = strSQL & ",UTILIZA_TASA_FLUCTUANTE = " & chkTasaFluctuante.Value & ",CONTROLA_SALDO = " & chkControlaSaldo.Value _
          & ",TIPO_CDP = " & chkCDP.Value & ",PERMITE_MOV_CAJAS = " & chkMovCajas.Value _
          & ",apl_liq_socio = " & chkLiqSocio.Value & ",cuenta_maestra = " & chkCuentaMaestra.Value _
          & ",visible_Ec = " & chkVisibleEC.Value & ",webSite = " & chkWebSite.Value & ", notas = '" & Trim(txtNotas.Text) _
          & "',COMISION_VTA_MONTO = " & CCur(txtComVentaMonto) & ",COMISION_VTA_INV = " & CCur(txtComVentaTasa) _
          & ",NUM_CONTRATOS_ACTIVOS = " & txtContratosPersona & ",TASA_COMISION_APORTES = " & CCur(txtTasaComAportes) _
          & ",TASA_COMISION_REND = " & CCur(txtTasaComRend) _
          & ",Cuenta_Ing_Retiros = '" & fxgCntCuentaFormato(False, txtCuentaIngRetiros, 0) & "',Cuenta_Gst_Comision = '" & fxgCntCuentaFormato(False, txtCuentaGstComision, 0) _
          & "',Cuenta_Comision_ADM = '" & fxgCntCuentaFormato(False, txtCuentaIngComision, 0) & "', CUENTA_RENDIMIENTO = '" & fxgCntCuentaFormato(False, txtCuentaRend, 0) _
          & "',patrimonio_enlace = " & chkEnlazaPatrimonio.Value & ", patrimonio_tipo = '" & Mid(cboTipoPatrimonio, 1, 1) _
          & "',requiere_beneficiarios = " & chkRequiereBeneficiario.Value & ",deduce_independiente = " & chkDeducIndependiente.Value & ", codigo_deduc = '" & txtCodigoDeduc _
          & "',TASA_AJUSTE_VENCIMIENTO = " & chkRenuevaTasa.Value & ",TASA_AJUSTE = " & CCur(txtNuevaTasa.Text) _
          & ",PERMITE_RETIROS_CAJAS = " & chkCajasRetiros.Value & ", PERMITE_GIRO_TERCEROS = " & chkPagoTercero.Value _
          & ",CUENTA_IMPUESTOS = '" & fxgCntCuentaFormato(False, txtCuentaImpuestos, 0) & "', IMPUESTO_RENTA = " & CCur(txtImpuestosRend.Text) _
          & ",SINPE_CUENTA = " & chkSINPE.Value & ", MOV_ENTRE_FONDOS = " & chkMovEntreFondos.Value _
          & ",SINPE_PRODUCTO = dbo.fxFndSinpeProducto('" & vCodigoP & "')" _
          & ",FORMA_PAGO_SERVICIOS = " & chkFP_Servicios.Value & ",FORMA_PAGO_POS = " & chkFP_POS.Value _
          & ",MOV_ENTRE_FONDOS_TERCEROS = " & chkMovEntreFondosTerceros.Value & ", RENTA_GLOBAL = " & chkRentaGlobal.Value _
          & ", APL_REND_AUTOMATICO = " & chkRendimientoAuto.Value _
          & ", PERMITE_RET_PARCIAL = " & chkRetiroParcial.Value & ", PATRIMONIO_UNIFICA = " & chkPat_Unifica.Value _
          & ", VENCE_ACCION = '" & cboVence_Accion.ItemData(cboVence_Accion.ListIndex) & "', VENCE_PLAN = '" & txtVence_Plan.Text _
          & "', VENCE_RENUEVA = " & chkVence_Renueva.Value & ", VENCE_PLAZO = " & chkVence_Plazo_Sol.Value _
          & ", COD_TIPO_PLAN = " & cboTipoPlan.ItemData(cboTipoPlan.ListIndex) & ", WEB_CREAR = " & chkWebCrea.Value _
          & ", WEB_MODIFICA_COUTA = " & chkWebModifica.Value & ", WEB_VALIDA_VENCE = 0" _
          & ", MOV_SINPE_TIPOS = " & cboSinpeTipos.ItemData(cboSinpeTipos.ListIndex) & ", MOV_SINPE = " & chkSINPE_Mov.Value
          
   strSQL = strSQL _
          & ", APLICAR_TASA_CONT_VENCIDOS = " & chkVence_AplTasaCntVencidos.Value & ", APLICAR_EN_PROCS_CONTRS_VENCIDOS = " & chkVence_ActivaControlVencimiento.Value _
          & ", SIF_LIQUIDA = " & chkLiqPlanesAhorros.Value & ", PAGO_CUPONES = " & chkCDP_PagaCupones.Value _
          & ", IndAplicarAMora = isnull(IndAplicarAMora, 0)" _
          & ", TASA_MARGEN_NEGOCIACION = " & CCur(txtTasaMargenNegociacion.Text) _
          & " where cod_operadora = " & vCodigo & " and Cod_plan='" & vCodigoP & "'"
          
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Plan:" & Trim(txtCodigo) & "-" & Trim(txtDescripcion) & " Operadora:" & Trim(cboOperadora))
    
Else

   vCodigo = cboOperadora.ItemData(cboOperadora.ListIndex)
   vCodigoP = Trim(txtCodigo)

   strSQL = "insert FND_Planes(Cod_operadora,Cod_plan,Descripcion,TIPO_DEDUC,PORC_DEDUC,Plazo_Tipo,Plazo_Minimo,Monto_Minimo,Inversion_minimo" _
          & ",Cuenta_Conta,Codigo_Ase,Estado,cod_moneda,Sirve_Garantia,Garantia_Porc_Disp,Garantia_Integrada,Garantia_TasaAD" _
          & ",Consecutivo,Cuenta_gasto,Rend_corte,Calcula_rend,Base_calculo,Tasa_base,Capitaliza_rendimientos,utiliza_TBP" _
          & ",Utiliza_Tasa_Fluctuante,tipo_CDP, Visible_ec,cuenta_maestra,apl_liq_socio,cod_grupo,webSite,Notas" _
          & ",Deducir_Planilla,Genera_Mora,Controla_Saldo,Permite_Mov_Cajas,Num_Contratos_Activos,Comision_Vta_Monto,Comision_Vta_Inv" _
          & ",Tasa_Comision_Aportes,Tasa_Comision_Rend,Cuenta_Ing_Retiros,Cuenta_Gst_Comision,Cuenta_Comision_ADM,CUENTA_RENDIMIENTO" _
          & " ,requiere_beneficiarios,deduce_independiente,codigo_deduc,patrimonio_enlace,patrimonio_tipo" _
          & ",TASA_AJUSTE_VENCIMIENTO,TASA_AJUSTE,WEB_LIQUIDA,WEB_VENCE, PERMITE_RETIROS_CAJAS,PERMITE_GIRO_TERCEROS" _
          & ", CUENTA_IMPUESTOS,IMPUESTO_RENTA,SINPE_CUENTA,SINPE_PRODUCTO,MOV_ENTRE_FONDOS,MOV_ENTRE_FONDOS_TERCEROS" _
          & ", FORMA_PAGO_SERVICIOS, FORMA_PAGO_POS, RENTA_GLOBAL, APL_REND_AUTOMATICO, PERMITE_RET_PARCIAL, PATRIMONIO_UNIFICA" _
          & ", VENCE_ACCION, VENCE_PLAN, VENCE_RENUEVA,VENCE_PLAZO, COD_TIPO_PLAN, WEB_CREAR , WEB_MODIFICA_COUTA, WEB_VALIDA_VENCE" _
          & ", MOV_SINPE_TIPOS, MOV_SINPE, APLICAR_TASA_CONT_VENCIDOS, APLICAR_EN_PROCS_CONTRS_VENCIDOS" _
          & ", SIF_LIQUIDA, PAGO_CUPONES, IndAplicarAMora,TASA_MARGEN_NEGOCIACION)"
   
   strSQL = strSQL & " values(" & cboOperadora.ItemData(cboOperadora.ListIndex) & ",'" & UCase(Trim(txtCodigo)) & "','" & UCase(Trim(txtDescripcion)) _
          & "','" & Mid(cboTipoAporte.Text, 1, 1) & "'," & CCur(txtPorcentaje.Text) & ",'" & Mid(cboPlazo.Text, 1, 1) & "'," & txtPlazo & "," & CCur(txtMonto) & "," & CCur(txtInversion) & ",'" & fxgCntCuentaFormato(False, txtCuentaCod, 0) & "','" _
          & Trim(txtAseCod) & "','" & Mid(cboEstado.Text, 1, 1) & "','" & cboMoneda.ItemData(cboMoneda.ListIndex) & "'," & chkGarantia.Value _
          & "," & CCur(txtBacktoBackDisponible) & "," & chkBacktoBackIntegra.Value & "," & CCur(txtBacktoBackTasaAdCredito) _
          & ",0,'" & fxgCntCuentaFormato(False, txtCuentaGasto, 0) & "','" & Format(fxFechaServidor, "yyyy/mm/dd") _
          & "'," & chkRendimientos.Value & "," & cboBaseCalculo.Text & "," & CCur(txtTasaBase) & "," & chkCapiltalizaRendimiento.Value & "," & chkUtilizaTBP.Value _
          & "," & chkTasaFluctuante.Value & "," & chkCDP.Value & "," & chkVisibleEC.Value & "," & chkCuentaMaestra.Value _
          & "," & chkLiqSocio.Value & ",'" & cboGrupo.ItemData(cboGrupo.ListIndex) & "'," & chkWebSite.Value & ",'" & Trim(txtNotas.Text) _
          & "'," & chkDeducirPlanilla.Value & "," & chkGeneraMora.Value & "," & chkControlaSaldo.Value & "," & chkMovCajas.Value _
          & "," & txtContratosPersona & "," & CCur(txtComVentaMonto) & "," & CCur(txtComVentaTasa) _
          & "," & CCur(txtTasaComAportes) & "," & CCur(txtTasaComRend) & ",'" & fxgCntCuentaFormato(False, txtCuentaIngRetiros, 0) _
          & "','" & fxgCntCuentaFormato(False, txtCuentaGstComision, 0) & "','" & fxgCntCuentaFormato(False, txtCuentaIngComision, 0) _
          & "','" & fxgCntCuentaFormato(False, txtCuentaRend, 0) & "'," & chkRequiereBeneficiario & "," & chkDeducIndependiente.Value & ",'" & Trim(txtCodigoDeduc) _
          & "'," & chkEnlazaPatrimonio.Value & ",'" & Mid(cboTipoPatrimonio, 1, 1) & "'," & chkRenuevaTasa.Value & " ," & CCur(txtNuevaTasa.Text) & "," & chkLiquidaWebSite.Value
          
    If chkVenceWeb.Value = vbChecked Then
        strSQL = strSQL & ",'" & Format(dtpVenceWeb.Value, "yyyymmdd") & "'"
    Else
        strSQL = strSQL & ",NULL"
    End If
    
    
        strSQL = strSQL & "," & chkCajasRetiros.Value & "," & chkPagoTercero.Value _
               & ",'" & fxgCntCuentaFormato(False, txtCuentaImpuestos.Text, 0) & "'," & CCur(txtImpuestosRend.Text) _
               & ", " & chkSINPE.Value & ",dbo.fxFndSinpeProducto('" & vCodigoP & "')," & chkMovEntreFondos.Value _
               & ", " & chkMovEntreFondosTerceros.Value & ", " & chkFP_Servicios.Value & ", " & chkFP_POS.Value _
               & ", " & chkRentaGlobal.Value & ", " & chkRendimientoAuto.Value _
               & ", " & chkRetiroParcial.Value & ", " & chkPat_Unifica.Value _
               & ",'" & cboVence_Accion.ItemData(cboVence_Accion.ListIndex) _
               & "', '" & txtVence_Plan.Text & "', " & chkVence_Renueva.Value _
               & ", " & chkVence_Plazo_Sol.Value & ", " & cboTipoPlan.ItemData(cboTipoPlan.ListIndex) _
               & ", " & chkWebCrea.Value & ", " & chkWebModifica.Value & ", 0" _
               & ", " & cboSinpeTipos.ItemData(cboSinpeTipos.ListIndex) & ", " & chkSINPE_Mov.Value _
               & ", " & chkVence_AplTasaCntVencidos.Value & ", " & chkVence_ActivaControlVencimiento.Value _
               & ", " & chkLiqPlanesAhorros.Value & ", " & chkCDP_PagaCupones.Value & ", 0" _
               & ", " & CCur(txtTasaMargenNegociacion.Text) & ")"
               
   
   Call ConectionExecute(strSQL)
   
   Call Bitacora("Registra", "Plan:" & Trim(txtCodigo) & "-" & Trim(txtDescripcion) & " Operadora:" & Trim(cboOperadora))
   
End If


vCodigo = cboOperadora.ItemData(cboOperadora.ListIndex)
vCodigoP = Trim(txtCodigo)

Call sbConsultaPlan(txtCodigo)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim strSQL As String, i As Integer

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este plan", vbYesNo)
If i = vbYes Then
  strSQL = "delete FND_Planes where cod_Operadora = " & vCodigo
  strSQL = strSQL & " And cod_plan='" & Trim(vCodigoP) & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Borra", "Plan:" & Trim(txtCodigo) & "-" & Trim(txtDescripcion) & " Operadora:" & Trim(cboOperadora))
  
  Call sbLimpiaPantalla
  Call sbToolBar(Me.tlb, "nuevo")
  vEdita = False
End If

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description)
    
  
End Sub

Public Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String, Optional vBorra As Boolean = True)
Dim rs As New ADODB.Recordset, i As Integer
Dim vNota As String


If vBorra Then
    vGrid.MaxCols = vGridMaxCol
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.Col = i
     vGrid.Text = ""
    Next i
End If

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    vGrid.Text = CStr(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1)))
       
     
    If vGrid.Col = 1 Then
          vGrid.TextTip = TextTipFixed
          vGrid.TextTipDelay = 1000
        
          vGrid.CellNote = "Registro: " & vbCrLf & "---> Fecha : " & rs!Registro_Fecha & vbCrLf & "---> Usuario : " & rs!Registro_Usuario _
                & vbCrLf & vbCrLf & "Actualiza: " & vbCrLf & "---> Fecha : " & rs!Actualiza_fecha & vbCrLf & "---> Usuario : " & rs!Actualiza_usuario
    End If
   
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

End Sub





Private Sub imgCopia_Click()

GLOBALES.gTag = cboOperadora.ItemData(cboOperadora.ListIndex)
GLOBALES.gTag2 = Trim(txtCodigo.Text)
GLOBALES.gTag3 = Trim(txtDescripcion.Text)

Call sbFormsCall("frmFNDPlanesCopia", 1, , , False, Me)

End Sub


Private Sub lswDestinos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert fnd_planes_destinos(cod_plan,cod_operadora,cod_destino,registro_usuario,registro_fecha)" _
          & " values('" & txtCodigo.Text & "'," & cboOperadora.ItemData(cboOperadora.ListIndex) & ",'" & Item.Tag _
          & "','" & glogon.Usuario & "',dbo.MyGetdate())"
   Call ConectionExecute(strSQL)
   Call Bitacora("Aplica", "Asignación Plan " & txtCodigo.Text & " -> Destino : " & Item.Tag)

Else
   strSQL = "delete fnd_planes_destinos where cod_destino = '" & Item.Tag _
          & "' and cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) & " and cod_plan = '" & txtCodigo.Text & "'"
   Call ConectionExecute(strSQL)
   Call Bitacora("Elimina", "Asignación Plan " & txtCodigo.Text & " -> Destino : " & Item.Tag)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswEstados_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, vMovimiento As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
  vMovimiento = "Registra"
  strSQL = "insert FND_PLANES_ESTADOS(cod_plan,cod_operadora,cod_estado,registro_usuario,registro_fecha) values('" & txtCodigo.Text _
         & "'," & cboOperadora.ItemData(cboOperadora.ListIndex) & " ,'" & Item.Tag & "','" & glogon.Usuario & "',dbo.MyGetdate())"
Else
  vMovimiento = "Borrar"
  strSQL = "delete FND_PLANES_ESTADOS where cod_plan = '" _
         & txtCodigo & "' and cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) & "" _
         & " and cod_estado = '" & Item.Tag & "' "

End If
Call ConectionExecute(strSQL)

Call Bitacora(vMovimiento, "Estado Persona : " & Item.Text & " del Plan:" & txtCodigo.Text)

Exit Sub


vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswR_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

Call sbRegla_Load(Item.Text)

End Sub

Private Sub lswVence_Plazos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, vMovimiento As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
  vMovimiento = "Registra"
  strSQL = "exec  spFnd_Planes_Plazos_Registro " & cboOperadora.ItemData(cboOperadora.ListIndex) & ", '" & txtCodigo.Text _
         & "', " & Item.Tag & ",'" & glogon.Usuario & "', 'A'"
Else
  vMovimiento = "Borrar"
  strSQL = "exec  spFnd_Planes_Plazos_Registro " & cboOperadora.ItemData(cboOperadora.ListIndex) & ", '" & txtCodigo.Text _
         & "', " & Item.Tag & ",'" & glogon.Usuario & "', 'B'"

End If
Call ConectionExecute(strSQL)

Call Bitacora(vMovimiento, "Plazos de Vencimiento: " & Item.Text & " del Plan:" & txtCodigo.Text)

Exit Sub


vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tcAux_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo vError

Select Case Item.Index
   Case 4
    Call sbCargaEstados
End Select

txtDescripcion.SetFocus
Exit Sub

vError:
End Sub

Private Sub sbReglas_New()

    txtR_Id.Text = "0"
    txtR_Vigente.Text = "No"
    txtR_Justifica.Text = ""
    cboR_Tipo.Text = "Todos"
    dtpR_Fecha.Value = Now
    
    vaGrid.MaxRows = 0

End Sub


Private Sub sbReglas_List()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lswR.ListItems.Clear

strSQL = "exec spFnd_Reglas_Tasas_List " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & ", '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswR.ListItems.Add(, , rs!ID_PER_TASA)
      itmX.SubItems(1) = rs!Tipo_Desc
      itmX.SubItems(2) = rs!Vigente_Desc
      itmX.SubItems(3) = Format(rs!Fecha_Inicio, "yyyy-mm-dd")
      itmX.SubItems(4) = rs!OBS_USUARIO
      itmX.SubItems(5) = rs!USR_REGISTRA & ""
      itmX.SubItems(6) = rs!FEC_REGISTRA & ""
      itmX.SubItems(7) = rs!MODIFICA_USUARIO & ""
      itmX.SubItems(8) = rs!MODIFICA_FECHA & ""
      itmX.SubItems(9) = rs!ACTIVA_USUARIO & ""
      itmX.SubItems(10) = rs!Activa_Fecha & ""
  rs.MoveNext
Loop
rs.Close

If lswR.ListItems.Count = 0 Then
    Call sbReglas_New
End If


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbRegla_Load(pRegla As Long)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass



strSQL = "exec spFnd_Reglas_Tasas_Load " & pRegla & ", " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & ", '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtR_Id.Text = rs!ID_PER_TASA
  txtR_Vigente = rs!Vigente_Desc
  txtR_Justifica.Text = rs!OBS_USUARIO
  cboR_Tipo.Text = rs!Tipo_Desc
  dtpR_Fecha.Value = rs!Fecha_Inicio
End If
rs.Close

strSQL = "select COD_TABLA_AUM,Tipo_Tasa,desde,hasta,plus,registro_usuario,registro_fecha,actualiza_usuario,actualiza_fecha" _
       & " from FND_TABLA_AUMENTOS" _
       & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & " And Cod_Plan = '" & Trim(txtCodigo) & "' and ID_PER_TASA = " & pRegla _
       & " order by desde"
Call sbCargaGridLocal(vaGrid, 5, strSQL)


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbReglas_Update()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spFnd_Reglas_Tasas " & txtR_Id.Text & ", " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & ", '" & txtCodigo.Text & "', '" & Format(dtpR_Fecha.Value, "yyyy-mm-dd") & "', '" & Mid(cboR_Tipo.Text, 1, 1) _
       & "', '" & txtR_Justifica.Text & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

Call Bitacora("Registra", "Regla de Tasas Id: " & rs!IdRegla)

Call sbRegla_Load(rs!IdRegla)

Call sbReglas_List

Me.MousePointer = vbDefault


MsgBox "Regla Registrada satisfactoriamente!", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbReglas_Activa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtR_Id.Text = "0" Or Not IsNumeric(txtR_Id.Text) Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spFnd_Reglas_Tasas_Activa " & txtR_Id.Text & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Pass = 1 Then
    Call Bitacora("Activa", "Regla de Tasas Id: " & rs!IdRegla)
    
    Call sbRegla_Load(txtR_Id.Text)
    
    Call sbReglas_List
    Me.MousePointer = vbDefault
    MsgBox "Regla Activada satisfactoriamente!", vbInformation
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Select Case Item.Index
  Case 1 'Tabla Retiros
    strSQL = "select cod_fnd_tabla_ret,desde,hasta,porcentaje" _
           & " , case when aplicar_a = 'R' then 'Rendimientos' when aplicar_a = 'A' then 'Aportes'  when aplicar_a = 'T' then 'Total Retiro' end as 'Aplicar_A' " _
           & " , registro_usuario,registro_fecha,actualiza_usuario,actualiza_fecha" _
           & " from fnd_tabla_retiros" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " And Cod_Plan = '" & Trim(txtCodigo) & "' order by desde"
    Call sbCargaGridLocal(vGrid, 5, strSQL)
  
  Case 2 'Tabla de Plus
    
    Call sbReglas_List
    
    If lswR.ListItems.Count > 0 Then
        Call sbRegla_Load(lswR.ListItems(1).Text)
    Else
        Call sbRegla_Load(0)
    End If
  
  Case 3 'Historial de Rendimientos
    strSQL = "select corte,tasa,TCP,usuario,fecha_sys from FND_HISTORIAL_REND" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " And Cod_Plan = '" & Trim(txtCodigo.Text) & "' order by IDx desc"
    Call sbCargaGrid(vhGrid, 5, strSQL)
    vhGrid.MaxRows = vhGrid.MaxRows - 1
    
  Case 4 'Destinos: Asignación
     vPaso = True
     strSQL = "select D.cod_destino,D.descripcion,A.cod_plan " _
            & " from fnd_destinos D left join fnd_planes_destinos A on D.cod_destino = A.cod_destino" _
            & " and A.cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
            & " and A.cod_plan = '" & Trim(txtCodigo.Text) & "' Where D.activo = 1"
     lswDestinos.ListItems.Clear
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
      Set itmX = lswDestinos.ListItems.Add(, , rs!Descripcion)
          itmX.Tag = rs!cod_destino
          
      If IsNull(rs!COD_PLAN) Then
         itmX.Checked = False
      Else
         itmX.ForeColor = vbBlue
         itmX.Checked = True
      End If
      rs.MoveNext
     Loop
     rs.Close
  
     vPaso = False
  
End Select


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      Call sbToolBar(Me.tlb, "edicion")
      txtCodigo.SetFocus

    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      vCodigo = cboOperadora.ItemData(cboOperadora.ListIndex)
      vCodigoP = Trim(txtCodigo)
      Call sbToolBar(Me.tlb, "edicion")
      txtDescripcion.SetFocus
      
    Case "BORRAR"
      Call sbBorrar
      
    Case "GUARDAR", "SALVAR"
      Call sbGuardar
          
    Case "DESHACER"
      vEdita = False
      Call sbToolBar(Me.tlb, "nuevo")
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      
    Case "CONSULTAR"
      Select Case vTipoBusca
        Case "1"
          Call sbgFNDBuscaConfiguracion(txtCuentaCod, "C")
        Case "3"
          Call sbgFNDBuscaConfiguracion(txtAseCod, "A")
        Case "5"
          Call sbgFNDBuscaConfiguracion(txtCuentaGasto, "C")
        Case Else
         Call sbConsulta
      End Select
      
    Case "REPORTES"
       
End Select

Call RefrescaTags(Me)

End Sub

Private Sub sbConsulta()

vEdita = True

gBusquedas.Convertir = "N"

If vTipoBusca = "D" Then
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
Else
  gBusquedas.Columna = "cod_plan"
  gBusquedas.Orden = "cod_plan"
End If
gBusquedas.Filtro = " And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
gBusquedas.Consulta = "select cod_plan,descripcion from fnd_Planes"
frmBusquedas.Show vbModal

txtCodigo = gBusquedas.Resultado
txtDescripcion.SetFocus
gBusquedas.Resultado = ""

txtCodigo_LostFocus

End Sub


Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
  Case "Planes"
   frmContenedor.Crt.ReportFileName = SIFGlobal.fxPathReportes("Fondos_Planes.rpt")
  
  Case "Retiros"
   frmContenedor.Crt.ReportFileName = SIFGlobal.fxPathReportes("Fondos_PlanesRetiros.rpt")
End Select

frmContenedor.Crt.Connect = glogon.ConectRPT

frmContenedor.Crt.Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
frmContenedor.Crt.Formulas(1) = "Usuario='" & Trim(glogon.Usuario) & "'"
frmContenedor.Crt.Formulas(2) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
frmContenedor.Crt.PrintReport

End Sub


Private Sub txtAseCod_GotFocus()
vTipoBusca = "3"
End Sub

Private Sub txtAseCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then tlb_ButtonClick tlb.Buttons(7)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAseDesc.SetFocus
End Sub


Private Sub txtAseCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

If Trim(txtAseCod) = "" Then
   txtAseDesc = ""
   Exit Sub
End If

strSQL = "Select Codigo,Descripcion from Catalogo where Codigo='" & Trim(txtAseCod) & "' And Retencion = 'S' "
With rs
 .Open strSQL, glogon.Conection, adOpenStatic
    If .EOF = True Then
       vGuardar = False
       MsgBox "Codigo Incorrecto", vbExclamation
       txtAseCod = ""
       txtAseDesc = ""
    Else
       vGuardar = True
       txtAseDesc = Trim(!Descripcion)
    End If
 .Close
End With
End Sub


Private Sub txtAseDesc_GotFocus()
vTipoBusca = "4"
End Sub

Private Sub txtAseDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkDeducirPlanilla.SetFocus
End Sub


Private Sub txtCodigo_GotFocus()
vTipoBusca = "C"
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Código"
    gBusquedas.Col2Name = "Descripción"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_plan, Descripcion from fnd_Planes"
    gBusquedas.Filtro = " And Cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
    
    If chkFiltra(0).Value = xtpChecked Then
        gBusquedas.Filtro = gBusquedas.Filtro & " And (WEB_CREAR = 1 or WEB_LIQUIDA = 1 or WEBSITE = 1)"
    End If
    
    If chkFiltra(1).Value = xtpChecked Then
        gBusquedas.Filtro = gBusquedas.Filtro & " And ESTADO = 'A'"
    End If
    
    If chkFiltra(2).Value = xtpChecked Then
        gBusquedas.Filtro = gBusquedas.Filtro & " And TIPO_CDP = 1"
    End If
    
    vSearch = True
    frmBusquedas.Show vbModal

    If gBusquedas.Resultado <> "" Then
        txtCodigo.Text = gBusquedas.Resultado
        txtDescripcion.Text = gBusquedas.Resultado2
    End If
    
    vSearch = False
    
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

End Sub

Private Sub txtCodigo_LostFocus()
0
If Trim(txtCodigo) <> "" And Not vSearch Then
   Call sbConsultaPlan(Trim(txtCodigo))
End If

End Sub


Private Sub sbConsultaPlan(strCodigo As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

   strSQL = "select * from vFnd_Planes" _
          & " where cod_Operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
          & " And Cod_plan = '" & strCodigo & "'"
          
   Call OpenRecordSet(rs, strSQL)
     If Not rs.EOF And Not rs.BOF Then
           Call sbToolBar(Me.tlb, "activo")
           vEdita = True
           vCodigo = rs!COD_OPERADORA
           vCodigoP = rs!COD_PLAN
           
           txtCodigo.Text = rs!COD_PLAN
           
           txtDescripcion.Text = Trim(rs!Descripcion)
           txtDescripcion.SetFocus
           
           txtUltContrato.Text = "Cont: " & rs!Consecutivo & " Ult.Tasa: " & rs!UltTasa & "% " & Format(rs!Rend_Corte, "dd/mm/yyyy")
           
           'Plan Contable
           txtCuentaCod.Text = Trim(rs!CtaPlan)
           txtCuentaDesc.Text = rs!CtaPlanDesc
           
           txtCuentaRend.Text = Trim(rs!CtaRnd)
           txtCuentaRendDesc.Text = rs!CtaRndDesc
           
           txtCuentaGasto = Trim(rs!CtaGasto)
           txtCuentaGasDesc.Text = rs!CtaGastoDesc
           
           txtCuentaGstComision = Trim(rs!CtaGstComision)
           txtCuentaGstComisionDesc.Text = rs!CtaGstComisionDesc
           
           txtCuentaIngComision = Trim(rs!CtaComisionAdm)
           txtCuentaIngComisionDesc.Text = rs!CtaComisionAdmDesc
           
           txtCuentaIngRetiros = Trim(rs!CtaIngRetiros)
           txtCuentaIngRetirosDesc.Text = rs!CtaIngRetirosDesc
           
           txtCuentaImpuestos = Trim(rs!CtaImpuesto)
           txtCuentaImpuestosDesc.Text = rs!CtaImpuestoDesc
           
           txtNotas.Text = Trim(rs!Notas)
           cboEstado.Text = IIf(rs!Estado = "A", "Activo", "InActivo")
           
           Call sbCboAsignaDato(cboMoneda, rs!DivisaDesc, True, rs!Cod_Moneda)
           Call sbCboAsignaDato(cboGrupo, rs!GrupoDesc, True, rs!Cod_Grupo)
           Call sbCboAsignaDato(cboTipoPlan, rs!TIPO_PLAN_DESC, True, rs!COD_TIPO_PLAN)
                      
                      
                      
           If rs!Tipo_Deduc = "M" Then
              cboTipoAporte.Text = "Monto"
           Else
              cboTipoAporte.Text = "Porcentaje"
           End If
           
           txtPorcentaje.Text = Format(rs!Porc_Deduc, "Standard")
           
           If Not IsNull(rs!Patrimonio_Tipo) Then
                cboTipoPatrimonio.Enabled = True
                Select Case rs!Patrimonio_Tipo
                  Case "O"
                    cboTipoPatrimonio.Text = "Obrero"
                  Case "P"
                    cboTipoPatrimonio.Text = "Patronal"
                  Case "C"
                    cboTipoPatrimonio.Text = "Capitalizado"
                  Case "E"
                    cboTipoPatrimonio.Text = "Excedente"
                End Select
           Else
                cboTipoPatrimonio.Enabled = False
           End If
           
           
           
           txtAseCod.Text = Trim(rs!codigo_ase)
           txtAseDesc.Text = rs!LineaDesc
           
           chkDeducirPlanilla.Value = rs!DEDUCIR_PLANILLA
           chkGeneraMora.Value = rs!GENERA_MORA
           
           
           txtPlazo.Text = CStr(rs!Plazo_Minimo)
           cboPlazo.Text = IIf(rs!plazo_Tipo = "D", "Días", "Meses")
           txtMonto.Text = Format(rs!MONTO_MINIMO, "Standard")
           txtInversion.Text = Format(rs!INVERSION_MINIMO, "Standard")
           
  
            'Caracteristicas
            chkCDP.Value = rs!tipo_cdp
            chkCDP_PagaCupones.Value = rs!PAGO_CUPONES
            chkControlaSaldo.Value = rs!CONTROLA_SALDO
            txtTasaMargenNegociacion.Text = Format(rs!TASA_MARGEN_NEGOCIACION, "Standard")
            
            chkCuentaMaestra.Value = rs!cuenta_maestra
            chkRequiereBeneficiario = IIf(IsNull(rs!REQUIERE_BENEFICIARIOS), vbUnchecked, rs!REQUIERE_BENEFICIARIOS)
            
            ' SINPE / Mov. Entre Fondos
            chkSINPE.Value = rs!SINPE_Cuenta
            chkSINPE.Caption = "Cuenta SINPE? Código Interno:" & rs!SINPE_PRODUCTO & ""
             
            'Rendimientos
            chkRendimientos.Value = rs!calcula_rend
            cboBaseCalculo.Text = rs!Base_Calculo
            txtTasaBase.Text = Format(rs!tasa_base, "Standard")
            chkTasaFluctuante.Value = rs!UTILIZA_TASA_FLUCTUANTE
            chkCapiltalizaRendimiento.Value = rs!Capitaliza_rendimientos
            chkUtilizaTBP.Value = rs!utiliza_tbp
            chkRendimientoAuto.Value = rs!APL_REND_AUTOMATICO
            
               
            'BackToBack
            chkGarantia.Value = rs!Sirve_Garantia
            txtBacktoBackDisponible.Text = Format(rs!GARANTIA_PORC_DISP, "Standard")
            txtBacktoBackTasaAdCredito = Format(rs!GARANTIA_TASAAD, "Standard")
            chkBacktoBackIntegra.Value = rs!GARANTIA_INTEGRADA
            
            
            'Otros
            chkMovCajas.Value = rs!PERMITE_MOV_CAJAS
            chkCajasRetiros.Value = rs!PERMITE_RETIROS_CAJAS
            chkPagoTercero.Value = rs!PERMITE_GIRO_TERCEROS
            
            chkFP_POS.Value = rs!FORMA_PAGO_POS
            chkFP_Servicios.Value = rs!FORMA_PAGO_SERVICIOS
            chkMovEntreFondos.Value = rs!MOV_ENTRE_FONDOS
            chkMovEntreFondosTerceros.Value = rs!MOV_ENTRE_FONDOS_TERCEROS
            

            chkVisibleEC.Value = rs!visible_EC
            chkLiqSocio.Value = rs!apl_liq_socio
            chkLiqPlanesAhorros.Value = rs!SIF_LIQUIDA
            
            chkDeducIndependiente = IIf(IsNull(rs!DEDUCE_INDEPENDIENTE), vbUnchecked, rs!DEDUCE_INDEPENDIENTE)
            chkEnlazaPatrimonio.Value = rs!patrimonio_enlace
            chkPat_Unifica.Value = rs!Patrimonio_Unifica
            
            chkRetiroParcial.Value = rs!PERMITE_RET_PARCIAL
            
            chkRenuevaTasa.Value = IIf(IsNull(rs!TASA_AJUSTE_VENCIMIENTO), vbUnchecked, rs!TASA_AJUSTE_VENCIMIENTO)
            
            txtNuevaTasa.Text = IIf(IsNull(rs!TASA_AJUSTE), "0", Format(rs!TASA_AJUSTE, "Standard"))
            txtCodigoDeduc.Text = IIf(IsNull(rs!CODIGO_DEDUC), "", rs!CODIGO_DEDUC)
            txtComVentaMonto.Text = Format(rs!COMISION_VTA_MONTO, "Standard")
            txtComVentaTasa.Text = Format(rs!COMISION_VTA_INV, "Standard")
            
            txtImpuestosRend.Text = Format(rs!impuesto_renta, "Standard")
            chkRentaGlobal.Value = rs!RENTA_GLOBAL
            
            txtContratosPersona.Text = CStr(rs!NUM_CONTRATOS_ACTIVOS)
            txtTasaComAportes.Text = Format(rs!TASA_COMISION_APORTES, "Standard")
            txtTasaComRend.Text = Format(rs!TASA_COMISION_REND, "Standard")
            
            'Auto Gestion
            chkWebSite.Value = rs!WebSite
            chkLiquidaWebSite.Value = IIf(IsNull(rs!WEB_LIQUIDA), vbUnchecked, rs!WEB_LIQUIDA)
            ''Controla la fecha de vencimiento en la web
            chkVenceWeb.Value = IIf(IsNull(rs!web_vence), vbUnchecked, vbChecked)
                If chkVenceWeb.Value = vbUnchecked Then
                    dtpVenceWeb.Enabled = False
                    dtpVenceWeb.Value = Format(rs!FechaServer, "dd/mm/yyyy")
                Else
                    chkVenceWeb.Enabled = True
                    dtpVenceWeb.Value = rs!web_vence
                End If
                        
            chkWebCrea.Value = rs!WEB_CREAR
            chkWebModifica.Value = rs!WEB_MODIFICA_COUTA
            
            
            'Vencimientos
 
            Call sbCboAsignaDato(cboVence_Accion, rs!VENCE_ACCION_DESC, True, rs!VENCE_ACCION)
            
            txtVence_Plan.Text = rs!VENCE_PLAN
            txtVence_PlanDesc.Text = rs!Vence_Plan_Desc
            
            chkVence_Plazo_Sol.Value = rs!VENCE_PLAZO
            chkVence_Renueva.Value = rs!VENCE_RENUEVA
            
            chkVence_AplTasaCntVencidos.Value = rs!APLICAR_TASA_CONT_VENCIDOS
            chkVence_ActivaControlVencimiento.Value = rs!APLICAR_EN_PROCS_CONTRS_VENCIDOS
            
            'Sinpe
            chkSINPE_Mov.Value = rs!MOV_SINPE
            Call sbCboAsignaDato(cboSinpeTipos, rs!Sinpe_Tipos_Desc, True, rs!MOV_SINPE_TIPOS)
            
            'Actualiza Bloqueos
            Call chkGarantia_Click
            Call chkRendimientos_Click
            Call chkUtilizaTBP_Click
            
            
            tcMain.Item(1).Enabled = True
            tcMain.Item(2).Enabled = True
            tcMain.Item(3).Enabled = True
            tcMain.Item(4).Enabled = True
            
            tcMain.Item(0).Selected = True
            
            tcAux.Item(4).Enabled = True
            tcAux.Item(0).Selected = True
            
            
            
    End If
   rs.Close

Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCuentaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaCod.Text = gCuenta
   txtCuentaDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaCod.Text = fxgCntCuentaFormato(True, txtCuentaCod, 0)
End If

End Sub

Private Sub txtCuentaCod_LostFocus()
   txtCuentaDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaCod, 0))
   txtCuentaCod.Text = fxgCntCuentaFormato(True, txtCuentaCod, 0)
End Sub

Private Sub txtCuentaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaRend.SetFocus
End Sub


Private Sub txtCuentaGasDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaIngComision.SetFocus
End Sub


Private Sub txtCuentaGasto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaGasDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaGasto.Text = gCuenta
   txtCuentaGasDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaGasto.Text = fxgCntCuentaFormato(True, txtCuentaGasto, 0)
End If
End Sub

Private Sub txtCuentaGasto_LostFocus()
   txtCuentaGasDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaGasto, 0))
   txtCuentaGasto.Text = fxgCntCuentaFormato(True, txtCuentaGasto, 0)
End Sub

Private Sub txtCuentaIngComision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaIngComisionDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaIngComision.Text = gCuenta
   txtCuentaIngComisionDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaIngComision.Text = fxgCntCuentaFormato(True, txtCuentaIngComision, 0)
End If
End Sub

Private Sub txtCuentaIngComision_LostFocus()
   txtCuentaIngComisionDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaIngComision, 0))
   txtCuentaIngComision.Text = fxgCntCuentaFormato(True, txtCuentaIngComision, 0)
End Sub

Private Sub txtCuentaIngComisionDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaIngRetiros.SetFocus
End Sub


Private Sub txtCuentaIngRetiros_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaIngRetirosDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaIngRetiros.Text = gCuenta
   txtCuentaIngRetirosDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaIngRetiros.Text = fxgCntCuentaFormato(True, txtCuentaIngRetiros, 0)
End If
End Sub


Private Sub txtCuentaIngRetiros_LostFocus()
   txtCuentaIngRetirosDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaIngRetiros, 0))
   txtCuentaIngRetiros.Text = fxgCntCuentaFormato(True, txtCuentaIngRetiros, 0)
End Sub

Private Sub txtCuentaIngRetirosDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaGstComision.SetFocus
End Sub

Private Sub txtCuentaGstComision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaGstComisionDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaGstComision.Text = gCuenta
   txtCuentaGstComisionDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaGstComision.Text = fxgCntCuentaFormato(True, txtCuentaGstComision, 0)
End If
End Sub


Private Sub txtCuentaGstComision_LostFocus()
   txtCuentaGstComisionDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaGstComision, 0))
   txtCuentaGstComision.Text = fxgCntCuentaFormato(True, txtCuentaGstComision, 0)
End Sub

Private Sub txtCuentaGstComisionDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaImpuestos.SetFocus
End Sub



Private Sub txtCuentaImpuestos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaImpuestosDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaImpuestos.Text = gCuenta
   txtCuentaImpuestosDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaImpuestos.Text = fxgCntCuentaFormato(True, txtCuentaImpuestos, 0)
End If
End Sub

Private Sub txtCuentaImpuestos_LostFocus()
   txtCuentaImpuestosDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaImpuestos, 0))
   txtCuentaImpuestos = fxgCntCuentaFormato(True, txtCuentaImpuestos, 0)
End Sub

Private Sub txtCuentaImpuestosDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaCod.SetFocus
End Sub




Private Sub txtCuentaRend_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaRendDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaRend = gCuenta
   txtCuentaRendDesc = fxgCntCuentaDesc(gCuenta)
   txtCuentaRend = fxgCntCuentaFormato(True, txtCuentaRend, 0)
End If
End Sub

Private Sub txtCuentaRend_LostFocus()
   txtCuentaRendDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaRend, 0))
   txtCuentaRend = fxgCntCuentaFormato(True, txtCuentaRend, 0)
End Sub

Private Sub txtCuentaRendDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaGasto.SetFocus
End Sub


Private Sub txtDescripcion_GotFocus()
vTipoBusca = "D"
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then tlb_ButtonClick tlb.Buttons(7)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboGrupo.SetFocus

End Sub


Private Sub txtInversion_GotFocus()
On Error GoTo vError
 txtInversion = CCur(txtInversion)
vError:
End Sub

Private Sub txtInversion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAseCod.SetFocus
End Sub

Private Sub txtInversion_LostFocus()
On Error GoTo vError
 txtInversion = Format(CCur(txtInversion), "Standard")
vError:
End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError
 txtMonto = CCur(txtMonto)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtInversion.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
 txtMonto = Format(CCur(txtMonto), "Standard")
vError:
End Sub


Private Function fxGuardar() As Integer
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If vGrid.Text = "" Then
   strSQL = "insert fnd_tabla_retiros(cod_operadora,cod_plan,desde,hasta,porcentaje, Aplicar_A, registro_usuario,registro_fecha)" _
          & " values(" & cboOperadora.ItemData(cboOperadora.ListIndex) & ",'" & txtCodigo & "',"
   vGrid.Col = 2
   strSQL = strSQL & CDbl(vGrid.Text) & ","
   vGrid.Col = 3
   strSQL = strSQL & CDbl(vGrid.Text) & ","
   vGrid.Col = 4
   strSQL = strSQL & CCur(vGrid.Text) & ",'"
   vGrid.Col = 5
   strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "','" & glogon.Usuario & "',dbo.MyGetdate())"
   Call ConectionExecute(strSQL)
   
   vGrid.Col = 1
   
   strSQL = "select max(cod_fnd_tabla_ret) as ultimo from fnd_tabla_retiros"
   Call OpenRecordSet(rs, strSQL)
    vGrid.Text = CStr(rs!ultimo)
   rs.Close
   
   Call Bitacora("Registra", "Multa x Retiro Anticipado # " & vGrid.Text)
   
 Else 'Actualizar
    vGrid.Col = 2
    strSQL = "update fnd_tabla_retiros set desde = " & CDbl(vGrid.Text) & ",hasta = "
    vGrid.Col = 3
    strSQL = strSQL & CDbl(vGrid.Text) & ", porcentaje = "
    vGrid.Col = 4
    strSQL = strSQL & CCur(vGrid.Text) & ", Aplicar_A = '"
    vGrid.Col = 5
    strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "', actualiza_usuario = '" & glogon.Usuario & "', actualiza_fecha = dbo.MyGetdate()"
    vGrid.Col = 1
    strSQL = strSQL & " where cod_fnd_tabla_ret = " & vGrid.Text
   
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Multa x Retiro Anticipado # " & vGrid.Text)
    
End If

vGrid.Col = 1
fxGuardar = vGrid.Text

Exit Function
   
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Function fxGuardarAum() As Integer
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardarAum = 0
vaGrid.Row = vaGrid.ActiveRow
vaGrid.Col = 1

If vaGrid.Text = "" Then
   strSQL = "insert FND_TABLA_AUMENTOS(cod_operadora,cod_plan,Tipo_Tasa,desde,hasta,plus,registro_usuario,registro_fecha, ID_PER_TASA)" _
          & " values(" & cboOperadora.ItemData(cboOperadora.ListIndex) & ",'" & txtCodigo & "','"
   vaGrid.Col = 2
   If vaGrid.Text = "" Then vaGrid.Text = "TBP"
   
   strSQL = strSQL & vaGrid.Text & "',"
   vaGrid.Col = 3
   strSQL = strSQL & CLng(vaGrid.Text) & ","
   vaGrid.Col = 4
   strSQL = strSQL & CLng(vaGrid.Text) & ","
   vaGrid.Col = 5
   strSQL = strSQL & CCur(vaGrid.Text) & ",'" & glogon.Usuario & "', dbo.MyGetdate(), " & txtR_Id.Text & ")"
   Call ConectionExecute(strSQL)
   
   vaGrid.Col = 1
   
   strSQL = "select max(COD_TABLA_AUM) as ultimo from FND_TABLA_AUMENTOS"
   Call OpenRecordSet(rs, strSQL)
    vaGrid.Text = CStr(rs!ultimo)
   rs.Close
   
   Call Bitacora("Registra", "Aumento de Tasa x Antiguedad # " & vaGrid.Text & ", Regla: " & txtR_Id.Text)
   
 Else 'Actualizar
    vaGrid.Col = 2
    If vaGrid.Text = "" Then vaGrid.Text = "TBP"
    
    strSQL = "update FND_TABLA_AUMENTOS set Tipo_Tasa = '" & vaGrid.Text & "', desde = "
    vaGrid.Col = 3
    strSQL = strSQL & CLng(vaGrid.Text) & ",hasta = "
    vaGrid.Col = 4
    strSQL = strSQL & CLng(vaGrid.Text) & ",plus = "
    vaGrid.Col = 5
    strSQL = strSQL & CCur(vaGrid.Text) & ", actualiza_usuario = '" & glogon.Usuario & "', actualiza_fecha = dbo.MyGetdate()"
    vaGrid.Col = 1
    strSQL = strSQL & " where COD_TABLA_AUM = " & vaGrid.Text
   
    
   
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Aumento de Tasa x Antiguedad # " & vaGrid.Text & ", Regla: " & txtR_Id.Text)
    
End If

vaGrid.Col = 1
fxGuardarAum = vaGrid.Text

Exit Function
   
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus
End Sub

Private Sub txtNuevaTasa_Change()
If Not IsNumeric(txtNuevaTasa) Then
   txtNuevaTasa = 0
End If

End Sub

Private Sub txtNuevaTasa_GotFocus()
txtNuevaTasa.Text = CCur(txtNuevaTasa.Text)
End Sub

Private Sub txtNuevaTasa_LostFocus()
txtNuevaTasa.Text = Format(txtNuevaTasa.Text, "Standard")
End Sub

Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboPlazo.SetFocus
End Sub

Private Sub txtVence_Plan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Código"
    gBusquedas.Col2Name = "Descripción"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_plan, Descripcion from fnd_Planes"
    gBusquedas.Filtro = " And Cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
                      & " And Estado = 'A' and Tipo_CDP = 0"
    frmBusquedas.Show vbModal

    If gBusquedas.Resultado <> "" Then
        txtVence_Plan.Text = gBusquedas.Resultado
        txtVence_PlanDesc.Text = gBusquedas.Resultado2
    End If
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

End Sub

Private Sub vaGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long, strSQL As String

On Error GoTo vError

If vaGrid.ActiveCol = vaGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarAum
  vaGrid.Row = vaGrid.ActiveRow
  vaGrid.Col = 1
  If vaGrid.MaxRows <= vaGrid.ActiveRow Then
    vaGrid.MaxRows = vaGrid.MaxRows + 1
    vaGrid.Row = vaGrid.MaxRows
  End If
End If

If vaGrid.ActiveCol = 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vaGrid.Col = vaGrid.ActiveCol
  vaGrid.Row = vaGrid.ActiveRow
  vaGrid.Text = vaGrid.Text
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vaGrid.MaxRows = vaGrid.MaxRows + 1
    vaGrid.InsertRows vaGrid.ActiveRow, 1
    vaGrid.Row = vaGrid.ActiveRow
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then

        vaGrid.Row = vaGrid.ActiveRow
        vaGrid.Col = 1

       If vaGrid.Text = "" Then Exit Sub

     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vaGrid.Col = 1
        strSQL = "delete FND_TABLA_AUMENTOS where cod_tabla_aum = " & vaGrid.Text
        Call ConectionExecute(strSQL)
        
        
        Call Bitacora("Elimina", "Aumento de Tasa x Antiguedad # " & vaGrid.Text)
        
        vaGrid.DeleteRows vaGrid.ActiveRow, 1
        vaGrid.MaxRows = vaGrid.MaxRows - 1
        If vaGrid.MaxRows = 0 Then vaGrid.MaxRows = 1
        
     End If
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

If vGrid.ActiveCol = 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = vGrid.Text
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then

        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1

       If vGrid.Text = "" Then Exit Sub

     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Col = 1
        strSQL = "delete FND_TABLA_RETIROS where cod_fnd_tabla_ret = " & vGrid.Text
        Call ConectionExecute(strSQL)
        
        
        Call Bitacora("Elimina", "Multa x Retiro Anticipado # " & vGrid.Text)
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
        
     End If
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbCargaEstados()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

vPaso = True

strSQL = "select P.COD_ESTADO,P.DESCRIPCION,F.cod_plan as Existe from  AFI_ESTADOS_PERSONA P" _
        & " left join FND_PLANES_ESTADOS E on P.COD_ESTADO  = E.COD_ESTADO  " _
        & " and  E.cod_plan = '" & txtCodigo & "' and E.cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) & " " _
        & " left Join FND_PLANES F  on E.Cod_plan = F.cod_plan and E.cod_operadora = F.Cod_Operadora" _
        & " order by existe desc,P.descripcion"
Call OpenRecordSet(rs, strSQL, 0)

lswEstados.ListItems.Clear

Do While Not rs.EOF
  Set itmX = lswEstados.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!cod_estado
      itmX.Checked = IIf(IsNull(rs!Existe), False, True)
      If itmX.Checked Then itmX.ForeColor = vbBlue
      
  rs.MoveNext
Loop
rs.Close


strSQL = "exec spFnd_Planes_Plazos_Consulta " & cboOperadora.ItemData(cboOperadora.ListIndex) & ", '" & txtCodigo.Text & "', 'T'"
Call OpenRecordSet(rs, strSQL, 0)

lswVence_Plazos.ListItems.Clear

Do While Not rs.EOF
  Set itmX = lswVence_Plazos.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!Plazo
      itmX.Checked = IIf((rs!Asignado = 1), vbChecked, vbUnchecked)
      If itmX.Checked Then itmX.ForeColor = vbBlue
      
  rs.MoveNext
Loop
rs.Close


vPaso = False

Me.MousePointer = vbDefault

End Sub


