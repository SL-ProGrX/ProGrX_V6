VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCxC_Clientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   11535
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   11280
      Top             =   0
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   10320
      TabIndex        =   0
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.TabControl ssTab 
      Height          =   6612
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   11292
      _Version        =   1572864
      _ExtentX        =   19918
      _ExtentY        =   11663
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
      Item(0).ControlCount=   10
      Item(0).Control(0)=   "txtRazon"
      Item(0).Control(1)=   "cboTipoId"
      Item(0).Control(2)=   "GroupBox2"
      Item(0).Control(3)=   "GroupBox3"
      Item(0).Control(4)=   "chkActivo"
      Item(0).Control(5)=   "cboClasificacion"
      Item(0).Control(6)=   "Label4(8)"
      Item(0).Control(7)=   "Label4(9)"
      Item(0).Control(8)=   "Label4(10)"
      Item(0).Control(9)=   "gbDireccion"
      Item(1).Caption =   "Operaciones"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "GroupBox4"
      Item(1).Control(1)=   "lswOperaciones"
      Item(1).Control(2)=   "Label2(7)"
      Item(1).Control(3)=   "txtMonto"
      Item(1).Control(4)=   "txtSaldo"
      Item(1).Control(5)=   "Label2(8)"
      Item(1).Control(6)=   "Label2(9)"
      Item(2).Caption =   "Contratos"
      Item(2).ControlCount=   7
      Item(2).Control(0)=   "txtDescripcion"
      Item(2).Control(1)=   "txtCodigo"
      Item(2).Control(2)=   "lswP"
      Item(2).Control(3)=   "lswC"
      Item(2).Control(4)=   "Label1(18)"
      Item(2).Control(5)=   "Label1(13)"
      Item(2).Control(6)=   "lswContratos"
      Item(3).Caption =   "Cuentas Bancarias"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "lswBancos"
      Item(4).Caption =   "Autorizados"
      Item(4).ControlCount=   12
      Item(4).Control(0)=   "feAut_Identificacion"
      Item(4).Control(1)=   "Label2(2)"
      Item(4).Control(2)=   "lswAutorizados"
      Item(4).Control(3)=   "feAut_Nombre"
      Item(4).Control(4)=   "Label2(3)"
      Item(4).Control(5)=   "chkAutorizadoActivo"
      Item(4).Control(6)=   "feAut_Profesion"
      Item(4).Control(7)=   "Label2(4)"
      Item(4).Control(8)=   "feAut_Condicion"
      Item(4).Control(9)=   "Label2(5)"
      Item(4).Control(10)=   "btnAutorizadoRegistra"
      Item(4).Control(11)=   "btnAutorizadoElimina"
      Begin XtremeSuiteControls.ListView lswAutorizados 
         Height          =   4692
         Left            =   -69880
         TabIndex        =   27
         Top             =   1800
         Visible         =   0   'False
         Width           =   11172
         _Version        =   1572864
         _ExtentX        =   19706
         _ExtentY        =   8276
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
         Sorted          =   -1  'True
      End
      Begin XtremeSuiteControls.ListView lswBancos 
         Height          =   6012
         Left            =   -69880
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   11052
         _Version        =   1572864
         _ExtentX        =   19494
         _ExtentY        =   10604
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
         Sorted          =   -1  'True
      End
      Begin XtremeSuiteControls.ListView lswC 
         Height          =   1692
         Left            =   -69880
         TabIndex        =   38
         Top             =   4800
         Visible         =   0   'False
         Width           =   11052
         _Version        =   1572864
         _ExtentX        =   19494
         _ExtentY        =   2984
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
         Sorted          =   -1  'True
      End
      Begin XtremeSuiteControls.ListView lswP 
         Height          =   1812
         Left            =   -69880
         TabIndex        =   37
         Top             =   2640
         Visible         =   0   'False
         Width           =   11052
         _Version        =   1572864
         _ExtentX        =   19494
         _ExtentY        =   3196
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
         Sorted          =   -1  'True
      End
      Begin XtremeSuiteControls.ListView lswContratos 
         Height          =   1572
         Left            =   -69880
         TabIndex        =   39
         Top             =   360
         Visible         =   0   'False
         Width           =   11052
         _Version        =   1572864
         _ExtentX        =   19494
         _ExtentY        =   2773
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
         Sorted          =   -1  'True
      End
      Begin XtremeSuiteControls.ListView lswOperaciones 
         Height          =   2772
         Left            =   -70000
         TabIndex        =   32
         Top             =   3240
         Visible         =   0   'False
         Width           =   11172
         _Version        =   1572864
         _ExtentX        =   19706
         _ExtentY        =   4890
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
      Begin XtremeSuiteControls.CheckBox chkAutorizadoActivo 
         Height          =   252
         Left            =   -61720
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Activo?"
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
      End
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   252
         Left            =   2640
         TabIndex        =   9
         Top             =   1200
         Width           =   1092
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Activo?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         Alignment       =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   2292
         Left            =   -69880
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   11052
         _Version        =   1572864
         _ExtentX        =   19494
         _ExtentY        =   4043
         _StockProps     =   79
         Caption         =   "Crédito y Adelantos"
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
         Begin XtremeSuiteControls.CheckBox chkExcento 
            Height          =   252
            Left            =   840
            TabIndex        =   23
            Top             =   360
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cliente Exento?"
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
         End
         Begin XtremeSuiteControls.CheckBox chkCredito 
            Height          =   252
            Left            =   840
            TabIndex        =   24
            Top             =   720
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Crédito Cerrado?"
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
         End
         Begin XtremeSuiteControls.CheckBox chkAdelantoAplica 
            Height          =   252
            Left            =   6000
            TabIndex        =   25
            Top             =   360
            Width           =   3372
            _Version        =   1572864
            _ExtentX        =   5948
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplica Adelantos?"
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
         End
         Begin XtremeSuiteControls.CheckBox chkAdelantoModifica 
            Height          =   252
            Left            =   6000
            TabIndex        =   26
            Top             =   720
            Width           =   3372
            _Version        =   1572864
            _ExtentX        =   5948
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Modifica Porcentaje de Adelanto?"
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
         End
         Begin XtremeSuiteControls.CheckBox chkAdelantoComisionApl 
            Height          =   252
            Left            =   6000
            TabIndex        =   28
            Top             =   1560
            Width           =   3372
            _Version        =   1572864
            _ExtentX        =   5948
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplica Comisión por Adelantos?"
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
         End
         Begin XtremeSuiteControls.CheckBox chkRol_Pagador 
            Height          =   252
            Left            =   840
            TabIndex        =   46
            Top             =   1560
            Width           =   4212
            _Version        =   1572864
            _ExtentX        =   7429
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "En Descuento de Facturas es un Pagador?"
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
         End
         Begin XtremeSuiteControls.CheckBox chkRol_Autorizador 
            Height          =   252
            Left            =   840
            TabIndex        =   47
            Top             =   1920
            Width           =   3732
            _Version        =   1572864
            _ExtentX        =   6583
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "En Descuento de Facturas es un Autorizado?"
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
         End
         Begin XtremeSuiteControls.FlatEdit txtCreditoLimite 
            Height          =   312
            Left            =   1080
            TabIndex        =   87
            Top             =   1080
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAdelantoPorcentaje 
            Height          =   312
            Left            =   6360
            TabIndex        =   88
            Top             =   1080
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
            Text            =   "90"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAdelantoComision 
            Height          =   312
            Left            =   6360
            TabIndex        =   89
            Top             =   1920
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            Caption         =   "Porcentaje de Comisión por Adelanto"
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
            Index           =   6
            Left            =   7080
            TabIndex        =   45
            Top             =   1920
            Width           =   3972
         End
         Begin VB.Label Label2 
            Caption         =   "Limite de Crédito"
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
            Left            =   2640
            TabIndex        =   8
            Top             =   1080
            Width           =   2532
         End
         Begin VB.Label Label2 
            Caption         =   "Porcentaje de Adelanto Autorizado"
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
            Index           =   0
            Left            =   7080
            TabIndex        =   3
            Top             =   1080
            Width           =   3612
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1572
         Left            =   240
         TabIndex        =   4
         Top             =   5040
         Width           =   10812
         _Version        =   1572864
         _ExtentX        =   19071
         _ExtentY        =   2773
         _StockProps     =   79
         Caption         =   "Datos Adicionales"
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
         Begin XtremeSuiteControls.ComboBox cboSexo 
            Height          =   312
            Left            =   1440
            TabIndex        =   29
            Top             =   360
            Width           =   2052
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
         End
         Begin XtremeSuiteControls.ComboBox cboEstado 
            Height          =   312
            Left            =   1440
            TabIndex        =   30
            Top             =   720
            Width           =   2052
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
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFecNac 
            Height          =   312
            Left            =   1440
            TabIndex        =   31
            Top             =   1080
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
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
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   1092
            Left            =   3840
            TabIndex        =   80
            Top             =   360
            Width           =   6852
            _Version        =   1572864
            _ExtentX        =   12086
            _ExtentY        =   1926
            _StockProps     =   77
            ForeColor       =   0
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   16
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Sexo"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   15
            Left            =   120
            TabIndex        =   78
            Top             =   720
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Est. Civil"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   14
            Left            =   120
            TabIndex        =   77
            Top             =   1080
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Fec. Nac."
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
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1812
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   10812
         _Version        =   1572864
         _ExtentX        =   19071
         _ExtentY        =   3196
         _StockProps     =   79
         Caption         =   "Información de Contacto"
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
         Begin XtremeSuiteControls.FlatEdit txtWebSite 
            Height          =   330
            Left            =   5400
            TabIndex        =   48
            Top             =   360
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEmail 
            Height          =   330
            Left            =   5400
            TabIndex        =   49
            Top             =   720
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEmail2 
            Height          =   330
            Left            =   5400
            TabIndex        =   50
            Top             =   1080
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAptoPostal 
            Height          =   330
            Left            =   5400
            TabIndex        =   51
            Top             =   1440
            Width           =   5292
            _Version        =   1572864
            _ExtentX        =   9334
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelefono 
            Height          =   330
            Left            =   1440
            TabIndex        =   56
            Top             =   360
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelefono2 
            Height          =   330
            Left            =   1440
            TabIndex        =   57
            Top             =   720
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelFax 
            Height          =   330
            Left            =   1440
            TabIndex        =   58
            Top             =   1440
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCelular 
            Height          =   330
            Left            =   1440
            TabIndex        =   62
            Top             =   1080
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   0
            Left            =   0
            TabIndex        =   63
            Top             =   1080
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Tel. Móvil"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   1
            Left            =   0
            TabIndex        =   61
            Top             =   360
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Teléfono (1)"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   2
            Left            =   0
            TabIndex        =   60
            Top             =   720
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Teléfono (2)"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   3
            Left            =   0
            TabIndex        =   59
            Top             =   1440
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Tel. Fax"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   4
            Left            =   3840
            TabIndex        =   55
            Top             =   360
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Web Site"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   5
            Left            =   3840
            TabIndex        =   54
            Top             =   720
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Email (1)"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   6
            Left            =   3840
            TabIndex        =   53
            Top             =   1080
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Email (2)"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   7
            Left            =   3840
            TabIndex        =   52
            Top             =   1440
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Apto. Postal"
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
      Begin XtremeSuiteControls.ComboBox cboTipoId 
         Height          =   312
         Left            =   1680
         TabIndex        =   10
         Top             =   480
         Width           =   2052
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
      End
      Begin XtremeSuiteControls.ComboBox cboClasificacion 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   840
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
      End
      Begin XtremeSuiteControls.FlatEdit feAut_Identificacion 
         Height          =   330
         Left            =   -69640
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnAutorizadoRegistra 
         Height          =   312
         Left            =   -61600
         TabIndex        =   13
         ToolTipText     =   "Detalle de Facturas"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Registra"
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
      Begin XtremeSuiteControls.PushButton btnAutorizadoElimina 
         Height          =   312
         Left            =   -60520
         TabIndex        =   14
         ToolTipText     =   "Detalle de Facturas"
         Top             =   1320
         Visible         =   0   'False
         Width           =   972
         _Version        =   1572864
         _ExtentX        =   1714
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Elimina"
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
      Begin XtremeSuiteControls.FlatEdit feAut_Nombre 
         Height          =   330
         Left            =   -67840
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   5892
         _Version        =   1572864
         _ExtentX        =   10393
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit feAut_Profesion 
         Height          =   330
         Left            =   -69640
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit feAut_Condicion 
         Height          =   330
         Left            =   -67840
         TabIndex        =   21
         Top             =   1320
         Visible         =   0   'False
         Width           =   5892
         _Version        =   1572864
         _ExtentX        =   10393
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRazon 
         Height          =   912
         Left            =   5640
         TabIndex        =   68
         Top             =   480
         Width           =   5292
         _Version        =   1572864
         _ExtentX        =   9334
         _ExtentY        =   1609
         _StockProps     =   77
         ForeColor       =   0
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
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbDireccion 
         Height          =   2052
         Left            =   240
         TabIndex        =   69
         Top             =   3360
         Width           =   10812
         _Version        =   1572864
         _ExtentX        =   19071
         _ExtentY        =   3619
         _StockProps     =   79
         Caption         =   "Dirección"
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
         Begin XtremeSuiteControls.ComboBox cboProvincia 
            Height          =   312
            Left            =   1440
            TabIndex        =   70
            Top             =   480
            Width           =   2052
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
               Weight          =   400
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
            Left            =   1440
            TabIndex        =   71
            Top             =   840
            Width           =   2052
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
               Weight          =   400
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
            Left            =   1440
            TabIndex        =   72
            Top             =   1200
            Width           =   2052
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
            Height          =   1092
            Left            =   3840
            TabIndex        =   73
            Top             =   480
            Width           =   6852
            _Version        =   1572864
            _ExtentX        =   12086
            _ExtentY        =   1926
            _StockProps     =   77
            ForeColor       =   0
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   13
            Left            =   120
            TabIndex        =   76
            Top             =   480
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   12
            Left            =   120
            TabIndex        =   75
            Top             =   840
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   11
            Left            =   120
            TabIndex        =   74
            Top             =   1200
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
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
      End
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         Height          =   312
         Left            =   -61000
         TabIndex        =   85
         Top             =   6120
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   -65320
         TabIndex        =   86
         Top             =   6120
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   -66160
         TabIndex        =   90
         Top             =   2040
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   -64360
         TabIndex        =   91
         Top             =   2040
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1572864
         _ExtentX        =   9758
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   10
         Left            =   240
         TabIndex        =   66
         Top             =   840
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Clasificacion"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   9
         Left            =   240
         TabIndex        =   65
         Top             =   480
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tipo Id"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   8
         Left            =   4080
         TabIndex        =   64
         Top             =   480
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Razón Social"
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   9
         Left            =   -64120
         TabIndex        =   35
         Top             =   6120
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Monto Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   8
         Left            =   -67720
         TabIndex        =   34
         Top             =   6120
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label2 
         Caption         =   "Operaciones activas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   7
         Left            =   -69640
         TabIndex        =   33
         Top             =   2520
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.Label Label2 
         Caption         =   "Condición (Representación)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   -67840
         TabIndex        =   22
         Top             =   1080
         Visible         =   0   'False
         Width           =   3612
      End
      Begin VB.Label Label2 
         Caption         =   "Profesión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   -69640
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   -67840
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label2 
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
         Height          =   252
         Index           =   2
         Left            =   -69640
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "Pagadores:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   13
         Left            =   -69760
         TabIndex        =   7
         Top             =   2400
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "Cargos por Suscripción:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   18
         Left            =   -69760
         TabIndex        =   6
         Top             =   4560
         Visible         =   0   'False
         Width           =   3972
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11400
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Clientes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Clientes.frx":3492
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Clientes.frx":6924
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Clientes.frx":6A42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   3480
      TabIndex        =   81
      Top             =   480
      Width           =   6732
      _Version        =   1572864
      _ExtentX        =   11874
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   1440
      TabIndex        =   83
      Top             =   480
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.GroupBox gbIdCambio 
      Height          =   4452
      Left            =   240
      TabIndex        =   40
      Top             =   2400
      Visible         =   0   'False
      Width           =   11052
      _Version        =   1572864
      _ExtentX        =   19494
      _ExtentY        =   7853
      _StockProps     =   79
      Caption         =   "Cambio  y Unificación de Identificación de la Persona"
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
      Begin XtremeSuiteControls.PushButton btnIdCambio 
         Height          =   372
         Left            =   7200
         TabIndex        =   41
         Top             =   2400
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cambio de Id"
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
      End
      Begin XtremeSuiteControls.PushButton btnIdCerrar 
         Height          =   372
         Left            =   8760
         TabIndex        =   42
         Top             =   2400
         Width           =   1092
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cerrar"
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
      End
      Begin XtremeSuiteControls.FlatEdit txtIdReferencia 
         Height          =   312
         Left            =   3120
         TabIndex        =   67
         Top             =   1800
         Width           =   6732
         _Version        =   1572864
         _ExtentX        =   11874
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIdNew 
         Height          =   312
         Left            =   1080
         TabIndex        =   82
         Top             =   1800
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Referencia encontrada?"
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
         Index           =   16
         Left            =   3240
         TabIndex        =   44
         Top             =   1560
         Width           =   2172
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Nueva Identificación"
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
         Index           =   19
         Left            =   1080
         TabIndex        =   43
         Top             =   1560
         Width           =   2052
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   330
      Left            =   120
      TabIndex        =   92
      Top             =   0
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
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
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbAux 
      Height          =   330
      Left            =   10440
      TabIndex        =   93
      Top             =   0
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   252
      Index           =   17
      Left            =   240
      TabIndex        =   84
      Top             =   480
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cliente"
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
   Begin VB.Image imgId_Cambio 
      Height          =   252
      Left            =   10920
      Picture         =   "frmCxC_Clientes.frx":6B79
      Stretch         =   -1  'True
      ToolTipText     =   "Cambiar el Número de Identificación"
      Top             =   480
      Width           =   252
   End
End
Attribute VB_Name = "frmCxC_Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCedula As String
Dim vFechaActual As Date
Dim vScroll As Boolean, vTipoJuridica As Integer, vPaso As Boolean


Private Sub sbAutorizado_Mantenimiento(Optional pMovimiento As String = "I")
Dim strSQL As String

On Error GoTo vError

If Len(feAut_Identificacion.Text) = 0 Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spCxC_Persona_Autorizador_Registra '" & txtCedula.Text & "','" & feAut_Identificacion.Text _
        & "','" & feAut_Profesion.Text _
        & "','" & feAut_Condicion.Text _
        & "'," & chkAutorizadoActivo.Value _
        & ",'" & glogon.Usuario & "','" & pMovimiento & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Call sbAutorizadores_Load

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnAutorizadoElimina_Click()
Call sbAutorizado_Mantenimiento("E")
End Sub

Private Sub btnAutorizadoRegistra_Click()
Call sbAutorizado_Mantenimiento("I")
End Sub

Private Sub btnIdCambio_Click()
Dim strSQL As String, i As Integer

On Error GoTo vError

If Trim(txtIdNew.Text) = "" Then Exit Sub

i = MsgBox("Esta seguro que cambiar el Id:" & txtCedula.Text & " por: " & txtIdNew.Text, vbYesNo)
If i = vbNo Then Exit Sub


strSQL = "exec spCxC_Persona_Id_Cambio '" & txtCedula.Text & "','" & txtIdNew.Text & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
If Not glogon.error Then
  Call Bitacora("Cambio", "Identificación: " & txtCedula.Text & " -> " & txtIdNew.Text)
End If

Call imgId_Cambio_Click
Call sbConsulta(txtIdNew.Text)
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnIdCerrar_Click()
Call imgId_Cambio_Click
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

Private Sub cboDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
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


Private Sub cboSexo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus
End Sub

Private Sub cboTipoId_Click()
If vPaso Then Exit Sub

If cboTipoId.ItemData(cboTipoId.ListIndex) = vTipoJuridica Then
   txtRazon.Enabled = True
   dtpFecNac.Enabled = False
   cboSexo.Enabled = False
   cboEstado.Enabled = False

   txtRazon.BackColor = vbWhite
  
Else
   txtRazon.Enabled = False
   dtpFecNac.Enabled = True
   cboSexo.Enabled = True
   cboEstado.Enabled = True
   txtRazon.BackColor = RGB(84, 153, 199)
End If


End Sub




Private Sub feAut_Identificacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then feAut_Nombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from CxC_Personas"
  gBusquedas.Filtro = " and tipo_Id = 1 and Rol_Autorizador = 1"
  frmBusquedas.Show vbModal
  feAut_Identificacion.Text = gBusquedas.Resultado
  feAut_Nombre.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub feAut_Nombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then feAut_Nombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from CxC_Personas"
  gBusquedas.Filtro = " and tipo_Id = 1"
  frmBusquedas.Show vbModal
  feAut_Identificacion.Text = gBusquedas.Resultado
  feAut_Nombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 cedula from CxC_Personas"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cedula > '" & txtCedula.Text & "' order by cedula asc"
    Else
       strSQL = strSQL & " where cedula < '" & txtCedula.Text & "' order by cedula desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCedula.Text = rs!Cedula
      Call txtCedula_LostFocus
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
vModulo = 31
End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 31

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 imgId_Cambio.Tag = 1

 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 
ssTab.Item(0).Selected = True

gbIdCambio.Left = ssTab.Left
gbIdCambio.top = ssTab.top
gbIdCambio.Height = ssTab.Height
gbIdCambio.Width = ssTab.Width
gbIdCambio.Visible = False

vEdita = False

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
vCedula = ""
txtCedula = ""

txtNombre = ""
txtTelefono.Text = ""
txtTelefono2.Text = ""
txtCelular.Text = ""
txtRazon.Text = ""
txtWebSite.Text = ""
txtEmail.Text = ""
txtEmail2.Text = ""
txtAptoPostal.Text = ""

txtDireccion = ""

chkActivo.Value = xtpChecked

dtpFecNac.Value = vFechaActual
'cboEstado.Text = "Soltero"
'cboSexo.Text = "Masculino"
txtNotas.Text = ""

chkCredito.Value = vbUnchecked
chkExcento.Value = vbUnchecked

chkRol_Pagador.Value = vbUnchecked
chkRol_Autorizador.Value = vbUnchecked

txtCreditoLimite.Text = "0"

chkAdelantoAplica.Value = vbUnchecked
chkAdelantoModifica.Value = vbUnchecked
chkAdelantoComisionApl.Value = xtpUnchecked

txtAdelantoPorcentaje.Text = "0"
txtAdelantoComision.Text = "0"


ssTab.Item(0).Selected = True
ssTab.Item(1).Enabled = False
ssTab.Item(2).Enabled = False
ssTab.Item(3).Enabled = False
ssTab.Item(4).Enabled = False

tlbAux.Buttons.Item(1).Enabled = False 'Nuevo
tlbAux.Buttons.Item(3).Enabled = False 'Borrar


End Sub



Private Sub sbConsultaContratoDetalle()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

'Carga Pagadores
lswP.ListItems.Clear
strSQL = " select P.nombre, C.*" _
     & " from CxC_Personas P inner join CxC_Personas_Contratos_Pagadores C on P.cedula = C.cedula_pagador" _
     & " where C.cod_contrato = '" & txtCodigo.Text & "' and C.cedula = '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    Set itmX = lswP.ListItems.Add(, , rs!cedula_pagador)
        itmX.SubItems(1) = rs!Nombre
        itmX.SubItems(2) = rs!Registro_Usuario & "..." & rs!Registro_Fecha
    rs.MoveNext
Loop
rs.Close


'Carga Cargos de Suscripción
lswC.ListItems.Clear
strSQL = " select C.descripcion,S.*" _
       & " from CxC_Cargos C inner join CxC_Personas_Contratos_Suscripciones S on C.cod_cargo = S.cod_cargo" _
       & " where S.cod_contrato = '" & txtCodigo.Text & "' and S.cedula = '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    Set itmX = lswC.ListItems.Add(, , rs!COD_CARGO)
        itmX.SubItems(1) = rs!Descripcion
        
        Select Case rs!Tipo
           Case "P"
             itmX.SubItems(2) = "Porcentual"
           Case "M"
             itmX.SubItems(2) = "Monto"
        End Select
        
        
        Select Case rs!Frecuencia_Tipo
          Case "O"
            itmX.SubItems(4) = "Operación"
          Case "D"
            itmX.SubItems(4) = "Días"
        End Select
        
        itmX.SubItems(3) = Format(rs!Valor, "Standard")
        itmX.SubItems(5) = rs!Frecuencia_dias
        itmX.SubItems(6) = Format(rs!Recaudado, "Standard")
        itmX.SubItems(7) = Format(rs!Pago_Ultimo, "dd/mm/yyyy")
        itmX.SubItems(8) = Format(rs!Pago_Proximo, "dd/mm/yyyy")
        itmX.SubItems(9) = rs!Modifica
        itmX.Checked = True
    rs.MoveNext
Loop
rs.Close


vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub imgId_Cambio_Click()
If gbIdCambio.Visible Then
    gbIdCambio.Visible = False
    ssTab.Visible = True
    gbIdCambio.top = gbIdCambio.top
Else
    gbIdCambio.Visible = True
    ssTab.Visible = False
End If
End Sub

Private Sub lswAutorizados_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
On Error GoTo vError

feAut_Identificacion.Text = Item.Text
feAut_Nombre.Text = Item.SubItems(1)
feAut_Profesion.Text = Item.SubItems(2)
feAut_Condicion.Text = Item.SubItems(3)

If Item.SubItems(4) = "Sí" Then
    chkAutorizadoActivo.Value = xtpChecked
Else
    chkAutorizadoActivo.Value = xtpUnchecked
End If


vError:
End Sub






Private Sub sbAutorizadores_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass
vPaso = True


feAut_Identificacion.Text = ""
feAut_Nombre.Text = ""
feAut_Profesion.Text = ""
feAut_Condicion.Text = ""

With lswAutorizados
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Identificación", 1200
    .ColumnHeaders.Add , , "Nombre", 3200
    .ColumnHeaders.Add , , "Profesión", 1200
    .ColumnHeaders.Add , , "Condición", 3200
    .ColumnHeaders.Add , , "Activo?", 1200, vbCenter

strSQL = "select Per.Nombre, A.*" _
       & " from CxC_Personas Per inner join  CxC_Personas_Autorizados A on Per.Cedula = A.cedula_Autorizado" _
       & " where A.cedula = '" & Trim(txtCedula) & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    Set itmX = .ListItems.Add(, , rs!Cedula_Autorizado)
        itmX.SubItems(1) = Trim(rs!Nombre)
        itmX.SubItems(2) = rs!Profesion
        itmX.SubItems(3) = rs!Condicion
        itmX.SubItems(4) = IIf((rs!Activo = 1), "Sí", "No")
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


Private Sub lswContratos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub
If lswContratos.ListItems.Count <= 0 Then Exit Sub

txtCodigo.Text = Item.SubItems(1)
txtDescripcion.Text = Item.SubItems(2)

Call sbConsultaContratoDetalle
End Sub





Private Sub ssTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curMonto As Currency, curSaldo As Currency


If vCedula = "" Then
  ssTab.Item(0).Selected = True
  tlbAux.Buttons.Item(1).Enabled = False 'Nuevo
  tlbAux.Buttons.Item(3).Enabled = False 'Borrar
  Exit Sub
End If

Me.MousePointer = vbHourglass

If Item.Index > 0 Then
   tlbAux.Buttons.Item(1).Enabled = True 'Nuevo
   tlbAux.Buttons.Item(3).Enabled = True 'Borrar
End If

vPaso = True
Select Case Item.Index
   Case 1 'Operaciones
       tlbAux.Buttons.Item(3).Enabled = False 'Borrar
       tlbAux.Buttons.Item(1).ToolTipText = "Nueva : Operación"
       tlbAux.Buttons.Item(3).ToolTipText = "Borra : Operación"
      
       vPaso = True
       lswOperaciones.ListItems.Clear
       curMonto = 0
       curSaldo = 0
       
       strSQL = "exec spCxC_PersonasCuentas '" & txtCedula.Text & "','A'"
       Call OpenRecordSet(rs, strSQL)
       Do While Not rs.EOF
         Set itmX = lswOperaciones.ListItems.Add(, , rs!Operacion)
             itmX.SubItems(1) = rs!Num_Documento
             itmX.SubItems(2) = Format(rs!Activa_Fecha, "dd/mm/yyyy")
             itmX.SubItems(3) = Format(rs!Fecha_Vencimiento, "dd/mm/yyyy")
             itmX.SubItems(4) = Format(rs!Fecha_Pago, "dd/mm/yyyy")
             itmX.SubItems(5) = Format(rs!Monto, "Standard")
             itmX.SubItems(6) = Format(rs!Saldo, "Standard")
             itmX.SubItems(7) = rs!Estado
             itmX.SubItems(8) = rs!ConceptoDesc
             itmX.SubItems(9) = rs!Nombre_Pagador
             itmX.SubItems(10) = rs!OficinaDesc
             
             curMonto = curMonto + rs!Monto
             curSaldo = curSaldo + rs!Saldo

          rs.MoveNext
       Loop
       rs.Close

    txtMonto.Text = Format(curMonto, "Standard")
    txtSaldo.Text = Format(curSaldo, "Standard")
       
       
       vPaso = False
   
   Case 2 'Contratos
       
       tlbAux.Buttons.Item(1).ToolTipText = "Nuevo : Contrato"
       tlbAux.Buttons.Item(3).ToolTipText = "Borra : Contrato"

       vPaso = True
       lswContratos.ListItems.Clear
       strSQL = " select P.descripcion,C.*" _
              & " from CxC_Contratos P inner join CxC_Personas_Contratos C on P.cod_contrato = C.cod_contrato" _
              & " where C.cedula = '" & vCedula & "' order by C.Activo desc,C.Registro_Fecha desc"
       Call OpenRecordSet(rs, strSQL)
       Do While Not rs.EOF
         Set itmX = lswContratos.ListItems.Add(, , rs!Contrato_Num)
             itmX.SubItems(1) = rs!COD_CONTRATO
             itmX.SubItems(2) = rs!Descripcion
             itmX.SubItems(3) = IIf(rs!Activo = 1, "Sí", "No")
             
             itmX.SubItems(4) = rs!Plazo
             itmX.SubItems(5) = Format(rs!Tasa_Corriente, "Standard")
             itmX.SubItems(6) = Format(rs!Tasa_Mora, "Standard")
             
             If rs!Contrato_Tipo = "D" Then
                 itmX.SubItems(7) = Format(rs!Contrato_Vence, "dd/mm/yyyy")
             Else
                 itmX.SubItems(7) = "Indefinido"
             End If
             
             itmX.SubItems(8) = rs!Registro_Usuario & "..." & rs!Registro_Fecha
          
          rs.MoveNext
       Loop
       rs.Close
       
       txtCodigo.Text = ""
       txtDescripcion.Text = ""
       lswP.ListItems.Clear
       lswC.ListItems.Clear
       
       vPaso = False
       
   Case 3 'Cuentas Bancarias
      
       tlbAux.Buttons.Item(1).ToolTipText = "Nuevo : Cuenta Bancaria"
       tlbAux.Buttons.Item(3).ToolTipText = "Borra : Cuenta Bancaria"
       
        lswBancos.ListItems.Clear
    
        strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
               & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
               & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
               & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
               & " where C.Identificacion = '" & Trim(txtCedula) & "' and C.Modulo = 'CxC'"
    
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
           Set itmX = lswBancos.ListItems.Add(, , rs!CUENTA_INTERNA)
               itmX.SubItems(1) = Trim(rs!Banco)
               itmX.SubItems(2) = rs!TipoDesc
               itmX.SubItems(3) = rs!COD_DIVISA
               itmX.SubItems(4) = rs!CUENTA_INTERBANCA
               itmX.SubItems(5) = IIf(rs!Activa = 1, "Activa", "Cerrada")
               itmX.SubItems(6) = rs!Registro_Fecha & ""
               itmX.SubItems(7) = rs!Registro_Usuario & ""
         
           rs.MoveNext
        Loop
        rs.Close
        
        
    Case 4 'Autorizados
      Call sbAutorizadores_Load
    
    
End Select

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

TimerX.Enabled = False
TimerX.Interval = 0

vFechaActual = Format(fxFechaServidor, "dd/mm/yyyy")


With lswOperaciones.ColumnHeaders
    .Clear
    .Add , , "No.Operación", 1300
    .Add , , "No.Documeno", 1300
    .Add , , "Fec.Activa", 1400, vbCenter
    .Add , , "Fec.Vence", 1400, vbCenter
    .Add , , "Fec.Pago", 1400, vbCenter
    .Add , , "Monto", 1400, vbRightJustify
    .Add , , "Saldo", 1400, vbRightJustify
    .Add , , "Estado", 1200, vbCenter
    .Add , , "Concepto", 2000
    .Add , , "Pagador", 3000
    .Add , , "Oficina", 3000
End With

With lswBancos.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 2500
    .Add , , "Banco", 3500
    .Add , , "Tipo", 1100, vbCenter
    .Add , , "Divisa", 1100, vbCenter
    .Add , , "Interbanca", 2500
    .Add , , "Activa", 1100, vbCenter
    .Add , , "Fecha", 2500
    .Add , , "Usuario", 2500
End With


With lswContratos.ColumnHeaders
    .Clear
    .Add , , "No.Contrato", 1300, vbCenter
    .Add , , "Código", 1100, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Activo?", 1000, vbCenter
    .Add , , "Plazo", 900, vbCenter
    .Add , , "Tasa Cor.", 1100, vbRightJustify
    .Add , , "Tasa Mor.", 1100, vbRightJustify
    .Add , , "Vence", 2400, vbCenter
    .Add , , "Registro", 4000
End With


With lswP.ColumnHeaders
    .Clear
    .Add , , "Identificación", 1400
    .Add , , "Nombre", 3000
    .Add , , "Registro", 4000
End With

With lswC.ColumnHeaders
    .Clear
    .Add , , "Código", 1100, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Tipo", 1200, vbCenter
    .Add , , "Valor", 1200, vbRightJustify
    .Add , , "Frecuencia", 1200, vbRightJustify
    .Add , , "Frec.(Dias)", 1200, vbCenter
    .Add , , "Recaudado", 2400, vbRightJustify
    .Add , , "Pago Ult.", 2400, vbCenter
    .Add , , "Pago Prox.", 2400, vbCenter
    .Add , , "Modifica?", 1200, vbCenter
End With


'Listas con CheckBoxes
lswBancos.Checkboxes = True
lswC.Checkboxes = True
lswP.Checkboxes = True



'Sexo
cboSexo.Clear
cboSexo.AddItem "Masculino"
cboSexo.ItemData(cboSexo.ListCount - 1) = "M"
cboSexo.AddItem "Femenido"
cboSexo.ItemData(cboSexo.ListCount - 1) = "F"
Call sbCboAsignaDato(cboSexo, "Femenino", True, "F")
 
strSQL = "select rtrim(Estado_Civil) as 'IdX', rtrim(DESCRIPCION) as 'ItmX'" _
       & " from SYS_ESTADO_CIVIL where ACTIVO = 1"
Call sbCbo_Llena_New(cboEstado, strSQL, False, True)
 
'Revisa cual Tipo de Identificacion es Juridica (Solo es Valido la Primera)
vTipoJuridica = 0
strSQL = "select TIPO_ID from AFI_TIPOS_IDS where Tipo_Personeria = 'J'"
rs.Open strSQL, glogon.Conection
If Not rs.EOF And Not rs.BOF Then
    vTipoJuridica = rs!Tipo_Id
End If
rs.Close

'Carga Tipos de Identificacion
vPaso = True
strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS"
    Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)
vPaso = False
Call cboTipoId_Click

'Carga Clasificaciones de Clientes
strSQL = "select rtrim(cod_categoria) as 'IdX' , rtrim(descripcion) as ItmX from CxC_Categoria_Clientes"
Call sbCbo_Llena_New(cboClasificacion, strSQL, False, True)

'Carga combo de Provincias
vPaso = True
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)
vPaso = False
 
Call sbLimpiaPantalla

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCedula.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCedula.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCedula = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCedula)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "nombre"
       gBusquedas.Orden = "nombre"
       gBusquedas.Consulta = "select cedula,nombre from CxC_Personas"
       frmBusquedas.Show vbModal
       txtCedula.SetFocus
       txtCedula = gBusquedas.Resultado
       txtNombre.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(pCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String


On Error GoTo vError

If Not fxSIFValidaCadena(txtCedula.Text) Then
   Exit Sub
End If

'Verifica el Enlace con Afiliación y CxP
Call sbFichaCliente(pCodigo)

Me.MousePointer = vbHourglass

strSQL = "select P.*,rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
       & ",Tid.Descripcion as TipoIdDesc,Tid.Tipo_Personeria, rtrim(Cat.Descripcion) as 'CatDesc'" _
       & ",isnull(Ec.Descripcion,'No. Identificado') as 'EstadoCivilDesc' " _
       & " from CxC_Personas P " _
       & " left join Provincias Prov on P.Provincia = Prov.Provincia" _
       & " left join Cantones Cant on P.Provincia = Cant.Provincia and P.Canton = Cant.Canton" _
       & " left join Distritos Dist on P.Provincia = Dist.Provincia and P.Canton = Dist.Canton and P.distrito = Dist.distrito" _
       & " left join AFI_TIPOS_IDS Tid on P.tipo_id = Tid.tipo_id" _
       & " left join CxC_Categoria_Clientes Cat on P.cod_categoria = Cat.cod_Categoria" _
       & " left join SYS_ESTADO_CIVIL Ec on P.EstadoCivil = Ec.Estado_Civil" _
       & " where P.cedula = '" & pCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  vCedula = rs!Cedula
  txtCedula = rs!Cedula
  
  txtIdNew.Text = ""
  txtIdReferencia.Text = ""

   If Not IsNull(rs!TipoIdDesc) Then
           cboTipoId.Text = Trim(rs!TipoIdDesc)
   End If

  txtNombre = rs!Nombre & ""
  

  txtTelefono.Text = rs!telefono1 & ""
  txtTelefono2.Text = rs!telefono2 & ""
  txtCelular.Text = rs!Celular & ""
  
  txtRazon.Text = rs!Razon_Social & ""
  txtWebSite.Text = rs!WebSite & ""
  txtEmail.Text = rs!Email_01 & ""
  txtEmail2.Text = rs!Email_02 & ""
  txtAptoPostal.Text = rs!apto_postal & ""


  dtpFecNac.Value = rs!fecha_nacimiento

  txtNotas = rs!Notas & ""
  
  Call sbCboAsignaDato(cboSexo, IIf((rs!sexo = "M"), "Masculino", "Femenino"), True, rs!sexo)
  Call sbCboAsignaDato(cboEstado, rs!EstadoCivilDesc, True, rs!EstadoCivil)
  Call sbCboAsignaDato(cboClasificacion, rs!CatDesc, True, rs!cod_categoria)
  
  Call sbCboAsignaDato(cboProvincia, rs!ProvDesc & "")  'Se activa el Click ->    Call cboProvincia_Click
  Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")   'Se activa el Click
  Call sbCboAsignaDato(cboDistrito, rs!DistDesc & "")
     
  cboDistrito.ToolTipText = Trim(rs!distrito) & ""
  txtDireccion.Text = rs!direccion


  chkCredito.Value = rs!Credito_Cerrado
  chkExcento.Value = rs!Cliente_Exento
  
  chkRol_Pagador.Value = rs!Rol_Pagador
  chkRol_Autorizador.Value = rs!Rol_Autorizador
  
  
  chkAdelantoAplica.Value = rs!ADELANTO_PERMITE
  chkAdelantoModifica.Value = rs!ADELANTO_MODIFICA
  txtAdelantoPorcentaje.Text = Format(rs!ADELANTO_PORCENTAJE, "Standard")
  chkAdelantoComisionApl.Value = rs!ADELANTO_COMISION_APL
  txtAdelantoComision.Text = Format(rs!ADELANTO_COMISION, "Standard")
  
  chkActivo.Value = rs!Activo
  
  txtCreditoLimite.Text = Format(rs!Credito_Limite, "Standard")
  

  ssTab.Item(0).Selected = True
  ssTab.Item(1).Enabled = True
  ssTab.Item(2).Enabled = True
  ssTab.Item(3).Enabled = True
  ssTab.Item(4).Enabled = True


Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

'Call RefrescaTags(Me)

imgId_Cambio.Enabled = True

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbFichaCliente(pCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

'Procedimiento
'1. Buscar si existe en la tabla CxC_Personas, si existe --> salir
'2.1. Si no existe buscar si existe en Socios, si no existe --> salir
'2.2. Si no existe buscar si existe en CxP_Proveedores, si no existe --> salir
'3. Si existe en (Socios o CxP_Proveedores) cargar los datos encontrados en pantalla


' Punto 1
strSQL = "select isnull(count(*),0) as Existe from cxc_personas where cedula = '" & pCedula & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
  Me.MousePointer = vbDefault
  Exit Sub
End If
rs.Close

'Punto 2.1
strSQL = "select P.*,rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
       & ",Tid.Descripcion as TipoIdDesc,Tid.Tipo_Personeria" _
       & ",dbo.fxAFITelefono(P.cedula,1) as 'TelHab',dbo.fxAFITelefono(P.cedula,2) as 'TelTra', dbo.fxAFITelefono(P.cedula,3) as 'TelCell'" _
       & " from socios P " _
       & " left join Provincias Prov on P.Provincia = Prov.Provincia" _
       & " left join Cantones Cant on P.Provincia = Cant.Provincia and P.Canton = Cant.Canton" _
       & " left join Distritos Dist on P.Provincia = Dist.Provincia and P.Canton = Dist.Canton and P.distrito = Dist.distrito" _
       & " left join AFI_TIPOS_IDS Tid on P.tipo_id = Tid.tipo_id" _
       & " where P.cedula = '" & pCedula & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  vCedula = rs!Cedula
  txtCedula = rs!Cedula

   If Not IsNull(rs!TipoIdDesc) Then
           cboTipoId.Text = Trim(rs!TipoIdDesc)
   End If

  txtNombre = rs!Nombre & ""
  
  txtRazon.Text = rs!Razon_Social & ""
  txtWebSite.Text = ""
  txtEmail.Text = rs!AF_Email & ""
  txtEmail2.Text = ""
  txtAptoPostal.Text = Trim(rs!apto & "")

  txtTelefono.Text = rs!TelHab
  txtTelefono2.Text = rs!TelTra
  txtCelular.Text = rs!TelCell

  cboSexo.Text = IIf((rs!sexo = "M"), "Masculino", "Femenino")
  cboEstado.Text = fxEstadoCivil(rs!EstadoCivil)
  dtpFecNac.Value = rs!fecha_nac

  Call sbCboAsignaDato(cboProvincia, rs!ProvDesc & "")  'Se activa el Click ->    Call cboProvincia_Click
  Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")   'Se activa el Click
  Call sbCboAsignaDato(cboDistrito, rs!DistDesc & "")
     
  cboDistrito.ToolTipText = Trim(rs!distrito) & ""
  txtDireccion.Text = rs!direccion & ""

  rs.Close
  Exit Sub
End If
rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim pLargoCedula As Integer
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Actualiza el Parametro de Validacion y Luego lo Aplica
strSQL = "select LARGO_MINIMO from AFI_TIPOS_IDS Where TIPO_ID = " & cboTipoId.ItemData(cboTipoId.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    pLargoCedula = rs!Largo_Minimo
End If
rs.Close

If Not vEdita Then
    If Len(Trim(txtCedula)) <> pLargoCedula Then vMensaje = vMensaje & " - Número de Identidad no es válido, se espera que sea de: " & pLargoCedula _
            & " caracteres, verifique...!" & vbCrLf
End If

If Len(Trim(txtCedula.Text)) > 20 Then vMensaje = vMensaje & " - Número de Identidad no es válido, verifique...!" & vbCrLf

If Not fxEmail_Valida(txtEmail.Text) Then
    vMensaje = vMensaje & " - El Email principal no es válido!" & vbCrLf
End If

If Len(Trim(txtEmail2.Text)) > 0 Then
    If Not fxEmail_Valida(txtEmail2.Text) Then
        vMensaje = vMensaje & " - El Email secundario no es válido!" & vbCrLf
    End If
End If

If Trim(cboProvincia.Text) = "" Then vMensaje = vMensaje & " - No se especificó la Provincia" & vbCrLf
If Trim(cboCanton.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Cantón" & vbCrLf
If Trim(cboDistrito.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Distrito en la dirección" & vbCrLf

'If Trim(txtDireccion) = "" Then vMensaje = vMensaje & " - No se especificó la Dirección" & vbCrLf
If Not fxDireccion_Valida(Trim(txtDireccion), "-,#,*") Then vMensaje = vMensaje & " - No se especificó la Dirección" & vbCrLf


If Trim(txtNombre.Text) = "" Then vMensaje = vMensaje & " - Indique el Nombre de la Entidad o Persona" & vbCrLf

If txtRazon.Enabled Then
    If Trim(txtRazon.Text) = "" Then vMensaje = vMensaje & " - Indique la Razón Social de la Sociedad" & vbCrLf
Else
    If Trim(cboSexo) = "" Then vMensaje = vMensaje & " - No se especificó el Sexo" & vbCrLf
    If Trim(cboEstado) = "" Then vMensaje = vMensaje & " - No se especificó el Estado Civil" & vbCrLf
End If

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vEstadoCivil As String, vSexo As String

On Error GoTo vError

If cboTipoId.ItemData(cboTipoId.ListIndex) = vTipoJuridica Then
    vEstadoCivil = "O"
    vSexo = "M"
Else
    vEstadoCivil = cboEstado.ItemData(cboEstado.ListIndex)
    vSexo = cboSexo.ItemData(cboSexo.ListIndex)
End If



If vEdita Then
  strSQL = "update CxC_Personas set nombre = '" & Trim(txtNombre.Text) & "',Razon_Social = '" & Trim(txtRazon.Text) & "', Tipo_Id = " & cboTipoId.ItemData(cboTipoId.ListIndex) _
         & ",telefono1 = '" & txtTelefono.Text & "',telefono2 = '" & txtTelefono2.Text & "',celular = '" & txtCelular.Text & "',Fax = '" & txtTelFax.Text & "',WebSite = '" _
         & txtWebSite.Text & "',apto_postal = '" & txtAptoPostal & "',email_01 = '" & txtEmail & "', email_02 = '" & txtEmail2.Text & "',direccion = '" & txtDireccion _
         & "',distrito = '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "',canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
         & "',provincia = " & cboProvincia.ItemData(cboProvincia.ListIndex) _
         & ",sexo = '" & vSexo & "',EstadoCivil = '" & vEstadoCivil & "',Fecha_nacimiento = '" & Format(dtpFecNac.Value, "yyyy/mm/dd") _
         & "',notas = '" & txtNotas & "',credito_cerrado = " & chkCredito.Value & ",Cliente_Exento = " & chkExcento.Value _
         & ",cod_categoria = '" & cboClasificacion.ItemData(cboClasificacion.ListIndex) & "',Categoria_Fecha = dbo.MyGetdate()" _
         & ", ADELANTO_PERMITE = " & chkAdelantoAplica.Value & ", ADELANTO_MODIFICA = " & chkAdelantoModifica.Value _
         & ", ADELANTO_PORCENTAJE = " & CCur(txtAdelantoPorcentaje.Text) & ", CREDITO_LIMITE = " & CCur(txtCreditoLimite.Text) _
         & ", ACTIVO = " & chkActivo.Value & ", ADELANTO_COMISION_APL = " & chkAdelantoComisionApl.Value & ", ADELANTO_COMISION = " & CCur(txtAdelantoComision.Text) _
         & ", ROL_PAGADOR = " & chkRol_Pagador.Value & ", ROL_AUTORIZADOR = " & chkRol_Autorizador.Value _
         & " where cedula = '" & vCedula & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Persona: " & vCedula & ", Estado: " & chkActivo.Value & ", Crd.Cerrado: " & chkCredito.Value & ", Crd. Limite: " & txtCreditoLimite.Text)

Else
  vCedula = txtCedula
  

   strSQL = "insert into CxC_Personas(cedula,Tipo_Id,nombre,razon_social,celular,telefono1,telefono2,fax,sexo,estadoCivil,fecha_nacimiento" _
          & ",apto_postal,email_01,email_02,webSite,notas,direccion,distrito,provincia,canton,credito_cerrado,Cliente_Exento,cod_categoria,categoria_fecha" _
          & ",ADELANTO_PERMITE, ADELANTO_MODIFICA,ADELANTO_PORCENTAJE, CREDITO_LIMITE, ACTIVO, ADELANTO_COMISION_APL" _
          & ", ADELANTO_COMISION, ROL_PAGADOR,ROL_AUTORIZADOR)" _
          & " values('" & vCedula & "'," & cboTipoId.ItemData(cboTipoId.ListIndex) & ",'" & txtNombre.Text & "','" & txtRazon.Text & "','" & txtCelular.Text _
          & "','" & txtTelefono & "','" & txtTelefono2 & "','" & txtTelFax.Text & "','" & vSexo & "','" & vEstadoCivil & "','" _
          & Format(dtpFecNac.Value, "yyyy/mm/dd") & "','" & txtAptoPostal.Text & "','" & txtEmail.Text & "','" & txtEmail2.Text & "','" & txtWebSite.Text _
          & "','" & txtNotas.Text & "','" & txtDireccion.Text & "','" & cboDistrito.ItemData(cboDistrito.ListIndex) & "','" _
          & cboProvincia.ItemData(cboProvincia.ListIndex) & "','" & cboCanton.ItemData(cboCanton.ListIndex) _
          & "'," & chkCredito.Value & "," & chkExcento.Value & ",'" & cboClasificacion.ItemData(cboClasificacion.ListIndex) _
          & "',dbo.MyGetdate()," & chkAdelantoAplica.Value & "," & chkAdelantoModifica.Value & "," & CCur(txtAdelantoPorcentaje.Text) _
          & "," & CCur(txtCreditoLimite.Text) & "," & chkActivo.Value & "," & chkAdelantoComisionApl.Value & "," & CCur(txtAdelantoComision.Text) _
          & "," & chkRol_Pagador.Value & "," & chkRol_Autorizador.Value & ")"
   Call ConectionExecute(strSQL)

  Call Bitacora("Registra", "Persona: " & vCedula & ", Estado: " & chkActivo.Value & ", Crd.Cerrado: " & chkCredito.Value & ", Crd. Limite: " & txtCreditoLimite.Text)

End If

ssTab.Item(1).Enabled = True
ssTab.Item(2).Enabled = True
ssTab.Item(3).Enabled = True

Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete CxC_Personas where cedula = '" & vCedula & "'"
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Persona: " & vCedula)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, i As Long

GLOBALES.gTag = Trim(vCedula)
GLOBALES.gTag2 = txtNombre.Text

Select Case Button.Key
 Case "nuevo"
    Select Case ssTab.Selected.Index
       Case 1 'Operaciones
            Call sbFormsCall("frmCxC_Cuentas", 1, , , False)
       Case 2 'Contratos
            Call sbFormsCall("frmCxC_ClientesContratos", 1, , , False)
       Case 3 'Cuentas Bancarias
            GLOBALES.gTag2 = "CxC"
            'Call sbFormsCall("frmCC_Cuentas_Bancarias", 1, , , False)
            frmCC_Cuentas_Bancarias.Show vbModal
            
            'Call sbFormsCall("frmCxC_ClientesBancos", 1, , , False)
    End Select
 
    'Refresca la información
    ssTab.Item(0).Selected = True
    
 Case "borrar"
    Select Case ssTab.Selected.Index
       Case 2 'Contratos
          
          With lswP.ListItems
          For i = 1 To .Count
            If .Item(i).Checked Then
               strSQL = "delete CxC_Personas_Contratos_Pagadores where cod_contrato = '" & txtCodigo.Text _
                      & "' and cedula = '" & vCedula & "' and cedula_pagador = '" & .Item(i).Text & "'"
               Call ConectionExecute(strSQL)
               
                Call Bitacora("Borra", "Pagador Id.:" & .Item(i).Text & " de Contrato No.:" & txtCodigo.Text & " Ced:" & vCedula)
               
            End If
          Next i
          End With
       
          With lswC.ListItems
          For i = 1 To .Count
            If .Item(i).Checked Then
               strSQL = "delete CxC_Personas_Contratos_Suscripciones where cod_contrato = '" & txtCodigo.Text _
                      & "' and cod_cargo = '" & .Item(i).Text & "' and cedula = '" & vCedula & "'"
               Call ConectionExecute(strSQL)
  
               Call Bitacora("Borra", "Cargo Suscripción Cod:" & .Item(i).Text & " Cnt: " & txtCodigo.Text)
            End If
          Next i
          End With
       
       
       Case 3 'Cuentas Bancarias
          With lswBancos.ListItems
          For i = 1 To .Count
            If .Item(i).Checked Then
               strSQL = "delete SYS_CUENTAS_BANCARIAS where Identificacion = '" & vCedula _
                      & "' and CUENTA_INTERNA = '" & .Item(i).Text _
                      & "' and Modulo = 'CxC'"
               Call ConectionExecute(strSQL)
               Call Bitacora("Elimina", "Cuenta Ahorros: " & .Item(i).SubItems(1) & " Id: " & .Item(i).Text & "_Ced:" & vCedula)

            End If
          Next i
          End With
  
    End Select
    
    Call sbConsultaContratoDetalle


End Select 'Toolbar

End Sub

Private Sub txtCelular_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelFax.SetFocus
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from CxC_Personas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  If txtCedula <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCedula_LostFocus()
  Call sbConsulta(txtCedula.Text)
'  txtNombre.SetFocus
End Sub




Private Sub txtCreditoLimite_GotFocus()
On Error GoTo vError

txtCreditoLimite.Text = CCur(txtCreditoLimite.Text)

vError:
End Sub

Private Sub txtCreditoLimite_LostFocus()
On Error GoTo vError

txtCreditoLimite.Text = Format(CCur(txtCreditoLimite.Text), "Standard")

vError:
End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If cboSexo.Enabled Then
     cboSexo.SetFocus
  Else
     txtNotas.SetFocus
  End If
End If
End Sub

Private Sub txtEmail2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAptoPostal.SetFocus
End Sub



Private Sub txtIdNew_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtIdReferencia.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from CxC_Personas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtIdNew = gBusquedas.Resultado
  txtIdReferencia = gBusquedas.Resultado2
End If

End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoId.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from CxC_Personas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  If txtCedula <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono2.SetFocus
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCelular.SetFocus
End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub

Private Sub txtAptoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboProvincia.SetFocus
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail2.SetFocus
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpFecNac.SetFocus
End Sub

Private Sub txtTelFax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtRazon.Enabled Then
     txtRazon.SetFocus
  Else
     txtWebSite.SetFocus
  End If
End If
End Sub

Private Sub txtWebSite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub
