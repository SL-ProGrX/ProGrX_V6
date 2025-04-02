VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCxPProvCargoPer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Cargos Flotantes a Proveedores"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   9345
   Begin XtremeSuiteControls.TabControl ssTab 
      Height          =   4932
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   9132
      _Version        =   1441792
      _ExtentX        =   16108
      _ExtentY        =   8700
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
      Item(0).Caption =   "Registro"
      Item(0).ControlCount=   18
      Item(0).Control(0)=   "cboSeq"
      Item(0).Control(1)=   "txtDetalle"
      Item(0).Control(2)=   "txtValor"
      Item(0).Control(3)=   "cboTipo"
      Item(0).Control(4)=   "txtConcepto"
      Item(0).Control(5)=   "txtCargoDesc"
      Item(0).Control(6)=   "txtCargoCod"
      Item(0).Control(7)=   "dtpVence"
      Item(0).Control(8)=   "dtpCobroAnticipo"
      Item(0).Control(9)=   "Label3(9)"
      Item(0).Control(10)=   "Label3(6)"
      Item(0).Control(11)=   "Label3(5)"
      Item(0).Control(12)=   "Label3(4)"
      Item(0).Control(13)=   "Label3(3)"
      Item(0).Control(14)=   "Label3(2)"
      Item(0).Control(15)=   "Label3(1)"
      Item(0).Control(16)=   "Label3(0)"
      Item(0).Control(17)=   "GroupBox1"
      Item(1).Caption =   "Cobros"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "ssTabAux"
      Item(2).Caption =   "Informes"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "GroupBox2"
      Item(2).Control(1)=   "GroupBox3"
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   4212
         Left            =   -65920
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   4332
         _Version        =   1441792
         _ExtentX        =   7641
         _ExtentY        =   7429
         _StockProps     =   79
         Caption         =   "Filtros"
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnReporte 
            Height          =   492
            Left            =   2760
            TabIndex        =   33
            Top             =   3600
            Width           =   1572
            _Version        =   1441792
            _ExtentX        =   2773
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Reporte"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCxPProvCargoPer.frx":0000
         End
         Begin VB.CheckBox chkTodas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Todas las Fechas"
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
            Height          =   375
            Left            =   2760
            TabIndex        =   30
            Top             =   3000
            Width           =   1575
         End
         Begin VB.ComboBox cboProveedor 
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
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   26
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   600
            Width           =   4092
         End
         Begin VB.ComboBox cboCargo 
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
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1200
            Width           =   4092
         End
         Begin VB.ComboBox cboEstado 
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
            ItemData        =   "frmCxPProvCargoPer.frx":07BC
            Left            =   240
            List            =   "frmCxPProvCargoPer.frx":07BE
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1800
            Width           =   4092
         End
         Begin XtremeSuiteControls.DateTimePicker dtpInicio 
            Height          =   312
            Left            =   3000
            TabIndex        =   50
            Top             =   2280
            Width           =   1332
            _Version        =   1441792
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
         Begin XtremeSuiteControls.DateTimePicker dtpCorte 
            Height          =   312
            Left            =   3000
            TabIndex        =   51
            Top             =   2640
            Width           =   1332
            _Version        =   1441792
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
         Begin VB.Label Label2 
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
            Height          =   252
            Index           =   2
            Left            =   2400
            TabIndex        =   32
            Top             =   2280
            Width           =   612
         End
         Begin VB.Label Label2 
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
            Height          =   252
            Index           =   3
            Left            =   2400
            TabIndex        =   31
            Top             =   2640
            Width           =   972
         End
         Begin VB.Label Label2 
            Caption         =   "Proveedor"
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
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   972
         End
         Begin VB.Label Label2 
            Caption         =   "Cargo"
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
            Left            =   240
            TabIndex        =   28
            Top             =   960
            Width           =   972
         End
         Begin VB.Label Label2 
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
            Index           =   4
            Left            =   240
            TabIndex        =   27
            Top             =   1560
            Width           =   972
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   3852
         Left            =   -69640
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   3372
         _Version        =   1441792
         _ExtentX        =   5948
         _ExtentY        =   6794
         _StockProps     =   79
         Caption         =   "Reportes"
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
         Appearance      =   16
         BorderStyle     =   1
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "Cargos Registrados"
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
            Left            =   600
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   3135
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "Saldos de Cargos Registrados"
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
            Left            =   600
            TabIndex        =   21
            Top             =   720
            Width           =   3135
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "Pagos de Cargos Periódicos"
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
            Index           =   2
            Left            =   600
            TabIndex        =   20
            Top             =   1080
            Width           =   3135
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "Pagos de Cargos Inmediatos"
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
            Index           =   3
            Left            =   600
            TabIndex        =   19
            Top             =   1440
            Width           =   3135
         End
      End
      Begin XtremeSuiteControls.TabControl ssTabAux 
         Height          =   4452
         Left            =   -70000
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1441792
         _ExtentX        =   16108
         _ExtentY        =   7853
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
         Item(0).Caption =   "Cargos Registrado"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "lsw"
         Item(1).Caption =   "Pagos"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "lswPagos"
         Begin XtremeSuiteControls.ListView lsw 
            Height          =   4092
            Left            =   0
            TabIndex        =   48
            Top             =   360
            Width           =   9132
            _Version        =   1441792
            _ExtentX        =   16108
            _ExtentY        =   7218
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
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.ListView lswPagos 
            Height          =   4092
            Left            =   -70000
            TabIndex        =   49
            Top             =   360
            Visible         =   0   'False
            Width           =   9132
            _Version        =   1441792
            _ExtentX        =   16108
            _ExtentY        =   7218
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
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   16
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1332
         Left            =   480
         TabIndex        =   13
         Top             =   3480
         Width           =   8292
         _Version        =   1441792
         _ExtentX        =   14626
         _ExtentY        =   2350
         _StockProps     =   79
         Caption         =   "Recaudado: "
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtSaldo 
            Height          =   312
            Left            =   5520
            TabIndex        =   45
            Top             =   240
            Width           =   2532
            _Version        =   1441792
            _ExtentX        =   4466
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
         Begin XtremeSuiteControls.FlatEdit txtRecaudado 
            Height          =   312
            Left            =   5520
            TabIndex        =   46
            Top             =   600
            Width           =   2532
            _Version        =   1441792
            _ExtentX        =   4466
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
         Begin XtremeSuiteControls.FlatEdit txtTipoCambioRegistro 
            Height          =   312
            Left            =   6720
            TabIndex        =   47
            Top             =   960
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
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
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Tipo Cambio"
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
            Height          =   312
            Index           =   1
            Left            =   5520
            TabIndex        =   16
            Top             =   960
            Width           =   1212
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Recaudado"
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
            Height          =   312
            Left            =   4440
            TabIndex        =   15
            Top             =   600
            Width           =   1092
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Saldo"
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
            Height          =   312
            Index           =   0
            Left            =   4440
            TabIndex        =   14
            Top             =   240
            Width           =   1092
         End
      End
      Begin XtremeSuiteControls.DateTimePicker dtpVence 
         Height          =   312
         Left            =   7200
         TabIndex        =   34
         Top             =   2520
         Width           =   1332
         _Version        =   1441792
         _ExtentX        =   2350
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.DateTimePicker dtpCobroAnticipo 
         Height          =   312
         Left            =   7200
         TabIndex        =   35
         Top             =   3000
         Width           =   1332
         _Version        =   1441792
         _ExtentX        =   2350
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   1560
         TabIndex        =   36
         Top             =   2520
         Width           =   1452
         _Version        =   1441792
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
      Begin XtremeSuiteControls.ComboBox cboSeq 
         Height          =   312
         Left            =   1560
         TabIndex        =   37
         Top             =   480
         Width           =   6972
         _Version        =   1441792
         _ExtentX        =   12303
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
      Begin XtremeSuiteControls.FlatEdit txtCargoCod 
         Height          =   312
         Left            =   1560
         TabIndex        =   40
         Top             =   840
         Width           =   1092
         _Version        =   1441792
         _ExtentX        =   1926
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCargoDesc 
         Height          =   312
         Left            =   2640
         TabIndex        =   41
         Top             =   840
         Width           =   5892
         _Version        =   1441792
         _ExtentX        =   10393
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
      Begin XtremeSuiteControls.FlatEdit txtConcepto 
         Height          =   312
         Left            =   1560
         TabIndex        =   42
         Top             =   1200
         Width           =   6972
         _Version        =   1441792
         _ExtentX        =   12298
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   912
         Left            =   1560
         TabIndex        =   43
         Top             =   1560
         Width           =   6972
         _Version        =   1441792
         _ExtentX        =   12298
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtValor 
         Height          =   312
         Left            =   4200
         TabIndex        =   44
         Top             =   2520
         Width           =   1692
         _Version        =   1441792
         _ExtentX        =   2984
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "Secuencia"
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
         TabIndex        =   12
         Top             =   480
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "Cargo"
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
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   972
      End
      Begin VB.Label Label3 
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
         Height          =   252
         Index           =   2
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   972
      End
      Begin VB.Label Label3 
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
         Height          =   252
         Index           =   3
         Left            =   480
         TabIndex        =   9
         Top             =   2520
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "Valor"
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
         Left            =   3240
         TabIndex        =   8
         Top             =   2520
         Width           =   612
      End
      Begin VB.Label Label3 
         Caption         =   "Vence"
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
         Left            =   6480
         TabIndex        =   7
         Top             =   2520
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "Detalle"
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
         Index           =   6
         Left            =   480
         TabIndex        =   6
         Top             =   1560
         Width           =   972
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cobro de Anticipo en próximos pagos a partir de:"
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
         Index           =   9
         Left            =   5040
         TabIndex        =   5
         Top             =   2880
         Width           =   2052
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   3
      Top             =   5976
      Width           =   9348
      _ExtentX        =   16484
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Divisa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Tipo de Cambio"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9596
            MinWidth        =   9596
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
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
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8760
      TabIndex        =   2
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1200
      TabIndex        =   38
      Top             =   480
      Width           =   1092
      _Version        =   1441792
      _ExtentX        =   1926
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   2280
      TabIndex        =   39
      Top             =   480
      Width           =   6372
      _Version        =   1441792
      _ExtentX        =   11239
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
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
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   972
   End
End
Attribute VB_Name = "frmCxPProvCargoPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vScroll As Boolean
Dim vDivisa As String, vTipoCambio As Currency, vPaso As Boolean

Private Sub cboProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "S"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado2 <> "" Then cboProveedor.Text = Trim(gBusquedas.Resultado2)
End If
End Sub


Private Sub cboSeq_Click()

If vPaso Then Exit Sub
If cboSeq.ListCount <= 0 Then Exit Sub

Call sbConsulta(cboSeq.ItemData(cboSeq.ListIndex))

End Sub

Private Sub cboTipo_Click()
If Mid(cboTipo.Text, 1, 1) = "P" Then
  dtpVence.Enabled = True
Else
  dtpVence.Enabled = False
End If
End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtValor.SetFocus
End Sub

Private Sub chkTodas_Click()
If chkTodas.Value = vbChecked Then
 dtpInicio.Enabled = False
 dtpCorte.Enabled = False
Else
 dtpInicio.Enabled = True
 dtpCorte.Enabled = True
End If
End Sub

Private Sub cmdReporte_Click()
   MsgBox "Opción en Desarrollo! Consulte los Informes del Módulo", vbExclamation
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 cod_proveedor,descripcion from cxp_proveedores"
           
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_proveedor > " & IIf(txtCodigo = "", 0, txtCodigo) & " order by cod_proveedor asc"
    Else
       strSQL = strSQL & " where cod_proveedor < " & IIf(txtCodigo = "", 0, txtCodigo) & " order by cod_proveedor desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!cod_proveedor
      txtNombre.Text = rs!Descripcion
      Call txtCodigo_LostFocus
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
vModulo = 30
End Sub

Private Sub Form_Load()

vModulo = 30

On Error GoTo vError
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
'Desactiva la Opcion de Reportes hasta que este desarrollada
ssTab.Item(2).Visible = False
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 500
    .Add , , "Cargo", 900, vbCenter
    .Add , , "Descripción", 2500
    .Add , , "Tipo", 900, vbCenter
    .Add , , "Valor", 1500, vbRightJustify
    .Add , , "Saldo", 1500, vbRightJustify
    .Add , , "Recaudado", 1500, vbRightJustify
    .Add , , "Concepto", 2000
    .Add , , "Vence", 2100, vbCenter
    .Add , , "Cobra a partir", 2100, vbCenter
End With

With lswPagos.ColumnHeaders
    .Clear
    .Add , , "No.Pago", 800
    .Add , , "No.Factura", 2500
    .Add , , "Fecha", 1800, vbCenter
    .Add , , "Monto", 1500, vbRightJustify
    .Add , , "Documento", 2000
End With


 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()
Dim strSQL As String, rs As New ADODB.Recordset

vCodigo = 0
txtCodigo = ""
txtNombre = ""

vPaso = True
    cboSeq.Clear
vPaso = False

vDivisa = "COL"
vTipoCambio = 1
txtCodigo.Enabled = True

txtCargoCod = ""
txtCargoDesc = ""
txtConcepto = ""
txtDetalle = ""

cboTipo.Clear
'cboTipo.AddItem "Porcentaje"
cboTipo.AddItem "Monto"
cboTipo.Text = "Monto"
txtValor = 0

dtpVence.Value = fxFechaServidor

dtpCobroAnticipo.Value = dtpVence.Value
dtpCobroAnticipo.MinDate = dtpCobroAnticipo.Value
dtpCobroAnticipo.MaxDate = DateAdd("d", 45, dtpCobroAnticipo.Value)

txtSaldo.Text = "0.00"
txtRecaudado.Text = "0.00"


ssTab.Item(0).Selected = True


End Sub


Private Sub sbCobrosRealizados()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

lsw.ListItems.Clear
lswPagos.ListItems.Clear

strSQL = "select C.*,D.descripcion as CargoDesc" _
       & " from cxp_cargosper C inner join cxp_cargos D on C.cod_cargo = D.cod_cargo" _
       & " where C.cod_proveedor = " & vCodigo & " order by C.ID desc"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Id)
     itmX.SubItems(1) = rs!COD_CARGO
     itmX.SubItems(2) = rs!cargodesc
     itmX.SubItems(3) = rs!Tipo
     itmX.SubItems(4) = Format(rs!Valor, "Standard")
     itmX.SubItems(5) = Format(rs!Saldo, "Standard")
     itmX.SubItems(6) = Format(rs!Recaudado, "Standard")
     itmX.SubItems(7) = rs!CONCEPTO
     itmX.SubItems(8) = Format(rs!Vence, "yyyy/mm/dd")
     itmX.SubItems(9) = Format(rs!Fecha_Cobro_Cargo & "", "yyyy/mm/dd")
 rs.MoveNext
Loop
rs.Close


End Sub

Private Sub sbInicializaReportes()
Dim strSQL As String, rs As New ADODB.Recordset

cboProveedor.Clear
cboEstado.Clear
cboCargo.Clear

strSQL = "select cod_cargo,descripcion from cxp_cargos"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboCargo.AddItem Trim(rs!COD_CARGO) & " - " & Trim(rs!Descripcion)
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
  rs.MoveFirst
  cboCargo.Text = Trim(rs!COD_CARGO) & " - " & Trim(rs!Descripcion)
End If
rs.Close

strSQL = "select cod_proveedor,descripcion from cxp_proveedores"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboProveedor.AddItem Trim(rs!Descripcion)
 cboProveedor.ItemData(cboProveedor.NewIndex) = rs!cod_proveedor
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
  rs.MoveFirst
  cboProveedor.Text = Trim(rs!Descripcion)
End If
rs.Close

cboEstado.AddItem "01 - En Cobro"
cboEstado.AddItem "02 - Cancelados"
cboEstado.AddItem "03 - Todos"
cboEstado.Text = "03 - Todos"

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

chkTodas.Value = vbUnchecked

End Sub



Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curMonto As Currency

If lsw.ListItems.Count = 0 Then Exit Sub
If Item.Text = "" Then Exit Sub

curMonto = 0
lswPagos.ListItems.Clear

ssTabAux.Item(1).Selected = True

strSQL = "select C.*,P.fecha_Traslada,P.tesoreria" _
       & " from cxp_pagoprov P inner join cxp_pagoprovcargos C" _
       & " on P.npago = C.npago and P.cod_proveedor = C.cod_proveedor" _
       & " and P.cod_factura = C.cod_factura and P.tesoreria is not null" _
       & " Where C.id = " & Item.Text & " And C.cod_proveedor = " & vCodigo
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswPagos.ListItems.Add(, , rs!Npago)
     itmX.SubItems(1) = rs!cod_Factura
     itmX.SubItems(2) = Format(rs!fecha_traslada, "yyyy/mm/dd")
     itmX.SubItems(3) = Format(rs!Monto, "Standard")
     itmX.SubItems(4) = IIf(IsNull(rs!tesoreria), 0, rs!tesoreria)
 curMonto = curMonto + rs!Monto
 rs.MoveNext
Loop
rs.Close

Set itmX = lswPagos.ListItems.Add(, , "")
     itmX.SubItems(3) = "____________"
Set itmX = lswPagos.ListItems.Add(, , "TOTAL")
     itmX.SubItems(3) = Format(curMonto, "Standard")



End Sub


Private Sub ssTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 1 'Pagos Realizados
    Call sbCobrosRealizados
  Case 2 'Reportes
    Call sbInicializaReportes
End Select

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
        vEdita = False
        txtCargoCod = ""
        txtCargoDesc = ""
        txtConcepto = ""
        txtDetalle = ""
        
        cboTipo.Clear
        'cboTipo.AddItem "Porcentaje"
        cboTipo.AddItem "Monto"
        cboTipo.Text = "Monto"
        txtValor = 0
        
        dtpVence.Value = fxFechaServidor
        
        txtSaldo.Text = "0.00"
        txtRecaudado.Text = "0.00"
        
        ssTab.Item(0).Selected = True

        txtCodigo.Enabled = False
        Call sbToolBar(tlb, "edicion")
        
        txtTipoCambioRegistro.Text = Format(vTipoCambio, "Standard")
    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.Enabled = False
      txtCargoCod.SetFocus
      Call sbToolBar(tlb, "edicion")
    
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If cboSeq.ListCount <= 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(cboSeq.ItemData(cboSeq.ListIndex))
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtNombre.SetFocus
    
    Case "REPORTES"
       ssTab.Item(2).Selected = True
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(pCargoID As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select C.*, isnull(C.Fecha_Cobro_Cargo, C.registro_Fecha) as 'FechaInicioCobro'" _
       & ",P.descripcion  as Proveedor,D.descripcion as CargoDesc" _
       & " from cxp_proveedores P inner join cxp_cargosper C on P.cod_proveedor = C.cod_proveedor" _
       & " inner join cxp_cargos D on C.cod_cargo = D.cod_cargo" _
       & " where C.ID = " & pCargoID & " and C.cod_proveedor = " & txtCodigo.Text
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!cod_proveedor
  txtCodigo = rs!cod_proveedor
  txtNombre = rs!Proveedor
  
  txtCargoCod = rs!COD_CARGO
  txtCargoDesc = rs!cargodesc
  
  txtConcepto = rs!CONCEPTO
  txtDetalle = rs!Detalle
  
  If UCase(rs!Tipo) = "P" Then
    cboTipo.Text = "Porcentaje"
    txtSaldo.Text = "0.00"
    txtRecaudado.Text = Format(fxRecaudado(rs!Id, rs!cod_proveedor), "Standard")
  Else
    cboTipo.Text = "Monto"
    txtSaldo.Text = Format(rs!Saldo, "Standard")
    txtRecaudado.Text = Format((rs!Valor - rs!Saldo), "Standard")
  End If
  
  txtValor = Format(rs!Valor, "Standard")
  txtTipoCambioRegistro.Text = Format(rs!TIPO_CAMBIO, "Standard")
  
  
  dtpVence.Value = rs!Vence
  dtpCobroAnticipo.MinDate = rs!REGISTRO_FECHA
  dtpCobroAnticipo.Value = rs!FechaInicioCobro
  dtpCobroAnticipo.MaxDate = DateAdd("d", 45, dtpCobroAnticipo.Value)
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxRecaudado(pCargoID As Long, pProveedor As Long) As Currency
Dim strSQL As String, rs As New ADODB.Recordset

fxRecaudado = 0

strSQL = "select isnull(sum(C.monto),0) as Monto" _
       & " from cxp_pagoprov P inner join cxp_pagoprovcargos C" _
       & " on P.npago = C.npago and P.cod_proveedor = C.cod_proveedor" _
       & " and P.cod_factura = C.cod_factura and P.tesoreria is not null" _
       & " Where C.id = " & pCargoID & " And C.cod_proveedor = " & pProveedor
Call OpenRecordSet(rs, strSQL)
    fxRecaudado = rs!Monto
rs.Close

End Function

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

txtConcepto.Text = fxSysCleanTxtInject(txtConcepto.Text)
txtDetalle.Text = fxSysCleanTxtInject(txtDetalle.Text)

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Proveedor no es válido ..."
If txtConcepto = "" Then vMensaje = vMensaje & vbCrLf & " - El concepto no es válido ..."
If Not IsNumeric(txtValor) Then
  vMensaje = vMensaje & vbCrLf & " - El valor no es válido ..."
Else
  If Mid(cboTipo.Text, 1, 1) = "P" And CCur(txtValor) > 100 Then vMensaje = vMensaje & vbCrLf & " - El valor porcentual es superior al maximo ..."
  If CCur(txtValor) <= 0 Then vMensaje = vMensaje & vbCrLf & " - El valor (no marca ningun rango de referencia) no es válido ..."
End If

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCargoID As Long

On Error GoTo vError

If vEdita Then
   
  vCargoID = cboSeq.ItemData(cboSeq.ListIndex)
   
  strSQL = "update cxp_cargosper set detalle = '" & UCase(txtDetalle) & "',concepto = '" _
         & UCase(txtConcepto) & "', Fecha_Cobro_Cargo = '" & Format(dtpCobroAnticipo.Value, "yyyy/mm/dd") _
         & "', Vence = '" & Format(dtpVence.Value, "yyyy/mm/dd") & "'" _
         & " where [id] = " & vCargoID & " and cod_proveedor = " & vCodigo
  Call ConectionExecute(strSQL)

  Call Bitacora("Modifica", "Cargo Adicional a Prov:" & vCodigo & " Sec: " & vCargoID)

Else
   strSQL = "select isnull(max([ID]),0) as ultimo from cxp_cargosper where cod_proveedor = " & vCodigo
   Call OpenRecordSet(rs, strSQL)
     vCargoID = rs!ultimo + 1
   rs.Close
   
   strSQL = "insert cxp_cargosper([id],cod_proveedor,cod_cargo,tipo,valor,vence,saldo,concepto,detalle,recaudado" _
          & ",importe_divisa_real,registro_fecha,registro_usuario,cod_divisa,tipo_cambio, FECHA_COBRO_CARGO)" _
          & " values(" & vCargoID & "," & vCodigo & ",'" & txtCargoCod & "','" & Mid(cboTipo.Text, 1, 1) _
          & "'," & CCur(txtValor) & ",'" & Format(dtpVence.Value, "yyyy/mm/dd") & "',"
   
   If Mid(cboTipo.Text, 1, 1) = "P" Then
     strSQL = strSQL & "0,'" & UCase(txtConcepto) & "','" & UCase(txtDetalle) & "',0,0"
   Else
     strSQL = strSQL & CCur(txtValor) & ",'" & UCase(txtConcepto) & "','" & UCase(txtDetalle) _
            & "',0," & CCur(txtValor.Text) / vTipoCambio
   End If
   
   strSQL = strSQL & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & vDivisa & "'," & vTipoCambio _
          & ",'" & Format(dtpCobroAnticipo.Value, "yyyy/mm/dd") & "')"
   
   Call ConectionExecute(strSQL)
    
    
    'Actualiza Saldo del Proveedor / Solo cuando es por Monto
   If Mid(cboTipo.Text, 1, 1) = "M" Then
        strSQL = "update cxp_proveedores set saldo = isnull(saldo,0) - " & CCur(txtValor) _
               & ",SALDO_DIVISA_REAL = isnull(SALDO_DIVISA_REAL,0) - " & CCur(txtValor.Text) / vTipoCambio _
               & " where cod_proveedor = " & vCodigo
        Call ConectionExecute(strSQL)
   End If
    
   Call Bitacora("Registra", "Cargo Adicional a Prov:" & vCodigo & " Sec: " & vCargoID)

End If
'Activa el codigo
txtCodigo.Enabled = True

'Actualiza todos los datos
Call txtCodigo_LostFocus
Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

If cboSeq.ListCount <= 0 Then Exit Sub
If txtRecaudado.Text > 0 Then
    MsgBox "No se puede eliminar el Cargo porque ya tiene recaudación aplicada...!", vbExclamation
    Exit Sub
End If

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
   'Actualiza Saldo del Proveedor / Solo cuando es por Monto
   If Mid(cboTipo.Text, 1, 1) = "M" Then
        strSQL = "update cxp_proveedores set saldo = isnull(saldo,0) + " & CCur(txtValor) _
               & ",SALDO_DIVISA_REAL = isnull(SALDO_DIVISA_REAL,0) + " & CCur(txtValor.Text) / CCur(txtTipoCambioRegistro.Text) _
               & " where cod_proveedor = " & vCodigo
        Call ConectionExecute(strSQL)
   End If
   
  'Elimina Registro del Cargo
  strSQL = "delete cxp_cargosper where cod_proveedor = " & vCodigo _
         & " and [id] = " & cboSeq.ItemData(cboSeq.ListIndex) & " and recaudado =0 "
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Cargo Adicional a Prov:" & vCodigo & " Sec: " & cboSeq.ItemData(cboSeq.ListIndex) & "..Mnt..:" & txtValor.Text)
  Call txtCodigo_LostFocus
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCargoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCargoDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_cargo"
  gBusquedas.Orden = "cod_cargo"
  gBusquedas.Consulta = "select cod_cargo,descripcion from cxp_cargos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCargoCod = gBusquedas.Resultado
  txtCargoDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCargoCod_LostFocus()
txtCargoDesc = fxSIFCCodigos("D", txtCargoCod, "CargosProv")
End Sub

Private Sub txtCargoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConcepto.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_cargo,descripcion from cxp_cargos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCargoCod = gBusquedas.Resultado
  txtCargoDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "S"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
  Call txtCodigo_LostFocus
End If
End Sub

Private Sub sbCboSeqLlena()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCadena As String, vPrimera As String
On Error GoTo vError

If Not IsNumeric(txtCodigo.Text) Then Exit Sub

vPaso = True
cboSeq.Clear
vPrimera = ""

strSQL = " select Top 200 [ID], registro_fecha, registro_usuario, saldo" _
       & " From cxp_cargosPer  Where COD_PROVEEDOR = " & vCodigo _
       & " order by  [id] desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vCadena = "[" & Format(rs!Id, "####000") & "] Registro: " & rs!REGISTRO_FECHA & " ¦ " & Trim(rs!REGISTRO_USUARIO) & " ¦ Saldo.:" & Format(rs!Saldo, "Standard")
 
  If vPrimera = "" Then vPrimera = vCadena
  cboSeq.AddItem vCadena
'  cboSeq.ItemData(cboSeq.NewIndex) = rs!Id
  cboSeq.ItemData(cboSeq.ListCount - 1) = CStr(rs!Id)
  rs.MoveNext
Loop
rs.Close
vPaso = False

'Activa último Cargo registrado
If vPrimera <> "" Then
   cboSeq.Text = vPrimera
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not IsNumeric(txtCodigo.Text) Then Exit Sub
strSQL = "select Top 1 cod_proveedor,descripcion,cod_divisa,saldo" _
       & ",dbo.fxCntXTipoCambio(" & GLOBALES.gEnlace & ",COD_DIVISA,dbo.MyGetdate(),'V') as 'TipoCambio'" _
       & " from cxp_proveedores where cod_proveedor = " & txtCodigo.Text

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  vDivisa = Trim(rs!COD_DIVISA)
  vTipoCambio = rs!TipoCambio
  txtNombre.Text = rs!Descripcion
  vCodigo = rs!cod_proveedor
  StatusBarX.Panels(1).Text = "Divisa..: " & vDivisa
  StatusBarX.Panels(2).Text = "Tipo Cambio ..: " & Format(vTipoCambio, "Standard")
  StatusBarX.Panels(3).Text = "Saldo Actual..: " & Format(rs!Saldo, "Standard")
  
  Call sbCboSeqLlena
  
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDetalle.SetFocus
End Sub

Private Sub txtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
End Sub

Private Sub txtValor_GotFocus()
On Error GoTo vError
 txtValor = CCur(txtValor)
vError:
End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 If dtpVence.Enabled Then
   dtpVence.SetFocus
 Else
   txtConcepto.SetFocus
 End If
End If
End Sub

Private Sub txtValor_LostFocus()
On Error GoTo vError
 txtValor = Format(CCur(txtValor), "Standard")
vError:
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboSeq.SetFocus
On Error GoTo vError
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  vCodigo = txtCodigo
  txtNombre = gBusquedas.Resultado2
  Call txtCodigo_LostFocus
End If
vError:
End Sub

