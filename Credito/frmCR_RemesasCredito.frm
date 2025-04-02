VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_RemesasCredito 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remesas de Crédito"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   13350
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6732
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   13092
      _Version        =   1441793
      _ExtentX        =   23093
      _ExtentY        =   11874
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
      ItemCount       =   4
      Item(0).Caption =   "General"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "tcAux"
      Item(0).Control(1)=   "vGrid"
      Item(0).Control(2)=   "chkTodos"
      Item(0).Control(3)=   "cmdBuscar"
      Item(0).Control(4)=   "cmdMicrofilm"
      Item(0).Control(5)=   "cmdCrear"
      Item(1).Caption =   "Informes"
      Item(1).ControlCount=   20
      Item(1).Control(0)=   "chkRepTodasFechas"
      Item(1).Control(1)=   "dtpRepCorte"
      Item(1).Control(2)=   "dtpRepInicio"
      Item(1).Control(3)=   "cboRepTags"
      Item(1).Control(4)=   "Label8(10)"
      Item(1).Control(5)=   "Label8(11)"
      Item(1).Control(6)=   "txtRepTagConsec"
      Item(1).Control(7)=   "btnRepBuscar"
      Item(1).Control(8)=   "txtRepRemesas"
      Item(1).Control(9)=   "lswRep"
      Item(1).Control(10)=   "lblRemesa"
      Item(1).Control(11)=   "Label16(2)"
      Item(1).Control(12)=   "fraRecibo"
      Item(1).Control(13)=   "Label16(4)"
      Item(1).Control(14)=   "cmdReporte"
      Item(1).Control(15)=   "chkRemesaInd"
      Item(1).Control(16)=   "opt(3)"
      Item(1).Control(17)=   "opt(2)"
      Item(1).Control(18)=   "opt(1)"
      Item(1).Control(19)=   "opt(0)"
      Item(2).Caption =   "Consultas"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "txtConRemesa"
      Item(2).Control(1)=   "txtOperacion"
      Item(2).Control(2)=   "Label8(16)"
      Item(2).Control(3)=   "Label8(15)"
      Item(3).Caption =   "Carga Listados"
      Item(3).ControlCount=   6
      Item(3).Control(0)=   "txtArchivo"
      Item(3).Control(1)=   "vGridConsulta"
      Item(3).Control(2)=   "Label8(17)"
      Item(3).Control(3)=   "btnArchivoLoad(0)"
      Item(3).Control(4)=   "btnArchivoLoad(1)"
      Item(3).Control(5)=   "btnArchivoLoad(2)"
      Begin XtremeSuiteControls.ListView lswRep 
         Height          =   2652
         Left            =   -69760
         TabIndex        =   51
         Top             =   1440
         Visible         =   0   'False
         Width           =   12852
         _Version        =   1441793
         _ExtentX        =   22669
         _ExtentY        =   4678
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
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   0
         Left            =   -69400
         TabIndex        =   52
         Top             =   4680
         Visible         =   0   'False
         Width           =   3612
         _Version        =   1441793
         _ExtentX        =   6371
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Detalle de Remesa"
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
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   1935
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   13095
         _Version        =   1441793
         _ExtentX        =   23098
         _ExtentY        =   3413
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
         Item(0).Caption =   "Filtros"
         Item(0).ControlCount=   18
         Item(0).Control(0)=   "dtpInicio"
         Item(0).Control(1)=   "dtpCorte"
         Item(0).Control(2)=   "Label8(1)"
         Item(0).Control(3)=   "Label8(0)"
         Item(0).Control(4)=   "Label8(2)"
         Item(0).Control(5)=   "Label8(3)"
         Item(0).Control(6)=   "Label8(4)"
         Item(0).Control(7)=   "Label8(5)"
         Item(0).Control(8)=   "Label8(6)"
         Item(0).Control(9)=   "Label8(7)"
         Item(0).Control(10)=   "cboFuente"
         Item(0).Control(11)=   "cboEOperacion"
         Item(0).Control(12)=   "cboUsuarios"
         Item(0).Control(13)=   "cboDestino"
         Item(0).Control(14)=   "cboGrupos"
         Item(0).Control(15)=   "cboOficina"
         Item(0).Control(16)=   "chkCreditosNoRevisados"
         Item(0).Control(17)=   "txtLinea"
         Item(1).Caption =   "Notas"
         Item(1).ControlCount=   4
         Item(1).Control(0)=   "cboTag"
         Item(1).Control(1)=   "Label8(8)"
         Item(1).Control(2)=   "Label8(9)"
         Item(1).Control(3)=   "txtNotas"
         Item(2).Caption =   "Secuencias/Tags"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "vGridTags"
         Begin XtremeSuiteControls.DateTimePicker dtpInicio 
            Height          =   315
            Left            =   1440
            TabIndex        =   3
            Top             =   480
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
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
         Begin XtremeSuiteControls.DateTimePicker dtpCorte 
            Height          =   315
            Left            =   2760
            TabIndex        =   4
            Top             =   480
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
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
         Begin XtremeSuiteControls.ComboBox cboFuente 
            Height          =   330
            Left            =   1440
            TabIndex        =   5
            Top             =   840
            Width           =   2655
            _Version        =   1441793
            _ExtentX        =   4683
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
         Begin XtremeSuiteControls.ComboBox cboEOperacion 
            Height          =   330
            Left            =   1440
            TabIndex        =   6
            Top             =   1200
            Width           =   2655
            _Version        =   1441793
            _ExtentX        =   4683
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
         Begin XtremeSuiteControls.ComboBox cboUsuarios 
            Height          =   330
            Left            =   1440
            TabIndex        =   7
            Top             =   1560
            Width           =   2655
            _Version        =   1441793
            _ExtentX        =   4683
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
         Begin XtremeSuiteControls.ComboBox cboDestino 
            Height          =   315
            Left            =   5400
            TabIndex        =   8
            Top             =   840
            Width           =   5415
            _Version        =   1441793
            _ExtentX        =   9551
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
         Begin XtremeSuiteControls.ComboBox cboGrupos 
            Height          =   315
            Left            =   5400
            TabIndex        =   9
            Top             =   1200
            Width           =   5415
            _Version        =   1441793
            _ExtentX        =   9551
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
         Begin XtremeSuiteControls.ComboBox cboOficina 
            Height          =   315
            Left            =   5400
            TabIndex        =   10
            Top             =   1560
            Width           =   5415
            _Version        =   1441793
            _ExtentX        =   9551
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
         Begin XtremeSuiteControls.CheckBox chkCreditosNoRevisados 
            Height          =   255
            Left            =   8040
            TabIndex        =   11
            Top             =   480
            Width           =   2775
            _Version        =   1441793
            _ExtentX        =   4890
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cargar Créditos sin Revisión"
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
            Appearance      =   16
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtLinea 
            Height          =   330
            Left            =   5400
            TabIndex        =   12
            Top             =   480
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.ComboBox cboTag 
            Height          =   312
            Left            =   -67720
            TabIndex        =   13
            Top             =   1320
            Visible         =   0   'False
            Width           =   5532
            _Version        =   1441793
            _ExtentX        =   9763
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   792
            Left            =   -67720
            TabIndex        =   14
            Top             =   480
            Visible         =   0   'False
            Width           =   10092
            _Version        =   1441793
            _ExtentX        =   17801
            _ExtentY        =   1397
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
            ScrollBars      =   2
            Appearance      =   2
         End
         Begin FPSpreadADO.fpSpread vGridTags 
            Height          =   1572
            Left            =   -68200
            TabIndex        =   15
            Top             =   360
            Visible         =   0   'False
            Width           =   8892
            _Version        =   524288
            _ExtentX        =   15685
            _ExtentY        =   2773
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   4
            ScrollBars      =   2
            SpreadDesigner  =   "frmCR_RemesasCredito.frx":0000
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   252
            Index           =   1
            Left            =   480
            TabIndex        =   25
            Top             =   480
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Corte:"
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
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   255
            Index           =   0
            Left            =   4560
            TabIndex        =   24
            Top             =   480
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Línea:"
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
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   252
            Index           =   2
            Left            =   480
            TabIndex        =   23
            Top             =   840
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Fuente:"
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
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   252
            Index           =   3
            Left            =   480
            TabIndex        =   22
            Top             =   1200
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Estado:"
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
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   252
            Index           =   4
            Left            =   480
            TabIndex        =   21
            Top             =   1560
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Usuario:"
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
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   255
            Index           =   5
            Left            =   4560
            TabIndex        =   20
            Top             =   1200
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Grupos:"
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
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   255
            Index           =   6
            Left            =   4560
            TabIndex        =   19
            Top             =   840
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Destino:"
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
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   255
            Index           =   7
            Left            =   4560
            TabIndex        =   18
            Top             =   1560
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Oficina:"
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
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   252
            Index           =   8
            Left            =   -68800
            TabIndex        =   17
            Top             =   480
            Visible         =   0   'False
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Notas:"
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
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   252
            Index           =   9
            Left            =   -68800
            TabIndex        =   16
            Top             =   1320
            Visible         =   0   'False
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Etiqueta:"
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
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3852
         Left            =   120
         TabIndex        =   26
         Top             =   3000
         Width           =   12972
         _Version        =   524288
         _ExtentX        =   22881
         _ExtentY        =   6795
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
         MaxCols         =   497
         SpreadDesigner  =   "frmCR_RemesasCredito.frx":064E
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   252
         Left            =   840
         TabIndex        =   27
         Top             =   2640
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   420
         Left            =   9360
         TabIndex        =   28
         Top             =   2520
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Buscar"
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
         Picture         =   "frmCR_RemesasCredito.frx":0E38
      End
      Begin XtremeSuiteControls.PushButton cmdMicrofilm 
         Height          =   420
         Left            =   5160
         TabIndex        =   29
         Top             =   2520
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Recibe Microfilm"
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
         Picture         =   "frmCR_RemesasCredito.frx":1538
      End
      Begin XtremeSuiteControls.PushButton cmdCrear 
         Height          =   420
         Left            =   10560
         TabIndex        =   30
         Top             =   2520
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Crear Remesa!"
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
         Picture         =   "frmCR_RemesasCredito.frx":1B6A
      End
      Begin XtremeSuiteControls.CheckBox chkRepTodasFechas 
         Height          =   255
         Left            =   -65320
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpRepCorte 
         Height          =   315
         Left            =   -66760
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.DateTimePicker dtpRepInicio 
         Height          =   315
         Left            =   -68080
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.ComboBox cboRepTags 
         Height          =   312
         Left            =   -68080
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   4932
         _Version        =   1441793
         _ExtentX        =   8705
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
      Begin XtremeSuiteControls.FlatEdit txtRepTagConsec 
         Height          =   330
         Left            =   -63160
         TabIndex        =   37
         Top             =   840
         Visible         =   0   'False
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
         _ExtentY        =   582
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
         Text            =   "1"
         Alignment       =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.PushButton btnRepBuscar 
         Height          =   420
         Left            =   -62320
         TabIndex        =   38
         Top             =   720
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Buscar"
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
         Picture         =   "frmCR_RemesasCredito.frx":2291
      End
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   420
         Left            =   -58960
         TabIndex        =   42
         Top             =   6240
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Informe"
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
         Picture         =   "frmCR_RemesasCredito.frx":2991
      End
      Begin XtremeSuiteControls.CheckBox chkRemesaInd 
         Height          =   375
         Left            =   -58960
         TabIndex        =   43
         Top             =   5640
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Indicar Remesa"
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
         Appearance      =   16
      End
      Begin FPSpreadADO.fpSpread vGridConsulta 
         Height          =   4932
         Left            =   -69880
         TabIndex        =   44
         Top             =   1560
         Visible         =   0   'False
         Width           =   12972
         _Version        =   524288
         _ExtentX        =   22881
         _ExtentY        =   8700
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
         MaxCols         =   15
         SpreadDesigner  =   "frmCR_RemesasCredito.frx":3098
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   432
         Left            =   -67840
         TabIndex        =   45
         Top             =   720
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   762
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
      Begin XtremeSuiteControls.FlatEdit txtConRemesa 
         Height          =   5112
         Left            =   -67840
         TabIndex        =   46
         Top             =   1320
         Visible         =   0   'False
         Width           =   9852
         _Version        =   1441793
         _ExtentX        =   17378
         _ExtentY        =   9017
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
         Locked          =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRepRemesas 
         Height          =   312
         Left            =   -57760
         TabIndex        =   50
         Top             =   4200
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441793
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
         Text            =   "15"
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   1
         Left            =   -69400
         TabIndex        =   53
         Top             =   5040
         Visible         =   0   'False
         Width           =   3612
         _Version        =   1441793
         _ExtentX        =   6371
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Detalle de Remesa Orden de Revisión"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   2
         Left            =   -69400
         TabIndex        =   54
         Top             =   5400
         Visible         =   0   'False
         Width           =   3612
         _Version        =   1441793
         _ExtentX        =   6371
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Detalle Agrupado de Remesa"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   3
         Left            =   -69400
         TabIndex        =   55
         Top             =   5760
         Visible         =   0   'False
         Width           =   3612
         _Version        =   1441793
         _ExtentX        =   6371
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Remesa de Readecuaciones"
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
      End
      Begin XtremeSuiteControls.FlatEdit txtArchivo 
         Height          =   672
         Left            =   -68200
         TabIndex        =   56
         Top             =   720
         Visible         =   0   'False
         Width           =   7932
         _Version        =   1441793
         _ExtentX        =   13991
         _ExtentY        =   1185
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
         Locked          =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox fraRecibo 
         Height          =   2172
         Left            =   -65800
         TabIndex        =   57
         Top             =   4560
         Visible         =   0   'False
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9334
         _ExtentY        =   3831
         _StockProps     =   79
         Caption         =   "Marcar como Recibido en Archivo Digital?"
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
         Begin XtremeSuiteControls.FlatEdit txtReciboRemesa 
            Height          =   312
            Left            =   2520
            TabIndex        =   58
            Top             =   480
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtReciboUsuario 
            Height          =   312
            Left            =   2520
            TabIndex        =   59
            Top             =   840
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtReciboFecha 
            Height          =   312
            Left            =   2520
            TabIndex        =   60
            Top             =   1200
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   2
         End
         Begin XtremeSuiteControls.PushButton btnArchivoDigital 
            Height          =   420
            Index           =   0
            Left            =   2280
            TabIndex        =   61
            Top             =   1680
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            Picture         =   "frmCR_RemesasCredito.frx":3A6B
         End
         Begin XtremeSuiteControls.PushButton btnArchivoDigital 
            Height          =   420
            Index           =   1
            Left            =   3600
            TabIndex        =   62
            Top             =   1680
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "Cerrar"
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
            Picture         =   "frmCR_RemesasCredito.frx":4192
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   252
            Index           =   12
            Left            =   360
            TabIndex        =   65
            Top             =   480
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Id. Remesa de Crédito:"
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
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   252
            Index           =   13
            Left            =   360
            TabIndex        =   64
            Top             =   840
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Recibido por: (Usuario):"
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
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   252
            Index           =   14
            Left            =   360
            TabIndex        =   63
            Top             =   1200
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Fecha de recibido:"
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
         End
      End
      Begin XtremeSuiteControls.PushButton btnArchivoLoad 
         Height          =   372
         Index           =   0
         Left            =   -60040
         TabIndex        =   66
         Top             =   720
         Visible         =   0   'False
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_RemesasCredito.frx":47D0
      End
      Begin XtremeSuiteControls.PushButton btnArchivoLoad 
         Height          =   372
         Index           =   1
         Left            =   -59560
         TabIndex        =   67
         Top             =   720
         Visible         =   0   'False
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_RemesasCredito.frx":4ED0
      End
      Begin XtremeSuiteControls.PushButton btnArchivoLoad 
         Height          =   372
         Index           =   2
         Left            =   -59080
         TabIndex        =   68
         Top             =   720
         Visible         =   0   'False
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_RemesasCredito.frx":55E9
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   17
         Left            =   -69280
         TabIndex        =   49
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Archivo:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   15
         Left            =   -69160
         TabIndex        =   48
         Top             =   720
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "No. Operación:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   16
         Left            =   -69160
         TabIndex        =   47
         Top             =   1200
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Remesa:"
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
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Remesas - visualizar últimas"
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
         Height          =   300
         Index           =   4
         Left            =   -60640
         TabIndex        =   41
         Top             =   4200
         Visible         =   0   'False
         Width           =   2892
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Seleccione la Remesa que Desea Visualizar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   2
         Left            =   -69760
         TabIndex        =   40
         Top             =   1200
         Visible         =   0   'False
         Width           =   12852
      End
      Begin VB.Label lblRemesa 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -69760
         TabIndex        =   39
         Top             =   4200
         Visible         =   0   'False
         Width           =   9132
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   252
         Index           =   11
         Left            =   -69640
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Corte:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   252
         Index           =   10
         Left            =   -69640
         TabIndex        =   35
         Top             =   840
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Etiqueta:"
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
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   20
      Left            =   12720
      Top             =   120
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Remesas de revisión y archivo de Créditos"
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
      Height          =   492
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   7572
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   13452
   End
End
Attribute VB_Name = "frmCR_RemesasCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim mReqTagRevision As Boolean, mTagRevision As String


Private Sub sbLlenaCbo(cboX As ComboBox, strSQL As String, Optional vTodos As Boolean = True)
Dim rs As New ADODB.Recordset

cboX.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboX.AddItem rs!itmX
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboX.Text = rs!itmX
End If
rs.Close

If vTodos Then
    cboX.AddItem "TODOS"
    cboX.Text = "TODOS"
End If

End Sub

Private Sub btnArchivoDigital_Click(Index As Integer)
Dim strSQL As String

Select Case Index
  Case 0 'aplicar
  
     If cmdMicrofilm.Enabled Then
        If Len(Trim(txtReciboFecha.Text)) = 0 Then
            strSQL = "update crd_remesas set Microfilm_Fecha = dbo.MyGetdate(), Microfilm_usuario = '" & glogon.Usuario _
                    & "' where remesa = " & txtReciboRemesa.Text
            Call ConectionExecute(strSQL)
            
            MsgBox "Recibo (Microfilm ) Satisfactoriamente...!", vbInformation
            Call sbCargaRemesas
        Else
            MsgBox "La remesa ya fue recibida en Microfilm o no existe...verifique!", vbExclamation
        End If
     Else
        MsgBox "No tiene los permisos para realizar esta opción, verifique...!!!", vbExclamation
     End If
  Case 1 'cancelar
    'Nada
End Select

    fraRecibo.Visible = False
    vPaso = False

End Sub

Private Sub btnArchivoLoad_Click(Index As Integer)
Select Case Index
  Case 0 'buscar
        
        txtArchivo.Text = ""
        
        With frmContenedor.CD
                
                .InitDir = "C:\"
                .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]"
                .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
                .ShowOpen
        
                If .FileName = "" Then
                    MsgBox "Archivo no válido...", vbExclamation
                    Exit Sub
                End If
        
                If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
                    'Ok
                Else
                    MsgBox "La Extensión del Archivo no es válido...", vbExclamation
                    Exit Sub
                End If
        
                
         txtArchivo.Text = .FileName
        
        End With


  Case 1 'cargar
    Call sbCarga_Listado
  
  
  Case 2 'Info
        MsgBox "Archivo de Carga: Microsoft Excel" & vbCrLf _
              & " - Columnas: OPERACION" & vbCrLf _
              & " - Nombre de la Hoja: IMPORT" _
          , vbInformation, "Información del Archivo de Carga"
End Select


End Sub

Private Sub btnRepBuscar_Click()
 Call sbCargaRemesas
End Sub

Private Sub cboGrupos_Click()
    Call sbCargarCboUsuario
End Sub

Private Sub cboRepTags_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

If Mid(cboRepTags.Text, 1, 2) = "TO" Then
   txtRepTagConsec.Text = 0
Else
   strSQL = "select isnull(consecutivo,0) as ConsecX from crd_remesas_tags where tag_codigo = '" & cboRepTags.ItemData(cboRepTags.ListIndex) & "'"
   Call OpenRecordSet(rs, strSQL)
   txtRepTagConsec.Text = rs!ConsecX
   rs.Close
End If

End Sub

Private Sub chkRepTodasFechas_Click()
 If chkRepTodasFechas.Value = vbChecked Then
   dtpRepInicio.Enabled = False
 Else
   dtpRepInicio.Enabled = True
 End If
 
 dtpRepCorte.Enabled = dtpRepInicio.Enabled
 
End Sub

Private Sub chkTodos_Click()
Dim i As Integer

For i = 1 To vGrid.MaxRows
 vGrid.Row = i
 vGrid.col = 1
 vGrid.Value = chkTodos.Value
Next i

End Sub

Private Sub sbBuscaCaso01(Optional vTraspaso As Boolean = False)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select V.id_solicitud,V.codigo,V.garantiax,V.fechaforp,R.Fecha_Registro,V.montoapr,V.cedula,V.nombre,V.estado,V.userfor,V.destinoX,R.observacion" _
       & " from vCRDCreditosReportes01 V inner join Reg_Creditos R on V.id_solicitud = R.id_solicitud" _
       & " left join CRD_REMESA_ASG Asg on V.id_solicitud = Asg.ID_SOLICITUD and Asg.REFERENCIA = 0"

If mReqTagRevision Then

    If Not chkCreditosNoRevisados.Value = vbChecked Then
        strSQL = strSQL & " left join CRD_OPERACION_TAGS T on V.ID_SOLICITUD = T.ID_SOLICITUD and T.TAG_CODIGO = '" & mTagRevision & "'" _
                & " inner join dbo.vCRDOperacionTagsMax OT on T.ID_SOLICITUD = OT.ID_SOLICITUD and T.LINEA = OT.LINEA"
'    Else
'        strSQL = strSQL & " inner join CRD_OPERACION_TAGS T on V.ID_SOLICITUD = T.ID_SOLICITUD and T.TAG_CODIGO = '" & mTagRevision & "'" _
'                & " inner join dbo.vCRDOperacionTagsMax OT on T.ID_SOLICITUD = OT.ID_SOLICITUD and T.LINEA = OT.LINEA"
    End If
    
End If

strSQL = strSQL & " where V.fechaforp between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") _
       & " 23:59:59' and Asg.ID_SOLICITUD is null" _

'Eliminado
'V.id_solicitud not in(select id_solicitud from CRD_REMESA_ASG where referencia = 0)"
       
If mReqTagRevision Then
    If Not chkCreditosNoRevisados.Value = vbChecked Then
        strSQL = strSQL & " and R.ANALISTAS_REVISION = 1 "
    End If
End If

Select Case cboEOperacion.Text
  Case "Activas"
     strSQL = strSQL & " and V.estado = 'A'"
  Case "Canceladas"
     strSQL = strSQL & " and V.estado = 'C'"
  Case "Activas y Canceladas"
     strSQL = strSQL & " and V.estado in('A','C')"
  Case "Nulas"
     strSQL = strSQL & " and V.estado = 'N'"
  Case "Todas"
End Select

If Mid(cboGrupos.Text, 1, 1) <> "T" Then
 strSQL = strSQL & " and V.cod_grupo = '" & cboGrupos.ItemData(cboGrupos.ListIndex) & "'"
End If

If txtLinea.Text <> "" Then
  strSQL = strSQL & " and V.codigo = '" & txtLinea.Text & "'"
End If

If Mid(cboOficina.Text, 1, 1) <> "T" Then
 strSQL = strSQL & " and V.cod_oficina_f = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
End If

If Mid(cboDestino.Text, 1, 1) <> "T" Then
 strSQL = strSQL & " and V.cod_Destino = '" & cboDestino.ItemData(cboDestino.ListIndex) & "'"
End If

If txtLinea.Text <> "" Then
  strSQL = strSQL & " and V.codigo = '" & txtLinea.Text & "'"
End If

If Mid(cboUsuarios.Text, 1, 5) <> "TODOS" Then
 strSQL = strSQL & " and T.REGISTRO_USUARIO = '" & Trim(cboUsuarios) & "'"
End If

If vTraspaso Then
  strSQL = strSQL & " and V.referencia is not null"
Else
  strSQL = strSQL & " and V.referencia is null"
End If

If mReqTagRevision Then
    If Not chkCreditosNoRevisados.Value = vbChecked Then
        strSQL = strSQL & " order by T.REGISTRO_USUARIO,T.REGISTRO_FECHA, V.codigo,V.nombre asc"
    Else
        strSQL = strSQL & " order by V.codigo,V.nombre asc"
    End If
Else
    strSQL = strSQL & " order by V.codigo,V.nombre asc"
End If


Call OpenRecordSet(rs, strSQL)

vGrid.MaxRows = 0
Do While Not rs.EOF
   vGrid.MaxRows = vGrid.MaxRows + 1
   vGrid.Row = vGrid.MaxRows
   
   vGrid.col = 1
   vGrid.Value = chkTodos.Value
   
   vGrid.col = 2
   vGrid.Text = CStr(rs!Id_Solicitud)
   vGrid.CellTag = 0
   
   vGrid.col = 3
   vGrid.Text = rs!Codigo
   
   vGrid.col = 4
   vGrid.Text = rs!GarantiaX
   
   
   vGrid.col = 5
   vGrid.Text = IIf(IsNull(rs!FECHA_REGISTRO), Format(rs!FechaForp, "dd/mm/yyyy"), Format(rs!FECHA_REGISTRO, "dd/mm/yyyy hh:mm:ss"))
   vGrid.col = 6
   vGrid.Text = Format(rs!montoapr, "Standard")
   vGrid.col = 7
   vGrid.Text = rs!Cedula
   vGrid.col = 8
   vGrid.Text = rs!Nombre
   vGrid.col = 9
   Select Case rs!Estado
     Case "A"
       vGrid.Text = "Activo"
     Case "C"
       vGrid.Text = "Activo"
     Case Else
       vGrid.Text = "Nulo"
   End Select
   vGrid.col = 10
   vGrid.Text = Trim(rs!Userfor)
   vGrid.col = 11
   vGrid.Text = Trim(rs!DestinoX & "")
   
   vGrid.col = 12
   vGrid.Text = Trim(rs!observacion & "")
   
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbBuscaCaso02()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select C.id_solicitud,C.codigo,T.descripcion as GarantiaX,C.fecha,R.montoapr,R.cedula,S.nombre,R.estado,C.usuario" _
       & ",C.id_credito_suBit,C.Detalle,D.descripcion as DestinoX" _
       & " from credito_suBit C inner join reg_Creditos R on C.id_solicitud = R.id_solicitud" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " inner join Crd_Garantia_Tipos T on R.garantia = T.garantia" _
       & " left join Catalogo_Destinos D on R.cod_destino = D.cod_destino" _
       & " where C.tipo = 'C' and C.Movimiento = '01' and C.id_credito_suBit not in(select referencia" _
       & " from CRD_REMESA_ASG) and C.fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

Select Case cboEOperacion.Text
  Case "Activas"
     strSQL = strSQL & " and R.estado = 'A'"
  Case "Canceladas"
     strSQL = strSQL & " and R.estado = 'C'"
  Case "Activas y Canceladas"
     strSQL = strSQL & " and R.estado in('A','C')"
  Case "Nulas"
     strSQL = strSQL & " and R.estado = 'N'"
  Case "Todas"
End Select


'If Mid(cboGrupos.Text, 1, 1) <> "T" Then
' strSQL = strSQL & " and cod_grupo = '" & fxCodigoCbo(cboGrupos) & "'"
'End If

If txtLinea.Text <> "" Then
  strSQL = strSQL & " and R.codigo = '" & txtLinea.Text & "'"
End If



strSQL = strSQL & " order by R.codigo,S.nombre asc"

vGrid.MaxRows = 0
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
   vGrid.MaxRows = vGrid.MaxRows + 1
   vGrid.Row = vGrid.MaxRows
   
   vGrid.col = 1
   vGrid.Value = chkTodos.Value
   
   vGrid.col = 2
   vGrid.Text = CStr(rs!Id_Solicitud)
   vGrid.CellTag = rs!id_credito_suBit
   
   vGrid.col = 3
   vGrid.Text = rs!Codigo
   
   vGrid.col = 4
   vGrid.Text = rs!GarantiaX
   
   vGrid.col = 5
   vGrid.Text = Format(rs!fecha, "dd/mm/yyyy")
   vGrid.col = 6
   vGrid.Text = Format(rs!montoapr, "Standard")
   vGrid.col = 7
   vGrid.Text = rs!Cedula
   vGrid.col = 8
   vGrid.Text = rs!Nombre
   vGrid.col = 9
   vGrid.Text = rs!Estado
   vGrid.col = 10
   vGrid.Text = Trim(rs!Usuario)
   
   vGrid.col = 11
   vGrid.Text = rs!DestinoX & ""
   
   vGrid.col = 12
   vGrid.Text = rs!Detalle
 
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbBuscaCaso03()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select R.id_solicitud,R.codigo,T.descripcion as GarantiaX,R.fechaforp,R.cuota * R.plazo as MontoApr,R.cedula,S.nombre,R.estado,R.userfor" _
       & " From Reg_creditos R inner join Catalogo C on R.codigo = C.codigo and C.retencion = 'S'" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " inner join Crd_Garantia_Tipos T on R.garantia = T.garantia" _
       & " where R.fechaforp between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
       & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59' and id_solicitud not in(select id_solicitud from CRD_REMESA_ASG where referencia = 0)"



Select Case cboEOperacion.Text
  Case "Activas"
     strSQL = strSQL & " and R.estado = 'A'"
  Case "Canceladas"
     strSQL = strSQL & " and R.estado = 'C'"
  Case "Activas y Canceladas"
     strSQL = strSQL & " and R.estado in('A','C')"
  Case "Nulas"
     strSQL = strSQL & " and R.estado = 'N'"
  Case "Todas"
End Select

'No aplica para Retenciones
'If Mid(cboGrupos.Text, 1, 1) <> "T" Then
' strSQL = strSQL & " and cod_grupo = '" & fxCodigoCbo(cboGrupos) & "'"
'End If

If txtLinea.Text <> "" Then
  strSQL = strSQL & " and R.codigo = '" & txtLinea.Text & "'"
End If

strSQL = strSQL & " order by R.codigo,S.nombre asc"

vGrid.MaxRows = 0
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
   vGrid.MaxRows = vGrid.MaxRows + 1
   vGrid.Row = vGrid.MaxRows
   
   vGrid.col = 1
   vGrid.Value = chkTodos.Value
   
   vGrid.col = 2
   vGrid.Text = CStr(rs!Id_Solicitud)
   vGrid.CellTag = 0
   
   vGrid.col = 3
   vGrid.Text = rs!Codigo
   
   vGrid.col = 4
   vGrid.Text = rs!GarantiaX
   
   vGrid.col = 5
   vGrid.Text = Format(rs!FechaForp, "dd/mm/yyyy")
   vGrid.col = 6
   vGrid.Text = Format(rs!montoapr, "Standard")
   vGrid.col = 7
   vGrid.Text = rs!Cedula
   vGrid.col = 8
   vGrid.Text = rs!Nombre
   vGrid.col = 9
   vGrid.Text = rs!Estado
   vGrid.col = 10
   vGrid.Text = Trim(rs!Userfor)
   
   vGrid.col = 11
   vGrid.Text = ""
   
   vGrid.col = 12
   vGrid.Text = ""

 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cmdBuscar_Click()

Select Case cboFuente.ItemData(cboFuente.ListIndex)
 Case 1
   Call sbBuscaCaso01(False)
 Case 2
   Call sbBuscaCaso02
 Case 3
   Call sbBuscaCaso01(True)
 Case 4
   Call sbBuscaCaso03
End Select

End Sub

Private Sub cmdCrear_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vRemesa As Long, vTag As String, vTagConsec As Long, i As Integer
Dim vLinea As String

On Error GoTo vError

Me.MousePointer = vbHourglass
'Saca la Ultima Remesa
strSQL = "select isnull(max(remesa),0) + 1 as Remesa from crd_remesas"
Call OpenRecordSet(rs, strSQL)
 vRemesa = rs!remesa
rs.Close


 If cboTag.Text <> "TODOS" Then
    
    vTag = SIFGlobal.fxCodText(cboTag.Text)
    
    'Saca Consecutivo del Tag
    strSQL = "select isnull(consecutivo,0) + 1 as Consec from crd_remesas_tags where tag_Codigo = '" & vTag & "'"
    Call OpenRecordSet(rs, strSQL)
     vTagConsec = rs!consec
    rs.Close
    strSQL = "update crd_remesas_tags set consecutivo = " & vTagConsec & " where tag_Codigo = '" & vTag & "'"
    Call ConectionExecute(strSQL)
    
    strSQL = "insert CRD_REMESAS(remesa,fecha,usuario,notas,Tag_Codigo,Tag_Consecutivo) values(" & vRemesa & ",dbo.MyGetdate(),'" _
           & glogon.Usuario & "','" & txtNotas.Text & "','" & vTag & "'," & vTagConsec & ")"
 Else
    strSQL = "insert CRD_REMESAS(remesa,fecha,usuario,notas) values(" & vRemesa & ",dbo.MyGetdate(),'" _
           & glogon.Usuario & "','" & txtNotas.Text & "')"
 End If
 Call ConectionExecute(strSQL)
  
 Dim Linea As Integer
 Linea = 1

'Asigna Operacion a la Remesa Nueva
For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.col = 1
  If vGrid.Value = vbChecked Then
     vGrid.col = 3
     vLinea = vGrid.Text
     vGrid.col = 2
     
     strSQL = "insert CRD_REMESA_ASG(remesa,id_solicitud,referencia,linea) values(" & vRemesa _
            & "," & vGrid.Text & "," & vGrid.CellTag & "," & Linea & ")"
     Call ConectionExecute(strSQL)
     
     Linea = Linea + 1
  
     'Tags de Seguimiento
     Call sbCrdOperacionTags(vGrid.Text, vLinea, "S05", "", "Remesa de Crédito No..:" & vRemesa)
  End If
Next i

Me.MousePointer = vbDefault
MsgBox "Remesa Creada Satisfactoriamente : Remesa(" & vRemesa & ")", vbInformation


Call TimerX_Timer

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaRemesas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pWhere As Boolean
Dim itmX  As ListViewItem

On Error GoTo vError

lswRep.ListItems.Clear
pWhere = False

strSQL = "select top " & txtRepRemesas & " * from crd_remesas"

If cboRepTags.Text <> "TODOS" Then
  pWhere = True
  strSQL = strSQL & " where tag_codigo = '" & cboRepTags.ItemData(cboRepTags.ListIndex) & "'"
End If

If IsNumeric(txtRepTagConsec.Text) Then
  If CLng(txtRepTagConsec.Text) > 0 Then
    If pWhere Then
       strSQL = strSQL & " and tag_consecutivo = " & txtRepTagConsec.Text
    Else
       strSQL = strSQL & " Where tag_consecutivo = " & txtRepTagConsec.Text
       pWhere = True
    End If
  End If
End If


If chkRepTodasFechas.Value = vbUnchecked Then
  If pWhere Then
     strSQL = strSQL & " and fecha between '" & Format(dtpRepInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" & Format(dtpRepCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  Else
     strSQL = strSQL & " Where fecha between '" & Format(dtpRepInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" & Format(dtpRepCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
     pWhere = True
  End If
End If

strSQL = strSQL & " order by fecha desc"


Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswRep.ListItems.Add(, , rs!remesa)
     itmX.SubItems(1) = rs!fecha
     itmX.SubItems(2) = rs!Usuario
     itmX.SubItems(3) = rs!Notas & ""
     itmX.SubItems(4) = rs!Microfilm_fecha & ""
     itmX.SubItems(5) = rs!microfilm_usuario & ""
     itmX.SubItems(6) = rs!TAG_CODIGO & ""
     itmX.SubItems(7) = rs!tag_consecutivo & ""
 rs.MoveNext
Loop
rs.Close

lblRemesa.Caption = ""
lblRemesa.Tag = ""

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cmdReporte_Click()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String
Dim strSQL As String, vTipoUser As String, xRemesa As String

On Error GoTo vError


If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Crédito")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If


If lblRemesa.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass


vSubTitulo = ""
vFiltro = ""
strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Créditos"

 .Connect = glogon.ConectRPT

 Select Case True
  Case opt.Item(0).Value 'Detalle Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Credito_RemesasDetalle.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO"
  Case opt.Item(1).Value 'Detalle Agrupado Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Credito_RemesasDetalleAgrupado.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO AGRUPADO"
     
  Case opt.Item(2).Value 'Readeuaciones
     .ReportFileName = SIFGlobal.fxPathReportes("Credito_RemesasReadecuaciones.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : READECUACIONES"
     
  Case opt.Item(3).Value 'Detalle Remesa orden revisión
     .ReportFileName = SIFGlobal.fxPathReportes("Credito_RemesasDetalleOrdenRevision.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO ORDEN REVISIÓN"
     
 End Select
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA DE CREDITOS'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(5) = "fxFiltro='" & vFiltro & "'"
 
 If Not opt.Item(3).Value = True Then
    .SelectionFormula = "{CRD_REMESAS.REMESA} = " & lblRemesa.Tag
 Else
    strSQL = "{CRD_REMESAS.REMESA} = " & lblRemesa.Tag
    strSQL = strSQL & " and {CRD_OPERACION_TAGS.TAG_CODIGO} = 'S10'"
    .SelectionFormula = strSQL
 End If
 
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

 tcMain.Item(0).Selected = True
 tcAux.Item(0).Selected = True
 
 With lswRep.ColumnHeaders
    .Clear
    .Add , , "Remesa Id", 1400
    .Add , , "Fecha", 2100
    .Add , , "Usuario", 1800
    .Add , , "Notas", 3400
    
    .Add , , "Archivo: Entregado", 2400
    .Add , , "Archivo: Recibe", 2400
    .Add , , "Etiqueta", 2400
    .Add , , "Tag. Id:", 1400
 End With


Call Formularios(Me)
Call RefrescaTags(Me)

Call sbParametrosTagRevision

cmdBuscar.Visible = True
cmdCrear.Visible = True
cmdReporte.Visible = False

chkRemesaInd.Visible = False
txtRepRemesas.Visible = False

cboTag.AddItem "TODOS"
cboTag.Text = "TODOS"

End Sub

Private Sub Form_Resize()

'On Error Resume Next
'
'tcMain.Width = Me.Width - 280
'tcMain.Height = Me.Height - 700
'
'vGrid.Width = tcMain.Width - 440
'vGrid.Height = tcMain.Height - 3530
'
'vGridConsulta.Width = tcMain.Width - 440
'vGridConsulta.Height = tcMain.Height - (txtArchivo.Top + txtArchivo.Height + 150)
'
'
'Label16.Item(1).Width = vGrid.Width
'Label16.Item(2).Width = vGrid.Width
'lblRemesa.Width = vGrid.Width
'
'Line1.Item(1).X2 = vGrid.Width
'lswRep.Width = vGrid.Width
'
'cmdReporte.Top = tcMain.Height - 495
'cmdReporte.Left = tcMain.Width - 1695
'
'Line1.Item(1).Y1 = cmdReporte.Top - 120
'Line1.Item(1).Y2 = Line1.Item(1).Y1
'
'chkRemesaInd.Top = Line1.Item(1).Y1 - 280
'chkRemesaInd.Left = cmdReporte.Left
'
'
'lswRep.Height = ssTab.Height - 4580
'
'opt.Item(0).Top = (lswRep.Height + lswRep.Top + 440) '2460
'opt.Item(3).Top = opt.Item(0).Top + 360
'opt.Item(1).Top = opt.Item(3).Top + 360
'opt.Item(2).Top = opt.Item(1).Top + 360
'
'lblRemesa.Top = (lswRep.Height + lswRep.Top + 80)
'Label16(4).Top = lblRemesa.Top
'txtRepRemesas.Top = lblRemesa.Top
'
'txtRepRemesas.Left = lswRep.Width - 340
'Label16(4).Left = lswRep.Width - 2900
'
'txtConRemesa.Width = ssTab.Width - 2260
'txtConRemesa.Height = ssTab.Height - 1350
'
End Sub



Private Sub lswRep_DblClick()
If lswRep.ListItems.Count > 0 Then
        
  
   vPaso = True
   
   fraRecibo.Visible = True
   
   txtReciboRemesa.Text = lswRep.SelectedItem
   txtReciboUsuario.Text = lswRep.SelectedItem.SubItems(5)
   txtReciboFecha.Text = lswRep.SelectedItem.SubItems(4)
   

End If

End Sub

Private Sub lswRep_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

lblRemesa.Caption = Item.Text & " ¦ " & Item.SubItems(1) _
            & " ¦ " & Item.SubItems(2)
lblRemesa.Tag = Item.Text


End Sub

Private Sub tcAux_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

Select Case Item.Index
  Case 0 'Filtros
  Case 1 'Notas y Tags
        strSQL = "select Tag_Codigo as 'IdX', rtrim(descripcion) as 'ItmX'" _
                 & " from  crd_remesas_tags where activo = 1"
        Call sbCbo_Llena_New(cboTag, strSQL, True, True)
  Case 2 'Tags
        strSQL = "select tag_codigo,descripcion,activo,consecutivo" _
               & " from crd_remesas_tags order by Tag_Codigo"
        Call sbCargaGrid(vGridTags, 4, strSQL)
End Select
End Sub

'Private Sub lswRep_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'If vPaso Then Exit Sub
'
'On Error Resume Next
' fraRecibo.Left = x
' fraRecibo.Top = y
'
'End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 0 Then
 TimerX_Timer
Else
 Call sbCargaRemesas
End If

Select Case Item.Index
 Case 0 'Remesas
   
   tcAux.Item(0).Selected = True
   
   
 Case 1 'Reportes
   vPaso = False
 Case 2 'Consultas
 Case 3 'Listados
End Select

End Sub

Private Sub TimerX_Timer()
Dim strSQL As String

TimerX.Interval = 0

tcMain.Item(0).Selected = True

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

cboEOperacion.Clear
cboEOperacion.AddItem "Activas"
cboEOperacion.AddItem "Canceladas"
cboEOperacion.AddItem "Nulas"
cboEOperacion.AddItem "Activas y Canceladas"
cboEOperacion.AddItem "Todas"
cboEOperacion.Text = "Activas"

strSQL = "select cod_grupo as 'IdX',rtrim(descripcion) as 'ItmX'" _
         & " from  crd_grupos"
Call sbCbo_Llena_New(cboGrupos, strSQL, True, True)

Call sbCargarCboUsuario

strSQL = "select cod_destino as 'IdX',rtrim(descripcion) as 'ItmX'" _
       & " from  catalogo_destinos"
Call sbCbo_Llena_New(cboDestino, strSQL, True, True)

strSQL = "select rtrim(cod_oficina) as 'IdX',rtrim(descripcion) as 'ItmX'" _
       & " from SIF_Oficinas order by cod_oficina"
Call sbCbo_Llena_New(cboOficina, strSQL, True, True)

dtpRepInicio.Value = dtpInicio.Value
dtpRepCorte.Value = dtpRepInicio.Value

vPaso = True
strSQL = "select rtrim(Tag_codigo) as 'IdX',rtrim(descripcion) as 'ItmX'" _
       & " from Crd_Remesas_Tags order by Tag_codigo"
Call sbCbo_Llena_New(cboRepTags, strSQL, True, True)
vPaso = False

txtNotas.Text = ""

cboFuente.Clear
cboFuente.AddItem "Formalizaciones"
cboFuente.ItemData(cboFuente.ListCount - 1) = CStr(1)

cboFuente.AddItem "Readecuaciones de Plazos"
cboFuente.ItemData(cboFuente.ListCount - 1) = CStr(2)
cboFuente.AddItem "Traspaso de Deudas"
cboFuente.ItemData(cboFuente.ListCount - 1) = CStr(3)
cboFuente.AddItem "Retenciones"
cboFuente.ItemData(cboFuente.ListCount - 1) = CStr(4)

cboFuente.Text = "Formalizaciones"

chkTodos.Value = vbUnchecked
vGrid.MaxCols = 12
vGrid.MaxRows = 0


Call chkRepTodasFechas_Click
Call cboRepTags_Click

End Sub

Private Sub sbCargarCboUsuario()
Dim strSQL As String
On Error GoTo vError
    
    cboUsuarios.Clear
    
    If Mid(cboGrupos.Text, 1, 5) <> "TODOS" Then
        strSQL = "SELECT UPPER(USUARIO) as ItmX from CRD_GRPUSERS WHERE COD_GRUPO = '" & cboGrupos.ItemData(cboGrupos.ListIndex) & "'"
        Call sbCbo_Llena_New(cboUsuarios, strSQL, True, False)
    Else
        cboUsuarios.AddItem "TODOS"
        cboUsuarios.Text = "TODOS"
    End If

    Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 txtConRemesa.Text = ""
    
End Sub

Private Sub sbConsultaOpRemesa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


strSQL = "select A.* from crd_remesas A inner join crd_remesa_asg X on A.remesa = X.remesa where id_solicitud = " & txtOperacion
Call OpenRecordSet(rs, strSQL)
If rs.BOF Or rs.EOF Then
 txtConRemesa.Text = "** No se encontró operación en las remesas registradas **"
Else
 txtConRemesa.Text = "Remesa   " & vbTab & " ...:" & rs!remesa & vbCrLf
 txtConRemesa.Text = txtConRemesa & "Fecha   " & vbTab & " ...:" & rs!fecha & vbCrLf
 txtConRemesa.Text = txtConRemesa & "Usuario  " & vbTab & " ...:" & rs!Usuario
End If
rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 txtConRemesa.Text = ""

End Sub

Private Sub sbParametrosTagRevision()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

    mReqTagRevision = False

    strSQL = "select isnull(valor,'') from CRD_PARAMETROS where cod_parametro = '25'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        If rs.Fields(0) = "S" Then
            mReqTagRevision = True
        Else
            mReqTagRevision = False
        End If
    End If
    rs.Close
    
    If mReqTagRevision = True Then
    
        strSQL = "select isnull(valor,'') from CRD_PARAMETROS where cod_parametro = '26'"
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            mTagRevision = rs.Fields(0)
        End If
        rs.Close
        
    End If
    
    Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCarga_Listado()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass

vGridConsulta.MaxRows = 0
       
Set rsExcel = Excel_Load(txtArchivo.Text, "Import")
        
        With vGridConsulta
        
            Do While Not rsExcel.EOF
                    strSQL = "select R.id_solicitud,R.codigo,D.descripcion as DestinoX,G.descripcion as GarantiaX, R.cedula, S.nombre, X.*, T.descripcion as TagDescripcion" _
                           & ", case when R.estado = 'A' then 'Activo' when R.estado = 'C' then 'Cancelado' when R.estado = 'N' then 'Anulada' else 'No Ident.' end as Estado " _
                           & ",R.montoapr,R.fechaforp" _
                           & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
                           & " inner join crd_garantia_tipos G on R.garantia = G.garantia" _
                           & " left join catalogo_destinos D on R.cod_destino = D.cod_destino" _
                           & " left join crd_remesa_asg A on R.id_solicitud = A.id_solicitud" _
                           & " left join crd_remesas X on A.remesa = X.remesa" _
                           & " left join crd_remesas_tags T on X.tag_codigo  = T.tag_codigo" _
                           & " Where R.id_solicitud = " & rsExcel!Operacion
                    
                    Call OpenRecordSet(rs, strSQL)
                    If Not rs.EOF And Not rs.BOF Then
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                        
                        .col = 1
                        .Text = CStr(rs!Id_Solicitud)
                        .col = 2
                        .Text = CStr(rs!Codigo & "")
                        .col = 3
                        .Text = CStr(rs!DestinoX & "")
                        .col = 4
                        .Text = CStr(rs!GarantiaX & "")
                        
                        .col = 5
                        .Text = Format(rs!montoapr, "Standard")
                        .col = 6
                        .Text = Format(rs!FechaForp, "dd/mm/yyyy")
                        
                        
                        .col = 7
                        .Text = CStr(rs!Cedula)
                        .col = 8
                        .Text = CStr(rs!Nombre)
                        .col = 9
                        .Text = CStr(rs!remesa & "")
                        .col = 10
                        .Text = CStr(rs!Usuario & "")
                        .col = 11
                        .Text = CStr(rs!Microfilm_fecha & "")
                        .col = 12
                        .Text = CStr(rs!microfilm_usuario & "")
                        .col = 13
                        .Text = CStr(rs!TagDescripcion & "")
                        .col = 14
                        .Text = CStr(rs!tag_consecutivo & "")
                        .col = 15
                        .Text = CStr(rs!Estado & "")
                    
                    End If
                    rs.Close
                    
              rsExcel.MoveNext
            Loop
        End With
        

Me.MousePointer = vbDefault
MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGridConsulta.MaxRows = 0
End Sub




Private Sub txtLinea_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String

If KeyCode = vbKeyReturn And txtLinea.Text <> "" Then
  strSQL = "select (R.cod_destino) as 'IdX',rtrim(R.descripcion) as 'ItmX'" _
         & " from catalogo_destinos R inner join catalogo_destinosAsg A on R.cod_destino = A.cod_destino" _
         & " where A.codigo = '" & txtLinea.Text & "' order by R.descripcion"
  Call sbCbo_Llena_New(cboDestino, strSQL, True, True)
End If


End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Call sbConsultaOpRemesa
End Sub


Private Sub sbMicrofilmRecibeLimpia()
   txtReciboUsuario.Text = ""
   txtReciboFecha.Text = ""
End Sub

Private Sub sbMicrofilmRecibeConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

On Error GoTo vError

txtReciboUsuario.Text = "No Existe!"
txtReciboFecha.Text = "No Existe!"

strSQL = "select * from crd_remesas where remesa = " & txtReciboRemesa.Text
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtReciboUsuario.Text = rs!microfilm_usuario & ""
   txtReciboFecha.Text = rs!Microfilm_fecha & ""
End If
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtReciboRemesa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    txtReciboUsuario.SetFocus
Else
  Call sbMicrofilmRecibeLimpia
End If

End Sub

Private Sub txtReciboRemesa_LostFocus()
 Call sbMicrofilmRecibeConsulta
End Sub

Private Sub txtRepRemesas_Change()
 Call sbCargaRemesas
End Sub


Private Function fxGuardarGridTags() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardarGridTags = 0

With vGridTags

.Row = .ActiveRow
.col = 1

strSQL = "select isnull(count(*),0) as Existe from crd_remesas_tags where tag_codigo = '" & .Text & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
   strSQL = "insert crd_remesas_tags(tag_codigo,descripcion,activo,consecutivo) values('" & .Text & "','"
   .col = 2
   strSQL = strSQL & .Text & "',"
   .col = 3
   strSQL = strSQL & .Value & ","
   .col = 4
   strSQL = strSQL & IIf((IsNumeric(.Text)), CLng(.Text), 0) & ")"
   
   Call ConectionExecute(strSQL)
   
   .col = 1
   Call Bitacora("Registra", "Remesas de Crédito [TAG] : " & .Text)
   
   Else 'Actualizar
    .col = 2
    strSQL = "update crd_remesas_tags set descripcion = '" & .Text & "',activo = "
    .col = 3
    strSQL = strSQL & .Value & ",consecutivo = "
    .col = 4
    strSQL = strSQL & IIf((IsNumeric(.Text)), CLng(.Text), 0)
    .col = 1
    strSQL = strSQL & " where tag_codigo = '" & .Text & "'"
   
    Call ConectionExecute(strSQL)
    
    .col = 1
    Call Bitacora("Modifica", "Remesas de Crédito [TAG] : " & .Text)
    
   End If

   .col = 1
   fxGuardarGridTags = 1
   
   
End With

Exit Function
   
vError:
 fxGuardarGridTags = 0
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Function


Private Sub vGridTags_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, i As Long

If vGridTags.ActiveCol = vGridTags.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
        i = fxGuardarGridTags
        vGridTags.Row = vGridTags.ActiveRow
        vGridTags.col = 1
        If vGridTags.MaxRows <= vGridTags.ActiveRow Then
          vGridTags.MaxRows = vGridTags.MaxRows + 1
          vGridTags.Row = vGridTags.MaxRows
        End If
End If

End Sub
