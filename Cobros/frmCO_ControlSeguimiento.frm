VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCO_ControlSeguimiento 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento : Control de Cobro"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11790
   Icon            =   "frmCO_ControlSeguimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   11790
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5532
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   11532
      _Version        =   1572864
      _ExtentX        =   20341
      _ExtentY        =   9758
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
      SelectedItem    =   1
      Item(0).Caption =   "Historial"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "tcAux"
      Item(1).Caption =   "Seguimiento"
      Item(1).ControlCount=   20
      Item(1).Control(0)=   "fraLista"
      Item(1).Control(1)=   "txtGestion"
      Item(1).Control(2)=   "txtArreglo"
      Item(1).Control(3)=   "txtArregloDesc"
      Item(1).Control(4)=   "txtCausa"
      Item(1).Control(5)=   "txtCausaDesc"
      Item(1).Control(6)=   "cboOperacion"
      Item(1).Control(7)=   "txtGestionMonto"
      Item(1).Control(8)=   "cmdAplica"
      Item(1).Control(9)=   "txtNotas"
      Item(1).Control(10)=   "txtGestionDesc"
      Item(1).Control(11)=   "dtpVence"
      Item(1).Control(12)=   "Label1(9)"
      Item(1).Control(13)=   "Label1(8)"
      Item(1).Control(14)=   "Label1(7)"
      Item(1).Control(15)=   "Label1(6)"
      Item(1).Control(16)=   "Label1(5)"
      Item(1).Control(17)=   "Label1(3)"
      Item(1).Control(18)=   "Label1(2)"
      Item(1).Control(19)=   "Label1(0)"
      Item(2).Caption =   "Fiadores"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "chkFiadoresEstado"
      Item(2).Control(1)=   "vgridFiadores"
      Item(3).Caption =   "Comisiones"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "lswRetencion"
      Item(3).Control(1)=   "lswComisiones"
      Item(3).Control(2)=   "Label2(0)"
      Item(3).Control(3)=   "Label2(1)"
      Item(4).Caption =   "Estado"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "txtEstado"
      Begin XtremeSuiteControls.PushButton cmdAplica 
         Height          =   615
         Left            =   9240
         TabIndex        =   20
         Top             =   4800
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         Picture         =   "frmCO_ControlSeguimiento.frx":6852
      End
      Begin VB.TextBox txtEstado 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4995
         Left            =   -69880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   11292
      End
      Begin VB.CheckBox chkFiadoresEstado 
         Appearance      =   0  'Flat
         Caption         =   "Solo operaciones atrasadas"
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
         Height          =   330
         Left            =   -65920
         TabIndex        =   15
         Top             =   360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cboOperacion 
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
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Frame fraLista 
         Caption         =   "Gestión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4335
         Left            =   5880
         TabIndex        =   5
         Top             =   360
         Width           =   5175
         Begin XtremeSuiteControls.ListView lswLista 
            Height          =   3372
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   4932
            _Version        =   1572864
            _ExtentX        =   8700
            _ExtentY        =   5948
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
         Begin XtremeSuiteControls.FlatEdit txtListaFiltro 
            Height          =   330
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   4935
            _Version        =   1572864
            _ExtentX        =   8705
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
      End
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   5412
         Left            =   -70000
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   11532
         _Version        =   1572864
         _ExtentX        =   20341
         _ExtentY        =   9546
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
         Item(0).Caption =   "Gestiones"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "vgCobro"
         Item(1).Caption =   "Detalle"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "lswDetalle"
         Begin FPSpreadADO.fpSpread vgCobro 
            Height          =   4452
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   11412
            _Version        =   524288
            _ExtentX        =   20129
            _ExtentY        =   7853
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
            MaxCols         =   11
            SpreadDesigner  =   "frmCO_ControlSeguimiento.frx":702A
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin MSComctlLib.ListView lswDetalle 
            Height          =   4695
            Left            =   -70000
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   8281
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   16711680
            BackColor       =   14737632
            BorderStyle     =   1
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "# Oper."
               Object.Width           =   1835
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Código"
               Object.Width           =   1482
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Cuotas"
               Object.Width           =   1835
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Mora"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Saldo"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Abono"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Est/Actual"
               Object.Width           =   2011
            EndProperty
         End
      End
      Begin FPSpreadADO.fpSpread vgridFiadores 
         Height          =   4572
         Left            =   -69880
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   11292
         _Version        =   524288
         _ExtentX        =   19918
         _ExtentY        =   8065
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
         MaxCols         =   7
         SpreadDesigner  =   "frmCO_ControlSeguimiento.frx":7CCC
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.ListView lswRetencion 
         Height          =   1932
         Left            =   -69880
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   11292
         _ExtentX        =   19923
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#Operación"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripción"
            Object.Width           =   5539
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Abonos"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Gestiones"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Monto"
            Object.Width           =   2011
         EndProperty
      End
      Begin MSComctlLib.ListView lswComisiones 
         Height          =   1932
         Left            =   -69880
         TabIndex        =   18
         Top             =   3240
         Visible         =   0   'False
         Width           =   11292
         _ExtentX        =   19923
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Remesa"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Usuario"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "# Tesoreria "
            Object.Width           =   2540
         EndProperty
      End
      Begin XtremeSuiteControls.DateTimePicker dtpVence 
         Height          =   315
         Left            =   1440
         TabIndex        =   23
         Top             =   1920
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.FlatEdit txtGestion 
         Height          =   330
         Left            =   1440
         TabIndex        =   27
         Top             =   480
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
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
      Begin XtremeSuiteControls.FlatEdit txtGestionDesc 
         Height          =   330
         Left            =   2280
         TabIndex        =   28
         Top             =   480
         Width           =   3495
         _Version        =   1572864
         _ExtentX        =   6165
         _ExtentY        =   582
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
         BackColor       =   16777152
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCausa 
         Height          =   330
         Left            =   1440
         TabIndex        =   29
         Top             =   840
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
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
      Begin XtremeSuiteControls.FlatEdit txtCausaDesc 
         Height          =   330
         Left            =   2280
         TabIndex        =   30
         Top             =   840
         Width           =   3495
         _Version        =   1572864
         _ExtentX        =   6165
         _ExtentY        =   582
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
         BackColor       =   16777152
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtArreglo 
         Height          =   330
         Left            =   1440
         TabIndex        =   31
         Top             =   1200
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
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
      Begin XtremeSuiteControls.FlatEdit txtArregloDesc 
         Height          =   330
         Left            =   2280
         TabIndex        =   32
         Top             =   1200
         Width           =   3495
         _Version        =   1572864
         _ExtentX        =   6165
         _ExtentY        =   582
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
         BackColor       =   16777152
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtGestionMonto 
         Height          =   330
         Left            =   1440
         TabIndex        =   33
         Top             =   1560
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   2415
         Left            =   1320
         TabIndex        =   34
         Top             =   3000
         Width           =   4455
         _Version        =   1572864
         _ExtentX        =   7858
         _ExtentY        =   4260
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   1
         Left            =   -69880
         TabIndex        =   22
         Top             =   2880
         Visible         =   0   'False
         Width           =   3492
         _Version        =   1572864
         _ExtentX        =   6159
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Comisiones:"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   0
         Left            =   -69880
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   3492
         _Version        =   1572864
         _ExtentX        =   6159
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Retenciones:"
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
      End
      Begin VB.Label Label1 
         Caption         =   "Gestión"
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
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pago"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
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
         Left            =   240
         TabIndex        =   12
         Top             =   3120
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "Monto"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   972
      End
      Begin VB.Label Label1 
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
         Height          =   252
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "a la que se le va a registrar el recargo"
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
         Index           =   7
         Left            =   3360
         TabIndex        =   9
         Top             =   2520
         Width           =   2412
      End
      Begin VB.Label Label1 
         Caption         =   "Causas"
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
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "Acuerdo"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   972
      End
   End
   Begin MSComctlLib.ImageList imgSemaforos 
      Left            =   10800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlSeguimiento.frx":869D
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlSeguimiento.frx":87BB
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlSeguimiento.frx":88E1
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlSeguimiento.frx":8A0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlSeguimiento.frx":8B1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlSeguimiento.frx":8C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlSeguimiento.frx":8D35
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlSeguimiento.frx":8E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlSeguimiento.frx":8F81
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlSeguimiento.frx":90A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlSeguimiento.frx":91CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2640
      TabIndex        =   24
      Top             =   480
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4440
      TabIndex        =   25
      Top             =   480
      Width           =   6012
      _Version        =   1572864
      _ExtentX        =   10604
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
   Begin VB.Image imgDatosPersonales 
      Height          =   360
      Left            =   11040
      Picture         =   "frmCO_ControlSeguimiento.frx":95B6
      Stretch         =   -1  'True
      ToolTipText     =   "Datos Personales"
      Top             =   480
      Width           =   360
   End
   Begin VB.Image imgExpediente 
      Height          =   360
      Left            =   10560
      Picture         =   "frmCO_ControlSeguimiento.frx":9D49
      Stretch         =   -1  'True
      Top             =   480
      Width           =   348
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Expediente"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "Cuenta de Inventarios para la Bodega"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   11892
   End
End
Attribute VB_Name = "frmCO_ControlSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean
Dim vDesviacionMax As Double, vDesviacionMin As Double
Dim vTipoGestion As String


Private Sub chkFiadoresEstado_Click()
Dim strSQL As String

Me.MousePointer = vbHourglass
On Error GoTo vError


    strSQL = "select M.ESTADO as 'EstadoMora','',F.Id_Solicitud, S.cedula,S.nombre,E.descripcion as Estado,I.descripcion as Inst " _
            & " from fiadores F inner join Socios S on F.cedulaf = S.cedula" _
            & " inner join Instituciones I on S.cod_institucion = I.cod_institucion" _
            & " inner join Reg_Creditos R on F.Id_Solicitud = R.Id_Solicitud" _
            & " inner join AFI_ESTADOS_PERSONA E on E.cod_estado = S.estadoActual" _
            & " left join MOROSIDAD M on F.Id_Solicitud = M.Id_Solicitud and M.Estado = 'A'" _
            & "  where F.estado = 'A' and R.cedula = '" & txtCedula.Text & "' and R.Estado = 'A'"
            
     If chkFiadoresEstado.Value = vbChecked Then
        strSQL = strSQL & " and M.ESTADO = 'A'"
     End If
     
     strSQL = strSQL & " group by F.Id_Solicitud,S.cedula,M.Estado,S.nombre,E.descripcion,I.descripcion"
           
     
    Call sbCargaGridFiadores(vgridFiadores, strSQL)

Me.MousePointer = vbDefault
Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cmdAplica_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String, mAccesoRestringido As Boolean

On Error GoTo vError

Me.MousePointer = vbHourglass

'Verifica datos
vMensaje = ""

'If txtEstado.Tag = "N" Then vMensaje = vMensaje & " - La persona no se encuentra morosa verifique..." & vbCrLf
If txtNotas.Text = "" Then vMensaje = vMensaje & " - No se especificó ninguna observación..." & vbCrLf

strSQL = "select isnull(count(*),0) as Existe from cbr_usuarios where usuario = '" _
       & glogon.Usuario & "' and estado = 1"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & " - El usuario actual no se encuentra activo..." & vbCrLf
rs.Close

strSQL = "select isnull(count(*),0) as Existe from cbr_gestiones where cod_gestion = '" _
       & txtGestion & "' and estado = 1 and NIVEL_GESTION = 'U'"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & " - La gestion actual no se encuentra activa..." & vbCrLf
rs.Close

'Preguntar si existe el parametro de sgt sin asignacion previa, de lo contrario buscar asignacion
strSQL = "select valor from cbr_parametros where cod_parametro = '05'"
Call OpenRecordSet(rs, strSQL)
If Mid(rs!Valor, 1, 1) <> "S" Then
  rs.Close
  strSQL = "select isnull(count(*),0) as Existe from cbr_asignacion where usuario = '" _
       & glogon.Usuario & "' and cedula = '" & txtCedula & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then vMensaje = vMensaje & " - Este expediente no se encuentra asignado al usuario actual, verifique..." & vbCrLf
End If
rs.Close

If vMensaje <> "" Then
  Me.MousePointer = vbDefault
  MsgBox vMensaje, vbExclamation
  Exit Sub
End If


strSQL = "exec spCBRControlSGT '" & txtCedula & "','" & glogon.Usuario & "','" & txtGestion _
       & "','" & Format(dtpVence.Value, "yyyy/mm/dd") & "','" & txtNotas & "','" & GLOBALES.gOficinaTitular _
       & "'," & CCur(txtGestionMonto.Text) & ""
       
If cboOperacion.Text = "+ Antigua" Then
   strSQL = strSQL & ",0"
Else
   strSQL = strSQL & "," & cboOperacion.Text
End If

strSQL = strSQL & "," & "'" & txtCausa.Text & "','" & txtArreglo.Text & "'"
       
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

MsgBox "Seguimiento Registrado Satisfactoriamente...", vbInformation
Call sbCargaDatos(txtCedula)

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub dtpVence_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboOperacion.SetFocus
End Sub

Private Sub Form_Activate()
 vModulo = 4
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

With lswLista.ColumnHeaders
    .Clear
    .Add , , "ID", 640
    .Add , , "Detalle", 3850
End With


txtGestionMonto.Locked = True
 
 vScroll = False
' FlatScrollBar.Value = 0
 vScroll = True
 
vgridFiadores.Enabled = True
vgridFiadores.Visible = True

i = 30

strSQL = "select valor from cbr_parametros where cod_parametro = '01'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  i = rs!Valor
End If
rs.Close

strSQL = "select tiempo_resolucion_com from cbr_usuarios where usuario = '" _
       & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
 If rs!tiempo_resolucion_com <= i Then i = rs!tiempo_resolucion_com
End If
rs.Close

'Solo se inicializan estas variales una vez
dtpVence.Value = fxFechaServidor
dtpVence.MinDate = dtpVence.Value
dtpVence.MaxDate = DateAdd("d", i, dtpVence.Value)


txtGestion = ""
txtGestionDesc = ""


tcMain.Item(0).Selected = True
     
End Sub

Private Sub sbLimpia()

Select Case tcMain.SelectedItem
  Case 0 'Historico
     vgCobro.MaxRows = 0
     lswDetalle.ListItems.Clear
  Case 1
     txtEstado.Tag = "N"
     txtEstado.Text = ""
     txtNotas = ""
     cboOperacion.Clear
     cboOperacion.AddItem "+ Antigua"
     cboOperacion.Text = "+ Antigua"
     
  Case 2
     lswRetencion.ListItems.Clear
     lswComisiones.ListItems.Clear
End Select

End Sub


Public Sub sbCargaDatos(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem, i As Integer
Dim curMonto As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbLimpia

If txtCedula.Text <> vCedula Then
    txtCedula.Text = vCedula
    txtNombre.Text = fxNombre(vCedula)
End If

Select Case tcMain.SelectedItem
  Case 0 'Historico
    Call vgCobro_SheetChanged(vgCobro.ActiveSheet, vgCobro.ActiveSheet)
  
  Case 1 'Seguimientos
    Call sbCBRControlEstadoTxtCbo(txtCedula, txtEstado, cboOperacion)
    Call sbTraeUltimaGestiones(txtCedula.Text, glogon.Usuario)
  
  Case 2 'Fiadores
     Call chkFiadoresEstado_Click
  
  Case 3 'Retenciones y Comisiones
       
       strSQL = "select C.usuario,C.cod_Remesa,C.monto,C.Tesoreria_Numero,C.Tesoreria_Fecha" _
              & " from cbr_comisiones_detalle C inner join cbr_Segdetalle D on C.Cod_remesa = D.cod_remesa" _
              & " inner join cbr_seguimiento S on D.cod_seg = S.cod_seg" _
              & " where S.cedula = '" & txtCedula & "'" _
              & " group by C.usuario,C.cod_Remesa,C.monto,C.Tesoreria_Numero,C.Tesoreria_Fecha"
       
       curMonto = 0
       Call OpenRecordSet(rs, strSQL)
       Do While Not rs.EOF
         Set itmX = lswComisiones.ListItems.Add(, , rs!cod_remesa)
             itmX.SubItems(1) = rs!Usuario
             itmX.SubItems(2) = Format(rs!Monto, "Standard")
             itmX.SubItems(3) = Format(rs!tesoreria_fecha & "", "dd/mm/yyyy")
             itmX.SubItems(4) = rs!tesoreria_numero & ""
            curMonto = curMonto + rs!Monto
         rs.MoveNext
       Loop
       rs.Close
         Set itmX = lswComisiones.ListItems.Add(, , "")
             itmX.SubItems(2) = "__________"

         Set itmX = lswComisiones.ListItems.Add(, , "")
             itmX.SubItems(2) = Format(curMonto, "Standard")


End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub imgDatosPersonales_Click()
     GLOBALES.gCedulaActual = Trim(txtCedula)
     frmCR_VerificaDatosPersonales.Show vbModal
End Sub

Private Sub imgExpediente_Click()

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Módulo de Cobro"
    
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "Empresa = '" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "Fecha = '" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    
    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ControlExpediente.rpt")
    .SelectionFormula = "{SOCIOS.CEDULA} = '" & txtCedula.Text & "'"
       
    .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

'Devuelve el codigo de la causa de mora o la gestion seleccionada
Private Sub lswLista_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
 
  
  Select Case Mid(vTipoGestion, 1, 1)
    Case "G"
        txtGestion.Text = Item.Text
        txtGestionDesc.Text = Item.SubItems(1)
        
        If Right(vTipoGestion, 1) = "C" Then
           txtGestion.SetFocus
        Else
           txtGestionDesc.SetFocus
        End If
       
    Case "C"
        txtCausa.Text = Item.Text
        txtCausaDesc.Text = Item.SubItems(1)
        
        If Right(vTipoGestion, 1) = "C" Then
           txtCausa.SetFocus
        Else
           txtCausaDesc.SetFocus
        End If
        
    Case "A"
        txtArreglo.Text = Item.Text
        txtArregloDesc.Text = Item.SubItems(1)
        
        If Right(vTipoGestion, 1) = "C" Then
           txtArreglo.SetFocus
        Else
           txtArregloDesc.SetFocus
        End If
        
  End Select
  
  
End Sub



Private Sub tcAux_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
  Call txtCedula_LostFocus
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

    Call sbCargaDatos(txtCedula.Text)

End Sub

Private Sub txtArregloDesc_GotFocus()
    vTipoGestion = "AD"
    Call sbCargaLista(vTipoGestion)
End Sub

Private Sub txtCausaDesc_GotFocus()
    vTipoGestion = "CD"
    Call sbCargaLista(vTipoGestion)
End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select cedula,nombre from socios"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    txtCedula = Trim(gBusquedas.Resultado)
End If

End Sub

Private Sub txtCedula_LostFocus()
  txtNombre = fxNombre(txtCedula)
  If txtNombre <> "" Then
    Call sbCargaDatos(txtCedula)
  Else
    Call sbLimpia
  End If
End Sub

Private Sub txtArreglo_GotFocus()
 vTipoGestion = "AC"
 Call sbCargaLista(vTipoGestion)
End Sub

Private Sub txtArreglo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   Call sbConsultaGestion(vTipoGestion)
   txtArregloDesc.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_ARREGLO,DESCRIPCION from CBR_TIPOS_ARREGLOS"
    gBusquedas.Columna = "COD_ARREGLO"
    gBusquedas.Orden = "COD_ARREGLO"
    gBusquedas.Filtro = " and ACTIVO = 1 "
    frmBusquedas.Show vbModal
    txtArreglo = Trim(gBusquedas.Resultado)
    txtArregloDesc = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtCausa_GotFocus()
 vTipoGestion = "CC"
 Call sbCargaLista(vTipoGestion)
End Sub

Private Sub txtCausa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   Call sbConsultaGestion(vTipoGestion)
   txtCausaDesc.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_CAUSA,DESCRIPCION from CBR_CAUSAS_MOROSIDAD"
    gBusquedas.Columna = "COD_CAUSA"
    gBusquedas.Orden = "COD_CAUSA"
    gBusquedas.Filtro = " and ACTIVA = 1  "
    frmBusquedas.Show vbModal
    txtCausa = Trim(gBusquedas.Resultado)
    txtCausaDesc = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtArregloDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_ARREGLO,DESCRIPCION from CBR_TIPOS_ARREGLOS"
    gBusquedas.Columna = "COD_ARREGLO"
    gBusquedas.Orden = "COD_ARREGLO"
    gBusquedas.Filtro = " and ACTIVO = 1 "
    frmBusquedas.Show vbModal
    txtArreglo = Trim(gBusquedas.Resultado)
    txtArregloDesc = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtCausaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   Call sbConsultaGestion(vTipoGestion)
   txtArreglo.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_CAUSA,DESCRIPCION from CBR_CAUSAS_MOROSIDAD"
    gBusquedas.Columna = "COD_CAUSA"
    gBusquedas.Orden = "COD_CAUSA"
    gBusquedas.Filtro = " and ACTIVA = 1  "
    frmBusquedas.Show vbModal
    txtCausa = Trim(gBusquedas.Resultado)
    txtCausaDesc = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtGestion_GotFocus()
If vTipoGestion <> "GC" Then
 vTipoGestion = "GC"
 Call sbCargaLista(vTipoGestion)
End If
End Sub

Private Sub txtGestion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtGestionDesc.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select cod_gestion,descripcion from cbr_gestiones"
    gBusquedas.Columna = "cod_gestion"
    gBusquedas.Orden = "cod_gestion"
    gBusquedas.Filtro = " and estado = 1 and nivel_gestion = 'U' "
    frmBusquedas.Show vbModal
    txtGestion.Text = Trim(gBusquedas.Resultado)
    txtGestionDesc.Text = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtGestion_LostFocus()
  Call sbCBRControlGestion
End Sub

Private Sub txtGestionDesc_GotFocus()
If vTipoGestion <> "GD" Then
 vTipoGestion = "GD"
 Call sbCargaLista(vTipoGestion)
End If
End Sub

Private Sub txtGestionDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtCausa.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select cod_gestion,descripcion from cbr_gestiones"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Filtro = " and estado = 1 and nivel_gestion = 'U' "
    frmBusquedas.Show vbModal
    txtGestion = Trim(gBusquedas.Resultado)
    txtGestionDesc = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtGestionDesc_LostFocus()
    Call sbCBRControlGestion
End Sub

Private Sub txtGestionMonto_LostFocus()
    If txtGestionMonto.Text = Empty Then
        txtGestionMonto.Text = Format(0, "Standard")
    End If
    
    If Not IsNumeric(txtGestionMonto) Then
        txtGestionMonto.Text = Format(0, "Standard")
    End If
    txtGestionMonto = Format(txtGestionMonto, "Standard")
    
    If txtGestionMonto < vDesviacionMin Then
        MsgBox "El monto es menor que la desviación mínima"
        txtGestionMonto = Format(vDesviacionMin, "Standard")
        txtGestionMonto.SetFocus
    End If

    If txtGestionMonto > vDesviacionMax Then
        MsgBox "El monto es mayor que la desviación máxima"
        txtGestionMonto = Format(vDesviacionMax, "Standard")
        txtGestionMonto.SetFocus
    End If
    
End Sub


Private Sub txtListaFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    Call sbCargaLista(vTipoGestion, txtListaFiltro.Text)
  End If
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select cedula,nombre from socios"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    txtCedula = Trim(gBusquedas.Resultado)
    txtCedula.SetFocus
End If

End Sub

Private Sub sbCBRControlGestion()
Dim strSQL As String, rs As New ADODB.Recordset

    If txtGestion.Text = Empty Then
        Exit Sub
    End If

    strSQL = "select descripcion, isnull(monto,0) as Monto, MODIFICA_USUARIO, isnull(MODIFICA_DESVIACION,0) as MODIFICA_DESVIACION " _
        & " from cbr_gestiones where estado = 1 " _
        & " and nivel_gestion = 'U' and cod_gestion = '" & Trim(txtGestion) & "'"
    Call OpenRecordSet(rs, strSQL)

    strSQL = ""
    If Not rs.EOF Then
    
        txtGestionDesc = Trim(rs!Descripcion)
        txtGestionMonto = Format(rs!Monto, "Standard")
        vDesviacionMax = rs!Monto + rs!MODIFICA_DESVIACION
        vDesviacionMin = rs!Monto - rs!MODIFICA_DESVIACION
        txtGestionMonto.ToolTipText = "Min: " & Format(vDesviacionMin, "Standard") & " Max: " & Format(vDesviacionMax, "Standard")
        If rs!MODIFICA_USUARIO = 1 Then
           txtGestionMonto.Locked = False
        Else
           txtGestionMonto.Locked = True
        End If
        
    Else
    
        txtGestion.Text = Empty
        txtGestionDesc.Text = Empty
        txtGestionMonto = Format(0, "Standard")
        txtGestionMonto.Locked = True
        vDesviacionMax = 0
        vDesviacionMin = 0
        
    End If
    rs.Close
    
    strSQL = "select dbo.fxCBRGestionUsuario('" & Trim(txtGestion) & "','" & glogon.Usuario & "') as acceso"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
         If rs!Acceso = 0 Then
            MsgBox "El usuario no tiene acceso a esta gestión"
            
            txtGestion.Text = Empty
            txtGestionDesc.Text = Empty
            txtGestionMonto = Format(0, "Standard")
            txtGestionMonto.Locked = True
            vDesviacionMax = 0
            vDesviacionMin = 0
            
            txtGestion.SetFocus
            Exit Sub
         End If
    End If
    rs.Close

End Sub

'Carga la lista con : Gestiones, Causas de morosidad o tipos de areglos
Private Sub sbCargaLista(vTipoGestion As String, Optional vFiltro As String = "")
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError

If vTipoGestion = "" Then Exit Sub

Select Case Mid(vTipoGestion, 1, 1)
   
   Case "G" 'Consulta de gestiones
     fraLista.Caption = "Gestiones"
     strSQL = "Select COD_GESTION as 'Codigo',DESCRIPCION from CBR_GESTIONES" _
            & " where ESTADO = 1 and  NIVEL_GESTION = 'U'"
    
        If vFiltro <> "" Then
            If Right(vTipoGestion, 1) = "C" Then
               strSQL = strSQL & " and COD_GESTION like '%" & txtListaFiltro.Text & "%'"
            Else
               strSQL = strSQL & " and DESCRIPCION like '%" & txtListaFiltro.Text & "%'"
            End If
        End If
            
    Case "C"  'Consulta de Causas de Mora
      fraLista.Caption = "Causas de Mora"
      strSQL = "Select COD_CAUSA as 'Codigo',DESCRIPCION from CBR_CAUSAS_MOROSIDAD" _
             & " where ACTIVA = 1"
      If vFiltro <> "" Then
         If Right(vTipoGestion, 1) = "C" Then
            strSQL = strSQL & " and COD_CAUSA like '%" & txtListaFiltro.Text & "%'"
         Else
            strSQL = strSQL & " and DESCRIPCION like '%" & txtListaFiltro.Text & "%'"
         End If
      End If
            
    Case "A" 'Consulta de Tipos de Arreglos
      fraLista.Caption = "Arreglos"
      strSQL = "Select COD_ARREGLO as 'Codigo',DESCRIPCION from CBR_TIPOS_ARREGLOS" _
             & " where ACTIVO = 1"
      
      If vFiltro <> "" Then
         If Right(vTipoGestion, 1) = "C" Then
          strSQL = strSQL & " and COD_ARREGLO like '%" & txtListaFiltro.Text & "%'"
         Else
          strSQL = strSQL & " and DESCRIPCION like '%" & txtListaFiltro.Text & "%'"
         End If
      End If
      
End Select

If Right(vTipoGestion, 1) = "C" Then
    fraLista.Caption = fraLista.Caption & " [Código]"
Else
    fraLista.Caption = fraLista.Caption & " [Descripción]"
End If

Call OpenRecordSet(rs, strSQL)
     
lswLista.ListItems.Clear
     
Do While Not rs.EOF
  Set itmX = lswLista.ListItems.Add(, , Trim(rs!Codigo))
      itmX.SubItems(1) = rs!Descripcion
  rs.MoveNext
Loop

rs.Close
Me.MousePointer = vbDefault


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

'Consulta el codigo y la descripcion de las gestiones
'la causa de mora y el tipo de arreglo
Private Sub sbConsultaGestion(vTipoGestion As String)
Dim strSQL As String, rs As New ADODB.Recordset
On Error GoTo vError

Select Case Mid(vTipoGestion, 1, 1)
   
   Case "G" 'Consulta de gestiones
      If Mid(vTipoGestion, 2, 1) = "C" Then
        strSQL = "Select COD_GESTION as 'Codigo',DESCRIPCION, MONTO from CBR_GESTIONES" _
               & " where COD_GESTION like '%" & txtGestion.Text & "%' and ESTADO = 1 and  NIVEL_GESTION = 'U'"
      Else
        strSQL = "Select COD_GESTION as 'Codigo',DESCRIPCION, MONTO from CBR_GESTIONES" _
               & " where DESCRIPCION like  '%" & Trim(txtGestionDesc.Text) & "%' and ESTADO = 1 and  NIVEL_GESTION = 'U'"
      End If
      
      Call OpenRecordSet(rs, strSQL)
      strSQL = ""

      If Not rs.EOF Then
        txtGestionDesc = Trim(rs!Descripcion)
        txtGestionMonto = Format(rs!Monto, "Standard")
      Else
        txtGestion.Text = Empty
        txtGestionDesc.Text = Empty
        txtGestionMonto = Format(0, "Standard")
      End If
   
   
   Case "C"
      If Mid(vTipoGestion, 2, 1) = "C" Then
        strSQL = "Select COD_CAUSA as 'Codigo',DESCRIPCION from CBR_CAUSAS_MOROSIDAD" _
               & " where COD_CAUSA like  '%" & Trim(txtCausa.Text) & "%' and ACTIVA = 1"
      Else
        strSQL = "Select COD_CAUSA as 'Codigo',DESCRIPCION from CBR_CAUSAS_MOROSIDAD" _
               & " where DESCRIPCION like '%" & Trim(txtCausaDesc.Text) & "%' and ACTIVA = 1"
      End If
      
      Call OpenRecordSet(rs, strSQL)
      strSQL = ""

      If Not rs.EOF Then
        txtCausaDesc = Trim(rs!Descripcion)
      Else
        txtCausa.Text = Empty
        txtCausaDesc.Text = Empty
      End If
    
    
    Case "A"
      If Mid(vTipoGestion, 2, 1) = "C" Then
        strSQL = "Select COD_ARREGLO as 'Codigo',DESCRIPCION from CBR_TIPOS_ARREGLOS" _
               & " where COD_ARREGLO like '%" & Trim(txtArreglo.Text) & "%' and ACTIVO = 1"
      Else
        strSQL = "Select COD_ARREGLO as 'Codigo',DESCRIPCION from CBR_TIPOS_ARREGLOS" _
               & " where DESCRIPCION like '%" & Trim(txtArregloDesc.Text) & "%' and ACTIVO = 1"
      End If
      
      Call OpenRecordSet(rs, strSQL)
      strSQL = ""

      If Not rs.EOF Then
        txtArregloDesc = Trim(rs!Descripcion)
      Else
        txtArreglo.Text = Empty
        txtArregloDesc.Text = Empty
      End If

End Select

rs.Close

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbTraeUltimaGestiones(vCedula As String, vUsuario As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "Select ULT_COD_GESTION, GestionDesc , COD_CAUSA, CausaDesc, COD_ARREGLO, ArregloDesc" _
       & " From dbo.vCBRControlListado" _
       & " where Cedula = '" & vCedula & "' and usuario = '" & vUsuario & "' "
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.EOF Then
  txtGestion.Text = rs!ULT_COD_GESTION
  txtGestionDesc.Text = rs!GestionDesc
  txtCausa.Text = rs!COD_CAUSA
  txtCausaDesc.Text = rs!CausaDesc
  txtArreglo.Text = rs!COD_ARREGLO
  txtArregloDesc.Text = rs!ArregloDesc
End If

Call sbConsultaGestion("GC")
Call sbConsultaGestion("CC")
Call sbConsultaGestion("AC")

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub sbCargaGridFiadores(vGrid As Object, strSQL As String)
Dim i As Integer, rs As New ADODB.Recordset

vGrid.MaxCols = 7
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.Col = i
 vGrid.Text = ""
Next i

Call OpenRecordSet(rs, strSQL)

With vGrid
  
.MaxRows = 0

Do While Not rs.EOF
  .MaxRows = .MaxRows + 1
  .Row = .MaxRows
  
  For i = 1 To .MaxCols
     .Col = i
     Select Case i
        Case 1 'Status
           If rs!EstadoMora = "A" Then
              .TypePictPicture = imgSemaforos.ListImages.Item(3).Picture
           Else
              .TypePictPicture = imgSemaforos.ListImages.Item(1).Picture
           End If
           
              
        Case 3 'Solicitud
           .Text = CStr(rs!ID_SOLICITUD)
           
        Case 4 'Cedula
           .Text = rs!Cedula
           
        Case 5 'Nombre
           .Text = rs!Nombre
           
        Case 6 ' Estado
           .Text = rs!Estado
           
        Case 7 ' Institución
           .Text = rs!Inst
           
     End Select
  Next i
  
  rs.MoveNext
Loop
rs.Close
End With


End Sub


Private Sub vgCobro_Click(ByVal Col As Long, ByVal Row As Long)
'Private Sub lswSGT_Click()
'Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListItem
'Dim curSaldo As Currency, curAbono As Currency, curMora As Currency
'
'If lswSGT.ListItems.Count = 0 Then Exit Sub
'If lswSGT.SelectedItem = "" Then Exit Sub
'
'Me.MousePointer = vbHourglass
'
'
'strSQL = "select R.codigo,D.*,dbo.MyGetdate() as FechaX" _
'       & " from cbr_SegDetalle D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
'       & " where D.cod_seg = " & lswSGT.SelectedItem
'Call OpenRecordSet(rs, strSQL)
'
'lblNotas.Caption = "ID: " & lswSGT.SelectedItem & " Nota : " & lswSGT.SelectedItem.ToolTipText
'
'rs.Close
'
'Me.MousePointer = vbDefault
'
'End Sub

End Sub

Private Sub vgCobro_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
        
With vgCobro
    .Sheet = NewSheet
    .MaxRows = 0
    
    Select Case NewSheet
      Case 1 'Gestiones
       strSQL = "select S.*, isnull(G.descripcion,'') as 'Gestion'" _
              & "   , isnull(C.DESCRIPCION,'') as 'Causa'" _
              & "   , isnull(A.descripcion,'') as 'Arreglo'" _
              & " from CBR_Seguimiento S  left join cbr_gestiones G on S.cod_gestion = G.cod_gestion" _
              & "  left join CBR_CAUSAS_MOROSIDAD C on S.COD_CAUSA = C.COD_CAUSA" _
              & "  left join CBR_TIPOS_ARREGLOS A on S.COD_ARREGLO = A.COD_ARREGLO" _
              & " where cedula = '" & txtCedula.Text & "' order by S.cod_seg desc"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          
          For i = 1 To 11
            .Col = i
            Select Case i
              Case 1 'ID
                .Text = CStr(rs!Cod_Seg)
              Case 2 'Fecha
                .Text = Format(rs!fecha, "dd/mm/yyyy")
              Case 3 'vencimiento
                .Text = Format(DateAdd("d", rs!tiempo_resolucion, rs!fecha), "dd/mm/yyyy")
              Case 4 'Gestión
                .Text = rs!Gestion
              Case 5 ' Detalle
                .Text = rs!Notas
                .RowHeight(.Row) = .MaxTextRowHeight(.Row)

              Case 6 ' Ejecutivo
                .Text = rs!Usuario
              Case 7 ' Monto
                .Text = Format(rs!Monto, "Standard")
              Case 8 ' Dias
                .Text = CStr(rs!tiempo_resolucion)
              Case 9  'Arrelgo de Pago
                .Text = rs!Arreglo
              Case 10 'Promesa de Pago
                .Text = Format(rs!Arreglo_Vence & "", "dd/mm/yyyy")
              Case 11 'Causa de Morosidad
                .Text = rs!Causa
                
            End Select
          Next i
          rs.MoveNext
        Loop
        rs.Close
      
      Case 2 'Oficiales
      
        strSQL = "select * from cbr_asignacion_h where cedula = '" & txtCedula.Text _
               & "' order by fecha_asignacion desc"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          
          For i = 1 To 5
            .Col = i
            Select Case i
              Case 1 'Fecha
                .Text = Format(rs!fecha_asignacion, "dd/mm/yyyy")
              Case 2 'Oficial
                .Text = UCase(rs!Usuario)
              Case 3 'Mantiene
                .Value = rs!mantener
              Case 4 ' Rebajo 2x
                .Value = rs!rebajo_doble
              Case 5 ' Mora
                .Value = rs!aplica_mora
            End Select
          Next i
          rs.MoveNext
        Loop
        rs.Close
      
    End Select
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub vgridFiadores_Click(ByVal Col As Long, ByVal Row As Long)
  With vgridFiadores
    .Row = Row
    
    Select Case Col
      Case 2
        .Col = 4
        If .Text = "" Then Exit Sub
        GLOBALES.gCedulaActual = .Text
        Call sbFormsCall("frmCR_VerificaDatosPersonales", 1, , , False, Me)
    
    End Select
  End With
End Sub
