VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCprOrdenes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Compra"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   12015
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   6315
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   14887
            MinWidth        =   14887
            Object.ToolTipText     =   "Estado de adjudicación"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Bevel           =   0
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Bevel           =   0
            TextSave        =   "NÚM"
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCprOrdenes.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   7080
      TabIndex        =   5
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2655
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   11895
      _Version        =   524288
      _ExtentX        =   20981
      _ExtentY        =   4683
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
      MaxCols         =   488
      ScrollBars      =   2
      SpreadDesigner  =   "frmCprOrdenes.frx":08DA
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   1680
      TabIndex        =   7
      Top             =   960
      Width           =   5175
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
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   0
      EndProperty
      Height          =   555
      Left            =   1680
      TabIndex        =   9
      Top             =   1320
      Width           =   10215
      _Version        =   1441793
      _ExtentX        =   18018
      _ExtentY        =   979
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   0
      EndProperty
      Height          =   450
      Left            =   1680
      TabIndex        =   10
      Top             =   480
      Width           =   2535
      _Version        =   1441793
      _ExtentX        =   4471
      _ExtentY        =   794
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
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   1455
      Left            =   0
      TabIndex        =   11
      Top             =   4800
      Width           =   12015
      _Version        =   1441793
      _ExtentX        =   21193
      _ExtentY        =   2566
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   2
      Begin XtremeSuiteControls.GroupBox gbAutorizacion 
         Height          =   1095
         Left            =   0
         TabIndex        =   27
         Top             =   360
         Width           =   6375
         _Version        =   1441793
         _ExtentX        =   11245
         _ExtentY        =   1931
         _StockProps     =   79
         Caption         =   "Autorización"
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
         Begin XtremeSuiteControls.FlatEdit txtAutorizadoUser 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   30
            Top             =   360
            Width           =   4455
            _Version        =   1441793
            _ExtentX        =   7858
            _ExtentY        =   556
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAutorizadoFecha 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   31
            Top             =   720
            Width           =   4455
            _Version        =   1441793
            _ExtentX        =   7858
            _ExtentY        =   556
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            Caption         =   "Autorizado Por:"
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
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha"
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
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   1575
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtSubTotal 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   12
         Top             =   0
         Width           =   1695
         _Version        =   1441793
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDescuento 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   13
         Top             =   360
         Width           =   1695
         _Version        =   1441793
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtImpuestos 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   14
         Top             =   720
         Width           =   1695
         _Version        =   1441793
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotal 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   15
         Top             =   1080
         Width           =   1695
         _Version        =   1441793
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label lblCantidad 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2880
         TabIndex        =   21
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblLineas 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total"
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
         Index           =   6
         Left            =   8400
         TabIndex        =   19
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Descuento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   7
         Left            =   8400
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Impuestos"
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
         Index           =   8
         Left            =   8400
         TabIndex        =   17
         Top             =   690
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Index           =   9
         Left            =   8400
         TabIndex        =   16
         Top             =   1050
         Width           =   975
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtFecha 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   9480
      TabIndex        =   22
      Top             =   600
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   9480
      TabIndex        =   23
      Top             =   960
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin MSComctlLib.Toolbar tlbProcesos 
      Height          =   330
      Left            =   4920
      TabIndex        =   24
      Top             =   0
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      ButtonWidth     =   1958
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Procesos"
            Key             =   "Procesos"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   330
      Left            =   480
      TabIndex        =   25
      Top             =   0
      Width           =   3675
      _ExtentX        =   6482
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repBoleta"
                  Text            =   "Boleta "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repListadoGeneral"
                  Text            =   "Listado General"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repSep1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repBoletaC"
                  Text            =   "Boleta vrs Compras"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   0
      EndProperty
      Height          =   435
      Left            =   4320
      TabIndex        =   26
      Top             =   480
      Width           =   2535
      _Version        =   1441793
      _ExtentX        =   4471
      _ExtentY        =   767
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
      Text            =   "..."
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   8280
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Orden"
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
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
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
      Left            =   8280
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Index           =   10
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmCprOrdenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vMascara As String, vScroll As Boolean

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 cod_orden from cpr_ordenes"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_orden > '" & Format(txtCodigo, vMascara) & "' order by cod_orden asc"
    Else
       strSQL = strSQL & " where cod_orden < '" & Format(txtCodigo, vMascara) & "' order by cod_orden desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!cod_orden)
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
vModulo = 35
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
 vModulo = 35
 vMascara = "0000000000"
 vEdita = True
 
 vGrid.AppearanceStyle = fxGridStyle

 Call sbCprCboTiposOrden(cbo)
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 
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
Dim i As Integer

vCodigo = ""
txtCodigo = ""

txtFecha = Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss")
txtEstado = "Solicitada"
txtUsuario = glogon.Usuario
txtNotas = ""

vGrid.MaxRows = 0
vGrid.MaxRows = 1
vGrid.MaxCols = 7

txtAutorizadoFecha = ""
txtAutorizadoUser = ""

txtSubTotal = 0
txtDescuento = 0
txtImpuestos = 0
txtTotal = 0

txtCodigo.Enabled = True

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      cbo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      cbo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
'       gBusquedas.Columna = "descripcion"
'       gBusquedas.Orden = "descripcion"
'       gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
'       frmBusquedas.Show vbModal
'       txtCodigo.SetFocus
'       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
'       txtNombre.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select O.*,rtrim(C.Tipo_Orden) as 'Causa_Id', rtrim(C.descripcion) as 'Causa_Desc', isnull(Prov.Descripcion,'') as 'Proveedor_Desc'" _
       & " from cpr_ordenes O inner join cpr_Tipo_Orden C on O.Tipo_Orden = C.Tipo_Orden" _
       & " left join CXP_Proveedores Prov on O.cod_Proveedor = Prov.Cod_Proveedor" _
       & " where O.cod_orden = '" & Format(xCodigo, vMascara) & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!cod_orden
  txtCodigo = rs!cod_orden
  
  Call sbCboAsignaDato(cbo, rs!Causa_Desc, True, rs!Causa_Id)

  Select Case rs!Estado
    Case "S"
      txtEstado = "Solicitada"
    Case "A"
      txtEstado = "Autorizada"
    Case "R"
      txtEstado = "Rechazada"
    Case Else
      txtEstado = "No Identificada"
  End Select

  Select Case rs!Proceso
    Case "P"
      txtEstado = txtEstado & " ¦ Pendiente"
    Case "C"
      txtEstado = txtEstado & " ¦ Cotizada"
    Case "A"
      txtEstado = txtEstado & " ¦ Adjudicada"
    Case "D"
      txtEstado = txtEstado & " ¦ Despacho Total"
    Case "X"
      txtEstado = txtEstado & " ¦ Despacho Parcial"
    Case "Y"
      txtEstado = txtEstado & " ¦ Cerrada"
    Case Else
      txtEstado = txtEstado & " ¦ No Identificada"
  End Select


  txtFecha = Format(rs!genera_fecha, "yyyy/mm/dd hh:mm:ss")
  txtUsuario = rs!genera_user & ""
  txtNotas = rs!nota & ""

  If rs!Estado = "A" Or rs!Estado = "R" Then
      txtAutorizadoFecha = Format(rs!Autoriza_Fecha, "yyyy/mm/dd hh:mm:ss")
      txtAutorizadoUser = rs!Autoriza_user & ""
  End If
  
  txtSubTotal = Format(rs!SubTotal, "Standard")
  txtDescuento = Format(rs!descuento, "Standard")
  txtImpuestos = Format(rs!imp_ventas, "Standard")
  txtTotal = Format(rs!Total, "Standard")
  
  If rs!Proveedor_Desc = "" Then
     StatusBarX.Panels.Item(1).Text = ""
  Else
     StatusBarX.Panels.Item(1).Text = "Adjudicada a: " & rs!Proveedor_Desc
  End If
  
'         & "((D.cantidad * D.precio) + ((D.cantidad * D.precio) * (D.imp_ventas / 100))) as Total" _

  strSQL = "select D.cod_producto,P.descripcion,D.cantidad,D.precio,isnull(D.descuento,0) as Descuento" _
         & ",D.imp_ventas, 0 as Total" _
         & " from cpr_ordenes_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " where D.cod_orden = '" & rs!cod_orden & "' order by D.Linea"
  Call sbCargaGrid(vGrid, 7, strSQL)
  Call sbCalculaTotales
  vGrid.Enabled = True
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

On Error GoTo vError

vMensaje = fxInvVerificaLineaDetalle(vGrid, 3, "E", 1)


vError:

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, i As Integer
Dim rs As New ADODB.Recordset

On Error GoTo vError

Call sbCalculaTotales

If vEdita Then
   
   strSQL = "update cpr_ordenes set nota = '" & txtNotas & "',descuento = " _
         & CCur(txtDescuento) & ",subtotal = " & CCur(txtSubTotal) _
         & ",imp_ventas = " & CCur(txtImpuestos) & ",total = " & CCur(txtTotal) _
         & " where cod_orden = " & vCodigo & " and tipo_orden = '" & cbo.ItemData(cbo.ListIndex) & "'"
  If Mid(txtEstado, 1, 1) = "S" Then
      Call ConectionExecute(strSQL)
  Else
      MsgBox "No puede Modificar esta Orden, ya que no se encuentra Solicitada...", vbExclamation
      Exit Sub
  End If
  
  Call Bitacora("Modifica", "Orden Compra : " & vCodigo)

Else

   'Consecutivo de la Orden
   strSQL = "select isnull(max(cod_orden),0) + 1 as Ultimo from cpr_Ordenes"
   Call OpenRecordSet(rs, strSQL)
     vCodigo = Format(rs!ultimo, vMascara)
   rs.Close
   txtCodigo = vCodigo

   strSQL = "insert cpr_ordenes(cod_orden,tipo_orden,estado,genera_fecha,nota,genera_user" _
          & ",subtotal,descuento,imp_ventas,total,pin_autorizacion,pin_entrada,proceso) values('" & vCodigo & "','" _
          & cbo.ItemData(cbo.ListIndex) & "','S',dbo.MyGetdate(),'" & txtNotas & "','" & glogon.Usuario & "'," & CCur(txtSubTotal) _
          & "," & CCur(txtDescuento) & "," & CCur(txtImpuestos) & "," & CCur(txtTotal) & ",0,'','P')"
   Call ConectionExecute(strSQL)

  Call Bitacora("Registra", "Orden Compra : " & vCodigo)

   txtCodigo.Enabled = True

End If

'Guardar Detalle de la Orden
strSQL = "delete cpr_ordenes_detalle where cod_orden = '" & vCodigo & "'"

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.col = 1
  If vGrid.Text <> "" Then
    strSQL = strSQL & Space(10) & "insert cpr_ordenes_detalle(linea,cod_orden,cod_producto,cantidad,estado,precio" _
           & ",descuento,imp_ventas) values(" & i & ",'" & vCodigo & "','" & vGrid.Text & "',"
    vGrid.col = 3
    strSQL = strSQL & CCur(vGrid.Text) & ",'P',"
    vGrid.col = 4
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.col = 5
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.col = 6
    strSQL = strSQL & CCur(vGrid.Text) & ")"
  End If
Next i

'Registra Detalle
Call ConectionExecute(strSQL)

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
   'no se pueden Ejecutar Borrados en Ordenes
'  strSQL = "delete cxp_proveedores where cod_proveedor = " & vCodigo
'  Call ConectionExecute(strSQL)

'  Call Bitacora("Elimina", "ER ESPECIAL : " & vCodigo & " EMP: " & vParametros.CodigoEmpresa)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim i As Integer, vSQL As String

vSQL = ""
i = MsgBox("Desea visualizar solo la Orden Actual", vbYesNo)

If i = vbYes Then
  vSQL = "{cpr_ordenes.COD_ORDEN} = '" & vCodigo & "'"
End If

Select Case ButtonMenu.Key
  Case "repBoleta"
     Call sbInvReportes("OrdenesBoleta", "Boleta", "", vSQL)
    
  Case "repBoletaC"
     Call sbInvReportes("OrdenesBoleta_C", "Boleta", "", vSQL)
  
  Case "repListadoGeneral"
    Call MuestraForms(frmCprReportesGenerales)
  
End Select

End Sub

Private Sub tlbProcesos_ButtonClick(ByVal Button As MSComctlLib.Button)

GLOBALES.gTag = txtCodigo

Call sbFormsCall("frmCprOrdenesProceso", vbModal, , , False, Me)

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_orden"
  gBusquedas.Orden = "cod_orden"
  gBusquedas.Consulta = "select cod_orden,genera_user,nota from cpr_ordenes"
  gBusquedas.Filtro = ""
  gBusquedas.Mascara = vMascara
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtDescuento_GotFocus()
On Error GoTo vError
txtDescuento = CCur(txtDescuento)
Exit Sub
vError:
  MsgBox "Información del Descuento no es válida...", vbCritical
End Sub

Private Sub txtDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTotal.SetFocus
End Sub

Private Sub txtDescuento_LostFocus()
On Error GoTo vError

txtDescuento = Format(CCur(txtDescuento), "Standard")
txtTotal = Format(CCur(txtSubTotal) + CCur(txtImpuestos) - CCur(txtDescuento), "Standard")

Exit Sub
vError:
  MsgBox "Información del Descuento no es válida...", vbCritical
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
End Sub


Private Sub sbCalculaTotales()
Dim curSubTotal As Currency, curIV As Currency
Dim curTmpPrecio As Currency, curTmpIV As Currency, curTmpCant As Currency
Dim i As Integer, lng As Long, curTmpDesc As Currency, curDescuento As Currency
Dim iLineas As Integer, curCantidad As Currency

'**********************************************    OJO
'Revisar esta formula por la situacion del descuento, si es antes o despues del
'impuesto de ventas, por ahora está despues del impuesto

curSubTotal = 0
curIV = 0
curDescuento = 0

iLineas = 0
curCantidad = 0


For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 3
 If vGrid.Text <> "" Then
    curTmpCant = CCur(vGrid.Text)
    vGrid.col = 4
    curTmpPrecio = CCur(vGrid.Text)
    vGrid.col = 5
    curTmpDesc = CCur(vGrid.Text)
    vGrid.col = 6
    curTmpIV = CCur(vGrid.Text)

    curSubTotal = curSubTotal + (curTmpCant * curTmpPrecio)
    curTmpDesc = ((curTmpCant * curTmpPrecio) * (curTmpDesc / 100))
    curDescuento = curDescuento + curTmpDesc
    
    curTmpIV = (((curTmpCant * curTmpPrecio) - curTmpDesc) * (curTmpIV / 100))
    curIV = curIV + curTmpIV
    
    vGrid.col = 7
    vGrid.Text = CStr((curTmpCant * curTmpPrecio) - curTmpDesc + curIV)
    
    curCantidad = curCantidad + curTmpCant
    iLineas = iLineas + 1
 
 End If
Next lng

txtSubTotal = Format(curSubTotal, "Standard")
txtImpuestos = Format(curIV, "Standard")
txtDescuento = Format(curDescuento, "Standard")

txtTotal = Format(curSubTotal + curIV - curDescuento, "Standard")

lblLineas.Caption = "Líneas   : " & iLineas
lblCantidad.Caption = "Cantidad : " & Format(curCantidad, "Standard")


End Sub

Private Sub sbConsultaArticulo(fila As Long, Columna As Integer, vCriterio As String)
Dim strSQL As String, rs As New ADODB.Recordset, vPaso As Boolean

'Busquedas
'1. Por Codigo del Articulo
'2. Por Codigo de Barras
'3. Por Codigo del Fabricante
vPaso = False

strSQL = "select cod_producto,descripcion,costo_regular,impuesto_ventas from pv_productos" _
       & " where cod_producto = '" & vCriterio & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then vPaso = True

If Not vPaso Then
  rs.Close
  strSQL = "select cod_producto,descripcion,costo_regular,impuesto_ventas from pv_productos" _
         & " where cod_barras = '" & vCriterio & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then vPaso = True
End If

If Not vPaso Then
  rs.Close
  strSQL = "select cod_producto,descripcion,costo_regular,impuesto_ventas from pv_productos" _
         & " where cod_fabricante = '" & vCriterio & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then vPaso = True
End If

If Not vPaso Then
  MsgBox "No se encontró el Articulo en la Base de Datos...", vbExclamation
Else
  vGrid.Row = fila
  vGrid.col = 1
  vGrid.Text = rs!Cod_Producto
  vGrid.col = 2
  vGrid.Text = rs!Descripcion
  vGrid.col = 3
  vGrid.Text = 1
  vGrid.col = 4
  vGrid.Text = CStr(rs!costo_regular)
  vGrid.col = 5
  vGrid.Text = "0"
  vGrid.col = 6
  vGrid.Text = CStr(rs!impuesto_ventas)
End If
rs.Close


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(7) As Variant, x As Integer

'Abrir Nueva Linea
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    Call sbCalculaTotales
  End If
End If

'Consulta Articulo
If vGrid.ActiveCol = 1 And KeyCode = vbKeyReturn Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  Call sbConsultaArticulo(vGrid.ActiveRow, vGrid.ActiveCol, vGrid.Text)
End If


'Consular Articulo
If vGrid.ActiveCol = 1 And KeyCode = vbKeyF4 Then
   frmBusquedaArticulos.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 1
   vGrid.Text = gBusquedas.Resultado
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 7
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To 7
        vGrid.col = x
        vTemp(x) = vGrid.Text
     Next x

     vGrid.Row = lng
     For x = 1 To 7
       vGrid.col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  Call sbCalculaTotales
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If



End Sub



Private Sub vGrid_KeyPress(KeyAscii As Integer)
Dim curCantidad As Currency, curPrecio As Currency, curIV As Currency
Dim curDesc As Currency

On Error GoTo vError
'Calcula Total
Select Case vGrid.ActiveCol
  Case 3, 4, 5, 6
    vGrid.Row = vGrid.ActiveRow
    vGrid.col = 3
    curCantidad = CCur(vGrid.Text)
    vGrid.col = 4
    curPrecio = CCur(vGrid.Text)
    vGrid.col = 5
    curDesc = CCur(vGrid.Text)
    vGrid.col = 6
    curIV = CCur(vGrid.Text)
    vGrid.col = 7
    
    curPrecio = curPrecio * curCantidad
    curDesc = curPrecio * (curDesc / 100)
    curIV = (curPrecio - curDesc) * (curIV / 100)
    
    vGrid.Text = curPrecio - curDesc + curIV
   
   Call sbCalculaTotales
  Case Else 'No Aplica
End Select
vError:
End Sub
