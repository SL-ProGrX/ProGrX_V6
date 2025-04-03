VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCprCompraDirecta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras Directas con Facturas Proveedor"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   12000
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   1452
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   11652
      _Version        =   1441793
      _ExtentX        =   20553
      _ExtentY        =   2561
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   2
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
         Height          =   312
         Left            =   9840
         TabIndex        =   20
         Top             =   0
         Width           =   1692
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
         Height          =   312
         Left            =   9840
         TabIndex        =   21
         Top             =   360
         Width           =   1692
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
         Height          =   312
         Left            =   9840
         TabIndex        =   22
         Top             =   720
         Width           =   1692
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
         Height          =   312
         Left            =   9840
         TabIndex        =   23
         Top             =   1080
         Width           =   1692
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
         Height          =   252
         Index           =   9
         Left            =   8160
         TabIndex        =   19
         Top             =   1056
         Width           =   972
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
         Height          =   252
         Index           =   8
         Left            =   8160
         TabIndex        =   18
         Top             =   696
         Width           =   1332
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
         Height          =   252
         Index           =   7
         Left            =   8160
         TabIndex        =   17
         Top             =   360
         Width           =   1452
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
         Height          =   252
         Index           =   6
         Left            =   8160
         TabIndex        =   16
         Top             =   0
         Width           =   972
      End
      Begin VB.Label lblLineas 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   216
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   2292
      End
      Begin VB.Label lblCantidad 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   216
         Left            =   0
         TabIndex        =   14
         Top             =   480
         Width           =   2292
      End
   End
   Begin VB.TextBox txtEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "..."
      ToolTipText     =   "Estado"
      Top             =   30
      Width           =   3768
   End
   Begin ComCtl3.CoolBar CoolBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   12000
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlb"
      MinHeight1      =   330
      Width1          =   3990
      NewRow1         =   0   'False
      Child2          =   "txtEstado"
      MinWidth2       =   1830
      MinHeight2      =   315
      Width2          =   1890
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   6
         Top             =   30
         Width           =   9825
         _ExtentX        =   17330
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
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repBoleta"
                     Text            =   "Boleta "
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repListadoGeneral"
                     Text            =   "Listado General"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3252
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   11892
      _Version        =   524288
      _ExtentX        =   20976
      _ExtentY        =   5736
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
      MaxCols         =   489
      ScrollBars      =   2
      SpreadDesigner  =   "frmCprCompraDirecta.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboCausa 
      Height          =   312
      Left            =   1440
      TabIndex        =   12
      Top             =   960
      Width           =   5652
      _Version        =   1441793
      _ExtentX        =   9975
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   3600
      TabIndex        =   10
      Top             =   480
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   5640
      TabIndex        =   11
      Top             =   480
      Width           =   1452
      _Version        =   1441793
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
   Begin XtremeSuiteControls.FlatEdit txtFactura 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   0
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   24
      Top             =   480
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
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
      Height          =   312
      Left            =   8880
      TabIndex        =   25
      Top             =   480
      Width           =   1452
      _Version        =   1441793
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCompra 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   0
      EndProperty
      Height          =   312
      Left            =   10320
      TabIndex        =   26
      Top             =   480
      Width           =   1452
      _Version        =   1441793
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtProvCod 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   0
      EndProperty
      Height          =   312
      Left            =   1440
      TabIndex        =   27
      Top             =   1320
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
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
   Begin XtremeSuiteControls.FlatEdit txtProvDesc 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   0
      EndProperty
      Height          =   312
      Left            =   3000
      TabIndex        =   28
      Top             =   1320
      Width           =   8772
      _Version        =   1441793
      _ExtentX        =   15473
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
      Height          =   552
      Left            =   1440
      TabIndex        =   29
      Top             =   1680
      Width           =   10332
      _Version        =   1441793
      _ExtentX        =   18224
      _ExtentY        =   974
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
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   315
      Left            =   8760
      TabIndex        =   30
      Top             =   960
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4043
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
      CustomFormat    =   "dd/mm/yyyy hh:mm:ss"
      Format          =   3
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
      Height          =   312
      Left            =   9480
      TabIndex        =   31
      Top             =   960
      Width           =   2292
      _Version        =   1441793
      _ExtentX        =   4043
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "# Orden/Compra"
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
      Left            =   7200
      TabIndex        =   7
      Top             =   480
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Causa"
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
      Index           =   10
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1092
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
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Height          =   252
      Index           =   4
      Left            =   8640
      TabIndex        =   2
      Top             =   960
      Width           =   732
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Factura"
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
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1092
   End
End
Attribute VB_Name = "frmCprCompraDirecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vCompra As String, vMascara As String

Private Sub cboCausa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProvCod.SetFocus
End Sub

Private Sub Form_Activate()
vModulo = 35
End Sub

Private Sub Form_Load()
Dim strSQL As String

On Error GoTo vError

Call sbCargaCboMonedas(cbo)
Call sbCprCboTiposOrden(cboCausa)


cboTipo.Clear
cboTipo.AddItem "Contado"
cboTipo.AddItem "Crédito"
cboTipo.Text = "Crédito"

 vModulo = 35
 vMascara = "0000000000"
 vEdita = True
 
 vGrid.AppearanceStyle = fxGridStyle
 
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


txtCompra = ""
vCompra = ""

txtFactura = ""
dtpFecha.Value = fxFechaServidor
txtFecha = Format(dtpFecha.Value, "yyyy/mm/dd hh:mm:ss")
dtpFecha.Visible = fxCprCambiaFecha(glogon.Usuario)


txtEstado = ""

txtNotas = ""


vGrid.MaxCols = 8
vGrid.MaxRows = 0
vGrid.MaxRows = 1

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
      txtFactura.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNotas.SetFocus
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

Private Sub sbConsulta(xCodigo As String, Optional xOrden As String = "")
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select E.*,rtrim(C.descripcion) as 'Causa_Desc', rtrim(C.Tipo_Orden) as 'Causa_Id'" _
       & ",P.descripcion as Proveedor,O.nota" _
       & " from cpr_ordenes O inner join cpr_Tipo_Orden C on O.Tipo_Orden = C.Tipo_Orden" _
       & " inner join cpr_compras E on O.cod_orden = E.cod_orden" _
       & " inner join cxp_proveedores P on E.cod_proveedor = P.cod_proveedor" _
       & " where E.cod_compra = '" & Format(xCodigo, vMascara) & "'"
       
If xOrden <> "" Then
  strSQL = strSQL & " and E.cod_orden = '" & Format(xOrden, vMascara) & "'"
Else
    If txtProvCod <> "" Then
       strSQL = strSQL & " and E.cod_proveedor = " & txtProvCod
    End If
End If

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!cod_orden
  txtCodigo = rs!cod_orden
  
  txtCompra = Format(xCodigo, vMascara)
  vCompra = Format(xCodigo, vMascara)
  
  txtFactura = rs!cod_Factura
  
  Call sbCboAsignaDato(cboCausa, rs!Causa_Desc, True, rs!Causa_Id)
  
  
  txtProvCod = rs!cod_Proveedor
  txtProvDesc = rs!Proveedor
  
  txtFecha = Format(rs!fecha, "yyyy/mm/dd hh:mm:ss")
  dtpFecha.Value = rs!fecha
  
  txtNotas = rs!nota & ""
  
  Select Case UCase(Trim(rs!forma_pago))
    Case "CO"
        cboTipo.Text = "Contado"
    Case "CR"
        cboTipo.Text = "Crédito"
  End Select
  
  Select Case rs!Estado
    Case "P"
      txtEstado = ">> PROCESADA <<"
      txtEstado.ForeColor = vbBlue
    Case "A"
      txtEstado = ">> ANULADA <<"
      txtEstado.ForeColor = vbRed
    Case "D"
      txtEstado = ">> PROCESADA CON DEVOLUCIONES <<"
      txtEstado.ForeColor = vbBlack
  End Select
  
  txtImpuestos = Format(rs!imp_ventas, "Standard")

  strSQL = "select D.cod_producto,P.descripcion,D.cantidad,D.cod_bodega,D.precio,isnull(D.descuento,0),D.imp_ventas,0" _
         & " from cpr_compras_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " where D.cod_factura = '" & rs!cod_Factura & "' and D.cod_proveedor = " & rs!cod_Proveedor _
         & " order by D.Linea"
  Call sbCargaGrid(vGrid, 8, strSQL)
  
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

vMensaje = fxInvVerificaLineaDetalle(vGrid, 3, "E", 1, 4)

If Len(txtFactura) = 0 Then vMensaje = vMensaje & vbCrLf & " - Número de Factura no es válido..."


If dtpFecha.Visible Then
    If Not fxInvPeriodos(dtpFecha.Value) Then vMensaje = vMensaje & vbCrLf & " - El periodo en el que desea realizar el movimiento se encuentra cerrado ..."
Else
    If Not fxInvPeriodos(fxFechaServidor) Then vMensaje = vMensaje & vbCrLf & " - El periodo en el que desea realizar el movimiento se encuentra cerrado ..."
End If


vError:

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, i As Integer, rs As New ADODB.Recordset
Dim curCantidad As Currency, vCodPro As String, vCodBodega As String
Dim curPrecio As Currency, curImpVentas As Currency, curImpConsumo As Currency
Dim vFecha As Date, curDescuento As Currency

On Error GoTo vError


'Solo se puede Insertar y no Editar
'01 - Guardar el registro en Ordenes y Detalles como procesado
'02 - Guardar el registro en las Entradas y Afectar Inventarios
If vEdita Then
   MsgBox "No se puede editar una Compra Guardada...", vbInformation
   Exit Sub
End If

Call sbCalculaTotales


'Consecutivo de la Orden
strSQL = "select isnull(max(cod_orden),0) + 1 as Ultimo from cpr_Ordenes"
Call OpenRecordSet(rs, strSQL)
  vCodigo = Format(rs!ultimo, vMascara)
rs.Close
txtCodigo = vCodigo

strSQL = "insert cpr_ordenes(cod_orden,Tipo_Orden,estado,nota,genera_user,genera_fecha" _
       & ",subtotal,descuento,imp_ventas,total,autoriza_fecha,autoriza_user,pin_autorizacion" _
       & ",pin_entrada,proceso,cod_proveedor) values('" & vCodigo & "','" & cboCausa.ItemData(cboCausa.ListIndex) & "','A','" _
       & txtNotas & "','" & glogon.Usuario & "',dbo.MyGetdate()," & CCur(txtSubTotal) _
       & "," & CCur(txtDescuento) & "," & CCur(txtImpuestos) & "," & CCur(txtTotal) _
       & ",dbo.MyGetdate(),'" & Mid(glogon.Usuario, 1, 23) & "',0,'','D','" & txtProvCod.Text & "')"
Call ConectionExecute(strSQL)

strSQL = "Insert CPR_ORDENES_PROCESO(COD_ORDEN, COD_PROVEEDOR, REGISTRO_FECHA, REGISTRO_USUARIO, COTIZA_FECHA,COTIZA_USUARIO" _
       & ",ADJUDICA_FECHA,ADJUDICA_USUARIO, NOTAS) " _
       & " values('" & vCodigo & "', " & txtProvCod.Text & ", dbo.Mygetdate(),'" & glogon.Usuario & "'" _
       & ", dbo.Mygetdate(),'" & glogon.Usuario & "', dbo.Mygetdate(),'" & glogon.Usuario & "','Compra Directa!')"
Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Orden Compra: " & vCodigo)

txtCodigo.Enabled = True

'Guardar Detalle de la Orden
strSQL = "delete cpr_ordenes_detalle" _
         & " where cod_orden = '" & vCodigo & "'"
'Call ConectionExecute(strSQL)

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.Col = 3
  curCantidad = CCur(IIf((vGrid.Text = ""), 0, vGrid.Text))
  
  vGrid.Col = 1
  If vGrid.Text <> "" And curCantidad > 0 Then
    strSQL = strSQL & Space(10) & "insert cpr_ordenes_detalle(linea,cod_orden,cod_producto,cantidad,estado" _
           & ",cantidad_despachada,precio,descuento,imp_ventas,imp_consumo) values(" & i & ",'" & vCodigo _
           & "','" & vGrid.Text & "',"
    vGrid.Col = 3
    strSQL = strSQL & CCur(vGrid.Text) & ",'D'," & CCur(vGrid.Text) & ","
    vGrid.Col = 5
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.Col = 6
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.Col = 7
    strSQL = strSQL & CCur(vGrid.Text) & ",0)"
    'Call ConectionExecute(strSQL)
  End If
Next i

'Registra Detalle
Call ConectionExecute(strSQL)

'*************** Guardar la Entrada a partir de aqui ****************

If dtpFecha.Visible Then
  vFecha = dtpFecha.Value
Else
  vFecha = fxFechaServidor
End If

'Consecutivo de la Orden
strSQL = "select isnull(max(cod_compra),0) + 1 as Ultimo from cpr_compras"
Call OpenRecordSet(rs, strSQL)
  vCompra = Format(rs!ultimo, vMascara)
rs.Close
txtCompra = vCompra

strSQL = "insert cpr_compras(estado,cod_factura,forma_pago,cod_proveedor,cod_compra,cod_orden,genera_user," _
       & "genera_fecha,fecha,sub_total,descuento,imp_ventas,imp_consumo,total,cxp_estado,asiento_estado)" _
       & " values('P','" & txtFactura & "','" & UCase(Mid(cboTipo.Text, 1, 2)) & "'," & txtProvCod & ",'" & vCompra & "','" & vCodigo _
       & "','" & glogon.Usuario & "','" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "','" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") _
       & "'," & CCur(txtSubTotal) & "," & CCur(txtDescuento) & "," & CCur(txtImpuestos) _
       & ",0," & CCur(txtTotal) & ",'" & IIf((UCase(Mid(cboTipo.Text, 1, 2)) = "CR"), "P", "G") & "','P')"
Call ConectionExecute(strSQL)

'Actualiza Saldo Proveedores
strSQL = "update cxp_proveedores set saldo = isnull(saldo,0) + " & CCur(txtTotal) _
       & " where cod_proveedor = " & txtProvCod
Call ConectionExecute(strSQL)
    
If UCase(Mid(cboTipo.Text, 1, 2)) = "CO" Then
   'Registrar Pagos al Contado, como pagados
  strSQL = "insert cxp_pagoProv(NPago,Cod_Proveedor,Cod_Factura,Fecha_Vencimiento,Monto,Frecuencia" _
         & ",Tipo_Transac,User_TrasLada,Fecha_Traslada,Tesoreria,Pago_Tercero,Apl_Cargo_Flotante" _
         & ",Pago_Anticipado,forma_pago, IMPORTE_DIVISA_REAL) values(1," & txtProvCod & ",'" & txtFactura & "',dbo.MyGetdate()," & CCur(txtTotal.Text) _
         & ",0,0,Null,Null,Null,'',0,0,'CO', " & CCur(txtTotal.Text) & ")"
  Call ConectionExecute(strSQL)
End If


Call Bitacora("Registra", "Compra Directa: " & vCompra)

txtCompra.Enabled = True


'Guardar Detalle de la Orden
strSQL = "delete cpr_compras_detalle" _
         & " where cod_factura = '" & txtFactura & "' and cod_proveedor = " & txtProvCod
'Call ConectionExecute(strSQL)

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  
  vGrid.Col = 3
  curCantidad = CCur(IIf((vGrid.Text = ""), 0, vGrid.Text))
  
  vGrid.Col = 1
  
  If vGrid.Text <> "" And curCantidad > 0 Then
    
    vGrid.Col = 1
    vCodPro = Trim(vGrid.Text)
    strSQL = strSQL & Space(10) & "insert cpr_compras_detalle(linea,cod_factura,cod_proveedor,cod_producto,cantidad,cod_bodega" _
           & ",precio,descuento,imp_ventas,imp_consumo) values(" & i & ",'" & txtFactura & "'," & txtProvCod & ",'" _
           & vGrid.Text & "'," & curCantidad & ",'"
    vGrid.Col = 4
    vCodBodega = Trim(vGrid.Text)
    strSQL = strSQL & vGrid.Text & "',"
    vGrid.Col = 5
    curPrecio = CCur(vGrid.Text)
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.Col = 6
    curDescuento = CCur(vGrid.Text)
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.Col = 7
    curImpVentas = CCur(vGrid.Text)
    curImpConsumo = 0
    strSQL = strSQL & CCur(vGrid.Text) & ",0)"
'    Call ConectionExecute(strSQL)
  
    'Actualizar Aqui el Inventario
    Call sbInvInventario(vCodPro, curCantidad, vCodBodega, vCompra, "Compra", vFecha _
             , curPrecio, curImpConsumo, curImpVentas, "E")
    
  End If
Next i

'Registra el Detalle
Call ConectionExecute(strSQL)




'Actualiza Costos de los articulos
strSQL = "exec spCRPActualizaCts '" & vCompra & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


'*********************************** fin
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

Select Case UCase(ButtonMenu.Key)
  Case "REPBOLETA"
     
     i = MsgBox("Desea visualizar solo la compra Actual", vbYesNo)
     If i = vbYes Then vSQL = "{cpr_compras.cod_compra} = '" & txtCompra & "'"
     
     Call sbInvReportes("COMPRA", "Boleta de Compra", "", vSQL)

  Case "REPLISTADOGENERAL"
    Call MuestraForms(frmCprReportesGenerales)

End Select


End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "E.cod_compra"
  gBusquedas.Orden = "E.cod_compra"
  gBusquedas.Consulta = "select E.cod_compra,E.cod_orden,E.cod_factura,P.descripcion as Proveedor" _
          & " from cpr_compras E inner join cxp_proveedores P on E.cod_proveedor = P.cod_proveedor"
  gBusquedas.Filtro = ""
  gBusquedas.Mascara = vMascara
  frmBusquedas.Show vbModal
  txtCompra = gBusquedas.Resultado
  If txtCompra <> "" Then Call sbConsulta(gBusquedas.Resultado, gBusquedas.Resultado2)
End If
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


Private Sub txtCompra_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCompra <> "" Then Call sbConsulta(txtCompra)
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "E.cod_compra"
  gBusquedas.Orden = "E.cod_compra"
  gBusquedas.Consulta = "select E.cod_compra,E.cod_orden,E.cod_factura,P.descripcion as Proveedor" _
          & " from cpr_compras E inner join cxp_proveedores P on E.cod_proveedor = P.cod_proveedor"
  gBusquedas.Filtro = ""
  gBusquedas.Mascara = vMascara
  frmBusquedas.Show vbModal
  txtCompra = gBusquedas.Resultado
  If txtCompra <> "" Then Call sbConsulta(gBusquedas.Resultado, gBusquedas.Resultado2)
End If

End Sub

Private Sub txtFactura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCausa.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "E.cod_compra"
  gBusquedas.Orden = "E.cod_compra"
  gBusquedas.Consulta = "select E.cod_compra,E.cod_orden,E.cod_factura,P.descripcion as Proveedor" _
          & " from cpr_compras E inner join cxp_proveedores P on E.cod_proveedor = P.cod_proveedor"
  gBusquedas.Filtro = ""
  gBusquedas.Mascara = vMascara
  frmBusquedas.Show vbModal
  txtCompra = gBusquedas.Resultado
  If txtCompra <> "" Then Call sbConsulta(gBusquedas.Resultado, gBusquedas.Resultado2)
End If

End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
End Sub

Private Sub txtProvCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProvDesc.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  txtProvCod = gBusquedas.Resultado
  txtProvDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtProvCod_LostFocus()
txtProvDesc = fxSIFCCodigos("D", txtProvCod, "proveedores")
End Sub

Private Sub txtProvDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  txtProvCod = gBusquedas.Resultado
  txtProvDesc = gBusquedas.Resultado2
End If

End Sub


Private Sub sbCalculaTotales()
Dim curSubTotal As Currency, curIV As Currency, curDescuento As Currency
Dim curTmpPrecio As Currency, curTmpIV As Currency, curTmpCant As Currency
Dim i As Integer, lng As Long, curTmpDesc As Currency
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
 vGrid.Col = 3
 If vGrid.Text <> "" Then
    curTmpCant = CCur(vGrid.Text)
    vGrid.Col = 5
    curTmpPrecio = CCur(vGrid.Text) * curTmpCant
    vGrid.Col = 6
    curTmpDesc = curTmpPrecio * (CCur(vGrid.Text) / 100)
    
    vGrid.Col = 7
    curTmpIV = (curTmpPrecio - curTmpDesc) * (CCur(vGrid.Text) / 100)

    curSubTotal = curSubTotal + curTmpPrecio
    curIV = curIV + curTmpIV
    curDescuento = curDescuento + curTmpDesc
    
    vGrid.Col = 8
    vGrid.Text = CStr(curTmpPrecio - curTmpDesc + curTmpIV)
    
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
  vGrid.Col = 1
  vGrid.Text = rs!Cod_Producto
  vGrid.Col = 2
  vGrid.Text = rs!Descripcion
  vGrid.Col = 5
  vGrid.Text = CStr(rs!costo_regular)
  vGrid.Col = 6
  vGrid.Text = "0"
  vGrid.Col = 7
  vGrid.Text = CStr(rs!impuesto_ventas)
End If
rs.Close


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(8) As Variant, x As Integer

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
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  Call sbConsultaArticulo(vGrid.ActiveRow, vGrid.ActiveCol, vGrid.Text)
End If

'Consular Bodegas
If vGrid.ActiveCol = 4 And KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_bodega"
   gBusquedas.Orden = "cod_bodega"
   gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
   gBusquedas.Filtro = " and permite_entradas = 1"
   frmBusquedas.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 4
   vGrid.Text = gBusquedas.Resultado
End If

'Consular Articulo
If vGrid.ActiveCol = 1 And KeyCode = vbKeyF4 Then
   frmBusquedaArticulos.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   vGrid.Text = gBusquedas.Resultado
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = vGrid.MaxCols
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To vGrid.MaxCols
        vGrid.Col = x
        vTemp(x) = vGrid.Text
     Next x

     vGrid.Row = lng
     For x = 1 To vGrid.MaxCols
       vGrid.Col = x
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
Dim curCantidad As Currency, curPrecio As Currency, curIV As Currency, curDescuento As Currency

On Error GoTo vError
'Calcula Total
Select Case vGrid.ActiveCol
  Case 3, 5, 6, 7
    vGrid.Row = vGrid.ActiveRow
    vGrid.Col = 3
    curCantidad = CCur(vGrid.Text)
    vGrid.Col = 5
    curPrecio = CCur(vGrid.Text) * curCantidad
    vGrid.Col = 6
    curDescuento = curPrecio * (CCur(vGrid.Text) / 100)
    vGrid.Col = 7
    curIV = (curPrecio - curDescuento) * (CCur(vGrid.Text) / 100)
    
    vGrid.Col = 8
    vGrid.Text = CStr(curPrecio + curIV - curDescuento)
   Call sbCalculaTotales
  Case Else 'No Aplica
End Select
vError:
End Sub


