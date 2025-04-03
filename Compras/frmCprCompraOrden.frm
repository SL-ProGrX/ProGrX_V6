VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCprComprasOrden 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras Con Orden de Compra"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   11685
   Begin VB.TextBox txtEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "..."
      Top             =   30
      Width           =   3972
   End
   Begin ComCtl3.CoolBar CoolBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   11685
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlb"
      MinHeight1      =   330
      Width1          =   3855
      NewRow1         =   0   'False
      Child2          =   "txtEstado"
      MinWidth2       =   2655
      MinHeight2      =   315
      Width2          =   2715
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   6
         Top             =   30
         Width           =   8685
         _ExtentX        =   15319
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
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   312
      Left            =   9240
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   2052
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
      Format          =   293470211
      CurrentDate     =   37791
   End
   Begin VB.TextBox txtFecha 
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
      Height          =   315
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   11535
      _Version        =   524288
      _ExtentX        =   20346
      _ExtentY        =   4471
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
      SpreadDesigner  =   "frmCprCompraOrden.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboCausa 
      Height          =   315
      Left            =   6840
      TabIndex        =   13
      Top             =   960
      Width           =   4815
      _Version        =   1572864
      _ExtentX        =   8493
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
      Height          =   315
      Left            =   4560
      TabIndex        =   14
      Top             =   480
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   315
      Left            =   3600
      TabIndex        =   15
      Top             =   960
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
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
      Left            =   9720
      TabIndex        =   16
      Top             =   4920
      Width           =   1695
      _Version        =   1572864
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
      Left            =   9720
      TabIndex        =   17
      Top             =   5280
      Width           =   1695
      _Version        =   1572864
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
      Left            =   9720
      TabIndex        =   18
      Top             =   5640
      Width           =   1695
      _Version        =   1572864
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
      Left            =   9720
      TabIndex        =   19
      Top             =   6000
      Width           =   1695
      _Version        =   1572864
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
      Height          =   315
      Left            =   1440
      TabIndex        =   24
      Top             =   1320
      Width           =   1575
      _Version        =   1572864
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
      Height          =   315
      Left            =   3000
      TabIndex        =   25
      Top             =   1320
      Width           =   8655
      _Version        =   1572864
      _ExtentX        =   15266
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtCompraNotas 
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
      Left            =   1440
      TabIndex        =   26
      Top             =   1680
      Width           =   10215
      _Version        =   1572864
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
      TabIndex        =   29
      Top             =   960
      Width           =   2175
      _Version        =   1572864
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
      Height          =   330
      Left            =   1440
      TabIndex        =   30
      Top             =   480
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
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
      Height          =   330
      Left            =   3000
      TabIndex        =   31
      Top             =   480
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
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
      Height          =   675
      Left            =   120
      TabIndex        =   32
      Top             =   5640
      Width           =   5895
      _Version        =   1572864
      _ExtentX        =   10398
      _ExtentY        =   1191
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
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   28
      Top             =   1440
      Width           =   1215
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
      Index           =   2
      Left            =   240
      TabIndex        =   27
      Top             =   1800
      Width           =   1095
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
      Left            =   8040
      TabIndex        =   23
      Top             =   4920
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
      Left            =   8040
      TabIndex        =   22
      Top             =   5280
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
      Left            =   8040
      TabIndex        =   21
      Top             =   5610
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
      Left            =   8040
      TabIndex        =   20
      Top             =   5970
      Width           =   975
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
      Left            =   120
      TabIndex        =   10
      Top             =   5040
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
      Left            =   2520
      TabIndex        =   9
      Top             =   5040
      Width           =   2292
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Notas de la Orden de Compra"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   5325
      Width           =   5895
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   10
      Left            =   6000
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
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
      Left            =   8280
      TabIndex        =   3
      Top             =   480
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Factura"
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
      TabIndex        =   1
      Top             =   960
      Width           =   852
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   9600
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "No. Orden"
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
      TabIndex        =   0
      Top             =   480
      Width           =   852
   End
End
Attribute VB_Name = "frmCprComprasOrden"
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

On Error GoTo vError
 vModulo = 35
 vEdita = True
 vMascara = "0000000000"
 
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
vCompra = ""
txtCompra = ""

Call sbCargaCboMonedas(cbo)

dtpFecha.Value = fxFechaServidor
txtFecha = Format(dtpFecha.Value, "yyyy/mm/dd hh:mm:ss")
dtpFecha.Visible = fxCprCambiaFecha(glogon.Usuario)

txtEstado = ""
txtNotas = ""
txtCompraNotas = ""

Call sbCprCboTiposOrden(cboCausa)

vGrid.MaxRows = 1
vGrid.MaxCols = 8
For i = 1 To vGrid.MaxCols
  vGrid.Col = i
  vGrid.Text = ""
Next

txtSubTotal = 0
txtDescuento = 0
txtImpuestos = 0
txtTotal = 0

txtCodigo.Enabled = True
txtCompra.Enabled = True

txtFactura = ""
txtProvCod = ""
txtProvDesc = ""

cboTipo.Clear
cboTipo.AddItem "Contado"
cboTipo.AddItem "Crédito"
cboTipo.Text = "Crédito"

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = True
      txtCodigo.SetFocus
      txtCompra.Enabled = False
      vGrid.Enabled = True
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtFactura.SetFocus
      vGrid.Enabled = False
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
        Call sbConsultaOrden(vCodigo)
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

Private Function sbDesActProdCancelados()
Dim lng As Long

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.Col = 3
 If vGrid.Text <> "" Then
   If CCur(vGrid.Text) <= 0 Then
      vGrid.Col = 1
      vGrid.Text = ""
   End If
 End If
Next lng

End Function

Private Sub sbConsultaOrden(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select O.*,rtrim(C.tipo_orden) as 'Causa_ID', Rtrim(C.descripcion) as 'Causa_Desc', Prov.Descripcion as 'Proveedor'" _
       & " from cpr_ordenes O inner join cpr_Tipo_Orden C on O.tipo_orden = C.tipo_orden" _
       & "  inner join CxP_Proveedores Prov on O.cod_Proveedor = Prov.cod_proveedor" _
       & " where O.cod_orden = '" & Format(xCodigo, vMascara) _
       & "' and O.estado = 'A' and O.Proceso in('A','D','X')"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
'  Call sbToolBar(tlb, "edicion")
  vEdita = False 'False
  Call sbLimpiaPantalla
  
  vCodigo = rs!cod_orden
  txtCodigo = rs!cod_orden
  
  txtProvCod.Text = rs!cod_Proveedor
  txtProvDesc.Text = rs!Proveedor
  
  txtNotas.Text = rs!nota
  
  Call sbCboAsignaDato(cboCausa, rs!Causa_Desc, True, rs!Causa_Id)
  
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
  txtNotas = rs!nota & ""
 'Tengo que conservar visibles aquellos que ya fueron despachados para conservar el consecutivo
 'de la linea del detalle
  strSQL = "select D.cod_producto,P.descripcion,(D.cantidad - isnull(D.cantidad_despachada,0)) as Cantidad" _
         & ",'',D.precio,isnull(D.descuento,0) as Descuento,D.imp_ventas,0 as Total" _
         & " from cpr_ordenes_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " where D.cod_orden = '" & rs!cod_orden & "' order by D.Linea"
'         & "' and (D.cantidad - isnull(D.cantidad_despachada,0)) > 0" _
'         & " order by D.Linea"
  Call sbCargaGrid(vGrid, 8, strSQL)
  Call sbDesActProdCancelados
  
  
  Call sbCalculaTotales
  
Else
  MsgBox "No se encontró la orden para procesar (Revisar su existencia o si el estado o proceso estan autorizados)", vbInformation
End If

rs.Close
Call RefrescaTags(Me)
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsulta(xCodigo As String, Optional xOrden As String = "")
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select E.*,(rtrim(C.Tipo_Orden) + ' - ' + C.descripcion) as Causa" _
       & ",P.descripcion as Proveedor,O.nota,E.notas as CompraNotas" _
       & " from cpr_ordenes O inner join cpr_Tipo_Orden C on O.Tipo_Orden = C.Tipo_Orden" _
       & " inner join cpr_Compras E on O.cod_orden = E.cod_orden" _
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
  txtCompraNotas = rs!CompraNotas & ""
  
  cboCausa.Text = Trim(rs!Causa)
  Select Case UCase(Trim(rs!forma_pago))
    Case "CO"
        cboTipo.Text = "Contado"
    Case "CR"
        cboTipo.Text = "Crédito"
  End Select
  txtProvCod = rs!cod_Proveedor
  txtProvDesc = rs!Proveedor
  
  If rs!Estado = "P" Then
    txtEstado = "Procesada"
  Else
    txtEstado = "Anulada"
  End If
  
  txtFecha = Format(rs!genera_fecha, "yyyy/mm/dd hh:mm:ss")
  dtpFecha.Value = rs!genera_fecha
  
  txtNotas = rs!nota & ""
  
  txtImpuestos = Format(rs!imp_ventas, "Standard")

  strSQL = "select D.cod_producto,P.descripcion,D.cantidad,D.cod_bodega,D.precio,isnull(D.descuento,0),D.imp_ventas,0 as Total" _
         & " from cpr_Compras_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
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

If Len(txtFactura) = 0 Then vMensaje = vMensaje & vbCrLf & " - Número de Factura no es válido..."

vMensaje = fxInvVerificaLineaDetalle(vGrid, 3, "E", 1, 4)

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
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer, vPin As String
Dim curCantidad As Currency, vCodPro As String, vCodBodega As String
Dim curPrecio As Currency, curImpVentas As Currency, curImpConsumo As Currency
Dim vFecha As Date, curDescuento As Currency

On Error GoTo vError

If vEdita Then
   MsgBox "No se puede editar una Compra Guardada...", vbInformation
   Exit Sub
End If

strSQL = "select pin_autorizacion from cpr_ordenes where cod_orden = '" & vCodigo & "'"
Call OpenRecordSet(rs, strSQL)
If rs!pin_autorizacion = 1 Then
 rs.Close
 vPin = InputBox("Digite el Pin de Autorización de Compra de Mercadería : ", "Pin de Compra ?")
 strSQL = "select isnull(count(*),0) as Existe from cpr_ordenes " _
        & " where cod_orden = '" & vCodigo _
        & "' and pin_entrada = '" & vPin & "'"
 Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "El Pin de Compra suministrado no es correcto...", vbExclamation
    Exit Sub
 End If
End If
rs.Close


If dtpFecha.Visible Then
  vFecha = dtpFecha.Value
Else
  vFecha = fxFechaServidor
End If


'Consecutivo de la Compra
strSQL = "select isnull(max(cod_orden),0) + 1 as Ultimo from cpr_compras"
Call OpenRecordSet(rs, strSQL)
  vCompra = Format(rs!ultimo, vMascara)
rs.Close
txtCompra = vCompra

Call sbCalculaTotales

txtCodigo = vCodigo
vCodigo = txtCodigo

strSQL = "insert cpr_Compras(estado,cod_factura,forma_pago,cod_proveedor,cod_compra,cod_orden,genera_user," _
       & "genera_fecha,fecha,sub_total,descuento,imp_ventas,imp_consumo,total,cxp_estado,asiento_estado,notas)" _
       & " values('P','" & txtFactura & "','" & UCase(Mid(cboTipo.Text, 1, 2)) & "'," & txtProvCod & ",'" _
       & vCompra & "','" & vCodigo & "','" & glogon.Usuario & "','" _
       & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "','" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") _
       & "'," & CCur(txtSubTotal) & "," & CCur(txtDescuento) & "," & CCur(txtImpuestos) _
       & ",0," & CCur(txtTotal) & ",'" & IIf((UCase(Mid(cboTipo.Text, 1, 2)) = "CR"), "P", "G") & "','P','" & txtCompraNotas & "')"
Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Compra: " & vCompra)

'Actualiza Saldo Proveedores / Si es A credito (CR)
If UCase(Mid(cboTipo.Text, 1, 2)) = "CR" Then
    strSQL = "update cxp_proveedores set saldo = isnull(saldo,0) + " & CCur(txtTotal) _
           & " where cod_proveedor = " & txtProvCod
    Call ConectionExecute(strSQL)
Else
   'Registrar Pagos al Contado, como pagados
  strSQL = "insert cxp_pagoProv(NPago,Cod_Proveedor,Cod_Factura,Fecha_Vencimiento,Monto,Frecuencia" _
         & ",Tipo_Transac,User_TrasLada,Fecha_Traslada,Tesoreria,Pago_Tercero,Apl_Cargo_Flotante" _
         & ",Pago_Anticipado,forma_pago, IMPORTE_DIVISA_REAL) values(1," & txtProvCod & ",'" & txtFactura & "',dbo.MyGetdate()," & CCur(txtTotal) _
         & ",0,0,'" & glogon.Usuario & "',dbo.MyGetdate(),0,'',0,0,'CO', " & CCur(txtTotal) & ")"
  Call ConectionExecute(strSQL)
End If

txtCompra.Enabled = True


'Guardar Detalle de la Orden
strSQL = "delete cpr_Compras_detalle" _
         & " where cod_factura = '" & txtFactura & "' and cod_proveedor = " & txtProvCod

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  
  vGrid.Col = 3
  curCantidad = CCur(IIf((vGrid.Text = ""), 0, vGrid.Text))
  
  vGrid.Col = 1
  
  If vGrid.Text <> "" And curCantidad > 0 Then
    
    vGrid.Col = 1
    vCodPro = Trim(vGrid.Text)
    strSQL = strSQL & Space(10) & "insert cpr_Compras_detalle(linea,cod_factura,cod_proveedor,cod_producto,cantidad,cod_bodega" _
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
  
    'Actualizar Aqui el Inventario y la Orden de Compra
    vGrid.Col = 3
    strSQL = strSQL & Space(10) & "update cpr_ordenes_detalle set cantidad_despachada = isnull(cantidad_despachada,0) + " _
           & CCur(vGrid.Text) & " where linea = " & i & " and cod_orden = '" & vCodigo & "'"
'    Call ConectionExecute(strSQL)
    
    Call sbInvInventario(vCodPro, curCantidad, vCodBodega, CStr(vCompra), "Compra", vFecha _
             , curPrecio, curImpConsumo, curImpVentas, "E")
    
  End If
Next i


'Registra Detalle
Call ConectionExecute(strSQL)

'Indica si la Orden fue Total/Parcialmente despachada
Call sbCprOrdenesDespacho(vCodigo)

'Actualiza Costos de los articulos
strSQL = "exec spCRPActualizaCts '" & vCompra & "', '" & glogon.Usuario & "'"
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

Select Case UCase(ButtonMenu.Key)
  Case "REPBOLETA"
     
     i = MsgBox("Desea visualizar solo la Compra Actual", vbYesNo)
     If i = vbYes Then vSQL = "{cpr_Compras.cod_compra} = '" & txtCompra & "' and {CPR_COMPRAS.COD_FACTURA} = '" & txtFactura.Text & "'"
     
     Call sbInvReportes("Compra", "Boleta de Compra", "", vSQL)

  Case "REPLISTADOGENERAL"
    Call MuestraForms(frmCprReportesGenerales)

End Select

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFactura.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_orden"
  gBusquedas.Orden = "cod_orden"
  gBusquedas.Consulta = "select cod_orden,genera_user,nota from cpr_ordenes"
  gBusquedas.Filtro = " and Estado in('A') and Proceso in('A','X')"
  gBusquedas.Mascara = vMascara
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsultaOrden(gBusquedas.Resultado)
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" Then Call sbConsultaOrden(txtCodigo)
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
  If txtCompra <> "" Then
   Call sbConsulta(txtCompra)
  Else
    txtFactura.SetFocus
  End If
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_compra"
  gBusquedas.Orden = "cod_compra"
  gBusquedas.Consulta = "select E.cod_compra,E.cod_orden,E.cod_factura,P.descripcion as Proveedor" _
          & " from cpr_Compras E inner join cxp_proveedores P on E.cod_proveedor = P.cod_proveedor"
  gBusquedas.Filtro = ""
  gBusquedas.Mascara = vMascara
  frmBusquedas.Show vbModal
  txtCompra = gBusquedas.Resultado
  If txtCompra <> "" Then Call sbConsulta(gBusquedas.Resultado, gBusquedas.Resultado2)
End If

End Sub

Private Sub txtCompraNotas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
vError:
End Sub

Private Sub txtFactura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProvCod.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_factura"
  gBusquedas.Orden = "cod_factura"
  gBusquedas.Consulta = "select E.cod_compra,E.cod_orden,E.cod_factura,P.descripcion as Proveedor" _
          & " from cpr_Compras E inner join cxp_proveedores P on E.cod_proveedor = P.cod_proveedor"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCompra = gBusquedas.Resultado
  If txtCompra <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
vError:
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
   If curTmpCant > 0 Then
    vGrid.Col = 5
    curTmpPrecio = CCur(vGrid.Text) * curTmpCant
    vGrid.Col = 6
    curTmpDesc = curTmpPrecio * (CCur(vGrid.Text) / 100)
    vGrid.Col = 7
    curTmpIV = (curTmpPrecio - curTmpDesc) * (CCur(vGrid.Text) / 100)
    
    
    curSubTotal = curSubTotal + curTmpPrecio
    curDescuento = curDescuento + curTmpDesc
    curIV = curIV + curTmpIV
   
    
    vGrid.Col = 8
    vGrid.Text = CStr(curTmpPrecio - curTmpDesc + curTmpIV)
    
    curCantidad = curCantidad + curTmpCant
    iLineas = iLineas + 1
   
   End If
 End If
Next lng

txtSubTotal = Format(curSubTotal, "Standard")
txtImpuestos = Format(curIV, "Standard")
txtDescuento = Format(curDescuento, "Standard")
txtTotal = Format(curSubTotal + curIV - CCur(txtDescuento), "Standard")

lblLineas.Caption = "Líneas   : " & iLineas
lblCantidad.Caption = "Cantidad : " & Format(curCantidad, "Standard")


End Sub

Private Sub txtProvCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProvDesc.SetFocus
'If KeyCode = vbKeyF4 Then
'  gBusquedas.Resultado = ""
'  gBusquedas.Resultado2 = ""
'  gBusquedas.Columna = "cod_proveedor"
'  gBusquedas.Orden = "cod_proveedor"
'  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
'  gBusquedas.Filtro = " and estado = 'A'"
'  frmBusquedas.Show vbModal
'  txtProvCod = gBusquedas.Resultado
'  txtProvDesc = gBusquedas.Resultado2
'End If
End Sub

Private Sub txtProvCod_LostFocus()
txtProvDesc = fxSIFCCodigos("D", txtProvCod, "proveedores")
End Sub

Private Sub txtProvDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCompraNotas.SetFocus
'If KeyCode = vbKeyF4 Then
'  gBusquedas.Resultado = ""
'  gBusquedas.Resultado2 = ""
'  gBusquedas.Columna = "descripcion"
'  gBusquedas.Orden = "descripcion"
'  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
'  gBusquedas.Filtro = " and estado = 'A'"
'  frmBusquedas.Show vbModal
'  txtProvCod = gBusquedas.Resultado
'  txtProvDesc = gBusquedas.Resultado2
'End If

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

'Consular Articulo
'If vGrid.ActiveCol = 1 And KeyCode = vbKeyF4 Then
'   frmBusquedaArticulos.Show vbModal
'   vGrid.Row = vGrid.ActiveRow
'   vGrid.Col = 1
'   vGrid.Text = gBusquedas.Resultado
'End If

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

'
''Borrar una linea
'If KeyCode = vbKeyDelete Then
'  vGrid.Row = vGrid.ActiveRow
'  vGrid.Col = 7
'  For lng = vGrid.ActiveRow To vGrid.MaxRows
'     vGrid.Row = lng + 1
'     For x = 1 To 7
'        vGrid.Col = x
'        vTemp(x) = vGrid.Text
'     Next x
'
'     vGrid.Row = lng
'     For x = 1 To 7
'       vGrid.Col = x
'       vGrid.Text = vTemp(x)
'     Next x
'  Next lng
'  vGrid.MaxRows = vGrid.MaxRows - 1
'  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
'  Call sbCalculaTotales
'End If


End Sub



Private Sub vGrid_KeyPress(KeyAscii As Integer)
Dim curCantidad As Currency, curPrecio As Currency, curIV As Currency
Dim curDescuento As Currency

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
    vGrid.Text = curPrecio - curDescuento + curIV
   Call sbCalculaTotales
  Case Else 'No Aplica
End Select
vError:
End Sub


