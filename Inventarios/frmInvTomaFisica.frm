VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmInvTomaFisica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Toma Física del Inventario"
   ClientHeight    =   7476
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11712
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7476
   ScaleWidth      =   11712
   Begin XtremeSuiteControls.FlatEdit txtFechaEjecucion 
      Height          =   312
      Left            =   3240
      TabIndex        =   32
      Top             =   6720
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5052
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   12372
      _Version        =   1245187
      _ExtentX        =   21823
      _ExtentY        =   8911
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
      Item(0).Caption =   "Toma Física"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "txtBodegaDes"
      Item(0).Control(1)=   "txtBodegaCod"
      Item(0).Control(2)=   "txtNotas"
      Item(0).Control(3)=   "vGrid"
      Item(0).Control(4)=   "Label1(4)"
      Item(0).Control(5)=   "Label1(2)"
      Item(1).Caption =   "Carga Archivo"
      Item(1).ControlCount=   0
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3612
         Left            =   0
         TabIndex        =   15
         Top             =   1320
         Width           =   11532
         _Version        =   524288
         _ExtentX        =   20341
         _ExtentY        =   6371
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
         MaxCols         =   484
         ScrollBars      =   2
         SpreadDesigner  =   "frmInvTomaFisica.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtBodegaCod 
         Height          =   312
         Left            =   2640
         TabIndex        =   23
         Top             =   360
         Width           =   1452
         _Version        =   1245187
         _ExtentX        =   2561
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtBodegaDes 
         Height          =   312
         Left            =   4080
         TabIndex        =   24
         Top             =   360
         Width           =   6372
         _Version        =   1245187
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   552
         Left            =   2640
         TabIndex        =   25
         Top             =   720
         Width           =   7812
         _Version        =   1245187
         _ExtentX        =   13779
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
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
         Index           =   2
         Left            =   1560
         TabIndex        =   17
         Top             =   720
         Width           =   612
      End
      Begin VB.Label Label1 
         Caption         =   "Bodega"
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
         Left            =   1560
         TabIndex        =   16
         Top             =   360
         Width           =   612
      End
   End
   Begin ComCtl3.CoolBar CoolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11712
      _ExtentX        =   20659
      _ExtentY        =   635
      BandCount       =   2
      _CBWidth        =   11712
      _CBHeight       =   360
      _Version        =   "6.7.9816"
      Child1          =   "tlb"
      MinHeight1      =   264
      Width1          =   3336
      NewRow1         =   0   'False
      Child2          =   "tlbProcesos"
      MinHeight2      =   312
      Width2          =   4224
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlb 
         Height          =   264
         Left            =   132
         TabIndex        =   8
         Top             =   48
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   466
         ButtonWidth     =   487
         ButtonHeight    =   466
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
                     Key             =   "Sep1"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repListadoGeneral"
                     Text            =   "Listado General"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repBoletaV"
                     Text            =   "Boleta Valorización"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbProcesos 
         Height          =   312
         Left            =   3492
         TabIndex        =   7
         Top             =   24
         Width           =   8148
         _ExtentX        =   14372
         _ExtentY        =   550
         ButtonWidth     =   2117
         ButtonHeight    =   550
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Procesos"
               Key             =   "Procesos"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Inv.Lógico"
               Key             =   "logico"
               Object.ToolTipText     =   "Actualiza Inventario Lógico"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Compara"
               Key             =   "compara"
               Object.ToolTipText     =   "Llena TF con faltantes en Bodega"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6240
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvTomaFisica.frx":07B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvTomaFisica.frx":108C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvTomaFisica.frx":13A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvTomaFisica.frx":16C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   5760
      TabIndex        =   10
      Top             =   600
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   432
      Left            =   1320
      TabIndex        =   18
      Top             =   600
      Width           =   2172
      _Version        =   1245187
      _ExtentX        =   3831
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
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
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   432
      Left            =   3480
      TabIndex        =   19
      Top             =   600
      Width           =   2172
      _Version        =   1245187
      _ExtentX        =   3831
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   8640
      TabIndex        =   21
      Top             =   720
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   9960
      TabIndex        =   22
      Top             =   720
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
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
   Begin XtremeSuiteControls.FlatEdit txtUdLogica 
      Height          =   312
      Left            =   9840
      TabIndex        =   26
      Top             =   6360
      Width           =   1452
      _Version        =   1245187
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUDDif 
      Height          =   312
      Left            =   9840
      TabIndex        =   27
      Top             =   7080
      Width           =   1452
      _Version        =   1245187
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUdFisica 
      Height          =   312
      Left            =   9840
      TabIndex        =   28
      Top             =   6720
      Width           =   1452
      _Version        =   1245187
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtHechoPor 
      Height          =   312
      Left            =   1680
      TabIndex        =   29
      Top             =   6360
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtFechaCrea 
      Height          =   312
      Left            =   3240
      TabIndex        =   30
      Top             =   6360
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtProcesadoPor 
      Height          =   312
      Left            =   1680
      TabIndex        =   31
      Top             =   6720
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtSalida 
      Height          =   312
      Left            =   6840
      TabIndex        =   33
      Top             =   6720
      Width           =   972
      _Version        =   1245187
      _ExtentX        =   1714
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEntrada 
      Height          =   312
      Left            =   6840
      TabIndex        =   34
      Top             =   6360
      Width           =   972
      _Version        =   1245187
      _ExtentX        =   1714
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtProveedor 
      Height          =   312
      Left            =   6840
      TabIndex        =   35
      Top             =   7080
      Width           =   972
      _Version        =   1245187
      _ExtentX        =   1714
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   20
      Top             =   600
      Width           =   1092
      _Version        =   1245187
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Boleta: "
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Diferencia Libros"
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
      Index           =   7
      Left            =   8280
      TabIndex        =   13
      Top             =   7080
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "UD´s Físicas"
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
      Index           =   3
      Left            =   8280
      TabIndex        =   12
      Top             =   6720
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "UD´s Libros"
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
      Index           =   1
      Left            =   8280
      TabIndex        =   11
      Top             =   6360
      Width           =   1572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aplicación:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   9960
      TabIndex        =   9
      Top             =   480
      Width           =   1572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   8640
      TabIndex        =   5
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cod.Prov."
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
      Index           =   6
      Left            =   5760
      TabIndex        =   4
      Top             =   7080
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Salida"
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
      Index           =   5
      Left            =   5760
      TabIndex        =   3
      Top             =   6720
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Entrada"
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
      Index           =   4
      Left            =   5760
      TabIndex        =   2
      Top             =   6360
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Procesado por"
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
      Left            =   120
      TabIndex        =   1
      Top             =   6720
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Hecho por"
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   1572
   End
End
Attribute VB_Name = "frmInvTomaFisica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vScroll As Boolean

Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtBodegaCod.SetFocus
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCorte.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCodigo = "" Then txtCodigo = 0

If vScroll Then
    strSQL = "select Top 1 Consecutivo from pv_invTomaFisica"

    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where consecutivo > " & txtCodigo & " order by Consecutivo asc"
    Else
       strSQL = strSQL & " where consecutivo < " & txtCodigo & " order by Consecutivo desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!Consecutivo)
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
 vModulo = 32
End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 32
 vGrid.AppearanceStyle = fxGridStyle

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True

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
Dim i As Integer

vCodigo = 0
txtCodigo = ""

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

txtBodegaCod = ""
txtBodegaDes = ""
txtNotas = ""

vGrid.MaxRows = 1
vGrid.MaxCols = 7
For i = 1 To vGrid.MaxCols
  vGrid.col = i
  vGrid.Text = ""
Next

txtHechoPor = ""
txtFechaCrea = ""
txtProcesadoPor = ""
txtFechaEjecucion = ""
txtEntrada = ""
txtSalida = ""
txtProveedor = ""

txtProveedor.ToolTipText = ""
txtEntrada.ToolTipText = ""
txtSalida.ToolTipText = ""

tlbProcesos.Buttons(1).Enabled = False
tlbProcesos.Buttons(2).Enabled = False
tlbProcesos.Buttons(3).Enabled = False

txtCodigo.Enabled = True


txtUDDif.Text = ""
txtUdFisica.Text = ""
txtUdLogica.Text = ""

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      dtpInicio.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      dtpInicio.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select consecutivo,user_crear,cod_bodega,fecha_inicio from pv_InvTomaFisica"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       Call sbConsulta(txtCodigo)

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbCalculaExistencias()
Dim i As Integer, vExLogica As Currency, vExFisica As Currency, vExDif As Currency

vExLogica = 0
vExFisica = 0
vExDif = 0

With vGrid
 For i = 1 To .MaxRows
     .Row = i
     .col = 5
     If IsNumeric(vGrid.Text) Then vExLogica = vExLogica + CCur(vGrid.Text)
     .col = 6
     If IsNumeric(vGrid.Text) Then vExFisica = vExFisica + CCur(vGrid.Text)
     .col = 7
     If IsNumeric(vGrid.Text) Then vExDif = vExDif + CCur(vGrid.Text)
 Next i
End With

txtUDDif.Text = Format(vExDif, "Standard")
txtUdFisica.Text = Format(vExFisica, "Standard")
txtUdLogica.Text = Format(vExLogica, "Standard")

If CCur(txtUDDif.Text) < 0 Then
   txtUDDif.ForeColor = vbRed
Else
   txtUDDif.ForeColor = vbBlack
End If


End Sub

Private Sub sbConsulta(lngCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select I.*,B.descripcion as Bodega" _
       & " from pv_invTomaFisica I inner join pv_Bodegas B on I.cod_bodega = B.cod_bodega" _
       & " where I.consecutivo = " & lngCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!Consecutivo
  txtCodigo = rs!Consecutivo
  
  Select Case rs!Estado
    Case "S"
      txtEstado = "Solicitada"
    Case "P"
      txtEstado = "Procesada"
    Case Else
      txtEstado = "No Identificada"
  End Select

    dtpInicio.Value = rs!FECHA_INICIO
    dtpCorte.Value = rs!fecha_corte
    
    txtBodegaCod.Text = rs!cod_bodega
    txtBodegaDes.Text = rs!Bodega
    txtNotas.Text = rs!notas
    
    txtHechoPor.Text = rs!user_crea
    txtFechaCrea.Text = Format(rs!fecha_crea, "yyyy/mm/dd")
 
 If rs!Estado = "P" Then
    tlbProcesos.Buttons(1).Enabled = False
    tlbProcesos.Buttons(2).Enabled = False
    tlbProcesos.Buttons(3).Enabled = False
    
    txtProcesadoPor.Text = rs!user_aplica
    txtFechaEjecucion.Text = Format(rs!fecha_aplica, "yyyy/mm/dd")
    txtEntrada = rs!causa_entrada
    txtSalida.Text = rs!causa_salida
    txtProveedor.Text = rs!cod_proveedor_entrada & ""
    txtEntrada.ToolTipText = rs!cod_entradag
    txtSalida.ToolTipText = rs!cod_salidag
  Else
    tlbProcesos.Buttons(1).Enabled = True
    tlbProcesos.Buttons(2).Enabled = True
    tlbProcesos.Buttons(3).Enabled = True
  End If


  strSQL = "select D.cod_producto,P.descripcion,P.tipo_producto,D.ubicacion,D.existencia_logica" _
         & ",D.existencia_fisica,(D.existencia_logica - D.existencia_fisica) as Diferencia" _
         & " from pv_invtf_Detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " where D.cod_bodega = '" & rs!cod_bodega & "' and D.consecutivo = " & rs!Consecutivo
 
  Call sbCargaGrid(vGrid, 7, strSQL)
  
  Call sbCalculaExistencias
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


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, i As Integer
Dim rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
   strSQL = "update pv_InvTomaFisica set notas = '" & Trim(txtNotas) _
         & "',fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
         & "',fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") _
         & "',cod_bodega = '" & Trim(txtBodegaCod) _
         & "' where consecutivo = " & vCodigo
  
  If Mid(txtEstado, 1, 1) = "S" Then
      Call ConectionExecute(strSQL)
  Else
      MsgBox "No puede Modificar esta Toma Fisica, ya que no se encuentra Solicitada...", vbExclamation
      Exit Sub
  End If
  Call Bitacora("Modifica", "Toma Fisica Cod: " & vCodigo)

Else
   strSQL = "select isnull(max(consecutivo),0) as Ultimo from pv_InvTomaFisica"
   Call OpenRecordSet(rs, strSQL)
    vCodigo = rs!ultimo + 1
    txtCodigo = vCodigo
   rs.Close
   
   strSQL = "insert pv_InvTomaFisica(consecutivo,cod_bodega,fecha_inicio,fecha_corte,estado" _
          & ",fecha_crea,user_crea,notas) values(" & vCodigo & ",'" & Trim(txtBodegaCod) _
          & "','" & Format(dtpInicio.Value, "yyyy/mm/dd") & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") _
          & "','S',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtNotas) & "')"
   Call ConectionExecute(strSQL)

    Call Bitacora("Registra", "Toma Fisica Cod: " & vCodigo)
    txtCodigo.Enabled = True

End If

'Guardar Detalle de la toma fisica
strSQL = "delete pv_InvTF_detalle where consecutivo = " & vCodigo
Call ConectionExecute(strSQL)

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.col = 1
  If vGrid.Text <> "" Then
    strSQL = strSQL & Space(10) & "insert pv_InvTF_detalle(consecutivo,cod_bodega,cod_producto,ubicacion" _
           & ",existencia_logica,existencia_fisica) values(" & vCodigo & ",'" & Trim(txtBodegaCod) _
           & "','" & vGrid.Text & "','"
    vGrid.col = 4
    strSQL = strSQL & vGrid.Text & "',"
    vGrid.col = 5
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, CCur(vGrid.Text)) & ","
    vGrid.col = 6
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, CCur(vGrid.Text)) & ")"
  End If

  If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
  End If

Next i

'Ultimo Lote
If Len(strSQL) > 0 Then
      Call ConectionExecute(strSQL)
      strSQL = ""
End If


vEdita = True
txtEstado = "Solicitada"
tlbProcesos.Buttons(1).Enabled = True
tlbProcesos.Buttons(2).Enabled = True
tlbProcesos.Buttons(3).Enabled = True


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
  If Mid(txtEstado, 1, 1) = "S" Then
    strSQL = "delete pv_InvTF_Detalle where consecutivo = " & vCodigo
    strSQL = strSQL & Space(10) & "delete pv_InvTomaFisica where consecutivo = " & vCodigo
    Call ConectionExecute(strSQL)
    
    MsgBox "Toma Fisica # " & vCodigo & " borrada satisfactoriamente...", vbInformation
  End If
  Call Bitacora("Elimina", "Toma Fisica Cod: " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim i As Byte, vSQL As String, vOrden As String

vSQL = ""

If ButtonMenu.Key = "repListadoGeneral" Then
     Call sbInvReportes("InvTomaFisicaListado", "Toma Física", "Listado General", vSQL)
     Exit Sub
End If

i = MsgBox("Desea Ordenar Líneas de Detalle por Descripcion del Producto?", vbYesNo)

vSQL = "{PV_INVTOMAFISICA.CONSECUTIVO} = " & txtCodigo


If i = vbYes Then
     vOrden = "+{PV_Productos.Descripcion}"
End If

i = MsgBox("Desea ver solo líneas con diferencias", vbYesNo)
If i = vbYes Then
  If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
  vSQL = vSQL & "{PV_INVTF_DETALLE.EXISTENCIA_LOGICA} <> {PV_INVTF_DETALLE.EXISTENCIA_FISICA}"
End If

Select Case ButtonMenu.Key
  Case "repBoleta"
     Call sbInvReportes("InvTomaFisicaBoleta", "Toma Física", "Boleta", vSQL, vOrden)
  
  Case "repBoletaV"
     Call sbInvReportes("InvTomaFisicaBoletaV", "Toma Física - Valorización", "Boleta", vSQL, vOrden)
End Select

End Sub

Private Sub tlbProcesos_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim lng As Long, strSQL As String, rs2 As New ADODB.Recordset
Dim vTempo As String, rs As New ADODB.Recordset, i As Integer

On Error GoTo vError

If vCodigo > 0 Then

    Select Case Button.Key
     Case "Procesos"
        GLOBALES.gTag = txtCodigo.Text
        frmInvTomaFisicaEjecucion.Show vbModal
        Call txtCodigo_LostFocus

     Case "logico"
        
        Me.MousePointer = vbHourglass
        
        vTempo = txtEstado.Text
        txtEstado = "Comparando..."
        
        
        Call sbInvInventarioProceso(dtpCorte.Value, txtBodegaCod, False, False)
        
        For lng = 1 To vGrid.MaxRows
            vGrid.col = 1
            vGrid.Row = lng
            If vGrid.Text <> "" Then
              strSQL = "select (existencia_inicial + entradas - salidas) as Existencia" _
                     & " from pv_inventario_proceso where usuario = '" & glogon.Usuario _
                     & "' and cod_producto = '" & vGrid.Text _
                     & "' and cod_bodega = '" & txtBodegaCod.Text & "'"
              Call OpenRecordSet(rs, strSQL)
              If Not rs.EOF And Not rs.BOF Then
                 vGrid.col = 5
                 vGrid.Text = CStr(rs!Existencia)
              Else
                 vGrid.col = 5
                 vGrid.Text = "0"
              End If
              rs.Close
            End If
        Next lng
        
        Call sbCalculaTotales
        Call sbCalculaExistencias
        
        txtEstado = vTempo
        
        Me.MousePointer = vbDefault
        
      Case "compara"
        'Generar el Inventario en proceso a la fecha de Corte
          
        i = MsgBox("Esta Seguro que desea comparar la toma física?" & vbCrLf & " >> Se Guardaran los Datos Actuales <<", vbYesNo)
        If i = vbNo Then Exit Sub
        
        Me.MousePointer = vbHourglass
        
        vTempo = txtEstado
        txtEstado = "Comparando..."
        
        'Guardar Datos Base
        strSQL = "update pv_InvTomaFisica set notas = '" & Trim(txtNotas) _
              & "',fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
              & "',fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") _
              & "',cod_bodega = '" & Trim(txtBodegaCod) _
              & "' where consecutivo = " & vCodigo
        Call ConectionExecute(strSQL)
        
        Call sbInvInventarioProceso(dtpCorte.Value, txtBodegaCod, False, False)
        
        'Registra los Productos Nuevos
        strSQL = "insert into pv_invTF_Detalle(consecutivo,cod_bodega,cod_producto,existencia_logica,Existencia_fisica,Ubicacion)" _
               & "(select " & vCodigo & ",'" & txtBodegaCod & "',I.cod_producto,(I.existencia_inicial + I.entradas - I.salidas),0,''" _
               & " from pv_inventario_proceso I inner join pv_productos P on I.cod_producto = P.cod_producto" _
               & " where I.cod_bodega = '" & txtBodegaCod & "' and I.usuario = '" & glogon.Usuario _
               & "' and I.cod_producto not in(select cod_producto from pv_invTF_detalle where consecutivo = " _
               & vCodigo & "))"
        Call ConectionExecute(strSQL)
                
        'Actualiza la Existencia en Libros
        strSQL = "Update D set D.existencia_logica =  isnull(I.existencia_inicial + I.entradas - I.salidas, 0)" _
               & " from pv_invTF_Detalle D left join pv_inventario_proceso I" _
               & " on D.Cod_Bodega = I.Cod_Bodega and D.cod_Producto = I.cod_Producto" _
               & " where I.usuario = '" & glogon.Usuario & "' and D.Consecutivo = " & vCodigo
        Call ConectionExecute(strSQL)
        
        txtEstado = vTempo
    
        Call sbConsulta(vCodigo)
        
        Call sbCalculaTotales
        Call sbCalculaExistencias
        
        Me.MousePointer = vbDefault
        
    End Select

End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtBodegaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtBodegaDes.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_bodega"
  gBusquedas.Orden = "cod_bodega"
  gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtBodegaCod = gBusquedas.Resultado
  txtBodegaDes = gBusquedas.Resultado2
End If

End Sub

Private Sub txtBodegaCod_LostFocus()
txtBodegaDes = fxSIFCCodigos("D", txtBodegaCod, "bodegas")
End Sub

Private Sub txtBodegaDes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtBodegaCod = gBusquedas.Resultado
  txtBodegaDes = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInicio.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "consecutivo"
  gBusquedas.Orden = "consecutivo"
  gBusquedas.Consulta = "select consecutivo,fecha_inicio,fecha_corte,user_crea,notas from pv_InvTomaFisica"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Public Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
vError:
End Sub


Private Sub sbCalculaTotales()
Dim curLogico As Currency, curFisico As Currency
Dim i As Integer, lng As Long

On Error GoTo vError

For lng = 1 To vGrid.MaxRows
 
 vGrid.Row = lng
 vGrid.col = 1
 
 curLogico = 0
 curFisico = 0
 
 If vGrid.Text <> "" Then
    vGrid.col = 5
    curLogico = IIf((vGrid.Text = ""), 0, CCur(vGrid.Text))
    vGrid.col = 6
    curFisico = IIf((vGrid.Text = ""), 0, CCur(vGrid.Text))
    
    vGrid.col = 7
    vGrid.Text = curLogico - curFisico
    
 End If
Next lng

vError:

End Sub

Private Sub sbConsultaArticulo(fila As Long, Columna As Integer, vCriterio As String)
Dim strSQL As String, rs As New ADODB.Recordset, vPaso As Boolean

'Busquedas
'1. Por Codigo del Articulo
'2. Por Codigo de Barras
'3. Por Codigo del Fabricante
vPaso = False

strSQL = "select P.cod_producto,P.descripcion,P.tipo_producto,0 as Existencia" _
       & " from pv_productos P where P.cod_producto = '" & vCriterio & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  MsgBox "No se encontró el Articulo en la Base de Datos...", vbExclamation
Else
  vGrid.Row = fila
  vGrid.col = 1
  vGrid.Text = rs!cod_producto
  vGrid.col = 2
  vGrid.Text = rs!Descripcion
  vGrid.col = 3
  vGrid.Text = rs!tipo_producto
  
  vGrid.col = 5
  vGrid.Text = CStr(rs!Existencia)

  vGrid.col = 6
  vGrid.Text = "0"


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
  If txtEstado.Text <> "Procesada" Then
      Call sbConsultaArticulo(vGrid.ActiveRow, vGrid.ActiveCol, vGrid.Text)
  End If
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
  Call sbCalculaExistencias
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

If KeyCode = vbKeyReturn And vGrid.ActiveCol = vGrid.MaxCols Then
     Call sbCalculaExistencias
End If

End Sub



Private Sub vGrid_KeyPress(KeyAscii As Integer)
Dim curLogico As Currency, curFisico As Currency

On Error GoTo vError

 curLogico = 0
 curFisico = 0
 
 vGrid.Row = vGrid.ActiveRow
 vGrid.col = 1
 
 If vGrid.Text <> "" Then
    vGrid.col = 5
    curLogico = curLogico + IIf((vGrid.Text = ""), 0, CCur(vGrid.Text))
    vGrid.col = 6
    curFisico = curFisico + IIf((vGrid.Text = ""), 0, CCur(vGrid.Text))
    
    vGrid.col = 7
    vGrid.Text = curLogico - curFisico
 End If

vError:

End Sub


