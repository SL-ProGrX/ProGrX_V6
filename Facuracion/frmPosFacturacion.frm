VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmPosFacturacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   11385
   Begin XtremeSuiteControls.CheckBox chkVentaExenta 
      Height          =   252
      Left            =   6840
      TabIndex        =   42
      Top             =   960
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Venta Exenta?"
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
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   312
      Left            =   9480
      TabIndex        =   26
      Top             =   480
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
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
   Begin XtremeSuiteControls.GroupBox gbNotas 
      Height          =   2172
      Left            =   120
      TabIndex        =   23
      Top             =   5400
      Width           =   6732
      _Version        =   1441793
      _ExtentX        =   11874
      _ExtentY        =   3831
      _StockProps     =   79
      Caption         =   "Notas"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnNotas 
         Height          =   495
         Left            =   6120
         TabIndex        =   25
         Top             =   1620
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         Appearance      =   2
         Picture         =   "frmPosFacturacion.frx":0000
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1272
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
         _ExtentY        =   2244
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
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   11040
      Top             =   5400
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   7785
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Estado de la Factura"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1658
            MinWidth        =   1658
            Object.ToolTipText     =   "Caja Registra"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4481
            MinWidth        =   4481
            Object.ToolTipText     =   "Usuario Registro"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1658
            MinWidth        =   1658
            Object.ToolTipText     =   "Caja Anula"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4481
            MinWidth        =   4481
            Object.ToolTipText     =   "Usuario Anula"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2716
            MinWidth        =   2716
            Object.ToolTipText     =   "Fecha Anula"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   688
      _CBWidth        =   11385
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlb"
      MinHeight1      =   330
      Width1          =   3480
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   4005
      NewRow2         =   0   'False
      Child3          =   "tlbAux02"
      MinHeight3      =   330
      Width3          =   2730
      NewRow3         =   0   'False
      Begin MSComctlLib.Toolbar tlbAux02 
         Height          =   330
         Left            =   7710
         TabIndex        =   8
         Top             =   30
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Pedidos"
               Object.ToolTipText     =   "Ingresar Pedidos"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paquetes"
               Object.ToolTipText     =   "Paquetes y Combos"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Precios"
               Object.ToolTipText     =   "Cambio de Precios"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   330
         Left            =   3675
         TabIndex        =   7
         Top             =   30
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   582
         ButtonWidth     =   1826
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Pago ?"
               Key             =   "Pago"
               Object.ToolTipText     =   "Forma de Pago de la Factura"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Clientes"
               Key             =   "Cliente"
               Object.ToolTipText     =   "Ficha de Cliente"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Crédito"
               Key             =   "Credito"
               Object.ToolTipText     =   "Ficha de Crédito"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   6
         Top             =   30
         Width           =   3285
         _ExtentX        =   5794
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
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repBoleta"
                     Text            =   "Boleta de Registro"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repSep1"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repReimpresion"
                     Text            =   "Re-Impresion"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImgAux01 
      Left            =   5400
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosFacturacion.frx":07D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosFacturacion.frx":10B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosFacturacion.frx":13D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosFacturacion.frx":16F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosFacturacion.frx":1A10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8760
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosFacturacion.frx":1D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosFacturacion.frx":51BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosFacturacion.frx":1BB80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosFacturacion.frx":30CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosFacturacion.frx":476B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosFacturacion.frx":5C826
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   6000
      TabIndex        =   10
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2652
      Left            =   0
      TabIndex        =   11
      Top             =   2640
      Width           =   11292
      _Version        =   524288
      _ExtentX        =   19918
      _ExtentY        =   4678
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
      MaxCols         =   490
      ScrollBars      =   2
      SpreadDesigner  =   "frmPosFacturacion.frx":731E8
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   3000
      TabIndex        =   12
      Top             =   480
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   582
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1440
      TabIndex        =   13
      Top             =   480
      Width           =   1572
      _Version        =   1441793
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
   Begin XtremeSuiteControls.ComboBox cboCobro 
      Height          =   312
      Left            =   4560
      TabIndex        =   14
      Top             =   480
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   582
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtPrecioId 
      Height          =   312
      Left            =   1440
      TabIndex        =   15
      Top             =   960
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAgenteId 
      Height          =   312
      Left            =   1440
      TabIndex        =   17
      Top             =   1320
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPrecioDesc 
      Height          =   312
      Left            =   2400
      TabIndex        =   16
      Top             =   960
      Width           =   4092
      _Version        =   1441793
      _ExtentX        =   7218
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
   Begin XtremeSuiteControls.FlatEdit txtAgenteDesc 
      Height          =   312
      Left            =   2400
      TabIndex        =   18
      Top             =   1320
      Width           =   4092
      _Version        =   1441793
      _ExtentX        =   7218
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1440
      TabIndex        =   19
      Top             =   1800
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3600
      TabIndex        =   20
      Top             =   1800
      Width           =   6972
      _Version        =   1441793
      _ExtentX        =   12298
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
   Begin XtremeSuiteControls.FlatEdit txtDocumento 
      Height          =   312
      Left            =   9240
      TabIndex        =   22
      Top             =   1320
      Width           =   1812
      _Version        =   1441793
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtFecha 
      Height          =   312
      Left            =   9480
      TabIndex        =   27
      Top             =   480
      Width           =   1572
      _Version        =   1441793
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTotalRecaudo 
      Height          =   312
      Left            =   1800
      TabIndex        =   28
      Top             =   0
      Width           =   3012
      _Version        =   1441793
      _ExtentX        =   5313
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtRecaudado 
      Height          =   312
      Left            =   9000
      TabIndex        =   30
      Top             =   6960
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDiferencia 
      Height          =   312
      Left            =   9000
      TabIndex        =   31
      Top             =   7320
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtSubTotal 
      Height          =   312
      Left            =   9000
      TabIndex        =   39
      Top             =   5400
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescuento 
      Height          =   312
      Left            =   9000
      TabIndex        =   40
      Top             =   5760
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtImpuestos 
      Height          =   312
      Left            =   9000
      TabIndex        =   41
      Top             =   6120
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTotal 
      Height          =   312
      Left            =   9000
      TabIndex        =   29
      Top             =   6480
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnClienteAsoc 
      Height          =   315
      Left            =   10680
      TabIndex        =   32
      Top             =   1800
      Width           =   372
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   556
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Picture         =   "frmPosFacturacion.frx":73DDE
   End
   Begin XtremeSuiteControls.CheckBox chkImprime 
      Height          =   252
      Left            =   6840
      TabIndex        =   43
      Top             =   1320
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Imprime ticket?"
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
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkElectronico 
      Height          =   492
      Left            =   6840
      TabIndex        =   46
      Top             =   480
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Comprobante Electrónico?"
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
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption lblBodegaDesc 
      Height          =   372
      Left            =   5160
      TabIndex        =   45
      ToolTipText     =   "Bodega"
      Top             =   2280
      Width           =   6132
      _Version        =   1441793
      _ExtentX        =   10816
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption lblProdDesc 
      Height          =   372
      Left            =   0
      TabIndex        =   44
      ToolTipText     =   "Producto"
      Top             =   2280
      Width           =   5172
      _Version        =   1441793
      _ExtentX        =   9123
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   372
      Index           =   5
      Left            =   7080
      TabIndex        =   38
      Top             =   5400
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Sub Total:"
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   372
      Index           =   4
      Left            =   7080
      TabIndex        =   37
      Top             =   5760
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "[-] Descuento:"
      ForeColor       =   255
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   372
      Index           =   3
      Left            =   7080
      TabIndex        =   36
      Top             =   6120
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "[+] Impuestos:"
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   372
      Index           =   2
      Left            =   7080
      TabIndex        =   35
      Top             =   6480
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Total:"
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   1
      Left            =   7080
      TabIndex        =   34
      Top             =   7320
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Pendiente:"
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   0
      Left            =   7080
      TabIndex        =   33
      Top             =   6960
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Total Recaudado:"
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      TabIndex        =   21
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Doc.Ref:"
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
      Index           =   11
      Left            =   9720
      TabIndex        =   4
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Agente"
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
      Index           =   10
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
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
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1212
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Factura"
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
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fecha"
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
      Index           =   4
      Left            =   8400
      TabIndex        =   0
      Top             =   480
      Width           =   1092
   End
End
Attribute VB_Name = "frmPosFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vScroll As Boolean, vFilaNote As Long, vPaso As Boolean
Dim mRow As Long

Private Sub btnClienteAsoc_Click()
    gBusquedas.Convertir = "N"
    gBusquedas.Col1Name = "Cédula Colilla"
    gBusquedas.Col2Name = "Cédula Real"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    txtCedula.Text = Trim(gBusquedas.Resultado)
    txtNombre.Text = Trim(gBusquedas.Resultado3)
    
    Call sbXFichaCliente(txtCedula.Text)

End Sub

Private Sub btnNotas_Click()

If vFilaNote > 0 Then
  vGrid.Row = vFilaNote
  vGrid.col = 1
  vGrid.CellNote = txtNotas.Text
End If

End Sub

Private Sub cboCobro_Click()
If cboCobro.Text = "Contado" Then
  tlbAux.Buttons.Item(1).Enabled = True
  tlbAux.Buttons.Item(4).Enabled = False
Else
  tlbAux.Buttons.Item(1).Enabled = False
  tlbAux.Buttons.Item(4).Enabled = True
End If
End Sub

Private Sub cboTipo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

If Not vEdita Then
    Select Case Mid(cboTipo.Text, 1, 1)
      Case "A" 'Automaticas
         strSQL = "select facturas_auto as Factura from pv_consecutivos"
         txtCodigo.Enabled = False
      Case "M" 'Manuales
         strSQL = "select facturas_man as Factura from pv_consecutivos"
         txtCodigo.Enabled = True
    End Select
    Call OpenRecordSet(rs, strSQL)
    If rs.EOF And rs.BOF Then
      txtCodigo = 1
    Else
      txtCodigo = rs!factura + 1
    End If
    rs.Close
End If 'vEdita

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 cod_factura from pv_Facturacion" _
           & " where tipo = '" & Mid(cboTipo.Text, 1, 1) & "'"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and cod_factura > '" & txtCodigo & "' order by cod_factura asc"
    Else
       strSQL = strSQL & " and cod_factura < '" & txtCodigo & "' order by cod_factura desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!cod_Factura)
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
vModulo = 33
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

 vModulo = 33
 vGrid.AppearanceStyle = fxGridStyle

 vPaso = True
 
If gCajas.VentaExenta Then
    chkVentaExenta.Enabled = True
    chkVentaExenta.Value = xtpChecked
Else
    chkVentaExenta.Enabled = False
    chkVentaExenta.Value = xtpUnchecked
End If
 
cboCobro.Clear
cboCobro.AddItem "Contado"
cboCobro.AddItem "Credito"
cboCobro.Text = "Contado"

cboTipo.Clear
cboTipo.AddItem "Automáticas"
cboTipo.AddItem "Manuales"
cboTipo.Text = "Automáticas"

 vPaso = False

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub

Private Sub sbLimpiaPantalla()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

Me.MousePointer = vbHourglass

On Error GoTo vError

gCajas.TicketId = fxPosCajaTicket

'Revisa el Estado de la Caja
If gCajas.Apertura = 0 Then
    Me.MousePointer = vbDefault
    MsgBox "La Caja no cuenta con una apertura válida!", vbExclamation
    Unload Me
Else


End If

vCodigo = 0
txtCodigo = ""

vFilaNote = 0

StatusBarX.Panels.Item(1) = "En Proceso"
StatusBarX.Panels.Item(2) = ""
StatusBarX.Panels.Item(3) = ""
StatusBarX.Panels.Item(4) = ""
StatusBarX.Panels.Item(5) = ""
StatusBarX.Panels.Item(6) = ""

dtpFecha.Visible = gCajas.ModFechas

txtDocumento.Text = ""
txtNombre.Text = ""
txtCedula.Text = ""

cboCobro.Text = "Contado"
cboTipo.Text = "Automáticas"


vGrid.MaxRows = 0
vGrid.MaxCols = 9
vGrid.MaxRows = 1
mRow = 1

txtSubTotal.Text = 0
txtDescuento.Text = 0
txtImpuestos.Text = 0
txtTotal.Text = 0
txtRecaudado.Text = 0
txtDiferencia.Text = 0


'Inicializa valores por defecto
strSQL = "select isnull(Bo.Descripcion,'') as 'Bodega_Desc', isnull(Cl.Nombre,'') as 'Cliente_Desc' " _
       & ", isnull(Tp.Descripcion,'') as 'Precio_Desc', isnull(Ag.Nombre,'') as 'Agente_Desc' " _
       & ", dbo.MyGetdate() as 'Fecha'" _
       & " from pv_cajas C left join PV_Bodegas Bo on C.Def_Bodega = Bo.cod_Bodega" _
       & " left join PV_Clientes Cl on C.def_cliente = Cl.Cedula" _
       & " left join pv_tipos_precios Tp on C.def_precio = Tp.cod_precio" _
       & " left join PV_Agentes Ag on C.def_agente = Ag.cod_Agente" _
       & " where C.cod_caja = '" & gCajas.Caja & "' and C.usuario = '" & gCajas.Usuario & "'"
       
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtCedula.Text = gCajas.Cliente
   txtNombre.Text = rs!Cliente_Desc & ""
   
   txtAgenteId.Text = gCajas.Agente
   txtAgenteDesc.Text = rs!Agente_Desc
   
   txtPrecioId.Text = gCajas.Precio
   txtPrecioDesc.Text = rs!Precio_Desc
   
   lblBodegaDesc.Caption = rs!Bodega_desc
   lblBodegaDesc.Tag = gCajas.Bodega
   
   
    dtpFecha.Value = rs!fecha
    txtFecha.Text = Format(rs!fecha, "yyyy/mm/dd hh:mm:ss")
   
End If
rs.Close

lblProdDesc.Caption = ""
lblBodegaDesc.Caption = ""

txtCodigo.Enabled = True
txtCodigo.SetFocus

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim strSQL As String
'Bloquear Caja al Entrar / Desbloquear al Salir
strSQL = "update pv_cajas set bloqueo = 0" _
       & " where cod_caja = '" & gCajas.Caja & "' and usuario = '" & gCajas.Usuario & "'"
Call ConectionExecute(strSQL)
End Sub


Private Sub sbNuevo()
vEdita = False
Call sbLimpiaPantalla
Call sbToolBar(tlb, "edicion")

'Si la factura es manual no proceso el numero de consecutivo
If Mid(cboTipo.Text, 1, 2) = "01" Then
    txtCodigo.Enabled = False
    cboTipo.SetFocus
Else
    txtCodigo.Enabled = True
    txtCodigo.SetFocus
End If
End Sub



Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

TimerX.Interval = 0

'frmPosCajaLogin.Show vbModal
If gCajas.Apertura > 0 Then

    'No puede darse EOF porque ya verificado
'    strSQL = "select nombre,def_cliente,def_bodega,def_precio,def_agente" _
'           & ",modifica_precio,modifica_fechas from pv_cajas" _
'           & " where cod_caja = '" & gCajas.Caja & "' and usuario = '" & gCajas.Usuario & "'"
'    Call OpenRecordSet(rs, strSQL)
'        gCajas.Agente = rs!def_agente
'        gCajas.Bodega = rs!def_bodega
'        gCajas.Cliente = rs!def_cliente
'        gCajas.Precio = rs!def_precio
'        gCajas.Nombre = rs!Nombre
'        gCajas.BodegaDesc = fxSIFCCodigos("D", rs!def_bodega, "bodegas")
'        gCajas.ModFechas = IIf((rs!modifica_fechas = 1), True, False)
'        gCajas.ModPrecios = IIf((rs!modifica_precio = 1), True, False)
'    rs.Close
'
    'Bloquear Caja al Entrar / Desbloquear al Salir
    strSQL = "update pv_cajas set bloqueo = 1" _
           & " where cod_caja = '" & gCajas.Caja & "' and usuario = '" & gCajas.Usuario & "'"
    Call ConectionExecute(strSQL)

    Me.Caption = "POS: Facturación" & Space(10) & "Caja: " & Trim(gCajas.Caja) & Space(5) & "Usuario: " _
               & gCajas.Usuario & Space(5) & " Nombre: " & gCajas.Nombre

Else
  Unload Me

End If

'Inicializa en Estado de Edicion
Call sbNuevo

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      Call sbNuevo
      
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
      If vCodigo = 0 Then
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

Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 1
        vGrid.CellNote = rs!Notas & ""
     Case 2
        vGrid.Text = CStr(rs!Cod_Producto)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = CStr(rs!producto)
        vGrid.TextTip = TextTipFixed
     Case 3
        vGrid.Text = CStr(rs!Cantidad)
     Case 4
        vGrid.Text = CStr(rs!cod_bodega)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = CStr(rs!Bodega)
        vGrid.TextTip = TextTipFixed
     Case 5
        vGrid.Text = CStr(rs!Precio)
        vGrid.CellNote = "Costo Registrado : " & Format(IIf(IsNull(rs!costo), 0, rs!costo), "Standard")
        vGrid.CellTag = CStr(IIf(IsNull(rs!costo), 0, rs!costo))
     Case 6
        vGrid.Text = CStr(rs!descuento)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = CStr((rs!Precio * rs!Cantidad) * (rs!descuento / 100))
        vGrid.TextTip = TextTipFixed
     Case 7
        vGrid.Text = CStr(rs!imp_consumo)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = CStr((rs!Precio * rs!Cantidad) * (rs!imp_consumo / 100))
        vGrid.TextTip = TextTipFixed
     Case 8
        vGrid.Text = CStr(rs!imp_ventas)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = CStr((rs!Precio * rs!Cantidad) * (rs!imp_ventas / 100))
        vGrid.TextTip = TextTipFixed
     Case 9
        vGrid.Text = CStr(rs!Total - (rs!Total * (rs!descuento / 100)))
    End Select
  
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub sbConsulta(vCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select F.*,C.nombre, case when isnull(CXC_TIPO,'CNT') = 'CRD' then 'Credito' else 'Contado' end as 'Tipo_Cobro'" _
       & ",rtrim(P.descripcion) as 'Precio_Desc'" _
       & ",rtrim(A.nombre) as 'Agente_Desc'" _
       & ",rtrim(R.descripcion) as FormaPago_Desc" _
       & " from pv_facturacion F inner join pv_clientes C on F.cedula = C.cedula" _
       & " inner join pv_tipos_precios P on F.cod_precio = P.cod_precio" _
       & " inner join pv_agentes A on F.cod_agente = A.cod_agente" _
       & "  left join pv_formas_pago R on F.cod_forma_pago = R.cod_forma_pago" _
       & " where cod_factura = '" & vCodigo & "' and tipo = '" _
       & Mid(cboTipo.Text, 1, 1) & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!cod_Factura
  txtCodigo = rs!cod_Factura
      
  Select Case UCase(rs!Estado)
    Case "A"
      StatusBarX.Panels.Item(1) = "Anulada"
    Case "P"
      StatusBarX.Panels.Item(1) = "Procesada"
  End Select
    
    StatusBarX.Panels.Item(2) = rs!Cod_Caja & ""
    StatusBarX.Panels.Item(3) = rs!Usuario & ""
    StatusBarX.Panels.Item(4) = rs!anu_CajaCod & ""
    StatusBarX.Panels.Item(5) = rs!anu_CajaUser & ""
    StatusBarX.Panels.Item(6) = rs!anu_fecha & ""
    
  Select Case UCase(rs!Tipo)
    Case "A"
      cboTipo.Text = "Automáticas"
    Case "M"
      cboTipo.Text = "Manuales"
  End Select
    
  cboCobro.Text = rs!Tipo_Cobro
    
  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre
  
  txtAgenteId.Text = rs!Cod_Agente
  txtAgenteDesc.Text = rs!Agente_Desc
  
  txtPrecioId.Text = rs!cod_precio
  txtPrecioDesc.Text = rs!Precio_Desc
  
  txtDocumento = rs!Documento
     
  txtFecha.Text = Format(rs!fecha, "yyyy/mm/dd hh:mm:ss")
  
  dtpFecha.Value = rs!fecha
  
  txtSubTotal = Format(rs!sub_Total, "Standard")
  txtDescuento = Format(rs!descuento, "Standard")
  txtImpuestos = Format(rs!imp_ventas + rs!imp_consumo, "Standard")
  txtTotal = Format(rs!Total, "Standard")

  strSQL = "select D.Notas,D.cod_producto,P.descripcion as Producto,D.cantidad,D.cod_bodega,B.descripcion as 'Bodega'" _
         & ",isnull(D.descuento,0) as 'Descuento',isnull(D.imp_consumo,0) as 'Imp_Consumo',D.precio,D.imp_ventas," _
         & "(D.cantidad * D.precio) + (D.cantidad * D.precio * (D.imp_ventas / 100)) as 'Total', D.costo" _
         & " from pv_factura_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " inner join pv_bodegas B on D.cod_bodega = B.cod_bodega" _
         & " where D.cod_factura = '" & rs!cod_Factura & "' and D.tipo = '" & rs!Tipo _
         & "' order by D.Linea"
  Call sbCargaGridLocal(vGrid, 9, strSQL)
  
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
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

vMensaje = ""
fxValida = True

On Error GoTo vError


If Len(txtCodigo) = 0 Then vMensaje = vMensaje & vbCrLf & " - Número de Factura no es válido..."

If gCajas.Apertura = 0 Then vMensaje = vMensaje & vbCrLf & " - La caja no se encuentra Abierta! favor reingresar al módulo"

If cboCobro.Text = "Contado" And CCur(txtDiferencia.Text) <> 0 Then vMensaje = vMensaje & vbCrLf & " - No se ha indicado los valores de pago de la factura"


vMensaje = fxInvVerificaLineaDetalle(vGrid, 4, "S", 2, 4)

If CCur(txtTotal.Text) <> CCur(txtRecaudado.Text) And cboCobro.Text = "Contado" Then
   vMensaje = vMensaje & vbCrLf & " - El Monto de Recaudo no Cancela la Factura!"
End If

'Validar Factura Manual, que no existe
If Mid(cboTipo.Text, 1, 1) = "M" Then
   strSQL = "select isnull(count(*),0) as Existe from pv_facturacion where tipo = 'M' and cod_factura = '" & txtCodigo.Text & "'"
   Call OpenRecordSet(rs, strSQL)
   If rs!Existe > 0 Then
       vMensaje = vMensaje & vbCrLf & " - El numero de factura manual ya ha sido registrado con anterioridad..."
   End If
   rs.Close
End If

'Validar Cliente, que exista
strSQL = "select isnull(count(*),0) as Existe from pv_clientes" _
       & " where cedula = '" & txtCedula & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    vMensaje = vMensaje & vbCrLf & " - No Existe registro del cliente especificado ..."
End If
rs.Close

'Validar que se haya facturado almenos un articulo
   
Select Case vGrid.MaxRows
  Case 1 'Solo hay una linea, verifica si existe un producto en ella
    vGrid.col = 2
    vGrid.Row = 1
    If Trim(vGrid.Text) = "" Then
       vMensaje = vMensaje & vbCrLf & " - No hay productos / articulos o servicios en el detalle ..."
    End If
  
  Case 0 'No hay linea de detalle
       vMensaje = vMensaje & vbCrLf & " - No hay productos / articulos o servicios en el detalle..."
End Select

If gCajas.ModFechas Then
  If Not fxInvPeriodos(dtpFecha.Value) Then vMensaje = vMensaje & vbCrLf & " - El periodo de la factura ya fue cerrado o no es válido ..."
Else
  If Not fxInvPeriodos(fxFechaServidor) Then vMensaje = vMensaje & vbCrLf & " - El periodo de la factura ya fue cerrado o no es válido..."
End If


vError:

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If


End Function

Private Sub sbImprime()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vSQL As String, vFormato As Integer, vFE_Utiliza As Integer, vFE_Archivo As String

Me.MousePointer = vbHourglass

vFormato = 1

strSQL = "select formato_factura, FORMATO_ESPECIAL,FORMATO_ESPECIAL_ARCHIVO " _
       & " from pv_cajas where cod_caja = '" & gCajas.Caja & "' and usuario = '" & gCajas.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  vFormato = rs!Formato_Factura
  vFE_Utiliza = rs!FORMATO_ESPECIAL
  vFE_Archivo = rs!FORMATO_ESPECIAL_ARCHIVO
End If
rs.Close


With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowGroupTree = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes POS"
 
 .Connect = glogon.ConectRPT
 If vFE_Utiliza = 0 Then
        Select Case vFormato
         Case 1
            strSQL = "select CEDULA_JURIDICA, EMAIL , TELEFONOEMP , PAG_NOMCORTO  From SIF_EMPRESA"
            Call OpenRecordSet(rs, strSQL)
            
    
            .Formulas(0) = "fxEmpresa = '" & Trim(rs!PAG_NOMCORTO) & "'"
            .Formulas(1) = "fxTitulo = 'Ced.Jur." & Trim(rs!cedula_juridica) & "'"
            .Formulas(2) = "fxSubTitulo = '" & Trim(rs!Email) & " ¦ " & Trim(rs!TelefonoEmp) & "'"
            
            .ReportFileName = SIFGlobal.fxPathReportes("POS_FacturaFlat.rpt")
            rs.Close
         
         Case 2
           .ReportFileName = SIFGlobal.fxPathReportes("POS_Factura.rpt")
         Case Else
            
            strSQL = "select CEDULA_JURIDICA, EMAIL , TELEFONOEMP , PAG_NOMCORTO  From SIF_EMPRESA"
            Call OpenRecordSet(rs, strSQL)
            
    
            .Formulas(0) = "fxEmpresa = '" & Trim(rs!PAG_NOMCORTO) & "'"
            .Formulas(1) = "fxTitulo = 'Ced.Jur." & Trim(rs!cedula_juridica) & "'"
            .Formulas(2) = "fxSubTitulo = '" & Trim(rs!Email) & "'"
            
            .ReportFileName = SIFGlobal.fxPathReportes("POS_FacturaFlat.rpt")
            rs.Close
        End Select
 Else
       .ReportFileName = SIFGlobal.fxPathReportes(vFE_Archivo)
 End If
    
    
 .SelectionFormula = "{vPOS_Factura.COD_FACTURA} = '" & txtCodigo.Text _
                   & "' AND {vPOS_Factura.TIPO} = '" & Mid(cboTipo.Text, 1, 1) & "'"
 If vFormato = 1 Then
    .Destination = crptToPrinter
 End If

 .PrintReport

End With

Me.MousePointer = vbDefault


End Sub

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, curCantidad As Currency
Dim vCodPro As String, vCodBodega As String, vFecha As Date, vNotas As String
Dim curPrecio As Currency, curImpVentas As Currency, curImpConsumo As Currency
Dim vImprime As Boolean, curDescuento As Currency, vCobro As String

On Error GoTo vError


'Solo se puede Insertar y no Editar
'01 - Guardar el registro en las Entradas y Afectar Inventarios

If vEdita Then
   MsgBox "No se puede editar una factura Guardada...", vbInformation
   Exit Sub
End If

Select Case Mid(cboTipo.Text, 1, 1)
  Case "A"
    vCodigo = fxSIFCConsecutivos("Facturas_Auto")
    txtCodigo = vCodigo
    vImprime = True
  Case "M"
    vCodigo = fxSIFCConsecutivos("Facturas_Man")
    txtDocumento = vCodigo
    vCodigo = txtCodigo
    vImprime = False
End Select

If chkImprime.Value = xtpUnchecked Then
   vImprime = False
End If

If cboCobro.Text = "Contado" Then
   vCobro = "CNT"
Else
   vCobro = "CRD"
End If


If gCajas.ModFechas Then
  vFecha = CDate(Format(dtpFecha.Value, "yyyy/mm/dd") & " " & Format(fxFechaServidor, "hh:mm:ss"))
Else
  vFecha = fxFechaServidor
End If

'glogon.Conection.BeginTrans

 
strSQL = "insert pv_facturacion(cod_factura,tipo,cedula,cod_agente,cod_precio,cod_caja,usuario,documento" _
       & ",cod_forma_pago,estado,fecha,sub_total,descuento,imp_ventas,imp_consumo,total,asiento_estado" _
       & ", CXC_TIPO, Venta_Exenta, FE_ESTADO, FE_TIPO)" _
       & " values('" & vCodigo & "','" & Mid(cboTipo.Text, 1, 1) _
       & "','" & txtCedula.Text & "','" & txtAgenteId.Text & "','" & txtPrecioId.Text & "','" _
       & gCajas.Caja & "','" & gCajas.Usuario & "','" & txtDocumento & "',1" _
       & ",'P','" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "'," & CCur(txtSubTotal) & "," & CCur(txtDescuento) & "," _
       & CCur(txtImpuestos) & ",0," & CCur(txtTotal) & ",'P','" & vCobro & "'," & chkVentaExenta.Value _
       & ",'E', 'NA')"
Call ConectionExecute(strSQL)

'Registra Movimiento en Cajas
'Call sbPosCajaMovRegistra(IIf((Mid(cboTipo.Text, 1, 1) = "A"), "FA", "FM"), gCajas.Caja, gCajas.Usuario, CInt(gCajas.Apertura) _
'         , CCur(txtTotal), 1, CStr(vCodigo), txtDocumento & " (" & txtCedula & " - " & txtNombre & ")")

strSQL = "exec spPV_Cajas_Mov_Tempo_Aplica '" & txtCedula.Text & "','" & gCajas.Caja & "','" & gCajas.TicketId & "'," _
                & gCajas.Apertura & ",'" & IIf((Mid(cboTipo.Text, 1, 1) = "A"), "FA", "FM") & "','" & vCodigo _
                & "','" & txtCedula.Text & " - " & txtNombre.Text & "','" & gCajas.Usuario & "'"
Call ConectionExecute(strSQL)

'spPV_Cajas_Mov_Tempo_Aplica(@Cedula varchar(30), @Caja varchar(10), @Ticket varchar(50), @CodAC int, @Origen varchar(10)
'                                      , @Comprobante varchar(30), @Detalle varchar(500), @Usuario varchar(30))

'Registro Bitacora
Call Bitacora("Registra", "Factura N.: " & vCodigo & " (Tipo: " & Mid(cboTipo.Text, 1, 1) & ")")

txtCodigo.Enabled = True

'Inicia Lote
strSQL = ""

'Guardar Detalle de la Factura y Registra Inventario
strSQL = "delete pv_factura_detalle" _
         & " where cod_factura = '" & vCodigo & "' and tipo = '" _
         & Mid(cboTipo.Text, 1, 1) & "'"

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  
  vGrid.col = 3
  curCantidad = CCur(IIf((vGrid.Text = ""), 0, vGrid.Text))
  
  vGrid.col = 2
  
  If vGrid.Text <> "" And curCantidad > 0 Then
      
    vGrid.col = 1
    vNotas = Trim(vGrid.CellNote)
    
    vGrid.col = 2
    vCodPro = Trim(vGrid.Text)
    
    strSQL = strSQL & Space(10) & "insert pv_factura_detalle(linea,cod_factura,tipo,notas, cod_producto,cantidad,cod_bodega" _
           & ",precio,imp_ventas,imp_consumo,descuento,costo) values(" & i & ",'" & vCodigo & "','" _
           & Mid(cboTipo.Text, 1, 1) & "','" & vNotas & "','" & vCodPro & "'," & curCantidad & ",'"
    
    vGrid.col = 4
    vCodBodega = Trim(vGrid.Text)
    strSQL = strSQL & vGrid.Text & "',"
    
    vGrid.col = 5
    curPrecio = CCur(vGrid.CellTag)
    strSQL = strSQL & CCur(vGrid.Text) & ","
    
    vGrid.col = 8
    curImpVentas = CCur(vGrid.Text)
    strSQL = strSQL & CCur(vGrid.Text) & ","
    
    vGrid.col = 7
    curImpConsumo = CCur(vGrid.Text)
    strSQL = strSQL & CCur(vGrid.Text) & ","
    
    
    vGrid.col = 6
    curDescuento = CCur(vGrid.Text)
    strSQL = strSQL & CCur(vGrid.Text) & "," & curPrecio & ")"
    
    If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If
    
  End If
Next i

If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If


'Comprobante Electrónico para Hacienda
If chkElectronico.Value = xtpChecked And Len(txtCedula.Text) > 8 Then


    Dim CodPais As String, FechaTransac As Date, pCedula As String _
        , codSucursal As String, TerminalPOS As String, ComprobanteInterno As String _
        , SituacionComprobante As String, TipoComprobante As String
    Dim pClave50 As String, pClave20 As String
    
    ', pFecha As Date, pLinea As Integer, idEmpresa As String


    strSQL = "exec spPOS_FE_SECUENCIA '01', 'FE'"
    Call OpenRecordSet(rs, strSQL)
    
    CodPais = Trim(rs!Pais)
    codSucursal = Trim(rs!Sucursal)
    TerminalPOS = Trim(rs!Terminal)
    
    SituacionComprobante = "1"
    
    TipoComprobante = Trim(rs!Tipo_Comprobante)
 
    ComprobanteInterno = rs!Consecutivo
    FechaTransac = vFecha
    
    pCedula = Trim(rs!CEDULA_ID)
    pClave50 = fxHacienda_Clave50("506", FechaTransac, pCedula, codSucursal, TerminalPOS, ComprobanteInterno, SituacionComprobante, TipoComprobante)
    pClave20 = fxHacienda_Clave20(codSucursal, TerminalPOS, ComprobanteInterno, TipoComprobante)
      
    rs.Close

    'Actualiza Factura
    strSQL = "update pv_facturacion set FE_NUMERO = '" & pClave20 & "', FE_CLAVE = '" & pClave50 _
           & "', FE_ESTADO = 'P', FE_TIPO = '" & TipoComprobante & "'" _
           & " Where COD_FACTURA = '" & vCodigo & "' AND TIPO = '" & Mid(cboTipo.Text, 1, 1) & "'"
    Call ConectionExecute(strSQL)
    
End If

'Guarda Transaccion
'glogon.Conection.CommitTrans

'Actualiza Inventario
strSQL = "exec spInv_Afectacion_POS '" & vCodigo & "','" & Mid(cboTipo.Text, 1, 1) & "','S'"
Call ConectionExecute(strSQL)

'Imprimir Aqui Factura
If vImprime Then Call sbImprime


'Activa Opciones de Credito
If cboCobro.Text = "Credito" Then
    GLOBALES.gTag = txtCodigo.Text
    GLOBALES.gTag2 = Mid(cboTipo.Text, 1, 1)
    GLOBALES.gTag3 = txtCedula.Text
    
    frmPosFichaCredito.Show vbModal
End If
 

'Call sbToolBar(tlb, "activo")
Call sbNuevo


Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
' glogon.Conection.RollbackTrans
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
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbRepBoleta()

On Error GoTo vError

MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del POS"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 
 .SelectionFormula = "{PV_FACTURACION.COD_FACTURA} = '" & txtCodigo & "' AND {PV_FACTURACION.TIPO} = '" _
       & Mid(cboTipo.Text, 1, 1) & "'"
 .ReportFileName = SIFGlobal.fxPathReportes("POS_FacturaBoletaRegistro.rpt")
 
 .SubreportToChange = "sbFiadores"
 .Connect = glogon.ConectRPT
 
 
 .PrintReport
End With

MousePointer = vbDefault

Exit Sub

vError:
  MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
  Case "repBoleta"
     Call sbRepBoleta
  Case "repReimpresion"
     Call sbImprime
     'Ojo con los niveles de Autorizacion y Preguntar por el Formato si es Boucher o Clasico
End Select
End Sub

Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset

If Len(txtCodigo.Text) = 0 Then Exit Sub

GLOBALES.gTag = txtCodigo.Text
GLOBALES.gTag2 = Mid(cboTipo.Text, 1, 1)


GLOBALES.gTag3 = txtCedula.Text
gCajas.TicketMonto = CCur(txtTotal.Text)

Select Case Button.Key
  Case "Cliente"
    Call sbFormsCall("frmPosFichaCliente", , , , False, Me)
  
  Case "Credito"
    If cboCobro.Text = "Credito" Then
        Call sbFormsCall("frmPosFichaCredito", , , , False, Me)
    End If
 
  Case "Pago"
    If cboCobro.Text = "Contado" Then
        Call sbFormsCall("frmPosCajaPagoDetalle", vbModal, , , False, Me)
    End If
    
End Select

txtRecaudado.Text = Format(gCajas.TicketAbono, "Standard")
txtDiferencia.Text = Format(CCur(txtTotal.Text) - (txtRecaudado.Text), "Standard")

End Sub

Private Sub sbCargaPedido(vPedido As Long)
Dim strSQL As String, rs As New ADODB.Recordset

Call tlb_ButtonClick(tlb.Buttons.Item(1))

strSQL = "select F.cod_pedido,F.fecha,F.vence,F.sub_total,F.descuento,F.imp_ventas,F.total,F.plantilla" _
       & ",F.cedula,C.nombre,rtrim(P.descripcion) as 'Precio_Desc'" _
       & ",rtrim(A.nombre) as 'Agente_Desc'" _
       & ",(rtrim(CONVERT(char, F.cod_forma_pago))+ ' - ' + R.descripcion) as FormaPago" _
       & " from pv_Pedidos F inner join pv_clientes C on F.cedula = C.cedula" _
       & " inner join pv_tipos_precios P on F.cod_precio = P.cod_precio" _
       & " inner join pv_agentes A on F.cod_agente = A.cod_agente" _
       & "  left join pv_formas_pago R on F.cod_forma_pago = R.cod_forma_pago" _
       & " where cod_pedido = '" & vPedido & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  
  txtCedula = rs!Cedula
  txtNombre = rs!Nombre
  
  txtAgenteId.Text = rs!Cod_Agente
  txtAgenteDesc.Text = rs!Agente_Desc
      
  txtPrecioId.Text = rs!cod_precio
  txtPrecioDesc.Text = rs!Precio_Desc
    
  txtSubTotal = Format(rs!sub_Total, "Standard")
  txtDescuento = Format(rs!descuento, "Standard")
  txtImpuestos = Format(rs!imp_ventas, "Standard")
  txtTotal = Format(rs!Total, "Standard")

  strSQL = "select '' as 'Notas', D.cod_producto,P.descripcion as Producto,D.cantidad,D.cod_bodega,B.descripcion as Bodega" _
         & ",0 as Descuento,0 as Imp_Consumo,D.precio,D.imp_ventas" _
         & ",(D.cantidad * D.precio) + (D.cantidad * D.precio * (D.imp_ventas / 100)) as Total, P.costo_regular as Costo" _
         & " from pv_pedidos_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " inner join pv_bodegas B on D.cod_bodega = B.cod_bodega" _
         & " where D.cod_pedido = '" & rs!cod_pedido & "' order by D.Linea"
  Call sbCargaGridLocal(vGrid, 9, strSQL)
  
  Call sbCalculaTotales
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

End Sub

Private Sub tlbAux02_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "Pedidos"
        frmPosFacPedidos.Show vbModal
        If gCajas.Pedido > 0 Then Call sbCargaPedido(gCajas.Pedido)
        
  Case "Paquetes"
        frmPosFacPaquetes.Show vbModal
  Case "Precios"
End Select
End Sub




Private Sub txtAgenteDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedula.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cod_agente,nombre from pv_Agentes"
  gBusquedas.Filtro = " and Estado = 'A'"
  frmBusquedas.Show vbModal
  txtAgenteId.Text = gBusquedas.Resultado
  txtAgenteDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtAgenteId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAgenteDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_agente"
  gBusquedas.Orden = "cod_agente"
  gBusquedas.Consulta = "select cod_agente,nombre from pv_Agentes"
  gBusquedas.Filtro = " and Estado = 'A'"
  frmBusquedas.Show vbModal
  txtAgenteId.Text = gBusquedas.Resultado
  txtAgenteDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtAgenteId_LostFocus()
txtAgenteDesc.Text = fxSIFCCodigos("D", txtAgenteId, "Agentes")
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula.Text = gBusquedas.Resultado
  txtNombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCedula_LostFocus()
'Verifica el Enlace con Cuentas Corrientes
Call sbXFichaCliente(txtCedula.Text)
txtNombre.Text = fxSIFCCodigos("D", txtCedula, "clientes")
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_factura"
  gBusquedas.Orden = "cod_factura"
  gBusquedas.Consulta = "select cod_factura,tipo,cod_caja,usuario from pv_facturacion"
  gBusquedas.Filtro = " and tipo = '" & Mid(cboTipo.Text, 1, 1) & "'"
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
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


Private Sub sbCalculaTotales()
Dim curSubTotal As Currency, curIV As Currency, curIC As Currency, curDes As Currency
Dim curTmpPrecio As Currency, curTmpIV As Currency, curTmpCant As Currency
Dim i As Integer, lng As Long, curTmpIC As Currency, curTmpDes As Currency

'**********************************************    OJO
'Revisar esta formula por la situacion del descuento, si es antes o despues del
'impuesto de ventas, por ahora está despues del impuesto

curSubTotal = 0
curIV = 0
curIC = 0
curDes = 0

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 3
 
 If vGrid.Text <> "" Then
    curTmpCant = CCur(vGrid.Text)
    vGrid.col = 5
    curTmpPrecio = CCur(vGrid.Text)
    vGrid.col = 6
    curTmpDes = CCur(vGrid.Text) / 100
    vGrid.col = 7
    curTmpIC = CCur(vGrid.Text) / 100
    vGrid.col = 8
    curTmpIV = CCur(vGrid.Text) / 100
    
    
    curTmpDes = curTmpCant * curTmpPrecio * curTmpDes
    curTmpIC = ((curTmpCant * curTmpPrecio) - curTmpDes) * curTmpIC
    curTmpIV = ((curTmpCant * curTmpPrecio) - curTmpDes) * curTmpIV
    
    vGrid.col = 9
    vGrid.Text = (curTmpCant * curTmpPrecio) - curTmpDes + curTmpIC + curTmpIV
     
     
    curSubTotal = curSubTotal + (curTmpCant * curTmpPrecio) '- curTmpDes
    curIV = curIV + curTmpIV
    curIC = curIC + curTmpIC
    curDes = curDes + curTmpDes

 End If
Next lng

txtSubTotal = Format(curSubTotal, "Standard")
txtDescuento = Format(curDes, "Standard")
txtImpuestos = Format(curIV + curIC, "Standard")
txtTotal = Format(curSubTotal + curIV + curIC - CCur(txtDescuento), "Standard")
txtDiferencia.Text = Format(CCur(txtTotal.Text) - (txtRecaudado.Text), "Standard")

End Sub

Private Sub sbConsultaArticulo(fila As Long, Columna As Integer, vCriterio As String)
Dim strSQL As String, rs As New ADODB.Recordset, vPaso As Boolean

'Busquedas
'1. Por Codigo del Articulo
'2. Por Codigo de Barras
'3. Por Codigo del Fabricante
vPaso = False

vGrid.Row = fila
vGrid.col = 5
vGrid.Lock = False

vGrid.Row = fila
vGrid.col = 6
vGrid.Lock = False

vGrid.Row = fila
vGrid.col = 7
vGrid.Lock = False

vGrid.Row = fila
vGrid.col = 8
vGrid.Lock = True


strSQL = "select cod_producto,descripcion,costo_regular,precio_regular,impuesto_ventas,impuesto_consumo from pv_productos" _
       & " where cod_producto = '" & vCriterio & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then vPaso = True

If Not vPaso Then
  rs.Close
  strSQL = "select cod_producto,descripcion,costo_regular,precio_regular,impuesto_ventas,impuesto_consumo from pv_productos" _
         & " where cod_barras = '" & vCriterio & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then vPaso = True
End If

If Not vPaso Then
  rs.Close
  strSQL = "select cod_producto,descripcion,costo_regular,precio_regular,impuesto_ventas,impuesto_consumo from pv_productos" _
         & " where cod_fabricante = '" & vCriterio & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then vPaso = True
End If

If Not vPaso Then
  MsgBox "No se encontró el Articulo en la Base de Datos...", vbExclamation
Else
  vGrid.Row = fila
  vGrid.col = 2
  vGrid.Text = rs!Cod_Producto
  vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
  vGrid.CellNote = CStr(rs!Descripcion)
  vGrid.TextTip = TextTipFixed
  
  vGrid.col = 3
  vGrid.Text = CStr(1)
  
  vGrid.col = 4
  vGrid.Text = gCajas.Bodega
  vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
  vGrid.CellNote = gCajas.BodegaDesc
  vGrid.TextTip = TextTipFixed
  
  vGrid.col = 5
  vGrid.Text = CStr(rs!precio_regular)
  vGrid.CellNote = "Costo Actual : " & Format(IIf(IsNull(rs!costo_regular), 0, rs!costo_regular), "Standard")
  vGrid.CellTag = CStr(IIf(IsNull(rs!costo_regular), 0, rs!costo_regular))
  
  vGrid.col = 6
  vGrid.Text = CStr(0)
  
  If chkVentaExenta.Value = xtpChecked Then
        vGrid.col = 7
        vGrid.Text = CStr(0)
        
        vGrid.col = 8
        vGrid.Text = CStr(0)
  
  Else
        vGrid.col = 7
        vGrid.Text = CStr(rs!impuesto_consumo)
        
        vGrid.col = 8
        vGrid.Text = CStr(rs!impuesto_ventas)
  End If
   
  
  'Verificar si existe precio especificado en combo y si es asi cambiarlo
  strSQL = "select monto from pv_producto_precios where cod_producto = '" _
         & rs!Cod_Producto & "' and cod_precio = '" & txtPrecioId.Text & "'"
  rs.Close
  rs.CursorLocation = adUseServer
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then
    vGrid.col = 5
    vGrid.Text = CStr(rs!Monto)
  End If
End If
rs.Close

'Si la caja no puede modificar los precios, bloquea las columnas de precios e Impuestos
If Not gCajas.ModPrecios Then
    vGrid.Row = fila
    vGrid.col = 5
    vGrid.Lock = True
    
    vGrid.Row = fila
    vGrid.col = 6
    vGrid.Lock = True
    
    vGrid.Row = fila
    vGrid.col = 7
    vGrid.Lock = True

    vGrid.Row = fila
    vGrid.col = 8
    vGrid.Lock = True
End If

End Sub

Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
vError:
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula.Text = gBusquedas.Resultado
  txtNombre.Text = gBusquedas.Resultado2
End If

End Sub



Private Sub txtPrecioDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAgenteId.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_precio,descripcion from pv_tipos_precios"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtPrecioId.Text = gBusquedas.Resultado
  txtPrecioDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtPrecioId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPrecioDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_precio"
  gBusquedas.Orden = "cod_precio"
  gBusquedas.Consulta = "select cod_precio,descripcion from pv_tipos_precios"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtPrecioId.Text = gBusquedas.Resultado
  txtPrecioDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtPrecioId_LostFocus()
txtPrecioDesc.Text = fxSIFCCodigos("D", txtPrecioId, "Precios")
End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

vFilaNote = Row
vGrid.Row = Row
vGrid.col = 1
txtNotas.Text = vGrid.CellNote
txtNotas.SetFocus

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(9) As Variant, x As Integer

'Abrir Nueva Linea
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    Call sbCalculaTotales
  End If
End If

'Buscar Articulo
If vGrid.ActiveCol = 2 And KeyCode = vbKeyReturn Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  Call sbConsultaArticulo(vGrid.ActiveRow, vGrid.ActiveCol, vGrid.Text)
End If

'Buscar Bodegas
If vGrid.ActiveCol = 4 And KeyCode = vbKeyReturn Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
  vGrid.CellNote = fxSIFCCodigos("D", vGrid.Text, "BODEGAS")
  vGrid.TextTip = TextTipFixed
End If


'Consular Bodegas
If vGrid.ActiveCol = 4 And KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_bodega"
   gBusquedas.Orden = "cod_bodega"
   gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
   gBusquedas.Filtro = " and permite_salidas = 1"
   frmBusquedas.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 4
   vGrid.Text = gBusquedas.Resultado
   vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
   vGrid.CellNote = gBusquedas.Resultado2
   vGrid.TextTip = TextTipFixed
End If

'Consular Articulo
If vGrid.ActiveCol = 2 And KeyCode = vbKeyF4 Then
   frmBusquedaArticulos.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 2
   vGrid.Text = gBusquedas.Resultado
End If



'Borrar una linea
If KeyCode = vbKeyDelete Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = vGrid.MaxCols
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 2 To vGrid.MaxCols
        vGrid.col = x
        vTemp(x) = vGrid.Text
     Next x

     vGrid.Row = lng
     For x = 2 To vGrid.MaxCols
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
Dim curIC As Currency, curDes As Currency

On Error GoTo vError
'Calcula Total
Select Case vGrid.ActiveCol
  Case 3, 5, 6, 7, 8
    vGrid.Row = vGrid.ActiveRow
    vGrid.col = 3
    curCantidad = CCur(vGrid.Text)
    vGrid.col = 5
    curPrecio = CCur(vGrid.Text)
    
    vGrid.col = 6
    curDes = CCur(vGrid.Text) / 100
    
    vGrid.col = 7
    curIC = CCur(vGrid.Text) / 100
    
    vGrid.col = 8
    curIV = CCur(vGrid.Text) / 100
    
    curDes = (curPrecio * curCantidad) * curDes
    curIC = ((curPrecio * curCantidad) - curDes) * curIC
    curIV = ((curPrecio * curCantidad) - curDes) * curIV
    
    vGrid.col = 9
    vGrid.Text = (curPrecio * curCantidad) - curDes + curIC + curIV
   
   Call sbCalculaTotales
  
  Case Else 'No Aplica
End Select
vError:
End Sub



Private Sub vGrid_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim pRow As Long


Select Case True
  Case NewRow = -1 And Row = 0
       pRow = mRow
  Case NewRow = -1 And Row > 0
        pRow = Row
        mRow = Row
        
  Case NewRow > 0
       pRow = NewRow
       mRow = NewRow
End Select

vGrid.Row = pRow
vFilaNote = pRow

vGrid.col = 1
txtNotas.Text = vGrid.CellNote

vGrid.col = 2
lblProdDesc.Caption = vGrid.CellNote

vGrid.col = 4
lblBodegaDesc.Caption = vGrid.CellNote

End Sub
