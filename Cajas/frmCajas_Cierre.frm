VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmCajas_Cierre 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cierre de Caja"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   14445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox chkAP_Compartida 
      Height          =   255
      Left            =   4200
      TabIndex        =   33
      Top             =   480
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Compartida ?"
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
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.TabControl tcFP 
      Height          =   5895
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   6015
      _Version        =   1441793
      _ExtentX        =   10610
      _ExtentY        =   10398
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
      ItemCount       =   1
      Item(0).Caption =   "Resumen"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "txtDP_DetalleCheques"
      Item(0).Control(1)=   "txtDP_ADepositar"
      Item(0).Control(2)=   "txtTotalEfectivoCheque"
      Item(0).Control(3)=   "txtTotalDeposito"
      Item(0).Control(4)=   "txtTotalCaja"
      Item(0).Control(5)=   "vGrid"
      Item(0).Control(6)=   "Label1(15)"
      Item(0).Control(7)=   "Label1(14)"
      Item(0).Control(8)=   "Label1(2)"
      Item(0).Control(9)=   "Label1(0)"
      Item(0).Control(10)=   "Label1(10)"
      Item(0).Control(11)=   "Label1(8)"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   2535
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   5655
         _Version        =   524288
         _ExtentX        =   9970
         _ExtentY        =   4466
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
         MaxCols         =   3
         ScrollBars      =   2
         SpreadDesigner  =   "frmCajas_Cierre.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalEfectivoCheque 
         Height          =   330
         Left            =   3240
         TabIndex        =   36
         Top             =   3240
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalCaja 
         Height          =   330
         Left            =   3240
         TabIndex        =   37
         Top             =   3600
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDP_ADepositar 
         Height          =   330
         Left            =   3240
         TabIndex        =   38
         Top             =   4560
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDP_DetalleCheques 
         Height          =   330
         Left            =   3240
         TabIndex        =   39
         Top             =   4920
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalDeposito 
         Height          =   330
         Left            =   3240
         TabIndex        =   40
         Top             =   5400
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total valores registrados..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   720
         TabIndex        =   15
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Depósitos Registrados..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   720
         TabIndex        =   14
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Efectivo + Cheques ..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   13
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto a Depositar..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   720
         TabIndex        =   12
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Información para Deposito..:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   120
         TabIndex        =   11
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Efectivo a Depositar + CKs..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   720
         TabIndex        =   10
         Top             =   4920
         Width           =   2655
      End
   End
   Begin XtremeSuiteControls.GroupBox gbAcciones 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   7320
      Width           =   14415
      _Version        =   1441793
      _ExtentX        =   25426
      _ExtentY        =   1720
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin MSComctlLib.Toolbar tblAplicar 
         Height          =   315
         Left            =   5280
         TabIndex        =   6
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         ButtonWidth     =   1931
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aplicar"
               Key             =   "Aplicar"
               Object.ToolTipText     =   "Aplica la Cierre "
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reporte"
               Key             =   "Reporte"
               Object.ToolTipText     =   "Reportes del Cierre"
               ImageIndex      =   9
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Resumen"
                     Text            =   "Resumen"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Cierre"
                     Text            =   "Informe de Cierre"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Movimientos"
                     Text            =   "Movimientos"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "Cancelar"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   12240
      Top             =   840
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13920
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Cierre.frx":0620
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Cierre.frx":0EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Cierre.frx":1031
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Cierre.frx":112D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Cierre.frx":1247
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Cierre.frx":134C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Cierre.frx":1CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Cierre.frx":26A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Cierre.frx":2E90
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Cierre.frx":364C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5895
      Left            =   6000
      TabIndex        =   8
      Top             =   1440
      Width           =   8415
      _Version        =   1441793
      _ExtentX        =   14843
      _ExtentY        =   10398
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
      SelectedItem    =   2
      Item(0).Caption =   "Efectivo + Mov"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "vGridDt"
      Item(0).Control(1)=   "Label1(9)"
      Item(0).Control(2)=   "txtTotalEfectivo"
      Item(1).Caption =   "Efectivo a Depositar"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "txtSobranteFaltante"
      Item(1).Control(1)=   "txtDP_EF_ADepositar"
      Item(1).Control(2)=   "txtDP_Detallado"
      Item(1).Control(3)=   "vGridDP"
      Item(1).Control(4)=   "Label1(17)"
      Item(1).Control(5)=   "Label1(16)"
      Item(1).Control(6)=   "Label1(13)"
      Item(2).Caption =   "Depósitos en Cajas"
      Item(2).ControlCount=   8
      Item(2).Control(0)=   "txtDep_No"
      Item(2).Control(1)=   "txtDep_Monto"
      Item(2).Control(2)=   "cboDP_Cuenta"
      Item(2).Control(3)=   "tlbDP"
      Item(2).Control(4)=   "lswDP"
      Item(2).Control(5)=   "Label1(1)"
      Item(2).Control(6)=   "Label1(11)"
      Item(2).Control(7)=   "Label1(12)"
      Begin XtremeSuiteControls.ListView lswDP 
         Height          =   3495
         Left            =   120
         TabIndex        =   47
         Top             =   2160
         Width           =   8175
         _Version        =   1441793
         _ExtentX        =   14420
         _ExtentY        =   6165
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
      End
      Begin FPSpreadADO.fpSpread vGridDt 
         Height          =   4695
         Left            =   -69880
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   8055
         _Version        =   524288
         _ExtentX        =   14208
         _ExtentY        =   8281
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
         MaxCols         =   4
         SpreadDesigner  =   "frmCajas_Cierre.frx":3CD6
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridDP 
         Height          =   4335
         Left            =   -69880
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   8175
         _Version        =   524288
         _ExtentX        =   14420
         _ExtentY        =   7646
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "frmCajas_Cierre.frx":47E9
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbDP 
         Height          =   270
         Left            =   5280
         TabIndex        =   22
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   476
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Aplicar"
               Object.ToolTipText     =   "Aplica la Depósito"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Eliminar"
               Object.ToolTipText     =   "Eliminar Depósito"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Anular"
               Object.ToolTipText     =   "Anular Depósito"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalEfectivo 
         Height          =   330
         Left            =   -65560
         TabIndex        =   35
         Top             =   5400
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDP_Detallado 
         Height          =   330
         Left            =   -64480
         TabIndex        =   41
         Top             =   5040
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSobranteFaltante 
         Height          =   330
         Left            =   -69520
         TabIndex        =   42
         Top             =   5280
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDP_EF_ADepositar 
         Height          =   330
         Left            =   -64480
         TabIndex        =   43
         Top             =   5400
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboDP_Cuenta 
         Height          =   330
         Left            =   1680
         TabIndex        =   44
         Top             =   480
         Width           =   6135
         _Version        =   1441793
         _ExtentX        =   10821
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
      Begin XtremeSuiteControls.FlatEdit txtDep_No 
         Height          =   330
         Left            =   1680
         TabIndex        =   45
         Top             =   960
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
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
      Begin XtremeSuiteControls.FlatEdit txtDep_Monto 
         Height          =   330
         Left            =   5400
         TabIndex        =   46
         Top             =   960
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Monto ..:"
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
         Index           =   12
         Left            =   4080
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Depósito No..:"
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
         Index           =   11
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta..:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Total del Detalle..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   -66520
         TabIndex        =   21
         Top             =   5040
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Efectivo a Depositar..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   -66520
         TabIndex        =   20
         Top             =   5400
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Sobrante/Faltante..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   -69520
         TabIndex        =   19
         Top             =   5040
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Efectivo Detallado..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   -67600
         TabIndex        =   17
         Top             =   5400
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtAP_Fecha 
      Height          =   330
      Left            =   8520
      TabIndex        =   26
      Top             =   120
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAP_EnUso_Fecha 
      Height          =   330
      Left            =   8520
      TabIndex        =   27
      Top             =   480
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAP_Numero 
      Height          =   330
      Left            =   2040
      TabIndex        =   28
      Top             =   120
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAP_Vence 
      Height          =   330
      Left            =   2040
      TabIndex        =   29
      Top             =   480
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAP_EnUso_Usuario 
      Height          =   330
      Left            =   10560
      TabIndex        =   30
      Top             =   480
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtAP_Usuario 
      Height          =   330
      Left            =   10560
      TabIndex        =   31
      Top             =   120
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtAP_Estado 
      Height          =   330
      Left            =   4080
      TabIndex        =   32
      Top             =   120
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.ComboBox cboDivisa 
      Height          =   330
      Left            =   2040
      TabIndex        =   34
      Top             =   840
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Divisa "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima Apertura"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   7080
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "En Uso"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   7080
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento ?"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmCajas_Cierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vCierreCiego As Boolean
Dim mTotalCK As Currency, mTotalEF As Currency


Private Sub sbAperturaCarga()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

'Carga el Tipo de Cierre
strSQL = "select CIERRE_TIPO From CAJAS_DEFINICION" _
       & " where cod_Caja = '" & ModuloCajas.mCaja & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Cierre_Tipo = "C" Then
    vCierreCiego = True
    txtTotalEfectivoCheque.PasswordChar = "*"
    txtTotalCaja.PasswordChar = "*"
    
    vGridDt.Sheet = 2
    vGridDt.SheetVisible = False
    
Else
    vCierreCiego = False
End If
rs.Close


strSQL = "select *, Case when Estado = 'A' then 'Abierta' else 'Cerrada' end as 'Estado'" _
       & " from Cajas_Aperturas_Main" _
       & " where cod_Caja = '" & ModuloCajas.mCaja & "' and Cod_Apertura = " & ModuloCajas.mApertura
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  txtAP_Numero.Text = rs!Cod_Apertura
  txtAP_Estado.Text = rs!Estado
  txtAP_Fecha.Text = rs!Apertura_Fecha
  txtAP_Usuario.Text = rs!Apertura_Usuario
  
  txtAP_EnUso_Fecha.Text = rs!En_Uso_Fecha & ""
  txtAP_EnUso_Usuario.Text = rs!En_Uso_Usuario & ""
  
  txtAP_Vence.Text = rs!Apertura_Vence & ""
  chkAP_Compartida.Value = rs!Apertura_Compartida
  
Else
  txtAP_Numero.Text = "0"
  txtAP_Estado.Text = ""
  txtAP_Fecha.Text = ""
  txtAP_Usuario.Text = ""

  txtAP_EnUso_Fecha.Text = ""
  txtAP_EnUso_Usuario.Text = ""
  
  txtAP_Vence.Text = ""
  chkAP_Compartida.Value = vbUnchecked
End If
rs.Close

'Actualiza datos
Call cboDivisa_Click

End Sub


Private Sub cboDivisa_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTotalEF As Currency, vTotalDP As Currency, vTotalCajas As Currency

If vPaso Or cboDivisa.ListCount = 0 Then Exit Sub

Me.MousePointer = vbHourglass

'Inicializa datos

tcMain.Item(0).Selected = True



mTotalCK = 0
mTotalEF = 0

vPaso = True
vTotalEF = 0
vTotalDP = 0
vTotalCajas = 0

vGrid.MaxRows = 0
vGridDP.MaxRows = 0

vGridDt.Sheet = 1
vGridDt.MaxRows = 0
vGridDt.Sheet = 0
vGridDt.MaxRows = 0

txtTotalCaja.Text = 0
txtTotalEfectivo.Text = 0
txtTotalDeposito.Text = 0
txtTotalEfectivoCheque.Text = 0

txtSobranteFaltante.Text = 0
txtDP_EF_ADepositar.Text = 0

txtDep_No.Text = ""
txtDep_Monto.Text = 0

txtDP_ADepositar.Text = 0
txtDP_Detallado.Text = 0



'1. Carga el Resumen de las formas de pago y Totaliza para el cierre
strSQL = "exec spCajas_CierreFPTotal '" & ModuloCajas.mCaja & "'," & txtAP_Numero.Text _
      & ", '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"

Call OpenRecordSet(rs, strSQL)
With vGrid
Do While Not rs.EOF
  .MaxRows = .MaxRows + 1
  .Row = .MaxRows
  .col = 1
  .CellTag = rs!Tipo
  .col = 2
  .CellTag = rs!Cod_Forma_Pago
  .Text = rs!Descripcion
  .col = 3
  .CellTag = rs!cod_Divisa
  .Text = Format(rs!Monto, "Standard")
     
  vTotalCajas = vTotalCajas + rs!Monto
  
    Select Case rs!Tipo
        Case "E", "C" 'Efectivo y Cheques
            If rs!Efectivo = 1 Then
                vTotalEF = vTotalEF + rs!Monto
            End If
                 
            If vCierreCiego Then
               .Text = "0.00"
            End If
                 
        Case "B", "T" 'Depositos y Tarjetas
            vTotalDP = vTotalDP + rs!Monto
    End Select
  
  If rs!Tipo = "E" Then
    mTotalEF = mTotalEF + rs!Monto
  End If
  
  If rs!Tipo = "C" Then
    mTotalCK = mTotalCK + rs!Monto
  End If
  
  rs.MoveNext
Loop
rs.Close
End With

txtTotalCaja.Text = Format(vTotalCajas, "Standard")

If Not vCierreCiego Then
    txtTotalCaja.ToolTipText = "Efectivo + Cheques  .: " + Format(vTotalEF, "Standard") & vbCrLf & vbCrLf _
                             & "Depósito + Tarjetas .: " + Format(vTotalDP, "Standard")
End If

txtTotalEfectivoCheque.Text = Format(vTotalEF, "Standard")

'2.1 Carga las Denominaciones del Efectivo
strSQL = "exec spCajas_CierreEFDetalle '" & ModuloCajas.mCaja & "'," & txtAP_Numero.Text _
      & ", '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "','E'"

Call OpenRecordSet(rs, strSQL)
With vGridDt
  .Sheet = 0
  vTotalCajas = 0
  
Do While Not rs.EOF
  .MaxRows = .MaxRows + 1
  .Row = .MaxRows
  .col = 1
  .Text = rs!Tipo
  .col = 2
  .CellTag = Format(rs!Denominacion, "###,###,###,##0.0000")
  .Text = rs!Descripcion
  .col = 3
  .Text = CStr(rs!Cantidad)
  .col = 4
  .Text = Format(rs!Monto, "Standard")
     
  vTotalCajas = vTotalCajas + rs!Monto
  
  rs.MoveNext
Loop
rs.Close
 
End With

txtTotalEfectivo.Text = Format(vTotalCajas, "Standard")


'2.2 Carga las Denominaciones del Efectivo para Efectos del Depósito
strSQL = "exec spCajas_CierreEFDetalle '" & ModuloCajas.mCaja & "'," & txtAP_Numero.Text _
      & ", '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "','D'"

Call OpenRecordSet(rs, strSQL)
With vGridDP
  .Sheet = 0
  vTotalCajas = 0
  
Do While Not rs.EOF
  .MaxRows = .MaxRows + 1
  .Row = .MaxRows
  .col = 1
  .Text = rs!Tipo
  .col = 2
  .CellTag = Format(rs!Denominacion, "###,###,###,##0.0000")
  .Text = rs!Descripcion
  .col = 3
  .Text = CStr(rs!Cantidad)
  .col = 4
  .Text = Format(rs!Monto, "Standard")
     
  vTotalCajas = vTotalCajas + rs!Monto
  
  rs.MoveNext
Loop
rs.Close
 
End With

txtDP_Detallado.Text = Format(vTotalCajas, "Standard")



'3. Carga Deposito de Cierre de Caja (Supuesto: Solo un deposito por Divisa)
Call sbConsultaDepositosRegistrados

'4. Calcular el Total a Depositar
strSQL = "select SI_EFECTIVO From CAJAS_APERTURAS_CIERRES" _
       & " where COD_CAJA = '" & ModuloCajas.mCaja & "' and COD_APERTURA = " & txtAP_Numero.Text _
       & " and COD_DIVISA = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtDP_ADepositar.Text = Format(CCur(txtTotalEfectivoCheque.Text), "Standard")
    txtSobranteFaltante.Text = Format((CCur(txtTotalEfectivo.Text) - rs!SI_EFECTIVO) - mTotalEF, "Standard")
    txtDP_EF_ADepositar.Text = Format(mTotalEF + CCur(txtSobranteFaltante.Text), "Standard")
End If
rs.Close

If IsNumeric(txtDP_Detallado.Text) Then
    txtDP_DetalleCheques.Text = Format(CCur(txtDP_Detallado.Text) + mTotalCK, "Standard")
End If
vPaso = False


txtTotalDeposito.SetFocus

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbConsultaDepositosRegistrados()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curMonto As Currency


On Error GoTo vError

strSQL = "exec spCajas_CierreDepositoDivisa '" & ModuloCajas.mCaja & "'," & txtAP_Numero.Text _
      & ", '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"

With lswDP.ListItems

    Call OpenRecordSet(rs, strSQL)
    curMonto = 0
    .Clear
    Do While Not rs.EOF
        Set itmX = .Add(, , rs!DP_Numero)
            itmX.SubItems(1) = Format(rs!Monto, "Standard")
            
        If rs!Estado = 1 Or rs!Estado = 2 Then
           curMonto = curMonto + rs!Monto
            itmX.SubItems(2) = IIf((rs!Estado = 1), "Activado", "En Bancos")
        
            txtDep_No.Text = rs!DP_Numero & ""
            txtDep_Monto.Text = Format(rs!Monto, "Standard")
            Call sbCboAsignaDato(cboDP_Cuenta, rs!itmX, True)
        Else
            itmX.SubItems(2) = "Anulado"
        End If
        
        itmX.SubItems(3) = Trim(rs!DP_Cuenta)
        itmX.SubItems(4) = rs!BancoDesc

       rs.MoveNext
    
    Loop
    rs.Close

    txtTotalDeposito.Text = Format(curMonto, "Standard")

End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub Form_Activate()
 vModulo = 5
End Sub

Private Sub Form_Load()
Dim strSQL As String
', rs As New ADODB.Recordset

vModulo = 5

ModuloCajas.mCierreActiva = True


With lswDP.ColumnHeaders
  .Clear
  .Add , , "No.Depósito", 2500
  .Add , , "Monto", 2500, vbRightJustify
  .Add , , "Estado", 1800, vbCenter
  .Add , , "Cuenta", 2500
  .Add , , "Banco", 2500
End With

       
vPaso = True

strSQL = "Select rtrim(cod_divisa) as 'IdX', rtrim(Descripcion) as 'Itmx'" _
       & " from CNTX_DIVISAS where COD_CONTABILIDAD = " & GLOBALES.gEnlace _
       & "  Order by DIVISA_LOCAL desc,COD_DIVISA"
Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)


strSQL = "exec spCajas_DepositosCuentasBancarias"
Call sbCbo_Llena_New(cboDP_Cuenta, strSQL, False, True)


vPaso = False


'Call OpenRecordSet(rs, strSQL)
'
'
'Do While Not rs.EOF
'  cboDP_Cuenta.AddItem rs!Cta
'  cboDP_Cuenta.Text = rs!Cta
'  rs.MoveNext
'Loop
'rs.Close

 
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub tblAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer

On Error GoTo vError

Select Case Button.Key
   Case "Aplicar"
      
      If fxValidaCierreCaja Then
        i = MsgBox("Esta seguro que desea realizar el cierre?", vbYesNo)
        If i = vbYes Then
           Call sbAplicaCierre
        End If
      End If
      
   Case "Cancelar"
     Unload Me
     Exit Sub
End Select

Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxCuentaDevolucion(vCaja As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select cod_cuenta_dev from cajas_definicion where cod_caja = '" & vCaja & "'"
Call OpenRecordSet(rs, strSQL)
fxCuentaDevolucion = rs!Cod_Cuenta_Dev
rs.Close

End Function

Private Function fxValidaCierreCaja() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

vMensaje = ""
fxValidaCierreCaja = True

strSQL = "Select count(*) as existe from cajas_aperturas_main where cod_caja  = '" & ModuloCajas.mCaja & "' and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
   vMensaje = vMensaje & " - Esta Caja NO se encuentra abierta!" & vbCrLf
   Exit Function
End If
rs.Close



If Len(vMensaje) > 0 Then
  fxValidaCierreCaja = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbAplicaCierre()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, iDiasVence As Long

On Error GoTo vError

If Not fxValidaCierreCaja() Then Exit Sub

strSQL = "exec spCajas_CierreCajaMain '" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura _
       & ",'" & glogon.Usuario & "',0"
Call ConectionExecute(strSQL)

'Reporte de Cierre
Call sbCajasCierreReportes(ModuloCajas.mCaja, ModuloCajas.mApertura, "Cierre", vCierreCiego)

MsgBox "Cierre de Caja Realizado Satisfactoriamente!", vbInformation

Unload Me

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub



Private Sub tblAplicar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim strSQL As String

'Aplica un Cierre Preliminar de los datos para ver el informe
If txtAP_Estado.Text = "Abierta" Then
   strSQL = "exec spCajas_CierreCajaMain '" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura _
       & ",'" & glogon.Usuario & "',1"
   Call ConectionExecute(strSQL)
End If

Select Case ButtonMenu.Key
  Case "Resumen"
     Call sbCajasCierreReportes(ModuloCajas.mCaja, ModuloCajas.mApertura, "Resumen", vCierreCiego)
  Case "Cierre"
     Call sbCajasCierreReportes(ModuloCajas.mCaja, ModuloCajas.mApertura, "Cierre", vCierreCiego)
  Case "Movimientos"
     Call sbCajasCierreReportes(ModuloCajas.mCaja, ModuloCajas.mApertura, "Movimientos", vCierreCiego)
End Select

End Sub

Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

Call sbCajaInicial

End Sub

Private Sub sbCajaInicial()
Dim strSQL As String

'Paso 1: Si la Caja no está abierta (Llamar pantalla de login de Caja)
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   Call sbFormsCall("frmCajas_Acceso", vbModal, , , False, Me)
End If

'Paso 2: Si despues del Login de Caja permanece sin Apertura Salir
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   MsgBox "No se ha indicado ninguna caja con Apertura disponible?", vbExclamation
   Unload Me
   Exit Sub
End If

Me.Caption = "Cierre de Caja       ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

Call sbAperturaCarga

End Sub

Private Sub sbCalculaEfectivo(Optional pTipo As String = "E")
Dim i As Integer, curMonto As Currency
Dim strSQL As String

curMonto = 0

If pTipo = "E" Then
    With vGridDt
        For i = 1 To .MaxRows
            .Row = i
            .col = 4
            curMonto = curMonto + CCur(.Text)
        Next i
    End With
    
    txtTotalEfectivo.Text = Format(curMonto, "Standard")
End If


If pTipo = "D" Then
    With vGridDP
        For i = 1 To .MaxRows
            .Row = i
            .col = 4
            curMonto = curMonto + CCur(.Text)
        Next i
    End With
    
    txtDP_Detallado.Text = Format(curMonto, "Standard")
    txtDP_DetalleCheques.Text = Format(mTotalCK + curMonto, "Standard")
    
End If


End Sub

Private Sub tlbDP_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, i As Integer


On Error GoTo vError


'ALTER proc [dbo].[spCajas_CierreRegistraDeposito](@Caja varchar(10), @Apertura int, @Divisa varchar(10), @Monto dec(18,2)
'                                         ,  @DP_Numero varchar(30) , @DP_Cuenta varchar(30), @Usuario varchar(30)
'                                         ,  @DP_Banco  smallint = 0, @Estado smallint = 1)

Select Case Button.Key
  Case "Aplicar"

        strSQL = ""
        
        If Trim(txtDep_No.Text) = "" Then strSQL = strSQL & " - Indique el número del depósito!" & vbCrLf
        If Not IsNumeric(txtDep_Monto.Text) Then strSQL = strSQL & " - Indique un monto válido para el depósito!" & vbCrLf
        
        strSQL = "exec spCajas_CierreRegistraDeposito '" & ModuloCajas.mCaja & "'," & txtAP_Numero.Text _
               & ",'" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'," & CCur(txtDep_Monto.Text) _
               & ",'" & txtDep_No.Text & "','','" & glogon.Usuario & "', " & cboDP_Cuenta.ItemData(cboDP_Cuenta.ListIndex) & ", 1"
        Call ConectionExecute(strSQL)
        
        txtTotalDeposito.Text = Format(CCur(txtDep_Monto.Text), "Standard")
        
        MsgBox "Deposito Registrado Satisfactoriamente...!", vbInformation




  Case "Anular"
        With lswDP.ListItems
            For i = 1 To .Count
               If .Item(i).Checked Then
                    strSQL = "exec spCajas_CierreRegistraDeposito '" & ModuloCajas.mCaja & "'," & txtAP_Numero.Text _
                           & ",'" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'," & CCur(.Item(i).SubItems(1)) _
                           & ",'" & .Item(i).Text & "','" & .Item(i).SubItems(3) & "','" & glogon.Usuario & "', 0, 0"
                    Call ConectionExecute(strSQL)
               End If
            Next i
        End With
        MsgBox "Deposito Marcados! Fueron Anulados...!", vbInformation
        
  Case "Eliminar"
        With lswDP.ListItems
            For i = 1 To .Count
               If .Item(i).Checked Then
                    strSQL = "exec spCajas_CierreRegistraDeposito '" & ModuloCajas.mCaja & "'," & txtAP_Numero.Text _
                           & ",'" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'," & CCur(.Item(i).SubItems(1)) _
                           & ",'" & .Item(i).Text & "','" & .Item(i).SubItems(3) & "','" & glogon.Usuario & "', 0 ,2"
                    Call ConectionExecute(strSQL)
               End If
            Next i
        End With
        MsgBox "Deposito Marcados! Fueron Eliminados...!", vbInformation
End Select


txtDep_No.Text = ""
txtDep_Monto.Text = "0"

Call sbConsultaDepositosRegistrados

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub





Private Sub txtDep_Monto_GotFocus()

On Error GoTo vError

txtDep_Monto.Text = CCur(txtDep_Monto.Text)

Exit Sub

vError:

End Sub

Private Sub txtDep_Monto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And txtDep_No.Enabled Then
    txtDep_No.SetFocus
End If

End Sub

Private Sub txtDep_Monto_LostFocus()
On Error GoTo vError

txtDep_Monto.Text = Format(CCur(txtDep_Monto.Text), "Standard")

Exit Sub

vError:

End Sub

Private Sub txtDep_No_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And txtDep_Monto.Enabled Then
    txtDep_Monto.SetFocus
End If

End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String


If vPaso Or vGrid.MaxRows = 0 Then Exit Sub

vGrid.Row = Row
vGrid.col = 2

strSQL = "exec spCajas_CierreFPDetalle '" & ModuloCajas.mCaja & "'," & txtAP_Numero.Text _
       & ",'" & cboDivisa.ItemData(cboDivisa.ListIndex) & "','" & vGrid.CellTag & "'"
       
vGridDt.Sheet = 2
Call sbCargaGridFps7(vGridDt, 11, strSQL, True, 2, 0)

End Sub


Private Sub vGridDP_KeyUp(KeyCode As Integer, Shift As Integer)
Dim vDenominacion As Currency, vMonto As Currency

On Error GoTo vError

vGridDP.Row = vGridDP.ActiveRow
vGridDP.col = 3


If IsNumeric(vGridDP.Text) Then
    vGridDP.col = 2
    vDenominacion = CCur(vGridDP.CellTag)
    
    vGridDP.col = 3
    vMonto = vDenominacion * CCur(vGridDP.Text)

    vGridDP.col = 4
    vGridDP.CellTag = "Cambia"
    vGridDP.Text = Format(vMonto, "Standard")
    
    Call sbCalculaEfectivo("D")
End If

Exit Sub

vError:

End Sub

Private Sub vGridDP_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim vDenominacion As Currency, vCantidad As Integer
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass


With vGridDP
 .Row = Row
 .col = 4
 If col = 3 And .CellTag = "Cambia" Then
    .col = 2
    vDenominacion = CCur(.CellTag)
    .col = 3
    vCantidad = .Text
    
    strSQL = "exec spCajas_CierreRegistraEFDenominacion '" & ModuloCajas.mCaja & "'," & txtAP_Numero.Text & ",'" & cboDivisa.ItemData(cboDivisa.ListIndex) _
           & "'," & vDenominacion & "," & vCantidad & ",'D'"
    Call ConectionExecute(strSQL)
        
    .col = 4
    .CellTag = ""
 End If
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub vGridDt_KeyUp(KeyCode As Integer, Shift As Integer)
Dim vDenominacion As Currency, vMonto As Currency

On Error GoTo vError

If vGridDt.ActiveSheet = 2 Then Exit Sub

vGridDt.Sheet = vGridDt.ActiveSheet
vGridDt.Row = vGridDt.ActiveRow
vGridDt.col = 3


If IsNumeric(vGridDt.Text) Then
    vGridDt.col = 2
    vDenominacion = CCur(vGridDt.CellTag)
    
    vGridDt.col = 3
    vMonto = vDenominacion * CCur(vGridDt.Text)

    vGridDt.col = 4
    vGridDt.CellTag = "Cambia"
    vGridDt.Text = Format(vMonto, "Standard")
    
    Call sbCalculaEfectivo("E")
End If

Exit Sub
vError:

End Sub

Private Sub vGridDt_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim vDenominacion As Currency, vCantidad As Integer
Dim strSQL As String

On Error GoTo vError

If vPaso Or vGridDt.Sheet = 2 Then Exit Sub

Me.MousePointer = vbHourglass


With vGridDt
 .Row = Row
 .col = 4
 If col = 3 And .CellTag = "Cambia" Then
    .col = 2
    vDenominacion = CCur(.CellTag)
    .col = 3
    vCantidad = .Text
    
    strSQL = "exec spCajas_CierreRegistraEFDenominacion '" & ModuloCajas.mCaja & "'," & txtAP_Numero.Text & ",'" & cboDivisa.ItemData(cboDivisa.ListIndex) _
           & "'," & vDenominacion & "," & vCantidad
    Call ConectionExecute(strSQL)
        
    .col = 4
    .CellTag = ""
 End If
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
