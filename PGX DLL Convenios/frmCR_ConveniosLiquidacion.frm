VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmCR_ConveniosLiquidacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidacion de Convenios [Ordenes de Pago]"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11250
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   7.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosLiquidacion.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosLiquidacion.frx":07B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosLiquidacion.frx":0EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosLiquidacion.frx":1585
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosLiquidacion.frx":1F71
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosLiquidacion.frx":272D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosLiquidacion.frx":2F0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   8355
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Fecha de Registro"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Usuario de Registro"
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9840
      TabIndex        =   0
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.TabControl ssTabX 
      Height          =   6972
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   10812
      _Version        =   1310722
      _ExtentX        =   19071
      _ExtentY        =   12298
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
      Item(0).Caption =   "Ordenes"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "Label3(16)"
      Item(0).Control(2)=   "feLineas"
      Item(1).Caption =   "Liquidación"
      Item(1).ControlCount=   46
      Item(1).Control(0)=   "txtFlotante"
      Item(1).Control(1)=   "txtDevoluciones"
      Item(1).Control(2)=   "txtContrato"
      Item(1).Control(3)=   "txtEstado"
      Item(1).Control(4)=   "txtReservas"
      Item(1).Control(5)=   "dtpInicio"
      Item(1).Control(6)=   "txtNetoLiquidar"
      Item(1).Control(7)=   "txtPlanesAhorros"
      Item(1).Control(8)=   "txtComisionCreditos"
      Item(1).Control(9)=   "txtComisionRecaudacion"
      Item(1).Control(10)=   "txtCreditos"
      Item(1).Control(11)=   "txtRecaudacion"
      Item(1).Control(12)=   "txtCargos"
      Item(1).Control(13)=   "txtNotas"
      Item(1).Control(14)=   "txtDocumento"
      Item(1).Control(15)=   "txtOrden"
      Item(1).Control(16)=   "dtpCorte"
      Item(1).Control(17)=   "dtpVencimiento"
      Item(1).Control(18)=   "Label3(15)"
      Item(1).Control(19)=   "Label3(14)"
      Item(1).Control(20)=   "Label3(12)"
      Item(1).Control(21)=   "Label3(13)"
      Item(1).Control(22)=   "Label3(10)"
      Item(1).Control(23)=   "Label3(11)"
      Item(1).Control(24)=   "Label3(9)"
      Item(1).Control(25)=   "Label3(8)"
      Item(1).Control(26)=   "Label3(7)"
      Item(1).Control(27)=   "Label3(5)"
      Item(1).Control(28)=   "Label3(4)"
      Item(1).Control(29)=   "Label3(3)"
      Item(1).Control(30)=   "Label3(2)"
      Item(1).Control(31)=   "Label3(1)"
      Item(1).Control(32)=   "Label3(0)"
      Item(1).Control(33)=   "Label3(19)"
      Item(1).Control(34)=   "Label3(20)"
      Item(1).Control(35)=   "txtCargosCxP"
      Item(1).Control(36)=   "Label3(6)"
      Item(1).Control(37)=   "chkComisionInformativa"
      Item(1).Control(38)=   "chkCreditosAnulados"
      Item(1).Control(39)=   "btnAjuste"
      Item(1).Control(40)=   "btnDetalle"
      Item(1).Control(41)=   "gbBarra"
      Item(1).Control(42)=   "txtIVA"
      Item(1).Control(43)=   "Label3(17)"
      Item(1).Control(44)=   "txtIVA_Referencia"
      Item(1).Control(45)=   "Label3(18)"
      Item(2).Caption =   "Flotante al Cobro"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   6252
         Left            =   -69760
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1310722
         _ExtentX        =   18224
         _ExtentY        =   11028
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.GroupBox gbBarra 
         Height          =   480
         Left            =   -66280
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1310722
         _ExtentX        =   4466
         _ExtentY        =   847
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin MSComctlLib.Toolbar tlbOrden 
            Height          =   264
            Left            =   20
            TabIndex        =   52
            Top             =   144
            Width           =   2292
            _ExtentX        =   4048
            _ExtentY        =   476
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   7
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Nuevo"
                  Object.ToolTipText     =   "Nueva Orden de Pago"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Guardar"
                  Object.ToolTipText     =   "Guarda Cambios"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Calcular"
                  Object.ToolTipText     =   "Actualizar Cálculos"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Cerrar"
                  Object.ToolTipText     =   "Cerrar la Orden de Pago"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Boleta"
                  ImageIndex      =   5
                  Style           =   5
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   3
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Boleta"
                        Text            =   "Boleta"
                     EndProperty
                     BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Asiento"
                        Text            =   "Boleta + Asiento"
                     EndProperty
                     BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Informe"
                        Text            =   "Informe detallado"
                     EndProperty
                  EndProperty
               EndProperty
            EndProperty
         End
      End
      Begin XtremeSuiteControls.PushButton btnAjuste 
         Height          =   312
         Left            =   -67120
         TabIndex        =   49
         Top             =   2040
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   550
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCR_ConveniosLiquidacion.frx":3871
      End
      Begin XtremeSuiteControls.CheckBox chkComisionInformativa 
         Height          =   972
         Left            =   -69640
         TabIndex        =   23
         Top             =   4080
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1310722
         _ExtentX        =   4254
         _ExtentY        =   1714
         _StockProps     =   79
         Caption         =   "Esta Orden utiliza la comision como dato informativo no aplica rebajo de la misma al convenio"
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
         Alignment       =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5892
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   10572
         _Version        =   524288
         _ExtentX        =   18648
         _ExtentY        =   10393
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
         MaxCols         =   457
         SpreadDesigner  =   "frmCR_ConveniosLiquidacion.frx":3F64
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit feLineas 
         Height          =   252
         Left            =   9960
         TabIndex        =   25
         Top             =   6600
         Width           =   732
         _Version        =   1310722
         _ExtentX        =   1291
         _ExtentY        =   444
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Text            =   "50"
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkCreditosAnulados 
         Height          =   852
         Left            =   -69640
         TabIndex        =   26
         Top             =   5160
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1310722
         _ExtentX        =   4254
         _ExtentY        =   1503
         _StockProps     =   79
         Caption         =   "Calcula el Cobro por Créditos Anulados en el periodo de fechas de la Liquidación?"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCargos 
         Height          =   312
         Left            =   -62800
         TabIndex        =   28
         Top             =   2040
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtRecaudacion 
         Height          =   312
         Left            =   -62800
         TabIndex        =   29
         Top             =   2400
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDevoluciones 
         Height          =   312
         Left            =   -62800
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtCreditos 
         Height          =   312
         Left            =   -62800
         TabIndex        =   31
         Top             =   3120
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtComisionRecaudacion 
         Height          =   312
         Left            =   -62800
         TabIndex        =   32
         Top             =   3480
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtComisionCreditos 
         Height          =   312
         Left            =   -62800
         TabIndex        =   33
         Top             =   3840
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtReservas 
         Height          =   312
         Left            =   -62800
         TabIndex        =   34
         Top             =   4200
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtCargosCxP 
         Height          =   312
         Left            =   -62800
         TabIndex        =   35
         Top             =   4560
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPlanesAhorros 
         Height          =   312
         Left            =   -62800
         TabIndex        =   36
         Top             =   4920
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNetoLiquidar 
         Height          =   312
         Left            =   -62800
         TabIndex        =   37
         Top             =   5760
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtFlotante 
         Height          =   312
         Left            =   -62800
         TabIndex        =   38
         Top             =   6120
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtOrden 
         Height          =   312
         Left            =   -68800
         TabIndex        =   41
         Top             =   600
         Visible         =   0   'False
         Width           =   972
         _Version        =   1310722
         _ExtentX        =   1714
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   312
         Left            =   -67840
         TabIndex        =   42
         Top             =   600
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1310722
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   312
         Left            =   -62200
         TabIndex        =   43
         Top             =   600
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
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
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   672
         Left            =   -68800
         TabIndex        =   44
         Top             =   1200
         Visible         =   0   'False
         Width           =   8532
         _Version        =   1310722
         _ExtentX        =   15049
         _ExtentY        =   1185
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
      End
      Begin XtremeSuiteControls.FlatEdit txtContrato 
         Height          =   312
         Left            =   -68800
         TabIndex        =   45
         Top             =   3120
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1310722
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   -68800
         TabIndex        =   46
         Top             =   2040
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1310722
         _ExtentX        =   2773
         _ExtentY        =   556
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
         Left            =   -68800
         TabIndex        =   47
         Top             =   2400
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1310722
         _ExtentX        =   2773
         _ExtentY        =   556
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpVencimiento 
         Height          =   312
         Left            =   -68800
         TabIndex        =   48
         Top             =   3480
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1310722
         _ExtentX        =   2773
         _ExtentY        =   556
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton btnDetalle 
         Height          =   312
         Left            =   -60640
         TabIndex        =   50
         ToolTipText     =   "Información detallado del rubro"
         Top             =   2040
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   550
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCR_ConveniosLiquidacion.frx":4BF9
      End
      Begin XtremeSuiteControls.FlatEdit txtIVA 
         Height          =   312
         Left            =   -62800
         TabIndex        =   53
         Top             =   5400
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtIVA_Referencia 
         Height          =   312
         Left            =   -62800
         TabIndex        =   55
         Top             =   6480
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin VB.Label Label3 
         Caption         =   "(i) IVA Referencia (Venta)"
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
         Index           =   18
         Left            =   -65680
         TabIndex        =   56
         Top             =   6480
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label3 
         Caption         =   "(-) I.V.A."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   252
         Index           =   17
         Left            =   -65680
         TabIndex        =   54
         Top             =   5400
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "No. Líneas"
         Height          =   252
         Index           =   16
         Left            =   8520
         TabIndex        =   24
         Top             =   6600
         Width           =   1212
      End
      Begin VB.Label Label3 
         Caption         =   "(i) Flotante por Cobrar"
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
         Index           =   15
         Left            =   -65680
         TabIndex        =   22
         Top             =   6120
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label3 
         Caption         =   "(+) Devoluciones"
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
         Index           =   14
         Left            =   -65680
         TabIndex        =   21
         Top             =   2760
         Visible         =   0   'False
         Width           =   2412
      End
      Begin VB.Label Label3 
         Caption         =   "No. Contrato"
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
         Index           =   12
         Left            =   -69760
         TabIndex        =   20
         Top             =   3120
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "Vencimiento"
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
         Index           =   13
         Left            =   -69760
         TabIndex        =   19
         Top             =   3480
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "(-) Reservas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   252
         Index           =   10
         Left            =   -65680
         TabIndex        =   18
         Top             =   4200
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label3 
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
         Index           =   11
         Left            =   -69760
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "No. Documento"
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
         Left            =   -63640
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label3 
         Caption         =   "Total a Pagar"
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
         Left            =   -65680
         TabIndex        =   15
         Top             =   5760
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label3 
         Caption         =   "(-) Planes de Ahorros"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   252
         Index           =   7
         Left            =   -65680
         TabIndex        =   14
         Top             =   4920
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label3 
         Caption         =   "(-) Comisión s/Nuevos Créditos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   252
         Index           =   5
         Left            =   -65680
         TabIndex        =   13
         Top             =   3840
         Visible         =   0   'False
         Width           =   2652
      End
      Begin VB.Label Label3 
         Caption         =   "(+) Giro Nuevos Créditos"
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
         Left            =   -65680
         TabIndex        =   12
         Top             =   3120
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.Label Label3 
         Caption         =   "(-) Comisión s/Recaudación"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   252
         Index           =   3
         Left            =   -65680
         TabIndex        =   11
         Top             =   3480
         Visible         =   0   'False
         Width           =   2652
      End
      Begin VB.Label Label3 
         Caption         =   "(+) Recaudación de Cuentas"
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
         Left            =   -65680
         TabIndex        =   10
         Top             =   2400
         Visible         =   0   'False
         Width           =   2412
      End
      Begin VB.Label Label3 
         Caption         =   "(i) Cargos de Formalización"
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
         Left            =   -65680
         TabIndex        =   9
         Top             =   2040
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.Label Label3 
         Caption         =   "Corte"
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
         Left            =   -69760
         TabIndex        =   8
         Top             =   2400
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio"
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
         Index           =   19
         Left            =   -69760
         TabIndex        =   7
         Top             =   2040
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label3 
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
         Index           =   20
         Left            =   -69760
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label3 
         Caption         =   "(-) Cargos de CxP"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   252
         Index           =   6
         Left            =   -65680
         TabIndex        =   5
         Top             =   4560
         Visible         =   0   'False
         Width           =   2052
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2400
      TabIndex        =   39
      Top             =   240
      Width           =   1092
      _Version        =   1310722
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3480
      TabIndex        =   40
      Top             =   240
      Width           =   6252
      _Version        =   1310722
      _ExtentX        =   11028
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
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   4080
      X2              =   9480
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Convenio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1212
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmCR_ConveniosLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vScroll As Boolean, vEdita As Boolean, vPaso As Boolean
Dim mConvenio As String, mOrden As Long, mOrdenEstado As String
Dim mFechaInicioMin As Date, mFechaCorteMax As Date
Dim mComisionRecaudacion As Currency, mComisionNuevosCrd As Currency
Dim mFechaCorteGeneral As Date, mComisionInformativa As Integer, mFlotante As Currency, mCreditosAnulados As Integer

Private Sub sbLimpiaDatos(Optional pTodo As Boolean = True)
  mOrden = 0
  mOrdenEstado = ""
  
  txtOrden.Text = ""
  txtEstado.Text = ""
  txtDocumento.Text = ""
  txtNotas.Text = ""
  
  chkComisionInformativa.Value = mComisionInformativa
  chkCreditosAnulados.Value = mCreditosAnulados
  
  txtFlotante.Text = Format(mFlotante, "Standard")
  
  If pTodo Then
    dtpInicio.MinDate = "1980/01/01"
    dtpInicio.Value = mFechaCorteGeneral
    dtpCorte.MaxDate = dtpInicio.Value
    dtpCorte.Value = DateAdd("d", -1, dtpInicio.Value)
    dtpVencimiento.Value = dtpInicio.Value
  End If
  
  btnDetalle.Top = 2040
  btnDetalle.Tag = ""
  
  
  txtCargos.Text = "0.00"
  txtRecaudacion.Text = "0.00"
  txtCreditos.Text = "0.00"
  
  txtComisionRecaudacion.Text = "0.00"
  txtComisionCreditos.Text = "0.00"
  txtReservas.Text = "0.00"
  txtCargosCxP.Text = "0.00"
  txtPlanesAhorros.Text = "0.00"
  txtDevoluciones.Text = "0.00"
  
  txtNetoLiquidar.Text = "0.00"
  txtIVA_Referencia.Text = "0.00"
  
  btnDetalle.Tag = "Cargos"
  btnDetalle.Top = txtCargos.Top
 
End Sub

Private Sub sbLimpiaMontos()
 
 txtCargos.Text = "0.00"
 txtRecaudacion.Text = "0.00"
 txtDevoluciones.Text = "0.00"
 txtCreditos.Text = "0.00"
 
 txtComisionRecaudacion.Text = "0.00"
 txtComisionCreditos.Text = "0.00"
 txtReservas.Text = "0.00"
 txtCargosCxP.Text = "0.00"
 txtPlanesAhorros.Text = "0.00"
 
 txtNetoLiquidar.Text = "0.00"
 
End Sub

Private Sub btnAjuste_Click()
dtpInicio.Enabled = True
dtpInicio.Tag = dtpInicio.Value
End Sub

Private Sub btnDetalle_Click()
  If mConvenio = "" Or mOrden = 0 Then Exit Sub
  
  GLOBALES.gTag = mConvenio
  GLOBALES.gTag2 = mOrden
  GLOBALES.gTag3 = btnDetalle.Tag
  
  Select Case btnDetalle.Tag
  
    Case "CargosCxP"
     ' Call sbFormsCall("frmCR_ConveniosCargosCxP", vbModal, , , False, Me)
      
       frmCR_ConveniosCargosCxP.Show vbModal, Me
      Call sbRecalculandoOrden
    
    Case "PlanesAhorro"
'      Call sbFormsCall("frmCR_ConveniosAhorros", vbModal, , , False, Me)
      frmCR_ConveniosAhorros.Show vbModal, Me
      Call sbRecalculandoOrden
    
    Case "Reservas"
    
    
    Case "Asiento"
        GLOBALES.gTag3 = txtDescripcion.Text
        frmCR_ConveniosOrdenAsiento.Show vbModal, Me
      
      
    Case Else
      Call sbFormsCall("frmCR_ConveniosLiqDetalle", vbModal, , , False, Me)
     
 End Select

End Sub

Private Sub feLineas_Change()

If IsNumeric(feLineas) Then
  Call sbConsultaOrdenes
End If

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset
    
On Error GoTo vError

If vScroll Then

    strSQL = "select Top 1 COD_CONVENIO from CRD_CONVENIOS"
        
    If Len(txtCodigo.Text) > 0 Then
        If FlatScrollBar.Value = 1 Then
           strSQL = strSQL & " where COD_CONVENIO > '" & txtCodigo.Text & "' order by COD_CONVENIO asc"
        Else
           strSQL = strSQL & " where COD_CONVENIO < '" & txtCodigo.Text & "' order by COD_CONVENIO desc"
        End If
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
        Call sbConsulta(rs!COD_CONVENIO)
    Else
        Call sbLimpiaDatos
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
  vModulo = 16
End Sub

Private Sub Form_Load()
vModulo = 16

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture
mFechaCorteGeneral = Format(DateAdd("d", -1, fxFechaServidor), "yyyy/mm/dd")


mComisionInformativa = 0
mCreditosAnulados = 0
mFlotante = 0


With lsw.ColumnHeaders
   .Add , , "Cargo", 2500
   .Add , , "Monto", 1850, vbRightJustify
   .Add , , "Registro", 1850, vbCenter
   .Add , , "Inicio Cobro", 1850, vbCenter
   .Add , , "Documento", 1500
   .Add , , "Usuario", 1500
End With


Call sbLimpiaDatos

ssTabX.Item(0).Selected = True


vGrid.MaxCols = 19
vGrid.MaxRows = 0

Call Formularios(Me)
Call RefrescaTags(Me)

   
End Sub

Private Sub sbConsultaOrdenes()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vEstado As String, i As Integer

If mConvenio = "" Then Exit Sub

On Error GoTo vError

ssTabX.Item(0).Selected = True

strSQL = "select Top " & feLineas.Text & " COD_ORDEN, COD_CONVENIO, ESTADO, NOTAS, FECHA_INICIO,FECHA_CORTE, RECAUDACION_CUOTAS" _
       & ", RECAUDACION_CARGOS, NUEVOS_CREDITOS, RETENCION_AHORROS, COMISIONES_RECAUDACION, COMISIONES_NUEVOS_CREDITOS" _
       & ", TOTAL_RESERVA, TOTAL_REBAJOS_CXP, TOTAL_PAGAR, DOCUMENTO, REGISTRO_USUARIO, REGISTRO_FECHA " _
       & " from CRD_CONVENIOS_ORDENES where COD_CONVENIO = '" & mConvenio & "' order by COD_ORDEN desc"
  
Call OpenRecordSet(rs, strSQL)
     
Me.MousePointer = vbHourglass

vPaso = True

With vGrid

.MaxRows = 1
Do While Not rs.EOF
 
 Select Case rs!estado
   Case "A"
     vEstado = "Abierta"
   Case "C"
     vEstado = "Cerrada"
 End Select
    

 .Row = .MaxRows
 For i = 2 To .MaxCols
  .Col = i
  Select Case i
    Case 2 'Orden
      .Text = rs!cod_orden
    Case 3 'Convenio
      .Text = rs!COD_CONVENIO
    Case 4 ' Estado
      .Text = vEstado
    Case 5 'Notas
      .Text = rs!NOTAS
    Case 6 'Fecha Inicio
      .Text = Format(rs!FECHA_INICIO, "dd/mm/yyyy")
    Case 7 'Fecha Corte
      .Text = Format(rs!FECHA_CORTE, "dd/mm/yyyy")
    Case 8 'Recaudación
      .Text = IIf(IsNull(rs!Recaudacion_Cuotas), 0, Format(rs!Recaudacion_Cuotas, "Standard"))
    Case 9 'Cargos
      .Text = IIf(IsNull(rs!Recaudacion_Cargos), 0, Format(rs!Recaudacion_Cargos, "Standard"))
    Case 10 'Monto Nuevos Créditos
      .Text = IIf(IsNull(rs!Nuevos_Creditos), 0, Format(rs!Nuevos_Creditos, "Standard"))
    Case 11 'Retenciones Ahorros
      .Text = IIf(IsNull(rs!Retencion_Ahorros), 0, Format(rs!Retencion_Ahorros, "Standard"))
    Case 12 'Comision Cuotas
      .Text = IIf(IsNull(rs!Comisiones_Recaudacion), 0, Format(rs!Comisiones_Recaudacion, "Standard"))
    Case 13 'Comision Nuevos Creditos
      .Text = IIf(IsNull(rs!COMISIONES_NUEVOS_CREDITOS), 0, Format(rs!COMISIONES_NUEVOS_CREDITOS, "Standard"))
    Case 14 'Reservas
      .Text = Format(rs!Total_Reserva, "Standard")
    Case 15 'Total Rebajos por Proveedor
      .Text = IIf(IsNull(rs!Total_Rebajos_CXP), 0, Format(rs!Total_Rebajos_CXP, "Standard"))
    Case 16 'Monto Total
      .Text = IIf(IsNull(rs!Total_Pagar), 0, Format(rs!Total_Pagar, "Standard"))
    Case 17 'Documento
      .Text = rs!Documento
    Case 18 'Usuario
      .Text = rs!registro_usuario
    Case 19 'Fecha Registro
      .Text = rs!registro_Fecha
  End Select
 Next i

 .MaxRows = .MaxRows + 1
 rs.MoveNext
Loop

End With
rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub ssTabX_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index < 2 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError
Me.MousePointer = vbHourglass

strSQL = "exec spConvenio_CxP_CargosFlotantes '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Descripcion)
     itmX.SubItems(1) = Format(rs!Monto, "Standard")
     itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
     itmX.SubItems(3) = Format(rs!Vence, "dd/mm/yyyy")
     itmX.SubItems(4) = rs!Documento & ""
     itmX.SubItems(5) = rs!Usuario & ""
 rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub




Private Sub tlbOrden_ButtonClick(ByVal Button As MSComctlLib.Button)

If mConvenio = "" Then
   MsgBox "No se ha indicado ningun convenio?", vbExclamation
   Exit Sub
End If

Select Case Button.Key
   
  Case "Nuevo" 'Crea una nueva orden, valida que la anterior este cerrada
    If Not fxValidaUltimaOrdenCerrada Then
        Call sbLimpiaDatos(True)
    Else
        Call sbLimpiaDatos(False)
    End If
    vEdita = False
    
    
  Case "Calcular" 'Calcula los montos utilizados en la liquidación
     If mOrden = 0 Or mOrdenEstado = "C" Then
        MsgBox "La orden no existe o ya fue cerrada...verifique!", vbExclamation
        Exit Sub
     End If
     
     Call sbRecalculandoOrden
                
  Case "Guardar" 'Cambia de estado a la orden
    If Not fxValida Then Exit Sub
    
    Call sbGuardaOrden

     
  Case "Cerrar" 'Cierra la orden de pago y envia a CxP
  
    If Not fxValida Then Exit Sub
    
    If mOrdenEstado = "C" Then
       MsgBox "La orden ya se encuentra cerrada, Verifique!!", vbExclamation
       Exit Sub
    End If
    
    If mConvenio = "" Or mOrden = 0 Then
       MsgBox "La Orden no existe...Verifique!!", vbExclamation
       Exit Sub
    End If
    
    'Actualiza Datos
    Call sbGuardaOrden
    
    If CCur(txtNetoLiquidar.Text) < 0 Then
      MsgBox " - No se puede Cerrar! Monto a Liquidar es NEGATIVO revisar las deducciones!", vbExclamation
      Exit Sub
    End If
    
    'Cierra la Orden válida
    Call sbCierraOrden
     
End Select
   
End Sub

Private Sub tlbOrden_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim vTitulo As String, vSubTitulo As String, vFiltro As String, vTemp As String, i As Integer
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vTitulo = ""
vSubTitulo = ""
strSQL = ""
Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = False
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = False
 .WindowState = crptMaximized
 
 .Connect = glogon.ConectRPT
 .WindowTitle = "Convenios: Boleta de Liquidación"
 
 vTitulo = "Boleta de Liquidación"


'Filtro Estado
 strSQL = "{CRD_CONVENIOS_ORDENES.COD_CONVENIO} = '" & txtCodigo.Text _
        & "' AND {CRD_CONVENIOS_ORDENES.COD_ORDEN} = " & txtOrden.Text
 
 Select Case ButtonMenu.Key
   Case "Boleta"
     .ReportFileName = SIFGlobal.fxPathReportes("Convenios_OrdenesBoletaRsm.rpt")
    vSubTitulo = "Resumen"
   Case "Informe"
     .ReportFileName = SIFGlobal.fxPathReportes("Convenios_OrdenesBoletaInforme.rpt")
     vSubTitulo = "Boleta + Informe detallado"
   Case "Asiento"
     .ReportFileName = SIFGlobal.fxPathReportes("Convenios_OrdenesBoletaAsiento.rpt")
     vSubTitulo = "Boleta + Asiento"
 End Select
 
 .Formulas(0) = "fxTitulo= '" & vTitulo & "'"
 .Formulas(1) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(2) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(4) = "fxUsuario='Usuario: " & glogon.Usuario & "'"

 .SelectionFormula = strSQL

 Select Case ButtonMenu.Key
   Case "Boleta"
   Case "Informe"
       
       .SubreportToChange = "sbCreditosNuevos"
       .StoredProcParam(0) = txtCodigo.Text
       .StoredProcParam(1) = txtOrden.Text
       .StoredProcParam(2) = "D"
       
       .SubreportToChange = "sbRecaudacion"
       .StoredProcParam(0) = txtCodigo.Text
       .StoredProcParam(1) = txtOrden.Text
       .StoredProcParam(2) = "D"
       
       .SubreportToChange = "sbDevoluciones"
       .StoredProcParam(0) = txtCodigo.Text
       .StoredProcParam(1) = txtOrden.Text
       .StoredProcParam(2) = "D"
 
  Case "Asiento"
       .SubreportToChange = "sbAsiento"
       .StoredProcParam(0) = txtCodigo.Text
       .StoredProcParam(1) = txtOrden.Text
 
 End Select


 .PrintReport

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtCargos_GotFocus()
  btnDetalle.Tag = "Cargos"
  btnDetalle.Top = txtCargos.Top
End Sub

Private Sub txtCargosCxP_GotFocus()
  btnDetalle.Tag = "CargosCxP"
  btnDetalle.Top = txtCargosCxP.Top
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Consulta = "select COD_CONVENIO,DESCRIPCION" _
                            & " from CRD_CONVENIOS"
        gBusquedas.Columna = "COD_CONVENIO"
        gBusquedas.Orden = "DESCRIPCION"
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        frmBusquedas.Show vbModal

        txtCodigo.Text = Trim(gBusquedas.Resultado)
        txtDescripcion.Text = Trim(gBusquedas.Resultado2)

        Call sbConsulta(txtCodigo.Text)
    
    End If
        
   If KeyCode = vbKeyReturn Then
      Call sbConsulta(txtCodigo.Text)
   End If
   

    
End Sub

Private Function fxConvenioDescrip(vConvenio As String)
Dim strSQL As String, rs As New ADODB.Recordset

 strSQL = "select Cod_Convenio,Descripcion" _
        & " from Crd_Convenios" _
        & " where Cod_Convenio = '" & vConvenio & "'"
 Call OpenRecordSet(rs, strSQL)
 If Not rs.EOF Then
    fxConvenioDescrip = rs!Descripcion
 Else
    fxConvenioDescrip = ""
 End If
 rs.Close
    
End Function


Public Sub sbConsultaExterna(pConvenio As String, Optional pOrden As Long = 0)

Call sbConsulta(pConvenio)
If pOrden > 0 Then
    Call sbConsultaOrden(pOrden)
End If

End Sub

Private Sub sbConsulta(pConvenio As String)
Dim strSQL As String, rs As New ADODB.Recordset
On Error GoTo vError

Call sbLimpiaDatos

strSQL = "Select COD_CONVENIO,DESCRIPCION,CONTRATO_NUMERO,isnull(FECHA_VENCIMIENTO,dateadd(YEAR,1,dbo.MyGetdate())) as 'FechaVence' " _
       & ",Porc_Comision_Creditos,Porc_Comision_Recaudacion, Fecha_Inicio,COMISION_INFORMATIVA,COBRA_CREDITOS_ANULADOS" _
       & ",dbo.fxConvenio_CxP_CargosFlotantes('" & pConvenio & "') AS 'CxP_Flotante'" _
       & " from CRD_CONVENIOS where COD_CONVENIO = '" & pConvenio & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  mConvenio = Trim(rs!COD_CONVENIO)
  txtCodigo.Text = rs!COD_CONVENIO
  txtDescripcion.Text = rs!Descripcion
  txtContrato.Text = rs!CONTRATO_NUMERO
  dtpVencimiento.Value = Format(rs!FechaVence, "dd/mm/yyyy")
  
  mFechaCorteMax = rs!FechaVence
  mFechaInicioMin = rs!FECHA_INICIO
  mComisionNuevosCrd = rs!Porc_Comision_Creditos
  mComisionRecaudacion = rs!Porc_Comision_Recaudacion
  mComisionInformativa = rs!COMISION_INFORMATIVA
  mFlotante = rs!CxP_flotante
  mCreditosAnulados = rs!COBRA_CREDITOS_ANULADOS
  
  Call sbConsultaOrdenes


Else
  mConvenio = ""
  txtDescripcion.Text = ""
  txtContrato.Text = ""
End If

rs.Close

Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Long

Dim vMensaje As String

vMensaje = ""
fxValida = True

If mOrdenEstado = "C" Then
   vMensaje = vMensaje & vbCrLf & "- No se puede modificar una orden Cerrada."
End If

If mConvenio = "" Then
   vMensaje = vMensaje & vbCrLf & "- No se ha consultado ningun convenio?"
End If

If vEdita And mOrden = 0 Then
   vMensaje = vMensaje & vbCrLf & "- No se puede Editar porque no se ha consultado ninguna orden!"
End If

If Trim(txtNotas.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Debe indicar una Nota para la Orden, verifique!"

If dtpCorte.Value > mFechaCorteGeneral Then
    vMensaje = vMensaje & vbCrLf & " - La fecha de corte de la remesa no puede ser igual o mayor a : " & Format(mFechaCorteGeneral, "dd/mm/yyyy") & ", verifique!"
End If

'Valida la Fecha de Inicio
If IsNumeric(txtOrden.Text) Then
   i = txtOrden.Text
Else
   i = 0
End If

If CCur(txtIVA_Referencia.Text) > CCur(txtNetoLiquidar.Text) Then
    vMensaje = vMensaje & vbCrLf & " - El Monto a Pagar es menor al monto del IVA de Referencia!"
End If

strSQL = "select  count(*) as 'Existe'" _
       & " From CRD_CONVENIOS_ORDENES" _
       & " where COD_CONVENIO = '" & mConvenio & "'" _
       & "   and cod_orden not in(" & i & ")" _
       & "   and '" & Format(dtpInicio.Value, "yyyy/mm/dd") & "' BETWEEN FECHA_INICIO AND FECHA_CORTE"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
    vMensaje = vMensaje & vbCrLf & " - Conflicto de Fechas con otra Liquidación, revise el rango de fechas!"
End If
rs.Close

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub txtComisionCreditos_GotFocus()
  btnDetalle.Tag = "ComisionCreditos"
  btnDetalle.Top = txtComisionCreditos.Top
End Sub

Private Sub txtComisionRecaudacion_GotFocus()
   btnDetalle.Tag = "ComisionRecaudacion"
   btnDetalle.Top = txtComisionRecaudacion.Top
End Sub

Private Sub txtCreditos_GotFocus()
  btnDetalle.Tag = "NuevosCreditos"
  btnDetalle.Top = txtCreditos.Top
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Consulta = "select COD_CONVENIO,DESCRIPCION" _
                            & " from CRD_CONVENIOS"
        gBusquedas.Columna = "Descripcion"
        gBusquedas.Orden = "DESCRIPCION"
        frmBusquedas.Show vbModal
        
        txtCodigo.Text = Trim(gBusquedas.Resultado)
        txtDescripcion.Text = Trim(gBusquedas.Resultado2)
      
          Call sbConsulta(txtCodigo.Text)
        
    End If
        
End Sub


Private Sub txtDevoluciones_GotFocus()
  btnDetalle.Tag = "Devoluciones"
  btnDetalle.Top = txtDevoluciones.Top
End Sub



Private Sub txtNetoLiquidar_GotFocus()
  btnDetalle.Tag = "Asiento"
  btnDetalle.Top = txtNetoLiquidar.Top
End Sub

Private Sub txtOrden_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then

    gBusquedas.Consulta = "select Cod_Orden, Estado, Notas from CRD_CONVENIOS_ORDENES"
    gBusquedas.Columna = "Cod_Orden"
    gBusquedas.Orden = "Cod_Orden"
    gBusquedas.Filtro = " and cod_Convenio = '" & mConvenio & "'"
    frmBusquedas.Show vbModal
    
    txtOrden.Text = gBusquedas.Resultado
    Call sbConsultaOrden(txtOrden.Text)
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   Call sbConsultaOrden(txtOrden.Text)
End If

End Sub

Private Function fxValidaUltimaOrdenCerrada() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As Boolean

On Error GoTo vError
  
  vResultado = True
  
  dtpInicio.MinDate = "1980/01/01"
  dtpInicio.Value = mFechaInicioMin
  dtpCorte.MaxDate = mFechaCorteMax
  dtpCorte.Value = mFechaCorteMax
  
  strSQL = "select top 1 COD_ORDEN,ESTADO,FECHA_CORTE, dateadd(d,1,Fecha_Corte) as 'Inicio', dateadd(d,-1,dbo.MyGetdate()) as 'Corte' " _
         & " From CRD_CONVENIOS_ORDENES " _
         & " where COD_CONVENIO = '" & mConvenio & "'" _
         & " Order by FECHA_CORTE desc"
         
  Call OpenRecordSet(rs, strSQL)

  If Not rs.EOF And Not rs.BOF Then
     If rs!estado = "A" Then
        vResultado = False
        MsgBox "La Orden " & rs!cod_orden & " no se encuentra cerrada,verifique!!", vbExclamation
     Else
        dtpInicio.MinDate = "1980/01/01"
        dtpInicio.Value = rs!Inicio
        dtpInicio.MinDate = rs!Inicio
        
        If rs!Corte > mFechaCorteMax Then
            dtpCorte.Value = mFechaCorteMax
            MsgBox "El Contrato de este convenio se encuentra vencido, la fecha de corte fue puesta al vencimiento!!", vbExclamation
        Else
            dtpCorte.Value = rs!Corte
            dtpCorte.MaxDate = rs!Corte
        End If
     End If
  End If
       
  If dtpCorte.Value > mFechaCorteGeneral Then
      dtpCorte.Value = mFechaCorteGeneral
  End If
  If rs.RecordCount > 0 Then
     dtpInicio.Enabled = False
  Else
     dtpInicio.Enabled = True
  End If
     
  rs.Close

fxValidaUltimaOrdenCerrada = vResultado

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbConsultaOrden(pOrden As Long)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select COD_ORDEN, COD_CONVENIO, REGISTRO_USUARIO, REGISTRO_FECHA, ESTADO, NOTAS, FECHA_INICIO" _
       & ", FECHA_CORTE, RECAUDACION_CUOTAS, RECAUDACION_CARGOS, NUEVOS_CREDITOS, RETENCION_AHORROS, COMISIONES_RECAUDACION" _
       & ", TOTAL_RESERVA, TOTAL_REBAJOS_CXP, TOTAL_PAGAR, DOCUMENTO,COMISIONES_NUEVOS_CREDITOS " _
       & ", RECAUDACION_DEVOLUCION,APL_COM_INFORMATIVA, FLOTANTE_ALCOBRO, COBRA_CREDITOS_ANULADOS" _
       & ", IVA_TOTAL, IVA_COM_CREDITOS, IVA_COM_RECAUDO, IVA_REFERENCIA_CRD" _
       & " from CRD_CONVENIOS_ORDENES where COD_ORDEN = '" & pOrden & "' and COD_CONVENIO = '" & mConvenio & "'"
Call OpenRecordSet(rs, strSQL)
     
If Not rs.EOF Then
 vEdita = True
 
 mOrden = rs!cod_orden
 
 txtOrden.Text = rs!cod_orden
 txtDocumento.Text = rs!Documento
 txtNotas.Text = rs!NOTAS
 
 mOrdenEstado = rs!estado
 
 'Inicial:
 chkComisionInformativa.Value = IIf(IsNull(rs!APL_COM_INFORMATIVA), 0, rs!APL_COM_INFORMATIVA)
 chkCreditosAnulados.Value = IIf(IsNull(rs!COBRA_CREDITOS_ANULADOS), 0, rs!COBRA_CREDITOS_ANULADOS)
 
 txtFlotante.Text = IIf(IsNull(rs!FLOTANTE_ALCOBRO), "0.00", Format(rs!FLOTANTE_ALCOBRO, "Standard"))
 
 Select Case mOrdenEstado
    Case "A"
      txtEstado.Text = "Abierta"
      
      'Actualiza Flotante y Comision Informativa
      txtFlotante.Text = Format(mFlotante, "Standard")
      chkComisionInformativa.Value = mComisionInformativa
      
    Case "C"
      txtEstado.Text = "Cerrada"
    Case "T"
      txtEstado.Text = "Tramitada"
 End Select

 dtpInicio.Value = rs!FECHA_INICIO
 dtpCorte.Value = Format(rs!FECHA_CORTE, "yyyy/mm/dd")
 
 txtCargos.Text = IIf(IsNull(rs!Recaudacion_Cargos), "0.00", Format(rs!Recaudacion_Cargos, "Standard"))
 txtRecaudacion.Text = IIf(IsNull(rs!Recaudacion_Cuotas), "0.00", Format(rs!Recaudacion_Cuotas, "Standard"))
 txtDevoluciones.Text = IIf(IsNull(rs!RECAUDACION_DEVOLUCION), "0.00", Format(rs!RECAUDACION_DEVOLUCION, "Standard"))
 txtCreditos.Text = IIf(IsNull(rs!Nuevos_Creditos), "0.00", Format(rs!Nuevos_Creditos, "Standard"))
 txtComisionRecaudacion.Text = IIf(IsNull(rs!Comisiones_Recaudacion), "0.00", Format(rs!Comisiones_Recaudacion, "Standard"))
 txtComisionCreditos.Text = IIf(IsNull(rs!COMISIONES_NUEVOS_CREDITOS), "0.00", Format(rs!COMISIONES_NUEVOS_CREDITOS, "Standard"))
 txtReservas.Text = IIf(IsNull(rs!Total_Reserva), "0.00", Format(rs!Total_Reserva, "Standard"))
 txtCargosCxP.Text = IIf(IsNull(rs!Total_Rebajos_CXP), "0.00", Format(rs!Total_Rebajos_CXP, "Standard"))
 txtPlanesAhorros.Text = IIf(IsNull(rs!Retencion_Ahorros), "0.00", Format(rs!Retencion_Ahorros, "Standard"))
 txtNetoLiquidar.Text = IIf(IsNull(rs!Total_Pagar), "0.00", Format(rs!Total_Pagar, "Standard"))

 txtIVA.Text = IIf(IsNull(rs!IVA_TOTAL), "0.00", Format(rs!IVA_TOTAL, "Standard"))
 txtIVA_Referencia.Text = IIf(IsNull(rs!IVA_REFERENCIA_CRD), "0.00", Format(rs!IVA_REFERENCIA_CRD, "Standard"))


 StatusBar.Panels.Item(1).Text = "Usuario: " & rs!registro_usuario
 StatusBar.Panels.Item(2).Text = "Fecha Registro: " & rs!registro_Fecha
   
 ssTabX.Item(1).Selected = True
   
End If

rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Call sbLimpiaDatos

End Sub

Private Sub txtPlanesAhorros_GotFocus()
  btnDetalle.Tag = "PlanesAhorro"
  btnDetalle.Top = txtPlanesAhorros.Top
End Sub

Private Sub txtPlanesAhorros_LostFocus()
  txtNetoLiquidar.Text = 0
  txtNetoLiquidar.Text = Format((CCur(txtCargos.Text) + CCur(txtRecaudacion.Text) + CCur(txtCreditos)) - (CCur(txtComisionRecaudacion.Text) + CCur(txtComisionCreditos.Text) + CCur(txtReservas.Text) + CCur(txtCargosCxP.Text) + CCur(txtPlanesAhorros.Text)), "Standard")
End Sub

Private Sub txtRecaudacion_GotFocus()
  btnDetalle.Tag = "Recaudacion"
  btnDetalle.Top = txtRecaudacion.Top
End Sub

Private Sub txtReservas_GotFocus()
  btnDetalle.Tag = "Reservas"
  btnDetalle.Top = txtReservas.Top
End Sub

Private Sub sbRecalculandoOrden()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spConvenio_Orden_Detalle '" & mConvenio & "'," & mOrden
Call ConectionExecute(strSQL)


Call sbConsultaOrden(mOrden)

Me.MousePointer = vbDefault

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

'Guarda la orden de pago
Private Sub sbGuardaOrden()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curTotalBruto As Currency

On Error GoTo vError


'vMontoReserva = fxVerificaReserva
'
'If vMontoReserva > txtReservas.Text Then
'    vMontoReserva = txtReservas.Text
'End If
   
Me.MousePointer = vbHourglass

curTotalBruto = CCur(txtCargos.Text) + CCur(txtRecaudacion.Text) + CCur(txtCreditos.Text) + CCur(txtDevoluciones.Text)

If vEdita Then
  strSQL = "update CRD_CONVENIOS_ORDENES set NOTAS = '" & txtNotas.Text & "', FECHA_INICIO = '" _
         & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00', FECHA_CORTE = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
         & ", RECAUDACION_CUOTAS = " & CCur(txtRecaudacion.Text) & ", RECAUDACION_CARGOS = " & CCur(txtCargos.Text) _
         & ", NUEVOS_CREDITOS = " & CCur(txtCreditos.Text) & ", RETENCION_AHORROS = " & CCur(txtPlanesAhorros.Text) _
         & ", COMISIONES_RECAUDACION = " & CCur(txtComisionRecaudacion) & ", COMISIONES_NUEVOS_CREDITOS = " & CCur(txtComisionCreditos.Text) _
         & ", TOTAL_RESERVA = " & CCur(txtReservas.Text) & ", TOTAL_ORDEN = " & curTotalBruto _
         & ", TOTAL_REBAJOS_CXP = " & CCur(txtCargosCxP.Text) & ", TOTAL_PAGAR = " & CCur(txtNetoLiquidar.Text) _
         & ", DOCUMENTO = '" & txtDocumento.Text & "',PORC_COMISION_RECAUDACION = " & mComisionRecaudacion _
         & ",PORC_COMISION_CREDITOS = " & mComisionNuevosCrd & ", RECAUDACION_DEVOLUCION = " & CCur(txtDevoluciones.Text) _
         & ",APL_COM_INFORMATIVA = " & chkComisionInformativa.Value & ", FLOTANTE_ALCOBRO = " & CCur(txtFlotante.Text) _
         & ",COBRA_CREDITOS_ANULADOS = " & chkCreditosAnulados.Value _
         & " Where COD_CONVENIO = '" & mConvenio & "' and COD_ORDEN = '" & mOrden & "'"
  Call ConectionExecute(strSQL)
      
  'Guarda los datos del detalle de la orden
  strSQL = "exec spConvenio_Orden_Detalle '" & mConvenio & "'," & mOrden
  Call ConectionExecute(strSQL)
      
  Call Bitacora("Modifica", "Liquidación Convenio.: " & mConvenio & " - Orden.:" & mOrden)
  
  MsgBox "Información actualizada satisfactoriamente!!!", vbInformation

Else
  
  strSQL = "select isnull(max(cod_Orden),0) + 1  Orden From CRD_CONVENIOS_ORDENES " _
         & " where COD_CONVENIO = '" & mConvenio & "'"
  Call OpenRecordSet(rs, strSQL)
    mOrden = rs!Orden
  rs.Close
  
  strSQL = "Insert CRD_CONVENIOS_ORDENES (COD_ORDEN, COD_CONVENIO, REGISTRO_USUARIO, REGISTRO_FECHA, ESTADO" _
         & ", NOTAS, FECHA_INICIO, FECHA_CORTE, RECAUDACION_CUOTAS, RECAUDACION_CARGOS, NUEVOS_CREDITOS, RETENCION_AHORROS" _
         & ", COMISIONES_RECAUDACION,TOTAL_ORDEN, TOTAL_RESERVA, TOTAL_REBAJOS_CXP, TOTAL_PAGAR, DOCUMENTO,COMISIONES_NUEVOS_CREDITOS" _
         & ",PORC_COMISION_RECAUDACION,PORC_COMISION_CREDITOS,RECAUDACION_DEVOLUCION,APL_COM_INFORMATIVA,FLOTANTE_ALCOBRO,COBRA_CREDITOS_ANULADOS)" _
         & " Values ('" & mOrden & "','" & mConvenio & "','" & glogon.Usuario & "',dbo.MyGetdate(),'A'" _
         & ", '" & txtNotas.Text & "','" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00','" _
         & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'," & CCur(txtRecaudacion.Text) _
         & "," & CCur(txtCargos.Text) & "," & CCur(txtCreditos.Text) & "," & CCur(txtPlanesAhorros.Text) _
         & "," & CCur(txtComisionRecaudacion) & "," & curTotalBruto & "," & CCur(txtReservas.Text) _
         & "," & CCur(txtCargosCxP.Text) & "," & CCur(txtNetoLiquidar.Text) & ",'" & txtDocumento.Text _
         & "','" & CCur(txtComisionCreditos.Text) & "'," & mComisionRecaudacion & "," & mComisionNuevosCrd _
         & "," & CCur(txtDevoluciones.Text) & "," & chkComisionInformativa.Value & "," & CCur(txtFlotante.Text) _
         & "," & chkCreditosAnulados.Value & ")"
  Call ConectionExecute(strSQL)
  
  'Guarda los datos del detalle de la orden
  strSQL = "Exec spConvenio_Orden_Detalle '" & mConvenio & "'," & mOrden & ""
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Liquidación Convenio.: " & mConvenio & " - Orden.:" & mOrden)
  
  MsgBox "Información Guardada satisfactoriamente!!!", vbInformation
  

End If

Me.MousePointer = vbDefault

Call sbConsultaOrden(mOrden)

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

'Verifica el monto de reservas por acumular
Private Function fxVerificaReserva()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxVerificaReserva = 0

strSQL = "Select (RESERVA_TOPE - isnull(RESERVAS_SALDO,0)) AS 'ReservaDiferencia'" _
       & " from CRD_CONVENIOS where COD_CONVENIO = '" & txtCodigo.Text & "' "
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
  If (rs!ReservaDiferencia) > 0 Then
     fxVerificaReserva = rs!ReservaDiferencia
  End If
End If

rs.Close


Exit Function
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbCierraOrden()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass
 
 strSQL = "exec spConvenio_Orden_Cierra '" & mConvenio & "'," & mOrden & ",'" & glogon.Usuario & "'"
 Call ConectionExecute(strSQL)
  
Me.MousePointer = vbDefault

MsgBox "La orden se cerro Correctamente!!", vbInformation
  
Call sbConsultaOrden(mOrden)
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub

vGrid.Row = Row
vGrid.Col = 2
Call sbConsultaOrden(CLng(vGrid.Text))

End Sub
