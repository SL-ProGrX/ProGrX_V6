VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmSeguros_PolizaAjustes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambios / Ajustes a Pólizas"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSeguros_PolizaAjustes.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.FlatScrollBar FlatScrollBarTipoSeguro 
      Height          =   255
      Left            =   9120
      TabIndex        =   2
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBarTipoCuenta 
      Height          =   255
      Left            =   9120
      TabIndex        =   3
      Top             =   2520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9360
      Top             =   120
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
            Picture         =   "frmSeguros_PolizaAjustes.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_PolizaAjustes.frx":6953
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_PolizaAjustes.frx":6A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_PolizaAjustes.frx":6B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_PolizaAjustes.frx":6CA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_PolizaAjustes.frx":6DC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_PolizaAjustes.frx":6EEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_PolizaAjustes.frx":7015
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_PolizaAjustes.frx":7131
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_PolizaAjustes.frx":722D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_PolizaAjustes.frx":7344
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   5730
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Usuario Registra"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Fecha de Registro"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Usuario Activa"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Fecha Activa"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Usuario - Cierra"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Fecha Cierre"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.PushButton cmdAjustar 
      Height          =   612
      Left            =   5640
      TabIndex        =   15
      Top             =   4920
      Width           =   2172
      _Version        =   1441792
      _ExtentX        =   3831
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Ajustar"
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
      Picture         =   "frmSeguros_PolizaAjustes.frx":746D
   End
   Begin XtremeSuiteControls.PushButton cmdEliminar 
      Height          =   612
      Left            =   7800
      TabIndex        =   16
      Top             =   4920
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Eliminar"
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
      Picture         =   "frmSeguros_PolizaAjustes.frx":7DD2
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3600
      TabIndex        =   18
      Top             =   1320
      Width           =   6015
      _Version        =   1441792
      _ExtentX        =   10610
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   1920
      TabIndex        =   19
      Top             =   1320
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtVendedorDesc 
      Height          =   315
      Left            =   3600
      TabIndex        =   20
      Top             =   1800
      Width           =   6015
      _Version        =   1441792
      _ExtentX        =   10610
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtVendedorCod 
      Height          =   315
      Left            =   1920
      TabIndex        =   21
      Top             =   1800
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtTipoSeguroDesc 
      Height          =   315
      Left            =   3600
      TabIndex        =   22
      Top             =   2160
      Width           =   5415
      _Version        =   1441792
      _ExtentX        =   9551
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtTipoSeguroCod 
      Height          =   315
      Left            =   1920
      TabIndex        =   23
      Top             =   2160
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtTipoCuentaDesc 
      Height          =   315
      Left            =   3600
      TabIndex        =   24
      Top             =   2520
      Width           =   5415
      _Version        =   1441792
      _ExtentX        =   9551
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtTipoCuentaCod 
      Height          =   315
      Left            =   1920
      TabIndex        =   25
      Top             =   2520
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   1575
      Left            =   4560
      TabIndex        =   26
      Top             =   3000
      Width           =   5055
      _Version        =   1441792
      _ExtentX        =   8916
      _ExtentY        =   2778
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
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   315
      Left            =   1920
      TabIndex        =   27
      Top             =   3000
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   1920
      TabIndex        =   28
      Top             =   3480
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtPlazo 
      Height          =   315
      Left            =   1920
      TabIndex        =   29
      Top             =   3840
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtCuota 
      Height          =   315
      Left            =   1920
      TabIndex        =   30
      Top             =   4200
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtPoliza 
      Height          =   435
      Left            =   1920
      TabIndex        =   31
      Top             =   600
      Width           =   2055
      _Version        =   1441792
      _ExtentX        =   3625
      _ExtentY        =   767
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
   Begin XtremeSuiteControls.ComboBox cboAseguradora 
      Height          =   330
      Left            =   1920
      TabIndex        =   32
      Top             =   120
      Width           =   5655
      _Version        =   1441792
      _ExtentX        =   9975
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
      Caption         =   "Aseguradora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   33
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   3840
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuota"
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
      Left            =   600
      TabIndex        =   11
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   10
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Estado"
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
      Index           =   10
      Left            =   600
      TabIndex        =   9
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblPlazo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Plazo"
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
      Left            =   600
      TabIndex        =   8
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblPagador 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tipo Cobro"
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
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblContrato 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tipo Seguro"
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
      Left            =   600
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Identificación"
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
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vendedor"
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
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Image ImgAutorizacion 
      Height          =   255
      Left            =   4680
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Poliza"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   0
      Top             =   600
      Width           =   4935
   End
End
Attribute VB_Name = "frmSeguros_PolizaAjustes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje        As String  'Envia Mensajes en Fallas de Verificacion
Dim vEdita          As Boolean 'Indica si se esta actualizando o insertando
Dim vPaso           As Boolean 'Control de Activacion de Controles en proceso de carga
Dim vScroll         As Boolean
Dim vFecha          As Date

Dim strSQL As String, rs As New ADODB.Recordset

Function fxPersonaNombre(strCedula As String) As String
Dim rsX As New ADODB.Recordset

glogon.strSQL = "select nombre from Socios where cedula = '" & strCedula & "'"
Call OpenRecordSet(rsX, glogon.strSQL)

If rsX.EOF And rsX.BOF Then
 fxPersonaNombre = ""
Else
 fxPersonaNombre = IIf(IsNull(rsX!Nombre), "", rsX!Nombre)
End If
rsX.Close

End Function




Private Function fxValida() As Boolean

fxValida = True
vMensaje = ""


If Len(txtPoliza.Text) = 0 Then vMensaje = vMensaje & vbCrLf & "- No se indicó el número de la póliza!"


If IsNumeric(txtPlazo) Then
 If txtPlazo < 1 Then vMensaje = vMensaje & vbCrLf & "- El Plazo NO es válido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Plazo es Inválido"
End If


If IsNumeric(txtMonto.Text) Then
 If txtMonto.Text < 1 Then vMensaje = vMensaje & vbCrLf & "- El Monto NO es válido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Monto No es Inválido"
End If


strSQL = "select count(*) as Existe from SEGUROS_TIPOS_PRODUCTOS where COD_PRODUCTO ='" & txtTipoSeguroCod.Text & "' and Activo = 1"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El tipo de seguro no se encuentra Activo!"
rs.Close

strSQL = "select count(*) as Existe from SEGUROS_TIPOS_COBRO where TIPO_COBRO ='" & txtTipoCuentaCod.Text & "' and Activo = 1"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El tipo de Cuenta no se encuentra Activa!"
rs.Close


strSQL = "select count(*) as Existe from SEGUROS_Vendedores where cod_vendedor = " & txtVendedorCod.Text & " and Activo = 1"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El vendedor no se encuentra Activo!"
rs.Close

strSQL = "select count(*) as Existe from Socios where cedula = '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- La persona no existe en la base de datos!"
rs.Close


If Len(vMensaje) > 0 Then fxValida = False

End Function


Private Sub FlatScrollBar_Change()

On Error GoTo vError

If txtPoliza.Text = "" Then txtPoliza.Text = "0"
If FlatScrollBar.Tag = "" Then FlatScrollBar.Tag = 0

strSQL = "select Top 1 num_poliza from SEGUROS_REGISTRO"



If FlatScrollBar.Value > CLng(FlatScrollBar.Tag) Then
   strSQL = strSQL & " where cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "' and  num_poliza > '" & txtPoliza & "' order by num_poliza asc"
Else
   strSQL = strSQL & " where cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "' and num_poliza < '" & txtPoliza & "' order by num_poliza desc"
End If

FlatScrollBar.Tag = FlatScrollBar.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtPoliza.Text = rs!Num_Poliza
  Call sbConsulta
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cmdAjustar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vExiste As Integer

On Error GoTo vError

       
If Mid(txtEstado.Text, 1, 1) <> "A" Then
    MsgBox "No se puede modificar esta póliza porque no se encuentra Activa", vbExclamation
    Exit Sub
End If
       
strSQL = "update SEGUROS_REGISTRO set cod_vendedor = " & txtVendedorCod.Text & ",COD_PRODUCTO = '" & txtTipoSeguroCod.Text & "',TIPO_COBRO = '" _
    & txtTipoCuentaCod.Text & "',notas = '" & txtNotas.Text & "',Monto = " & CCur(txtMonto.Text) & ", Cuota =  " & CCur(txtCuota.Text) _
    & ", Plazo = " & txtPlazo.Text & ", cedula = '" & Trim(txtCedula.Text) _
    & "' where num_poliza = '" & txtPoliza.Text & "' and cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
Call ConectionExecute(strSQL)

'Actualiza datos de Cobranza General y Variaciones en las Polizas
strSQL = "exec spSeguros_Poliza_SincronizaCambios '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtPoliza.Text & "'"
Call ConectionExecute(strSQL)


MsgBox "Ajuste a Póliza realizado satisfactoriamente!", vbInformation
'TODO: Crear Bitácora

Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cmdEliminar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vExiste As Integer

On Error GoTo vError

       
If Mid(txtEstado.Text, 1, 1) <> "A" Then
    MsgBox "No se puede Eliminar esta póliza porque no se encuentra Activa", vbExclamation
    Exit Sub
End If
       
strSQL = "exec spSeguros_PolizaActivaElimina '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtPoliza.Text & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
       
MsgBox "Póliza Eliminada / Ajustadas todas las referencias...", vbInformation
       
Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub FlatScrollBarTipoSeguro_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarTipoSeguro.Tag = "" Then FlatScrollBarTipoSeguro.Tag = 0

strSQL = "select Top 1 COD_PRODUCTO,Descripcion from SEGUROS_TIPOS_PRODUCTOS"

If FlatScrollBarTipoSeguro.Value > CLng(FlatScrollBarTipoSeguro.Tag) Then
   strSQL = strSQL & " where cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "' and COD_PRODUCTO > '" & txtTipoSeguroCod.Text & "' and Activo = 1 order by COD_PRODUCTO asc"
Else
   strSQL = strSQL & " where cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "' and COD_PRODUCTO < '" & txtTipoSeguroCod.Text & "' and Activo = 1 order by COD_PRODUCTO asc"
End If

FlatScrollBarTipoSeguro.Tag = FlatScrollBarTipoSeguro.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
        txtTipoSeguroCod.Text = rs!COD_PRODUCTO
        txtTipoSeguroDesc.Text = rs!Descripcion
Else
        txtTipoSeguroCod.Text = ""
        txtTipoSeguroDesc.Text = ""
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub FlatScrollBarTipoCuenta_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarTipoCuenta.Tag = "" Then FlatScrollBarTipoCuenta.Tag = 0

strSQL = "select Top 1 TIPO_COBRO,Descripcion from SEGUROS_TIPOS_COBRO"

If FlatScrollBarTipoCuenta.Value > CLng(FlatScrollBarTipoCuenta.Tag) Then
   strSQL = strSQL & " where TIPO_COBRO > '" & txtTipoCuentaCod.Text & "' and Activo = 1 order by TIPO_COBRO asc"
Else
   strSQL = strSQL & " where TIPO_COBRO < '" & txtTipoCuentaCod.Text & "' and Activo = 1 order by TIPO_COBRO asc"
End If

FlatScrollBarTipoCuenta.Tag = FlatScrollBarTipoCuenta.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtTipoCuentaCod.Text = rs!TIPO_COBRO
  txtTipoCuentaDesc.Text = rs!Descripcion
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 17
End Sub

Private Sub Form_Load()
 
Dim strSQL As String

strSQL = "select cod_aseguradora as 'IdX', rtrim(nombre) as 'ItmX' from seguros_Aseguradoras"
Call sbCbo_Llena_New(cboAseguradora, strSQL, False, True)

 vModulo = 17

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbLimpiaPantalla


End Sub

Private Sub sbLimpiaPantalla()
 

Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(11).Picture
ImgAutorizacion.ToolTipText = "Pendiente: Consulta/Nuevo"
 
 txtCedula = ""
 txtNombre = ""
 lblNombre.Caption = txtNombre.Text
 
 txtVendedorCod.Text = ""
 txtVendedorDesc.Text = ""

 txtTipoSeguroCod.Text = ""
 txtTipoSeguroDesc.Text = ""

 txtTipoCuentaCod.Text = ""
 txtTipoCuentaDesc.Text = ""
   
 txtEstado.Text = "Pendiente"
   
 txtMonto = "0"
 txtPlazo = "60"
 txtCuota = "0"
  

 txtNotas = ""
 

 StatusBarX.Panels(1).Text = ""
 StatusBarX.Panels(2).Text = ""
 StatusBarX.Panels(3).Text = ""
 StatusBarX.Panels(4).Text = ""
 StatusBarX.Panels(5).Text = ""

End Sub



Private Sub sbConsulta()


On Error GoTo vError

vPaso = True

strSQL = "select Pol.*,Ts.Descripcion as 'TipoSeguroDesc', Per.Nombre, isnull(Pol.Estado,'P') as 'Estado'" _
       & ",Ven.Nombre as 'VendedorNombre',Tc.descripcion as 'TipoCuentaDesc',dbo.MyGetdate() as 'FechaServer'" _
       & " from SEGUROS_REGISTRO Pol inner join SEGUROS_TIPOS_PRODUCTOS Ts on Pol.COD_PRODUCTO = Ts.COD_PRODUCTO" _
       & " inner join Socios Per on Pol.cedula = Per.cedula" _
       & " left join SEGUROS_Vendedores Ven on Pol.cod_Vendedor = Ven.cod_Vendedor" _
       & " left join SEGUROS_TIPOS_COBRO Tc on Pol.TIPO_COBRO = Tc.TIPO_COBRO" _
       & " where Pol.cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "' and Pol.num_poliza = '" & txtPoliza.Text & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
 
 vFecha = rs!FechaServer
 
 txtCedula.Text = rs!Cedula
 txtNombre.Text = rs!Nombre
 lblNombre.Caption = txtNombre.Text
 
 txtVendedorCod.Text = rs!cod_vendedor
 txtVendedorDesc.Text = rs!VendedorNombre
 
 txtTipoSeguroCod.Text = rs!COD_PRODUCTO
 txtTipoSeguroDesc.Text = rs!TipoSeguroDesc
 txtTipoCuentaCod.Text = rs!TIPO_COBRO
 txtTipoCuentaDesc.Text = rs!TipoCuentaDesc
 
 
 txtMonto.Text = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "Standard")
 txtPlazo.Text = rs!Plazo
 txtCuota.Text = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
 
 txtNotas = IIf(IsNull(rs!notas), "", rs!notas)



 Select Case rs!estado
   Case "P" 'Pendiente
      txtEstado.Text = "Pendiente"
      Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(7).Picture
      ImgAutorizacion.ToolTipText = "Activación: Pendiente"
      
   Case "A"
      txtEstado.Text = "Activa"
      Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(5).Picture
      ImgAutorizacion.ToolTipText = "Póliza Activada!"
   Case "C"
      txtEstado.Text = "Cerrada"
      Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(6).Picture
      ImgAutorizacion.ToolTipText = "Póliza Cerrada (Inactivada)"
  End Select

 StatusBarX.Panels(1).Text = rs!Registro_Usuario
 StatusBarX.Panels(2).Text = rs!Registro_Fecha
 StatusBarX.Panels(3).Text = rs!Activa_Usuario & ""
 StatusBarX.Panels(4).Text = rs!ACTIVA_FECHA & ""
 StatusBarX.Panels(5).Text = rs!Cierra_usuario & ""
 StatusBarX.Panels(6).Text = rs!Cierra_fecha & ""
 

Else
 If vEdita Then
    MsgBox "No existe la Póliza, verifique!", vbCritical
 End If
End If
rs.Close

vPaso = False

Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub






Private Sub txtCuota_GotFocus()
On Error GoTo vError

txtCuota.Text = CCur(txtCuota.Text)

vError:
End Sub

Private Sub txtCuota_LostFocus()
On Error GoTo vError

txtCuota.Text = Format(CCur(txtCuota.Text), "Standard")

vError:

End Sub

Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuota.SetFocus
End Sub

Private Sub txtPoliza_LostFocus()
  Call sbConsulta
End Sub

Private Sub txtTipoSeguroCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoSeguroDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "COD_PRODUCTO"
   gBusquedas.Orden = "COD_PRODUCTO"
   gBusquedas.Filtro = " and Activo = 1 and cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
   gBusquedas.Consulta = "Select COD_PRODUCTO,Descripcion  from SEGUROS_TIPOS_PRODUCTOS"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoSeguroCod.Text = gBusquedas.Resultado
      txtTipoSeguroDesc.Text = gBusquedas.Resultado2
      Call txtTipoSeguroCod_LostFocus
   End If
End If

End Sub

Private Sub txtTipoSeguroCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from SEGUROS_TIPOS_PRODUCTOS where COD_PRODUCTO = '" & txtTipoSeguroCod.Text _
       & "' and cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtTipoSeguroCod.Text = rs!COD_PRODUCTO
   txtTipoSeguroDesc.Text = rs!Descripcion
End If
rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub txtTipoSeguroDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoCuentaCod.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Orden = "Descripcion"
   gBusquedas.Filtro = " and Activo = 1 and cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
   gBusquedas.Consulta = "Select COD_PRODUCTO,Descripcion  from SEGUROS_TIPOS_PRODUCTOS"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoSeguroCod.Text = gBusquedas.Resultado
      txtTipoSeguroDesc.Text = gBusquedas.Resultado2
      Call txtTipoSeguroCod_LostFocus
   End If
End If
End Sub


'--Tipo de Cuenta
Private Sub txtTipoCuentaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "TIPO_COBRO"
   gBusquedas.Orden = "TIPO_COBRO"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select TIPO_COBRO,Descripcion  from SEGUROS_TIPOS_COBRO"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoCuentaCod.Text = gBusquedas.Resultado
      txtTipoCuentaDesc.Text = gBusquedas.Resultado2
      Call txtTipoCuentaCod_LostFocus
   End If
End If

End Sub

Private Sub txtTipoCuentaCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from SEGUROS_TIPOS_COBRO where TIPO_COBRO = '" & txtTipoCuentaCod.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtTipoCuentaCod.Text = rs!TIPO_COBRO
   txtTipoCuentaDesc.Text = rs!Descripcion
End If
rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub txtTipoCuentaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Orden = "Descripcion"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select TIPO_COBRO,Descripcion  from SEGUROS_TIPOS_COBRO"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoCuentaCod.Text = gBusquedas.Resultado
      txtTipoCuentaDesc.Text = gBusquedas.Resultado2
      Call txtTipoCuentaCod_LostFocus
   End If
End If
End Sub




Private Sub txtCuota_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus

End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError

txtMonto.Text = CCur(txtMonto.Text)

vError:

End Sub


Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlazo.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError

txtMonto.Text = Format(txtMonto.Text, "Standard")

vError:

End Sub



Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtCedula.Text = gBusquedas.Resultado
      txtNombre.Text = gBusquedas.Resultado2
      Call txtCedula_LostFocus
   End If
End If


End Sub

Private Sub txtCedula_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


txtNombre.Text = fxPersonaNombre(txtCedula)
lblNombre.Caption = txtNombre.Text


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtVendedorCod.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtCedula.Text = gBusquedas.Resultado
      txtNombre.Text = gBusquedas.Resultado2
      Call txtCedula_LostFocus
   End If
End If
End Sub

'--Vendedor
Private Sub txtVendedorCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtVendedorDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cod_Vendedor"
   gBusquedas.Orden = "Cod_Vendedor"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cod_Vendedor,Nombre from SEGUROS_Vendedores"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtVendedorCod.Text = gBusquedas.Resultado
      txtVendedorDesc.Text = gBusquedas.Resultado2
      Call txtVendedorCod_LostFocus
   End If
End If


End Sub

Private Sub txtVendedorCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "Select Cod_Vendedor,Nombre from SEGUROS_Vendedores where cod_Vendedor = " & txtVendedorCod.Text
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtVendedorDesc.Text = rs!Nombre
End If
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtVendedorDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoSeguroCod.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cod_Vendedor,Nombre from SEGUROS_Vendedores"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtVendedorCod.Text = gBusquedas.Resultado
      txtVendedorDesc.Text = gBusquedas.Resultado2
      Call txtVendedorCod_LostFocus
   End If
End If
End Sub


Public Sub sbConsultaExterna(xOpTemp As String)
 txtPoliza.Text = xOpTemp
 Call sbConsulta
End Sub


Private Sub txtPoliza_Change()
 Call sbLimpiaPantalla

End Sub

Private Sub txtPoliza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedula.SetFocus
End Sub




