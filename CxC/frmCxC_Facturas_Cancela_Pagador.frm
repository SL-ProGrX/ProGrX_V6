VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmCxC_Facturas_Cancela_Pagador 
   Caption         =   "Cancelación de Facturas por Pagador"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6372
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   11772
      _Version        =   1441793
      _ExtentX        =   20764
      _ExtentY        =   11239
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
      SelectedItem    =   1
      Item(0).Caption =   "Facturas"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "lswFacturas"
      Item(0).Control(1)=   "gbFacturas"
      Item(1).Caption =   "Cancelación"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "fraFormaPago"
      Item(1).Control(2)=   "txtTotalPagar"
      Item(1).Control(3)=   "txtDiferencia"
      Item(1).Control(4)=   "lblTotal(0)"
      Item(1).Control(5)=   "lblTotal(1)"
      Item(1).Control(6)=   "btnExport(1)"
      Begin XtremeSuiteControls.ListView lswFacturas 
         Height          =   4932
         Left            =   -69880
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   11772
         _Version        =   1441793
         _ExtentX        =   20764
         _ExtentY        =   8700
         _StockProps     =   77
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
         View            =   3
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   7
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3732
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   11532
         _Version        =   1441793
         _ExtentX        =   20341
         _ExtentY        =   6583
         _StockProps     =   77
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
         View            =   3
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   7
      End
      Begin XtremeSuiteControls.GroupBox gbFacturas 
         Height          =   732
         Left            =   -69880
         TabIndex        =   23
         Top             =   5400
         Visible         =   0   'False
         Width           =   11652
         _Version        =   1441793
         _ExtentX        =   20553
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Detalle de Facturas:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtFacturas_Pendiente_Casos 
            Height          =   312
            Left            =   4320
            TabIndex        =   26
            Top             =   360
            Width           =   852
            _Version        =   1441793
            _ExtentX        =   1503
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
         Begin XtremeSuiteControls.FlatEdit txtFacturas_Pendiente_Total 
            Height          =   312
            Left            =   2400
            TabIndex        =   24
            Top             =   360
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFacturas_Sel_Total 
            Height          =   312
            Left            =   7560
            TabIndex        =   27
            Top             =   360
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFacturas_Sel_Casos 
            Height          =   312
            Left            =   9480
            TabIndex        =   29
            Top             =   360
            Width           =   852
            _Version        =   1441793
            _ExtentX        =   1503
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
         Begin XtremeSuiteControls.PushButton btnExport 
            Height          =   255
            Index           =   0
            Left            =   10560
            TabIndex        =   31
            ToolTipText     =   "Exportar a Excel"
            Top             =   360
            Width           =   255
            _Version        =   1441793
            _ExtentX        =   444
            _ExtentY        =   444
            _StockProps     =   79
            Appearance      =   16
            Picture         =   "frmCxC_Facturas_Cancela_Pagador.frx":0000
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "Seleccionadas:"
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
            Left            =   5160
            TabIndex        =   28
            Top             =   360
            Width           =   2292
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "Pendientes:"
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
            Index           =   2
            Left            =   360
            TabIndex        =   25
            Top             =   360
            Width           =   1932
         End
      End
      Begin VB.TextBox txtDiferencia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   324
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   $"frmCxC_Facturas_Cancela_Pagador.frx":08D1
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox txtTotalPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   324
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   $"frmCxC_Facturas_Cancela_Pagador.frx":08D8
         Top             =   4320
         Width           =   1812
      End
      Begin VB.Frame fraFormaPago 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   960
         TabIndex        =   9
         Top             =   4680
         Width           =   9612
         Begin VB.TextBox txtNotas 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   720
            Width           =   5415
         End
         Begin VB.TextBox txtTotalCajas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   1815
         End
         Begin MSComctlLib.Toolbar tlbCajas 
            Height          =   810
            Left            =   6720
            TabIndex        =   12
            Top             =   720
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   1429
            ButtonWidth     =   1402
            ButtonHeight    =   1429
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Pago"
                  Key             =   "Desgloce"
                  Object.ToolTipText     =   "Desglosa la transacción"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Aplicar"
                  Key             =   "Aplicar"
                  Object.ToolTipText     =   "Aplica la Transacción"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Cancelar"
                  Key             =   "Cancelar"
                  Object.ToolTipText     =   "Cancelar la Transacción"
                  ImageIndex      =   3
               EndProperty
            EndProperty
         End
         Begin XtremeSuiteControls.ComboBox cboTipoDoc 
            Height          =   312
            Left            =   1320
            TabIndex        =   13
            Top             =   240
            Width           =   2772
            _Version        =   1441793
            _ExtentX        =   4895
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin VB.Label Label3 
            Caption         =   "Documento ..:"
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
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Notas ..:"
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
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Total ..:"
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
            Index           =   4
            Left            =   4200
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   255
         Index           =   1
         Left            =   10680
         TabIndex        =   32
         ToolTipText     =   "Exportar a Excel"
         Top             =   4320
         Width           =   255
         _Version        =   1441793
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmCxC_Facturas_Cancela_Pagador.frx":08DF
      End
      Begin VB.Label lblTotal 
         Caption         =   "Diferencia ...:"
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
         Left            =   7800
         TabIndex        =   20
         Top             =   4320
         Width           =   1332
      End
      Begin VB.Label lblTotal 
         Caption         =   "Total a Pagar.:"
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
         Left            =   4440
         TabIndex        =   19
         Top             =   4320
         Width           =   1332
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8760
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Facturas_Cancela_Pagador.frx":11B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Facturas_Cancela_Pagador.frx":7A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Facturas_Cancela_Pagador.frx":81EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Facturas_Cancela_Pagador.frx":89B7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboPagador 
      Height          =   312
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   6492
      _Version        =   1441793
      _ExtentX        =   11456
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboDivisas 
      Height          =   312
      Left            =   8280
      TabIndex        =   1
      Top             =   360
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3413
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtFactura 
      Height          =   315
      Left            =   8280
      TabIndex        =   5
      Top             =   960
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
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
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCliente 
      Height          =   312
      Left            =   1800
      TabIndex        =   21
      Top             =   960
      Width           =   6492
      _Version        =   1441793
      _ExtentX        =   11451
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
   Begin VB.Label lblLoading 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "...Cargando Facturas!"
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
      Left            =   10320
      TabIndex        =   30
      Top             =   960
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente..:"
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
      Index           =   1
      Left            =   1800
      TabIndex        =   22
      Top             =   720
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Factura No..:"
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
      Left            =   8280
      TabIndex        =   4
      Top             =   720
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Divisa...:"
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
      Index           =   7
      Left            =   8280
      TabIndex        =   3
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pagador..:"
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
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1212
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12012
   End
End
Attribute VB_Name = "frmCxC_Facturas_Cancela_Pagador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mClienteId As String, mOperacion As Long, mFactura As String
Dim vPaso As Boolean, pCharRelleno As String

Private Sub sbLimpiaDatos()
 
 txtDiferencia.Text = 0

 txtTotalCajas.Text = 0

 txtNotas.Text = ""
 
 lsw.ListItems.Clear
 lswFacturas.ListItems.Clear
 
With lsw.ColumnHeaders
 .Clear
 .Add , , "No. Operación", 1200
 .Add , , "No. Factura", 2200
 .Add , , "Monto", 1800, vbRightJustify
 .Add , , "Fecha Pago", 1200
 .Add , , "Divisa", 1100, vbCenter
 .Add , , "Importe", 1400, vbRightJustify
 .Add , , "Fecha Emite", 1200
 .Add , , "Fecha Activa", 1200
 .Add , , "Cliente Id", 2000
 .Add , , "Cliente Nombre", 5000
End With

With lswFacturas.ColumnHeaders
 .Clear
 .Add , , "No. Operación", 1200
 .Add , , "No. Factura", 2200
 .Add , , "Monto", 1800, vbRightJustify
 .Add , , "Fecha Pago", 1200
 .Add , , "Divisa", 1100, vbCenter
 .Add , , "Importe", 1400, vbRightJustify
 .Add , , "Fecha Emite", 1200
 .Add , , "Fecha Activa", 1200
 .Add , , "Cliente Id", 2000
 .Add , , "Cliente Nombre", 5000
End With


End Sub

Private Function fxFactura_Seleccionada(pOperacion As Long, pFactura As String) As Boolean
Dim i As Long, vResultado As Boolean

vResultado = False

With lsw.ListItems
For i = 1 To .Count
    If CLng(.Item(i).Text) = pOperacion And .Item(i).SubItems(1) = pFactura Then
            vResultado = True
            Exit For
    End If
Next i
End With

fxFactura_Seleccionada = vResultado

End Function

Private Sub sbFacturas_Carga(Optional pFiltro As Boolean = False)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotal As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Item(0).Selected = True

lblLoading.Visible = True

'Carga Facturas
strSQL = "select Operacion, Cod_Factura, Monto, Cod_Divisa, Fecha_Pago, Importe, Fecha_Emision, Activa_Fecha, Cedula, Nombre from vCxC_Facturas_Pendientes_Cancelacion" _
       & " where Cedula_Pagador = '" & cboPagador.ItemData(cboPagador.ListIndex) & "'" _
       & " and cod_divisa = '" & cboDivisas.ItemData(cboDivisas.ListIndex) & "'"
       
curTotal = 0
lswFacturas.ListItems.Clear

If pFiltro Then
   strSQL = strSQL & " and cod_factura like '%" & txtFactura.Text _
          & "%' and Nombre like '%" & txtCliente.Text & "%'"
Else
    lsw.ListItems.Clear
End If
       
'Orden
strSQL = strSQL & " order by Nombre, COD_FACTURA, FECHA_PAGO"


Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  If Not fxFactura_Seleccionada(rs!Operacion, rs!cod_Factura) Then
        Set itmX = lswFacturas.ListItems.Add(, , rs!Operacion)
            itmX.SubItems(1) = rs!cod_Factura
            itmX.SubItems(2) = Format(rs!Monto, "Standard")
            itmX.SubItems(3) = Format(rs!Fecha_Pago, "dd/mm/yyyy")
            itmX.SubItems(4) = rs!cod_Divisa
            itmX.SubItems(5) = Format(rs!Importe, "Standard")
            itmX.SubItems(6) = Format(rs!Fecha_Emision, "dd/mm/yyyy")
            itmX.SubItems(7) = Format(rs!Activa_Fecha, "dd/mm/yyyy")
            itmX.SubItems(8) = rs!Cedula
            itmX.SubItems(9) = rs!Nombre
            
            curTotal = curTotal + rs!Monto
  End If
  rs.MoveNext
Loop
rs.Close

txtFacturas_Pendiente_Casos.Text = Format(lswFacturas.ListItems.Count, "###,##0")
txtFacturas_Pendiente_Total.Text = Format(curTotal, "Standard")

lblLoading.Visible = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnExport_Click(Index As Integer)
Select Case Index
    Case 0 'Lista
        Call Excel_Exportar_Lsw(lswFacturas)
    Case 1 'Cancela
        Call Excel_Exportar_Lsw(lsw)
End Select
End Sub

Private Sub cboDivisas_Click()
If vPaso Or cboDivisas.ListCount <= 0 Then Exit Sub

    ModuloCajas.mTotalDetallado = 0
    txtTotalCajas.Text = 0
    
    ModuloCajas.mClienteId = cboPagador.ItemData(cboPagador.ListIndex)
    ModuloCajas.mCliente = cboPagador.Text
    
    ModuloCajas.mDivisa = cboDivisas.ItemData(cboDivisas.ListIndex)
    ModuloCajas.mConceptoValida = True 'IIf((rs!Caja_Valida_Concepto > 0), True, False)
    
    txtFactura.Text = ""
    txtCliente.Text = ""
    
    txtFacturas_Pendiente_Casos.Text = "0"
    txtFacturas_Pendiente_Total.Text = "0"
    
    txtFacturas_Sel_Casos.Text = "0"
    txtFacturas_Sel_Total.Text = "0"
    
    Call sbFacturas_Carga(False)

End Sub

Private Sub cboPagador_Click()
If vPaso Or cboPagador.ListCount <= 0 Then Exit Sub

Dim strSQL As String

    'Re Inicia Calculos de Cajas
    ModuloCajas.mTiquete = Mid(cboPagador.ItemData(cboPagador.ListIndex), 1, 10) & "." & Format(Time, "HH:mm:ss")
    
    txtTotalPagar.Text = 0
    txtDiferencia.Text = 0
    
    'Carga Divisas de Facturas Pendientes con el Pagador
    strSQL = "select Cod_Divisa as 'IdX', Cod_Divisa as 'ItmX'" _
           & " from vCxC_Facturas_Pendientes_Cancelacion" _
           & " where cedula_pagador = '" & cboPagador.ItemData(cboPagador.ListIndex) & "'" _
           & " group by Cod_Divisa"
     
     vPaso = True
         Call sbCbo_Llena_New(cboDivisas, strSQL, False, True)
     vPaso = False
     
     Call cboDivisas_Click


End Sub

Private Sub Form_Activate()
 vModulo = 31

End Sub


Private Sub sbPagagor_Load()
Dim strSQL As String

'Carga Pagadores Con Facturas Pendientes de Pago
strSQL = "select Per.Cedula as 'IdX', Per.Nombre as 'ItmX'" _
      & " from vCxC_Facturas_Pendientes_Cancelacion Ft inner join CxC_Personas Per on Ft.Cedula_Pagador = Per.Cedula" _
      & " group by Per.Cedula, Per.Nombre"
vPaso = True
    cboPagador.Clear
    Call sbCbo_Llena_New(cboPagador, strSQL, False, True)
vPaso = False
 
'Call cboPagador_Click

End Sub

Private Sub Form_Load()
 
 vModulo = 31
 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
   
 tcMain.Item(0).Selected = True

 Call sbLimpiaDatos
 Call sbPagagor_Load
  
 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

tcMain.Width = Me.Width - 300
tcMain.Height = Me.Height - (tcMain.top + 400)

lswFacturas.Width = tcMain.Width - 300
lswFacturas.Height = tcMain.Height - (lswFacturas.top + gbFacturas.Height + 400)

gbFacturas.top = lswFacturas.top + lswFacturas.Height + 100
gbFacturas.Width = lswFacturas.Width

lsw.Width = tcMain.Width - 300
lsw.Height = tcMain.Height - (lsw.top + fraFormaPago.Height + 800)

lblTotal.Item(0).top = lsw.top + lsw.Height + 110
lblTotal.Item(1).top = lblTotal.Item(0).top

txtTotalPagar.top = lblTotal.Item(0).top
txtDiferencia.top = txtTotalPagar.top

btnExport.Item(1).top = lblTotal.Item(0).top

fraFormaPago.top = txtTotalPagar.top + 340

End Sub



Private Sub sbFacturas_Totales()
Dim i As Long, curTotal As Currency

On Error GoTo vError

curTotal = 0
With lsw.ListItems
    For i = 1 To .Count
      curTotal = curTotal + CCur(.Item(i).SubItems(5))
    Next i
End With

txtTotalPagar.Text = Format(curTotal, "Standard")
txtDiferencia.Text = Format(CCur(txtTotalCajas.Text) - CCur(txtTotalPagar), "Standard")

txtFacturas_Sel_Casos.Text = Format(lsw.ListItems.Count, "###,##0")
txtFacturas_Sel_Total.Text = Format(curTotal, "Standard")

curTotal = 0
With lswFacturas.ListItems
    For i = 1 To .Count
      curTotal = curTotal + CCur(.Item(i).SubItems(5))
    Next i
End With

txtFacturas_Pendiente_Casos.Text = Format(lswFacturas.ListItems.Count, "###,##0")
txtFacturas_Pendiente_Total.Text = Format(curTotal, "Standard")


vError:
End Sub



Private Sub lsw_DblClick()
Dim itmX As ListViewItem

On Error GoTo vError

If lsw.ListItems.Count = 0 Then Exit Sub


With lsw.SelectedItem
    Set itmX = lswFacturas.ListItems.Add(, , .Text)
        itmX.SubItems(1) = .SubItems(1)
        itmX.SubItems(2) = .SubItems(2)
        itmX.SubItems(3) = .SubItems(3)
        itmX.SubItems(4) = .SubItems(4)
        itmX.SubItems(5) = .SubItems(5)
        itmX.SubItems(6) = .SubItems(6)
        itmX.SubItems(7) = .SubItems(7)
        itmX.SubItems(8) = .SubItems(8)
        itmX.SubItems(9) = .SubItems(9)
   

    lsw.ListItems.Remove lsw.SelectedItem.Index
End With

Call sbFacturas_Totales

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswFacturas_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswFacturas.SortKey = ColumnHeader.Index - 1
  If lswFacturas.SortOrder = 0 Then lswFacturas.SortOrder = 1 Else lswFacturas.SortOrder = 0
  lswFacturas.Sorted = True
End Sub

Private Sub lswFacturas_DblClick()
Dim itmX As ListViewItem

If lswFacturas.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

With lswFacturas.SelectedItem
If Not fxFactura_Seleccionada(.Text, .SubItems(1)) Then
    Set itmX = lsw.ListItems.Add(, , .Text)
        itmX.SubItems(1) = .SubItems(1)
        itmX.SubItems(2) = .SubItems(2)
        itmX.SubItems(3) = .SubItems(3)
        itmX.SubItems(4) = .SubItems(4)
        itmX.SubItems(5) = .SubItems(5)
        itmX.SubItems(6) = .SubItems(6)
        itmX.SubItems(7) = .SubItems(7)
        itmX.SubItems(8) = .SubItems(8)
        itmX.SubItems(9) = .SubItems(9)
   
   lswFacturas.ListItems.Remove .Index
   
   Call sbFacturas_Totales
End If

End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub



Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

Call sbCajaInicial

Call cboPagador_Click

End Sub


Private Sub sbCajaInicial()
Dim strSQL As String

'Paso 1: Si la Caja no está abierta (Llamar pantalla de login de Caja)
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Or ModuloCajas.mUsuario <> glogon.Usuario Then
   Call sbFormsCall("frmCajas_Acceso", vbModal, , , False, Me)
End If

'Paso 2: Si despues del Login de Caja permanece sin Apertura Salir
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   MsgBox "No se ha indicado ninguna caja con Apertura disponible?", vbExclamation
   Unload Me
   Exit Sub
End If

pCharRelleno = fxCajasParametros("05")

Me.Caption = "Abonos a Cuentas por Cobrar    ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

txtTotalCajas.Text = 0
txtNotas.Text = ""
strSQL = "select rTrim(C.tipo_documento) as 'IdX' , rtrim(D.Descripcion) as 'itmX'" _
       & " from SIF_DOCUMENTOS D inner join CAJAS_DOCUMENTOS C on D.TIPO_DOCUMENTO = C.TIPO_DOCUMENTO " _
       & " Where C.cod_caja =  '" & ModuloCajas.mCaja & "' and D.Tipo_Movimiento in('A','C')" _
       & " order by C.tipo_documento"
Call sbCbo_Llena_New(cboTipoDoc, strSQL, False, True)

ModuloCajas.mServicio = "Abonos a Cuentas por Cobrar"

End Sub


Private Sub sbDocumentoAbono(pTipoAbono As String, pTipoDoc As String, pNumDoc As String _
                                , pConcepto As String, pCuenta As String)
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String
Dim strCliente As String, vCuenta As String
Dim rsTmp As New ADODB.Recordset, vCuentaPoliza As String, pTipoCambio As Currency
Dim curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency

'vCuenta = pCuenta
'
'pTipoCambio = fxCajasTipoCambio(ModuloCajas.mDivisa)
'
'
''Cuentas
'strSQL = "exec spCxC_OperacionCtas " & txtOperacion.Text
'Call OpenRecordSet(rs, strSQL)
'
'
'strSQL = "exec spCxC_DocumentoAfectacion '" & fxTipoASENumero(pTipoDoc) & "','" & pNumDoc & "','R'"
'Call OpenRecordSet(rsTmp, strSQL, 0)
'If rsTmp.EOF And rsTmp.BOF Then
'  curIntC = 0
'  curIntM = 0
'  curAmortiza = 0
'  curCargo = 0
'Else
'  curIntC = rsTmp!IntCor
'  curIntM = rsTmp!IntMor
'  curAmortiza = rsTmp!Principal
'  curCargo = rsTmp!Cargos
'End If
'rsTmp.Close
'
'
'
''Lineas de Comprobante
'strLinea(1) = "Saldo Anterior    ..: " & SIFGlobal.fxStringRelleno(lblSaldo.Caption, "I", pCharRelleno, 15) '
'strLinea(2) = "Saldo Actual      ..: " & SIFGlobal.fxStringRelleno(Format(CCur(lblSaldo.Caption) - curAmortiza, "Standard"), "I", pCharRelleno, 15) '
'strLinea(3) = "Interes Corriente ..: " & SIFGlobal.fxStringRelleno(Format(curIntC, "Standard"), "I", pCharRelleno, 15) '
'strLinea(4) = "Interes Atrasado  ..: " & SIFGlobal.fxStringRelleno(Format(curIntM, "Standard"), "I", pCharRelleno, 15) '
'strLinea(5) = "Amortización      ..: " & SIFGlobal.fxStringRelleno(Format(curAmortiza, "Standard"), "I", pCharRelleno, 15) '
'strLinea(6) = "Cargos Totales    ..: " & SIFGlobal.fxStringRelleno(Format(curCargo, "Standard"), "I", pCharRelleno, 15) '
'strLinea(7) = "Operacion/Concepto..: " & "Op.:" & txtOperacion.Text & " Cpt.:" & txtCodigo.Text
'
'If cboDiferenciaApl.Enabled Then
'    strLinea(8) = "Aplica Diferencia ..: " & cboDiferenciaApl.Text
'Else
'    strLinea(8) = "Descripción       ..: " & lblDescripcion.Caption
'
'End If
'
'
'
'strLinea(9) = ""
'strLinea(10) = "Num. Documento    ..:" & lblDocumento.Caption
'strLinea(11) = ""
'
'strSQL = "exec spCxC_OperacionFechaProxPago " & txtOperacion.Text
'Call OpenRecordSet(rsTmp, strSQL, 0)
'  If Not IsNull(rsTmp!fecha_corte) Then
'       strLinea(9) = "Prox.Pago..:" & Format(rsTmp!fecha_corte, "dd/mm/yyyy") & " Cta.(" & rsTmp!Linea & ") " & Format(rsTmp!Monto, "Standard")
'  Else
'       strLinea(9) = "Prox.Pago..: >> <<"
'  End If
'  strLinea(10) = "Notas: " & rsTmp!Notas & ""
'rsTmp.Close
'
'strLinea(10) = Mid(strLinea(10), 1, 80)
'
'
'If dtpFechaCancelacion.Enabled Then
'   strLinea(11) = "Fecha Real Abono  ..: " & Format(dtpFechaCancelacion.Value, "dd/mm/yyyy")
'End If
'
''Registro del Comprobante
'strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
'         & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
'         & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle,documento)" _
'         & " values('" & pNumDoc & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
'         & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curIntC + curIntM + curAmortiza + curCargo & ",'P','" & txtOperacion.Text _
'         & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
'         & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
'         & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
'         & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
'         & strLinea(11) & "','" & vAseDocDetalle & "','" & vAseDocDeposito & "')"
' Call ConectionExecute(strSQL)
'
' 'ASIENTO
' If curIntC > 0 Then
'   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntC * pTipoCambio & ",'C','" & rs!cod_Divisa _
'          & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintc _
'          & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
'   Call ConectionExecute(strSQL)
' End If
'
' If curIntM > 0 Then
'   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntM * pTipoCambio & ",'C','" & rs!cod_Divisa _
'          & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintm _
'          & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
'   Call ConectionExecute(strSQL)
' End If
'
' If curCargo > 0 Then
' 'Detallar Cargos
'   strSQL = "exec spCxC_DocumentoAfectacionCargos '" & pTipoDoc & "','" & pNumDoc & "'"
'   Call OpenRecordSet(rsTmp, strSQL, 0)
'   Do While Not rsTmp.EOF
'         strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & rsTmp!Monto * pTipoCambio & ",'C','" & rs!cod_Divisa _
'                & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rsTmp!cod_unidad & "','" & rsTmp!cod_centro_costo & "','" & rsTmp!cod_cuenta _
'                & "','" & rsTmp!Operacion & "','" & rsTmp!cod_Concepto & "','" & vAseDocDeposito & "'"
'         Call ConectionExecute(strSQL)
'         rsTmp.MoveNext
'   Loop
'   rsTmp.Close
' End If
'
'
' If curAmortiza > 0 Then
'   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curAmortiza * pTipoCambio & ",'C','" & rs!cod_Divisa _
'          & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaamortiza _
'          & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
'   Call ConectionExecute(strSQL)
' End If
'
'
'  If curIntC + curIntM + curCargo + curAmortiza > 0 Then
'     'Procesa Formas de Pago (Registro Final / Asiento de Pago)
'      strSQL = "exec spCajas_DesglocePagosDocFinal '" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ",'" & ModuloCajas.mTiquete _
'              & "','" & ModuloCajas.mUsuario & "','" & pTipoDoc & "','" & pNumDoc & "','" & ModuloCajas.mUnidad _
'              & "','" & rs!Operacion & "','" & rs!cod_Concepto & "'"
'      Call ConectionExecute(strSQL)
' End If
'
'rs.Close


End Sub



Private Sub sbAbono()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipoDoc As String, vNumDoc As String, i As Integer
Dim pOperacion As Long, pFactura As String, pAbono As Currency


Me.MousePointer = vbHourglass

On Error GoTo vError

vTipoDoc = cboTipoDoc.ItemData(cboTipoDoc.ListIndex)
vNumDoc = fxDocumentoConsecutivo(vTipoDoc)


'Procesa las Facturas Canceladas
strSQL = ""
With lsw.ListItems
  For i = 1 To .Count
        pOperacion = .Item(i)
        pFactura = .Item(i).SubItems(1)
        pAbono = CCur(.Item(i).SubItems(2))
         
        strSQL = strSQL & Space(10) & "exec spCxC_Operacion_Factura_Cancela " & pOperacion & ",'" & pFactura & "'," & pAbono _
               & ",'" & vTipoDoc & "','" & vNumDoc & "','" & glogon.Usuario & "'"
  Next i
End With

If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
End If


'Procesa Abono + Documento + Asiento
strSQL = "exec spCxC_Operacion_Factura_Cancela_Abono '" & vTipoDoc & "','" & vNumDoc & "','" & ModuloCajas.mCaja _
       & "'," & ModuloCajas.mApertura & ",'" & ModuloCajas.mTiquete & "','" & glogon.Usuario & "',1"
Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Registra Cancelación de Facturas> Pagador Id: " & cboPagador.ItemData(cboPagador.ListIndex))

'Imprime el Comprobante
Call sbImprimeRecibo(vNumDoc, vTipoDoc)

Me.MousePointer = vbDefault

strSQL = " - Abono aplicado, con : " & cboTipoDoc.Text & " ...No.: " & vNumDoc & vbCrLf _
       & " - Desea Realizar Otra Transacción?"

i = MsgBox(strSQL, vbYesNo)
If i = vbYes Then
    Call sbPagagor_Load
    txtTotalCajas.Text = 0
Else
    Unload Me
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tlbCajas_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim iRespuesta As Integer

On Error GoTo vError

Select Case Button.Key
  Case "Cancelar"
     Call sbPagagor_Load
     
  Case "Desgloce"
        If Not IsNumeric(txtTotalPagar.Text) Then txtTotalPagar.Text = 0
        If Not ModuloCajas.mConceptoValida Then
           MsgBox "Esta caja no está autorizada para registrar movimientos a este Concepto de Cuentas por Cobrar", vbExclamation
           Exit Sub
        End If
                
        If cboDivisas.ListCount = 0 Then
            ModuloCajas.mDivisa = "COL"
        Else
            ModuloCajas.mDivisa = cboDivisas.Text
        End If
        ModuloCajas.mTotalAplicar = CCur(txtTotalPagar.Text)
        
        If ModuloCajas.mTotalAplicar = 0 Then
            MsgBox "No se ha especificado ningún monto a detallar?", vbExclamation
            Exit Sub
        End If
        
        ModuloCajas.mServicio = "Abonos a Cuentas por Cobrar"
        
        Call sbFormsCall("frmCajas_DetallePago", vbModal, 0, 0, False, Me)
        
        txtTotalCajas.Text = Format(ModuloCajas.mTotalDetallado, "Standard")
        
        
        If txtTotalCajas.Text <> txtTotalPagar.Text Then
           txtTotalCajas.BackColor = vbRed
        Else
           txtTotalCajas.BackColor = vbWhite
        End If

        txtDiferencia.Text = Format((CCur(txtTotalCajas.Text) - CCur(txtTotalPagar.Text)), "Standard")
  Case "Aplicar"

'    If Not fxVerifica Then Exit Sub
     If CCur(txtTotalCajas.Text) <> CCur(txtTotalPagar.Text) Then
           MsgBox "No se ha recaudado el total a cancelar/pagar de las facturas seleccionadas!", vbInformation
           Exit Sub
     End If
     
     If CCur(txtTotalPagar.Text) = 0 Then
           MsgBox "No se ha indicado ninguna factura a cancelar!", vbInformation
           Exit Sub
     End If
     
     iRespuesta = MsgBox("Esta seguro de realizar el abono a las facturas?", vbYesNo)
     If iRespuesta = vbYes Then
        Call sbAbono
'        Call sbConsultaCliente
     End If


End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCliente_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  Call sbFacturas_Carga(True)
End If
End Sub

Private Sub txtFactura_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  Call sbFacturas_Carga(True)
End If
End Sub
