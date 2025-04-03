VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmGenParametrosPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros de Procedimientos Generales"
   ClientHeight    =   5148
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5148
   ScaleWidth      =   8100
   Begin TabDlg.SSTab ssTab 
      Height          =   3612
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   7692
      _ExtentX        =   13568
      _ExtentY        =   6371
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmGenParametrosPro.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCambiaParGen"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkNotifica"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkCalcularImp"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkPermiteActualizarCostos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkUtilizaCostoUltCompra"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "optIV(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "optIV(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkUtilizaModoTrasaccional"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "CxP"
      TabPicture(1)   =   "frmGenParametrosPro.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboPago"
      Tab(1).Control(1)=   "cboND"
      Tab(1).Control(2)=   "cboNC"
      Tab(1).Control(3)=   "cmdCambiaParCxP"
      Tab(1).Control(4)=   "Label1(8)"
      Tab(1).Control(5)=   "Label1(6)"
      Tab(1).Control(6)=   "Label1(5)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Inv / Compras"
      TabPicture(2)   =   "frmGenParametrosPro.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboCompras"
      Tab(2).Control(1)=   "cboTraslado"
      Tab(2).Control(2)=   "cboSalida"
      Tab(2).Control(3)=   "cboEntrada"
      Tab(2).Control(4)=   "cmdCambiaParInv"
      Tab(2).Control(5)=   "Label1(4)"
      Tab(2).Control(6)=   "Label1(2)"
      Tab(2).Control(7)=   "Label1(1)"
      Tab(2).Control(8)=   "Label1(0)"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "POS"
      TabPicture(3)   =   "frmGenParametrosPro.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtAutoCambioCL"
      Tab(3).Control(1)=   "txtAutoCambioUS"
      Tab(3).Control(2)=   "txtAutoReImpCL"
      Tab(3).Control(3)=   "txtAutoReImpUS"
      Tab(3).Control(4)=   "cboRecibos"
      Tab(3).Control(5)=   "cboFactura"
      Tab(3).Control(6)=   "cmdCambiaParPos"
      Tab(3).Control(7)=   "cmdCambiaUserReImp"
      Tab(3).Control(8)=   "cmdCambioPrecioUser"
      Tab(3).Control(9)=   "Label1(16)"
      Tab(3).Control(10)=   "Label1(15)"
      Tab(3).Control(11)=   "Label1(14)"
      Tab(3).Control(12)=   "Label1(13)"
      Tab(3).Control(13)=   "Label1(12)"
      Tab(3).Control(14)=   "Label1(11)"
      Tab(3).Control(15)=   "Label1(3)"
      Tab(3).Control(16)=   "Label1(7)"
      Tab(3).ControlCount=   17
      TabCaption(4)   =   "Tipos de Cambio"
      TabPicture(4)   =   "frmGenParametrosPro.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtTCVenta"
      Tab(4).Control(1)=   "txtTCCompra"
      Tab(4).Control(2)=   "cmdCambiaTC"
      Tab(4).Control(3)=   "Label1(10)"
      Tab(4).Control(4)=   "Label1(9)"
      Tab(4).ControlCount=   5
      Begin VB.CheckBox chkUtilizaModoTrasaccional 
         Caption         =   "Ultiliza Modo de Generacion de Asientos Transaccional y no por Notas"
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
         Left            =   240
         TabIndex        =   40
         Top             =   1920
         Value           =   1  'Checked
         Width           =   6372
      End
      Begin VB.TextBox txtTCVenta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -72240
         TabIndex        =   39
         Text            =   "403.15"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtTCCompra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -72240
         TabIndex        =   38
         Text            =   "402.07"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtAutoCambioCL 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         IMEMode         =   3  'DISABLE
         Left            =   -72720
         PasswordChar    =   "*"
         TabIndex        =   37
         Text            =   "system"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox txtAutoCambioUS 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -72720
         TabIndex        =   36
         Text            =   "sa"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox txtAutoReImpCL 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         IMEMode         =   3  'DISABLE
         Left            =   -72720
         PasswordChar    =   "*"
         TabIndex        =   35
         Text            =   "system"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txtAutoReImpUS 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -72720
         TabIndex        =   34
         Text            =   "sa"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.ComboBox cboRecibos 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   -72720
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cboFactura 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   -72720
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cboCompras 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   -72240
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1920
         Width           =   2172
      End
      Begin VB.ComboBox cboTraslado 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   -72240
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1560
         Width           =   2172
      End
      Begin VB.ComboBox cboSalida 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   -72240
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1200
         Width           =   2172
      End
      Begin VB.ComboBox cboEntrada 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   -72240
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   840
         Width           =   2172
      End
      Begin VB.ComboBox cboPago 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   -72120
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1560
         Width           =   2172
      End
      Begin VB.ComboBox cboND 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   -72120
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1200
         Width           =   2172
      End
      Begin VB.ComboBox cboNC 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   -72120
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   840
         Width           =   2172
      End
      Begin VB.OptionButton optIV 
         Caption         =   "Sub Total + Consumo"
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
         Left            =   1800
         TabIndex        =   24
         Top             =   2880
         Width           =   2055
      End
      Begin VB.OptionButton optIV 
         Caption         =   "Sub Total"
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
         Left            =   1800
         TabIndex        =   23
         Top             =   2520
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CheckBox chkUtilizaCostoUltCompra 
         Caption         =   "Utilizar Costo del Articulo, según última comprar cuando el costo es Cero "
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
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   6252
      End
      Begin VB.CheckBox chkPermiteActualizarCostos 
         Caption         =   "Permite Actualizar Costos de los articulos según última compra"
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
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Value           =   1  'Checked
         Width           =   6372
      End
      Begin VB.CheckBox chkCalcularImp 
         Caption         =   "Cálcular Impuestos después de Aplicados los descuentos"
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
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Value           =   1  'Checked
         Width           =   6372
      End
      Begin VB.CheckBox chkNotifica 
         Caption         =   "Notificar si una factura deja por debajo del minimo un articulo"
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
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   6372
      End
      Begin XtremeSuiteControls.PushButton cmdCambiaParGen 
         Height          =   540
         Left            =   6000
         TabIndex        =   41
         Top             =   3000
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   952
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmGenParametrosPro.frx":008C
      End
      Begin XtremeSuiteControls.PushButton cmdCambiaParCxP 
         Height          =   540
         Left            =   -69000
         TabIndex        =   42
         Top             =   3000
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   952
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmGenParametrosPro.frx":0791
      End
      Begin XtremeSuiteControls.PushButton cmdCambiaParInv 
         Height          =   540
         Left            =   -69000
         TabIndex        =   43
         Top             =   3000
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   952
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmGenParametrosPro.frx":0E96
      End
      Begin XtremeSuiteControls.PushButton cmdCambiaParPos 
         Height          =   540
         Left            =   -69120
         TabIndex        =   44
         Top             =   720
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   952
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmGenParametrosPro.frx":159B
      End
      Begin XtremeSuiteControls.PushButton cmdCambiaUserReImp 
         Height          =   540
         Left            =   -69120
         TabIndex        =   45
         Top             =   1800
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   952
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmGenParametrosPro.frx":1CA0
      End
      Begin XtremeSuiteControls.PushButton cmdCambioPrecioUser 
         Height          =   540
         Left            =   -69120
         TabIndex        =   46
         Top             =   2880
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   952
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmGenParametrosPro.frx":23A5
      End
      Begin XtremeSuiteControls.PushButton cmdCambiaTC 
         Height          =   540
         Left            =   -69480
         TabIndex        =   47
         Top             =   1080
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   952
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmGenParametrosPro.frx":2AAA
      End
      Begin VB.Label Label2 
         Caption         =   "Aplicar Impuesto de Ventas sobre"
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
         Left            =   240
         TabIndex        =   22
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Clave"
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
         Index           =   16
         Left            =   -74040
         TabIndex        =   17
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ID Usuario"
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
         Index           =   15
         Left            =   -74040
         TabIndex        =   16
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Clave"
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
         Index           =   14
         Left            =   -74040
         TabIndex        =   15
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ID Usuario"
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
         Index           =   13
         Left            =   -74040
         TabIndex        =   14
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Autorización Cambio Precio"
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
         Index           =   12
         Left            =   -74760
         TabIndex        =   13
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Autorización ReImpresión"
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
         Index           =   11
         Left            =   -74760
         TabIndex        =   12
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Cambio de Venta"
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
         Left            =   -74760
         TabIndex        =   11
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Cambio de Compra"
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
         Index           =   9
         Left            =   -74760
         TabIndex        =   10
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "T.C. para Facturas "
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
         Index           =   3
         Left            =   -74760
         TabIndex        =   9
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "T.C. para Recibos"
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
         Index           =   7
         Left            =   -74760
         TabIndex        =   8
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "T.C. para Pagos"
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
         Index           =   8
         Left            =   -74760
         TabIndex        =   7
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "T.C. para Notas de Débito"
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
         Index           =   6
         Left            =   -74760
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "T.C. para Notas de Crédito"
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
         Left            =   -74760
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "T.C. para Compras"
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
         Left            =   -74760
         TabIndex        =   4
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "T.C. para Traslados"
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
         Left            =   -74760
         TabIndex        =   3
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "T.C. para Salidas"
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
         Left            =   -74760
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "T.C. para Entradas"
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
         Left            =   -74760
         TabIndex        =   1
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Parámetros Generales para los módulos comerciales"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   852
      Left            =   2040
      TabIndex        =   48
      Top             =   120
      Width           =   4932
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13092
   End
End
Attribute VB_Name = "frmGenParametrosPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbCargaCboTC(cbo As ComboBox, Optional i As Byte = 1)

cbo.AddItem "T.C. Venta"
cbo.ItemData(cbo.NewIndex) = 1
cbo.AddItem "T.C. Compra"
cbo.ItemData(cbo.NewIndex) = 2
cbo.AddItem "T.C. Registro"
cbo.ItemData(cbo.NewIndex) = 3

Select Case i
 Case 1
   cbo.Text = "T.C. Venta"
 Case 2
   cbo.Text = "T.C. Compra"
 Case 3
   cbo.Text = "T.C. Registro"
End Select


End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select * from pv_parametros_mod"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
   strSQL = "insert pv_parametros_mod(chk_factura_min,chk_Descuento_Bifiv,chk_costo_ultComp,Chk_Costo_cero" _
          & ",chk_modo_asiento,aplica_iv_sobre,cxp_tc_nc,cxp_tc_nd,cxp_tc_pago,inv_tc_entrada,inv_tc_salida" _
          & ",inv_tc_traslado,inv_tc_compra,pos_tc_factura,pos_tc_recibo,pos_rei_user,pos_rei_clave,pos_cp_user,pos_cp_clave" _
          & ",tc_compra,tc_venta,tc_fecha,tc_usuario) values(0,1,0,1,1,'SB','C','V','C','C','C','C','C','V'" _
          & ",'C','','','','',0,0,dbo.MyGetdate(),'InI')"
  Call ConectionExecute(strSQL)
End If
rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Function fxTipoCambioTxT(vValor As String)
Select Case UCase(Trim(vValor))
 Case "V"
   fxTipoCambioTxT = "T.C. Venta"
 Case "C"
   fxTipoCambioTxT = "T.C. Compra"
 Case "R"
   fxTipoCambioTxT = "T.C. Registro"
End Select
End Function

Private Sub sbCargaParGen()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select * from pv_parametros_mod"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  chkNotifica.Value = rs!chk_factura_min
  chkCalcularImp = rs!chk_descuento_BifIV
  chkPermiteActualizarCostos = rs!chk_costo_ultComp
  chkUtilizaCostoUltCompra = rs!chk_costo_cero
  chkUtilizaModoTrasaccional = rs!chk_modo_asiento
  Select Case rs!aplica_iv_sobre
     Case "SB"  'Sobre SubTotal - Descuento
          optIV.Item(0).Value = True
     Case "SC"  'Sobre SubTotal + Impuesto de Consumo
          optIV.Item(1).Value = True
  End Select
  
  'CxP
  cboNC.Text = fxTipoCambioTxT(rs!cxp_tc_nc)
  cboND.Text = fxTipoCambioTxT(rs!cxp_tc_nd)
  cboPago.Text = fxTipoCambioTxT(rs!cxp_tc_pago)
  
  'Inv / Compras
  cboEntrada.Text = fxTipoCambioTxT(rs!inv_tc_entrada)
  cboSalida.Text = fxTipoCambioTxT(rs!inv_tc_salida)
  cboTraslado.Text = fxTipoCambioTxT(rs!inv_tc_traslado)
  cboCompras.Text = fxTipoCambioTxT(rs!inv_tc_compra)
  
  'POS
  
  cboFactura.Text = fxTipoCambioTxT(rs!pos_tc_factura)
  cboRecibos.Text = fxTipoCambioTxT(rs!pos_tc_recibo)
  
  txtAutoReImpUS.Text = rs!pos_rei_user
  txtAutoReImpCL.Tag = rs!pos_rei_clave
  txtAutoReImpCL.Text = ""
  
  txtAutoCambioUS.Text = rs!pos_cp_user
  txtAutoCambioCL.Tag = rs!pos_cp_clave
  txtAutoCambioCL.Text = ""
  
  'TC
  txtTCCompra.Text = Format(rs!tc_compra, "Standard")
  txtTCVenta.Text = Format(rs!tc_venta, "Standard")
  
End If
rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub



Private Sub cmdCambiaParCxP_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "update pv_parametros_mod set cxp_tc_nc = '" & Mid(cboND.Text, 6, 1) _
       & "',cxp_tc_nd = '" & Mid(cboNC.Text, 6, 1) _
       & "',cxp_tc_pago = '" & Mid(cboPago.Text, 6, 1) & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Actualiza", "Parámetros Generales de CxP")


MsgBox "Parámetros Actualizados Satisfactoriamente...", vbInformation
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdCambiaParGen_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "update pv_parametros_mod set chk_factura_min = " & chkNotifica.Value _
       & ",chk_descuento_bifIV = " & chkCalcularImp.Value _
       & ",chk_costo_UltComp = " & chkUtilizaCostoUltCompra.Value _
       & ",chk_costo_cero = " & chkPermiteActualizarCostos.Value _
       & ",chk_modo_asiento = " & chkUtilizaModoTrasaccional.Value _
       & ",aplica_iv_sobre = '" & IIf((optIV.Item(0).Value = True), "SB", "SC") & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Actualiza", "Parámetros Generales del SIF.C")

MsgBox "Parámetros Actualizados Satisfactoriamente...", vbInformation
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdCambiaParInv_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "update pv_parametros_mod set inv_tc_entrada = '" & Mid(cboEntrada.Text, 6, 1) _
       & "',inv_tc_salida = '" & Mid(cboSalida.Text, 6, 1) _
       & "',inv_tc_traslado = '" & Mid(cboTraslado.Text, 6, 1) _
       & "',inv_tc_compra = '" & Mid(cboCompras.Text, 6, 1) & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Actualiza", "Parámetros Generales de Inventario/Compras")

MsgBox "Parámetros Actualizados Satisfactoriamente...", vbInformation
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cmdCambiaParPos_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "update pv_parametros_mod set pos_tc_factura = '" & Mid(cboFactura.Text, 6, 1) _
       & "',pos_tc_recibo = '" & Mid(cboRecibos.Text, 6, 1)
Call ConectionExecute(strSQL)

Call Bitacora("Actualiza", "Parámetros Generales de POS/TC:Fac/Rec")

MsgBox "Parámetros Actualizados Satisfactoriamente...", vbInformation
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 34
End Sub

Private Sub Form_Load()

vModulo = 34

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

ssTab.Tab = 0

Call sbCargaCboTC(cboNC, 1)
Call sbCargaCboTC(cboND, 2)
Call sbCargaCboTC(cboPago, 3)

Call sbCargaCboTC(cboEntrada, 2)
Call sbCargaCboTC(cboSalida, 2)
Call sbCargaCboTC(cboTraslado, 3)
Call sbCargaCboTC(cboCompras, 2)


Call sbCargaCboTC(cboFactura, 1)
Call sbCargaCboTC(cboRecibos, 2)

Call sbInicializa
Call sbCargaParGen

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

