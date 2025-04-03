VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmPosCajas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Cajas"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   8775
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   6720
      Width           =   8535
      _Version        =   1310723
      _ExtentX        =   15055
      _ExtentY        =   2143
      _StockProps     =   79
      Caption         =   "Información de Saldos y Cierres"
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
      Begin XtremeSuiteControls.FlatEdit txtUltApertura 
         Height          =   312
         Left            =   1920
         TabIndex        =   44
         Top             =   360
         Width           =   2052
         _Version        =   1310723
         _ExtentX        =   3619
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
      Begin XtremeSuiteControls.FlatEdit txtUltCierre 
         Height          =   312
         Left            =   1920
         TabIndex        =   45
         Top             =   720
         Width           =   2052
         _Version        =   1310723
         _ExtentX        =   3619
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
      Begin XtremeSuiteControls.FlatEdit txtSaldoEfectivo 
         Height          =   312
         Left            =   6240
         TabIndex        =   46
         Top             =   360
         Width           =   2052
         _Version        =   1310723
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSaldoDocumentos 
         Height          =   312
         Left            =   6240
         TabIndex        =   47
         Top             =   720
         Width           =   2052
         _Version        =   1310723
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ultima Apertura"
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
         Index           =   8
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1692
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ultimo Cierre"
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
         Index           =   9
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1692
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Saldo Actual Documentos"
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
         Index           =   11
         Left            =   3960
         TabIndex        =   9
         Top             =   720
         Width           =   2292
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Saldo Actual Efectivo"
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
         Index           =   12
         Left            =   3960
         TabIndex        =   8
         Top             =   360
         Width           =   2292
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
               NumButtonMenus  =   5
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
                  Key             =   "repAntiguedadSaldos"
                  Text            =   "Antiguedad Saldos"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repPagosPendientes"
                  Text            =   "Pagos Pendientes"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2652
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   8532
      _Version        =   1310723
      _ExtentX        =   15049
      _ExtentY        =   4678
      _StockProps     =   79
      Caption         =   "Comportamiento para la Caja"
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
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   252
         Left            =   7920
         TabIndex        =   23
         Top             =   1920
         Width           =   372
         _Version        =   1310723
         _ExtentX        =   656
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "..."
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkBloqueo 
         Height          =   252
         Left            =   5760
         TabIndex        =   18
         Top             =   360
         Width           =   2172
         _Version        =   1310723
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Caja Bloqueada?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkModFechas 
         Height          =   252
         Left            =   5760
         TabIndex        =   19
         Top             =   720
         Width           =   2172
         _Version        =   1310723
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Modifica Fechas?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkModPrecios 
         Height          =   252
         Left            =   5760
         TabIndex        =   20
         Top             =   1080
         Width           =   2172
         _Version        =   1310723
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Modifica Precios?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkVentaExenta 
         Height          =   252
         Left            =   5760
         TabIndex        =   26
         Top             =   1440
         Width           =   2172
         _Version        =   1310723
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Venta Exenta Default?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   2160
         TabIndex        =   28
         Top             =   360
         Width           =   2412
         _Version        =   1310723
         _ExtentX        =   4260
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboDisplay 
         Height          =   312
         Left            =   2160
         TabIndex        =   29
         Top             =   720
         Width           =   2412
         _Version        =   1310723
         _ExtentX        =   4260
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtAutorizador 
         Height          =   312
         Left            =   2160
         TabIndex        =   34
         Top             =   1080
         Width           =   2412
         _Version        =   1310723
         _ExtentX        =   4254
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
      Begin XtremeSuiteControls.CheckBox chkFormatoEspecial 
         Height          =   375
         Left            =   2160
         TabIndex        =   21
         Top             =   1440
         Width           =   3495
         _Version        =   1310723
         _ExtentX        =   6165
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Utiliza formato de Impresión Especial?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         MultiLine       =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtArchivoEspecial 
         Height          =   552
         Left            =   2160
         TabIndex        =   35
         Top             =   1920
         Width           =   5652
         _Version        =   1310723
         _ExtentX        =   9970
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
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Autorizador de esta Caja"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   15
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Archivo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   14
         Left            =   480
         TabIndex        =   22
         Top             =   1920
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Display  POS"
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
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1692
      End
      Begin VB.Label Label1 
         Caption         =   "Formato de la Factura"
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
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1692
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   8535
      _Version        =   1310723
      _ExtentX        =   15055
      _ExtentY        =   4260
      _StockProps     =   79
      Caption         =   "Valores Predeterminados en el POS"
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
      Begin XtremeSuiteControls.FlatEdit txtCodCliente 
         Height          =   315
         Left            =   1080
         TabIndex        =   36
         Top             =   840
         Width           =   1935
         _Version        =   1310723
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDescCliente 
         Height          =   315
         Left            =   3000
         TabIndex        =   37
         Top             =   840
         Width           =   5415
         _Version        =   1310723
         _ExtentX        =   9546
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
      Begin XtremeSuiteControls.FlatEdit txtCodBodega 
         Height          =   315
         Left            =   1080
         TabIndex        =   38
         Top             =   1200
         Width           =   1935
         _Version        =   1310723
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDescBodega 
         Height          =   315
         Left            =   3000
         TabIndex        =   39
         Top             =   1200
         Width           =   5415
         _Version        =   1310723
         _ExtentX        =   9546
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
      Begin XtremeSuiteControls.FlatEdit txtCodPrecio 
         Height          =   315
         Left            =   1080
         TabIndex        =   40
         Top             =   1560
         Width           =   1935
         _Version        =   1310723
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDescPrecio 
         Height          =   315
         Left            =   3000
         TabIndex        =   41
         Top             =   1560
         Width           =   5415
         _Version        =   1310723
         _ExtentX        =   9546
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
      Begin XtremeSuiteControls.FlatEdit txtCodAgente 
         Height          =   315
         Left            =   1080
         TabIndex        =   42
         Top             =   1920
         Width           =   1935
         _Version        =   1310723
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDescAgente 
         Height          =   315
         Left            =   3000
         TabIndex        =   43
         Top             =   1920
         Width           =   5415
         _Version        =   1310723
         _ExtentX        =   9546
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
      Begin XtremeSuiteControls.FlatEdit txtCentroCosto 
         Height          =   315
         Left            =   1080
         TabIndex        =   48
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   480
         Width           =   1935
         _Version        =   1310723
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
      Begin XtremeSuiteControls.FlatEdit txtCentroCostoDesc 
         Height          =   315
         Left            =   3000
         TabIndex        =   49
         Top             =   480
         Width           =   5415
         _Version        =   1310723
         _ExtentX        =   9546
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
      Begin VB.Label Label1 
         Caption         =   "C.Costo"
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
         Left            =   240
         TabIndex        =   50
         ToolTipText     =   "Centro de Costo"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Agente"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Precio"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   735
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
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8160
      TabIndex        =   24
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   6360
      TabIndex        =   27
      Top             =   840
      Width           =   1692
      _Version        =   1310723
      _ExtentX        =   2990
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   960
      TabIndex        =   30
      Top             =   480
      Width           =   1212
      _Version        =   1310723
      _ExtentX        =   2138
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   2160
      TabIndex        =   31
      Top             =   480
      Width           =   5892
      _Version        =   1310723
      _ExtentX        =   10393
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   960
      TabIndex        =   32
      Top             =   840
      Width           =   1572
      _Version        =   1310723
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtClave 
      Height          =   312
      Left            =   3840
      TabIndex        =   33
      Top             =   840
      Width           =   1692
      _Version        =   1310723
      _ExtentX        =   2984
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
      PasswordChar    =   "*"
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Index           =   4
      Left            =   5640
      TabIndex        =   4
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave Inicial"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
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
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Caja"
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
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmPosCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vUsuario As String, vScroll As Boolean

Private Sub btnArchivo_Click()
frmContenedor.CD.FileName = "*.rpt"
frmContenedor.CD.ShowOpen


If frmContenedor.CD.FileName <> "" And frmContenedor.CD.FileName <> "*.rpt" Then
  txtArchivoEspecial.Text = Dir(frmContenedor.CD.FileName)
Else
   MsgBox "No selecciono ningun archivo"
End If
frmContenedor.CD.FileName = ""

End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodCliente.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 COD_CAJA,USUARIO from PV_CAJAS"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COD_CAJA > '" & txtCodigo & "' order by COD_CAJA asc"
    Else
       strSQL = strSQL & " where COD_CAJA < '" & txtCodigo & "' order by COD_CAJA desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo = rs!cod_caja
      Call sbConsulta(rs!cod_caja, rs!Usuario)
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

On Error GoTo vError
 
 vModulo = 33
 
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 
cboEstado.Clear
cboEstado.AddItem "Activa"
cboEstado.AddItem "Inactiva"


cbo.Clear
cbo.AddItem "Factura Tiquete"
cbo.ItemData(cbo.ListCount - 1) = CStr(1)
cbo.AddItem "Factura Clasica"
cbo.ItemData(cbo.ListCount - 1) = CStr(2)

cboDisplay.Clear
cboDisplay.AddItem "Default"
cboDisplay.AddItem "Market"
cboDisplay.AddItem "Vnt.Prd."
cboDisplay.AddItem "Servicios"
 
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()
vCodigo = ""
vUsuario = ""

txtCodigo = ""

txtUsuario = ""
txtClave = ""

cboEstado.Text = "Activa"

cbo.Text = "Factura Clasica"

cboDisplay.Text = "Default"

txtCodAgente = ""
txtCodBodega = ""
txtCodCliente = ""
txtCodPrecio = ""

txtDescAgente = ""
txtDescBodega = ""
txtDescCliente = ""
txtDescPrecio = ""

txtCentroCosto.Text = ""
txtCentroCostoDesc.Text = ""

txtUltApertura = ""
txtUltCierre = ""

txtSaldoDocumentos = "0"
txtSaldoEfectivo = "0"

chkBloqueo.Value = xtpUnchecked
chkModFechas.Value = xtpUnchecked
chkModPrecios.Value = xtpUnchecked
chkVentaExenta.Value = xtpUnchecked



chKFormatoEspecial.Value = xtpUnchecked
txtArchivoEspecial.Text = ""
txtAutorizador.Text = ""


End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
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
        Call sbConsulta(vCodigo, vUsuario)
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "cod_caja"
       gBusquedas.Orden = "cod_caja"
       gBusquedas.Consulta = "select cod_caja,usuario,nombre from pv_cajas"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtUsuario = gBusquedas.Resultado2
       Call sbConsulta(txtCodigo, txtUsuario)
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(xCodigo As String, xUsuario As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select C.*, isnull(Bo.Descripcion,'') as 'Bodega_Desc', isnull(Cl.Nombre,'') as 'Cliente_Desc' " _
       & " , isnull(Tp.Descripcion,'') as 'Precio_Desc', isnull(Ag.Nombre,'') as 'Agente_Desc' " _
       & ", isnull(Cc.Descripcion,'') as  'CentroCostoDesc'" _
       & " from pv_cajas C left join PV_Bodegas Bo on C.Def_Bodega = Bo.cod_Bodega" _
       & " left join PV_Clientes Cl on C.def_cliente = Cl.Cedula" _
       & " left join pv_tipos_precios Tp on C.def_precio = Tp.cod_precio" _
       & " left join PV_Agentes Ag on C.def_agente = Ag.cod_Agente" _
       & " left join CNTX_CENTRO_COSTOS Cc on C.COD_CENTRO_COSTO = Cc.COD_CENTRO_COSTO and Cc.COD_CONTABILIDAD = " & GLOBALES.gEnlace _
       & " where C.cod_caja = '" & Trim(xCodigo) & "' and C.usuario = '" & txtUsuario & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = Trim(rs!cod_caja)
  vUsuario = rs!Usuario & ""
      
  txtCodigo = Trim(rs!cod_caja)
  txtUsuario = rs!Usuario & ""
  txtNombre = rs!Nombre & ""
  
  If rs!Estado = "A" Then
    cboEstado.Text = "Activa"
  Else
    cboEstado.Text = "InActiva"
  End If
  
  txtCodAgente = Trim(rs!def_agente)
  txtDescAgente = rs!Agente_Desc
  
  txtCodBodega = Trim(rs!def_bodega)
  txtDescBodega.Text = rs!Bodega_desc
'  txtDescBodega = fxSIFCCodigos("D", rs!def_bodega, "bodegas")
   
  txtCodPrecio = Trim(rs!def_precio)
  txtDescPrecio.Text = rs!Precio_Desc
  
  txtCodCliente = Trim(rs!def_cliente)
  txtDescCliente.Text = rs!Cliente_Desc
  
  txtCentroCosto.Text = Trim(rs!cod_centro_costo & "")
  txtCentroCostoDesc.Text = rs!CentroCostoDesc
  
  txtUltApertura = rs!ult_apertura & ""
  txtUltCierre = rs!ult_cierre & ""
  
  txtSaldoDocumentos = Format(IIf(IsNull(rs!saldo_documentos), 0, rs!saldo_documentos), "Standard")
  txtSaldoEfectivo = Format(IIf(IsNull(rs!saldo_efectivo), 0, rs!saldo_efectivo), "Standard")
    

   chkBloqueo.Value = rs!bloqueo
   chkModPrecios.Value = rs!modifica_precio
   chkModFechas.Value = rs!modifica_fechas
   chkVentaExenta.Value = rs!venta_Exenta
   
  Select Case rs!Formato_Factura
     Case 0 'Boucher
       cbo.Text = "Factura Tiquete"
     Case 1 'Clasica
        cbo.Text = "Factura Clasica"
  End Select
   
   
  Select Case rs!Formato_Factura
     Case 0 'Boucher
       cbo.Text = "Factura Tiquete"
     Case 1 'Clasica
        cbo.Text = "Factura Clasica"
  End Select
    
  txtAutorizador.Text = rs!Autorizador & ""
  chKFormatoEspecial.Value = rs!FORMATO_ESPECIAL
  txtArchivoEspecial.Text = rs!FORMATO_ESPECIAL_ARCHIVO

  Select Case Trim(rs!POS_DISPLAY)
    Case "E01" 'Default
        cboDisplay.Text = "Default"
    Case "M01" 'Market
        cboDisplay.Text = "Market"
    Case "VPD" 'Retail Productos Normal
        cboDisplay.Text = "Vnt.Prd."
    Case "SRV"
        cboDisplay.Text = "Servicios"
  End Select
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

If txtUsuario = "" Then vMensaje = vMensaje & vbCrLf & " - El Usuario no es válido ..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset, pPosDisplay As String

On Error GoTo vError

  Select Case cboDisplay.Text
    Case "Default"
        pPosDisplay = "E01"
    Case "Market"
        pPosDisplay = "M01"
    Case "Vnt.Prd." 'Retail Productos Normal
        pPosDisplay = "VPD"
    Case "Servicios"
        pPosDisplay = "SRV"
  End Select
   

If vEdita Then
  strSQL = "update pv_cajas set usuario = '" & txtUsuario & "',estado = '" _
         & Mid(cboEstado.Text, 1, 1) & "',nombre = '" & UCase(Trim(txtNombre)) & "',def_cliente = '" & txtCodCliente _
         & "',def_bodega = '" & txtCodBodega & "',def_agente = '" & txtCodAgente _
         & "',def_precio = '" & txtCodPrecio & "', bloqueo = " & chkBloqueo.Value _
         & ",formato_factura = " & cbo.ItemData(cbo.ListIndex) _
         & ",modifica_precio = " & chkModPrecios.Value _
         & ",modifica_fechas = " & chkModFechas.Value _
         & ",venta_exenta = " & chkVentaExenta.Value _
         & ",Autorizador = '" & txtAutorizador.Text & "', Pos_Display = '" & pPosDisplay _
         & "',FORMATO_ESPECIAL = " & chKFormatoEspecial.Value & ",FORMATO_ESPECIAL_ARCHIVO = '" & txtArchivoEspecial.Text & "'" _
         & ", Cod_Centro_Costo = '" & Trim(txtCentroCosto.Text) & "'" _
         & " where cod_caja = '" & vCodigo & "' and usuario = '" & vUsuario & "'"
  
  Call ConectionExecute(strSQL)
  Call Bitacora("Modifica", "Cajas: " & vCodigo & " US: " & vUsuario)

Else
  vCodigo = txtCodigo
  vUsuario = txtUsuario
  
  strSQL = "insert into pv_cajas(cod_caja,usuario,clave,nombre,estado,def_cliente" _
         & ",def_bodega,def_precio,def_agente,saldo_efectivo,saldo_documentos,bloqueo" _
         & ",formato_factura,modifica_precio,modifica_fechas,Autorizador,Pos_Display" _
         & ",FORMATO_ESPECIAL,FORMATO_ESPECIAL_ARCHIVO, Venta_Exenta, cod_centro_costo ) values('" _
         & vCodigo & "','" & vUsuario & "','" & fxPosEncrypta(txtClave) & "','" & UCase(txtNombre) _
         & "','" & Mid(cboEstado.Text, 1, 1) & "','" & txtCodCliente _
         & "','" & txtCodBodega & "','" & txtCodPrecio & "','" & txtCodAgente & "',0,0," & chkBloqueo.Value _
         & "," & cbo.ItemData(cbo.ListIndex) & "," & chkModPrecios.Value & "," & chkModFechas.Value _
         & ",'" & txtAutorizador.Text & "','" & pPosDisplay & "'," & chKFormatoEspecial.Value _
         & ",'" & txtArchivoEspecial.Text & "'," & chkVentaExenta.Value & ",'" & Trim(txtCentroCosto.Text) & "')"
   
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Registra", "Cajas: " & vCodigo & " US: " & vUsuario)
    
 
End If


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
  strSQL = "delete pv_cajas where cod_caja = '" & vCodigo & "' and usuario = '" & vUsuario & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Cajas: " & vCodigo & " US: " & vUsuario)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
  
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtAutorizador_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select nombre,descripcion from usuarios"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  txtUsuario = gBusquedas.Resultado
End If

End Sub



Private Sub txtCentroCosto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "COD_CENTRO_COSTO"
  gBusquedas.Orden = "COD_CENTRO_COSTO"
  gBusquedas.Consulta = "SELECT COD_CENTRO_COSTO, DESCRIPCION " _
                     & " From CNTX_CENTRO_COSTOS"
  gBusquedas.Filtro = " AND COD_CONTABILIDAD = " & GLOBALES.gEnlace & " And ACTIVO = 1"
  frmBusquedas.Show vbModal
  txtCentroCosto.Text = gBusquedas.Resultado
  txtCentroCostoDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus
End Sub

Private Sub txtCodAgente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescAgente.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_agente"
  gBusquedas.Orden = "cod_agente"
  gBusquedas.Consulta = "select cod_agente,nombre from pv_agentes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodAgente = gBusquedas.Resultado
  txtDescAgente = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCodAgente_LostFocus()
txtDescAgente = fxSIFCCodigos("D", txtCodAgente, "agentes")
End Sub

Private Sub txtCodBodega_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescBodega.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_bodega"
  gBusquedas.Orden = "cod_bodega"
  gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodBodega = gBusquedas.Resultado
  txtDescBodega = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCodBodega_LostFocus()
txtDescBodega = fxSIFCCodigos("D", txtCodBodega, "bodegas")
End Sub

Private Sub txtCodCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescCliente.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodCliente = gBusquedas.Resultado
  txtDescCliente = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCodCliente_LostFocus()
txtDescCliente = fxSIFCCodigos("D", txtCodCliente, "clientes")
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo, txtUsuario)
  txtNombre.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_caja"
  gBusquedas.Orden = "cod_caja"
  gBusquedas.Consulta = "select cod_caja,usuario,nombre from pv_cajas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtUsuario = gBusquedas.Resultado2
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado, gBusquedas.Resultado2)
End If

End Sub

Private Sub txtCodigo_LostFocus()
txtNombre = fxSIFCCodigos("D", txtCodigo, "Cajas")
End Sub

Private Sub txtCodPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescPrecio.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_precio"
  gBusquedas.Orden = "cod_precio"
  gBusquedas.Consulta = "select cod_precio,descripcion from pv_tipos_precios"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodPrecio = gBusquedas.Resultado
  txtDescPrecio = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCodPrecio_LostFocus()
txtDescPrecio = fxSIFCCodigos("D", txtCodPrecio, "Precios")
End Sub

Private Sub txtDescAgente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cod_agente,nombre from pv_agentes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodAgente = gBusquedas.Resultado
  txtDescAgente = gBusquedas.Resultado2
End If
End Sub

Private Sub txtDescBodega_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodPrecio.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodBodega = gBusquedas.Resultado
  txtDescBodega = gBusquedas.Resultado2
End If
End Sub

Private Sub txtDescCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodBodega.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodCliente = gBusquedas.Resultado
  txtDescCliente = gBusquedas.Resultado2
End If

End Sub

Private Sub txtDescPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodAgente.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_precio,descripcion from pv_tipos_precios"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodPrecio = gBusquedas.Resultado
  txtDescPrecio = gBusquedas.Resultado2
End If
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUsuario.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cod_caja,usuario,nombre from pv_cajas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtUsuario = gBusquedas.Resultado2
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado, gBusquedas.Resultado2)
End If
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtClave.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select nombre,descripcion from usuarios"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  txtUsuario = gBusquedas.Resultado
End If
End Sub
