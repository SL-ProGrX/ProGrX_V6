VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmCR_Convenios_Consultas 
   Caption         =   "Consultas de Convenios"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15285
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   15285
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   3480
      Top             =   120
   End
   Begin VB.CheckBox chkFechas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2880
      TabIndex        =   56
      Top             =   4560
      Width           =   200
   End
   Begin VB.CheckBox chkTipos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2880
      TabIndex        =   45
      Top             =   120
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.Frame fraDetallePago 
      Caption         =   "[Orden] Convenio"
      Height          =   7335
      Left            =   4440
      TabIndex        =   2
      Top             =   480
      Width           =   8175
      Begin XtremeSuiteControls.ListView lswOrdenesCanceladas 
         Height          =   1332
         Left            =   1800
         TabIndex        =   62
         Top             =   5880
         Width           =   5892
         _Version        =   1310722
         _ExtentX        =   10393
         _ExtentY        =   2350
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
         Appearance      =   16
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenes Canceladas"
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
         Height          =   312
         Index           =   27
         Left            =   360
         TabIndex        =   44
         Top             =   5640
         Width           =   3012
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   43
         Top             =   4920
         Width           =   1932
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   26
         Left            =   4200
         TabIndex        =   42
         Top             =   4920
         Width           =   1692
      End
      Begin VB.Label lblFechaEmision 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1800
         TabIndex        =   41
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Emisión"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   25
         Left            =   360
         TabIndex        =   40
         Top             =   4920
         Width           =   1692
      End
      Begin VB.Label lblBanco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1800
         TabIndex        =   39
         Top             =   4560
         Width           =   5772
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   24
         Left            =   360
         TabIndex        =   38
         Top             =   4560
         Width           =   1572
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "N. Documento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   23
         Left            =   4200
         TabIndex        =   37
         Top             =   4200
         Width           =   1692
      End
      Begin VB.Label lblTesoreriaDocumento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   36
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label lblTipoDocumento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1800
         TabIndex        =   35
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Documento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   21
         Left            =   360
         TabIndex        =   34
         Top             =   4200
         Width           =   1692
      End
      Begin VB.Label lblBeneficiario 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1800
         TabIndex        =   33
         Top             =   3840
         Width           =   5772
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Beneficiario"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   20
         Left            =   360
         TabIndex        =   32
         Top             =   3840
         Width           =   1692
      End
      Begin VB.Label lblFechaSolicitud 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   31
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Solicitud"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   19
         Left            =   4200
         TabIndex        =   30
         Top             =   3480
         Width           =   1572
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "N. Solicitud"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   18
         Left            =   360
         TabIndex        =   29
         Top             =   3480
         Width           =   1452
      End
      Begin VB.Label lblNSolicitud 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1800
         TabIndex        =   28
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Desembolso"
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
         Height          =   312
         Index           =   17
         Left            =   360
         TabIndex        =   27
         Top             =   3120
         Width           =   2172
      End
      Begin VB.Label lblMontoDesembolso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   26
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Desembolso"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   552
         Index           =   16
         Left            =   4200
         TabIndex        =   25
         Top             =   2520
         Width           =   1452
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Neto "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   15
         Left            =   360
         TabIndex        =   24
         Top             =   2520
         Width           =   1812
      End
      Begin VB.Label lblFactTotalNeto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1800
         TabIndex        =   23
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblCargosAplicados 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1800
         TabIndex        =   22
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cargos Aplicados"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   14
         Left            =   360
         TabIndex        =   21
         Top             =   2160
         Width           =   1692
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   13
         Left            =   4200
         TabIndex        =   20
         Top             =   1800
         Width           =   1452
      End
      Begin VB.Label lblFacturaMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   19
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblFactura 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1800
         TabIndex        =   18
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   12
         Left            =   360
         TabIndex        =   17
         Top             =   1800
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuentas por Pagar"
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
         Height          =   312
         Index           =   11
         Left            =   360
         TabIndex        =   16
         Top             =   1440
         Width           =   3252
      End
      Begin VB.Label lblOrdenDocumento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   15
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   10
         Left            =   4200
         TabIndex        =   14
         Top             =   1080
         Width           =   1452
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   9
         Left            =   4200
         TabIndex        =   13
         Top             =   720
         Width           =   1452
      End
      Begin VB.Label lblUsuario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   12
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblFechaCorte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5640
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Corte"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   8
         Left            =   4200
         TabIndex        =   10
         Top             =   360
         Width           =   1452
      End
      Begin VB.Label lblOrdenTotalNeto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1800
         TabIndex        =   9
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Neto"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   7
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   1692
      End
      Begin VB.Label lblTotalDeducciones 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1800
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deducciones"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   1932
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bruto"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1452
      End
      Begin VB.Label lblTotalBruto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5928
      Left            =   3480
      TabIndex        =   0
      Top             =   600
      Width           =   11292
      _Version        =   524288
      _ExtentX        =   19918
      _ExtentY        =   10456
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
      MaxCols         =   19
      SpreadDesigner  =   "frmCR_Convenios_Consultas.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   7695
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Casos..:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Total Pagar..:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Total Orden..:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Total Deducciones..:"
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
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3852
      Left            =   120
      TabIndex        =   46
      Top             =   480
      Width           =   3012
      _Version        =   1310722
      _ExtentX        =   5313
      _ExtentY        =   6794
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
      Checkboxes      =   -1  'True
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   1560
      TabIndex        =   47
      Top             =   5760
      Width           =   1572
      _Version        =   1310722
      _ExtentX        =   2778
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
   Begin XtremeSuiteControls.ComboBox cboUsuarios 
      Height          =   312
      Left            =   1560
      TabIndex        =   48
      Top             =   6120
      Width           =   1572
      _Version        =   1310722
      _ExtentX        =   2778
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1560
      TabIndex        =   49
      Top             =   4920
      Width           =   1572
      _Version        =   1310722
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   1560
      TabIndex        =   50
      Top             =   5280
      Width           =   1572
      _Version        =   1310722
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   372
      Left            =   11040
      TabIndex        =   57
      Top             =   120
      Width           =   1212
      _Version        =   1310722
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCR_Convenios_Consultas.frx":0DBB
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnInforme 
      Height          =   372
      Left            =   12240
      TabIndex        =   58
      Top             =   120
      Width           =   1572
      _Version        =   1310722
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Informe"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCR_Convenios_Consultas.frx":14BB
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   372
      Left            =   13800
      TabIndex        =   59
      Top             =   120
      Width           =   1572
      _Version        =   1310722
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Exportar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCR_Convenios_Consultas.frx":1BC2
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   4560
      TabIndex        =   60
      Top             =   120
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   5520
      TabIndex        =   61
      Top             =   120
      Width           =   5412
      _Version        =   1310722
      _ExtentX        =   9546
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
      Appearance      =   2
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   22
      Left            =   120
      TabIndex        =   55
      Top             =   4920
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   5
      Left            =   120
      TabIndex        =   54
      Top             =   5280
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos Convenios ...:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   4
      Left            =   120
      TabIndex        =   53
      Top             =   120
      Width           =   1812
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   6
      Left            =   120
      TabIndex        =   52
      Top             =   5760
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   3
      Left            =   120
      TabIndex        =   51
      Top             =   6120
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "Convenio"
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
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgBanner 
      Height          =   9390
      Left            =   0
      Picture         =   "frmCR_Convenios_Consultas.frx":2493
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3285
   End
End
Attribute VB_Name = "frmCR_Convenios_Consultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim vNombreReporte As String

Private Sub chkDetallePago_Click()

End Sub

Private Sub btnBuscar_Click()
    Call sbConsulta
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders

vHeaders.Columnas = 19
vHeaders.Headers(3) = "Orden"
vHeaders.Headers(4) = "Factura"
vHeaders.Headers(5) = "Convenio"
vHeaders.Headers(6) = "Descripción"
vHeaders.Headers(7) = "Estado"
vHeaders.Headers(8) = "Tipo Convenio"
vHeaders.Headers(9) = "Fecha"
vHeaders.Headers(10) = "Usuario"
vHeaders.Headers(11) = "Monto Pagar"
vHeaders.Headers(12) = "Monto Orden"
vHeaders.Headers(13) = "Total Deducciones"
vHeaders.Headers(14) = "Comisión Recaudación"
vHeaders.Headers(15) = "Comisión Nuevos Créditos"
vHeaders.Headers(16) = "Reservas"
vHeaders.Headers(17) = "Cargos Cuentas x Pagar"
vHeaders.Headers(18) = "Fondos de Ahorro"
vHeaders.Headers(19) = "IVA"


    
Call sbSIFGridExportar(vGrid, vHeaders, "Convenios_ListadoOrdenes")

End Sub

Private Sub btnInforme_Click()
Dim i As Integer

i = MsgBox("Desea el visualizar el Informe Resumen?", vbYesNo)

If i = vbYes Then
      Call sbReporteCRD("Rsm")
Else
      Call sbReporteCRD("Dt")
End If

End Sub

Private Sub chkFechas_Click()

  If chkFechas.Value = vbChecked Then
     dtpInicio.Enabled = False
  Else
     dtpInicio.Enabled = True
  End If
  
  dtpCorte.Enabled = dtpInicio.Enabled
    
End Sub

Private Sub chkTipos_Click()
Dim i As Integer

If chkTipos.Value = vbChecked Then
  lsw.Enabled = False
Else
  lsw.Enabled = True
End If

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTipos.Value
Next i

End Sub

Private Sub Form_Activate()
  vModulo = 16
End Sub

Private Sub Form_Load()

vModulo = 16
lsw.ColumnHeaders.Add , , "", 3150

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
vGrid.MaxRows = 0

With lswOrdenesCanceladas.ColumnHeaders
    .Clear
    .Add , , "No.Orden", 1200
    .Add , , "Fecha Pago", 2100
    .Add , , "Monto", 1500, vbRightJustify
End With

fraDetallePago.Visible = False
chkTipos.Value = vbChecked
lsw.Enabled = False
 

End Sub

Private Sub sbInicial()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


'Carga los tipos de convenios
lsw.ListItems.Clear
strSQL = "select TIPO_CONVENIO, DESCRIPCION from CRD_CONVENIO_TIPO order by DESCRIPCION"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Descripcion)
     itmX.Tag = rs!TIPO_CONVENIO
     itmX.Checked = chkTipos.Value
 rs.MoveNext
Loop
rs.Close

cboEstado.Clear
cboEstado.AddItem "Abierta"
cboEstado.AddItem "Cerrada"
cboEstado.AddItem "Tramitada"
cboEstado.AddItem "TODOS"
cboEstado.Text = "TODOS"


strSQL = " select Nombre as Itmx, Nombre as IdX from Usuarios" _
       & " where Estado = 'A'"
Call sbCbo_Llena_New(cboUsuarios, strSQL, True, True)

Call chkFechas_Click

End Sub

Private Sub Form_Resize()

On Error Resume Next

imgBanner.Height = Me.Height

vGrid.Height = Me.Height - (vGrid.Top + StatusBarX.Height + 650)
vGrid.Width = Me.Width - (vGrid.Left + 450)


End Sub



Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim vHeaders As vGridHeaders

vHeaders.Columnas = 19
vHeaders.Headers(3) = "Orden"
vHeaders.Headers(4) = "Factura"
vHeaders.Headers(5) = "Convenio"
vHeaders.Headers(6) = "Descripción"
vHeaders.Headers(7) = "Estado"
vHeaders.Headers(8) = "Tipo Convenio"
vHeaders.Headers(9) = "Fecha"
vHeaders.Headers(10) = "Usuario"
vHeaders.Headers(11) = "Monto Pagar"
vHeaders.Headers(12) = "Monto Orden"
vHeaders.Headers(13) = "Total Deducciones"
vHeaders.Headers(14) = "Comisión Recaudación"
vHeaders.Headers(15) = "Comisión Nuevos Créditos"
vHeaders.Headers(16) = "Reservas"
vHeaders.Headers(17) = "Cargos Cuentas x Pagar"
vHeaders.Headers(18) = "Fondos de Ahorro"
vHeaders.Headers(19) = "IVA"

    
Select Case ButtonMenu.Key
   Case "Excel"
      Call sbSIFGridExportar(vGrid, vHeaders, "Convenios_ListadoOrdenes")
   Case "HTML"
      Call sbSIFGridExportar(vGrid, vHeaders, "Convenios_ListadoOrdenes", "HTML")
      
   Case "Resumen"
      Call sbReporteCRD("Rsm")

   Case "Detalle"
      Call sbReporteCRD("Dt")
      
  End Select
  
End Sub

'Oculta el Frame
Private Sub tlbCerrar_ButtonClick(ByVal Button As MSComctlLib.Button)
  fraDetallePago.Visible = False
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False
Call sbInicial
End Sub

Private Sub txtCodigo_Change()
  txtDescripcion.Text = ""
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then

    gBusquedas.Consulta = "Select COD_CONVENIO,DESCRIPCION" _
                        & " from CRD_CONVENIOS"
    gBusquedas.Columna = "COD_CONVENIO"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    frmBusquedas.Show vbModal
    txtCodigo.Text = Trim(gBusquedas.Resultado)
    txtDescripcion.Text = Trim(gBusquedas.Resultado2)
    
    txtDescripcion.SetFocus

End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    txtDescripcion.Text = ""
    Call sbConvenioDatos(txtCodigo.Text)
End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then

    gBusquedas.Consulta = "Select COD_CONVENIO,DESCRIPCION" _
                        & " from CRD_CONVENIOS"
    gBusquedas.Columna = "Descripcion"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    frmBusquedas.Show vbModal
    
    txtCodigo.Text = Trim(gBusquedas.Resultado)
    txtDescripcion.Text = Trim(gBusquedas.Resultado2)

End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  txtCodigo.Text = ""
  Call sbConvenioDatos(txtDescripcion.Text)
End If

End Sub

Private Sub sbConvenioDatos(ByVal vValor As String)
Dim strSQL As String, rs As New ADODB.Recordset
   
On Error GoTo vError
   
   strSQL = " Select COD_CONVENIO,DESCRIPCION" _
          & " from CRD_CONVENIOS"
          
   If txtCodigo.Text <> "" Then strSQL = strSQL & " where COD_CONVENIO = '" & vValor & "' "
   If txtDescripcion.Text <> "" Then strSQL = strSQL & " where DESCRIPCION = '" & vValor & "' "
          
   Call OpenRecordSet(rs, strSQL)
   
   If Not rs.EOF Then
      txtCodigo.Text = rs!COD_CONVENIO
      txtDescripcion.Text = rs!Descripcion
   End If
   
   rs.Close

Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Sub


Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim vCodFactura As String
Dim frm As Form

If Col > 2 Or vPaso Then Exit Sub

vGrid.Row = Row
vGrid.Col = 3
GLOBALES.gTag2 = vGrid.Text

vGrid.Col = 5
GLOBALES.gTag = vGrid.Text

Select Case Col
   Case 1 'Orden
        Call sbSIFForms("frmCR_ConveniosLiquidacion", , , , False)
        
        For Each frm In Forms
        If UCase(frm.Name) = UCase("frmCR_ConveniosLiquidacion") Then
          Call frm.sbConsultaExterna(GLOBALES.gTag, CLng(GLOBALES.gTag2))
          Exit For
        End If
        Next frm
    Case 2 'Pago
        vCodFactura = ""
        
        If fraDetallePago.Visible Then
           fraDetallePago.Visible = False
           
        Else
'           vGrid.Row = vGrid.ActiveRow
           vGrid.Col = 4
           vCodFactura = vGrid.Text
           
           If vGrid.ActiveRow <= 0 Or vCodFactura = "" Then Exit Sub
           
           Call sbCargaDetallePago(vCodFactura)
           fraDetallePago.Visible = True
        End If

End Select

End Sub

Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vTotalPagar As Currency, vTotalOrden As Currency
Dim vTotalDeducciones As Currency

On Error GoTo vError
    
strSQL = " select O.COD_ORDEN, O.COD_CONVENIO, C.DESCRIPCION,C.TIPO_CONVENIO,T.DESCRIPCION as 'TipoConvenio', O.ESTADO, O.REGISTRO_FECHA, O.REGISTRO_USUARIO, O.TOTAL_PAGAR, O.TOTAL_ORDEN" _
       & ",(O.RETENCION_AHORROS + O.COMISIONES_NUEVOS_CREDITOS + O.COMISIONES_RECAUDACION + O.TOTAL_REBAJOS_CXP + O.TOTAL_RESERVA) as 'TotalDeducciones'" _
       & ", O.COMISIONES_NUEVOS_CREDITOS, O.COMISIONES_RECAUDACION, O.TOTAL_RESERVA, O.TOTAL_REBAJOS_CXP, O.RETENCION_AHORROS,O.COD_FACTURA" _
       & ", O.IVA_TOTAL" _
       & " from CRD_CONVENIOS_ORDENES O" _
       & "  inner join CRD_CONVENIOS C on C.COD_CONVENIO = O.COD_CONVENIO" _
       & "  inner join CRD_CONVENIO_TIPO T on C.TIPO_CONVENIO = T.TIPO_CONVENIO"

'Filtro Convenio
If txtCodigo.Text <> "" Then
  strSQL = strSQL & " where O.COD_CONVENIO = '" & txtCodigo.Text & "'"
End If


'Filtro Tipo
i = 0
If chkTipos.Value = Unchecked Then

  If InStr(strSQL, "where") Then
     strSQL = strSQL & " and C.TIPO_CONVENIO in('"
     For i = 1 To lsw.ListItems.Count
       If lsw.ListItems.Item(i).Checked Then
         strSQL = strSQL & "','" & lsw.ListItems.Item(i).Tag
       End If
     Next i
     strSQL = strSQL & "')"
  Else
     strSQL = strSQL & " where C.TIPO_CONVENIO in('"
     For i = 1 To lsw.ListItems.Count
       If lsw.ListItems.Item(i).Checked Then
         strSQL = strSQL & "','" & lsw.ListItems.Item(i).Tag
       End If
     Next i
     strSQL = strSQL & "')"
  End If

End If


'Filtro de fechas
If chkFechas.Value = Unchecked Then

 If InStr(strSQL, "where") Then
   strSQL = strSQL & " and O.REGISTRO_FECHA between ('" & Format(dtpInicio.Value, "yyyy-mm-dd 00:00:00") & "') and ('" & Format(dtpCorte.Value, "yyyy-mm-dd 23:59:59") & "')"
 Else
   strSQL = strSQL & " where O.REGISTRO_FECHA between('" & Format(dtpInicio.Value, "yyyy-mm-dd 00:00:00") & "') and ('" & Format(dtpCorte.Value, "yyyy-mm-dd 23:59:59") & "')"
 End If

End If

'Filtro Usuario
If cboUsuarios.Text <> "TODOS" Then

  If InStr(strSQL, "where") Then
     strSQL = strSQL & " and O.REGISTRO_USUARIO = '" & cboUsuarios.Text & "' "
  Else
    strSQL = strSQL & " where O.REGISTRO_USUARIO = '" & cboUsuarios.Text & "' "
  End If
  
End If


'Filtro Estado
If cboEstado.Text <> "TODOS" Then
  If InStr(strSQL, "where") Then
     strSQL = strSQL & " and O.ESTADO = '" & Mid(cboEstado.Text, 1, 1) & "' "
  Else
    strSQL = strSQL & " where O.ESTADO = '" & Mid(cboEstado.Text, 1, 1) & "' "
  End If
  
End If


strSQL = strSQL & " group by O.COD_ORDEN, O.COD_CONVENIO, C.DESCRIPCION,C.TIPO_CONVENIO,T.DESCRIPCION, O.ESTADO, O.REGISTRO_FECHA, O.REGISTRO_USUARIO, O.TOTAL_PAGAR, O.TOTAL_ORDEN" _
       & ",O.RETENCION_AHORROS,O.COMISIONES_NUEVOS_CREDITOS,O.COMISIONES_RECAUDACION,O.TOTAL_REBAJOS_CXP, O.TOTAL_RESERVA" _
       & ", O.COMISIONES_NUEVOS_CREDITOS, O.COMISIONES_RECAUDACION, O.TOTAL_RESERVA, O.TOTAL_REBAJOS_CXP, O.RETENCION_AHORROS,O.COD_FACTURA" _
       & ", O.IVA_TOTAL"


Call OpenRecordSet(rs, strSQL)


With vGrid

.MaxRows = 1
Do While Not rs.EOF
  .Row = .MaxRows
  
  .Col = 3
  .Text = rs!cod_orden
  
  .Col = 4
  .Text = IIf(IsNull(rs!cod_factura), "", rs!cod_factura)
  
  .Col = 5
  .Text = rs!COD_CONVENIO
  
  .Col = 6
  .Text = rs!Descripcion
  
  .Col = 7
    Select Case rs!estado
      Case "A"
         .Text = "Abierta"
      Case "C"
         .Text = "Cerrada"
      Case "T"
         .Text = "Tramitada"
    End Select
    
  .Col = 8
  .Text = rs!TipoConvenio
  
  .Col = 9
  .Text = Format(rs!registro_Fecha, "dd/mm/yyyy")
  
  .Col = 10
  .Text = rs!registro_usuario
  
  .Col = 11
  .Text = IIf(IsNull(rs!Total_Pagar), 0, Format(rs!Total_Pagar, "standard"))
  vTotalPagar = vTotalPagar + CCur(rs!Total_Pagar)
 
  .Col = 12
  .Text = IIf(IsNull(rs!TOTAL_ORDEN), 0, Format(rs!TOTAL_ORDEN, "Standard"))
  vTotalOrden = vTotalOrden + CCur(rs!TOTAL_ORDEN)
 
  .Col = 13
  .Text = IIf(IsNull(rs!TotalDeducciones), 0, Format(rs!TotalDeducciones, "Standard"))
  vTotalDeducciones = vTotalDeducciones + CCur(rs!TotalDeducciones)
 
  .Col = 14
  .Text = IIf(IsNull(rs!COMISIONES_NUEVOS_CREDITOS), 0, Format(rs!COMISIONES_NUEVOS_CREDITOS, "Standard"))
 
  .Col = 15
  .Text = IIf(IsNull(rs!Comisiones_Recaudacion), 0, Format(rs!Comisiones_Recaudacion, "Standard"))
 
  .Col = 16
  .Text = IIf(IsNull(rs!Total_Reserva), 0, Format(rs!Total_Reserva, "Standard"))
 
  .Col = 17
  .Text = IIf(IsNull(rs!Total_Rebajos_CXP), 0, Format(rs!Total_Rebajos_CXP, "Standard"))
 
  .Col = 18
  .Text = IIf(IsNull(rs!Retencion_Ahorros), 0, Format(rs!Retencion_Ahorros, "Standard"))
  
  .Col = 19
  .Text = IIf(IsNull(rs!IVA_TOTAL), 0, Format(rs!IVA_TOTAL, "Standard"))
  
  .MaxRows = .MaxRows + 1
  rs.MoveNext
Loop

StatusBarX.Panels(1).Text = "Casos ..: " & Format(rs.RecordCount, "###,###,##0")
StatusBarX.Panels(2).Text = "Total Pagar ..: " & Format(vTotalPagar, "###,###,##0")
StatusBarX.Panels(3).Text = "Total Orden ..: " & Format(vTotalOrden, "###,###,##0")
StatusBarX.Panels(4).Text = "Total Deducciones ..: " & Format(vTotalDeducciones, "###,###,##0")

rs.Close
.MaxRows = .MaxRows - 1

End With

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbCargaDetallePago(ByVal vCodFactura As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vNSolicitud As Long, itmX As ListViewItem

On Error GoTo vError

  Call sbLimpiaDetallePago

  strSQL = "exec spConvenios_Ordenes_DetallePago '" & vCodFactura & "' "
  Call OpenRecordSet(rs, strSQL)

  If Not rs.EOF Then
   fraDetallePago.Caption = "[Orden: " & rs!cod_orden & "]  " & rs!COD_CONVENIO & "-" & rs!Descripcion
   lblTotalBruto.Caption = IIf(IsNull(rs!OrdenTotalNeto), 0, Format(rs!OrdenTotalNeto, "Standard"))
   lblTotalDeducciones.Caption = IIf(IsNull(rs!TotalDeducciones), 0, Format(rs!TotalDeducciones, "Standard"))
   lblOrdenTotalNeto.Caption = IIf(IsNull(rs!Total_Pagar), 0, Format(rs!Total_Pagar, "Standard"))
   lblFechaCorte.Caption = Format(rs!FECHA_CORTE, "dd/mm/yyyy")
   lblUsuario.Caption = rs!registro_usuario
   lblOrdenDocumento.Caption = rs!OrdenDocumento
   lblFactura.Caption = rs!cod_factura
   lblFacturaMonto.Caption = IIf(IsNull(rs!FacturaTotal), 0, Format(rs!FacturaTotal, "Standard"))
   lblCargosAplicados.Caption = IIf(IsNull(rs!Cargos), 0, Format(rs!Cargos, "Standard"))
   lblFactTotalNeto.Caption = 0
   lblMontoDesembolso.Caption = IIf(IsNull(rs!MontoDesembolso), 0, Format(rs!MontoDesembolso, "Standard"))
   lblNSolicitud.Caption = IIf(IsNull(rs!NSolicitud), 0, rs!NSolicitud)
   vNSolicitud = IIf(IsNull(rs!NSolicitud), 0, rs!NSolicitud)
   lblFechaSolicitud.Caption = Format(rs!fecha_solicitud & "", "dd/mm/yyyy")
   lblBeneficiario.Caption = rs!beneficiario & ""
   lblTipoDocumento.Caption = rs!Tipo & ""
   lblTesoreriaDocumento.Caption = IIf(IsNull(rs!nDocumento), "", rs!nDocumento)
   lblBanco.Caption = rs!Banco & ""
   lblFechaEmision.Caption = IIf(IsNull(rs!fecha_emision), "", Format(rs!fecha_emision, "dd/mm/yyyy"))
   
   Select Case rs!estado
     Case "S"
        lblEstado.Caption = "Solicitado"
     Case "I", "T", "E"
        lblEstado.Caption = "Impreso"
     Case "A"
        lblEstado.Caption = "Anulado"
     Case "P"
        lblEstado.Caption = "Pendiente"
   End Select
  End If
  rs.Close
  
  If vNSolicitud <= 0 Then Exit Sub
  strSQL = "Exec spConvenios_OrdenesCanceladas " & vNSolicitud & ""
  Call OpenRecordSet(rs, strSQL)

  Do While Not rs.EOF
     Set itmX = lswOrdenesCanceladas.ListItems.Add(, , rs!cod_orden)
      itmX.SubItems(1) = IIf(IsNull(rs!fecha_emision), "", Format(rs!fecha_emision, "dd/mm/yyyy"))
      itmX.SubItems(2) = Format(rs!Total_Pagar, "Standard")
      rs.MoveNext
  Loop
  rs.Close
    
Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbLimpiaDetallePago()
   lblTotalBruto.Caption = 0
   lblTotalDeducciones.Caption = 0
   lblOrdenTotalNeto.Caption = 0
   lblFechaCorte.Caption = ""
   lblUsuario.Caption = ""
   lblOrdenDocumento.Caption = ""
   lblFactura.Caption = ""
   lblFacturaMonto.Caption = 0
   lblCargosAplicados.Caption = ""
   lblFactTotalNeto.Caption = 0
   lblMontoDesembolso.Caption = 0
   lblNSolicitud.Caption = ""
   lblFechaSolicitud.Caption = ""
   lblBeneficiario.Caption = ""
   lblTipoDocumento.Caption = ""
   lblTesoreriaDocumento.Caption = ""
   lblBanco.Caption = ""
   lblFechaEmision.Caption = ""
   lblEstado.Caption = ""
   
   lswOrdenesCanceladas.ListItems.Clear
End Sub


Private Sub sbReporteCRD(ByVal vTipoReporte As String)
Dim vTitulo As String, vSubTitulo As String, vFiltro As String, vTemp As String, i As Integer
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vTitulo = ""
vSubTitulo = ""
strSQL = ""
Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .Connect = glogon.ConectRPT
 .WindowTitle = "Reporte Listados de Convenios"
 vTitulo = "Listado General de Convenios"
 vSubTitulo = "Listado de ordenes "
 vNombreReporte = "Convenios_Ordenes"
 
 If txtCodigo.Text <> "" Then
    vSubTitulo = vSubTitulo & " del convenio: " & "" & txtCodigo.Text & ""
    strSQL = strSQL & "{CRD_CONVENIOS_ORDENES.COD_CONVENIO} = '" & txtCodigo.Text & "'"
 Else
    vSubTitulo = vSubTitulo & " por convenio"
 End If
 
 
 'Filtro Tipo
i = 0
If chkTipos.Value = Unchecked Then

  If Len(strSQL) > 0 Then
     strSQL = strSQL & " and {CRD_CONVENIOS.TIPO_CONVENIO} in ['"
     For i = 1 To lsw.ListItems.Count
       If lsw.ListItems.Item(i).Checked Then
         strSQL = strSQL & "','" & lsw.ListItems.Item(i).Tag
       End If
     Next i
     strSQL = strSQL & "']"
  Else
     strSQL = strSQL & " {CRD_CONVENIOS.TIPO_CONVENIO} in ['"
     For i = 1 To lsw.ListItems.Count
       If lsw.ListItems.Item(i).Checked Then
         strSQL = strSQL & "','" & lsw.ListItems.Item(i).Tag
       End If
     Next i
     strSQL = strSQL & "']"
  End If

End If


'Filtro de fechas
If chkFechas.Value = Unchecked Then
  
  If Len(strSQL) > 0 Then
   vSubTitulo = vSubTitulo & " / Registradas entre  " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
   strSQL = strSQL & " and {CRD_CONVENIOS_ORDENES.REGISTRO_FECHA}" & " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ") to date(" _
           & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
  Else
   vSubTitulo = vSubTitulo & " Registradas entre  " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
   strSQL = strSQL & " {CRD_CONVENIOS_ORDENES.REGISTRO_FECHA}" & " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ") to date(" _
           & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
  End If
  
End If


'Filtro Usuario

If cboUsuarios.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & "and"
  vSubTitulo = vSubTitulo & " / Usuario registra: " & "" & cboUsuarios.Text & ""
  strSQL = strSQL & " {CRD_CONVENIOS_ORDENES.REGISTRO_USUARIO} = '" & cboUsuarios.Text & "' "
End If


'Filtro Estado
If cboEstado.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & "and"
   vSubTitulo = vSubTitulo & " / Estado de la Orden: " & "" & cboEstado.Text & ""
   strSQL = strSQL & " {CRD_CONVENIOS_ORDENES.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "' "
End If
 
 .ReportFileName = SIFGlobal.fxPathReportes(Trim(vNombreReporte) & vTipoReporte & ".rpt")

  If vTipoReporte = "Rsm" Then
    vTitulo = vTitulo & " Resumen"
  Else
    vTitulo = vTitulo & " Detalle"
  End If
  
 .Formulas(0) = "fxTitulo= '" & vTitulo & "'"
 .Formulas(1) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(2) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(4) = "fxUsuario='Usuario: " & glogon.Usuario & "'"

 .SelectionFormula = strSQL

 .PrintReport

End With

Me.MousePointer = vbDefault

Call Bitacora("Imprime", vNombreReporte & vTipoReporte)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

