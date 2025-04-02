VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmFNDConsultaDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle del Contrato"
   ClientHeight    =   7608
   ClientLeft      =   2208
   ClientTop       =   1488
   ClientWidth     =   11232
   Icon            =   "frmFNDConsultaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7608
   ScaleWidth      =   11232
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3372
      Left            =   120
      TabIndex        =   40
      Top             =   3360
      Width           =   10932
      _Version        =   1245187
      _ExtentX        =   19283
      _ExtentY        =   5948
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
      Item(0).Caption =   "Movimientos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lsw"
      Item(1).Caption =   "Beneficiarios"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lswBeneficiarios"
      Item(2).Caption =   "SINPE/Tránsito"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "lswTransito"
      Begin XtremeSuiteControls.ListView lswBeneficiarios 
         Height          =   2892
         Left            =   -69880
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   10692
         _Version        =   1245187
         _ExtentX        =   18860
         _ExtentY        =   5101
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
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswTransito 
         Height          =   2892
         Left            =   -69880
         TabIndex        =   42
         Top             =   360
         Visible         =   0   'False
         Width           =   10692
         _Version        =   1245187
         _ExtentX        =   18860
         _ExtentY        =   5101
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
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   2892
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   10692
         _Version        =   1245187
         _ExtentX        =   18860
         _ExtentY        =   5101
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
         ShowBorder      =   0   'False
      End
   End
   Begin XtremeSuiteControls.PushButton cmdEstadoCuenta 
      Height          =   612
      Left            =   9480
      TabIndex        =   39
      Top             =   6840
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Estado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmFNDConsultaDetalle.frx":030A
   End
   Begin VB.TextBox txtCuentaCliente 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   120
      Width           =   3012
   End
   Begin VB.Timer timerX 
      Interval        =   50
      Left            =   6960
      Top             =   0
   End
   Begin VB.TextBox txtBeneficiario 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   840
      Width           =   6252
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   480
      Width           =   6252
   End
   Begin VB.TextBox txtOperadora 
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
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   6252
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtContrato 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
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
      Height          =   315
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2052
   End
   Begin VB.TextBox txtDescripcion 
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
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1800
      Width           =   6252
   End
   Begin VB.TextBox txtEstado 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtSubCuenta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
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
      Height          =   315
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   840
      Width           =   2052
   End
   Begin VB.TextBox txtCedula 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
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
      Height          =   315
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   480
      Width           =   2052
   End
   Begin VB.Label lblDisponible 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   9600
      TabIndex        =   38
      Top             =   2880
      Width           =   1452
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Disponible:"
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
      Left            =   8400
      TabIndex        =   37
      Top             =   2880
      Width           =   852
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total.:"
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
      Index           =   9
      Left            =   5400
      TabIndex        =   36
      Top             =   2880
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rendimiento.:"
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
      Left            =   5400
      TabIndex        =   35
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Aportes:"
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
      Index           =   7
      Left            =   5400
      TabIndex        =   34
      Top             =   2160
      Width           =   1092
   End
   Begin VB.Label lblTasaPtsAdd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   4440
      TabIndex        =   33
      Top             =   2880
      Width           =   732
   End
   Begin VB.Label lblTasa 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   4440
      TabIndex        =   32
      Top             =   2520
      Width           =   732
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Renovación:"
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
      Index           =   6
      Left            =   360
      TabIndex        =   31
      Top             =   2880
      Width           =   1452
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pts Add:"
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
      Index           =   5
      Left            =   3360
      TabIndex        =   30
      Top             =   2880
      Width           =   1092
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tasa Ref:"
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
      Left            =   3360
      TabIndex        =   29
      Top             =   2520
      Width           =   852
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Plazo:"
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
      Left            =   360
      TabIndex        =   28
      Top             =   2520
      Width           =   1452
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mensualidad:"
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
      Left            =   360
      TabIndex        =   27
      Top             =   2160
      Width           =   1452
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "En Tránsito:"
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
      Left            =   8400
      TabIndex        =   26
      Top             =   2520
      Width           =   972
   End
   Begin VB.Label lblMontoTransito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   9600
      TabIndex        =   25
      Top             =   2520
      Width           =   1452
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IBAN:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   5280
      TabIndex        =   24
      Top             =   120
      Width           =   2532
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   6600
      TabIndex        =   22
      Top             =   2880
      Width           =   1452
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1320
      TabIndex        =   21
      Top             =   840
      Width           =   1692
   End
   Begin VB.Label lblOperacionASE 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operación asociada:"
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
      Left            =   6000
      TabIndex        =   18
      Top             =   6960
      Width           =   3252
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1320
      TabIndex        =   17
      Top             =   480
      Width           =   1332
   End
   Begin VB.Label lblMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   16
      Top             =   2160
      Width           =   1332
   End
   Begin VB.Label lblPlazo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2280
      TabIndex        =   15
      Top             =   2520
      Width           =   852
   End
   Begin VB.Label lblRenueva 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2280
      TabIndex        =   14
      Top             =   2880
      Width           =   852
   End
   Begin VB.Label lblAportes 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   6600
      TabIndex        =   13
      Top             =   2160
      Width           =   1452
   End
   Begin VB.Label lblRendimiento 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   6600
      TabIndex        =   12
      Top             =   2520
      Width           =   1452
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   312
      Left            =   8520
      TabIndex        =   9
      Top             =   1800
      Width           =   852
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   0
      Left            =   8520
      TabIndex        =   8
      Top             =   1440
      Width           =   852
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   1452
   End
   Begin VB.Label lblX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   1692
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11772
   End
End
Attribute VB_Name = "frmFNDConsultaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbDetalleContrato()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError

lsw.ListItems.Clear

strSQL = "select C.*, C.Aportes + C.Rendimiento - isnull(C.Monto_Transito,0) as 'Disponible'" _
       & " ,S.nombre,O.descripcion as Operadora,P.descripcion as PlanX" _
       & " from fnd_contratos C inner join Socios S on C.cedula = S.cedula" _
       & " inner join fnd_planes P on C.cod_plan = P.cod_plan and C.cod_operadora = P.cod_operadora" _
       & " inner join fnd_operadoras O on C.cod_operadora = O.cod_operadora" _
       & " where C.cod_operadora = " & gFondos.Operadora _
       & " and C.cod_plan = '" & gFondos.Plan & "' and C.cod_contrato = " & gFondos.Contrato
Call OpenRecordSet(rs, strSQL)

 txtCedula.Text = rs!Cedula
 txtNombre.Text = rs!Nombre
 txtOperadora.Text = rs!Operadora
 txtOperadora.Tag = rs!cod_Operadora
 txtDescripcion.Text = rs!PlanX
 txtDescripcion.Tag = rs!cod_Plan
 txtContrato = rs!COD_CONTRATO
 txtEstado.Text = fxFndEstadoContrato(rs!Estado & "")
 txtFecha.Text = Format(rs!Fecha_Inicio, "dd/mm/yyyy")
 

 txtCuentaCliente.Text = rs!CUENTA_CLIENTE & ""
 lblMontoTransito.Caption = Format(rs!Monto_Transito & "", "Standard")
 lblDisponible.Caption = Format(rs!Disponible & "", "Standard")

 lblMonto.Caption = Format(rs!Monto, "Standard")
 lblPlazo.Caption = rs!Plazo
 lblRenueva.Caption = IIf(rs!Renueva = "S", "SI", "NO")
 If rs!Tasa_Tipo = "V" Then
     lblTasa.Caption = "TBP/TCP"
 Else
     lblTasa.Caption = rs!Tasa_Referencia
 End If
 lblTasaPtsAdd.Caption = IIf(IsNull(rs!Tasa_PtsAdd), 0, rs!Tasa_PtsAdd)
 
' lblIncAnual.Caption = Format(rs!Inc_Anual, "Standard")
' lblIncTipo.Caption = IIf(rs!Inc_Tipo = "P", "Porcentaje", "Monto")
 lblAportes.Caption = Format(rs!aportes, "Standard")
 lblRendimiento.Caption = Format(rs!rendimiento, "Standard")
 
 lblOperacionASE = "Operación : " & IIf(IsNull(rs!Operacion), "", rs!Operacion)
rs.Close


If gFondos.SubCuenta = 0 Then

  strSQL = "Select Det.*,isnull(Doc.Descripcion,'') as 'DocDesc', isnull(Con.Descripcion,'') as 'ConceptoDesc'" _
         & " from fnd_contratos_detalle Det left join SIF_DOCUMENTOS Doc on Det.Tcon = Doc.Tipo_Documento" _
         & " left join SIF_Conceptos Con on Det.Cod_Concepto = Con.Cod_Concepto" _
         & " where Det.cod_operadora = " & gFondos.Operadora _
         & " And Det.cod_plan = '" & gFondos.Plan & "' and Det.Cod_Contrato = " & gFondos.Contrato _
         & " order by Det.cod_fnd_detalle desc"
  
  txtSubCuenta.BackColor = lblX.BackColor
  txtBeneficiario.BackColor = lblX.BackColor

Else
  
  txtCedula.BackColor = lblX.BackColor
  txtNombre.BackColor = lblX.BackColor
 
  
  strSQL = "select * from fnd_subCuentas  where cod_operadora = " & gFondos.Operadora _
         & " And cod_plan='" & gFondos.Plan & "' and Cod_Contrato = " & gFondos.Contrato _
         & " and IdX = " & gFondos.SubCuenta
  Call OpenRecordSet(rs, strSQL)
    txtSubCuenta = rs!Cedula
    txtBeneficiario = Trim(rs!Nombre)
    'Reemplazo los valores del contrato por el de la subCuenta
    lblMonto = Format(rs!Cuota, "Standard")
    lblAportes = Format(rs!aportes, "Standard")
    lblRendimiento = Format(rs!rendimiento, "Standard")
  rs.Close
  
  strSQL = "Select Det.*,isnull(Doc.Descripcion,'') as 'DocDesc', '' as 'ConceptoDesc'" _
         & ",'' as 'Usuario'" _
         & " from fnd_SubCuentas_detalle Det left join SIF_DOCUMENTOS Doc on Det.Tcon = Doc.Tipo_Documento" _
         & " where Det.cod_operadora = " & gFondos.Operadora _
         & " And Det.cod_plan = '" & gFondos.Plan & "' and Det.Cod_Contrato = " & gFondos.Contrato _
         & " and Det.IDx = " & gFondos.SubCuenta _
         & " order by Det.cod_fnd_detalle desc"

End If



Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!cod_fnd_Detalle)
       itmX.SubItems(1) = Format(rs!Monto, "Standard")
       itmX.SubItems(2) = Format(rs!Fecha_Proceso, "####-##")
       itmX.SubItems(3) = rs!fecha
       itmX.SubItems(4) = rs!DocDesc
       itmX.SubItems(5) = rs!nCon
       itmX.SubItems(6) = Format(IIf(IsNull(rs!Fecha_Acredita), rs!fecha, rs!Fecha_Acredita), "dd/mm/yyyy")
       itmX.SubItems(7) = rs!Usuario & ""
       itmX.SubItems(8) = rs!ConceptoDesc
   
       If rs!Monto < 0 Then
              itmX.ForeColor = vbRed
              itmX.Bold = True
              itmX.TextBackColor = RGB(250, 219, 216)
       End If
   rs.MoveNext
Loop
rs.Close

lblTotal.Caption = Format(CCur(lblAportes.Caption) + CCur(lblRendimiento.Caption), "Standard")

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cmdEstadoCuenta_Click()
Dim strSQL As String

strSQL = "{FND_CONTRATOS.COD_OPERADORA} =" & gFondos.Operadora & "And " _
        & "{FND_CONTRATOS.COD_PLAN} ='" & gFondos.Plan & "' and " _
        & "{FND_CONTRATOS.COD_CONTRATO} = " & gFondos.Contrato

   
With frmContenedor.Crt
  .Reset
  .WindowShowPrintSetupBtn = True
  .WindowShowExportBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowShowZoomCtl = True
  .WindowTitle = "Fondos de Ahorros e Inversiones"
  .WindowState = crptMaximized

  .Connect = glogon.ConectRPT

  .ReportFileName = SIFGlobal.fxPathReportes("Fondos_EstadoDetallado.rpt")
  .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
  .Formulas(1) = "Usuario='" & Trim(glogon.Usuario) & "'"
  .Formulas(2) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
  .Formulas(3) = "SubTitulo='" & Format(fxFechaServidor, "yyyy/mm/dd") & "'"
  .SelectionFormula = strSQL
  .PrintReport
End With

End Sub





Private Sub Form_Load()

tcMain.Item(0).Selected = True

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lsw.ColumnHeaders
   .Clear
   .Add 1, , "ID", 1100
   .Add 2, , "Monto", 1500, vbRightJustify
   .Add 3, , "Proceso", 1100, vbCenter
   .Add 4, , "Fecha", 2000
   .Add 5, , "Tipo", 2600
   .Add 6, , "No. Documento", 2500
   .Add 7, , "Acredita", 1200, vbCenter
   .Add 8, , "Usuario", 1800
   .Add 9, , "Concepto", 2500
End With

With lswBeneficiarios.ColumnHeaders
    .Clear
    .Add 1, , "Identificación", 1600
    .Add 2, , "Nombre", 3300
    .Add 3, , "Porcentaje", 1200
    .Add 4, , "Parentesco", 1800
    .Add 5, , "Fecha Nac.", 1800
End With


With lswTransito.ColumnHeaders
    .Clear
    .Add 1, , "Transac.", 1400
    .Add 2, , "Fecha", 1800
    .Add 3, , "Monto", 1800, vbRightJustify
    .Add 4, , "Divisa", 900, vbCenter
    .Add 5, , "Referencia", 1800
    .Add 6, , "Transac.", 900
    .Add 7, , "Transac. Desc", 2500
    .Add 8, , "Id. Origen", 1500
    .Add 9, , "CC. Origen", 1800
    .Add 10, , "Banco Origen", 2500
    
End With
End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Select Case Item.Index

Case 0 'Movimientos del Contrato
  Call sbDetalleContrato

Case 1 'Detalle de Beneficiarios

    strSQL = "Select CedulaBn,Nombre,Porcentaje,parentesco,fechanac From FND_CONTRATOS_BENEFICIARIOS where " _
           & " Cedula = '" & Trim(txtCedula) & "' and cod_contrato = " & txtContrato & "" _
           & " and cod_operadora = " & txtOperadora.Tag _
           & " and cod_plan='" & Trim(txtDescripcion.Tag) & "'"
           
           
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswBeneficiarios.ListItems.Add(, , rs!cedulaBn)
           itmX.SubItems(1) = Trim(rs!Nombre)
           itmX.SubItems(2) = Trim(rs!Porcentaje) & "%"
           itmX.SubItems(3) = fxParentesco(rs!parentesco)
           itmX.SubItems(4) = Format(rs!FechaNac, "dd/mm/yyyy")
       rs.MoveNext
    Loop
    rs.Close
    
 Case 2 'Movimientos en Tránsito

    lswTransito.ListItems.Clear

    strSQL = "exec spFndSinpeMovTransito '" & txtCuentaCliente.Text & "'"

    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswTransito.ListItems.Add(, , rs!COD_TRANSITO)
           itmX.SubItems(1) = rs!Registro_Fecha & ""
           itmX.SubItems(2) = Format(rs!Monto, "Standard")
           itmX.SubItems(4) = rs!Cod_Moneda & ""
           itmX.SubItems(5) = rs!COD_REFERENCIA & ""
           itmX.SubItems(6) = rs!TRANSAC_TIPO & ""
           itmX.SubItems(7) = rs!TRANSAC_DESC & ""
           itmX.SubItems(8) = rs!CEDULA_ORIGEN & ""
           itmX.SubItems(9) = rs!CUENTA_CLIENTE_ORIGEN & ""
           itmX.SubItems(10) = rs!BANCO_ORIGEN_DESC & ""
          
       rs.MoveNext
    Loop
    rs.Close

End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
Call sbDetalleContrato
End Sub

