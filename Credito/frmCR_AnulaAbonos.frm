VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCR_AnulaAbonos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de Abonos"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   HelpContextID   =   3002
   Icon            =   "frmCR_AnulaAbonos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   10560
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2052
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   10332
      _Version        =   1572864
      _ExtentX        =   18224
      _ExtentY        =   3619
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Timer TimerVerificaPlanPagos 
      Left            =   3720
      Top             =   240
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   372
      Left            =   1920
      TabIndex        =   8
      Top             =   240
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3196
      _ExtentY        =   656
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3720
      TabIndex        =   9
      Top             =   840
      Width           =   5412
      _Version        =   1572864
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3720
      TabIndex        =   10
      Top             =   1200
      Width           =   5412
      _Version        =   1572864
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1920
      TabIndex        =   11
      Top             =   1200
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3196
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
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1920
      TabIndex        =   12
      Top             =   840
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3196
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
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   312
      Left            =   9120
      TabIndex        =   13
      Top             =   1200
      Width           =   1092
      _Version        =   1572864
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOpex 
      Height          =   312
      Left            =   9120
      TabIndex        =   14
      Top             =   840
      Width           =   1092
      _Version        =   1572864
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   216
      Left            =   240
      TabIndex        =   19
      Top             =   1716
      Width           =   216
      _Version        =   1572864
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
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
   End
   Begin XtremeSuiteControls.ComboBox cboVisualiza 
      Height          =   312
      Left            =   6480
      TabIndex        =   25
      Top             =   240
      Width           =   3732
      _Version        =   1572864
      _ExtentX        =   6588
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
   Begin XtremeSuiteControls.FlatEdit txtIntCor 
      Height          =   312
      Left            =   1800
      TabIndex        =   26
      Top             =   4560
      Width           =   1572
      _Version        =   1572864
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtIntMor 
      Height          =   312
      Left            =   1800
      TabIndex        =   27
      Top             =   4920
      Width           =   1572
      _Version        =   1572864
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAmortizacion 
      Height          =   312
      Left            =   1800
      TabIndex        =   28
      Top             =   5640
      Width           =   1572
      _Version        =   1572864
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtCargos 
      Height          =   312
      Left            =   1800
      TabIndex        =   29
      Top             =   5280
      Width           =   1572
      _Version        =   1572864
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtTotal 
      Height          =   312
      Left            =   1800
      TabIndex        =   30
      Top             =   6120
      Width           =   1572
      _Version        =   1572864
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAbIntCor 
      Height          =   312
      Left            =   3360
      TabIndex        =   31
      Top             =   4560
      Width           =   1572
      _Version        =   1572864
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
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAbIntMor 
      Height          =   312
      Left            =   3360
      TabIndex        =   32
      Top             =   4920
      Width           =   1572
      _Version        =   1572864
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
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAbAmortizacion 
      Height          =   312
      Left            =   3360
      TabIndex        =   33
      Top             =   5640
      Width           =   1572
      _Version        =   1572864
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
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAbCargos 
      Height          =   312
      Left            =   3360
      TabIndex        =   34
      Top             =   5280
      Width           =   1572
      _Version        =   1572864
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
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAbTotal 
      Height          =   312
      Left            =   3360
      TabIndex        =   35
      Top             =   6120
      Width           =   1572
      _Version        =   1572864
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtFecha 
      Height          =   312
      Left            =   4920
      TabIndex        =   36
      Top             =   4560
      Width           =   1212
      _Version        =   1572864
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdAnular 
      Height          =   732
      Left            =   8760
      TabIndex        =   37
      Top             =   5760
      Width           =   1452
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "&Anular"
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
      Picture         =   "frmCR_AnulaAbonos.frx":000C
   End
   Begin XtremeSuiteControls.CheckBox chkGeneraMorosidad 
      Height          =   216
      Left            =   6120
      TabIndex        =   15
      Top             =   5160
      Width           =   3456
      _Version        =   1572864
      _ExtentX        =   6096
      _ExtentY        =   381
      _StockProps     =   79
      Caption         =   "Generar Registro de Morosidad"
      ForeColor       =   255
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
   Begin XtremeSuiteControls.CheckBox chkTrasladar 
      Height          =   216
      Left            =   6120
      TabIndex        =   16
      Top             =   5400
      Width           =   3456
      _Version        =   1572864
      _ExtentX        =   6096
      _ExtentY        =   381
      _StockProps     =   79
      Caption         =   "Trasladar Saldo al Plazo Restante"
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
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.PushButton cmdMarcas 
      Height          =   312
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Calcular Marcas!"
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   3
      Left            =   4440
      TabIndex        =   24
      Top             =   240
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3196
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Tipo Movimiento"
      ForeColor       =   16777215
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
      Index           =   0
      Left            =   840
      TabIndex        =   23
      Top             =   240
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Operación"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   1
      Left            =   840
      TabIndex        =   22
      Top             =   840
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Identificación"
      ForeColor       =   16777215
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
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   2
      Left            =   840
      TabIndex        =   21
      Top             =   1200
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Línea"
      ForeColor       =   16777215
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
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCapMov 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Movimientos Registrados"
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
      Height          =   312
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   10332
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Totales...:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Proceso"
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
      Height          =   315
      Index           =   7
      Left            =   4920
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargos Registrados"
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
      Index           =   6
      Left            =   240
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Amortización"
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
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Int.Morosidad"
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
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Int.Corriente"
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
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos Anulación"
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
      Height          =   315
      Index           =   5
      Left            =   3360
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos Originales"
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
      Height          =   315
      Index           =   1
      Left            =   1800
      TabIndex        =   0
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Image imgBanner 
      Height          =   1572
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12012
   End
End
Attribute VB_Name = "frmCR_AnulaAbonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vLlave As Long, vRetencion As Boolean, vPaso As Boolean

Private Sub cboVisualiza_Click()
If vPaso Then Exit Sub

txtFecha.Text = ""

txtIntCor.Text = "0.00"
txtIntMor.Text = "0.00"
txtAmortizacion.Text = "0.00"
txtCargos.Text = "0.00"
txtTotal.Text = "0.00"

txtABCargos.Text = 0
txtABTotal.Text = 0

txtABIntCor.Text = 0
txtABIntMor.Text = 0
txtABAmortizacion.Text = 0

chkGeneraMorosidad.Value = vbUnchecked

Call sbCargoMovimientos

If cboVisualiza.Text = "Abonos Ordinarios" Then
  txtABIntMor.Enabled = False
  txtABCargos.Enabled = False
  
  chkGeneraMorosidad.Enabled = True
  chkGeneraMorosidad.Value = vbUnchecked
    
Else
  txtABIntMor.Enabled = True
  txtABCargos.Enabled = True
  chkGeneraMorosidad.Enabled = False
  chkGeneraMorosidad.Value = vbChecked

End If

End Sub

Private Sub chkGeneraMorosidad_Click()
If chkGeneraMorosidad.Value = vbChecked Then
  chkTrasladar.Enabled = False
  chkTrasladar.Value = vbUnchecked
Else
  chkTrasladar.Enabled = True
End If

End Sub

Private Sub chkTodos_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub

Private Sub cmdAnular_Click()
Dim strSQL As String, rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
Dim vIntM As Currency, vIntC As Currency, vAMORTIZAm As Currency
Dim vANUIntM As Currency, vANUIntC As Currency, vANUAMORTIZAm As Currency
Dim vCuenta As String, vFecha As Date, lngRecibo As Long
Dim lngOperacion As Long, vTipo As String, vFechaProceso As Long
Dim vTipoDoc As String, vConcepto As String

'Verificar Congelamiento
If fxgCongelamiento(txtCedula, "per_abono_cajas") Then
  MsgBox "Esta Persona se encuentra CONGELADA, verifique...", vbExclamation
  Exit Sub
End If


If Not fxValidaInformacion Then
  MsgBox "La información suministrada no es válida...", vbCritical
  Exit Sub
End If 'Validacion
    
lngOperacion = txtOperacion
   
vTipo = "ND"
vConcepto = "CRD008"
lngRecibo = 0
vFecha = fxFechaServidor

If GLOBALES.SysDocVersion = 1 Then
 vTipoDoc = "8"
Else
 'Ctrl Doc v2
 vTipoDoc = "ND"
End If

'Configuracion del Documento
vTipo = "ND"
vCuenta = Trim(fxDocumentoCuenta(vTipo))
lngRecibo = fxDocumentoConsecutivo(vTipo)

If vAseDocValido = False Then
  MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
        & " válida para esta operación...", vbCritical
  Exit Sub
End If


'Establece la Fecha de Proceso
If txtFecha.Text = "" Then
   vFechaProceso = GLOBALES.glngFechaCR
Else
   vFechaProceso = txtFecha.Text
End If

Me.MousePointer = vbHourglass
  
  
If Not vRetencion Then
  strSQL = "update reg_creditos set estado = case when " & CCur(txtABAmortizacion.Text) & " > 0 then 'A' else Estado end" _
       & ",interesc = interesc - " & CCur(txtABIntCor) _
       & ", amortiza = amortiza - " & CCur(txtABAmortizacion) _
       & ", saldo = saldo + " & CCur(txtABAmortizacion) _
       & ", saldo_mes = saldo_mes + " & CCur(txtABAmortizacion) _
       & ", cuotas_anuladas = cuotas_anuladas + 1" _
       & " where id_solicitud = " & lngOperacion
  Call ConectionExecute(strSQL)


    If chkTrasladar.Enabled And chkTrasladar.Value = vbChecked Then
       strSQL = "update REG_CREDITOS  set CUOTA =  dbo.fxCrdCalculoCuota(Saldo_Mes,(Plazo - DATEDIFF(mm, dbo.fxSIFCorteAFecha(PriDeduc) , dbo.MyGetdate())), interesv )" _
              & ",FECULT = " & GLOBALES.glngFechaCR _
              & " where ID_SOLICITUD = " & txtOperacion.Text
       Call ConectionExecute(strSQL)
    End If

Else 'Retenciones
  strSQL = "update reg_creditos set estado = 'A',interesc = interesc - " & CCur(txtABIntCor) _
       & ", amortiza = amortiza - " & CCur(txtABAmortizacion) _
       & ", cuotas_anuladas = cuotas_anuladas + 1" _
       & " where id_solicitud = " & lngOperacion
  Call ConectionExecute(strSQL)
End If
  

 
If cboVisualiza.Text = "Abonos Ordinarios" Then
    strSQL = "insert creditos_dt(codigo,id_solicitud,cuota,abono,intcp,amortiza,fechas,fechap,estado" _
        & ",tcon,ncon,usuario,cod_concepto,cod_caja) values('" & txtCodigo & "'," & lngOperacion & ",0," _
        & CCur(txtABIntCor) + CCur(txtABAmortizacion) & "," & CCur(txtABIntCor) & "," _
        & CCur(txtABAmortizacion) & ",dbo.MyGetdate()" _
        & "," & IIf((Len(Trim(txtFecha)) = 0), GLOBALES.glngFechaCR, txtFecha) _
        & ",'N','" & vTipoDoc & "','" & lngRecibo & "','" & glogon.Usuario & "','" & vConcepto & "','')"
    Call ConectionExecute(strSQL)
   
   
   
    'Genera Registro de Morosidad
    If chkGeneraMorosidad.Value = vbChecked Then
      strSQL = "insert morosidad(id_solicitud,codigo,fechap,fecap,fecult,cuota_morosa,intc,intm,amortiza,cargo,estado,estadoi,tcon,ncon,usuario,cod_concepto,cod_caja)" _
             & " values(" & lngOperacion & ",'" & UCase(txtCodigo) & "'," & vFechaProceso & "," & vFechaProceso _
             & ",dbo.MyGetdate()," & (CCur(txtABAmortizacion.Text) + CCur(txtABIntCor.Text) + CCur(txtABIntMor.Text) + CCur(txtABCargos.Text)) _
             & "," & CCur(txtABIntCor) & "," & CCur(txtABIntMor.Text) & "," & CCur(txtABAmortizacion) & "," & CCur(txtABCargos) _
             & ",'A','A','" & vTipoDoc & "','" & lngRecibo & "','" & glogon.Usuario & "','" & vConcepto & "','')"
      Call ConectionExecute(strSQL)
    End If
   
   Call Bitacora("Anula", "OP: " & txtOperacion & " INTC: " & CCur(txtABIntCor) & " AMORT: " & CCur(txtABAmortizacion))
  
Else 'Morosidad
  
    If txtFecha = "" Then
      'Esquema Personalizado de Anulacion de Morosidad
         
         Call Bitacora("Anula", "OP: " & txtOperacion & " INTC: " & txtABIntCor & " INTM: " & txtABIntMor & " AMORT: " & txtABAmortizacion)
        
         strSQL = "insert morosidad(id_solicitud,codigo,fechap,fecap,fecult,cuota_morosa,intc,intm,amortiza,cargo,estado,estadoi,tcon,ncon" _
                & ",abintc,abintm,abamortiza,abCargo,usuario,cod_concepto,cod_caja) values(" & txtOperacion & ",'" & txtCodigo & "'," & vFechaProceso & "," & vFechaProceso _
                & ",dbo.MyGetdate()," & (CCur(txtABIntCor) + CCur(txtABIntMor) + CCur(txtABAmortizacion) + CCur(txtABCargos.Text)) _
                & "," & CCur(txtABIntCor) & "," & CCur(txtABIntMor) & "," & CCur(txtABAmortizacion) & "," & CCur(txtABCargos.Text) _
                & ",'N','C','" & vTipoDoc & "','" & lngRecibo & "'," & CCur(txtABIntCor) & "," & CCur(txtABIntMor) & "," & CCur(txtABAmortizacion) _
                & "," & CCur(txtABCargos.Text) & ",'" & glogon.Usuario & "','" & vConcepto & "','')"
         Call ConectionExecute(strSQL)
          
        'Genera Registro de Morosidad
        If chkGeneraMorosidad.Value = vbChecked Then
          strSQL = "insert morosidad(id_solicitud,codigo,fechap,fecap,fecult,cuota_morosa,intc,intm,amortiza,cargo,estado,estadoi,tcon,ncon,usuario,cod_concepto,cod_caja)" _
                 & " values(" & lngOperacion & ",'" & UCase(txtCodigo) & "'," & vFechaProceso & "," & vFechaProceso _
                 & ",dbo.MyGetdate()," & (CCur(txtABAmortizacion.Text) + CCur(txtABIntCor.Text) + CCur(txtABIntMor.Text) + CCur(txtABCargos.Text)) _
                 & "," & CCur(txtABIntCor) & "," & CCur(txtABIntMor.Text) & "," & CCur(txtABAmortizacion) & "," & CCur(txtABCargos) _
                 & ",'A','A','" & vTipoDoc & "','" & lngRecibo & "','" & glogon.Usuario & "','" & vConcepto & "','')"
          Call ConectionExecute(strSQL)
        End If
          
    Else
      
      'Esquema de Anulacion de Cuotas
        If Abs(CCur(txtIntCor) - CCur(txtABIntCor)) < 1 And _
           Abs(CCur(txtCargos) - CCur(txtABCargos)) < 1 And _
           Abs(CCur(txtAmortizacion) - CCur(txtABAmortizacion)) < 1 And _
           Abs(CCur(txtIntCor) - CCur(txtABIntCor)) < 1 Then 'Anulacion Total o Parcial
             
            'todo el registro
            'RECUPERA EL ESTADO ORIGINAL PARA EL ASIENTO
             rs2.CursorLocation = adUseServer
             rs2.Open "select * from morosidad where id_moro =" & vLlave, glogon.Conection, adOpenStatic
             
             
            'INICIA ANULACION
            
             'Establece la cuota de Cancelada a Nula
             strSQL = "update morosidad set estado = 'N' where id_moro =" & vLlave
             Call ConectionExecute(strSQL)
                    
             'Ingresa el detalle de la ND con la Anulacion de la cuota
             'Tambien como Nula, para pista del movimiento
             
             strSQL = "insert morosidad(id_solicitud,codigo,fechap,fecap,fecult,cuota_morosa,intc,intm,amortiza,cargo,estado,estadoi,tcon,ncon" _
                    & ",abintc,abintm,abamortiza,abCargo,usuario,cod_concepto,cod_caja) values(" & rs2!ID_SOLICITUD & ",'" & rs2!Codigo & "'," & rs2!fechap & "," & rs2!fecap _
                    & ",dbo.MyGetdate()," & IIf(IsNull(rs2!cuota_morosa), 0, rs2!cuota_morosa) & "," & rs2!IntC & "," & rs2!IntM _
                    & "," & rs2!Amortiza & "," & rs2!Cargo & ",'N','C','" & vTipoDoc & "','" & lngRecibo & "'," & rs2!abintc & "," & rs2!abintm _
                    & "," & rs2!abAmortiza & "," & rs2!AbCargo & ",'" & glogon.Usuario & "','" & vConcepto & "','')"
             Call ConectionExecute(strSQL)
            
             'Ingresa Cuota de mora Activa, con referencia a ND de la anulacion
             'la cual puede ser reemplazada por un futuro abono
             If chkGeneraMorosidad.Value = vbChecked Then
                strSQL = "insert morosidad(id_solicitud,codigo,fechap,fecap,fecult,cuota_morosa,intc,intm,amortiza,cargo,estado,estadoi,tcon,ncon,usuario,cod_concepto,cod_caja)" _
                       & " values(" & rs2!ID_SOLICITUD & ",'" & rs2!Codigo & "'," & rs2!fechap & "," & vFechaProceso & ",dbo.MyGetdate()" _
                       & "," & rs2!abintc + rs2!abintm + rs2!abAmortiza & "," & rs2!abintc & "," & rs2!abintm & "," & rs2!abAmortiza _
                       & "," & rs2!AbCargo & ",'A','A','" & vTipoDoc & "','" & lngRecibo & "','" & glogon.Usuario & "','" & vConcepto & "','')"
                Call ConectionExecute(strSQL)
             End If
            
            Call Bitacora("Anula", "OP: " & txtOperacion & " INTC: " & rs2!abintc & " INTM: " & rs2!abintm & " AMORT: " & rs2!abAmortiza)
           
            rs2.Close
           
           
           Else 'anulacion parcial
             
               
             Call Bitacora("Anula", "OP: " & txtOperacion & " INTC: " & txtABIntCor & " INTM: " & txtABIntMor & " AMORT: " & txtABAmortizacion)
               
            
            'Solo Hay Que insertar los datos en pantalla
             rs.Open "select * from morosidad where id_moro = " & vLlave, glogon.Conection, adOpenStatic
               
              
             'Ingresa el detalle de la ND con la Anulacion de la cuota
             'Tambien como Nula, para pista del movimiento
             
             strSQL = "insert morosidad(id_solicitud,codigo,fechap,fecap,fecult,cuota_morosa,intc,intm,amortiza,cargo,estado,estadoi,tcon,ncon" _
                    & ",abintc,abintm,abamortiza,abCargo,usuario,cod_concepto,cod_caja) values(" & rs!ID_SOLICITUD & ",'" & rs!Codigo & "'," & rs!fechap & "," & rs!fecap _
                    & ",dbo.MyGetdate()," & rs!cuota_morosa & "," & CCur(txtABIntCor) & "," & CCur(txtABIntMor) _
                    & "," & CCur(txtABAmortizacion) & "," & CCur(txtABCargos.Text) & ",'N','C','8','" & lngRecibo & "'," & CCur(txtABIntCor) _
                    & "," & CCur(txtABIntMor) & "," & CCur(txtABAmortizacion) & "," & CCur(txtABCargos.Text) & ",'" & glogon.Usuario & "','" & vConcepto & "','')"
             Call ConectionExecute(strSQL)
              
             'Ingresa Cuota de mora Activa, con referencia a ND de la anulacion
             'la cual puede ser reemplazada por un futuro abono
             
             If chkGeneraMorosidad.Value = vbChecked Then
                strSQL = "insert morosidad(id_solicitud,codigo,estado,estadoi,fechap,fecap,fecult,intc,intm,amortiza,cargo,cuota_morosa,tcon,ncon)" _
                        & " values(" & rs!ID_SOLICITUD & ",'" & rs!Codigo & "','A','C'," & rs!fechap & "," & vFechaProceso _
                        & ",dbo.MyGetdate()," & CCur(txtABIntCor) & "," & CCur(txtABIntMor) _
                        & "," & CCur(txtABAmortizacion) & "," & CCur(txtABCargos.Text) & "," & CCur(txtABIntCor) + CCur(txtABIntMor) + CCur(txtABAmortizacion) + CCur(txtABCargos.Text) _
                        & ",'" & vTipoDoc & "','" & lngRecibo & "')"
                Call ConectionExecute(strSQL)
             End If
             rs.Close
            
            End If 'Si es todo el registro o parcial
        
    End If 'Mora Personalizada o Por Cuota
   

 End If 'Visualizacion
 
 Call sbDocumento("ND", lngRecibo, vConcepto, vCuenta, "ANULACION DE ABONOS", CCur(txtABIntCor.Text), CCur(txtABIntMor.Text), CCur(txtABCargos.Text), CCur(txtABAmortizacion.Text))
   
 MsgBox "Anulación Realizada ... Con Nota Debito #" & lngRecibo, vbInformation

 'Imprime Comprobante
 Call sbImprimeRecibo(lngRecibo, "ND")
 Call cboVisualiza_Click

Exit Sub

vError:
  Me.MousePointer = vbDefault
  
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Function fxValidaInformacion() As Boolean
 
 fxValidaInformacion = True
 
 If Len(Trim(txtABIntCor)) = 0 Then
   txtABIntCor = 0
 End If
 

 If Len(Trim(txtABIntMor)) = 0 Then
   txtABIntMor = 0
 End If
 
 If Len(Trim(txtABCargos)) = 0 Then
   txtABCargos = 0
 End If
 
 If Len(Trim(txtABAmortizacion)) = 0 Then
   txtABAmortizacion = 0
 End If
 
 If Len(Trim(txtFecha)) > 0 Then
    If CCur(txtIntCor) < CCur(txtABIntCor) Or _
       CCur(txtAmortizacion) < CCur(txtABAmortizacion) Or _
       CCur(txtIntMor) < CCur(txtABIntMor) Then
        fxValidaInformacion = False
    End If
 End If
 
  If (CCur(txtABAmortizacion) + CCur(txtABIntCor) + CCur(txtABIntMor)) = 0 Then
    fxValidaInformacion = False
 End If

 If Len(Trim(txtNombre.Text)) = 0 Then
    fxValidaInformacion = False
 End If

End Function


Private Sub sbDocumento(pTipoDoc As String, pNumDoc As Long, pConcepto As String, pCuenta As String, pDetalle As String _
                      , curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency)
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
Dim strCliente As String


'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)


strCliente = Trim(txtCedula.Text) & " - " & Trim(txtNombre.Text)
strCliente = Mid(strCliente, 1, 45)


strLinea(1) = "Saldo Anterior    " & Format(rs!Saldo - curAmortiza, "Standard")
strLinea(2) = "Interes Corriente " & Format(curIntC * -1, "Standard")
strLinea(3) = "Interes Moratorio " & Format(curIntM * -1, "Standard")
strLinea(4) = "Cargos            " & Format(curCargo * -1, "Standard")
strLinea(5) = "Amortizacion      " & Format(curAmortiza * -1, "Standard")
strLinea(6) = "Saldo Actual      " & Format(rs!Saldo, "Standard")
strLinea(7) = "Operación         " & txtOperacion & ".." & txtCodigo & ".." & UCase(txtOpex.Text)
strLinea(8) = "Divisa: " & rs!cod_Divisa & " / Tipo Cambio: " & rs!TipoCambio
strLinea(9) = "Proc.Retencion    " & IIf(vRetencion, "SI", "NO")
strLinea(10) = "Usuario           " & glogon.Usuario

'Control de Documentos 2

strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
        & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
        & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento)" _
        & " values('" & pNumDoc & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
        & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curIntC + curIntM + curAmortiza + curCargo & ",'P','" & txtOperacion.Text _
        & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
        & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
        & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
        & vAseDocDetalle & "','" & vAseDocDeposito & "')"
Call ConectionExecute(strSQL)


If curIntC > 0 Then
  strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntC * rs!TipoCambio & ",'D','" & rs!cod_Divisa _
         & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintc _
         & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)
End If

If curIntM > 0 Then
  strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntM * rs!TipoCambio & ",'D','" & rs!cod_Divisa _
         & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintm _
         & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)
End If

If curCargo > 0 Then
  strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curCargo * rs!TipoCambio & ",'D','" & rs!cod_Divisa _
         & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!CtaCargos _
         & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)
End If


If curAmortiza > 0 Then
  strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curAmortiza * rs!TipoCambio & ",'D','" & rs!cod_Divisa _
         & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaamortiza _
         & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)
End If

'ASIENTO
If curIntC + curIntM + curAmortiza + curCargo > 0 Then
  strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & (curIntC + curIntM + curCargo + curAmortiza) * rs!TipoCambio & ",'C','" & rs!cod_Divisa _
         & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & pCuenta _
         & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)
End If

rs.Close

End Sub

Private Sub cmdMarcas_Click()
Dim i As Integer, curIntMor As Currency, curIntCor As Currency
Dim curAmortiza As Currency, curCargo As Currency

curIntCor = 0
curIntMor = 0
curCargo = 0
curAmortiza = 0

For i = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(i).Checked Then
     'Ordinario
       curIntCor = curIntCor + CCur(lsw.ListItems.Item(i).SubItems(4))
       curIntMor = curIntMor + CCur(lsw.ListItems.Item(i).SubItems(5))
       curCargo = curCargo + CCur(lsw.ListItems.Item(i).SubItems(6))
       curAmortiza = curAmortiza + CCur(lsw.ListItems.Item(i).SubItems(7))
 End If
Next i

txtABIntCor = Format(curIntCor, "Standard")
txtABIntMor = Format(curIntMor, "Standard")
txtABCargos.Text = Format(curCargo, "Standard")
txtABAmortizacion = Format(curAmortiza, "Standard")

txtABTotal.Text = Format(curIntCor + curIntMor + curCargo + curAmortiza, "Standard")

'Personalizado
txtFecha = ""
txtIntCor = 0
txtIntMor = 0
txtCargos = 0
txtAmortizacion = 0
txtTotal.Text = 0

End Sub

Private Sub Form_Activate()
 vModulo = 3
End Sub

Private Sub Form_Load()
 vModulo = 3
 
 Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 
vPaso = True

 cboVisualiza.Clear
 cboVisualiza.AddItem "Abonos Ordinarios"
 cboVisualiza.AddItem "Abonos a Cuotas Morosas"


With lsw.ColumnHeaders
    .Clear
    .Add , , "Proceso", 1200
    .Add , , "Fecha", 1200, vbCenter
    .Add , , "Cuota", 1500, vbRightJustify
    .Add , , "Abono", 1500, vbRightJustify
    .Add , , "Int.Cor.", 1500, vbRightJustify
    .Add , , "Int.Mor.", 1500, vbRightJustify
    .Add , , "Cargos", 1500, vbRightJustify
    .Add , , "Amortización", 1500, vbRightJustify
    .Add , , "Tipo Doc.", 1200, vbCenter
    .Add , , "Num.Doc.", 2100
    .Add , , "Id", 1200, vbCenter
End With

lsw.Checkboxes = True
 
vPaso = False
 
 If GLOBALES.SysPlanPagos = 1 Then
    TimerVerificaPlanPagos.Interval = 10
 Else
   'Carga Load Normalmente
    cboVisualiza.Text = "Abonos Ordinarios"
    Call Formularios(Me)
    Call RefrescaTags(Me)
 End If

End Sub

Private Sub LimpiaDatos()
    txtOperacion.Text = ""
    txtCedula.Text = ""
    txtNombre.Text = ""
    
    
    txtCodigo.Text = ""
    txtDescripcion.Text = ""
    
    txtFecha.Text = ""
    
    txtIntMor.Text = 0
    txtIntCor.Text = 0
    txtCargos.Text = 0
    txtAmortizacion.Text = 0
    
    txtABIntCor.Text = 0
    txtABIntMor.Text = 0
    txtABCargos.Text = 0
    txtABAmortizacion.Text = 0
    
    txtTotal.Text = 0
    txtABTotal.Text = 0
    
    chkGeneraMorosidad.Value = vbUnchecked
End Sub




Private Sub sbCalculaTotal()
On Error GoTo vError
    
txtABTotal.Text = Format(CCur(txtABAmortizacion.Text) + CCur(txtABIntCor.Text) + CCur(txtABIntMor.Text) + CCur(txtABCargos.Text), "Standard")

vError:
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If Item.Checked = True Then

   txtFecha = Mid(Item.Text, 1, 4) + Mid(Item.Text, 6, 2)
   txtIntCor.Text = Format(CCur(Item.ListSubItems.Item(4).Text), "Standard")
   txtIntMor.Text = Format(CCur(Item.ListSubItems.Item(5).Text), "Standard")
   txtCargos.Text = Format(CCur(Item.ListSubItems.Item(6).Text), "Standard")
   txtAmortizacion.Text = Format(CCur(Item.ListSubItems.Item(7).Text), "Standard")
       
   vLlave = Item.ListSubItems(10).Text
       
   txtTotal.Text = Format(CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtCargos.Text) + CCur(txtAmortizacion.Text), "Standard")
   
   
Else
  txtFecha.Text = ""
  txtIntCor.Text = "0.00"
  txtIntMor.Text = "0.00"
  txtCargos.Text = "0.00"
  txtAmortizacion.Text = "0.00"
  txtTotal.Text = "0.00"
End If

txtABIntCor.Text = txtIntCor.Text
txtABIntMor.Text = txtIntMor.Text
txtABCargos.Text = txtCargos.Text
txtABAmortizacion.Text = txtAmortizacion.Text
txtABTotal.Text = txtTotal.Text

End Sub

Private Sub TimerVerificaPlanPagos_Timer()

TimerVerificaPlanPagos.Interval = 0
Call sbFormsCall("frmCR_AnulaAbonosNew", 0, , , False)

Unload Me
End Sub

Private Sub txtABAmortizacion_GotFocus()
On Error GoTo vError
    txtABAmortizacion.Text = CCur(txtABAmortizacion.Text)
vError:
End Sub

Private Sub txtABAmortizacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtABTotal.SetFocus
End Sub

Private Sub txtABAmortizacion_LostFocus()
On Error GoTo vError

txtABAmortizacion.Text = Format(CCur(txtABAmortizacion.Text), "Standard")
Call sbCalculaTotal

vError:
End Sub

Private Sub txtABCargos_GotFocus()
On Error GoTo vError
    txtABCargos.Text = CCur(txtABCargos.Text)
vError:
End Sub

Private Sub txtABCargos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtABAmortizacion.SetFocus
End Sub

Private Sub txtABCargos_LostFocus()
On Error GoTo vError

txtABCargos.Text = Format(CCur(txtABCargos.Text), "Standard")
Call sbCalculaTotal

vError:
End Sub

Private Sub txtABIntCor_GotFocus()
On Error GoTo vError
    txtABIntCor.Text = CCur(txtABIntCor.Text)
vError:
End Sub

Private Sub txtABIntCor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtABIntMor.SetFocus

vError:
End Sub

Private Sub txtABIntCor_LostFocus()
On Error GoTo vError

txtABIntCor.Text = Format(CCur(txtABIntCor.Text), "Standard")
Call sbCalculaTotal

vError:
End Sub

Private Sub txtABIntMor_GotFocus()
On Error GoTo vError
txtABIntMor.Text = CCur(txtABIntMor.Text)
vError:
End Sub

Private Sub txtABIntMor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtABCargos.SetFocus
End Sub

Private Sub txtABIntMor_LostFocus()
On Error GoTo vError

txtABIntMor.Text = Format(CCur(txtABIntMor.Text), "Standard")
Call sbCalculaTotal

vError:
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 txtABIntCor.SetFocus
End If
End Sub


Public Sub sbConsultaExterna(xOpTemp As Long)
 txtOperacion.Text = xOpTemp
 Call txtOperacion_KeyPress(vbKeyReturn)
End Sub


Private Sub txtOperacion_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
  Call sbConsulta(txtOperacion)
Else
  txtOpex.Text = ""
  txtCedula.Text = ""
  txtNombre.Text = ""
  txtCodigo = ""
  txtDescripcion.Text = ""
  txtFecha.Text = ""
  txtIntCor.Text = 0
  txtIntMor.Text = 0
  txtCargos.Text = 0
  txtABCargos.Text = 0
  txtAmortizacion.Text = 0
  txtABIntCor.Text = 0
  txtABIntMor.Text = 0
  txtABAmortizacion.Text = 0
  
  txtTotal.Text = 0
  txtABTotal.Text = 0
  
  chkGeneraMorosidad.Value = vbUnchecked
  
  lsw.ListItems.Clear
End If

End Sub


Private Sub sbConsulta(pOperacion As String)
Dim strSQL As String, rs As New ADODB.Recordset

 strSQL = "select R.cedula,S.nombre,R.codigo,C.descripcion,C.retencion,C.poliza" _
        & ",R.estado,R.proceso,R.opex" _
        & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
        & " inner join Socios S on R.cedula = S.cedula where R.id_solicitud = " & pOperacion
 Call OpenRecordSet(rs, strSQL)

 If rs.EOF And rs.BOF Then
   MsgBox "No se encontró número de operación solicitada:" & pOperacion
   Call LimpiaDatos
   Exit Sub
 End If
 
 If rs!Estado = "N" Or IsNull(rs!Estado) Then
   MsgBox "Se encontró número de operación pero se encuentra en tramite o se anulo del estado de cuenta Anulada:" & pOperacion
   Call LimpiaDatos
   Exit Sub
 End If
 
 If rs!Proceso = "J" Then
   MsgBox "Esta operación se encuentra en Cobro Judicial no puede recibir movimientos...", vbInformation
   Call LimpiaDatos
   Exit Sub
 End If
 
 txtCodigo = rs!Codigo
 txtDescripcion.Text = rs!Descripcion
 txtCedula = rs!Cedula
 txtNombre.Text = rs!Nombre
 
    Select Case rs!Proceso
      Case "N"
        txtProceso.Text = "Normal"
      Case "T"
        txtProceso.Text = "Traspaso Deuda"
      Case "J"
        txtProceso.Text = "Cobro Judicial"
      Case "I"
        txtProceso.Text = "Incobrable"
    End Select
    
    
    If (rs!opex & "") = 1 Then
        txtOpex.Text = "Op.Ex."
    Else
        txtOpex.Text = "Interno"
    End If
 
 
 
 If rs!retencion = "S" Or rs!Poliza = "S" Then
   vRetencion = True
 Else
   vRetencion = False
 End If
   
 rs.Close
 
 Call sbCargoMovimientos
 

End Sub


Private Sub sbCargoMovimientos()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

If txtOperacion = "" Or Not IsNumeric(txtOperacion) Then
  Exit Sub
End If

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

 If cboVisualiza.Text = "Abonos Ordinarios" Then
   
   'Excluye Tcon = 8, que son ND de Anulaciones Anteriores
   
   strSQL = "select C.*,isnull(D.Descripcion,C.Tcon) as 'TipoDoc' " _
          & " from creditos_dt C left join SIF_DOCUMENTOS D on C.tcon = D.tipo_Documento" _
          & " where C.id_solicitud = " & txtOperacion _
          & " and C.estado = 'A' and C.tcon not in('6','8','ND') " _
          & " order by C.fechas desc"
   Call OpenRecordSet(rs, strSQL)
      
      Do While Not rs.EOF
       Set itmX = lsw.ListItems.Add(, , Format(rs!fechap, "####-##"))
         itmX.SubItems(1) = Format(rs!fechas, "dd/mm/yyyy")
         itmX.SubItems(2) = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
         itmX.SubItems(3) = Format(IIf(IsNull(rs!Abono), 0, rs!Abono), "Standard")
         itmX.SubItems(4) = Format(IIf(IsNull(rs!intcp), 0, rs!intcp), "Standard")
         itmX.SubItems(5) = Format(0, "Standard")
         itmX.SubItems(6) = Format(0, "Standard")
         
         itmX.SubItems(7) = Format(IIf(IsNull(rs!Amortiza), 0, rs!Amortiza), "Standard")
         itmX.SubItems(8) = rs!TipoDoc
         itmX.SubItems(9) = IIf(IsNull(rs!nCon), 0, rs!nCon)
         itmX.SubItems(10) = rs!consec
         
        rs.MoveNext
      Loop
 Else
 
   strSQL = "select C.*,isnull(D.Descripcion,C.Tcon) as 'TipoDoc' " _
          & " from morosidad C left join SIF_DOCUMENTOS D on C.tcon = D.tipo_Documento" _
          & " where C.id_solicitud = " & txtOperacion _
          & " and C.estado = 'C' and C.tcon not in('6','8','ND') " _
          & " order by C.fechap desc"
 
    Call OpenRecordSet(rs, strSQL)
      
      Do While Not rs.EOF
       Set itmX = lsw.ListItems.Add(, , Format(rs!fechap, "####-##"))
         itmX.SubItems(1) = IIf(IsNull(rs!FecUlt), "", Format(rs!FecUlt, "dd/mm/yyyy"))
         itmX.SubItems(2) = Format(rs!IntC + rs!IntM + rs!Amortiza + rs!Cargo, "Standard")
         itmX.SubItems(3) = Format(IIf(IsNull(rs!abintc), 0, rs!abintc) + IIf(IsNull(rs!abintm), 0, rs!abintm) _
                            + IIf(IsNull(rs!abAmortiza), 0, rs!abAmortiza) + IIf(IsNull(rs!AbCargo), 0, rs!AbCargo), "Standard")
         
         itmX.SubItems(4) = Format(IIf(IsNull(rs!abintc), 0, rs!abintc), "Standard")
         itmX.SubItems(5) = Format(IIf(IsNull(rs!abintm), 0, rs!abintm), "Standard")
         itmX.SubItems(6) = Format(IIf(IsNull(rs!AbCargo), 0, rs!AbCargo), "Standard")
         itmX.SubItems(7) = Format(IIf(IsNull(rs!abAmortiza), 0, rs!abAmortiza), "Standard")
         
         
         itmX.SubItems(8) = rs!TipoDoc
         itmX.SubItems(9) = IIf(IsNull(rs!nCon), 0, rs!nCon)
         itmX.SubItems(10) = rs!id_moro
         
        rs.MoveNext
      Loop
  
  End If
 
 rs.Close

Me.MousePointer = vbDefault

End Sub

