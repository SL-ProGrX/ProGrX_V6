VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCR_Retenciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retenciones"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "frmCR_Retenciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   8820
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   6060
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3951
            MinWidth        =   3951
            Object.ToolTipText     =   "Usuario > Registro"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3246
            MinWidth        =   3246
            Object.ToolTipText     =   "Fecha > Registro"
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
   Begin VB.Frame fraOperacion 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   8535
      Begin VB.TextBox txtPendiente 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
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
         Height          =   314
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   4800
         Width           =   1935
      End
      Begin VB.TextBox txtPagado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
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
         Height          =   314
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox txtProyectado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
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
         Height          =   314
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox txtPlazoTrasnscurrido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4800
         Width           =   1692
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   4440
         Width           =   1692
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   4080
         Width           =   1692
      End
      Begin XtremeSuiteControls.ComboBox cboDestino 
         Height          =   315
         Left            =   3120
         TabIndex        =   26
         Top             =   720
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9128
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.ComboBox cboGarantia 
         Height          =   315
         Left            =   3120
         TabIndex        =   27
         Top             =   1080
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2990
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   315
         Left            =   3120
         TabIndex        =   30
         Top             =   0
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9123
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   315
         Left            =   3120
         TabIndex        =   31
         Top             =   360
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9123
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboMes 
         Height          =   315
         Left            =   6480
         TabIndex        =   32
         Top             =   3000
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   315
         Left            =   6360
         TabIndex        =   33
         Top             =   1080
         Width           =   1935
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   315
         Left            =   3120
         TabIndex        =   34
         Top             =   1440
         Width           =   1695
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   315
         Left            =   6360
         TabIndex        =   35
         Top             =   1440
         Width           =   1935
         _Version        =   1572864
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtObservaciones 
         Height          =   675
         Left            =   3120
         TabIndex        =   36
         Top             =   1800
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9123
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAnio 
         Height          =   330
         Left            =   5760
         TabIndex        =   37
         Top             =   3000
         Width           =   750
         _Version        =   1572864
         _ExtentX        =   1323
         _ExtentY        =   573
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   315
         Left            =   1440
         TabIndex        =   28
         Top             =   0
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2984
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   315
         Left            =   1440
         TabIndex        =   29
         Top             =   360
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2984
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   315
         Left            =   5640
         TabIndex        =   39
         Top             =   1440
         Width           =   735
         _Version        =   1572864
         _ExtentX        =   1291
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboFrecuencia 
         Height          =   315
         Left            =   6480
         TabIndex        =   40
         Top             =   3360
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.ComboBox cboDeductora 
         Height          =   330
         Left            =   3120
         TabIndex        =   41
         Top             =   2640
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9128
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin VB.Label Label5 
         Caption         =   "Deductora"
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
         Left            =   1800
         TabIndex        =   42
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Destino"
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
         Index           =   28
         Left            =   1800
         TabIndex        =   23
         Top             =   720
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "Garantía"
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
         Left            =   1800
         TabIndex        =   22
         Top             =   1080
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "Primer Deducción (aaaa/mm)"
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
         Left            =   3000
         TabIndex        =   18
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pendiente"
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
         Index           =   11
         Left            =   4800
         TabIndex        =   17
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proyectado"
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
         Index           =   10
         Left            =   4800
         TabIndex        =   16
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
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
         Left            =   4920
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo Trans."
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
         Index           =   8
         Left            =   1680
         TabIndex        =   14
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Index           =   7
         Left            =   1680
         TabIndex        =   12
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pagado"
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
         Index           =   5
         Left            =   4800
         TabIndex        =   11
         Top             =   4440
         Width           =   975
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
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   8
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Observación"
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
         Left            =   1800
         TabIndex        =   7
         Top             =   1800
         Width           =   1212
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   16
         Left            =   5040
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Height          =   252
         Index           =   6
         Left            =   1800
         TabIndex        =   5
         Top             =   1440
         Width           =   492
      End
      Begin VB.Label Label1 
         Caption         =   "Línea"
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
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   264
      Left            =   4800
      TabIndex        =   0
      Top             =   240
      Width           =   2916
      _ExtentX        =   5133
      _ExtentY        =   476
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
            Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
            Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
            Object.ToolTipText     =   "Imprime el listado seleccionado"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
            Object.Tag             =   "1"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3480
      TabIndex        =   25
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   315
      Left            =   1560
      TabIndex        =   38
      Top             =   240
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2984
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   8400
      X2              =   120
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmCR_Retenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje As String 'Envia Mensajes en Fallas de Verificacion
Dim vEdita As Boolean 'Indica si se esta actualizando o insertando
Dim vPaso As Boolean, vScroll As Boolean
Dim mFrecuenciaPago As String, mBaseCalculo As String

Private Function fxVerificaRetencion() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngPriDeduc As Long, pAnio As Long

On Error GoTo vError

fxVerificaRetencion = True

vMensaje = ""

If txtDocumento.Text = "" Then vMensaje = vMensaje & vbCrLf & "- No se especificó el # Documento ? "

pAnio = Year(Date)

If Len(txtAnio.Text) <> 4 Or Not IsNumeric(txtAnio.Text) Then
    vMensaje = vMensaje & vbCrLf & "- El año de Primer Deducción es inválido, verifique que tenga este format [aaaa]"
Else
    If CLng(txtAnio.Text) < pAnio Or CLng(txtAnio.Text) > pAnio + 1 Then
        vMensaje = vMensaje & vbCrLf & "- El año de Primer Deducción no puede ser menor al actual o mayor a un año adicional!"
    End If
End If



If IsNumeric(txtPlazo) Then
 If txtPlazo < 1 Then vMensaje = vMensaje & vbCrLf & "- El Plazo Solicitado es Inválido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Plazo Solicitado es Inválido"
End If

If cboGarantia.Text = "" Or cboGarantia.ListCount <= 0 Then vMensaje = vMensaje & vbCrLf & "- No se especificó el tipo de garantía"

If Not fxVerificaCodigoDoble(txtCodigo, txtCedula) _
   Then vMensaje = vMensaje & vbCrLf & "- No se permite sobrepasar el número máximo de operaciones o el Rango (Monto) máximo de la línea"

If IsNumeric(txtMonto.Text) Then
 If txtMonto.Text < 1 Then vMensaje = vMensaje & vbCrLf & "- El Monto Solicitado Solicitado es Inválido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Monto Solicitado es Inválido"
End If

'Verifica que sea un codigo de retencion
strSQL = "select isnull(count(*),0) as Existe from catalogo where (Retencion = 'S' or Poliza = 'S') and codigo ='" & txtCodigo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
 If rsX!Existe = 0 Then
   vMensaje = vMensaje & vbCrLf & "- El código de retencion no existe o no es válido..."
 End If
rsX.Close

'Verifica que exista la persona
strSQL = "select isnull(count(*),0) as Existe from socios where cedula ='" & txtCedula & "'"
Call OpenRecordSet(rsX, strSQL, 0)
 If rsX!Existe = 0 Then
   vMensaje = vMensaje & vbCrLf & "- No Existe el cliente definido (debe de Ingresarlo como No Socio)..."
 End If
rsX.Close


''Verifica que no existe una retencion con el mismo codigo
'strSQL = "select isnull(count(*),0) as Existe from reg_creditos where estado = 'A' and codigo ='" & txtCodigo & "' and cedula = '" & txtCedula & "'"
'Call OpenRecordSet(rsX, strSQL, 0)
' If rsX!existe > 0 Then
'   vMensaje = vMensaje & vbCrLf & "- Ya existe una retencion activa para esta persona..."
' End If
'rsX.Close

strSQL = "select ctaNintC from catalogo where codigo ='" & txtCodigo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
   vMensaje = vMensaje & vbCrLf & "- El código de la retencion no existe"
 Else
  If IsNull(rsX!ctaNintC) Then vMensaje = vMensaje & vbCrLf & "- El código no se encuentra codificado contablemente"
End If
rsX.Close

If fxConvierteMES(cboMes.Text) = cboMes.Text Then vMensaje = vMensaje & vbCrLf & "- El Mes para la primer deduccion no es válido"

lngPriDeduc = txtAnio.Text & Format(fxConvierteMES(cboMes.Text), "00")

'If lngPriDeduc <= GLOBALES.glngFechaCR Then vMensaje = vMensaje & vbCrLf & "- La primer deducción no es válida porque es igual o menor a la fecha de proceso actual"


If Len(vMensaje) > 0 Then fxVerificaRetencion = False


Exit Function

vError:
  vMensaje = vMensaje & vbCrLf & fxSys_Error_Handler(Err.Description)
  fxVerificaRetencion = False

End Function



Private Sub cboDeductora_Click()
If vPaso Then Exit Sub

On Error GoTo vError

Dim strSQL As String, rs As New ADODB.Recordset
Dim vProceso As Currency, pProcesoClean As Long

strSQL = "select rtrim(descripcion) as 'Descripcion', isnull(Frecuencia,'M') as 'Frecuencia_Id'" _
       & " from instituciones " _
       & " where cod_institucion = " & cboDeductora.ItemData(cboDeductora.ListIndex)
Call OpenRecordSet(rs, strSQL)
    mFrecuenciaPago = rs!Frecuencia_ID
rs.Close

cboFrecuencia.Clear
Select Case mFrecuenciaPago
    Case "M" 'Mensual
        cboFrecuencia.AddItem "Mensual"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "0"
        cboFrecuencia.Text = "Mensual"
    
    Case "Q" 'Quincenal
        cboFrecuencia.AddItem "1er Quincena"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "1"
        cboFrecuencia.AddItem "2da Quincena"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "2"
End Select
  
  
vProceso = fxPrimerDeduccion(txtCodigo.Text, cboDeductora.ItemData(cboDeductora.ListIndex))
pProcesoClean = vProceso

cboMes.Text = fxConvierteMES(Val(Mid(pProcesoClean, 5, 2)))
txtAnio.Text = Mid(pProcesoClean, 1, 4)

If mFrecuenciaPago = "Q" Then
    If (vProceso - pProcesoClean) = 0.1 Then
        cboFrecuencia.Text = "1er Quincena"
    Else
        cboFrecuencia.Text = "2da Quincena"
    End If
End If

Exit Sub

vError:

End Sub

Private Sub cboDestino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboGarantia.SetFocus
End Sub



Private Sub cboGarantia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtOperacion.Text = "" Then txtOperacion.Text = "0"

If vScroll Then
    strSQL = "select Top 1 R.id_solicitud from reg_creditos R inner join Catalogo C on R.codigo = C.codigo"

    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where (C.retencion = 'S' or C.poliza = 'S') and R.id_solicitud > " & txtOperacion & " order by R.id_solicitud asc"
    Else
       strSQL = strSQL & " where (C.retencion = 'S' or C.poliza = 'S') and R.id_solicitud < " & txtOperacion & " order by R.id_solicitud desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtOperacion = rs!Id_Solicitud
      Call sbCargaOperacion
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

Private Sub Form_Load()
 
 vModulo = 3
 
 mFrecuenciaPago = "M"
 
 Call sbToolBarIconos(tlbPrincipal, False)
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 With cboMes
    .Clear
    .AddItem "Enero"
    .AddItem "Febrero"
    .AddItem "Marzo"
    .AddItem "Abril"
    .AddItem "Mayo"
    .AddItem "Junio"
    .AddItem "Julio"
    .AddItem "Agosto"
    .AddItem "Septiembre"
    .AddItem "Octubre"
    .AddItem "Noviembre"
    .AddItem "Diciembre"
 End With

 
 Call Formularios(Me)
 
 Call sbLimpia
 
 With tlbPrincipal.Buttons
   .Item(2).Enabled = False
   .Item(3).Enabled = False
 End With
 
Call RefrescaTags(Me)

End Sub

Private Sub sbLimpia()
 
 txtCedula = ""
 txtCodigo = ""
 txtDescripcion = ""
 txtNombre = ""
 txtObservaciones = ""
 txtPlazo.Text = "1"
 txtMonto.Text = "0"
 txtPagado.Text = "0"
 txtPendiente.Text = "0"
 txtProyectado.Text = "0"
 txtDocumento = ""
 txtFecha = ""
 txtEstado = ""
 txtPlazoTrasnscurrido = ""
 
 cboMes.Text = fxConvierteMES(Val(Mid(fxPrimerDeduccion, 5, 2)))
 txtAnio.Text = Mid(fxPrimerDeduccion, 1, 4)
 
 StatusBarX.Panels(1).Text = ""
 StatusBarX.Panels(2).Text = ""
 
 
End Sub

Private Function fxOperacionDestino(vDestino As String) As String
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select rtrim(cod_destino) + ' - ' + descripcion as ItemX from catalogo_destinos where cod_destino = '" & vDestino & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  fxOperacionDestino = " -"
Else
  fxOperacionDestino = rs!itemx
End If
rs.Close

End Function

Private Sub sbCargaOperacion()

Dim strSQL As String, rs As New ADODB.Recordset
Dim vTemp As String

On Error GoTo vError

strSQL = "select R.id_solicitud,R.codigo,C.descripcion,R.cedula,S.nombre,R.cuota,R.estado" _
       & " ,R.observacion,R.fechaforp,R.plazo,R.amortiza,R.cuotas_planilla,R.cuotas_directas" _
       & " ,R.documento_referido,R.prideduc,R.userRec,R.cod_destino,R.garantia" _
       & " ,RTRIM(isnull(Gt.DESCRIPCION,'')) as 'GarantiaDesc'" _
       & " ,RTRIM(isnull(Cd.DESCRIPCION,'')) as 'DestinoDesc', R.Base_Calculo, R.Cod_Divisa" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo " _
       & " inner join Socios S on R.cedula = S.cedula " _
       & "  left join CRD_GARANTIA_TIPOS Gt on R.GARANTIA = Gt.GARANTIA" _
       & "  left join CATALOGO_DESTINOS Cd on R.COD_DESTINO = Cd.COD_DESTINO" _
       & " where R.estadosol = 'F' and (C.retencion = 'S' or C.poliza = 'S')" _
       & "   and R.id_solicitud = " & txtOperacion
       
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
 
 txtCedula = rs!Cedula
 txtNombre = rs!Nombre
 txtCodigo = rs!Codigo
 txtDescripcion = rs!Descripcion
 
 
 txtDivisa.Text = rs!COD_DIVISA & ""
 mBaseCalculo = rs!Base_Calculo & ""
 
 txtMonto = Format(rs!Cuota, "Standard")
 txtPlazo = CStr(rs!Plazo)
 
 txtEstado = fxEstadoCuota(rs!Estado)
 txtObservaciones = IIf(IsNull(rs!observacion), "", rs!observacion)
 
 txtFecha = Format((rs!FechaForp & ""), "dd/mm/yyyy")
 txtPlazoTrasnscurrido = rs!cuotas_planilla + rs!cuotas_directas
 txtDocumento = rs!documento_referido & ""
  
 
 txtPagado.Text = Format(rs!Amortiza, "Standard")
 If rs!Plazo >= 999 Then
    txtProyectado.Text = Format(rs!Cuota, "Standard")
    txtPendiente = Format(rs!Cuota, "Standard")
 Else
    txtProyectado.Text = Format(rs!Cuota * rs!Plazo, "Standard")
    txtPendiente = Format((rs!Cuota * rs!Plazo) - rs!Amortiza, "Standard")
 End If
 
 With tlbPrincipal.Buttons
   .Item(1).Enabled = True
   .Item(2).Enabled = False
   .Item(3).Enabled = False
 End With
 Me.fraOperacion.Enabled = False

 If IsNull(rs!PriDeduc) Then
  cboMes.Text = fxConvierteMES(Val(Mid(fxPrimerDeduccion, 5, 2)))
  txtAnio.Text = Mid(fxPrimerDeduccion, 1, 4)
 Else
  cboMes.Text = fxConvierteMES(Val(Mid(rs!PriDeduc, 5, 2)))
  txtAnio.Text = Mid(rs!PriDeduc, 1, 4)
 End If


 'Carga Destino
 Call sbSTCargaCboDestinos(cboDestino, rs!Codigo)
 Call sbSTCargaCboGarantia(cboGarantia, rs!Codigo)
 
 If (rs!cod_destino & "") <> "" Then
    Call sbCboAsignaDato(cboDestino, rs!DestinoDesc, True, rs!cod_destino)
 End If
 
 Call sbCboAsignaDato(cboGarantia, rs!GarantiaDesc, True, rs!Garantia)


 StatusBarX.Panels(1).Text = rs!userRec & ""
 StatusBarX.Panels(2).Text = Format(rs!FechaForp, "dd/mm/yyyy")

Else
 
 MsgBox "No existe esta Operación..?", vbExclamation

End If
rs.Close

Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbGuardar()
Dim strSQL As String, vFecha As Date
Dim lngOP As Long, lngPriDeduc As Currency
Dim vComite As Integer, vDestino As String
Dim curMonto As Currency

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

On Error GoTo vError

lngPriDeduc = txtAnio.Text & Format(fxConvierteMES(cboMes.Text), "00") & "." & cboFrecuencia.ItemData(cboFrecuencia.ListIndex)

vComite = fxCrdIdComiteLinea(txtCodigo.Text)

vDestino = cboDestino.ItemData(cboDestino.ListIndex)
If Trim(vDestino) = "" Then
  vDestino = "Null"
Else
  vDestino = "'" & vDestino & "'"
End If

If CInt(txtPlazo) < 999 Then
    curMonto = CCur(txtMonto.Text) * CInt(txtPlazo)
Else
    curMonto = CCur(txtMonto.Text)
End If



'Insertar la operacion
strSQL = "insert reg_creditos(codigo,id_comite,cedula,montosol,montoapr,monto_girado" _
       & ",saldo,amortiza,interesc,saldo_mes,cuota,int,interesv,plazo,userrec,userres" _
       & ",userfor,usertesoreria,tesoreria,fechasol,fechares,fechaforp,fechaforf" _
       & ",fecha_calculo_int,garantia,primer_cuota,tdocumento,ndocumento,pagare" _
       & ",firma_deudor,premio,observacion,estado,prideduc,fecult,estadosol,documento_referido" _
       & ",cod_destino, cod_divisa, base_calculo)" _
       & " values('" & UCase(txtCodigo.Text) & "'," & vComite & ",'" _
       & Trim(txtCedula) & "'," & curMonto & "," & curMonto & ",0," & curMonto & ",0,0," _
       & curMonto & "," & CCur(txtMonto.Text) & ",0,0," & txtPlazo & ",'" & glogon.Usuario & "','" & glogon.Usuario _
       & "','" & glogon.Usuario & "'," & "'" & glogon.Usuario & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
       & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
       & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'" _
       & ",'N','OT','',0,1,0,'" & UCase(txtObservaciones) & "','A'," & lngPriDeduc _
       & "," & fxFechaProcesoAnterior(lngPriDeduc) & ",'F','" & txtDocumento & "'," _
       & vDestino & ",'" & txtDivisa.Text & "','" & mBaseCalculo & "')"
  
 Call ConectionExecute(strSQL)
 lngOP = fxUltimaOperacion(txtCedula)
 
 If GLOBALES.SysPlanPagos = 1 Then
    strSQL = "exec spCrdPlanPagos " & lngOP
    Call ConectionExecute(strSQL)
 End If
 
 
 'Bitacora General
 Call Bitacora("Registra", "Retencion en la OP : " & lngOP)
 
 'Bitacora de Retenciones
 Call sbBitacoraCredito("08", "Op: " & lngOP & " - Monto " & CCur(txtMonto) _
        & " - Plazo: " & txtPlazo, "R", lngOP, UCase(txtCodigo))
 
 txtOperacion = lngOP
 
 
 Me.MousePointer = vbDefault
 
 MsgBox "Retención Grabada Satisfactoriamente...", vbInformation

 Exit Sub
 
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlazo.SetFocus
End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError
  txtMonto.Text = CCur(txtMonto.Text)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservaciones.SetFocus
End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

Select Case Button.Key
 Case "nuevo"
  txtOperacion.Text = ""
  txtOperacion.Enabled = False
  Call sbLimpia
  tlbPrincipal.Buttons(1).Enabled = False
  tlbPrincipal.Buttons(2).Enabled = True
  tlbPrincipal.Buttons(3).Enabled = True
  fraOperacion.Enabled = True
  txtCedula.SetFocus
  
 Case "guardar"
  
  If fxVerificaRetencion Then
    Call sbGuardar
    Call sbCargaOperacion
    txtOperacion.Enabled = True
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = False
    tlbPrincipal.Buttons(3).Enabled = False
    fraOperacion.Enabled = False
  Else
    MsgBox vMensaje, vbCritical
  End If
  
 Case "deshacer"
    txtOperacion.Enabled = True
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = False
    tlbPrincipal.Buttons(3).Enabled = False
    fraOperacion.Enabled = False
    If txtOperacion <> "" Then Call sbCargaOperacion
    txtOperacion.SetFocus
 
 Case "reportes"
    Call ReporteBoletaRetencion
 Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
 
 End Select
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus


If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Col1Name = "Cédula Colilla"
    gBusquedas.Col2Name = "Cédula Real"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    txtCedula = Trim(gBusquedas.Resultado)
    txtNombre.Text = gBusquedas.Resultado3

End If


End Sub


Private Sub sbDeductoras_Load(pInstitucion As Long)
Dim strSQL As String

strSQL = "select COD_DEDUCTORA AS 'IdX', DESCRIPCION AS 'ItmX'" _
       & " From vAFI_Deductoras" _
       & " Where cod_institucion = " & pInstitucion

vPaso = True
    Call sbCbo_Llena_New(cboDeductora, strSQL, False, True)
vPaso = False

End Sub

Private Sub txtCedula_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select S.nombre, isnull(I.DEDUCCION_PLANILLA,0) as 'Deduccion' " _
       & ",S.cod_institucion, Ed.Cod_Institucion as 'DeductoraCod', Ed.Descripcion as 'DeductoraDesc'" _
       & " from Socios S inner join Instituciones I on S.cod_institucion = I.cod_Institucion" _
       & " left join Instituciones Ed on isnull(S.cod_deductora,S.cod_institucion) = Ed.cod_Institucion" _
       & " Where S.cedula = '" & txtCedula.Text & "'"
       
Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
    txtNombre.Text = ""
Else
    txtNombre.Text = Trim(rs!Nombre)

    'Carga Deductoras por Institucion
    Call sbDeductoras_Load(rs!cod_institucion)
    Call sbCboAsignaDato(cboDeductora, rs!DeductoraDesc, True, rs!DeductoraCod)

    cboDeductora.Tag = CStr(rs!DeductoraCod)
    
End If
rs.Close


End Sub

Private Sub ReporteBoletaRetencion()
Dim strRuta As String, rs As New ADODB.Recordset

strRuta = SIFGlobal.fxPathReportes("Credito_BoletaFormalizacion.rpt")
Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Boleta de Formalización / Retenciones"
   
 .Connect = glogon.ConectRPT
 
 .ReportFileName = strRuta
 .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD}=" & txtOperacion
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy HH:MM:ss") & "'"
 .Formulas(1) = "Usuario='" & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 
 .SubreportToChange = "sbAsiento"
 .StoredProcParam(0) = "FRM"
 .StoredProcParam(1) = Operacion.Operacion
 .StoredProcParam(2) = 0
 
 
 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  txtCodigo = UCase(txtCodigo)
  
  vPaso = True
        Call sbSTCargaCboGarantia(cboGarantia, txtCodigo)
        Call sbSTCargaCboDestinos(cboDestino, txtCodigo)
  vPaso = False
  
  cboDestino.SetFocus
End If



If KeyCode = vbKeyF4 Then

        gBusquedas.Consulta = "select Codigo,Descripcion from catalogo"
        gBusquedas.Columna = "Codigo"
        gBusquedas.Orden = "Codigo"
        gBusquedas.Filtro = " and Retencion = 'S' and Codigo not in(select CODIGO_ASE from FND_PLANES group by CODIGO_ASE)"
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        
        frmBusquedas.Show vbModal
        If gBusquedas.Resultado <> "" Then
            txtCodigo = gBusquedas.Resultado
            txtDescripcion = gBusquedas.Resultado2
        End If
End If



Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

txtDivisa.Text = ""
txtDescripcion.Text = ""
mBaseCalculo = "01"

strSQL = "select Cat.CODIGO, Cat.DESCRIPCION, Cat.MONEDA as 'COD_DIVISA'" _
       & " , Cat.ID_COMITE, isnull(Com.DESCRIPCION,'') as 'COMITE_DESC', Cat.BASE_CALCULO" _
       & " from CATALOGO Cat left join COMITES Com on Cat.ID_COMITE = Com.ID_COMITE" _
       & " where Cat.CODIGO = '" & txtCodigo.Text & "'"

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.BOF Then
    txtDescripcion.Text = rs!Descripcion & ""
    txtDivisa.Text = rs!COD_DIVISA
    mBaseCalculo = rs!Base_Calculo
End If
rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtMonto_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
  If CInt(txtPlazo.Text) < 900 Then
      txtProyectado.Text = Format(CInt(txtPlazo.Text) * CCur(txtMonto.Text), "Standard")
      txtPendiente.Text = Format(CCur(txtProyectado.Text) - CCur(txtPagado.Text), "Standard")
  Else
      txtProyectado.Text = Format(CCur(txtMonto.Text), "Standard")
      txtPendiente.Text = Format(CCur(txtProyectado.Text), "Standard")
  End If
vError:
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
  txtMonto.Text = Format(CCur(txtMonto.Text), "Standard")
vError:
End Sub

Private Sub txtObservaciones_LostFocus()
On Error GoTo vError
If txtAnio.Enabled = True Then
 txtAnio.SetFocus
End If
vError:
End Sub

Private Sub txtOperacion_Change()
 Call sbLimpia
  With tlbPrincipal.Buttons
   .Item(2).Enabled = False
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With
End Sub

Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call sbCargaOperacion
End Sub


Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub txtPlazo_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
  If CInt(txtPlazo.Text) < 900 Then
      txtProyectado.Text = Format(CInt(txtPlazo.Text) * CCur(txtMonto.Text), "Standard")
      txtPendiente.Text = Format(CCur(txtProyectado.Text) - CCur(txtPagado.Text), "Standard")
  Else
      txtProyectado.Text = Format(CCur(txtMonto.Text), "Standard")
      txtPendiente.Text = Format(CCur(txtProyectado.Text), "Standard")
  End If
vError:
End Sub
