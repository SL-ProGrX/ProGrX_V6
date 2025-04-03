VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCxC_Reportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de CxC"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   10230
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3132
      Left            =   0
      TabIndex        =   19
      Top             =   3960
      Width           =   5412
      _Version        =   1572864
      _ExtentX        =   9546
      _ExtentY        =   5524
      _StockProps     =   77
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
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   252
      Left            =   8640
      TabIndex        =   6
      Top             =   1680
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkConcepto 
      Height          =   252
      Left            =   8640
      TabIndex        =   7
      Top             =   2640
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkPagador 
      Height          =   252
      Left            =   8640
      TabIndex        =   8
      Top             =   2040
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtClienteId 
      Height          =   330
      Left            =   1440
      TabIndex        =   9
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1680
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
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
      Enabled         =   0   'False
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPagadorId 
      Height          =   330
      Left            =   1440
      TabIndex        =   10
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   2040
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
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
      Enabled         =   0   'False
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtConceptoId 
      Height          =   330
      Left            =   1440
      TabIndex        =   11
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   2640
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
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
      Enabled         =   0   'False
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtClienteNombre 
      Height          =   330
      Left            =   3480
      TabIndex        =   12
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1680
      Width           =   5052
      _Version        =   1572864
      _ExtentX        =   8911
      _ExtentY        =   582
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
      Enabled         =   0   'False
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPagadorNombre 
      Height          =   330
      Left            =   3480
      TabIndex        =   13
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   2040
      Width           =   5052
      _Version        =   1572864
      _ExtentX        =   8911
      _ExtentY        =   582
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
      Enabled         =   0   'False
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtConceptoDesc 
      Height          =   330
      Left            =   3480
      TabIndex        =   14
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   2640
      Width           =   5052
      _Version        =   1572864
      _ExtentX        =   8911
      _ExtentY        =   582
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
      Enabled         =   0   'False
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkCargos 
      Height          =   252
      Left            =   8640
      TabIndex        =   15
      Top             =   3000
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCargoId 
      Height          =   330
      Left            =   1440
      TabIndex        =   16
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   3000
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
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
      Enabled         =   0   'False
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCargoDesc 
      Height          =   330
      Left            =   3480
      TabIndex        =   17
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   3000
      Width           =   5052
      _Version        =   1572864
      _ExtentX        =   8911
      _ExtentY        =   582
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
      Enabled         =   0   'False
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkFecTodas 
      Height          =   372
      Left            =   7800
      TabIndex        =   20
      Top             =   5160
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Todas las fechas"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   7440
      TabIndex        =   21
      Top             =   5640
      Width           =   2292
      _Version        =   1572864
      _ExtentX        =   4048
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   8400
      TabIndex        =   22
      Top             =   4440
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   8400
      TabIndex        =   23
      Top             =   4800
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.PushButton cmdReporte 
      Height          =   612
      Left            =   8160
      TabIndex        =   24
      Top             =   6240
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Reporte"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCxC_Reportes.frx":0000
   End
   Begin XtremeSuiteControls.Label Label8 
      Height          =   132
      Left            =   120
      TabIndex        =   30
      Top             =   3600
      Width           =   2532
      _Version        =   1572864
      _ExtentX        =   4466
      _ExtentY        =   233
      _StockProps     =   79
      Caption         =   "Informes:"
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
   End
   Begin XtremeShortcutBar.ShortcutCaption lblReporte 
      Height          =   372
      Left            =   0
      TabIndex        =   29
      Top             =   3480
      Width           =   10212
      _Version        =   1572864
      _ExtentX        =   18013
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Seleccione un informe!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   2
   End
   Begin VB.Label lblx02 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas"
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
      Left            =   5520
      TabIndex        =   28
      Top             =   4440
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      Left            =   5520
      TabIndex        =   27
      Top             =   5640
      Width           =   1212
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Inicio"
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
      Index           =   0
      Left            =   7440
      TabIndex        =   26
      Top             =   4440
      Width           =   972
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Corte"
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
      Index           =   1
      Left            =   7440
      TabIndex        =   25
      Top             =   4800
      Width           =   972
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
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
      Left            =   0
      TabIndex        =   18
      Top             =   3000
      Width           =   1212
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto"
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
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes del Módulo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   2
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   6852
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pagador"
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
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   1212
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nombre"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   5052
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Identificación"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Width           =   2052
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmCxC_Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSubTitulo As String


Private Sub chkConcepto_Click()
If chkConcepto.Value = vbChecked Then
 txtConceptoId.Enabled = False
Else
 txtConceptoId.Enabled = True
 txtConceptoId.SetFocus
End If

txtConceptoDesc.Enabled = txtConceptoId.Enabled
End Sub

Private Sub chkFecTodas_Click()
If chkFecTodas.Value = xtpChecked Then
 dtpInicio.Enabled = False
Else
 dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub


Private Sub chkPagador_Click()
If chkPagador.Value = vbChecked Then
 txtPagadorId.Enabled = False
Else
 txtPagadorId.Enabled = True
 txtPagadorId.SetFocus
End If

txtPagadorNombre.Enabled = txtPagadorId.Enabled
End Sub

Private Sub chkTodos_Click()
If chkTodos.Value = vbChecked Then
 txtClienteId.Enabled = False
Else
 txtClienteId.Enabled = True
 txtClienteId.SetFocus
End If

txtClienteNombre.Enabled = txtClienteId.Enabled

End Sub


Private Sub sbReport_Main()
Dim vSQL As String, vTipo As String, vDisplay As String
Dim vConcepto As String, vClienteId As String, vPagadorId  As String
         

On Error GoTo vError

Me.MousePointer = vbHourglass
         
If chkTodos.Value = xtpChecked Then
  vClienteId = "-1x"
Else
  vClienteId = Trim(txtClienteId.Text)
End If

If chkPagador.Value = xtpChecked Then
  vPagadorId = "-1x"
Else
  vPagadorId = Trim(txtPagadorId.Text)
End If

If chkConcepto.Value = xtpChecked Then
  vConcepto = "-1x"
Else
  vConcepto = Trim(txtConceptoId.Text)
End If


Select Case cboTipo.Text
  Case "Detallado"
    vTipo = "Det"
    vDisplay = "D"
  Case "Resumido"
    vTipo = "Rsm"
    vDisplay = "R"
  Case "Contable"
    vTipo = "Contable"
    vDisplay = "C"
End Select


With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowGroupTree = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes de Cuentas por Cobrar"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 

'lsw.ListItems.Add , "x00", "Listado de Personas"
'lsw.ListItems.Add , "x01", "Auxiliar: General"
'lsw.ListItems.Add , "x01-1", Space(10) & "Auxiliar..Concepto"
'lsw.ListItems.Add , "x01-2", Space(10) & "Auxiliar..Cliente"
'lsw.ListItems.Add , "x01-3", Space(10) & "Auxiliar..Pagador"
'lsw.ListItems.Add , "x01-4", Space(10) & "Auxiliar..Facturas Descontadas"
'lsw.ListItems.Add , "x01-5", Space(10) & "Auxiliar..Facturas Adelantadas"
'lsw.ListItems.Add , "x01-6", Space(10) & "Auxiliar..Antiguedad de Saldos"
'lsw.ListItems.Add , "x02", "Facturas Descontadas"
'lsw.ListItems.Add , "x02-1", Space(10) & "..Facturas En Tramite"
'lsw.ListItems.Add , "x02-2", Space(10) & "..Facturas Canceladas"
'lsw.ListItems.Add , "x03", "Adelantos Registrados"
'lsw.ListItems.Add , "x03-1", Space(10) & "..Adelantos Pendientes del Descuento"
'lsw.ListItems.Add , "x04", "Movimientos Realizados"
'lsw.ListItems.Add , "x05", "Informe de Colocación"



 Select Case lblReporte.Tag
    Case "x00" 'Listado de Clientes
         vSQL = fxSQL(0)
         .Formulas(3) = "fxTitulo = 'LISTADO DE CLIENTES'"
         .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_Clientes.rpt")
         .SelectionFormula = vSQL
    
    
    Case "x01" 'Auxiliar: General
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_Auxiliar_General_" & vTipo & ".rpt")
         
    Case "x01-1" 'Auxiliar: Concepto
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_Auxiliar_Concepto_" & vTipo & ".rpt")
         
    Case "x01-2" 'Auxiliar: Clientes
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_Auxiliar_Cliente_" & vTipo & ".rpt")
         
    Case "x01-3" 'Auxiliar: Pagador
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_Auxiliar_Pagador_" & vTipo & ".rpt")
         
    Case "x01-4" 'Auxiliar: Facturas Descontadas
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_Auxiliar_Factura_Descuento_" & vTipo & ".rpt")
         
    Case "x01-5" 'Auxiliar: Facturas Adelantadas
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_Auxiliar_Factura_Adelanto_" & vTipo & ".rpt")
         
    Case "x01-6" 'Auxiliar: Antiguedad de Saldos
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_Auxiliar_Antiguedad_Saldos_" & vTipo & ".rpt")
    
 End Select

'Aplica Filtros Reportes de Auxiliar
If Mid(lblReporte.Tag, 1, 3) = "x01" Then
    .Formulas(3) = "fxTitulo = 'CxC: " & UCase(lblReporte.Caption) & "'"
    .Formulas(4) = "fxSubTitulo = 'CORTE: " & Format(dtpCorte.Value, "dd/mm/yyyy") & " ¦ INFORME: " & UCase(cboTipo.Text) & "'"
    .SelectionFormula = ""

    .StoredProcParam(0) = vConcepto
    .StoredProcParam(1) = vClienteId
    .StoredProcParam(2) = vPagadorId
    .StoredProcParam(3) = glogon.Usuario
    .StoredProcParam(4) = Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59"
    .StoredProcParam(5) = vDisplay

'.StoredProcParam(4) = "@Corte;" & Format(dtpCorte.Value, "dd/MM/yyyy") & " 11:59:59PM" ' Format(dtpCorte.Value, "dd,MM,yyyy") '& " 23,59,59"

End If



.Action = 1


End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub sbReport_Cargos()
Dim vSQL As String, vTipo As String, vDisplay As String
Dim vConcepto As String, vClienteId As String, vPagadorId  As String
         

On Error GoTo vError

Me.MousePointer = vbHourglass
         
If chkTodos.Value = xtpChecked Then
  vClienteId = "-1x"
Else
  vClienteId = Trim(txtClienteId.Text)
End If

If chkPagador.Value = xtpChecked Then
  vPagadorId = "-1x"
Else
  vPagadorId = Trim(txtPagadorId.Text)
End If

If chkConcepto.Value = xtpChecked Then
  vConcepto = "-1x"
Else
  vConcepto = Trim(txtConceptoId.Text)
End If


Select Case cboTipo.Text
  Case "Detallado"
    vTipo = "Det"
    vDisplay = "D"
  Case "Resumido"
    vTipo = "Rsm"
    vDisplay = "R"
  Case "Contable"
    vTipo = "Contable"
    vDisplay = "C"
End Select


With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowGroupTree = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes de Cuentas por Cobrar"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "fxTitulo = 'CxC: " & UCase(lblReporte.Caption) & "'"
 .Formulas(4) = "fxSubTitulo = 'CORTE: " & Format(dtpInicio.Value, "dd/mm/yyyy") & "..." & Format(dtpCorte.Value, "dd/mm/yyyy") & " ¦ INFORME: " & UCase(cboTipo.Text) & "'"
 
 vSQL = fxSQL(3)

.Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
.SelectionFormula = vSQL

 Select Case lblReporte.Tag
    Case "c01" 'Cargos Registados
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_Cargos_" & vTipo & ".rpt")
    
    Case "c02" 'Cargos Registados: Clientes
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_Cargos_Cliente_" & vTipo & ".rpt")
    
    Case "c03" 'Cargos Registados: Pagador
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_Cargos_Pagador_" & vTipo & ".rpt")
    
    Case "c04" 'Cargos Registados: Concepto
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_Cargos_Concepto_" & vTipo & ".rpt")
    
    Case Else
    
 End Select




.Action = 1


End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cmdReporte_Click()
 
Select Case Mid(lblReporte.Tag, 1, 1)
  Case "x" 'Reportes de CxC
    Call sbReport_Main
  Case "c" 'Reportes de Cargos Registrados
    Call sbReport_Cargos
End Select

End Sub

Private Function fxSQL(i As Integer) As String
Dim vSQL As String

vSQL = ""
vSubTitulo = ""

Select Case i
  Case 0 'Listado de Clientes
     
     If chkTodos.Value = vbUnchecked Then
       vSQL = "{vCxC_Cuentas_Consulta.cedula} = '" & txtClienteId & "'"
       vSubTitulo = "Cliente: " & UCase(txtClienteNombre)
     Else
       vSubTitulo = "TODOS LOS Clientes"
     End If
          
     
     
     
     fxSQL = vSQL
     Exit Function
     
  Case 1 'Antiguedad de Saldos / Programacion
     If chkTodos.Value = vbUnchecked Then
       vSQL = "{vCxC_Cuentas_Consulta.cedula} = '" & txtClienteId & "'"
       vSubTitulo = "Cliente: " & UCase(txtClienteNombre)
     Else
       vSubTitulo = "TODOS LOS Clientes"
     End If
          
'     Select Case Mid(cbo.Text, 1, 1)
'        Case "A"
'           If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
'           vSQL = vSQL & "{vCxC_Cuentas_Consulta.ESTADO} = 'A'"
'        Case "I"
'           If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
'           vSQL = vSQL & "{vCxC_Cuentas_Consulta.ESTADO} = 'I'"
'        Case "F"
'           If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
'           vSQL = vSQL & "isnull({vCxC_Cuentas_Consulta.FUSION}) = TRUE"
'     End Select
'
'     vSubTitulo = vSubTitulo & " [ESTADO : " & UCase(cbo.Text) & "]"
'
     fxSQL = vSQL
     Exit Function

  Case 2 'Pagos Realizados
     
     vSQL = "ISNULL({CxC_PAGOPROV.TESORERIA}) = FALSE"
     
     If chkTodos.Value = vbUnchecked Then
       If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
       vSQL = "{vCxC_Cuentas_Consulta.cedula} = '" & txtClienteId & "'"
       vSubTitulo = "Cliente: " & UCase(txtClienteNombre)
     Else
       vSubTitulo = "TODOS LOS Clientes"
     End If
          
'     Select Case Mid(cbo.Text, 1, 1)
'        Case "A"
'           If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
'           vSQL = vSQL & "{vCxC_Cuentas_Consulta.ESTADO} = 'A'"
'        Case "I"
'           If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
'           vSQL = vSQL & "{vCxC_Cuentas_Consulta.ESTADO} = 'I'"
'        Case "F"
'           If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
'           vSQL = vSQL & "isnull({vCxC_Cuentas_Consulta.FUSION}) = TRUE"
'     End Select
'
'     vSubTitulo = vSubTitulo & " [ESTADO : " & UCase(cbo.Text) & "]"
     
     If chkFecTodas.Value = vbUnchecked Then
         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
         vSQL = vSQL & "{CxC_PAGOPROV.FECHA_TRASLADA}" & fxFechaReportes
         vSubTitulo = vSubTitulo & " [Pagos Entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy") & "]"
     Else
         vSubTitulo = vSubTitulo & " [TODAS LAS FECHAS]"
     End If
     
     fxSQL = vSQL
     Exit Function

    '------------------------------------------------------------------------------------------------------------
    Case 3 'Cargos Registrados
     
     If chkTodos.Value = vbUnchecked Then
       If Len(vSQL) > 0 Then
            vSQL = vSQL & " AND "
            vSubTitulo = vSubTitulo & " ¦ "
       End If
       vSQL = vSQL & "{vCxC_Cargos_Registrados.CEDULA} = '" & txtClienteId & "'"
       vSubTitulo = vSubTitulo & "Cliente Id: " & txtClienteId.Text
     Else
       vSubTitulo = vSubTitulo & "Cliente Id: [T]"
     End If


     If chkPagador.Value = vbUnchecked Then
       If Len(vSQL) > 0 Then
            vSQL = vSQL & " AND "
            vSubTitulo = vSubTitulo & " ¦ "
       End If
       vSQL = vSQL & "{vCxC_Cargos_Registrados.CEDULA_PAGADOR} = '" & txtPagadorId.Text & "'"
       vSubTitulo = vSubTitulo & " ¦ Pagador Id: " & txtPagadorId.Text
     Else
       vSubTitulo = vSubTitulo & " ¦ Pagador Id: [T]"
     End If

     If chkConcepto.Value = vbUnchecked Then
       If Len(vSQL) > 0 Then
            vSQL = vSQL & " AND "
            vSubTitulo = vSubTitulo & " ¦ "
       End If
       vSQL = vSQL & "{vCxC_Cargos_Registrados.COD_CONCEPTO} = '" & txtConceptoId.Text & "'"
       vSubTitulo = vSubTitulo & " ¦ Concepto Id: " & txtConceptoId.Text
     Else
       vSubTitulo = vSubTitulo & " ¦ Concepto Id: [T]"
     End If

     If chkFecTodas.Value = xtpUnchecked Then
       If Len(vSQL) > 0 Then
            vSQL = vSQL & " AND "
            vSubTitulo = vSubTitulo & " ¦ "
       End If
       vSQL = vSQL & "{vCxC_Cargos_Registrados.ACTIVA_FECHA}" _
               & " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ") to date(" _
               & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
       
       
       vSubTitulo = vSubTitulo & " Corte: " & Format(dtpInicio.Value, "dd/mm/yyyy") & "..." & Format(dtpCorte.Value, "dd/mm/yyyy")
     Else
       vSubTitulo = vSubTitulo & " ¦ Corte: [T]"
     End If

     vSubTitulo = vSubTitulo & " ¦ INFORME: " & UCase(cboTipo.Text)

     fxSQL = vSQL
     Exit Function
     
End Select

End Function


Private Function fxFechaReportes(Optional vTipo As Integer = 0) As String

fxFechaReportes = " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ")" _
                & " to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"

End Function


Private Sub Form_Load()

vModulo = 31

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

cboTipo.AddItem "Detallado"
cboTipo.AddItem "Resumido"
cboTipo.Text = "Detallado"


lsw.ColumnHeaders.Add , , "", 4500
lsw.HideColumnHeaders = True

lsw.ListItems.Clear
lsw.ListItems.Add , "x00", "Listado de Personas"
lsw.ListItems.Add , "x01", "Auxiliar: General"
lsw.ListItems.Add , "x01-1", Space(10) & "Auxiliar..Concepto"
lsw.ListItems.Add , "x01-2", Space(10) & "Auxiliar..Cliente"
lsw.ListItems.Add , "x01-3", Space(10) & "Auxiliar..Pagador"
lsw.ListItems.Add , "x01-4", Space(10) & "Auxiliar..Facturas Descontadas"
lsw.ListItems.Add , "x01-5", Space(10) & "Auxiliar..Facturas Adelantadas"
lsw.ListItems.Add , "x01-6", Space(10) & "Auxiliar..Antiguedad de Saldos"
lsw.ListItems.Add , "x02", "Facturas Descontadas"
lsw.ListItems.Add , "x02-1", Space(10) & "..Facturas En Tramite"
lsw.ListItems.Add , "x02-2", Space(10) & "..Facturas Canceladas"
lsw.ListItems.Add , "x03", "Adelantos Registrados"
lsw.ListItems.Add , "x03-1", Space(10) & "..Adelantos Pendientes del Descuento"
lsw.ListItems.Add , "x04", "Movimientos Realizados"
lsw.ListItems.Add , "x05", "Informe de Colocación"

lsw.ListItems.Add , "c01", "Cargos Registrados: General"
lsw.ListItems.Add , "c02", Space(10) & "Cargos..Cliente"
lsw.ListItems.Add , "c03", Space(10) & "Cargos..Pagador"
lsw.ListItems.Add , "c04", Space(10) & "Cargos..Concepto"

lblReporte.Tag = "x00"
lblReporte.Caption = "Listado de Personas"

Call chkFecTodas_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub






Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
cboTipo.Enabled = False
dtpInicio.Enabled = False
dtpCorte.Enabled = False
chkFecTodas.Enabled = False

lblReporte.Tag = Item.Key
lblReporte.Caption = Trim(Replace(Item.Text, "..", " "))


cboTipo.Clear
cboTipo.AddItem "Detallado"
cboTipo.AddItem "Resumido"
cboTipo.Text = "Detallado"

Select Case Mid(Item.Key, 1, 3)
  Case "x01"
    cboTipo.AddItem "Contable"
    dtpCorte.Enabled = True
    cboTipo.Enabled = True
  Case Else
    cboTipo.Enabled = True
    dtpInicio.Enabled = True
    dtpCorte.Enabled = True
    chkFecTodas.Enabled = True
    chkFecTodas_Click
End Select


End Sub



Private Sub txtCargoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And cmdReporte.Enabled Then cmdReporte.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_cargo,descripcion from CXC_CARGOS"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCargoId.Text = gBusquedas.Resultado
  txtCargoDesc.Text = gBusquedas.Resultado2
  
  If cmdReporte.Enabled Then cmdReporte.SetFocus
End If

End Sub

Private Sub txtCargoId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And cmdReporte.Enabled Then cmdReporte.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_cargo"
  gBusquedas.Orden = "cod_cargo"
  gBusquedas.Consulta = "select cod_cargo,descripcion from CXC_CARGOS"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCargoId.Text = gBusquedas.Resultado
  txtCargoDesc.Text = gBusquedas.Resultado2
  If cmdReporte.Enabled Then cmdReporte.SetFocus
End If
End Sub

Private Sub txtClienteId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtClienteNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from CxC_Personas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtClienteId = gBusquedas.Resultado
  txtClienteNombre = gBusquedas.Resultado2
  cmdReporte.SetFocus
End If

End Sub

Private Sub txtClienteId_LostFocus()
'txtClienteNombre = fxSIFCCodigos("D", txtClienteId, "Clientes")
End Sub




Private Sub txtConceptoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And cmdReporte.Enabled Then cmdReporte.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_concepto,descripcion from CxC_Conceptos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtConceptoId.Text = gBusquedas.Resultado
  txtConceptoDesc.Text = gBusquedas.Resultado2
  
  If cmdReporte.Enabled Then cmdReporte.SetFocus
End If
End Sub

Private Sub txtConceptoId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And cmdReporte.Enabled Then cmdReporte.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_concepto"
  gBusquedas.Orden = "cod_concepto"
  gBusquedas.Consulta = "select cod_concepto,descripcion from CxC_Conceptos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtConceptoId.Text = gBusquedas.Resultado
  txtConceptoDesc.Text = gBusquedas.Resultado2
  
  If cmdReporte.Enabled Then cmdReporte.SetFocus
End If
End Sub

Private Sub txtClienteNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdReporte.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from CxC_Personas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtClienteId = gBusquedas.Resultado
  txtClienteNombre = gBusquedas.Resultado2
  cmdReporte.SetFocus
End If

End Sub


Private Sub txtPagadorId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPagadorNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from CxC_Personas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtPagadorId.Text = gBusquedas.Resultado
  txtPagadorNombre.Text = gBusquedas.Resultado2
  cmdReporte.SetFocus
End If
End Sub

Private Sub txtPagadorNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdReporte.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from CxC_Personas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtPagadorId.Text = gBusquedas.Resultado
  txtPagadorNombre.Text = gBusquedas.Resultado2
  cmdReporte.SetFocus
End If
End Sub
