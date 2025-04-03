VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAH_AnulaAhorros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Patrimonio: Anulación de Movimientos"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9645
   HelpContextID   =   2001
   Icon            =   "frmAH_AnulaAhorros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   1932
      Left            =   0
      TabIndex        =   25
      Top             =   2640
      Width           =   9612
      _Version        =   1441793
      _ExtentX        =   16954
      _ExtentY        =   3408
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbAnulacion 
      Height          =   2172
      Left            =   120
      TabIndex        =   19
      Top             =   4680
      Width           =   9372
      _Version        =   1441793
      _ExtentX        =   16531
      _ExtentY        =   3831
      _StockProps     =   79
      Caption         =   "Datos de la anulación:"
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdAnulaAhorro 
         Height          =   612
         Left            =   7800
         TabIndex        =   20
         Top             =   1320
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Anular"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_AnulaAhorros.frx":08CA
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   5880
         TabIndex        =   21
         Top             =   720
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   550
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   5880
         TabIndex        =   22
         Top             =   360
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.ComboBox cboAccion 
         Height          =   312
         Left            =   1800
         TabIndex        =   28
         Top             =   720
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   1800
         TabIndex        =   30
         Top             =   1200
         Width           =   5772
         _Version        =   1441793
         _ExtentX        =   10181
         _ExtentY        =   1397
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
      Begin XtremeSuiteControls.ComboBox cboTipoAnulacion 
         Height          =   312
         Left            =   1800
         TabIndex        =   32
         Top             =   360
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2990
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Anulación..:"
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
         Left            =   -600
         TabIndex        =   31
         Top             =   360
         Width           =   2292
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Notas..:"
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
         Left            =   -600
         TabIndex        =   29
         Top             =   1200
         Width           =   2292
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Anula y Procesar..:"
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
         Left            =   -600
         TabIndex        =   26
         Top             =   720
         Width           =   2292
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Monto del aporte ..:"
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
         Left            =   3240
         TabIndex        =   24
         Top             =   720
         Width           =   2412
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rubro de Patrimonio..:"
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
         Left            =   3360
         TabIndex        =   23
         Top             =   360
         Width           =   2292
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9000
      Top             =   240
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
      _ExtentY        =   550
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   550
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1572
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   9612
      _Version        =   1441793
      _ExtentX        =   16954
      _ExtentY        =   2773
      _StockProps     =   79
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
      Appearance      =   16
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtObrero 
         Height          =   312
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPatronal 
         Height          =   312
         Left            =   1680
         TabIndex        =   5
         Top             =   840
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCapitalizacion 
         Height          =   312
         Left            =   6480
         TabIndex        =   6
         Top             =   480
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCustodia 
         Height          =   312
         Left            =   6480
         TabIndex        =   7
         Top             =   840
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotal 
         Height          =   312
         Left            =   6480
         TabIndex        =   8
         Top             =   1200
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   312
         Left            =   8280
         TabIndex        =   27
         Top             =   1200
         Width           =   612
         _Version        =   1441793
         _ExtentX        =   1080
         _ExtentY        =   550
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   372
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   9612
         _Version        =   1441793
         _ExtentX        =   16954
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Estado actual del Patrimonio"
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
         Alignment       =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Aporte Patronal"
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
         Index           =   5
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Aporte Obrero"
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
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Total"
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
         Left            =   5040
         TabIndex        =   15
         Top             =   1200
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Ap.Pat/Custodia"
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
         Left            =   5040
         TabIndex        =   14
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Capitalización"
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
         Left            =   5040
         TabIndex        =   13
         Top             =   480
         Width           =   1332
      End
      Begin XtremeSuiteControls.Label lblFechaObrero 
         Height          =   312
         Left            =   3600
         TabIndex        =   12
         Top             =   480
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblFechaPatronal 
         Height          =   312
         Left            =   3600
         TabIndex        =   11
         Top             =   840
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblCapitalizado 
         Height          =   312
         Left            =   8400
         TabIndex        =   10
         Top             =   480
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblFechaCustodia 
         Height          =   312
         Left            =   8400
         TabIndex        =   9
         Top             =   840
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1452
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmAH_AnulaAhorros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim pCharRelleno As String


Public Sub sbConsulta(pCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
 
On Error GoTo vError

strSQL = "select *" _
       & ",(select COD_DIVISA from vSys_Divisas where DIVISA_LOCAL = 1) AS 'COD_DIVISA'" _
       & " from vPAT_Consolidado" _
       & " where cedula = '" & pCedula & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  txtNombre.Text = rs!Nombre
  txtObrero.Text = Format(rs!Obrero, "Standard")
  txtPatronal.Text = Format(rs!Patronal, "Standard")
  txtCustodia.Text = Format(rs!Custodia, "Standard")
  txtCapitalizacion.Text = Format(rs!capitaliza, "Standard")
  txtDivisa.Text = Trim(rs!cod_Divisa)
  
  txtTotal.Text = Format(rs!Obrero + rs!Patronal + rs!Custodia + rs!capitaliza, "Standard")
  
  lblFechaObrero.Caption = IIf(IsNull(rs!fecAhorro), "", Format(rs!fecAhorro, "dd/mm/yyyy"))
  lblFechaPatronal.Caption = IIf(IsNull(rs!fecaporte), "", Format(rs!fecaporte, "dd/mm/yyyy"))
  lblFechaCustodia.Caption = IIf(IsNull(rs!fecCustodia), "", Format(rs!fecCustodia, "dd/mm/yyyy"))
  lblCapitalizado.Caption = IIf(IsNull(rs!fecCapitaliza), "", Format(rs!fecCapitaliza, "dd/mm/yyyy"))
  
 
Else
    txtNombre.Text = ""
    
    txtObrero.Text = 0
    txtPatronal.Text = 0
    txtCustodia.Text = 0
    txtCapitalizacion.Text = 0
    txtTotal.Text = 0
    
    
    lblFechaObrero.Caption = ""
    lblFechaPatronal.Caption = ""
    lblFechaCustodia.Caption = ""
    lblCapitalizado.Caption = ""
End If
rs.Close

Call cboTipoAnulacion_Click

Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cboTipo_Click()
If vPaso Then Exit Sub

Call cboTipoAnulacion_Click

End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub cboTipoAnulacion_Click()
    
If vPaso Then Exit Sub
    
    Select Case cboTipoAnulacion
     Case "Monto"
        txtMonto = "0"
        txtMonto.Locked = False
        
        lsw.ListItems.Clear
        lsw.Enabled = False
    
     Case "Movimiento"
     
        txtMonto.Text = "0"
        txtMonto.Locked = True
        
        lsw.Enabled = True
        Call sbMovimientos_Load
    
    End Select

End Sub


Private Sub cboTipoAnulacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
End Sub

Private Function fxVerificaAnulacion() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Vericar Que todos los datos de la anulación sean válidos
'REFERENCIAS:   Ninguna
'OBSERVACIONES: Ninguna
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim vValida As Boolean

On Error GoTo vError


If Not IsNumeric(txtMonto.Text) Then
    fxVerificaAnulacion = False
    Exit Function
End If

If CCur(txtMonto.Text) <= 0 Then
    fxVerificaAnulacion = False
    Exit Function
End If

vValida = True

Select Case cboTipo.ItemData(cboTipo.ListIndex)
   Case "C", "CAP"  'Capitalizacion
     If CCur(txtMonto.Text) > CCur(txtCapitalizacion.Text) Then vValida = False
   Case "O", "OBR" 'Obrero
     If CCur(txtMonto.Text) > CCur(txtObrero.Text) Then vValida = False
   Case "P", "PAT" 'Patronal
     If CCur(txtMonto.Text) > CCur(txtPatronal.Text) Then vValida = False
   Case "X", "CST" 'Custodia
     If CCur(txtMonto.Text) > CCur(txtCustodia.Text) Then vValida = False
 End Select

fxVerificaAnulacion = vValida

Exit Function

vError:
   fxVerificaAnulacion = False
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub sbDocumento(vTipoDoc As String, vNumDoc As String, vCuenta As String, vFP_SF As String)
Dim strSQL As String, rs As New ADODB.Recordset

Dim strLinea(10) As String
Dim pDivisa As String, pTipoCambio As Currency
Dim vColCuenta As String, vColAporte As String

Dim pUnidad As String, pCentroCosto As String, pOficina As String

On Error GoTo vError

pDivisa = txtDivisa.Text
pTipoCambio = fxCajasTipoCambio(pDivisa)

pUnidad = GLOBALES.gOficinaUnidad
pOficina = GLOBALES.gOficinaTitular
pCentroCosto = ""

Call sbSIFCleanTxtInject(txtNotas)

Select Case cboTipo.ItemData(cboTipo.ListIndex)
  Case "P" 'Aporte Patronal
     vColAporte = "Aporte"
     vColCuenta = "Cta_Patronal"
  
  Case "X" 'Aporte en Custodia"
     vColAporte = "Custodia"
     vColCuenta = "cta_custodia"
 
  Case "O" 'Aporte Obrero"
     vColAporte = "ahorro"
     vColCuenta = "cta_obrero"
  
  Case "C" 'Capitalización"
     vColAporte = "capitaliza"
     vColCuenta = "cta_capitaliza"
End Select


strSQL = "select P." & vColAporte & " as 'Aporte'" _
       & ", (select " & vColCuenta & " as Cuenta from par_afah Where Cod_Divisa = P.Cod_Divisa) as 'Cuenta'" _
       & " from vPAT_Consulta_Integrada P" _
       & " where P.cedula = '" & txtCedula & "'"

Call OpenRecordSet(rs, strSQL)


strLinea(1) = "Plan            : " & cboTipo.Text
strLinea(2) = "                  "
strLinea(3) = "Saldo Anterior  : " & SIFGlobal.fxStringRelleno(Format(rs!Aporte, "Standard"), "I", pCharRelleno, 20)
strLinea(4) = "Monto Anulación : " & SIFGlobal.fxStringRelleno(txtMonto.Text, "I", pCharRelleno, 20)
strLinea(5) = "Saldo Actual    : " & SIFGlobal.fxStringRelleno(Format(rs!Aporte - CCur(txtMonto), "Standard"), "I", pCharRelleno, 20)
strLinea(6) = "                  "
strLinea(7) = "Divisa          : " & pDivisa
strLinea(8) = ""
strLinea(9) = "Usuario         : " & glogon.Usuario
strLinea(10) = "Acción: " & cboAccion.Text

    
   strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
             & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
             & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento,cod_caja)" _
             & " values('" & vNumDoc & "','" & vTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" _
             & txtCedula.Text & "','" & txtNombre & "','PAT002'," & CCur(txtMonto) * -1 & ",'P','" _
             & txtCedula.Text & "','','','" & pOficina & "', " _
             & "'" & strLinea(1) & "','" & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
             & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
             & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
             & txtNotas.Text & "','" & vAseDocDeposito & "','')"
    
    'ASIENTO
    If CCur(txtMonto) > 0 Then
        strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & CCur(txtMonto) * fxSys_Tipo_Cambio_Apl(pTipoCambio) & "" _
                & ",'D','" & pDivisa & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & pUnidad & "'," _
                & " '','" & rs!Cuenta & "','Pat:" & cboTipo.ItemData(cboTipo.ListIndex) & "','" & txtCedula.Text & "','" & vAseDocDeposito & "'"
        
        strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & CCur(txtMonto) * fxSys_Tipo_Cambio_Apl(pTipoCambio) & "" _
                & ",'C','" & pDivisa & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & pUnidad & "'," _
                & " '','" & vCuenta & "','Pat:" & cboTipo.ItemData(cboTipo.ListIndex) & "','" & txtCedula.Text & "','" & vAseDocDeposito & "'"

    End If
    
    'Registrar
    Call ConectionExecute(strSQL)
    
    
    Select Case cboAccion.ItemData(cboAccion.ListIndex)
        Case "C" 'Cuenta
        Case "S" 'Saldo a Favor
             strSQL = "exec spCajas_SaldoFavor_Registra '" & vFP_SF & "','" & vTipoDoc & "-" & vNumDoc & "'," & CCur(txtMonto.Text) _
                    & ",'" & txtCedula.Text & "','" & txtNombre.Text & "','" & glogon.Usuario & "','" & txtDivisa.Text & "'"
           
             Call OpenRecordSet(rs, strSQL)
             
             'Insertar Format de Pago
            strSQL = "exec spPAT_Anulacion_Saldo_Favor '" & vTipoDoc & "','" & vNumDoc & "','" & glogon.Usuario _
                    & "','" & vFP_SF & "','" & txtDivisa.Text & "'," & CCur(txtMonto.Text) _
                    & ",'" & pUnidad & "','" & vCuenta & "','" & vTipoDoc & "-" & vNumDoc _
                    & "'," & rs!SF_ID
            
            Call ConectionExecute(strSQL)
    End Select
   


rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub cmdAnulaAhorro_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCuenta As String, curMonto As Currency, vFP_SF As String
Dim vTipoDoc As String, vNumDoc As String, vConcepto As String

If Not fxVerificaAnulacion Then
    MsgBox "Verifique la Información!", vbExclamation
    Exit Sub
End If

vCuenta = ""
vFP_SF = ""

 Select Case cboAccion.ItemData(cboAccion.ListIndex)
  Case "C" 'Cuenta Contable
       vCuenta = Trim(fxDocumentoCuenta("ND"))
       
       If vAseDocValido = False Then
         MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
              & " válida para esta operación...", vbCritical
         Exit Sub
       End If
   
       txtNotas.Text = vAseDocDetalle
       
   Case "S" 'Saldo a favor en Cajas
       strSQL = "select COD_FORMA_PAGO, COD_CUENTA " _
              & " From SIF_FORMAS_PAGO" _
              & " where TIPO = 'S' and Activa = 1"
      Call OpenRecordSet(rs, strSQL)
        vCuenta = Trim(rs!cod_cuenta)
        vFP_SF = Trim(rs!Cod_Forma_Pago)
      rs.Close
 End Select

vTipoDoc = "ND"
vNumDoc = fxDocumentoConsecutivo(vTipoDoc)

Call sbDocumento(vTipoDoc, vNumDoc, vCuenta, vFP_SF)
  
strSQL = "exec spPAT_Anulacion '" & txtCedula.Text & "','" & cboTipo.ItemData(cboTipo.ListIndex) & "'," & CCur(txtMonto.Text) _
       & ",'" & vTipoDoc & "','" & vNumDoc & "','" & glogon.Usuario _
       & "','','" & vConcepto & "',0"
Call ConectionExecute(strSQL)
 
 Call Bitacora("Anula", cboTipo.Text & " Anula: " & curMonto & ", Id:" & txtCedula.Text)
 
 Call sbImprimeRecibo(vNumDoc, vTipoDoc)
 
 MsgBox "Anulación Realizada ... Con Nota Debito: " & vNumDoc, vbInformation
 
 txtMonto.Text = "0"
 Call sbConsulta(txtCedula.Text)
 
End Sub

Private Sub Form_Load()

vModulo = 2

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

txtCedula.Text = GLOBALES.gCedulaActual

pCharRelleno = "*"

vPaso = True
    cboTipoAnulacion.Clear
    cboTipoAnulacion.AddItem "Monto"
    cboTipoAnulacion.AddItem "Movimiento"
    cboTipoAnulacion.Text = "Monto"


    cboTipo.AddItem "Aporte Obrero"
    cboTipo.ItemData(cboTipo.ListCount - 1) = "O"
     
    cboTipo.AddItem "Aporte Patronal"
    cboTipo.ItemData(cboTipo.ListCount - 1) = "P"
    
    cboTipo.AddItem "Capitalización"
    cboTipo.ItemData(cboTipo.ListCount - 1) = "C"

    cboTipo.AddItem "Aporte en Custodia"
    cboTipo.ItemData(cboTipo.ListCount - 1) = "X"

    cboTipo.Text = "Aporte Obrero"
    
    
    cboAccion.Clear
    cboAccion.AddItem "Cuenta Contable"
    cboAccion.ItemData(cboAccion.ListCount - 1) = "C"
    cboAccion.AddItem "Saldo a Favor"
    cboAccion.ItemData(cboAccion.ListCount - 1) = "S"
    cboAccion.Text = "Cuenta Contable"
    
vPaso = False



Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

Dim curMonto As Currency

On Error GoTo vError

curMonto = CCur(txtMonto.Text)

If Item.Checked Then
    curMonto = curMonto + CCur(Item.SubItems(3))
Else
    curMonto = curMonto - CCur(Item.SubItems(3))
End If

txtMonto.Text = Format(curMonto, "Standard")

Exit Sub

vError:

End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError
 txtMonto.Text = CCur(txtMonto)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And cmdAnulaAhorro.Enabled Then cmdAnulaAhorro.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
  txtMonto = Format(CCur(txtMonto.Text), "Standard")
vError:
End Sub

Private Sub sbMovimientos_Load()
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError

With lsw
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Fecha", 1750
    .ColumnHeaders.Add , , "Proceso", 1120, vbCenter
    .ColumnHeaders.Add , , "Tipo", 2400
    .ColumnHeaders.Add , , "Monto", 1600, vbRightJustify
    .ColumnHeaders.Add , , "Tipo Doc.", 1000, vbCenter
    .ColumnHeaders.Add , , "Num. Doc.", 2100
    .ColumnHeaders.Add , , "Concepto", 900, vbCenter
    
    strSQL = "select top 24 *" _
           & "From vPAT_Movimientos" _
           & " where cedula = '" & txtCedula.Text & "'" _
           & " and Tipo = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'" _
           & " and Monto > 0" _
           & " order by Fecha desc"
    Call OpenRecordSet(rs, strSQL)
            
    Do While Not rs.EOF
        Set itmX = .ListItems.Add(, , Format(rs!fecha, "dd/mm/yyyy"))
            itmX.SubItems(1) = Format(rs!Fecha_Proceso, "####-##")
            itmX.SubItems(2) = rs!DESCRIPCION
            itmX.SubItems(3) = Format(rs!Monto, "Standard")
            itmX.SubItems(4) = rs!Tcon
            itmX.SubItems(5) = rs!nCon
            itmX.SubItems(6) = rs!cod_Concepto
        rs.MoveNext
    Loop
    rs.Close
End With
    
Me.MousePointer = vbDefault
    
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

Call sbConsulta(txtCedula.Text)
End Sub
