VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCC_ReportesEstudio 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Estudio"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13560
   HelpContextID   =   9011
   Icon            =   "frmCC_ReportesEstudio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   13560
   Begin XtremeSuiteControls.ListView lswRep 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3975
      _Version        =   1441793
      _ExtentX        =   7011
      _ExtentY        =   8916
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
      HideSelection   =   0   'False
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbHistory 
      Height          =   855
      Left            =   120
      TabIndex        =   24
      Top             =   6960
      Visible         =   0   'False
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   1508
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.DateTimePicker dtpHistory 
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   25
         Top             =   120
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHistory 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   26
         Top             =   480
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Corte"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Inicio"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6375
      Left            =   4200
      TabIndex        =   4
      Top             =   1440
      Width           =   9255
      _Version        =   1441793
      _ExtentX        =   16325
      _ExtentY        =   11245
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
      Item(0).Caption =   "Parámetros"
      Item(0).ControlCount=   15
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "txtProyeccion"
      Item(0).Control(2)=   "cboInstitucion"
      Item(0).Control(3)=   "cboCartera"
      Item(0).Control(4)=   "cboEstado"
      Item(0).Control(5)=   "Label3(3)"
      Item(0).Control(6)=   "Label3(0)"
      Item(0).Control(7)=   "Label3(2)"
      Item(0).Control(8)=   "lblProyecta"
      Item(0).Control(9)=   "chkTodos"
      Item(0).Control(10)=   "chkInternas"
      Item(0).Control(11)=   "chkActivas"
      Item(0).Control(12)=   "Label2"
      Item(0).Control(13)=   "dtpProyectaInicio"
      Item(0).Control(14)=   "Label4"
      Item(1).Caption =   "Resultados"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "vGrid"
      Item(1).Control(1)=   "btnResultado"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4215
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   9015
         _Version        =   1441793
         _ExtentX        =   15901
         _ExtentY        =   7435
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
         HideSelection   =   0   'False
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Todos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5295
         Left            =   -69880
         TabIndex        =   5
         Top             =   810
         Visible         =   0   'False
         Width           =   9135
         _Version        =   524288
         _ExtentX        =   16113
         _ExtentY        =   9340
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
         MaxCols         =   494
         ScrollBars      =   2
         SpreadDesigner  =   "frmCC_ReportesEstudio.frx":000C
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboInstitucion 
         Height          =   330
         Left            =   1440
         TabIndex        =   7
         Top             =   5520
         Width           =   5895
         _Version        =   1441793
         _ExtentX        =   10398
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboCartera 
         Height          =   330
         Left            =   1440
         TabIndex        =   8
         Top             =   5160
         Width           =   5895
         _Version        =   1441793
         _ExtentX        =   10398
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   330
         Left            =   1440
         TabIndex        =   9
         Top             =   5880
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkInternas 
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   480
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Lineas Internas?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.CheckBox chkActivas 
         Height          =   255
         Left            =   7320
         TabIndex        =   16
         Top             =   480
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Solo con Saldos?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.PushButton btnResultado 
         Height          =   375
         Left            =   -62080
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Exportar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCC_ReportesEstudio.frx":04C0
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.FlatEdit txtProyeccion 
         Height          =   315
         Left            =   6600
         TabIndex        =   19
         Top             =   5880
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
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
         Text            =   "12"
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpProyectaInicio 
         Height          =   315
         Left            =   7440
         TabIndex        =   22
         Top             =   5880
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   375
         Left            =   7440
         TabIndex        =   23
         Top             =   5400
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Inicio de Proyección"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   480
         Width           =   4455
         _Version        =   1441793
         _ExtentX        =   7858
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Filtras las Líneas de crédito:     "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label lblProyecta 
         Alignment       =   1  'Right Justify
         Caption         =   "Proyección de Líneas en Meses ?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   5880
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Instituciones"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cartera"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   5160
         Width           =   1335
      End
   End
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   7920
      Width           =   13575
      _Version        =   1441793
      _ExtentX        =   23945
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9960
      Top             =   120
   End
   Begin XtremeSuiteControls.PushButton cmdGenera 
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   7080
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Generar"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCC_ReportesEstudio.frx":062A
   End
   Begin XtremeSuiteControls.Label lblEstatus 
      Height          =   375
      Left            =   12120
      TabIndex        =   21
      Top             =   840
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   ".."
      ForeColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption lblReporte 
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   3975
      _Version        =   1441793
      _ExtentX        =   7011
      _ExtentY        =   661
      _StockProps     =   14
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
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Estudio de Auxiliares"
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
      Height          =   615
      Index           =   3
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   7335
   End
   Begin VB.Image imgBanner 
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmCC_ReportesEstudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vHeaders As vGridHeaders
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub sbCreaArchivoTexto(xTabla As String, xArchivo As String)
Dim fn, pRuta As String, strCadena As String
Dim pArchivo As String, i As Integer
'Genera Archivo de Texto

fn = FreeFile


lblEstatus.Caption = "Creando Archivo de Texto"
DoEvents

PrgBar.Value = 1

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados

pRuta = SIFGlobal.DirectorioDeResultados & "\" & xArchivo
pArchivo = Dir(pRuta, vbArchive)

If pArchivo = xArchivo Then 'El archivo existe
  Kill pRuta
End If

Open pRuta For Output As #fn  ' Crea Archivo.

strSQL = "Select * from " & xTabla
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1

strCadena = ""

Print #fn, vbTab & "ProGrX: " & lblReporte.Caption & ", Usuario: " & glogon.Usuario _
           & vbTab & " Fecha : " & Format(fxFechaServidor, "dd/mm/yyyy") & vbTab & "Hora : " _
           & Format(Time, "hh:mm:ss AMPM")

For i = 0 To rs.Fields.Count - 1
 strCadena = strCadena & rs.Fields(i).Name & vbTab
Next

Print #fn, strCadena

Do While Not rs.EOF
 strCadena = ""
 For i = 0 To rs.Fields.Count - 1
   strCadena = strCadena & rs.Fields(i).Value & vbTab
 Next
 Print #fn, strCadena
 PrgBar.Value = PrgBar.Value + 1
 rs.MoveNext
Loop
rs.Close
Close #fn   ' Cierra Archivo.

MsgBox "Se Creó el Siguiente Archivo : " & pRuta, vbInformation

End Sub

Private Sub sbLLenaTabla(vInstitucion As Integer, strCodigo As String, iMeses As Integer, strDescripcion As String, vRetencion As Boolean)
Dim curInt As Currency, curAmortiza As Currency, curSaldo As Currency
Dim curTotalInt As Currency, curTotalAmortiza As Currency, i As Integer
Dim vPlazo As Integer
Dim xSaldo As Currency, x As Integer, xInt(12) As Currency, xAmort(12) As Currency 'Temporales

Dim vInt() As Currency, vAmortiza() As Currency, vSaldo() As Currency 'Mensuales
Dim ySaldo() As Currency, yInt() As Currency, yAmortiza() As Currency 'Totales x Codigo


lblEstatus.Caption = "Procesando Código: " & strCodigo
DoEvents

PrgBar.Value = 1

If vRetencion Then
  strSQL = "Select ((R.cuota * R.plazo) - R.amortiza) as Saldo,R.Cuota,R.interesv,R.plazo" _
         & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
         & " where R.saldo > 0 and R.Estado = 'A' and R.codigo ='" & strCodigo & "'"
Else
  strSQL = "Select R.Saldo,R.Cuota,R.interesv,R.plazo" _
         & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
         & " where R.saldo > 0 and R.Estado = 'A' and R.codigo ='" & strCodigo & "'"
End If
If vInstitucion > 0 Then strSQL = strSQL & " and S.cod_institucion = " & vInstitucion

Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
'Insertar Codigo

curTotalInt = 0
curTotalAmortiza = 0

strSQL = "Insert #ProyectaCartera(cod_institucion,linea,descripcion,saldo_inicial,plazo"

For i = 1 To iMeses
    strSQL = strSQL & ",saldo" & i
    strSQL = strSQL & ",int" & i
    strSQL = strSQL & ",amortiza" & i
Next i

strSQL = strSQL & ",TOTAL_INT,TOTAL_AMORTIZA"
strSQL = strSQL & ") values(" & vInstitucion & ",'" & strCodigo & "','" & strDescripcion & "',0,0"
For i = 1 To iMeses
    strSQL = strSQL & ",0,0,0"
Next i
strSQL = strSQL & ",0,0)"

Call ConectionExecute(strSQL)

ReDim vInt(iMeses) As Currency
ReDim vAmortiza(iMeses) As Currency
ReDim vSaldo(iMeses) As Currency

ReDim yInt(iMeses) As Currency
ReDim yAmortiza(iMeses) As Currency
ReDim ySaldo(iMeses) As Currency


For i = 1 To iMeses
  yInt(i) = 0
  yAmortiza(i) = 0
  ySaldo(i) = 0
Next i

On Error Resume Next

vPlazo = 0

Do While Not rs.EOF

  vPlazo = vPlazo + rs!Plazo
  
  lblEstatus.Caption = "Procesando Línea: " & strCodigo & " Registro : " & PrgBar.Value & " de " & PrgBar.Max
  DoEvents
  
 xSaldo = rs!Saldo
 vSaldo(1) = rs!Saldo
    
 For i = 1 To iMeses
  
  vInt(i) = 0
  vAmortiza(i) = 0
  xSaldo = vSaldo(i)
  
  If vSaldo(i) > 0 Then
           xInt(x) = 0
           xAmort(x) = 0
          
          If rs!interesv > 0 Then
           xInt(x) = xSaldo * (rs!interesv / 1200)
          Else
           xInt(x) = 0
          End If
          
          If xSaldo >= rs!Cuota - xInt(x) Then
            xAmort(x) = rs!Cuota - xInt(x)
            xSaldo = xSaldo - xAmort(x)
          Else
            xAmort(x) = xSaldo
            xSaldo = 0
          End If
          
          vAmortiza(i) = vAmortiza(i) + xAmort(x)
          vInt(i) = vInt(i) + xInt(x)
        
          
          curTotalAmortiza = curTotalAmortiza + vAmortiza(i)
          curTotalInt = curTotalInt + vInt(i)
    
  End If
  
  If i < iMeses Then
     vSaldo(i + 1) = xSaldo
  End If
 
  yInt(i) = yInt(i) + vInt(i)
  yAmortiza(i) = yAmortiza(i) + vAmortiza(i)
  ySaldo(i) = ySaldo(i) + vSaldo(i)
 
 Next i
    
 rs.MoveNext
 PrgBar.Value = PrgBar.Value + 1
 vPlazo = 0

Loop

 strSQL = "UPDATE #ProyectaCartera set saldo_inicial = saldo_inicial + " & ySaldo(1) _
        & ",plazo = " & vPlazo / PrgBar.Value
 For i = 1 To iMeses
   strSQL = strSQL & ",saldo" & i & " = saldo" & i & " + " & ySaldo(i) & ","
   strSQL = strSQL & "int" & i & " = int" & i & " + " & yInt(i) & ","
   strSQL = strSQL & "amortiza" & i & " = amortiza" & i & " + " & yAmortiza(i)
 Next i
 strSQL = strSQL & ",TOTAL_INT = " & curTotalInt
 strSQL = strSQL & ",TOTAL_AMORTIZA = " & curTotalAmortiza
 strSQL = strSQL & " where linea = '" & strCodigo & "'"

 Call ConectionExecute(strSQL)


rs.Close

End Sub



Private Sub sbLLenaTablaAnual(vInstitucion As Integer, strCodigo As String, iAnios As Integer, strDescripcion As String, vRetencion As Boolean)
Dim curInt As Currency, curAmortiza As Currency, curSaldo As Currency
Dim curTotalInt As Currency, curTotalAmortiza As Currency, i As Integer
Dim vPlazo As Integer
Dim xSaldo As Currency, x As Integer, xInt(12) As Currency, xAmort(12) As Currency 'Mensuales

Dim vInt() As Currency, vAmortiza() As Currency, vSaldo() As Currency 'Anuales
Dim ySaldo() As Currency, yInt() As Currency, yAmortiza() As Currency 'Totales x Codigo



lblEstatus.Caption = "Procesando Línea: " & strCodigo
DoEvents

PrgBar.Value = 1

If vRetencion Then
  strSQL = "Select ((R.cuota * R.plazo) - R.amortiza) as Saldo,R.Cuota,R.interesv,R.plazo" _
         & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
         & " where R.saldo > 0 and R.Estado = 'A' and R.codigo ='" & strCodigo & "'"
Else
  strSQL = "Select R.Saldo,R.Cuota,R.interesv,R.plazo" _
         & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
         & " where R.saldo > 0 and R.Estado = 'A' and R.codigo ='" & strCodigo & "'"
End If
If vInstitucion > 0 Then strSQL = strSQL & " and S.cod_institucion = " & vInstitucion

Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
'Insertar Codigo

curTotalInt = 0
curTotalAmortiza = 0

strSQL = "Insert #ProyectaCartera(cod_institucion,codigo,descripcion,saldo_inicial,plazo"

For i = 1 To iAnios
    strSQL = strSQL & ",saldo" & i
    strSQL = strSQL & ",int" & i
    strSQL = strSQL & ",amortiza" & i
Next i

strSQL = strSQL & ",TOTAL_INT,TOTAL_AMORTIZA"
strSQL = strSQL & ") values(" & vInstitucion & ",'" & strCodigo & "','" & strDescripcion & "',0,0"
For i = 1 To iAnios
    strSQL = strSQL & ",0,0,0"
Next i
strSQL = strSQL & ",0,0)"

Call ConectionExecute(strSQL)

ReDim vInt(iAnios) As Currency
ReDim vAmortiza(iAnios) As Currency
ReDim vSaldo(iAnios) As Currency

ReDim yInt(iAnios) As Currency
ReDim yAmortiza(iAnios) As Currency
ReDim ySaldo(iAnios) As Currency


For i = 1 To iAnios
  yInt(i) = 0
  yAmortiza(i) = 0
  ySaldo(i) = 0
Next i

On Error Resume Next

vPlazo = 0

Do While Not rs.EOF

  vPlazo = vPlazo + rs!Plazo
  
  lblEstatus.Caption = "Procesando Línea: " & strCodigo & " Registro : " & PrgBar.Value & " de " & PrgBar.Max
  DoEvents
  
 xSaldo = rs!Saldo
 vSaldo(1) = rs!Saldo
    
 For i = 1 To iAnios
  vInt(i) = 0
  vAmortiza(i) = 0
  xSaldo = vSaldo(i)
  
  If vSaldo(i) > 0 Then
  
        For x = 1 To 12
           xInt(x) = 0
           xAmort(x) = 0
          
          If rs!interesv > 0 Then
           xInt(x) = xSaldo * (rs!interesv / 1200)
          Else
           xInt(x) = 0
          End If
          
          If xSaldo >= rs!Cuota - xInt(x) Then
            xAmort(x) = rs!Cuota - xInt(x)
            xSaldo = xSaldo - xAmort(x)
          Else
            xAmort(x) = xSaldo
            xSaldo = 0
          End If
          
          vAmortiza(i) = vAmortiza(i) + xAmort(x)
          vInt(i) = vInt(i) + xInt(x)
        
        Next x
          curTotalAmortiza = curTotalAmortiza + vAmortiza(i)
          curTotalInt = curTotalInt + vInt(i)
    
  End If
  
  If i < iAnios Then
     vSaldo(i + 1) = xSaldo
  End If
 
  yInt(i) = yInt(i) + vInt(i)
  yAmortiza(i) = yAmortiza(i) + vAmortiza(i)
  ySaldo(i) = ySaldo(i) + vSaldo(i)
 
 Next i
    
 rs.MoveNext
 PrgBar.Value = PrgBar.Value + 1
 vPlazo = 0

Loop

 strSQL = "UPDATE #ProyectaCartera set saldo_inicial = saldo_inicial + " & ySaldo(1) _
        & ",plazo = " & vPlazo / PrgBar.Value
 For i = 1 To iAnios
   strSQL = strSQL & ",saldo" & i & " = saldo" & i & " + " & ySaldo(i) & ","
   strSQL = strSQL & "int" & i & " = int" & i & " + " & yInt(i) & ","
   strSQL = strSQL & "amortiza" & i & " = amortiza" & i & " + " & yAmortiza(i)
 Next i
 strSQL = strSQL & ",TOTAL_INT = " & curTotalInt
 strSQL = strSQL & ",TOTAL_AMORTIZA = " & curTotalAmortiza
 strSQL = strSQL & " where codigo = '" & strCodigo & "'"

 Call ConectionExecute(strSQL)


rs.Close

End Sub


Private Sub sbLLenaTablaTasas(vInstitucion As Integer, strCodigo As String, strDescripcion As String)
Dim curSaldoTotal As Currency, iCasos As Long
Dim dbTasa As Double, dbPlazo As Double

lblEstatus.Caption = "Procesando Línea: " & strCodigo
DoEvents

PrgBar.Value = 1


'Sacar Totales x Linea -> Para Base de Factores

curSaldoTotal = 0
iCasos = 0

strSQL = "Select isnull(sum(R.Saldo),0) as Saldo,count(*) as Casos" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " where R.saldo > 0 and R.Estado = 'A' and R.codigo ='" & strCodigo & "'" _
       & " and R.proceso <> 'J'"

If vInstitucion > 0 Then strSQL = strSQL & " and S.cod_institucion = " & vInstitucion

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
   curSaldoTotal = rs!Saldo
   iCasos = rs!Casos
End If
rs.Close

If curSaldoTotal = 0 Then Exit Sub


'Calculando Tasas y Plazos
strSQL = "Select sum( (R.saldo / " & curSaldoTotal & ") * R.interesv) as TasaPond" _
       & ", sum( (R.saldo / " & curSaldoTotal & ") * dbo.fxCrdPlazoRestante(R.Plazo,R.Prideduc," & GLOBALES.glngFechaCR & ") ) as PlazoPond" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " where R.saldo > 0 and R.Estado = 'A' and R.codigo ='" & strCodigo & "'" _
       & " and R.proceso <> 'J'"

If vInstitucion > 0 Then strSQL = strSQL & " and S.cod_institucion = " & vInstitucion

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  dbTasa = rs!tasaPond
  dbPlazo = rs!plazoPond
End If
rs.Close

'Registra
strSQL = "insert #TasasPlazos(COD_INSTITUCION,LINEA,DESCRIPCION,SALDO,CASOS,PLAZO,TASA)" _
       & " values(" & vInstitucion & ",'" & strCodigo & "','" & strDescripcion & "'," & curSaldoTotal _
       & "," & iCasos & "," & dbPlazo & "," & dbTasa & ")"
Call ConectionExecute(strSQL)


End Sub


Private Sub sbRecuperacionCarteraAnual(vRetencion As Boolean, Optional vInstitucion As Integer = 0)
Dim strSQL As String, i As Integer, lng As Long
Dim itmX As ListItem

Me.MousePointer = vbHourglass

On Error Resume Next

glogon.Conection.Execute "drop table #ProyectaCartera"

On Error GoTo vError

strSQL = "CREATE TABLE #ProyectaCartera (COD_INSTITUCION INT NOT NULL,CODIGO varchar(6) NOT NULL," _
        & "DESCRIPCION varchar(40) NULL,SALDO_INICIAL float NULL,PLAZO int Null,"

For i = 1 To Val(txtProyeccion) 'Aqui el Rango
 strSQL = strSQL & "SALDO" & i & " float NULL,"
 strSQL = strSQL & "INT" & i & " float NULL,"
 strSQL = strSQL & "AMORTIZA" & i & " float NULL,"
Next
 strSQL = strSQL & "TOTAL_INT FLOAT NULL,TOTAL_AMORTIZA FLOAT NULL"
 strSQL = strSQL & ")"

Call ConectionExecute(strSQL)

With lsw.ListItems
    For lng = 1 To .Count
      If .Item(lng).Checked Then
            Call sbLLenaTablaAnual(vInstitucion, .Item(lng).Text, Val(txtProyeccion), .Item(lng).SubItems(1), vRetencion)
      End If
    Next lng
End With

Call sbCreaArchivoTexto("#ProyectaCartera", "ProyCarteraAnual.txt")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbRecuperacionCartera(vRetencion As Boolean, Optional vInstitucion As Integer = 0)
Dim strSQL As String, i As Integer, lng As Long
Dim itmX As ListItem

Me.MousePointer = vbHourglass

On Error Resume Next
glogon.Conection.Execute "drop table #ProyectaCartera"

On Error GoTo vError

strSQL = "CREATE TABLE #ProyectaCartera (COD_INSTITUCION INT NOT NULL,LINEA varchar(6) NOT NULL," _
        & "DESCRIPCION varchar(60) NULL,SALDO_INICIAL float NULL,PLAZO int Null,"

For i = 1 To Val(txtProyeccion) 'Aqui el Rango
 strSQL = strSQL & "SALDO" & i & " float NULL,"
 strSQL = strSQL & "INT" & i & " float NULL,"
 strSQL = strSQL & "AMORTIZA" & i & " float NULL,"
Next
 strSQL = strSQL & "TOTAL_INT FLOAT NULL,TOTAL_AMORTIZA FLOAT NULL"
 strSQL = strSQL & ")"

Call ConectionExecute(strSQL)

With lsw.ListItems
    For lng = 1 To .Count
      If .Item(lng).Checked Then
            Call sbLLenaTabla(vInstitucion, .Item(lng).Text, Val(txtProyeccion), .Item(lng).SubItems(1), vRetencion)
      End If
    Next lng
End With

If vRetencion Then
    Call sbCreaArchivoTexto("#ProyectaCartera", "ProyRetencionMensual.txt")
Else
    Call sbCreaArchivoTexto("#ProyectaCartera", "ProyCarteraMensual.txt")
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbTasasPlazoPond(Optional vInstitucion As Integer = 0)
Dim strSQL As String, i As Integer, lng As Long
Dim itmX As ListItem

Me.MousePointer = vbHourglass

On Error Resume Next
glogon.Conection.Execute "drop table #TasasPlazos"

On Error GoTo vError

strSQL = "CREATE TABLE #TasasPlazos (COD_INSTITUCION INT NOT NULL,LINEA varchar(6) NOT NULL," _
        & "DESCRIPCION varchar(40) NULL,SALDO float NULL, CASOS int null, PLAZO dec(10,4) Null,TASA dec(10,4))"
Call ConectionExecute(strSQL)

With lsw.ListItems
    For lng = 1 To .Count
      If .Item(lng).Checked Then
            Call sbLLenaTablaTasas(vInstitucion, .Item(lng).Text, .Item(lng).SubItems(1))
      End If
    Next lng
End With

Call sbCreaArchivoTexto("#TasasPlazos", "TasasPlazosPond.txt")


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Function fxValidaDatos() As Boolean
fxValidaDatos = True

If Len(Trim(txtProyeccion)) = 0 Then
 fxValidaDatos = False
 MsgBox "Verifique el Rango de Proyección", vbCritical
 Exit Function
End If

If IsNumeric(txtProyeccion) = False Then
 fxValidaDatos = False
 MsgBox "Verifique el Rango de Proyección", vbCritical
 Exit Function
End If

If Val(txtProyeccion) < 2 Then
 fxValidaDatos = False
 MsgBox "El rango mínimo de Proyección es 2 - Verifique", vbCritical
 Exit Function
End If

If Val(txtProyeccion) > 60 Then
 fxValidaDatos = False
 MsgBox "El rango máximo de Proyección es 60 - Verifique", vbCritical
 Exit Function
End If

End Function


Private Sub sbEndeudamiento()
Dim strSQL As String, rs As New ADODB.Recordset
Dim fn, pRuta As String, strCadena As String, pPaso As Boolean
Dim pArchivo As String, i As Integer, vFecha As Date
'Genera Archivo de Texto

fn = FreeFile

Me.MousePointer = vbHourglass

lblEstatus.Caption = "Montando Información Principal..."
DoEvents

PrgBar.Value = 1

On Error GoTo vError


vFecha = fxFechaServidor

pArchivo = "Endeudamiento_" & Format(vFecha, "yyyymmdd") & ".txt"
pRuta = SIFGlobal.DirectorioDeResultados & "\" & pArchivo

If Dir(SIFGlobal.DirectorioDeResultados, vbDirectory) = "" Then
    MkDir SIFGlobal.DirectorioDeResultados
End If

 
If Dir(pRuta, vbArchive) <> "" Then 'El archivo existe
  Kill pRuta
End If


strSQL = "select S.Cedula as 'Identificación',S.Nombre,S.fechaingreso as 'Ingreso',Est.descripcion as 'Estado',I.descripcion as 'Institución'" _
       & ",A.fecAhorro as 'Ult.Aporte',A.AhorroMes as 'Aporte.Mes',A.Ahorro as 'Obrero',A.Capitaliza as 'Capitalización',A.Aporte as 'Patronal'" _
       & ",isnull(sum(R.saldo),0) as 'Saldos',dbo.fxCRDFianzas(S.cedula) as Fianzas" _
       & " from Socios S inner join Instituciones I on S.cod_institucion = I.cod_Institucion" _
       & " inner join Afi_Estados_Persona Est on S.estadoActual = Est.cod_estado" _
       & " inner join Ahorro_Consolidado A on S.cedula = A.cedula" _
       & " left join reg_creditos R on S.cedula = R.cedula and R.estado = 'A'"
       
pPaso = False
If cboEstado.Text <> "TODOS" Then
  pPaso = True
  strSQL = strSQL & " Where S.estadoActual = '" & cboEstado.ItemData(cboEstado.ListIndex) & "'"
End If

If cboInstitucion.Text <> "TODOS" Then
  If pPaso Then
    strSQL = strSQL & " and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
  Else
    strSQL = strSQL & " where S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
  End If

  pPaso = True
End If
       
       
strSQL = strSQL & " group by S.cedula,S.nombre,S.fechaingreso,S.estadoactual,A.fecAhorro" _
       & ",A.ahorro,A.capitaliza,A.aporte,A.ahorroMes,I.descripcion,Est.descripcion"

Call OpenRecordSet(rs, strSQL)
PrgBar.Max = rs.RecordCount + 1

strCadena = ""

lblEstatus.Caption = "Creando Archivo de Texto"
DoEvents


Open pRuta For Output As #fn  ' Crea Archivo.

    Print #fn, "Listado de Estado de Endeudamiento"
    Print #fn, "Usuario: " & vbTab & glogon.Usuario & vbTab & "Fecha : " & Format(vFecha, "dd/mm/yyyy") & vbTab & "Hora : " _
               & Format(vFecha, "hh:mm:ss AMPM")
    Print #fn, "Institución:" & vbTab & cboInstitucion.Text
    Print #fn, "Estado:" & vbTab & cboEstado.Text
    
    Print #fn, ""
    Print #fn, ""
    For i = 0 To rs.Fields.Count - 1
     strCadena = strCadena & rs.Fields(i).Name & vbTab
    Next
    Print #fn, strCadena
    
    
    Do While Not rs.EOF
     strCadena = ""
     For i = 0 To rs.Fields.Count - 1
       strCadena = strCadena & rs.Fields(i).Value & vbTab
     Next
     Print #fn, strCadena
     
     PrgBar.Value = PrgBar.Value + 1
     lblEstatus.Caption = "Creando Archivo de Texto     (" & PrgBar.Value & ")"
     DoEvents
     rs.MoveNext
    Loop
    rs.Close

Close #fn   ' Cierra Archivo.

Me.MousePointer = vbDefault
MsgBox "Se creó el siguiente Archivo : " & pRuta, vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbListaPersonasvsDeuda(Optional pDeuda As Boolean = False)
Dim strSQL As String, rs As New ADODB.Recordset
Dim fn, pRuta As String, strCadena As String
Dim pArchivo As String, i As Integer, vFecha As Date
Dim pPaso As Boolean
'Genera Archivo de Texto

fn = FreeFile

Me.MousePointer = vbHourglass

lblEstatus.Caption = "Montando Información Principal..."
DoEvents

PrgBar.Value = 1

On Error GoTo vError

vFecha = fxFechaServidor

If Dir(SIFGlobal.DirectorioDeResultados, vbDirectory) = "" Then
    MkDir SIFGlobal.DirectorioDeResultados
End If


If pDeuda Then
    pArchivo = "PersonasConDeuda_" & Format(vFecha, "yyyymmdd") & ".txt"
Else
    pArchivo = "PersonasSinDeuda_" & Format(vFecha, "yyyymmdd") & ".txt"
End If

pRuta = SIFGlobal.DirectorioDeResultados & "\" & pArchivo

If Dir(pRuta, vbArchive) <> "" Then  'El archivo existe
  Kill pRuta
End If

If Not GLOBALES.SysASEVersion Then
    strSQL = "select S.Cedula as 'Identificación',S.Nombre,S.fechaingreso as 'Ingreso',Est.Descripcion as 'Estado'" _
           & ",A.Ahorro,A.Capitaliza,A.Aporte, case when I.porc_ahorro = 0 then 0 else  (isnull(A.ahorroMes,0) /(I.porc_ahorro/100)) end as 'Salario'" _
           & ",datediff(yyyy,S.Fecha_nac,dbo.MyGetdate()) as Edad" & IIf(pDeuda, ",dbo.fxCrdSaldo(S.cedula) as 'Saldos'", "") _
           & ",dbo.fxCRDClasificacion(S.cedula,dbo.MyGetdate()) as Categoria,I.descripcion as 'Institución',Dept.Descripcion as 'Departamento',Sec.Descripcion as 'Sección'" _
           & ",dbo.fxAFITelefono(S.cedula,1) as TelHab" _
           & ",dbo.fxAFITelefono(S.cedula,2) as TelTrab, dbo.fxAFITelefono(S.cedula,3) as TelCell" _
           & " from Socios S inner join instituciones I on S.cod_institucion = I.cod_institucion" _
           & " inner join Afi_Estados_Persona Est on S.estadoActual = Est.cod_Estado" _
           & " left join AFDepartamentos Dept on S.cod_Institucion = Dept.cod_Institucion and S.cod_departamento = S.cod_Departamento" _
           & " left join AFSecciones Sec on S.cod_Institucion = Sec.cod_Institucion and S.cod_departamento = Sec.cod_Departamento and S.cod_Seccion = Sec.cod_seccion" _
           & " left join Ahorro_Consolidado A on S.cedula = A.cedula"
Else
    strSQL = "select S.Cedula as 'Identificación',S.Nombre,S.fechaingreso as 'Ingreso',Est.Descripcion as 'Estado'" _
           & ",A.Ahorro,A.Capitaliza,A.Aporte, case when I.porc_ahorro = 0 then 0 else  (isnull(A.ahorroMes,0) /(I.porc_ahorro/100)) end as 'Salario'" _
           & ",datediff(yyyy,S.Fecha_nac,dbo.MyGetdate()) as Edad" & IIf(pDeuda, ",dbo.fxCrdSaldo(S.cedula) as 'Saldos'", "") _
           & ",dbo.fxCRDClasificacion(S.cedula,dbo.MyGetdate()) as Categoria,I.descripcion as 'Institución'" _
           & ",U.descripcion as 'U.Programatica',Tra.UT_Descripcion as 'U.Trabajo' ,dbo.fxAFITelefono(S.cedula,1) as TelHab" _
           & ",dbo.fxAFITelefono(S.cedula,2) as TelTrab, dbo.fxAFITelefono(S.cedula,3) as TelCell" _
           & " from Socios S inner join instituciones I on S.cod_institucion = I.cod_institucion" _
           & " inner join Afi_Estados_Persona Est on S.estadoActual = Est.cod_Estado" _
           & " left join Ahorro_Consolidado A on S.cedula = A.cedula" _
           & " left join uprogramatica U on S.up = U.codigo" _
           & " left join utrabajo Tra on S.ut = Tra.Ut_Codigo"
End If


pPaso = False
If cboEstado.Text <> "TODOS" Then
  pPaso = True
  strSQL = strSQL & " Where S.estadoActual = '" & cboEstado.ItemData(cboEstado.ListIndex) & "'"
End If

If cboInstitucion.Text <> "TODOS" Then
  If pPaso Then
    strSQL = strSQL & " and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
  Else
    strSQL = strSQL & " where S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
  End If

  pPaso = True
End If

If pDeuda Then
    If pPaso Then
        strSQL = strSQL & " and S.cedula in(select R.cedula from reg_creditos R" _
                   & " inner join catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
                   & " where R.estado = 'A')"
    Else
        strSQL = strSQL & " where S.cedula in(select R.cedula from reg_creditos R" _
                   & " inner join catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
                   & " where R.estado = 'A')"
    End If
Else
    If pPaso Then
        strSQL = strSQL & " and S.cedula not in(select R.cedula from reg_creditos R" _
                   & " inner join catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
                   & " where R.estado = 'A')"
    Else
        strSQL = strSQL & " where S.cedula not in(select R.cedula from reg_creditos R" _
                   & " inner join catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
                   & " where R.estado = 'A')"
    End If
End If
Call OpenRecordSet(rs, strSQL)
PrgBar.Max = rs.RecordCount + 1

strCadena = ""

lblEstatus.Caption = "Creando Archivo de Texto"
DoEvents


Open pRuta For Output As #fn  ' Crea Archivo.

    If pDeuda Then
        Print #fn, "Listado de Personas con Deudas"
    Else
        Print #fn, "Listado de Personas sin Deudas"
    End If
    
    Print #fn, "Usuario: " & vbTab & glogon.Usuario & vbTab & "Fecha : " & Format(vFecha, "dd/mm/yyyy") & vbTab & "Hora : " _
               & Format(vFecha, "hh:mm:ss AMPM")
    Print #fn, "Institución:" & vbTab & cboInstitucion.Text
    Print #fn, "Estado:" & vbTab & cboEstado.Text
    
    Print #fn, ""
    Print #fn, ""
    For i = 0 To rs.Fields.Count - 1
     strCadena = strCadena & rs.Fields(i).Name & vbTab
    Next
    Print #fn, strCadena
    
    
    Do While Not rs.EOF
     strCadena = ""
     For i = 0 To rs.Fields.Count - 1
       strCadena = strCadena & rs.Fields(i).Value & vbTab
     Next
     Print #fn, strCadena
     
     PrgBar.Value = PrgBar.Value + 1
     lblEstatus.Caption = "Creando Archivo de Texto     (" & PrgBar.Value & ")"
     DoEvents
     rs.MoveNext
    Loop
    rs.Close

Close #fn   ' Cierra Archivo.

Me.MousePointer = vbDefault
MsgBox "Se creó el siguiente Archivo : " & pRuta, vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxFechas(vFecha As Date, ByVal Meses As Integer, Optional Signo As String = "-", Optional Dias As Integer = 0) As Date
Dim iMes As Integer, lngAnio As Long, iDias As Integer
Dim i As Integer

iMes = Month(vFecha)
lngAnio = Year(vFecha)
iDias = Day(vFecha)


If Signo = "-" Then

  Do While Meses <> 0
  
     If iMes = 1 Then
        iMes = 12
        lngAnio = lngAnio - 1
     Else
        iMes = iMes - 1
     End If
     
     Meses = Meses - 1
     
  Loop

Else

  Do While Meses <> 0
  
     If iMes = 12 Then
        iMes = 1
        lngAnio = lngAnio + 1
     Else
        iMes = iMes + 1
     End If
     
     Meses = Meses - 1
     
  Loop


End If



For i = 1 To Dias
   
   If iDias = 1 Then
        iDias = 30
        If iMes = 1 Then
         iMes = 12
         lngAnio = lngAnio - 1
        Else
         iMes = iMes - 1
        End If
    
    Else
        iDias = iDias - 1
   End If

Next i

fxFechas = iDias & "/" & iMes & "/" & lngAnio

End Function


Private Sub sbAntiguedadPersonaAhorro()
Dim strSQL As String, rs As New ADODB.Recordset
Dim fn, pArchivo As String, pRuta As String
Dim vFecha As Date, i As Integer, i2 As Integer

fn = FreeFile

Me.MousePointer = vbHourglass

lblEstatus.Caption = "Montando Información Principal..."
DoEvents

PrgBar.Value = 1
PrgBar.Max = 7


On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados

vFecha = fxFechaServidor

pRuta = SIFGlobal.DirectorioDeResultados & "\AntiguedadAhorros.TXT"
pArchivo = Dir(pRuta, vbArchive)

If pArchivo = "ANTAHORRO.TXT" Then 'El archivo existe
  Kill pRuta
End If

Open pRuta For Output As #fn  ' Crea Archivo.

Print #fn, ",,ProGrX: Antiguedad Personas (Ahorro), Usuario: " & glogon.Usuario _
           & ", Fecha : " & Format(vFecha, "dd/mm/yyyy") & ",Hora : " _
           & Format(Time, "hh:mm:ss AMPM")


Print #fn, "Antiguedad,Casos,Ahorros"

On Error GoTo vError

lblEstatus.Caption = "Creando Archivo de Texto"
DoEvents

i = 6

strSQL = "select isnull(sum(A.ahorro),0) as Ahorro, isnull(count(*),0) as casos" _
       & " from Ahorro_Consolidado A inner join Socios S on A.cedula = S.cedula" _
       & " and S.estadoactual = 'S' and S.fechaingreso BETWEEN '" _
       & Format(fxFechas(vFecha, i), "yyyy/mm/dd") & "' and '" & Format(vFecha, "yyyy/mm/dd") & "'"


Call OpenRecordSet(rs, strSQL)

Print #fn, i & ","; rs!ahorro & ","; rs!Casos

rs.Close

PrgBar.Value = PrgBar.Value + 1

i = 12
i2 = 6

Do While i < 60
    strSQL = "select isnull(sum(A.ahorro),0) as Ahorro, isnull(count(*),0) as casos" _
           & " from Ahorro_Consolidado A inner join Socios S on A.cedula = S.cedula" _
           & " and S.estadoactual = 'S' and S.fechaingreso BETWEEN '" _
           & Format(fxFechas(vFecha, i), "yyyy/mm/dd") _
           & "' and '" & Format(fxFechas(vFecha, i2, , 1), "yyyy/mm/dd") & "'"

    Call OpenRecordSet(rs, strSQL)

    Print #fn, i & ","; rs!ahorro & ","; rs!Casos

    rs.Close

  i2 = i
  i = i + 12
  PrgBar.Value = PrgBar.Value + 1
  
Loop



strSQL = "select isnull(sum(A.ahorro),0) as Ahorro, isnull(count(*),0) as casos" _
       & " from Ahorro_Consolidado A inner join Socios S on A.cedula = S.cedula" _
       & " and S.estadoactual = 'S' and S.fechaingreso < '" _
       & Format(fxFechas(vFecha, i, , 1), "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL)

Print #fn, i & "+,"; rs!ahorro & ","; rs!Casos

rs.Close

PrgBar.Value = PrgBar.Value + 1



Close #fn
Me.MousePointer = vbDefault

MsgBox "Se Creó el Siguiente Archivo : " & pRuta, vbInformation

lblEstatus.Caption = ""
PrgBar.Value = 1

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Sub sbAntiguedadPersonaSaldos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim fn, pArchivo As String, pRuta As String
Dim vFecha As Date, rs2 As New ADODB.Recordset
Dim curTotal(10, 2) As Currency, i As Integer, i2 As Integer


fn = FreeFile

Me.MousePointer = vbHourglass

lblEstatus.Caption = "Montando Información Principal..."
DoEvents

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados

vFecha = fxFechaServidor

pRuta = SIFGlobal.DirectorioDeResultados & "\Antiguedad Ahorros vrs Saldos.TXT"
pArchivo = Dir(pRuta, vbArchive)

If pArchivo = "AHvrsSLD.TXT" Then 'El archivo existe
  Kill pRuta
End If

Open pRuta For Output As #fn  ' Crea Archivo.

Print #fn, ",,ProGrX: Cuadro de Ahorros vrs Saldos, Usuario: " & glogon.Usuario _
           & ", Fecha : " & Format(vFecha, "dd/mm/yyyy") & ",Hora : " _
           & Format(Time, "hh:mm:ss AMPM")


Print #fn, "Rango,Saldos,Casos"

On Error GoTo vError

lblEstatus.Caption = "Creando Archivo de Texto"
DoEvents

strSQL = "select cedula,ahorro from ahorro_consolidado where ahorro > 0"
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount
PrgBar.Value = 1


For i = 1 To 10
 For i2 = 1 To 2
   curTotal(i, i2) = 0
 Next i2
Next i

Do While Not rs.EOF
 strSQL = "select isnull(sum(saldo),0) as saldo from reg_creditos where estado = 'A'" _
        & " and saldo > 0 and cedula = '" & Trim(rs!Cedula) & "'"
 rs2.CursorLocation = adUseServer
 rs2.Open strSQL, glogon.Conection, adOpenStatic
 If rs2!Saldo > 0 Then
    If rs!ahorro >= 1000000 Then
       curTotal(10, 1) = curTotal(10, 1) + rs2!Saldo
       curTotal(10, 2) = curTotal(10, 2) + 1
    Else
       curTotal(Mid(Trim(CStr(rs!ahorro)), 1, 1), 1) = curTotal(Mid(Trim(CStr(rs!ahorro)), 1, 1), 1) + rs2!Saldo
       curTotal(Mid(Trim(CStr(rs!ahorro)), 1, 1), 2) = curTotal(Mid(Trim(CStr(rs!ahorro)), 1, 1), 2) + 1
    End If
 End If
 rs2.Close
 rs.MoveNext
 If PrgBar.Max > PrgBar.Value Then PrgBar.Value = PrgBar.Value + 1
Loop

rs.Close

Print #fn, "1000000," & curTotal(1, 1) & "," & curTotal(1, 2)
Print #fn, "2000000," & curTotal(2, 1) & "," & curTotal(2, 2)
Print #fn, "3000000," & curTotal(3, 1) & "," & curTotal(3, 2)
Print #fn, "4000000," & curTotal(4, 1) & "," & curTotal(4, 2)
Print #fn, "5000000," & curTotal(5, 1) & "," & curTotal(5, 2)
Print #fn, "6000000," & curTotal(6, 1) & "," & curTotal(6, 2)
Print #fn, "7000000," & curTotal(7, 1) & "," & curTotal(7, 2)
Print #fn, "8000000," & curTotal(8, 1) & "," & curTotal(8, 2)
Print #fn, "9000000," & curTotal(9, 1) & "," & curTotal(9, 2)
Print #fn, "10000000+," & curTotal(10, 1) & "," & curTotal(10, 2)

Close #fn

Me.MousePointer = vbDefault

MsgBox "Se Creó el Siguiente Archivo : " & pRuta, vbInformation

lblEstatus.Caption = ""
PrgBar.Value = 1

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub btnResultado_Click()
 
        Call sbSIFGridExportar(vGrid, vHeaders, lblReporte.Caption)

End Sub

Private Sub cboCartera_Click()
If vPaso Then Exit Sub

Select Case lblReporte.Tag
 Case "x00", "x01", "x02"
   Call sbCargaLsw(False)
 Case "x03"
    Call sbCargaLsw(True)
End Select

End Sub

Private Sub chkActivas_Click()
Select Case lblReporte.Tag
 Case "x00", "x01", "x02"
   Call sbCargaLsw(False)
 Case "x03"
    Call sbCargaLsw(True)
End Select
End Sub

Private Sub chkInternas_Click()
Select Case lblReporte.Tag
 Case "x00", "x01", "x02"
   Call sbCargaLsw(False)
 Case "x03"
    Call sbCargaLsw(True)
End Select
End Sub

Private Sub chkTodos_Click()
Dim lng As Long

Me.MousePointer = vbHourglass
For lng = 1 To lsw.ListItems.Count
    lsw.ListItems.Item(lng).Checked = chkTodos.Value
Next lng
Me.MousePointer = vbDefault

End Sub


Private Sub sbvGridColSize()
Dim i As Integer

For i = 1 To vGrid.MaxCols
  vGrid.ColWidth(i) = vGrid.MaxTextColWidth(i)
Next i

End Sub


Private Sub sbvGridCol(Index As Integer)

vGrid.DAutoSizeCols = DAutoSizeColsMax
vGrid.DAutoHeadings = False
vGrid.MaxRows = 0

Select Case Index
  Case 0 'Estadistica Afiliacion
      vGrid.MaxCols = 10
      vGrid.SetText 1, 0, "Corte"
      vGrid.SetText 2, 0, "Asociados"
      vGrid.SetText 3, 0, "Nuevos"
      vGrid.SetText 4, 0, "Reincorporacion"
      vGrid.SetText 5, 0, "Activaciones"
      vGrid.SetText 6, 0, "Inactivaciones"
      vGrid.SetText 7, 0, "Afi.Automatica"
      
      vGrid.SetText 8, 0, "Salidas"
      vGrid.SetText 9, 0, "Ex Asociados"
      vGrid.SetText 10, 0, "Aporte Custodia"
      
      vHeaders.Columnas = 10
      vHeaders.Headers(1) = "Corte"
      vHeaders.Headers(2) = "Asociados"
      vHeaders.Headers(3) = "Nuevos"
      vHeaders.Headers(4) = "Reincorporacion"
      vHeaders.Headers(5) = "Activaciones"
      vHeaders.Headers(6) = "Inactivacion"
      vHeaders.Headers(7) = "Afi.Automática"
      
      vHeaders.Headers(8) = "Salidas"
      vHeaders.Headers(9) = "Ex Asociados"
      vHeaders.Headers(10) = "Aporte Custodia"

  
  Case 1 'Estadistica Liquidaciones
  
      vGrid.MaxCols = 14
      vGrid.SetText 1, 0, "Año"
      vGrid.SetText 2, 0, "Mes"
      vGrid.SetText 3, 0, "Corte"
      vGrid.SetText 4, 0, "Liquidaciones"
      vGrid.SetText 5, 0, "Obrero Liq."
      vGrid.SetText 6, 0, "Patronal Liq."
      vGrid.SetText 7, 0, "Custodia Liq."
      vGrid.SetText 8, 0, "Capitaliza Liq."
      vGrid.SetText 9, 0, "Impuesto Retenido"
      vGrid.SetText 10, 0, "Créditos Aplicados"
      vGrid.SetText 11, 0, "Fondos Liq."
      vGrid.SetText 12, 0, "Salida Neta"
      vGrid.SetText 13, 0, "Liq. Internas"
      vGrid.SetText 14, 0, "Liq. Patronales"
      
      vHeaders.Columnas = 14
      vHeaders.Headers(1) = "Año"
      vHeaders.Headers(2) = "Mes"
      vHeaders.Headers(3) = "Corte"
      vHeaders.Headers(4) = "Liquidaciones"
      vHeaders.Headers(5) = "Obrero Liq."
      vHeaders.Headers(6) = "Patronal Liq."
      vHeaders.Headers(7) = "Custodia Liq."
      vHeaders.Headers(8) = "Capitaliza Liq."
      vHeaders.Headers(9) = "Impuesto Retenido"
      vHeaders.Headers(10) = "Créditos Aplicados"
      vHeaders.Headers(11) = "Fondos Liq."
      vHeaders.Headers(12) = "Salida Neta"
      vHeaders.Headers(13) = "Liq. Internas"
      vHeaders.Headers(14) = "Liq. Patronales"
      
  
  Case 5 'Proyeccion de Cartera
      vGrid.MaxCols = 14
      vGrid.SetText 1, 0, "Periodo"
      vGrid.SetText 2, 0, "Inicio"
      vGrid.SetText 3, 0, "Corte"
      vGrid.SetText 4, 0, "Línea"
      vGrid.SetText 5, 0, "Garantía"
      vGrid.SetText 6, 0, "Recurso"
      vGrid.SetText 7, 0, "Deductor"
      vGrid.SetText 8, 0, "Destino"
      vGrid.SetText 9, 0, "Oficina"
      vGrid.SetText 10, 0, "Int.Cor."
      vGrid.SetText 11, 0, "Principal"
      vGrid.SetText 12, 0, "Cargos"
      vGrid.SetText 13, 0, "Pólizas"
      vGrid.SetText 14, 0, "Saldo"
  
      vHeaders.Columnas = 14
      vHeaders.Headers(1) = "Periodo"
      vHeaders.Headers(2) = "Inicio"
      vHeaders.Headers(3) = "Corte"
      vHeaders.Headers(4) = "Línea"
      vHeaders.Headers(5) = "Garantía"
      vHeaders.Headers(6) = "Recurso"
      vHeaders.Headers(7) = "Deductor"
      vHeaders.Headers(8) = "Destino"
      vHeaders.Headers(9) = "Oficina"
      vHeaders.Headers(10) = "Int.Cor."
      vHeaders.Headers(11) = "Principal"
      vHeaders.Headers(12) = "Cargos"
      vHeaders.Headers(13) = "Pólizas"
      vHeaders.Headers(14) = "Saldo"
      
'      vGrid.col = 10
'      vGrid.CellType = CellTypeNumber
'      vGrid.col = 11
'      vGrid.CellType = CellTypeNumber
'      vGrid.col = 12
'      vGrid.CellType = CellTypeNumber
      
      
  Case 6 'Analisis de Endeudamiento
      vGrid.MaxCols = 6
      vGrid.SetText 1, 0, "Rango"
      vGrid.SetText 2, 0, "Casos"
      vGrid.SetText 3, 0, "Ahorros"
      vGrid.SetText 4, 0, "Aportes"
      vGrid.SetText 5, 0, "Salarios"
      vGrid.SetText 6, 0, "Saldos"
  
      vHeaders.Columnas = 6
      vHeaders.Headers(1) = "Rango"
      vHeaders.Headers(2) = "Casos"
      vHeaders.Headers(3) = "Ahorros"
      vHeaders.Headers(4) = "Aportes"
      vHeaders.Headers(5) = "Salarios"
      vHeaders.Headers(6) = "Saldos"
    
  
  Case 7 'Analisis Personas con Saldos
      vGrid.MaxCols = 7
      vGrid.SetText 1, 0, "Rango"
      vGrid.SetText 2, 0, "Casos"
      vGrid.SetText 3, 0, "Ahorros"
      vGrid.SetText 4, 0, "Aportes"
      vGrid.SetText 5, 0, "Saldos"
      vGrid.SetText 6, 0, "Monto"
      vGrid.SetText 7, 0, "Desembolsos"
      
      vHeaders.Columnas = 7
      vHeaders.Headers(1) = "Rango"
      vHeaders.Headers(2) = "Casos"
      vHeaders.Headers(3) = "Ahorros"
      vHeaders.Headers(4) = "Aportes"
      vHeaders.Headers(5) = "Saldos"
      vHeaders.Headers(6) = "Monto"
      vHeaders.Headers(7) = "Desembolsos"
    
      
  Case 8 'Analisis Patrimonio x Membresia
      vGrid.MaxCols = 4
      vGrid.SetText 1, 0, "Rango"
      vGrid.SetText 2, 0, "Casos"
      vGrid.SetText 3, 0, "Ahorros"
      vGrid.SetText 4, 0, "Aportes"
      
      vGrid.col = 1
      vGrid.CellType = CellTypeNumber
      vGrid.col = 2
      vGrid.CellType = CellTypeNumber
      vGrid.col = 3
      vGrid.CellType = CellTypeNumber
      vGrid.col = 4
      vGrid.CellType = CellTypeNumber
      
      vHeaders.Columnas = 4
      vHeaders.Headers(1) = "Rango"
      vHeaders.Headers(2) = "Casos"
      vHeaders.Headers(3) = "Ahorros"
      vHeaders.Headers(4) = "Aportes"
      
  Case 9 'Analisis Patrimonio x Institucion
      vGrid.MaxCols = 4
      vGrid.SetText 1, 0, "Institución"
      vGrid.SetText 2, 0, "Casos"
      vGrid.SetText 3, 0, "Ahorros"
      vGrid.SetText 4, 0, "Aportes"
      
      vGrid.col = 1
      vGrid.CellType = CellTypeNumber
      vGrid.col = 2
      vGrid.CellType = CellTypeNumber
      vGrid.col = 3
      vGrid.CellType = CellTypeNumber
      vGrid.col = 4
      vGrid.CellType = CellTypeNumber
      
      
      vHeaders.Columnas = 4
      vHeaders.Headers(1) = "Rango"
      vHeaders.Headers(2) = "Casos"
      vHeaders.Headers(3) = "Ahorros"
      vHeaders.Headers(4) = "Aportes"

  Case 10 'Cartera x Categoria Crediticia
      vGrid.MaxCols = 7
      vGrid.SetText 1, 0, "Estado"
      vGrid.SetText 2, 0, "Categoria"
      vGrid.SetText 3, 0, "Casos"
      vGrid.SetText 4, 0, "Montos"
      vGrid.SetText 5, 0, "Saldos"
      vGrid.SetText 6, 0, "Plazo Prom"
      vGrid.SetText 7, 0, "Tasa Prom"
      
      
      vHeaders.Columnas = 7
      vHeaders.Headers(1) = "Estado"
      vHeaders.Headers(2) = "Categoria"
      vHeaders.Headers(3) = "Casos"
      vHeaders.Headers(4) = "Montos"
      vHeaders.Headers(5) = "Saldos"
      vHeaders.Headers(6) = "Plazo Prom."
      vHeaders.Headers(7) = "Tasa Prom."
      
      
      vGrid.col = 1
      vGrid.CellType = CellTypeEdit
      vGrid.col = 2
      vGrid.CellType = CellTypeEdit
      
      vGrid.col = 3
      vGrid.CellType = CellTypeNumber
    
      vGrid.col = 4
      vGrid.CellType = CellTypeNumber
      vGrid.col = 5
      vGrid.CellType = CellTypeNumber
      vGrid.col = 6
      vGrid.CellType = CellTypeNumber
      vGrid.col = 7
      vGrid.CellType = CellTypeNumber
      
  Case 15 'Frecuencia de Creditos x Persona
      
      vGrid.MaxCols = 16
      vGrid.SetText 1, 0, "Cedula"
      vGrid.SetText 2, 0, "Nombre"
      vGrid.SetText 3, 0, "Ingreso"
      vGrid.SetText 4, 0, "Estado"
      vGrid.SetText 5, 0, "Años"
      vGrid.SetText 6, 0, "Ahorro"
      vGrid.SetText 7, 0, "Salario"
      vGrid.SetText 8, 0, "Saldo"
      vGrid.SetText 9, 0, "Cuota"
      vGrid.SetText 10, 0, "Tel.Hab"
      vGrid.SetText 11, 0, "Tel.Trab"
      vGrid.SetText 12, 0, "Tel.Cel"
      vGrid.SetText 13, 0, "%Liq."
      vGrid.SetText 14, 0, "Salario Liq"
      vGrid.SetText 15, 0, "Clasificacion Hoy"
      vGrid.SetText 16, 0, "Clasifica.PreAna."
      
      vHeaders.Columnas = 16
      vHeaders.Headers(1) = "Cédula"
      vHeaders.Headers(2) = "Nombre"
      vHeaders.Headers(3) = "Ingreso"
      vHeaders.Headers(4) = "Estado"
      vHeaders.Headers(5) = "Años"
      vHeaders.Headers(6) = "Ahorro"
      vHeaders.Headers(7) = "Salario"
      vHeaders.Headers(8) = "Saldo"
      vHeaders.Headers(9) = "Cuota"
      vHeaders.Headers(10) = "Tel.Hab."
      vHeaders.Headers(11) = "Tel.Hab."
      vHeaders.Headers(12) = "Tel.Cel."
      vHeaders.Headers(13) = "(%)Liq."
      vHeaders.Headers(14) = "Salario Liq."
      vHeaders.Headers(15) = "Clasifica. Actual"
      vHeaders.Headers(16) = "Clasifica. Registro"
      
      
      vGrid.col = 1
      vGrid.CellType = CellTypeEdit
      vGrid.col = 2
      vGrid.CellType = CellTypeEdit
      vGrid.col = 3
      vGrid.CellType = CellTypeEdit
      vGrid.col = 4
      vGrid.CellType = CellTypeEdit
      vGrid.col = 5
      vGrid.CellType = CellTypeNumber
      vGrid.col = 6
      vGrid.CellType = CellTypeNumber
      vGrid.col = 7
      vGrid.CellType = CellTypeNumber
      vGrid.col = 8
      vGrid.CellType = CellTypeNumber
      vGrid.col = 9
      vGrid.CellType = CellTypeNumber
      vGrid.col = 10
      vGrid.CellType = CellTypeEdit
      vGrid.col = 11
      vGrid.CellType = CellTypeEdit
      vGrid.col = 12
      vGrid.CellType = CellTypeEdit
      vGrid.col = 13
      vGrid.CellType = CellTypeNumber
      vGrid.col = 14
      vGrid.CellType = CellTypeNumber
      vGrid.col = 15
      vGrid.CellType = CellTypeEdit
      vGrid.col = 16
      vGrid.CellType = CellTypeEdit
      


  Case 16 'Disponibles
      vGrid.MaxCols = 8
      vGrid.SetText 1, 0, "Cédula"
      vGrid.SetText 2, 0, "d Alterno"
      vGrid.SetText 3, 0, "Nombre"
      vGrid.SetText 4, 0, "Ingreso"
      vGrid.SetText 5, 0, "Email"
      vGrid.SetText 6, 0, "Móvil"
      vGrid.SetText 7, 0, "Disponible"
      vGrid.SetText 8, 0, "Divisa"
      
      
      vHeaders.Columnas = 8
      vHeaders.Headers(1) = "Cédula"
      vHeaders.Headers(2) = "Id Alterno"
      vHeaders.Headers(3) = "Nombre"
      vHeaders.Headers(4) = "Ingreso"
      vHeaders.Headers(5) = "Email"
      vHeaders.Headers(6) = "Móvil"
      vHeaders.Headers(7) = "Disponible"
      vHeaders.Headers(8) = "Divisa"
      
      
      vGrid.col = 1
      vGrid.CellType = CellTypeEdit
      vGrid.col = 2
      vGrid.CellType = CellTypeEdit
      vGrid.col = 3
      vGrid.CellType = CellTypeEdit
      vGrid.col = 4
      vGrid.CellType = CellTypeDate
      vGrid.col = 5
      vGrid.CellType = CellTypeEdit
      vGrid.col = 6
      vGrid.CellType = CellTypeEdit
      vGrid.col = 7
      vGrid.CellType = CellTypeNumber
      vGrid.col = 8
      vGrid.CellType = CellTypeEdit
      

  Case 17 'Listado de Fiadores
      vGrid.MaxCols = 16
      vGrid.SetText 1, 0, "Operación"
      vGrid.SetText 2, 0, "Línea"
      vGrid.SetText 3, 0, "F. Id"
      vGrid.SetText 4, 0, "F. Nombre"
      vGrid.SetText 5, 0, "F. Estado"
      vGrid.SetText 6, 0, "F. Est.Lab."
      vGrid.SetText 7, 0, "F. Empresa - Corto"
      vGrid.SetText 8, 0, "F. Empresa - Largo"
      vGrid.SetText 9, 0, "Cuota"
      vGrid.SetText 10, 0, "Saldo"
      vGrid.SetText 11, 0, "Antiguedad"
      vGrid.SetText 12, 0, "Mora Financiera"
      vGrid.SetText 13, 0, "Deudor Id"
      vGrid.SetText 14, 0, "Deudor Nombre"
      vGrid.SetText 15, 0, "Deudor Estado"
      vGrid.SetText 16, 0, "Línea Desc"
  
      vHeaders.Columnas = 16
      vHeaders.Headers(1) = "Operación"
      vHeaders.Headers(2) = "Línea"
      vHeaders.Headers(3) = "F. Id"
      vHeaders.Headers(4) = "F. Nombre"
      vHeaders.Headers(5) = "F. Estado"
      vHeaders.Headers(6) = "F. Est.Lab."
      vHeaders.Headers(7) = "F. Empresa - Corto"
      vHeaders.Headers(8) = "F. Empresa - Largo"
      vHeaders.Headers(9) = "Cuota"
      vHeaders.Headers(10) = "Saldo"
      vHeaders.Headers(11) = "Antiguedad"
      vHeaders.Headers(12) = "Mora Financiera"
      vHeaders.Headers(13) = "Deudor Id"
      vHeaders.Headers(14) = "Deudor Nombre"
      vHeaders.Headers(15) = "Deudor Estado"
      vHeaders.Headers(16) = "Línea Desc"


End Select

End Sub

Private Sub sbProcesa(pSQL As String, pGridCol As Integer)


On Error GoTo vError

Me.MousePointer = vbHourglass
    
tcMain.Item(1).Selected = True
    
lblEstatus.Caption = "Procesando, Espere!"
DoEvents

Call sbvGridCol(pGridCol)

Call sbCargaGrid(vGrid, vGrid.MaxCols, pSQL)
vGrid.MaxRows = vGrid.MaxRows - 1

Call sbvGridColSize

lblEstatus.Caption = ""

Me.MousePointer = vbDefault


MsgBox "Informe Procesado!", vbInformation

'Exporta a Excel Automático
Call btnResultado_Click
Exit Sub

vError:
  Me.MousePointer = vbDefault


End Sub



Private Sub cmdGenera_Click()
Dim strSQL As String, x As Long


If Not fxValidaDatos Then Exit Sub
    
On Error GoTo vError
        
    
Select Case lblReporte.Tag
 Case "H00" 'Estadistica de Afiliacion
   
   
    strSQL = "exec spAfi_History_Afiliacion '" & Format(dtpHistory(0).Value, "yyyy-MM-dd") & " 00:00:00'" _
           & ", '" & Format(dtpHistory(1).Value, "yyyy-MM-dd") & " 23:59:59'"
    Call sbProcesa(strSQL, 0)

 Case "H01" 'Estadistica de Liquidacion
   
   
    strSQL = "exec spAfi_History_Liquidacion '" & Format(dtpHistory(0).Value, "yyyy-MM-dd") & " 00:00:00'" _
           & ", '" & Format(dtpHistory(1).Value, "yyyy-MM-dd") & " 23:59:59'"
    Call sbProcesa(strSQL, 1)


 Case "x00" 'Proyección de Cartera Mensual
   
   
    Call sbRecuperacionCartera(False, 0)
   
'    strSQL = "exec spCrdProyectaCartera '" & Format(dtpProyectaInicio.Value, "yyyy-MM-dd") & "', " & txtProyeccion.Text & ", 'M'"
'    Call sbProcesa(strSQL, 5)
   

 Case "x01" 'Proyección de Cartera Anual
   
   
    strSQL = "exec spCrdProyectaCartera '" & Format(dtpProyectaInicio.Value, "yyyy-MM-dd") & "', " & txtProyeccion.Text & ", 'A'"
    Call sbProcesa(strSQL, 5)
   
 
 Case "x02" 'Tasas y Plazos Ponderados
   If cboInstitucion.Text = "TODOS" Then
     Call sbTasasPlazoPond(0)
   Else
     Call sbTasasPlazoPond(cboInstitucion.ItemData(cboInstitucion.ListIndex))
   End If
 
 Case "x03" 'Proyección de Retenciones
   If cboInstitucion.Text = "TODOS" Then
     Call sbRecuperacionCartera(True)
   Else
     Call sbRecuperacionCartera(True, cboInstitucion.ItemData(cboInstitucion.ListIndex))
   End If

 
 Case "x04" 'Endeudamiento (General Detallado)
   Call sbEndeudamiento
   
 Case "x05" 'Antiguedad Persona (Ahorros
   Call sbAntiguedadPersonaAhorro
 
 Case "x06" 'Antiguedad de Saldos Personas vrs Ahorros
   Call sbAntiguedadPersonaSaldos
 

 Case "x08" 'Análisis de Endeudamiento
    
    strSQL = "exec spSIFEstudioEndeuda"
    Call sbProcesa(strSQL, 6)
   

 Case "x09" 'Análisis de Personas/Membresía/Saldos
   
    strSQL = "exec spSIFEstudioPersonasSaldos"
    Call sbProcesa(strSQL, 7)
    

 Case "x10" 'Analisis de Patrimonio x Membresia
   
    strSQL = "exec spSIFEstudioPatrimonioMemb"
    Call sbProcesa(strSQL, 8)
   
 Case "x11" 'Analisis de Patrimonio x Institucion
   
    strSQL = "exec spSIFEstudioPatrimonioInst"
    Call sbProcesa(strSQL, 9)
    

 Case "x12" 'Analisis de Personas / Edades / Patrimonio / Deuda
   
    strSQL = "exec spSIFEstudioEdadesAsociados"
    Call sbProcesa(strSQL, 7)
   

 Case "x13" 'Personas Sin Deudas
   Call sbListaPersonasvsDeuda(False)
 
 Case "x13.2" 'Personas con Deudas
   Call sbListaPersonasvsDeuda(True)

 Case "x13.3" 'Listado de Fiadores
   
   
    strSQL = "exec spCrd_Listado_Fiadores 'A'"
    Call sbProcesa(strSQL, 17)



 Case "x14" 'Cartera x Categoria Crediticia
   
    strSQL = "exec spSIFEstudioCarteraCategoria"
    Call sbProcesa(strSQL, 10)
   

 Case "x15" 'Frecuencia de Creditos x Persona
   
   x = InputBox("Digite la frecuencia de creditos otorgados por persona: ", "Frecuencias de Otorgamiento de Credito", 5)
   
    strSQL = "exec spSIFEstudioFrecuenciaCreditos 6," & x
    Call sbProcesa(strSQL, 15)
        


 Case "x16.1", "x16.2", "x16.3" 'Disponibles
        
        Select Case lblReporte.Tag
            Case "x16.1"
                strSQL = "select * from vCrd_Disponible_List_sAhorros"
            Case "x16.2"
                strSQL = "select * from vCrd_Disponible_List_sExcedentes"
            Case "x16.3"
                strSQL = "select * from vCrd_Disponible_List_sFondos"
        End Select
        
        Call sbProcesa(strSQL, 16)

End Select


lblEstatus.Caption = ""
PrgBar.Max = 100000000
PrgBar.Value = 1

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaLsw(vRetencion As Boolean)
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

Me.MousePointer = vbHourglass

If vRetencion Then
    strSQL = "select Cat.codigo,Cat.descripcion,Count(*) as Operaciones, Sum(Reg.Saldo) as Saldo" _
           & " from catalogo Cat inner join Reg_Creditos Reg on Cat.codigo = Reg.codigo and Cat.Retencion = 'S' or Cat.Poliza = 'S'"
Else
    strSQL = "select Cat.codigo,Cat.descripcion,Count(*) as Operaciones, Sum(Reg.Saldo) as Saldo" _
           & " from catalogo Cat inner join Reg_Creditos Reg on Cat.codigo = Reg.codigo and Cat.Retencion = 'N' and Cat.Poliza = 'N'"

End If

strSQL = strSQL & " where Reg.Estado = 'A'"

If chkInternas.Value = vbChecked Then
  strSQL = strSQL & " and Cat.linea_interna = 1"
End If

If cboCartera.Text <> "TODOS" Then
    strSQL = strSQL & " and Cat.Codigo in(select codigo from CBR_CLASIFICACION_DETALLE where COD_CLASIFICACION = '" _
           & cboCartera.ItemData(cboCartera.ListIndex) & "')"
End If



strSQL = strSQL & " group by Cat.codigo,Cat.descripcion"

If chkActivas.Value = vbChecked Then
  strSQL = strSQL & " having Sum(Reg.Saldo)  > 1"
End If

strSQL = strSQL & " order by Cat.codigo"

Call OpenRecordSet(rs, strSQL, 0)

lsw.ListItems.Clear
Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!Codigo)
       itmX.SubItems(1) = Trim(rs!DESCRIPCION)
       itmX.SubItems(2) = Format(rs!Saldo, "Standard")
       itmX.SubItems(3) = Format(rs!Operaciones, "###,###,###")
       
       itmX.Checked = chkTodos.Value
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

vGrid.AppearanceStyle = fxGridStyle

dtpProyectaInicio.Value = fxFechaServidor

dtpHistory(1).Value = dtpProyectaInicio.Value
dtpHistory(0).Value = DateAdd("m", -12, dtpHistory(1).Value)

lblEstatus.Caption = ""
lswRep.ColumnHeaders.Add , , "Seleccionar:", lswRep.Width - 50

With lsw.ColumnHeaders
   .Clear
   .Add , , "Código", 1100, vbCenter
   .Add , , "Descripción", 4100
   .Add , , "Saldos", 2400, vbRightJustify
   .Add , , "Operaciones", 1300, vbCenter
End With

strSQL = "select cod_institucion as Idx,descripcion as ItmX from instituciones"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)


strSQL = "select rtrim(cod_estado) as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  afi_Estados_Persona"
Call sbCbo_Llena_New(cboEstado, strSQL, True, True)

vPaso = True
    strSQL = "select rtrim(COD_CLASIFICACION) as 'IdX', rtrim(descripcion) as ItmX from CBR_CLASIFICACION_CARTERA" _
           & " order by COD_CLASIFICACION"
    Call sbCbo_Llena_New(cboCartera, strSQL, True, True)
vPaso = False

tcMain.Item(0).Selected = True
vGrid.MaxRows = 0

gbHistory.Visible = False

lswRep.ListItems.Clear
lswRep.ListItems.Add , "H00", "Estadística de Afiliación"
lswRep.ListItems.Add , "H01", "Estadística de Liquidaciones"
lswRep.ListItems.Add , "x00", "Proyección de Cartera Mensual"
lswRep.ListItems.Add , "x01", "Proyección de Cartera Anual"
lswRep.ListItems.Add , "x02", "Tasas y Plazos Ponderados"
lswRep.ListItems.Add , "x03", "Proyección de Retenciones"
lswRep.ListItems.Add , "x04", "Endeudamiento (General Detallado)"
lswRep.ListItems.Add , "x05", "Antiguedad Persona (Ahorros)"
lswRep.ListItems.Add , "x06", "Antiguedad de Saldos Personas vrs Ahorros"
lswRep.ListItems.Add , "x07", "Creditos Activos vrs Ahorros"

lswRep.ListItems.Add , "x16.1", "Disponible Garantía Sobre Ahorros"
lswRep.ListItems.Add , "x16.2", "Disponible Garantía s/Excedentes"
lswRep.ListItems.Add , "x16.3", "Disponible Garantía Planes de Ahorros"

lswRep.ListItems.Add , "x08", "Análisis de Endeudamiento"
lswRep.ListItems.Add , "x09", "Análisis de Personas/Membresía/Saldos"
lswRep.ListItems.Add , "x10", "Análisis de Patrimonio x Membresía"
lswRep.ListItems.Add , "x11", "Análisis de Patrimonio x Institución"
lswRep.ListItems.Add , "x12", "Análisis de Personas Edades/Patrimonio/Saldos"
lswRep.ListItems.Add , "x13", "Listado de Personas sin Deudas"
lswRep.ListItems.Add , "x13.2", "Listado de Personas con Deudas"
lswRep.ListItems.Add , "x13.3", "Listado de Fiadores Activos"


lswRep.ListItems.Add , "x14", "Cartera x Categoria Crediticia"
lswRep.ListItems.Add , "x15", "Frecuencias de Créditos x Persona"


End Sub


Private Sub lswRep_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError
    
lblReporte.Tag = Item.Key
lblReporte.Caption = Item.Text

tcMain.Item(0).Selected = True

gbHistory.Visible = False

Select Case lblReporte.Tag
  Case "H00", "H01" 'Estadisticas Historicas
    gbHistory.Visible = True
  
  Case "x00" 'Proyección de Cartera Mensual Creditos
    lblProyecta.Caption = "Proyección de Líneas en Meses ?"
    txtProyeccion.Text = 12
    Call sbCargaLsw(False)
  
  Case "x01" 'Proyección de Cartera Anual
    lblProyecta.Caption = "Proyección de Líneas en Años ?"
    txtProyeccion.Text = 5
    Call sbCargaLsw(False)
    
  Case "x02" 'Tasas y Plazos Ponderados
    Call sbCargaLsw(False)
 
  Case "x03" 'Proyección de Cartera Mensual Retenciones
    lblProyecta.Caption = "Proyección de Líneas en Meses ?"
    txtProyeccion.Text = 6
    Call sbCargaLsw(True)
  
  Case Else

End Select


vError:

End Sub

Private Sub Timer1_Timer()
lblReporte.Tag = "x00"
lblReporte.Caption = "Proyección de Cartera Mensual"

Timer1.Interval = 0
Timer1.Enabled = False

End Sub

