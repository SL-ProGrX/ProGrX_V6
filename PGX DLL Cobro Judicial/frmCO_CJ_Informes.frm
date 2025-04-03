VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCO_CJ_Informes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros: Informes"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10965
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   375
      Left            =   7440
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Todas"
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
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9600
      Top             =   2760
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   6000
      TabIndex        =   3
      Top             =   2640
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   550
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
      Left            =   6000
      TabIndex        =   4
      Top             =   3000
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   1710
      Width           =   4485
      _Version        =   1310723
      _ExtentX        =   7902
      _ExtentY        =   6371
      _StockProps     =   77
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
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   972
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   10812
      _Version        =   1310723
      _ExtentX        =   19071
      _ExtentY        =   1714
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   492
         Left            =   7560
         TabIndex        =   7
         Top             =   240
         Width           =   1572
         _Version        =   1310723
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Reporte"
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
         Picture         =   "frmCO_CJ_Informes.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnCubo 
         Height          =   492
         Left            =   9120
         TabIndex        =   8
         Top             =   240
         Width           =   1572
         _Version        =   1310723
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Cubo"
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
         Picture         =   "frmCO_CJ_Informes.frx":07BC
      End
      Begin XtremeSuiteControls.Label lblStatus 
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   4695
         _Version        =   1310723
         _ExtentX        =   8281
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Este proceso puede tardar varios minutos, espere el mensaje de proceso concluido."
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
   Begin XtremeSuiteControls.ComboBox cboProceso 
      Height          =   330
      Left            =   6000
      TabIndex        =   11
      Top             =   2160
      Width           =   4830
      _Version        =   1310723
      _ExtentX        =   8520
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   6000
      TabIndex        =   12
      Top             =   1800
      Width           =   4830
      _Version        =   1310723
      _ExtentX        =   8520
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboJuzgado 
      Height          =   330
      Left            =   6000
      TabIndex        =   21
      Top             =   3840
      Width           =   4830
      _Version        =   1310723
      _ExtentX        =   8520
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboUsuarios 
      Height          =   330
      Left            =   6000
      TabIndex        =   22
      Top             =   3480
      Width           =   4830
      _Version        =   1310723
      _ExtentX        =   8520
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboBufete 
      Height          =   330
      Left            =   6000
      TabIndex        =   23
      Top             =   4560
      Width           =   4830
      _Version        =   1310723
      _ExtentX        =   8520
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboJuicio 
      Height          =   330
      Left            =   6000
      TabIndex        =   24
      Top             =   4200
      Width           =   4830
      _Version        =   1310723
      _ExtentX        =   8520
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboAbogado 
      Height          =   330
      Left            =   6000
      TabIndex        =   25
      Top             =   4920
      Width           =   4830
      _Version        =   1310723
      _ExtentX        =   8520
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   8
      Left            =   4680
      TabIndex        =   20
      Top             =   4920
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Abogado"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Firma/Bufete"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   18
      Top             =   4200
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo Juicio"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   4680
      TabIndex        =   17
      Top             =   3840
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Juzgado"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   16
      Top             =   3480
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Proceso"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo"
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Cobros Judiciales"
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
      Height          =   312
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   4812
   End
   Begin XtremeShortcutBar.ShortcutCaption lblReporte 
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4452
      _Version        =   1310723
      _ExtentX        =   7853
      _ExtentY        =   656
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
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Left            =   4560
      TabIndex        =   0
      Top             =   1320
      Width           =   6372
      _Version        =   1310723
      _ExtentX        =   11239
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Filtros:"
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
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmCO_CJ_Informes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbCubo()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim vMensaje As String

On Error GoTo vError

Me.MousePointer = vbHourglass

lblStatus.Caption = "Procesando Información Espere!....Este proceso puede durar varios minutos."
lblStatus.Refresh

vMensaje = "Cobros_Judicial"

If chkFechas.Value = vbChecked Then
  vFechaInicio = "1900/01/01"
  vFechaCorte = fxFechaServidor
Else
  vFechaInicio = dtpInicio.Value
  vFechaCorte = dtpCorte.Value
End If

'strSQL = "exec spCbrControlRecuperacionAnalisisCubo '" & Format(vFechaInicio, "yyyy/mm/dd") & "','" & Format(dtpCorte, "yyyy/mm/dd") & "'"
'Call ConectionExecute(strSQL)

lblStatus.Caption = "Proceso Concluido con éxito, la información puede ser utilizada desde la base de datos de análisis, cubo: " & vMensaje

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnCubo_Click()
    lblStatus.Visible = True
    Call sbCubo
End Sub

Private Sub btnReporte_Click()

    lblStatus.Visible = False
    Call sbReportes
End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = vbUnchecked Then
  dtpInicio.Enabled = True
Else
  dtpInicio.Enabled = False
End If
  
dtpCorte.Enabled = dtpInicio.Enabled
  
End Sub



Private Sub sbReporteGestiones()
'Dim strSQL As String, vSubTitulo As String
'Dim i As Byte
'
'Me.MousePointer = vbHourglass
'
'Select Case Mid(cboEPersona.Text, 1, 2)
' Case "00" 'Todos
'   strSQL = ""
' Case "01" 'Socios
'   strSQL = "{SOCIOS.ESTADOACTUAL} = 'S'"
' Case "02" 'Opex
'   strSQL = "({SOCIOS.ESTADOACTUAL} = 'A' OR {SOCIOS.ESTADOACTUAL} = 'P')"
' Case "03" 'No Socios
'   strSQL = "{SOCIOS.ESTADOACTUAL} = 'N'"
' Case "04" 'Ren.Interna
'   strSQL = "{SOCIOS.ESTADOACTUAL} = 'A'"
' Case "05" 'Ren.Patronal
'   strSQL = "{SOCIOS.ESTADOACTUAL} = 'P'"
'End Select
'
'
'If cboUsuarios.Text <> "TODOS" Then
'  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'  strSQL = strSQL & "{CBR_USUARIOS.USUARIO} = '" & cboUsuarios.Text & "'"
'End If
'
'If cboGestion.Text <> "TODOS" Then
'  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'  strSQL = strSQL & "{CBR_GESTIONES.COD_GESTION} = '" & fxCodigoCbo(cboGestion) & "'"
'End If
'
'
'vSubTitulo = "Gestiones : " & cboGestion.Text & "  Estado : " & cboEPersona.Text _
'                 & "  Usuario : " & cboUsuarios.Text & "  Fechas: "
'
'
'If chkFechas.Value = vbChecked Then
'  vSubTitulo = vSubTitulo & " Todas"
'Else
'  vSubTitulo = vSubTitulo & " I." & Format(dtpInicio.Value, "dd/mm/yyyy") & " C." & Format(dtpCorte.Value, "dd/mm/yyyy")
'  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'  strSQL = strSQL & "CDATE({CBR_SEGUIMIENTO.FECHA}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ")" _
'                & " to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
'End If
'
'With frmContenedor.Crt
'    .Reset
'    .WindowShowGroupTree = True
'    .WindowShowPrintSetupBtn = True
'    .WindowShowRefreshBtn = True
'    .WindowShowSearchBtn = True
'    .WindowState = crptMaximized
'    .WindowTitle = "Reportes del Módulo de Cobro"
'
'    .Connect = glogon.ConectRPT
'
'  Select Case lblReporte.Tag
'   Case "01" 'Gestiones Realizadas
'        If cboTipo.Text = "Resumen" Then
'               .ReportFileName = SIFGlobal.fxPathReportes("CbrControlGestionesRealizadasRsm.rpt")
'               .Formulas(1) = "Titulo='GESTIONES REALIZADAS'"
'        Else
'               .ReportFileName = SIFGlobal.fxPathReportes("CbrControlGestionesRealizadas.rpt")
'               .Formulas(1) = "Titulo='GESTIONES REALIZADAS'"
'        End If
'        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
'        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
'        .SelectionFormula = strSQL
'
'   Case "02" 'Personas bajo Control
'   Case "03" 'Personas sin Control
'   Case "04" 'Gestiones x Usuarios
'        If cboTipo.Text = "Resumen" Then
'               .ReportFileName = SIFGlobal.fxPathReportes("CbrControlGestionesUsuariosRsm.rpt")
'               .Formulas(1) = "Titulo='GESTIONES REALIZADAS x USUARIOS'"
'        Else
'               .ReportFileName = SIFGlobal.fxPathReportes("CbrControlGestionesUsuarios.rpt")
'               .Formulas(1) = "Titulo='GESTIONES REALIZADAS x USUARIOS'"
'        End If
'        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
'        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
'        .SelectionFormula = strSQL
'
'  End Select
'
'    .PrintReport
'End With
'
'Me.MousePointer = vbDefault

End Sub

Private Sub sbReporteRecuperacion()
'Dim strSQL As String, vSubTitulo As String
'Dim i As Byte
'
'Me.MousePointer = vbHourglass
'
'Select Case Mid(cboEPersona.Text, 1, 2)
' Case "00" 'Todos
'   strSQL = ""
' Case "01" 'Socios
'   strSQL = "{vCBRControlRecuperacion.ESTADOACTUAL} = 'S'"
' Case "02" 'Opex
'   strSQL = "({vCBRControlRecuperacion.ESTADOACTUAL} = 'A' OR {vCBRControlRecuperacion.ESTADOACTUAL} = 'P')"
' Case "03" 'No Socios
'   strSQL = "{vCBRControlRecuperacion.ESTADOACTUAL} = 'N'"
' Case "04" 'Ren.Interna
'   strSQL = "{SOCIvCBRControlRecuperacionOS.ESTADOACTUAL} = 'A'"
' Case "05" 'Ren.Patronal
'   strSQL = "{SOCIOS.ESTADOACTUAL} = 'P'"
'End Select
'
'
'If cboUsuarios.Text <> "TODOS" Then
'  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'  strSQL = strSQL & "{vCBRControlRecuperacion.USUARIO} = '" & cboUsuarios.Text & "'"
'End If
'
'If cboGestion.Text <> "TODOS" Then
'  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'  strSQL = strSQL & "{vCBRControlRecuperacion.COD_GESTION} = '" & fxCodigoCbo(cboGestion) & "'"
'End If
'
'
'vSubTitulo = "Gestiones : " & cboGestion.Text & "  Estado : " & cboEPersona.Text _
'                 & "  Usuario : " & cboUsuarios.Text & "  Fechas: "
'
'
'If chkFechas.Value = vbChecked Then
'  vSubTitulo = vSubTitulo & " Todas"
'Else
'  vSubTitulo = vSubTitulo & " I." & Format(dtpInicio.Value, "dd/mm/yyyy") & " C." & Format(dtpCorte.Value, "dd/mm/yyyy")
'  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'  strSQL = strSQL & "CDATE({vCBRControlRecuperacion.FECHAGestion}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ")" _
'                & " to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
'End If
'
'With frmContenedor.Crt
'    .Reset
'    .WindowShowGroupTree = True
'    .WindowShowPrintSetupBtn = True
'    .WindowShowRefreshBtn = True
'    .WindowShowSearchBtn = True
'    .WindowState = crptMaximized
'    .WindowTitle = "Reportes del Módulo de Cobro"
'
'    .Connect = glogon.ConectRPT
'
'  Select Case lblReporte.Tag
'   Case "09" 'Recuperación x Gestión
'        If cboTipo.Text = "Resumen" Then
'               .ReportFileName = SIFGlobal.fxPathReportes("CbrControlRecuperacionPorGestion.rpt")
'               .Formulas(1) = "Titulo='RECUPERACION X GESTION'"
'        Else
'               .ReportFileName = SIFGlobal.fxPathReportes("CbrControlRecuperacionPorGestion.rpt")
'               .Formulas(1) = "Titulo='RECUPERACION X GESTION'"
'        End If
'        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
'        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
'        .SelectionFormula = strSQL
'
'   Case "10" 'Recuperación x Usuario
'
'
'   Case "11" 'Recuperación x Línea
'   Case "12" 'Recuperación x Garantía
'   Case "13" 'Recuperación Estadística
'
'
'
'
'  End Select
'
'    .PrintReport
'
'End With
'
'Me.MousePointer = vbDefault

End Sub

Private Sub sbReportes()

'Select Case lblReporte.Tag
'   Case "01" 'Gestiones Realizadas
'        Call sbReporteGestiones
'   Case "02" 'Personas bajo Control
'   Case "03" 'Personas sin Control
'   Case "04" 'Gestiones x Usuarios
'        Call sbReporteGestiones
'   Case "05" 'Comisiones x Gestión
'   Case "06" 'Comisiones x Usuario
'   Case "07" 'Cobro x Gestión
'   Case "08" 'Cobro x Usuario
'   Case "09" 'Recuperación x Gestión
'        Call sbReporteRecuperacion
'   Case "10" 'Recuperación x Usuario
'        Call sbReporteRecuperacion
'   Case "11" 'Recuperación x Línea
'        Call sbReporteRecuperacion
'   Case "12" 'Recuperación x Garantía
'        Call sbReporteRecuperacion
'   Case "13" 'Recuperación Estadística
'        Call sbReporteRecuperacion
'End Select


End Sub

Private Sub Form_Activate()
vModulo = 6

End Sub

Private Sub Form_Load()

vModulo = 6

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

lsw.ColumnHeaders.Add , , "", 4352

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
lblReporte.Caption = Item.Text
lblReporte.Tag = Item.Tag
End Sub


Private Sub TimerX_Timer()
Dim strSQL As String, itmX As ListViewItem

TimerX.Interval = 0

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

lblReporte.Tag = ""
lblReporte.Caption = ">>> Seleccione Un Reporte <<<"

cboTipo.Clear
cboTipo.AddItem "Detalle"
cboTipo.AddItem "Resumen"
cboTipo.Text = "Detalle"

strSQL = "select cod_proceso as 'IdX',  rtrim(descripcion) as ItmX" _
         & " from  cbr_cj_proceso where activo = 1"
Call sbCbo_Llena_New(cboProceso, strSQL, True, True)

strSQL = "select cod_juzgado as 'IdX', rtrim(nombre) as ItmX" _
         & " from  cbr_cj_juzgados where activo = 1"
Call sbCbo_Llena_New(cboJuzgado, strSQL, True, True)

strSQL = "select Tipo_Juicio as 'IdX', rtrim(Descripcion) as ItmX" _
         & " from  cbr_cj_Tipos_Juicios where activo = 1"
Call sbCbo_Llena_New(cboJuicio, strSQL, True, True)


strSQL = "select cod_bufete as 'IdX', rtrim(nombre) as ItmX" _
         & " from cbr_cj_bufetes where activo = 1"
Call sbCbo_Llena_New(cboBufete, strSQL, True, True)


strSQL = "select cod_abogado as 'IdX' ,rtrim(nombre) as ItmX" _
         & " from  cbr_cj_abogados where activo = 1"
Call sbCbo_Llena_New(cboAbogado, strSQL, True, True)


strSQL = "select usuario as 'Itmx', usuario as 'IdX' from cbr_usuarios"
Call sbCbo_Llena_New(cboUsuarios, strSQL, True, True)

With lsw.ListItems
  .Clear
  Set itmX = .Add(, , "Listado General")
      itmX.Tag = "01"
  Set itmX = .Add(, , "General por Línea de crédito")
      itmX.Tag = "02"
  Set itmX = .Add(, , "General por Garantía")
      itmX.Tag = "03"
  Set itmX = .Add(, , "General por Cartera de Cobro")
      itmX.Tag = "04"
  Set itmX = .Add(, , "Informe Estadístico")
      itmX.Tag = "05"
  Set itmX = .Add(, , "Informe por Proceso")
      itmX.Tag = "06"
  Set itmX = .Add(, , "Informe por Juzgado")
      itmX.Tag = "07"
  Set itmX = .Add(, , "Informe por Tipo de Juicio")
      itmX.Tag = "08"
  Set itmX = .Add(, , "Informe por Abogados")
      itmX.Tag = "09"
  Set itmX = .Add(, , "Informe por Bufete")
      itmX.Tag = "10"
  Set itmX = .Add(, , "Informe de Gastos Generados")
      itmX.Tag = "11"
  Set itmX = .Add(, , "Informe de Honorarios Aplicados")
      itmX.Tag = "12"
  Set itmX = .Add(, , "Informe de Recuperación")
      itmX.Tag = "13"
End With


End Sub


