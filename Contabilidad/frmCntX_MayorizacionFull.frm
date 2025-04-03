VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmCntX_MayorizacionFull 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mayorización/Reversión en Lote de Asientos"
   ClientHeight    =   4404
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4404
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.RadioButton Opt 
      Height          =   372
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   1920
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Periodo"
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
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdProcesar 
      Height          =   492
      Left            =   6360
      TabIndex        =   1
      Top             =   3876
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Procesar"
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
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmCntX_MayorizacionFull.frx":0000
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   2760
      TabIndex        =   2
      Top             =   1440
      Width           =   4092
      _Version        =   1245187
      _ExtentX        =   7218
      _ExtentY        =   550
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.RadioButton Opt 
      Height          =   372
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Top             =   2400
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Fechas"
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
   End
   Begin XtremeSuiteControls.RadioButton Opt 
      Height          =   372
      Index           =   2
      Left            =   720
      TabIndex        =   6
      Top             =   2880
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Tipo de Asiento"
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
   End
   Begin XtremeSuiteControls.RadioButton Opt 
      Height          =   372
      Index           =   3
      Left            =   720
      TabIndex        =   7
      Top             =   3360
      Width           =   1692
      _Version        =   1245187
      _ExtentX        =   2984
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Tipo + Fechas"
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
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   2760
      TabIndex        =   9
      Top             =   2400
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
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
      Left            =   4920
      TabIndex        =   10
      Top             =   2400
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   2760
      TabIndex        =   11
      Top             =   2880
      Width           =   4092
      _Version        =   1245187
      _ExtentX        =   7218
      _ExtentY        =   550
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
      Appearance      =   2
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Left            =   2760
      TabIndex        =   12
      Top             =   1920
      Width           =   3972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo aplicación:"
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
      Left            =   720
      TabIndex        =   8
      Top             =   1440
      Width           =   1692
   End
   Begin XtremeShortcutBar.ShortcutCaption lbl 
      Height          =   624
      Left            =   0
      TabIndex        =   3
      Top             =   3840
      Width           =   12732
      _Version        =   1245187
      _ExtentX        =   22458
      _ExtentY        =   1101
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
      VisualTheme     =   6
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recuerde que para aplicación en Lote de asientos  se recomienda que ningún usuario esté modificando asientos al mismos tiempo."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1880
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Image imgBanner 
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frmCntX_MayorizacionFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProcesar_Click()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case True
  Case opt.Item(0).Value 'Periodo
        lbl.Caption = "Mayorizando (Espere!)..."
        
        strSQL = "exec spCntX_AsientosAplicacionLote_Todo " & gCntX_Parametros.CodigoConta _
               & "," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes _
               & ",'" & Mid(cbo.Text, 1, 1) & "','" & glogon.Usuario & "'"
  Case opt.Item(1).Value 'Fechas
        lbl.Caption = "Reversando (Espere!)..."
        
        strSQL = "exec spCntX_AsientosAplicacionLote_Fechas " & gCntX_Parametros.CodigoConta _
               & "," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes _
               & ",'" & Mid(cbo.Text, 1, 1) & "','" & glogon.Usuario & "','" _
               & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00','" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"


  Case opt.Item(2).Value 'Tipo de Asiento
        lbl.Caption = "Mayorizando (Espere!)..."
        
        strSQL = "exec spCntX_AsientosAplicacionLote_TipoAsiento " & gCntX_Parametros.CodigoConta _
               & "," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes _
               & ",'" & Mid(cbo.Text, 1, 1) & "','" & glogon.Usuario & "','" & cboTipo.ItemData(cboTipo.ListIndex) & "'"


  Case opt.Item(3).Value 'Tipo de Asiento + Fechas
        lbl.Caption = "Mayorizando (Espere!)..."
        
        strSQL = "exec spCntX_AsientosAplicacionLote_TipoAsientoFechas " & gCntX_Parametros.CodigoConta _
               & "," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes _
               & ",'" & Mid(cbo.Text, 1, 1) & "','" & glogon.Usuario & "','" & cboTipo.ItemData(cboTipo.ListIndex) & "','" _
               & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00','" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

End Select

Call ConectionExecute(strSQL, 0)

Me.MousePointer = vbDefault
lbl.Caption = "Aplicación en LOTE Completa..."

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 20
End Sub

Private Sub Form_Load()
Dim vFecha  As Date, strSQL As String
Dim vPeriodoDesc As String

vModulo = 20

On Error GoTo vError

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture


vPeriodoDesc = fxCntX_PeriodoDesc(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes)

lblX.Caption = vPeriodoDesc


vFecha = CDate(gCntX_Parametros.PeriodoAnio & "/" & Format(gCntX_Parametros.PeriodoMes, "00") & "/01")

dtpInicio.MinDate = vFecha
dtpCorte.MinDate = vFecha

vFecha = DateAdd("d", -1, DateAdd("m", 1, vFecha))

dtpInicio.MaxDate = vFecha
dtpCorte.MaxDate = vFecha

cbo.Clear
cbo.AddItem "Mayorización de Asientos"
cbo.AddItem "Reversión de Asientos"
cbo.Text = "Mayorización de Asientos"


strSQL = "select rtrim(Tipo_Asiento) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from CntX_Tipos_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta

Call sbCbo_Llena_New(cboTipo, strSQL, False, True)

Call Formularios(Me)
Call RefrescaTags(Me)


dtpInicio.Value = CDate(gCntX_Parametros.PeriodoAnio & "/" & Format(gCntX_Parametros.PeriodoMes, "00") & "/" & Format(Day(fxFechaServidor), "00"))
dtpCorte.Value = dtpInicio.Value

vError:

End Sub

Private Sub opt_Click(Index As Integer)

dtpInicio.Enabled = opt.Item(1).Value
dtpCorte.Enabled = opt.Item(1).Value
cboTipo.Enabled = opt.Item(2).Value

If opt.Item(3).Value Then
    dtpInicio.Enabled = opt.Item(3).Value
    dtpCorte.Enabled = opt.Item(3).Value
    cboTipo.Enabled = opt.Item(3).Value
End If

End Sub
