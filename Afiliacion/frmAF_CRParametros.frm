VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmAF_CRParametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control de Renuncias"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.RadioButton opt 
      Height          =   252
      Index           =   0
      Left            =   2640
      TabIndex        =   8
      Top             =   1800
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Dias de Ingreso"
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
      Appearance      =   16
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkZonas 
      Height          =   252
      Left            =   720
      TabIndex        =   5
      Top             =   2640
      Width           =   5052
      _Version        =   1441793
      _ExtentX        =   8911
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Utiliza Segregación de Funciones x Zonas (Seguimiento) ?"
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
      Appearance      =   16
      Alignment       =   1
   End
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   612
      Left            =   5400
      TabIndex        =   2
      Top             =   3960
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Guardar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmAF_CRParametros.frx":0000
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   4440
      TabIndex        =   3
      Top             =   2160
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtDias 
      Height          =   312
      Left            =   4440
      TabIndex        =   4
      Top             =   1800
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
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
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkAportePatronal 
      Height          =   252
      Left            =   720
      TabIndex        =   6
      Top             =   3000
      Width           =   5052
      _Version        =   1441793
      _ExtentX        =   8911
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Desea dar Seguimiento a Renuncias Patronales ?"
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
      Appearance      =   16
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkActivar 
      Height          =   252
      Left            =   720
      TabIndex        =   7
      Top             =   3360
      Width           =   5052
      _Version        =   1441793
      _ExtentX        =   8911
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Activar Proceso de Control de Renuncias (Liquidaciones)"
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
      Appearance      =   16
      Alignment       =   1
   End
   Begin XtremeSuiteControls.RadioButton opt 
      Height          =   252
      Index           =   1
      Left            =   2640
      TabIndex        =   9
      Top             =   2160
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Fecha Corte"
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
      Appearance      =   16
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Renuncias: Parámetros de Integración"
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
      Height          =   735
      Index           =   0
      Left            =   1880
      TabIndex        =   1
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Vencimiento Renuncias:"
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
      Height          =   252
      Index           =   1
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   2652
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_CRParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbCargaDatos()
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select * from afi_cr_parametros"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtDias = rs!dias_vence & ""
  dtpCorte.Value = rs!fecha_limite
  If rs!tipo_vencimiento = "D" Then
     opt.Item(0).Value = True
  Else
     opt.Item(1).Value = True
  End If
  
  chkAportePatronal.Value = rs!liq_pat_control
  chkZonas.Value = rs!utiliza_zonas
  chkActivar.Value = rs!activar_control
End If
rs.Close

Call opt_Click(0)

End Sub

Private Sub cmdGuardar_Click()
Dim strSQL As String

On Error GoTo vError

If opt.Item(0).Value = True Then
  strSQL = "D"
Else
  strSQL = "F"
End If

strSQL = "update afi_cr_parametros set dias_vence = " & txtDias _
       & ",liq_pat_control = " & chkAportePatronal.Value _
       & ",tipo_vencimiento = '" & strSQL & "'" _
       & ",fecha_limite = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'" _
       & ",utiliza_zonas = " & chkZonas.Value _
       & ",activar_control = " & chkActivar.Value
Call ConectionExecute(strSQL)

MsgBox "Parámetros Actualizados Satisfactoriamente...", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "select isnull(count(*),0) as Existe from afi_cr_parametros"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
   strSQL = "insert afi_cr_parametros(dias_vence,liq_pat_control,fecha_limite,tipo_vencimiento" _
          & ",utiliza_zonas,activar_control) values(20,0,dbo.MyGetdate(),'D',0,0)"
   Call ConectionExecute(strSQL)
End If
rs.Close

Call sbCargaDatos

Call Formularios(Me)
Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub opt_Click(Index As Integer)
txtDias.Enabled = False
dtpCorte.Enabled = False

Select Case True
  Case opt.Item(0).Value
     txtDias.Enabled = True
  Case opt.Item(1).Value
     dtpCorte.Enabled = True
End Select
End Sub
