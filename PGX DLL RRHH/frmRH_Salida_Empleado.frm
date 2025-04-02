VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmRH_Salida_Empleado 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RRHH: Salidas"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   10080
      Top             =   1800
   End
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   615
      Left            =   7800
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
      _Version        =   1310723
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Aplicar"
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
      Picture         =   "frmRH_Salida_Empleado.frx":0000
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
      Width           =   6855
      _Version        =   1310723
      _ExtentX        =   12091
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   915
      Left            =   2520
      TabIndex        =   5
      Top             =   2640
      Width           =   6855
      _Version        =   1310723
      _ExtentX        =   12086
      _ExtentY        =   1609
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   315
      Left            =   2520
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   315
      Left            =   6960
      TabIndex        =   7
      Top             =   3720
      Width           =   2415
      _Version        =   1310723
      _ExtentX        =   4260
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.DateTimePicker dtpHasta 
      Height          =   315
      Left            =   2520
      TabIndex        =   12
      Top             =   4080
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
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
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   15
      Top             =   5400
      Width           =   6855
      _Version        =   1310723
      _ExtentX        =   12091
      _ExtentY        =   450
      _StockProps     =   14
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      ForeColor       =   4210752
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   14
      Top             =   5400
      Width           =   3855
      _Version        =   1310723
      _ExtentX        =   6800
      _ExtentY        =   450
      _StockProps     =   14
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      ForeColor       =   4210752
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Trabaja Hasta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   13
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Notas"
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
      Index           =   4
      Left            =   1320
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Salida"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Presenta"
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
      Left            =   840
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label15 
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
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   8
      Top             =   3720
      Width           =   1095
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   8175
      _Version        =   1310723
      _ExtentX        =   14420
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "NOMBRE_COMPLETO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
      _Version        =   1310723
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "CEDULA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de la Tipo de Salida del Empleado"
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
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   7935
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "frmRH_Salida_Empleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim mEmpleadoId As String


Private Sub btnAplicar_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spRH_Persona_Salida_Registro '" & mEmpleadoId & "','" & Mid(cboTipo.Text, 1, 1) _
        & "','" & Mid(cboEstado.Text, 1, 1) & "','" & txtNotas.Text _
        & "','" & Format(dtpFecha.Value, "yyyy-MM-dd") & "','" & Format(dtpHasta.Value, "yyyy-MM-dd") _
        & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Salida, Empleado Id: " & mEmpleadoId & ", Tipo: " & cboTipo.Text & ", Estado: " & cboEstado.Text)

Me.MousePointer = vbDefault

MsgBox "Salida Registrada Satisfactoriamente!", vbInformation

UnLoad Me

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()
vModulo = 23

On Error GoTo vError

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

mEmpleadoId = GLOBALES.gTag

cboTipo.Clear
cboTipo.AddItem "Despido"
cboTipo.AddItem "Renuncia"
cboTipo.AddItem "Terminación"
cboTipo.Text = "Renuncia"

cboEstado.Clear
cboEstado.AddItem "Proceso"
cboEstado.AddItem "Descartada"
cboEstado.Text = "Proceso"


Call Formularios(Me)
Call RefrescaTags(Me)

Exit Sub

vError:

End Sub

Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

Call sbInicializa

End Sub


Private Sub sbInicializa()
  
On Error GoTo vError

strSQL = "exec spRH_Persona_Salida_Consulta '" & mEmpleadoId & "'"
Call OpenRecordSet(rs, strSQL)

scMain.Item(0).Caption = rs!IDENTIFICACION
scMain.Item(1).Caption = rs!NOMBRE_COMPLETO

If rs!SALIDA_ESTADO_DESC = "Liquidado" Then
    cboEstado.AddItem "Liquidado"
    btnAplicar.Visible = False
End If

cboTipo.Text = rs!SALIDA_TIPO_DESC
cboEstado.Text = rs!SALIDA_ESTADO_DESC

If IsNull(rs!SALIDA_FECHA_PRESENTA) Then
    dtpFecha.Value = rs!Fecha_Server
Else
    dtpFecha.Value = rs!SALIDA_FECHA_PRESENTA
End If

If IsNull(rs!SALIDA_FECHA) Then
    dtpHasta.Value = rs!Fecha_Server
Else
    dtpHasta.Value = rs!SALIDA_FECHA
End If


txtNotas.Text = RTrim(rs!SALIDA_NOTAS & "")


scMain(2).ToolTipText = "Fecha de Registro de la Salida"
scMain(3).ToolTipText = "Usuario que Registra de la Salida"

scMain(2).Caption = rs!SALIDA_REGISTRO_FECHA & ""
scMain(3).Caption = rs!SALIDA_REGISTRO_USUARIO & ""

txtNotas.SetFocus

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub
