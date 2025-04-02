VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_CR_NoAumentoTasas_Autorizacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Autorización de Renuncias con No Aumento de Tasas"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   13455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   2910
      Width           =   13215
      _Version        =   1441793
      _ExtentX        =   23310
      _ExtentY        =   9975
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
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   720
   End
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Todas las fechas"
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
   Begin XtremeSuiteControls.CheckBox chkTodas 
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   2265
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2535
      Width           =   13215
      _Version        =   1441793
      _ExtentX        =   23310
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   360
      Left            =   12840
      TabIndex        =   4
      Top             =   2160
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   635
      _StockProps     =   79
      Appearance      =   6
      Picture         =   "frmAF_CR_NoAumentoTasas_Autorizacion.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   330
      Left            =   6240
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin XtremeSuiteControls.PushButton btnAutorizar 
      Height          =   495
      Left            =   11640
      TabIndex        =   7
      Top             =   8760
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Autorizar"
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
      Picture         =   "frmAF_CR_NoAumentoTasas_Autorizacion.frx":016A
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   120
      TabIndex        =   8
      Top             =   8520
      Visible         =   0   'False
      Width           =   13215
      _Version        =   1441793
      _ExtentX        =   23310
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboUsuarios 
      Height          =   330
      Left            =   3840
      TabIndex        =   10
      Top             =   1440
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   960
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   582
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
      Height          =   330
      Left            =   2280
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   8640
      TabIndex        =   16
      Top             =   1440
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmAF_CR_NoAumentoTasas_Autorizacion.frx":0891
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   15
      Top             =   1200
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario Específico"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo de Usuarios"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   9
      Top             =   1200
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
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
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   13215
      _Version        =   1441793
      _ExtentX        =   23310
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Renuncias con No Aumento de Tasas - Pendientes de Autorización"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Autorización de Renuncias con NO Aumento de Tasas"
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
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
   Begin VB.Image imgBanner 
      Height          =   1125
      Left            =   0
      Top             =   0
      Width           =   15135
   End
End
Attribute VB_Name = "frmAF_CR_NoAumentoTasas_Autorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean


Private Sub sbListas_Load()

On Error GoTo vError

Dim pFI As String, pFC As String

Me.MousePointer = vbHourglass

If chkFechas.Value = xtpChecked Then
    pFI = "1900-01-01 00:00:00"
    pFC = "2300-01-01 23:59:59"
Else
    pFI = Format(dtpInicio.Value, "yyyy-mm-dd") & " 00:00:00"
    pFC = Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59"
End If

strSQL = "exec spAFI_Renuncias_NAT_Control_Consulta '" & pFI & "', '" & pFC & "', '" & Mid(cboUsuarios.Text, 1, 1) _
       & "', '" & txtFiltro.Text & "', '" & txtUsuario.Text & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
    
With lsw.ListItems
    .Clear
    Do While Not rs.EOF
        Set itmX = .Add(, , rs!cod_Renuncia)
            itmX.SubItems(1) = rs!Cedula
            itmX.SubItems(2) = rs!Nombre
            itmX.SubItems(3) = Format(rs!Registro_Fecha, "yyyy-mm-dd")
            itmX.SubItems(4) = Format(rs!Vencimiento, "yyyy-mm-dd")
            itmX.SubItems(5) = rs!Tipo
            itmX.SubItems(6) = rs!Estado_Desc
            itmX.SubItems(7) = rs!Causa_Desc
            itmX.SubItems(8) = rs!Registro_Usuario & ""
            
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




Private Sub btnAutorizar_Click()
Dim i As Long, pEstado As String, pRefresh As Integer, pNota As String

On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

strSQL = ""
'spAFI_Renuncia_NAT_Tag_Autoriza(@RenunciaId  int = NULL, @Usuario  varchar(30) = NULL)

With lsw.ListItems
    For i = 1 To .Count
        If .Item(i).Checked Then
            strSQL = strSQL & Space(10) & "exec spAFI_Renuncia_NAT_Tag_Autoriza " & .Item(i).Text & ", '" & glogon.Usuario & "'"
        End If
        
        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
    Next i

End With
   
 

'Ultimo Lote
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If


'Actualiza Lista
 Call sbListas_Load

ProgressBarX.Visible = False
Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBuscar_Click()
 Call sbListas_Load
End Sub

Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub chkFechas_Click()
If chkFechas.Value = xtpChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub chkTodas_Click()
Dim i As Long

With lsw.ListItems
   For i = 1 To .Count
       .Item(i).Checked = chkTodas.Value
   Next i
End With

End Sub

Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

lsw.Checkboxes = True
With lsw.ColumnHeaders
    .Clear
    .Add , , "Id Renuncia", 1800
    .Add , , "Cédula", 1500, vbCenter
    .Add , , "Nombre", 3150
    .Add , , "Fecha", 1800, vbCenter
    .Add , , "Vence", 1800, vbCenter
    .Add , , "Tipo", 1800, vbCenter
    .Add , , "Estado", 1800, vbCenter
    .Add , , "Causa", 3100
    .Add , , "Usuario", 3100
End With

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -30, dtpCorte.Value)

cboUsuarios.Clear
cboUsuarios.AddItem "TODOS"
cboUsuarios.AddItem "Cobros"
cboUsuarios.AddItem "Otros"
cboUsuarios.Text = "TODOS"


Call chkFechas_Click

Call Formularios(Me)

Call RefrescaTags(Me)

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Call sbListas_Load

End Sub

Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)

Call sbListas_Load

End Sub




