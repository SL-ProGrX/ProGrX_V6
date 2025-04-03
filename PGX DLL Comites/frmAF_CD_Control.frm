VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmAF_CD_Control 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control y Seguimiento de Cuentas de Comités"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   13410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   3630
      Width           =   13215
      _Version        =   1572864
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
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
      _Version        =   1572864
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
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3255
      Width           =   13215
      _Version        =   1572864
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
      TabIndex        =   3
      Top             =   2880
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   635
      _StockProps     =   79
      Appearance      =   6
      Picture         =   "frmAF_CD_Control.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnAutorizar 
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Aprobar / Rechazar"
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
      Picture         =   "frmAF_CD_Control.frx":016A
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   120
      TabIndex        =   5
      Top             =   9240
      Visible         =   0   'False
      Width           =   13215
      _Version        =   1572864
      _ExtentX        =   23310
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   330
      Left            =   10920
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
      _Version        =   1572864
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
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
      _Version        =   1572864
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
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
      _Version        =   1572864
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
      Left            =   6480
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
      _Version        =   1572864
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
      Picture         =   "frmAF_CD_Control.frx":0891
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.ComboBox cboProceso 
      Height          =   330
      Left            =   10920
      TabIndex        =   15
      Top             =   1800
      Width           =   2295
      _Version        =   1572864
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
   Begin XtremeSuiteControls.FlatEdit txtComite 
      Height          =   330
      Left            =   3960
      TabIndex        =   18
      ToolTipText     =   "Presione F4"
      Top             =   1440
      Width           =   5175
      _Version        =   1572864
      _ExtentX        =   9128
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboEmite 
      Height          =   330
      Left            =   3960
      TabIndex        =   19
      Top             =   1800
      Width           =   2295
      _Version        =   1572864
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
   Begin XtremeSuiteControls.FlatEdit txtTesoreriaId 
      Height          =   330
      Left            =   7560
      TabIndex        =   21
      Top             =   1800
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   2400
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Registro"
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
      Picture         =   "frmAF_CD_Control.frx":0F91
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   23
      Top             =   2400
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Pre-Cálculo"
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
      Picture         =   "frmAF_CD_Control.frx":1699
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   5
      Left            =   6360
      TabIndex        =   20
      Top             =   1800
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Id Tesorería"
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
      Index           =   4
      Left            =   9720
      TabIndex        =   17
      Top             =   1800
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   3
      Left            =   9720
      TabIndex        =   16
      Top             =   1440
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Estado"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Control y Seguimiento de Cuentas de Comités Sede"
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
      TabIndex        =   14
      Top             =   360
      Width           =   6255
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   13215
      _Version        =   1572864
      _ExtentX        =   23310
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Casos Localizados"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   975
      _Version        =   1572864
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Comités"
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
      Index           =   2
      Left            =   2880
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo Desemb"
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
   Begin VB.Image imgBanner 
      Height          =   1125
      Left            =   0
      Top             =   0
      Width           =   15135
   End
End
Attribute VB_Name = "frmAF_CD_Control"
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


Dim pEmite As String, pEstado As String, pProceso As String, pTesoreriaId As String

If cboEmite.Text = "TODOS" Then
    pEmite = ""
Else
    pEmite = cboEmite.ItemData(cboEmite.ListIndex)
End If

If cboEstado.Text = "TODOS" Then
    pEstado = ""
Else
    pEstado = cboEstado.ItemData(cboEstado.ListIndex)
End If

If cboProceso.Text = "TODOS" Then
    pProceso = ""
Else
    pProceso = cboProceso.ItemData(cboProceso.ListIndex)
End If

If IsNumeric(txtTesoreriaId.Text) Then
  pTesoreriaId = txtTesoreriaId.Text
Else
  pTesoreriaId = 0
End If

strSQL = "exec spAFI_CD_Cuenta_List '" & txtComite.Tag & "', '" & pEmite & "', '" & pFI & "', '" & pFC & "', '" & pProceso _
       & "', '" & pEstado & "', " & pTesoreriaId
Call OpenRecordSet(rs, strSQL)
    
    
'With lsw.ColumnHeaders
'    .Clear
'    .Add , , "N.Operación", 1800
'    .Add , , "Comité", 3500
'    .Add , , "Monto", 2200, vbRightJustify
'    .Add , , "Estado", 2100, vbCenter
'    .Add , , "Proceso", 2100, vbCenter
'    .Add , , "Fecha", 2100
'    .Add , , "Usuario", 2100, vbCenter
'End With
    
With lsw.ListItems
    .Clear
    Do While Not rs.EOF
        Set itmX = .Add(, , rs!Noperacion)
            itmX.SubItems(1) = rs!COMITE_DESC
            itmX.SubItems(2) = Format(rs!Monto, "Standard")
            itmX.SubItems(3) = rs!Estado_Desc
            itmX.SubItems(4) = rs!Proceso_Desc
            itmX.SubItems(4) = rs!Registro_Fecha
            itmX.SubItems(5) = rs!Registro_Usuario
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




Private Sub btnAccion_Click(Index As Integer)
Select Case Index
    Case 0 'Registro
        Call sbFormsCall("frmAF_CD_Cuentas", , , , False, Me, False)
    Case 1 'PreCalculo
        Call sbFormsCall("frmAF_CD_PreCalculo", , , , False, Me, False)

End Select
End Sub

Private Sub btnAutorizar_Click()
Dim i As Long, pEstado As String, pRefresh As Integer, pNota As String

On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

strSQL = ""
'spAFI_Renuncia_NAT_Tag_Autoriza(@RenunciaId  int = NULL, @Usuario  varchar(30) = NULL)

'With lsw.ListItems
'    For i = 1 To .Count
'        If .Item(i).Checked Then
'            strSQL = strSQL & Space(10) & "exec spAFI_Renuncia_NAT_Tag_Autoriza " & .Item(i).Text & ", '" & glogon.Usuario & "'"
'        End If
'
'        If Len(strSQL) > 20000 Then
'            Call ConectionExecute(strSQL)
'            strSQL = ""
'        End If
'    Next i
'
'End With
'
'
'
''Ultimo Lote
'If Len(strSQL) > 0 Then
'    Call ConectionExecute(strSQL)
'    strSQL = ""
'End If

frmAF_CD_Aprobaciones.Show vbModal

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



Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "select CodTipoCuenta as 'IdX', NombreTipoCuenta as 'itmX' from AFI_CD_TIPO_CUENTA where Activo = 1"
Call sbCbo_Llena_New(cboEmite, strSQL, True, True)

strSQL = "select CodTipoProceso as 'IdX', NombreTipoProceso as 'itmX' from AFI_CD_TIPO_PROCESO where Activo = 1"
Call sbCbo_Llena_New(cboProceso, strSQL, True, True)

strSQL = "select CodEstado as 'IdX', NombreEstado as 'itmX' from AFI_CD_TIPOS_ESTADOS_CUENTAS where Activo = 1"
Call sbCbo_Llena_New(cboEstado, strSQL, True, True)


lsw.Checkboxes = False
With lsw.ColumnHeaders
    .Clear
    .Add , , "N.Operación", 1800
    .Add , , "Comité", 3500
    .Add , , "Monto", 2200, vbRightJustify
    .Add , , "Estado", 2100, vbCenter
    .Add , , "Proceso", 2100, vbCenter
    .Add , , "Fecha", 2100
    .Add , , "Usuario", 2100, vbCenter
End With

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -30, dtpCorte.Value)

Call chkFechas_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub lsw_DblClick()
If lsw.ListItems.Count = 0 Then Exit Sub

Dim pOperacion As Long, frm As Form


pOperacion = lsw.SelectedItem.Text

Call sbFormsCall("frmAF_CD_Cuentas", , , , False, Me, True)
Call sbFormActivo("frmAF_CD_Cuentas", frm)

On Error Resume Next
Call frm.sbConsulta_Externa(pOperacion)


End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Call sbListas_Load

End Sub



Private Sub txtComite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select COD_COMITE, DESCRIPCION from AFI_CD_COMITES"
       gBusquedas.Filtro = " AND ACTIVO = 1"
       frmBusquedas.Show vbModal
       txtComite.Tag = gBusquedas.Resultado
       txtComite.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)

Call sbListas_Load

End Sub






