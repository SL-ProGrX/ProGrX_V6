VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAH_Autorizaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Patrimonio: Autorizaciones de movimientos"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   13995
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   10335
      _Version        =   1441793
      _ExtentX        =   18224
      _ExtentY        =   8064
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
      FlatScrollBar   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   0
      Left            =   8160
      TabIndex        =   15
      Top             =   1560
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
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
      Picture         =   "frmAH_Autorizaciones.frx":0000
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   10080
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   5040
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   6360
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
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
      Left            =   7680
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
      _Version        =   1441793
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
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   1440
      TabIndex        =   8
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1200
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   3360
      TabIndex        =   9
      Top             =   1200
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   330
      Left            =   1440
      TabIndex        =   10
      Top             =   1560
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   1
      Left            =   8640
      TabIndex        =   16
      Top             =   1560
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmAH_Autorizaciones.frx":0700
   End
   Begin XtremeSuiteControls.PushButton btnResolucion 
      Height          =   375
      Index           =   0
      Left            =   10080
      TabIndex        =   17
      Top             =   1560
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Picture         =   "frmAH_Autorizaciones.frx":0FD1
   End
   Begin XtremeSuiteControls.PushButton btnResolucion 
      Height          =   375
      Index           =   1
      Left            =   11280
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Denegar"
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
      Picture         =   "frmAH_Autorizaciones.frx":16F8
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   14
      Top             =   1560
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   1200
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Persona"
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
   Begin XtremeShortcutBar.ShortcutCaption lblX 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   10335
      _Version        =   1441793
      _ExtentX        =   18224
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "                                Seleccione las Solicitudes  a Autorizar o Desautorizar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Autorización de Solicitudes"
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
      Height          =   492
      Left            =   2004
      TabIndex        =   6
      Top             =   360
      Width           =   6852
   End
   Begin VB.Image imgBanner 
      Height          =   1125
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAH_Autorizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim vPaso As Boolean


Private Sub btnAccion_Click(Index As Integer)
Select Case Index
  Case 0 'Buscar
    Call sbBuscar
  
  Case 1 'Exportar
    Call Excel_Exportar_Lsw(lsw, ProgressBarX)

End Select

End Sub

Private Function fxAutorizador_Valida(pAutorizador As String) As Boolean
Dim pResult As Boolean

pResult = False

strSQL = "select dbo.fxPAT_Autorizado_Valida('" & pAutorizador & "') as 'Estado'"
Call OpenRecordSet(rs, strSQL)
If rs!Estado = 1 Then
   pResult = True
End If

fxAutorizador_Valida = pResult

End Function

Private Sub btnResolucion_Click(Index As Integer)

If Not fxAutorizador_Valida(glogon.Usuario) Then
    MsgBox "El usuario actual no es un Autorizador de Gestiones de Patrimonio!", vbExclamation
    Exit Sub
End If

Select Case Index
  Case 0 'Autorizar
    Call sbAutoriza("A")
  Case 1 'Denegar
    Call sbAutoriza("D")
End Select
End Sub

Private Sub cboEstado_Click()
If vPaso Then Exit Sub

Call sbBuscar

End Sub


Private Sub chkTodos_Click()
Dim i As Long

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub Form_Activate()
vModulo = 2

End Sub

Private Sub Form_Load()

vModulo = 2

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True

cboEstado.Clear
cboEstado.AddItem "Pendientes"
cboEstado.AddItem "Autorizadas"
cboEstado.AddItem "Denegadas"
cboEstado.AddItem "Vencidas"

cboEstado.Text = "Pendientes"

vPaso = False

Call Formularios(Me)

btnResolucion.Item(1).Tag = btnResolucion.Item(0).Tag
dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value


Call RefrescaTags(Me)

End Sub


Private Sub sbBuscar()

Me.MousePointer = vbHourglass

On Error GoTo vError

lsw.ListItems.Clear


txtUsuario.Text = fxSysCleanTxtInject(txtUsuario.Text)
txtCedula.Text = fxSysCleanTxtInject(txtCedula.Text)
txtNombre.Text = fxSysCleanTxtInject(txtNombre.Text)

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id Gestión", 2100
    .Add , , "Estado", 1800, vbCenter
    .Add , , "Identificación", 1500
    .Add , , "Nombre", 3200
    .Add , , "Tipo", 2200
    .Add , , "Mnt.Calculado", 2100, vbRightJustify
    .Add , , "Mnt.Solicitado", 2100, vbRightJustify
    .Add , , "Diferencia", 2100, vbRightJustify
    .Add , , "Usuario", 1800, vbCenter
    .Add , , "Fecha", 2100

    .Add , , "Res.Usuario", 1800, vbCenter
    .Add , , "Res.Fecha", 2100

    .Add , , "Apl.Usuario", 1800, vbCenter
    .Add , , "Apl.Fecha", 2100

    .Add , , "Apl.TDoc", 1800, vbCenter
    .Add , , "Apl.NDoc", 2100, vbCenter


End With

strSQL = "select * from vPAT_Gestiones_List"

strSQL = strSQL & " Where Estado = '" & Mid(cboEstado.Text, 1, 1) _
       & "' and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"


'Filtros
If txtUsuario.Text = "" Then
    strSQL = strSQL & " and Registro_Usuario like '%" & txtUsuario.Text & "%'"
End If

If txtCedula.Text = "" Then
    strSQL = strSQL & " and Cedula like '%" & txtCedula.Text & "%'"
End If

If txtNombre.Text = "" Then
    strSQL = strSQL & " and Nombre like '%" & txtNombre.Text & "%'"
End If



Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Id_Autorizacion)
     itmX.SubItems(1) = rs!Estado_Desc
     itmX.SubItems(2) = rs!Cedula
     itmX.SubItems(3) = rs!Nombre
     itmX.SubItems(4) = rs!Tipo_Desc
     itmX.SubItems(5) = Format(rs!Monto_Calculado, "Standard")
     itmX.SubItems(6) = Format(rs!Monto_Solicitado, "Standard")
     itmX.SubItems(7) = Format(rs!Monto_Dif, "Standard")
     
     
     itmX.SubItems(8) = rs!Registro_Usuario & ""
     itmX.SubItems(9) = Format(rs!Registro_Fecha & "", "dd/mm/yyyy")
     
     itmX.SubItems(10) = rs!Resuelve_Usuario & ""
     itmX.SubItems(11) = Format(rs!Resuelve_Fecha & "", "dd/mm/yyyy")
     
     itmX.SubItems(12) = rs!Aplica_Usuario & ""
     itmX.SubItems(13) = Format(rs!Aplica_Fecha & "", "dd/mm/yyyy")
     
     itmX.SubItems(14) = rs!Tcon & ""
     itmX.SubItems(15) = rs!nCon & ""
     
     
     Select Case rs!Estado
         Case "P"
         Case "V"
         Case "A"
              itmX.Bold = True
              itmX.TextBackColor = RGB(252, 243, 207)
         Case "D"
              itmX.ForeColor = vbRed
              itmX.Bold = True
              itmX.TextBackColor = RGB(250, 219, 216)
     End Select

 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbAutoriza(pGestion As String)
Dim i As Long

On Error GoTo vError

Me.MousePointer = vbHourglass


pGestion = Mid(pGestion, 1, 1)


With lsw.ListItems
  For i = 1 To .Count
      If .Item(i).Checked Then
      
         strSQL = "exec spPAT_Autorizaciones_Registro '" & .Item(i).Text & "','" & pGestion & "','" & glogon.Usuario & "'"
         Call ConectionExecute(strSQL)

         Call Bitacora("Aplica", IIf((pGestion = "A"), "Autoriza", "Deniega") & " de Gestion Id:" & .Item(i).Text _
                 & "..Id: " & .Item(i).SubItems(2) & "..Nombre: " & .Item(i).SubItems(3))

      End If
  Next i
End With

Me.MousePointer = vbDefault
MsgBox IIf((pGestion = "A"), "Autorización", "Denegación") & " realizada satisfactoriamente.!", vbInformation

Call sbBuscar

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub Form_Resize()
On Error Resume Next


imgBanner.Width = Me.Width

lsw.Width = Me.Width - 450
lblX.Width = lsw.Width
lsw.Height = Me.Height - (lsw.top + 880)


End Sub



Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If Item.SubItems(14) <> "" Then
 Call sbImprimeRecibo(Item.SubItems(15), Item.SubItems(14))
End If

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Identificación"
   gBusquedas.Col2Name = "Id Alterno"
   gBusquedas.Col3Name = "Nombre"
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Consulta = "Select Cedula,CedulaR, Nombre From Socios"
   gBusquedas.Filtro = ""

   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
        txtCedula.Text = Trim(gBusquedas.Resultado)
        txtNombre.Text = Trim(gBusquedas.Resultado2)
   End If
End If

End Sub
