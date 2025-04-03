VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCprOrdenesAutorizacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Autorización de Ordenes de Compra"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   12030
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6015
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   12015
      _Version        =   1441793
      _ExtentX        =   21193
      _ExtentY        =   10610
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
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkTodas 
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   840
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Sel. Todas"
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
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "&Buscar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCprOrdenesAutorizacion.frx":0000
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton cmdRechazar 
      Height          =   495
      Left            =   9720
      TabIndex        =   3
      Top             =   120
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4043
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Rechazar Ordenes"
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
      Picture         =   "frmCprOrdenesAutorizacion.frx":0700
   End
   Begin XtremeSuiteControls.PushButton cmdAutorizar 
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   120
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4043
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Autoriza Ordenes"
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
      Picture         =   "frmCprOrdenesAutorizacion.frx":0CA4
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
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
      CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   3600
      TabIndex        =   6
      Top             =   480
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3408
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
      CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
      Format          =   3
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Width           =   3855
      _Version        =   1441793
      _ExtentX        =   6800
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
   Begin XtremeSuiteControls.CheckBox chkTodosPendientes 
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   840
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Todas las Fechas"
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
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   600
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Exportar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCprOrdenesAutorizacion.frx":13CB
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   7440
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   4575
      _Version        =   1441793
      _ExtentX        =   8070
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      Index           =   0
      Left            =   285
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitadas Entre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   288
      TabIndex        =   0
      Top             =   480
      Width           =   1212
   End
End
Attribute VB_Name = "frmCprOrdenesAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem



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

Private Sub chkTodas_Click()
Dim lng As Long

For lng = 1 To lsw.ListItems.Count
 lsw.ListItems.Item(lng).Checked = chkTodas.Value
Next lng

End Sub

Private Sub chkTodosPendientes_Click()
If chkTodosPendientes.Value = vbChecked Then
  dtpInicio.Enabled = False
  dtpCorte.Enabled = False
Else
  dtpInicio.Enabled = True
  dtpCorte.Enabled = True
End If
End Sub

Private Sub cmdAutorizar_Click()
Dim lng As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = ""

For lng = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(lng).Checked Then
    strSQL = strSQL & Space(10) & "update cpr_ordenes set autoriza_fecha = dbo.MyGetdate(),autoriza_user = '" _
           & glogon.Usuario & "',estado = 'A' where cod_orden = '" & lsw.ListItems.Item(lng).Text & "'"
    
        
    If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If
 
 End If
Next lng

'Lote Final
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If


Me.MousePointer = vbDefault
MsgBox "Solicitudes Autorizadas Satisfactoriamente...", vbInformation

Call cmdBuscar_Click

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Call cmdBuscar_Click

End Sub

Private Sub cmdBuscar_Click()


On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select O.cod_orden,C.Descripcion as 'TipoOrdenDesc',O.total,O.genera_user,O.genera_fecha" _
       & ",C.descripcion as TipoOrden,O.nota,O.proceso" _
       & " from cpr_ordenes O inner join cpr_Tipo_Orden C on O.Tipo_Orden = C.Tipo_Orden" _
       & " where O.autoriza_fecha is null and O.estado = 'S' and O.tipo_orden = '" _
       & cboTipo.ItemData(cboTipo.ListIndex) & "' and O.genera_user in(" _
       & "select usuario_asignado from cpr_orden_autousers where usuario = '" _
       & glogon.Usuario & "')"

If chkTodosPendientes.Value = vbUnchecked Then
  strSQL = strSQL & " and O.genera_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd hh:mm:ss") _
         & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd hh:mm:ss") & "'"
End If
       
       
lsw.ListItems.Clear
lsw.Checkboxes = True
Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!cod_orden)
     itmX.SubItems(1) = rs!TipoOrdenDesc
   Select Case rs!Proceso
     Case "P"
        itmX.SubItems(2) = "Pendiente"
     Case "C"
        itmX.SubItems(2) = "Cotizada"
   End Select
     
     itmX.SubItems(3) = Format(rs!Total, "Standard")
     itmX.SubItems(4) = rs!genera_user
     itmX.SubItems(5) = Format(rs!genera_fecha, "yyyy/mm/dd hh:mm:ss")
     itmX.SubItems(6) = rs!tipoOrden
     itmX.SubItems(7) = rs!nota
     
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cmdRechazar_Click()
Dim lng As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = ""

For lng = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(lng).Checked Then
    strSQL = strSQL & Space(10) & "update cpr_ordenes set autoriza_fecha = dbo.MyGetdate(),autoriza_user = '" _
           & glogon.Usuario & "',estado = 'R' where cod_orden = " & lsw.ListItems.Item(lng).Text & "'"

    If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If
 


 End If
Next lng


'Lote Final
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If

Me.MousePointer = vbDefault

MsgBox "Solicitudes Rechazadas Satisfactoriamente...", vbInformation

Call cmdBuscar_Click

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Call cmdBuscar_Click
End Sub

Private Sub Form_Activate()
vModulo = 35
End Sub

Private Sub Form_Load()

vModulo = 35

Call sbCprCboTiposOrden(cboTipo)

With lsw.ColumnHeaders
    .Clear
    .Add , , "No Orden", 1600
    .Add , , "Tipo", 2500
    .Add , , "Proceso", 2500, vbCenter
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Usuario", 2500, vbCenter
    .Add , , "Fecha", 2500, vbCenter
    .Add , , "Causa", 2500
    .Add , , "Notas", 4500
    
End With

dtpInicio.Value = Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss")
dtpCorte.Value = dtpInicio.Value

Call chkTodosPendientes_Click

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Sub Form_Resize()
On Error Resume Next

lsw.Width = Me.Width - 150
lsw.Height = Me.Height - 1700

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub
