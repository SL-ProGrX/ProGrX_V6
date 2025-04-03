VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInvOrdenesAutorizacion 
   Caption         =   "Autorización en Lote de Entradas/Salidas/Traslados/Requisiciones"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   9480
   Begin VB.CheckBox chkTodas 
      Caption         =   "Todas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   50
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "Rechazar Solicitudes Marcadas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   9
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton cmdAutorizar 
      Caption         =   "&Autorizar Solicitudes Marcadas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   8
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      Picture         =   "frmInvOrdenesAutorizacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   4695
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#Orden"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Tipo"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Proceso"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Monto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Usuario"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Fecha"
         Object.Width           =   3598
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Causa"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Notas"
         Object.Width           =   7832
      EndProperty
   End
   Begin VB.CheckBox chkTodosPendientes 
      Caption         =   "Todos los Pendientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   680
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Value           =   1  'Checked
      Width           =   1572
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   312
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   1932
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
      Format          =   289013763
      CurrentDate     =   37566
   End
   Begin VB.ComboBox cboTipo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      ItemData        =   "frmInvOrdenesAutorizacion.frx":016F
      Left            =   1440
      List            =   "frmInvOrdenesAutorizacion.frx":017C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   312
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   1932
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
      Format          =   289013763
      CurrentDate     =   37566
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   5160
      X2              =   5160
      Y1              =   0
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   5160
      X2              =   5160
      Y1              =   0
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Solicitadas Entre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   48
      TabIndex        =   2
      Top             =   480
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
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
      Index           =   0
      Left            =   50
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmInvOrdenesAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Dim strSQL As String, lng As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

For lng = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(lng).Checked Then
    If lsw.ListItems(lng).SubItems(1) = "R" Then
        strSQL = "update pv_requisiciones set autoriza_fecha = dbo.MyGetdate(),autoriza_user = '" _
               & glogon.Usuario & "',estado = 'A' where cod_requisicion = " & lsw.ListItems.Item(lng).Text
    Else
        strSQL = "update pv_InvTranSac set autoriza_fecha = dbo.MyGetdate(),autoriza_user = '" _
               & glogon.Usuario & "',estado = 'A' where cod_orden = " & lsw.ListItems.Item(lng).Text _
               & " and tipo_orden = '" & lsw.ListItems.Item(lng).SubItems(1) & "'"
    End If
    Call ConectionExecute(strSQL)
 End If
Next lng

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
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

If Mid(cboTipo.Text, 1, 2) <> "03" Then
    strSQL = "select O.cod_orden,O.tipo_orden,O.total,O.user_solicita,O.fecha" _
           & ",C.descripcion as Causa,O.nota,O.proceso" _
           & " from pv_InvTranSac O inner join pv_entrada_salida C on O.cod_entsal = C.cod_entsal" _
           & " where O.autoriza_fecha is null and O.estado = 'S' and O.tipo_orden = '" _
           & IIf((Mid(cboTipo, 1, 2) = "01"), "E", "S") & "' and O.user_solicita in(" _
           & "select usuario_asignado from pv_orden_autousers where usuario = '" _
           & glogon.Usuario & "')"
    
    If chkTodosPendientes.Value = vbUnchecked Then
      strSQL = strSQL & " and O.fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd hh:mm:ss") _
             & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd hh:mm:ss") & "'"
    End If
Else
 'Requisiciones
  strSQL = "select R.cod_requisicion as Cod_Orden,'R' as Tipo_Orden,0 as Total,R.Genera_User as User_Solicita" _
          & ",R.Genera_Fecha as Fecha,C.descripcion as Causa,R.notas as Nota,'P' as proceso" _
          & " from pv_requisiciones R inner join pv_entrada_salida C on R.cod_entsal = C.cod_entsal" _
          & " where R.autoriza_fecha is null and R.Genera_User in(" _
          & "select usuario_asignado from pv_orden_autousers where usuario = '" _
          & glogon.Usuario & "')"

    If chkTodosPendientes.Value = vbUnchecked Then
      strSQL = strSQL & " and R.Genera_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd hh:mm:ss") _
             & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd hh:mm:ss") & "'"
    End If
End If
       
       
'strSQL = "select Boleta,Tipo"
If chkTodosPendientes.Value = vbUnchecked Then
  strSQL = strSQL & " and R.Genera_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd hh:mm:ss") _
         & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd hh:mm:ss") & "'"
End If
       
lsw.ListItems.Clear

Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!cod_orden)
     itmX.SubItems(1) = rs!tipo_Orden
   Select Case rs!Proceso
     Case "P"
        itmX.SubItems(2) = "Pendiente"
     Case "C"
        itmX.SubItems(2) = "Cotizada"
   End Select
     
     itmX.SubItems(3) = Format(rs!Total, "Standard")
     itmX.SubItems(4) = rs!user_solicita
     itmX.SubItems(5) = Format(rs!fecha, "yyyy/mm/dd hh:mm:ss")
     itmX.SubItems(6) = rs!Causa
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
Dim strSQL As String, lng As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

For lng = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(lng).Checked Then
    If lsw.ListItems(lng).SubItems(1) = "R" Then
        strSQL = "update pv_requisiciones set autoriza_fecha = dbo.MyGetdate(),autoriza_user = '" _
               & glogon.Usuario & "',estado = 'R' where cod_requisicion = " & lsw.ListItems.Item(lng).Text
    Else
        strSQL = "update pv_InvTranSac set autoriza_fecha = dbo.MyGetdate(),autoriza_user = '" _
               & glogon.Usuario & "',estado = 'R' where cod_orden = " & lsw.ListItems.Item(lng).Text _
               & " and tipo_orden = '" & lsw.ListItems.Item(lng).SubItems(1) & "'"
    End If
    Call ConectionExecute(strSQL)
 End If
Next lng

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
vModulo = 32
End Sub

Private Sub Form_Load()

vModulo = 32

cboTipo.Clear
cboTipo.AddItem "Entradas"
cboTipo.AddItem "Salidas"
cboTipo.AddItem "Traspasos"
cboTipo.AddItem "Requisiciones"
cboTipo.Text = "Entradas"

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
