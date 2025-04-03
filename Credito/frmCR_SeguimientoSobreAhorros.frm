VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCR_SeguimientoSobreAhorros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indique cuales códigos va a Refundir"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVerificar 
      Caption         =   "&Verifica"
      Height          =   315
      Left            =   5040
      TabIndex        =   1
      Top             =   3000
      Width           =   855
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Operación"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   1536
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descripción"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Saldo"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmCR_SeguimientoSobreAhorros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curMontoDisponible As Currency

Private Function fxMoraPendiente(lngOp As Long) As Currency
Dim strSQL As String, rsX As New ADODB.Recordset

fxMoraPendiente = 0

strSQL = "select * from vista_morosidad where id_solicitud = " & lngOp
rsX.CursorLocation = adUseServer
rsX.Open strSQL, glogon.Conection, adOpenStatic
If Not rsX.EOF And Not rsX.BOF Then
  fxMoraPendiente = rsX!intc + rsX!intm
End If
rsX.Close
End Function

Private Sub sbLlenaLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

strSQL = "select id_solicitud,codigo,saldo from reg_creditos where " _
       & " garantia = 'A' and estado = 'A' and proceso = 'N' and cedula = '" _
       & Operacion.Cedula & "'"
       
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic
lsw.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(lsw.ListItems.Count + 1, , rs!id_solicitud)
     itmX.SubItems(1) = rs!Codigo
     itmX.SubItems(2) = fxDescribeCodigo(rs!Codigo)
     itmX.SubItems(3) = Format(rs!Saldo + fxMoraPendiente(rs!id_solicitud), "###,###,###,##0.00")
     If rs!Codigo = Operacion.Codigo Then itmX.Checked = True
  rs.MoveNext
Loop
rs.Close

'Se supone que el monto aprobado ya se valido contra el porcentaje
curMontoDisponible = Operacion.MontoAprobado

End Sub

Private Sub cmdVerificar_Click()
Dim itmX As ListItem, curMonto As Currency, lng As Long
Dim strSQL As String, rs As New ADODB.Recordset, vFecha As Date
Dim strSQL2 As String
Me.MousePointer = vbHourglass

curMonto = 0
vFecha = fxFechaServidor

strSQL = "delete refundiciones where id_solicitudr = " & Operacion.Operacion
glogon.Conection.Execute strSQL

For lng = 1 To lsw.ListItems.Count
  lsw.SelectedItem = lsw.ListItems(lng)
  If lsw.SelectedItem.Checked = True Then
     curMonto = curMonto + CCur(lsw.SelectedItem.SubItems(3))
     strSQL = "select saldo from reg_creditos where id_solicitud = " & lsw.SelectedItem.Text
     rs.Open strSQL, glogon.Conection, adOpenStatic
     strSQL = "insert refundiciones(ID_SOLICITUD,CODIGO,CODIGOR,MONTO,FECHA,ID_SOLICITUDR" _
            & ",INTCOR,INTMOR) values(" & lsw.SelectedItem.Text & ",'" _
            & lsw.SelectedItem.SubItems(1) & "','" & Operacion.Codigo & "'," & rs!Saldo
     rs.Close
     
     strSQL2 = "select coalesce(sum(intc),0) as intc, coalesce(sum(intm),0) as intm " _
            & "from morosidad where estado = 'A' and id_solicitud = " & lsw.SelectedItem.Text
     rs.Open strSQL2, glogon.Conection, adOpenStatic
     strSQL = strSQL & ",'" & Format(vFecha, "yyyy/mm/dd") & "'," & Operacion.Operacion _
            & "," & rs!intc & "," & rs!intm & ")"
     rs.Close
     glogon.Conection.Execute strSQL
  End If
Next lng

Me.MousePointer = vbDefault

If curMonto < curMontoDisponible Then
  Operacion.Valida = True
  Unload Me
Else
  Operacion.Valida = False
  MsgBox "El monto de operaciones a refundir no se cumple por ser mayor al autorizado...", vbCritical
End If

End Sub

Private Sub Form_Load()
    Operacion.Valida = False
    Call sbLlenaLsw
End Sub
