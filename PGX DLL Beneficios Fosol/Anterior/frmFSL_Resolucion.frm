VERSION 5.00
Begin VB.Form frmFSL_Resolucion 
   Caption         =   "Resolución"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14610
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   14610
End
Attribute VB_Name = "frmFSL_Resolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim vFecha As String
'Dim strSQL As String
'Dim rs As New ADODB.Recordset
'
'Private Sub Form_Activate()
' vModulo = 22
'End Sub
'
'Private Sub Form_Load()
' vModulo = 22
' ssTab.Tab = 0
'
' vFecha = fxFechaServidor
'End Sub
'
'Private Sub sbCargaEstado()
'    cboResolucion.Clear
'    cboResolucion.AddItem "Aprobado"
'    cboResolucion.AddItem "Rechazadp"
'    cboResolucion.AddItem "Pendiente"
'    cboResolucion.Text = "Pendiente"
'End Sub

'
'Private Sub sbGuardaResolucion(vExpediente As Integer, vEstado As String)
'
'If vEstado = "A" Then Exit Sub
'
'strSQL = "Update FSL_EXPEDIENTE_COMITE set ESTADO= '" & Mid(cboResolucion.Text, 1, 1) & "',RESOLUCION_ESTADO='" & Mid(cboResolucion.Text, 1, 1) & "', RESOLUCION_NOTAS='" & txtObservaciones.Text & "' ,RESOLUCION_FECHA='" & Format(vFecha, "standard") & "' " _
'       & ", RESOLUCION_USUARIO='" & glogon.Usuario & "' where COD_EXPEDIENTE = " & vExpediente & ""
'
'glogon.Conection.Execute strSQL
'
'End Sub
'
'
'Private Sub sbCargaMiembrosComite()
'Dim vItem As MSComctlLib.ListItem
'Dim vLvw As MSComctlLib.ListView
'Dim vKey As String
'Dim rs As New ADODB.Recordset
'On Error GoTo vError
'
'Me.lswComite.ColumnHeaders.Clear
'Me.lswComite.ListItems.Clear
'
'Set vLvw = Me.lswComite
'vLvw.ColumnHeaders.Add , , "Miembro Comite", 1000
'vLvw.ColumnHeaders.Add , , "Nombre", 5000
'
'strSQL = "Select COD_MIEMBRO, NOMBRE" _
'       & " from FSL_COMITE_MIEMBROS where ACTIVO = 1"
'rs.Open strSQL, glogon.Conection, adOpenStatic
'
'If rs.EOF Then
'   MsgBox "No se tienen miembros de Comite cargados"
'   rs.Close
'   Exit Sub
'End If
'
'Do While Not rs.EOF
'    vKey = Trim(rs.Fields("COD_MIEMBRO")) & "(CA)"
'
'    Set vItem = lswComite.ListItems.Add(, vKey, Trim(rs!Nombre))
'    rs.MoveNext
'Loop
'
'rs.Close
'
'Exit Sub
'vError:
'  MsgBox Err.Description, vbCritical
'
'End Sub
'
'Private Sub SSTab_Click(PreviousTab As Integer)
'   Select Case ssTab.Tab
'      Case 2
'        Call sbCargaMiembrosComite
'   End Select
'End Sub
'
'
'Private Sub sbGuardaComiteAprueba(vExpediente As Integer, vEstado As String)
'On Error GoTo vError
'Dim vCodMiembro As Integer, i As Integer
'Dim vExiste As Boolean
'
'If vEstado = "A" Then Exit Sub
'
'With lswComite
'For i = 1 To .ListItems.Count
'    If .ListItems.Item(i).Checked Then
'        vCodMiembro = DeCodificaPrimaryKey(.ListItems.Item(i).Key, 1, "(CA)")
'
'        strSQL = "Select count(COD_MIEMBRO) as 'Miembros' from FSL_EXPEDIENTE_COMITE" _
'               & " where COD_EXPEDIENTE = " & vExpediente & " and COD_MIEMBRO = " & vCodMiembro & " "
'        rs.Open strSQL, glogon.Conection, adOpenStatic
'
'        If rs!Miembros <= 0 Then vExiste = False
'        rs.Close
'
'        If vExiste = False Then
'            strSQL = "INSERT FSL_EXPEDIENTE_COMITE (COD_EXPEDIENTE,COD_MIEMBRO)" _
'                   & " Values (" & vExpediente & "," & vCodMiembro & ")"
'            glogon.Conection.Execute strSQL
'        End If
'    End If '.ListItems.Item(i).Checked
'Next i
'End With
'
'Exit Sub
'vError:
'     MsgBox Err.Description
'End Sub
'
'
