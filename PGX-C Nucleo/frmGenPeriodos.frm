VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmGenPeriodos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Periodos de Cierre"
   ClientHeight    =   5160
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   6792
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   6792
   Begin VB.OptionButton opt 
      Caption         =   "Cerrados"
      Height          =   375
      Index           =   1
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      Width           =   1335
   End
   Begin VB.OptionButton opt 
      Caption         =   "Pendientes"
      Height          =   375
      Index           =   0
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtMes 
      Height          =   315
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtAnio 
      Height          =   315
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   3855
      _ExtentX        =   6795
      _ExtentY        =   6371
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Año"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Mes"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Estado"
         Object.Width           =   3951
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   6480
      X2              =   4200
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   4200
      X2              =   6480
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Mes"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Año"
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   $"frmGenPeriodos.frx":0000
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmGenPeriodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbLlenaLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

strSQL = "select * from pv_periodos where estado = '"

Select Case True
  Case opt.Item(0)
     strSQL = strSQL & "P'"
  Case opt.Item(1)
     strSQL = strSQL & "C'"
End Select
strSQL = strSQL & " order by proceso desc"

Call OpenRecordSet(rs, strSQL, 0)
Lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = Lsw.ListItems.Add(, , rs!Anio)
      itmX.SubItems(1) = rs!Mes
      itmX.SubItems(2) = IIf((rs!Estado = "P"), "Pendiente", "Cerrado")
  rs.MoveNext
Loop
rs.Close

End Sub

Private Function fxExisteProducto(vBodega As String, vProducto As String, vAnio As Long, vMes As Integer) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from pv_inventario where cod_bodega = '" & vBodega _
       & "' and cod_producto = '" & vProducto & "' and anio = " & vAnio & " and mes = " & vMes
Call OpenRecordSet(rs, strSQL)
fxExisteProducto = IIf((rs!Existe = 0), False, True)
rs.Close
End Function

Private Sub cmdCerrar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngAnioX As Long, iMesX As Integer
Dim i As Byte

On Error GoTo vError


i = MsgBox("Esta seguro que desea cerrar el periodo seleccionado", vbYesNo)
If i = vbNo Then Exit Sub

'Verifica que no se haya cerrado el periodo anteriormente
strSQL = "select isnull(count(*),0) as Existe from pv_periodos where estado = 'P'" _
       & " and anio = " & txtAnio & " and Mes = " & txtMes
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
   MsgBox "El periodo ya se encuentra cerrado, verifique...", vbExclamation
   Exit Sub
End If
rs.Close

'Verifica que no se hayan cerrado periodos posteriores
strSQL = "select isnull(count(*),0) as Existe from pv_periodos" _
       & " where mes >" & txtMes & " and anio = " & txtAnio & " and estado = 'C'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
   MsgBox "Existen periodos posteriores ya cerrados, verifique...", vbExclamation
   Exit Sub
End If
rs.Close

strSQL = "select isnull(count(*),0) as Existe from pv_periodos" _
       & " where anio > " & txtAnio & " and estado = 'C'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
   MsgBox "Existen periodos posteriores ya cerrados, verifique...", vbExclamation
   Exit Sub
End If
rs.Close


'Verifica que el periodo anterior este cerrado, por cuestiones de orden
'Si no existe no hay problema.

lngAnioX = txtAnio
iMesX = txtMes

If iMesX = 1 Then
   iMesX = 12
   lngAnioX = lngAnioX - 1
Else
   iMesX = iMesX - 1
End If

strSQL = "select Estado from pv_periodos" _
       & " where anio = " & lngAnioX & " and Mes = " & iMesX
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  If rs!Estado = "P" Then
    MsgBox "El periodo Anterior no se ha cerrado, proceda en el mismo orden...", vbExclamation
    Exit Sub
  End If
End If
rs.Close

'Inicia Proceso

Me.MousePointer = vbHourglass



'Verificar si existe el Periodo Siguiente, de lo contrario crearlo


lngAnioX = txtAnio
iMesX = txtMes

If iMesX = 12 Then
   iMesX = 1
   lngAnioX = lngAnioX + 1
Else
   iMesX = iMesX + 1
End If
 
' /* desde aqui lo realiza el Sp */
strSQL = "exec spINVCierrePeriodo " & txtAnio & "," & txtMes & "," & lngAnioX & "," & iMesX
Call ConectionExecute(strSQL)

Call opt_Click(0)
Me.MousePointer = vbDefault
MsgBox "El Cierre del Periodo se Realizó Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 34
End Sub

Private Sub Form_Load()
vModulo = 34
Call Formularios(Me)

Call sbLlenaLsw

Call RefrescaTags(Me)
End Sub

Private Sub lsw_Click()
If Lsw.ListItems.Count > 0 Then
   txtAnio = Lsw.SelectedItem
   txtMes = Lsw.SelectedItem.SubItems(1)
End If
End Sub

Private Sub opt_Click(Index As Integer)
Call sbLlenaLsw
End Sub
