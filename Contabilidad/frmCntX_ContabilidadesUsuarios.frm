VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCntX_ContabilidadesUsuarios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Usuarios autorizados por Contabilidad"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptX 
      Caption         =   "Contabilidades"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton OptX 
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CheckBox chkTodos 
      Appearance      =   0  'Flat
      Caption         =   "&Todos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer TimerX 
      Interval        =   20
      Left            =   4680
      Top             =   120
   End
   Begin MSComctlLib.ListView lswLista 
      Height          =   7455
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   13150
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   7832
      EndProperty
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   7455
      Left            =   5400
      TabIndex        =   4
      Top             =   600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   13150
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   7832
      EndProperty
   End
   Begin VB.Label lblDato 
      Alignment       =   1  'Right Justify
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmCntX_ContabilidadesUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

On Error GoTo vError

lblDato.Caption = "Seleccione!"
lblDato.Tag = "0x0"

Select Case True
  Case optX.Item(0).Value 'CntX_Contabilidades
       strSQL = "select cod_contabilidad as 'IdX',nombre as 'ItmX' from CntX_Contabilidades"
       Call OpenRecordSet(rs, strSQL, 0)
  Case optX.Item(1).Value 'Usuarios
       strSQL = "exec spCntX_UsuariosAutorizados " & gPortal.Empresa_Id
       Call OpenRecordSet(rs, strSQL, 1)
End Select

lswLista.ListItems.Clear
lsw.ListItems.Clear
vPaso = True

Do While Not rs.EOF
  Set itmX = lswLista.ListItems.Add(, , rs!itmX)
      itmX.Tag = rs!idX
  rs.MoveNext
Loop
rs.Close
vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub chkTodos_Click()
Dim strSQL As String
Dim i As Long

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

On Error GoTo vError

If chkTodos.Value = vbUnchecked Then
   If optX.Item(0).Value Then 'CntX_Contabilidades
       strSQL = "delete CNTX_CONTA_USUARIOS where cod_contabilidad = " & lblDato.Tag
   Else
    'Usuario
       strSQL = "delete CNTX_CONTA_USUARIOS where usuario = '" & lblDato.Tag & "'"
   End If
   Call ConectionExecute(strSQL, 0)
End If

For i = 1 To lsw.ListItems.Count
  If chkTodos.Value = vbChecked And Not lsw.ListItems.Item(i).Checked Then
        If optX.Item(0).Value Then 'CntX_Contabilidades
           strSQL = "insert CNTX_CONTA_USUARIOS(cod_contabilidad,usuario,registro_fecha,registro_usuario)" _
                  & " values(" & lblDato.Tag & ",'" & lsw.ListItems.Item(i).Tag & "',getdate(),'" & glogon.Usuario & "')"
        Else
           strSQL = "delete CNTX_CONTA_USUARIOS where cod_contabilidad = " & lsw.ListItems.Item(i).Tag _
                  & " and usuario = '" & lblDato.Tag & "'"
        End If
        
        Call ConectionExecute(strSQL, 0)
  End If
  
  'Marca / Desmarca
  If lsw.ListItems.Item(i).Checked <> chkTodos.Value Then
     lsw.ListItems.Item(i).Checked = chkTodos.Value
  End If
Next i

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lsw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

Me.MousePointer = vbHourglass

On Error GoTo vError

Select Case True
  Case optX.Item(0).Value 'CntX_Contabilidades
        If Item.Checked Then
           strSQL = "insert CNTX_CONTA_USUARIOS(cod_contabilidad,usuario,registro_fecha,registro_usuario)" _
                  & " values(" & lblDato.Tag & ",'" & Item.Tag & "',getdate(),'" & glogon.Usuario & "')"
        Else
           strSQL = "delete CNTX_CONTA_USUARIOS where cod_contabilidad = " & lblDato.Tag _
                  & " and usuario = '" & Item.Tag & "'"
        End If
      
  Case optX.Item(1).Value 'Usuarios
        If Item.Checked Then
           strSQL = "insert CNTX_CONTA_USUARIOS(cod_contabilidad,usuario,registro_fecha,registro_usuario)" _
                  & " values(" & Item.Tag & ",'" & lblDato.Tag & "',getdate(),'" & glogon.Usuario & "')"
        Else
           strSQL = "delete CNTX_CONTA_USUARIOS where cod_contabilidad = " & Item.Tag _
                  & " and usuario = '" & lblDato.Tag & "'"
        End If

End Select

Call ConectionExecute(strSQL, 0)


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswLista_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If vPaso Then Exit Sub
If lswLista.ListItems.Count = 0 Then Exit Sub

Me.MousePointer = vbHourglass

On Error GoTo vError

lblDato.Caption = lswLista.SelectedItem.Text
lblDato.Tag = lswLista.SelectedItem.Tag

Select Case True
  Case optX.Item(0).Value 'CntX_Contabilidades
       strSQL = "select rtrim(U.Nombre) as 'IdX',rtrim(U.Nombre) as 'ItmX',isnull(A.cod_contabilidad,-1) as 'Marca' " _
              & " from vCntX_UsuariosAutorizados U left join CNTX_CONTA_USUARIOS A on U.nombre = A.usuario" _
              & " and A.cod_contabilidad = " & lblDato.Tag
              
      
  Case optX.Item(1).Value 'Usuarios
       strSQL = "select I.cod_contabilidad as 'IdX',I.nombre as 'ItmX', isnull(A.cod_contabilidad,-1) as 'Marca'" _
              & " from CntX_Contabilidades I left join CNTX_CONTA_USUARIOS A on I.cod_contabilidad = A.cod_contabilidad" _
              & " and A.usuario = '" & lblDato.Tag & "'"

End Select

lsw.ListItems.Clear
vPaso = True

chkTodos.Value = vbUnchecked

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!itmX)
      itmX.Tag = rs!idX
  If rs!Marca <> -1 Then
      itmX.Checked = True
  End If
  rs.MoveNext
Loop
rs.Close
vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub OptX_Click(Index As Integer)
Call sbInicializa
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa


End Sub


