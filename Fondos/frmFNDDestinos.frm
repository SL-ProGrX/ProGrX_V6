VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.ShortcutBar.v19.1.0.ocx"
Begin VB.Form frmFNDDestinos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Destinos"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   13428
   Icon            =   "frmFNDDestinos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6672
   ScaleWidth      =   13428
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4812
      Left            =   7320
      TabIndex        =   2
      Top             =   1800
      Width           =   6012
      _Version        =   1245185
      _ExtentX        =   10604
      _ExtentY        =   8488
      _StockProps     =   77
      BackColor       =   -2147483643
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
      Appearance      =   16
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4692
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   7092
      _Version        =   524288
      _ExtentX        =   12509
      _ExtentY        =   8276
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   485
      ScrollBars      =   2
      SpreadDesigner  =   "frmFNDDestinos.frx":6852
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scDestino 
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   13212
      _Version        =   1245185
      _ExtentX        =   23304
      _ExtentY        =   656
      _StockProps     =   14
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
      VisualTheme     =   3
      Alignment       =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Destinos de Planes de Ahorros e Inversión"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   7332
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmFNDDestinos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbPlanes_Consulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

lsw.ListItems.Clear

If vPaso Then Exit Sub
If scDestino.Tag = "" Then Exit Sub

strSQL = "select D.cod_operadora,D.COD_PLAN,D.descripcion,A.cod_destino" _
       & " from fnd_Planes D left join fnd_planes_destinos A on D.COD_OPERADORA = A.cod_operadora" _
       & " and D.COD_PLAN = A.cod_plan and A.cod_destino = '" & scDestino.Tag & "'" _
       & " where D.Estado = 'A'"
Call OpenRecordSet(rs, strSQL)

vPaso = True

With lsw.ListItems
  Do While Not rs.EOF
     Set itmX = .Add(, , rs!cod_Plan)
         itmX.SubItems(1) = rs!Descripcion
         itmX.SubItems(2) = rs!cod_Operadora
         
     If IsNull(rs!cod_destino) Then
        itmX.Checked = False
     Else
        itmX.Checked = True
        itmX.ForeColor = vbBlue
     End If
     
     rs.MoveNext
  Loop
End With
rs.Close

vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 18
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 18
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Plan", 1500
    .Add , , "Descripción", 4000
    .Add , , "Operadora Id", 300
End With
lsw.Checkboxes = True

vPaso = True
    strSQL = "select cod_destino,descripcion,activo,0 from fnd_destinos order by cod_destino"
    Call sbCargaGrid(vGrid, 4, strSQL)
vPaso = False

scDestino.Tag = ""
scDestino.Caption = "(Seleccione un destino para asignar Planes)"



Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If Trim(vGrid.Text) = "" Then
   MsgBox "No se especifico ningún código....verifique..!!!", vbExclamation
   Exit Function
End If

strSQL = "select isnull(count(*),0) as Existe from FND_DESTINOS where cod_destino = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  vGrid.Col = 1
  strSQL = "insert into FND_DESTINOS(cod_destino,descripcion,Activo,registro_usuario,registro_fecha) values('"
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
  

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Destinos de Planes: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update FND_DESTINOS set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & ",actualiza_usuario = '" & glogon.Usuario _
        & "', actualiza_fecha = dbo.MyGetdate() where cod_destino = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 
 Call Bitacora("Modifica", "Destinos de Planes: " & vGrid.Text)

End If
rs.Close
fxGuardar = 1


Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError



If Item.Checked Then
   strSQL = "insert fnd_planes_destinos(cod_plan,cod_operadora,cod_destino,registro_usuario,registro_fecha)" _
          & " values('" & Item.Text & "'," & Item.SubItems(2) & ",'" & scDestino.Tag _
          & "','" & glogon.Usuario & "',dbo.MyGetdate())"
   Call ConectionExecute(strSQL)
   Call Bitacora("Aplica", "Asignación Plan " & Item.Text & " -> Destino : " & scDestino.Tag)

Else
   strSQL = "delete fnd_planes_destinos where cod_destino = '" & scDestino.Tag _
          & "' and cod_operadora = " & Item.SubItems(2) & " and cod_plan = '" & Item.Text & "'"
   Call ConectionExecute(strSQL)
   Call Bitacora("Elimina", "Asignación Plan " & Item.Text & " -> Destino : " & scDestino.Tag)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub
If Col <> 4 Then Exit Sub

vGrid.Row = Row
vGrid.Col = 1
scDestino.Tag = vGrid.Text
vGrid.Col = 2
scDestino.Caption = vGrid.Text

Call sbPlanes_Consulta

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol >= (vGrid.MaxCols - 1) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow

  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then

        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1

       If vGrid.Text = "" Then Exit Sub

     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Col = 1
        strSQL = "delete fnd_destinos where cod_destino = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        
        Call Bitacora("Elimina", "Destino de Planes : " & vGrid.Text)
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
        
     End If
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

