VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmFnd_SeguridadTasaPreferencial 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fondos: Niveles de Seguridad para Autorización de Tasas Preferenciales"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ListView lswUsuarios 
      Height          =   3495
      Left            =   5400
      TabIndex        =   0
      Top             =   5040
      Width           =   5295
      _Version        =   1572864
      _ExtentX        =   9334
      _ExtentY        =   6159
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
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   5295
      _Version        =   1572864
      _ExtentX        =   9334
      _ExtentY        =   6159
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
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   10455
      _Version        =   524288
      _ExtentX        =   18441
      _ExtentY        =   5318
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
      MaxCols         =   492
      ScrollBars      =   2
      SpreadDesigner  =   "frmFnd_SeguridadTasaPreferencial.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   5295
      _Version        =   1572864
      _ExtentX        =   9340
      _ExtentY        =   635
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
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   360
      Index           =   1
      Left            =   5400
      TabIndex        =   4
      Top             =   4680
      Width           =   5295
      _Version        =   1572864
      _ExtentX        =   9340
      _ExtentY        =   635
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
   Begin XtremeShortcutBar.ShortcutCaption scNivel 
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   10572
      _Version        =   1572864
      _ExtentX        =   18648
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
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Niveles de Seguridad: Autorización de Tasas Preferenciales"
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
      Height          =   480
      Index           =   0
      Left            =   2040
      TabIndex        =   5
      Top             =   360
      Width           =   7575
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmFnd_SeguridadTasaPreferencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean


Private Sub sbPlanes_Load()

On Error GoTo vError

If vPaso Or scNivel.Tag = "" Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

vPaso = True

lsw.ListItems.Clear

txtFiltro(0).Text = fxSysCleanTxtInject(txtFiltro(0).Text)

'Planes
strSQL = "select Pl.cod_Operadora, Pl.cod_Plan,Pl.Descripcion,Asg.registro_Fecha,ASg.Registro_Usuario" _
       & " from Fnd_Planes Pl left join FND_TP_ROLES_PLANES Asg on Pl.cod_operadora = Asg.cod_Operadora" _
       & " and Pl.cod_Plan = Asg.Cod_Plan and Asg.TP_ROL = '" & scNivel.Tag _
       & "' where Estado = 'A' and (Pl.Cod_Plan like '%" & txtFiltro(0).Text & "%' or Pl.Descripcion like '%" & txtFiltro(0).Text & "%')" _
       & " order by isnull(Asg.Cod_Plan,'ZZZZZZZZZZZZ') asc,Pl.cod_Plan asc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Cod_Plan)
     itmX.Tag = rs!COD_OPERADORA
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!Registro_Usuario & ""
     itmX.SubItems(3) = rs!Registro_Fecha & ""
     
     If Not IsNull(rs!Registro_Fecha) Then itmX.Checked = True
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

Private Sub sbUsuarios_Load()

On Error GoTo vError

If vPaso Or scNivel.Tag = "" Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

vPaso = True

lswUsuarios.ListItems.Clear

txtFiltro(1).Text = fxSysCleanTxtInject(txtFiltro(1).Text)

'Usuarios
strSQL = "select Us.Nombre,Us.Descripcion,Asg.registro_Fecha,ASg.Registro_Usuario" _
       & " from Usuarios Us left join FND_TP_ROLES_AUTORIZADORES Asg on Us.Nombre = Asg.Usuario" _
       & " and Asg.TP_ROL = '" & scNivel.Tag _
       & "' Where Us.Estado = 'A' and (Us.Nombre like '%" & txtFiltro(1).Text & "%' or Us.Descripcion like '%" & txtFiltro(1).Text & "%')" _
       & " Order by isnull(Asg.Usuario,'ZZZZZZZZZZZ') asc, Us.Nombre asc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswUsuarios.ListItems.Add(, , rs!Nombre)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!Registro_Usuario & ""
     itmX.SubItems(3) = rs!Registro_Fecha & ""
     
     If Not IsNull(rs!Registro_Fecha) Then itmX.Checked = True
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



Private Sub sbDetalle_Consulta()
 Call sbPlanes_Load
 Call sbUsuarios_Load
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
   .Add , , "Planes", 1500
   .Add , , "Descripción", 3500
   .Add , , "Usuario", 10
   .Add , , "Fecha", 10
End With
lsw.Checkboxes = True

With lswUsuarios.ColumnHeaders
   .Clear
   .Add , , "Usuarios", 1500
   .Add , , "Descripción", 3500
   .Add , , "Usuario", 10
   .Add , , "Fecha", 10
End With
lswUsuarios.Checkboxes = True


scNivel.Tag = ""
scNivel.Caption = "- Seleccione un Nivel - "

strSQL = "select * from FND_TP_ROLES" _
      & " order by TP_ROL"
Call sbCargaGridLocal(vGrid, 4, strSQL)


Call Formularios(Me)
Call RefrescaTags(Me)

lsw.Enabled = vGrid.Enabled
lswUsuarios.Enabled = vGrid.Enabled

End Sub

Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1 'Codigo
        vGrid.Text = CStr(rs!TP_ROL)
     Case 2 'descripcion
        vGrid.Text = CStr(rs!Descripcion)
     Case 3 ' Activo
        vGrid.Value = rs!Activo
    End Select
  
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

Me.MousePointer = vbDefault

End Sub



Private Function fxGuardar() As Long
Dim vCuenta As String, vCuentaSalida As String

On Error GoTo vError
        
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

fxGuardar = 0
If Trim(vGrid.Text) = "" Then Exit Function


vGrid.Col = 1
strSQL = "select isnull(count(*),0) as Existe from FND_TP_ROLES " _
       & " where TP_ROL = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar

  strSQL = "insert into FND_TP_ROLES(TP_ROL, Descripcion, Activo, registro_fecha, registro_usuario)" _
         & " values('" & vGrid.Text & "', '"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "', "
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ", dbo.MyGetdate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Grupo de Seguridad Tasa Preferencial:  " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update FND_TP_ROLES set descripcion = '" & vGrid.Text & "', Activo = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & " where TP_ROL = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Grupo de Seguridad Tasa Preferencial:  " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim vCodigo As String

If vPaso Or scNivel.Tag = "" Then Exit Sub

On Error GoTo vError

vCodigo = scNivel.Tag

If Item.Checked Then
   strSQL = "insert FND_TP_ROLES_PLANES(cod_operadora,cod_plan,TP_ROL,registro_usuario,registro_fecha)" _
          & " values(" & Item.Tag & ",'" & Item.Text & "','" & vCodigo & "','" & glogon.Usuario & "',dbo.MyGetdate())"
   Item.SubItems(2) = glogon.Usuario
   Item.SubItems(3) = Date

Else
   strSQL = "delete FND_TP_ROLES_PLANES where cod_operadora  = " & Item.Tag & " and  cod_Plan = '" & Item.Text _
          & "' and TP_ROL = '" & vCodigo & "'"
   
   Item.SubItems(2) = ""
   Item.SubItems(3) = ""
   
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub lswUsuarios_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim vCodigo As String

If vPaso Or scNivel.Tag = "" Then Exit Sub

On Error GoTo vError

vCodigo = scNivel.Tag

If Item.Checked Then
   strSQL = "insert FND_TP_ROLES_AUTORIZADORES(usuario,TP_ROL,registro_usuario,registro_fecha)" _
          & " values('" & Item.Text & "','" & vCodigo & "','" & glogon.Usuario & "',dbo.MyGetdate())"
   Item.SubItems(2) = glogon.Usuario
   Item.SubItems(3) = Date

Else
   strSQL = "delete FND_TP_ROLES_AUTORIZADORES where usuario = '" & Item.Text _
          & "' and TP_ROL = '" & vCodigo & "'"
   
   Item.SubItems(2) = ""
   Item.SubItems(3) = ""
   
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub txtFiltro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
  If Index = 0 Then
     Call sbPlanes_Load
  Else
     Call sbUsuarios_Load
  End If
End If

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub
If Col <> 4 Then Exit Sub

vGrid.Row = Row
vGrid.Col = 1
scNivel.Tag = vGrid.Text
vGrid.Col = 2
scNivel.Caption = vGrid.Text

Call sbDetalle_Consulta

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

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "exec spFndSeguridad_ApAnul_Delete '" & vGrid.Text & "', '" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Grupo de Seguridad Tasa Preferencial:  " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

        scNivel.Tag = ""
        scNivel.Caption = ""
        lsw.ListItems.Clear
        lswUsuarios.ListItems.Clear

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


