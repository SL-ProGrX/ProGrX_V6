VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmFNDGrupos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos de Planes"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   10320
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3492
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   10092
      _Version        =   1572864
      _ExtentX        =   17801
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
      Appearance      =   16
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2772
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   9972
      _Version        =   524288
      _ExtentX        =   17590
      _ExtentY        =   4890
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
      MaxCols         =   495
      ScrollBars      =   2
      SpreadDesigner  =   "frmFNDGrupos.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scPlan 
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   10092
      _Version        =   1572864
      _ExtentX        =   17801
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
      Caption         =   "Grupos o Clasificación de Planes de ahorros"
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
      Height          =   372
      Index           =   3
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   6372
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmFNDGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

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
    .Add , , "Plan", 2000
    .Add , , "Descripción", 6000
    .Add , , "Operadora Id", 1500
End With
lsw.Checkboxes = True

scPlan.Tag = ""
scPlan.Caption = "Seleccione un Grupo para asignación de Planes"

vPaso = True
    strSQL = "select cod_grupo,descripcion,categoria,interno,prioridad,0 from fnd_grupos order by cod_grupo"
    Call sbCargaGridLocal(vGrid, 6, strSQL)
vPaso = False

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

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
     Case 1
        vGrid.Text = CStr(rs!Cod_Grupo)
     Case 2
        vGrid.Text = CStr(rs!Descripcion)
     Case 3
        Select Case rs!Categoria
          Case "00"
             vGrid.Text = "00 Modelo General"
          Case "01"
             vGrid.Text = "01 Patrimonio"
          Case "02"
             vGrid.Text = "02 Ahorros"
          Case "03"
             vGrid.Text = "03 Fondos"
          Case "04"
             vGrid.Text = "04 Administrados"
        End Select
     Case 4
        vGrid.Text = CStr(rs!interno)
     Case 5
        vGrid.Text = CStr(rs!prioridad)
    End Select
  
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

Me.MousePointer = vbDefault

End Sub



Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from fnd_grupos " _
       & " where cod_grupo = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)


If rs!Existe = 0 Then 'Insertar
  vGrid.Col = 1
  strSQL = "insert into fnd_grupos(cod_grupo,descripcion,categoria,interno,prioridad) values('"
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & Mid(vGrid.Text, 1, 2) & "',"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ",'"
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Text & "')"
  

  Call ConectionExecute(strSQL)


  vGrid.Col = 1
  Call Bitacora("Registra", "Grupo de Fondo Cod: " & vGrid.Text)


Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update fnd_grupos set descripcion = '" & vGrid.Text & "',categoria = '"
 vGrid.Col = 3
 strSQL = strSQL & Mid(vGrid.Text, 1, 2) & "',interno = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & ",prioridad = '"
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Text & "' where cod_grupo = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 
 Call Bitacora("Modifica", "Grupo de Fondo Cod: " & vGrid.Text)

End If

rs.Close
fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String

On Error Resume Next

Select Case UCase(Button.Key)
  Case "NUEVO"
    vGrid.MaxRows = vGrid.MaxRows + 1

  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = 6 Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete fnd_grupos where cod_grupo = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 2
        Call Bitacora("Elimina", "Grupo de Fondo : " & strSQL & " - " & vGrid.Text)
        vGrid.Col = 1
        strSQL = "select cod_grupo,descripcion,categoria,interno,prioridad from fnd_grupos order by cod_grupo"
        Call sbCargaGridLocal(vGrid, 5, strSQL)
     End If
  Case "REPORTES"

  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "update fnd_Planes set cod_Grupo  = '" & scPlan.Tag & "'" _
          & " where cod_plan = '" & Item.Text & "' and cod_Operadora = " & Item.SubItems(2)
   Call ConectionExecute(strSQL)
   
   Call Bitacora("Aplica", "Asignación del Plan: " & Item.Text & " al Grupo: " & scPlan.Tag)
   
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbPlanes_Consulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

lsw.ListItems.Clear

If vPaso Then Exit Sub
If scPlan.Tag = "" Then Exit Sub

strSQL = "select P.cod_operadora,P.COD_PLAN,P.descripcion,G.cod_Grupo" _
       & " from fnd_Planes P left join fnd_grupos G on P.cod_Grupo = G.cod_Grupo" _
       & " and P.cod_Grupo = '" & scPlan.Tag & "'" _
       & " where P.Estado = 'A'"
Call OpenRecordSet(rs, strSQL)

vPaso = True

With lsw.ListItems
  Do While Not rs.EOF
     Set itmX = .Add(, , rs!COD_PLAN)
         itmX.SubItems(1) = rs!Descripcion
         itmX.SubItems(2) = rs!COD_OPERADORA
         
     If IsNull(rs!Cod_Grupo) Then
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

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub
If Col <> 6 Then Exit Sub

vGrid.Row = Row
vGrid.Col = 1
scPlan.Tag = vGrid.Text
vGrid.Col = 2
scPlan.Caption = vGrid.Text

Call sbPlanes_Consulta

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol >= (vGrid.MaxCols - 1) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
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


End Sub



