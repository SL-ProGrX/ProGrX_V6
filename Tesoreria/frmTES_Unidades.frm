VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmTES_Unidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unidades de Negocios"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   9240
   Begin TabDlg.SSTab ssTab 
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Unidades"
      TabPicture(0)   =   "frmTES_Unidades.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Centros de Costos"
      TabPicture(1)   =   "frmTES_Unidades.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lsw"
      Tab(1).Control(1)=   "lswAsg"
      Tab(1).Control(2)=   "lblX"
      Tab(1).Control(3)=   "Label2"
      Tab(1).Control(4)=   "Label1"
      Tab(1).ControlCount=   5
      Begin MSComctlLib.ListView lsw 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   2
         Top             =   1440
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   7223
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Unidad"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descipción"
            Object.Width           =   5362
         EndProperty
      End
      Begin MSComctlLib.ListView lswAsg 
         Height          =   4095
         Left            =   -70440
         TabIndex        =   3
         Top             =   1440
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   7223
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "C.C."
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descipción"
            Object.Width           =   5362
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4935
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   8535
         _Version        =   524288
         _ExtentX        =   15055
         _ExtentY        =   8705
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   487
         ScrollBars      =   2
         SpreadDesigner  =   "frmTES_Unidades.frx":0038
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   0
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -70440
         TabIndex        =   6
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidades"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   5
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   $"frmTES_Unidades.frx":05AC
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
         Left            =   -74760
         TabIndex        =   4
         Top             =   480
         Width           =   8535
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTES_Unidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 9
vGrid.AppearanceStyle = fxGridStyle


Set Me.Icon = frmContenedor.Icon
Call sbToolBarIconos(tlb, False)

Call Formularios(Me)
Call RefrescaTags(Me)

If tlb.Buttons(1).Enabled = False Then vGrid.Enabled = False

strSQL = "select cod_unidad,descripcion,case when estado = 'A' then 1 else 0 end As Estado,ALTERNO from tes_unidades" _
      & " order by cod_unidad"
Call sbCargaGrid(vGrid, 4, strSQL)

ssTab.Tab = 0


End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select coalesce(count(*),0) as Existe from tes_unidades " _
       & " where cod_unidad = '" & vGrid.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs!existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into tes_unidades(cod_unidad,descripcion,estado,alterno) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "','"
  vGrid.Col = 3
  strSQL = strSQL & IIf((vGrid.Value = 0), "I", "A") & "','"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Text & "')"

  glogon.Conection.Execute strSQL

  vGrid.Col = 1
  Call Bitacora("Registra", "Unidad Negocio : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update tes_unidades set descripcion = '" & vGrid.Text & "',estado = '"
 vGrid.Col = 3
 strSQL = strSQL & IIf((vGrid.Value = 0), "I", "A") & "',alterno  = '"
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Text & "' where cod_unidad = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 glogon.Conection.Execute strSQL

 vGrid.Col = 1
 Call Bitacora("Modifica", "Unidad Negocio : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function

Private Sub sbCargaAsignacion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem
  
Me.MousePointer = vbHourglass
  
lswAsg.ListItems.Clear
  
strSQL = "select C.*,A.cod_cc as ExisteX" _
       & " from tes_centros_costos C left join Tes_Unidades_CC A on C.cod_cc = A.cod_cc" _
       & " and A.cod_unidad = '" & lblX.Tag & "' order by ExisteX desc,C.cod_cc"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 Set itmX = lswAsg.ListItems.Add(, , rs!cod_cc)
     itmX.SubItems(1) = rs!Descripcion
     itmX.Checked = IIf(IsNull(rs!ExisteX), vbUnchecked, vbChecked)
 If itmX.Checked Then itmX.ForeColor = vbBlue
 rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault
  
End Sub


Private Sub lsw_Click()

If lsw.ListItems.Count = 0 Then Exit Sub

lblX.Tag = lsw.SelectedItem
lblX.Caption = lsw.SelectedItem.SubItems(1)

Call sbCargaAsignacion

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo vError
    
    lsw.SortKey = ColumnHeader.Index - 1
    
    If (lsw.SortOrder = lvwAscending) Then
        lsw.SortOrder = lvwDescending
    Else
        lsw.SortOrder = lvwAscending
    End If
    
    lsw.Sorted = True
    Exit Sub

vError:
   MsgBox "Ocurrió un error al ordenar los datos de la columna seleccionada.", vbCritical
End Sub

Private Sub lswAsg_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo vError
    
    lswAsg.SortKey = ColumnHeader.Index - 1
    
    If (lswAsg.SortOrder = lvwAscending) Then
        lswAsg.SortOrder = lvwDescending
    Else
        lswAsg.SortOrder = lvwAscending
    End If
    
    lswAsg.Sorted = True
    Exit Sub

vError:
   MsgBox "Ocurrió un error al ordenar los datos de la columna seleccionada.", vbCritical
End Sub

Private Sub lswAsg_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

On Error GoTo vError

If Item.Checked Then
  strSQL = "insert tes_unidades_cc(cod_unidad,cod_cc) values('" & lblX.Tag _
         & "','" & Item.Text & "')"
Else
  strSQL = "delete tes_unidades_cc where cod_unidad = '" & lblX.Tag _
         & "' and cod_cc = '" & Item.Text & "'"

End If

glogon.Conection.Execute strSQL

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Select Case ssTab.Tab
  Case 0
        strSQL = "select cod_unidad,descripcion,case when estado = 'A' then 1 else 0 end As Estado,ALTERNO from tes_unidades" _
            & " order by cod_unidad"
        Call sbCargaGrid(vGrid, 4, strSQL)

  Case 1
       lsw.ListItems.Clear
       lswAsg.ListItems.Clear
       lblX.Caption = ">> Seleccione una Unidad <<"
       lblX.Tag = "(x)"
       
       strSQL = "select cod_unidad,descripcion from tes_unidades where estado = 'A'"
       rs.Open strSQL, glogon.Conection, adOpenStatic
       Do While Not rs.EOF
         Set itmX = lsw.ListItems.Add(, , rs!cod_unidad)
             itmX.SubItems(1) = rs!Descripcion
         rs.MoveNext
       Loop
       rs.Close

End Select

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String

On Error Resume Next

Select Case UCase(Button.Key)
  Case "NUEVO"
    vGrid.MaxRows = vGrid.MaxRows + 1

  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete tes_unidades where cod_unidad = '" & vGrid.Text & "'"
        glogon.Conection.Execute strSQL
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Unidad Negocio : " & vGrid.Text)
        
        vGrid.Col = 1
        strSQL = "select cod_unidad,descripcion,case when estado = 'A' then 1 else 0 end As Estado, Alterno from tes_unidades" _
              & " order by cod_unidad"
        Call sbCargaGrid(vGrid, 4, strSQL)
     End If
  
  Case "REPORTES"
'     Call sbReportes("Caracteristicas", Me)

  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
'MsgBox "Columna : " & vGrid.Col
'MsgBox "Columna Activa: " & vGrid.ActiveCol
'MsgBox "Fila : " & vGrid.Row
'MsgBox "Fila Activa: " & vGrid.ActiveRow

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
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


End Sub






