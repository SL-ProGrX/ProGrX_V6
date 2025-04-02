VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmInvDepartamentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departamentos"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   8250
   Begin TabDlg.SSTab ssTab 
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Definición"
      TabPicture(0)   =   "frmInvDepartamentos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Asignación de Líneas"
      TabPicture(1)   =   "frmInvDepartamentos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(0)"
      Tab(1).Control(1)=   "lbl"
      Tab(1).Control(2)=   "lswLineas"
      Tab(1).Control(3)=   "lswDep"
      Tab(1).ControlCount=   4
      Begin MSComctlLib.ListView lswDep 
         Height          =   1452
         Left            =   -74760
         TabIndex        =   2
         Top             =   600
         Width           =   7692
         _ExtentX        =   13573
         _ExtentY        =   2566
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   9067
         EndProperty
      End
      Begin MSComctlLib.ListView lswLineas 
         Height          =   3492
         Left            =   -74760
         TabIndex        =   3
         Top             =   2280
         Width           =   7692
         _ExtentX        =   13573
         _ExtentY        =   6165
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   9067
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5172
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   7572
         _Version        =   524288
         _ExtentX        =   13356
         _ExtentY        =   9123
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
         MaxCols         =   487
         ScrollBars      =   2
         SpreadDesigner  =   "frmInvDepartamentos.frx":0038
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Líneas Asignadas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   -74760
         TabIndex        =   5
         Top             =   2040
         Width           =   7692
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Departamentos Disponibles"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Index           =   0
         Left            =   -74760
         TabIndex        =   4
         Top             =   360
         Width           =   7692
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8250
      _ExtentX        =   14552
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LisDep"
                  Text            =   "Listado de Departamentos"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LisLineas"
                  Text            =   "Líneas x Departamento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInvDepartamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDepartamento As String

Private Sub Form_Activate()
vModulo = 32
End Sub

Private Sub Form_Load()
Dim strSQL As String

Set Me.Icon = frmContenedor.Icon

ssTab.Tab = 0

vModulo = 32
Call Formularios(Me)

vGrid.AppearanceStyle = fxGridStyle

Call sbToolBarIconos(tlb)

strSQL = "select cod_departamento,descripcion,activo from pv_departamentos" _
      & " order by cod_departamento"
Call sbCargaGrid(vGrid, 3, strSQL)

Call RefrescaTags(Me)
If tlb.Buttons(1).Enabled = False Then vGrid.Enabled = False

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from pv_departamentos " _
       & " where cod_departamento = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into pv_departamentos(cod_departamento,descripcion,activo) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Value & ")"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Departamento : " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update pv_departamentos set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.col = 3
 strSQL = strSQL & vGrid.Value & " where cod_departamento = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Modifica", "Departamento : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub lswDep_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

vDepartamento = lswDep.SelectedItem.Text

strSQL = "select C.*,L.cod_departamento" _
       & " from pv_prod_clasifica C left join pv_lineasdep L" _
       & " on C.cod_prodclas = L.cod_prodclas and L.cod_departamento = '" _
       & lswDep.SelectedItem.Text & "' order by L.cod_departamento desc"
Call OpenRecordSet(rs, strSQL, 0)

lswLineas.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lswLineas.ListItems.Add(, , rs!cod_prodclas)
     itmX.SubItems(1) = rs!Descripcion
     itmX.Checked = IIf(IsNull(rs!cod_departamento), vbUnchecked, vbChecked)
     itmX.ForeColor = IIf(IsNull(rs!cod_departamento), vbBlack, vbBlue)
 rs.MoveNext
Loop
rs.Close

lbl.Caption = "Líneas Asignadas a " & lswDep.SelectedItem.SubItems(1)

End Sub


Private Sub lswLineas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

If Item.Checked = True Then
   strSQL = "insert PV_LINEASDEP(cod_departamento,cod_prodclas) values('" _
          & vDepartamento & "'," & Item.Text & ")"
Else
   strSQL = "delete PV_LINEASDEP where cod_departamento = '" & vDepartamento _
          & "' and cod_prodclas = " & Item.Text
End If
Call ConectionExecute(strSQL)

End Sub

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
        vGrid.col = 1
        strSQL = "delete pv_departamentos where cod_departamento = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.col = 2
        Call Bitacora("Elimina", "Departamento : " & strSQL & " - " & vGrid.Text)
        
        vGrid.col = 1
        strSQL = "select cod_departamento,descripcion from pv_departamentos" _
              & " order by cod_departamento"
        Call sbCargaGrid(vGrid, 2, strSQL)

     End If
  
  Case "REPORTES"
'     Call sbReportes("Caracteristicas", Me)

  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub


Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
  Case "LisDep"
     Call sbInvReportes("Departamentos", "Departamentos", "Listado", "")
  Case "LisLineas"
     Call sbInvReportes("DeptLineas", "Líneas x Departamentos", "Listado", "")
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

Private Sub SSTab_Click(PreviousTab As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX  As ListItem, curSum As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case ssTab.Tab
  Case 1 'Asignación
    lswDep.ListItems.Clear
    lswLineas.ListItems.Clear
    strSQL = "select cod_departamento,descripcion from pv_departamentos order by cod_departamento"
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
      Set itmX = lswDep.ListItems.Add(, , rs!cod_departamento)
          itmX.SubItems(1) = rs!Descripcion
      rs.MoveNext
    Loop
    rs.Close
End Select

vError:
Me.MousePointer = vbDefault

End Sub


