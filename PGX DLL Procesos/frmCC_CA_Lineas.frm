VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCC_CA_Lineas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cargos Automáticos: Lineas de asociadas"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   5775
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   10186
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tipos de Líneas"
      TabPicture(0)   =   "frmCC_CA_Lineas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Asignación de Código"
      TabPicture(1)   =   "frmCC_CA_Lineas.frx":06B4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cboLineas"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lswLineas"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin MSComctlLib.ListView lswLineas 
         Height          =   4215
         Left            =   -74160
         TabIndex        =   5
         Top             =   720
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7435
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   11360
         EndProperty
      End
      Begin VB.ComboBox cboLineas 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   -72360
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   5655
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4812
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   10572
         _Version        =   524288
         _ExtentX        =   18648
         _ExtentY        =   8488
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
         SpreadDesigner  =   "frmCC_CA_Lineas.frx":0E92
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Líneas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73800
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Líneas de Crédito y Retención Autorizadas"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "frmCC_CA_Lineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cboLineas_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If vPaso Or cboLineas.ListCount = 0 Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True
lswLineas.ListItems.Clear

strSQL = "select Cat.Codigo,Cat.Descripcion, isnull(Dt.Codigo,'-1') as 'Existe' " _
       & " from Catalogo Cat left join prm_Ca_Lineas_Dt Dt on Cat.codigo = Dt.Codigo and Dt.cod_Linea = '" _
       & SIFGlobal.fxCodText(cboLineas.Text) & "'" _
       & " Order by isnull(Dt.Codigo,'ZZZZZZZ'),Cat.Codigo"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswLineas.ListItems.Add(, , rs!codigo)
      itmX.SubItems(1) = rs!Descripcion
      If rs!Existe <> "-1" Then
          itmX.Checked = vbChecked
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

Private Sub Form_Activate()
vModulo = 10
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 10
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


Call Formularios(Me)
Call RefrescaTags(Me)

ssTab.Tab = 0

Call ssTab_Click(0)


End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from PRM_CA_LINEAS " _
       & " where Cod_Linea = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into PRM_CA_LINEAS(Cod_Linea,descripcion,cod_plan,Activo,Registro_Usuario,Registro_Fecha) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Cargo Automatico - Tipo Linea: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update PRM_CA_LINEAS set descripcion = '" & vGrid.Text & "',Cod_Plan = '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "', Activo = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & " where Cod_Linea = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Cargo Automatico - Tipo Linea: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function





Private Sub lswLineas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String, vMovimiento As String, vLinea As String


On Error GoTo vError

Me.MousePointer = vbHourglass
vLinea = SIFGlobal.fxCodText(cboLineas.Text)

If Item.Checked Then
    strSQL = "insert prm_ca_lineas_dt(cod_linea,codigo,registro_usuario,registro_Fecha) values('" _
           & vLinea & "','" & Item.Text & "','" & glogon.Usuario _
           & "',dbo.mygetdate())"
    vMovimiento = "Registra"
Else
    strSQL = "delete prm_ca_lineas_dt where cod_linea = '" & vLinea & "' and codigo = '" & Item.Text & "'"
    vMovimiento = "Elimina"
End If

Call ConectionExecute(strSQL)
Call Bitacora(vMovimiento, "Cargo Automatico: Linea:" & vLinea & " Cod:" & Item.Text)

Me.MousePointer = vbDefault


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case ssTab.Tab
    Case 0
    
        strSQL = "select Cod_Linea,descripcion,cod_plan, activo from PRM_CA_LINEAS" _
              & " order by Cod_Linea"
        Call sbCargaGrid(vGrid, 4, strSQL)
    
    Case 1
        
        strSQL = "select rtrim(Cod_Linea) + ' - ' + descripcion as 'ItmX' from PRM_CA_LINEAS where activo = 1"
        
        vPaso = True
            Call sbLlenaCbo(cboLineas, strSQL, False, False)
        vPaso = False
        
End Select

 Me.MousePointer = vbDefault


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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
        strSQL = "delete PRM_CA_LINEAS where Cod_Linea = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Cargo Automatico - Tipo Linea: " & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





