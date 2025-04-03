VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_Zonas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Zonas - Mantenimiento"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lswI 
      Height          =   3012
      Left            =   6000
      TabIndex        =   5
      Top             =   4920
      Width           =   6012
      _Version        =   1441793
      _ExtentX        =   10604
      _ExtentY        =   5313
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
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.ListView lswU 
      Height          =   3012
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   6012
      _Version        =   1441793
      _ExtentX        =   10604
      _ExtentY        =   5313
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
      ShowBorder      =   0   'False
   End
   Begin VB.Timer Timerx 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2772
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   9012
      _Version        =   524288
      _ExtentX        =   15896
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
      MaxCols         =   497
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_Zonas.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scZona 
      Height          =   372
      Left            =   0
      TabIndex        =   2
      Top             =   4200
      Width           =   12012
      _Version        =   1441793
      _ExtentX        =   21188
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   372
      Left            =   0
      TabIndex        =   3
      Top             =   4560
      Width           =   6012
      _Version        =   1441793
      _ExtentX        =   10604
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Usuario [Ejectivos]"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mantenimiento de Zonas de atención"
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
      Index           =   2
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   7068
   End
   Begin VB.Image imgBanner 
      Height          =   1248
      Left            =   0
      Top             =   0
      Width           =   12492
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Left            =   6000
      TabIndex        =   6
      Top             =   4560
      Width           =   6012
      _Version        =   1441793
      _ExtentX        =   10604
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Entidades [Centros de Trabajo]"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAF_Zonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub Form_Activate()
vModulo = 1

End Sub

Private Sub sbZonas_Consulta()
Dim strSQL As String

vPaso = True
    strSQL = "select cod_zona,descripcion,activa,'' from afi_zonas" _
          & " order by cod_zona"
    Call sbCargaGrid(vGrid, 4, strSQL)
vPaso = False

scZona.Caption = ""
scZona.Tag = ""

lswU.ListItems.Clear
lswI.ListItems.Clear


End Sub


Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lswU.ColumnHeaders
  .Clear
  .Add , , "Id", 2300
  .Add , , "Descripción", 3600
End With

lswU.Checkboxes = True

With lswI.ColumnHeaders
  .Clear
  .Add , , "Id", 900
  .Add , , "Desc. Corta ", 1400, vbCenter
  .Add , , "Descripción", 3600
End With
lswI.Checkboxes = True

vGrid.AppearanceStyle = fxGridStyle

Call Formularios(Me)
Call RefrescaTags(Me)


lswU.Enabled = vGrid.Enabled
lswI.Enabled = vGrid.Enabled

End Sub



Private Sub sbZona_Elimina(pZona As String)
Dim strSQL As String

On Error GoTo vError

strSQL = "delete afi_zonas where cod_zona = '" & pZona & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Elimina", "Zonas : " & pZona)

Call sbZonas_Consulta

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbZonas_Detalle()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If scZona.Tag = "" Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True


strSQL = "exec spAfi_Zonas_Usuario_Asigna_Consulta '" & scZona.Tag & "'"
Call OpenRecordSet(rs, strSQL)

lswU.ListItems.Clear

Do While Not rs.EOF
   Set itmX = lswU.ListItems.Add(, , rs!Codigo)
       itmX.SubItems(1) = rs!Descripcion
       itmX.Checked = IIf(rs!Asignado = 1, True, False)
       
    rs.MoveNext
Loop
rs.Close


strSQL = "exec spAfi_Zonas_Inst_Asigna_Consulta '" & scZona.Tag & "'"
Call OpenRecordSet(rs, strSQL)

lswI.ListItems.Clear

Do While Not rs.EOF
   Set itmX = lswI.ListItems.Add(, , rs!Codigo)
       itmX.SubItems(1) = rs!Desc_Corta & ""
       itmX.SubItems(2) = rs!Descripcion
       itmX.Checked = IIf(rs!Asignado = 1, True, False)
       
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

Private Sub lswI_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswI.SortKey = ColumnHeader.Index - 1
  If lswI.SortOrder = 0 Then lswI.SortOrder = 1 Else lswI.SortOrder = 0
  lswI.Sorted = True
End Sub

Private Sub lswI_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spAfi_Zonas_Inst_Asigna_Registra '" & scZona.Tag & "'," & Item.Text & ", '" _
        & glogon.Usuario & "','" & IIf(Item.Checked, "I", "E") & "'"
Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswU_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswU.SortKey = ColumnHeader.Index - 1
  If lswU.SortOrder = 0 Then lswU.SortOrder = 1 Else lswU.SortOrder = 0
  lswU.Sorted = True
End Sub

Private Sub lswU_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spAfi_Zonas_Usuario_Asigna_Registra '" & scZona.Tag & "','" & Item.Text & "', '" _
        & glogon.Usuario & "','" & IIf(Item.Checked, "I", "E") & "'"
Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbZonas_Consulta
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub

vGrid.Row = Row
vGrid.Col = 1
If vGrid.Text <> "" Then
  scZona.Tag = vGrid.Text
  
  vGrid.Col = 2
  scZona.Caption = vGrid.Text
  
  vGrid.Col = 3
  If vGrid.Value = vbChecked Then
      Call sbZonas_Detalle
  End If
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

If vGrid.ActiveCol = (vGrid.MaxCols - 1) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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

'Elimina Linea
If KeyCode = vbKeyDelete Then
     
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
                
        Call sbZona_Elimina(vGrid.Text)
     End If

End If

End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select coalesce(count(*),0) as Existe from afi_zonas " _
       & " where cod_zona = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function

  strSQL = "insert into afi_zonas(cod_zona,descripcion, activa, registro_fecha, registro_usuario) values('" _
         & Trim(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & Trim(vGrid.Text) & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",dbo.mygetdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Zonas : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update afi_zonas set descripcion = '" & vGrid.Text & "', Activa = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value
 vGrid.Col = 1
 strSQL = strSQL & " where cod_zona = '" & vGrid.Text & "'"

 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", " Zonas : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function
