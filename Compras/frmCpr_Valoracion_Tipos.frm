VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCpr_Valoracion_Tipos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compras: Tipos de Valoración"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   9495
      _Version        =   1572864
      _ExtentX        =   16748
      _ExtentY        =   11245
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Esquema de Valoración"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Items"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "Label2"
      Item(1).Control(1)=   "cboEsquema"
      Item(1).Control(2)=   "scTitulo"
      Item(1).Control(3)=   "vgItems"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5775
         Left            =   0
         TabIndex        =   2
         Top             =   480
         Width           =   9495
         _Version        =   524288
         _ExtentX        =   16748
         _ExtentY        =   10186
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
         MaxCols         =   3
         ScrollBars      =   2
         SpreadDesigner  =   "frmCpr_Valoracion_Tipos.frx":0000
         VScrollSpecialType=   2
         Appearance      =   1
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vgItems 
         Height          =   4935
         Left            =   -70000
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   9495
         _Version        =   524288
         _ExtentX        =   16748
         _ExtentY        =   8705
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
         MaxCols         =   3
         ScrollBars      =   2
         SpreadDesigner  =   "frmCpr_Valoracion_Tipos.frx":05A3
         VScrollSpecialType=   2
         Appearance      =   1
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEsquema 
         Height          =   330
         Left            =   -68440
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   5895
         _Version        =   1572864
         _ExtentX        =   10398
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Left            =   -70000
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   9495
         _Version        =   1572864
         _ExtentX        =   16748
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Indique los Ítems a calificar y sus pesos en la nota final"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Left            =   -69520
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Esquema: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Esquemas de Valoración"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmCpr_Valoracion_Tipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Private Sub sbLoad_Items_List()

On Error GoTo vError

strSQL = "select VAL_ITEM, descripcion, Peso from CPR_VALORA_ITEMS Where VAL_ID = '" & cboEsquema.ItemData(cboEsquema.ListIndex) & "'" _
      & " order by VAL_ITEM"
Call sbCargaGrid(vgItems, 3, strSQL)

Exit Sub

vError:

End Sub


Private Sub cboEsquema_Click()
If vPaso Then Exit Sub

Call sbLoad_Items_List


End Sub

Private Sub Form_Activate()
vModulo = 35
End Sub

Private Sub sbLoad_List()

On Error GoTo vError

strSQL = "select VAL_ID, descripcion, Activo from CPR_VALORA_ESQUEMA" _
      & " order by VAL_ID"
Call sbCargaGrid(vGrid, 3, strSQL)

Exit Sub

vError:

End Sub

Private Sub Form_Load()

vModulo = 35

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
vGrid.AppearanceStyle = fxGridStyle

tcMain.Item(0).Selected = True

Call Formularios(Me)
Call RefrescaTags(Me)

Call sbLoad_List

End Sub


Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1
If Trim(vGrid.Text) = "" Then
  MsgBox "Indique un Código Válido!", vbExclamation
  Exit Function
End If

strSQL = "select isnull(count(*),0) as Existe from CPR_VALORA_ESQUEMA " _
       & " where VAL_ID = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into CPR_VALORA_ESQUEMA(VAL_ID, descripcion, Activo, Registro_Fecha, Registro_Usuario) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "',"
  vGrid.Col = 3
  strSQL = strSQL & CCur(vGrid.Text) & ", dbo.MyGetdate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Esquema de Valoración Id: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CPR_VALORA_ESQUEMA set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & ", Modifica_Fecha = dbo.MyGetdate(), Modifica_Usuario = '" _
        & glogon.Usuario & "' where VAL_ID = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Esquema de Valoración Id: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Function fxItem_Guardar() As Long

On Error GoTo vError

fxItem_Guardar = 0
vgItems.Row = vgItems.ActiveRow
vgItems.Col = 1

If Trim(vgItems.Text) = "" Then
  MsgBox "Indique un Código Válido!", vbExclamation
  Exit Function
End If

strSQL = "select isnull(count(*),0) as Existe from CPR_VALORA_ITEMS " _
       & " where VAL_ID = '" & cboEsquema.ItemData(cboEsquema.ListIndex) & "' and VAL_ITEM = '" & vgItems.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into CPR_VALORA_ITEMS(VAL_ID, VAL_ITEM, descripcion, Peso, Registro_Fecha, Registro_Usuario) values('" _
         & cboEsquema.ItemData(cboEsquema.ListIndex) & "', '" & vgItems.Text & "', '"
  vgItems.Col = 2
  strSQL = strSQL & vgItems.Text & "', "
  vgItems.Col = 3
  strSQL = strSQL & CCur(vgItems.Text) & ", dbo.MyGetdate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vgItems.Col = 1
  Call Bitacora("Registra", "Esquema de Valoración Id: " & cboEsquema.ItemData(cboEsquema.ListIndex) & ", Item Id: " & vgItems.Text)

Else 'Actualizar

 vgItems.Col = 2
 strSQL = "update CPR_VALORA_ITEMS set descripcion = '" & vgItems.Text & "', Peso = "
 vgItems.Col = 3
 strSQL = strSQL & CCur(vgItems.Text) & ", Modifica_Fecha = dbo.MyGetdate(), Modifica_Usuario = '" _
        & glogon.Usuario & "' where VAL_ID = '" & cboEsquema.ItemData(cboEsquema.ListIndex) _
        & "' and VAL_ITEM = '"
 vgItems.Col = 1
 strSQL = strSQL & vgItems.Text & "'"
 
 Call ConectionExecute(strSQL)

 vgItems.Col = 1
 Call Bitacora("Modifica", "Esquema de Valoración Id: " & cboEsquema.ItemData(cboEsquema.ListIndex) & ", Item Id: " & vgItems.Text)

End If
rs.Close

fxItem_Guardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vgItems_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vgItems.ActiveCol = vgItems.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxItem_Guardar
  If i = 0 Then Exit Sub
  vgItems.Row = vgItems.ActiveRow
  If vgItems.MaxRows <= vgItems.ActiveRow Then
    vgItems.MaxRows = vgItems.MaxRows + 1
    vgItems.Row = vgItems.MaxRows
  End If
End If

'Elimina
If KeyCode = vbKeyDelete Then
   vgItems.Row = vgItems.ActiveRow
   vgItems.Col = 1
     i = MsgBox("Está Seguro que desea borrar este registro?", vbYesNo)
     If i = vbYes Then
        strSQL = "delete CPR_VALORA_ITEMS where VAL_ID = '" & cboEsquema.ItemData(cboEsquema.ListIndex) & "' and Val_ITEM = '" & vgItems.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vgItems.Text
        vgItems.Col = 1
        Call Bitacora("Elimina", "Esquema de Valoración Id:  " & cboEsquema.ItemData(cboEsquema.ListIndex) & ", Item: " & vgItems.Text)
        
        Call sbLoad_Items_List
     End If

End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vgItems.MaxRows = vgItems.MaxRows + 1
    vgItems.InsertRows vgItems.ActiveRow, 1
    vgItems.Row = vgItems.ActiveRow
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Elimina
If KeyCode = vbKeyDelete Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
     i = MsgBox("Está Seguro que desea borrar este registro?", vbYesNo)
     If i = vbYes Then
        strSQL = "delete CPR_VALORA_ESQUEMA where VAL_ID = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Esquema de Valoración Id:  " & vGrid.Text)
        
        Call sbLoad_List
     End If

End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


End Sub

Private Sub sbCbo_Esquema_Load()

vPaso = True

vgItems.MaxRows = 0

strSQL = "select VAL_ID as 'IdX', Descripcion as 'ItmX' from CPR_VALORA_ESQUEMA where activo = 1 order by VAL_ID"
Call sbCbo_Llena_New(cboEsquema, strSQL, False, True)

vPaso = False

If cboEsquema.ListCount > 0 Then
   Call cboEsquema_Click
End If


End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
 Select Case Item.Index
   Case 0
    Call sbLoad_List
  Case 1
    Call sbCbo_Esquema_Load
 End Select
End Sub
