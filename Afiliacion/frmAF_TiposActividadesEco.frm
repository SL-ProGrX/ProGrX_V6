VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_TiposActividadesEco 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actividades Económicas"
   ClientHeight    =   7884
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   9732
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7884
   ScaleWidth      =   9732
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   120
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6612
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   9732
      _Version        =   1245187
      _ExtentX        =   17166
      _ExtentY        =   11663
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
      Item(0).Caption =   "Actividades"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Sub Actividades"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "vGridSub"
      Item(1).Control(1)=   "txtNombre"
      Item(1).Control(2)=   "txtCodigo"
      Item(1).Control(3)=   "Label2"
      Begin FPSpreadADO.fpSpread vGridSub 
         Height          =   5292
         Left            =   -69880
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   9492
         _Version        =   524288
         _ExtentX        =   16743
         _ExtentY        =   9335
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
         MaxCols         =   496
         ScrollBars      =   2
         SpreadDesigner  =   "frmAF_TiposActividadesEco.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5892
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   9492
         _Version        =   524288
         _ExtentX        =   16743
         _ExtentY        =   10393
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
         MaxCols         =   496
         ScrollBars      =   2
         SpreadDesigner  =   "frmAF_TiposActividadesEco.frx":054F
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   312
         Left            =   -67120
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   6732
         _Version        =   1245187
         _ExtentX        =   11874
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   -68320
         TabIndex        =   5
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   480
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Left            =   -69640
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Actividad"
         BackColor       =   -2147483633
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
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Actividades Económicas"
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
      Height          =   480
      Index           =   0
      Left            =   1880
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_TiposActividadesEco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As Integer, vPaso As Boolean

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()

vModulo = 1
tcMain.Item(0).Selected = True

vGrid.AppearanceStyle = fxGridStyle
vGridSub.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


End Sub



Private Sub TimerX_Timer()


TimerX.Interval = 0
TimerX.Enabled = False


End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_actividad"
  gBusquedas.Orden = "cod_actividad"
  gBusquedas.Consulta = "select cod_actividad,descripcion from AFI_ACTIVIDADES_ECO"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  If txtCodigo.Text <> "" Then Call sbSubActividad_Load
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  
  strSQL = "select isnull(count(*),0) as Existe from AFI_ACTIVIDADES_ECO where cod_actividad = '" & vGrid.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  
  If rs!Existe = 0 Then
     
     vGrid.Col = 1
     strSQL = "insert AFI_ACTIVIDADES_ECO(cod_actividad,descripcion,Activa)" _
             & " values('" & vGrid.Text & "','"
     vGrid.Col = 2
     strSQL = strSQL & vGrid.Text & "',"
     vGrid.Col = 3
     strSQL = strSQL & vGrid.Value & ")"
     
     
  Else
     vGrid.Col = 2
     strSQL = "update AFI_ACTIVIDADES_ECO set descripcion = '" & vGrid.Text & "',Activa = "
     vGrid.Col = 3
     strSQL = strSQL & vGrid.Value & " where cod_actividad = '"
     vGrid.Col = 1
     strSQL = strSQL & vGrid.Value & "'"
  End If
  
  Call ConectionExecute(strSQL)
  
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
          strSQL = "delete AFI_ACTIVIDADES_ECO where cod_actividad = '" & vGrid.Text & "'"
          Call ConectionExecute(strSQL)
        
          vGrid.DeleteRows vGrid.ActiveRow, 1
          If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
          vGrid.Row = vGrid.ActiveRow

  End If
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Resume
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Item.Index
  Case 0 'Actividades
     strSQL = "select cod_actividad,descripcion,activa from AFI_ACTIVIDADES_ECO" _
            & " order by cod_actividad"
     Call sbCargaGrid(vGrid, 3, strSQL)
     
  Case 1 'Sub Actividades
    Call sbSubActividad_Load
End Select

Me.MousePointer = vbDefault
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Me.MousePointer = vbDefault

End Sub

Private Sub vGridSub_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If vGridSub.ActiveCol = vGridSub.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGridSub.Row = vGridSub.ActiveRow
  vGridSub.Col = 1
  
  strSQL = "select isnull(count(*),0) as Existe from AFI_ACTIVIDADES_ECO_SUB where cod_actividad = '" _
         & txtCodigo & "' and COD_SUB_ACT = '" & vGridSub.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  
  If rs!Existe = 0 Then
     
     vGridSub.Col = 1
     strSQL = "insert AFI_ACTIVIDADES_ECO_SUB(cod_actividad,COD_SUB_ACT,descripcion,Activa)" _
             & " values('" & txtCodigo & "','" & vGridSub.Text & "','"
     vGridSub.Col = 2
     strSQL = strSQL & vGridSub.Text & "',"
     vGridSub.Col = 3
     strSQL = strSQL & vGridSub.Value & ")"
     
  Else
     vGridSub.Col = 2
     strSQL = "update AFI_ACTIVIDADES_ECO_SUB set descripcion = '" & Trim(vGridSub.Text) & "',Activa = "
     vGridSub.Col = 3
     strSQL = strSQL & vGridSub.Value & " Where cod_actividad = '" & txtCodigo _
            & "' and COD_SUB_ACT = '"
     vGridSub.Col = 1
     strSQL = strSQL & vGridSub.Text & "'"
     
  
  End If
  
  Call ConectionExecute(strSQL)
  
  If vGridSub.MaxRows <= vGridSub.ActiveRow Then
    vGridSub.MaxRows = vGridSub.MaxRows + 1
    vGridSub.Row = vGridSub.MaxRows
  End If
  
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridSub.MaxRows = vGridSub.MaxRows + 1
    vGridSub.InsertRows vGridSub.ActiveRow, 1
    vGridSub.Row = vGridSub.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
  i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
  If i = vbYes Then
        vGridSub.Row = vGridSub.ActiveRow
        vGridSub.Col = 1
        strSQL = "delete AFI_ACTIVIDADES_ECO_SUB where cod_actividad = '" & txtCodigo _
               & "' and COD_SUB_ACT = '" & vGridSub.Text & "'"
        Call ConectionExecute(strSQL)
        
        vGridSub.DeleteRows vGridSub.ActiveRow, 1
        If vGridSub.MaxRows > 1 Then vGridSub.MaxRows = vGridSub.MaxRows - 1
        vGridSub.Row = vGridSub.ActiveRow
  End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbSubActividad_Load()
Dim strSQL As String

vGridSub.MaxRows = 0

If vPaso Then Exit Sub
If txtCodigo.Text = "" Then Exit Sub

strSQL = "select COD_SUB_ACT,descripcion,Activa from AFI_ACTIVIDADES_ECO_SUB" _
       & " where cod_actividad = '" & txtCodigo.Text & "' order by COD_SUB_ACT"
Call sbCargaGrid(vGridSub, 3, strSQL)

End Sub




