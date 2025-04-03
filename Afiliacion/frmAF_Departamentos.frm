VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_Departamentos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departamentos y Secciones"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   8565
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7815
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   8295
      _Version        =   1441793
      _ExtentX        =   14631
      _ExtentY        =   13785
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
      SelectedItem    =   1
      Item(0).Caption =   "Departamentos"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "txtFiltro(0)"
      Item(1).Caption =   "Secciones"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "vGridSec"
      Item(1).Control(1)=   "cboDept"
      Item(1).Control(2)=   "txtFiltro(1)"
      Begin FPSpreadADO.fpSpread vGridSec 
         Height          =   6255
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   7815
         _Version        =   524288
         _ExtentX        =   13785
         _ExtentY        =   11033
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
         SpreadDesigner  =   "frmAF_Departamentos.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6735
         Left            =   -69640
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   7695
         _Version        =   524288
         _ExtentX        =   13573
         _ExtentY        =   11880
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
         SpreadDesigner  =   "frmAF_Departamentos.frx":04ED
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboDept 
         Height          =   330
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   7815
         _Version        =   1441793
         _ExtentX        =   13785
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   330
         Index           =   0
         Left            =   -69640
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   7575
         _Version        =   1441793
         _ExtentX        =   13361
         _ExtentY        =   582
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   330
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   7815
         _Version        =   1441793
         _ExtentX        =   13785
         _ExtentY        =   582
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   1920
      TabIndex        =   5
      Top             =   480
      Width           =   6372
      _Version        =   1441793
      _ExtentX        =   11245
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Instituciones:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmAF_Departamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As Integer, vPaso As Boolean

Private Sub cbo_Click()
If vPaso Then Exit Sub

vCodigo = cbo.ItemData(cbo.ListIndex)
Call sbDepartamentos_Load

End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_institucion"
  gBusquedas.Orden = "cod_institucion"
  gBusquedas.Consulta = "select cod_institucion,desc_Corta, descripcion from instituciones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    Call sbCboAsignaDato(cbo, gBusquedas.Resultado3, True, gBusquedas.Resultado)
  End If
  
End If

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 1
vGrid.AppearanceStyle = fxGridStyle
vGridSec.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vPaso = True

strSQL = "Select cod_institucion as 'IdX',descripcion as 'ItmX' from instituciones"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

vPaso = False

Call cbo_Click

End Sub



Private Sub txtFiltro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Index = 0 Then
    Call sbDepartamentos_Load
End If

If KeyCode = vbKeyReturn And Index = 1 Then
    Call cboDept_Click
End If


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  
  strSQL = "select isnull(count(*),0) as Existe from AfDepartamentos where cod_institucion = " _
         & vCodigo & " and cod_departamento = '" & vGrid.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  
  If rs!Existe = 0 Then
     
     vGrid.col = 1
     strSQL = "insert AfDepartamentos(cod_institucion,cod_departamento,descripcion)" _
             & " values(" & vCodigo & ",'" & vGrid.Text & "','"
     vGrid.col = 2
     strSQL = strSQL & vGrid.Text & "')"
     
    vGrid.col = 1
     strSQL = strSQL & Space(10) & "insert AfSecciones(cod_institucion,cod_departamento,cod_seccion,descripcion)" _
             & " values(" & vCodigo & ",'" & vGrid.Text & "','','Sin Descripción')"
     
  Else
     vGrid.col = 2
     strSQL = "update AfDepartamentos set descripcion = '" & Trim(vGrid.Text) & "'"
     vGrid.col = 1
     strSQL = strSQL & " where cod_institucion = " & vCodigo & " and cod_departamento = '" _
            & vGrid.Text & "'"
  End If
  
  Call ConectionExecute(strSQL)
  
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
  
End If

If KeyCode = vbKeyDelete Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  
   i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
   If i = vbYes Then
  
          strSQL = "delete AfSecciones where cod_institucion = " & vCodigo _
                 & " and cod_departamento = '" & vGrid.Text & "'"
          
          strSQL = strSQL & Space(10) & "delete AfDepartamentos where cod_institucion = " & vCodigo _
                 & " and cod_departamento = '" & vGrid.Text & "'"
          
          Call ConectionExecute(strSQL)
        
          strSQL = "select cod_departamento,descripcion from AfDepartamentos" _
                 & " where cod_institucion = " & vCodigo & " order by cod_departamento"
          Call sbCargaGrid(vGrid, 2, strSQL)
  End If
    
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbDepartamentos_Load()
Dim strSQL As String

On Error GoTo vError

tcMain.Item(0).Selected = True

Me.MousePointer = vbHourglass
     
txtFiltro(0).Text = fxSysCleanTxtInject(txtFiltro(0).Text)
     
strSQL = "select cod_departamento,descripcion from AFDepartamentos" _
       & " where cod_institucion = " & vCodigo _
       & " and descripcion like '%" & txtFiltro(0).Text & "%'" _
       & " order by cod_departamento"
Call sbCargaGrid(vGrid, 2, strSQL)


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Item.Index
  Case 0 'Departamentos
    Call sbDepartamentos_Load
     
  Case 1 'Secciones
     vPaso = True
     strSQL = "select cod_departamento as 'IdX',rtrim(descripcion) as 'ItmX' from AFDepartamentos" _
            & " where cod_institucion = " & vCodigo & " order by cod_departamento"
     
     Call sbCbo_Llena_New(cboDept, strSQL, False, True)
     vPaso = False
     
     Call cboDept_Click
     
End Select

vError:
Me.MousePointer = vbDefault

End Sub

Private Sub vGridSec_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If vGridSec.ActiveCol = vGridSec.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGridSec.Row = vGridSec.ActiveRow
  vGridSec.col = 1
  
  strSQL = "select isnull(count(*),0) as Existe from afSecciones where cod_institucion = " _
         & vCodigo & " and cod_departamento = '" & cboDept.ItemData(cboDept.ListIndex) _
         & "' and cod_seccion = '" & vGridSec.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  
  If rs!Existe = 0 Then
     
     vGridSec.col = 1
     strSQL = "insert AfSecciones(cod_institucion,cod_departamento,cod_seccion,descripcion)" _
             & " values(" & vCodigo & ",'" & cboDept.ItemData(cboDept.ListIndex) & "','" & vGridSec.Text & "','"
     vGridSec.col = 2
     strSQL = strSQL & vGridSec.Text & "')"
     
  Else
     vGridSec.col = 2
     strSQL = "update AfSecciones set descripcion = '" & Trim(vGridSec.Text) & "'"
     vGridSec.col = 1
     strSQL = strSQL & " where cod_institucion = " & vCodigo & " and cod_departamento = '" _
            & cboDept.ItemData(cboDept.ListIndex) & "' and cod_seccion = '" & vGridSec.Text & "'"
  End If
  
  Call ConectionExecute(strSQL)
  
  If vGridSec.MaxRows <= vGridSec.ActiveRow Then
    vGridSec.MaxRows = vGridSec.MaxRows + 1
    vGridSec.Row = vGridSec.MaxRows
  End If
  
End If

If KeyCode = vbKeyDelete Then
  vGridSec.Row = vGridSec.ActiveRow
  vGridSec.col = 1
  
   i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
   If i = vbYes Then
          strSQL = "delete AfSecciones where cod_institucion = " & vCodigo _
                 & " and cod_departameto = '" & cboDept.ItemData(cboDept.ListIndex) & "' and cod_seccion = '" & vGridSec.Text & "'"
          Call ConectionExecute(strSQL)
        
          strSQL = "select cod_seccion,descripcion from AfSecciones" _
                 & " where cod_institucion = " & vCodigo & " and cod_departamento = '" _
                 & cboDept.ItemData(cboDept.ListIndex) & "' order by cod_seccion"
          Call sbCargaGrid(vGridSec, 2, strSQL)
    End If

End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub cboDept_Click()
Dim strSQL As String

If vPaso Then Exit Sub


txtFiltro(1).Text = fxSysCleanTxtInject(txtFiltro(1).Text)

strSQL = "select cod_seccion,descripcion from AfSecciones" _
       & " where cod_institucion = " & vCodigo & " and cod_departamento = '" & cboDept.ItemData(cboDept.ListIndex) _
       & "' and Descripcion like '%" & txtFiltro(1).Text & "%'" _
       & " order by cod_seccion"
Call sbCargaGrid(vGridSec, 2, strSQL)

End Sub


