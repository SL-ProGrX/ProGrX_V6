VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmAF_PlanMutual 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Plan Mutual y de Beneficios a Asociados"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   13935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6735
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   13695
      _Version        =   1441793
      _ExtentX        =   24156
      _ExtentY        =   11880
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
      Item(0).Caption =   "Personas"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "scMain"
      Item(0).Control(2)=   "btnExportar"
      Item(1).Caption =   "Planes"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5775
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   13455
         _Version        =   1441793
         _ExtentX        =   23733
         _ExtentY        =   10186
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
         ShowBorder      =   0   'False
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5655
         Left            =   -69400
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   12375
         _Version        =   524288
         _ExtentX        =   21828
         _ExtentY        =   9975
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
         SpreadDesigner  =   "frmAF_PlanMutual.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   375
         Left            =   13080
         TabIndex        =   17
         ToolTipText     =   "Exportar a Excel"
         Top             =   360
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_PlanMutual.frx":0764
      End
      Begin XtremeShortcutBar.ShortcutCaption scMain 
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   360
         Width           =   13695
         _Version        =   1441793
         _ExtentX        =   24156
         _ExtentY        =   661
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
         Alignment       =   1
      End
   End
   Begin VB.Timer Timerx 
      Interval        =   10
      Left            =   13200
      Top             =   1320
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   11280
      TabIndex        =   1
      Top             =   1320
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_PlanMutual.frx":1035
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtIdAlterna 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   4200
      TabIndex        =   4
      Top             =   1320
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   12240
      TabIndex        =   5
      Top             =   8760
      Visible         =   0   'False
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   330
      Left            =   7320
      TabIndex        =   6
      Top             =   360
      Width           =   3735
      _Version        =   1441793
      _ExtentX        =   6588
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
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   315
      Left            =   10080
      TabIndex        =   19
      Top             =   8640
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   556
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
      Text            =   "100"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnActualizar 
      Height          =   375
      Left            =   11280
      TabIndex        =   20
      Top             =   360
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Actualizar Retenciones"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_PlanMutual.frx":1735
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   9000
      TabIndex        =   18
      Top             =   8640
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Qty Lineas:"
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
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
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Id Alterna"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
   Begin XtremeSuiteControls.Label lblItems 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   8640
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado de la Persona"
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
      Height          =   255
      Index           =   5
      Left            =   7320
      TabIndex        =   7
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "frmAF_PlanMutual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub btnActualizar_Click()

If vPaso Then Exit Sub
If cbo.ListCount = 0 Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spAFI_PM_Recaudos_Update '" & cbo.ItemData(cbo.ListIndex) & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


Call Bitacora("Aplica", "Plan Mutual/Beneficios: " & vGrid.Text & ", Actualización de Recaudos")

Me.MousePointer = vbDefault

MsgBox "Actualización de Operaciones de Recaudo actualizadas!", vbInformation

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBuscar_Click()
Call sbBuscar
End Sub


Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbPlanes_Load()

On Error GoTo vError

strSQL = "select COD_PLAN, DESCRIPCION, MONTO, CODIGIO_RETENCION, ACTIVO, REGISTRO_FECHA, REGISTRO_USUARIO" _
       & " from AFI_PLAN_MUTUAL" _
       & " order by COD_PLAN"
vPaso = True
    Call sbCargaGrid(vGrid, 7, strSQL)
vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cbo_Click()
If vPaso Then Exit Sub
If cbo.ListCount = 0 Then Exit Sub
If cboEstado.ListCount = 0 Then Exit Sub

Call sbBuscar

End Sub

Private Sub cboEstado_Click()
If vPaso Then Exit Sub
If cbo.ListCount = 0 Then Exit Sub
If cboEstado.ListCount = 0 Then Exit Sub

Call sbBuscar

End Sub

Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

scMain.Caption = "(Casos con Check estan excluídos del proceso)"

With lsw.ColumnHeaders
    .Add , , "Identificación", 2000
    .Add , , "Id. Alterna", 2000
    .Add , , "Nombre", 3500
    .Add , , "Excluye?", 1000, vbCenter
    .Add , , "Reg.Usuario", 2000, vbCenter
    .Add , , "Reg.Fecha", 2000, vbCenter
End With


Call Formularios(Me)
Call RefrescaTags(Me)
        
End Sub



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub
If cbo.ListCount = 0 Then Exit Sub
If Not vGrid.Enabled Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spAFI_PM_Excluye '" & cbo.ItemData(cbo.ListIndex) & "', '" & Item.Text _
        & "', " & IIf((Item.Checked = True), 1, 0) & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Consulta
        Call sbBuscar
    
    Case 1 'Planes
        Call sbPlanes_Load
End Select

End Sub

Private Sub TimerX_Timer()

Timerx.Interval = 0
Timerx.Enabled = False


cbo.Clear

vPaso = True

strSQL = "select COD_PLAN as 'Idx',  RTRIM(DESCRIPCION) as 'ItmX'" _
       & "  from AFI_PLAN_MUTUAL where ACTIVO = 1" _
       & "  order by COD_PLAN"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

cboEstado.Clear
cboEstado.AddItem "Todos"
cboEstado.ItemData(cboEstado.ListCount - 1) = "T"

cboEstado.AddItem "Casos Excluídos"
cboEstado.ItemData(cboEstado.ListCount - 1) = "E"

cboEstado.AddItem "Casos Activos"
cboEstado.ItemData(cboEstado.ListCount - 1) = "A"

cboEstado.Text = "Todos"

vPaso = False

Call sbBuscar

End Sub



Private Sub sbBuscar()

On Error GoTo vError


Me.MousePointer = vbHourglass


tcMain.Item(0).Selected = True

lsw.ListItems.Clear

txtCedula.Text = fxSysCleanTxtInject(txtCedula.Text)
txtIdAlterna.Text = fxSysCleanTxtInject(txtIdAlterna.Text)
txtNombre.Text = fxSysCleanTxtInject(txtNombre.Text)

If Not IsNumeric(txtLineas.Text) Then
    txtLineas.Text = "100"
End If

strSQL = "exec spAFI_PM_Consulta '" & cbo.ItemData(cbo.ListIndex) & "', '" & txtCedula.Text _
            & "', '" & txtIdAlterna.Text & "', '" & txtNombre.Text & "', '" & cboEstado.ItemData(cboEstado.ListIndex) _
            & "', " & txtLineas.Text

vPaso = True

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!CEDULA)
      itmX.SubItems(1) = rs!ID_ALTERNA & ""
      itmX.SubItems(2) = rs!Nombre & ""
      itmX.SubItems(3) = IIf(rs!Excluye = 1, "Sí", "")
      itmX.SubItems(4) = rs!REGISTRO_USUARIO & ""
      itmX.SubItems(5) = Format(rs!REGISTRO_FECHA & "", "yyyy-MM-dd")
      
      If rs!Excluye = 1 Then
        itmX.Checked = True
      End If
      
  rs.MoveNext
Loop
rs.Close
    
vPaso = False
    
lblItems.Caption = "Total de Líneas: " & lsw.ListItems.Count
    

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    lsw.ListItems.Clear

End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn And Len(txtCedula.Text) > 0 Then
    Call sbBuscar
End If
End Sub


Private Sub txtIdAlterna_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn And Len(txtIdAlterna.Text) > 0 Then
    Call sbBuscar
End If

End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn And Len(txtNombre.Text) > 0 Then
    Call sbBuscar
End If
End Sub



Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0

Dim pCodigo As String, pDescripcion As String, pMonto As Currency, pActivo As Integer, pRetencion As String


vGrid.Row = vGrid.ActiveRow

vGrid.Col = 1
pCodigo = vGrid.Text
vGrid.Col = 2
pDescripcion = vGrid.Text
vGrid.Col = 3
pMonto = CCur(vGrid.Text)
vGrid.Col = 4
pRetencion = vGrid.Text
vGrid.Col = 5
pActivo = vGrid.Value


strSQL = "exec spAFI_PM_Registro '" & pCodigo & "', '" & pDescripcion & "', " & pMonto _
        & " , '" & pRetencion & "', " & pActivo & ", '" & glogon.Usuario & "', 'A'"
Call ConectionExecute(strSQL)

vGrid.Col = 6
vGrid.Text = Date

vGrid.Col = 7
vGrid.Text = glogon.Usuario

Call Bitacora("Registra", "Plan Mutual/Beneficios: " & pCodigo)


fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If (vGrid.ActiveCol = vGrid.MaxCols Or vGrid.ActiveCol = 5) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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

        strSQL = "exec spAFI_PM_Registro '" & vGrid.Text & "', '', 0, '', 0, '" & glogon.Usuario & "', 'E'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Plan Mutual/Beneficios: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


