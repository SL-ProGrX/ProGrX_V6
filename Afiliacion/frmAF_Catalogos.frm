VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmAF_Catalogos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Catálogos para Clientes"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   9600
      Top             =   720
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   11055
      _Version        =   1441793
      _ExtentX        =   19500
      _ExtentY        =   11668
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
      Item(0).Caption =   "Catálogos"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "Label2"
      Item(0).Control(2)=   "cbo"
      Item(1).Caption =   "Tipos"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGridt"
      Begin FPSpreadADO.fpSpread vGridt 
         Height          =   6135
         Left            =   -69280
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   9735
         _Version        =   524288
         _ExtentX        =   17171
         _ExtentY        =   10821
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
         MaxCols         =   490
         ScrollBars      =   2
         SpreadDesigner  =   "frmAF_Catalogos.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5535
         Left            =   0
         TabIndex        =   3
         Top             =   1080
         Width           =   11055
         _Version        =   524288
         _ExtentX        =   19500
         _ExtentY        =   9763
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
         MaxCols         =   491
         ScrollBars      =   2
         SpreadDesigner  =   "frmAF_Catalogos.frx":06BA
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   330
         Left            =   960
         TabIndex        =   5
         Top             =   480
         Width           =   4575
         _Version        =   1441793
         _ExtentX        =   8070
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Tipo"
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
      Caption         =   "Catálogos Generales para Clientes"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   7812
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "frmAF_Catalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Private Sub cbo_Click()
If vPaso Then Exit Sub

strSQL = "select Linea_Id, Catalogo_Id, Descripcion, Activo, Registro_Fecha, REgistro_Usuario" _
       & " from AFI_CATALOGOS Where Tipo_Id = " & cbo.ItemData(cbo.ListIndex) _
       & " Order by Catalogo_Id"
Call sbCargaGrid(vGrid, 6, strSQL)

End Sub

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub sbInicial()


tcMain.Item(0).Selected = True
      
vPaso = True

      
strSQL = "select Tipo_Id as 'IdX', Descripcion as 'ItmX'" _
       & " from AFI_CATALOGOS_TIPOS Where Activo = 1" _
       & " order by descripcion"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

vPaso = False

Call cbo_Click


End Sub


Private Sub Form_Load()

vModulo = 1

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 2
If Trim(vGrid.Text) = "" Then
    MsgBox "Código No es Válido!", vbExclamation
    Exit Function
End If

strSQL = "select Count(*) as 'Existe', max(Linea_Id) as 'Linea_Id' from AFI_CATALOGOS " _
       & " where Catalogo_Id = '" & vGrid.Text & "' and Tipo_Id = " & cbo.ItemData(cbo.ListIndex)
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into AFI_CATALOGOS(Catalogo_Id, descripcion, Activo, Tipo_Id, Registro_fecha, Registro_usuario) values('" _
         & Trim(vGrid.Text) & "', '"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "', "
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ", " & cbo.ItemData(cbo.ListIndex) & ", Getdate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  
    strSQL = "select Linea_Id, registro_Fecha, Registro_Usuario from AFI_CATALOGOS " _
           & " where Catalogo_Id = '" & vGrid.Text & "' and Tipo_Id = " & cbo.ItemData(cbo.ListIndex)
    Call OpenRecordSet(rs, strSQL)
  
  
  vGrid.Col = 5
  vGrid.Text = rs!Registro_Fecha & ""
  vGrid.Col = 6
  vGrid.Text = rs!Registro_Usuario & ""
  
  
  vGrid.Col = 1
  vGrid.Text = CStr(rs!Linea_Id)
  
  Call Bitacora("Registra", "Catalogo Cliente Id:  " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 3
 strSQL = "update AFI_CATALOGOS set Descripcion = '" & vGrid.Text & "', Activo = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & ", Modifica_Fecha = getdate(), Modifica_Usuario = '" _
        & glogon.Usuario & "' where Linea_Id = "
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Catalogo Cliente Id:  " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 1 Then
    strSQL = "select Tipo_Id, Descripcion, Activo, Registro_Fecha, REgistro_Usuario" _
           & " from AFI_CATALOGOS_TIPOS Order by Descripcion"
    Call sbCargaGrid(vGridt, 5, strSQL)
End If

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False
Call sbInicial
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If (vGrid.ActiveCol = vGrid.MaxCols Or vGrid.ActiveCol = 4) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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
        strSQL = "delete AFI_CATALOGOS where Linea_Id = " & vGrid.Text
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Catalogo Cliente Id:  " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




