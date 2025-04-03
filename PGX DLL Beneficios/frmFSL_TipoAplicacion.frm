VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmFSL_TipoAplicacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FOSOL: Tipo de aplicación"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5775
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   10455
      _Version        =   1441793
      _ExtentX        =   18441
      _ExtentY        =   10186
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
      Item(0).Caption =   "Plan"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Causas"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "cboTipo"
      Item(1).Control(1)=   "vGridCausas"
      Item(1).Control(2)=   "Label6"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5175
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   9735
         _Version        =   524288
         _ExtentX        =   17171
         _ExtentY        =   9128
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   4
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_TipoAplicacion.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridCausas 
         Height          =   4575
         Left            =   -69760
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   10095
         _Version        =   524288
         _ExtentX        =   17806
         _ExtentY        =   8070
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   5
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_TipoAplicacion.frx":06AF
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   -64240
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
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
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Plan"
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
         Left            =   -65920
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Planes de Aplicación del Fondo"
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
Attribute VB_Name = "frmFSL_TipoAplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub Form_Activote()
vModulo = 7
End Sub

Private Sub cboTipo_Click()
Dim strSQL As String

If vPaso Then Exit Sub

strSQL = "select COD_CAUSA,descripcion" _
       & ", case when MONTO_BASE = 'F' then 'Formalizado' else 'Saldo' end as 'MontoBase'" _
       & ", case when TIPO_TABLA = 'F' then 'Fallecimiento' when TIPO_TABLA = 'I' then 'Incapacidad' " _
       & "       when TIPO_TABLA = 'X' then '100 %' when TIPO_TABLA = 'S' then 'Suicidio' Else 'Fallecimiento' end as 'TipoTabla'" _
       & ",Activa" _
       & " from FSL_PLANES_CAUSAS" _
       & " where COD_PLAN = '" & cboTipo.ItemData(cboTipo.ListIndex) _
       & "' order by COD_CAUSA"
Call sbCargaGrid(vGridCausas, 5, strSQL)

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 7
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

tcMain.Item(0).Selected = True

strSQL = "select COD_PLAN,descripcion,case when isnull(Tipo_Desembolso,'F') = 'F' then 'Fondos' else 'Tesorería' end as 'TIPO' " _
       & " ,Activo" _
       & " from FSL_PLANES order by COD_PLAN"
Call sbCargaGrid(vGrid, 4, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim pExiste As Long, pCodigo As String, pTabla As String

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

pCodigo = "COD_PLAN"
pTabla = "FSL_PLANES"

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If Trim(vGrid.Text) = "" Then Exit Function

strSQL = "select isnull(count(*),0) as Existe from " & pTabla _
       & " where " & pCodigo & " = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)


If rs!Existe = 0 Then

   
  strSQL = "insert " & pTabla & "(" & pCodigo & " ,Descripcion,Tipo_Desembolso, Activo,registro_fecha,registro_usuario) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "',"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ",getdate(),'" & glogon.Usuario & "')"
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  
  Call Bitacora("Registra", "Planes de Aplicación Id.:" & vGrid.Text)

Else 'Actualizar

  vGrid.Col = 2
  strSQL = "update " & pTabla & " set Descripcion = '" & vGrid.Text & "', Tipo_Desembolso = '"
  vGrid.Col = 3
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "', Activo = "
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & " where " & pCodigo & " = '"
  vGrid.Col = 1
  strSQL = strSQL & vGrid.Text & "'"
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Modifica", "Planes de Aplicación Id.:" & vGrid.Text)

End If

rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Function fxGuardarCausa() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim pExiste As Long, pCodigo As String, pTabla As String, pTipo As String

On Error GoTo vError

pCodigo = "COD_CAUSA"
pTabla = "FSL_PLANES_CAUSAS"

fxGuardarCausa = 0
vGridCausas.Row = vGridCausas.ActiveRow
vGridCausas.Col = 1
 
If Trim(vGridCausas.Text) = "" Then Exit Function

strSQL = "select isnull(count(*),0) as Existe from " & pTabla _
       & " where " & pCodigo & " = '" & vGridCausas.Text & "' AND COD_PLAN = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)


If rs!Existe = 0 Then

   
  strSQL = "insert " & pTabla & "(" & pCodigo & ",cod_plan, Descripcion,Monto_Base,Tipo_Tabla,  Activa,registro_fecha,registro_usuario) values('" _
         & vGridCausas.Text & "','" & cboTipo.ItemData(cboTipo.ListIndex) & "','"
  vGridCausas.Col = 2
  strSQL = strSQL & vGridCausas.Text & "','"
  vGridCausas.Col = 3
  strSQL = strSQL & Mid(vGridCausas.Text, 1, 1) & "','"
  vGridCausas.Col = 4
  Select Case Mid(vGridCausas.Text, 1, 1)
     Case "F"
      pTipo = "F"
     Case "I"
      pTipo = "I"
     Case "S"
      pTipo = "S"
     Case "1"
      pTipo = "X"
  End Select
  strSQL = strSQL & pTipo & "',"
  vGridCausas.Col = 5
  strSQL = strSQL & vGridCausas.Value & ",getdate(),'" & glogon.Usuario & "')"
  
  Call ConectionExecute(strSQL)

  vGridCausas.Col = 1
  
  Call Bitacora("Registra", "Planes de Apl: " & cboTipo.ItemData(cboTipo.ListIndex) & "..Causa Id.:" & vGridCausas.Text)

Else 'Actualizar

  vGridCausas.Col = 2
  strSQL = "update " & pTabla & " set Descripcion = '" & vGridCausas.Text & "', Monto_Base = '"
  vGridCausas.Col = 3
  strSQL = strSQL & Mid(vGridCausas.Text, 1, 1) & "',Tipo_Tabla = '"
  vGridCausas.Col = 4
  Select Case Mid(vGridCausas.Text, 1, 1)
     Case "F"
      pTipo = "F"
     Case "I"
      pTipo = "I"
     Case "S"
      pTipo = "S"
     Case "1"
      pTipo = "X"
  End Select
  strSQL = strSQL & pTipo & "',Activa = "
  vGridCausas.Col = 5
  strSQL = strSQL & vGridCausas.Value & " where " & pCodigo & " = '"
  vGridCausas.Col = 1
  strSQL = strSQL & vGridCausas.Text & "' and COD_PLAN = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
  Call ConectionExecute(strSQL)

  vGridCausas.Col = 1
  Call Bitacora("Modifica", "Planes de Apl: " & cboTipo.ItemData(cboTipo.ListIndex) & "..Causa Id.:" & vGridCausas.Text)

End If

rs.Close

fxGuardarCausa = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

If Item.Index = 1 Then

    vPaso = True
        strSQL = "select RTRIM(COD_PLAN) as 'IdX', rtrim(DESCRIPCION) as ItmX FROM FSL_PLANES WHERE ACTIVO = 1"
        Call sbCbo_Llena_New(cboTipo, strSQL, False, True)
    vPaso = False
    Call cboTipo_Click
    
End If

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
        strSQL = "delete FSL_PLANES where COD_PLAN = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Planes de Aplicación Id.:" & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Private Sub vGridCausas_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGridCausas.ActiveCol = vGridCausas.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarCausa
  If i = 0 Then Exit Sub
  vGridCausas.Row = vGridCausas.ActiveRow
  If vGridCausas.MaxRows <= vGridCausas.ActiveRow Then
    vGridCausas.MaxRows = vGridCausas.MaxRows + 1
    vGridCausas.Row = vGridCausas.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridCausas.MaxRows = vGridCausas.MaxRows + 1
    vGridCausas.InsertRows vGridCausas.ActiveRow, 1
    vGridCausas.Row = vGridCausas.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGridCausas.Row = vGridCausas.ActiveRow
        vGridCausas.Col = 1
        strSQL = "delete FSL_PLANES_CAUSAS where COD_PLAN = '" & cboTipo.ItemData(cboTipo.ListIndex) _
                & "' AND COD_CAUSA = '" & vGridCausas.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGridCausas.Text
        vGridCausas.Col = 1
        Call Bitacora("Elimina", "Planes Apl: " & cboTipo.ItemData(cboTipo.ListIndex) & " .. Causa Id.:" & vGridCausas.Text)

        vGridCausas.DeleteRows vGridCausas.ActiveRow, 1
        vGridCausas.MaxRows = vGridCausas.MaxRows - 1
        vGridCausas.Row = vGridCausas.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


