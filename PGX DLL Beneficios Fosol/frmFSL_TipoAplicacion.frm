VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_TipoAplicacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FOSOL: Tipo de aplicación"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Plan"
      TabPicture(0)   =   "frmFSL_TipoAplicacion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Causas"
      TabPicture(1)   =   "frmFSL_TipoAplicacion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboTipo"
      Tab(1).Control(1)=   "vGridCausas"
      Tab(1).Control(2)=   "Label6"
      Tab(1).ControlCount=   3
      Begin VB.ComboBox cboTipo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68880
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   3855
      End
      Begin FPSpreadADO.fpSpread vGridCausas 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   2
         Top             =   960
         Width           =   10095
         _Version        =   524288
         _ExtentX        =   17806
         _ExtentY        =   6376
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   5
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_TipoAplicacion.frx":0038
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4095
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   9735
         _Version        =   524288
         _ExtentX        =   17171
         _ExtentY        =   7223
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   4
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_TipoAplicacion.frx":0724
         VScrollSpecialType=   2
         AppearanceStyle =   1
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
         Left            =   -70440
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   10440
      X2              =   0
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Planes de Aplicación del Fondo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Image Image3 
      Height          =   720
      Left            =   0
      Picture         =   "frmFSL_TipoAplicacion.frx":0DAB
      Top             =   0
      Width           =   720
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
vModulo = 22
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
       & " where COD_PLAN = '" & SIFGlobal.fxSIFCodText(cboTipo.Text) _
       & "' order by COD_CAUSA"
Call sbCargaGrid(vGridCausas, 5, strSQL)

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 22
vGrid.AppearanceStyle = fxGridStyle

ssTab.Tab = 0
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
rs.Open strSQL, glogon.Conection, adOpenStatic


If rs!Existe = 0 Then

   
  strSQL = "insert " & pTabla & "(" & pCodigo & " ,Descripcion,Tipo_Desembolso, Activo,registro_fecha,registro_usuario) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "',"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ",getdate(),'" & glogon.Usuario & "')"
  
  glogon.Conection.Execute strSQL

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
  glogon.Conection.Execute strSQL

  vGrid.Col = 1
  Call Bitacora("Modifica", "Planes de Aplicación Id.:" & vGrid.Text)

End If

rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

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
       & " where " & pCodigo & " = '" & vGridCausas.Text & "' AND COD_PLAN = '" & SIFGlobal.fxSIFCodText(cboTipo.Text) & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic


If rs!Existe = 0 Then

   
  strSQL = "insert " & pTabla & "(" & pCodigo & ",cod_plan, Descripcion,Monto_Base,Tipo_Tabla,  Activa,registro_fecha,registro_usuario) values('" _
         & vGridCausas.Text & "','" & SIFGlobal.fxSIFCodText(cboTipo.Text) & "','"
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
  
  glogon.Conection.Execute strSQL

  vGridCausas.Col = 1
  
  Call Bitacora("Registra", "Planes de Apl: " & SIFGlobal.fxSIFCodText(cboTipo.Text) & "..Causa Id.:" & vGridCausas.Text)

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
  strSQL = strSQL & vGridCausas.Text & "' and COD_PLAN = '" & SIFGlobal.fxSIFCodText(cboTipo.Text) & "'"
  glogon.Conection.Execute strSQL

  vGridCausas.Col = 1
  Call Bitacora("Modifica", "Planes de Apl: " & SIFGlobal.fxSIFCodText(cboTipo.Text) & "..Causa Id.:" & vGridCausas.Text)

End If

rs.Close

fxGuardarCausa = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function




Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String

If ssTab.Tab = 1 Then

    vPaso = True
        strSQL = "select RTRIM(COD_PLAN) + ' - ' + DESCRIPCION as ItmX FROM FSL_PLANES WHERE ACTIVO = 1"
        Call sbLlenaCbo(cboTipo, strSQL, False, False)
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
        glogon.Conection.Execute strSQL

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
  MsgBox Err.Description, vbCritical

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
        strSQL = "delete FSL_PLANES_CAUSAS where COD_PLAN = '" & SIFGlobal.fxSIFCodText(cboTipo.Text) _
                & "' AND COD_CAUSA = '" & vGridCausas.Text & "'"
        glogon.Conection.Execute strSQL

        strSQL = vGridCausas.Text
        vGridCausas.Col = 1
        Call Bitacora("Elimina", "Planes Apl: " & SIFGlobal.fxSIFCodText(cboTipo.Text) & " .. Causa Id.:" & vGridCausas.Text)

        vGridCausas.DeleteRows vGridCausas.ActiveRow, 1
        vGridCausas.MaxRows = vGridCausas.MaxRows - 1
        vGridCausas.Row = vGridCausas.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub


