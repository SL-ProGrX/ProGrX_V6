VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_TablaCoberturas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabla de aplicación de coberturas"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboTabla 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4575
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   5955
      _Version        =   524288
      _ExtentX        =   10504
      _ExtentY        =   8070
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
      SpreadDesigner  =   "frmFSL_TablaCoberturas.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label lblDescripcion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmFSL_TablaCoberturas.frx":06B1
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   6135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   6360
      X2              =   120
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmFSL_TablaCoberturas.frx":075B
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmFSL_TablaCoberturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mTipo As String


Private Sub cboTabla_Click()
Dim strSQL As String

If vPaso Then Exit Sub

strSQL = "select Linea,Mes_Inicio,Mes_Corte,Cobertura from FSL_TABLAS_APLICACION"
Select Case cboTabla.Text
  Case "Tabla de Fallecimiento"
    mTipo = "F"
  Case "Tabla de Incapacidad"
    mTipo = "I"
  Case "Tabla de Suicidios"
    mTipo = "S"
  Case "Tabla de 100%"
    mTipo = "X"
End Select

strSQL = strSQL & " Where Tipo = '" & mTipo & "' order by Mes_Inicio"
Call sbCargaGrid(vGrid, 4, strSQL)

End Sub

Private Sub Form_Activate()
vModulo = 22
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 22
vGrid.AppearanceStyle = fxGridStyle

vPaso = True

cboTabla.Clear
cboTabla.AddItem "Tabla de Fallecimiento"
cboTabla.AddItem "Tabla de Incapacidad"
cboTabla.AddItem "Tabla de Suicidios"
cboTabla.AddItem "Tabla de 100%"
cboTabla.Text = "Tabla de Fallecimiento"

vPaso = False
Call cboTabla_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim pCodigo As Long

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1


If Trim(vGrid.Text) = "" Then

   strSQL = "select isnull(Max(Linea),0) + 1 as Ultimo from FSL_TABLAS_APLICACION " _
          & " where Tipo = '" & mTipo & "'"
   rs.Open strSQL, glogon.Conection, adOpenStatic
       pCodigo = rs!ultimo
   rs.Close
   
  strSQL = "insert into FSL_TABLAS_APLICACION(Tipo,Linea,Mes_Inicio,Mes_Corte,Cobertura,registra_fecha,registra_usuario) values('" _
         & mTipo & "'," & pCodigo & ","
  vGrid.Col = 2
  strSQL = strSQL & CLng(vGrid.Text) & ","
  vGrid.Col = 3
  strSQL = strSQL & CLng(vGrid.Text) & ","
  vGrid.Col = 4
  strSQL = strSQL & CCur(vGrid.Text) & ",getdate(),'" & glogon.Usuario & "')"
  
  glogon.Conection.Execute strSQL

  vGrid.Col = 1
  vGrid.Text = CStr(pCodigo)
  
  Call Bitacora("Registra", cboTabla.Text & " Id.:" & vGrid.Text)

Else 'Actualizar

  vGrid.Col = 2
  strSQL = "update FSL_TABLAS_APLICACION set Mes_Inicio = " & CLng(vGrid.Text) & ",Mes_Corte = "
  vGrid.Col = 3
  strSQL = strSQL & CLng(vGrid.Text) & ", Cobertura = "
  vGrid.Col = 4
  strSQL = strSQL & CCur(vGrid.Text) & " where Tipo = '" & mTipo & "' and Linea = "
  vGrid.Col = 1
  strSQL = strSQL & vGrid.Text
  glogon.Conection.Execute strSQL

  vGrid.Col = 1
  Call Bitacora("Modifica", cboTabla.Text & " Id.:" & vGrid.Text)

End If

fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function




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
        strSQL = "delete FSL_TABLAS_APLICACION where Tipo = '" & mTipo & "' and Linea = " & vGrid.Text
        glogon.Conection.Execute strSQL

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", cboTabla.Text & " Id.:" & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub


