VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_TablasTipos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mantenimiento de Tipos (Enfermedades/Gestiones/Apelaciones)"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7785
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
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7515
      _Version        =   524288
      _ExtentX        =   13256
      _ExtentY        =   8705
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
      MaxCols         =   3
      MoveActiveOnFocus=   0   'False
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmFSL_PorcentajeIncapacidad.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Caption         =   "TIPOS DE .:"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   260
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmFSL_PorcentajeIncapacidad.frx":0612
      Top             =   120
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   6360
      X2              =   120
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "frmFSL_TablasTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mTipo As String


Private Sub cboTabla_Click()
Dim strSQL As String

If vPaso Then Exit Sub

mTipo = Mid(cboTabla.Text, 1, 1)

Select Case mTipo
  Case "G"
    strSQL = "select COD_GESTION,descripcion,Activa from FSL_TIPOS_GESTIONES order by COD_GESTION"
  Case "A"
    strSQL = "select COD_APELACION,descripcion,Activa from FSL_TIPOS_APELACIONES order by COD_APELACION"
  Case "E"
    strSQL = "select COD_ENFERMEDAD,descripcion,Activa from FSL_TIPOS_ENFERMEDADES order by COD_ENFERMEDAD"
End Select

Call sbCargaGrid(vGrid, 3, strSQL)

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
cboTabla.AddItem "Gestiones"
cboTabla.AddItem "Apelaciones"
cboTabla.AddItem "Enfermedades"
cboTabla.Text = "Gestiones"

vPaso = False
Call cboTabla_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim pExiste As Long, pCodigo As String, pTabla As String

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

Select Case mTipo
  Case "G"
        pCodigo = "COD_GESTION"
        pTabla = "FSL_TIPOS_GESTIONES"
  Case "A"
        pCodigo = "COD_APELACION"
        pTabla = "FSL_TIPOS_APELACIONES"
  Case "E"
        pCodigo = "COD_ENFERMEDAD"
        pTabla = "FSL_TIPOS_ENFERMEDADES"
End Select


fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If Trim(vGrid.Text) = "" Then Exit Function

strSQL = "select isnull(count(*),0) as Existe from " & pTabla _
       & " where " & pCodigo & " = '" & vGrid.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic


If rs!Existe = 0 Then

   
  strSQL = "insert " & pTabla & "(" & pCodigo & " ,Descripcion, Activa,registro_fecha,registro_usuario) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",getdate(),'" & glogon.Usuario & "')"
  
  glogon.Conection.Execute strSQL

  vGrid.Col = 1
  
  Call Bitacora("Registra", "Tipos de " & cboTabla.Text & " Id.:" & vGrid.Text)

Else 'Actualizar

  vGrid.Col = 2
  strSQL = "update " & pTabla & " set Descripcion = '" & vGrid.Text & "', Activa = "
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & " where " & pCodigo & " = '"
  vGrid.Col = 1
  strSQL = strSQL & vGrid.Text & "'"
  glogon.Conection.Execute strSQL

  vGrid.Col = 1
  Call Bitacora("Modifica", "Tipos de " & cboTabla.Text & " Id.:" & vGrid.Text)

End If

rs.Close

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
        Select Case mTipo
            Case "G"
                    strSQL = "delete FSL_TIPOS_GESTIONES where COD_GESTION = '" & vGrid.Text & "'"
            Case "A"
                    strSQL = "delete FSL_TIPOS_APELACIONES where COD_APELACION = '" & vGrid.Text & "'"
            Case "E"
                    strSQL = "delete FSL_TIPOS_ENFERMEDADES where COD_ENFERMEDAD = '" & vGrid.Text & "'"
        End Select
        glogon.Conection.Execute strSQL

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Tipos de " & cboTabla.Text & " Id.:" & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub
