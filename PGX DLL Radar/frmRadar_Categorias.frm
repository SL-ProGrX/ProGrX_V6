VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmRadar_Categorias 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Radar: Mantenimientos de Categorías"
   ClientHeight    =   6576
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11436
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6576
   ScaleWidth      =   11436
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   960
      Top             =   5520
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   3612
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1812
      _Version        =   1245185
      _ExtentX        =   3196
      _ExtentY        =   6371
      _StockProps     =   79
      Caption         =   "Categorías (Tipos): "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnOpciones 
         Height          =   420
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Contactos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmRadar_Categorias.frx":0000
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnOpciones 
         Height          =   420
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Colocadores"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmRadar_Categorias.frx":07CD
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnOpciones 
         Height          =   420
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Centros"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmRadar_Categorias.frx":1132
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnOpciones 
         Height          =   420
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Resultados"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmRadar_Categorias.frx":18EE
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnOpciones 
         Height          =   420
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Eventos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmRadar_Categorias.frx":1FD6
         ImageAlignment  =   0
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5052
      Left            =   2160
      TabIndex        =   7
      Top             =   1440
      Width           =   9132
      _Version        =   524288
      _ExtentX        =   16108
      _ExtentY        =   8911
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
      SpreadDesigner  =   "frmRadar_Categorias.frx":2963
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label lblTipo 
      BackStyle       =   0  'Transparent
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   2124
      TabIndex        =   6
      Top             =   360
      Width           =   7332
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   12372
   End
End
Attribute VB_Name = "frmRadar_Categorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mTabla As String, mCodigo As String, mReferencia As String

Private Sub btnOpciones_Click(Index As Integer)

Select Case Index
  Case 0 'Tipos de Contactos
    mReferencia = "Tipos de Contactos"
    mTabla = "RADAR_CONTACTOS_TIPOS"
    mCodigo = "CONTACTO_TIPO"
  
  Case 1 'Tipos de Colocadores
    mReferencia = "Tipos de Colocadores"
    mTabla = "RADAR_COLOCADORES_TIPOS"
    mCodigo = "COLOCADOR_TIPO"
  
  Case 2 'Tipos de Centros
    mReferencia = "Tipos de Centros de Trabajo"
    mTabla = "RADAR_CENTROS_TIPOS"
    mCodigo = "CENTRO_TIPO"
  
  Case 3 'Resultados
    mReferencia = "Tipos de Resultados"
    mTabla = "RADAR_RESULTADOS_TIPOS"
    mCodigo = "RESULTADO_TIPO"
  
  Case 4 'Tipos de Eventos
    mReferencia = "Tipos de Eventos"
    mTabla = "RADAR_EVENTOS_TIPOS"
    mCodigo = "EVENTO_TIPO"
  
   Case Else
    mReferencia = "Tipos de Contactos"
    mTabla = "RADAR_CONTACTOS_TIPOS"
    mCodigo = "CONTACTO_TIPO"
   
End Select

lblTipo.Caption = mReferencia

Call sbConsulta

End Sub

Private Sub Form_ACTIVOte()
vModulo = 37
End Sub

Private Sub sbConsulta()
Dim strSQL As String
      
strSQL = "select " & mCodigo & ",descripcion,ACTIVO,Registro_Fecha,Registro_Usuario" _
       & " from " & mTabla _
       & " order by " & mCodigo
vPaso = True
    Call sbCargaGrid(vGrid, 5, strSQL)
vPaso = False
End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 1
vGrid.AppearanceStyle = fxGridStyle

mTabla = ""


Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from " & mTabla _
       & " where " & mCodigo & " = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function

  strSQL = "insert into " & mTabla & "(" & mCodigo & ",descripcion,ACTIVO,registro_fecha,registro_usuario) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",Getdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 4
  vGrid.Text = fxFechaServidor
  vGrid.Col = 5
  vGrid.Text = glogon.Usuario
  
  Call Bitacora("Registra", mReferencia & ": " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update " & mTabla & " set Descripcion = '" & vGrid.Text & "',ACTIVO= "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & " where " & mCodigo & " = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", mReferencia & ": " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False
Call btnOpciones_Click(0)


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If (vGrid.ActiveCol = vGrid.MaxCols Or vGrid.ActiveCol = 3) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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
        strSQL = "delete " & mTabla & " where " & mCodigo & " = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", mReferencia & ": " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
