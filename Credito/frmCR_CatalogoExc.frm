VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCR_CatalogoExc 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabla de Cálculo de Diponible Creditos Sobre Excedentes"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton cmdModifica 
      Height          =   615
      Left            =   11400
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3413
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Actualiza Tabla"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   14175
      _Version        =   524288
      _ExtentX        =   25003
      _ExtentY        =   6800
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
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_CatalogoExc.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tabla de Porcentajes Disponibles para Crédito con Garantía en Excedentes"
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
      Height          =   720
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   300
      Width           =   6615
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   14532
   End
End
Attribute VB_Name = "frmCR_CatalogoExc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub Form_Load()

vModulo = 3

Set imgBanner = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle

strSQL = "select MES,ACUMULADO_MES,ACUMULADO_PORC,CAPGEN, REGISTRO_FECHA, REGISTRO_USUARIO, MODIFICA_FECHA, MODIFICA_USUARIO" _
       & " from EXC_DISPONIBLE order by mes"
Call sbCargaGrid(vGrid, 8, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

If Not cmdModifica.Enabled Then vGrid.Enabled = False

End Sub


Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If Trim(vGrid.Text) = "" Then Exit Function

        strSQL = "select MES,ACUMULADO_MES,ACUMULADO_PORC,CAPGEN, REGISTRO_FECHA, REGISTRO_USUARIO, MODIFICA_FECHA, MODIFICA_USUARIO" _
               & " from EXC_DISPONIBLE order by mes"

strSQL = "select isnull(count(*),0) as Existe from EXC_DISPONIBLE " _
       & " where MES = " & vGrid.Text
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into EXC_DISPONIBLE(MES, ACUMULADO_MES, ACUMULADO_PORC, CAPGEN, REGISTRO_FECHA, REGISTRO_USUARIO)" _
         & " values( " & vGrid.Text & ", "
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & ", "
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & ", "
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Text & ", dbo.MyGetDate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Disponible Excedentes Mes: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update EXC_DISPONIBLE set ACUMULADO_MES = " & vGrid.Text & ", ACUMULADO_PORC = "
 vGrid.Col = 3
 strSQL = strSQL & CCur(vGrid.Text) & ", CAPGEN = "
 vGrid.Col = 4
 strSQL = strSQL & CCur(vGrid.Text) & ", Modifica_Fecha = dbo.mygetdate(), Modifica_Usuario = '" & glogon.Usuario & "'"
 vGrid.Col = 1
 strSQL = strSQL & " Where Mes = " & vGrid.Text
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Disponible Excedentes Mes: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = 4 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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
      If vGrid.Text <> "" Then
        strSQL = "delete EXC_DISPONIBLE where MES = " & vGrid.Text
        Call ConectionExecute(strSQL)
      End If
    
      strSQL = vGrid.Text
      vGrid.Col = 1
      Call Bitacora("Elimina", "Disponible Excedentes Mes: " & vGrid.Text)
      
        strSQL = "select MES,ACUMULADO_MES,ACUMULADO_PORC,CAPGEN, REGISTRO_FECHA, REGISTRO_USUARIO, MODIFICA_FECHA, MODIFICA_USUARIO" _
               & " from EXC_DISPONIBLE order by mes"
        Call sbCargaGrid(vGrid, 8, strSQL)
    End If
 
End If

End Sub

