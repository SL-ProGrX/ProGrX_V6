VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmMarcasDetalle 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle de Marcas"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox chkTodasFec 
      Height          =   252
      Left            =   4560
      TabIndex        =   14
      Top             =   120
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todas ?"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton cmdReporte 
      Height          =   372
      Left            =   10440
      TabIndex        =   5
      Top             =   960
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Informe"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   7335
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   13575
      _Version        =   524288
      _ExtentX        =   23945
      _ExtentY        =   12938
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
      MaxCols         =   7
      ScrollBars      =   2
      SpreadDesigner  =   "frmMarcasDetalle.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   372
      Left            =   9120
      TabIndex        =   6
      Top             =   960
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Consultar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton cmdExcel 
      Height          =   372
      Left            =   11760
      TabIndex        =   7
      Top             =   960
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Exportar Excel"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.ComboBox cboHorario 
      Height          =   312
      Left            =   6960
      TabIndex        =   9
      Top             =   480
      Width           =   3252
      _Version        =   1441793
      _ExtentX        =   5741
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
   Begin XtremeSuiteControls.ComboBox cboMovimiento 
      Height          =   312
      Left            =   1800
      TabIndex        =   10
      Top             =   480
      Width           =   2652
      _Version        =   1441793
      _ExtentX        =   4683
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
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   6960
      TabIndex        =   11
      Top             =   120
      Width           =   3252
      _Version        =   1441793
      _ExtentX        =   5736
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
      Text            =   "(Presione F4)"
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1800
      TabIndex        =   12
      Top             =   120
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   550
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   3120
      TabIndex        =   13
      Top             =   120
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   550
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.CheckBox chkTodosMov 
      Height          =   252
      Left            =   4560
      TabIndex        =   15
      Top             =   480
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos ?"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.CheckBox chkTodosUsu 
      Height          =   252
      Left            =   10440
      TabIndex        =   16
      Top             =   120
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos ?"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.CheckBox chkTodosHorario 
      Height          =   252
      Left            =   10440
      TabIndex        =   8
      Top             =   480
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos ?"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   1800
      TabIndex        =   17
      Top             =   960
      Width           =   5175
      _Version        =   1441793
      _ExtentX        =   9128
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
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   18
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Horario"
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
      Height          =   312
      Index           =   0
      Left            =   6180
      TabIndex        =   3
      Top             =   480
      Width           =   852
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
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
      Height          =   312
      Index           =   1
      Left            =   6180
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Marca"
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
      Height          =   315
      Left            =   780
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas"
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
      Height          =   315
      Index           =   0
      Left            =   780
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgBanner 
      Height          =   876
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13920
   End
End
Attribute VB_Name = "frmMarcasDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vTamanoForm As Double
Dim aColumnas(6) As Double
Dim vHanchoGrid As Double, vAltoGrid As Double

Private Sub chkTodasFec_Click()

If chkTodasFec.Value = vbChecked Then
  dtpInicio.Enabled = False
  dtpCorte.Enabled = False
Else
  dtpInicio.Enabled = True
  dtpCorte.Enabled = True
End If

End Sub

Private Sub chkTodosHorario_Click()
If chkTodosHorario.Value = vbChecked Then
  cboHorario.Enabled = False
Else
  cboHorario.Enabled = True
End If
End Sub

Private Sub chkTodosMov_Click()
If chkTodosMov.Value = vbChecked Then
  cboMovimiento.Enabled = False
Else
  cboMovimiento.Enabled = True
End If

End Sub


Private Sub chkTodosUsu_Click()
 If chkTodosUsu.Value = vbChecked Then
   txtUsuario.Enabled = False
 Else
   txtUsuario.Enabled = True
   txtUsuario = "(Presione F4)"
 End If
 
End Sub

Private Sub cmdBuscar_Click()
Dim rs As New ADODB.Recordset

vGrid.MaxRows = 0

Call OpenRecordSet(rs, fxSQL)
Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  vGrid.Col = 1
  vGrid.Text = rs!Usuario
  vGrid.Col = 2
  vGrid.Text = Format(rs!fecha, "dd/mm/yyyy")
  vGrid.Col = 3
  vGrid.Text = Format(rs!fecha, "hh:mm:ss AMPM")
  vGrid.Col = 4
  vGrid.Text = rs!TIPO_MARCA
  vGrid.Col = 5
  vGrid.Text = rs!HorarioDesc
  vGrid.Col = 6
  vGrid.Text = rs!estacion
  vGrid.Col = 7
  vGrid.Text = rs!UserName
  
  rs.MoveNext
Loop
rs.Close
End Sub

Private Sub cmdExcel_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 8
    vHeaders.Headers(1) = "Usuario"
    vHeaders.Headers(2) = "Fecha"
    vHeaders.Headers(3) = "Hora"
    vHeaders.Headers(4) = "Monto"
    vHeaders.Headers(5) = "Tipo de Marca"
    vHeaders.Headers(6) = "Horario"
    vHeaders.Headers(7) = "Estación"
    vHeaders.Headers(8) = "Nombre"

Call sbSIFGridExportar(vGrid, vHeaders, "Marcas_Detalle")
End Sub

Private Sub cmdReporte_Click()
Me.MousePointer = vbHourglass

vGrid.PrintFooter = "fecha desde  " & Format(dtpInicio.Value, "dd/mm/yyyy") _
                  & " hasta " & Format(dtpCorte.Value, "dd/mm/yyyy") & "Usuario : " & glogon.Usuario

vGrid.PrintHeader = Me.Caption

If vGrid.MaxCols > 5 Then
    vGrid.PrintOrientation = PrintOrientationLandscape
Else
    vGrid.PrintOrientation = PrintOrientationPortrait
End If
vGrid.PrintSheet
  
Me.MousePointer = vbDefault
  
End Sub

Private Sub Form_Activate()
vModulo = 21
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vModulo = 21

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture


dtpInicio = fxFechaServidor
dtpCorte = dtpInicio

With cboMovimiento
   .AddItem "Entrada"
   .ItemData(.ListCount - 1) = CStr(1)
   .AddItem "Salida Almuerzo"
   .ItemData(.ListCount - 1) = CStr(2)
   .AddItem "Entrada Almuerzo"
   .ItemData(.ListCount - 1) = CStr(3)
   .AddItem "Salida"
   .ItemData(.ListCount - 1) = CStr(4)
   .Text = "Entrada"
End With

strSQL = "select rtrim(COD_HORARIO) as 'IdX', rtrim(DESCRIPCION) AS 'ItmX' from marcas_horarios where estado = 1"
Call sbCbo_Llena_New(cboHorario, strSQL, False, True)

chkTodosMov.Value = vbChecked
cboMovimiento.Enabled = False

chkTodosUsu.Value = vbChecked
txtUsuario.Enabled = False

chkTodosHorario.Value = vbChecked
cboHorario.Enabled = False

Call Formularios(Me)
Call RefrescaTags(Me)

vError:

End Sub




Private Function fxMovimientos(strMovimiento) As String
     
Select Case strMovimiento
  Case "1"
    fxMovimientos = "Entrada"
  Case "2"
    fxMovimientos = "Salida Almuerzo"
  Case "3"
    fxMovimientos = "Entrada Almuerzo"
  Case "4"
    fxMovimientos = "Salida"
  End Select

End Function

Private Function fxSQL() As String
Dim vPaso As Boolean, strSQL As String

vPaso = False
     
strSQL = "select Mr.USUARIO, Mr.FECHA, Mr.ESTACION, Mr.COD_HORARIO, Mh.DESCRIPCION as 'HorarioDesc', Us.DESCRIPCION as 'UserName'" _
        & "         ,case when Mr.TIPO_MARCA = 1 then 'Entrada'" _
        & "               when Mr.TIPO_MARCA = 2 then 'Salida Almuerzo'" _
        & "               when Mr.TIPO_MARCA = 3 then 'Entrada Almuerzo'" _
        & "               when Mr.TIPO_MARCA = 4 then 'Salida'  end as 'TIPO_MARCA'" _
        & "    from marcas_registro Mr inner join MARCAS_HORARIOS Mh on Mr.COD_HORARIO = Mh.COD_HORARIO" _
        & "      inner join USUARIOS Us on Mr.USUARIO = Us.NOMBRE"
     
If chkTodasFec.Value = vbUnchecked Then
  If vPaso Then
    strSQL = strSQL & " and Mr.fecha between '" & Format(dtpInicio.Value, "yyyymmdd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyymmdd") & " 23:59:00'"
  Else
    vPaso = True
    strSQL = strSQL & " where Mr.fecha between '" & Format(dtpInicio.Value, "yyyymmdd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyymmdd") & " 23:59:00'"
  End If
End If 'Clausulas de Fechas

If chkTodosMov.Value = vbUnchecked Then
  If vPaso Then
    strSQL = strSQL & " and Mr.tipo_marca = " & cboMovimiento.ItemData(cboMovimiento.ListIndex) & ""
  Else
    vPaso = True
    strSQL = strSQL & " where Mr.tipo_marca = " & cboMovimiento.ItemData(cboMovimiento.ListIndex) & ""
  End If
End If

'
If chkTodosUsu.Value = vbUnchecked Then
  If txtUsuario <> "" And txtUsuario <> "(Presione F4)" Then
   If vPaso Then
     strSQL = strSQL & " and Usuario = '" & txtUsuario.Text & "'"
   Else
     vPaso = True
     strSQL = strSQL & " where Mr.usuario = '" & txtUsuario.Text & "'"
   End If
  End If
End If

If chkTodosHorario.Value = vbUnchecked Then
 If vPaso Then
     strSQL = strSQL & " and Mr.cod_horario = '" & cboHorario.ItemData(cboHorario.ListIndex) & "'"
 Else
     strSQL = strSQL & " Where Mr.cod_horario = '" & cboHorario.ItemData(cboHorario.ListIndex) & "'"
 End If
End If

txtNombre.Text = fxSysCleanTxtInject(txtNombre.Text)

  If txtNombre.Text <> "" Then
   If vPaso Then
     strSQL = strSQL & " and Us.DESCRIPCION like '%" & txtNombre.Text & "%'"
   Else
     vPaso = True
     strSQL = strSQL & " where Us.DESCRIPCION like '%" & txtNombre.Text & "%'"
   End If
  End If

strSQL = strSQL & " order by Mr.Usuario,Mr.fecha"

fxSQL = strSQL

End Function


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call cmdBuscar_Click
End If
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = Trim(txtUsuario)
    gBusquedas.Consulta = "Select Nombre,Descripcion From Usuarios"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    gBusquedas.Filtro = " and Estado = 'A'"
    frmBusquedas.Show vbModal
    txtUsuario = Trim(gBusquedas.Resultado)
End If
End Sub

