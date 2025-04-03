VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmSUGEF_Informes_Archivos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Módulo de SUGEF"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnAdministrador 
      Height          =   375
      Left            =   9840
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Seg. Administracion"
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
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   11415
      _Version        =   1572864
      _ExtentX        =   20135
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7455
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   11415
      _Version        =   1572864
      _ExtentX        =   20135
      _ExtentY        =   13150
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Cortes"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "scCorte"
      Item(0).Control(2)=   "btnCorte(0)"
      Item(0).Control(3)=   "btnCorte(1)"
      Item(0).Control(4)=   "btnCorte(2)"
      Item(0).Control(5)=   "vGrid"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   9
      Item(1).Control(0)=   "lblX"
      Item(1).Control(1)=   "dtpCorte"
      Item(1).Control(2)=   "Label1(0)"
      Item(1).Control(3)=   "dtpRngInicio"
      Item(1).Control(4)=   "dtpRngCorte"
      Item(1).Control(5)=   "Label1(1)"
      Item(1).Control(6)=   "txtDescripcion"
      Item(1).Control(7)=   "btnCorte_Add"
      Item(1).Control(8)=   "lblStatus"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   2775
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   11415
         _Version        =   1572864
         _ExtentX        =   20135
         _ExtentY        =   4895
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   21
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnCorte_Add 
         Height          =   495
         Left            =   -66160
         TabIndex        =   15
         Top             =   3600
         Visible         =   0   'False
         Width           =   3615
         _Version        =   1572864
         _ExtentX        =   6376
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Crear Corte y Procesar Base de Datos"
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
         Picture         =   "frmSUGEF_Informes_Archivos.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnCorte 
         Height          =   375
         Index           =   0
         Left            =   6960
         TabIndex        =   5
         Top             =   3120
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Volver a Generar"
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
         Picture         =   "frmSUGEF_Informes_Archivos.frx":0708
      End
      Begin XtremeSuiteControls.PushButton btnCorte 
         Height          =   375
         Index           =   1
         Left            =   8520
         TabIndex        =   6
         Top             =   3120
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Archivo XML"
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
         Picture         =   "frmSUGEF_Informes_Archivos.frx":0E21
      End
      Begin XtremeSuiteControls.PushButton btnCorte 
         Height          =   375
         Index           =   2
         Left            =   10080
         TabIndex        =   7
         Top             =   3120
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Exportar"
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
         Picture         =   "frmSUGEF_Informes_Archivos.frx":1552
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3735
         Left            =   0
         TabIndex        =   8
         Top             =   3600
         Width           =   11415
         _Version        =   524288
         _ExtentX        =   20135
         _ExtentY        =   6588
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
         MaxCols         =   19
         ScrollBars      =   2
         SpreadDesigner  =   "frmSUGEF_Informes_Archivos.frx":16BC
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   330
         Left            =   -68080
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.DateTimePicker dtpRngInicio 
         Height          =   330
         Left            =   -68080
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.DateTimePicker dtpRngCorte 
         Height          =   330
         Left            =   -66760
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   975
         Left            =   -68080
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
         _ExtentY        =   1720
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label lblStatus 
         Height          =   735
         Left            =   -67480
         TabIndex        =   16
         Top             =   4320
         Visible         =   0   'False
         Width           =   6135
         _Version        =   1572864
         _ExtentX        =   10821
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Se estan creado el corte y procesando la base de datos de resultados, este proceso puede durar varios minutos, espere!"
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   1
         Left            =   -69520
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Descripción"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   0
         Left            =   -69520
         TabIndex        =   10
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Rango de Fecha"
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
      Begin XtremeShortcutBar.ShortcutCaption scCorte 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   3120
         Width           =   11415
         _Version        =   1572864
         _ExtentX        =   20135
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "(Seleccione un Corte)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label lblX 
         Height          =   255
         Left            =   -69520
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Corte"
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
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Módulo de Informes y Generación de XML para SUGEF"
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
      Height          =   612
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   7572
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmSUGEF_Informes_Archivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub sbCortes_List()

On Error GoTo vError

Me.MousePointer = vbHourglass


lsw.ListItems.Clear

vGrid.MaxRows = 0
scCorte.Caption = "(Seleccione un Corte)"
scCorte.Tag = ""

btnCorte(0).Enabled = False
btnCorte(1).Enabled = False
btnCorte(2).Enabled = False

strSQL = "select Corte, Descripcion, Genera_Base, Genera_Fecha, Genera_Usuario, Archivo_Genera, Archivo_Fecha, Archivo_Usuario" _
       & ", Rango_Inicio, Rango_Corte  From SUGEF_Facilidades_Crediticias_Cortes order by Corte desc"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , Format(rs!Corte, "yyyy-mm-dd"))
        itmX.SubItems(1) = rs!Descripcion
        itmX.SubItems(2) = Format(rs!Rango_Inicio, "yyyy-mm-dd")
        itmX.SubItems(3) = Format(rs!Rango_Corte, "yyyy-mm-dd")
        itmX.SubItems(4) = rs!Genera_Base & ""
        itmX.SubItems(5) = rs!Genera_Fecha & ""
        itmX.SubItems(6) = rs!Genera_Usuario & ""
        
        itmX.SubItems(7) = rs!Archivo_Genera & ""
        itmX.SubItems(8) = rs!Archivo_Fecha & ""
        itmX.SubItems(9) = rs!Archivo_Usuario & ""
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbCorte_Movimientos(pCorte As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

vGrid.MaxRows = 0
scCorte.Caption = "Corte : " & pCorte
scCorte.Tag = pCorte

btnCorte(0).Enabled = True
btnCorte(1).Enabled = True
btnCorte(2).Enabled = True

strSQL = "select Id, Accion, NumerdoIdentificacion, TipoIdentificacion, NombreCliente, PrimerApellidoCliente, SegundoApellidoCliente, NombreEmpresa" _
       & "             , TipoReporte, TipoOperacion, TipoMovimiento, TipoIngreso, TipoSalida, TipoMonedaMovimiento" _
       & "             , MontoMovimiento, FechaTransaccion, MotivoTransaccion, OrigenRecursos, MotivoCredito" _
       & " from SUGEF_Facilidades_Crediticias Where Corte = '" & scCorte.Tag & "' order by Id"
       
Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL)
       
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbCorte_Procesar(pCorte As Date, pDescripcion As String, pRngInicio As Date, pRngCorte As Date)
On Error GoTo vError

lblStatus.Visible = True
DoEvents

Me.MousePointer = vbHourglass

pDescripcion = fxSysCleanTxtInject(pDescripcion)


strSQL = "exec spSUGEF_Facilidades_Crediticias_Corte '" & Format(pCorte, "yyyy-mm-dd") & "', '" & pDescripcion _
       & "', '" & Format(pRngInicio, "yyyy-mm-dd") & "', '" & Format(pRngCorte, "yyyy-mm-dd") & " 23:59', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

lblStatus.Visible = False

Me.MousePointer = vbDefault

MsgBox "Corte Generado Satisfactoriamente!", vbInformation

tcMain.Item(0).Selected = True

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCorte_Archivo(pCorte As String)

Dim fn
Dim strArchivo As String, strPath As String

On Error GoTo vError

fn = FreeFile


On Error Resume Next

strArchivo = SIFGlobal.DirectorioDeResultados & "\SUGEF"
strPath = Dir(strArchivo, vbDirectory)

If strPath = "" Then
   ChDir ("C:\")
   MkDir (strArchivo)
Else
   strPath = Dir(strArchivo, vbDirectory)
   
   If strPath = "" Then
      ChDir ("C:\")
      MkDir (strArchivo)
   Else
      strPath = Dir(strArchivo, vbDirectory)
      
      If strPath = "" Then
         ChDir ("C:\")
         MkDir (strArchivo)
      End If
   End If
End If

ChDir (strArchivo)

Me.MousePointer = vbHourglass


strArchivo = strArchivo & "\Facilidades_Crediticia_" & pCorte & ".xml"
Open strArchivo For Output As #1


strSQL = "exec spSUGEF_Facilidades_Crediticias_Archivo '" & pCorte & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

ProgressBarX.Visible = True

ProgressBarX.Max = rs.RecordCount + 1
ProgressBarX.Value = 1

Do While Not rs.EOF
    
    If rs!XML_TEXT <> "" Then
        Print #1, rs!XML_TEXT
    End If
    
 ProgressBarX.Value = ProgressBarX.Value + 1
 rs.MoveNext
Loop
rs.Close

Close #1   ' Close file.
 
 
ProgressBarX.Visible = False
Me.MousePointer = vbDefault

MsgBox "Archivo Generado en: " & strArchivo, vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  ProgressBarX.Visible = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnCorte_Add_Click()

Call sbCorte_Procesar(dtpCorte.Value, txtDescripcion.Text, dtpRngInicio.Value, dtpRngCorte.Value)

End Sub

Private Sub sbExportar()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = vGrid.MaxCols
    vHeaders.Headers(1) = "Id"
    vHeaders.Headers(2) = "Acción"
    vHeaders.Headers(3) = "Idenficación"
    vHeaders.Headers(4) = "Tipo Id"
    vHeaders.Headers(5) = "Nombre"
    vHeaders.Headers(6) = "Apellido 1"
    vHeaders.Headers(7) = "Apellido 2"
    vHeaders.Headers(8) = "Empresa"
    vHeaders.Headers(9) = "Tipo Reporte"
    vHeaders.Headers(10) = "Tipo Operación"
    vHeaders.Headers(11) = "Tipo Movimiento"
    vHeaders.Headers(12) = "Tipo Ingreso"
    vHeaders.Headers(13) = "Tipo Salida"
    vHeaders.Headers(14) = "Tipo Moneda"
    vHeaders.Headers(15) = "Monto"
    vHeaders.Headers(16) = "Fecha"
    vHeaders.Headers(17) = "Motivo"
    vHeaders.Headers(18) = "Origen Recursos"
    vHeaders.Headers(19) = "Motivo Credito"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_SUGEF_Facilidades_Crediticia_" & scCorte.Tag)

End Sub

Private Sub btnCorte_Click(Index As Integer)
Select Case Index
    Case 0 'Volver a Crear
        Call sbCorte_Procesar(dtpCorte.Value, txtDescripcion.Text, dtpRngInicio.Value, dtpRngCorte.Value)
        Call sbCorte_Movimientos(dtpCorte.Value)
    Case 1 'Archivo XML
        Call sbCorte_Archivo(scCorte.Tag)
    Case 2 'Exportar
        Call sbExportar
End Select
End Sub

Private Sub Form_Load()

vModulo = 10
 
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
 
With lsw.ColumnHeaders
  .Clear
  .Add , , "Corte", 1200
  .Add , , "Descripción", 3000
  .Add , , "Rng Inicio", 1400, vbCenter
  .Add , , "Rng Corte", 1400, vbCenter
  
  .Add , , "Generado?", 1200, vbCenter
  .Add , , "Gen.Fecha", 2100, vbCenter
  .Add , , "Gen.Usuario", 2100, vbCenter
  
  .Add , , "Archivo?", 1200, vbCenter
  .Add , , "Arc.Fecha", 2100, vbCenter
  .Add , , "Arc.Usuario", 2100, vbCenter
End With

vGrid.MaxRows = 0
scCorte.Caption = "(Seleccione un Corte)"
scCorte.Tag = ""

tcMain.Item(0).Selected = True

Call Formularios(Me)

btnCorte(0).Tag = btnAdministrador.Tag
btnCorte(1).Tag = btnAdministrador.Tag
btnCorte(2).Tag = btnAdministrador.Tag

btnCorte_Add.Tag = btnAdministrador.Tag

Call RefrescaTags(Me)
 
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

dtpCorte.Value = Item.Text
txtDescripcion.Text = Item.SubItems(1)
dtpRngInicio.Value = Item.SubItems(2)
dtpRngCorte.Value = Item.SubItems(3)

Call sbCorte_Movimientos(Item.Text)

End Sub

Private Sub sbLimpia()

dtpCorte.Value = fxFechaServidor
dtpRngInicio.Value = dtpCorte.Value
dtpRngCorte.Value = dtpCorte.Value

txtDescripcion.Text = ""

lblStatus.Visible = False

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Cortes
        Call sbCortes_List
        
    Case 1 'Registro
        Call sbLimpia
End Select

End Sub

Private Sub sbInicial()
    Call sbCortes_List
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False
Call sbInicial
End Sub
