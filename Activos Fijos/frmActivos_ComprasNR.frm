VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmActivos_ComprasNR 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Compras de Activos Fijos No Registrados"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   14595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   14415
      _Version        =   1572864
      _ExtentX        =   25426
      _ExtentY        =   11456
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
      Appearance      =   21
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnConsultar 
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   600
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmActivos_ComprasNR.frx":0000
   End
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   135
      Left            =   120
      TabIndex        =   4
      Top             =   7560
      Width           =   14415
      _Version        =   1572864
      _ExtentX        =   25426
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   600
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
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   600
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmActivos_ComprasNR.frx":0700
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   2520
      TabIndex        =   8
      Top             =   600
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeShortcutBar.ShortcutCaption lblY 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14655
      _Version        =   1572864
      _ExtentX        =   25850
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   $"frmActivos_ComprasNR.frx":086A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Periodo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label lblX 
      Height          =   255
      Left            =   7680
      TabIndex        =   1
      Top             =   600
      Width           =   5895
      _Version        =   1572864
      _ExtentX        =   10398
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmActivos_ComprasNR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub btnConsultar_Click()
 Call sbBuscar
End Sub

Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

PrgBar.Visible = True

Call Excel_Exportar_Lsw(lsw, PrgBar)

PrgBar.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub dtpFecha_Change()
 
lblX.Caption = fxActivos_Periodo(dtpFecha.Value)
 
End Sub

Private Sub Form_Load()
vModulo = 36

With lsw.ColumnHeaders
    .Clear
    .Add , , "No. Factura", 4200
    .Add , , "Línea", 1200, vbCenter
    .Add , , "Proveedor", 3200
    .Add , , "Activo Nombre", 3200
    .Add , , "Cantidad", 1200, vbCenter
    .Add , , "Registrada", 1200, vbCenter
    .Add , , "Pendiente", 1200, vbCenter
    .Add , , "Valor Adq.", 2200, vbRightJustify
End With

cbo.Clear
cbo.AddItem "Pendientes"
cbo.AddItem "Registrados"
cbo.AddItem "Todos"
cbo.Text = "Pendientes"

dtpFecha.Value = fxFechaServidor
Call dtpFecha_Change

End Sub


Private Sub lsw_DblClick()

If lsw.ListItems.Count <= 0 Then Exit Sub

On Error GoTo vError

gAsistente.Documento = lsw.SelectedItem.Text
gAsistente.Proveedor = lsw.SelectedItem.Tag
gAsistente.VU = CCur(lsw.SelectedItem.SubItems(7))
gAsistente.Tipo = "C"

Call sbFormsCall("frmActivos_Main", , , , , Me, True)

vError:

End Sub


Private Sub sbBuscar()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spActivos_Compras_Pendientes_Registro '" & Format(dtpFecha.Value, "yyyy-mm-dd") & "', '" & Mid(cbo.Text, 1, 1) & "'"
Call OpenRecordSet(rs, strSQL, 0)

lsw.ListItems.Clear

PrgBar.Max = rs.RecordCount + 2
PrgBar.Value = 1
PrgBar.Visible = True

Do While Not rs.EOF
 
    Set itmX = lsw.ListItems.Add(, , rs!cod_factura)
        itmX.Tag = rs!COD_PROVEEDOR
        itmX.SubItems(1) = rs!Linea
        itmX.SubItems(2) = rs!Proveedor
        itmX.SubItems(3) = rs!producto
        itmX.SubItems(4) = rs!cantidad
        itmX.SubItems(5) = rs!Registrados
        itmX.SubItems(6) = rs!Pendientes
        itmX.SubItems(7) = Format(rs!Precio, "Standard")
   
 PrgBar.Value = PrgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
PrgBar.Visible = False

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub
