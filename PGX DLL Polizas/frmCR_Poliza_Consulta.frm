VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_Poliza_Consulta 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Pólizas"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   12705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox gbList 
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   12495
      _Version        =   1441793
      _ExtentX        =   22040
      _ExtentY        =   3836
      _StockProps     =   79
      Caption         =   "Operaciones:"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   1695
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   12495
         _Version        =   1441793
         _ExtentX        =   22040
         _ExtentY        =   2990
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4200
      TabIndex        =   0
      Top             =   480
      Width           =   5532
      _Version        =   1441793
      _ExtentX        =   9758
      _ExtentY        =   550
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.GroupBox gbList 
      Height          =   2175
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   12495
      _Version        =   1441793
      _ExtentX        =   22040
      _ExtentY        =   3836
      _StockProps     =   79
      Caption         =   "Pólizas"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ListView lswPolizas 
         Height          =   1695
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   12495
         _Version        =   1441793
         _ExtentX        =   22040
         _ExtentY        =   2990
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
   End
   Begin XtremeSuiteControls.GroupBox gbList 
      Height          =   3015
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   12495
      _Version        =   1441793
      _ExtentX        =   22040
      _ExtentY        =   5318
      _StockProps     =   79
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ListView lswReclamos 
         Height          =   1935
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   12495
         _Version        =   1441793
         _ExtentX        =   22040
         _ExtentY        =   3413
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   10
         Top             =   2520
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_Poliza_Consulta.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   1
         Left            =   6000
         TabIndex        =   11
         Top             =   2520
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_Poliza_Consulta.frx":0720
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   2
         Left            =   6480
         TabIndex        =   12
         Top             =   2520
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_Poliza_Consulta.frx":0CC4
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   3
         Left            =   7080
         TabIndex        =   13
         Top             =   2520
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_Poliza_Consulta.frx":12BF
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   12495
         _Version        =   1441793
         _ExtentX        =   22040
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Reclamos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
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
      Height          =   252
      Index           =   5
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmCR_Poliza_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, vPaso As Boolean

Private Sub btnAccion_Click(Index As Integer)
Call sbFormsCall("frmPoliza_Reclamo", vbModal, , , False, Me, True)
End Sub

Private Sub Form_Load()

vModulo = 2

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture


With lsw.ColumnHeaders
    .Clear
    .Add , , "No. Operación", 2000
    .Add , , "Código", 1000, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Mensualidad", 2100, vbRightJustify
    .Add , , "Plazo", 1000, vbCenter
    .Add , , "Formalización", 2500, vbCenter
    .Add , , "Estado", 1500, vbCenter
End With


With lswPolizas.ColumnHeaders
    .Clear
    .Add , , "Id Póliza", 2000
    .Add , , "Cód. Póliza", 1100, vbCenter
    .Add , , "Cód. Retención", 1100, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Mensualidad", 2100, vbRightJustify
    .Add , , "Estado", 1500, vbCenter
End With

With lswReclamos.ColumnHeaders
    .Clear
    .Add , , "Id Reclamo", 2000
    .Add , , "Id Póliza", 2000, vbCenter
    .Add , , "Cód. Póliza", 1100, vbCenter
    .Add , , "F. Registro", 2500, vbCenter
    .Add , , "Estado", 1500, vbCenter
    .Add , , "Identificación", 1500, vbCenter
    .Add , , "Nombre", 3000
End With

Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

Private Sub sbOperaciones_Load()

On Error GoTo vError

lsw.ListItems.Clear
lswPolizas.ListItems.Clear
lswReclamos.ListItems.Clear

Me.MousePointer = vbHourglass

'strSQL = "exec spPolizas_Persona_Consulta '" & txtCedula.Text & "'"
'Call OpenRecordSet(rs, strSQL)
'Do While Not rs.EOF
' Set itmX = lsw.ListItems.Add(, , rs!x)
'     itmX.SubItems(1) = rs!x
'     itmX.SubItems(2) = rs!x
'     itmX.SubItems(3) = rs!x
'     itmX.SubItems(4) = rs!x
'     itmX.SubItems(5) = rs!x
'     itmX.SubItems(6) = rs!x
'     itmX.SubItems(7) = rs!x
'     itmX.SubItems(8) = rs!x
'     itmX.SubItems(9) = rs!x
' rs.MoveNext
'Loop
'rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbPolizas_Load(pOperacion As Long)

On Error GoTo vError

lswPolizas.ListItems.Clear
lswReclamos.ListItems.Clear

Me.MousePointer = vbHourglass

'strSQL = "exec spPolizas_Persona_Consulta '" & txtCedula.Text & "'"
'Call OpenRecordSet(rs, strSQL)
'Do While Not rs.EOF
' Set itmX = lsw.ListItems.Add(, , rs!x)
'     itmX.SubItems(1) = rs!x
'     itmX.SubItems(2) = rs!x
'     itmX.SubItems(3) = rs!x
'     itmX.SubItems(4) = rs!x
'     itmX.SubItems(5) = rs!x
'     itmX.SubItems(6) = rs!x
'     itmX.SubItems(7) = rs!x
'     itmX.SubItems(8) = rs!x
'     itmX.SubItems(9) = rs!x
' rs.MoveNext
'Loop
'rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbReclamos_Load(pOperacion As Long, Optional pPoliza As Long = 0)

On Error GoTo vError

lswReclamos.ListItems.Clear

Me.MousePointer = vbHourglass

'strSQL = "exec spPolizas_Persona_Consulta '" & txtCedula.Text & "'"
'Call OpenRecordSet(rs, strSQL)
'Do While Not rs.EOF
' Set itmX = lsw.ListItems.Add(, , rs!x)
'     itmX.SubItems(1) = rs!x
'     itmX.SubItems(2) = rs!x
'     itmX.SubItems(3) = rs!x
'     itmX.SubItems(4) = rs!x
'     itmX.SubItems(5) = rs!x
'     itmX.SubItems(6) = rs!x
'     itmX.SubItems(7) = rs!x
'     itmX.SubItems(8) = rs!x
'     itmX.SubItems(9) = rs!x
' rs.MoveNext
'Loop
'rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id Alterna"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "cedula"
    gBusquedas.Orden = "cedula"
    frmBusquedas.Show vbModal
    
    txtCedula.Text = gBusquedas.Resultado
    txtNombre.Text = gBusquedas.Resultado2
    
    Call sbOperaciones_Load
End If
End Sub
