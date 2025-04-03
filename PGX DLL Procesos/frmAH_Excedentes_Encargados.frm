VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAH_Excedentes_Encargados 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Encargados del Proceso de Excedentes"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   11655
      _Version        =   1441793
      _ExtentX        =   20558
      _ExtentY        =   8493
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
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   1800
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Activo ?"
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
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   11280
      Top             =   720
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   1
      Top             =   1400
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAH_Excedentes_Encargados.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1440
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   582
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   3240
      TabIndex        =   3
      Top             =   1440
      Width           =   6255
      _Version        =   1441793
      _ExtentX        =   11033
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   1
      Left            =   10080
      TabIndex        =   4
      Top             =   1400
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAH_Excedentes_Encargados.frx":0720
   End
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   375
      Left            =   10680
      TabIndex        =   5
      Top             =   1400
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAH_Excedentes_Encargados.frx":0CC4
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   11655
      _Version        =   1441793
      _ExtentX        =   20558
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtEmail 
      Height          =   330
      Left            =   1080
      TabIndex        =   11
      Top             =   1800
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11456
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Email"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Encargados del Proceso de Excedentes"
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
      Height          =   495
      Index           =   2
      Left            =   1995
      TabIndex        =   9
      Top             =   360
      Width           =   7215
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario"
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
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   11655
      _Version        =   1441793
      _ExtentX        =   20558
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Lista de Usuarios Encargados del Proceso de Excedentes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmAH_Excedentes_Encargados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub sbConsulta()
On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spExc_Encargados"
Call OpenRecordSet(rs, strSQL)

With lsw.ListItems
  .Clear
  Do While Not rs.EOF
    Set itmX = .Add(, , rs!Usuario)
        itmX.SubItems(1) = rs!Descripcion
        itmX.SubItems(2) = rs!Activo_Desc & ""
        itmX.SubItems(3) = rs!Email & ""
        itmX.SubItems(4) = rs!Registro_Fecha & ""
        itmX.SubItems(5) = rs!Registro_Usuario & ""
        itmX.SubItems(6) = rs!Modifica_Fecha & ""
        itmX.SubItems(7) = rs!Modifica_Usuario & ""
    rs.MoveNext
  Loop
  rs.Close

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnAccion_Click(Index As Integer)

On Error GoTo vError

If txtUsuario.Text = "" Then Exit Sub

strSQL = "exec  spExc_Encargados_Add '" & txtUsuario.Text & "','" & IIf(Index = 0, "A", "B") & "','" & glogon.Usuario _
        & "', '" & txtEmail.Text & "', " & chkActivo.Value
Call ConectionExecute(strSQL)

Call Bitacora(IIf(Index = 0, "Registra", "Elimina"), "Usuario Encargo de Excedentes: " & txtUsuario.Text)

txtUsuario.Text = ""
txtNombre.Text = ""
txtEmail.Text = ""
chkActivo.Value = xtpChecked

Call sbConsulta

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExport_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
 vModulo = 2

End Sub

Private Sub Form_Load()

vModulo = 2

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2500
    .Add , , "Nombre", 3600
    .Add , , "Activo?", 1100, vbCenter
    .Add , , "Notifica", 4100
    .Add , , "Reg.Fecha", 2800
    .Add , , "Reg.Usuario", 2500
    .Add , , "Act.Fecha", 2800
    .Add , , "Act.Usuario", 2500
End With

Call Formularios(Me)

btnAccion(1).Tag = btnAccion(0).Tag

Call RefrescaTags(Me)

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

txtUsuario.Text = Item.Text
txtNombre.Text = Item.SubItems(1)
chkActivo.Value = IIf((Mid(Item.SubItems(2), 1, 1) = "S"), 1, 0)
txtEmail.Text = Item.SubItems(3)


End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False
Call sbConsulta
End Sub


Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Col1Name = "Usuario"
   gBusquedas.Col2Name = "Nombre"
   gBusquedas.Consulta = "Select Nombre, Descripcion from Usuarios"
   gBusquedas.Filtro = " and Estado = 'A'"
   
   frmBusquedas.Show vbModal
   
   txtUsuario.Text = gBusquedas.Resultado
   txtNombre.Text = gBusquedas.Resultado2
  
End If

End Sub

