VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_CR_NoAumentoTasas_Autorizadores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Autorizadores para Gestión de Renuncia sin Aumento de Tasas"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   11535
      _Version        =   1441793
      _ExtentX        =   20346
      _ExtentY        =   8916
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
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   11280
      Top             =   720
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   6
      Top             =   1400
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_CR_NoAumentoTasas_Autorizadores.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   330
      Left            =   1080
      TabIndex        =   4
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
      TabIndex        =   5
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
      TabIndex        =   7
      Top             =   1400
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_CR_NoAumentoTasas_Autorizadores.frx":0720
   End
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   375
      Left            =   10680
      TabIndex        =   8
      Top             =   1400
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_CR_NoAumentoTasas_Autorizadores.frx":0CC4
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   11535
      _Version        =   1441793
      _ExtentX        =   20346
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   11535
      _Version        =   1441793
      _ExtentX        =   20346
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Lista de Usuarios autorizados para Resolución de Renuncias con NO aumento de tasas"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   240
      TabIndex        =   1
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios Autorizados"
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
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmAF_CR_NoAumentoTasas_Autorizadores"
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

strSQL = "exec spAFI_Renuncia_NAT_Autorizadores"
Call OpenRecordSet(rs, strSQL)

With lsw.ListItems
  .Clear
  Do While Not rs.EOF
    Set itmX = .Add(, , rs!Usuario)
        itmX.SubItems(1) = rs!Descripcion
        itmX.SubItems(2) = rs!Registro_Fecha & ""
        itmX.SubItems(3) = rs!Registro_Usuario & ""
        
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

strSQL = "exec  spAFI_Renuncia_NAT_Autorizadores_Add '" & txtUsuario.Text & "','" & IIf(Index = 0, "A", "B") & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call Bitacora(IIf(Index = 0, "Registra", "Elimina"), "Usuario Autorizador para Renuncias con No Aumento de Tasas: " & txtUsuario.Text)

txtUsuario.Text = ""
txtNombre.Text = ""

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

Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2500
    .Add , , "Nombre", 3600
    .Add , , "Reg.Fecha", 2800
    .Add , , "Reg.Usuario", 2500
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
