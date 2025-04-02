VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmSYS_RA_Usuarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RA Expedientes: Usuarios Autorizados"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   12075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6255
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   12015
      _Version        =   1310723
      _ExtentX        =   21193
      _ExtentY        =   11033
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   375
      Left            =   10920
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1931
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
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   10695
      _Version        =   1310723
      _ExtentX        =   18865
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios Autorizados a Consultar Expedientes Restringidos"
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
      Height          =   480
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   10455
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "frmSYS_RA_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Private Sub btnExport_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

Call Excel_Exportar_Lsw(lsw)

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 10
End Sub

Private Sub sbConsulta()
      
On Error GoTo vError
      
Me.MousePointer = vbHourglass
      
txtFiltro.Text = fxSysCleanTxtInject(txtFiltro.Text)
      

lsw.ListItems.Clear

strSQL = "exec spSYS_RA_Usuarios_Consulta '" & txtFiltro.Text & "'"
Call OpenRecordSet(rs, strSQL)

vPaso = True

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Usuario)
     itmX.SubItems(1) = rs!Nombre
     itmX.SubItems(2) = Format(rs!registro_Fecha & "", "yyyy-mm-dd")
     itmX.SubItems(3) = rs!Registro_Usuario & ""
     
     itmX.SubItems(4) = Format(rs!Activa_Fecha & "", "yyyy-mm-dd")
     itmX.SubItems(5) = rs!Activa_Usuario & ""
     
     itmX.SubItems(6) = Format(rs!Inactiva_Fecha & "", "yyyy-mm-dd")
     itmX.SubItems(7) = rs!Inactiva_Usuario & ""
     
     
     itmX.Checked = IIf((rs!Activo = 1), vbChecked, vbUnchecked)
       
 rs.MoveNext
Loop
rs.Close

vPaso = False

Call Formularios(Me)
Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Load()

vModulo = 10

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2200
    .Add , , "Nombre", 3500
    .Add , , "Rec.Fecha", 1800, vbCenter
    .Add , , "Rec.Usuario", 2200, vbCenter
    .Add , , "Act.Fecha", 1800, vbCenter
    .Add , , "Act.Usuario", 2200, vbCenter
    .Add , , "Ina.Fecha", 1800, vbCenter
    .Add , , "Ina.Usuario", 2200, vbCenter
End With

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)


If vPaso Then Exit Sub
If lsw.ListItems.Count = 0 Then Exit Sub


On Error GoTo vError
      
Me.MousePointer = vbHourglass

strSQL = "exec spSYS_RA_Usuarios_Add '" & Item.Text & "', " _
        & IIf(Item.Checked, 1, 0) & ", '" & glogon.Usuario & "'"

Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "RA Usuario Autorizado: " & Item.Text & ", " & IIf(Item.Checked, "Activa", "Inactiva"))

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbConsulta

End Sub

Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then sbConsulta
End Sub
