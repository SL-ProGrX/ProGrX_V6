VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmPres_Usuarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Usuarios de Presupuesto"
   ClientHeight    =   7635
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2412
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   10692
      _Version        =   1441793
      _ExtentX        =   18860
      _ExtentY        =   4254
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
      Appearance      =   16
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ListView lswContabilidad 
      Height          =   2772
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
      _ExtentY        =   4890
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
      Appearance      =   16
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ListView lswUnidades 
      Height          =   2772
      Left            =   5520
      TabIndex        =   3
      Top             =   4440
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
      _ExtentY        =   4890
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
      Appearance      =   16
      UseVisualStyle  =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   600
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   7380
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCriterio 
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   10692
      _Version        =   1441793
      _ExtentX        =   18860
      _ExtentY        =   656
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios de Presupuesto y Nivel de Visualización"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   1800
      TabIndex        =   7
      Top             =   360
      Width           =   9132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Unidades: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   5520
      TabIndex        =   5
      Top             =   4200
      Width           =   1812
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Contabilidades: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   1812
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmPres_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub Form_Load()
 vModulo = 12
 
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
 
 
lsw.ColumnHeaders.Add , , "Usuario", 2000
lsw.ColumnHeaders.Add , , "Nombre", 4000
lsw.ColumnHeaders.Add , , "Registro: Fecha", 2400
lsw.ColumnHeaders.Add , , "Registro: Usuario", 2000

lswContabilidad.ColumnHeaders.Add , , "Contabilidad", 5000

lswUnidades.ColumnHeaders.Add , , "Código", 1400
lswUnidades.ColumnHeaders.Add , , "Unidad", 3500
 
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
End Sub

Private Sub sbLista_Usuarios()

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

lsw.ListItems.Clear
lswContabilidad.ListItems.Clear
lswUnidades.ListItems.Clear

StatusBarX.Panels(1).Text = ""
StatusBarX.Panels(2).Text = ""

strSQL = "exec spPres_Usuarios_Modulo '%" & txtCriterio.Text & "%'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Usuario)
     itmX.SubItems(1) = rs!Nombre
     itmX.SubItems(2) = rs!REGISTRO_FECHA & ""
     itmX.SubItems(3) = rs!Registro_Usuario & ""
     
     
     itmX.Checked = IIf((rs!Activo = 1), True, False)
 rs.MoveNext
Loop
rs.Close

vPaso = False


Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbLista_Contabilidades()

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

lswContabilidad.ListItems.Clear
lswUnidades.ListItems.Clear

StatusBarX.Panels(2).Text = ""
StatusBarX.Panels(2).Tag = "0"

strSQL = "exec spPres_Usuarios_Modulo_Contabilidades '" & StatusBarX.Panels(1).Text & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswContabilidad.ListItems.Add(, , rs!Nombre)
     itmX.Tag = rs!COD_CONTABILIDAD
     itmX.Checked = IIf((rs!Activo = 1), True, False)
 
 rs.MoveNext
Loop
rs.Close

vPaso = False


Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbLista_Unidades()

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

lswUnidades.ListItems.Clear

strSQL = "exec spPres_Usuarios_Modulo_Unidades '" & StatusBarX.Panels(1).Text & "'," & StatusBarX.Panels(2).Tag
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswUnidades.ListItems.Add(, , rs!Cod_Unidad)
     itmX.Tag = rs!COD_CONTABILIDAD
     itmX.SubItems(1) = rs!Descripcion
     itmX.Checked = IIf((rs!Activo = 1), True, False)
   rs.MoveNext
Loop
rs.Close

vPaso = False


Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spPres_Usuarios_Modulo_Registro '" & Item.Text & "','" & glogon.Usuario & "'," & IIf(Item.Checked, 1, 0)
Call ConectionExecute(strSQL)

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub


StatusBarX.Panels(1).Text = Item.Text

If Item.Checked Then
    Call sbLista_Contabilidades
End If

End Sub



Private Sub lswContabilidad_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

StatusBarX.Panels(2).Text = Item.Text
StatusBarX.Panels(2).Tag = Item.Tag

If Item.Checked Then
    Call sbLista_Unidades
End If

End Sub


Private Sub lswUnidades_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spPres_Usuarios_Unidades_Registro '" & StatusBarX.Panels(1).Text & "'," & Item.Tag _
        & ",'" & Item.Text & "','" & glogon.Usuario & "'," & IIf(Item.Checked, 1, 0)
Call ConectionExecute(strSQL)

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbLista_Usuarios
End Sub

Private Sub txtCriterio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbLista_Usuarios

End Sub
