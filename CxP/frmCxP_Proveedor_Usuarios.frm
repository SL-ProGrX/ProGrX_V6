VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCxP_Proveedor_Usuarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Usuarios en Línea del Proveedor"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   3840
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   5292
      Begin XtremeSuiteControls.PushButton btnRenovar 
         Height          =   735
         Left            =   1320
         TabIndex        =   4
         Top             =   2640
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Renovar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCxP_Proveedor_Usuarios.frx":0000
      End
      Begin XtremeSuiteControls.PushButton cmdCerrar 
         Height          =   735
         Left            =   3000
         TabIndex        =   5
         Top             =   2640
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Cerrar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCxP_Proveedor_Usuarios.frx":07D8
      End
      Begin XtremeSuiteControls.FlatEdit txtClaveNotas 
         Height          =   975
         Left            =   720
         TabIndex        =   3
         Top             =   1440
         Width           =   3975
         _Version        =   1572864
         _ExtentX        =   7011
         _ExtentY        =   1720
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   $"frmCxP_Proveedor_Usuarios.frx":0FA5
         BackColor       =   12648447
         Alignment       =   2
         MultiLine       =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   852
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12132
         _Version        =   1572864
         _ExtentX        =   21399
         _ExtentY        =   1503
         _StockProps     =   14
         Caption         =   "Renovar Contraseña de acceso"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin VB.Label lblUsuario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   1080
         Width           =   3855
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   9240
      Top             =   840
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   13575
      _Version        =   524288
      _ExtentX        =   23945
      _ExtentY        =   12303
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
      MaxCols         =   486
      ScrollBars      =   2
      SpreadDesigner  =   "frmCxP_Proveedor_Usuarios.frx":1046
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Proveedor"
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
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   7335
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   14055
   End
End
Attribute VB_Name = "frmCxP_Proveedor_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim pEmail As String

Private Sub sbUsuarios_List()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select USUARIO, NOMBRE, EMAIL, WEB_AUTO_GESTION, WEB_FERIAS, ACTIVO" _
       & " from CXP_AG_USUARIOS where cod_Proveedor = " & lblX.Tag

Call sbCargaGrid(vGrid, 6, strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub





Private Function fxGuardar() As Long
Dim pUsuario As String, pNombre As String, pEmail As String, pPortal As Integer, pFerias As Integer, pActivo As Integer

On Error GoTo vError

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1
  
fxGuardar = 0

If vGrid.Text = "" Then Exit Function
 
 vGrid.Col = 1
 pUsuario = vGrid.Text
 vGrid.Col = 2
 pNombre = vGrid.Text
 vGrid.Col = 3
 pEmail = vGrid.Text
 vGrid.Col = 4
 pPortal = vGrid.Value
 vGrid.Col = 5
 pFerias = vGrid.Value
 vGrid.Col = 6
 pActivo = vGrid.Value

 strSQL = "exec spCxP_Proveedores_Usuario_Add " & lblX.Tag & ", '" & pUsuario & "', '" & pNombre & "', '" & pEmail _
        & "', " & pPortal & ", " & pFerias & ", " & pActivo & ", '" & glogon.Usuario & "'"
 Call OpenRecordSet(rs, strSQL)
 If rs!Pass = 1 Then
    Call Bitacora(rs!Movimiento, rs!Mensaje)
 Else
    MsgBox rs!Mensaje, vbExclamation
 End If
 
fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub btnRenovar_Click()


On Error GoTo vError

strSQL = "exec spuProGrX_MOBILE_Proveedor_WebKey_Renueva " & lblX.Tag & ",'" & lblUsuario.Caption & "', '" & pEmail _
        & "', '" & glogon.Usuario & "',''"
Call ConectionExecute(strSQL)

If Not glogon.error Then
    Call Bitacora("Modifica", "Clave para Portal y Ferias Web del Usuario: " & lblUsuario.Caption & "..Proveedor:" & lblX.Tag)
    
    MsgBox "Clave de AutoGestion Renovada satisfactoriamente (Enviada por E-mail)", vbInformation
End If


Call cmdCerrar_Click

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdCerrar_Click()
 fraKey.Visible = False
 vGrid.Enabled = True
End Sub


Private Sub Form_Load()

On Error GoTo vError

vModulo = 30

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle

lblX.Tag = GLOBALES.gTag
lblX.Caption = GLOBALES.gTag2

vError:

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbUsuarios_List

End Sub

Private Sub txtClaveNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then btnRenovar.SetFocus
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

If vGrid.ActiveCol = 1 And KeyCode = vbKeyF2 Then
 vGrid.Row = vGrid.ActiveRow
 vGrid.Col = vGrid.ActiveCol
 If vGrid.Text <> "" Then
    fraKey.Visible = True
    fraKey.Caption = vGrid.Text
    lblUsuario.Caption = vGrid.Text
    
    vGrid.Col = 3
    pEmail = vGrid.Text
    
    vGrid.Enabled = False
  End If
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

End Sub

