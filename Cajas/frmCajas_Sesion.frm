VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCajas_Sesion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sesión de Cliente en Cajas"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerCaja 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   4680
      TabIndex        =   0
      Top             =   1440
      Width           =   5895
      _Version        =   1441793
      _ExtentX        =   10393
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3619
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
   Begin XtremeSuiteControls.ComboBox cboTipoId 
      Height          =   330
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
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
   End
   Begin XtremeSuiteControls.GroupBox gbAccion 
      Height          =   975
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   9975
      _Version        =   1441793
      _ExtentX        =   17595
      _ExtentY        =   1720
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSesion 
         Height          =   495
         Index           =   0
         Left            =   4200
         TabIndex        =   6
         Top             =   240
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Iniciar Sesión"
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
         Picture         =   "frmCajas_Sesion.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnSesion 
         Height          =   495
         Index           =   1
         Left            =   5880
         TabIndex        =   7
         Top             =   240
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cerrar Sesión"
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
         Picture         =   "frmCajas_Sesion.frx":0727
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnSesion 
         Height          =   495
         Index           =   2
         Left            =   9000
         TabIndex        =   8
         Top             =   240
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Salir"
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
         Picture         =   "frmCajas_Sesion.frx":0E3D
         ImageAlignment  =   4
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sesión de Cliente en Cajas"
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
      Index           =   2
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "No. Identificación"
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
      Index           =   3
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmCajas_Sesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub btnSesion_Click(Index As Integer)

Select Case Index
    Case 0 'Inicia
        'Exec Inicio de Sesion
        ModuloCajas.mSesionId = 1000
        ModuloCajas.mSesionCedula = txtCedula.Text
        ModuloCajas.mSesionNombre = txtNombre.Text
    Case 1 'Cierra
        'Exec Cierre de Sesion
        ModuloCajas.mSesionId = 0
        ModuloCajas.mSesionCedula = ""
        ModuloCajas.mSesionNombre = ""
    Case 2 'Salir
End Select

If Index = 0 Or Index = 2 Then
 Unload Me
End If
End Sub

Private Sub Form_Load()

vModulo = 5

If ModuloCajas.mSesionId > 0 Then
    txtCedula.Text = ModuloCajas.mSesionCedula
    txtNombre.Text = ModuloCajas.mSesionNombre
Else
    txtCedula.Text = ModuloCajas.mClienteId
    txtNombre.Text = ModuloCajas.mCliente
End If

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Me.Caption = "Sesión de Cliente ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
       & " order by Tipo_Id"
Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)



End Sub


Private Sub TimerCaja_Timer()
TimerCaja.Interval = 0
TimerCaja.Enabled = False

'Paso 1: Si la Caja no está abierta (Llamar pantalla de login de Caja)
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   Call sbFormsCall("frmCajas_Acceso", vbModal, , , False, Me)
End If

'Paso 2: Si despues del Login de Caja permanece sin Apertura Salir
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   MsgBox "No se ha indicado ninguna caja con Apertura disponible?", vbExclamation
   Unload Me
   Exit Sub
End If

'Paso 3: Continuar con Barra de Información
'lblInfoApertura.Caption = ModuloCajas.mApertura
'lblInfoCaja.Caption = ModuloCajas.mCaja
'lblInfoUsuario.Caption = ModuloCajas.mUsuario


Me.Caption = "Sesión de Cliente en Cajas ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
End Sub


Private Sub txtCedula_LostFocus()
On Error GoTo vError
        
If txtCedula.Text <> "" Then
    Call gBase_Padron(txtCedula.Text, "General", rs, "CRI")
    
    If rs.RecordCount > 0 Then
       txtNombre.Text = Trim(rs!Apellido_1) & " " & Trim(rs!Apellido_2) & " " & Trim(rs!Nombre)
    End If
End If

vError:
End Sub
