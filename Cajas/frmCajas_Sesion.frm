VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCajas_Sesion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sesión de Cliente en Cajas"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2415
      Left            =   0
      TabIndex        =   15
      Top             =   3360
      Width           =   10695
      _Version        =   1572864
      _ExtentX        =   18865
      _ExtentY        =   4260
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
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   10455
      _Version        =   1572864
      _ExtentX        =   18441
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Datos de la Sesión"
      ForeColor       =   16711680
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
      Appearance      =   21
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         Top             =   480
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtInicio 
         Height          =   315
         Left            =   6000
         TabIndex        =   13
         Top             =   480
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label lblROE 
         Height          =   615
         Left            =   8520
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Requiere ROE"
         ForeColor       =   16777215
         BackColor       =   33023
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
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
         Left            =   5040
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
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
         Index           =   0
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
   End
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
      _Version        =   1572864
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
      _Version        =   1572864
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
      _Version        =   1572864
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
      Left            =   240
      TabIndex        =   5
      Top             =   5760
      Width           =   10215
      _Version        =   1572864
      _ExtentX        =   18018
      _ExtentY        =   1720
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSesion 
         Height          =   495
         Index           =   0
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   1695
         _Version        =   1572864
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
         Left            =   5040
         TabIndex        =   7
         Top             =   360
         Width           =   1695
         _Version        =   1572864
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
         Left            =   8400
         TabIndex        =   8
         Top             =   360
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Continuar con la sesión actual"
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
      Begin XtremeSuiteControls.FlatEdit txtTotal 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Movimientos"
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
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   3000
      Width           =   10695
      _Version        =   1572864
      _ExtentX        =   18865
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Movimientos Registrados"
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
Dim itmX As ListViewItem

Private Sub btnSesion_Click(Index As Integer)
On Error GoTo vError

If txtCedula.Text = "" Or txtNombre.Text = "" Then
    MsgBox "Faltan Datos!", vbExclamation
    Exit Sub
End If

Select Case Index
    Case 0 'Inicia
        'Exec Inicio de Sesion
        
        strSQL = "exec spCajas_Sesion_Inicia '" & ModuloCajas.mCaja & "', '" & glogon.Usuario _
               & "', " & ModuloCajas.mApertura & ", " & cboTipoId.ItemData(cboTipoId.ListIndex) _
               & ", '" & txtCedula.Text & "', '" & txtNombre.Text & "'"
        Call OpenRecordSet(rs, strSQL)
        
        
        ModuloCajas.mSesionId = rs!SesionId
        ModuloCajas.mSesionCedula = txtCedula.Text
        ModuloCajas.mSesionNombre = txtNombre.Text
    Case 1 'Cierra
            
        strSQL = "exec spCajas_Sesion_Finaliza " & ModuloCajas.mSesionId & ", '" & glogon.Usuario & "'"
        Call OpenRecordSet(rs, strSQL)
        
        ModuloCajas.mSesionId = 0
        ModuloCajas.mSesionCedula = ""
        ModuloCajas.mSesionNombre = ""
        
        txtCedula.Locked = False
        txtNombre.Locked = False
        
        txtCedula.Text = ModuloCajas.mClienteId
        txtNombre.Text = ModuloCajas.mCliente
        
        txtEstado.Text = ""
        txtInicio.Text = ""
        
        lsw.ListItems.Clear
        txtTotal.Text = Format(0, "Standard")
    
        btnSesion(0).Enabled = True
        btnSesion(1).Enabled = False
        
        If rs!ROE > 0 Then
           GLOBALES.gTag = "ROE_" & rs!ROE
           MsgBox "Sesión finalizada Correctamente, Este caso necesita el registro del ROE!", vbInformation
           Call sbFormsCall("frmCajas_ROE", vbModal, , , False, Me, True)
        Else
            MsgBox "Sesión finalizada Correctamente!", vbInformation
        End If
        
        
    
    
    Case 2 'Salir
End Select

If Index = 0 Or Index = 2 Then
 Unload Me
End If

Exit Sub


vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

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


btnSesion(0).Enabled = True
btnSesion(1).Enabled = True

strSQL = "exec spCajas_Sesion_Info " & ModuloCajas.mSesionId & ", '" & ModuloCajas.mCaja _
       & "', '" & ModuloCajas.mUsuario & "', " & ModuloCajas.mApertura
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    If rs!estado = 1 Then
        btnSesion(0).Enabled = False
    End If
    
    If rs!estado = 2 Then
        btnSesion(1).Enabled = False
    End If
    
    Call sbCboAsignaDato(cboTipoId, rs!TipoId_Desc, True, rs!Tipo_Id)
    
    ModuloCajas.mSesionId = rs!ID_SESION
    
    ModuloCajas.mSesionCedula = rs!Identificacion
    ModuloCajas.mSesionNombre = rs!Nombre
    
    txtCedula.Text = rs!Identificacion
    txtNombre.Text = rs!Nombre
    
    txtEstado.Text = rs!ESTADO_DESC
    txtInicio.Text = rs!Fecha_Inicio

    If rs!ROE = 1 Then
        lblROE.Visible = True
    Else
        lblROE.Visible = False
    End If
    
    txtCedula.Locked = True
    txtNombre.Locked = True
Else
    lblROE.Visible = False
    
    txtCedula.Locked = False
    txtNombre.Locked = False
    
    txtCedula.Text = ModuloCajas.mClienteId
    txtNombre.Text = ModuloCajas.mCliente
    
    txtEstado.Text = ""
    txtInicio.Text = ""

    btnSesion(1).Enabled = False
End If

Dim pMonto As Currency

pMonto = 0
lsw.ListItems.Clear


With lsw.ColumnHeaders
    .Clear
    .Add , , "Tipo   Doc.", 1800
    .Add , , "Número Doc.", 2500, vbCenter
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Fecha", 2500
    .Add , , "Identificación", 2500, vbCenter
    .Add , , "Nombre", 3500
    .Add , , "Concepto", 3500, vbCenter
    .Add , , "Documento", 3500, vbCenter
    .Add , , "Ref No. 1", 1500, vbCenter
    .Add , , "Ref No. 2", 1500, vbCenter
    .Add , , "Ref No. 3", 1500, vbCenter
    
    .Add , , "Usuario", 1500, vbCenter
    .Add , , "Caja", 1000, vbCenter
    .Add , , "Apertura", 1000, vbCenter
    
End With

strSQL = "exec spCajas_Sesion_Aplicaciones " & ModuloCajas.mSesionId
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Tipo_documento)
     itmX.SubItems(1) = rs!Cod_Transaccion
     itmX.SubItems(2) = Format(rs!Monto, "Standard")
     itmX.SubItems(3) = rs!Registro_Fecha
     itmX.SubItems(4) = rs!Cliente_Identificacion
     itmX.SubItems(5) = rs!Cliente_Nombre
     itmX.SubItems(6) = rs!Concepto_Desc
     itmX.SubItems(7) = rs!Documento_Desc
     itmX.SubItems(8) = rs!Referencia_01 & ""
     itmX.SubItems(9) = rs!Referencia_02 & ""
     itmX.SubItems(10) = rs!Referencia_03 & ""
     
     itmX.SubItems(11) = rs!Registro_Usuario
     itmX.SubItems(12) = rs!COD_CAJA
     itmX.SubItems(13) = rs!Cod_Apertura
     
     pMonto = pMonto + rs!Monto
 
 rs.MoveNext
Loop
rs.Close

txtTotal.Text = Format(pMonto, "Standard")


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
