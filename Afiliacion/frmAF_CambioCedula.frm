VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_CambioCedula 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Identificación"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   Icon            =   "frmAF_CambioCedula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   11160
   Begin XtremeSuiteControls.GroupBox gbNueva 
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   10695
      _Version        =   1441793
      _ExtentX        =   18865
      _ExtentY        =   2355
      _StockProps     =   79
      Caption         =   "Cambiar por:"
      ForeColor       =   16711680
      BackColor       =   16777215
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
      Begin XtremeSuiteControls.FlatEdit txtCedNew 
         Height          =   330
         Left            =   0
         TabIndex        =   7
         Top             =   720
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboTipoId 
         Height          =   330
         Left            =   2160
         TabIndex        =   10
         Top             =   720
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.FlatEdit txtNombreNew 
         Height          =   330
         Left            =   4080
         TabIndex        =   15
         Top             =   720
         Width           =   4455
         _Version        =   1441793
         _ExtentX        =   7858
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEstadoNew 
         Height          =   330
         Left            =   8520
         TabIndex        =   17
         Top             =   720
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Id"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   2160
         TabIndex        =   19
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   8520
         TabIndex        =   18
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   4080
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Identificación Nueva"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   10695
      _Version        =   1441793
      _ExtentX        =   18865
      _ExtentY        =   2355
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   615
         Left            =   8640
         TabIndex        =   4
         Top             =   360
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         Picture         =   "frmAF_CambioCedula.frx":000C
      End
      Begin XtremeSuiteControls.CheckBox chkUsuarioValida 
         Height          =   210
         Left            =   4200
         TabIndex        =   20
         Top             =   480
         Width           =   210
         _Version        =   1441793
         _ExtentX        =   360
         _ExtentY        =   360
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   556
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuarioClave 
         Height          =   315
         Left            =   2160
         TabIndex        =   22
         Top             =   480
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   556
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
         Alignment       =   2
         PasswordChar    =   "*"
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label lblAutorizador 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Clave"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   24
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblAutorizador 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2055
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   240
      TabIndex        =   2
      Top             =   1680
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTipoId 
      Height          =   330
      Left            =   2400
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
      Left            =   4320
      TabIndex        =   11
      Top             =   1680
      Width           =   4455
      _Version        =   1441793
      _ExtentX        =   7858
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   330
      Left            =   8760
      TabIndex        =   13
      Top             =   1680
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   8760
      TabIndex        =   14
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   12
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Id"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   9
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cambio de Identificación"
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
      Left            =   2115
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación Actual"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmAF_CambioCedula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub sbLimpia()

txtCedula.Text = ""
txtCedNew.Text = ""

txtTipoId.Text = ""
txtTipoId.Tag = ""
txtNombre.Text = ""
txtEstado.Text = ""

txtNombreNew.Text = ""
txtEstadoNew.Text = ""

chkUsuarioValida.Visible = False

lblAutorizador(0).Visible = False
lblAutorizador(1).Visible = False

txtUsuario.Text = ""
txtUsuarioClave.Text = ""

txtUsuario.Visible = False
txtUsuarioClave.Visible = False


End Sub

Private Sub chkUsuarioValida_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If chkUsuarioValida.Value = vbChecked Then
    
    'Valida el Usuario
    If UCase(Trim(txtUsuario.Text)) = UCase(glogon.Usuario) Then
        chkUsuarioValida.Value = vbUnchecked
        MsgBox "El usuario Autorizador  no puede se igual al usuario actual: " & glogon.Usuario & ", proceda a cambiarlo"
        Exit Sub
    End If
    
   
    'Verifica que el usuario Autorizador tambien tenga acceso al cambio de Identificación
    If Derecho("cmdAplicar", Me) <> 1 Then
        chkUsuarioValida.Value = vbUnchecked
        MsgBox "El Usuario: " & Trim(txtUsuario.Text) & ", no es tiene permisos de cambio de Identificación de Personas!", vbExclamation
        Exit Sub
    End If

    'Verifica Usuario / Cifrado Actual
    strSQL = "exec spSEG_Logon '" & Trim(txtUsuario.Text) & "','" & SIFGlobal.fxStringCifrado(Trim(txtUsuarioClave.Text)) & "'"
    Call OpenRecordSet(rs, strSQL, 1)
    If Not rs!Existe = 0 Then
        chkUsuarioValida.Value = vbChecked
    Else
        chkUsuarioValida.Value = vbUnchecked
        MsgBox "Clave de Usuario incorrecta, intente de nuevo"
        txtUsuarioClave.SetFocus
    End If
    rs.Close
    
End If 'Check

Exit Sub
    
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub CmdAplicar_Click()

On Error GoTo vError


If txtCedula.Text = txtCedNew.Text Then
    MsgBox "La identificación de cambio es la misma a la actual!", vbExclamation
    Exit Sub
End If


If Len(txtCedula.Text) = 0 Or Len(txtCedNew.Text) = 0 Then
    MsgBox "Faltan datos, verifique!", vbExclamation
    Exit Sub
End If


'Actualiza el Parametro de Validacion y Luego lo Aplica
strSQL = "select LARGO_MINIMO from AFI_TIPOS_IDS Where TIPO_ID = " & cboTipoId.ItemData(cboTipoId.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    If Len(txtCedNew.Text) <> rs!Largo_Minimo Then
            MsgBox "El número de identificación nuevo no cumplen con los caracteres requeridos [" & rs!Largo_Minimo & "], verifique!", vbExclamation
            Exit Sub
    End If
End If
rs.Close


If txtNombreNew.Text <> "Idem" And chkUsuarioValida.Value = xtpUnchecked Then
    MsgBox "Este cambio implica fusionar dos Registros de Clientes, para ello necesita indicar a un usuario autorizador (Metodo Mancomunado)!", vbExclamation
    Exit Sub
End If



'----------------------------------
Dim i As Integer

i = MsgBox("Esta Seguro que desea realizar el cambio de identificación?", vbYesNo)
If i = vbNo Then
    Exit Sub
End If

If txtNombreNew.Text <> "Idem" Then
    i = MsgBox("!!!!! ESTA ACCIÓN FUSIONARÁ LOS DATOS DE AMBAS PERSONAS !!!! Está Seguro que desea realizar el cambio de identificación?", vbYesNo)
    If i = vbNo Then
        Exit Sub
    End If
End If

Me.MousePointer = vbHourglass

'****************************************************
'*  Mantener el Triger de la Tabla bien Actualizado *
'****************************************************


strSQL = "exec spAFI_Identificacion_Cambio '" & txtCedNew.Text & "','" & txtCedula.Text & "','" & glogon.Usuario & "', " & cboTipoId.ItemData(cboTipoId.ListIndex)
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Cambio de Cedula : " & txtCedula & " a " & txtCedNew & " : " & txtNombre.Text)


'If vParametros.BitacoraEspecial Then
'   Call sbgAFIBitacora("18", "Aplica Cambio de Cedula : " & txtCedula & " a " & txtCedNew, Trim(txtCedula))
'End If

Me.MousePointer = vbDefault

MsgBox "Cambio de Cédula realizado satisfactoriamente...", vbInformation

Call sbLimpia

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 1

End Sub

Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call sbLimpia

strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
       & " order by Tipo_Id"
Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub txtCedNew_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdAplicar.SetFocus
vError:
End Sub

Private Sub txtCedNew_LostFocus()
On Error GoTo vError


chkUsuarioValida.Visible = False

lblAutorizador(0).Visible = False
lblAutorizador(1).Visible = False

txtUsuario.Text = ""
txtUsuarioClave.Text = ""

txtUsuario.Visible = False
txtUsuarioClave.Visible = False


strSQL = "select S.Cedula, S.NOMBRE, S.TIPO_ID, Tip.DESCRIPCION as 'TipoId_Desc'" _
       & "     , Ep.COD_ESTADO, Ep.DESCRIPCION as 'Estado_Persona'" _
       & " from socios S " _
       & " inner join AFI_TIPOS_IDS Tip on S.TIPO_ID = Tip.TIPO_ID" _
       & " inner join AFI_ESTADOS_PERSONA Ep on S.ESTADOACTUAL = Ep.COD_ESTADO" _
       & " Where S.cedula = '" & txtCedNew.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   Call sbCboAsignaDato(cboTipoId, rs!TipoId_Desc, False, rs!Tipo_Id)
   txtNombreNew.Text = rs!Nombre
   txtEstadoNew.Text = rs!Estado_Persona

    chkUsuarioValida.Visible = True
    lblAutorizador(0).Visible = True
    lblAutorizador(1).Visible = True
    txtUsuario.Visible = True
    txtUsuarioClave.Visible = True
Else
   Call sbCboAsignaDato(cboTipoId, txtTipoId.Text, False, txtTipoId.Tag)
   txtNombreNew.Text = "Idem"
   txtEstadoNew.Text = "Idem"
End If
rs.Close

vError:

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedNew.SetFocus
End Sub

Private Sub txtCedula_LostFocus()

On Error GoTo vError

strSQL = "select S.Cedula, S.NOMBRE, S.TIPO_ID, Tip.DESCRIPCION as 'TipoId_Desc'" _
       & "     , Ep.COD_ESTADO, Ep.DESCRIPCION as 'Estado_Persona'" _
       & " from socios S " _
       & " inner join AFI_TIPOS_IDS Tip on S.TIPO_ID = Tip.TIPO_ID" _
       & " inner join AFI_ESTADOS_PERSONA Ep on S.ESTADOACTUAL = Ep.COD_ESTADO" _
       & " Where S.cedula = '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtTipoId.Text = rs!TipoId_Desc
   txtTipoId.Tag = rs!Tipo_Id
   txtNombre.Text = rs!Nombre
   txtEstado.Text = rs!Estado_Persona
Else
   txtTipoId.Text = ""
   txtTipoId.Tag = ""
   txtNombre.Text = ""
   txtEstado.Text = ""
End If
rs.Close


vError:

End Sub


Private Sub txtUsuario_Change()
chkUsuarioValida.Value = vbUnchecked
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUsuarioClave.SetFocus
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "descripcion"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Filtro = " and estado = 'A'"
        gBusquedas.Consulta = "select nombre, DESCRIPCION " _
                            & " from usuarios"
        frmBusquedas.Show vbModal
        txtUsuario.Text = gBusquedas.Resultado
        txtUsuarioClave.SetFocus
        
    End If

End Sub

Private Sub txtUsuarioClave_Change()
chkUsuarioValida.Value = vbUnchecked
End Sub
