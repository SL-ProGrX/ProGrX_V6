VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmSYS_RA_Autorizadores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RA Expedientes: Usuarios Autorizadores"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3960
      TabIndex        =   0
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   1800
      TabIndex        =   1
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1680
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
   Begin XtremeSuiteControls.FlatEdit txtClave 
      Height          =   312
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   2040
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
      PasswordChar    =   "*"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   1800
      TabIndex        =   3
      Top             =   2400
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3625
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
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   1032
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1680
      Width           =   3972
      _Version        =   1441793
      _ExtentX        =   7006
      _ExtentY        =   1820
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   960
      Width           =   2055
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
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   4680
      TabIndex        =   12
      ToolTipText     =   "Nuevo"
      Top             =   960
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_RA_Autorizadores.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   5760
      TabIndex        =   13
      ToolTipText     =   "Editar"
      Top             =   960
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_RA_Autorizadores.frx":0632
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   6120
      TabIndex        =   14
      ToolTipText     =   "Guardar"
      Top             =   960
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_RA_Autorizadores.frx":0C2D
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   6480
      TabIndex        =   15
      ToolTipText     =   "Deshacer"
      Top             =   960
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_RA_Autorizadores.frx":135E
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   6840
      TabIndex        =   16
      ToolTipText     =   "Reporte"
      Top             =   960
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_RA_Autorizadores.frx":1A5E
      ImageAlignment  =   6
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Autorizado Id"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Notas:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   6720
      TabIndex        =   6
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Autorizadores (RA)"
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
      Height          =   492
      Left            =   1800
      TabIndex        =   5
      Top             =   240
      Width           =   4692
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmSYS_RA_Autorizadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vEdita As Boolean, vCodigo As Long, vScroll As Boolean


Private Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
'btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR", "EDICION"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
'        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub

Private Sub btnBarra_Click(Index As Integer)


Select Case Index
 Case 0  'nuevo
  txtCodigo.Text = "0"
  txtCodigo.Enabled = False
  
  Call sbLimpiaPantalla
  
  Call sbBarra_Accion("edicion")
  
  vEdita = False
  
  txtUsuario.SetFocus
  
  
  
 Case 1 'editar
      vEdita = True
      Call sbBarra_Accion("edicion")
      
      txtCodigo.Enabled = False
      txtUsuario.SetFocus
 
 Case 3 'guardar
  
    Call sbGuardar
    
    Call sbConsulta(txtCodigo.Text)
 
 Case 4 'deshacer
    Call sbLimpiaPantalla
    
    Call sbBarra_Accion("nuevo")
    
    txtCodigo.SetFocus
 

   
End Select



End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If Not IsNumeric(txtCodigo.Text) Then
    txtCodigo.Text = "0"
End If

If vScroll Then
    strSQL = "select Autorizador_Id from SYS_EXP_AUTORIZADORES"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where Autorizador_Id > " & txtCodigo.Text & " order by Autorizador_Id asc"
    Else
        If txtCodigo.Text = "0" Then
                txtCodigo.Text = "999999999999"
        End If
    
       strSQL = strSQL & " where Autorizador_Id < " & txtCodigo.Text & " order by Autorizador_Id desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Autorizador_Id
      Call sbConsulta(txtCodigo.Text)
    End If

End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
 vModulo = 10
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
 vModulo = 10
 
 Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 vEdita = True
  
    cbo.Clear
    cbo.AddItem "Activo"
    cbo.AddItem "Inactivo"
  
 Call sbLimpiaPantalla
 Call sbBarra_Accion("nuevo")

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()

vCodigo = 0

txtCodigo.Text = ""
txtCodigo.Enabled = True

txtUsuario.Text = ""

cbo.Text = "Activo"

txtClave.Text = ""
txtNotas.Text = ""

End Sub



Private Sub sbConsulta(pAutorizadoId As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from SYS_EXP_AUTORIZADORES where Autorizador_Id = " & pAutorizadoId
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbBarra_Accion("Activo")
  vEdita = True
  
  vCodigo = rs!Autorizador_Id
  txtCodigo.Text = rs!Autorizador_Id
 
 
  txtUsuario.Text = rs!Aut_Usuario
 
  If rs!Estado = "A" Then
    cbo.Text = "Activo"
  Else
    cbo.Text = "Inactivo"
  End If
       
  txtNotas = rs!Notas
  txtClave = ""


Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtUsuario.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Usuario no es válido ..."


  
If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

txtNotas.Text = fxSysCleanTxtInject(txtNotas.Text)
txtUsuario.Text = fxSysCleanTxtInject(txtUsuario.Text)


'spSYS_RA_Autorizador_Add(@AutorizaId int,  @Estado char(1), @Aut_Usuario varchar(30), @Aut_Clave varchar(200), @Notas varchar(500), @Usuario varchar(30))


If vEdita Then
    strSQL = "exec spSYS_RA_Autorizador_Add " & vCodigo & ", '" & Mid(cbo.Text, 1, 1) & "', '" & txtUsuario.Text _
            & "', '" & SIFGlobal.fxStringCifrado(txtClave.Text) _
            & "', '" & txtNotas.Text & "', '" & glogon.Usuario & "'"
  
    Call ConectionExecute(strSQL)
    Call Bitacora("Modifica", "RA: Usuario Autorizador Id: " & vCodigo)

Else
    
    strSQL = "exec spSYS_RA_Autorizador_Add 0, '" & Mid(cbo.Text, 1, 1) & "', '" & txtUsuario.Text _
            & "', '" & SIFGlobal.fxStringCifrado(txtClave.Text) _
            & "', '" & txtNotas.Text & "', '" & glogon.Usuario & "'"
    Call OpenRecordSet(rs, strSQL)
    
    txtCodigo.Text = rs!Autorizador_Id
    vCodigo = txtCodigo.Text
    
   Call Bitacora("Registra", "RA: Usuario Autorizador Id: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbBarra_Accion("activo")

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete SYS_EXP_AUTORIZADORES where nombre = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "RA: Usuario Autorizador Id: " & vCodigo)
  Call sbLimpiaPantalla
  Call sbBarra_Accion("Nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If IsNumeric(txtCodigo.Text) And vEdita Then Call sbConsulta(txtCodigo.Text)
  txtClave.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Autoriador Id"
  gBusquedas.Col2Name = "Usuario"
  gBusquedas.Col3Name = "Estado"
  gBusquedas.Columna = "AUT_USUARIO"
  gBusquedas.Orden = "AUT_USUARIO"
  gBusquedas.Consulta = "select AUTORIZADOR_ID, AUT_USUARIO, ESTADO from SYS_EXP_AUTORIZADORES"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  If IsNumeric(txtCodigo.Text) Then Call sbConsulta(txtCodigo.Text)
End If

End Sub


Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus
End Sub


Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Usuario"
  gBusquedas.Col2Name = "Nombre"
  gBusquedas.Col3Name = ""
  gBusquedas.Columna = "NOMBRE"
  gBusquedas.Orden = "NOMBRE"
  gBusquedas.Consulta = "select NOMBRE, DESCRIPCION from USUARIOS"
  gBusquedas.Filtro = " AND ESTADO = 'A'"
  frmBusquedas.Show vbModal
  txtUsuario.Text = gBusquedas.Resultado
End If
End Sub
