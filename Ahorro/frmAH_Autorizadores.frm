VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAH_Autorizadores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patrimonio: Autorizadores"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   8325
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
      Height          =   315
      Left            =   1800
      TabIndex        =   1
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
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
      TabIndex        =   3
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
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   4680
      TabIndex        =   4
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
      Picture         =   "frmAH_Autorizadores.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   5760
      TabIndex        =   5
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
      Picture         =   "frmAH_Autorizadores.frx":0632
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   6120
      TabIndex        =   6
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
      Picture         =   "frmAH_Autorizadores.frx":0C2D
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   6480
      TabIndex        =   7
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
      Picture         =   "frmAH_Autorizadores.frx":135E
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   6840
      TabIndex        =   8
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
      Picture         =   "frmAH_Autorizadores.frx":1A5E
      ImageAlignment  =   6
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Autorizadores (PAT)"
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
      TabIndex        =   12
      Top             =   240
      Width           =   4692
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
      TabIndex        =   11
      Top             =   1440
      Width           =   1212
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
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
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
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmAH_Autorizadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vEdita As Boolean, vCodigo As String, vScroll As Boolean


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
  
  Call sbLimpiaPantalla
  
  Call sbBarra_Accion("edicion")
  
  vEdita = False
  
  txtUsuario.SetFocus
  
  
  
 Case 1 'editar
      vEdita = True
      Call sbBarra_Accion("edicion")
      txtUsuario.SetFocus
 
 Case 3 'guardar
  
    Call sbGuardar
    
    Call sbConsulta(txtUsuario.Text)
 
 Case 4 'deshacer
    Call sbLimpiaPantalla
    
    Call sbBarra_Accion("nuevo")
    
    txtUsuario.SetFocus
 

   
End Select



End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Usuario from PAT_AUTORIZADORES"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where Usuario > '" & txtUsuario.Text & "' order by Usuario asc"
    Else
       strSQL = strSQL & " where Usuario < '" & txtUsuario.Text & "' order by Usuario desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtUsuario.Text = rs!Usuario
      Call sbConsulta(txtUsuario.Text)
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
 vModulo = 2
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
 vModulo = 2
 
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

vCodigo = ""

txtUsuario.Text = ""
cbo.Text = "Activo"
txtNotas.Text = ""

End Sub



Private Sub sbConsulta(pAutorizadoId As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select USUARIO, ESTADO, NOTAS from PAT_AUTORIZADORES where Usuario = '" & pAutorizadoId & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbBarra_Accion("Activo")
  vEdita = True
  
  vCodigo = rs!Usuario
  txtUsuario.Text = rs!Usuario
 
  If rs!Estado = "A" Then
    cbo.Text = "Activo"
  Else
    cbo.Text = "Inactivo"
  End If
       
  txtNotas = rs!Notas


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

On Error GoTo vError

txtNotas.Text = fxSysCleanTxtInject(txtNotas.Text)
txtUsuario.Text = fxSysCleanTxtInject(txtUsuario.Text)


strSQL = "exec spPAT_Autorizador_Add '" & txtUsuario.Text & "', '" & Mid(cbo.Text, 1, 1) _
        & "', '" & txtNotas.Text & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

If vEdita Then
    Call Bitacora("Modifica", "PAT: Usuario Autorizador: " & txtUsuario.Text)
Else
    Call Bitacora("Registra", "PAT: Usuario Autorizador: " & txtUsuario.Text)
End If

vCodigo = txtUsuario.Text

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbBarra_Accion("activo")

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete PAT_AUTORIZADORES where Usuario = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "PAT: Usuario Autorizador: " & vCodigo)
  Call sbLimpiaPantalla
  Call sbBarra_Accion("Nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
End Sub


Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Call sbConsulta(txtUsuario.Text)
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Usuario"
  gBusquedas.Col2Name = "Estado"
  gBusquedas.Columna = "USUARIO"
  gBusquedas.Orden = "USUARIO"
  gBusquedas.Consulta = "select USUARIO, ESTADO from PAT_AUTORIZADORES"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtUsuario.Text = gBusquedas.Resultado
  Call sbConsulta(txtUsuario.Text)
End If
End Sub


