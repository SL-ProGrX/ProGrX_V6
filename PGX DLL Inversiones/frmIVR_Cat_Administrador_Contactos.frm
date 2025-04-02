VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.ShortcutBar.v20.0.0.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmIVR_Cat_Administrador_Contactos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SGCI Contactos"
   ClientHeight    =   3384
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10752
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3384
   ScaleWidth      =   10752
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   312
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Nuevo"
      Top             =   40
      Width           =   1092
      _Version        =   1310720
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmIVR_Cat_Administrador_Contactos.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   312
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Editar"
      Top             =   40
      Width           =   372
      _Version        =   1310720
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmIVR_Cat_Administrador_Contactos.frx":0632
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   312
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Eliminar"
      Top             =   40
      Width           =   372
      _Version        =   1310720
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmIVR_Cat_Administrador_Contactos.frx":0C2D
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   312
      Index           =   3
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "Guardar"
      Top             =   40
      Width           =   372
      _Version        =   1310720
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmIVR_Cat_Administrador_Contactos.frx":11D1
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   312
      Index           =   4
      Left            =   2520
      TabIndex        =   4
      ToolTipText     =   "Deshacer"
      Top             =   40
      Width           =   372
      _Version        =   1310720
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmIVR_Cat_Administrador_Contactos.frx":1902
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   312
      Index           =   5
      Left            =   3000
      TabIndex        =   5
      ToolTipText     =   "Reporte"
      Top             =   36
      Width           =   372
      _Version        =   1310720
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmIVR_Cat_Administrador_Contactos.frx":2002
      ImageAlignment  =   6
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   10080
      TabIndex        =   7
      Top             =   720
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1560
      TabIndex        =   8
      Top             =   720
      Width           =   1572
      _Version        =   1310720
      _ExtentX        =   2773
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
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3120
      TabIndex        =   9
      Top             =   720
      Width           =   6852
      _Version        =   1310720
      _ExtentX        =   12086
      _ExtentY        =   550
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   1452
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   10332
      _Version        =   1310720
      _ExtentX        =   18224
      _ExtentY        =   2561
      _StockProps     =   79
      Caption         =   "Información de Contacto"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   312
         Left            =   4920
         TabIndex        =   12
         Top             =   360
         Width           =   5292
         _Version        =   1310720
         _ExtentX        =   9334
         _ExtentY        =   550
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail2 
         Height          =   312
         Left            =   4920
         TabIndex        =   13
         Top             =   720
         Width           =   5292
         _Version        =   1310720
         _ExtentX        =   9334
         _ExtentY        =   550
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono1 
         Height          =   312
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono2 
         Height          =   312
         Left            =   1320
         TabIndex        =   15
         Top             =   720
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   1
         Left            =   0
         TabIndex        =   19
         Top             =   360
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Movil"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   2
         Left            =   0
         TabIndex        =   18
         Top             =   720
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Teléfono"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   5
         Left            =   3600
         TabIndex        =   17
         Top             =   360
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Email (1)"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   6
         Left            =   3600
         TabIndex        =   16
         Top             =   720
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Email (2)"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   372
      Left            =   0
      TabIndex        =   20
      Top             =   3000
      Width           =   10812
      _Version        =   1310720
      _ExtentX        =   19071
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Administrador"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   6
      Alignment       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Contacto Id:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   1452
   End
   Begin XtremeShortcutBar.ShortcutCaption scBarra 
      Height          =   372
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11772
      _Version        =   1310720
      _ExtentX        =   20764
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   6
      Alignment       =   1
   End
End
Attribute VB_Name = "frmIVR_Cat_Administrador_Contactos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vScroll As Boolean

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, vPaso As Boolean

Public Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub


Private Sub btnBarra_Click(Index As Integer)

Select Case Index
    Case 0 'NUEVO
        vEdita = False
        Call sbLimpiaPantalla
        txtNombre.SetFocus

        Call sbBarra_Accion("Editar")
        
    Case 1 'MODIFICAR", "EDITAR"
      If vCodigo = "" Then
        MsgBox "Consulte un Contacto primero!", vbInformation
      Else
        vEdita = True
        txtNombre.SetFocus
        Call sbBarra_Accion("Editar")
      End If
      
    Case 2 'BORRAR"
      Call sbBorrar
      Call sbBarra_Accion("Nuevo")
    
    Case 3 'GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case 4 'DESHACER"
      Call sbBarra_Accion("Editar")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbBarra_Accion("Nuevo")
        vEdita = True
      End If
    
    Case 5 'REPORTES
   
End Select

End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 COD_CONTACTO from IVR_CONTACTOS"
           
    If txtCodigo.Text = "" Then
       txtCodigo.Text = "0"
    End If
           
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COD_ADMINISTRADOR = '" & scTitulo.Tag & "' AND COD_CONTACTO > " & txtCodigo.Text & " order by COD_CONTACTO asc"
    Else
       strSQL = strSQL & " where COD_ADMINISTRADOR = '" & scTitulo.Tag & "' AND COD_CONTACTO < " & txtCodigo.Text & " order by COD_CONTACTO desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!COD_CONTACTO)
    End If

End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Load()

scTitulo.Tag = GLOBALES.gTag
scTitulo.Caption = GLOBALES.gTag2


End Sub



Private Sub sbLimpiaPantalla()

vCodigo = 0
txtCodigo.Text = ""

txtNombre.Text = ""

txtEmail.Text = ""
txtEmail2.Text = ""

txtTelefono1.Text = ""
txtTelefono2.Text = ""

End Sub


Private Sub sbConsulta(pCodigo As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select *" _
       & " from IVR_CONTACTOS" _
       & " Where COD_ADMINISTRADOR = '" & scTitulo.Tag & "' and COD_CONTACTO = " & pCodigo

Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
    
    Call sbBarra_Accion("activo")
    
    vEdita = True
    
    vCodigo = rs!COD_CONTACTO
    txtCodigo.Text = CStr(rs!COD_CONTACTO)
  
 
    txtNombre.Text = rs!Nombre & ""
        
    txtEmail.Text = rs!Email_01 & ""
    txtEmail2.Text = rs!Email_02 & ""
    
    txtTelefono1.Text = rs!Celular & ""
    txtTelefono2.Text = rs!telefono & ""
Else
  
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Me.MousePointer = vbDefault

'Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Contacto no es válido ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()

On Error GoTo vError


If txtCodigo.Text = "" Then
   
    strSQL = "select isnull(max(cod_Contacto),0) + 1 as 'ContactoID'" _
           & " from IVR_CONTACTOS"
    Call OpenRecordSet(rs, strSQL)
    
    txtCodigo.Text = rs!ContactoID
    vCodigo = txtCodigo.Text

   strSQL = "insert into IVR_CONTACTOS(COD_CONTACTO, COD_ADMINISTRADOR, Nombre" _
          & ", Celular, telefono, email_01, email_02, registro_usuario, registro_fecha)" _
          & " values(" & vCodigo & ",'" & scTitulo.Tag & "','" & Trim(txtNombre.Text) _
          & "','" & txtTelefono1 & "','" & txtTelefono2 _
          & "','" & txtEmail.Text & "','" & txtEmail2.Text _
          & "','" & glogon.Usuario & "', dbo.mygetdate())"
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Contacto:  " & vCodigo)

Else
   
    
  strSQL = "update IVR_CONTACTOS set Nombre = '" & Trim(txtNombre.Text) _
         & "', email_01 = '" & txtEmail.Text & "',  Celular = '" & txtTelefono1.Text _
         & "', email_02 = '" & txtEmail2.Text & "', telefono= '" & txtTelefono2.Text _
         & "'  where COD_CONTACTO = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Contacto:  " & vCodigo)
 
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbConsulta(vCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete IVR_CONTACTOS where COD_CONTACTO = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Contacto:  " & vCodigo)
  Call sbLimpiaPantalla
 
  Call sbBarra_Accion("NUEVO")
  
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Id. Contacto"
  gBusquedas.Col2Name = "Nombre"
  gBusquedas.Col3Name = ""
  gBusquedas.Columna = "Nombre"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Consulta = "select Cod_Contacto,Nombre from IVR_CONTACTOS"
  gBusquedas.Filtro = " AND COD_ADMINISTRADOR = '" & scTitulo.Tag & "'"
  frmBusquedas.Show vbModal

  If gBusquedas.Resultado <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail2.SetFocus
End Sub

Private Sub txtEmail2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono1.SetFocus
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono1.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Id. Contacto"
  gBusquedas.Col2Name = "Nombre"
  gBusquedas.Col3Name = ""
  gBusquedas.Columna = "Nombre"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Consulta = "select Cod_Contacto,Nombre from IVR_CONTACTOS"
  gBusquedas.Filtro = " AND COD_ADMINISTRADOR = '" & scTitulo.Tag & "'"
  frmBusquedas.Show vbModal

  If gBusquedas.Resultado <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtTelefono1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono2.SetFocus
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub


