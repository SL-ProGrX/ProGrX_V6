VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmPGX_Servicios_Asignados 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Servicios Asignados"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.CheckBox chkAplicaUsuario 
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   3000
      Width           =   3495
      _Version        =   1441792
      _ExtentX        =   6165
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Aplica Costo por Cantidad de Usuarios"
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
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   13
      Top             =   3600
      Width           =   495
      _Version        =   1441792
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
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
      Picture         =   "frmPGX_Servicios_Asignados.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtQtyUsers 
      Height          =   330
      Left            =   4320
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
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
   Begin XtremeSuiteControls.FlatEdit txtServicioDesc 
      Height          =   330
      Left            =   3000
      TabIndex        =   6
      Top             =   2520
      Width           =   6135
      _Version        =   1441792
      _ExtentX        =   10821
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtServicioCod 
      Height          =   330
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9240
      TabIndex        =   0
      Top             =   2520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   330
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
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
      Left            =   3000
      TabIndex        =   3
      Top             =   2040
      Width           =   6135
      _Version        =   1441792
      _ExtentX        =   10821
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   14
      Top             =   3600
      Width           =   495
      _Version        =   1441792
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
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
      Picture         =   "frmPGX_Servicios_Asignados.frx":0720
   End
   Begin XtremeSuiteControls.FlatEdit txtFecha 
      Height          =   330
      Left            =   5040
      TabIndex        =   17
      Top             =   1560
      Width           =   2055
      _Version        =   1441792
      _ExtentX        =   3625
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
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   330
      Left            =   7080
      TabIndex        =   18
      Top             =   1560
      Width           =   2055
      _Version        =   1441792
      _ExtentX        =   3625
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
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   330
      Left            =   1440
      TabIndex        =   12
      Top             =   3600
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtCosto 
      Height          =   330
      Left            =   1440
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   330
      Left            =   3000
      TabIndex        =   21
      Top             =   1560
      Width           =   2055
      _Version        =   1441792
      _ExtentX        =   3625
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
      BackStyle       =   0  'Transparent
      Caption         =   "Suscribir Servicios"
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
      Height          =   480
      Index           =   0
      Left            =   1440
      TabIndex        =   20
      Top             =   240
      Width           =   4452
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   6
      Left            =   7080
      TabIndex        =   16
      Top             =   1320
      Width           =   975
      _Version        =   1441792
      _ExtentX        =   1720
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   15
      Top             =   1320
      Width           =   975
      _Version        =   1441792
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Fecha"
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   495
      Index           =   4
      Left            =   480
      TabIndex        =   11
      Top             =   3480
      Width           =   975
      _Version        =   1441792
      _ExtentX        =   1720
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Monto / Mensual"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   9
      Top             =   3000
      Width           =   975
      _Version        =   1441792
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Qty Users"
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   975
      _Version        =   1441792
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Costo Ud"
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   975
      _Version        =   1441792
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Servicio"
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   975
      _Version        =   1441792
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cliente"
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
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11772
   End
End
Attribute VB_Name = "frmPGX_Servicios_Asignados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean, vPaso As Boolean

Private Sub btnAccion_Click(Index As Integer)
Dim strSQL As String, vMensaje As String

On Error GoTo vError

Select Case Index
  Case 0 'nuevo
        vMensaje = "Servicios Registrados Satisfactoriamente!"
        strSQL = "exec spPGX_Servicio_Asigna " & txtCodigo.Text & ",'" & txtServicioCod.Text & "'," & CCur(txtMonto.Text) _
               & "," & CCur(txtCosto.Text) & "," & CInt(txtQtyUsers.Text) & "," & chkAplicaUsuario.Value _
               & ",'" & glogon.Usuario & "','M'"
  Case 1 'borrar
        vMensaje = "Servicios Eliminados!"
        strSQL = "exec spPGX_Servicio_Asigna " & txtCodigo.Text & ",'" & txtServicioCod.Text & "'," & CCur(txtMonto.Text) _
               & "," & CCur(txtCosto.Text) & "," & CInt(txtQtyUsers.Text) & "," & chkAplicaUsuario.Value _
               & ",'" & glogon.Usuario & "','E'"

End Select

Call ConectionExecute(strSQL)

MsgBox vMensaje, vbInformation

Call sbConsulta(txtServicioCod.Text)

Exit Sub

vError:

   MsgBox Err.Description, vbCritical
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 S.cod_Servicio" _
           & " from PGX_Servicios S left join PGX_Servicios_Asg A on S.cod_Servicio = A.cod_Servicio" _
           & " and A.cod_Empresa = " & txtCodigo.Text
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where S.cod_Servicio > '" & txtServicioCod.Text & "' order by S.cod_Servicio asc"
    Else
       strSQL = strSQL & " where S.cod_Servicio < '" & txtServicioCod.Text & "' order by S.cod_Servicio desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtServicioCod.Text = rs!cod_Servicio
      Call txtServicioCod_LostFocus
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox Err.Description, vbCritical


End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

txtCodigo.Text = GLOBALES.gTag
txtNombre.Text = GLOBALES.gTag2


End Sub

Private Sub sbLimpia()

txtQtyUsers.Text = 4
txtMonto.Text = 0
txtCosto.Text = 0
txtFecha.Text = ""
txtUsuario.Text = ""
chkAplicaUsuario.Value = vbUnchecked

End Sub

Private Sub sbConsulta(pCodigo As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Call sbLimpia

strSQL = "select S.Cod_Servicio,S.Descripcion,A.Monto,A.Costo,A.Cantidad_Usuarios,A.Registro_Fecha,A.Registro_Usuario" _
    & " , S.Aplica_Por_Usuario,S.Costo as 'CostoServicio',A.Activo" _
    & " from PGX_Servicios S left join PGX_Servicios_ASG A on S.cod_Servicio = A.Cod_Servicio" _
    & " and A.cod_Empresa = " & txtCodigo.Text _
    & " where S.cod_Servicio = '" & pCodigo & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
     
    txtServicioCod.Text = rs!cod_Servicio
    txtServicioDesc.Text = rs!Descripcion
    chkAplicaUsuario.Value = rs!Aplica_por_Usuario
    
    If Not IsNull(rs!Monto) And rs!Activo = 1 Then
        txtCosto.Text = Format(rs!Costo, "Standard")
        txtQtyUsers.Text = rs!Cantidad_Usuarios
        txtMonto.Text = Format(rs!Monto, "Standard")
        
        txtFecha.Text = rs!Registro_Fecha & ""
        txtUsuario.Text = rs!Registro_usuario & ""
        
        If rs!Activo = 1 Then
            txtEstado.Text = "Activa"
        Else
            txtEstado.Text = "Eliminada"
        End If
        
    Else
        txtCosto.Text = Format(rs!CostoServicio, "Standard")
        txtQtyUsers.Text = 1
        txtMonto.Text = Format(rs!CostoServicio, "Standard")
        txtEstado.Text = "Sin Asignar"
    End If
End If
rs.Close

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub



Private Sub txtQtyUsers_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If chkAplicaUsuario.Value = vbChecked Then
    txtMonto.Text = Format(CInt(txtQtyUsers.Text) * CCur(txtCosto.Text), "Standard")
Else
    txtMonto.Text = txtCosto.Text
End If

Exit Sub

vError:
    txtMonto.Text = txtCosto.Text

End Sub

Private Sub txtServicioCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtServicioDesc.SetFocus

If KeyCode = vbKeyF4 Then
 gBusquedas.Col1Name = "Id Servicio"
 gBusquedas.Col2Name = "Descripción"
 gBusquedas.Columna = "cod_servicio"
 gBusquedas.Orden = "cod_servicio"
 gBusquedas.Consulta = "select cod_Servicio, descripcion from PGX_Servicios"
 gBusquedas.Filtro = " and activo = 1"
 frmBusquedas.Show vbModal
 
 If gBusquedas.Resultado <> "" Then
    txtServicioCod.Text = gBusquedas.Resultado
    txtServicioCod_LostFocus
 End If
 
 
End If

End Sub

Private Sub txtServicioCod_LostFocus()
Call sbConsulta(txtServicioCod.Text)
End Sub


