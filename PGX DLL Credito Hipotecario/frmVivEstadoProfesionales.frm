VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmVivEstadoProfesionales 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Suspensión de Profesionales"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox gbSuspension 
      Height          =   2172
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   10332
      _Version        =   1441793
      _ExtentX        =   18224
      _ExtentY        =   3831
      _StockProps     =   79
      Caption         =   "Datos de la Suspensión"
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.RadioButton rbSuspende 
         Height          =   252
         Index           =   0
         Left            =   5880
         TabIndex        =   18
         Top             =   1440
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Suspender?"
         BackColor       =   16777215
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
         Appearance      =   16
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   492
         Left            =   8520
         TabIndex        =   4
         Top             =   1440
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Aplicar"
         BackColor       =   -2147483633
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
         Picture         =   "frmVivEstadoProfesionales.frx":0000
      End
      Begin XtremeSuiteControls.FlatEdit txtObservaciones 
         Height          =   792
         Left            =   1800
         TabIndex        =   8
         Top             =   480
         Width           =   8172
         _Version        =   1441793
         _ExtentX        =   14414
         _ExtentY        =   1397
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   4080
         TabIndex        =   10
         Top             =   1440
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   550
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   1800
         TabIndex        =   9
         Top             =   1440
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   550
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.RadioButton rbSuspende 
         Height          =   252
         Index           =   1
         Left            =   7320
         TabIndex        =   19
         Top             =   1440
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Reactivar?"
         BackColor       =   16777215
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
         Appearance      =   16
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   960
         TabIndex        =   7
         Top             =   1440
         Width           =   792
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   3360
         TabIndex        =   6
         Top             =   1440
         Width           =   672
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   6
         Left            =   960
         TabIndex        =   5
         Top             =   480
         Width           =   672
      End
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   2160
      TabIndex        =   0
      Top             =   1680
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   672
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   1185
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   312
      Left            =   2160
      TabIndex        =   13
      Top             =   1320
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
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
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   5040
      TabIndex        =   14
      Top             =   1320
      Width           =   5652
      _Version        =   1441793
      _ExtentX        =   9970
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
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.Label lblTipo 
      Height          =   252
      Left            =   4440
      TabIndex        =   17
      Top             =   600
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "..."
      ForeColor       =   16777215
      BackColor       =   -2147483633
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   2
      Left            =   4080
      TabIndex        =   16
      Top             =   1320
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Nombre"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   0
      Left            =   960
      TabIndex        =   15
      Top             =   1320
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Identificación"
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
   Begin XtremeSuiteControls.Label lblSuspendido 
      Height          =   252
      Left            =   4440
      TabIndex        =   12
      Top             =   360
      Width           =   6132
      _Version        =   1441793
      _ExtentX        =   10816
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "..."
      ForeColor       =   16777215
      BackColor       =   -2147483633
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   12
      Left            =   720
      TabIndex        =   11
      Top             =   240
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Persona Id.:"
      ForeColor       =   16777215
      BackColor       =   -2147483633
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   3
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Estado"
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmVivEstadoProfesionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Public Sub sbConsulta_Externa_IdPersona(pIdentificacion As String)
Dim pContacto As Long

strSQL = "select idContacto from ViviendaContactos where Identificacion = '" & pIdentificacion & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
 pContacto = 0
Else
 pContacto = rs!IdContacto
End If
rs.Close

If pContacto > 0 Then
    Call sbConsulta(pContacto)
End If

End Sub


Public Sub sbConsulta_Externa_IdContacto(pContacto As Long)

Call sbConsulta(pContacto)

End Sub



Private Function fxValida() As Boolean
Dim vEstado As String
Dim vSuspendeInicio As String
Dim vSuspendeCorte As String
ReDim gParametros(1 To 6)

On Error GoTo error

fxValida = False


vEstado = Mid(cboEstado.Text, 1, 1)

If Len(Trim(txtObservaciones.Text)) = 0 Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe de ingresa una observación para la suspensión.")
    Exit Function
End If


'compruebar si el valor de value es marcado o desmarcado
Select Case True
    Case rbSuspende.Item(0).Value 'Suspende
        If dtpInicio.Value > dtpCorte.Value Then
            MsgBox "La fecha de finalización de suspensión no debe ser menor a la fecha de inicio de suspensión", vbExclamation
            dtpInicio.SetFocus
            Exit Function
        End If
        vSuspendeInicio = Format(dtpInicio.Value, "yyyy/mm/dd")
        vSuspendeCorte = Format(dtpCorte.Value, "yyyy/mm/dd")

    Case rbSuspende.Item(1).Value 'Re-Activar
        vSuspendeInicio = "1900/01/01"
        vSuspendeCorte = "1900/01/01"
End Select

gParametros(1) = txtCodigo.Text
gParametros(2) = vEstado
gParametros(3) = vSuspendeInicio
gParametros(4) = vSuspendeCorte
gParametros(5) = Trim(txtObservaciones.Text)
gParametros(6) = glogon.Usuario
    
fxValida = True
salir:
    Exit Function
error:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbSuspendeContacto()
Dim vIdGarantia As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

If Not fxValida() Then Exit Sub

strSQL = "exec spCRDVivEstadoContacto_M " & gParametros(1) & ",'" & gParametros(2) _
       & "'," & IIf(gParametros(3) = "1900/01/01", "Null", "'" & gParametros(3) & "'") _
       & "," & IIf(gParametros(4) = "1900/01/01", "Null", "'" & gParametros(4) & "'") _
       & ",'" & gParametros(5) & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

If Not glogon.error Then
                                               
    Call Bitacora("Aplica", "Hipotecario> Supensión de Persona: " & gParametros(1) & " estado:" & gParametros(2))
    
    MsgBox "Información fue registrada corretamente.", vbInformation

End If

Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbLimpiaDatos()
    
    txtIdentificacion.Text = ""
    txtNombre.Text = ""
    txtObservaciones.Text = ""
    
    cboEstado.Text = "Activo"

End Sub

Private Sub sbConsulta(pContacto As Long)
On Error GoTo vError

strSQL = "SELECT IdContacto, IdEmpresa, TipoContacto, Identificacion, Nombre" _
       & ",Case TipoProfesional  WHEN 'A' THEN 'Abogado' WHEN 'I' THEN 'Ingeniero' ELSE 'Contacto' END AS 'Profesional'" _
       & ",TipoProfesional, ESTADO" _
       & ",SuspensionInicio, SuspensionCorte, Observacion" _
       & ",dbo.fxCrd_Viv_Profesional_Suspendido(P.IdContacto) as 'SuspendeActual'" _
       & " from ViviendaContactos P" _
       & " where P.idContacto = " & pContacto

Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  
    txtCodigo.Text = CStr(rs!IdContacto)
    
    If rs!SuspendeActual = 1 Then
        lblSuspendido.Caption = "Suspendido desde: " & Format(rs!SuspensionInicio, "dd/MM/yyyy")
    Else
        lblSuspendido.Caption = ""
    End If
     
    lblTipo.Caption = rs!Profesional
     
    txtNombre.Text = rs!Nombre & ""
    txtObservaciones.Text = rs!observacion & ""
    
    txtIdentificacion.Text = rs!Identificacion & ""
    
    If rs!Estado = "A" Then
      cboEstado.Text = "Activo"
    Else
      cboEstado.Text = "Inactivo"
    End If


    dtpInicio.Value = IIf(IsNull(rs!SuspensionInicio), Null, Trim(rs!SuspensionInicio))
    If IsNull(rs!SuspensionInicio) Then
        dtpCorte.Enabled = False
    Else
        dtpCorte.Value = IIf(IsNull(rs!SuspensionCorte), Now, Trim(rs!SuspensionCorte))
        dtpCorte.Enabled = True
    End If
        
End If

Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnAplicar_Click()
    Call sbSuspendeContacto
End Sub



Private Sub dtpInicio_Click()
If IsNull(dtpInicio.Value) Then
    txtObservaciones.Text = Empty
    txtObservaciones.Enabled = False
    If cboEstado.ItemData(cboEstado.ListIndex) = 2 Then
        txtObservaciones.Enabled = True
    End If
    
    dtpCorte.Enabled = False
Else
    txtObservaciones.Enabled = True
    dtpCorte.Enabled = True
    dtpCorte.Value = Now
End If

End Sub


Private Sub Form_Load()
vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

cboEstado.Clear
cboEstado.AddItem "Activo"
cboEstado.AddItem "Inactivo"
cboEstado.Text = "Activo"

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
      txtIdentificacion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Persona Id"
  gBusquedas.Col2Name = "Identificación"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Columna = "idContacto"
  gBusquedas.Orden = "idContacto"
  gBusquedas.Consulta = "select idContacto,Identificacion,Nombre from ViviendaContactos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub


Private Sub txtCodigo_LostFocus()
If IsNumeric(txtCodigo.Text) Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtIdentificacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Identificación"
  gBusquedas.Col2Name = "Id Persona"
  gBusquedas.Columna = "Identificacion"
  gBusquedas.Orden = "Identificacion"
  gBusquedas.Consulta = "select Identificacion,idContacto,nombre from ViviendaContactos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado2
  If txtCodigo.Text <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado2))
End If
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Identificación"
  gBusquedas.Col2Name = "Id Persona"
  gBusquedas.Columna = "Identificacion"
  gBusquedas.Orden = "Identificacion"
  gBusquedas.Consulta = "select Identificacion,idContacto,nombre from ViviendaContactos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado2
  If txtCodigo.Text <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado2))
End If

End Sub
