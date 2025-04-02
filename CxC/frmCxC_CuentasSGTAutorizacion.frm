VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCxC_CuentasSGTAutorizacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Autorización (Resolución) de Operaciones"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.RadioButton optAutoriza 
      Height          =   372
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   5880
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Autorizar"
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
      Appearance      =   16
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   615
      Left            =   6840
      TabIndex        =   4
      Top             =   5880
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "&Aplicar"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCxC_CuentasSGTAutorizacion.frx":0000
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   4332
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   8412
      _Version        =   1572864
      _ExtentX        =   14838
      _ExtentY        =   7641
      _StockProps     =   79
      Caption         =   "Autorizaciones"
      ForeColor       =   8421504
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   435
         Left            =   1560
         TabIndex        =   10
         Top             =   480
         Width           =   1935
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   762
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   1935
         Left            =   1560
         TabIndex        =   8
         Top             =   960
         Width           =   6855
         _Version        =   1572864
         _ExtentX        =   12091
         _ExtentY        =   3413
         _StockProps     =   77
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1335
         Left            =   1560
         TabIndex        =   9
         Top             =   3000
         Width           =   6855
         _Version        =   1572864
         _ExtentX        =   12091
         _ExtentY        =   2355
         _StockProps     =   77
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
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
         Height          =   312
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Width           =   1332
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Operación"
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
         Height          =   312
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1332
      End
   End
   Begin XtremeSuiteControls.RadioButton optAutoriza 
      Height          =   372
      Index           =   1
      Left            =   3120
      TabIndex        =   6
      Top             =   5880
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Denegar"
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton optAutoriza 
      Height          =   372
      Index           =   2
      Left            =   4560
      TabIndex        =   7
      Top             =   5880
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Desautorizar"
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
      Appearance      =   16
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Autorización para Activación de Cuenta"
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
      Height          =   480
      Index           =   2
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   6852
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmCxC_CuentasSGTAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As Boolean

vResultado = True

fxValida = vResultado


End Function


Private Sub sbAplicar()
Dim strSQL As String, vEstado As String

On Error GoTo vError

If txtDetalle.Tag = "N" Then
    MsgBox "Es solicitud no ha pasado la validación, verifique...", vbExclamation
    Exit Sub
End If


If Len(Trim(txtNotas.Text)) = 0 Then
    MsgBox "No se ha especificado ninguna nota para la autorización, verifique...", vbExclamation
    Exit Sub
End If


'If Not fxValida Then
'   MsgBox "...."
'   Exit Sub
'End If

Select Case True
  Case optAutoriza.Item(0).Value
        vEstado = "A"
  Case optAutoriza.Item(1).Value
        vEstado = "D"
  Case optAutoriza.Item(2).Value
        vEstado = "R"
End Select




If vEstado = "R" Then
    strSQL = "update CxC_Cuentas set Autoriza_Usuario = '" & glogon.Usuario & "',Autoriza_fecha = Null" _
           & ",Autoriza_notas = 'Reversada la Autorización', Autoriza_Estado = '" & vEstado _
           & "' where Operacion = " & txtOperacion
Else
    strSQL = "update CxC_Cuentas set Autoriza_Usuario = '" & glogon.Usuario & "',Autoriza_fecha = Null" _
           & ",Autoriza_notas = '" & txtNotas.Text & "', Autoriza_Estado = '" & vEstado _
           & "' where Operacion = " & txtOperacion
End If
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Resolución de la Operación: " & txtOperacion.Text & " -> Estado :" & vEstado)

MsgBox "Operación Resolucionada Satisfactoriamente...", vbInformation

Unload Me

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub



Private Sub btnAplicar_Click()
Call sbAplicar
End Sub

Private Sub Form_Activate()
vModulo = 31
End Sub

Private Sub Form_Load()

vModulo = 31

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

txtOperacion.Text = GLOBALES.gTag
Call txtOperacion_Change
Call sbConsulta

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub



Private Sub optAutoriza_Click(Index As Integer)
Select Case Index
  Case 0
     txtNotas.Text = "La Operación cumple con todos los requisitos."
  Case 1
     txtNotas.Text = "La Operación NO cumple con todos los requisitos y disposiciones internas."
  Case 2
     txtNotas.Text = "La Operación se resolucionó equivocadamente!"

End Select
End Sub


Private Sub txtOperacion_Change()
    txtDetalle.Text = ""
    txtDetalle.Tag = "N"
    txtNotas.Text = ""
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbConsulta

Exit Sub

vError:

End Sub


Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

txtDetalle.Text = ""
txtDetalle.Tag = "N"
txtNotas.Text = ""


strSQL = "Select R.Operacion,R.cod_concepto,R.cedula,S.nombre,R.Monto,R.Dias_plazo,R.Tasa_Corriente, R.cuota,R.cod_Contrato" _
       & ",D.descripcion as ContratoDesc,C.descripcion as ConceptoDesc,R.Registro_Usuario,R.Registro_Fecha,R.Notas" _
       & " from CxC_Cuentas R inner join CxC_Personas S on R.cedula = S.cedula" _
       & " inner join CxC_Conceptos C on R.cod_concepto = C.cod_concepto" _
       & " left join CxC_Contratos D on R.cod_Contrato = D.cod_Contrato" _
       & " where R.Autoriza_Fecha is null and R.Estado = 'R' and R.Operacion = " & txtOperacion.Text

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then

  txtDetalle = txtDetalle & vbCrLf & "Operación   : " & vbTab & rs!Operacion
  txtDetalle = txtDetalle & vbCrLf & "Concepto    : " & vbTab & rs!cod_Concepto & " - " & rs!ConceptoDesc
  txtDetalle = txtDetalle & vbCrLf & "Contrato    : " & vbTab & rs!COD_CONTRATO & " - " & rs!ContratoDesc
  txtDetalle = txtDetalle & vbCrLf & "Cédula      : " & vbTab & Trim(rs!Cedula) & " - " & rs!Nombre & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & "Monto       : " & vbTab & Format(rs!Monto, "Standard")
  txtDetalle = txtDetalle & vbCrLf & "Plazo       : " & vbTab & rs!Dias_Plazo & "(Días)"
  txtDetalle = txtDetalle & vbCrLf & "Tasa        : " & vbTab & rs!Tasa_Corriente
  txtDetalle = txtDetalle & vbCrLf & "Cuota       : " & vbTab & Format(rs!Cuota, "Standard") & vbCrLf

  txtDetalle = txtDetalle & vbCrLf & "Fecha       : " & vbTab & Format(rs!Registro_Fecha, "dd/mm/yyyy")
  txtDetalle = txtDetalle & vbCrLf & "Usuario     : " & vbTab & rs!Registro_Usuario & vbCrLf

  txtDetalle = txtDetalle & vbCrLf & "Notas : " & rs!Notas & ""

  txtDetalle.Tag = "S"
  txtOperacion.Tag = rs!cod_Concepto
  
  
    'Validaciones Finales> Consolida Varias Del disponible y Contabilizacion
    strSQL = "select dbo.fxCxC_Persona_Disponible_Valida('" & Trim(rs!Cedula) & "', " & CCur(rs!Monto) _
           & ", '" & rs!cod_Concepto & "') as 'Resultado'"
    Call OpenRecordSet(rs, strSQL)
    If Len(rs!Resultado) > 0 Then
      txtDetalle.Tag = "N"
      txtDetalle.Text = txtDetalle.Text & vbCrLf & rs!Resultado
    End If
    rs.Close
    
    'Verifica que no Existan Facturas Duplicadas
    strSQL = "exec spCxC_Operacion_Facturas_Verifica " & txtOperacion.Text
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      txtDetalle.Tag = "N"
      txtDetalle.Text = txtDetalle.Text & vbCrLf & "- Factura No.: " & Trim(rs!cod_Factura) _
                      & ", se encuentra registrada en la Operación: " & rs!Operacion
      rs.MoveNext
    Loop
   
   
End If
rs.Close

Me.MousePointer = vbDefault

If txtDetalle.Tag = "N" Then
   txtDetalle.ForeColor = vbRed
   MsgBox " La Solicitud no cumple con alguno(s) de los siguientes parámetros:" _
          & vbCrLf & " 1. No se encuentra recibida" & vbCrLf & " 2. No Existe la Operación" _
          & vbCrLf & " 3. Ya se encuentra Resolucionada!", vbExclamation
Else
   txtDetalle.ForeColor = vbBlue
End If

Call optAutoriza_Click(0)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 Call txtOperacion_Change

End Sub

