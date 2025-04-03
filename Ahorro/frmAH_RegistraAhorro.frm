VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAH_RegistraAhorro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Ahorro Extraordinario"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   HelpContextID   =   2006
   Icon            =   "frmAH_RegistraAhorro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   5655
      Begin VB.TextBox txtAno 
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin MSMask.MaskEdBox medMonto 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Monto que desea ahorrar"
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "###,###,###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox chkindefinido 
         Caption         =   "Devolucion Indefinida"
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmAH_RegistraAhorro.frx":030A
         Left            =   1200
         List            =   "frmAH_RegistraAhorro.frx":0335
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Monto"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Fec.Dev."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   4560
      Picture         =   "frmAH_RegistraAhorro.frx":039E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancela el registro extraordinario y vuelve a la ventana principal de ahorros"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerarRecibo 
      Caption         =   "&Aplicar"
      Height          =   855
      Left            =   3480
      Picture         =   "frmAH_RegistraAhorro.frx":06A8
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Genera el recibo para validar el movimiento"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmAH_RegistraAhorro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strfechadev As String
Sub habilita()
       medMonto.Enabled = True
       cboMes.Enabled = True
       txtAno.Enabled = True
       chkindefinido = True
       
End Sub
Sub deshabilita()
       medMonto.Enabled = False
       cboMes.Enabled = False
       txtAno.Enabled = False
       chkindefinido.Enabled = False
End Sub
Private Sub chkindefinido_Click()
 
    If chkindefinido.Value = 0 Then
       cboMes.Enabled = True
       txtAno.Enabled = True
    End If
    
    If chkindefinido.Value = 1 Then
       cboMes.Enabled = False
       txtAno.Enabled = False
       txtAno.Text = ""
       cboMes.ListIndex = -1
       strfechadev = "999999"
       
     End If
 End Sub

Private Sub cmdCancelar_Click()
'Este formulario no se descarga solo se oculta
'cuando regresa al formulario principal de ahorros
'este ultimo evalua el propiedad tag del boton de recibo y
'luego lo descarga
  Unload frmAH_RegistraAhorro
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PROCEDIMIENTO cmdGenerarRecibo_click
'DESCRIPCION
'Funcion que registra los ahorros extraordinarios de un socio.
'los cuales pueden ser indefinidos o pueden tener una fecha de devolucion
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdGenerarRecibo_Click()
' se debe generar el recibo para que este devuelva el # del comprobante
' y active el registro del ahorro

On Error GoTo CapturaError

cmdGenerarRecibo.Tag = "1"

Dim strFechapro As String
Dim strFechasis As String, strFpromes As String, strFproanio As String
Dim strSQL As String, strCedu As String, strFdev As String, strFecha As String
Dim de As Integer, hasta As Integer, strfpro As String

Dim recRegs As New ADODB.Recordset

strSQL = "select * from PAR_AHCR"

recRegs.ActiveConnection = glogon.Conection
recRegs.Source = strSQL
recRegs.Open

If Not recRegs.EOF And Not IsNull(recRegs!ah_fec) Then
 strfpro = Trim(recRegs!ah_fec)
Else
 MsgBox "No existe fecha de proceso. Ingresela", 0
 recRegs.Close
 Exit Sub
End If
recRegs.Close

If (medMonto.Text <> "" And cboMes.ListIndex <> -1 And txtAno.Text <> "") Or (medMonto.Text <> "" And cboMes.ListIndex = -1 And Trim(txtAno.Text) = "" And chkindefinido = 1) Then
'calculo automaticamente la fecha de y la fecha hasta a partir de la fecha de proceso.
    If chkindefinido = 0 Then
        strFpromes = Trim(str(cboMes.ItemData(cboMes.ListIndex)))
        strFproanio = Trim(txtAno.Text)
        Select Case strFpromes
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9
                 strFpromes = "0" & Trim(strFpromes)
            Case 10, 11, 12
                 strFpromes = strFpromes
        End Select
        strFdev = strFproanio & strFpromes
    End If

    strCedu = frmAH_Principal.txtCedula.Text
    
    strSQL = "insert into ahorro_detallado(cedula,tipo,monto,fecha,fechaproc,estado,fecdev)"
    strSQL = strSQL & "values('" & strCedu & "'," & "'E'," & medMonto.Text & ",'"
    
    If chkindefinido = 1 Then 'fecha de devolucion
      strSQL = strSQL & Format(Date, "yyyy/mm/dd") & "'," & strfpro & "," & "'A'," & "999999)"
    Else
      strSQL = strSQL & Format(Date, "yyyy/mm/dd") & "'," & strfpro & "," & "'A'," & strFdev & ")"
    End If
    
    glogon.Conection.Execute strSQL
        
    strSQL = "UPDATE ahorro_consolidado set"
    strSQL = strSQL & " extra= extra + " & medMonto.Text
    strSQL = strSQL & ",fecextra='" & Format(Date, "yyyy/mm/dd") & "'"
    strSQL = strSQL & " where cedula='" & strCedu & "'"
    
    glogon.Conection.Execute strSQL
    
    Call Bitacora("Registra", ("Ahor.Extra. Ced:" & frmAH_Principal.txtCedula & " Mto: " & medMonto.Text))
    
    
    'ASIENTO AQUI ********************************* SIN HACER ********************
    
    cmdGenerarRecibo.Enabled = False
    deshabilita
    
Else
    MsgBox "ingrese datos", 0
End If
Exit Sub

CapturaError:
 Call ProcedimientoErrores(Me.Name, Err)

End Sub

Private Sub Form_DblClick()
Set Conlsw.frmX = Me
Conlsw.ImprimeForm
End Sub

Private Sub Form_Load()
lblNombre.Caption = Trim(frmAH_Principal.txtNombre)
vModulo = 2
Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

Private Sub medMonto_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
   If IsNumeric(medMonto.Text) Then
       cboMes.SetFocus
   End If
 End If
End Sub
