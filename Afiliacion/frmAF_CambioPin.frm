VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_CambioPin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambio de PIN y Clave de Auto Gestión"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   1812
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   9492
      _Version        =   1441793
      _ExtentX        =   16743
      _ExtentY        =   3196
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Renovación de PIN"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "fraTipo"
      Item(1).Caption =   "Renovación de Clave"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "txtEmailValida"
      Item(1).Control(1)=   "Label2(3)"
      Begin VB.Frame fraTipo 
         BorderStyle     =   0  'None
         Height          =   1212
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   8175
         Begin VB.TextBox txtPIN 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtTicket 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   5880
            TabIndex        =   7
            Top             =   120
            Width           =   1815
         End
         Begin XtremeSuiteControls.PushButton btnPIN 
            Height          =   372
            Left            =   7680
            TabIndex        =   14
            Top             =   600
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   656
            _StockProps     =   79
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Picture         =   "frmAF_CambioPin.frx":0000
         End
         Begin XtremeSuiteControls.PushButton btnTicket 
            Height          =   372
            Left            =   7680
            TabIndex        =   3
            Top             =   120
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   656
            _StockProps     =   79
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Picture         =   "frmAF_CambioPin.frx":0719
         End
         Begin VB.Label Label2 
            Caption         =   "Número de Ticket (Talonario)"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   3120
            TabIndex        =   10
            Top             =   120
            Width           =   2652
         End
         Begin VB.Label Label2 
            Caption         =   "P.I.N. de Auto Gestión"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   3120
            TabIndex        =   9
            Top             =   600
            Width           =   2412
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtEmailValida 
         Height          =   312
         Left            =   -67480
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   6732
         _Version        =   1441793
         _ExtentX        =   11874
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirme el Email en donde se enviará la renovación de la clave:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   -67480
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   6852
      End
   End
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   612
      Left            =   7200
      TabIndex        =   1
      Top             =   3720
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   1080
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
      Appearance      =   14
      Picture         =   "frmAF_CambioPin.frx":0E32
   End
   Begin XtremeSuiteControls.PushButton btnDatosPersonales 
      Height          =   612
      Left            =   4920
      TabIndex        =   2
      Top             =   3720
      Width           =   2292
      _Version        =   1441793
      _ExtentX        =   4043
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Verifica Datos Personales"
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
      Appearance      =   14
      Picture         =   "frmAF_CambioPin.frx":1610
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2880
      TabIndex        =   11
      Top             =   480
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4800
      TabIndex        =   12
      Top             =   480
      Width           =   4812
      _Version        =   1441793
      _ExtentX        =   8488
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
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEmail 
      Height          =   312
      Left            =   2640
      TabIndex        =   13
      Top             =   3240
      Width           =   6732
      _Version        =   1441793
      _ExtentX        =   11874
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
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   3240
      Width           =   732
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_CambioPin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mTipo As String
Dim mSMTP As String, mAsunto As String

Private Sub sbLimpia()

tcMain.Item(0).Selected = True

Select Case mTipo
  Case "I" 'Impreso
    txtPin.PasswordChar = "*"
  
  Case "T" 'Talonario
    txtPin.PasswordChar = ""
  
  Case "E" 'Email
    txtPin.PasswordChar = "*"

End Select

txtTicket.Text = ""
txtPin.Text = ""

txtCedula.Text = ""
txtNombre.Text = ""
txtEmail.Text = ""


End Sub

Private Sub btnDatosPersonales_Click()
GLOBALES.gCedulaActual = txtCedula.Text
Call sbFormsCall("frmCR_VerificaDatosPersonales", vbModal, , , False, Me)

Call txtCedula_LostFocus

End Sub

Private Sub btnPIN_Click()
If Len(txtTicket.Text) <= 1 Then
    MsgBox "Suministre un número de Ticket de cambio válido!", vbExclamation
    Exit Sub
End If

If Len(txtNombre.Text) <= 1 Then
    MsgBox "No se han suministrado datos válidos de la persona!", vbExclamation
    Exit Sub
End If


If fxTicketValida(txtTicket.Text) Then
    txtPin.Text = fxGeneraPin
Else
    MsgBox "Este número de Ticket no es válido o ya ha sido utilizado por otro usuario!", vbExclamation
End If
End Sub

Private Sub btnTicket_Click()
Dim strSQL As String, vTicket As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

vTicket = fxgAFIParametro("13")
vTicket = vTicket + 1

Call Bitacora("Aplica", "Generación de Ticket para PIN de AutoGestión No.: " & vTicket)

txtTicket.Text = vTicket
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub CmdAplicar_Click()
Dim strSQL As String, vNumPin As String
Dim vDetalle As String

On Error GoTo vError



vNumPin = Trim(txtPin.Text)

If mTipo = "E" And txtEmail.Text = "" Then
    MsgBox "La persona no tiene un email válido para envío reporte del cambio!", vbInformation
    Exit Sub
End If


If Len(txtNombre.Text) <= 1 Then
    MsgBox "No se han suministrado datos válidos de la persona!", vbExclamation
    Exit Sub
End If





Me.MousePointer = vbHourglass



Select Case tcMain.Selected.Index
    Case 0 'Cambio de PIN
        If vNumPin = "" Then
          MsgBox "No se ha generado ningún PIN, Revise!", vbExclamation
          Exit Sub
        End If

        'Revisa si el Ticket fue actualizado por otro usuario
        If Not fxTicketValida(txtTicket.Text) Then
            Me.MousePointer = vbDefault
            MsgBox "El Ticket fue utilizado por otro usuario! Vuelva a Generarlo.", vbExclamation
            Exit Sub
        End If

        mAsunto = gPortal.Empresa_Name & ": Cambio de PIN de Autogestión"
        
        
        'Portal
        strSQL = "exec spPersona_PIN_WebApp " & gPortal.Empresa_Id & ",'" & txtCedula.Text & "','" & SIFGlobal.fxStringCifrado(vNumPin) & "','" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL, 1)
        
        strSQL = "update afi_parametros set valor = '" & txtTicket.Text & "' where cod_parametro = '13'"
        Call ConectionExecute(strSQL)
        
        vDetalle = "Ticket: " & txtTicket.Text & Space(15 - Len(txtTicket.Text)) & " PIN AutoGestión Renovado (Tipo:" & mTipo & ")"
        'La bitácora Especial es Obligatoria en la Generación de PINs
        Call sbgAFIBitacora("28", vDetalle, txtCedula.Text)
        
        'Notifica al correo el PIN
        If mTipo = "E" Then
            vDetalle = "<html> <body> <div class=WordSection1> <p class=MsoNormal>Se ha cambiado su PIN para el APP de autogestión. Ahora " _
                     & "puede ingresar con el PIN: <b>" & txtPin.Text & "</b></p> </div> </body> </html>"
            
            strSQL = "exec spSys_CORREO_POOL '" & vDetalle & "','" & mAsunto & "','P','" & txtEmail.Text & "'"
            Call ConectionExecute(strSQL)
        End If
        
        
        Me.MousePointer = vbDefault
        
        
        Select Case mTipo
          Case "I" 'Impreso
            Call sbImprimeBoletaPin(txtCedula)
            MsgBox "Pin de AutoGestion (Impresión de Seguridad)", vbInformation
          Case "T" 'Talonario
            MsgBox "Pin de AutoGestion Renovado satisfactoriamente (Registro en Talonario)", vbInformation
          Case "E" 'Email
            MsgBox "Pin de AutoGestion Renovado satisfactoriamente (Enviado por E-mail)", vbInformation
        End Select


    Case 1 'Clave Web/App
    
        If txtEmail.Text <> txtEmailValida.Text Then
            Me.MousePointer = vbDefault
            MsgBox "El E-Mail no concuerda con el real!", vbExclamation
            Exit Sub
        End If
        
        
        strSQL = "exec spuProGrX_MOBILE_Persona_WebKey_Renueva " & gPortal.Empresa_Id & ",'" & txtCedula.Text _
                & "','" & txtEmail.Text & "','" & glogon.Usuario & "',''"
        Call ConectionExecute(strSQL)

        If Not glogon.error Then
            MsgBox "Clave de AutoGestion Renovada satisfactoriamente (Enviada por E-mail)", vbInformation
        End If
End Select

Call sbLimpia
Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

mTipo = fxgAFIParametro("12")

Call sbLimpia

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Function fxTicketValida(vTicket As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select count(*) as 'Existe'  From AFI_BITACORA_ESPECIAL" _
       & "  where MOVIMIENTO = '28' and SUBSTRING(detalle,9,16) = '" & vTicket & Space(15 - Len(txtTicket.Text)) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    fxTicketValida = True
Else
    fxTicketValida = False
End If
rs.Close

End Function





Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdAplicar.SetFocus
vError:
End Sub

Private Sub txtCedula_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset


txtNombre.Text = ""
txtEmail.Text = ""
txtPin.Text = ""
txtTicket.Text = ""

strSQL = "select nombre,isnull(af_email,'') as 'Email' from socios where cedula = '" & Mid(Trim(txtCedula.Text), 1, 20) & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
    txtNombre.Text = Trim(rs!Nombre)
    txtEmail.Text = Trim(rs!Email)
End If

End Sub

'Genera Pin, Número aleatorio de 4 digitos
Private Function fxGeneraPin() As String
Dim vResultado As String

On Error GoTo vError

Randomize
vResultado = Format(Int((9999 * Rnd + 1)), "0000")

fxGeneraPin = vResultado

Exit Function
   
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbImprimeBoletaPin(ByVal vCedula As String)

On Error GoTo vError
    
    With frmContenedor.Crt
        .Reset
        .WindowShowGroupTree = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowTitle = "Reportes"
        .Destination = crptToPrinter
        
        .Connect = glogon.ConectRPT
        
        .ReportFileName = SIFGlobal.fxPathReportes("Personas_BoletaPinAutogestion.rpt")
        .SelectionFormula = "{SOCIOS.CEDULA} =  '" & vCedula & "'"
        .PrintReport

    End With

 Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
