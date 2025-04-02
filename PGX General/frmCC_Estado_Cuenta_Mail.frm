VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCC_Estado_Cuenta_Mail 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estado de Cuenta"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Left            =   -15
      TabIndex        =   9
      Top             =   3960
      Width           =   10590
      _Version        =   1441793
      _ExtentX        =   18680
      _ExtentY        =   2566
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnEnviar 
         Height          =   735
         Left            =   4080
         TabIndex        =   10
         Top             =   480
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Enviar Estado de Cuenta"
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
         Picture         =   "frmCC_Estado_Cuenta_Mail.frx":0000
         ImageAlignment  =   0
      End
   End
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   5
      Top             =   2640
      Width           =   5175
      _Version        =   1441793
      _ExtentX        =   9128
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Estado de Cuenta Resumen, Corte a hoy"
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
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtEmail 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   2040
      Width           =   7695
      _Version        =   1441793
      _ExtentX        =   13573
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   6
      Top             =   3120
      Width           =   5175
      _Version        =   1441793
      _ExtentX        =   9128
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Estado de Cuenta Detallado (MEIC) "
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
   End
   Begin XtremeSuiteControls.ComboBox cboCorte 
      Height          =   330
      Left            =   4080
      TabIndex        =   7
      Top             =   3480
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   8
      Top             =   3480
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Corte"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Email"
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
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   1560
      Width           =   7935
      _Version        =   1441793
      _ExtentX        =   13996
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "NOMBRE_COMPLETO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
      _Version        =   1441793
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "CEDULA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      _Version        =   1441793
      _ExtentX        =   18653
      _ExtentY        =   2566
      _StockProps     =   14
      Caption         =   "Envío de Estado de Cuenta por Correo Electrónico"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
End
Attribute VB_Name = "frmCC_Estado_Cuenta_Mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Private Sub btnEnviar_Click()


If txtEmail.Text = "" Then
  MsgBox "La persona no cuenta con un correo registrado, verifique!", vbExclamation
  Exit Sub
End If

On Error GoTo vError
  
' txtEmail.Text = "pbaltodano@mpbsystemlogic.com"
  
Select Case True
    Case rbInforme(0).Value 'Estado Resumen
        strSQL = "exec spuProGrX_MOBILE_CUENTAS_ENVIAESTADO '" & scMain(0).Caption & "'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Aplica", "Estado de Cuenta: [email] " & Trim(scMain(0).Caption))
        
        MsgBox "Estado de Cuenta enviado al Correo Electrónico registrado de la persona!", vbInformation
    
    Case rbInforme(1).Value 'MEIC por Cortes
    
      Call sbEstadoCuenta_Email_Corte(scMain(0).Caption, txtEmail.Text, cboCorte.ItemData(cboCorte.ListIndex))

End Select

Unload Me

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
On Error GoTo vError

vModulo = 10


scMain(0).Caption = GLOBALES.gTag
scMain(1).Caption = GLOBALES.gTag2

strSQL = "select rtrim(isnull(AF_Email,'')) as 'Email' from socios where cedula = '" & GLOBALES.gTag & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtEmail.Text = rs!Email
End If
strSQL = "exec spSys_Periodos_Cierre_Consulta"
Call sbCbo_Llena_New(cboCorte, strSQL, False, False)

'Remueve Fecha Actual
cboCorte.RemoveItem (0)
If cboCorte.ListCount > 1 Then
    cboCorte.Text = cboCorte.ItemData(0)
End If
Call Formularios(Me)
Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
