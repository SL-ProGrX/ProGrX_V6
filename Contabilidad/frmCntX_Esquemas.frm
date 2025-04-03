VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCntX_Esquemas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copia estructura contable de una Contabilidad a Otra!"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9120
   HelpContextID   =   1
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1092
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   8652
      _Version        =   1310723
      _ExtentX        =   15261
      _ExtentY        =   1926
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   615
         Left            =   6960
         TabIndex        =   7
         Top             =   360
         Width           =   1575
         _Version        =   1310723
         _ExtentX        =   2778
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCntX_Esquemas.frx":0000
      End
      Begin VB.Label lblX 
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   5532
      End
   End
   Begin XtremeSuiteControls.CheckBox chkInicializa 
      Height          =   372
      Left            =   1680
      TabIndex        =   5
      Top             =   2280
      Width           =   6012
      _Version        =   1310723
      _ExtentX        =   10604
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Inicializa la configuración en el destino?"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.ComboBox cboFuente 
      Height          =   312
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   7092
      _Version        =   1310723
      _ExtentX        =   12515
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboDestino 
      Height          =   312
      Left            =   1680
      TabIndex        =   4
      Top             =   1800
      Width           =   7092
      _Version        =   1310723
      _ExtentX        =   12515
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
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copia de Estructuras Contables"
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
      Height          =   372
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   8532
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fuente"
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
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   972
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmCntX_Esquemas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAplicar_Click()
Dim strSQL As String, rsFuente As New ADODB.Recordset
Dim lngCodFuente As Long, lngCodDestino As Long
Dim rsDestino As New ADODB.Recordset

If cboFuente.Text = "" Then Exit Sub
If cboDestino.Text = "" Then Exit Sub

If cboFuente.ItemData(cboFuente.ListIndex) = cboDestino.ItemData(cboDestino.ListIndex) Then
    MsgBox "El origen y destino son el mismo!", vbExclamation
    Exit Sub
End If


On Error GoTo vError

Me.MousePointer = vbHourglass

lblX.Caption = "Verificando..."
lblX.Refresh


strSQL = "select * from CntX_Contabilidades where cod_contabilidad = " & cboFuente.ItemData(cboFuente.ListIndex)
Call OpenRecordSet(rsFuente, strSQL, 0)

  lngCodFuente = rsFuente!COD_CONTABILIDAD

strSQL = "select * from CntX_Contabilidades where cod_contabilidad = " & cboDestino.ItemData(cboDestino.ListIndex)
Call OpenRecordSet(rsDestino, strSQL, 0)
  
  lngCodDestino = rsDestino!COD_CONTABILIDAD

If rsFuente!Nivel1 <> rsDestino!Nivel1 Then
  Me.MousePointer = vbDefault
  MsgBox "La estructura de la cuenta contable no es la misma en el Fuente-Destino...", vbCritical
  Exit Sub
End If

If rsFuente!Nivel1 <> rsDestino!Nivel1 Then
  Me.MousePointer = vbDefault
  MsgBox "La estructura de la cuenta contable no es la misma en el Fuente-Destino...", vbCritical
  Exit Sub
End If

If rsFuente!Nivel2 <> rsDestino!Nivel2 Then
  Me.MousePointer = vbDefault
  MsgBox "La estructura de la cuenta contable no es la misma en el Fuente-Destino...", vbCritical
  Exit Sub
End If

If rsFuente!Nivel3 <> rsDestino!Nivel3 Then
  Me.MousePointer = vbDefault
  MsgBox "La estructura de la cuenta contable no es la misma en el Fuente-Destino...", vbCritical
  Exit Sub
End If

If rsFuente!Nivel4 <> rsDestino!Nivel4 Then
  Me.MousePointer = vbDefault
  MsgBox "La estructura de la cuenta contable no es la misma en el Fuente-Destino...", vbCritical
  Exit Sub
End If

If rsFuente!Nivel5 <> rsDestino!Nivel5 Then
  Me.MousePointer = vbDefault
  MsgBox "La estructura de la cuenta contable no es la misma en el Fuente-Destino...", vbCritical
  Exit Sub
End If

If rsFuente!Nivel6 <> rsDestino!Nivel6 Then
  Me.MousePointer = vbDefault
  MsgBox "La estructura de la cuenta contable no es la misma en el Fuente-Destino...", vbCritical
  Exit Sub
End If

If rsFuente!Nivel7 <> rsDestino!Nivel7 Then
  Me.MousePointer = vbDefault
  MsgBox "La estructura de la cuenta contable no es la misma en el Fuente-Destino...", vbCritical
  Exit Sub
End If

If rsFuente!Nivel8 <> rsDestino!Nivel8 Then
  Me.MousePointer = vbDefault
  MsgBox "La estructura de la cuenta contable no es la misma en el Fuente-Destino...", vbCritical
  Exit Sub
End If

rsFuente.Close
rsDestino.Close


'Inicia Copia de Información

lblX.Caption = "Copiando Información [Espere]"
lblX.Refresh

strSQL = "exec spCntX_Util_Contabilidad_Copia " & cboFuente.ItemData(cboFuente.ListIndex) _
       & "," & cboDestino.ItemData(cboDestino.ListIndex) _
       & "," & chkInicializa.Value & ",'" & glogon.Usuario & "','*xHM1tOk3n$'"
Call ConectionExecute(strSQL)

lblX.Caption = ""

Me.MousePointer = vbDefault

MsgBox "Copia de Esquemas Terminada Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub Form_Load()
Dim strSQL As String


vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


strSQL = "select COD_CONTABILIDAD as 'Idx', NOMBRE as 'ItmX' from CntX_Contabilidades"

Call sbCbo_Llena_New(cboFuente, strSQL, False, True)
Call sbCbo_Copia(cboFuente, cboDestino)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub
