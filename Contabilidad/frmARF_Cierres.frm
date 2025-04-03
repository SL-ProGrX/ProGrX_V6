VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmARF_Cierres 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cierre del Periodo"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.DateTimePicker dtpPeriodo 
      Height          =   312
      Left            =   3600
      TabIndex        =   0
      Top             =   480
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1572
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   7812
      _Version        =   1572864
      _ExtentX        =   13779
      _ExtentY        =   2773
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnCierre 
         Height          =   1092
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Cerrar"
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
         Picture         =   "frmARF_Cierres.frx":0000
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmARF_Cierres.frx":09C3
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1092
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   4692
      End
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmARF_Cierres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub btnCierre_Click()
On Error GoTo vError


Dim i As Integer
i = MsgBox("Esta seguro que desea >> Cerrar << este periodo: " & Format(dtpPeriodo.Value, "yyyy-mm-dd"), vbYesNo)
If i = vbNo Then
    Exit Sub
End If

Me.MousePointer = vbHourglass

'strSQL = "exec spIVR_Cierre '" & Format(dtpPeriodo.Value, "yyyy-mm-dd") & " 23:59:59', '" & glogon.Usuario & "'"
'Call ConectionExecute(strSQL)
'
'Call Bitacora("Aplica", "Cierre de Arrendamientos: " & Format(dtpPeriodo.Value, "yyyy-mm-dd"))

Call sbCorte_Actual

Me.MousePointer = vbDefault

MsgBox "Mes Cerrado Satisfactoriamente...", vbInformation
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbCorte_Actual()

strSQL = "select isnull( max(Periodo_Corte), dbo.mygetdate() )   as 'CORTE'" _
       & "  From CNTX_PERIODOS Where Cod_Contabilidad = " & GLOBALES.gEnlace & " and estado = 'P'"
Call OpenRecordSet(rs, strSQL)
    dtpPeriodo.Value = rs!Corte
rs.Close

End Sub


Private Sub Form_Load()
  vModulo = 20
   
 Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

 Call sbCorte_Actual

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
End Sub




