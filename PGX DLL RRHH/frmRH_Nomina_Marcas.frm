VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Begin VB.Form frmRH_Nomina_Marcas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RRHH: Registro de Marcas"
   ClientHeight    =   2976
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2976
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   9360
      Top             =   0
   End
   Begin XtremeSuiteControls.ComboBox cboNomina 
      Height          =   312
      Left            =   1920
      TabIndex        =   1
      Top             =   1560
      Width           =   5532
      _Version        =   1310720
      _ExtentX        =   9758
      _ExtentY        =   550
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   672
      Left            =   1920
      TabIndex        =   3
      Top             =   1920
      Width           =   5532
      _Version        =   1310720
      _ExtentX        =   9758
      _ExtentY        =   1185
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
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   0
      Left            =   7560
      TabIndex        =   4
      Top             =   1920
      Width           =   492
      _Version        =   1310720
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmRH_Nomina_Marcas.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   1
      Left            =   8040
      TabIndex        =   5
      Top             =   1920
      Width           =   492
      _Version        =   1310720
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmRH_Nomina_Marcas.frx":0700
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   2
      Left            =   8520
      TabIndex        =   6
      Top             =   1920
      Width           =   492
      _Version        =   1310720
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmRH_Nomina_Marcas.frx":0E19
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   720
      TabIndex        =   7
      Top             =   1920
      Width           =   1332
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Nómina"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   7
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Carga Archivo para Control de Marcas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   -120
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmRH_Nomina_Marcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnArchivo_Click(Index As Integer)
Dim vMensaje As String

        
Select Case Index
  
  Case 0 'buscar
  
    txtArchivo.Text = ""
    Call sbBuscaArchivo(1)
  
  Case 1 'Cargar
'       Call sbCargaDeducciones(1)
    
  Case 2 'info
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: IMPORT" & vbCrLf _
              & " 3. Columnas.: IDENTIFICACION, FECHA, ENTRADA_1, SALIDA_1, ENTRADA_2, SALIDA_2,ENTRADA_3, SALIDA_3, ENTRADA_4, SALIDA_4"
     
     MsgBox vMensaje, vbInformation
         
End Select


End Sub



Private Sub sbBuscaArchivo(vTipo As Integer)


With frmContenedor.CD
'    If vTipo = 1 Or chkExcel.Value = vbChecked Then
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]..."
        .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If

        If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
            'Ok
        Else
            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
            Exit Sub
        End If
        
        txtArchivo.Text = .FileName
    
'    Else
'        .InitDir = "C:\"
'        .DialogTitle = "Localice Archivo de Deducciones [Texto]..."
'        .Filter = "*.txt"
'        .ShowOpen
'
'        If .FileName = "" Then
'            MsgBox "Archivo no válido...", vbExclamation
'            Exit Sub
'        End If
'
'        If UCase(Right(.FileName, 3)) = "XLS" Then
'            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
'            Exit Sub
'        End If
'
'        'If UCase(Right(.FileName, 3)) <> "TXT" Or UCase(Right(.FileName, 3)) <> "DAT" Then
'         '   MsgBox "La Extensión del Archivo no es válido...", vbExclamation
'         '   Exit Sub
'        'End If
'
'        txtArchivo.Text = .FileName
'
'End If
End With

End Sub


Private Sub Form_Load()
vModulo = 23

 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


'Nomina
    strSQL = "select COD_NOMINA as Idx, rtrim(Descripcion) as ItmX from RH_NOMINAS_CATALOGO"
    Call sbCbo_Llena_New(cboNomina, strSQL, False, True)




Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


