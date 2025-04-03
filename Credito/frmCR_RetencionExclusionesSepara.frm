VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCR_RetencionExclusionesSepara 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Separa Exclusiones del Archivo de Deducciones"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.PushButton btnSearch_X 
      Height          =   252
      Left            =   7080
      TabIndex        =   7
      Top             =   360
      Width           =   372
      _Version        =   1245185
      _ExtentX        =   656
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "..."
      Appearance      =   2
   End
   Begin VB.TextBox txtResultados 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2520
      Width           =   4455
   End
   Begin VB.TextBox txtDeducciones 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1440
      Width           =   4455
   End
   Begin VB.TextBox txtExclusiones 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   360
      Width           =   4455
   End
   Begin MSComctlLib.Toolbar tlbProceso 
      Height          =   336
      Left            =   5520
      TabIndex        =   6
      Top             =   3840
      Width           =   2604
      _ExtentX        =   4604
      _ExtentY        =   582
      ButtonWidth     =   1461
      ButtonHeight    =   550
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar"
            Key             =   "aplicar"
            Object.ToolTipText     =   "Aplicar Archivo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "cancelar"
            Object.ToolTipText     =   "cancelar operacion"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.PushButton btnSearch_Y 
      Height          =   252
      Left            =   7080
      TabIndex        =   8
      Top             =   1440
      Width           =   372
      _Version        =   1245185
      _ExtentX        =   656
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "..."
      Appearance      =   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8280
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo de Resultados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo de Deducciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo de Exclusiones    (Microsoft Excel)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmCR_RetencionExclusionesSepara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxExclusion(pCedula As String) As Boolean
Dim vResultado As Boolean, rs As New ADODB.Recordset

vResultado = False



Set rs = Excel_Load(txtExclusiones.Text, "Import")

With rs
  .MoveFirst
  Do While Not .EOF
    If Trim(pCedula) = Trim(CStr(!Cedula)) Then
        vResultado = True
        Exit Do
    End If
    .MoveNext
  Loop
  .Close
End With

fxExclusion = vResultado

End Function


Private Sub sbAplicar()
Dim i As Integer, vCadena As String, vArchivo(1) As String
Dim fn, fe, fp

On Error GoTo vError

If txtDeducciones.Text = "" Or txtExclusiones.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass

fn = FreeFile

Open txtDeducciones.Text For Input As #fn    ' Lee el archivo.

fp = FreeFile
vArchivo(0) = Mid(txtDeducciones.Text, 1, Len(txtDeducciones.Text) - 4) & "_Real.txt"
Open vArchivo(0) For Output As #fp    ' Lee el archivo.

fe = FreeFile
vArchivo(1) = Mid(txtDeducciones.Text, 1, Len(txtDeducciones.Text) - 4) & "_Exclusiones.txt"
Open vArchivo(1) For Output As #fe    ' Lee el archivo.

 Do While Not EOF(fn)
   Input #fn, vCadena
   
   If fxExclusion(Mid(vCadena, 1, 11)) Then
      Print #fe, vCadena
   Else
      Print #fp, vCadena
   End If
   
 
 Loop

Close 'Cierra todos los archivos

txtResultados.Text = vArchivo(0) & vbCrLf & vArchivo(1)


Me.MousePointer = vbDefault
MsgBox "Información Procesada Satisfactoriamente...!", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Close
End Sub



Private Sub btnSearch_X_Click()

        
txtExclusiones.Text = ""

With frmContenedor.CD
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
    
    txtExclusiones.Text = .FileName

End With

txtResultados.Text = ""


End Sub

Private Sub btnSearch_Y_Click()

txtDeducciones.Text = ""

With frmContenedor.CD
         .InitDir = "C:\"
         .DialogTitle = "Localice Archivo de Deducciones..."
         .Filter = "*.txt"
         .ShowOpen
         
         If .FileName = "" Then
           MsgBox "Archivo no válido...", vbExclamation
           Exit Sub
         End If
         
         If UCase(Right(.FileName, 3)) <> "TXT" Then
           MsgBox "La Extensión del Archivo no es válido...", vbExclamation
           Exit Sub
         End If
        
   
 txtDeducciones.Text = .FileName

End With
 
txtResultados.Text = ""

End Sub

Private Sub tlbProceso_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "aplicar"
    Call sbAplicar
  Case "cancelar"
    txtResultados.Text = ""
End Select
End Sub


