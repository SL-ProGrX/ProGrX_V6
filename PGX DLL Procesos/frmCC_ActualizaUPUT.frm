VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCC_ActualizaUPUT 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualizacion de Unidades Programaticas "
   ClientHeight    =   3564
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   9408
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3564
   ScaleWidth      =   9408
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   156
      Left            =   0
      TabIndex        =   0
      Top             =   3408
      Width           =   9408
      _ExtentX        =   16595
      _ExtentY        =   275
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton cmdActualiza 
      Height          =   852
      Left            =   7320
      TabIndex        =   3
      Top             =   1920
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Actualiza Unidades"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmCC_ActualizaUPUT.frx":0000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualiza Unidades Programaticas/Trabajo"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   360
      Width           =   7212
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   4692
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCC_ActualizaUPUT.frx":09C3
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   852
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   5772
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCC_ActualizaUPUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdActualiza_Click()
Dim strSQL As String
Dim fn, strCadena As String, lng As Long
Dim vCedula As String, vUP As String, vUT As String

On Error GoTo vError

fn = FreeFile

With frmContenedor.CD
 .DialogTitle = "Localice archivo con las deducciones de Planilla..."
 .Filter = "*.*"
 .InitDir = "C:\"
 .ShowOpen
End With

If frmContenedor.CD.FileName = "" Then
 MsgBox "Seleccione el Archivo de Deducciones del Proceso " & Format(GLOBALES.glngFechaCR, "####-##"), vbInformation
 Exit Sub
End If


MsgBox "Se procederá a cargar los registros del archivo :" & frmContenedor.CD.FileName, vbInformation

Me.MousePointer = vbHourglass
prgBar.Min = 1
prgBar.Max = 2
Open frmContenedor.CD.FileName For Input As #fn   ' Lee el archivo.
 Do While Not EOF(fn)
   Input #fn, strCadena
   prgBar.Max = prgBar.Max + 1
 Loop
Close #fn

lbl = "Actualizando ..."
lbl.Refresh

prgBar.Min = 1
DoEvents
Open frmContenedor.CD.FileName For Input As #fn   'Lee el Archivo y lo compara
Do While Not EOF(fn)
   Input #fn, strCadena
    
   vCedula = Trim(Format(Mid(strCadena, 1, 11), "###########"))
   vUP = Format(Mid(strCadena, 54, 4), "####")
   vUT = Format(Mid(strCadena, 58, 4), "####")
         
   strSQL = "update socios set up='" & vUP & "',ut='" & vUT _
          & "' where cedula='" & vCedula & "'"
   Call ConectionExecute(strSQL)
   
   If prgBar.Max > prgBar.Value Then prgBar.Value = prgBar.Value + 1
   lbl.Caption = "Cargando..Registro # " & prgBar.Value & " de " & prgBar.Max & "     " & Format((prgBar.Value / prgBar.Max) * 100, "##0") & "%"
   lbl.Refresh

Loop
Close #fn

lbl.Caption = ""

Me.MousePointer = vbDefault

prgBar.Value = 1

lbl.Caption = ""

MsgBox "Información Actualizada ...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 10
 Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
End Sub

