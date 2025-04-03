VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCC_ActualizaUPUT 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualizacion de Unidades Programaticas "
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   156
      Left            =   0
      TabIndex        =   0
      Top             =   3408
      Width           =   9408
      _ExtentX        =   16589
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton cmdActualiza 
      Height          =   852
      Left            =   7320
      TabIndex        =   3
      Top             =   1920
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Actualiza Unidades"
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
      Appearance      =   21
      Picture         =   "frmCC_ActualizaUPUT.frx":0000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualiza Unidades Programaticas/Trabajo"
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
      Caption         =   "..."
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   6855
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
Dim strCadena As String, lng As Long
Dim vCedula As String, vUP As String, vUT As String

On Error GoTo vError


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

Dim Contenido As String
Dim lineas() As String
Dim iCount As Long


' Abrir el archivo y leer todo su contenido en una sola variable
Open frmContenedor.CD.FileName For Binary As #1
    Contenido = Space$(LOF(1)) ' Reservar espacio suficiente
    Get #1, , Contenido ' Leer todo el contenido del archivo
Close #1

' Dividir el contenido en líneas utilizando el carácter LF (Chr(10))
lineas = Split(Contenido, Chr(10))


If prgBar.Max < UBound(lineas) Then
    prgBar.Max = UBound(lineas) + 1
End If

prgBar.Min = 1
DoEvents


lbl = "Actualizando ..."
lbl.Refresh

Dim vInicial As Boolean

vInicial = True
strSQL = ""

' Recorrer cada línea y procesarla
For iCount = 0 To UBound(lineas)
    strCadena = lineas(iCount) ' Esto imprime cada línea en la ventana de depuración

   vCedula = Trim(Format(Mid(strCadena, 1, 11), "###########"))
   vUP = Format(Mid(strCadena, 54, 4), "####")
   vUT = Format(Mid(strCadena, 58, 4), "####")
         
   If vInicial Then
        strSQL = strSQL & Space(10) & "exec spPrmProcAddUPUT_Actualiza_Manual '" & vCedula & "','" & vUP & "','" & vUT & "', 1"
        vInicial = False
   Else
        strSQL = strSQL & Space(10) & "exec spPrmProcAddUPUT_Actualiza_Manual '" & vCedula & "','" & vUP & "','" & vUT & "', 2"
   End If
   
   If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
   End If
   
   If prgBar.Max > prgBar.Value Then prgBar.Value = prgBar.Value + 1
   lbl.Caption = "Cargando..Registro # " & prgBar.Value & " de " & prgBar.Max & "     " & Format((prgBar.Value / prgBar.Max) * 100, "##0") & "%"
   lbl.Refresh

Next iCount


lbl.Caption = "Procesando Cambios, Espere!"
lbl.Refresh

strSQL = strSQL & Space(10) & "exec spPrmProcAddUPUT_Actualiza_Manual '" & vCedula & "','" & vUP & "','" & vUT & "', 3"
Call ConectionExecute(strSQL)


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

