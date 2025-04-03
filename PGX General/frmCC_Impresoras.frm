VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCC_Impresoras 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Impresoras"
   ClientHeight    =   3684
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8232
   Icon            =   "frmCC_Impresoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3684
   ScaleWidth      =   8232
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   612
      Index           =   5
      Left            =   6480
      TabIndex        =   4
      Top             =   2760
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Guardar"
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.ComboBox cboRecibos 
      Height          =   312
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   6252
      _Version        =   1245187
      _ExtentX        =   11028
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboLetras 
      Height          =   312
      Left            =   1800
      TabIndex        =   6
      Top             =   1800
      Width           =   6252
      _Version        =   1245187
      _ExtentX        =   11028
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboCheques 
      Height          =   312
      Left            =   1800
      TabIndex        =   7
      Top             =   2160
      Width           =   6252
      _Version        =   1245187
      _ExtentX        =   11028
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheques"
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
      Height          =   315
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Configuración de Impresoras"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   1880
      TabIndex        =   2
      Top             =   360
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Letras de Cambio"
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
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Recibos"
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
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmCC_Impresoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPath As String



Private Sub btnGuardar_Click(Index As Integer)
 Call EscribeArchivoImpresoras
 MsgBox "Configuración de Impresoras Guardada..."
End Sub

Private Sub Form_Load()
 
vPath = SIFGlobal.fxCarpetaEspecial(CSIDL_PERSONAL)
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

 Call LeeArchivoImpresoras
End Sub


Private Sub LeeArchivoImpresoras()
Dim fn, strCadena As String, i As Integer

fn = FreeFile

Call LlenaCombosImpresoras(cboRecibos)
Call LlenaCombosImpresoras(cboLetras)
Call LlenaCombosImpresoras(cboCheques)

On Error GoTo vError
i = 1
Open vPath + "\PGX_Impresoras.ini" For Input As #fn
Do While Not EOF(fn)
  Input #fn, strCadena
  If i = 1 Then cboRecibos.Text = strCadena
  If i = 2 Then cboLetras.Text = strCadena
  If i = 3 Then cboCheques.Text = strCadena
  i = i + 1
Loop
Close #fn

Exit Sub
vError:
MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub EscribeArchivoImpresoras()
Dim fn
fn = FreeFile
On Error Resume Next
Kill vPath + "\PGX_Impresoras.ini"
Open vPath + "\PGX_Impresoras.ini" For Output As #fn
  Print #fn, cboRecibos.Text
  Print #fn, cboLetras.Text
  Print #fn, cboCheques.Text
  
Close #fn

End Sub

Private Sub LlenaCombosImpresoras(cbo As Object)
Dim xPrinter As Printer

For Each xPrinter In Printers
 cbo.AddItem xPrinter.DeviceName + ";" + xPrinter.Port + ";" + xPrinter.DriverName
Next xPrinter

End Sub

