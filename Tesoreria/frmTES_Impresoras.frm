VERSION 5.00
Begin VB.Form frmTES_Impresoras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Impresoras"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmTES_Impresoras.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   8520
   Begin VB.ComboBox cboRecibos 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1920
      TabIndex        =   2
      Text            =   "cboRecibos"
      Top             =   960
      Width           =   6495
   End
   Begin VB.ComboBox cboCheques 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1920
      TabIndex        =   1
      Text            =   "cboLetras"
      Top             =   1320
      Width           =   6495
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   975
      Left            =   7080
      Picture         =   "frmTES_Impresoras.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   8400
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Configuración de Impresoras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Recibos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cheques"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   8520
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8160
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "frmTES_Impresoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGuardar_Click()
 Call EscribeArchivoImpresoras
 MsgBox "Configuración de Impresoras Guardada..."
End Sub


Private Sub Form_Load()
 Call LeeArchivoImpresoras
End Sub


Private Sub LeeArchivoImpresoras()
Dim fn, strCadena As String, i As Integer
fn = FreeFile

Call LlenaCombosImpresoras(cboRecibos)
Call LlenaCombosImpresoras(cboCheques)
On Error GoTo vError
i = 1
Open App.Path & "\TesPrinters.ini" For Input As #fn
Do While Not EOF(fn)
  Input #fn, strCadena
  If i = 1 Then cboRecibos.Text = strCadena
  If i = 2 Then cboCheques.Text = strCadena
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

Open App.Path & "\TesPrinters.ini" For Output As #fn
  Print #fn, cboRecibos.Text
  Print #fn, cboCheques.Text
Close #fn

End Sub

Private Sub LlenaCombosImpresoras(cbo As ComboBox)
Dim xPrinter As Printer

For Each xPrinter In Printers
 cbo.AddItem xPrinter.DeviceName + "; " + xPrinter.Port + "; " + xPrinter.DriverName
Next xPrinter

End Sub

