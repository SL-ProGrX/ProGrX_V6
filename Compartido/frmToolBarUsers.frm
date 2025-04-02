VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmToolBarUsers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Personalización de las Barras de Herramientas"
   ClientHeight    =   2064
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   6588
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2064
   ScaleWidth      =   6588
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbo 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   4815
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   264
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   4284
      _ExtentX        =   7557
      _ExtentY        =   466
      ButtonWidth     =   487
      ButtonHeight    =   466
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   $"frmToolBarUsers.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Barras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmToolBarUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cbo_Click()
If vPaso Then Exit Sub
    gToolBar = fxCodigoCbo(cbo)
    Call sbToolBarIconos(tlb)
End Sub

Private Sub Form_Load()
vPaso = True

cbo.AddItem "00 - Win Clasico"
cbo.AddItem "01 - XP Color"
cbo.AddItem "02 - Green Format"
cbo.AddItem "03 - XP Clasico"
cbo.AddItem "04 - SIF Old"


Select Case gToolBar
  Case "00"
    cbo.Text = "00 - Win Clasico"
  Case "01"
    cbo.Text = "01 - XP Color"
  Case "02"
    cbo.Text = "02 - Green Format"
  Case "03"
    cbo.Text = "03 - XP Clasico"
  Case "04"
    cbo.Text = "04 - SIF Old"
End Select

vPaso = False

End Sub

Public Sub sbToolBarWrite()
Dim fn

On Error GoTo vError

fn = FreeFile

'Verifica si el Archivo Existe
gToolBar = fxCodigoCbo(cbo)
  Open App.Path & "\meToolBar.ini" For Output As #fn  ' Escribe el archivo.
    Print #fn, gToolBar
  Close #fn

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call sbToolBarWrite
End Sub
