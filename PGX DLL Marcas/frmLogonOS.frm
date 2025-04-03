VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogonOS 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Versión del Sistema Operativo ?"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogonOS.frx":0000
   ScaleHeight     =   2655
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlb 
      Height          =   330
      Left            =   3720
      TabIndex        =   5
      Top             =   2160
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   582
      ButtonWidth     =   2434
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Conectar"
            Key             =   "Conectar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Desconectar"
            Key             =   "Desconectar"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6480
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogonOS.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogonOS.frx":D0B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboIdioma 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1440
      Width           =   3975
   End
   Begin VB.ComboBox cboOS 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   7320
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   7320
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Sistema Operativo ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Idioma"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Sistema Operativo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogonOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

cboOS.AddItem "01 - Microsoft Windows 95/ 98 / Me"
cboOS.AddItem "02 - Microsoft Windows NT / 2000 / XP"
cboOS.AddItem "03 - Microsoft Windows Vista / Windows 7"
cboOS.Text = "03 - Microsoft Windows Vista / Windows 7"

cboIdioma.AddItem "01 - Español"
cboIdioma.AddItem "02 - Ingles"
cboIdioma.Text = "01 - Español"

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim vCadena As String, fn

Select Case Button.Key
    Case "Conectar"

        fn = FreeFile
        Open App.Path & "\System.Ini" For Output As #fn  ' Crea Archivo.
        
        Print #fn, cboOS.Text
        Print #fn, cboIdioma.Text
        
        Close #fn
    
        MsgBox "Sistema Operativo : " & Mid(cboOS.Text, 3, Len(cboOS.Text)), vbInformation
    
        Unload Me
            
            
    Case "Desconectar"
        Unload Me
End Select

End Sub
