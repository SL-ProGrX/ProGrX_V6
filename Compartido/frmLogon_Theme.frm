VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Begin VB.Form frmLogon_Theme 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambio de Tema de la Aplicación:"
   ClientHeight    =   3330
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
      _Version        =   1310720
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Cambiar"
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.ComboBox cboTema 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   3135
      _Version        =   1310720
      _ExtentX        =   5530
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tema:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cambiar de Tema de Aplicación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   3972
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmLogon_Theme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAplicar_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSEG_Usuario_Theme_Registra '" & glogon.Usuario & "','" & cboTema.Text & "'"
Call ConectionExecute(strSQL, 1)

MsgBox "Cambio de Tema realizado satisfactoriamente! Debe Cerrar la aplicación y volver a ingresar para ver el cambio!", vbInformation

Unload Me

vError:

End Sub

Private Sub Form_Load()
 Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento
 
 With cboTema
   .AddItem "Default"
   .AddItem "Aqua Normal"
   .AddItem "ProGrx Blue"
   .AddItem "Office Blue"
   .AddItem "Aero Normal"
   .AddItem "Le5 Blue"
   .AddItem "iTunes"
   .Text = glogon.ProGrX_Theme
 End With
 
 
 
End Sub
