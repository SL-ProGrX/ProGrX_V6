VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmSYS_Portal_WebApp_Sincroniza 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sincronizador para la Web/App"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ProgressBar prgBar 
      Height          =   135
      Left            =   0
      TabIndex        =   2
      Top             =   3120
      Width           =   10215
      _Version        =   1441793
      _ExtentX        =   18018
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin XtremeSuiteControls.PushButton btnSincronizar 
      Height          =   735
      Left            =   7440
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Sincronizar"
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
      Picture         =   "frmSYS_Portal_WebApp_Sincroniza.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   360
      Width           =   7575
      _Version        =   1441793
      _ExtentX        =   13361
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Sincronizar Portal con Usuarios Web/Apps nuevos"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblStatus 
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   4695
      _Version        =   1441793
      _ExtentX        =   8281
      _ExtentY        =   2355
      _StockProps     =   79
      Caption         =   $"frmSYS_Portal_WebApp_Sincroniza.frx":0719
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14055
   End
End
Attribute VB_Name = "frmSYS_Portal_WebApp_Sincroniza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String

Private Sub btnSincronizar_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

prgBar.Max = 4

lblStatus.Caption = "Sincronizando Servidor Principal..."
prgBar.Value = 1

strSQL = "exec spPersona_Portal_Sincroniza"
Call ConectionExecute(strSQL)


lblStatus.Caption = "Sincronizando Servidor South-Central 1 [Espere!]"
prgBar.Value = 2
Call sbWebApps_Sincroniza("progrx.southcentralus.cloudapp.azure.com", 1)

lblStatus.Caption = "Sincronizando Servidor South-Central 1 Subiendo Casos [Espere!]"
prgBar.Value = 3
Call sbWebApps_Sincroniza("progrx.southcentralus.cloudapp.azure.com", 2)

lblStatus.Caption = "Sincronizando Servidor South-Central 1 Sincronizando [Espere!]"
prgBar.Value = 4
Call sbWebApps_Sincroniza("progrx.southcentralus.cloudapp.azure.com", 3)


Me.MousePointer = vbDefault

MsgBox "Prceso de Sincronización finalizado satisfactoriamente!", vbInformation

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    

End Sub

Private Sub Form_Load()
vModulo = 13

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture



End Sub
