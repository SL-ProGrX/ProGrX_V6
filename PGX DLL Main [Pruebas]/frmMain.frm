VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modulo de Pruebas de DLLs"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   8625
   Begin VB.CommandButton cmdProbar 
      Caption         =   "Probar Opción"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2280
      Width           =   5055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProbar_Click()
'Ejemplo
        Dim SIFNucleo As clsSIFNucleo
  
        Set SIFNucleo = New clsSIFNucleo
        
        Call SIFNucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
            , App.Path, glogon.ConectRPT, 1)
        
        Set SIFNucleo = Nothing
End Sub
