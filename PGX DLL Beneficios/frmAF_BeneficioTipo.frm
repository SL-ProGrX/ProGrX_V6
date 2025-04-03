VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_BeneficioTipo 
   BorderStyle     =   0  'None
   ClientHeight    =   1950
   ClientLeft      =   5715
   ClientTop       =   4830
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTipo 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.OptionButton OptMOnetario 
         Caption         =   "Tipo Monetario"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton OptProducto 
         Caption         =   "Tipo Producto"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Width           =   2295
      End
      Begin XtremeSuiteControls.PushButton cmdCerrar 
         Height          =   372
         Left            =   2760
         TabIndex        =   4
         Top             =   1440
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cerrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   " Seleccione el Tipo de Beneficio     "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3705
      End
   End
End
Attribute VB_Name = "frmAF_BeneficioTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
If OptMOnetario.Value = True Then
   GLOBALES.gTag3 = "M"
Else
   GLOBALES.gTag3 = "P"
End If

' frmAF_BeneficioAsg.lblTipo.Caption =GLOBALES.gTag3
 UnLoad Me

End Sub

