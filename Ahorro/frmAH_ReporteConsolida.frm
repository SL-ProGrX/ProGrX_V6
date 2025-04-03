VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAH_ReportesRangosFecha 
   Caption         =   "Reportes de Aportes de Socios y Exsocios"
   ClientHeight    =   3090
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5985
   Begin MSComCtl2.DTPicker dtpde 
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   35473
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Socio"
      Height          =   2055
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   5415
      Begin VB.OptionButton Opt 
         Caption         =   "Liquidaciones"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   8
         Top             =   1440
         Width           =   3255
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Anulaciones"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   7
         Top             =   1080
         Width           =   3255
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Aplicacion de deudas a aportes"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dtpde 
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   35473
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "De:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_ReporteConsolida.frx":0000
            Key             =   "Ejecutar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_ReporteConsolida.frx":031C
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_ReporteConsolida.frx":0638
            Key             =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblrep 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   1535
      ButtonWidth     =   1376
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ejecutar"
            Key             =   "Ejecutar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Key             =   "Salir"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmAH_ReportesRangosFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub tblrep_ButtonClick(ByVal Button As MSComctlLib.Button)
 Dim strRuta As String, strsql As String, strmes As String, strmes1 As String
 
 strRuta = "C:\proyectos\is-ase\ahorro\reportes\"
 
 With frmCC_MenuPrincipal.Crt
     frmCC_MenuPrincipal.Crt.Reset
 Select Case Button.Key
   Case "Ejecutar"
   If Opt(0).Value = True Then
    frmCC_MenuPrincipal.Crt.ReportFileName = strRuta & "ahapl.rpt"
    'strsql = "{AHORRO_APL.FECHA}>=" & "Date(" & str(Year(dtpde(0))) & "," & str(Month(dtpde(0))) & "," & str(Day(dtpde(0))) & ")"
    'strsql = strsql & " and {AHORRO_APL.FECHA}<=" & "Date(" & str(Year(dtpde(1))) & "," & str(Month(dtpde(1))) & "," & str(Day(dtpde(1))) & ")"
    Select Case Month(dtpde(0))
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9
          strmes = "0" & Trim(Month(dtpde(0)))
    End Select
    Select Case Month(dtpde(1))
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9
          strmes1 = "0" & Trim(Month(dtpde(1)))
    End Select
    strsql = "{AHORRO_APL.FECPRO}>=" & Trim(Year(dtpde(0))) & strmes
    strsql = strsql & " and {AHORRO_APL.FECPRO}<=" & Trim(Year(dtpde(1))) & strmes1
    frmCC_MenuPrincipal.Crt.SelectionFormula = strsql
    MsgBox strsql
    frmCC_MenuPrincipal.Crt.PrintReport
   Else
    If Opt(1).Value = True Then
        frmCC_MenuPrincipal.Crt.ReportFileName = strRuta & "ahanu.rpt"
        strsql = "{ANULA.FECHA}>=" & "Date(" & str(Year(dtpde(0))) & "," & str(Month(dtpde(0))) & "," & str(Day(dtpde(0))) & ")"
        strsql = strsql & " and {ANULA.FECHA}<=" & "Date(" & str(Year(dtpde(1))) & "," & str(Month(dtpde(1))) & "," & str(Day(dtpde(1))) & ")"
        frmCC_MenuPrincipal.Crt.SelectionFormula = strsql
        MsgBox strsql
        frmCC_MenuPrincipal.Crt.PrintReport
    Else
      If Opt(2).Value = True Then
        frmCC_MenuPrincipal.Crt.ReportFileName = strRuta & "ahliq.rpt"
        strsql = "{LIQUIDACION.FECLIQ}>=" & "Date(" & str(Year(dtpde(0))) & "," & str(Month(dtpde(0))) & "," & str(Day(dtpde(0))) & ")"
        strsql = strsql & " and {LIQUIDACION.FECLIQ}<=" & "Date(" & str(Year(dtpde(1))) & "," & str(Month(dtpde(1))) & "," & str(Day(dtpde(1))) & ")"
        frmCC_MenuPrincipal.Crt.SelectionFormula = strsql
        MsgBox strsql
        frmCC_MenuPrincipal.Crt.PrintReport
      End If
    End If
   End If
   
   Case "Salir"
     Unload Me
 End Select
 End With



End Sub
