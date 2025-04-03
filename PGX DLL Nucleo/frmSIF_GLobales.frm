VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmSIF_GLobales 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Variables Globales"
   ClientHeight    =   3540
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   10092
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSIF_GLobales.frx":0000
   ScaleHeight     =   3540
   ScaleWidth      =   10092
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlb 
      Height          =   312
      Left            =   8400
      TabIndex        =   1
      Top             =   480
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   550
      ButtonWidth     =   1905
      ButtonHeight    =   550
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Visualizar"
            Key             =   "Visualizar"
            Object.ToolTipText     =   "Visualizar Variables Globales"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9600
      Top             =   360
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_GLobales.frx":6852
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2292
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   9492
      _Version        =   524288
      _ExtentX        =   16743
      _ExtentY        =   4043
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   496
      ScrollBars      =   2
      SpreadDesigner  =   "frmSIF_GLobales.frx":694E
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Variables Globales del Sistema"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   10080
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmSIF_GLobales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

vModulo = 10
vGrid.AppearanceStyle = fxGridStyle


Call Formularios(Me)
Call RefrescaTags(Me)


With vGrid
  .MaxCols = 2
  .MaxRows = 4
  
  .Row = 1
  .Col = 1
  .Text = "Directorio de Resultados"
  .Col = 2
  .Text = SIFGlobal.DirectorioDeResultados

  .Row = 2
  .Col = 1
  .Text = "Reportes Personalizados"
  .Col = 2
  .Text = SIFGlobal.ReportesPersonalizados

  .Row = 3
  .Col = 1
  .Text = "Puertos Disponibles"
  .Col = 2
  .Text = SIFGlobal.PuertosDisponibles
  
  .Row = 4
  .Col = 1
  .Text = "Fondo de Pantalla"
  .Col = 2
  .Text = SIFGlobal.FondoDePantalla
  
End With


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim vCadena As String, fn
Dim i As Integer

On Error GoTo vError

fn = FreeFile

Open GLOBALES.gAppRuta & "\Global.ini" For Output As #fn  ' Crea el Archivo.
    With vGrid
         For i = 1 To .MaxRows
          Select Case i
            Case 3 'Puertos
                .Row = 3
                .Col = 1
                vCadena = SIFGlobal.fxStringRelleno(.Text, "D", " ", 30) & "="
                .Col = 2
                vCadena = vCadena & SIFGlobal.fxEncryptaNumero(.Text, True)
                Print #fn, vCadena
            
            Case Else
                .Row = i
                .Col = 1
                vCadena = SIFGlobal.fxStringRelleno(.Text, "D", " ", 30) & "="
                .Col = 2
                vCadena = vCadena & .Text
                Print #fn, vCadena
          
          End Select
        Next i
      
    End With
Close #fn

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description)
 Cancel = 1
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
  vGrid.Visible = True
End Sub


