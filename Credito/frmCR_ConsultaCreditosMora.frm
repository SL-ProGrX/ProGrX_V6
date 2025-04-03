VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmCR_ConsultaCreditosMora 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Morosidad ?"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerImg 
      Interval        =   200
      Left            =   10920
      Top             =   0
   End
   Begin VB.Timer TimerX 
      Interval        =   20
      Left            =   10920
      Top             =   360
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4452
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   11412
      _Version        =   524288
      _ExtentX        =   20129
      _ExtentY        =   7853
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
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
      MaxCols         =   15
      SpreadDesigner  =   "frmCR_ConsultaCreditosMora.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.ImageList imgSemaforos 
      Left            =   10200
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCreditosMora.frx":0F97
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCreditosMora.frx":10B5
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCreditosMora.frx":11DB
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCreditosMora.frx":1305
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCreditosMora.frx":1417
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCreditosMora.frx":152E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCreditosMora.frx":162F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCreditosMora.frx":1766
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCreditosMora.frx":187B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCuotas 
      Height          =   312
      Left            =   3000
      TabIndex        =   13
      Top             =   5880
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMoraInt 
      Height          =   312
      Left            =   3000
      TabIndex        =   14
      Top             =   6240
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMoraCargos 
      Height          =   312
      Left            =   3000
      TabIndex        =   15
      Top             =   6600
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMoraPoliza 
      Height          =   312
      Left            =   3000
      TabIndex        =   16
      Top             =   6960
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMoraPrincipal 
      Height          =   312
      Left            =   6480
      TabIndex        =   17
      Top             =   5880
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMoraFinanciera 
      Height          =   312
      Left            =   6480
      TabIndex        =   18
      Top             =   6240
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMoraLegal 
      Height          =   312
      Left            =   6480
      TabIndex        =   19
      Top             =   6600
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCbrJudicial 
      Height          =   312
      Left            =   6480
      TabIndex        =   20
      Top             =   6960
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pólizas Registradas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   960
      TabIndex        =   12
      Top             =   6960
      Width           =   1932
   End
   Begin VB.Image imgMoraFinanciera 
      Height          =   240
      Left            =   8400
      Picture         =   "frmCR_ConsultaCreditosMora.frx":199F
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No incluye Cobro Judicial"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   8880
      TabIndex        =   11
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Incluye Cobro Judicial"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   8880
      TabIndex        =   10
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Label lblCorte 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Corte : "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lblCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   360
      Width           =   8175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "En Cobro Judicial"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Index           =   6
      Left            =   4680
      TabIndex        =   7
      Top             =   6960
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mora Legal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   4680
      TabIndex        =   6
      Top             =   6600
      Width           =   1692
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mora Financiera"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   4680
      TabIndex        =   5
      Top             =   6240
      Width           =   1692
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Principal Atrasado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   4
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargos Registrados"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   960
      TabIndex        =   3
      Top             =   6600
      Width           =   1932
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Intereses Atrasados"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   6240
      Width           =   1932
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Cuotas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmCR_ConsultaCreditosMora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
 
 vGrid.Sheet = 1
 vGrid.MaxRows = 0
 vGrid.MaxCols = 14
 imgBanner.Picture = frmContenedor.imgBanner_01.Picture

strSQL = "select cedula,nombre,dbo.MyGetdate() as Fecha from socios where cedula = '" _
       & GLOBALES.gCedulaActual & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  lblCliente.Caption = Trim(rs!Cedula) & " - " & Trim(rs!Nombre)
  lblCorte.Caption = Format(rs!fecha, "dd/mm/yyyy")
End If
rs.Close

End Sub


Private Sub TimerImg_Timer()
imgMoraFinanciera.Visible = Not imgMoraFinanciera.Visible
End Sub

Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curMora(7) As Currency, i As Integer, x As Integer

TimerX.Interval = 0
TimerX.Enabled = False

For i = 0 To 7
  curMora(i) = 0
Next i


With vGrid
    .Sheet = 1
    .MaxRows = 0
    strSQL = "exec spSIFEstadoCreditosMora '" & GLOBALES.gCedulaActual & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      If rs!ProcesoCod = "J" Or (rs!MoraCuota) > 0 Then
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        For i = 1 To 14
          .col = i
          Select Case i
            Case 1 'Estado
              .TypePictPicture = imgSemaforos.ListImages.Item(1).Picture
        
                     Select Case rs!ProcesoCod
                      Case "N"
               
                        If Not IsNull(rs!Referencia) Then
                            If rs!MoraCuota = 0 Then .TypePictPicture = imgSemaforos.ListImages.Item(2).Picture
                            .FontBold = True
                        End If
                
                      
                      Case "J"
                          .TypePictPicture = imgSemaforos.ListImages.Item(7).Picture
                               
                          .TextTip = TextTipFixed
                          .TextTipDelay = 1000
                        
                          .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
                          .CellNoteIndicatorColor = vbRed
                          
                          .CellNote = ">> Cobro Judicial <<" & vbCrLf _
                                    & "Fecha : " & Format(rs!fecha_enviaproceso, "dd/mm/yyyy") & vbCrLf _
                                    & "Nota  : " & rs!observacion_proceso & ""
                      
                      Case "T"
                            If rs!MoraCuota = 0 Then .TypePictPicture = imgSemaforos.ListImages.Item(2).Picture
                            
                            If rs!IndicadorCbr > 0 Then
                               .TypePictPicture = imgSemaforos.ListImages.Item(9).Picture
                            End If
                
                     End Select
            
                    ' Si esta moroso indicar Mora siempre y cuando no este en cobro Judicial
                    If rs!MoraCuota > 0 And rs!ProcesoCod <> "J" Then
                      
                      .TypePictPicture = imgSemaforos.ListImages.Item(3).Picture
'                      vMora = True
                    
                      .TextTip = TextTipFixed
                      .TextTipDelay = 1000
                    
                      .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
                      .CellNoteIndicatorColor = vbBlue
                      
                      .CellNote = "Morosidad:  Cuotas: " & rs!MoraCuota & vbCrLf _
                                & "   Intereses : " & Format(rs!MoraInt, "Standard") & vbCrLf _
                                & "   Cargos    : " & Format(rs!MoraCargos, "Standard") & vbCrLf _
                                & "   Póliza    : " & Format(rs!MoraPoliza, "Standard") & vbCrLf _
                                & "   Principal : " & Format(rs!MoraPrincipal, "Standard") & vbCrLf _
                                & "   Cta.+ Vieja : " & Format(rs!MoraAntigua, "####-##") & vbCrLf _
                                & "   Cta. Ultima : " & Format(rs!MoraUltima, "####-##") & vbCrLf & vbCrLf _
                                & "   Total Mora  : " & Format(rs!MoraInt + rs!MoraCargos + rs!MoraPrincipal + rs!MoraPoliza, "Standard") & vbCrLf
                    
                    End If
            
            
            Case 2 'Operacion

              .Text = CStr(rs!Id_Solicitud)

        
             Case 3 'Linea
              .Text = CStr(rs!Codigo)
              .TextTip = TextTipFixed
              .TextTipDelay = 1000
              .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
              .CellNoteIndicatorColor = vbBlue
  
            .CellNote = Trim(rs!LineaX) & vbCrLf & vbCrLf & "Formaliza: " & Format(rs!FechaForp, "dd/mm/yyyy") & vbCrLf & "Usuario: " & Trim(rs!Userfor) & vbCrLf & "Oficina:" & rs!OficinaX & ""

            Case 4 '# Cuotas
              .Text = CStr(rs!MoraCuota)
              curMora(0) = curMora(0) + rs!MoraCuota
            Case 5 'Mora Intereses
              .Text = Format(rs!MoraInt, "Standard")
              curMora(1) = curMora(1) + rs!MoraInt
            Case 6 'Mora Cargos
              .Text = Format(rs!MoraCargos, "Standard")
              curMora(2) = curMora(2) + rs!MoraCargos
            
            Case 7 'Mora Poliza
              .Text = Format(rs!MoraPoliza, "Standard")
              curMora(7) = curMora(7) + rs!MoraPoliza
            
            
            Case 8 'Mora Principal
              .Text = Format(rs!MoraPrincipal, "Standard")
              curMora(3) = curMora(3) + rs!MoraPrincipal
            Case 9 'Mora Financiera

                  .Text = Format(rs!MoraPrincipal + rs!MoraCargos + rs!MoraInt + rs!MoraPoliza, "Standard")
                  curMora(4) = curMora(4) + rs!MoraPrincipal + rs!MoraCargos + rs!MoraInt + rs!MoraPoliza
            
            Case 10 'Mora Legal
              If rs!ProcesoCod = "J" Then
                  .Text = Format(rs!Saldo + rs!MoraCargos + rs!CbrIntereses, "Standard")
                  curMora(5) = curMora(5) + rs!Saldo + rs!MoraCargos + rs!CbrIntereses + rs!MoraPoliza
                  curMora(6) = curMora(6) + rs!Saldo + rs!MoraCargos + rs!CbrIntereses + rs!MoraPoliza
              Else
                  .Text = Format(rs!Saldo + rs!MoraCargos + rs!MoraInt, "Standard")
                  curMora(5) = curMora(5) + rs!Saldo + rs!MoraCargos + rs!MoraInt + rs!MoraPoliza
              End If
              
            Case 11 'Mora Cuota + Antigua
              .Text = Format(rs!MoraAntigua, "####-##")
            Case 12 'Ultimo Movimiento
              .Text = Format(rs!MoraUltima, "####-##")
            Case 13 'Garantia
              .Text = CStr(rs!Garantia)
            Case 14 'Proceso
              .Text = CStr(rs!Proceso)
            Case 15 'Destino
              .Text = CStr(rs!DestinoX & "")
          End Select
        
        Next i
      End If
      rs.MoveNext
    Loop
    rs.Close

End With


txtCuotas.Text = Format(curMora(0), "###,###,##0")
txtMoraInt.Text = Format(curMora(1), "Standard")
txtMoraCargos.Text = Format(curMora(2), "Standard")
txtMoraPoliza.Text = Format(curMora(7), "Standard")

txtMoraPrincipal.Text = Format(curMora(3), "Standard")
txtMoraFinanciera.Text = Format(curMora(4), "Standard")
txtMoraLegal.Text = Format(curMora(5), "Standard")
txtCbrJudicial.Text = Format(curMora(6), "Standard")

End Sub


Private Sub vGrid_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
Dim curMora(7) As Currency

If NewSheet = 1 Then Exit Sub
 
Me.MousePointer = vbHourglass
 
For i = 0 To 7
  curMora(i) = 0
Next i


With vGrid
    .Sheet = 2
    .MaxRows = 0
    strSQL = "exec spCbrPersonaMoraGarantia '" & GLOBALES.gCedulaActual & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        For i = 1 To 10
          .col = i
          Select Case i
            Case 1 'Garantia
              .Text = CStr(rs!Garantia)
            
            Case 2 'Saldos
              .Text = CStr(rs!Saldo)
             
             Case 3 'Operaciones
              .Text = CStr(rs!Operaciones)
            
            Case 4 'Mora Intereses Corrientes
              .Text = Format(rs!MorIntCor, "Standard")
              curMora(1) = curMora(1) + rs!MorIntCor
            
            Case 5 'Mora Intereses Moratorio
              .Text = Format(rs!MorIntMor, "Standard")
              curMora(1) = curMora(1) + rs!MorIntMor
            
            Case 6 'Mora Cargos
              .Text = Format(rs!MorCargos, "Standard")
              curMora(2) = curMora(2) + rs!MorCargos
            
            Case 7 'Mora Principal
              .Text = Format(rs!MorPrincipal, "Standard")
              curMora(3) = curMora(3) + rs!MorPrincipal
            
            Case 8 '# Cuotas
              .Text = CStr(rs!MorCuotas * 30)
              curMora(0) = curMora(0) + rs!MorCuotas
            
            Case 9 'Mora Financiera

                 .Text = Format(rs!MorPrincipal + rs!MorCargos + rs!MorIntCor + rs!MorIntMor, "Standard")
                 curMora(4) = curMora(4) + rs!MorPrincipal + rs!MorCargos + rs!MorIntCor + rs!MorIntMor
            
            Case 10 'Mora Legal
              
                 .Text = Format(rs!Saldo + rs!MorCargos + rs!MorIntCor + rs!MorIntMor, "Standard")
                 curMora(5) = curMora(5) + rs!Saldo + rs!MorCargos + rs!MorIntCor + rs!MorIntMor
              
              
'            Case 11 'Mora Cuota + Antigua
'              .Text = Format(rs!MoraAntigua, "####-##")
'            Case 12 'Ultimo Movimiento
'              .Text = Format(rs!MoraUltima, "####-##")
'            Case 13 'Garantia
'              .Text = CStr(rs!Garantia)
'            Case 14 'Proceso
'              .Text = CStr(rs!Proceso)
          End Select
        
        Next i
      rs.MoveNext
    Loop
    rs.Close

End With
 
Me.MousePointer = vbDefault

End Sub

