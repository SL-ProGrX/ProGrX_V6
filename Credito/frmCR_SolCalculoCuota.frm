VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_SolCalculoCuota 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cálculo de Cuotas"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9990
   HelpContextID   =   3015
   Icon            =   "frmCR_SolCalculoCuota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox gbBallons 
      Height          =   2532
      Left            =   2760
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   4092
      _Version        =   1441793
      _ExtentX        =   7218
      _ExtentY        =   4466
      _StockProps     =   79
      Caption         =   "Definición de Cuota Ballon"
      ForeColor       =   16711680
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtBallonCta 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   1320
         TabIndex        =   22
         Top             =   600
         Width           =   2532
         _Version        =   1441793
         _ExtentX        =   4466
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtBallonAjuste 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   1920
         TabIndex        =   23
         Top             =   1080
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "1"
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBallonCta 
         Height          =   432
         Left            =   2640
         TabIndex        =   24
         Top             =   1680
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   762
         _StockProps     =   79
         Caption         =   "Calcular"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   12
      End
      Begin VB.Label Label2 
         Caption         =   "Cuota Fija"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "Ajustar Saldos  faltando"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1692
      End
      Begin VB.Label Label2 
         Caption         =   "Cuota para Finalizar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   2520
         TabIndex        =   19
         Top             =   1080
         Width           =   1452
      End
   End
   Begin XtremeSuiteControls.PushButton btnCalcular 
      Height          =   312
      Left            =   8640
      TabIndex        =   16
      Top             =   480
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Calcular"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   12
   End
   Begin XtremeSuiteControls.CheckBox chkBallons 
      Height          =   492
      Left            =   8640
      TabIndex        =   15
      Top             =   1080
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cuota Ballon"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5412
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   9732
      _Version        =   524288
      _ExtentX        =   17166
      _ExtentY        =   9546
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
      MaxCols         =   498
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_SolCalculoCuota.frx":000C
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   8
      Top             =   120
      Width           =   2532
      _Version        =   1441793
      _ExtentX        =   4466
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCuota 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   1
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   9
      Top             =   1200
      Width           =   2532
      _Version        =   1441793
      _ExtentX        =   4466
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPlazo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   1
      EndProperty
      Height          =   312
      Left            =   2640
      TabIndex        =   10
      Top             =   480
      Width           =   852
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "1"
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTasa 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   1
      EndProperty
      Height          =   312
      Left            =   2640
      TabIndex        =   11
      Top             =   840
      Width           =   852
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTotalIntereses 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   1
      EndProperty
      Height          =   312
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Width           =   2532
      _Version        =   1441793
      _ExtentX        =   4466
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTermina 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   0
      EndProperty
      Height          =   312
      Left            =   6000
      TabIndex        =   13
      Top             =   480
      Width           =   2532
      _Version        =   1441793
      _ExtentX        =   4466
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboFactor 
      Height          =   312
      Left            =   6000
      TabIndex        =   14
      Top             =   1200
      Width           =   2532
      _Version        =   1441793
      _ExtentX        =   4471
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnReporte 
      Height          =   312
      Left            =   8640
      TabIndex        =   17
      Top             =   120
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Reporte"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   12
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Factor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   7
      Left            =   4680
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Termina"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   6
      Left            =   4680
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Intereses"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   5
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuota"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tasa"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plazo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmCR_SolCalculoCuota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mFrecuenciaPago As String

Private Sub sbCalculoCuota()
Dim curSaldo As Currency, i As Integer
Dim curIntx As Currency, curAmortiza As Currency
Dim vFecha As Date, curCuota As Currency
Dim curTasa As Currency, iDias As Integer
Dim curCuotaTmp As Currency, iPlazo As Integer
Dim TempInt As Currency, TempAmort As Currency

Dim TempCta As Currency, pQuincena As String, pDias As Integer

On Error GoTo vError

If Not fxValida Then Exit Sub

Me.MousePointer = vbHourglass

If CInt(txtPlazo) <= 0 Then
    Me.MousePointer = vbDefault
    MsgBox "Plazo debe ser > a cero", vbOKOnly
    txtPlazo = ""
    Exit Sub
End If

vFecha = fxFechaServidor

'Revision del Periodo de Ajuste de Cuotas Bullet
If CInt(txtBallonAjuste.Text) <= 0 Then
    txtBallonAjuste.Text = 1
End If

If CInt(txtBallonAjuste.Text) >= CInt(txtPlazo.Text) Then
    txtBallonAjuste.Text = 1
End If


If fxCrd_Factor_Calculo(cboFactor.Text) = "03" Then
    txtCuota.Text = Format(fxCrdCuotaNivelada(txtMonto, txtPlazo, txtTasa, DateAdd("m", 1, vFecha)), "Standard")
Else
    txtCuota.Text = Format(fxCalcula_Cuota(txtMonto, txtPlazo, txtTasa, mFrecuenciaPago), "Standard")
End If
 
 curCuota = CCur(txtCuota.Text)
 curTasa = CCur(txtTasa.Text) / 100
 curSaldo = txtMonto
 curIntx = 0
 curAmortiza = 0
  
  
 vGrid.MaxRows = 0
 txtTotalIntereses.Text = "0.00"
 txtTermina.Text = ""
 
 If mFrecuenciaPago = "Q" Then
    pQuincena = "_Q1"
    pDias = 15
 Else
    pQuincena = ""
    pDias = 30
 End If
 
 For i = 1 To txtPlazo
    
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    vGrid.col = 1
    
    If pQuincena = "" Then
        vFecha = DateAdd("m", 1, vFecha)
    End If
    
    If pQuincena = "_Q1" Then
        vFecha = DateAdd("m", 1, vFecha)
    End If
    
    Select Case Month(vFecha)
      Case 1
          vGrid.Text = "Enero de " & Year(vFecha) & pQuincena
      Case 2
          vGrid.Text = "Febrero de " & Year(vFecha) & pQuincena
      Case 3
          vGrid.Text = "Marzo de " & Year(vFecha) & pQuincena
      Case 4
          vGrid.Text = "Abril de " & Year(vFecha) & pQuincena
      Case 5
          vGrid.Text = "Mayo de " & Year(vFecha) & pQuincena
      Case 6
          vGrid.Text = "Junio de " & Year(vFecha) & pQuincena
      Case 7
          vGrid.Text = "Julio de " & Year(vFecha) & pQuincena
      Case 8
          vGrid.Text = "Agosto de " & Year(vFecha) & pQuincena
      Case 9
          vGrid.Text = "Septiembre de " & Year(vFecha) & pQuincena
      Case 10
          vGrid.Text = "Octubre de " & Year(vFecha) & pQuincena
      Case 11
          vGrid.Text = "Noviembre de " & Year(vFecha) & pQuincena
      Case 12
          vGrid.Text = "Diciembre de " & Year(vFecha) & pQuincena
    End Select
    
    vGrid.col = 2
    vGrid.Text = Format(curCuota, "Standard")
    
    Select Case fxCrd_Factor_Calculo(cboFactor.Text)
    
    Case "01", "06"
        vGrid.col = 3
        vGrid.Text = curSaldo * curTasa * pDias / 360
        curIntx = curIntx + CCur(vGrid.Text)
        
        vGrid.col = 4
        vGrid.Text = curCuota - (curSaldo * curTasa * pDias / 360)
        curAmortiza = curAmortiza + CCur(vGrid.Text)
    
        curSaldo = curSaldo - CCur(vGrid.Text)
        vGrid.col = 5
        vGrid.Text = curSaldo
        
        vGrid.col = 6
        vGrid.Text = pDias
        
        If pQuincena <> "" Then
            If pQuincena = "_Q1" Then
               pQuincena = "_Q2"
            Else
               pQuincena = "_Q1"
            End If
        End If
     
     Case Else
       '365 / 360 , Nivelado y Bullet
         iDias = fxMesDias(Month(vFecha), Year(vFecha))
        
        If vGrid.Row = CInt(txtPlazo.Text) Then
          curCuota = curSaldo + (curSaldo * curTasa * iDias / 360)
        End If
        
        vGrid.col = 3
        vGrid.Text = curSaldo * curTasa * iDias / 360
        curIntx = curIntx + CCur(vGrid.Text)
        
        vGrid.col = 4
        vGrid.Text = curCuota - (curSaldo * curTasa * iDias / 360)
        curAmortiza = curAmortiza + CCur(vGrid.Text)
        
        curSaldo = curSaldo - CCur(vGrid.Text)
        vGrid.col = 5
        vGrid.Text = curSaldo
        
        vGrid.col = 6
        vGrid.Text = iDias
        
        'Recalcula Cuota al Plazo Restante
        If vGrid.Row < CInt(txtPlazo.Text) And vGrid.Row > 1 Then
         curCuota = fxCalcula_Cuota(CDbl(curSaldo), CInt(txtPlazo) - vGrid.Row, txtTasa)
        End If
        
        If vGrid.Row = CInt(txtPlazo.Text) And vGrid.Row > 1 Then
         curCuota = fxCalcula_Cuota(CDbl(curSaldo), 1, txtTasa)
        End If
    
    End Select
    
      
    
 
 Next i
 
vGrid.col = 1
txtTotalIntereses.Text = Format(curIntx, "Standard")
txtTermina.Text = vGrid.Text


'Bullet Datos Iniciales
If txtBallonCta.Tag = "0" Then
  txtBallonAjuste.Text = 1
  If cboFactor.Text = "365 / 360 Ballon" Then
    txtBallonCta.Text = CCur(txtMonto.Text) * curTasa * 31 / 360
    txtBallonCta.Tag = CCur(txtMonto.Text) * curTasa * 31 / 360
  Else
    txtBallonCta.Text = CCur(txtMonto.Text) * curTasa * 30 / 360
    txtBallonCta.Tag = CCur(txtMonto.Text) * curTasa * 30 / 360
  End If
  txtBallonCta.ToolTipText = "Cuota Mínima :" & Format(txtBallonCta.Text, "Standard")
End If




curSaldo = txtMonto.Text
If cboFactor.Text = "365 / 360 Nivelado" Then
'' 'Cuota Promedio
'' For i = 1 To vGrid.MaxRows
''   vGrid.Row = i
''   vGrid.Col = 2
''   curCuota = curCuota + CCur(vGrid.Text)
''  Next i
''  curCuota = curCuota / vGrid.MaxRows
  'Cuota al 30%  de Avance
 vGrid.Row = CLng(vGrid.MaxRows / 3)
 vGrid.col = 2
 curCuota = CCur(vGrid.Text)
 
 'Cambios
 For i = 1 To vGrid.MaxRows
   vGrid.Row = i
   vGrid.col = 2
   vGrid.Text = curCuota
   
   vGrid.col = 6
   iDias = CInt(vGrid.Text)
   
   If i = CInt(txtPlazo.Text) Then 'Igual al Plazo Restante o Ultima Cuota
    vGrid.col = 3
    TempInt = curSaldo * curTasa * iDias / 360
    curCuota = curSaldo + TempInt
    
    TempAmort = curSaldo
    curSaldo = curSaldo - TempAmort
    
    vGrid.col = 2
    vGrid.Text = curCuota
    
    
    vGrid.col = 3
    vGrid.Text = TempInt
    vGrid.col = 4
    vGrid.Text = TempAmort
    vGrid.col = 5
    vGrid.Text = curSaldo
   
   Else
    vGrid.col = 2
    vGrid.Text = curCuota
    
    TempInt = curSaldo * curTasa * iDias / 360
    TempAmort = curCuota - TempInt
    curSaldo = curSaldo - TempAmort
    
    vGrid.col = 3
    vGrid.Text = TempInt
    vGrid.col = 4
    vGrid.Text = TempAmort
    vGrid.col = 5
    vGrid.Text = curSaldo
   
   End If
 Next i
 
 
End If 'Nivelado



'Ajuste de la Cuota con Bullet
curSaldo = txtMonto
   
'Verificacion de Cuota Minima
vGrid.Row = 1
vGrid.col = 6
iDias = CInt(vGrid.Text)
TempInt = curSaldo * curTasa * iDias / 360
If TempInt > CCur(txtBallonCta.Text) Then
  txtBallonCta.Text = Format(TempInt, "Standard")
  txtBallonCta.Tag = TempInt
  txtBallonCta.ToolTipText = "Cuota Mínima :" & Format(TempInt, "Standard")
Else
  If cboFactor.Text = "360 / 360 Bullet" Then
    TempInt = curSaldo * curTasa * 30 / 360
    txtBallonCta.Text = Format(TempInt, "Standard")
    txtBallonCta.Tag = TempInt
    txtBallonCta.ToolTipText = "Cuota Mínima :" & Format(TempInt, "Standard")
  End If
End If

'Procesa el Ajuste BULLET
If InStr(cboFactor.Text, "Balloon") > 0 Then

 For i = 1 To vGrid.MaxRows
   vGrid.Row = i
   vGrid.col = 2
   
   'Revision de los Intereses
   If i > 1 Then
        vGrid.col = 6
        iDias = CInt(vGrid.Text)
        vGrid.col = 3
        vGrid.Text = curSaldo * curTasa * iDias / 360
   End If
   
   If i >= (CInt(txtPlazo.Text) - CInt(txtBallonAjuste.Text) + 1) Then
    vGrid.col = 3
    curCuota = (curSaldo / (CInt(txtPlazo.Text) - i + 1)) + CCur(vGrid.Text)
    
    
    vGrid.col = 2
    vGrid.Text = curCuota
    
    
    vGrid.col = 3
    TempAmort = curCuota - CCur(vGrid.Text)
    vGrid.col = 4
    vGrid.Text = TempAmort
    
    curSaldo = curSaldo - TempAmort
    
    
    vGrid.col = 5
    vGrid.Text = curSaldo
   
   Else
    vGrid.col = 3
    txtBallonCta = vGrid.Text
    
    vGrid.col = 2
    vGrid.Text = txtBallonCta.Text
    
    
    
    vGrid.col = 3
    TempAmort = 0 'CCur(txtBallonCta.Text) - CCur(vGrid.Text)
    vGrid.col = 4
    vGrid.Text = TempAmort
    
    curSaldo = curSaldo - TempAmort
    
    vGrid.col = 5
    vGrid.Text = curSaldo
   
   End If
 Next i

End If 'BULLET


curIntx = 0
 For i = 1 To vGrid.MaxRows
   vGrid.Row = i
   vGrid.col = 3
   curIntx = curIntx + CCur(vGrid.Text)
 Next i

 txtTotalIntereses.Text = Format(curIntx, "Standard")


Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxValida()

fxValida = True

If Trim(txtMonto) = "" Or Not IsNumeric(txtMonto) Then fxValida = False
If Trim(txtPlazo) = "" Or Not IsNumeric(txtPlazo) Then fxValida = False
If Trim(txtTasa) = "" Or Not IsNumeric(txtTasa) Then fxValida = False

vGrid.MaxRows = 0
txtTotalIntereses.Text = "0.00"
txtTermina.Text = ""
 
End Function



Private Sub btnCalcular_Click()
Call sbCalculoCuota
End Sub

Private Sub btnReporte_Click()
Dim vTitulo As String, vPie As String

On Error GoTo vError


    Me.MousePointer = vbHourglass
    
    
    vPie = GLOBALES.gstrNombreEmpresa & vbTab & vbTab & "[Plan de Pagos] [Factor Tiempo : " & cboFactor.Text & "]"
    
    vTitulo = "Monto ...:" & txtMonto.Text & Space(5) _
            & "Plazo ...:" & txtPlazo.Text & Space(5) _
            & "Tasa  ...:" & txtTasa.Text & Space(5) _
            & "Cuota ...:" & txtCuota.Text & Space(5) _
            & "Termina .:" & txtTermina.Text
    
    
    vGrid.PrintColor = True
    vGrid.PrintFooter = vPie
    vGrid.PrintHeader = vTitulo
    vGrid.PrintOrientation = PrintOrientationPortrait
    vGrid.PrintSheet
      

vError:

    Me.MousePointer = vbDefault

End Sub

Private Sub cboFactor_Click()
 
 chkBallons.Enabled = False
 mFrecuenciaPago = "M"
 
 Select Case fxCrd_Factor_Calculo(cboFactor.Text)
    Case "04", "05" 'Ballon
        chkBallons.Enabled = True
    Case "06" 'Quincenal
         mFrecuenciaPago = "Q"
 End Select
 Call sbCalculoCuota
End Sub


Private Sub chkBallons_Click()
If chkBallons.Value = vbChecked Then
   gbBallons.Visible = True
Else
   gbBallons.Visible = False
End If
End Sub

Private Sub Form_Load()

vModulo = 3
vGrid.AppearanceStyle = fxGridStyle

Call sbCrd_Factor_Calculo(cboFactor)

mFrecuenciaPago = "M"

vGrid.MaxCols = 6

End Sub



Private Sub txtBallonCta_GotFocus()
On Error GoTo vError
  txtBallonCta.Text = CCur(txtBallonCta)
vError:
End Sub


Private Sub txtBallonCta_LostFocus()
On Error GoTo vError
  txtBallonCta.Text = Format(CCur(txtBallonCta.Text), "Standard")
  
  If CCur(txtBallonCta.Text) < CCur(txtBallonCta.Tag) Then
     MsgBox "La cuota Bullet no puede ser inferior a los intereses de la primer cuota ordinaria...!", vbExclamation
     txtBallonCta.Text = Format(txtBallonCta.Tag)
     Exit Sub
  End If
vError:
End Sub

Private Sub txtCuota_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTotalIntereses.SetFocus
End Sub

Private Sub txtTasa_GotFocus()
On Error GoTo vError
  txtTasa.Text = CCur(txtTasa.Text)
vError:
End Sub

Private Sub txtTasa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTotalIntereses.SetFocus
End Sub

Private Sub txtTasa_LostFocus()
On Error GoTo vError
  txtBallonAjuste.Text = 1
  txtBallonCta.Text = 0
  txtBallonCta.Tag = 0
  txtBallonCta.ToolTipText = ""
   Call sbCalculoCuota
vError:

End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError
  txtMonto = CCur(txtMonto)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlazo.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
  txtMonto.Text = Format(CCur(txtMonto.Text), "Standard")
  txtBallonAjuste.Text = 1
  txtBallonCta.Text = 0
  txtBallonCta.Tag = 0
  txtBallonCta.ToolTipText = ""
  
  
  Call sbCalculoCuota
vError:
End Sub

Private Sub txtPlazo_GotFocus()
On Error GoTo vError
  txtPlazo = CInt(txtPlazo)
vError:
End Sub

Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTasa.SetFocus
End Sub

Private Sub txtPlazo_LostFocus()
On Error GoTo vError
  txtPlazo = CInt(txtPlazo)
  
  txtBallonAjuste.Text = 1
  txtBallonCta.Text = 0
  txtBallonCta.Tag = 0
  txtBallonCta.ToolTipText = ""
  Call sbCalculoCuota
vError:
End Sub
