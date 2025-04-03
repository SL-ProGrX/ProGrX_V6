VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFNDRenuevaContratos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Renovación de Contratos Fondos"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   10260
   Begin XtremeSuiteControls.GroupBox gbTitle01 
      Height          =   852
      Index           =   0
      Left            =   360
      TabIndex        =   15
      Top             =   2280
      Width           =   9612
      _Version        =   1572864
      _ExtentX        =   16954
      _ExtentY        =   1503
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   570
         Left            =   8040
         TabIndex        =   17
         Top             =   240
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   1005
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   16777215
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
         Appearance      =   21
         Picture         =   "frmFNDRenuevaContratos.frx":0000
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   252
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
         BackColor       =   16777215
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
         Appearance      =   21
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkRenueva 
         Height          =   252
         Left            =   2520
         TabIndex        =   19
         Top             =   240
         Width           =   3852
         _Version        =   1572864
         _ExtentX        =   6794
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "&Solo los que Renuevan"
         BackColor       =   16777215
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
         Value           =   1
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3732
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Width           =   9732
      _Version        =   524288
      _ExtentX        =   17166
      _ExtentY        =   6583
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
      MaxCols         =   6
      ScrollBars      =   2
      SpreadDesigner  =   "frmFNDRenuevaContratos.frx":0A1E
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   132
      Left            =   0
      TabIndex        =   5
      Top             =   8400
      Visible         =   0   'False
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   312
      Left            =   2520
      TabIndex        =   6
      Top             =   240
      Width           =   6492
      _Version        =   1572864
      _ExtentX        =   11456
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2520
      TabIndex        =   7
      Top             =   600
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3840
      TabIndex        =   8
      Top             =   600
      Width           =   5172
      _Version        =   1572864
      _ExtentX        =   9123
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPlanDestino 
      Height          =   312
      Left            =   2520
      TabIndex        =   11
      Top             =   960
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescDestino 
      Height          =   312
      Left            =   3840
      TabIndex        =   12
      Top             =   960
      Width           =   5172
      _Version        =   1572864
      _ExtentX        =   9123
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   2760
      TabIndex        =   13
      Top             =   1920
      Width           =   2292
      _Version        =   1572864
      _ExtentX        =   4048
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.DateTimePicker dtpVence 
      Height          =   312
      Left            =   2760
      TabIndex        =   14
      Top             =   1560
      Width           =   2292
      _Version        =   1572864
      _ExtentX        =   4043
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.GroupBox gbTitle01 
      Height          =   1092
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   7200
      Width           =   9732
      _Version        =   1572864
      _ExtentX        =   17166
      _ExtentY        =   1926
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      BorderStyle     =   1
      Begin VB.TextBox txtMontoExisten 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtMontoAplica 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtCasosAplicar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtCasosExisten 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   612
         Left            =   8280
         TabIndex        =   26
         Top             =   240
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Aplicar"
         BackColor       =   16777215
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
         Appearance      =   21
         Picture         =   "frmFNDRenuevaContratos.frx":1126
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Existe"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1212
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Aplicar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1212
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtPlazo 
      Height          =   312
      Left            =   6600
      TabIndex        =   27
      Top             =   1560
      Width           =   852
      _Version        =   1572864
      _ExtentX        =   1503
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Origen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   1320
      TabIndex        =   10
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   0
      Left            =   1320
      TabIndex        =   9
      Top             =   240
      Width           =   1332
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Plazo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   5880
      TabIndex        =   4
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha vencimiento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   2
      Left            =   720
      TabIndex        =   3
      Top             =   1560
      Width           =   1884
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado de la Persona"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   3
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   1764
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1092
   End
   Begin VB.Image imgBanner 
      Height          =   1335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmFNDRenuevaContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnBuscar_Click()
Dim strSQL As String

Me.MousePointer = vbHourglass

txtMontoAplica.Text = 0
txtMontoExisten.Text = 0
txtCasosExisten.Text = 0
txtCasosAplicar.Text = 0


If Trim(txtCodigo.Text) = "" Or txtPlanDestino.Text = "" Then
    MsgBox "Falta definir el plan de origen o destino...", vbCritical
    txtCodigo.SetFocus
Else
    strSQL = "Select " & chkTodos.Value & " as 'Aplicar',F.Cedula,S.nombre,F.monto" _
            & ",dbo.fxFndExisteContratoPersona(F.cedula," & cboOperadora.ItemData(cboOperadora.ListIndex) & ",'" & txtPlanDestino.Text & "') as 'Existe'" _
            & ", F.cod_Contrato" _
            & " from fnd_contratos F inner join Socios S on F.cedula = s.cedula" _
            & " where F.estado= 'A' and  F.cod_plan = '" & txtCodigo & "' and F.cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)

    If chkRenueva.Value = vbChecked Then
       strSQL = strSQL & " and F.renueva = 'S' "
    End If
    
    If cboEstado.Text <> "TODOS" Then
          strSQL = strSQL & " and S.EstadoActual = '" & cboEstado.ItemData(cboEstado.ListIndex) & "'"
    End If
    
   
    Call sbCargaGridLocal(vGrid, 6, strSQL)

End If

Me.MousePointer = vbDefault

End Sub

Private Sub cboOperadora_Change()
If vPaso Then Exit Sub
End Sub

Private Sub chkTodos_Click()
Dim i As Long
Dim curExiste As Currency, curMarcados As Currency
Dim iExiste As Long, iMarcados As Long

If vPaso Then Exit Sub

curMarcados = 0
curExiste = 0
iExiste = 0
iMarcados = 0

With vGrid
    For i = 1 To .MaxRows
       .Row = i
       .Col = 1
       .Value = chkTodos.Value
       If .Value = vbChecked Then
          .Col = 4
          curMarcados = curMarcados + CCur(.Text)
          iMarcados = iMarcados + 1
       
            .Col = 5
            If .Value = vbChecked Then
               .Col = 4
               curExiste = curExiste + CCur(.Text)
               iExiste = iExiste + 1
            End If
       
       End If
       
    Next i
End With

txtMontoAplica.Text = Format(curMarcados, "Standard")
txtCasosAplicar.Text = Format(iMarcados, "###,###,##0")

txtMontoExisten.Text = Format(curExiste, "Standard")
txtCasosExisten.Text = Format(iExiste, "###,###,##0")

End Sub

Private Sub chkRenueva_Click()
If vPaso Then Exit Sub
End Sub

Private Sub CmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim iContrato As Long, i As Long, vContratoActual As Long
Dim strCedula As String, curMonto As Currency, iExiste As Integer

On Error GoTo vError

If vGrid.MaxRows <= 0 Then Exit Sub

PrgBar.Max = vGrid.MaxRows + 1
PrgBar.Value = 1

PrgBar.Visible = True


strSQL = ""
For i = 1 To vGrid.MaxRows

    vGrid.Row = i
    vGrid.Col = 2
    strCedula = vGrid.Text
    
    vGrid.Col = 4
    curMonto = vGrid.Text
    vGrid.Col = 5
    iExiste = vGrid.Value
    vGrid.Col = 6
    vContratoActual = vGrid.Text
    
    vGrid.Col = 1
    If vGrid.Value = vbChecked And iExiste = 0 Then
    
        strSQL = strSQL & Space(10) & "exec spFnd_RenuevaContratos " & cboOperadora.ItemData(cboOperadora.ListIndex) _
               & ",'" & txtCodigo.Text & "','" & txtPlanDestino.Text & "'," & vContratoActual _
               & "," & txtPlazo.Text & ",'" & Format(dtpVence.Value, "yyyy/mm/dd") _
               & "','" & glogon.Usuario & "','" & GLOBALES.gOficinaTitular & "'"
        
        
        If Len(strSQL) > 20000 Then
           Call ConectionExecute(strSQL)
           strSQL = ""
        End If
        PrgBar.Value = PrgBar.Value + 1
    End If

Next i

'Procesa Lote Restante
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If


MsgBox "Se generaron los fondos satisfactoritamente..."
PrgBar.Visible = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub dtpVence_Change()
txtPlazo.Text = DateDiff("M", fxFechaServidor, dtpVence.Value)
If Val(txtPlazo.Text) <= 0 Then
  MsgBox "Fecha de vencimiento invalida..."
  dtpVence.Value = DateAdd("m", 1, Format(fxFechaServidor, "dd/mm/yyyy"))
End If
End Sub

Private Sub dtpVence_Click()
txtPlazo.Text = DateDiff("M", fxFechaServidor, dtpVence.Value)
If Val(txtPlazo.Text) <= 0 Then
  MsgBox "Fecha de vencimiento invalida..."
  dtpVence.Value = DateAdd("m", 1, Format(fxFechaServidor, "dd/mm/yyyy"))
End If
End Sub

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 18 'Fondo de Inversion

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True
vGrid.MaxRows = 0
dtpVence.Value = DateAdd("m", 1, Format(fxFechaServidor, "dd/mm/yyyy"))

Call dtpVence_Click

strSQL = "select rtrim(descripcion) as 'ItmX',cod_operadora as 'IdX' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)


strSQL = "select rtrim(COD_ESTADO) as 'IdX' , RTRIM(DESCRIPCION) as 'ItmX'" _
       & " From AFI_ESTADOS_PERSONA  Where ACTIVO = 1"
Call sbCbo_Llena_New(cboEstado, strSQL, True, True)

txtMontoAplica.Text = 0
txtMontoExisten.Text = 0

vPaso = False

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub txtPlanDestino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = "And estado = 'A' And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   cboEstado.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtPlanDestino = Trim(gBusquedas.Resultado)
      txtDescDestino = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = "And estado = 'A' And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub


Public Sub sbCargaGridLocal(pGrid As Object, MaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer


vPaso = True
PrgBar.Visible = True

With pGrid
    .MaxRows = 0
    .MaxCols = MaxCol
    
    Call OpenRecordSet(rs, strSQL)
    
    PrgBar.Max = rs.RecordCount + 1
    PrgBar.Value = 1

    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
       .Col = 1
       .Value = rs!Aplicar
       .Col = 2
       .Text = rs!Cedula
       .Col = 3
       .Text = rs!Nombre
       .Col = 4
       .Text = Format(rs!Monto, "Standard")
       .Col = 5
       
       If rs!Existe > 0 Then
         .Value = vbChecked
            txtCasosExisten.Text = txtCasosExisten.Text + 1
            txtMontoExisten.Text = Format(txtMontoExisten.Text + rs!Monto, "Standard")
         .Col = 1
         .Value = 0
       Else
         txtCasosAplicar.Text = txtCasosAplicar.Text + 1
         txtMontoAplica.Text = Format(txtMontoAplica.Text + rs!Monto, "Standard")
       End If
       
       .Col = 6
       .Text = rs!COD_Contrato
       
      PrgBar.Value = PrgBar.Value + 1
      rs.MoveNext
    Loop
    rs.Close

End With

PrgBar.Visible = False
vPaso = False

End Sub


Private Sub txtPlanDestino_LostFocus()
If Trim(txtCodigo) = Trim(txtPlanDestino) Then
   MsgBox "No se puede aplicar el mismo plan..."
   txtPlanDestino.Text = ""
   txtDescDestino.Text = ""
   txtPlanDestino.SetFocus
End If
End Sub



Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim i As Long
Dim curExiste As Currency, curMarcados As Currency
Dim iExiste As Long, iMarcados As Long

If vPaso Or Col > 1 Then Exit Sub

curMarcados = 0
curExiste = 0
iExiste = 0
iMarcados = 0

With vGrid
       .Row = Row
       .Col = 4
       curMarcados = curMarcados + CCur(.Text)
       iMarcados = iMarcados + 1
    
         .Col = 5
         If .Value = vbChecked Then
            .Col = 4
            curExiste = curExiste + CCur(.Text)
            iExiste = iExiste + 1
         End If
       
       
       .Col = 1
       If .Value = vbUnchecked Then
          curExiste = curExiste * -1
          iExiste = iExiste * -1
          
          curMarcados = curMarcados * -1
          iMarcados = iMarcados * -1
       End If

End With

txtMontoAplica.Text = Format(CCur(txtMontoAplica.Text) + curMarcados, "Standard")
txtCasosAplicar.Text = Format(CLng(txtCasosAplicar.Text) + iMarcados, "###,###,##0")

txtMontoExisten.Text = Format(CCur(txtMontoExisten.Text) + curExiste, "Standard")
txtCasosExisten.Text = Format(CLng(txtCasosExisten.Text) + iExiste, "###,###,##0")


End Sub
