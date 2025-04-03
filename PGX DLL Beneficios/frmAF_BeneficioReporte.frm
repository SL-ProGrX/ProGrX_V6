VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmAF_BeneficioReporte 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Beneficios"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8130
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1212
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   7932
      _Version        =   1441793
      _ExtentX        =   13991
      _ExtentY        =   2138
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdGenerar 
         Height          =   612
         Left            =   6240
         TabIndex        =   18
         Top             =   360
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   -2147483633
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
         Picture         =   "frmAF_BeneficioReporte.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.ComboBox cboTipoReporte 
         Height          =   315
         Left            =   4560
         TabIndex        =   19
         Top             =   360
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   1680
      TabIndex        =   7
      Top             =   2880
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   1680
      TabIndex        =   8
      Top             =   1440
      Width           =   4932
      _Version        =   1441793
      _ExtentX        =   8705
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboFecha 
      Height          =   312
      Left            =   1680
      TabIndex        =   9
      Top             =   1800
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboUsuario 
      Height          =   312
      Left            =   1680
      TabIndex        =   10
      Top             =   2520
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   3360
      TabIndex        =   11
      Top             =   2520
      Width           =   3252
      _Version        =   1441793
      _ExtentX        =   5736
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   5160
      TabIndex        =   12
      Top             =   1800
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   5160
      TabIndex        =   13
      Top             =   2160
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.CheckBox chkBeneficios 
      Height          =   372
      Left            =   6840
      TabIndex        =   14
      Top             =   1440
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Todos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.CheckBox chkFecha 
      Height          =   372
      Left            =   6840
      TabIndex        =   15
      Top             =   1800
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Todos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.CheckBox chkUsuario 
      Height          =   492
      Left            =   6840
      TabIndex        =   16
      Top             =   2400
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Todos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes del Beneficios y Ayudas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Index           =   1
      Left            =   1560
      TabIndex        =   6
      Top             =   360
      Width           =   5892
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   720
      TabIndex        =   5
      Top             =   2880
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   4200
      TabIndex        =   3
      Top             =   2160
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   4200
      TabIndex        =   2
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Beneficio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   972
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmAF_BeneficioReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset

Private Sub chkBeneficios_Click()
If chkBeneficios.Value Then
   cbo.Enabled = False
Else
   cbo.Enabled = True
End If
End Sub

Private Sub chkFecha_Click()
If chkFecha.Value Then
   cboFecha.Enabled = False
   dtpInicio.Enabled = False
   dtpCorte.Enabled = False
Else
   dtpInicio.Enabled = True
   dtpCorte.Enabled = True
   cboFecha.Enabled = True
   
End If
End Sub

Private Sub chkUsuario_Click()
If chkUsuario.Value Then
   cboUsuario.Enabled = False
   txtUsuario.Enabled = False
Else
   cboUsuario.Enabled = True
   txtUsuario.Enabled = True
End If
End Sub

Private Sub cmdGenerar_Click()

Dim vDetalle As String

vDetalle = UCase(cboTipoReporte.Text) & " :"

     If chkBeneficios.Value = 0 Then
            strSQL = "{AFI_BENE_OTORGA.COD_BENEFICIO} = '" & cbo.ItemData(cbo.ListIndex) & "'"
            vDetalle = vDetalle & "Beneficio Id: " & cbo.ItemData(cbo.ListIndex)
     Else
        strSQL = Empty
        vDetalle = vDetalle & " Todos los Beneficios"
     End If
     
     
     
     If chkFecha.Value = xtpUnchecked Then
        vDetalle = vDetalle & " ¦ Fechas: " & Format(dtpInicio, "mm/dd/yyyy") & " - " & Format(dtpCorte, "mm/dd/yyyy")
        
        If strSQL <> Empty And strSQL <> " and" Then strSQL = strSQL & " and "
        If UCase(Mid(cboFecha.Text, 1, 1)) = "R" Then
        'En caso de que seleccione la fecha de registro
            
            strSQL = strSQL & "cdate({afi_bene_otorga.registra_fecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
            strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
        Else
        'En caso de que seleccione la fecha de autorización
            strSQL = strSQL & "cdate({afi_bene_otorga.autoriza_fecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
            strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
            
         End If
        
     Else
      vDetalle = vDetalle & " ¦ Todas las Fechas"
      strSQL = strSQL & Empty
     End If
    
   
    
     If chkUsuario.Value = xtpUnchecked Then
         
      If txtUsuario <> Empty Then
         vDetalle = vDetalle & " ¦ Usuario: " & txtUsuario.Text
          If strSQL <> Empty And strSQL <> "and" Then strSQL = strSQL & " and "
            If UCase(Mid(cboUsuario.Text, 1, 1)) = "R" Then
            ' En caso de que seleccione usuario que registra
                strSQL = strSQL & "{afi_bene_otorga.registra_user} = '" & Trim(txtUsuario) & "'"
            Else
            ' En caso de que seleccione usario que autoriza
                strSQL = strSQL & "{afi_bene_otorga.autoriza_user} = '" & Trim(txtUsuario) & "'"
            End If
            
        Else
           MsgBox "Seleccione un usuario", vbInformation
           Exit Sub
        End If
    Else
            vDetalle = vDetalle & " ¦ Todos los Usuarios"
    End If
        
    
  
    
    If Mid(cboEstado.Text, 2, 1) <> "T" Then
          vDetalle = vDetalle & " ¦ Estado: " & cboEstado.Text
          If strSQL <> Empty And strSQL <> "and" Then strSQL = strSQL & " and "
        strSQL = strSQL & "{afi_bene_otorga.estado} = '" & Mid(cboEstado.Text, 1, 1) & "'"
    Else
       vDetalle = vDetalle & " ¦ Todos los Estados"
    End If
    
     
     
    
    
       With frmContenedor.Crt
            .Reset
            .WindowShowGroupTree = True
            .WindowShowPrintSetupBtn = True
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            .WindowState = crptMaximized
            .WindowTitle = "Reportes del Módulo de Beneficios y Ayudas Sociales"
            .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
            .Formulas(2) = "Usuario='" & glogon.Usuario & "'"
            
            .Formulas(3) = "DETALLE = '" & vDetalle & "'"
            
            .Connect = glogon.ConectRPT
            
            If Mid(cboTipoReporte.Text, 1, 1) = "D" Then
                .ReportFileName = SIFGlobal.fxPathReportes("Beneficios_Beneficios.rpt")
            Else
                .ReportFileName = SIFGlobal.fxPathReportes("Beneficios_Beneficios_Resumen.rpt")
            End If
            
            .SelectionFormula = strSQL
            .PrintReport

        End With
        
        
End Sub

Private Sub Form_Activate()
vModulo = 7
End Sub

Private Sub Form_Load()
vModulo = 7

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

Call sbInicializa
 
End Sub


Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Or KeyCode = vbKeyReturn Then
        
    gBusquedas.Resultado2 = ""
    gBusquedas.Resultado = Trim(txtUsuario)
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    gBusquedas.Consulta = "Select Nombre,Descripcion From usuarios"
    gBusquedas.Filtro = " and Estado = 'A'"
    frmBusquedas.Show vbModal
    
    txtUsuario.Text = Trim(gBusquedas.Resultado)
End If

End Sub



Private Sub sbInicializa()


strSQL = "select rtrim(cod_Beneficio) as 'IdX', rtrim(descripcion) as 'ItmX' from afi_beneficios" _
              & " where estado = 'A'"
   
Call sbCbo_Llena_New(cbo, strSQL, False, True)
   
   cboFecha.Clear
   cboFecha.AddItem "REGISTRO"
   cboFecha.AddItem "AUTORIZA"
   cboFecha.Text = "REGISTRO"

   cboUsuario.Clear
   cboUsuario.AddItem "REGISTRA"
   cboUsuario.AddItem "AUTORIZA"
   cboUsuario.Text = "REGISTRA"

  cboEstado.Clear
  cboEstado.AddItem "[TODOS]"
  cboEstado.AddItem "APROBADO"
  cboEstado.AddItem "SOLICITADO"
  cboEstado.AddItem "RECHAZADO"
  cboEstado.AddItem "EJECUTADO"
  cboEstado.AddItem "PENDIENTE"
  cboEstado.Text = "[TODOS]"
  
  cboTipoReporte.Clear
  cboTipoReporte.AddItem "Detalle"
  cboTipoReporte.AddItem "Resumen"
  cboTipoReporte.Text = "Detalle"
  
  
  dtpCorte.Value = fxFechaServidor
  dtpInicio.Value = DateAdd("m", -1, dtpCorte.Value)
  
  chkBeneficios.Value = xtpChecked
  chkUsuario.Value = xtpChecked
  
  Call chkBeneficios_Click
  Call chkUsuario_Click

End Sub
