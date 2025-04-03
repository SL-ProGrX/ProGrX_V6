VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCO_Gestones_Masivas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gesiones Masivas de Cobros"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15480
   LinkTopic       =   "frmCO_Gestiones_Masivas"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   15480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5895
      Left            =   11520
      TabIndex        =   21
      Top             =   1200
      Width           =   3735
      _Version        =   1572864
      _ExtentX        =   6588
      _ExtentY        =   10398
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   135
      Left            =   3960
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   6855
      _Version        =   1572864
      _ExtentX        =   12091
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin VB.ComboBox cboOperacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3240
      Width           =   1815
   End
   Begin XtremeSuiteControls.DateTimePicker dtpVence 
      Height          =   330
      Left            =   1560
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3196
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   2415
      Left            =   1560
      TabIndex        =   10
      Top             =   3840
      Width           =   4455
      _Version        =   1572864
      _ExtentX        =   7858
      _ExtentY        =   4260
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtGestion 
      Height          =   330
      Left            =   1560
      TabIndex        =   11
      Top             =   1200
      Width           =   855
      _Version        =   1572864
      _ExtentX        =   1508
      _ExtentY        =   582
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtGestionDesc 
      Height          =   330
      Left            =   2400
      TabIndex        =   12
      Top             =   1200
      Width           =   3495
      _Version        =   1572864
      _ExtentX        =   6165
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCausa 
      Height          =   330
      Left            =   1560
      TabIndex        =   13
      Top             =   1560
      Width           =   855
      _Version        =   1572864
      _ExtentX        =   1508
      _ExtentY        =   582
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCausaDesc 
      Height          =   330
      Left            =   2400
      TabIndex        =   14
      Top             =   1560
      Width           =   3495
      _Version        =   1572864
      _ExtentX        =   6165
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtArreglo 
      Height          =   330
      Left            =   1560
      TabIndex        =   15
      Top             =   1920
      Width           =   855
      _Version        =   1572864
      _ExtentX        =   1508
      _ExtentY        =   582
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtArregloDesc 
      Height          =   330
      Left            =   2400
      TabIndex        =   16
      Top             =   1920
      Width           =   3495
      _Version        =   1572864
      _ExtentX        =   6165
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtGestionMonto 
      Height          =   330
      Left            =   1560
      TabIndex        =   17
      Top             =   2280
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   582
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox fraLista 
      Height          =   6015
      Left            =   6120
      TabIndex        =   18
      Top             =   1080
      Width           =   5295
      _Version        =   1572864
      _ExtentX        =   9340
      _ExtentY        =   10610
      _StockProps     =   79
      Caption         =   "Gestion"
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.ListView lswLista 
         Height          =   5295
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   4935
         _Version        =   1572864
         _ExtentX        =   8705
         _ExtentY        =   9340
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtListaFiltro 
         Height          =   330
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   4935
         _Version        =   1572864
         _ExtentX        =   8705
         _ExtentY        =   582
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
   End
   Begin XtremeSuiteControls.PushButton cmdAplica 
      Height          =   615
      Left            =   4080
      TabIndex        =   22
      Top             =   6360
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3408
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
      Picture         =   "frmCO_Gestones_Masivas.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   435
      Left            =   3960
      TabIndex        =   23
      Top             =   360
      Width           =   6855
      _Version        =   1572864
      _ExtentX        =   12086
      _ExtentY        =   762
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
      Alignment       =   2
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   375
      Index           =   0
      Left            =   10920
      TabIndex        =   24
      Top             =   360
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmCO_Gestones_Masivas.frx":07D8
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   375
      Index           =   1
      Left            =   11400
      TabIndex        =   25
      Top             =   360
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmCO_Gestones_Masivas.frx":0ED8
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   375
      Index           =   2
      Left            =   11880
      TabIndex        =   26
      Top             =   360
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmCO_Gestones_Masivas.frx":15F1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   27
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gestión"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Pago"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "a la que se le va a registrar el recargo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3480
      TabIndex        =   4
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Causas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Acuerdo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   15735
   End
End
Attribute VB_Name = "frmCO_Gestones_Masivas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean, vTipoGestion As String


Private Sub cmdAplica_Click()
Dim vMensaje As String, mAccesoRestringido As Boolean
Dim i As Long

On Error GoTo vError


If lsw.ListItems.Count = 0 Then
    MsgBox "Carga la Lista de Cédulas antes de proceder!", vbExclamation
    Exit Sub
End If



Me.MousePointer = vbHourglass

'Verifica datos
vMensaje = ""

'If txtEstado.Tag = "N" Then vMensaje = vMensaje & " - La persona no se encuentra morosa verifique..." & vbCrLf
If txtNotas.Text = "" Then vMensaje = vMensaje & " - No se especificó ninguna observación..." & vbCrLf

strSQL = "select isnull(count(*),0) as Existe from cbr_usuarios where usuario = '" _
       & glogon.Usuario & "' and estado = 1"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & " - El usuario actual no se encuentra activo..." & vbCrLf
rs.Close

strSQL = "select isnull(count(*),0) as Existe from cbr_gestiones where cod_gestion = '" _
       & txtGestion.Text & "' and estado = 1 and NIVEL_GESTION = 'U'"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & " - La gestion actual no se encuentra activa..." & vbCrLf
rs.Close



'------No Aplica en Masivos
''Preguntar si existe el parametro de sgt sin asignacion previa, de lo contrario buscar asignacion
'strSQL = "select valor from cbr_parametros where cod_parametro = '05'"
'Call OpenRecordSet(rs, strSQL)
'If Mid(rs!Valor, 1, 1) <> "S" Then
'  rs.Close
'  strSQL = "select isnull(count(*),0) as Existe from cbr_asignacion where usuario = '" _
'       & glogon.Usuario & "' and cedula = '" & txtCedula & "'"
'  Call OpenRecordSet(rs, strSQL)
'  If rs!Existe = 0 Then vMensaje = vMensaje & " - Este expediente no se encuentra asignado al usuario actual, verifique..." & vbCrLf
'End If
'rs.Close

If vMensaje <> "" Then
  Me.MousePointer = vbDefault
  MsgBox vMensaje, vbExclamation
  Exit Sub
End If


With lsw.ListItems
    PrgBar.Max = .Count + 1
    PrgBar.Visible = True
    
    strSQL = ""
    
    For i = 1 To .Count
        PrgBar.Value = i
        
        strSQL = strSQL & Space(10) & "exec spCBRControlSGT '" & .Item(i).Text & "','" & glogon.Usuario & "','" & txtGestion.Text _
               & "','" & Format(dtpVence.Value, "yyyy/mm/dd") & "','" & txtNotas & "','" & GLOBALES.gOficinaTitular _
               & "'," & CCur(txtGestionMonto.Text) & ", 0, '" & txtCausa.Text & "','" & txtArreglo.Text & "'"
        
        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
    Next i
    
        'Lote Final
        If Len(strSQL) > 0 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
    
    PrgBar.Visible = False

End With

Me.MousePointer = vbDefault

MsgBox "Seguimiento Registrado Satisfactoriamente...", vbInformation

lsw.ListItems.Clear

Exit Sub

vError:
  PrgBar.Visible = False
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswLista_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
  
  Select Case Mid(vTipoGestion, 1, 1)
    Case "G"
        txtGestion.Text = Item.Text
        txtGestionDesc.Text = Item.SubItems(1)
        
        If Right(vTipoGestion, 1) = "C" Then
           txtGestion.SetFocus
        Else
           txtGestionDesc.SetFocus
        End If
       
    Case "C"
        txtCausa.Text = Item.Text
        txtCausaDesc.Text = Item.SubItems(1)
        
        If Right(vTipoGestion, 1) = "C" Then
           txtCausa.SetFocus
        Else
           txtCausaDesc.SetFocus
        End If
        
    Case "A"
        txtArreglo.Text = Item.Text
        txtArregloDesc.Text = Item.SubItems(1)
        
        If Right(vTipoGestion, 1) = "C" Then
           txtArreglo.SetFocus
        Else
           txtArregloDesc.SetFocus
        End If
        
  End Select
  

  
End Sub



Private Sub sbCargaLista(vTipoGestion As String, Optional vFiltro As String = "")

Me.MousePointer = vbHourglass

On Error GoTo vError

If vTipoGestion = "" Then Exit Sub

txtListaFiltro.Text = vFiltro

Select Case Mid(vTipoGestion, 1, 1)
   
   Case "G" 'Consulta de gestiones
     fraLista.Caption = "Gestiones"
     strSQL = "Select COD_GESTION as 'Codigo',DESCRIPCION from CBR_GESTIONES" _
            & " where ESTADO = 1 and  NIVEL_GESTION = 'U'"
    
        If vFiltro <> "" Then
            If Right(vTipoGestion, 1) = "C" Then
               strSQL = strSQL & " and COD_GESTION like '%" & txtListaFiltro.Text & "%' order by COD_GESTION"
            Else
               strSQL = strSQL & " and DESCRIPCION like '%" & txtListaFiltro.Text & "%' order by DESCRIPCION"
            End If
        End If
            
    Case "C"  'Consulta de Causas de Mora
      fraLista.Caption = "Causas de Mora"
      strSQL = "Select COD_CAUSA as 'Codigo',DESCRIPCION from CBR_CAUSAS_MOROSIDAD" _
             & " where ACTIVA = 1"
      If vFiltro <> "" Then
         If Right(vTipoGestion, 1) = "C" Then
            strSQL = strSQL & " and COD_CAUSA like '%" & txtListaFiltro.Text & "%' order by COD_CAUSA"
         Else
            strSQL = strSQL & " and DESCRIPCION like '%" & txtListaFiltro.Text & "%' order by DESCRIPCION"
         End If
      End If
            
    Case "A" 'Consulta de Tipos de Arreglos
      fraLista.Caption = "Arreglos"
      strSQL = "Select COD_ARREGLO as 'Codigo',DESCRIPCION from CBR_TIPOS_ARREGLOS" _
             & " where ACTIVO = 1"
      
      If vFiltro <> "" Then
         If Right(vTipoGestion, 1) = "C" Then
          strSQL = strSQL & " and COD_ARREGLO like '%" & txtListaFiltro.Text & "%' order by COD_CAUSA"
         Else
          strSQL = strSQL & " and DESCRIPCION like '%" & txtListaFiltro.Text & "%' order by DESCRIPCION"
         End If
      End If
      
End Select

If Right(vTipoGestion, 1) = "C" Then
    fraLista.Caption = fraLista.Caption & " [Código]"
Else
    fraLista.Caption = fraLista.Caption & " [Descripción]"
End If

Call OpenRecordSet(rs, strSQL)
     
lswLista.ListItems.Clear
     
Do While Not rs.EOF
  Set itmX = lswLista.ListItems.Add(, , Trim(rs!Codigo))
      itmX.SubItems(1) = rs!Descripcion
  rs.MoveNext
Loop

rs.Close

Me.MousePointer = vbDefault


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Sub Form_Load()

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

With lswLista.ColumnHeaders
    .Clear
    .Add , , "ID", 640
    .Add , , "Detalle", 3850
End With


With lsw.ColumnHeaders
    .Clear
    .Add , , "Cédula", 2640
End With

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub txtArreglo_GotFocus()

If vTipoGestion <> "AC" Then
 vTipoGestion = "AC"
 Call sbCargaLista(vTipoGestion)
End If
 
End Sub

Private Sub txtArreglo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtArregloDesc.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_ARREGLO,DESCRIPCION from CBR_TIPOS_ARREGLOS"
    gBusquedas.Columna = "COD_ARREGLO"
    gBusquedas.Orden = "COD_ARREGLO"
    gBusquedas.Filtro = " and ACTIVO = 1 "
    frmBusquedas.Show vbModal
    txtArreglo = Trim(gBusquedas.Resultado)
    txtArregloDesc = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtCausa_GotFocus()
If vTipoGestion <> "CC" Then
 vTipoGestion = "CC"
 Call sbCargaLista(vTipoGestion)
End If
End Sub

Private Sub txtCausa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtCausaDesc.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_CAUSA,DESCRIPCION from CBR_CAUSAS_MOROSIDAD"
    gBusquedas.Columna = "COD_CAUSA"
    gBusquedas.Orden = "COD_CAUSA"
    gBusquedas.Filtro = " and ACTIVA = 1  "
    frmBusquedas.Show vbModal
    txtCausa = Trim(gBusquedas.Resultado)
    txtCausaDesc = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtArregloDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtGestionMonto.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_ARREGLO,DESCRIPCION from CBR_TIPOS_ARREGLOS"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "COD_ARREGLO"
    gBusquedas.Filtro = " and ACTIVO = 1 "
    frmBusquedas.Show vbModal
    txtArreglo = Trim(gBusquedas.Resultado)
    txtArregloDesc = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtCausaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtArreglo.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_CAUSA,DESCRIPCION from CBR_CAUSAS_MOROSIDAD"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "COD_CAUSA"
    gBusquedas.Filtro = " and ACTIVA = 1  "
    frmBusquedas.Show vbModal
    txtCausa = Trim(gBusquedas.Resultado)
    txtCausaDesc = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtGestion_GotFocus()
If vTipoGestion <> "GC" Then
 vTipoGestion = "GC"
 Call sbCargaLista(vTipoGestion)
End If
End Sub

Private Sub txtGestion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtGestionDesc.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select cod_gestion,descripcion from cbr_gestiones"
    gBusquedas.Columna = "cod_gestion"
    gBusquedas.Orden = "cod_gestion"
    gBusquedas.Filtro = " and estado = 1 and nivel_gestion = 'U' "
    frmBusquedas.Show vbModal
    txtGestion.Text = Trim(gBusquedas.Resultado)
    txtGestionDesc.Text = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtGestionDesc_GotFocus()
If vTipoGestion <> "GD" Then
 vTipoGestion = "GD"
 Call sbCargaLista(vTipoGestion)
End If
End Sub

Private Sub txtGestionDesc_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtCausa.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select cod_gestion,descripcion from cbr_gestiones"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Filtro = " and estado = 1 and nivel_gestion = 'U' "
    frmBusquedas.Show vbModal
    txtGestion = Trim(gBusquedas.Resultado)
    txtGestionDesc = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtGestionDesc_LostFocus()
'    Call sbCBRControlGestion
End Sub

Private Sub txtGestionMonto_LostFocus()

    If txtGestionMonto.Text = Empty Then
        txtGestionMonto.Text = Format(0, "Standard")
    End If
    
    If Not IsNumeric(txtGestionMonto) Then
        txtGestionMonto.Text = Format(0, "Standard")
    End If
    
    txtGestionMonto = Format(txtGestionMonto, "Standard")
    
'    If txtGestionMonto < vDesviacionMin Then
'        MsgBox "El monto es menor que la desviación mínima"
'        txtGestionMonto = Format(vDesviacionMin, "Standard")
'        txtGestionMonto.SetFocus
'    End If
'
'    If txtGestionMonto > vDesviacionMax Then
'        MsgBox "El monto es mayor que la desviación máxima"
'        txtGestionMonto = Format(vDesviacionMax, "Standard")
'        txtGestionMonto.SetFocus
'    End If
    
End Sub

