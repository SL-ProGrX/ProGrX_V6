VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_Tipos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FOSOL: Tipo de aplicación"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   4815
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tipos"
      TabPicture(0)   =   "frmFSL_Tipos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGridTipos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Causas"
      TabPicture(1)   =   "frmFSL_Tipos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboTipo"
      Tab(1).Control(1)=   "vGrid"
      Tab(1).Control(2)=   "Label6"
      Tab(1).ControlCount=   3
      Begin VB.ComboBox cboTipo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68520
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   3255
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   2
         Top             =   960
         Width           =   9375
         _Version        =   524288
         _ExtentX        =   16536
         _ExtentY        =   6376
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         MaxCols         =   5
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Tipos.frx":0038
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridTipos 
         Height          =   4095
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   9015
         _Version        =   524288
         _ExtentX        =   15901
         _ExtentY        =   7223
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         MaxCols         =   4
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Tipos.frx":1D24
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label6 
         Caption         =   "Tipos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69000
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   11280
      X2              =   0
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Aplicación del Fondo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image3 
      Height          =   720
      Left            =   0
      Picture         =   "frmFSL_Tipos.frx":237E
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmFSL_Tipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim vCodigoCausa, VCodigoTipo As Integer
Dim vActiva, vEstadoTipo As Integer
Dim vDescripcionCausa, vDescripcionTipo As String, vFecha As String
Dim vCancelaTotal, vMontoFormalizacion As Integer
Dim vTabla As String
Dim vPaso As Boolean

Private Sub sbCargaCausas()
On Error GoTo vError
Dim i As Integer
   
If vPaso = True Then Exit Sub

strSQL = "Select COD_CAUSA, DESCRIPCION,MONTO_FORMALIZADO, SALDO_OPERACION, ACTIVA" _
       & " from FSL_CAUSAS where COD_PLAN = " & SIFGlobal.fxSIFCodText(cboTipo) & " order by ACTIVA desc"
       
rs.Open strSQL, glogon.Conection, adOpenDynamic
  
With vGrid

.MaxRows = 1
.Row = .MaxRows
For i = 1 To .MaxCols
 .Col = i
 .Text = ""
Next i

Do While Not rs.EOF
 .Row = .MaxRows
 .Col = 1
 .Text = rs!COD_CAUSA
 .Col = 2
 .Text = rs!Descripcion
 .Col = 3
 .Value = rs!MONTO_FORMALIZADO
 .Col = 4
 .Value = rs!SALDO_OPERACION
 .Col = 5
 .Value = rs!activa
 
 .MaxRows = .MaxRows + 1

 rs.MoveNext
Loop
.Row = .MaxRows
End With

rs.Close

Exit Sub
vError:
   MsgBox Err.Description

End Sub

Private Sub sbGuardaCausa()
   strSQL = "Insert FSL_CAUSAS (COD_CAUSA, COD_PLAN, DESCRIPCION, ACTIVA, MONTO_FORMALIZADO, SALDO_OPERACION" _
          & ", REGISTRO_FECHA, REGISTRO_USUARIO) " _
          & " values (" & vCodigoCausa & "," & SIFGlobal.fxSIFCodText(cboTipo) & ",'" & UCase(vDescripcionCausa) & "'," & vActiva & "" _
          & ", " & CInt(vMontoFormalizacion) & "," & CInt(vCancelaTotal) & ",'" & Format(vFecha, "yyyymmdd") & "','" & glogon.Usuario & "')"

   glogon.Conection.Execute strSQL
End Sub

Private Sub sbEliminaCausa()
On Error GoTo vError
   strSQL = "Delete FSL_CAUSAS where cod_causa=" & vCodigoCausa & " " _
          & "and COD_PLAN =" & Trim(SIFGlobal.fxSIFCodText(cboTipo)) & ""

   glogon.Conection.Execute strSQL

Exit Sub

vError:
   MsgBox Err.Description
End Sub

Private Sub sbModificaCausa()
On Error GoTo error

   strSQL = "update FSL_CAUSAS set COD_CAUSA = " & vCodigoCausa & ",COD_PLAN=" & SIFGlobal.fxSIFCodText(cboTipo) & ", " _
          & "DESCRIPCION='" & UCase(vDescripcionCausa) & "', MONTO_FORMALIZADO=" & vMontoFormalizacion & ", ACTIVA=" & vActiva & " " _
          & ",SALDO_OPERACION=" & vCancelaTotal & ", REGISTRO_USUARIO='" & glogon.Usuario & "', REGISTRO_FECHA='" & Format(fxFechaServidor, "yyyymmdd") & "' " _
          & "where cod_causa=" & vCodigoCausa & " and COD_PLAN =" & Trim(SIFGlobal.fxSIFCodText(cboTipo)) & ""

   glogon.Conection.Execute strSQL

Exit Sub

error:
   MsgBox Err.Description

End Sub

Private Sub sbCargaTipos()
On Error GoTo vError

cboTipo.Clear
    
strSQL = "select COD_PLAN,DESCRIPCION From FSL_PLANES_APLICACION Where ACTIVO = 1"
rs.Open strSQL, glogon.Conection, adOpenStatic

Do While Not rs.EOF
   cboTipo.AddItem (rs!cod_plan & " - " & Trim(rs!Descripcion))
   rs.MoveNext
Loop
    
If rs.RecordCount > 0 Then
   rs.MoveFirst
   vPaso = True
   cboTipo.Text = rs!cod_plan & " - " & rs!Descripcion
   vPaso = False
End If

rs.Close

Exit Sub
vError:
  MsgBox Err.Description

End Sub


Private Sub sbCargaGridTipo()
  Dim i As Integer
  
  With vGridTipos
    .MaxRows = 1
    .Row = .MaxRows
    For i = 1 To .MaxCols
     .Col = i
     .Text = ""
    Next i
      
    strSQL = "select COD_PLAN,DESCRIPCION,TABLA_ASIGNADA, ACTIVO " _
           & "From FSL_PLANES_APLICACION Where ACTIVO = 1"
         
    rs.Open strSQL, glogon.Conection, adOpenStatic
    
    Do While Not rs.EOF
      .Row = .MaxRows
      .Col = 1
      .Text = rs!cod_plan
      
      .Col = 2
      .Text = rs!Descripcion
      
      .Col = 3
      Select Case rs!TABLA_ASIGNADA
      
       Case "F"
          .Text = "Fallecimiento"
          
       Case "P"
           .Text = "Pension"
      
      End Select

      .Col = 4
      .Value = rs!Activo
    
      rs.MoveNext
      .MaxRows = .MaxRows + 1
    Loop
    
    rs.Close
  End With
  
End Sub

Private Sub cboTipo_Click()
  Call sbCargaCausas
End Sub

Private Sub Form_Activate()
  vModulo = 22
End Sub

Private Sub Form_Load()
  vModulo = 22
  vGridTipos.MaxRows = 1
  ssTab.Tab = 0
  Call sbCargaGridTipo
  vFecha = fxFechaServidor
  
End Sub

Private Function fxValidaCausa() As Boolean
  fxValidaCausa = True
  With vGrid
     .Row = .ActiveRow
     .Col = 1
     
     If .Text = Empty Then
         strSQL = "Select isnull(MAX(COD_CAUSA + 1), 1) as Codigo from FSL_CAUSAS"
         rs.Open strSQL, glogon.Conection, adOpenStatic
         
         vCodigoCausa = rs!Codigo
         rs.Close
         
        .Text = vCodigoCausa
         fxValidaCausa = False
     Else
         vCodigoCausa = .Text
     End If
     
     .Col = 2
      If .Text = Empty Then
        MsgBox "Falta la descripción de la Causa"
      Else
        vDescripcionCausa = .Text
      End If
      
      .Col = 3
      vMontoFormalizacion = .Value
      
      .Col = 4
      vCancelaTotal = .Value
      
      .Col = 5
      vActiva = CInt(.Text)
      
  End With
  
End Function

Private Sub SSTab_Click(PreviousTab As Integer)
  If ssTab.Tab = 1 Then
     vGrid.MaxRows = 1
     Call sbCargaTipos
     Call sbCargaCausas
  End If
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
  
On Error GoTo vError

  With vGrid
    .Row = .ActiveRow
    If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
      If fxValidaCausa = False Then
         Call sbGuardaCausa
         .MaxRows = .MaxRows + 1
         .Col = 1
         .Text = vCodigoCausa
      Else
         Call sbModificaCausa
      End If
    ElseIf KeyCode = vbKeyDelete Then
     If fxValidaCausa = True Then
        Call sbEliminaCausa
        Call sbCargaCausas
     End If
    End If
  
  End With
    
Exit Sub

vError:
  MsgBox Err.Description

End Sub

Private Sub vGridTipos_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

  With vGridTipos
    .Row = .ActiveRow
    If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
      If fxValidaTipo = False Then
        Call sbGuardaTipo
      Else
        Call sbModificaTipo
      End If
    ElseIf .ActiveCol = .MaxCols And (KeyCode = vbKeyDelete) Then
      If fxValidaTipo = True Then
        Call sbEliminaTipo
      End If
      Call sbCargaGridTipo
    End If
  End With
      
Exit Sub

vError:
  MsgBox Err.Description
End Sub

Private Function fxValidaTipo() As Boolean
On Error GoTo vError
  
  fxValidaTipo = True
  With vGridTipos
    .Row = .ActiveRow
    .Col = 1
    
    If .Text = Empty Then
        strSQL = "Select isnull(max(COD_PLAN) + 1,1) as CodigoTipo from FSL_PLANES_APLICACION "
        rs.Open strSQL, glogon.Conection, adOpenStatic
        
        VCodigoTipo = rs!CodigoTipo
        rs.Close
        
        .Text = VCodigoTipo
        fxValidaTipo = False
    Else
        VCodigoTipo = .Text
    End If
    
    .Col = 2
    vDescripcionTipo = .Text
    
    .Col = 3
    vTabla = Mid(.Text, 1, 1)
    
    .Col = 4
    vEstadoTipo = .Value
  
  End With
  
  Exit Function
  
vError:
  MsgBox Err.Description
  
End Function

Private Sub sbGuardaTipo()

On Error GoTo vError

    strSQL = "Insert FSL_PLANES_APLICACION (COD_PLAN, DESCRIPCION,TABLA_ASIGNADA ,ACTIVO, REGISTRO_FECHA, REGISTRO_USUARIO) " _
          & "values('" & Trim(VCodigoTipo) & "','" & UCase(Trim(vDescripcionTipo)) & "','" & vTabla & "'," & vEstadoTipo & ", " _
          & "'" & Format(vFecha, "yyyymmdd") & "','" & Trim(glogon.Usuario) & "')"
    glogon.Conection.Execute strSQL
      
    vGridTipos.MaxRows = vGridTipos.MaxRows + 1
      
    Exit Sub

vError:
  MsgBox Err.Description
  
End Sub

Private Sub sbModificaTipo()
On Error GoTo vError

    strSQL = "UPDATE FSL_PLANES_APLICACION SET COD_PLAN = '" & Trim(VCodigoTipo) & "',DESCRIPCION = '" & UCase(Trim(vDescripcionTipo)) & "', " _
           & "ACTIVO = " & vEstadoTipo & ",TABLA_ASIGNADA ='" & vTabla & "', REGISTRO_FECHA = '" & Format(vFecha, "yyyymmdd") & "',REGISTRO_USUARIO = '" & glogon.Usuario & "' " _
           & "Where COD_PLAN = '" & Trim(VCodigoTipo) & "'"
    
    glogon.Conection.Execute strSQL
      
    Exit Sub
    
vError:
  MsgBox Err.Description

End Sub

Private Sub sbEliminaTipo()
On Error GoTo vError

    strSQL = "DELETE FROM FSL_PLANES_APLICACION WHERE COD_PLAN = '" & Trim(VCodigoTipo) & "'"
    glogon.Conection.Execute strSQL
      
    Exit Sub
    
vError:
  MsgBox Err.Description
End Sub
