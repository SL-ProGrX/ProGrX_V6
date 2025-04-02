VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmTES_BancosGrupos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos Bancarios"
   ClientHeight    =   8424
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   12924
   Icon            =   "frmTES_BancosGrupos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8424
   ScaleWidth      =   12924
   Begin XtremeSuiteControls.GroupBox gbFirmas 
      Height          =   2892
      Left            =   2520
      TabIndex        =   2
      Top             =   5520
      Width           =   7812
      _Version        =   1245187
      _ExtentX        =   13779
      _ExtentY        =   5101
      _StockProps     =   79
      Caption         =   "Firma No.1 "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnFirma_Guarda 
         Height          =   492
         Left            =   5400
         TabIndex        =   3
         Top             =   1680
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmTES_BancosGrupos.frx":000C
      End
      Begin XtremeSuiteControls.PushButton btnFirma_Buscar 
         Height          =   492
         Left            =   5400
         TabIndex        =   4
         Top             =   2280
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmTES_BancosGrupos.frx":0703
      End
      Begin XtremeSuiteControls.FlatEdit txtImagenLogo 
         Height          =   492
         Left            =   1080
         TabIndex        =   5
         Top             =   2280
         Width           =   4212
         _Version        =   1245187
         _ExtentX        =   7429
         _ExtentY        =   868
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton rbFirmas 
         Height          =   372
         Index           =   0
         Left            =   5640
         TabIndex        =   6
         Top             =   240
         Width           =   2532
         _Version        =   1245187
         _ExtentX        =   4466
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Registro de Firma No 1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbFirmas 
         Height          =   372
         Index           =   1
         Left            =   5640
         TabIndex        =   7
         Top             =   600
         Width           =   2532
         _Version        =   1245187
         _ExtentX        =   4466
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Registro de Firma No 2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin VB.Image picImagen 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1932
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   240
         Width           =   4212
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3612
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   12732
      _Version        =   524288
      _ExtentX        =   22458
      _ExtentY        =   6371
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
      MaxCols         =   492
      ScrollBars      =   2
      SpreadDesigner  =   "frmTES_BancosGrupos.frx":1121
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption lblBanco 
      Height          =   372
      Left            =   120
      TabIndex        =   8
      Top             =   5040
      Width           =   12732
      _Version        =   1245187
      _ExtentX        =   22458
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Seleccione un Grupo Bancario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grupos Bancarios"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmTES_BancosGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnFirma_Buscar_Click()
        
On Error GoTo vError

        frmContenedor.CD.ShowOpen
        frmContenedor.CD.DialogTitle = "Buscar Imagen de Firma..."
        frmContenedor.CD.InitDir = "C:\"
        
        
        
        txtImagenLogo.Text = frmContenedor.CD.FileName
        If txtImagenLogo.Text <> "" Then
            picImagen.Picture = LoadPicture(frmContenedor.CD.FileName)
        End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation


End Sub

Private Sub btnFirma_Guarda_Click()
Dim strSQL As String, pFirma As String

If lblBanco.Tag = "" Or txtImagenLogo.Text = "" Then Exit Sub

strSQL = "select * from Tes_Bancos_Grupos where cod_Grupo = '" & lblBanco.Tag & "'"

If gbFirmas.Tag = "1" Then
  pFirma = "Firma_N1"
Else
  pFirma = "Firma_N2"
End If

If fxImagen_Guardar(strSQL, pFirma, txtImagenLogo.Text) Then
   MsgBox gbFirmas.Caption & ", guardada satisfactoriamente!", vbInformation
End If
    
End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 9
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select '',cod_grupo,desc_Corta,ID_SFN,descripcion,LCta_Interna, LCta_Interbancaria, TCta_UTiliza,Activo from Tes_Bancos_Grupos" _
      & " order by cod_grupo"
      
vPaso = True
Call sbCargaGrid(vGrid, 9, strSQL)
vPaso = False

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 2

strSQL = "select isnull(count(*),0) as Existe from Tes_Bancos_Grupos " _
       & " where cod_grupo = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into Tes_Bancos_Grupos(cod_grupo,desc_Corta,ID_SFN,descripcion,LCta_Interna, LCta_Interbancaria, TCta_UTiliza" _
         & ",Activo,Registro_Usuario,Registro_Fecha) values('" _
         & vGrid.Text & "','"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.col = 4
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.col = 5
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 6
  strSQL = strSQL & vGrid.Text & ","
  vGrid.col = 7
  strSQL = strSQL & vGrid.Text & ",'"
  vGrid.col = 8
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 9
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
  
  Call ConectionExecute(strSQL)

  vGrid.col = 2
  Call Bitacora("Registra", "Grupos Bancarios : " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 3
 strSQL = "update Tes_Bancos_Grupos set  desc_corta = '" & vGrid.Text & "',ID_SFN = '"
 vGrid.col = 4
 strSQL = strSQL & vGrid.Text & "',Descripcion = '"
 vGrid.col = 5
 strSQL = strSQL & vGrid.Text & "',LCta_Interna = "
 vGrid.col = 6
 strSQL = strSQL & vGrid.Text & ",LCta_Interbancaria = "
 vGrid.col = 7
 strSQL = strSQL & vGrid.Text & ",TCta_UTiliza = '"
 vGrid.col = 8
 strSQL = strSQL & vGrid.Text & "', Activo = "
 vGrid.col = 9
 strSQL = strSQL & vGrid.Value & " where cod_grupo = '"
 vGrid.col = 2
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.col = 2
 Call Bitacora("Modifica", "Grupos Bancarios : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub rbFirmas_Click(Index As Integer)
Dim strSQL As String

strSQL = "select Firma_N1, Firma_N2 from Tes_Bancos_Grupos where cod_grupo ='" & lblBanco.Tag & "'"

txtImagenLogo.Text = ""

Select Case Index
   Case 0 'Firma 1
      gbFirmas.Tag = 1
      Set picImagen.Picture = fxImagen_Leer(strSQL, "Firma_N1")
   Case 1 'Firma 2
      gbFirmas.Tag = 2
      Set picImagen.Picture = fxImagen_Leer(strSQL, "Firma_N2")
End Select

gbFirmas.Caption = "Registro de Firmas No." & gbFirmas.Tag

End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub

vGrid.Row = Row
vGrid.col = 2

If vGrid.Text <> "" Then
  lblBanco.Tag = vGrid.Text
  vGrid.col = 5
  lblBanco.Caption = vGrid.Text
  rbFirmas.Item(0).Value = True
  Call rbFirmas_Click(0)
End If


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 2
        strSQL = "delete Tes_Bancos_Grupos where cod_grupo = " & vGrid.Text
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.col = 2
        Call Bitacora("Elimina", "Grupos Bancarios : " & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
