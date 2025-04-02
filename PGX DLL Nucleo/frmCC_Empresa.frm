VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSIF_Empresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Empresa Principal del Sistema"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "frmCC_Empresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Empresa"
      TabPicture(0)   =   "frmCC_Empresa.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "img1(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line1(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line1(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtEmail"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtAptoPostal"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtFax"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtTelefono"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCedJur"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtNombre"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cbo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Pagaré"
      TabPicture(1)   =   "frmCC_Empresa.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPagNomLargo"
      Tab(1).Control(1)=   "txtPagNomCorto"
      Tab(1).Control(2)=   "txtPagCedJur"
      Tab(1).Control(3)=   "txtPagDomicilio"
      Tab(1).Control(4)=   "Line1(2)"
      Tab(1).Control(5)=   "Label3(1)"
      Tab(1).Control(6)=   "img1(1)"
      Tab(1).Control(7)=   "Label1(6)"
      Tab(1).Control(8)=   "Label1(7)"
      Tab(1).Control(9)=   "Label1(8)"
      Tab(1).Control(10)=   "Label1(9)"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Estado de Cuenta"
      TabPicture(2)   =   "frmCC_Empresa.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkVisibleFianzas"
      Tab(2).Control(1)=   "chkVisibleCreditos"
      Tab(2).Control(2)=   "chkVisibleFondos"
      Tab(2).Control(3)=   "chkVisiblePatrimonio"
      Tab(2).Control(4)=   "txtEC_PiePagina"
      Tab(2).Control(5)=   "txtEC_Encabezado"
      Tab(2).Control(6)=   "chkEstadoCuenta"
      Tab(2).Control(7)=   "Label1(12)"
      Tab(2).Control(8)=   "Label1(11)"
      Tab(2).Control(9)=   "Line1(3)"
      Tab(2).Control(10)=   "Label3(2)"
      Tab(2).Control(11)=   "img1(2)"
      Tab(2).ControlCount=   12
      Begin VB.CheckBox chkVisibleFianzas 
         Appearance      =   0  'Flat
         Caption         =   "Visualizar Sección de Fianzas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72000
         TabIndex        =   36
         Top             =   4560
         Width           =   5295
      End
      Begin VB.CheckBox chkVisibleCreditos 
         Appearance      =   0  'Flat
         Caption         =   "Visualizar Sección de Créditos y Retenciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72000
         TabIndex        =   35
         Top             =   4320
         Width           =   5295
      End
      Begin VB.CheckBox chkVisibleFondos 
         Appearance      =   0  'Flat
         Caption         =   "Visualizar Sección de Fondos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72000
         TabIndex        =   34
         Top             =   4080
         Width           =   5295
      End
      Begin VB.CheckBox chkVisiblePatrimonio 
         Appearance      =   0  'Flat
         Caption         =   "Visualizar Sección de Patrimonio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72000
         TabIndex        =   33
         Top             =   3840
         Width           =   5295
      End
      Begin VB.TextBox txtEC_PiePagina 
         Height          =   1155
         Left            =   -72240
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   2280
         Width           =   5295
      End
      Begin VB.TextBox txtEC_Encabezado 
         Height          =   1155
         Left            =   -72240
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   1080
         Width           =   5295
      End
      Begin VB.CheckBox chkEstadoCuenta 
         Appearance      =   0  'Flat
         Caption         =   "Utilizar Estado de Cuenta Comercial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -72240
         TabIndex        =   24
         Top             =   3480
         Width           =   5295
      End
      Begin VB.TextBox txtPagNomLargo 
         Height          =   315
         Left            =   -72360
         TabIndex        =   19
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox txtPagNomCorto 
         Height          =   315
         Left            =   -72360
         TabIndex        =   18
         Top             =   1560
         Width           =   5295
      End
      Begin VB.TextBox txtPagCedJur 
         Height          =   315
         Left            =   -72360
         TabIndex        =   17
         Top             =   1920
         Width           =   5295
      End
      Begin VB.TextBox txtPagDomicilio 
         Height          =   1155
         Left            =   -72360
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   2280
         Width           =   5295
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3720
         Width           =   5295
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox txtCedJur 
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   1560
         Width           =   5295
      End
      Begin VB.TextBox txtTelefono 
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtFax 
         Height          =   315
         Left            =   5520
         TabIndex        =   4
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtAptoPostal 
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         Top             =   2280
         Width           =   5295
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Top             =   2640
         Width           =   5295
      End
      Begin VB.Label Label1 
         Caption         =   "Pie de Página"
         Height          =   255
         Index           =   12
         Left            =   -73560
         TabIndex        =   32
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Encabezado"
         Height          =   255
         Index           =   11
         Left            =   -73560
         TabIndex        =   30
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   -74880
         X2              =   -66960
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Notas para Estados de Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   -73680
         TabIndex        =   28
         Top             =   240
         Width           =   6135
      End
      Begin VB.Image img1 
         Height          =   720
         Index           =   2
         Left            =   -74760
         Picture         =   "frmCC_Empresa.frx":035E
         Top             =   120
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   -74880
         X2              =   -66960
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Datos Base para Pagares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   -73680
         TabIndex        =   27
         Top             =   240
         Width           =   6135
      End
      Begin VB.Image img1 
         Height          =   720
         Index           =   1
         Left            =   -74760
         Picture         =   "frmCC_Empresa.frx":37E0
         Top             =   120
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   1200
         X2              =   8040
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contabilidad de Enlace"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   1200
         TabIndex        =   26
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   120
         X2              =   8040
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Datos de la Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1320
         TabIndex        =   25
         Top             =   240
         Width           =   6135
      End
      Begin VB.Image img1 
         Height          =   720
         Index           =   0
         Left            =   240
         Picture         =   "frmCC_Empresa.frx":6C62
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre Largo"
         Height          =   255
         Index           =   6
         Left            =   -73680
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre Corto"
         Height          =   255
         Index           =   7
         Left            =   -73680
         TabIndex        =   22
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Ced.Jur. Letras"
         Height          =   255
         Index           =   8
         Left            =   -73680
         TabIndex        =   21
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   9
         Left            =   -73680
         TabIndex        =   20
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Contabilidad"
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ced.Jur."
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   12
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Apt.Postal"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   11
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Fax"
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   10
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail"
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   9
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   8
         Top             =   1920
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      Picture         =   "frmCC_Empresa.frx":A0E4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   1095
   End
End
Attribute VB_Name = "frmSIF_Empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean

Private Sub cmdGuardar_Click()
Dim strSQL As String

On Error GoTo vError

If vEdita Then
   strSQL = "update sif_empresa set nombre = '" & txtNombre & "',email = '" & txtEMail & "', apto_postal = '" _
          & txtAptoPostal & "',cedula_juridica = '" & txtCedJur & "',telefonoemp = '" & txtTelefono & "',fax = '" _
          & txtFax & "',EstadoCuenta = '" & IIf((chkEstadoCuenta.Value = vbChecked), "C", "S") _
          & "',PAG_NOMLARGO = '" & txtPagNomLargo & "',pag_NomCorto = '" & txtPagNomCorto _
          & "',PAG_CedJurLE = '" & txtPagCedJur & "',PAG_Domicilio = '" & txtPagDomicilio _
          & "',cod_empresa_enlace = " & cbo.ItemData(cbo.ListIndex) & ",ec_Nota01 = '" & txtEC_Encabezado _
          & "',ec_Nota02 = '" & txtEC_PiePagina & "', ec_Visible_Patrimonio = " & chkVisiblePatrimonio.Value _
          & ",ec_visible_creditos = " & chkVisibleCreditos.Value & ",ec_visible_fondos = " & chkVisibleFondos.Value _
          & ",ec_visible_fianzas = " & chkVisibleFianzas.Value
Else
   strSQL = "insert into sif_empresa(nombre,cedula_juridica,email," _
          & "telefonoemp,fax,apto_postal,cod_empresa_enlace,EstadoCuenta,pag_nomLargo,pag_nomCorto" _
          & ",pag_CedJurLe,pag_Domicilio,EC_Nota01,EC_Nota02,Ec_visible_creditos,Ec_visible_patrimonio" _
          & ",Ec_visible_fondos,Ec_visible_fianzas) values('" _
          & txtNombre & "','" & txtCedJur & "','" & txtEMail & "','" _
          & txtTelefono & "','" & txtFax & "','" & txtAptoPostal _
          & "'," & cbo.ItemData(cbo.ListIndex) & ",'" _
          & IIf((chkEstadoCuenta.Value = vbChecked), "C", "S") & "','" & txtPagNomLargo & "','" & txtPagNomCorto _
          & "','" & txtPagCedJur & "','" & txtPagDomicilio & "','" & txtEC_Encabezado & "','" & txtEC_PiePagina _
          & "'," & chkVisibleCreditos.Value & "," & chkVisiblePatrimonio.Value & "," & chkVisibleFondos.Value _
          & "," & chkVisibleFianzas.Value & ")"
End If

glogon.Conection.Execute strSQL
GLOBALES.gEnlace = cbo.ItemData(cbo.ListIndex)

MsgBox "Infomación Guardada Satisfactoriamente...", vbInformation

Unload Me
Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 10

Call Formularios(Me)
Call RefrescaTags(Me)

SSTab.Tab = 0

'Carga Combo de Contabilidades
strSQL = "select * from empresas"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
  rs.Close
  MsgBox "No Existe ninguna contabilidad creada para enlazar el sistema...", vbExclamation
  Unload Me
End If
Do While Not rs.EOF
  cbo.AddItem IIf(IsNull(rs!Nombre), "", rs!Nombre)
  cbo.ItemData(cbo.NewIndex) = rs!cod_empresa
  rs.MoveNext
Loop
rs.MoveFirst
cbo.Text = IIf(IsNull(rs!Nombre), "", rs!Nombre)
rs.Close


If GLOBALES.gEnlace = 0 Then
   
   vEdita = False
      
   
Else 'Editar
  vEdita = True
  strSQL = "select * from sif_empresa"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  
  txtNombre = rs!Nombre & ""
  txtCedJur = rs!cedula_juridica & ""
  txtTelefono = rs!telefonoemp & ""
  txtFax = rs!fax & ""
  txtAptoPostal = rs!apto_postal & ""
  txtEMail = rs!email
  
  txtPagNomLargo = rs!pag_nomlargo & ""
  txtPagNomCorto = rs!pag_nomCorto & ""
  txtPagCedJur = rs!pag_cedJurLe & ""
  txtPagDomicilio = rs!pag_domicilio & ""
  
  chkEstadoCuenta.Value = IIf((rs!estadoCuenta = "S"), vbUnchecked, vbChecked)
    
  chkVisibleCreditos.Value = rs!ec_visible_creditos
  chkVisiblePatrimonio.Value = rs!ec_visible_patrimonio
  chkVisibleFondos.Value = rs!ec_visible_fondos
  chkVisibleFianzas.Value = rs!ec_visible_fianzas
    
  txtEC_Encabezado = rs!EC_Nota01 & ""
  txtEC_PiePagina = rs!EC_Nota02 & ""
  
  
  strSQL = "select nombre from empresas where cod_empresa = " & rs!cod_empresa_enlace
  rs.Close
  rs.Open strSQL, glogon.Conection, adOpenStatic
  
  cbo.Text = rs!Nombre
  
  rs.Close

End If

End Sub




