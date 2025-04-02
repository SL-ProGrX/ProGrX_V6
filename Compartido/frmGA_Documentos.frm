VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmGA_Documentos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestor de Archivos ¦ Documentos"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   13785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3975
      Left            =   0
      TabIndex        =   6
      Top             =   4080
      Width           =   13815
      _Version        =   1441793
      _ExtentX        =   24368
      _ExtentY        =   7011
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   12960
      Top             =   720
   End
   Begin XtremeSuiteControls.CheckBox chkVence 
      Height          =   255
      Left            =   5760
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Vence?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   390
      Left            =   1560
      TabIndex        =   8
      Top             =   1080
      Width           =   7215
      _Version        =   1441793
      _ExtentX        =   12726
      _ExtentY        =   688
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   435
      Left            =   1560
      TabIndex        =   9
      Top             =   1560
      Width           =   7215
      _Version        =   1441793
      _ExtentX        =   12726
      _ExtentY        =   767
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
      Left            =   8880
      TabIndex        =   10
      Top             =   1560
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmGA_Documentos.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnNuevo 
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   480
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Documento"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmGA_Documentos.frx":0700
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   495
      Left            =   9240
      TabIndex        =   15
      Top             =   3120
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Guardar"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmGA_Documentos.frx":0E20
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.DateTimePicker dtpVence 
      Height          =   345
      Left            =   7200
      TabIndex        =   16
      Top             =   2640
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   609
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpEmite 
      Height          =   345
      Left            =   7200
      TabIndex        =   18
      Top             =   2160
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   609
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtLlave_01 
      Height          =   345
      Left            =   1560
      TabIndex        =   11
      Top             =   2160
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLlave_02 
      Height          =   345
      Left            =   1560
      TabIndex        =   12
      Top             =   2640
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLlave_03 
      Height          =   345
      Left            =   1560
      TabIndex        =   13
      Top             =   3120
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label lblLoading 
      Height          =   1695
      Left            =   9360
      TabIndex        =   20
      Top             =   1320
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "xxxxxxxxxx"
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   8
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   19
      Top             =   2160
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Fecha Emisión?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scModulo 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13815
      _Version        =   1441793
      _ExtentX        =   24368
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Modulo General"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   3720
      Width           =   13815
      _Version        =   1441793
      _ExtentX        =   24368
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Documentos Registrados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Llave 03"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Llave 02"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Llave 01"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Archivo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmGA_Documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, cnPGA As String, vPaso As Boolean
Dim mFecha As Date, itmX As ListViewItem

Private Sub btnArchivo_Click()

With frmContenedor.CD
        
        .InitDir = "C:\"
        .DialogTitle = "Localice el Archivo.."
        .Filter = "*.*"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If

        txtArchivo.Text = .FileName
End With

End Sub

Function ObtenerExtension(ByVal RutaCompleta As String) As String
    Dim PuntoPosicion As Integer
    Dim NombreArchivo As String
    
    ' Extraer solo el nombre del archivo de la ruta completa
    NombreArchivo = Dir(RutaCompleta)
    
    ' Buscar la posición del último punto en el nombre del archivo
    PuntoPosicion = InStrRev(NombreArchivo, ".")
    
    ' Si se encontró un punto, extraer la extensión
    If PuntoPosicion > 0 Then
        ObtenerExtension = Right(NombreArchivo, Len(NombreArchivo) - PuntoPosicion + 1)
    Else
        ObtenerExtension = "" ' No se encontró extensión o el archivo no tiene extensión
    End If
End Function

Private Sub sbPGX_GA_Conection(pCn As ADODB.Connection)

If gGA.Empresa = 0 Then
 gGA.Empresa = gPortal.Empresa_Id
End If

If gGA.Conexion = "" Then
    gGA.Conexion = "PROVIDER=MSDASQL;Driver={SQL Server};Server=20.81.197.231" _
          & ";Database=PGX_GA;APP=PGX_Portal_Access"

End If

With pCn
  
  .CommandTimeout = 15
  .Mode = adModeReadWrite
  .CursorLocation = adUseClient
  
  .Open gGA.Conexion, "PGX_Interface", "f@M#5$f1lm4$KF*m1n0x0f."
  .CommandTimeout = 360
End With

End Sub


Private Function fxFileName_Valido(pNombre As String) As String

pNombre = Replace(pNombre, "(", "_")
pNombre = Replace(pNombre, ")", "_")

fxFileName_Valido = pNombre

End Function


Private Sub btnGuardar_Click()

'Validaciones
If txtArchivo.Text = "" Then
    MsgBox "Por favor, selecciona un archivo primero.", vbExclamation, "Error"
    Exit Sub
End If

txtLlave_01.Text = fxSysCleanTxtInject(txtLlave_01.Text)
txtLlave_02.Text = fxSysCleanTxtInject(txtLlave_02.Text)
txtLlave_03.Text = fxSysCleanTxtInject(txtLlave_03.Text)


If txtLlave_01.Text = "" Then
    MsgBox "Llave 1 es Obligatoria!", vbExclamation, "Error"
    Exit Sub
End If



On Error GoTo vError



Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

lblLoading.Caption = "Subiendo Archivo..Espere!"
DoEvents

Me.MousePointer = vbHourglass


Call sbPGX_GA_Conection(cn)


'Version 1
'    Dim objStream As New ADODB.Stream
'
'    objStream.Type = adTypeBinary
'    objStream.Open
'    objStream.LoadFromFile txtArchivo.Text
'
'
'    Set rs = New ADODB.Recordset
'    rs.Open "GA_Files", cn, adOpenDynamic, adLockOptimistic
'    rs.AddNew
'
'
'    rs.Fields("EmpresaId") = gGA.Empresa
'    rs.Fields("ModuloId") = gGA.Modulo
'    rs.Fields("TypeId") = cboTipo.ItemData(cboTipo.ListIndex)
'
'    rs.Fields("Llave_01") = txtLlave_01.Text
'    rs.Fields("Llave_02") = txtLlave_02.Text
'    rs.Fields("Llave_03") = txtLlave_03.Text
'
'    rs.Fields("FileType") = ObtenerExtension(txtArchivo.Text)
'    rs.Fields("FileName") = fxFileName_Valido(Dir(txtArchivo.Text, vbArchive))
'    rs.Fields("FileContent") = objStream.Read
'
'    If chkVence.Value = xtpChecked Then
'        rs.Fields("Vencimiento") = Format(dtpVence.Value, "yyyy-mm-dd") & " " & Format(dtpVence.Value, "hh:mm:ss")
'    End If
'
'    rs.Fields("FechaEmision") = Format(dtpEmite.Value, "yyyy-mm-dd") & " " & Format(dtpEmite.Value, "hh:mm:ss")
'
'    rs.Fields("RegistroFecha") = Format(mFecha, "yyyy-mm-dd") & " " & Format(mFecha, "hh:mm:ss")
'    rs.Fields("RegistroUsuario") = glogon.Usuario
'
'
'    rs.Update
'    rs.Close

'    objStream.Close
'    Set objStream = Nothing

'-------Fin Version 1


'Version 2
Dim Cmd As ADODB.Command
Dim fileData() As Byte

Dim pFVence As String, pFEmite As String, pFRegistro As String


If chkVence.Value = xtpChecked Then
    pFVence = Format(dtpVence.Value, "yyyy-mm-dd") & " " & Format(dtpVence.Value, "hh:mm:ss")
Else
    pFVence = Format(DateAdd("yyyy", 100, dtpVence.Value), "yyyy-mm-dd")
End If

pFEmite = Format(dtpEmite.Value, "yyyy-mm-dd") & " " & Format(dtpEmite.Value, "hh:mm:ss")
pFRegistro = Format(mFecha, "yyyy-mm-dd") & " " & Format(mFecha, "hh:mm:ss")

' Leer el contenido del archivo en un arreglo de bytes
Open txtArchivo.Text For Binary Access Read As #1
ReDim fileData(LOF(1) - 1)
Get #1, , fileData
Close #1


' Preparar la consulta SQL para insertar el archivo
Set Cmd = New ADODB.Command
Cmd.ActiveConnection = cn

Cmd.CommandText = "INSERT INTO GA_Files (EmpresaId, ModuloId, TypeId, Llave_01, Llave_02, Llave_03, FileType, FileName, FileContent" _
                & ", Vencimiento, FechaEmision, RegistroFecha, RegistroUsuario ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

Cmd.Parameters.Append Cmd.CreateParameter("@EmpresaId", adInteger, adParamInput, , gGA.Empresa)
Cmd.Parameters.Append Cmd.CreateParameter("@ModuloId", adVarChar, adParamInput, 10, gGA.Modulo)
Cmd.Parameters.Append Cmd.CreateParameter("@TypeId", adVarChar, adParamInput, 10, cboTipo.ItemData(cboTipo.ListIndex))

Cmd.Parameters.Append Cmd.CreateParameter("@Llave_01", adVarChar, adParamInput, 30, txtLlave_01.Text)
Cmd.Parameters.Append Cmd.CreateParameter("@Llave_02", adVarChar, adParamInput, 30, txtLlave_02.Text)
Cmd.Parameters.Append Cmd.CreateParameter("@Llave_03", adVarChar, adParamInput, 30, txtLlave_03.Text)

Cmd.Parameters.Append Cmd.CreateParameter("@FileType", adVarChar, adParamInput, 100, ObtenerExtension(txtArchivo.Text))
Cmd.Parameters.Append Cmd.CreateParameter("@FileName", adVarChar, adParamInput, 1000, fxFileName_Valido(Dir(txtArchivo.Text, vbArchive)))
Cmd.Parameters.Append Cmd.CreateParameter("@FileContent", adLongVarBinary, adParamInput, UBound(fileData) + 1, fileData)

Cmd.Parameters.Append Cmd.CreateParameter("@Vencimiento", adVarChar, adParamInput, 20, pFVence)
Cmd.Parameters.Append Cmd.CreateParameter("@FechaEmision", adVarChar, adParamInput, 20, pFEmite)

Cmd.Parameters.Append Cmd.CreateParameter("@RegistroFecha", adVarChar, adParamInput, 20, pFRegistro)
Cmd.Parameters.Append Cmd.CreateParameter("@RegistroUsuario", adVarChar, adParamInput, 30, glogon.Usuario)

Cmd.Execute


'-------Fin Version 2

cn.Close
Set cn = Nothing

    
Me.MousePointer = vbDefault

MsgBox "Archivo subido exitosamente.", vbInformation, "Éxito"
    
txtArchivo.Text = ""
lblLoading.Caption = ""

Call sbDocumentos_List

    
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    lblLoading.Caption = "Error en la carga!"
End Sub




Private Sub cboTipo_Click()
If vPaso Then Exit Sub
If cboTipo.ListCount = 0 Then Exit Sub


Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


Call sbPGX_GA_Conection(cn)


rs.CursorLocation = adUseClient
strSQL = "select dbo.fxGA_Interface_Tipos_Documento_Vence(" & gGA.Empresa & ", '" & glogon.Usuario _
       & "', '" & cboTipo.ItemData(cboTipo.ListIndex) & "') as 'Vence'"
rs.Open strSQL, cn, adOpenStatic, adLockReadOnly
If Not rs.EOF And Not rs.BOF Then
    If rs!Vence = 1 Then
        chkVence.Value = xtpChecked
    Else
        chkVence.Value = xtpUnchecked
    End If

End If

rs.Close
cn.Close

Call chkVence_Click

Me.MousePointer = vbDefault

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub chkVence_Click()

If chkVence.Value = xtpChecked Then
    dtpVence.Enabled = True
Else
    dtpVence.Enabled = False
End If


End Sub

Private Sub Form_Load()

mFecha = fxFechaServidor


lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Archivo Id", 1200
lsw.ColumnHeaders.Add , , "Tipo Adjunto", 3200
lsw.ColumnHeaders.Add , , "Nombre Archivo", 5200
lsw.ColumnHeaders.Add , , "Extensión", 1400, vbCenter
lsw.ColumnHeaders.Add , , "Vence?", 1400, vbCenter
lsw.ColumnHeaders.Add , , "Fecha", 1400, vbCenter
lsw.ColumnHeaders.Add , , "Usuario", 1400, vbCenter

lblLoading.Caption = ""

End Sub

Private Sub lsw_DblClick()
Dim cn As New ADODB.Connection

Dim sql As String, Campo_Imagen As String
Dim rs As New ADODB.Recordset, Stream As New ADODB.Stream
Dim pPath As String

Dim vArchivo As String

If lsw.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

Set itmX = lsw.SelectedItem

On Error Resume Next

vArchivo = "GA_" & Format(itmX.Text, "00000") & " [" & gGA.Empresa & "_" & gGA.Modulo & "] [" & txtLlave_01.Text & "]" _
        & " " & itmX.SubItems(1) & " - " & itmX.SubItems(2)


MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Adjuntos\"

pPath = SIFGlobal.DirectorioDeResultados & "\Adjuntos\" & vArchivo

'------------------------------------------------------------------------
  
On Error GoTo vError

Dim pPass As Boolean
  
lblLoading.Caption = "Abriendo archivo..Espere!"
DoEvents
  
Me.MousePointer = vbHourglass
  
pPass = False
Campo_Imagen = "FileContent"
  
sql = "select " & Campo_Imagen & " from GA_Files Where FileId = " & itmX.Text

Call sbPGX_GA_Conection(cn)

'---Version 2
Dim fileData() As Byte

Set rs = New ADODB.Recordset
rs.Open sql, cn, adOpenStatic, adLockReadOnly

' Guardar los datos del archivo en el disco
If Not rs.EOF Then
    fileData = rs.Fields("FileContent").Value
    Open pPath For Binary Access Write As #1
    Put #1, , fileData
    Close #1
    
    pPass = True
End If
rs.Close

'---Version 2: Fin

''---Version 1
'        rs.Open sql, cn, adOpenKeyset, adLockOptimistic
'
'        ' Si no hay registros sale de la función y retorna como _
'         resultado un valor Nothing, es decir ninguna imagen
'
'        If rs.RecordCount = 0 Then
'           Exit Sub
'        End If
'
'        ' Especifica el tipo de datos ( binario )
'        Stream.Type = adTypeBinary
'        Stream.Open
'
'        ' verifica con la función IsNull que el campo no tenga _
'         un valor Nulo ya que si no da error, en ese caso sale de la función
'        If IsNull(rs.Fields(Campo_Imagen).Value) Then
'            GoTo vError
'        End If
'        ' Graba los datos en el objeto stream
'        Stream.Write rs.Fields(Campo_Imagen).Value
'
'        ' este método graba un  archivo temporal  en disco _
'         ( en el pPath que luego se elimina )
'        Stream.SaveToFile pPath, adSaveCreateOverWrite
'
'        'Cierra el recordset y el objeto Stream
'        If rs.State = adStateOpen Then
'            rs.Close
'        End If
'        If Not rs Is Nothing Then
'            Set rs = Nothing
'        End If
'
'        If Stream.State = adStateOpen Then
'            Stream.Close
'        End If
'        If Not Stream Is Nothing Then
'            Set Stream = Nothing
'        End If
'    pPass = True
''---Version 1: Fin


cn.Close

'MsgBox "Adjunto guardado en: " & pPath, vbInformation

lblLoading.Caption = "Archivo Guardado en: " & pPath

Me.MousePointer = vbDefault

If pPass Then
    'Abre el Archivo
    Call Shell("Explorer.exe /e," & pPath, vbNormalFocus)
Else
    MsgBox "No fue posible visualizar el documento! ", vbExclamation
End If


Exit Sub

vError:
  lblLoading.Caption = ""

Me.MousePointer = vbDefault
 
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
 
 On Error Resume Next
 
 'Si no abre el archivo automáticamente, entonces abre el directorio
 Call Shell("Explorer.exe /select," & pPath, vbNormalFocus)
 
 
End Sub

Private Sub sbDocumentos_List()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

On Error GoTo vError


Me.MousePointer = vbHourglass

lsw.ListItems.Clear

Call sbPGX_GA_Conection(cn)

strSQL = "exec spGA_Interface_Documentos_Consulta " & gGA.Empresa & ", '" & glogon.Usuario & "', '" & txtLlave_01.Text _
        & "', '" & txtLlave_02.Text & "', '" & txtLlave_03.Text & "', '" & gGA.Modulo & "'"

rs.CursorLocation = adUseClient

rs.Open strSQL, cn, adOpenStatic, adLockReadOnly
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!FileId)
      itmX.SubItems(1) = rs!Tipo_Desc
      itmX.SubItems(2) = rs!FileName
      itmX.SubItems(3) = rs!FileType
      itmX.SubItems(4) = rs!Vencimiento & ""
      itmX.SubItems(5) = rs!RegistroFecha & ""
      itmX.SubItems(6) = rs!RegistroUsuario & ""
 rs.MoveNext
Loop

rs.Close
cn.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub



Private Sub sbInicializa()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


Call sbPGX_GA_Conection(cn)


rs.CursorLocation = adUseClient

strSQL = "exec spGA_Interface_Modulo_Load " & gGA.Empresa & ", '" & glogon.Usuario & "', '" & gGA.Modulo & "'"
rs.Open strSQL, cn, adOpenStatic, adLockReadOnly
If Not rs.EOF And Not rs.BOF Then
    
    scModulo.Tag = rs!ModuloId
    scModulo.Caption = rs!Descripcion

End If
rs.Close


strSQL = "exec spGA_Interface_Tipos_Documentos " & gGA.Empresa & ", '" & glogon.Usuario & "', '" & gGA.Modulo & "'"

vPaso = True

rs.Open strSQL, cn, adOpenStatic, adLockReadOnly
Do While Not rs.EOF
 cboTipo.AddItem rs!Descripcion & ""
 cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(rs!TypeId)
 rs.MoveNext
Loop

If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboTipo.Text = rs!Descripcion & ""
End If

vPaso = False

rs.Close
cn.Close

txtLlave_01.Text = gGA.Llave_01
txtLlave_02.Text = gGA.Llave_02
txtLlave_03.Text = gGA.Llave_03



dtpEmite.Value = mFecha
dtpVence.Value = DateAdd("m", 120, mFecha)

Call cboTipo_Click
Call sbDocumentos_List

Me.MousePointer = vbDefault

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub
