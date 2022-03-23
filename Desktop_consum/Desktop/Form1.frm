VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   10725
   ClientLeft      =   4980
   ClientTop       =   3615
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   10725
   ScaleWidth      =   9540
   Begin VB.Frame Frame3 
      Caption         =   "AUTOMATIC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9495
      Begin VB.CommandButton cmdproceso 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   2640
         Picture         =   "Form1.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "CONTROL DE PROGRAMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2415
      Left            =   0
      TabIndex        =   4
      Top             =   8280
      Width           =   9495
      Begin VB.CommandButton Command1 
         Caption         =   "Apagar el PC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   600
         Picture         =   "Form1.frx":5E40
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   3135
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Sortir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   5040
         Picture         =   "Form1.frx":6F72
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "MANUAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   9495
      Begin VB.CommandButton cmdConverttopdf 
         Caption         =   "1er Convertir a PDF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   240
         Picture         =   "Form1.frx":845D
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   3735
      End
      Begin VB.CommandButton cmdSendMail 
         Caption         =   "3er Enviar Correu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   2520
         Picture         =   "Form1.frx":1E42F
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2760
         Width           =   3615
      End
      Begin VB.CommandButton cmdImprimirPdf 
         Caption         =   "2on Imprimir Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   4320
         Picture         =   "Form1.frx":1FBD6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Libreria paracontrol apagado PC
Private Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags&, ByVal dwReserved&)
'Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

'Declaracion dll para ejecutar aplicaciones mdeiante API's windows
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Declare Function FormatMessage Lib "kernel32" _
Alias "FormatMessageA" _
(ByVal dwFlags As Long, _
lpSource As Any, _
ByVal dwMessageId As Long, _
ByVal dwLanguageId As Long, _
ByVal lpBuffer As String, _
ByVal nSize As Long, _
Arguments As Long) As Long






'CONSTANTS DEL PROGRAMA:
Const PATH_ORIGEN = "c:\factura.xlsm"
Const PATH_DESTI = "c:"

    
'constants per la funcio print file
    
Const SW_HIDE = 0&
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
' función que imprime un documento de cualquier aplicación
Public Function PrintFile(FileName As String) As Variant
 
Dim RetVal As Long
Dim sError As String
Dim LenMsg As Long

' se manda imprimir el documento
RetVal = ShellExecute(0&, "print", FileName, 0&, vbNullString, SW_HIDE)

' si se ha producido algún error
If RetVal < 33 Then
sError = Space(1024)
' obtenemos el mensaje de error que manda el sistema
LenMsg = FormatMessage( _
FORMAT_MESSAGE_FROM_SYSTEM, _
ByVal 0&, _
RetVal, _
0&, _
sError, _
Len(sError), _
0&)
' devolvemos el mensaje de error
PrintFile = Left(sError, LenMsg - 1)
Else
' la función tuvo éxito
PrintFile = True
End If

End Function

Private Sub cmdImprimirPdf_Click()
Dim Resultado4 As Variant

' enviamos el fichero a la impresora
Resultado4 = PrintFile(PATH_DESTI)
If Resultado4 = True Then
MsgBox "El documento se envió a la impresora"
Else
MsgBox Resultado4
End If
End Sub

Private Sub cmdproceso_Click()
cmdConverttopdf_Click
Sleep 500
cmdImprimirPdf_Click
Sleep 500
cmdSendMail_Click
Sleep 500
End Sub

Private Sub cmdSalir_Click()
Unload Form1
End Sub

Private Sub cmdSendMail_Click()
Resultado5 = Functions.Prepara_Mail
If Resultado5 Then
Resultado6 = Functions.Send_Mail
End If

If Resultado And Resultado2 And Resultado3 And Resultado4 And Resultado5 And Resultado6 Then
MsgBox ("Enhorabona campió!! Acabes de fer tot el procediment correctament, ara sols et queda comprovar que s'ha enviat el mail correctament i fer el seguiment del ingrés de diners.")
End If

End Sub

Private Sub cmdConverttopdf_Click() 'Esta comentario es un ejemplo

'Fem la crida a la funció obrir el fitxer Excel, Reportem funció utilitzada

Resultado = AbrirLibro(PATH_ORIGEN)

'Si s'ha obert be fitxer després cridem funció convertir

If Resultado Then
   Resultado2 = Convert_to_pdf(PATH_DESTI)
End If

'Cridem la funció per guardar i tancar fitxer
If Resultado And Resultado2 Then
Resultado3 = CerrarYGuardar
End If
End Sub


Private Sub Command1_Click()
Functions.ApagaPC
End Sub


