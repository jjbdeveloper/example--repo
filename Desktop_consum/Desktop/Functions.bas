Attribute VB_Name = "Functions21"

'Libreria paracontrol apagado PC
Private Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags&, ByVal dwReserved&)
'Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

'Declaracion dll para ejecutar aplicaciones mdeiante API's windows
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    


    
Public Function AbrirLibro(PATH_ORIGEN As String) As Boolean
    Dim TestWorkbook As Workbook
    Set TestWorkbook = Nothing
    On Error Resume Next
    Set TestWorkbook = Workbooks(PATH_ORIGEN)
    On Error GoTo 0
    If TestWorkbook Is Nothing Then
       'ruta = "C:\"
       Workbooks.Open FileName:=PATH_ORIGEN
        AbrirLibro = True
       'Workbooks.Open FileName:=ruta & "\hola.xlsx"
    Else
       MsgBox "El archivo ya estaba abierto o no se pudo abrir"
    End If
End Function

Public Function CerrarYGuardar() As Boolean
   ActiveWorkbook.Close savechanges:=False
    CerrarYGuardar = False
End Function

Public Function Convert_to_pdf(desti As String) As Boolean
Range("A1:I41").Select
Selection.ExportAsFixedFormat Type:=xlTypePDF, FileName:=PATH_DESTI & "\hola.pdf", _
Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
Convert_to_pdf = True
End Function




'Aquesta funcio sols es per imprimir fitxers txt i falta provar-la

'Public Function Imprimir_Fitxer(Path As String)
' Dim Free_File As Integer
'    Dim Datos As String
'    Dim pos As Integer
'    Dim L As String
'    Dim Palabra As String
'
'    ' número de archivo libre
'    Free_File = FreeFile
'
'    ' abre el archivo para leerlo
'    Open Path For Input As Free_File
'
'    ' Almacena los datos del archivo en la variable
'    Datos = Input(LOF(Free_File), Free_File)
'    ' cierra el archivo
'    Close Free_File
'
'
'    Do While Len(Datos) > 0
'
'        pos = InStr(Datos, vbCrLf)
'        If pos = 0 Then
'            L = Datos
'            Datos = ""
'        Else
'            ' linea
'            L = Left$(Datos, pos - 1)
'
'            Datos = Mid$(Datos, pos + 2)
'        End If
'
'    ' palabras
'    Do While Len(L) > 0
'        ' posición para extraer la palabra
'        pos = InStr(L, " ")
'        If pos = 0 Then
'            Palabra = L
'            L = ""
'        Else
'            Palabra = Left$(L, pos)
'            L = Mid$(L, pos + 1)
'        End If
'
'    ' verifica que no se pase del ancho de la hoja
'    If (Printer.CurrentX + Printer.TextWidth(Palabra)) <= Printer.ScaleWidth Then
'        ' imprime la palabra
'        Printer.Print Palabra;
'    ' si no imprime en la siguiente linea
'    Else
'        Printer.Print
'        ' verifica que no se pase del alto de la hoja
'        If (Printer.CurrentY + Printer.Font.Size) > Printer.ScaleHeight Then
'            ' nueva hoja
'            Printer.NewPage
'        End If
'        ' imprime la palabra
'        Printer.Print Palabra;
'    End If
'    Loop
'    Printer.Print
'    Loop
'
'    ' Fin. Manda a imprimir
'    Printer.EndDoc
'
'End Function
'
'



Public Function Send_Mail() As Boolean
Dim oMsg, oConf, Flds
Set oMsg = CreateObject("CDO.Message")
Set oConf = CreateObject("CDO.Configuration")
oConf.Load -1
Set Flds = oConf.Fields
With Flds
.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Compte_Correu
.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Contraseña
.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Direcció_Server
.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = Port
.Update
End With

With oMsg
Set .Configuration = oConf
.From = Compte_Correu
.To = Direcció_Denviament
.Subject = Assumpte
.TextBody = Cos_del_missatge
.AddAttachment (PATH_DESTI)


' Path_Adjunto = "c:\hola.pdf"
' If Path_Adjunto <> vbNullString Then
'        Obj_Email.AddAttachment (Path_Adjunto)
'    End If
.Send
End With
Send_Mail = True
End Function



Public Function Prepara_Mail() As Boolean

'exemple v1 = Range("I2").Value

Cos_del_missatge = Range("R45").Value
Nom_del_Fitxer = ActiveWorkbook.ActiveSheet.Range("L44").Value
Direcció_Denviament = ActiveWorkbook.ActiveSheet.Range("O47").Value
Assumpte = ActiveWorkbook.ActiveSheet.Range("O48").Value

If Not Cos_del_missatge = Empty And Nom_del_Fitxer = Empty And Direcció_Denviament = Empty And Assumpte = Empty Then
Prepara_Mail = True
Else
Prepara_Mail = False
MsgBox ("Error Estem enviant algun paràmetre sense càrregar")
End If
End Function

Public Function ApagaPC()
Dim i As Integer
    i = ExitWindowsEx(Apagar, 0&)   'Apaga el Sistema
    Shell "shutdown -s -t 0" ' 'Apaga el equipo en Win XP
End Function





