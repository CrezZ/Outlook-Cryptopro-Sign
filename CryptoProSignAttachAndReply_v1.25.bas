Attribute VB_Name = "CryptoProSignAttachAndReply1"
'---------------------------------------------------------------------------------
' The sample scripts are not supported under any Microsoft standard support
' program or service. The sample scripts are provided AS IS without warranty
' of any kind. Microsoft further disclaims all implied warranties including,
' without limitation, any implied warranties of merchantability or of fitness for
' a particular purpose. The entire risk arising out of the use or performance of
' the sample scripts and documentation remains with you. In no event shall
' Microsoft, its authors, or anyone else involved in the creation, production, or
' delivery of the scripts be liable for any damages whatsoever (including,
' without limitation, damages for loss of business profits, business interruption,
' loss of business information, or other pecuniary loss) arising out of the use
' of or inability to use the sample scripts or documentation, even if Microsoft
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------

'''########################
'''English translate
'''########################

'Const ErrorSelectCertificate = "Error select certificate"
'Const ErrorSelectPath = "File 'cryptcp.exe' not found"
'Const ErrorSelectPath2 = "Path to file like 'cryprcp.exe' is bad or file not found. Retry select path ?"
'Const MsgEnterCertificate = "Enter NUMBER of certificate to select:"
'Const MsgEnterManuallyCertificate = "Enter certificate thumbprint (without spaces) like this '1b70c1f8a25453b6ad9dacfd349c05e2988bdbec' :"
'Const MsgChangeCert = "Change certificate?"
'Const MsgRequestSelectCertFromList = "Select from list (if no - enter manually) ?"
'Const MsgSelectPath = "Choose a folder with 'cryptocp.exe'"
'Const MsgRequestChangePath = "Change path to 'cryptcp.exe' (or 'cryptcp.x86.exe' or 'cryptcp.x64.exe') ?"
'Const MsgRequestToSign = "Sign  files in this email?"
'Const MsgRequestDebug = "Enable debug mode?"
'Const ReplyTemplate = "Hello," & vbCrLf & "All signs are ready!"
'Const ErrorFailedToSearchOutlook = "Failed to get the handle of Outlook window!"
'Const MsgRequestAttached = "Use deattached sign?"

'''########################
'''Russian translate
'''########################

Const ErrorSelectCertificate = "Ошибка выбора сертификата"
Const ErrorSelectPath = "Файл 'cryptcp.exe','cryptcp.x86.exe' или 'cryptcp.x64.exe', не найден"
Const ErrorSelectPath2 = "Путь до утилиты 'cryprcp.exe' плохой или файл не найден. Выбрать другой путь ?"
Const MsgEnterCertificate = "Введите ЦИФРУ номера сертификата:"
Const MsgEnterManuallyCertificate = "Введите отпечаток(thumbprint) без пробелов, типа '1b70c1f8a25453b6ad9dacfd349c05e2988bdbec' :"
Const MsgChangeCert = "Поменять сертификат для подписи?"
Const MsgRequestSelectCertFromList = "Выбрать сертификат из списка (Если нет - ввести отпечаток вручную) ?"
Const MsgSelectPath = "Выберите папку с 'cryptocp.exe'"
Const MsgRequestChangePath = "Пoменять путь до 'cryptcp.exe' (или 'cryptcp.x86.exe' или 'cryptcp.x64.exe') ?"
Const MsgRequestToSign = "Подписать файлы в этом письме и подготовить ответное письмо?"
Const MsgRequestDebug = "Нужен ли режим отладки (не закрывать окно КриптоПро)?"
Const MsgRequestAttached = "Создавать открепленную подпись ?"

Const ReplyTemplate = "Здравствуйте," & vbCrLf & " Подписи документов готовы и прикреплены во вложении."
Const ErrorFailedToSearchOutlook = "Не могу найти окно Outlook!"


Option Explicit

' *****************
' For Outlook 2010.
' *****************
#If VBA7 Then
    ' The window handle of Outlook.
    Private lHwnd As LongPtr
    
    ' /* API declarations. */
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
    
' *****************************************
' For the previous version of Outlook 2010.
' *****************************************
#Else
    ' The window handle of Outlook.
    Private lHwnd As Long
    
    ' /* API declarations. */
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
#End If

' The class name of Outlook window.
Private Const olAppCLSN As String = "rctrl_renwnd32"
' Windows desktop - the virtual folder that is the root of the namespace.
Private Const CSIDL_DESKTOP = &H0
' Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed.
Private Const BIF_RETURNONLYFSDIRS = &H1
' Do not include network folders below the domain level in the dialog box's tree view control.
Private Const BIF_DONTGOBELOWDOMAIN = &H2
' The maximum length for a path is 260 characters.
Private Const MAX_PATH = 260

' ######################################################
'  Returns the number of attachements in the selection.
' ######################################################
Public Function SaveAttachmentsFromSelection(folder As String) As Long
    Dim objFSO              As Object       ' Computer's file system object.
    Dim objShell            As Object       ' Windows Shell application object.
    Dim objFolder           As Object       ' The selected folder object from Browse for Folder dialog box.
    Dim objItem             As Object       ' A specific member of a Collection object either by position or by key.
    Dim selItems            As Selection    ' A collection of Outlook item objects in a folder.
    Dim atmt                As Attachment   ' A document or link to a document contained in an Outlook item.
    Dim strAtmtPath         As String       ' The full saving path of the attachment.
    Dim strAtmtFullName     As String       ' The full name of an attachment.
    Dim strAtmtName(1)      As String       ' strAtmtName(0): to save the name; strAtmtName(1): to save the file extension. They are separated by dot of an attachment file name.
    Dim strAtmtNameTemp     As String       ' To save a temporary attachment file name.
    Dim intDotPosition      As Integer      ' The dot position in an attachment name.
    Dim atmts               As Attachments  ' A set of Attachment objects that represent the attachments in an Outlook item.
    Dim lCountEachItem      As Long         ' The number of attachments in each Outlook item.
    Dim lCountAllItems      As Long         ' The number of attachments in all Outlook items.
    Dim strFolderPath       As String       ' The selected folder path.
    Dim blnIsEnd            As Boolean      ' End all code execution.
    Dim blnIsSave           As Boolean      ' Consider if it is need to save.
    
    blnIsEnd = False
    blnIsSave = False
    lCountAllItems = 0
    
    On Error Resume Next
    
    Set selItems = ActiveExplorer.Selection
    
    If Err.Number = 0 Then
        
        ' Get the handle of Outlook window.
        lHwnd = FindWindow(olAppCLSN, vbNullString)
        
        If lHwnd <> 0 Then
            
            ' /* Create a Shell application object to pop-up BrowseForFolder dialog box. */
            Set objShell = CreateObject("Shell.Application")
            Set objFSO = CreateObject("Scripting.FileSystemObject")
          '  Set objFolder = objShell.BrowseForFolder(lHwnd, "Select folder to save attachments:", _
           '                                          BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN, CSIDL_DESKTOP)
            objFolder = folder
            
            ' /* Failed to create the Shell application. */
          '  If Err.Number <> 0 Then
          '      MsgBox "Run-time error '" & CStr(Err.Number) & " (0x" & CStr(Hex(Err.Number)) & ")':" & vbNewLine & _
          '             Err.Description & ".", vbCritical, "Error from Attachment Saver"
          '      blnIsEnd = True
          '      GoTo PROC_EXIT
          '  End If
            
         '   If objFolder Is Nothing Then
         '       strFolderPath = ""
         '       blnIsEnd = True
         '       GoTo PROC_EXIT
          '  Else
                'strFolderPath = CGPath(objFolder.Self.Path)
                strFolderPath = CGPath(folder)
                
                ' /* Go through each item in the selection. */
                For Each objItem In selItems
                    lCountEachItem = objItem.Attachments.Count
                    
                    ' /* If the current item contains attachments. */
                    If lCountEachItem > 0 Then
                        Set atmts = objItem.Attachments
                        
                        ' /* Go through each attachment in the current item. */
                        For Each atmt In atmts
                            
                            
                            ' Get the full name of the current attachment.
                            strAtmtFullName = atmt.fileName
                            
                            ' Find the dot postion in atmtFullName.
                            intDotPosition = InStrRev(strAtmtFullName, ".")
                            
                            ' Get the name.
                            strAtmtName(0) = Left$(strAtmtFullName, intDotPosition - 1)
                            ' Get the file extension.
                            strAtmtName(1) = Right$(strAtmtFullName, Len(strAtmtFullName) - intDotPosition)
                            ' Get the full saving path of the current attachment.
                            strAtmtPath = strFolderPath & atmt.fileName
                            
                            ' /* If the length of the saving path is not larger than 260 characters.*/
                            If Len(strAtmtPath) <= MAX_PATH Then
                                ' True: This attachment can be saved.
                                blnIsSave = True
                                
                                ' /* Loop until getting the file name which does not exist in the folder. */
                                Do While objFSO.FileExists(strAtmtPath)
                                    strAtmtNameTemp = strAtmtName(0) & _
                                                      Format(Now, "_mmddhhmmss") & _
                                                      Format(Timer * 1000 Mod 1000, "000")
                                    strAtmtPath = strFolderPath & strAtmtNameTemp & "." & strAtmtName(1)
                                        
                                    ' /* If the length of the saving path is over 260 characters.*/
                                    If Len(strAtmtPath) > MAX_PATH Then
                                        lCountEachItem = lCountEachItem - 1
                                        ' False: This attachment cannot be saved.
                                        blnIsSave = False
                                        Exit Do
                                    End If
                                Loop
                                
                                ' /* Save the current attachment if it is a valid file name. */
                                If blnIsSave Then atmt.SaveAsFile strAtmtPath
                            Else
                                lCountEachItem = lCountEachItem - 1
                            End If
                        Next
                    End If
                    
                    ' Count the number of attachments in all Outlook items.
                    lCountAllItems = lCountAllItems + lCountEachItem
                Next
            End If
        Else
            MsgBox "Failed to get the handle of Outlook window!", vbCritical, "Error from Attachment Saver"
            blnIsEnd = True
            GoTo PROC_EXIT
        End If
        
    ' /* For run-time error:
    '    The Explorer has been closed and cannot be used for further operations.
    '    Review your code and restart Outlook. */
'    Else
'        MsgBox "Please select an Outlook item at least.", vbExclamation, "Message from Attachment Saver"
'        blnIsEnd = True
  '  End If
    
PROC_EXIT:
    SaveAttachmentsFromSelection = lCountAllItems
    
    ' /* Release memory. */
    If Not (objFSO Is Nothing) Then Set objFSO = Nothing
    If Not (objItem Is Nothing) Then Set objItem = Nothing
    If Not (selItems Is Nothing) Then Set selItems = Nothing
    If Not (atmt Is Nothing) Then Set atmt = Nothing
    If Not (atmts Is Nothing) Then Set atmts = Nothing
    
    ' /* End all code execution if the value of blnIsEnd is True. */
    If blnIsEnd Then End
End Function

' #####################
' Convert general path.
' #####################
Public Function CGPath(ByVal path As String) As String
    If Right(path, 1) <> "\" Then path = path & "\"
    CGPath = path
End Function

' Create New Mail And Attach files

Sub MakeAllWork(path As String)
On Error Resume Next

Dim objFS As New FileSystemObject

'objFS = CreateObject("Scripting.FileSystemObject")
Dim objFolder As Scripting.folder
Dim objFile As File
Dim objMessage As Object, intX As Integer
Dim strFolderPath As String
Dim olItem As Outlook.MailItem
    Dim olReply As MailItem ' Reply
    Dim olRecip As Recipient ' Add Recipient
    
'If ActiveInspector Is Nothing Then Exit Sub

'Set objMessage = ActiveInspector.CurrentItem
strFolderPath = path

Set objFolder = objFS.GetFolder(strFolderPath)
 
 For Each olItem In Application.ActiveExplorer.Selection 'begin Reply
    Set olReply = olItem.ReplyAll
    'Set olRecip = olReply.Recipients.Add("Email Address Here") ' Recipient Address
        olRecip.Type = olCC
            olReply.HTMLBody = ReplyTemplate & vbCrLf & olReply.HTMLBody
        
'attach files
For Each objFile In objFolder.Files
'If Right(objFile.Name, 3) = "pdf" Then
olReply.Attachments.Add (objFile.path)
'objFile.Delete (True)
'End If
Next

olReply.Display
        'olReply.Send
    Next olItem

''MsgBox "Your PDF files have been attached, and the sources deleted.", vbOKOnly, vbInformation, "Operation Complete"

Set objFS = Nothing
Set objFolder = Nothing
Set objFile = Nothing
Set objMessage = Nothing
End Sub



Private Function FileExists(FilePath As String) As Boolean
     FileExists = (Dir(FilePath) > "")
End Function

' ######################################
' Run this macro for change  settings on path to cryptocp.exe
' ######################################
Private Function ChangePath() As String

 
    Dim path As String
    path = ""
    Dim addpath As String
    
    Dim objFSO              As Object       ' Computer's file system object.
    Dim objShell            As Object       ' Windows Shell application object.
    Dim objFolder           As Object       ' The selected folder object from Browse for Folder dialog box.
    
'''CHANGE PATH TO cryptocp.exe
If MsgBox(MsgRequestChangePath, vbYesNo) = vbYes Then
        
         lHwnd = FindWindow(olAppCLSN, vbNullString) ''search outlook
        
        If lHwnd <> 0 Then
                        Set objShell = CreateObject("Shell.Application")
                        
                        
RETRY_SELECT:
                ''openfile dialog
              Set objFolder = objShell.BrowseForFolder(lHwnd, MsgSelectPath, 0, 17)
           
            
          
            If objFolder Is Nothing Then
                path = ""
                
           Else
        path = objFolder.Self.path
        addpath = ""
        If FileExists(path + "\cryptcp.exe") Then
        path = path + "\cryptcp.exe"
        Else
          If FileExists(path + "\cryptcp.x86.exe") Then
              path = path + "\cryptcp.x86.exe"
            Else
              If FileExists(path + "\cryptcp.x64.exe") Then path = path + "\cryptcp.x64.exe"
          End If
        End If
          If path = "" Then
               If MsgBox(ErrorSelectPath2, vbYesNo) = vbYes Then GoTo RETRY_SELECT
          Else
            SaveSetting "CryptcpOutlook", "Startup", "Path", path
            ChangePath = path
            
          End If
        End If
    End If
End If
If path = "" Then ChangePath = ""
End Function

' ######################################
' Run this macro for change  settings on certificate
' ######################################
Private Sub ChangeCer()
    Dim email As String

'''CHANGE CERTIFICATE

        If MsgBox(MsgChangeCert, vbYesNo) = vbYes Then
          If MsgBox(MsgRequestSelectCertFromList, vbYesNo) = vbYes Then
            email = ListCert
          Else
            email = InputBox(MsgEnterManuallyCertificate)
            email = Replace(email, " ", "") ''remove spaces
          End If
        SaveSetting "CryptcpOutlook", "Startup", "Cert", email
        
        End If

End Sub

' ######################################
' Run this macro for change  settings on debug
' ######################################
Private Sub ChangeDebug()
    Dim email As String

'''CHANGE CERTIFICATE
        If MsgBox(MsgRequestDebug, vbYesNo) = vbYes Then
        
          SaveSetting "CryptcpOutlook", "Startup", "Debug", 1  '' enable debug mode
        Else
          SaveSetting "CryptcpOutlook", "Startup", "Debug", 0
        End If

End Sub

Private Sub ChangeAttachedSign()
    Dim email As String
  Dim result As VbMsgBoxResult
  
'''CHANGE CERTIFICATE type
        result = MsgBox(MsgRequestAttached, vbYesNoCancel)
        If result = vbYes Then
        
          SaveSetting "CryptcpOutlook", "Startup", "AttachedSign", "-signf" ''deattached sign
        Else
        If result = vbNo Then
            SaveSetting "CryptcpOutlook", "Startup", "AttachedSign", "-sign"
        End If
        End If

End Sub


Public Sub ChangeCryptoSettings()

ChangeDebug
ChangeCer
ChangePath
ChangeAttachedSign
End Sub
' ######################################
' Run this macro for Sign all attachments in selected email
' ######################################


Public Sub ExecuteSign()
    Dim lNum As Long
    Dim temp As String  ''path to temo folder
    Dim temp2 As String  '' path to temo sign folder
    Dim result As Double  '' result in hex
    Dim action As VbMsgBoxResult '' for dialogs
    Dim objShell            As Object  '' for file checks
    Dim signf As String  '' attacher or deattached sign
    Dim email As String '' thumbprint of certificate
    Dim path As String  '' path to cryptocp utility
     
     
     
    On Error Resume Next
   ' Dim x509Store As New X509Store("My")
   
    action = MsgBox(MsgRequestToSign, vbYesNo, "Singning")
     If action = vbNo Then Exit Sub
     
     
    ''Temp directories
    
    temp = Environ("Temp") + "\cryptocp"
    temp2 = Environ("Temp") + "\cryptocp\signs"
    
    MkDir (temp)
    MkDir (temp2)
    
    
    lNum = SaveAttachmentsFromSelection(temp)
    
     email = GetSetting("CryptcpOutlook", "Startup", "Cert")
    If email = "" Then
        ChangeCer
    End If
    
     path = GetSetting("CryptcpOutlook", "Startup", "Path")
    If path = "" Or FileExists(path) = False Then
        MsgBox (ErrorSelectPath)
        path = ChangePath
        If path = "" Then Exit Sub
    End If
    
    Dim cmd As String * 1000
    
    signf = GetSetting("CryptcpOutlook", "Startup", "AttachedSign")
    If signf = "" Then
        signf = "-signf"
    End If
    
    cmd = path + " " + signf + "  -cert -thumbprint """ + email + """ -dir " + temp2 + " " + temp + "\*.* "
    'cmd = "c:\cryptocp.cmd c:\ -signf -thumbprint """ + email + """ -dir " + temp2 + " " + temp + "\*.* "
    'result = Shell(cmd, vbNormalFocus)
    
    
    Set objShell = CreateObject("WScript.shell")
    ''if debug is set
    If GetSetting("CryptcpOutlook", "Startup", "Debug") = 1 Then
        objShell.Run "%comspec% /k " + cmd, 1, True
    Else
        objShell.Run cmd, 1, True
    End If
    
'        If result <> 0 Then
'        MsgBox ("Error Signing!")
'        Exit Sub
 '   Kill (temp2 + "\*.*")
 '   Kill (temp + "\*.*")
 '       End If
        
    MakeAllWork (temp2)
    Kill (temp2 + "\*.*")
    Kill (temp + "\*.*")
  '  RmDir (temp2)
  '  RmDir (temp)
'    If lNum > 0 Then
'        MsgBox CStr(lNum) & " attachment(s) was successfully worked", vbInformation, "Message from Attachment Saver"
'    Else
'        MsgBox "No attachment(s) in the selected Outlook items.", vbInformation, "Message from Attachment Saver"
'    End If
End Sub

Private Function GetCert(thumbprint As String) As Variant
''Get certificate object fot thumbprint
Dim root As Variant   ' The FPCLib.FPC root object
    Dim result() As Variant
     Dim i As Integer
     Dim build
     Set root = CreateObject("CAdESCOM.Store")
     build = root.Open(2, "MY")
    ' Declare the other objects needed.
    'Dim server        ' An FPCServer object
    Dim certificates  ' An FPCCertificates collection
    Dim certificate   ' An FPCCertificate object

    Set certificates = root.certificates
    ' Get references to the server object
    ' and the applicable certificates collections.
    'Display some properties of each certificate.
    If certificates.Count = 0 Then
        ''ListCertificates = Nothing
        GetCert = Nothing
        root = Nothing
        certificates = Nothing
        Exit Function
    End If
    
    '' Find certificate
    For Each certificate In certificates
     If certificate.thumbprint = thumbprint Then
      ''return cert
     GetCert = certificate
        root = Nothing
        certificates = Nothing
     Exit Function
     End If
     
    Next
    GetCert = Nothing
End Function

Public Sub MakeSign2(doc As Variant, thumbprint As String)
    Dim certificate
    '' Assign cert
    certificate = GetCert(thumbprint)
    
    Dim oSigner As Variant
    Dim oSignedData  As Variant
    Dim sSignedData
    
    Set oSigner = CreateObject("CAdESCOM.CPSigner")
    Set oSignedData = CreateObject("CAdESCOM.CadesSignedData")
    Set oSigner.certificate = certificate
    oSignedData.Content = doc  'what to sign
    
    sSignedData = oSignedData.Sign(oSigner, False)

End Sub

Function ListCert() As String
'' return tumbprint of selected certificate
Dim certs() As Variant
Dim c As String
Dim text As String
Dim i As Integer

certs = ListCertificates("CLIENT")
i = 1
text = MsgEnterCertificate & vbCrLf
For i = 1 To UBound(certs) - 1
c = certs(i, 1)

text = text & CStr(i) & ". " & c
Next

c = InputBox(text, "Select certificate")

If Not IsNumeric(c) Then
MsgBox ("Error, invalid cert number")
ListCert = ""
Else
  ''save new cert
  i = CInt(c)
 If UBound(certs) > i And i > 0 Then
  'SaveSetting "CryptcpOutlook", "Startup", "Cert", certs(i, 2)
   ListCert = certs(i, 2)
   Else
   MsgBox ("Error, invalid cert number")
   ListCert = ""
  End If
End If

End Function

Public Function ListCertificates(storeType) As Variant
''' Return 2 dim array:
''' Result(i)(1) - cert description
''' Result(i)(2) - cert tumbprint

    ' Create the root object.
    Dim root As Variant   ' The FPCLib.FPC root object
    Dim result() As Variant
     Dim i As Integer
     Dim build
   ' Set root = CreateObject("FPC.Root")
    'Set root = New FPCLib.FPC
    'Set root = CreateObject("CADESCOM.Store")
    Set root = CreateObject("CAdESCOM.Store")
     build = root.Open(2, "MY")
   

    ' Declare the other objects needed.
    Dim server        ' An FPCServer object
    Dim certificates  ' An FPCCertificates collection
    Dim certificate   ' An FPCCertificate object

    Set certificates = root.certificates
    ' Get references to the server object
    ' and the applicable certificates collections.
    'Display some properties of each certificate.
    If certificates.Count = 0 Then
        ListCertificates = Nothing
        Exit Function
    End If
    i = 1
    ReDim result(1 To certificates.Count, 1 To 2)
    For Each certificate In certificates
        Dim text As String
        text = ""
        
        'text = text + "Issued to: " & certificate.IssuerName & vbCrLf
        text = text + " Person: " & Left(certificate.SubjectName, 55) & vbCrLf
        text = text + "    Valid from: " & certificate.ValidFromDate & vbCrLf
        text = text + "    Valid to: " & certificate.ValidToDate & vbCrLf
        
        result(i, 1) = text
        result(i, 2) = certificate.thumbprint
        
     '   If storeType = "ALLSSL" Then
     '       Select Case certificate.CertificateStore
     '           Case fpcLocalMachinePersonalStore
     '               WScript.Echo "Store: Personal for local computer"
     '           Case fpcCurrentUserPersonalStore
     '               WScript.Echo "Store: Personal for current user"
     '           Case fpcFirewallServicePersonalStore
     '               WScript.Echo "Store: Personal for Firewall service"
     '       End Select
     '   End If
      i = i + 1
    Next
    ListCertificates = result
    
End Function




Public Sub InstallButton()
   

End Sub


 

