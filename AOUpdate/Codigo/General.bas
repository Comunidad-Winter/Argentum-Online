Attribute VB_Name = "General"
''
'  This module has the general functions and routines of the Argentum Online Auto Updater
'
' @author Marco Vanotti (marco@vanotti.com.ar)
' @version 0.0.1
' @date 20081005


Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL As Long = 1

Public Caller As String
Public NoExecute As Boolean
Public UPDATES_SITE As String
Public Const UPDATE_URL As String = "http://aotest.argentuuum.com.ar/Aoupdate/"
Public Const UPDATE_URL_MIRROR As String = "http://aoupdate.argentuuum.com.ar/updates/"
Public Const AOUPDATE_FILE As String = "AoUpdate.ini"
Public Const PARAM_UPDATED As String = "/uptodate"


Public Type tAoUpdateFile
    name As String              'File name
    version As Integer          'The version of the file
    MD5 As String * 32          'It's checksum
    Path As String              'Path in the client to the file from App.Path (the server path is the same, changing '\' with '/')
    HasPatches As Boolean       'Weather if patches are available for this file or not (if not the complete file has to be downloaded)
    Comment As String           'Any comments regarding this file.
    Critical As Boolean         'If is a critical file
End Type

Public Type tAoUpdatePatches
    name As String          'its location in the server
    MD5 As String * 32      'It's Checksum
End Type

Public DownloadsPath As String

Public AoUpdatePatches() As tAoUpdatePatches

Public AoUpdateRemote() As tAoUpdateFile
Public DownloadQueue() As Long
Public DownloadQueueIndex As Long
Public PatchQueueIndex As Long
Public ClientParams As String

''
' Loads the AoUpdate Ini File to an struct array
'
' @param file Specifies reference to AoUpdateIniFile
' @return an array of tAoUpdate

Public Function ReadAoUFile(ByVal file As String) As tAoUpdateFile()
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 27/10/2008
'
'*************************************************
    Dim Leer As New clsIniReader
    Dim NumFiles As Integer
    Dim tmpAoUFile() As tAoUpdateFile
    Dim i As Integer
    
On Error GoTo Error
    
    Call Leer.Initialize(file)
    
    NumFiles = Leer.GetValue("INIT", "NumFiles")
    
    ReDim tmpAoUFile(NumFiles - 1) As tAoUpdateFile
    
    For i = 1 To NumFiles
        tmpAoUFile(i - 1).name = Leer.GetValue("File" & i, "Name")
        tmpAoUFile(i - 1).version = CInt(Leer.GetValue("File" & i, "Version"))
        tmpAoUFile(i - 1).MD5 = Leer.GetValue("File" & i, "MD5")
        tmpAoUFile(i - 1).Path = Leer.GetValue("File" & i, "Path")
        tmpAoUFile(i - 1).HasPatches = CBool(Val(Leer.GetValue("File" & i, "HasPatches")))
        tmpAoUFile(i - 1).Comment = Leer.GetValue("File" & i, "Comment")
        tmpAoUFile(i - 1).Critical = CBool(Val(Leer.GetValue("File" & i, "Critical")))
    Next i
    
    ReadAoUFile = tmpAoUFile
    
    Set Leer = Nothing
Exit Function

Error:
    Call MsgBox(Err.Description, vbCritical, Err.Number)
    Set Leer = Nothing
End Function

''
' Compares the local AoUpdate file with the one in the server
'
' @param localUpdateFile Specifies reference to Local Update File
' @param remoteUpdateFile Specifies reference to Remote Update File

Public Sub CompareUpdateFiles(ByRef remoteUpdateFile() As tAoUpdateFile)
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 27/10/2008
'
'*************************************************
    Dim i As Long
    Dim j As Long
    Dim tmpArrIndex As Long
    
' TODO : Check what happens if no files are to be downloaded....
    'ReDim DownloadQueue(0) As Long
    tmpArrIndex = -1
    
    For i = 0 To UBound(remoteUpdateFile)
        If Not FileExist(App.Path & remoteUpdateFile(i).Path & "\" & remoteUpdateFile(i).name, vbNormal) Then
            tmpArrIndex = tmpArrIndex + 1
            ReDim Preserve DownloadQueue(tmpArrIndex) As Long
            DownloadQueue(tmpArrIndex) = i
        ElseIf remoteUpdateFile(i).MD5 <> MD5File(App.Path & remoteUpdateFile(i).Path & "\" & remoteUpdateFile(i).name) Then
            tmpArrIndex = tmpArrIndex + 1
            ReDim Preserve DownloadQueue(tmpArrIndex) As Long
            DownloadQueue(tmpArrIndex) = i
        End If
    Next i
End Sub

''
' Downloads the Updates from the UpdateQueue.
'
' @param DownloadQueue Specifies reference to UpdateQueue
' @param remoteUpdateFile Specifies reference to Remote Update File

Public Sub NextDownload()
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 27/10/2008
'
'*************************************************
On Error GoTo noqueue
    
    If DownloadQueueIndex > UBound(DownloadQueue) Then

On Error GoTo Error
        ClientParams = PARAM_UPDATED & " " & ClientParams
        Call AddtoRichTextBox(frmDownload.rtbDetalle, "Cliente de Argentum Online actualizado correctamente.", 255, 255, 255, True, False, False)
        frmDownload.cmdComenzar.Enabled = True
        
        If Not NoExecute Then
            Call ShellArgentum
        End If
        'End
    Else
        With AoUpdateRemote(DownloadQueue(DownloadQueueIndex))
            If .HasPatches Then
                Dim localVersion As Long
                
                localVersion = -1
                
                If FileExist(App.Path & "\" & .Path & "\" & .name, vbArchive) Then 'Check if local version is too old to be patched.
                    localVersion = GetVersion(App.Path & "\" & .Path & "\" & .name)
                End If
                
                If ReadPatches(DownloadQueue(DownloadQueueIndex) + 1, localVersion, .version, DownloadsPath & AOUPDATE_FILE) Then
                    'Reset index and download patches!
                    PatchQueueIndex = 0
                    Call frmDownload.DownloadPatch(AoUpdatePatches(PatchQueueIndex).name)
                Else
                    'Our version is too old to be patched (it doesn't exist in the server). Overwrite it!
                    .HasPatches = False
                    
                    Call AddtoRichTextBox(frmDownload.rtbDetalle, "Descargando " & .name & " - " & .Comment, 255, 255, 255, True, False, False)
                    
                    Call frmDownload.DownloadFile(Replace(.Path, "\", "/") & "/" & .name)
                End If
            Else
                'Downlaod file. Map local paths to urls.
                
                Call AddtoRichTextBox(frmDownload.rtbDetalle, "Descargando " & .name & " - " & .Comment, 255, 255, 255, True, False, False)
                
                Call frmDownload.DownloadFile(Replace(.Path, "\", "/") & .name)
            End If
        End With
        
        'Move on to the next one
        DownloadQueueIndex = DownloadQueueIndex + 1
    End If
Exit Sub

noqueue: 'If we get here, it means that there isn't any update.
    
    Call AddtoRichTextBox(frmDownload.rtbDetalle, "Descargas finalizadas", 255, 255, 255, True, False, False)
    frmDownload.cmdComenzar.Enabled = True
    
    ClientParams = PARAM_UPDATED & " " & ClientParams
    
    If Not NoExecute Then
        Call ShellArgentum
        End
    End If
Exit Sub

Error:
    Call MsgBox(Err.Description, vbCritical, Err.Number)
End Sub

Public Sub PatchDownloaded()
    Dim localVersion As Long
    
    localVersion = -1

    Call AddtoRichTextBox(frmDownload.rtbDetalle, "Parcheando Archivo de recursos Ao puede demorar unos minutos..", 255, 255, 255, True, False, False)
    With AoUpdateRemote(DownloadQueue(DownloadQueueIndex - 1))
        'Apply downlaoded patch!
            
#If seguridadalkon Then
        If Apply_Patch(App.Path & "\" & .Path & "\", DownloadsPath & "\", UCase(AoUpdatePatches(PatchQueueIndex).MD5), frmDownload.pbDownload) Then
#Else
        If Apply_Patch(App.Path & "\" & .Path & "\", DownloadsPath & "\", frmDownload.pbDownload) Then
#End If
            Call AddtoRichTextBox(frmDownload.rtbDetalle, "Archivo de recursos Ao parcheado correctamente", 255, 255, 255, True, False, False)
        Else
            Call AddtoRichTextBox(frmDownload.rtbDetalle, "No se pudo parchear el archivo de recursos Ao", 255, 255, 255, True, False, False)
        End If
        'Delete patch after patching!
        Kill DownloadsPath & "\" & Right(AoUpdatePatches(PatchQueueIndex).name, Len(AoUpdatePatches(PatchQueueIndex).name) - InStrRev(AoUpdatePatches(PatchQueueIndex).name, "/"))
        
        localVersion = GetVersion(App.Path & "\" & .Path & "\" & .name)
        
        If .version = localVersion Then
            'We finished patching this file, continue!
            Call NextDownload
        Else
            PatchQueueIndex = PatchQueueIndex + 1
            Call frmDownload.DownloadPatch(AoUpdatePatches(PatchQueueIndex).name)
        End If
    End With
End Sub

Private Sub CheckAoUpdateIntegrity()
    Dim nF As Integer
    
    'Look if exists the TEMP folder, if not, create it.
    If Dir$(DownloadsPath, vbDirectory) = vbNullString Then
        Call MkDir(DownloadsPath)
    End If
End Sub

Public Sub ConfgFileDownloaded()
    AoUpdateRemote = ReadAoUFile(DownloadsPath & AOUPDATE_FILE) 'Load the Remote file
    
    Call CompareUpdateFiles(AoUpdateRemote)  'Compare local vs remote.
    
    'Start downloads!
    Call NextDownload
End Sub

Public Sub Main()
    Dim i As Long
    Dim Pos As Byte
    
    DownloadsPath = App.Path & "\TEMP\"
    frmDownload.filePath = DownloadsPath
    
    
    'Nos fijamos si estamos ejecutando la copia del aoupdate o el original, si ejecutamos el original lo copiamos y llamamos al otro con shellexecute
    If UCase(App.EXEName) = "AOUPDATE" And Command = vbNullString Then
        'Nos copiamos..
        On Error GoTo tmpInUse
        
        FileCopy App.Path & "\" & App.EXEName & ".exe", App.Path & "\" & App.EXEName & "tmp" & ".exe"
        Call ShellExecute(0, "OPEN", App.Path & "\" & App.EXEName & "tmp" & ".exe", "NoExecute", App.Path, SW_SHOWNORMAL)    'We open AoUpdateTemp.exe updated
        
        End
    Else
    
        Select Case Command 'Si estamos ejecutando AoUpdateTMP leemos la linea de comandos
            Case vbNullString
                End
            Case "NoExecute"    'El AoUpdateComun nos pasa por parametro que no tenemos que ejecutar automaticamente el Ao al finalizar.
                NoExecute = True
                Caller = ""
            Case "UpDated"
                'Look & kill AoupdateTMP.exe
                On Error GoTo Error
                               
                If FileExist(App.Path & "\" & App.EXEName & "TMP" & ".exe", vbArchive) Then Kill App.Path & "\" & App.EXEName & "TMP" & ".exe"
                
                End
            Case Else
                Pos = InStr(1, Command, " ")
                If Pos Then
                    Caller = Left$(Command, Pos - 1)
                    ClientParams = Right$(Command, Len(Command) - Pos)
                Else
                    Caller = Command
                End If
        End Select
    End If
    
    'Display form
    Call frmDownload.Show
    
    Call CheckAoUpdateIntegrity
    
    'Download the remote AoUpdate.ini to the TEMP folder and let the magic begin
    Call frmDownload.DownloadConfigFile
    
    Exit Sub
tmpInUse:
    MsgBox Err.Description & vbCrLf, vbInformation, "[ " & Err.Number & " ]" & " Error "
    
    Exit Sub
Error:
    If Err.Number = 75 Or Err.Number = 70 Then 'Si el archivo AoUpdateTMP.exe está en uso, entonces esperamos 10 ms y volvemos a intentarlo hasta que nos deje.
        Sleep 10
        Resume
    Else
        MsgBox Err.Description & vbCrLf, vbInformation, "[ " & Err.Number & " ]" & " Error "
        'MsgBox "Error al verificar las actualizaciones, vuelva a intentarlo", vbCritical
        End
    End If
    
End Sub

Public Sub ShellArgentum()
On Error GoTo Error
    
    Call frmDownload.Download.Cancel
    
    If Not FileExist(App.Path & "\" & Caller, vbArchive) Or Caller = "" Then Caller = "Argentum.exe"
    Call ShellExecute(0, "OPEN", App.Path & "\" & Caller, ClientParams, App.Path, SW_SHOWNORMAL)   'We open Argentum.exe updated
    End
    Exit Sub
Error:
    MsgBox "Error al ejecutar el juego", vbCritical
End Sub

Public Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

''
' Loads the patches and their md5. Check if a file isn't too old to be patched
'
' @param NumFile Specifies reference to File in AoUpdateFile file.
' @param begininVersion Specifies reference to LocalVersion
' @param endingVersion Specifies reference to last version of the file
' @param sFile specifies reference to ConfigFile to read data from.
'
' @returns True if the file can be patcheable or false if the file can't be patcheable

Private Function ReadPatches(ByVal numFile As Integer, ByVal beginingVersion As Long, ByVal endingVersion As Long, ByVal sFile As String) As Boolean
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 27/10/2008
'
'*************************************************
    Dim nF As Integer
    Dim i As Long
    Dim Leer As New clsIniReader
    
    nF = FreeFile
    
    Call Leer.Initialize(sFile)
    
    If Not Leer.KeyExists("PATCHES" & numFile & "-" & beginingVersion) Or beginingVersion = -1 Then Exit Function
    ReadPatches = True
    
    ReDim AoUpdatePatches(endingVersion - beginingVersion - 1) As tAoUpdatePatches
    
    For i = beginingVersion To endingVersion - 1
        AoUpdatePatches(i - beginingVersion).name = Leer.GetValue("PATCHES" & numFile & "-" & i, "name")
        AoUpdatePatches(i - beginingVersion).MD5 = Leer.GetValue("PATCHES" & numFile & "-" & i, "md5")
    Next i
End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************
    With RichTextBox
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        RichTextBox.Refresh
    End With
End Sub
