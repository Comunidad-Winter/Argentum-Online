Attribute VB_Name = "wskapiAO"
Option Explicit

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long


Public Const GWL_WNDPROC = (-4)


Public OldWProc As Long
Public ActualWProc As Long
Public hWndMsg As Long
'===========================================
Public SockListen As Long


Public Sub IniciaWsApi()
#If UsarAPI Then

Call LogApiSock("IniciaWsApi")

hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", 0, 0, 0, 0, 0, 0, 0, App.hInstance, ByVal 0&)
OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

Dim Desc As String
Call StartWinsock(Desc)

#End If
End Sub

Public Sub LimpiaWsApi()
#If UsarAPI Then

Call LogApiSock("LimpiaWsApi")

If WSAStartedUp Then
    Call EndWinsock
End If

'If OldWProc <> 0 Then
'    SetWindowLong hWnd, GWL_WNDPROC, OldWProc
'    OldWProc = 0
'End If

#End If
End Sub

Public Function BuscaSlotSock(S As Long) As Long
#If UsarAPI Then

Dim i As Long

For i = 1 To MaxUsers
    If UserList(i).ConnID = S Then
        BuscaSlotSock = i
        Exit Function
    End If
Next i

BuscaSlotSock = -1

#End If
End Function


Public Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#If UsarAPI Then

Dim Ret As Long
Dim Tmp As String

WndProc = 0

Select Case msg
Case 1025
    Dim S As Long, E As Long
    Dim N As Integer
    S = wParam
    E = WSAGetSelectEvent(lParam)
    Debug.Print "Msg: " & msg & " W: " & wParam & " L: " & lParam
    'Call LogApiSock("Msg: " & Msg & " W: " & wParam & " L: " & lParam)
    
    Select Case E
    Case FD_ACCEPT
        Call EventoSockAccept(S)
    Case FD_READ
        
        N = BuscaSlotSock(S)
        If N < 0 Then
            Call apiclosesocket(S)
            Exit Function
        End If
        
        '4k de buffer
        Tmp = Space(4096)
        
        Ret = recv(S, Tmp, Len(Tmp), 0)
        If Ret < 0 Then
            Debug.Print "Error en Recv"
            Call LogApiSock("Error en Recv:N=" & N & ":S=" & S)
            Call CloseSocket(N)
            Exit Function
        End If
        
        Tmp = Left(Tmp, Ret)
        
        'Call LogApiSock("WndProc:FD_READ:N=" & N & ":TMP=" & Tmp)
        
        Call EventoSockRead(N, Tmp)
    Case FD_CLOSE
        N = BuscaSlotSock(S)
        
        Call LogApiSock("WndProc:FD_CLOSE:N=" & N & ":Err=" & WSAGetAsyncError(lParam))
        
        If N < 0 Then
            Call apiclosesocket(S)
        Else
            Call EventoSockClose(N)
        End If
    End Select
Case Else
    WndProc = CallWindowProc(OldWProc, hWnd, msg, wParam, lParam)
End Select

#End If
End Function

Public Sub WsApiEnviar(Slot As Integer, str As String)
#If UsarAPI Then
Dim Ret As String

If UserList(Slot).ConnID > -1 Then
    Ret = send(ByVal UserList(Slot).ConnID, ByVal str, ByVal Len(str), ByVal 0)
    If Ret < 0 Then
        'Debug.Print "Error en Send"
        'LogApiSock ("Error en Send, slot: " & Slot)
        'Call CloseSocket(Slot)
        Exit Sub
    End If
End If
#End If

End Sub



Public Sub LogApiSock(str As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\wsapi.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

Exit Sub

errhandler:

End Sub
