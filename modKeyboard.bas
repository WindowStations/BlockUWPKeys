Attribute VB_Name = "modKeyboard"
Private Const WH_KEYBOARD_LL As Long = 13
Private Const HC_GETNEXT As Long = 1
Private Const HC_ACTION As Long = 0
Private Type KBDLLHOOKSTRUCT
   vkCode As Long
   scanCode As Long
   Flags As Long
   time As Long
   dwExtraInfo As Long
End Type
Private Declare Function apiSetWindowsKeyHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function apiUnhookWindowsHookEx Lib "user32" Alias "UnhookWindowsHookEx" (ByVal hHook As Long) As Long
Private Declare Function apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As KBDLLHOOKSTRUCT, ByVal pSource As Long, ByVal cb As Long) As Long
Private Declare Function apiCallNextKeyHookEx Lib "user32" Alias "CallNextHookEx" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private hKey As Long
Private hStruct As KBDLLHOOKSTRUCT
Public Function HookKeyboard() As Long
    On Error Resume Next
   If hKey <> 0 Then
      If apiUnhookWindowsHookEx(hKey) <> 0 Then hKey = 0
   End If
   hKey = apiSetWindowsKeyHookEx(WH_KEYBOARD_LL, AddressOf Callback, App.hInstance, 0)
   HookKeyboard = hKey
End Function
Public Function UnhookKeyboard() As Long
   On Error Resume Next
   If hKey = 0 Then Exit Function
   If apiUnhookWindowsHookEx(hKey) <> 0 Then hKey = 0
   UnhookKeyboard = hKey
End Function
Private Function Callback(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   On Error GoTo callnxt
   If Code <> HC_ACTION Then
      Callback = apiCallNextKeyHookEx(0, Code, wParam, lParam)
      Exit Function
   End If
   apiCopyMemory hStruct, lParam, Len(hStruct)
   On Error Resume Next
   If hStruct.vkCode > 194 Then
      If hStruct.vkCode < 219 Then 'if xinput to uwp apps
         Callback = HC_GETNEXT 'block input from being dispatched to the target UWP window
         Exit Function
      End If
   End If
callnxt:
   Callback = apiCallNextKeyHookEx(hKey, Code, wParam, lParam)
End Function

