Attribute VB_Name = "Module1"
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - August 2002
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au

'Linking Richtextbox scrolling example
'This could be tweaked for use on many controls
'other than RichTextBox's

Option Explicit
'Used to locate scroll position
Public Type POINTL
    x As Long
    y As Long
End Type
'used to set scroll position
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_USER = &H400
Private Const EM_GETSCROLLPOS = (WM_USER + 221)
Private Const EM_SETSCROLLPOS = (WM_USER + 222)
'Hooking API to recieve messages from the RichTextBox's when they scroll
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private glPrevWndProc1 As Long
Private glPrevWndProc2 As Long
Public Function fSubClass(hwnd1 As Long, hwnd2 As Long) As Long
    'Hook the RichTextBox's
    glPrevWndProc1 = SetWindowLong(hwnd1, GWL_WNDPROC, AddressOf pMyWindowProc)
    glPrevWndProc2 = SetWindowLong(hwnd2, GWL_WNDPROC, AddressOf pMyWindowProc)
End Function
Public Sub pUnSubClass(hwnd1 As Long, hwnd2 As Long)
    'Unhook the RichTextBox's
    Call SetWindowLong(hwnd1, GWL_WNDPROC, glPrevWndProc1)
    Call SetWindowLong(hwnd2, GWL_WNDPROC, glPrevWndProc2)
End Sub
Private Function pMyWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim P As POINTL
    Select Case hw
        Case Form1.RichTextBox1.hwnd 'RichTextBox that is being scrolled
            Select Case uMsg
                Case 533, 8465 'scroll messages
                    SendMessage hw, EM_GETSCROLLPOS, 0, P 'new scroll position
                    SendMessage Form1.RichTextBox2.hwnd, EM_SETSCROLLPOS, 0, P 'make the other RichTextBox match
            End Select
            pMyWindowProc = CallWindowProc(glPrevWndProc1, hw, uMsg, wParam, lParam)
        Case Form1.RichTextBox2.hwnd 'RichTextBox that is being scrolled
            Select Case uMsg
                Case 533, 8465 'scroll messages
                    SendMessage hw, EM_GETSCROLLPOS, 0, P 'new scroll position
                    SendMessage Form1.RichTextBox1.hwnd, EM_SETSCROLLPOS, 0, P 'make the other RichTextBox match
            End Select
            pMyWindowProc = CallWindowProc(glPrevWndProc2, hw, uMsg, wParam, lParam)
    End Select
End Function

