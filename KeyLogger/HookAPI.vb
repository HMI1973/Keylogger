Imports System.Reflection
Imports System.Runtime.CompilerServices.RuntimeHelpers
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar

Public Class HookAPI
    Private Const WH_MOUSE As Integer = 14 '7, 14 = low level mouse hook, For other hook types, obtain these values from Winuser.h in Microsoft SDK.
    Private Const WH_KEYBOARD As Integer = 13 '2, 13 = low level keyboard hook, For other hook types, obtain these values from Winuser.h in Microsoft SDK.

    Private mousehookproc As CallBack
    Private keyboardhookproc As CallBack
    Private p_mouseHook As Integer = 0
    Private p_keyboardHook As Integer = 0
    Private p_callbackFn As PressHandle

    Public m_X As Integer
    Public m_Y As Integer
    Public m_LB As Boolean
    Public m_RB As Boolean
    Public m_MB As Boolean
    Public m_Wheel As Integer
    Public m_KeyCodes As List(Of Integer)

#Region "DLL"
    Public Delegate Sub PressHandle(ByVal NewChar As String) 'use in the class to pass call back function key press/mouse handle
    Private Delegate Function CallBack(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    'Import for the SetWindowsHookEx function.
    <DllImport("user32.dll")>
    Private Shared Function SetWindowsHookEx _
          (ByVal idHook As Integer, ByVal HookProc As CallBack, ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
    End Function
    'Import for the CallNextHookEx function.
    <DllImport("user32.dll")>
    Private Shared Function CallNextHookEx(ByVal idHook As Integer, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    End Function
    'Import for the UnhookWindowsHookEx function.
    <DllImport("user32.dll")>
    Private Shared Function UnhookWindowsHookEx(ByVal idHook As Integer) As Boolean
    End Function
    Private Enum MouseMessages
        WM_Move = &H200
        WM_LeftButtonDown = &H201
        WM_LeftButtonUp = &H202
        WM_LeftDblClick = &H203
        WM_RightButtonDown = &H204
        WM_RightButtonUp = &H205
        WM_RightDblClick = &H206
        WM_MiddleButtonDown = &H207
        WM_MiddleButtonUp = &H208
        WM_MiddleDblClick = &H209
        WM_MouseWheel = &H20A
    End Enum
    Private Enum KeyboardMessages
        WM_KEYDOWN = &H100
        WM_KEYUP = &H101
        WM_SYSKEYDOWN = &H104
        WM_SYSKEYUP = &H105
    End Enum
    'Point structure declaration.
    Private Structure Point
        Public x As Integer
        Public y As Integer
    End Structure
    'MouseHookStruct structure declaration.
    Private Structure MouseHookStruct
        Public pt As Point
        Public mouseData As UInteger
        Public flags As UInteger
        Public time As UInteger
        Public dwExtraInfo As IntPtr
        Public ReadOnly Property WheelDelta() As Integer
            Get
                Dim v As Integer = CInt((mouseData And &HFFFF0000) >> 16)
                If v > SystemInformation.MouseWheelScrollDelta Then v = v - (UShort.MaxValue + 1)
                Return v
            End Get
        End Property
    End Structure
    'KeyboardHookStruct structure declaration.
    Private Structure KeyboardHookStruct
        Public vkCode As UInteger
        Public scanCode As UInteger
        Public flags As UInteger
        Public time As UInteger
        Public dwExtraInfo As IntPtr
    End Structure
#End Region
#Region "Private Functions"
    Private Sub Init(ByVal Optional mycallbackFN As PressHandle = Nothing)
        m_X = 0
        m_Y = 0
        m_Wheel = 0
        m_LB = False
        m_RB = False
        m_MB = False
        m_KeyCodes = New List(Of Integer)
        p_mouseHook = 0
        p_keyboardHook = 0
        p_callbackFn = mycallbackFN
    End Sub
    Private Sub updateWheel(ByVal delta As Integer)
        If delta > 0 Then
            If m_Wheel >= 0 Then m_Wheel += 1 Else m_Wheel = 1
        End If
        If delta < 0 Then
            If m_Wheel <= 0 Then m_Wheel -= 1 Else m_Wheel = -1
        End If
        If delta = 0 Then m_Wheel = 0
    End Sub
    Private Function addKeyboardCode(ByVal KeyCode As Integer) As Boolean
        For i = 0 To m_KeyCodes.Count - 1
            If m_KeyCodes(i) = KeyCode Then Return False
        Next
        m_KeyCodes.Add(KeyCode)
        Return True
    End Function
    Private Sub removeKeyboardCode(ByVal KeyCode As Integer)
        For i = 0 To m_KeyCodes.Count - 1
            If m_KeyCodes(i) = KeyCode Then
                m_KeyCodes.RemoveAt(i)
                Return
            End If
        Next
    End Sub
    Private Function MapKey(ByVal vkey As Integer) As String
        Console.WriteLine(vkey)
        If vkey = 48 And isShift() Then Return ")"
        If vkey = 49 And isShift() Then Return "!"
        If vkey = 50 And isShift() Then Return "@"
        If vkey = 51 And isShift() Then Return "#"
        If vkey = 52 And isShift() Then Return "$"
        If vkey = 53 And isShift() Then Return "%"
        If vkey = 54 And isShift() Then Return "^"
        If vkey = 55 And isShift() Then Return "&"
        If vkey = 56 And isShift() Then Return "*"
        If vkey = 57 And isShift() Then Return "("
        If vkey = 96 Then Return "0"
        If vkey = 97 Then Return "1"
        If vkey = 98 Then Return "2"
        If vkey = 99 Then Return "3"
        If vkey = 100 Then Return "4"
        If vkey = 101 Then Return "5"
        If vkey = 102 Then Return "6"
        If vkey = 103 Then Return "7"
        If vkey = 104 Then Return "8"
        If vkey = 105 Then Return "9"

        If vkey = 192 And isShift() Then Return "~"
        If vkey = 192 Then Return "`"

        If vkey = 189 And isShift() Then Return "_"
        If vkey = 189 Then Return "-"
        If vkey = 187 And isShift() Then Return "+"
        If vkey = 187 Then Return "="

        If vkey = 219 And isShift() Then Return "{"
        If vkey = 219 Then Return "["
        If vkey = 221 And isShift() Then Return "}"
        If vkey = 221 Then Return "]"

        If vkey = 220 And isShift() Then Return "|"
        If vkey = 220 Then Return "\"
        If vkey = 226 And isShift() Then Return "|"
        If vkey = 226 Then Return "\"

        If vkey = 186 And isShift() Then Return ":"
        If vkey = 186 Then Return ";"
        If vkey = 222 And isShift() Then Return """"
        If vkey = 222 Then Return "'"

        If vkey = 188 And isShift() Then Return "<"
        If vkey = 188 Then Return ","
        If vkey = 190 And isShift() Then Return ">"
        If vkey = 190 Then Return "."
        If vkey = 191 And isShift() Then Return "?"
        If vkey = 191 Then Return "/"


        If vkey = 110 Then Return "."
        If vkey = 111 Then Return "/"
        If vkey = 106 Then Return "*"
        If vkey = 109 Then Return "-"
        If vkey = 107 Then Return "+"
        If vkey = 12 Then Return "5"

        'Arrow
        If vkey = 40 Then Return "{DOWN}"
        If vkey = 37 Then Return "{LEFT}"
        If vkey = 39 Then Return "{RIGHT}"
        If vkey = 38 Then Return "{UP}"

        If vkey = 112 Then Return "{F1}"
        If vkey = 113 Then Return "{F2}"
        If vkey = 114 Then Return "{F3}"
        If vkey = 115 Then Return "{F4}"
        If vkey = 116 Then Return "{F5}"
        If vkey = 117 Then Return "{F6}"
        If vkey = 118 Then Return "{F7}"
        If vkey = 119 Then Return "{F8}"
        If vkey = 120 Then Return "{F9}"
        If vkey = 121 Then Return "{F10}"
        If vkey = 122 Then Return "{F11}"
        If vkey = 123 Then Return "{F12}"

        If vkey = 32 Then Return " " 'Space
        If vkey = 13 Then Return "{ENTER}" 'Enter
        If isAlt(vkey) Then Return "{ALT}"
        If vkey = 33 Then Return "{PAGEUP}"
        If vkey = 34 Then Return "{PAGEDOWN}"
        If vkey = 144 Then Return "{NUM}" 'Num Lock
        If vkey = 45 Then Return "{INS}"
        If vkey = 46 Then Return "{DEL}"
        If vkey = 36 Then Return "{HOM}"
        If vkey = 35 Then Return "{END}"
        If vkey = 20 Then Return "{CAPS}" 'Caps Lock key
        If vkey = 27 Then Return "{ESC}" 'Esc Key
        If vkey = 8 Then Return "{BAK}" 'BACKSPACE
        If vkey = 9 Then Return "{TAB}" 'TAB
        If vkey = 93 Then Return "{WIN2}" 'Application key

        If vkey = 91 Then Return "{WIN}" 'Window Key
        If vkey = 92 Then Return "{WIN}" 'Window Key
        If isShift(vkey) Then Return "{SHIFT}"
        If isCtrl(vkey) Then Return "{CTRL}"

        If (vkey > 64 And vkey < 91) Or (vkey > 47 And vkey < 58) Then
            If Control.IsKeyLocked(Keys.CapsLock) Then
                If isShift() Then Return Char.ToLower(Chr(vkey))
                Return Char.ToUpper(Chr(vkey))
            Else
                If isShift() Then Return Char.ToUpper(Chr(vkey))
                Return Char.ToLower(Chr(vkey))
            End If
        End If
        Return ""
    End Function

    Private Function mouseHandleProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
        Dim MyHookStruct As New MouseHookStruct()
        Dim ActionFlag As Boolean = False
        Dim ActionBTN As Integer = -1

        If (nCode >= 0) Then
            MyHookStruct = CType(Marshal.PtrToStructure(lParam, MyHookStruct.GetType()), MouseHookStruct)
            Select Case wParam
                Case MouseMessages.WM_MouseWheel
                    updateWheel(MyHookStruct.WheelDelta)
                    ActionFlag = True
                Case MouseMessages.WM_Move
                    m_X = MyHookStruct.pt.x
                    m_Y = MyHookStruct.pt.y
                    updateWheel(MyHookStruct.WheelDelta)
                    ActionFlag = True
                Case MouseMessages.WM_LeftButtonDown
                    m_LB = True : ActionFlag = True
                Case MouseMessages.WM_LeftButtonUp
                    m_LB = False : ActionFlag = True : ActionBTN = 1
                Case MouseMessages.WM_LeftDblClick
                Case MouseMessages.WM_RightButtonDown
                    m_RB = True : ActionFlag = True
                Case MouseMessages.WM_RightButtonUp
                    m_RB = False : ActionFlag = True : ActionBTN = 2
                Case MouseMessages.WM_RightDblClick
                Case MouseMessages.WM_MiddleButtonDown
                    m_MB = True : ActionFlag = True
                Case MouseMessages.WM_MiddleButtonUp
                    m_MB = False : ActionFlag = True : ActionBTN = 3
                Case MouseMessages.WM_MiddleDblClick
            End Select
            If Not p_callbackFn Is Nothing And ActionFlag = True Then p_callbackFn("")
        End If

        Return CallNextHookEx(p_mouseHook, nCode, wParam, lParam)
    End Function
    Private Function KeyboardHandleProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
        Dim MyHookStruct As New KeyboardHookStruct()
        Dim NewChar As String = ""

        If (nCode >= 0) Then
            MyHookStruct = CType(Marshal.PtrToStructure(lParam, MyHookStruct.GetType()), KeyboardHookStruct)
            Select Case wParam
                Case KeyboardMessages.WM_KEYDOWN
                    If addKeyboardCode(MyHookStruct.vkCode) Then NewChar = getChar(MyHookStruct.vkCode)
                Case KeyboardMessages.WM_KEYUP
                    removeKeyboardCode(MyHookStruct.vkCode)
                Case KeyboardMessages.WM_SYSKEYDOWN
                    If addKeyboardCode(MyHookStruct.vkCode) Then NewChar = getChar(MyHookStruct.vkCode)
                Case KeyboardMessages.WM_SYSKEYUP
                    removeKeyboardCode(MyHookStruct.vkCode)
            End Select
        End If
        If Not p_callbackFn Is Nothing And Not NewChar = "" Then p_callbackFn.Invoke(NewChar)

        Return CallNextHookEx(p_keyboardHook, nCode, wParam, lParam)
    End Function
    Private Function UnHookMouse() As Boolean
        If Not p_mouseHook.Equals(0) Then
            Dim ret As Boolean = UnhookWindowsHookEx(p_mouseHook)
            If ret.Equals(False) Then Return False
        End If
        Init()
        Return True
    End Function
    Private Function UnHookKeyboard() As Boolean
        If Not p_keyboardHook.Equals(0) Then
            Dim ret As Boolean = UnhookWindowsHookEx(p_keyboardHook)
            If ret.Equals(False) Then Return False
        End If
        Init()
        Return True
    End Function
#End Region

    Public ReadOnly Property isMouseHooked As Boolean
        Get
            If p_mouseHook.Equals(0) Then
                Return False
            Else
                Return True
            End If

        End Get
    End Property
    Public ReadOnly Property isKeyboardHooked As Boolean
        Get
            If p_keyboardHook.Equals(0) Then
                Return False
            Else
                Return True
            End If

        End Get
    End Property
    Public Function isCtrl(ByVal Optional vKey As Integer = -1) As Boolean
        If vKey = -1 Then
            Return isKey(&H11) Or isKey(&HA2) Or isKey(&HA3)
        Else
            If vKey = &H11 Or vKey = &HA2 Or vKey = &HA3 Then Return True
            Return False
        End If
    End Function
    Public Function isAlt(ByVal Optional vKey As Integer = -1) As Boolean
        If vKey = -1 Then
            Return isKey(&H12) Or isKey(&HA4) Or isKey(&HA5)
        Else
            If vKey = &H12 Or vKey = &HA4 Or vKey = &HA5 Then Return True
            Return False
        End If
    End Function
    Public Function isShift(ByVal Optional vKey As Integer = -1) As Boolean
        If vKey = -1 Then
            Return isKey(&H10) Or isKey(&HA0) Or isKey(&HA1)
        Else
            If vKey = &H10 Or vKey = &HA0 Or vKey = &HA1 Then Return True
            Return False
        End If
    End Function
    Public Function isKey(ByVal vKey As Integer) As Boolean
        For i = 0 To m_KeyCodes.Count - 1
            If m_KeyCodes(i) = vKey Then Return True
        Next
        Return False
    End Function
    Public Function getChar(ByVal vKey As Integer) As String
        Return MapKey(vKey)
    End Function

    Public Sub New(ByVal Optional mycallbackFN As PressHandle = Nothing)
        Init(mycallbackFN)
    End Sub
    Public Function HookMouse() As Boolean
        If isMouseHooked Then UnHookMouse()
        If p_mouseHook.Equals(0) Then
            'Keep the reference so that the delegate is not garbage collected.
            mousehookproc = AddressOf mouseHandleProc

            p_mouseHook = SetWindowsHookEx(WH_MOUSE, mousehookproc, Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly.GetModules()(0)), IntPtr.Zero)
            If p_mouseHook.Equals(0) Then Return False
        End If
        Return True
    End Function
    Public Function HookKeyboard() As Boolean
        If isKeyboardHooked Then UnHookKeyboard()
        If p_keyboardHook.Equals(0) Then
            'Keep the reference so that the delegate is not garbage collected.
            keyboardhookproc = AddressOf KeyboardHandleProc

            p_keyboardHook = SetWindowsHookEx(WH_KEYBOARD, keyboardhookproc, Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly.GetModules()(0)), IntPtr.Zero)
            If p_keyboardHook.Equals(0) Then Return False
        End If
        Return True
    End Function
    Public Function UnHook() As Boolean
        If Not UnHookMouse() Then Return False
        Return UnHookKeyboard()
    End Function
End Class
