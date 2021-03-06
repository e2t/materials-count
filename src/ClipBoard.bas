Attribute VB_Name = "ClipBoard"
Option Explicit

'https://www.thespreadsheetguru.com/blog/2015/1/13/how-to-use-vba-code-to-copy-text-to-the-clipboard

'Handle 64-bit and 32-bit Office
#If Win64 Then
    Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hmem As LongPtr) As LongPtr
    Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hmem As LongPtr) As LongPtr
    Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As LongPtr
    Declare PtrSafe Function CloseClipboard Lib "User32" () As LongPtr
    Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As LongPtr
    Declare PtrSafe Function EmptyClipboard Lib "User32" () As LongPtr
    Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
    Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As LongPtr, ByVal hmem As LongPtr) As LongPtr
#Else
    Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Declare Function CloseClipboard Lib "User32" () As Long
    Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
    Declare Function EmptyClipboard Lib "User32" () As Long
    Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
#End If

'PURPOSE: API function to copy text to clipboard
'SOURCE: www.msdn.microsoft.com/en-us/library/office/ff192913.aspx
Sub CopyInClipBoard(MyString As String)
    Const GHND = &H42
    Const CF_TEXT = 1
    Const GMEM_MOVEABLE As Long = &H2
    Const CF_LOCALE As Long = 16

#If Win64 Then
    Dim hGlobalMemory As LongPtr
    Dim lpGlobalMemory As LongPtr
    Dim hClipMemory As LongPtr
    Dim X As LongPtr
    Dim hmem As LongPtr
#Else
    Dim hGlobalMemory As Long
    Dim lpGlobalMemory As Long
    Dim hClipMemory As Long
    Dim X As Long
    Dim hmem As Long
#End If

    'Allocate moveable global memory
    hGlobalMemory = GlobalAlloc(GHND, LenB(MyString) + 1)
    
    'Lock the block to get a far pointer to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    
    'Copy the string to this global memory.
    lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)
    
    'Unlock the memory.
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "Could not unlock memory location. Copy aborted."
        GoTo OutOfHere2
    End If

    'Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
        MsgBox "Could not open the Clipboard. Copy aborted."
        Exit Sub
    End If

    'Clear the Clipboard.
    X = EmptyClipboard()

    'Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)
  
    'Copy default locale
    hmem = GlobalAlloc(GMEM_MOVEABLE, 4)
    If hmem <> 0 Then
        SetClipboardData CF_LOCALE, hmem
    End If

OutOfHere2:
    If CloseClipboard() = 0 Then
        MsgBox "Could not close Clipboard."
    End If
End Sub

