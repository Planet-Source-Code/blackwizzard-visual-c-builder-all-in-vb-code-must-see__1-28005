Attribute VB_Name = "modAFS"
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Const RGN_DIFF = 4
Public Const SC_CLICKMOVE = &HF012&
                
Public Const WM_SYSCOMMAND = &H112

Dim CurRgn, TempRgn As Long

Public Function AutoFormShape(bg As Form, transColor)
Dim x, y As Integer

CurRgn = CreateRectRgn(0, 0, bg.ScaleWidth, bg.ScaleHeight)

While y <= bg.ScaleHeight
    While x <= bg.ScaleWidth
        If GetPixel(bg.hdc, x, y) = transColor Then
            TempRgn = CreateRectRgn(x, y, x + 1, y + 1)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
        End If
        x = x + 1
    Wend
        y = y + 1
        x = 0
Wend
success = SetWindowRgn(bg.hwnd, CurRgn, True)
DeleteObject (CurRgn) 'ne pas supprimer!!! ça libere les ressources!

End Function


