Attribute VB_Name = "translucide"
Option Explicit

'form translucide
Declare Function ReleaseDC Lib "USER32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
Declare Function GetDesktopWindow Lib "USER32" () As Long
Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function JoueWav Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SRCCOPY = &HCC0020

'voir l'ombre de deplacement...
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'fonctions recurentes
Global iRecursion As Boolean
Global tColor As Long

Public Sub DragForm(Who As Form) 'permet de deplacer la form sans barre de titre!
On Local Error Resume Next
'deplacer la form
Call ReleaseCapture
Call SendMessage(Who.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Public Sub MakeTranslucent(Who As Form, Optional tColor As Long)
'who est le  om de la form à rendre tanslucide.
'ici, who est "form1"
'tcolor reste "tcolor", on va definir la couleur ci-dessous...
On Local Error Resume Next

Dim HW As Long
Dim HA As Long
Dim iLeft As Integer
Dim iTop As Integer
Dim iWidth As Integer
Dim iHeight As Integer

If IsMissing(tColor) Or tColor = 0 Then
'ici on entre la couleur translucide...
    tColor = RGB(0, 0, 255) 'en hexa: 0000FF => bleu standard W3C
End If

Who.AutoRedraw = True
Who.Hide

DoEvents

HW = GetDesktopWindow()
HA = GetDC(HW)

'prend les mesures de la form...
iLeft = Who.Left / Screen.TwipsPerPixelX
iTop = Who.Top / Screen.TwipsPerPixelY '+ 25    si la form possede une barre de titre...
iWidth = Who.ScaleWidth
iHeight = Who.ScaleHeight

'met l'image de l'ecrant en image de font de la form :
'c'est ça qui crée l'illusion que la form est translucide!
Call BitBlt(Who.hdc, 0, 0, iWidth, iHeight, HA, iLeft, iTop, SRCCOPY) 'iLeft + 4    si la form possede une barre de titre...

'montre la forme (qui etait caché pour lui introduire l'image de font)
Who.Picture = Who.Image
Who.Show

Call ReleaseDC(HW, HA)

'Ajouter la couleur de font...
Who.DrawMode = 9
Who.ForeColor = tColor
Who.Line (0, 0)-(iWidth, iHeight), , BF

End Sub


Public Sub PicTrans(Who As PictureBox, Optional tColor As Long)
'#Si vous voulez faire de meme avec une PictureBox,
'#(rendre translucide)
'#reprennez les commentaires de la form...

On Local Error Resume Next

Dim HW As Long
Dim HA As Long
Dim iLeft As Integer
Dim iTop As Integer
Dim iWidth As Integer
Dim iHeight As Integer

If IsMissing(tColor) Or tColor = 0 Then
    tColor = RGB(255, 255, 255)
End If

Who.AutoRedraw = True


DoEvents

HW = GetDesktopWindow()
HA = GetDC(HW)

iLeft = Who.Left / Screen.TwipsPerPixelX
iTop = Who.Top / Screen.TwipsPerPixelY
iWidth = Who.ScaleWidth
iHeight = Who.ScaleHeight

Call BitBlt(Who.hdc, 0, 0, iWidth, iHeight, HA, iLeft, iTop, SRCCOPY)

Who.Picture = Who.Image

Call ReleaseDC(HW, HA)

Who.DrawMode = 9
Who.ForeColor = tColor
Who.Line (0, 0)-(iWidth, iHeight), , BF

End Sub
