Attribute VB_Name = "OnTop"
Option Explicit

'API nécessaire pour le mode "toujours visible"
Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SystemParametersInfo Lib "USER32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
'toujours visible


Public Function CopyControl(Control As Variant, Visible As Boolean, Top As Integer, Left As Integer, Width As Integer, Height As Integer)
    Dim NewIndex As Integer
    NewIndex = Control.Count + 1
    Load Control(NewIndex)
    With Control(NewIndex)
        .Visible = Visible
        .Top = Top
        .Left = Left
        .Width = Width
        .Height = Height
    End With
End Function

'Function CopyControlWithResize, optional API
Public Function CopyControlWithResize(Control As Variant, Visible As Boolean, Resize As Boolean, Handles As Variant, Top As Integer, Left As Integer, Width As Integer, Height As Integer)
    Dim NewIndex, x As Integer
    NewIndex = Control.Count + 1
    Load Control(NewIndex)
    With Control(NewIndex)
        .Visible = Visible
        .Top = Top
        .Left = Left
        .Width = Width
        .Height = Height
    End With
    If Resize = True Then
    x = 0
    Do Until x = 8
    Handles(x).Visible = True
    x = x + 1
    Loop
    HandlesMove Control(NewIndex), Handles
    End If
End Function

'Function ControlResize, API
Public Function ControlResize(ControlWithAPIHandle As Control, Handles As Variant, Index As Variant)
On Error Resume Next
    ReleaseCapture
    Select Case Index
        Case 0
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeNW, 0
        Case 1
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeN, 0
        Case 2
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeNE, 0
        Case 7
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeE, 0
        Case 3
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeSE, 0
        Case 6
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeS, 0
        Case 5
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeSW, 0
        Case 4
            SendMessage ControlWithAPIHandle.hWnd, MousePress, SizeW, 0
    End Select
    HandlesMove ControlWithAPIHandle, Handles
End Function

'Function HandlesMove, API
Public Function HandlesMove(ByVal Control As Control, Handles As Variant)
    Form1.rect(0).Left = Control.Left - Form1.rect(0).Width
    Form1.rect(0).Top = Control.Top - Form1.rect(0).Height
    Form1.rect(1).Left = (Control.Width - Form1.rect(7).Width) / 2 + Control.Left
    Form1.rect(1).Top = Control.Top - Form1.rect(1).Height
    Form1.rect(2).Left = Control.Left + Control.Width
    Form1.rect(2).Top = Control.Top - Form1.rect(0).Height
    Form1.rect(7).Left = Control.Left + Control.Width
    Form1.rect(7).Top = (Control.Height - Form1.rect(7).Height) / 2 + Control.Top
    Form1.rect(3).Left = Control.Left + Control.Width
    Form1.rect(3).Top = Control.Top + Control.Height
    Form1.rect(6).Left = (Control.Width - Form1.rect(6).Width) / 2 + Control.Left
    Form1.rect(6).Top = Control.Top + Control.Height
    Form1.rect(5).Left = Control.Left - Form1.rect(5).Width
    Form1.rect(5).Top = Control.Top + Control.Height
    Form1.rect(4).Left = Control.Left - Form1.rect(4).Width
    Form1.rect(4).Top = (Control.Height - Form1.rect(4).Height) / 2 + Control.Top
End Function


Public Function forward(Who As Form) 'who correspond au nom de la form  | exemple: form1
Dim Resultat As Long
Const Flags = &H2 Or &H1 Or &H40 Or &H10
Resultat = SetWindowPos(Who.hWnd, -1, 0, 0, 0, 0, Flags)
End Function

'annuler toujours visible
Public Function backward(Who As Form)
Dim Resultat As Long
Const Flags = &H2 Or &H1 Or &H40 Or &H10
Resultat = SetWindowPos(Who.hWnd, -2, 0, 0, 0, 0, Flags)
End Function

'execution d'un lien...
Public Function WeB(WebPage As String, actualfrmHWND As String)
On Error Resume Next
Dim cod
cod = ShellExecute(actualfrmHWND, vbNullString, WebPage, "", vbNullString, 1)
End Function

'restart...
Public Sub reload(frm As Form)
Unload frm
Load frm
frm.Show
End Sub

'verifier l'existence d'un fichier...
Public Function ExistFile(strPath As String) As Boolean
  Dim fs As Object
  Dim blnFExiste As Boolean
  Set fs = CreateObject("Scripting.FileSystemObject")
  If Not (fs.FileExists(strPath)) Then
    blnFExiste = False
  Else
    blnFExiste = True
  End If
  ExistFile = blnFExiste
End Function
 



Public Sub dragoon(Text As String, title As String)
Form3.Show
Form3.title.Caption = title
Form3.txt.Caption = Text
End Sub
