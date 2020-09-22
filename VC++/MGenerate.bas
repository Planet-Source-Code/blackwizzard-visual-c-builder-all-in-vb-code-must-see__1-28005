Attribute VB_Name = "MGenerate"
'ce module genere le code...

'Public Declare Function ReleaseCapture Lib "USER32" () As Long
Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const MousePress = &HA1
Public Const SizeN = 12
Public Const SizeS = 15
Public Const SizeW = 10
Public Const SizeE = 11
Public Const SizeNW = 13
Public Const SizeSW = 16
Public Const SizeNE = 14
Public Const SizeSE = 17



Dim I As Integer

Dim xName As String
Dim xValue As String
Dim xBgcol As String
Dim xFgcol As String
Dim xTitle As String
Dim xAlt As String
Dim xRo As String
Dim xBorder As String
Dim xPath As String
Dim xClick As String
Dim xDblclick As String
Dim xOver As String
Dim xOut As String
Dim xDown As String
Dim xMove As String
Dim xLoad As String
Dim xUnload As String
Dim xKeydown As String
Dim xKeyup As String
Dim xKeypress As String
Dim xSelect As String
Dim xFocus As String
Dim xChange As String
Dim xBlur As String
Dim xError  As String
Dim xAbord As String
Dim hwcb As Integer



Public Sub Generate()
SetVariables "Page", "N/A"
Form1.RTCode.text = ""
'Form1.RTCode.Text = "<html>" & Form1.space.Text & _
"<head>" & "<title>" & xTitle & "</title>" & Form1.space.Text & _
"<script language='javascript' src='vhtml.js'></SCRIPT></head>" & "<body" & xFgcol & xBgcol & xClick & xDblclick & xDown & xOut & xOver & xKeypress & xKeydown & xKeyup & xLoad & xUnload & ">"
Form1.RTCode.text = Form1.BWC0.text
SetVariables "Page", "N/A"
Form1.RTCode.text = Form1.RTCode.text & "winhWnd = CreateWindowEx(0, " & Chr(34) & "BWVCPP" & Chr(34) & ", " & Chr(34) & getstring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "title") & Chr(34) & ", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT," & getstring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "width") & ", " & getstring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "height") & ", HWND_DESKTOP, NULL, FirstInstance, NULL);"
Form1.RTCode.text = Form1.RTCode.text & Form1.BWC1.text
For I = 1 To Form1.cb.Count - 1
Add2Coding Form1.cb(I).Left, Form1.cb(I).Top, "cb", I, Form1.cb(I).Width, Form1.cb(I).Height
Next I
For I = 1 To Form1.cc.Count - 1
Add2Coding Form1.cc(I).Left, Form1.cc(I).Top, "cc", I, Form1.cc(I).Width, Form1.cc(I).Height
Next I
For I = 1 To Form1.ci.Count - 1
Add2Coding Form1.ci(I).Left, Form1.ci(I).Top, "ci", I, Form1.ci(I).Width, Form1.ci(I).Height
Next I
For I = 1 To Form1.cli.Count - 1
Add2Coding Form1.cli(I).Left, Form1.cli(I).Top, "cli", I, Form1.cli(I).Width, Form1.cli(I).Height
Next I
For I = 1 To Form1.clist.Count - 1
Add2Coding Form1.clist(I).Left, Form1.clist(I).Top, "clist", I, Form1.clist(I).Width, Form1.clist(I).Height
Next I
For I = 1 To Form1.ccombo.Count - 1
Add2Coding Form1.ccombo(I).Left, Form1.ccombo(I).Top, "ccombo", I, Form1.ccombo(I).Width, Form1.ccombo(I).Height
Next I
For I = 1 To Form1.ct.Count - 1
Add2Coding Form1.ct(I).Left, Form1.ct(I).Top, "ct", I, Form1.ct(I).Width, Form1.ct(I).Height
Next I
For I = 1 To Form1.cta.Count - 1
Add2Coding Form1.cta(I).Left, Form1.cta(I).Top, "cta", I, Form1.cta(I).Width, Form1.cta(I).Height
Next I
For I = 1 To Form1.cp.Count - 1
Add2Coding Form1.cp(I).Left, Form1.cp(I).Top, "cp", I, Form1.cp(I).Width, Form1.cp(I).Height
Next I
For I = 1 To Form1.ch.Count - 1
Add2Coding Form1.ch(I).Left, Form1.ch(I).Top, "ch", I, Form1.ch(I).Width, Form1.ch(I).Height
Next I
For I = 1 To Form1.cl.Count - 1
Add2Coding Form1.cl(I).X1, Form1.cl(I).Y1, "cl", I, (Form1.cl(I).X2 - Form1.cl(I).X1), 0
Next I
'Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"</body></html>"
Form1.RTCode.text = Form1.RTCode.text & Form1.BWC2.text

For I = 1 To Form1.cb.Count - 1
SetVariables "cb", I
Form1.RTCode.text = Form1.RTCode.text & "SetWindowText(cb" & I & ", " & Chr(34) & xValue & Chr(34) & ");" & vbCrLf
Next I

For I = 1 To Form1.ct.Count - 1
SetVariables "ct", I
Form1.RTCode.text = Form1.RTCode.text & "SetWindowText(ct" & I & ", " & Chr(34) & xValue & Chr(34) & ");" & vbCrLf
Next I

For I = 1 To Form1.cta.Count - 1
SetVariables "cta", I
Form1.RTCode.text = Form1.RTCode.text & "SetWindowText(cta" & I & ", " & Chr(34) & xValue & Chr(34) & ");" & vbCrLf
Next I

For I = 1 To Form1.cli.Count - 1
SetVariables "cli", I
Form1.RTCode.text = Form1.RTCode.text & "SetWindowText(cli" & I & ", " & Chr(34) & xValue & Chr(34) & ");" & vbCrLf
Next I

For I = 1 To Form1.ci.Count - 1
SetVariables "ci", I
Form1.RTCode.text = Form1.RTCode.text & "SendMessage(ci" & I & ", STM_SETIMAGE, IMAGE_ICON, (LPARAM)LoadIcon(NULL, IDI_APPLICATION));" & vbCrLf
Next I

For I = 1 To Form1.clist.Count - 1
SetVariables "clist", I
Form1.RTCode.text = Form1.RTCode.text & "SendMessage(clist" & I & ", LB_ADDSTRING, 0, (LPARAM)(LPCTSTR)" & Chr(34) & xValue & Chr(34) & ");" & vbCrLf
Next I

For I = 1 To Form1.ccombo.Count - 1
SetVariables "ccombo", I
Form1.RTCode.text = Form1.RTCode.text & "SendMessage(ccombo" & I & ", LB_ADDSTRING, 0, (LPARAM)(LPCTSTR)" & Chr(34) & xValue & Chr(34) & ");" & vbCrLf
Next I

Form1.RTCode.text = Form1.RTCode.text & Form1.BWC3.text
Open "c:\preview.html" For Output As #1
Print #1, Form1.RTCode.text
Close #1
Open "c:\vhtml.js" For Output As #1
Print #1, Form1.js.text
Close #1

Form1.RTCode.Colorize
End Sub

Public Sub Add2Coding(X As Integer, Y As Integer, Wtype As String, Index As Integer, Width As Integer, Height As Integer)
SetVariables Wtype, Index
Select Case Wtype


Case "cb"
Form1.RTCode.text = Form1.RTCode.text & Form1.space.text & _
"HWND " & Wtype & Index & " =  CreateWindowEx(0, " & Chr(34) & "BUTTON" & Chr(34) & ", " & Chr(34) & "" & Chr(34) & ", WS_VISIBLE|WS_CHILD|BS_PUSHBUTTON|BS_NOTIFY|BS_TEXT|ES_NOHIDESEL, " & X & "," & Y & "," & Width & "," & Height & ", winhWnd, (HMENU)ID_BUTTON, FirstInstance, NULL);"


Case "cc"
'Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"HWND " & Wtype & Index & " =  CreateWindowEx(0, " & Chr(34) & "OPTIONBOX" & Chr(34) & ", " & Chr(34) & "" & Chr(34) & ", WS_VISIBLE|WS_CHILD|BS_PUSHBUTTON|BS_NOTIFY|BS_TEXT|ES_NOHIDESEL, " & x & "," & y & "," & Width & "," & Height & ", winhWnd, (HMENU)ID_BUTTON, FirstInstance, NULL);"


Case "ci"
Form1.RTCode.text = Form1.RTCode.text & Form1.space.text & _
"HWND " & Wtype & Index & " =  CreateWindowEx(0, " & Chr(34) & "STATIC" & Chr(34) & ", " & Chr(34) & "" & Chr(34) & ", WS_VISIBLE|WS_CHILD|SS_ICON, " & X & "," & Y & "," & Width & "," & Height & ", winhWnd, (HMENU)ID_BUTTON, FirstInstance, NULL);"


Case "cli"
Form1.RTCode.text = Form1.RTCode.text & Form1.space.text & _
"HWND " & Wtype & Index & " =  CreateWindowEx(0, " & Chr(34) & "STATIC" & Chr(34) & ", " & Chr(34) & "" & Chr(34) & ", WS_VISIBLE|WS_CHILD, " & X & "," & Y & "," & Width & "," & Height & ", winhWnd, (HMENU)ID_BUTTON, FirstInstance, NULL);"


Case "clist"
Form1.RTCode.text = Form1.RTCode.text & Form1.space.text & _
"HWND " & Wtype & Index & " = CreateWindowEx(WS_EX_OVERLAPPEDWINDOW, " & Chr(34) & "LISTBOX" & Chr(34) & ", " & Chr(34) & "" & Chr(34) & ",WS_VISIBLE|WS_CHILD|CBS_AUTOHSCROLL|CBS_DISABLENOSCROLL|CBS_HASSTRINGS|CBS_SORT|CBS_DROPDOWN," & X & "," & Y & "," & Width & "," & Height & ", winhWnd, (HMENU)ID_EDITBOX, FirstInstance, NULL);"



Case "ccombo"
Form1.RTCode.text = Form1.RTCode.text & Form1.space.text & _
"HWND " & Wtype & Index & " = CreateWindowEx(WS_EX_OVERLAPPEDWINDOW, " & Chr(34) & "COMBOBOX" & Chr(34) & ", " & Chr(34) & "" & Chr(34) & ",WS_VISIBLE|WS_CHILD|CBS_AUTOHSCROLL|CBS_DISABLENOSCROLL|CBS_HASSTRINGS|CBS_SORT|CBS_DROPDOWN," & X & "," & Y & "," & Width & "," & Height & ", winhWnd, (HMENU)ID_EDITBOX, FirstInstance, NULL);"

'WS_VISIBLE|WS_CHILD|ES_AUTOHSCROLL|ES_NOHIDESEL|ES_AUTOVSCROLL|ES_MULTILINE|ES_WANTRETURN|ES_LEFT|WS_VSCROLL|WS_HSCROLL

Case "ct"
Form1.RTCode.text = Form1.RTCode.text & Form1.space.text & _
"HWND " & Wtype & Index & " = CreateWindowEx(WS_EX_OVERLAPPEDWINDOW, " & Chr(34) & "EDIT" & Chr(34) & ", " & Chr(34) & "" & Chr(34) & ", WS_VISIBLE|WS_CHILD|ES_AUTOHSCROLL|ES_NOHIDESEL," & X & "," & Y & "," & Width & "," & Height & ", winhWnd, (HMENU)ID_EDITBOX, FirstInstance, NULL);"


Case "cta"
Form1.RTCode.text = Form1.RTCode.text & Form1.space.text & _
"HWND " & Wtype & Index & " = CreateWindowEx(WS_EX_OVERLAPPEDWINDOW, " & Chr(34) & "EDIT" & Chr(34) & ", " & Chr(34) & "" & Chr(34) & ", WS_VISIBLE|WS_CHILD|ES_AUTOHSCROLL|ES_NOHIDESEL|ES_AUTOVSCROLL|ES_MULTILINE|ES_WANTRETURN|ES_LEFT|WS_VSCROLL|WS_HSCROLL," & X & "," & Y & "," & Width & "," & Height & ", winhWnd, (HMENU)ID_EDITBOX, FirstInstance, NULL);"


Case "cp"
'Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"HWND " & Wtype & Index & " = CreateWindowEx(WS_EX_OVERLAPPEDWINDOW, " & Chr(34) & "EDIT" & Chr(34) & ", " & Chr(34) & "" & Chr(34) & ", WS_VISIBLE|WS_CHILD|ES_AUTOHSCROLL|ES_NOHIDESEL," & x & "," & y & "," & Width & "," & Height & ", winhWnd, (HMENU)ID_EDITBOX, FirstInstance, NULL);"


Case "ch"
'Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"<input type='hidden'" & xName & _
"' style='position:absolute;width:" & Width & ";left:" & x & ";top:" & y & ";" & xValue & "'>"


Case "cl"
'Form1.RTCode.Text = Form1.RTCode.Text & Form1.space.Text & _
"<hr style='position:absolute;width:" & Width & ";left:" & x & ";top:" & y & ";'>"
Case "csub"

End Select
End Sub


Public Sub SetVariables(Wtype As String, Index)
If pro.nam.text <> "" Then
xName = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "name")
End If

If pro.valu.text <> "" Then
    If Wtype = "cli" Or wlist = "clist" Then
    xValue = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "value")
    Else
    xValue = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "value")
    End If
End If

If pro.bgcol.text <> "" Then
xBgcol = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "bgcolor")
End If

If pro.fgcol.text <> "" Then
xFgcol = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "fgcolor")
End If

If pro.title.text <> "" Then
    If Wtype = "Page" Then
    xTitle = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "title")
    Else
    xTitle = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "title")
    End If
End If

If pro.alt.text <> "" Then
    xAlt = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "alt")
End If

If pro.RO.Value = 1 Or pro.RO.Value = 2 Then
xRo = " READONLY" 'getstring(HKEY_CURRENT_USER, "Software\VHTML\" & wType & index, "ro")
ElseIf pro.RO.Value = 0 Then
xRo = ""
End If

If pro.border.text <> "" Then
xBorder = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "border")
End If

If pro.path.text <> "" Then
    If Wtype = "cli" Then
        xPath = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "path")
    Else
        xPath = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "path")
    End If
End If

If pro.onclick.text <> "" Then
xClick = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "click")
End If

If pro.ondblclick.text <> "" Then
xDblclick = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "dblclick")
End If

If pro.over.text <> "" Then
xOver = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "over")
End If

If pro.out.text <> "" Then
xOut = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "out")
End If

If pro.down.text <> "" Then
xDown = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "down")
End If

If pro.mouve.text <> "" Then
xMove = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "move")
End If

If pro.onload.text <> "" Then
xLoad = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "load")
End If

If pro.onunload.text <> "" Then
xUnload = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "unload")
End If

If pro.onkeydown.text <> "" Then
xKeydown = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "keydown")
End If

If pro.onkeyup.text <> "" Then
xKeyup = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "keyup")
End If

If pro.onkeypress.text <> "" Then
xKeypress = " onKeyPress='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "keypress")
End If

If pro.onselect.text <> "" Then
xSelect = " onSelect='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "select")
End If

If pro.onfocus.text <> "" Then
xFocus = " onFocus='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "focus")
End If

If pro.onchange.text <> "" Then
xChange = " onChange='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "change")
End If

If pro.onblur.text <> "" Then
xBlur = " onBlur='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "name")
End If

If pro.onerror.text <> "" Then
xError = " onError='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "error")
End If

If pro.onabord.text <> "" Then
xAbord = " onAbord='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Index, "abord")
End If
End Sub



Public Sub SetVariablesBody()
If pro.bgcol.text <> "" Then
xbgcolor = " bgcolor='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "bgcolor") & "'"
End If

If pro.fgcol.text <> "" Then
xfgcolor = " text='" & getstring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "fgcolor") & "'"
End If
End Sub
