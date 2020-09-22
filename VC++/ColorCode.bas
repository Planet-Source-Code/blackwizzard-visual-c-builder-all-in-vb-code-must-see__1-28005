Attribute VB_Name = "ColorCode"
Option Explicit
Public m_TextCol As String
Public m_AttribCol As String
Public m_TagCol As String
Public m_CommentCol As String
Public m_AspCol As String

Public Sub HtmlHighlight()
On Error Resume Next
    'frmMain.trapUndo = False
    ' Color Html and asp
    HtmlColorCode
    
    ' Move back to the start of the thing
    Form1.js.SelStart = 0
    'frmMain.trapUndo = True
End Sub

Public Function KeyPressEvent(KeyAscii As Integer) As Integer
    Static cInAttrib As Boolean, cInTag As Boolean
    Static cInAttribQuote As Boolean, cTypedIn As Boolean
    Static cInComment As Boolean
    Static cInASP As Boolean
    Static cInFunction As Boolean
    
    'frmMain.trapUndo = False
    
    Dim cChar As String
'form1.js
    With Form1.js
        cChar = Chr$(KeyAscii)
        
        If cInTag = False And cInAttrib = False And cInComment = False And cInASP = False Then
            .SelColor = m_TextCol
        End If

        If cInTag = True And (cInAttrib = True Or cInAttribQuote = True) Then
            .SelColor = m_AttribCol
        End If

        If cChar = "<" Then
            .SelColor = m_TagCol
            cInTag = True
            cTypedIn = True
        End If

        If cChar = "=" And cInTag = True Then
            cInAttrib = True
        End If

        If cChar = Chr$(34) And cInAttrib = True And cInAttribQuote = True Then
            cInAttrib = False
            cInAttribQuote = False
        ElseIf cChar = Chr$(34) And cInAttrib = True And cInAttribQuote = False Then
            cInAttribQuote = True
        End If

        If cChar = " " And (cInAttribQuote = False And cInTag = True) Then
            .SelColor = m_TagCol
            cInAttrib = False
        End If

        If cChar = "!" And Mid$(.Text, .SelStart, 1) = "<" Then

            .SelStart = .SelStart - 1
            .SelLength = 1
            .SelColor = m_CommentCol
            .SelText = "<!--"

            cInTag = False
            cInAttrib = False
            cInASP = False
            cInComment = True

            KeyAscii = 0
        End If
        
        If cChar = "%" And Mid$(.Text, .SelStart, 1) = "<" Then

            .SelStart = .SelStart - 1
            .SelLength = 1
            .SelColor = m_AspCol
            .SelText = "<%"

            cInTag = False
            cInAttrib = False
            cInASP = True
            cInComment = False

            KeyAscii = 0
        End If

        If cChar = ">" Then
            If cInComment = False And cInASP = True Then
                .SelColor = m_AspCol
            ElseIf cInComment = True And cInASP = False Then
                .SelColor = m_CommentCol
            ElseIf cInComment = False And cInASP = False Then
                .SelColor = m_TagCol
            End If
            
            cInTag = False
            cInASP = False
            cInComment = False
            cTypedIn = False
        End If

    End With

    KeyPressEvent = KeyAscii
    
    'frmMain.trapUndo = True
ErrExit:
    Exit Function
End Function

' Insert text w/tag coloring if necessary

Public Sub InsertTag(Tag$, StopAsp As Boolean)
Dim S As Long
    'frmMain.trapUndo = False
    S = Form1.js.SelStart
    If Len(Form1.js.SelText) > 0 Then Form1.js.SelText = ""
    Form1.js.SelText = Tag$
    
    If StopAsp = True Then
        HtmlColorCode S, S + Len(Tag), True
Else

        HtmlColorCode S, S + Len(Tag), False

    End If
    
    'frmMain.trapUndo = True
End Sub

' Insert Asp code with asp coloring

Public Sub InsertAspTag(Tag$)
Dim U As Long
    U = frmMain.rtfText.SelStart
    If Len(frmMain.rtfText.SelText) > 0 Then frmMain.rtfText.SelText = ""
    frmMain.rtfText.SelText = Tag$
    
    frmMain.trapUndo = False
    ASPColorCode U, U + Len(Tag)
    frmMain.trapUndo = True
End Sub

' This function determines whether the caret is currently outside a tag. This was a royal pain in the ass.

Public Function IsOutsideTag()
On Error Resume Next
Dim LastGT As Long, LastLT As Long, NextGT As Long, NextLT As Long
Dim EndTag As Long, StartTag As Long
Dim txt$, Start As Long, Start2 As Long
Dim InMainTag As Boolean, InEndTag As Boolean
    
    txt = Form1.js.Text
    Start = Form1.js.SelStart
    
    If Start = 0 Then
        m_TextCol = vbBlack
        Exit Function
    Else
        EndTag = InStr(Start + 1, txt, ">")
        StartTag = InStr(Start + 1, txt, "<")

        If StartTag > EndTag Then
            InMainTag = True
        Else
            InMainTag = False
        End If
        
        LastLT = RevInStr(txt, "<", Start + 1)
        LastGT = RevInStr(txt, ">", Start + 1)

        If LastLT < LastGT Then
            InEndTag = True
        Else
            InEndTag = False
        End If

        If InMainTag = True Or InEndTag = True Then
            m_TextCol = Form1.js.SelColor
        Else
            m_TextCol = vbBlack
        End If
    End If
End Function

' ##########################################################################################
' These are the main color coding functions. These are not called ever by the user.
' ##########################################################################################

' This is the main color coding function. This does everything html, comments, and attributes. It also calls
' the ASP color coding function if nessasary

Public Function HtmlColorCode(Optional startchar As Long = 1, Optional endchar As Long = -1, Optional StopAsp As Boolean = False)
On Error GoTo ErrHandler
    ' These are the variables for the tags for ColorCoding
    Dim CommentOpenTag As String
    Dim CommentCloseTag As String

    Dim oldselstart As Long, oldsellen As Long
    
    ' These are place holders for the color coding
    Dim tag_open As Long
    Dim tag_close As Long
    Dim bef As String
    Dim Curr As String
    Dim ci As Integer
    'frmMain.trapUndo = False
    
    ' Find out where the cursor is
    oldselstart = Form1.js.SelStart
    oldsellen = Form1.js.SelLength
    
    If endchar = -1 Then endchar = Len(Form1.js.Text)
    If startchar = 0 Then startchar = 1

    ' These are the close tags for colorcoding
    
    tag_close = startchar
    
    ' Lets try to hide the color coding from the user:
    Form1.js.HideSelection = True
    ci = 0
    Form1.js.Visible = False
    'frmDocument.PrgBar.Visible = True
    ' Now lets loop through the tags and color code it
    Do
    ci = ci + 1
    If ci = 100 Then
    ci = 0
    End If
    
    'frmDocument.PrgBar.Value = ci
        ' See where the next tag starts. if any
        tag_open = InStr(tag_close, Form1.js.Text, "<")
        
        'If so, then color it...
        If tag_open <> 0 Then  'Found a tag
            
            'Get everything before the tag we're on...
            bef = Mid$(Form1.js.Text, 1, tag_open - 1)
            
            'Find the end of the next tag...
            tag_close = InStr(tag_open, Form1.js.Text, ">")

            'Get the current HTML tag...
            Curr = Mid$(Form1.js.Text, tag_open, tag_close - tag_open + 1)
            
            If tag_close <> 0 Then
                Select Case Left$(Curr, 3)
                    Case "<!-"
                        ' It's a comment...
                        tag_close = InStr(tag_open, Form1.js.Text, "->") + 1
                            Form1.js.SelStart = tag_open - 1
                            Form1.js.SelLength = tag_close - tag_open + 1
                            Form1.js.SelColor = m_CommentCol
                    Case Else
                        ' This colors basic Html tags and then colors the attributes
                        cycleAttrib Curr, tag_open, tag_close
                End Select
            End If
            
            If tag_close = 0 Or tag_close >= endchar Then
                ' If we are coloring tags and it's over the end tag then
                ' get me out of this loop and don't color anymore
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    'frmDocument.PrgBar.Visible = False
    Form1.js.Visible = True
    ' Color ASP Stuff only if we need to. We have a special function for coloring ASP tags so we won't
    ' worry if this deals with it or not.
    If StopAsp = False Then
        ASPColorCode startchar, endchar
    End If
    
    Form1.js.SelStart = oldselstart
    Form1.js.SelLength = oldsellen
    Form1.js.HideSelection = False
    Form1.js.SetFocus
    
    'frmMain.trapUndo = True
    Exit Function
    
ErrHandler:
    Exit Function
End Function

' This function colorizes ASP code

Private Function ASPColorCode(Optional startchar As Long = 1, Optional endchar As Long = -1)
On Error GoTo ErrHandler
    Dim oldselstart As Long, oldsellen As Long
    
    ' These are place holders for the color coding
    Dim tag_open As Long
    Dim tag_close As Long
    Dim bef As String
    Dim Curr As String
    
    'frmMain.trapUndo = False
    
    ' Find out where the cursor is
    oldselstart = Form1.js.SelStart
    oldsellen = Form1.js.SelLength
    
    If endchar = -1 Then endchar = Len(Form1.js.Text)
    If startchar = 0 Then startchar = 1

    ' These are the close tags for colorcoding
    
    tag_close = startchar
    
    ' Lets try to hide the color coding from the user:
    Form1.js.HideSelection = True
    
    ' Now lets loop through the tags and color code it
    Do
        ' See where the next tag starts. if any
        tag_open = InStr(tag_close, Form1.js.Text, "<%")
        
        'If so, then color it...
        If tag_open <> 0 Then  'Found a tag
            
            'Get everything before the tag we're on...
            bef = Mid$(Form1.js.Text, 1, tag_open - 1)
            
            'Find the end of the next tag...
            tag_close = InStr(tag_open, Form1.js.Text, "%>")

            'Get the current HTML tag...
            Curr = Mid$(Form1.js.Text, tag_open, tag_close - tag_open + 1)
            
            If tag_close <> 0 Then
                Select Case Left$(Curr, 2)
                    Case "<%"
                        ' It's asp
                        tag_close = InStr(tag_open, Form1.js.Text, "%>") + 1
                            Form1.js.SelStart = tag_open - 1
                            Form1.js.SelLength = tag_close - tag_open + 1
                            Form1.js.SelColor = m_AspCol
                    Case Else
                        ' it's not an asp tag so do nothing
                End Select
            End If
            
            If tag_close = 0 Or tag_close >= endchar Then
                ' If we are coloring tags and it's over the end tag then
                ' get me out of this loop and don't color anymore
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    Form1.js.SelStart = oldselstart
    Form1.js.SelLength = oldsellen
    Form1.js.HideSelection = False
    Form1.js.SetFocus
    
    'frmMain.trapUndo = True
    
    Exit Function
    
ErrHandler:
    Exit Function
End Function

' This cycles through the html and comes back with the right tag colors for the tag and all of it's
' attributes

Private Function cycleAttrib(CurrTag As String, opentag As Long, closetag As Long)
    
    Dim fPos As Long, sPos As Long, qPos As Long, qnPos As Long, aPos As Long, tBeg As Long, tEnd As Long
    Dim isFirstCycle As Boolean
    Dim eTag As String
    Dim sPosTxt As String
    Dim LeftOver As Long
    Dim EndTag As Long, QuotePos As Long, QuoteEndPos As Long
    
    'frmDocument.trapUndo = False
    
    eTag = CurrTag
    isFirstCycle = True

    Do While Len(eTag) > 0
        fPos = InStr(1, eTag, "=")

        If (fPos = 0 And isFirstCycle = True) Then
            ' This just checks to see if it's a basic html tag w/ no attributes and if so colors that
            ' without going through the rest of the junk.
            Form1.js.SelStart = opentag - 1
            Form1.js.SelLength = closetag - opentag + 1
            Form1.js.SelColor = m_TagCol
            Exit Function
        ' It looks like we have an attribute. Here comes the hard part...
        ElseIf fPos <> 0 Then 'Put in the color info...
            If Left$(eTag, 1) = "<" Then
                ' This brings back the entire tag. something like:
                ' <img src="blah.jpg" onclick="blah">
                ' and then color codes the entire thing
                tBeg = opentag
                tEnd = opentag + fPos

                ' Color Code the entire tag first
                Form1.js.SelStart = tBeg - 1
                Form1.js.SelLength = closetag - tBeg + 1
                Form1.js.SelColor = m_TagCol

                ' This brings back the text that is past the attribute. in the previous example:
                ' "blah.jpg" onclick="blah">
                eTag = Mid$(eTag, fPos + 1)
                LeftOver = closetag - Len(eTag)
            End If
        End If
        
        'Find the first instance of a space in the
        'part of the tag that we have left...
        sPos = InStr(1, eTag, Chr$(32))

        'Gets the text up to the next space...
        sPosTxt = Mid$(eTag, 1, sPos)
        
        'Checks to see if there's a quote in the text...
        qPos = InStr(1, sPosTxt, Chr$(34))

        'If there's a quote found, then we need to find
        'its end...
        If qPos <> 0 Then
            'Look for the next quote...
            qnPos = InStr(2, eTag, Chr$(34))

            If qnPos <> 0 Then
                sPosTxt = Mid$(eTag, 1, qnPos)
            End If
        End If

        LeftOver = closetag - Len(eTag)
        Form1.js.SelStart = LeftOver
        Form1.js.SelLength = Len(sPosTxt)
        Form1.js.SelColor = m_AttribCol
        
        'Truncates the tag so there's no attrib value left...
        eTag = Mid$(eTag, Len(sPosTxt) + 1)

        'Find the next position of an equal sign...
        sPos = InStr(1, eTag, "=")

        'If there's no =, then we know we're on the last
        'attrib value, so we need to put in some final
        'info...all that's left is something like:
        '"#ffffff">
        If sPos = 0 Then
            'Put in the attrib color before the ">"
            'if it's the last attribute...
            eTag = Mid$(eTag, 1, Len(eTag) - 1)

            'Insert the RTF info...
            'bef = bef & infoRTF & AttribInfo & eTag
            Form1.js.SelStart = LeftOver
            Form1.js.SelLength = Len(eTag)
            Form1.js.SelColor = m_AttribCol

            'Truncate the end...
            sPos = Len(eTag)
            Exit Do
        End If

        'Truncates the tag appropriately...
        eTag = Mid$(eTag, sPos + 1)
        isFirstCycle = False

        'If there's nothing left, then we need to exit
        'the loop so it doesn't loop infinitely...
        If sPos = 0 And qPos = 0 Then Exit Do
    Loop
    
    'frmMain.trapUndo = True
    Exit Function
End Function

