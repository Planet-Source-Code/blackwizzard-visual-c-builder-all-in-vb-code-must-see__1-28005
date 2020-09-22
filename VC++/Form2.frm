VERSION 5.00
Begin VB.Form pro2 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Property Board"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   2160
   ClientWidth     =   3105
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Cleft 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   60
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox Ctop 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   61
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox Cheight 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   56
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   63
      Text            =   "Top"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   62
      Text            =   "Left"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   59
      Text            =   "Width"
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "Height"
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox Cwidth 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   57
      Top             =   6240
      Width           =   1935
   End
   Begin VB.ComboBox onabord 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   54
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox zonabord 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "onAbord"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.ComboBox onerror 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   46
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox zonerror 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   "onError"
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox zonblur 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "onBlur"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ComboBox onblur 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   49
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox zonchange 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "onChange"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.ComboBox onchange 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   48
      Top             =   5280
      Width           =   1935
   End
   Begin VB.ComboBox onfocus 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   45
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox zonfocus 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "onFocus"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ComboBox onselect 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   40
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox zonselect 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "onSelect"
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ComboBox onkeypress 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   38
      Top             =   4560
      Width           =   1935
   End
   Begin VB.ComboBox onkeyup 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   36
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ComboBox onkeydown 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   30
      Top             =   4080
      Width           =   1935
   End
   Begin VB.ComboBox onunload 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   31
      Top             =   3840
      Width           =   1935
   End
   Begin VB.ComboBox onload 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   32
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox zonkeypress 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "onKeypress"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox zonkeyup 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "onKeyup"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox zonkeydown 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "onKeydown"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox zonunload 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "onUnload"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox zonload 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "onLoad"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ComboBox mouve 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   41
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox zmouve 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "onMouseMove"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ComboBox down 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   27
      Top             =   3120
      Width           =   1935
   End
   Begin VB.ComboBox out 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   26
      Top             =   2880
      Width           =   1935
   End
   Begin VB.ComboBox over 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   25
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox zover 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "onMouseOver"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox zout 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "onMouseOut"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox zdown 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "onMouseDown"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ComboBox ondblclick 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   47
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox zondblclick 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   51
      Text            =   "onDblClick"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox onclick 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   24
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox zclick 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "onClick"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox zpath 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Path"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox path 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox alt 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1185
      Width           =   1935
   End
   Begin VB.TextBox title 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox border 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CheckBox RO 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox zborder 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Border"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox zreadonly 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "ReadOnly"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox zalt 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Alt"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox ztitle 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Title"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox zfgcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "FGColor"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox fgcol 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   28
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox ztype 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Type"
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox zitem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Item"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox bgcol 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox val 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox nam 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.TextBox zbgcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "BGColor"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox zvalue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Value"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox zname 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Name"
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Wtype 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox Witem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   7560
      Width           =   1935
   End
End
Attribute VB_Name = "pro2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alt_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "alt", alt.Text)
Call ApplyVal
End Sub

Private Sub bgcol_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "bgcolor", bgcol.Text)
Call ApplyVal
End Sub

Private Sub border_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "border", border.Text)
Call ApplyVal
End Sub

Private Sub Cheight_Change()
On Error Resume Next
Select Case Wtype.Text
Case "cb"
Form1.cb(Witem.Text).Height = Cheight.Text
Case "cc"
'non modifiable...
Case "ct"
Form1.ct(Witem.Text).Height = Cheight.Text
Case "ci"
Form1.ci(Witem.Text).Height = Cheight.Text
Case "cp"
Form1.cp(Witem.Text).Height = Cheight.Text
Case "cta"
Form1.cta(Witem.Text).Height = Cheight.Text
Case "ch"
Form1.ch(Witem.Text).Height = Cheight.Text
Case "clist"
Form1.clist(Witem.Text).Height = Cheight.Text
Case "ccombo"
Form1.ccombo(Witem.Text).Height = Cheight.Text
Case "cli"
Form1.cli(Witem.Text).Height = Cheight.Text
Case "Page"
'Form1.zone.height = Cheight.Text
End Select
End Sub

Private Sub Cleft_Change()
On Error Resume Next
Select Case Wtype.Text
Case "cb"
Form1.cb(Witem.Text).Left = Cleft.Text
Case "cc"
'non modifiable...
Case "ct"
Form1.ct(Witem.Text).Left = Cleft.Text
Case "ci"
Form1.ci(Witem.Text).Left = Cleft.Text
Case "cp"
Form1.cp(Witem.Text).Left = Cleft.Text
Case "cta"
Form1.cta(Witem.Text).Left = Cleft.Text
Case "ch"
Form1.ch(Witem.Text).Left = Cleft.Text
Case "clist"
Form1.clist(Witem.Text).Left = Cleft.Text
Case "ccombo"
Form1.ccombo(Witem.Text).Left = Cleft.Text
Case "cli"
Form1.cli(Witem.Text).Left = Cleft.Text
Case "Page"
'Form1.zone.Left = Cleft.Text
End Select
End Sub

Private Sub Ctop_Change()
On Error Resume Next
Select Case Wtype.Text
Case "cb"
Form1.cb(Witem.Text).Top = Ctop.Text
Case "cc"
'non modifiable...
Case "ct"
Form1.ct(Witem.Text).Top = Ctop.Text
Case "ci"
Form1.ci(Witem.Text).Top = Ctop.Text
Case "cp"
Form1.cp(Witem.Text).Top = Ctop.Text
Case "cta"
Form1.cta(Witem.Text).Top = Ctop.Text
Case "ch"
Form1.ch(Witem.Text).Top = Ctop.Text
Case "clist"
Form1.clist(Witem.Text).Top = Ctop.Text
Case "ccombo"
Form1.ccombo(Witem.Text).Top = Ctop.Text
Case "cli"
Form1.cli(Witem.Text).Top = Ctop.Text
Case "Page"
'Form1.zone.Top = Ctop.Text
End Select
End Sub

Private Sub Cwidth_Change()
On Error Resume Next
Select Case Wtype.Text
Case "cb"
Form1.cb(Witem.Text).Width = Cwidth.Text
Case "cc"
'non modifiable...
Case "ct"
Form1.ct(Witem.Text).Width = Cwidth.Text
Case "ci"
Form1.ci(Witem.Text).Width = Cwidth.Text
Case "cp"
Form1.cp(Witem.Text).Width = Cwidth.Text
Case "cta"
Form1.cta(Witem.Text).Width = Cwidth.Text
Case "ch"
Form1.ch(Witem.Text).Width = Cwidth.Text
Case "clist"
Form1.clist(Witem.Text).Width = Cwidth.Text
Case "ccombo"
Form1.ccombo(Witem.Text).Width = Cwidth.Text
Case "cli"
Form1.cli(Witem.Text).Width = Cwidth.Text
Case "Page"
'Form1.zone.width = Cwidth.Text
End Select
End Sub

Private Sub down_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "down", down.Text)
Call ApplyVal
End Sub

Private Sub fgcol_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "fgcolor", fgcol.Text)
Call ApplyVal
End Sub

Private Sub Form_Load()
forward Me
End Sub

Private Sub mouve_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "move", mouve.Text)
Call ApplyVal
End Sub

Private Sub nam_Change()
On Error Resume Next
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "name", nam.Text)
Call ApplyVal
End Sub

Private Sub onabord_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "abord", onabord.Text)
Call ApplyVal
End Sub

Private Sub onblur_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "blur", onblur.Text)
Call ApplyVal
End Sub

Private Sub onchange_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "change", onchange.Text)
Call ApplyVal
End Sub

Private Sub onclick_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "click", onclick.Text)
Call ApplyVal
End Sub

Private Sub ondblclick_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "dblclick", ondblclick.Text)
Call ApplyVal
End Sub

Private Sub onerror_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "error", onerror.Text)
Call ApplyVal
End Sub

Private Sub onfocus_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "focus", onfocus.Text)
Call ApplyVal
End Sub

Private Sub onkeydown_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "keydown", onkeydown.Text)
Call ApplyVal
End Sub

Private Sub onkeypress_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "keypress", onkeypress.Text)
Call ApplyVal
End Sub

Private Sub onkeyup_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "keyup", onkeyup.Text)
Call ApplyVal
End Sub

Private Sub onload_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "load", onload.Text)
Call ApplyVal
End Sub

Private Sub onselect_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "select", onselect.Text)
Call ApplyVal
End Sub

Private Sub onunload_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "unload", onunload.Text)
Call ApplyVal
End Sub

Private Sub out_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "out", out.Text)
Call ApplyVal
End Sub

Private Sub over_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype.Text & Witem.Text, "over", over.Text)
Call ApplyVal
End Sub

Private Sub path_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "path", path.Text)
Call ApplyVal
End Sub

Private Sub RO_Click()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "ro", RO.Value)
Call ApplyVal
End Sub

Private Sub title_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "title", title.Text)
Call ApplyVal
End Sub

Private Sub valu_Change()
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "value", valu.Text)
Call ApplyVal
End Sub

Public Sub ApplyVal()
Dim fileimg
Select Case Wtype
Case "cb"
Form1.cb(Witem.Text).Caption = valu.Text
Form1.cb(Witem.Text).ToolTipText = nam.Text
Case "cc"
Form1.cc(Witem.Text).ToolTipText = nam.Text
Case "ci"
Form1.ci(Witem.Text).ToolTipText = alt.Text & " - (" & nam.Text & ")"
If ExistFile(Form1.cd1.FileName & "\" & path.Text) = True Then
fileimg = path.Text
Form1.ci(Witem.Text).Picture = fileimg
End If
Case "cli"
Form1.cli(Witem.Text).Caption = valu.Text
Form1.cli(Witem.Text).ToolTipText = "URL: " & path.Text
Case "clist"
Form1.clist(Witem.Text).ToolTipText = nam.Text
Form1.clist(Witem.Text).Text = valu.Text
Case "ccombo"
Form1.ccombo(Witem.Text).ToolTipText = nam.Text
Form1.ccombo(Witem.Text).Text = valu.Text
Case "ct"
Form1.ct(Witem.Text).ToolTipText = nam.Text
Form1.ct(Witem.Text).Text = valu.Text
Case "cta"
Form1.cta(Witem.Text).ToolTipText = nam.Text
Form1.cta(Witem.Text).Text = valu.Text
Case "cp"
Form1.cp(Witem.Text).ToolTipText = nam.Text
Form1.cp(Witem.Text).Text = valu.Text
Case "ch"
Form1.ct(Witem.Text).ToolTipText = nam.Text
Form1.ct(Witem.Text).Text = valu.Text
Case "cl"

Case "csub"

End Select
End Sub
