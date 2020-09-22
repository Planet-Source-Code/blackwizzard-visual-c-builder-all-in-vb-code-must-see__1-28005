VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Builder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C++ Project Builder"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "Dev_cpp.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "VC++ Project Builder"
      Height          =   2295
      Left            =   0
      TabIndex        =   20
      Top             =   2520
      Width           =   5175
      Begin VB.TextBox m_name 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   30
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox m_output 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   29
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox m_icon 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   28
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox m_compiler 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   27
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox m_cpp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   26
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CommandButton Command6 
         Caption         =   "..."
         Height          =   195
         Left            =   4800
         TabIndex        =   25
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   195
         Left            =   4800
         TabIndex        =   24
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox m_folder 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   23
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox m_exe 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   22
         Text            =   "C:\Program Files\Microsoft Visual Studio\Common\MSDev98\Bin\MSDEV.EXE"
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   195
         Left            =   4800
         TabIndex        =   21
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Output Filename:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project icon:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compiler Options:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cpp Filename:"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generate VC++ Project"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         MouseIcon       =   "Dev_cpp.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project folder:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generate and launch VC++ Project"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2280
         MouseIcon       =   "Dev_cpp.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   2040
         Width           =   2475
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compiler Folder:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dev-Cpp Project Builder"
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   195
         Left            =   4800
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox d_exe 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   17
         Text            =   "C:\Dev-C++\DevCpp.exe"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox d_folder 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   14
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   195
         Left            =   4800
         TabIndex        =   13
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   195
         Left            =   4800
         TabIndex        =   11
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox d_cpp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   10
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox d_compiler 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   8
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox d_icon 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   6
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox d_output 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   4
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox d_name 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compiler Folder:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generate and launch Dev-cpp Project"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2280
         MouseIcon       =   "Dev_cpp.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   2040
         Width           =   2700
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project folder:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generate Dev-cpp Project"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         MouseIcon       =   "Dev_cpp.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   2040
         Width           =   1860
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cpp Filename:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compiler Options:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project icon:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Output Filename:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1005
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Builder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cd.Filter = "Icon File (*.ico)|*.ico|All Files|*.*"
cd.DialogTitle = "Choose a icon"
cd.ShowOpen
If cd.FileName <> "" Then
d_icon.Text = cd.FileName
End If
End Sub

Private Sub Command2_Click()
Folder.Show
End Sub

Private Sub Command3_Click()
cd.DialogTitle = "choose a compiler/editor"
cd.ShowOpen
If cd.FileName <> "" Then
d_exe = cd.FileName
End If
End Sub

Private Sub Command4_Click()
cd.DialogTitle = "choose a compiler/editor"
cd.ShowOpen
If cd.FileName <> "" Then
m_exe.Text = cd.FileName
End If
End Sub

Private Sub Command5_Click()
Folder.Show
End Sub

Private Sub Command6_Click()
cd.Filter = "Icon File (*.ico)|*.ico|All Files|*.*"
cd.DialogTitle = "Choose a icon"
cd.ShowOpen
If cd.FileName <> "" Then
m_icon.Text = cd.FileName
End If
End Sub

Private Sub Form_Load()
forward Me
Call Generate
d_name.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "title")
d_output.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "title") & ".dev"
d_cpp.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "title") & ".cpp"
m_name.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "title")
m_output.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "title") & ".dsp"
m_cpp.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "title") & ".cpp"
End Sub

Public Sub generate_dev(launch As Boolean)
'generate Project File (*.dev)
Open d_folder.Text & "\" & d_output.Text For Output As #1
Print #1, "[Project]"
Print #1, "FileName=" & d_folder.Text & "\" & d_output.Text
Print #1, "Name=" & d_name.Text
Print #1, "Use_gpp=1"
Print #1, "UnitCount=1"
Print #1, "ResFiles=" & d_folder.Text & "\rsrc.rc"
Print #1, "NoConsole=1"
Print #1, "IsDll=0"
Print #1, "Icon=" & d_icon.Text
Print #1, "CompilerOptions="
Print #1, "IncludeDirs="
Print #1, "ObjFile="
Print #1, ""
Print #1, "[Unit1]"
Print #1, "FileName=" & d_folder.Text & "\" & d_cpp.Text
Print #1, "FileTime=726441358"
Close #1

'generate Resource File (*.rc)
Open d_folder.Text & "\rsrc.rc" For Output As #1
Print #1, "500 ICON MOVEABLE PURE LOADONCALL DISCARDABLE " & Chr(34) & "C:/DEV-C++/Icon/Window.ico" & Chr(34)
Close #1

'generate c++ File (*.cpp)
Open d_folder.Text & "\" & d_cpp.Text For Output As #1
Print #1, Form1.RTCode.Text
Close #1

MsgBox "Dev-Cpp Project File was created successfully on " & d_folder.Text
If launch = True Then
If d_exe.Text <> "" Then
Shell d_exe.Text & " " & d_folder.Text & "\" & d_output.Text
Unload Me
Else
MsgBox "please select a compiler and try again!"
d_exe.ForeColor = RGB(255, 0, 0)
d_exe.Text = "select a compiler!"
End If
Else
Unload Me
End If
End Sub

Private Sub Label11_Click()
generate_vcpp True
End Sub

Private Sub Label13_Click()
generate_vcpp False
End Sub

Private Sub Label6_Click()
generate_dev False
End Sub

Private Sub Label8_Click()
generate_dev True
End Sub

Public Sub generate_vcpp(launch As Boolean)
'generate c++ File (*.cpp)
Open m_folder.Text & "\" & m_cpp.Text For Output As #1
Print #1, Form1.RTCode.Text
Close #1

'generate Project File (*.dsp)
Open m_folder.Text & "\" & m_output.Text For Output As #1
Print #1, "# Microsoft Developer Studio Project File - Name=" & Chr(34) & m_name.Text & Chr(34) & " - Package Owner=<4>"
Print #1, "# Microsoft Developer Studio Generated Build File, Format Version 6.00"
Print #1, "# ** DO NOT EDIT **"
Print #1, ""
Print #1, "# TARGTYPE " & Chr(34) & "Win32 (x86) Application" & Chr(34) & " 0x0101"
Print #1, ""
Print #1, "CFG=" & m_name.Text & " - Win32 Debug"
Print #1, "!MESSAGE This is not a valid makefile. To build this project using NMAKE,"
Print #1, "!MESSAGE use the Export Makefile command and run"
Print #1, "!MESSAGE"
Print #1, "!MESSAGE NMAKE /f " & Chr(34) & m_name.Text & ".mak" & Chr(34) & "."
Print #1, "!MESSAGE"
Print #1, "!MESSAGE You can specify a configuration when running NMAKE"
Print #1, "!MESSAGE by defining the macro CFG on the command line. For example:"
Print #1, "!MESSAGE"
Print #1, "!MESSAGE NMAKE /f " & Chr(34) & m_name.Text & ".mak" & Chr(34) & " CFG=" & Chr(34) & m_name.Text & " - Win32 Debug"
Print #1, "!MESSAGE"
Print #1, "!MESSAGE Possible choices for configuration are:"
Print #1, "!MESSAGE"
Print #1, "!MESSAGE " & Chr(34) & m_name.Text & " - Win32 Release" & Chr(34) & " (based on " & Chr(34) & "Win32 (x86) Application" & Chr(34) & ")"
Print #1, "!MESSAGE " & Chr(34) & m_name.Text & " - Win32 Debug" & Chr(34) & " (based on " & Chr(34) & "Win32 (x86) Application" & Chr(34) & ")"
Print #1, "!MESSAGE"
Print #1, ""
Print #1, "# Begin Project"
Print #1, "# PROP AllowPerConfigDependencies 0"
Print #1, "# PROP Scc_ProjName " & Chr(34) & "" & Chr(34)
Print #1, "# PROP Scc_LocalPath " & Chr(34) & "" & Chr(34)
Print #1, "CPP=cl.exe"
Print #1, "MTL=midl.exe"
Print #1, "RSC=rc.exe"
Print #1, ""
Print #1, "!IF  " & Chr(34) & "$(CFG)" & Chr(34) & " == " & Chr(34) & m_name.Text & " - Win32 Release" & Chr(34)
Print #1, ""
Print #1, "# PROP BASE Use_MFC 0"
Print #1, "# PROP BASE Use_Debug_Libraries 0"
Print #1, "# PROP BASE Output_Dir " & Chr(34) & "Release" & Chr(34)
Print #1, "# PROP BASE Intermediate_Dir " & Chr(34) & "Release" & Chr(34) & ""
Print #1, "# PROP BASE Target_Dir " & Chr(34) & Chr(34)
Print #1, "# PROP Use_MFC 0"
Print #1, "# PROP Use_Debug_Libraries 0"
Print #1, "# PROP Output_Dir " & Chr(34) & "Release" & Chr(34)
Print #1, "# PROP Intermediate_Dir " & Chr(34) & "Release" & Chr(34)
Print #1, "# PROP Target_Dir " & Chr(34) & Chr(34)
Print #1, "# ADD BASE CPP /nologo /W3 /GX /O2 /D " & Chr(34) & "WIN32" & Chr(34) & " /D " & Chr(34) & "NDEBUG" & Chr(34) & " /D " & Chr(34) & "_WINDOWS" & Chr(34) & " /D " & Chr(34) & "_MBCS" & Chr(34) & " /YX /FD /c"
Print #1, "# ADD CPP /nologo /W3 /GX /O2 /D " & Chr(34) & "WIN32" & Chr(34) & " /D " & Chr(34) & "NDEBUG" & Chr(34) & " /D " & Chr(34) & "_WINDOWS" & Chr(34) & " /D " & Chr(34) & "_MBCS" & Chr(34) & " /YX /FD /c"
Print #1, "# ADD BASE MTL /nologo /D " & Chr(34) & "NDEBUG" & Chr(34) & " /mktyplib203 /win32"
Print #1, "# ADD MTL /nologo /D " & Chr(34) & "NDEBUG" & Chr(34) & " /mktyplib203 /win32"
Print #1, "# ADD BASE RSC /l 0x40c /d " & Chr(34) & "NDEBUG" & Chr(34) & ""
Print #1, "# ADD RSC /l 0x40c /d " & Chr(34) & "NDEBUG" & Chr(34) & ""
Print #1, "BSC32=bscmake.exe"
Print #1, "# ADD BASE BSC32 /nologo"
Print #1, "# ADD BSC32 /nologo"
Print #1, "LINK32=link.exe"
Print #1, "# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /machine:I386"
Print #1, "# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /machine:I386"
Print #1, ""
Print #1, "!ELSEIF  " & Chr(34) & "$(CFG)" & Chr(34) & " == " & Chr(34) & m_name.Text & " - Win32 Debug" & Chr(34) & ""
Print #1, ""
Print #1, "# PROP BASE Use_MFC 0"
Print #1, "# PROP BASE Use_Debug_Libraries 1"
Print #1, "# PROP BASE Output_Dir " & Chr(34) & "Debug" & Chr(34) & ""
Print #1, "# PROP BASE Intermediate_Dir " & Chr(34) & "Debug" & Chr(34) & ""
Print #1, "# PROP BASE Target_Dir " & Chr(34) & "" & Chr(34) & ""
Print #1, "# PROP Use_MFC 0"
Print #1, "# PROP Use_Debug_Libraries 1"
Print #1, "# PROP Output_Dir " & Chr(34) & "Debug" & Chr(34) & ""
Print #1, "# PROP Intermediate_Dir " & Chr(34) & "Debug" & Chr(34) & ""
Print #1, "# PROP Target_Dir " & Chr(34) & "" & Chr(34) & ""
Print #1, "# ADD BASE CPP /nologo /W3 /Gm /GX /ZI /Od /D " & Chr(34) & "WIN32" & Chr(34) & " /D " & Chr(34) & "_DEBUG" & Chr(34) & " /D " & Chr(34) & "_WINDOWS" & Chr(34) & " /D " & Chr(34) & "_MBCS" & Chr(34) & " /YX /FD /GZ /c"
Print #1, "# ADD CPP /nologo /W3 /Gm /GX /ZI /Od /D " & Chr(34) & "WIN32" & Chr(34) & " /D " & Chr(34) & "_DEBUG" & Chr(34) & " /D " & Chr(34) & "_WINDOWS" & Chr(34) & " /D " & Chr(34) & "_MBCS" & Chr(34) & " /YX /FD /GZ /c"
Print #1, "# ADD BASE MTL /nologo /D " & Chr(34) & "_DEBUG" & Chr(34) & " /mktyplib203 /win32"
Print #1, "# ADD MTL /nologo /D " & Chr(34) & "_DEBUG" & Chr(34) & " /mktyplib203 /win32"
Print #1, "# ADD BASE RSC /l 0x40c /d " & Chr(34) & "_DEBUG" & Chr(34) & ""
Print #1, "# ADD RSC /l 0x40c /d " & Chr(34) & "_DEBUG" & Chr(34) & ""
Print #1, "BSC32=bscmake.exe"
Print #1, "# ADD BASE BSC32 /nologo"
Print #1, "# ADD BSC32 /nologo"
Print #1, "LINK32=link.exe"
Print #1, "# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /debug /machine:I386 /pdbtype:sept"
Print #1, "# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /debug /machine:I386 /pdbtype:sept"
Print #1, ""
Print #1, "!ENDIF"
Print #1, ""
Print #1, "# Begin Target"
Print #1, ""
Print #1, "# Name " & Chr(34) & m_name.Text & "- Win32 Release" & Chr(34) & ""
Print #1, "# Name " & Chr(34) & m_name.Text & " - Win32 Debug" & Chr(34) & ""
Print #1, "# Begin Group " & Chr(34) & "Source Files" & Chr(34) & ""
Print #1, ""
Print #1, "# PROP Default_Filter " & Chr(34) & "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat" & Chr(34) & ""
Print #1, "# Begin Source File"
Print #1, ""
Print #1, "SOURCE=.\" & m_cpp.Text
Print #1, "# End Source File"
Print #1, "# End Group"
Print #1, "# Begin Group " & Chr(34) & "Header Files" & Chr(34) & ""
Print #1, ""
Print #1, "# PROP Default_Filter " & Chr(34) & "h;hpp;hxx;hm;inl" & Chr(34) & ""
Print #1, "# End Group"
Print #1, "# Begin Group " & Chr(34) & "Resource Files" & Chr(34) & ""
Print #1, ""
Print #1, "# PROP Default_Filter " & Chr(34) & "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe" & Chr(34) & ""
Print #1, "# End Group"
Print #1, "# End Target"
Print #1, "# End Project"
Close #1

MsgBox "VC++ Project File was created successfully on " & m_folder.Text
If launch = True Then
If m_exe.Text <> "" Then
Shell m_exe.Text & " " & m_folder.Text & "\" & m_output.Text
Unload Me
Else
MsgBox "please select a compiler and try again!"
m_exe.ForeColor = RGB(255, 0, 0)
m_exe.Text = "select a compiler!"
End If
Else
Unload Me
End If
End Sub
