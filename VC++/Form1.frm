VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visual CPP editor"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13335
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox BWC0 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   62
      Text            =   "Form1.frx":1D2A
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox BWC3 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   58
      Text            =   "Form1.frx":2348
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox BWC2 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   57
      Text            =   "Form1.frx":291D
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox BWC1 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   56
      Text            =   "Form1.frx":2A89
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   855
      Left            =   120
      TabIndex        =   55
      Top             =   2880
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":2BEB
   End
   Begin VB.TextBox Ycoord 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   8760
      Width           =   735
   End
   Begin VB.TextBox Xcoord 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   8520
      Width           =   735
   End
   Begin VB.ListBox listfunc 
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox space 
      Height          =   285
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "Form1.frx":2CC0
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton o 
      Height          =   375
      Index           =   11
      Left            =   480
      Picture         =   "Form1.frx":2CC4
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1920
      Width           =   375
   End
   Begin VB.OptionButton o 
      Enabled         =   0   'False
      Height          =   375
      Index           =   10
      Left            =   120
      Picture         =   "Form1.frx":3066
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1920
      Width           =   375
   End
   Begin VB.OptionButton o 
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   480
      Picture         =   "Form1.frx":3468
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1560
      Width           =   375
   End
   Begin VB.OptionButton o 
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   120
      Picture         =   "Form1.frx":36BA
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1560
      Width           =   375
   End
   Begin VB.OptionButton o 
      Height          =   375
      Index           =   7
      Left            =   480
      Picture         =   "Form1.frx":3ABC
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1200
      Width           =   375
   End
   Begin VB.OptionButton o 
      Height          =   375
      Index           =   6
      Left            =   120
      Picture         =   "Form1.frx":3F72
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1200
      Width           =   375
   End
   Begin VB.OptionButton o 
      Height          =   375
      Index           =   5
      Left            =   480
      Picture         =   "Form1.frx":42FC
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton o 
      Height          =   375
      Index           =   4
      Left            =   120
      Picture         =   "Form1.frx":47B2
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton o 
      Caption         =   "text"
      Height          =   375
      Index           =   3
      Left            =   480
      Picture         =   "Form1.frx":4BB4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton o 
      Height          =   375
      Index           =   2
      Left            =   120
      Picture         =   "Form1.frx":4F7A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton o 
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   480
      Picture         =   "Form1.frx":5430
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton o 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":586E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton b 
      Height          =   375
      Index           =   11
      Left            =   480
      Picture         =   "Form1.frx":5BF8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton b 
      Height          =   375
      Index           =   10
      Left            =   120
      Picture         =   "Form1.frx":5F9A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton b 
      Height          =   375
      Index           =   9
      Left            =   480
      Picture         =   "Form1.frx":639C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton b 
      Height          =   375
      Index           =   8
      Left            =   120
      Picture         =   "Form1.frx":67DA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton b 
      Height          =   375
      Index           =   7
      Left            =   480
      Picture         =   "Form1.frx":6BDC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton b 
      Height          =   375
      Index           =   6
      Left            =   120
      Picture         =   "Form1.frx":7092
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton b 
      Height          =   375
      Index           =   5
      Left            =   480
      Picture         =   "Form1.frx":741C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton b 
      Height          =   375
      Index           =   4
      Left            =   120
      Picture         =   "Form1.frx":78D2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton b 
      Height          =   375
      Index           =   3
      Left            =   480
      Picture         =   "Form1.frx":7CD4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton b 
      Height          =   375
      Index           =   2
      Left            =   120
      Picture         =   "Form1.frx":809A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton b 
      Height          =   375
      Index           =   1
      Left            =   480
      Picture         =   "Form1.frx":8550
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton b 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":898E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   360
      Top             =   9000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9615
      Left            =   960
      TabIndex        =   28
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   16960
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   529
      BackColor       =   12632256
      TabCaption(0)   =   "Visual Dev"
      TabPicture(0)   =   "Form1.frx":8D18
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "zone"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Source Code"
      TabPicture(1)   =   "Form1.frx":8D34
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "xcv"
      Tab(1).Control(1)=   "RTCode"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Functions"
      TabPicture(2)   =   "Form1.frx":8D50
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(2)=   "Label5"
      Tab(2).Control(3)=   "js"
      Tab(2).Control(4)=   "fname"
      Tab(2).Control(5)=   "addfunc"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Preview"
      TabPicture(3)   =   "Form1.frx":8D6C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label4"
      Tab(3).Control(1)=   "wb"
      Tab(3).Control(2)=   "pscode"
      Tab(3).ControlCount=   3
      Begin SHDocVwCtl.WebBrowser pscode 
         Height          =   9015
         Left            =   -74880
         TabIndex        =   63
         Top             =   480
         Width           =   12135
         ExtentX         =   21405
         ExtentY         =   15901
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VHTML.htmSyntaxBox RTCode 
         Height          =   8535
         Left            =   -74880
         TabIndex        =   54
         Top             =   960
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   15055
         AutoVerbMenu    =   -1  'True
         BackColor       =   16777215
         CommentColor    =   49152
         EntityColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form1.frx":8D88
         PropNameColor   =   16711680
         PropValColor    =   33023
         TagColor        =   16744703
         Text            =   ""
      End
      Begin VB.CommandButton addfunc 
         Caption         =   "ok"
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70920
         TabIndex        =   48
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox fname 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72720
         TabIndex        =   47
         Text            =   "init"
         Top             =   480
         Width           =   1815
      End
      Begin VB.PictureBox zone 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   9135
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   607
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   807
         TabIndex        =   29
         Top             =   360
         Width           =   12135
         Begin VB.PictureBox rect 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   105
            Index           =   7
            Left            =   5760
            MousePointer    =   8  'Size NW SE
            Picture         =   "Form1.frx":8DA4
            ScaleHeight     =   75
            ScaleWidth      =   75
            TabIndex        =   46
            Top             =   4680
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.PictureBox rect 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   105
            Index           =   6
            Left            =   5400
            MousePointer    =   7  'Size N S
            Picture         =   "Form1.frx":8E36
            ScaleHeight     =   75
            ScaleWidth      =   75
            TabIndex        =   45
            Top             =   4680
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.PictureBox rect 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   105
            Index           =   5
            Left            =   5040
            MousePointer    =   6  'Size NE SW
            Picture         =   "Form1.frx":8EC8
            ScaleHeight     =   75
            ScaleWidth      =   75
            TabIndex        =   44
            Top             =   4680
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.PictureBox rect 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   105
            Index           =   4
            Left            =   5760
            MousePointer    =   9  'Size W E
            Picture         =   "Form1.frx":8F5A
            ScaleHeight     =   75
            ScaleWidth      =   75
            TabIndex        =   43
            Top             =   4320
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.PictureBox rect 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   105
            Index           =   3
            Left            =   5040
            MousePointer    =   9  'Size W E
            Picture         =   "Form1.frx":8FEC
            ScaleHeight     =   75
            ScaleWidth      =   75
            TabIndex        =   42
            Top             =   4320
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.PictureBox rect 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   105
            Index           =   2
            Left            =   5760
            MousePointer    =   6  'Size NE SW
            Picture         =   "Form1.frx":907E
            ScaleHeight     =   75
            ScaleWidth      =   75
            TabIndex        =   41
            Top             =   3840
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.PictureBox rect 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   105
            Index           =   1
            Left            =   5400
            MousePointer    =   7  'Size N S
            Picture         =   "Form1.frx":9110
            ScaleHeight     =   75
            ScaleWidth      =   75
            TabIndex        =   40
            Top             =   3840
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.PictureBox rect 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   105
            Index           =   0
            Left            =   5040
            MousePointer    =   8  'Size NW SE
            Picture         =   "Form1.frx":91A2
            ScaleHeight     =   75
            ScaleWidth      =   75
            TabIndex        =   39
            Top             =   3840
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.CommandButton csub 
            Caption         =   "Submit"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   38
            ToolTipText     =   "unidentified Submit Button"
            Top             =   840
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox ch 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   0
            Left            =   1440
            TabIndex        =   37
            Text            =   "hidden"
            ToolTipText     =   "unidentified Hidden Field"
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox cta 
            Height          =   615
            Index           =   0
            Left            =   3480
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Text            =   "Form1.frx":9234
            ToolTipText     =   "unidentified Textarea"
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox cp 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1440
            PasswordChar    =   "*"
            TabIndex        =   35
            Text            =   "password"
            ToolTipText     =   "unidentified Password Field"
            Top             =   600
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox ccombo 
            Height          =   315
            Index           =   0
            Left            =   2880
            TabIndex        =   34
            ToolTipText     =   "unidentified ComboBox (<select>)"
            Top             =   240
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.ListBox clist 
            Height          =   255
            Index           =   0
            ItemData        =   "Form1.frx":923F
            Left            =   2520
            List            =   "Form1.frx":9246
            TabIndex        =   33
            ToolTipText     =   "unidentified ListBox(<select>)"
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox cc 
            Height          =   195
            Index           =   0
            Left            =   2040
            TabIndex        =   32
            ToolTipText     =   "unidentified CheckBox"
            Top             =   240
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox ct 
            Height          =   285
            Index           =   0
            Left            =   1440
            TabIndex        =   31
            ToolTipText     =   "unidentified TextBox"
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cb 
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   30
            Tag             =   "uèçuèi-iuy"
            ToolTipText     =   "unidentified button"
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label cli 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New Label"
            Height          =   195
            Index           =   0
            Left            =   1920
            TabIndex        =   59
            Top             =   960
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Image ci 
            DataSource      =   "dhfkdhf"
            Height          =   240
            Index           =   0
            Left            =   960
            Picture         =   "Form1.frx":9251
            Stretch         =   -1  'True
            Tag             =   "c:\image\cool.jpg"
            ToolTipText     =   "unidentified image"
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Line cl 
            Index           =   0
            Visible         =   0   'False
            X1              =   40
            X2              =   72
            Y1              =   48
            Y2              =   48
         End
         Begin VB.Line Line1 
            BorderStyle     =   3  'Dot
            Index           =   0
            Visible         =   0   'False
            X1              =   24
            X2              =   24
            Y1              =   0
            Y2              =   56
         End
         Begin VB.Line Line1 
            BorderStyle     =   3  'Dot
            Index           =   1
            Visible         =   0   'False
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   56
         End
         Begin VB.Line Line1 
            BorderStyle     =   3  'Dot
            Index           =   2
            Visible         =   0   'False
            X1              =   8
            X2              =   8
            Y1              =   0
            Y2              =   56
         End
         Begin VB.Line Line1 
            BorderStyle     =   3  'Dot
            Index           =   3
            Visible         =   0   'False
            X1              =   16
            X2              =   16
            Y1              =   0
            Y2              =   56
         End
      End
      Begin RichTextLib.RichTextBox js 
         Height          =   135
         Left            =   -62880
         TabIndex        =   49
         Top             =   9360
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   238
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"Form1.frx":26753
      End
      Begin RichTextLib.RichTextBox xcv 
         Height          =   3135
         Left            =   -66960
         TabIndex        =   50
         Top             =   6375
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5530
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form1.frx":26828
      End
      Begin SHDocVwCtl.WebBrowser wb 
         Height          =   120
         Left            =   -62880
         TabIndex        =   51
         Top             =   9375
         Visible         =   0   'False
         Width           =   135
         ExtentX         =   238
         ExtentY         =   212
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.Label Label5 
         Caption         =   "Sorry but this panel is for Visual HTML, not for Visual CPP Editor!"
         Height          =   375
         Left            =   -71280
         TabIndex        =   61
         Top             =   4680
         Width           =   4695
      End
      Begin VB.Label Label4 
         Caption         =   "Sorry but this panel is for Visual HTML, not for Visual CPP Editor!"
         Height          =   375
         Left            =   -71280
         TabIndex        =   60
         Top             =   4680
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   $"Form1.frx":26933
         Enabled         =   0   'False
         Height          =   435
         Left            =   -69840
         TabIndex        =   53
         Top             =   480
         Width           =   6585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ajouter une nouvelle fonction:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   -74880
         TabIndex        =   52
         Top             =   540
         Width           =   2100
      End
   End
   Begin VB.Menu file 
      Caption         =   "Fichier"
      Begin VB.Menu f 
         Caption         =   "New"
         Index           =   0
      End
      Begin VB.Menu f 
         Caption         =   "Save as"
         Index           =   1
      End
      Begin VB.Menu f 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu f 
         Caption         =   "ShellExecute Code"
         Index           =   3
      End
      Begin VB.Menu f 
         Caption         =   "Quit"
         Index           =   4
      End
   End
   Begin VB.Menu pjc 
      Caption         =   "Project"
      Begin VB.Menu msP 
         Caption         =   "BlackWizzard Project Builder"
      End
   End
   Begin VB.Menu toolz 
      Caption         =   "Tools"
      Enabled         =   0   'False
      Begin VB.Menu dial 
         Caption         =   "Dialogs"
         Begin VB.Menu dialog 
            Caption         =   "Alert [X]"
            Index           =   0
         End
         Begin VB.Menu dialog 
            Caption         =   "Confirm [?]"
            Index           =   1
         End
         Begin VB.Menu dialog 
            Caption         =   "Prompt [input]"
            Index           =   2
         End
      End
      Begin VB.Menu loopz 
         Caption         =   "loops, op, ..."
         Begin VB.Menu loo 
            Caption         =   "if"
            Index           =   0
         End
         Begin VB.Menu loo 
            Caption         =   "if...else"
            Index           =   1
         End
         Begin VB.Menu loo 
            Caption         =   "if...else if..else"
            Index           =   2
         End
         Begin VB.Menu loo 
            Caption         =   "with"
            Index           =   3
         End
         Begin VB.Menu loo 
            Caption         =   "switch"
            Index           =   4
         End
         Begin VB.Menu loo 
            Caption         =   "while"
            Index           =   5
         End
         Begin VB.Menu loo 
            Caption         =   "for"
            Index           =   6
         End
         Begin VB.Menu loo 
            Caption         =   "break"
            Index           =   7
         End
      End
   End
   Begin VB.Menu other 
      Caption         =   "Other"
      Begin VB.Menu about 
         Caption         =   "About..."
      End
      Begin VB.Menu download 
         Caption         =   "Download Lattest version"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ccheck As Integer
Dim cbut As Integer
Dim cimg As Integer
Dim clink As Integer
Dim clistB As Integer
Dim ccomboB As Integer
Dim ctext As Integer
Dim cpassw As Integer
Dim ctexta As Integer
Dim chiddenf As Integer
Dim cline As Integer
Dim csubmitbut As Integer
Dim X1 As Long
Dim Y1 As Long
Dim X3 As Long
Dim Y3 As Long
Dim X4 As Long
Dim Y4 As Long
Dim bol As Boolean
Dim I
Dim aa As Boolean
Dim bb As Boolean
Dim c As Boolean
Dim dd As Boolean
Dim ee As Boolean
Dim ff As Boolean
Dim gg As Boolean
Dim hh As Boolean


Private Sub about_Click()
AboutVHTML.Show
End Sub

Private Sub addfunc_Click()
Call Verif_Func_Name
End Sub

Private Sub b_Click(Index As Integer)
o(Index).Value = True
If Index <> 11 Then
Me.zone.MousePointer = 2
ElseIf Index = 11 Then
Me.zone.MousePointer = 0
End If
End Sub

Private Sub cb_Click(Index As Integer)
pro.Wtype.Text = "cb"
pro.Witem = Index
'pro.Caption = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "caption")
pro.nam.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "name")
pro.valu.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "value")
pro.bgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "bgcolor")
pro.title.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "title")
pro.alt.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "alt")
pro.RO.Value = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "ro")
pro.border.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "border")
pro.path.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "path")
pro.onclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "click")
pro.out.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "out")
pro.over.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "over")
pro.down.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "down")
pro.fgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "fgcolor")
pro.ondblclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "dblclick")
pro.mouve.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "move")
pro.onload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "load")
pro.onunload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "unload")
pro.onkeydown.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keydown")
pro.onkeyup.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keyup")
pro.onkeypress.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keypress")
pro.onselect.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "select")
pro.onfocus.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "focus")
pro.onchange.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "change")
pro.onblur.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "blur")
pro.onerror.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "error")
pro.onabord.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "abord")
pro.ondblclick.Enabled = True
pro.mouve.Enabled = True
pro.onload.Enabled = False
pro.onunload.Enabled = False
pro.onkeydown.Enabled = False
pro.onkeyup.Enabled = False
pro.onkeypress.Enabled = False
pro.onselect.Enabled = True
pro.onfocus.Enabled = False
pro.onchange.Enabled = False
pro.onblur.Enabled = False
pro.onerror.Enabled = False
pro.onabord.Enabled = False
pro.zondblclick.Enabled = pro.ondblclick.Enabled
pro.zmouve.Enabled = pro.mouve.Enabled
pro.zonload.Enabled = pro.onload.Enabled
pro.zonunload.Enabled = pro.onunload.Enabled
pro.zonkeydown.Enabled = pro.onkeydown.Enabled
pro.zonkeyup.Enabled = pro.onkeyup.Enabled
pro.zonkeypress.Enabled = pro.onkeypress.Enabled
pro.zonselect.Enabled = pro.onselect.Enabled
pro.zonfocus.Enabled = pro.onfocus.Enabled
pro.zonchange.Enabled = pro.onchange.Enabled
pro.zonblur.Enabled = pro.onblur.Enabled
pro.zonerror.Enabled = pro.onerror.Enabled
pro.zonabord.Enabled = pro.onabord.Enabled
pro.fgcol.Enabled = False
pro.zfgcolor.Enabled = pro.fgcol.Enabled
pro.nam.Enabled = True
pro.valu.Enabled = True
pro.bgcol.Enabled = False
pro.title.Enabled = False
pro.alt.Enabled = False
pro.RO.Enabled = False
pro.border.Enabled = False
pro.path.Enabled = False
pro.zname.Enabled = pro.nam.Enabled
pro.zvalue.Enabled = pro.valu.Enabled
pro.zbgcolor.Enabled = pro.bgcol.Enabled
pro.ztitle.Enabled = pro.title.Enabled
pro.zalt.Enabled = pro.alt.Enabled
pro.zreadonly.Enabled = pro.RO.Enabled
pro.zborder.Enabled = pro.border.Enabled
pro.zpath.Enabled = pro.path.Enabled
pro.onclick.Enabled = True
pro.over.Enabled = True
pro.out.Enabled = True
pro.down.Enabled = True
pro.zclick.Enabled = pro.onclick.Enabled
pro.zover.Enabled = pro.over.Enabled
pro.zout.Enabled = pro.out.Enabled
pro.zdown.Enabled = pro.down.Enabled

rect(0).Left = cb(Index).Left - 8
rect(0).Top = cb(Index).Top - 8
rect(0).Visible = True
rect(1).Left = cb(Index).Left + cb(Index).Width / 2 - rect(1).Width / 2
rect(1).Top = cb(Index).Top - 8
rect(1).Visible = True
rect(2).Left = cb(Index).Left + cb(Index).Width
rect(2).Top = cb(Index).Top - 8
rect(2).Visible = True
rect(3).Left = cb(Index).Left - 8
rect(3).Top = cb(Index).Top + cb(Index).Height / 2 - rect(3).Height / 2
rect(3).Visible = True
rect(4).Left = cb(Index).Left + cb(Index).Width
rect(4).Top = cb(Index).Top + cb(Index).Height / 2 - rect(4).Height / 2
rect(4).Visible = True
rect(5).Left = cb(Index).Left - 8
rect(5).Top = cb(Index).Top + cb(Index).Height
rect(5).Visible = True
rect(6).Left = cb(Index).Left + cb(Index).Width / 2 - rect(1).Width / 2
rect(6).Top = cb(Index).Top + cb(Index).Height
rect(6).Visible = True
rect(7).Left = cb(Index).Left + cb(Index).Width
rect(7).Top = cb(Index).Top + cb(Index).Height
rect(7).Visible = True
For I = 0 To 3
Line1(I).BorderColor = zone.ForeColor
Line1(I).Visible = True
Next I
Line1(0).X1 = rect(0).Left + rect(0).Width
Line1(0).Y1 = rect(0).Top + rect(0).Height
Line1(0).X2 = rect(2).Left + rect(2).Width
Line1(0).Y2 = rect(2).Top + rect(2).Height
Line1(1).X1 = rect(2).Left
Line1(1).Y1 = rect(2).Top
Line1(1).X2 = rect(7).Left
Line1(1).Y2 = rect(7).Top
Line1(2).X1 = rect(7).Left
Line1(2).Y1 = rect(7).Top
Line1(2).X2 = rect(5).Left + rect(0).Width
Line1(2).Y2 = rect(5).Top
Line1(3).X1 = rect(5).Left + rect(0).Width
Line1(3).Y1 = rect(5).Top
Line1(3).X2 = rect(0).Left + rect(0).Width
Line1(3).Y2 = rect(0).Top + rect(0).Height
pro.Cheight.Text = cb(Index).Height
pro.Cwidth.Text = cb(Index).Width
pro.Ctop.Text = cb(Index).Top
pro.Cleft.Text = cb(Index).Left
End Sub

Private Sub cc_Click(Index As Integer)
pro.Wtype.Text = "cc"
pro.Witem = Index
'pro.Caption = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "caption")
pro.nam.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "name")
pro.valu.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "value")
pro.bgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "bgcolor")
pro.title.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "title")
pro.alt.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "alt")
pro.RO.Value = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "ro")
pro.border.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "border")
pro.path.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "path")
pro.onclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "click")
pro.out.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "out")
pro.over.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "over")
pro.down.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "down")
pro.fgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "fgcolor")
pro.ondblclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "dblclick")
pro.mouve.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "move")
pro.onload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "load")
pro.onunload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "unload")
pro.onkeydown.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keydown")
pro.onkeyup.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keyup")
pro.onkeypress.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keypress")
pro.onselect.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "select")
pro.onfocus.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "focus")
pro.onchange.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "change")
pro.onblur.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "blur")
pro.onerror.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "error")
pro.onabord.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "abord")
pro.ondblclick.Enabled = True
pro.mouve.Enabled = True
pro.onload.Enabled = False
pro.onunload.Enabled = False
pro.onkeydown.Enabled = False
pro.onkeyup.Enabled = False
pro.onkeypress.Enabled = False
pro.onselect.Enabled = True
pro.onfocus.Enabled = False
pro.onchange.Enabled = False
pro.onblur.Enabled = False
pro.onerror.Enabled = False
pro.onabord.Enabled = False
pro.zondblclick.Enabled = pro.ondblclick.Enabled
pro.zmouve.Enabled = pro.mouve.Enabled
pro.zonload.Enabled = pro.onload.Enabled
pro.zonunload.Enabled = pro.onunload.Enabled
pro.zonkeydown.Enabled = pro.onkeydown.Enabled
pro.zonkeyup.Enabled = pro.onkeyup.Enabled
pro.zonkeypress.Enabled = pro.onkeypress.Enabled
pro.zonselect.Enabled = pro.onselect.Enabled
pro.zonfocus.Enabled = pro.onfocus.Enabled
pro.zonchange.Enabled = pro.onchange.Enabled
pro.zonblur.Enabled = pro.onblur.Enabled
pro.zonerror.Enabled = pro.onerror.Enabled
pro.zonabord.Enabled = pro.onabord.Enabled
pro.fgcol.Enabled = False
pro.zfgcolor.Enabled = pro.fgcol.Enabled
pro.nam.Enabled = True
pro.valu.Enabled = True
pro.bgcol.Enabled = False
pro.title.Enabled = False
pro.alt.Enabled = False
pro.RO.Enabled = False
pro.border.Enabled = False
pro.path.Enabled = False
pro.zname.Enabled = pro.nam.Enabled
pro.zvalue.Enabled = pro.valu.Enabled
pro.zbgcolor.Enabled = pro.bgcol.Enabled
pro.ztitle.Enabled = pro.title.Enabled
pro.zalt.Enabled = pro.alt.Enabled
pro.zreadonly.Enabled = pro.RO.Enabled
pro.zborder.Enabled = pro.border.Enabled
pro.zpath.Enabled = pro.path.Enabled
pro.onclick.Enabled = True
pro.over.Enabled = True
pro.out.Enabled = True
pro.down.Enabled = True
pro.zclick.Enabled = pro.onclick.Enabled
pro.zover.Enabled = pro.over.Enabled
pro.zout.Enabled = pro.out.Enabled
pro.zdown.Enabled = pro.down.Enabled

rect(0).Left = cc(Index).Left - 8
rect(0).Top = cc(Index).Top - 8
rect(0).Visible = True
rect(1).Left = cc(Index).Left + cc(Index).Width / 2 - rect(1).Width / 2
rect(1).Top = cc(Index).Top - 8
rect(1).Visible = True
rect(2).Left = cc(Index).Left + cc(Index).Width
rect(2).Top = cc(Index).Top - 8
rect(2).Visible = True
rect(3).Left = cc(Index).Left - 8
rect(3).Top = cc(Index).Top + cc(Index).Height / 2 - rect(3).Height / 2
rect(3).Visible = True
rect(4).Left = cc(Index).Left + cc(Index).Width
rect(4).Top = cc(Index).Top + cc(Index).Height / 2 - rect(4).Height / 2
rect(4).Visible = True
rect(5).Left = cc(Index).Left - 8
rect(5).Top = cc(Index).Top + cc(Index).Height
rect(5).Visible = True
rect(6).Left = cc(Index).Left + cc(Index).Width / 2 - rect(1).Width / 2
rect(6).Top = cc(Index).Top + cc(Index).Height
rect(6).Visible = True
rect(7).Left = cc(Index).Left + cc(Index).Width
rect(7).Top = cc(Index).Top + cc(Index).Height
rect(7).Visible = True
For I = 0 To 3
Line1(I).BorderColor = zone.ForeColor
Line1(I).Visible = True
Next I
Line1(0).X1 = rect(0).Left + rect(0).Width
Line1(0).Y1 = rect(0).Top + rect(0).Height
Line1(0).X2 = rect(2).Left + rect(2).Width
Line1(0).Y2 = rect(2).Top + rect(2).Height
Line1(1).X1 = rect(2).Left
Line1(1).Y1 = rect(2).Top
Line1(1).X2 = rect(7).Left
Line1(1).Y2 = rect(7).Top
Line1(2).X1 = rect(7).Left
Line1(2).Y1 = rect(7).Top
Line1(2).X2 = rect(5).Left + rect(0).Width
Line1(2).Y2 = rect(5).Top
Line1(3).X1 = rect(5).Left + rect(0).Width
Line1(3).Y1 = rect(5).Top
Line1(3).X2 = rect(0).Left + rect(0).Width
Line1(3).Y2 = rect(0).Top + rect(0).Height
pro.Cheight.Text = cc(Index).Height
pro.Cwidth.Text = cc(Index).Width
pro.Ctop.Text = cc(Index).Top
pro.Cleft.Text = cc(Index).Left
End Sub

Private Sub ccombo_Change(Index As Integer)
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "value", ccombo(Index).Text)
pro.valu.Text = ccombo(Index).Text
End Sub

Private Sub ccombo_GotFocus(Index As Integer)
pro.Wtype.Text = "ccombo"
pro.Witem = Index
'pro.Caption = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "caption")
pro.nam.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "name")
pro.valu.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "value")
pro.bgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "bgcolor")
pro.title.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "title")
pro.alt.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "alt")
pro.RO.Value = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "ro")
pro.border.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "border")
pro.path.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "path")
pro.onclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "click")
pro.out.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "out")
pro.over.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "over")
pro.down.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "down")
pro.fgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "fgcolor")
pro.ondblclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "dblclick")
pro.mouve.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "move")
pro.onload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "load")
pro.onunload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "unload")
pro.onkeydown.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keydown")
pro.onkeyup.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keyup")
pro.onkeypress.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keypress")
pro.onselect.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "select")
pro.onfocus.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "focus")
pro.onchange.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "change")
pro.onblur.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "blur")
pro.onerror.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "error")
pro.onabord.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "abord")
pro.ondblclick.Enabled = True
pro.mouve.Enabled = True
pro.onload.Enabled = False
pro.onunload.Enabled = False
pro.onkeydown.Enabled = False
pro.onkeyup.Enabled = False
pro.onkeypress.Enabled = False
pro.onselect.Enabled = True
pro.onfocus.Enabled = False
pro.onchange.Enabled = False
pro.onblur.Enabled = False
pro.onerror.Enabled = False
pro.onabord.Enabled = False
pro.zondblclick.Enabled = pro.ondblclick.Enabled
pro.zmouve.Enabled = pro.mouve.Enabled
pro.zonload.Enabled = pro.onload.Enabled
pro.zonunload.Enabled = pro.onunload.Enabled
pro.zonkeydown.Enabled = pro.onkeydown.Enabled
pro.zonkeyup.Enabled = pro.onkeyup.Enabled
pro.zonkeypress.Enabled = pro.onkeypress.Enabled
pro.zonselect.Enabled = pro.onselect.Enabled
pro.zonfocus.Enabled = pro.onfocus.Enabled
pro.zonchange.Enabled = pro.onchange.Enabled
pro.zonblur.Enabled = pro.onblur.Enabled
pro.zonerror.Enabled = pro.onerror.Enabled
pro.zonabord.Enabled = pro.onabord.Enabled
pro.fgcol.Enabled = False
pro.zfgcolor.Enabled = pro.fgcol.Enabled
pro.nam.Enabled = True
pro.valu.Enabled = True
pro.bgcol.Enabled = False
pro.title.Enabled = False
pro.alt.Enabled = False
pro.RO.Enabled = False
pro.border.Enabled = False
pro.path.Enabled = False
pro.zname.Enabled = pro.nam.Enabled
pro.zvalue.Enabled = pro.valu.Enabled
pro.zbgcolor.Enabled = pro.bgcol.Enabled
pro.ztitle.Enabled = pro.title.Enabled
pro.zalt.Enabled = pro.alt.Enabled
pro.zreadonly.Enabled = pro.RO.Enabled
pro.zborder.Enabled = pro.border.Enabled
pro.zpath.Enabled = pro.path.Enabled
pro.onclick.Enabled = True
pro.over.Enabled = True
pro.out.Enabled = True
pro.down.Enabled = True
pro.zclick.Enabled = pro.onclick.Enabled
pro.zover.Enabled = pro.over.Enabled
pro.zout.Enabled = pro.out.Enabled
pro.zdown.Enabled = pro.down.Enabled

rect(0).Left = ccombo(Index).Left - 8
rect(0).Top = ccombo(Index).Top - 8
rect(0).Visible = True
rect(1).Left = ccombo(Index).Left + ccombo(Index).Width / 2 - rect(1).Width / 2
rect(1).Top = ccombo(Index).Top - 8
rect(1).Visible = True
rect(2).Left = ccombo(Index).Left + ccombo(Index).Width
rect(2).Top = ccombo(Index).Top - 8
rect(2).Visible = True
rect(3).Left = ccombo(Index).Left - 8
rect(3).Top = ccombo(Index).Top + ccombo(Index).Height / 2 - rect(3).Height / 2
rect(3).Visible = True
rect(4).Left = ccombo(Index).Left + ccombo(Index).Width
rect(4).Top = ccombo(Index).Top + ccombo(Index).Height / 2 - rect(4).Height / 2
rect(4).Visible = True
rect(5).Left = ccombo(Index).Left - 8
rect(5).Top = ccombo(Index).Top + ccombo(Index).Height
rect(5).Visible = True
rect(6).Left = ccombo(Index).Left + ccombo(Index).Width / 2 - rect(1).Width / 2
rect(6).Top = ccombo(Index).Top + ccombo(Index).Height
rect(6).Visible = True
rect(7).Left = ccombo(Index).Left + ccombo(Index).Width
rect(7).Top = ccombo(Index).Top + ccombo(Index).Height
rect(7).Visible = True
For I = 0 To 3
Line1(I).BorderColor = zone.ForeColor
Line1(I).Visible = True
Next I
Line1(0).X1 = rect(0).Left + rect(0).Width
Line1(0).Y1 = rect(0).Top + rect(0).Height
Line1(0).X2 = rect(2).Left + rect(2).Width
Line1(0).Y2 = rect(2).Top + rect(2).Height
Line1(1).X1 = rect(2).Left
Line1(1).Y1 = rect(2).Top
Line1(1).X2 = rect(7).Left
Line1(1).Y2 = rect(7).Top
Line1(2).X1 = rect(7).Left
Line1(2).Y1 = rect(7).Top
Line1(2).X2 = rect(5).Left + rect(0).Width
Line1(2).Y2 = rect(5).Top
Line1(3).X1 = rect(5).Left + rect(0).Width
Line1(3).Y1 = rect(5).Top
Line1(3).X2 = rect(0).Left + rect(0).Width
Line1(3).Y2 = rect(0).Top + rect(0).Height
pro.Cheight.Text = ccombo(Index).Height
pro.Cwidth.Text = ccombo(Index).Width
pro.Ctop.Text = ccombo(Index).Top
pro.Cleft.Text = ccombo(Index).Left
End Sub

Private Sub ch_Change(Index As Integer)
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "value", ch(Index).Text)
pro.valu.Text = ch(Index).Text
End Sub

Private Sub ch_Click(Index As Integer)
pro.Wtype.Text = "ch"
pro.Witem = Index
'pro.Caption = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "caption")
pro.nam.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "name")
pro.valu.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "value")
pro.bgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "bgcolor")
pro.title.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "title")
pro.alt.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "alt")
pro.RO.Value = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "ro")
pro.border.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "border")
pro.path.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "path")
pro.onclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "click")
pro.out.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "out")
pro.over.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "over")
pro.down.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "down")
pro.fgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "fgcolor")
pro.ondblclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "dblclick")
pro.mouve.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "move")
pro.onload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "load")
pro.onunload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "unload")
pro.onkeydown.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keydown")
pro.onkeyup.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keyup")
pro.onkeypress.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keypress")
pro.onselect.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "select")
pro.onfocus.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "focus")
pro.onchange.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "change")
pro.onblur.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "blur")
pro.onerror.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "error")
pro.onabord.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "abord")
pro.ondblclick.Enabled = False
pro.mouve.Enabled = False
pro.onload.Enabled = False
pro.onunload.Enabled = False
pro.onkeydown.Enabled = False
pro.onkeyup.Enabled = False
pro.onkeypress.Enabled = False
pro.onselect.Enabled = False
pro.onfocus.Enabled = False
pro.onchange.Enabled = False
pro.onblur.Enabled = False
pro.onerror.Enabled = False
pro.onabord.Enabled = False
pro.zondblclick.Enabled = pro.ondblclick.Enabled
pro.zmouve.Enabled = pro.mouve.Enabled
pro.zonload.Enabled = pro.onload.Enabled
pro.zonunload.Enabled = pro.onunload.Enabled
pro.zonkeydown.Enabled = pro.onkeydown.Enabled
pro.zonkeyup.Enabled = pro.onkeyup.Enabled
pro.zonkeypress.Enabled = pro.onkeypress.Enabled
pro.zonselect.Enabled = pro.onselect.Enabled
pro.zonfocus.Enabled = pro.onfocus.Enabled
pro.zonchange.Enabled = pro.onchange.Enabled
pro.zonblur.Enabled = pro.onblur.Enabled
pro.zonerror.Enabled = pro.onerror.Enabled
pro.zonabord.Enabled = pro.onabord.Enabled
pro.fgcol.Enabled = False
pro.zfgcolor.Enabled = pro.fgcol.Enabled
pro.nam.Enabled = True
pro.valu.Enabled = True
pro.bgcol.Enabled = False
pro.title.Enabled = False
pro.alt.Enabled = False
pro.RO.Enabled = False
pro.border.Enabled = False
pro.path.Enabled = False
pro.zname.Enabled = pro.nam.Enabled
pro.zvalue.Enabled = pro.valu.Enabled
pro.zbgcolor.Enabled = pro.bgcol.Enabled
pro.ztitle.Enabled = pro.title.Enabled
pro.zalt.Enabled = pro.alt.Enabled
pro.zreadonly.Enabled = pro.RO.Enabled
pro.zborder.Enabled = pro.border.Enabled
pro.zpath.Enabled = pro.path.Enabled
pro.onclick.Enabled = False
pro.over.Enabled = False
pro.out.Enabled = False
pro.down.Enabled = False
pro.zclick.Enabled = pro.onclick.Enabled
pro.zover.Enabled = pro.over.Enabled
pro.zout.Enabled = pro.out.Enabled
pro.zdown.Enabled = pro.down.Enabled

rect(0).Left = ch(Index).Left - 8
rect(0).Top = ch(Index).Top - 8
rect(0).Visible = True
rect(1).Left = ch(Index).Left + ch(Index).Width / 2 - rect(1).Width / 2
rect(1).Top = ch(Index).Top - 8
rect(1).Visible = True
rect(2).Left = ch(Index).Left + ch(Index).Width
rect(2).Top = ch(Index).Top - 8
rect(2).Visible = True
rect(3).Left = ch(Index).Left - 8
rect(3).Top = ch(Index).Top + ch(Index).Height / 2 - rect(3).Height / 2
rect(3).Visible = True
rect(4).Left = ch(Index).Left + ch(Index).Width
rect(4).Top = ch(Index).Top + ch(Index).Height / 2 - rect(4).Height / 2
rect(4).Visible = True
rect(5).Left = ch(Index).Left - 8
rect(5).Top = ch(Index).Top + ch(Index).Height
rect(5).Visible = True
rect(6).Left = ch(Index).Left + ch(Index).Width / 2 - rect(1).Width / 2
rect(6).Top = ch(Index).Top + ch(Index).Height
rect(6).Visible = True
rect(7).Left = ch(Index).Left + ch(Index).Width
rect(7).Top = ch(Index).Top + ch(Index).Height
rect(7).Visible = True
For I = 0 To 3
Line1(I).BorderColor = zone.ForeColor
Line1(I).Visible = True
Next I
Line1(0).X1 = rect(0).Left + rect(0).Width
Line1(0).Y1 = rect(0).Top + rect(0).Height
Line1(0).X2 = rect(2).Left + rect(2).Width
Line1(0).Y2 = rect(2).Top + rect(2).Height
Line1(1).X1 = rect(2).Left
Line1(1).Y1 = rect(2).Top
Line1(1).X2 = rect(7).Left
Line1(1).Y2 = rect(7).Top
Line1(2).X1 = rect(7).Left
Line1(2).Y1 = rect(7).Top
Line1(2).X2 = rect(5).Left + rect(0).Width
Line1(2).Y2 = rect(5).Top
Line1(3).X1 = rect(5).Left + rect(0).Width
Line1(3).Y1 = rect(5).Top
Line1(3).X2 = rect(0).Left + rect(0).Width
Line1(3).Y2 = rect(0).Top + rect(0).Height
pro.Cheight.Text = ch(Index).Height
pro.Cwidth.Text = ch(Index).Width
pro.Ctop.Text = ch(Index).Top
pro.Cleft.Text = ch(Index).Left
End Sub

Private Sub ci_Click(Index As Integer)
pro.Wtype.Text = "ci"
pro.Witem = Index
'pro.Caption = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "caption")
pro.nam.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "name")
pro.valu.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "value")
pro.bgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "bgcolor")
pro.title.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "title")
pro.alt.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "alt")
pro.RO.Value = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "ro")
pro.border.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "border")
pro.path.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "path")
pro.onclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "click")
pro.out.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "out")
pro.over.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "over")
pro.down.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "down")
pro.fgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "fgcolor")
pro.ondblclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "dblclick")
pro.mouve.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "move")
pro.onload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "load")
pro.onunload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "unload")
pro.onkeydown.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keydown")
pro.onkeyup.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keyup")
pro.onkeypress.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keypress")
pro.onselect.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "select")
pro.onfocus.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "focus")
pro.onchange.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "change")
pro.onblur.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "blur")
pro.onerror.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "error")
pro.onabord.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "abord")
pro.ondblclick.Enabled = True
pro.mouve.Enabled = True
pro.onload.Enabled = False
pro.onunload.Enabled = False
pro.onkeydown.Enabled = False
pro.onkeyup.Enabled = False
pro.onkeypress.Enabled = False
pro.onselect.Enabled = True
pro.onfocus.Enabled = False
pro.onchange.Enabled = False
pro.onblur.Enabled = False
pro.onerror.Enabled = True
pro.onabord.Enabled = True
pro.zondblclick.Enabled = pro.ondblclick.Enabled
pro.zmouve.Enabled = pro.mouve.Enabled
pro.zonload.Enabled = pro.onload.Enabled
pro.zonunload.Enabled = pro.onunload.Enabled
pro.zonkeydown.Enabled = pro.onkeydown.Enabled
pro.zonkeyup.Enabled = pro.onkeyup.Enabled
pro.zonkeypress.Enabled = pro.onkeypress.Enabled
pro.zonselect.Enabled = pro.onselect.Enabled
pro.zonfocus.Enabled = pro.onfocus.Enabled
pro.zonchange.Enabled = pro.onchange.Enabled
pro.zonblur.Enabled = pro.onblur.Enabled
pro.zonerror.Enabled = pro.onerror.Enabled
pro.zonabord.Enabled = pro.onabord.Enabled
pro.fgcol.Enabled = False
pro.zfgcolor.Enabled = pro.fgcol.Enabled
pro.nam.Enabled = True
pro.valu.Enabled = False
pro.bgcol.Enabled = False
pro.title.Enabled = False
pro.alt.Enabled = True
pro.RO.Enabled = False
pro.border.Enabled = True
pro.path.Enabled = True
pro.zname.Enabled = pro.nam.Enabled
pro.zvalue.Enabled = pro.valu.Enabled
pro.zbgcolor.Enabled = pro.bgcol.Enabled
pro.ztitle.Enabled = pro.title.Enabled
pro.zalt.Enabled = pro.alt.Enabled
pro.zreadonly.Enabled = pro.RO.Enabled
pro.zborder.Enabled = pro.border.Enabled
pro.zpath.Enabled = pro.path.Enabled
pro.onclick.Enabled = True
pro.over.Enabled = True
pro.out.Enabled = True
pro.down.Enabled = True
pro.zclick.Enabled = pro.onclick.Enabled
pro.zover.Enabled = pro.over.Enabled
pro.zout.Enabled = pro.out.Enabled
pro.zdown.Enabled = pro.down.Enabled

rect(0).Left = ci(Index).Left - 8
rect(0).Top = ci(Index).Top - 8
rect(0).Visible = True
rect(1).Left = ci(Index).Left + ci(Index).Width / 2 - rect(1).Width / 2
rect(1).Top = ci(Index).Top - 8
rect(1).Visible = True
rect(2).Left = ci(Index).Left + ci(Index).Width
rect(2).Top = ci(Index).Top - 8
rect(2).Visible = True
rect(3).Left = ci(Index).Left - 8
rect(3).Top = ci(Index).Top + ci(Index).Height / 2 - rect(3).Height / 2
rect(3).Visible = True
rect(4).Left = ci(Index).Left + ci(Index).Width
rect(4).Top = ci(Index).Top + ci(Index).Height / 2 - rect(4).Height / 2
rect(4).Visible = True
rect(5).Left = ci(Index).Left - 8
rect(5).Top = ci(Index).Top + ci(Index).Height
rect(5).Visible = True
rect(6).Left = ci(Index).Left + ci(Index).Width / 2 - rect(1).Width / 2
rect(6).Top = ci(Index).Top + ci(Index).Height
rect(6).Visible = True
rect(7).Left = ci(Index).Left + ci(Index).Width
rect(7).Top = ci(Index).Top + ci(Index).Height
rect(7).Visible = True
For I = 0 To 3
Line1(I).BorderColor = zone.ForeColor
Line1(I).Visible = True
Next I
Line1(0).X1 = rect(0).Left + rect(0).Width
Line1(0).Y1 = rect(0).Top + rect(0).Height
Line1(0).X2 = rect(2).Left + rect(2).Width
Line1(0).Y2 = rect(2).Top + rect(2).Height
Line1(1).X1 = rect(2).Left
Line1(1).Y1 = rect(2).Top
Line1(1).X2 = rect(7).Left
Line1(1).Y2 = rect(7).Top
Line1(2).X1 = rect(7).Left
Line1(2).Y1 = rect(7).Top
Line1(2).X2 = rect(5).Left + rect(0).Width
Line1(2).Y2 = rect(5).Top
Line1(3).X1 = rect(5).Left + rect(0).Width
Line1(3).Y1 = rect(5).Top
Line1(3).X2 = rect(0).Left + rect(0).Width
Line1(3).Y2 = rect(0).Top + rect(0).Height
pro.Cheight.Text = ci(Index).Height
pro.Cwidth.Text = ci(Index).Width
pro.Ctop.Text = ci(Index).Top
pro.Cleft.Text = ci(Index).Left
End Sub

Private Sub cli_Click(Index As Integer)
pro.Wtype.Text = "cli"
pro.Witem = Index
'pro.Caption = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "caption")
pro.nam.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "name")
pro.valu.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "value")
pro.bgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "bgcolor")
pro.title.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "title")
pro.alt.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "alt")
pro.RO.Value = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "ro")
pro.border.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "border")
pro.path.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "path")
pro.onclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "click")
pro.out.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "out")
pro.over.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "over")
pro.down.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "down")
pro.fgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "fgcolor")
pro.ondblclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "dblclick")
pro.mouve.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "move")
pro.onload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "load")
pro.onunload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "unload")
pro.onkeydown.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keydown")
pro.onkeyup.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keyup")
pro.onkeypress.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keypress")
pro.onselect.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "select")
pro.onfocus.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "focus")
pro.onchange.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "change")
pro.onblur.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "blur")
pro.onerror.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "error")
pro.onabord.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "abord")
pro.ondblclick.Enabled = True
pro.mouve.Enabled = True
pro.onload.Enabled = False
pro.onunload.Enabled = False
pro.onkeydown.Enabled = False
pro.onkeyup.Enabled = False
pro.onkeypress.Enabled = False
pro.onselect.Enabled = False
pro.onfocus.Enabled = False
pro.onchange.Enabled = False
pro.onblur.Enabled = False
pro.onerror.Enabled = False
pro.onabord.Enabled = False
pro.zondblclick.Enabled = pro.ondblclick.Enabled
pro.zmouve.Enabled = pro.mouve.Enabled
pro.zonload.Enabled = pro.onload.Enabled
pro.zonunload.Enabled = pro.onunload.Enabled
pro.zonkeydown.Enabled = pro.onkeydown.Enabled
pro.zonkeyup.Enabled = pro.onkeyup.Enabled
pro.zonkeypress.Enabled = pro.onkeypress.Enabled
pro.zonselect.Enabled = pro.onselect.Enabled
pro.zonfocus.Enabled = pro.onfocus.Enabled
pro.zonchange.Enabled = pro.onchange.Enabled
pro.zonblur.Enabled = pro.onblur.Enabled
pro.zonerror.Enabled = pro.onerror.Enabled
pro.zonabord.Enabled = pro.onabord.Enabled
pro.fgcol.Enabled = False
pro.zfgcolor.Enabled = pro.fgcol.Enabled
pro.nam.Enabled = False
pro.valu.Enabled = True
pro.bgcol.Enabled = False
pro.title.Enabled = True
pro.alt.Enabled = False
pro.RO.Enabled = False
pro.border.Enabled = False
pro.path.Enabled = True
pro.zname.Enabled = pro.nam.Enabled
pro.zvalue.Enabled = pro.valu.Enabled
pro.zbgcolor.Enabled = pro.bgcol.Enabled
pro.ztitle.Enabled = pro.title.Enabled
pro.zalt.Enabled = pro.alt.Enabled
pro.zreadonly.Enabled = pro.RO.Enabled
pro.zborder.Enabled = pro.border.Enabled
pro.zpath.Enabled = pro.path.Enabled
pro.onclick.Enabled = True
pro.over.Enabled = True
pro.out.Enabled = True
pro.down.Enabled = True
pro.zclick.Enabled = pro.onclick.Enabled
pro.zover.Enabled = pro.over.Enabled
pro.zout.Enabled = pro.out.Enabled
pro.zdown.Enabled = pro.down.Enabled

rect(0).Left = cli(Index).Left - 8
rect(0).Top = cli(Index).Top - 8
rect(0).Visible = True
rect(1).Left = cli(Index).Left + cli(Index).Width / 2 - rect(1).Width / 2
rect(1).Top = cli(Index).Top - 8
rect(1).Visible = True
rect(2).Left = cli(Index).Left + cli(Index).Width
rect(2).Top = cli(Index).Top - 8
rect(2).Visible = True
rect(3).Left = cli(Index).Left - 8
rect(3).Top = cli(Index).Top + cli(Index).Height / 2 - rect(3).Height / 2
rect(3).Visible = True
rect(4).Left = cli(Index).Left + cli(Index).Width
rect(4).Top = cli(Index).Top + cli(Index).Height / 2 - rect(4).Height / 2
rect(4).Visible = True
rect(5).Left = cli(Index).Left - 8
rect(5).Top = cli(Index).Top + cli(Index).Height
rect(5).Visible = True
rect(6).Left = cli(Index).Left + cli(Index).Width / 2 - rect(1).Width / 2
rect(6).Top = cli(Index).Top + cli(Index).Height
rect(6).Visible = True
rect(7).Left = cli(Index).Left + cli(Index).Width
rect(7).Top = cli(Index).Top + cli(Index).Height
rect(7).Visible = True
For I = 0 To 3
Line1(I).BorderColor = zone.ForeColor
Line1(I).Visible = True
Next I
Line1(0).X1 = rect(0).Left + rect(0).Width
Line1(0).Y1 = rect(0).Top + rect(0).Height
Line1(0).X2 = rect(2).Left + rect(2).Width
Line1(0).Y2 = rect(2).Top + rect(2).Height
Line1(1).X1 = rect(2).Left
Line1(1).Y1 = rect(2).Top
Line1(1).X2 = rect(7).Left
Line1(1).Y2 = rect(7).Top
Line1(2).X1 = rect(7).Left
Line1(2).Y1 = rect(7).Top
Line1(2).X2 = rect(5).Left + rect(0).Width
Line1(2).Y2 = rect(5).Top
Line1(3).X1 = rect(5).Left + rect(0).Width
Line1(3).Y1 = rect(5).Top
Line1(3).X2 = rect(0).Left + rect(0).Width
Line1(3).Y2 = rect(0).Top + rect(0).Height
pro.Cheight.Text = cli(Index).Height
pro.Cwidth.Text = cli(Index).Width
pro.Ctop.Text = cli(Index).Top
pro.Cleft.Text = cli(Index).Left
End Sub

Private Sub clist_Click(Index As Integer)
pro.Wtype.Text = "clist"
pro.Witem = Index
'pro.Caption = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "caption")
pro.nam.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "name")
pro.valu.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "value")
pro.bgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "bgcolor")
pro.title.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "title")
pro.alt.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "alt")
pro.RO.Value = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "ro")
pro.border.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "border")
pro.path.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "path")
pro.onclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "click")
pro.out.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "out")
pro.over.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "over")
pro.down.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "down")
pro.fgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "fgcolor")
pro.ondblclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "dblclick")
pro.mouve.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "move")
pro.onload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "load")
pro.onunload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "unload")
pro.onkeydown.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keydown")
pro.onkeyup.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keyup")
pro.onkeypress.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keypress")
pro.onselect.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "select")
pro.onfocus.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "focus")
pro.onchange.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "change")
pro.onblur.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "blur")
pro.onerror.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "error")
pro.onabord.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "abord")
pro.ondblclick.Enabled = True
pro.mouve.Enabled = True
pro.onload.Enabled = False
pro.onunload.Enabled = False
pro.onkeydown.Enabled = False
pro.onkeyup.Enabled = False
pro.onkeypress.Enabled = False
pro.onselect.Enabled = True
pro.onfocus.Enabled = False
pro.onchange.Enabled = False
pro.onblur.Enabled = False
pro.onerror.Enabled = False
pro.onabord.Enabled = False
pro.zondblclick.Enabled = pro.ondblclick.Enabled
pro.zmouve.Enabled = pro.mouve.Enabled
pro.zonload.Enabled = pro.onload.Enabled
pro.zonunload.Enabled = pro.onunload.Enabled
pro.zonkeydown.Enabled = pro.onkeydown.Enabled
pro.zonkeyup.Enabled = pro.onkeyup.Enabled
pro.zonkeypress.Enabled = pro.onkeypress.Enabled
pro.zonselect.Enabled = pro.onselect.Enabled
pro.zonfocus.Enabled = pro.onfocus.Enabled
pro.zonchange.Enabled = pro.onchange.Enabled
pro.zonblur.Enabled = pro.onblur.Enabled
pro.zonerror.Enabled = pro.onerror.Enabled
pro.zonabord.Enabled = pro.onabord.Enabled
pro.fgcol.Enabled = False
pro.zfgcolor.Enabled = pro.fgcol.Enabled
pro.nam.Enabled = True
pro.valu.Enabled = True
pro.bgcol.Enabled = False
pro.title.Enabled = True
pro.alt.Enabled = False
pro.RO.Enabled = False
pro.border.Enabled = False
pro.path.Enabled = False
pro.zname.Enabled = pro.nam.Enabled
pro.zvalue.Enabled = pro.valu.Enabled
pro.zbgcolor.Enabled = pro.bgcol.Enabled
pro.ztitle.Enabled = pro.title.Enabled
pro.zalt.Enabled = pro.alt.Enabled
pro.zreadonly.Enabled = pro.RO.Enabled
pro.zborder.Enabled = pro.border.Enabled
pro.zpath.Enabled = pro.path.Enabled
pro.onclick.Enabled = True
pro.over.Enabled = True
pro.out.Enabled = True
pro.down.Enabled = True
pro.zclick.Enabled = pro.onclick.Enabled
pro.zover.Enabled = pro.over.Enabled
pro.zout.Enabled = pro.out.Enabled
pro.zdown.Enabled = pro.down.Enabled

rect(0).Left = clist(Index).Left - 8
rect(0).Top = clist(Index).Top - 8
rect(0).Visible = True
rect(1).Left = clist(Index).Left + clist(Index).Width / 2 - rect(1).Width / 2
rect(1).Top = clist(Index).Top - 8
rect(1).Visible = True
rect(2).Left = clist(Index).Left + clist(Index).Width
rect(2).Top = clist(Index).Top - 8
rect(2).Visible = True
rect(3).Left = clist(Index).Left - 8
rect(3).Top = clist(Index).Top + clist(Index).Height / 2 - rect(3).Height / 2
rect(3).Visible = True
rect(4).Left = clist(Index).Left + clist(Index).Width
rect(4).Top = clist(Index).Top + clist(Index).Height / 2 - rect(4).Height / 2
rect(4).Visible = True
rect(5).Left = clist(Index).Left - 8
rect(5).Top = clist(Index).Top + clist(Index).Height
rect(5).Visible = True
rect(6).Left = clist(Index).Left + clist(Index).Width / 2 - rect(1).Width / 2
rect(6).Top = clist(Index).Top + clist(Index).Height
rect(6).Visible = True
rect(7).Left = clist(Index).Left + clist(Index).Width
rect(7).Top = clist(Index).Top + clist(Index).Height
rect(7).Visible = True
For I = 0 To 3
Line1(I).BorderColor = zone.ForeColor
Line1(I).Visible = True
Next I
Line1(0).X1 = rect(0).Left + rect(0).Width
Line1(0).Y1 = rect(0).Top + rect(0).Height
Line1(0).X2 = rect(2).Left + rect(2).Width
Line1(0).Y2 = rect(2).Top + rect(2).Height
Line1(1).X1 = rect(2).Left
Line1(1).Y1 = rect(2).Top
Line1(1).X2 = rect(7).Left
Line1(1).Y2 = rect(7).Top
Line1(2).X1 = rect(7).Left
Line1(2).Y1 = rect(7).Top
Line1(2).X2 = rect(5).Left + rect(0).Width
Line1(2).Y2 = rect(5).Top
Line1(3).X1 = rect(5).Left + rect(0).Width
Line1(3).Y1 = rect(5).Top
Line1(3).X2 = rect(0).Left + rect(0).Width
Line1(3).Y2 = rect(0).Top + rect(0).Height
pro.Cheight.Text = clist(Index).Height
pro.Cwidth.Text = clist(Index).Width
pro.Ctop.Text = clist(Index).Top
pro.Cleft.Text = clist(Index).Left
End Sub

Private Sub cp_Change(Index As Integer)
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "value", cp(Index).Text)
pro.valu.Text = cp(Index).Text
End Sub

Private Sub cp_Click(Index As Integer)
pro.Wtype.Text = "cp"
pro.Witem = Index
'pro.Caption = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "caption")
pro.nam.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "name")
pro.valu.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "value")
pro.bgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "bgcolor")
pro.title.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "title")
pro.alt.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "alt")
pro.RO.Value = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "ro")
pro.border.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "border")
pro.path.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "path")
pro.onclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "click")
pro.out.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "out")
pro.over.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "over")
pro.down.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "down")
pro.fgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "fgcolor")
pro.ondblclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "dblclick")
pro.mouve.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "move")
pro.onload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "load")
pro.onunload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "unload")
pro.onkeydown.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keydown")
pro.onkeyup.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keyup")
pro.onkeypress.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keypress")
pro.onselect.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "select")
pro.onfocus.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "focus")
pro.onchange.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "change")
pro.onblur.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "blur")
pro.onerror.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "error")
pro.onabord.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "abord")
pro.ondblclick.Enabled = True
pro.mouve.Enabled = True
pro.onload.Enabled = False
pro.onunload.Enabled = False
pro.onkeydown.Enabled = False
pro.onkeyup.Enabled = False
pro.onkeypress.Enabled = False
pro.onselect.Enabled = True
pro.onfocus.Enabled = True
pro.onchange.Enabled = True
pro.onblur.Enabled = True
pro.onerror.Enabled = False
pro.onabord.Enabled = False
pro.zondblclick.Enabled = pro.ondblclick.Enabled
pro.zmouve.Enabled = pro.mouve.Enabled
pro.zonload.Enabled = pro.onload.Enabled
pro.zonunload.Enabled = pro.onunload.Enabled
pro.zonkeydown.Enabled = pro.onkeydown.Enabled
pro.zonkeyup.Enabled = pro.onkeyup.Enabled
pro.zonkeypress.Enabled = pro.onkeypress.Enabled
pro.zonselect.Enabled = pro.onselect.Enabled
pro.zonfocus.Enabled = pro.onfocus.Enabled
pro.zonchange.Enabled = pro.onchange.Enabled
pro.zonblur.Enabled = pro.onblur.Enabled
pro.zonerror.Enabled = pro.onerror.Enabled
pro.zonabord.Enabled = pro.onabord.Enabled
pro.fgcol.Enabled = False
pro.zfgcolor.Enabled = pro.fgcol.Enabled
pro.nam.Enabled = True
pro.valu.Enabled = True
pro.bgcol.Enabled = False
pro.title.Enabled = False
pro.alt.Enabled = True
pro.RO.Enabled = True
pro.border.Enabled = False
pro.path.Enabled = False
pro.zname.Enabled = pro.nam.Enabled
pro.zvalue.Enabled = pro.valu.Enabled
pro.zbgcolor.Enabled = pro.bgcol.Enabled
pro.ztitle.Enabled = pro.title.Enabled
pro.zalt.Enabled = pro.alt.Enabled
pro.zreadonly.Enabled = pro.RO.Enabled
pro.zborder.Enabled = pro.border.Enabled
pro.zpath.Enabled = pro.path.Enabled
pro.onclick.Enabled = True
pro.over.Enabled = True
pro.out.Enabled = True
pro.down.Enabled = True
pro.zclick.Enabled = pro.onclick.Enabled
pro.zover.Enabled = pro.over.Enabled
pro.zout.Enabled = pro.out.Enabled
pro.zdown.Enabled = pro.down.Enabled

rect(0).Left = cp(Index).Left - 8
rect(0).Top = cp(Index).Top - 8
rect(0).Visible = True
rect(1).Left = cp(Index).Left + cp(Index).Width / 2 - rect(1).Width / 2
rect(1).Top = cp(Index).Top - 8
rect(1).Visible = True
rect(2).Left = cp(Index).Left + cp(Index).Width
rect(2).Top = cp(Index).Top - 8
rect(2).Visible = True
rect(3).Left = cp(Index).Left - 8
rect(3).Top = cp(Index).Top + cp(Index).Height / 2 - rect(3).Height / 2
rect(3).Visible = True
rect(4).Left = cp(Index).Left + cp(Index).Width
rect(4).Top = cp(Index).Top + cp(Index).Height / 2 - rect(4).Height / 2
rect(4).Visible = True
rect(5).Left = cp(Index).Left - 8
rect(5).Top = cp(Index).Top + cp(Index).Height
rect(5).Visible = True
rect(6).Left = cp(Index).Left + cp(Index).Width / 2 - rect(1).Width / 2
rect(6).Top = cp(Index).Top + cp(Index).Height
rect(6).Visible = True
rect(7).Left = cp(Index).Left + cp(Index).Width
rect(7).Top = cp(Index).Top + cp(Index).Height
rect(7).Visible = True
For I = 0 To 3
Line1(I).BorderColor = zone.ForeColor
Line1(I).Visible = True
Next I
Line1(0).X1 = rect(0).Left + rect(0).Width
Line1(0).Y1 = rect(0).Top + rect(0).Height
Line1(0).X2 = rect(2).Left + rect(2).Width
Line1(0).Y2 = rect(2).Top + rect(2).Height
Line1(1).X1 = rect(2).Left
Line1(1).Y1 = rect(2).Top
Line1(1).X2 = rect(7).Left
Line1(1).Y2 = rect(7).Top
Line1(2).X1 = rect(7).Left
Line1(2).Y1 = rect(7).Top
Line1(2).X2 = rect(5).Left + rect(0).Width
Line1(2).Y2 = rect(5).Top
Line1(3).X1 = rect(5).Left + rect(0).Width
Line1(3).Y1 = rect(5).Top
Line1(3).X2 = rect(0).Left + rect(0).Width
Line1(3).Y2 = rect(0).Top + rect(0).Height
pro.Cheight.Text = cp(Index).Height
pro.Cwidth.Text = cp(Index).Width
pro.Ctop.Text = cp(Index).Top
pro.Cleft.Text = cp(Index).Left
End Sub

Private Sub csub_Click(Index As Integer)
pro.Wtype.Text = "csub"
pro.Witem = Index
'pro.Caption = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "caption")
pro.nam.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "name")
pro.valu.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "value")
pro.bgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "bgcolor")
pro.title.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "title")
pro.alt.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "alt")
pro.RO.Value = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "ro")
pro.border.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "border")
pro.path.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "path")
pro.onclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "click")
pro.out.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "out")
pro.over.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "over")
pro.down.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "down")
pro.fgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "fgcolor")
pro.ondblclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "dblclick")
pro.mouve.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "move")
pro.onload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "load")
pro.onunload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "unload")
pro.onkeydown.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keydown")
pro.onkeyup.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keyup")
pro.onkeypress.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keypress")
pro.onselect.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "select")
pro.onfocus.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "focus")
pro.onchange.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "change")
pro.onblur.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "blur")
pro.onerror.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "error")
pro.onabord.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "abord")
pro.ondblclick.Enabled = True
pro.mouve.Enabled = True
pro.onload.Enabled = False
pro.onunload.Enabled = False
pro.onkeydown.Enabled = False
pro.onkeyup.Enabled = False
pro.onkeypress.Enabled = False
pro.onselect.Enabled = True
pro.onfocus.Enabled = False
pro.onchange.Enabled = False
pro.onblur.Enabled = False
pro.onerror.Enabled = False
pro.onabord.Enabled = False
pro.zondblclick.Enabled = pro.ondblclick.Enabled
pro.zmouve.Enabled = pro.mouve.Enabled
pro.zonload.Enabled = pro.onload.Enabled
pro.zonunload.Enabled = pro.onunload.Enabled
pro.zonkeydown.Enabled = pro.onkeydown.Enabled
pro.zonkeyup.Enabled = pro.onkeyup.Enabled
pro.zonkeypress.Enabled = pro.onkeypress.Enabled
pro.zonselect.Enabled = pro.onselect.Enabled
pro.zonfocus.Enabled = pro.onfocus.Enabled
pro.zonchange.Enabled = pro.onchange.Enabled
pro.zonblur.Enabled = pro.onblur.Enabled
pro.zonerror.Enabled = pro.onerror.Enabled
pro.zonabord.Enabled = pro.onabord.Enabled
pro.fgcol.Enabled = False
pro.zfgcolor.Enabled = pro.fgcol.Enabled
pro.nam.Enabled = True
pro.valu.Enabled = True
pro.bgcol.Enabled = False
pro.title.Enabled = True
pro.alt.Enabled = False
pro.RO.Enabled = False
pro.border.Enabled = False
pro.path.Enabled = False
pro.zname.Enabled = pro.nam.Enabled
pro.zvalue.Enabled = pro.valu.Enabled
pro.zbgcolor.Enabled = pro.bgcol.Enabled
pro.ztitle.Enabled = pro.title.Enabled
pro.zalt.Enabled = pro.alt.Enabled
pro.zreadonly.Enabled = pro.RO.Enabled
pro.zborder.Enabled = pro.border.Enabled
pro.zpath.Enabled = pro.path.Enabled
pro.onclick.Enabled = True
pro.over.Enabled = True
pro.out.Enabled = True
pro.down.Enabled = True
pro.zclick.Enabled = pro.onclick.Enabled
pro.zover.Enabled = pro.over.Enabled
pro.zout.Enabled = pro.out.Enabled
pro.zdown.Enabled = pro.down.Enabled

rect(0).Left = csub(Index).Left - 8
rect(0).Top = csub(Index).Top - 8
rect(0).Visible = True
rect(1).Left = csub(Index).Left + csub(Index).Width / 2 - rect(1).Width / 2
rect(1).Top = csub(Index).Top - 8
rect(1).Visible = True
rect(2).Left = csub(Index).Left + csub(Index).Width
rect(2).Top = csub(Index).Top - 8
rect(2).Visible = True
rect(3).Left = csub(Index).Left - 8
rect(3).Top = csub(Index).Top + csub(Index).Height / 2 - rect(3).Height / 2
rect(3).Visible = True
rect(4).Left = csub(Index).Left + csub(Index).Width
rect(4).Top = csub(Index).Top + csub(Index).Height / 2 - rect(4).Height / 2
rect(4).Visible = True
rect(5).Left = csub(Index).Left - 8
rect(5).Top = csub(Index).Top + csub(Index).Height
rect(5).Visible = True
rect(6).Left = csub(Index).Left + csub(Index).Width / 2 - rect(1).Width / 2
rect(6).Top = csub(Index).Top + csub(Index).Height
rect(6).Visible = True
rect(7).Left = csub(Index).Left + csub(Index).Width
rect(7).Top = csub(Index).Top + csub(Index).Height
rect(7).Visible = True
For I = 0 To 3
Line1(I).BorderColor = zone.ForeColor
Line1(I).Visible = True
Next I
Line1(0).X1 = rect(0).Left + rect(0).Width
Line1(0).Y1 = rect(0).Top + rect(0).Height
Line1(0).X2 = rect(2).Left + rect(2).Width
Line1(0).Y2 = rect(2).Top + rect(2).Height
Line1(1).X1 = rect(2).Left
Line1(1).Y1 = rect(2).Top
Line1(1).X2 = rect(7).Left
Line1(1).Y2 = rect(7).Top
Line1(2).X1 = rect(7).Left
Line1(2).Y1 = rect(7).Top
Line1(2).X2 = rect(5).Left + rect(0).Width
Line1(2).Y2 = rect(5).Top
Line1(3).X1 = rect(5).Left + rect(0).Width
Line1(3).Y1 = rect(5).Top
Line1(3).X2 = rect(0).Left + rect(0).Width
Line1(3).Y2 = rect(0).Top + rect(0).Height
pro.Cheight.Text = csub(Index).Height
pro.Cwidth.Text = csub(Index).Width
pro.Ctop.Text = csub(Index).Top
pro.Cleft.Text = csub(Index).Left
End Sub

Private Sub ct_Change(Index As Integer)
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "value", ct(Index).Text)
pro.valu.Text = ct(Index).Text
End Sub

Private Sub ct_Click(Index As Integer)
pro.Wtype.Text = "ct"
pro.Witem = Index
'pro.Caption = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "caption")
pro.nam.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "name")
pro.valu.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "value")
pro.bgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "bgcolor")
pro.title.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "title")
pro.alt.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "alt")
pro.RO.Value = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "ro")
pro.border.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "border")
pro.path.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "path")
pro.onclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "click")
pro.out.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "out")
pro.over.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "over")
pro.down.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "down")
pro.fgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "fgcolor")
pro.ondblclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "dblclick")
pro.mouve.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "move")
pro.onload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "load")
pro.onunload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "unload")
pro.onkeydown.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keydown")
pro.onkeyup.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keyup")
pro.onkeypress.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keypress")
pro.onselect.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "select")
pro.onfocus.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "focus")
pro.onchange.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "change")
pro.onblur.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "blur")
pro.onerror.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "error")
pro.onabord.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "abord")
pro.ondblclick.Enabled = True
pro.mouve.Enabled = True
pro.onload.Enabled = False
pro.onunload.Enabled = False
pro.onkeydown.Enabled = False
pro.onkeyup.Enabled = False
pro.onkeypress.Enabled = False
pro.onselect.Enabled = True
pro.onfocus.Enabled = True
pro.onchange.Enabled = True
pro.onblur.Enabled = True
pro.onerror.Enabled = False
pro.onabord.Enabled = False
pro.zondblclick.Enabled = pro.ondblclick.Enabled
pro.zmouve.Enabled = pro.mouve.Enabled
pro.zonload.Enabled = pro.onload.Enabled
pro.zonunload.Enabled = pro.onunload.Enabled
pro.zonkeydown.Enabled = pro.onkeydown.Enabled
pro.zonkeyup.Enabled = pro.onkeyup.Enabled
pro.zonkeypress.Enabled = pro.onkeypress.Enabled
pro.zonselect.Enabled = pro.onselect.Enabled
pro.zonfocus.Enabled = pro.onfocus.Enabled
pro.zonchange.Enabled = pro.onchange.Enabled
pro.zonblur.Enabled = pro.onblur.Enabled
pro.zonerror.Enabled = pro.onerror.Enabled
pro.zonabord.Enabled = pro.onabord.Enabled
pro.fgcol.Enabled = False
pro.zfgcolor.Enabled = pro.fgcol.Enabled
pro.nam.Enabled = True
pro.valu.Enabled = True
pro.bgcol.Enabled = False
pro.title.Enabled = False
pro.alt.Enabled = True
pro.RO.Enabled = True
pro.border.Enabled = False
pro.path.Enabled = False
pro.zname.Enabled = pro.nam.Enabled
pro.zvalue.Enabled = pro.valu.Enabled
pro.zbgcolor.Enabled = pro.bgcol.Enabled
pro.ztitle.Enabled = pro.title.Enabled
pro.zalt.Enabled = pro.alt.Enabled
pro.zreadonly.Enabled = pro.RO.Enabled
pro.zborder.Enabled = pro.border.Enabled
pro.zpath.Enabled = pro.path.Enabled
pro.onclick.Enabled = True
pro.over.Enabled = True
pro.out.Enabled = True
pro.down.Enabled = True
pro.zclick.Enabled = pro.onclick.Enabled
pro.zover.Enabled = pro.over.Enabled
pro.zout.Enabled = pro.out.Enabled
pro.zdown.Enabled = pro.down.Enabled

rect(0).Left = ct(Index).Left - 8
rect(0).Top = ct(Index).Top - 8
rect(0).Visible = True
rect(1).Left = ct(Index).Left + ct(Index).Width / 2 - rect(1).Width / 2
rect(1).Top = ct(Index).Top - 8
rect(1).Visible = True
rect(2).Left = ct(Index).Left + ct(Index).Width
rect(2).Top = ct(Index).Top - 8
rect(2).Visible = True
rect(3).Left = ct(Index).Left - 8
rect(3).Top = ct(Index).Top + ct(Index).Height / 2 - rect(3).Height / 2
rect(3).Visible = True
rect(4).Left = ct(Index).Left + ct(Index).Width
rect(4).Top = ct(Index).Top + ct(Index).Height / 2 - rect(4).Height / 2
rect(4).Visible = True
rect(5).Left = ct(Index).Left - 8
rect(5).Top = ct(Index).Top + ct(Index).Height
rect(5).Visible = True
rect(6).Left = ct(Index).Left + ct(Index).Width / 2 - rect(1).Width / 2
rect(6).Top = ct(Index).Top + ct(Index).Height
rect(6).Visible = True
rect(7).Left = ct(Index).Left + ct(Index).Width
rect(7).Top = ct(Index).Top + ct(Index).Height
rect(7).Visible = True
For I = 0 To 3
Line1(I).BorderColor = zone.ForeColor
Line1(I).Visible = True
Next I
Line1(0).X1 = rect(0).Left + rect(0).Width
Line1(0).Y1 = rect(0).Top + rect(0).Height
Line1(0).X2 = rect(2).Left + rect(2).Width
Line1(0).Y2 = rect(2).Top + rect(2).Height
Line1(1).X1 = rect(2).Left
Line1(1).Y1 = rect(2).Top
Line1(1).X2 = rect(7).Left
Line1(1).Y2 = rect(7).Top
Line1(2).X1 = rect(7).Left
Line1(2).Y1 = rect(7).Top
Line1(2).X2 = rect(5).Left + rect(0).Width
Line1(2).Y2 = rect(5).Top
Line1(3).X1 = rect(5).Left + rect(0).Width
Line1(3).Y1 = rect(5).Top
Line1(3).X2 = rect(0).Left + rect(0).Width
Line1(3).Y2 = rect(0).Top + rect(0).Height
pro.Cheight.Text = ct(Index).Height
pro.Cwidth.Text = ct(Index).Width
pro.Ctop.Text = ct(Index).Top
pro.Cleft.Text = ct(Index).Left
End Sub

Private Sub cta_Change(Index As Integer)
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\" & Wtype & Witem, "value", cta(Index).Text)
pro.valu.Text = cta(Index).Text
End Sub

Private Sub cta_Click(Index As Integer)
pro.Wtype.Text = "cta"
pro.Witem = Index
'pro.Caption = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "caption")
pro.nam.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "name")
pro.valu.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "value")
pro.bgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "bgcolor")
pro.title.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "title")
pro.alt.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "alt")
pro.RO.Value = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "ro")
pro.border.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "border")
pro.path.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "path")
pro.onclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "click")
pro.out.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "out")
pro.over.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "over")
pro.down.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "down")
pro.fgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "fgcolor")
pro.ondblclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "dblclick")
pro.mouve.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "move")
pro.onload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "load")
pro.onunload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "unload")
pro.onkeydown.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keydown")
pro.onkeyup.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keyup")
pro.onkeypress.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keypress")
pro.onselect.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "select")
pro.onfocus.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "focus")
pro.onchange.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "change")
pro.onblur.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "blur")
pro.onerror.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "error")
pro.onabord.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "abord")
pro.ondblclick.Enabled = True
pro.mouve.Enabled = True
pro.onload.Enabled = False
pro.onunload.Enabled = False
pro.onkeydown.Enabled = False
pro.onkeyup.Enabled = False
pro.onkeypress.Enabled = False
pro.onselect.Enabled = True
pro.onfocus.Enabled = True
pro.onchange.Enabled = True
pro.onblur.Enabled = True
pro.onerror.Enabled = False
pro.onabord.Enabled = False
pro.zondblclick.Enabled = pro.ondblclick.Enabled
pro.zmouve.Enabled = pro.mouve.Enabled
pro.zonload.Enabled = pro.onload.Enabled
pro.zonunload.Enabled = pro.onunload.Enabled
pro.zonkeydown.Enabled = pro.onkeydown.Enabled
pro.zonkeyup.Enabled = pro.onkeyup.Enabled
pro.zonkeypress.Enabled = pro.onkeypress.Enabled
pro.zonselect.Enabled = pro.onselect.Enabled
pro.zonfocus.Enabled = pro.onfocus.Enabled
pro.zonchange.Enabled = pro.onchange.Enabled
pro.zonblur.Enabled = pro.onblur.Enabled
pro.zonerror.Enabled = pro.onerror.Enabled
pro.zonabord.Enabled = pro.onabord.Enabled
pro.fgcol.Enabled = False
pro.zfgcolor.Enabled = pro.fgcol.Enabled
pro.nam.Enabled = True
pro.valu.Enabled = True
pro.bgcol.Enabled = False
pro.title.Enabled = False
pro.alt.Enabled = False
pro.RO.Enabled = True
pro.border.Enabled = False
pro.path.Enabled = False
pro.zname.Enabled = pro.nam.Enabled
pro.zvalue.Enabled = pro.valu.Enabled
pro.zbgcolor.Enabled = pro.bgcol.Enabled
pro.ztitle.Enabled = pro.title.Enabled
pro.zalt.Enabled = pro.alt.Enabled
pro.zreadonly.Enabled = pro.RO.Enabled
pro.zborder.Enabled = pro.border.Enabled
pro.zpath.Enabled = pro.path.Enabled
pro.onclick.Enabled = False
pro.over.Enabled = False
pro.out.Enabled = False
pro.down.Enabled = False
pro.zclick.Enabled = pro.onclick.Enabled
pro.zover.Enabled = pro.over.Enabled
pro.zout.Enabled = pro.out.Enabled
pro.zdown.Enabled = pro.down.Enabled

rect(0).Left = cta(Index).Left - 8
rect(0).Top = cta(Index).Top - 8
rect(0).Visible = True
rect(1).Left = cta(Index).Left + cta(Index).Width / 2 - rect(1).Width / 2
rect(1).Top = cta(Index).Top - 8
rect(1).Visible = True
rect(2).Left = cta(Index).Left + cta(Index).Width
rect(2).Top = cta(Index).Top - 8
rect(2).Visible = True
rect(3).Left = cta(Index).Left - 8
rect(3).Top = cta(Index).Top + cta(Index).Height / 2 - rect(3).Height / 2
rect(3).Visible = True
rect(4).Left = cta(Index).Left + cta(Index).Width
rect(4).Top = cta(Index).Top + cta(Index).Height / 2 - rect(4).Height / 2
rect(4).Visible = True
rect(5).Left = cta(Index).Left - 8
rect(5).Top = cta(Index).Top + cta(Index).Height
rect(5).Visible = True
rect(6).Left = cta(Index).Left + cta(Index).Width / 2 - rect(1).Width / 2
rect(6).Top = cta(Index).Top + cta(Index).Height
rect(6).Visible = True
rect(7).Left = cta(Index).Left + cta(Index).Width
rect(7).Top = cta(Index).Top + cta(Index).Height
rect(7).Visible = True
For I = 0 To 3
Line1(I).BorderColor = zone.ForeColor
Line1(I).Visible = True
Next I
Line1(0).X1 = rect(0).Left + rect(0).Width
Line1(0).Y1 = rect(0).Top + rect(0).Height
Line1(0).X2 = rect(2).Left + rect(2).Width
Line1(0).Y2 = rect(2).Top + rect(2).Height
Line1(1).X1 = rect(2).Left
Line1(1).Y1 = rect(2).Top
Line1(1).X2 = rect(7).Left
Line1(1).Y2 = rect(7).Top
Line1(2).X1 = rect(7).Left
Line1(2).Y1 = rect(7).Top
Line1(2).X2 = rect(5).Left + rect(0).Width
Line1(2).Y2 = rect(5).Top
Line1(3).X1 = rect(5).Left + rect(0).Width
Line1(3).Y1 = rect(5).Top
Line1(3).X2 = rect(0).Left + rect(0).Width
Line1(3).Y2 = rect(0).Top + rect(0).Height
pro.Cheight.Text = cta(Index).Height
pro.Cwidth.Text = cta(Index).Width
pro.Ctop.Text = cta(Index).Top
pro.Cleft.Text = cta(Index).Left
End Sub

Private Sub dialog_Click(Index As Integer)
Select Case Index
Case 0
js.SelRTF = "alert('votre texte');" & space.Text
Case 1
js.SelRTF = "var input = confirm('votre texte');" & space.Text & _
"if(input == true) {" & space.Text & _
"//si click sur ok..." & space.Text & _
"} else {" & space.Text & _
"//si click sur annuler..." & space.Text & _
"}" & space.Text
Case 3
js.SelRTF = "var input = prompt('votre texte','');" & space.Text
End Select
End Sub

Private Sub download_Click()
WeB "http://www.planet-source-code.com/vb/scripts/ShowZip.asp?lngWId=1&lngCodeId=28005&strZipAccessCode=isua280058101", Me.hWnd
End Sub

Private Sub f_Click(Index As Integer)
Select Case Index
Case 0 'new
reload Me
Form_Load
Case 1 'save as
cd1.DialogTitle = "save as (html page)"
cd1.Filter = "C++ file (*.CPP)|*.CPP|C File (*.c)|*.c|H File (*.h)|*.h"
cd1.ShowSave
If cd1.FileName <> "" Then
Call Generate
Open cd1.FileName For Output As #1
Print #1, Me.RTCode.Text
Close #1
Open Left(cd1.FileName, Len(cd1.FileName) - Len(cd1.FileTitle)) & "vhtml.js" For Output As #1
Print #1, Me.js.Text
Close #1
End If
Case 3
Open "c:\VCPP_defaut.cpp" For Output As #1
Print #1, RTCode.Text
Close #1
Shell Form2.Text1.Text & " c:\VCPP_defaut.cpp"
Case 4 'quit
End
End Select
End Sub

Private Sub Form_Load()
pro.Show
Open "c:\preview.html" For Output As #1
Print #1, ""
Close #1
Call resetREGEDIT
pro.onclick.AddItem ""
pro.ondblclick.AddItem ""
pro.over.AddItem ""
pro.down.AddItem ""
pro.out.AddItem ""
pro.mouve.AddItem ""
pro.onload.AddItem ""
pro.onunload.AddItem ""
pro.onkeydown.AddItem ""
pro.onkeyup.AddItem ""
pro.onkeypress.AddItem ""
pro.onselect.AddItem ""
pro.onfocus.AddItem ""
pro.onchange.AddItem ""
pro.onblur.AddItem ""
pro.onerror.AddItem ""
pro.onabord.AddItem ""
o(11).Value = True
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "width", "800")
Call savestring(HKEY_CURRENT_USER, "Software\VHTML\PageN/A", "height", "600")
pscode.Navigate "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=28005&lngWId=1"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call resetREGEDIT
End Sub

Private Sub js_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu Form1.toolz
End If
End Sub

Private Sub loo_Click(Index As Integer)
Select Case Index
Case 0 'if
js.SelRTF = "if (variable)" & space.Text & _
"{" & space.Text & _
"// action..." & space.Text & _
"}" & space.Text
Case 1 'if...else
js.SelRTF = "if (variable)" & space.Text & _
"{" & space.Text & _
"// action..." & space.Text & _
"} else {" & space.Text & _
"// sinon, action" & space.Text & _
"}"
Case 2 'if...else if...else
js.SelRTF = "if (variable)" & space.Text & _
"{" & space.Text & _
"// action..." & space.Text & _
"} else if(variable) {" & space.Text & _
"// sinon, action" & space.Text & _
"} else {" & space.Text & _
"// sinon, action" & space.Text & _
"}"
Case 3 'with
js.SelRTF = "with (object)" & space.Text & _
"{" & space.Text & _
"// code..." & space.Text & _
"}" & space.Text
Case 4 'switch
js.SelRTF = "switch (variable)" & space.Text & _
"{" & space.Text & _
"case '':" & space.Text & _
"//code..." & space.Text & _
"}" & space.Text
Case 5 'while
js.SelRTF = "while (variable)" & space.Text & _
"{" & space.Text & _
"// code..." & space.Text & _
"}" & space.Text
Case 6 'for
js.SelRTF = "for (Initialisation; Condition; Incrementation)" & space.Text & _
"{" & space.Text & _
"// Instructions..." & space.Text & _
"}" & space.Text
Case 7 'break
js.SelRTF = "break;" & space.Text
End Select
End Sub

Private Sub msP_Click()
Builder.Show
End Sub

Private Sub o_Click(Index As Integer)
If Index <> 11 Then
zone.MousePointer = 2
Else
zone.MousePointer = 0
End If
End Sub

Private Sub opt_Click()
Form2.Show
End Sub

Private Sub rect_Click(Index As Integer)
Select Case Index
Case 0
aa = True
Case 1
bb = True
Case 2
c = True
Case 3
dd = True
Case 4
ee = True
Case 5
ff = True
Case 6
gg = True
Case 7
hh = True
End Select
For I = 0 To 7
rect(I).Visible = False
Next I
End Sub

Private Sub rect_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case pro.Wtype.Text
Case "cb"
ControlResize cb(pro.Witem), Handle, Index
Case "ci"
ControlResize ci(pro.Witem), Handle, Index
Case "cl"
ControlResize cl(pro.Witem), Handle, Index
Case "ct"
ControlResize ct(pro.Witem), Handle, Index
Case "cp"
ControlResize cp(pro.Witem), Handle, Index
Case "cc"
ControlResize cc(pro.Witem), Handle, Index
Case "clist"
ControlResize clist(pro.Witem), Handle, Index
Case "ccombo"
ControlResize ccombo(pro.Witem), Handle, Index
Case "cli"
ControlResize cli(pro.Witem), Handle, Index
Case "ch"
ControlResize ch(pro.Witem), Handle, Index
Case "cta"
ControlResize cta(pro.Witem), Handle, Index
End Select
End Sub

Private Sub SSTab1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SSTab1.Tab = 1 Then
Call Generate
ElseIf SSTab1.Tab = 3 Then
Call Generate
wb.Navigate "c:/preview.html"
End If
End Sub

Private Sub zone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
bol = True
X3 = X
Y3 = Y
pro.Wtype.Text = "Page"
pro.Witem = "N/A"
Index = "N/A"
'pro.Caption = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "caption")
pro.nam.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "name")
pro.valu.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "value")
pro.bgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "bgcolor")
pro.title.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "title")
pro.alt.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "alt")
pro.RO.Value = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "ro")
pro.border.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "border")
pro.path.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "path")
pro.onclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "click")
pro.out.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "out")
pro.over.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "over")
pro.down.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "down")
pro.fgcol.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "fgcolor")
pro.ondblclick.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "dblclick")
pro.mouve.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "move")
pro.onload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "load")
pro.onunload.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "unload")
pro.onkeydown.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keydown")
pro.onkeyup.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keyup")
pro.onkeypress.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "keypress")
pro.onselect.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "select")
pro.onfocus.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "focus")
pro.onchange.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "change")
pro.onblur.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "blur")
pro.onerror.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "error")
pro.onabord.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "abord")
pro.ondblclick.Enabled = False
pro.mouve.Enabled = False
pro.onload.Enabled = False
pro.onunload.Enabled = False
pro.onkeydown.Enabled = False
pro.onkeyup.Enabled = False
pro.onkeypress.Enabled = False
pro.onselect.Enabled = False
pro.onfocus.Enabled = False
pro.onchange.Enabled = False
pro.onblur.Enabled = False
pro.onerror.Enabled = False
pro.onabord.Enabled = False
pro.zondblclick.Enabled = pro.ondblclick.Enabled
pro.zmouve.Enabled = pro.mouve.Enabled
pro.zonload.Enabled = pro.onload.Enabled
pro.zonunload.Enabled = pro.onunload.Enabled
pro.zonkeydown.Enabled = pro.onkeydown.Enabled
pro.zonkeyup.Enabled = pro.onkeyup.Enabled
pro.zonkeypress.Enabled = pro.onkeypress.Enabled
pro.zonselect.Enabled = pro.onselect.Enabled
pro.zonfocus.Enabled = pro.onfocus.Enabled
pro.zonchange.Enabled = pro.onchange.Enabled
pro.zonblur.Enabled = pro.onblur.Enabled
pro.zonerror.Enabled = pro.onerror.Enabled
pro.zonabord.Enabled = pro.onabord.Enabled
pro.onclick.Enabled = False
pro.over.Enabled = False
pro.out.Enabled = False
pro.down.Enabled = False
pro.fgcol.Enabled = False
pro.nam.Enabled = False
pro.valu.Enabled = False
pro.bgcol.Enabled = True
pro.title.Enabled = True
pro.alt.Enabled = False
pro.RO.Enabled = False
pro.border.Enabled = False
pro.path.Enabled = False
pro.zfgcolor.Enabled = pro.fgcol.Enabled
pro.zname.Enabled = pro.nam.Enabled
pro.zvalue.Enabled = pro.valu.Enabled
pro.zbgcolor.Enabled = pro.bgcol.Enabled
pro.ztitle.Enabled = pro.title.Enabled
pro.zalt.Enabled = pro.alt.Enabled
pro.zreadonly.Enabled = pro.RO.Enabled
pro.zborder.Enabled = pro.border.Enabled
pro.zpath.Enabled = pro.path.Enabled
pro.zclick.Enabled = pro.onclick.Enabled
pro.zover.Enabled = pro.over.Enabled
pro.zout.Enabled = pro.out.Enabled
pro.zdown.Enabled = pro.down.Enabled
pro.Cheight.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "height")
pro.Cwidth.Text = getstring(HKEY_CURRENT_USER, "Software\VHTML\" & pro.Wtype.Text & Index, "width")
pro.Ctop.Text = "N/A"
pro.Cleft.Text = "N/A"
End Sub

Private Sub zone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1 = X
Y1 = Y
Xcoord.Text = X
Ycoord.Text = Y
If bol = True Then
For I = 0 To 3
Line1(I).BorderColor = zone.ForeColor
Line1(I).Visible = True
Next I
Line1(0).X1 = X3
Line1(0).Y1 = Y3
Line1(0).X2 = X3
Line1(0).Y2 = Y
Line1(1).X1 = X3
Line1(1).Y1 = Y3
Line1(1).X2 = X
Line1(1).Y2 = Y3
Line1(2).X1 = X
Line1(2).Y1 = Y3
Line1(2).X2 = X
Line1(2).Y2 = Y
Line1(3).X1 = X3
Line1(3).Y1 = Y
Line1(3).X2 = X
Line1(3).Y2 = Y
End If
If bb = True Then
rect(1).Top = Y
End If
End Sub

Private Sub zone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
X4 = X
Y4 = Y
If o(0).Value = True And bol = True Then
Call wbut
ElseIf o(1).Value = True And bol = True Then
Call wcheck
ElseIf o(2).Value = True And bol = True Then
Call wimage
ElseIf o(3).Value = True And bol = True Then
Call wlink
ElseIf o(4).Value = True And bol = True Then
Call wlist
ElseIf o(5).Value = True And bol = True Then
Call wcombo
ElseIf o(6).Value = True And bol = True Then
Call wtext
ElseIf o(7).Value = True And bol = True Then
Call wtextarea
ElseIf o(8).Value = True And bol = True Then
Call wpass
ElseIf o(9).Value = True And bol = True Then
Call whidden
ElseIf o(10).Value = True And bol = True Then
Call wline
ElseIf o(11).Value = True And bol = True Then
Call wsubmit
End If
For I = 0 To 3
Line1(I).Visible = False
Next I
bol = False
o(11).Value = True
Me.MousePointer = 0
Me.zone.MousePointer = 0
aa = False
bb = False
c = False
dd = False
ee = False
ff = False
gg = False
hh = False
End Sub

Public Sub wbut()
cbut = cbut + 1
    Load cb(cbut)
    If X4 > X3 Then cb(cbut).Left = X3
    If X4 < X3 Then cb(cbut).Left = X4
    If Y4 > Y3 Then cb(cbut).Top = Y3
    If Y4 < Y3 Then cb(cbut).Top = Y4
    If X4 > X3 Then cb(cbut).Width = X4 - X3
    If X4 < X3 Then cb(cbut).Width = X3 - X4
    If Y4 > Y3 Then cb(cbut).Height = Y4 - Y3
    If Y4 < Y3 Then cb(cbut).Height = Y3 - Y4
    cb(cbut).Caption = ""
    cb(cbut).Visible = True
End Sub

Public Sub wcheck()
ccheck = ccheck + 1
    Load cc(ccheck)
    If X4 > X3 Then cc(ccheck).Left = X3
    If X4 < X3 Then cc(ccheck).Left = X4
    If Y4 > Y3 Then cc(ccheck).Top = Y3
    If Y4 < Y3 Then cc(ccheck).Top = Y4
    If X4 > X3 Then cc(ccheck).Width = 13
    If X4 < X3 Then cc(ccheck).Width = 13
    If Y4 > Y3 Then cc(ccheck).Height = 13
    If Y4 < Y3 Then cc(ccheck).Height = 13
    cc(ccheck).Visible = True
End Sub

Public Sub wimage()
cimg = cimg + 1
    Load ci(cimg)
    If X4 > X3 Then ci(cimg).Left = X3
    If X4 < X3 Then ci(cimg).Left = X4
    If Y4 > Y3 Then ci(cimg).Top = Y3
    If Y4 < Y3 Then ci(cimg).Top = Y4
    If X4 > X3 Then ci(cimg).Width = X4 - X3
    If X4 < X3 Then ci(cimg).Width = X3 - X4
    If Y4 > Y3 Then ci(cimg).Height = Y4 - Y3
    If Y4 < Y3 Then ci(cimg).Height = Y3 - Y4
    ci(cimg).Visible = True
End Sub

Public Sub wlink()
clink = clink + 1
    Load cli(clink)
    If X4 > X3 Then cli(clink).Left = X3
    If X4 < X3 Then cli(clink).Left = X4
    If Y4 > Y3 Then cli(clink).Top = Y3
    If Y4 < Y3 Then cli(clink).Top = Y4
    If X4 > X3 Then cli(clink).Width = X4 - X3
    If X4 < X3 Then cli(clink).Width = X3 - X4
    If Y4 > Y3 Then cli(clink).Height = Y4 - Y3
    If Y4 < Y3 Then cli(clink).Height = Y3 - Y4
    cli(clink).Caption = "New link !"
    cli(clink).Visible = True
End Sub

Public Sub wlist()
clistB = clistB + 1
    Load clist(clistB)
    If X4 > X3 Then clist(clistB).Left = X3
    If X4 < X3 Then clist(clistB).Left = X4
    If Y4 > Y3 Then clist(clistB).Top = Y3
    If Y4 < Y3 Then clist(clistB).Top = Y4
    If X4 > X3 Then clist(clistB).Width = X4 - X3
    If X4 < X3 Then clist(clistB).Width = X3 - X4
    If Y4 > Y3 Then clist(clistB).Height = Y4 - Y3
    If Y4 < Y3 Then clist(clistB).Height = Y3 - Y4
    clist(clistB).AddItem "New List !"
    clist(clistB).Visible = True
End Sub

Public Sub wcombo()
ccomboB = ccomboB + 1
    Load ccombo(ccomboB)
    If X4 > X3 Then ccombo(ccomboB).Left = X3
    If X4 < X3 Then ccombo(ccomboB).Left = X4
    If Y4 > Y3 Then ccombo(ccomboB).Top = Y3
    If Y4 < Y3 Then ccombo(ccomboB).Top = Y4
    If X4 > X3 Then ccombo(ccomboB).Width = X4 - X3
    If X4 < X3 Then ccombo(ccomboB).Width = X3 - X4
    ccombo(ccomboB).AddItem "New Combo !"
    ccombo(ccomboB).Visible = True
End Sub

Public Sub wtext()
ctext = ctext + 1
    Load ct(ctext)
    If X4 > X3 Then ct(ctext).Left = X3
    If X4 < X3 Then ct(ctext).Left = X4
    If Y4 > Y3 Then ct(ctext).Top = Y3
    If Y4 < Y3 Then ct(ctext).Top = Y4
    If X4 > X3 Then ct(ctext).Width = X4 - X3
    If X4 < X3 Then ct(ctext).Width = X3 - X4
    ct(ctext).Text = ""
    ct(ctext).Visible = True
End Sub

Public Sub wtextarea()
ctexta = ctexta + 1
    Load cta(ctexta)
    If X4 > X3 Then cta(ctexta).Left = X3
    If X4 < X3 Then cta(ctexta).Left = X4
    If Y4 > Y3 Then cta(ctexta).Top = Y3
    If Y4 < Y3 Then cta(ctexta).Top = Y4
    If X4 > X3 Then cta(ctexta).Width = X4 - X3
    If X4 < X3 Then cta(ctexta).Width = X3 - X4
    If Y4 > Y3 Then cta(ctexta).Height = Y4 - Y3
    If Y4 < Y3 Then cta(ctexta).Height = Y3 - Y4
    cta(ctexta).Text = ""
    cta(ctexta).Visible = True
End Sub

Public Sub wpass()
cpassw = cpassw + 1
    Load cp(cpassw)
    If X4 > X3 Then cp(cpassw).Left = X3
    If X4 < X3 Then cp(cpassw).Left = X4
    If Y4 > Y3 Then cp(cpassw).Top = Y3
    If Y4 < Y3 Then cp(cpassw).Top = Y4
    If X4 > X3 Then cp(cpassw).Width = X4 - X3
    If X4 < X3 Then cp(cpassw).Width = X3 - X4
    cp(cpassw).Text = ""
    cp(cpassw).Visible = True
End Sub

Public Sub whidden()
chiddenf = chiddenf + 1
    Load ch(chiddenf)
    If X4 > X3 Then ch(chiddenf).Left = X3
    If X4 < X3 Then ch(chiddenf).Left = X4
    If Y4 > Y3 Then ch(chiddenf).Top = Y3
    If Y4 < Y3 Then ch(chiddenf).Top = Y4
    If X4 > X3 Then ch(chiddenf).Width = X4 - X3
    If X4 < X3 Then ch(chiddenf).Width = X3 - X4
    ch(chiddenf).Text = ""
    ch(chiddenf).Visible = True
End Sub

Public Sub wline()
cline = cline + 1
    Load cl(cline)
    cl(cline).X1 = X3
    cl(cline).X2 = X4
    cl(cline).Y1 = Y3
    cl(cline).Y2 = Y3
    cl(cline).Visible = True
End Sub

Public Sub wsubmit()

End Sub

Public Sub resetREGEDIT()
Call DeleteKey(HKEY_CURRENT_USER, "Software\VHTML")
End Sub

Public Sub Verif_Func_Name()
For I = 0 To Form1.listfunc.ListCount
If fname.Text = Left(Form1.listfunc.List(I), Len(fname.Text)) Then
GoTo er1
End If
If I = Form1.listfunc.ListCount Then
js.Text = js.Text & space.Text & _
"function " & fname.Text & "() {" & space.Text & _
space.Text & _
"//tapez ici le code de la fonction " & fname.Text & "." & space.Text & _
space.Text & _
"}" & space.Text
pro.onclick.AddItem fname.Text & "()"
pro.ondblclick.AddItem fname.Text & "()"
pro.over.AddItem fname.Text & "()"
pro.down.AddItem fname.Text & "()"
pro.out.AddItem fname.Text & "()"
pro.mouve.AddItem fname.Text & "()"
pro.onload.AddItem fname.Text & "()"
pro.onunload.AddItem fname.Text & "()"
pro.onkeydown.AddItem fname.Text & "()"
pro.onkeyup.AddItem fname.Text & "()"
pro.onkeypress.AddItem fname.Text & "()"
pro.onselect.AddItem fname.Text & "()"
pro.onfocus.AddItem fname.Text & "()"
pro.onchange.AddItem fname.Text & "()"
pro.onblur.AddItem fname.Text & "()"
pro.onerror.AddItem fname.Text & "()"
pro.onabord.AddItem fname.Text & "()"
Form1.listfunc.AddItem fname.Text & "()"
End If
Next I
Exit Sub
er1:
MsgBox "error: fonction deja existante.", vbCritical, "error!"
Exit Sub
End Sub


