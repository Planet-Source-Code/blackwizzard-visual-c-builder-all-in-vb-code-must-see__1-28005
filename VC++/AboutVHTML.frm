VERSION 5.00
Begin VB.Form AboutVHTML 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "AboutVHTML.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "AboutVHTML.frx":058A
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "AboutVHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
    Dim lY As Long
    Dim lScaleHeight As Long
    Dim lScaleWidth As Long
    ScaleMode = vbPixels
    lScaleHeight = ScaleHeight
    lScaleWidth = ScaleWidth
    DrawStyle = vbInvisible
    FillStyle = vbFSSolid
    For lY = 0 To lScaleHeight
        FillColor = RGB(0, 0, 255 - (lY * 255) \ lScaleHeight)
        Line (-1, lY - 1)-(lScaleWidth, lY + 1), , B
    Next lY
End Sub
