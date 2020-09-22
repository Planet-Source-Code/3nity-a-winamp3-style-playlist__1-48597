VERSION 5.00
Object = "*\AProjectStdSS.vbp"
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   5925
      TabIndex        =   18
      Text            =   "New Title"
      Top             =   3525
      Width           =   2265
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Refresh Title"
      Height          =   315
      Left            =   4350
      TabIndex        =   17
      Top             =   3525
      Width           =   1440
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   6750
      TabIndex        =   16
      Text            =   "147"
      Top             =   3150
      Width           =   1440
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Refresh Time (only numers!!!)"
      Height          =   315
      Left            =   4350
      TabIndex        =   15
      Top             =   3150
      Width           =   2265
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   6300
      TabIndex        =   14
      Text            =   "File2"
      Top             =   2700
      Width           =   1890
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add File2 (to selected)"
      Height          =   315
      Left            =   4350
      TabIndex        =   13
      Top             =   2700
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "frmMain.frx":0000
      Left            =   4275
      List            =   "frmMain.frx":0002
      TabIndex        =   12
      Top             =   5325
      Width           =   3990
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GetSelectedData"
      Height          =   315
      Left            =   4275
      TabIndex        =   11
      Top             =   4950
      Width           =   3990
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6750
      Picture         =   "frmMain.frx":0004
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   0
      Top             =   1650
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Always Show Scroll Bar"
      Height          =   315
      Left            =   4350
      TabIndex        =   10
      Top             =   1425
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.CheckBox Check2 
      Caption         =   "AutoResize Scroller"
      Height          =   315
      Left            =   4350
      TabIndex        =   9
      Top             =   1725
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.CheckBox Check3 
      Caption         =   "ShowUP"
      Height          =   315
      Left            =   4350
      TabIndex        =   8
      Top             =   2025
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.CheckBox Check4 
      Caption         =   "ShowDown"
      Height          =   315
      Left            =   4350
      TabIndex        =   7
      Top             =   2325
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove All"
      Height          =   315
      Left            =   4275
      TabIndex        =   6
      Top             =   975
      Width           =   3990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Selected (Del)"
      Height          =   315
      Left            =   4275
      TabIndex        =   5
      Top             =   600
      Width           =   3990
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   7725
      TabIndex        =   4
      Text            =   "1:23"
      Top             =   75
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   5625
      TabIndex        =   3
      Text            =   "Item - Title"
      Top             =   75
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Item"
      Height          =   315
      Left            =   4275
      TabIndex        =   2
      Top             =   75
      Width           =   1290
   End
   Begin ProjectStdSS.stdSS stdSS1 
      Height          =   6315
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   4065
      _extentx        =   7170
      _extenty        =   11139
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is just a simple example... You can use it in
'your medaiplayers and so on...

Option Explicit

Private Sub Check1_Click()
'Shows the right scroller if true all the time else just when needed

If Check1.Value = 1 Then
    Me.stdSS1.AlwaysShowScroller True
Else
    Me.stdSS1.AlwaysShowScroller False
End If


End Sub

Private Sub Check2_Click()
'autosizes the scroller depending on the lenght of the list if true

Dim CHK2 As Boolean
Dim CHK3 As Boolean
Dim CHK4 As Boolean

If Check2.Value = 1 Then
    CHK2 = True
Else
    CHK2 = False
End If

'shows btnUp in scroll bar if true

If Check3.Value = 1 Then
    CHK3 = True
Else
    CHK3 = False
End If

'shows btnDown in scroll bar if true

If Check4.Value = 1 Then
    CHK4 = True
Else
    CHK4 = False
End If

Me.stdSS1.SetScroller stdSS1.GUI_SSOzadjeŠirina, CHK3, CHK4, CHK2, stdSS1.GUI_SSGorŠirina, stdSS1.GUI_SSGorVišina, stdSS1.GUI_SSDolŠirina, stdSS1.GUI_SSDolVišina, stdSS1.GUI_SSDrsnikMiniVišina


End Sub

Private Sub Check3_Click()
Check2_Click
End Sub

Private Sub Check4_Click()
Check2_Click
End Sub

Private Sub Command1_Click()
'add entry (filename, filenam2, title,time,timeinseconds)
Me.stdSS1.AddItem "", "", Text1, Text2, 0

End Sub

Private Sub Command2_Click()
'removes selected entry
Me.stdSS1.Remove (Me.stdSS1.Selected)

End Sub

Private Sub Command3_Click()
'clears the list
Me.stdSS1.Clear

End Sub

Private Sub Command4_Click()
'shows data from selected string u can set any index as well
List1.Clear
If Me.stdSS1.Selected > 0 Then
    Me.stdSS1.GetData Me.stdSS1.Selected
    
    List1.AddItem "Title: " & Me.stdSS1.gTitle
    List1.AddItem "File1: " & Me.stdSS1.gFileName
    List1.AddItem "File2: " & Me.stdSS1.gFileName2
    List1.AddItem "Time (seconds): " & Me.stdSS1.gTimeInSeconds
    List1.AddItem "Time: " & Me.stdSS1.gTime
    
End If

End Sub

Private Sub Command5_Click()
'adds a second filename - i use it for subtitles in my player
If Me.stdSS1.Selected > 0 Then
    Me.stdSS1.AddFileName2 Text3, Me.stdSS1.Selected
    
End If

End Sub


Private Sub Command6_Click()
'updates time in selected entry
If Me.stdSS1.Selected > 0 Then
Dim cc As String
Dim mm As Long
Dim ss As Integer
Dim seconds As Long
On Error GoTo err
seconds = Text4

mm = seconds / 60

If seconds - mm + 60 < 0 Then mm = mm - 1

ss = seconds - mm * 60

If ss < 10 Then
    cc = mm & ":0" & ss
Else
    cc = mm & ":" & ss
End If

    Me.stdSS1.RefreshTime Text4, cc, Me.stdSS1.Selected
    
End If
Exit Sub
err:
Text4 = 0

End Sub


Private Sub Command7_Click()
'updates title in selcted entry
If Me.stdSS1.Selected > 0 Then
    Me.stdSS1.RefreshTitle Text5, Me.stdSS1.Selected
    
End If

End Sub


Private Sub Form_Load()
'picture property must be set befor calling gui
Set Me.stdSS1.PictureData = Picture1

'constants must be set before calling the gui

LoadGUIConstants
Me.stdSS1.AlwaysShowScroller True

Me.stdSS1.GUI
'this must be call each time the list resizes!!!
Me.stdSS1.SetScroller stdSS1.GUI_SSOzadjeŠirina, True, True, True, stdSS1.GUI_SSGorŠirina, stdSS1.GUI_SSGorVišina, stdSS1.GUI_SSDolŠirina, stdSS1.GUI_SSDolVišina, stdSS1.GUI_SSDrsnikMiniVišina

End Sub

Private Sub LoadGUIConstants()
'constants for the gui

stdSS1.GUI_SSOzadjeŠirina = 13
stdSS1.GUI_SSOzadjeVišina = 40
stdSS1.GUI_SSOzadjeX = 26
stdSS1.GUI_SSOzadjeY = 0
stdSS1.GUI_SSOzadjeXD = 26
stdSS1.GUI_SSOzadjeYD = 0

stdSS1.GUI_SSDrsnikOzadjeŠirina = 13
stdSS1.GUI_SSDrsnikOzadjeVišina = 40
stdSS1.GUI_SSDrsnikOzadjeX = 65
stdSS1.GUI_SSDrsnikOzadjeY = 0
stdSS1.GUI_SSDrsnikOzadjeXD = 78
stdSS1.GUI_SSDrsnikOzadjeYD = 0

stdSS1.GUI_SSDrsnikGorŠirina = 13
stdSS1.GUI_SSDrsnikGorVišina = 21
stdSS1.GUI_SSDrsnikGorX = 39
stdSS1.GUI_SSDrsnikGorY = 0
stdSS1.GUI_SSDrsnikGorXD = 52
stdSS1.GUI_SSDrsnikGorYD = 0

stdSS1.GUI_SSDrsnikDolŠirina = 13
stdSS1.GUI_SSDrsnikDolVišina = 20
stdSS1.GUI_SSDrsnikDolX = 39
stdSS1.GUI_SSDrsnikDolY = 21
stdSS1.GUI_SSDrsnikDolXD = 52
stdSS1.GUI_SSDrsnikDolYD = 21

stdSS1.GUI_SSGorŠirina = 13
stdSS1.GUI_SSGorVišina = 17
stdSS1.GUI_SSGorX = 0
stdSS1.GUI_SSGorY = 0
stdSS1.GUI_SSGorXD = 13
stdSS1.GUI_SSGorYD = 0

stdSS1.GUI_SSDolŠirina = 13
stdSS1.GUI_SSDolVišina = 17
stdSS1.GUI_SSDolX = 0
stdSS1.GUI_SSDolY = 17
stdSS1.GUI_SSDolXD = 13
stdSS1.GUI_SSDolYD = 17

stdSS1.GUI_SSDrsnikMiniVišina = 41
stdSS1.GUI_SSDrsnikScale = False
stdSS1.GUI_SSVednoKaži = True


End Sub
