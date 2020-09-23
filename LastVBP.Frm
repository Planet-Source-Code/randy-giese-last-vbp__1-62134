VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLastVBP 
   Caption         =   "Last saved:"
   ClientHeight    =   5130
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7620
   ForeColor       =   &H00000000&
   Icon            =   "LastVBP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrLastVBP 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H00C8D0D4&
      ForeColor       =   &H00000000&
      Height          =   1080
      Left            =   90
      ScaleHeight     =   1020
      ScaleWidth      =   7380
      TabIndex        =   9
      Top             =   3945
      Visible         =   0   'False
      Width           =   7440
      Begin VB.CommandButton cmdRestart 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Restart"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5625
         TabIndex        =   10
         Top             =   90
         Width           =   1500
      End
      Begin RichTextLib.RichTextBox rtbDisplay 
         Height          =   600
         Left            =   90
         TabIndex        =   11
         Top             =   390
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   1058
         _Version        =   393217
         BackColor       =   -2147483648
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"LastVBP.frx":0CCA
      End
      Begin VB.Label lblDisplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Date          Time        File Name  and  Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   90
         TabIndex        =   12
         Top             =   90
         Width           =   5685
      End
   End
   Begin VB.PictureBox picLastVBP 
      BackColor       =   &H00C8D0D4&
      ForeColor       =   &H00000000&
      Height          =   3735
      Left            =   90
      ScaleHeight     =   3675
      ScaleWidth      =   7380
      TabIndex        =   0
      Top             =   90
      Width           =   7440
      Begin VB.PictureBox picFileType 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   3900
         ScaleHeight     =   345
         ScaleWidth      =   3240
         TabIndex        =   6
         Top             =   720
         Width           =   3300
         Begin VB.TextBox txtFileType 
            BackColor       =   &H00D8E9EC&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2220
            TabIndex        =   7
            Text            =   "*.Vbp"
            Top             =   30
            Width           =   1020
         End
         Begin VB.Label lblFileType 
            AutoSize        =   -1  'True
            BackColor       =   &H00D8E9EC&
            Caption         =   "Current File Type:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   60
            TabIndex        =   8
            Top             =   30
            Width           =   1860
         End
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   5100
         TabIndex        =   4
         Top             =   0
         Width           =   2100
      End
      Begin VB.FileListBox filLastVBP 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2370
         Hidden          =   -1  'True
         Left            =   3900
         MultiSelect     =   2  'Extended
         Pattern         =   "*.Vbp"
         TabIndex        =   3
         Top             =   1155
         Width           =   3300
      End
      Begin VB.DirListBox dirLastVBP 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2295
         Left            =   120
         TabIndex        =   2
         Top             =   1155
         Width           =   3600
      End
      Begin VB.DriveListBox drvLastVBP 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   3600
      End
      Begin VB.Label lblLastVBP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "Sort VBPs in the ""Last Saved"" order."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   6000
      End
   End
End
Attribute VB_Name = "frmLastVBP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Program Name:  LastVBP                                       ****
'   ****                                                                 ****
'   ****    Written by:    Randy Giese                                   ****
'   ****    Written at:    www.RandyGrams.com                            ****
'   ****    Written on:    August 09, 2005                               ****
'   ****    E-Mail me:     WebMaster@RandyGrams.com                      ****
'   ****                                                                 ****
'   ****    Purpose:                                                     ****
'   ****                                                                 ****
'   ****    I originally wrote this to scan my directories and           ****
'   ****    sub-directories, then list the most recently saved VB        ****
'   ****    Projects.  I know that when I open VB, I can choose          ****
'   ****    'Recent' and pick from that list.  However, when I           ****
'   ****    download and try many programs from PSC, plus a few more     ****
'   ****    of my own, that list quickly becomes quite cumbersome and    ****
'   ****    it really doesn't tell me which ones I have edited.  It      ****
'   ****    also can list programs that have already been moved or       ****
'   ****    deleted.  So I wrote this.  But please read on to see        ****
'   ****    what else I discovered.                                      ****
'   ****                                                                 ****
'   ****    I had planned to just list the 10 most recent projects,      ****
'   ****    but hadn't gotten around to that yet when I discovered       ****
'   ****    something interesting.  Besides listing the most recently    ****
'   ****    used projects, it also lists the oldest projects             ****
'   ****    (actually, it lists ALL projects in ALL sub-directories      ****
'   ****    that you choose).  So by checking the bottom of the list,    ****
'   ****    you may find some oldies that either need your attention     ****
'   ****    or may not be of any use anymore.  You may also find some    ****
'   ****    oldies that were stuck away in a sub-directory and           ****
'   ****    forgotten about.                                             ****
'   ****                                                                 ****
'   ****    Then I thought that this program might be useful for         ****
'   ****    finding other File Types too.  Such as *.Frm, *.Bas, etc,    ****
'   ****    so I added  the 'txtFileType' Text Box.  You may enter       ****
'   ****    any File Type you like and scan for it.  The program will    ****
'   ****    list All files matching your criteria, sorted in Date        ****
'   ****    order (most recent to oldest).                               ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    October 18, 2005 - Version 1.2.4                             ****
'   ****                                                                 ****
'   ****    This had some cosmetic changes.  I added a label,            ****
'   ****    "lblHeader", at the top of the Text Box and put the          ****
'   ****    "Headings" there.  This way the "Headings" stay put if       ****
'   ****    you have to scroll the screen.  I also downsized the         ****
'   ****    "ReStart" button and moved it onto "lblHeader".              ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    Inputs:                                                      ****
'   ****    You may change the File Type or use the default (*.Vbp).     ****
'   ****                                                                 ****
'   ****    Returns:                                                     ****
'   ****    A sorted list matching your criteria.                        ****
'   ****                                                                 ****
'   ****    Side Effects:                                                ****
'   ****    None.                                                        ****
'   ****                                                                 ****
'   ****    Assumptions:                                                 ****
'   ****    None.                                                        ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

Option Explicit
Option Base 1

Private Const ABM_GETSTATE         As Long = &H4
Private Const ABM_GETTASKBARPOS    As Long = &H5
Private Const ABS_ALWAYSONTOP      As Long = &H2
Private Const ABS_AUTOHIDE         As Long = &H1
Private Const CCHILDREN_TITLEBAR   As Long = 5
Private Const FILE_SHARE_READ      As Long = &H1
Private Const FILE_SHARE_WRITE     As Long = &H2
Private Const GENERIC_WRITE        As Long = &H40000000
Private Const MF_SEPARATOR         As Long = &H800&
Private Const MF_STRING            As Long = &H0&
Private Const OPEN_EXISTING        As Long = 3
Private Const SM_CYCAPTION         As Long = 4             'Height of windows caption
Private Const SM_CYMENU            As Long = 15            'Height of menu
Private Const SM_CYSCREEN          As Long = 1             'Height of Screen
Private Const TPM_LEFTALIGN        As Long = &H0&
Private Const TPM_RETURNCMD        As Long = &H100&
Private Const TPM_RIGHTBUTTON      As Long = &H2&
Private Const VK_RBUTTON           As Long = &H2

Private Type POINTAPI
    lngX                           As Long
    lngY                           As Long
End Type

Private Type RECT
    Left                           As Long
    Top                            As Long
    Right                          As Long
    Bottom                         As Long
End Type

Private Type APPBARDATA
    cbSize                         As Long
    HWnd                           As Long
    uCallbackMessage               As Long
    uEdge                          As Long
    rc                             As RECT
    lParam                         As Long  'message specific
End Type

Private Type FILETIME
    dwLowDateTime                  As Long
    dwHighDateTime                 As Long
End Type

Private Type SHFILEOPSTRUCT
    HWnd                           As Long
    wFunc                          As Long
    pFrom                          As String
    pTo                            As String
    fFlags                         As Integer
    fAborted                       As Boolean
    hNameMaps                      As Long
    sProgress                      As String
End Type

Private Type SYSTEMTIME
    wYear                          As Integer
    wMonth                         As Integer
    wDayOfWeek                     As Integer
    wDay                           As Integer
    wHour                          As Integer
    wMinute                        As Integer
    wSecond                        As Integer
    wMilliseconds                  As Integer
End Type

Private Type TITLEBARINFO
    cbSize                         As Long
    rcTitleBar                     As RECT
    rgstate(CCHILDREN_TITLEBAR)    As Long
End Type

Private ABD                        As APPBARDATA
Private arrFileInfo()              As String
Private lngFileCounter             As Long
Private lngLoop1                   As Long
Private lngState                   As Long
Private lngTop                     As Long
Private strFirstFolder             As String

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Function GetMenu Lib "user32" (ByVal HWnd As Long) As Long
Private Declare Function GetMenuItemRect Lib "user32" (ByVal HWnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

'These are the Menu Functions.
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal HWnd As Long, ByVal lptpm As Any) As Long

Private Sub cmdRestart_Click()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    cmdRestart came from a suggestion from 'Mark Goodfellow'.    ****
'   ****                                                                 ****
'   ****    Thank you for the idea Mark!                                 ****
'   ****                                                                 ****
'   *************************************************************************
'
    cmdStart.Enabled = True
    picLastVBP.Visible = True
    picDisplay.Visible = False
    ReDim arrFileInfo(1)

End Sub

Private Sub cmdStart_Click()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Make sure that the File List Box, the Directory List Box     ****
'   ****    and the Drive List Box are in sync.                          ****
'   ****                                                                 ****
'   ****    Then set a couple default values and start the party.        ****
'   ****                                                                 ****
'   *************************************************************************
'
    cmdStart.Enabled = False
    ChDrive drvLastVBP.Drive
    ChDir dirLastVBP.Path
    strFirstFolder = CurDir
    filLastVBP.Pattern = txtFileType.Text

    Scan
    Sort

End Sub

Private Sub dirLastVBP_Change()

    ChDir dirLastVBP.Path
    filLastVBP.Path = dirLastVBP.Path

End Sub

Private Sub drvLastVBP_Change()

    With dirLastVBP
        ChDrive .Path
        .Path = drvLastVBP.Drive
        .Refresh
    End With  'dirLastVBP

    Focus

End Sub

Private Sub FileDate()

Dim FT1                            As FILETIME
Dim FT2                            As FILETIME
Dim lngHandle                      As Long
Dim strFile                        As String
Dim strPath                        As String
Dim SysTime                        As SYSTEMTIME

    strPath = UCase$(Left$(filLastVBP.Path, 1) & Trim$(Mid$(filLastVBP.Path, 2, 999)))
    If Right$(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    strFile = strPath & filLastVBP.List(lngLoop1)

    lngHandle = CreateFile(strFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    GetFileTime lngHandle, FT1, FT1, FT2
    FileTimeToLocalFileTime FT2, FT1
    FileTimeToSystemTime FT1, SysTime
    CloseHandle lngHandle
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    The file's (Last Saved) Date & Time is stored in the         ****
'   ****    following format: YYYY-MM-DD HH:MM:SS                        ****
'   ****                                                                 ****
'   *************************************************************************
'
    lngFileCounter = lngFileCounter + 1
    ReDim Preserve arrFileInfo(lngFileCounter)
    arrFileInfo(lngFileCounter) = LTrim$(Str$(SysTime.wYear)) & "-" & Right$("0" & LTrim$(Str$(SysTime.wMonth)), 2) & "-" & Right$("0" & LTrim$(Str$(SysTime.wDay)), 2) & "     " & Right$("0" & LTrim$(Str$(SysTime.wHour)), 2) & ":" & Right$("0" & LTrim$(Str$(SysTime.wMinute)), 2) & ":" & Right$("0" & LTrim$(Str$(SysTime.wSecond)), 2) & "     " & strFile

End Sub

Private Sub Focus()

    If cmdStart.Visible Then
        cmdStart.SetFocus
    End If

End Sub

Private Sub Form_Load()

    ChDrive App.Path
    ChDir App.Path
    ScrnInfo

End Sub

Private Sub Form_Resize()

Dim abdChanged                     As APPBARDATA
Dim lngAuto                        As Long
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Set up the controls on the main form.                        ****
'   ****                                                                 ****
'   *************************************************************************
'
    With picDisplay
        .Move 90, 90, Screen.Width - 180, Screen.Height - 630
        lblDisplay.Move 90, 90, .Width - 240
        rtbDisplay.Move lblDisplay.Left, lblDisplay.Height, lblDisplay.Width, .Height - lblDisplay.Top - lblDisplay.Height - 60
        cmdRestart.Move lblDisplay.Left + lblDisplay.Width - cmdRestart.Width, lblDisplay.Top
    End With  'picDisplay

    With picLastVBP
'       .Move 300, 300, Screen.Width - 600, Screen.Height - 900
        .Move 120, 120, Screen.Width - 240, Screen.Height - 660
        lblLastVBP.Move (.Width - lblLastVBP.Width) \ 2, 0
    End With  'picLastVBP

    With drvLastVBP
        .Move 120, lblLastVBP.Top + lblLastVBP.Height, rtbDisplay.Width \ 2
        picFileType.Move .Left + .Width + 240, .Top, .Width - 360, .Height
        dirLastVBP.Move .Left, .Top + .Height + 60, .Width, picLastVBP.Height - .Top - .Height - 120
    End With  'drvLastVBP

    With picFileType
        filLastVBP.Move .Left, dirLastVBP.Top, .Width, dirLastVBP.Height
        cmdStart.Move .Left + .Width - 2100, 0, 2100, lblLastVBP.Height
    End With  'picFileType

'   Minimize the Height difference between the Directory ListBox and the File ListBox.
    lngAuto = (dirLastVBP.Height - filLastVBP.Height) \ 2
    picFileType.Top = picFileType.Top + lngAuto
    filLastVBP.Top = filLastVBP.Top + lngAuto

'   Locate the Cursor in the "Start" button.
    With cmdStart
        SetCursorPos (Me.Left + picLastVBP.Left + .Left + (.Width * 4 \ 5)) \ Screen.TwipsPerPixelX, ((Me.Top + picLastVBP.Top + .Top + (.Height \ 3)) \ Screen.TwipsPerPixelY) + lngTop + 3
        .SetFocus
    End With  'cmdLastVBP

'   Initialize the appbardata structure
    abdChanged.cbSize = Len(abdChanged)
'   See if the taskbar's state has changed (Did the AutoHide feature change?)
    lngAuto = SHAppBarMessage(ABM_GETSTATE, abdChanged)

    If lngAuto <> lngState Then
        ScrnInfo
    End If

'   See if the taskbar's position has changed
    SHAppBarMessage ABM_GETTASKBARPOS, abdChanged

    If abdChanged.rc.Left <> ABD.rc.Left Then
        ScrnInfo
    End If

    If abdChanged.rc.Top <> ABD.rc.Top Then
        ScrnInfo
    End If

    If abdChanged.rc.Right <> ABD.rc.Right Then
        ScrnInfo
    End If

    If abdChanged.rc.Bottom <> ABD.rc.Bottom Then
        ScrnInfo
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Quit

End Sub

Private Sub lblFileType_Click()

    txtFileType.SetFocus

End Sub

Private Sub MoveUp()

    With dirLastVBP
        If .List(-1) <> strFirstFolder Then
            ChDir .List(-2)
            .Path = .List(-2)
        End If
    End With  'dirLastVBP

End Sub

Private Sub picFileType_Click()

    txtFileType.SetFocus

End Sub

Private Sub Quit()

    On Error Resume Next
'   Unload the form
    Unload frmLastVBP

    End
    On Error GoTo 0

End Sub

Private Sub Scan()

Dim lngDirCounter                  As Long
Dim lngLoop2                       As Long
Dim lngNumberOfFiles               As Long
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    If there are any files in this Directory that match the      ****
'   ****    criteria in the 'txtFileType' Text Box, process them         ****
'   ****    through 'FileDate' to retrieve the file's (Last Saved)       ****
'   ****    Date & Time.                                                 ****
'   ****                                                                 ****
'   *************************************************************************
'
    lngNumberOfFiles = filLastVBP.ListCount
    If lngNumberOfFiles > 0 Then
        For lngLoop1 = 0 To lngNumberOfFiles - 1
            FileDate
        Next lngLoop1
    End If

    lngDirCounter = dirLastVBP.ListCount
    If lngDirCounter > 0 Then
        For lngLoop2 = 0 To lngDirCounter - 1
            With dirLastVBP
                ChDir .List(lngLoop2)
                .Path = .List(lngLoop2)
                filLastVBP.Path = .List(lngLoop2)
            End With         'dirLastVBP

            filLastVBP.Refresh
            Scan
        Next lngLoop2
    End If

    filLastVBP.Path = dirLastVBP.Path
    MoveUp

End Sub

Private Sub ScrnInfo()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    ScrnInfo is used to get information about some specific      ****
'   ****    items of the "Frm":                                          ****
'   ****                                                                 ****
'   ****    lngTitleBar gets the Height of the TitleBar, if there is     ****
'   ****    one.                                                         ****
'   ****                                                                 ****
'   ****    lngMenu gets the Height of the Menu, if there is one.        ****
'   ****                                                                 ****
'   ****    lngTaskBar gets the Height of the TaskBar if it's            ****
'   ****    visible, but only if it's at the Top of the screen.  Any     ****
'   ****    other location won't matter to this application.             ****
'   ****                                                                 ****
'   ****    These three figures (lngTop) are then used to help           ****
'   ****    determine the location of the Mouse Pointer.                 ****
'   ****    Note:  the Mouse Pointer is located on the "Start"           ****
'   ****    command, but does not obscure the caption.                   ****
'   ****                                                                 ****
'   ****    I know I could have just entered a number and set the        ****
'   ****    Mouse Pointer there, but it may not be accurate with         ****
'   ****    other screen resolutions, font sizes, etc.                   ****
'   ****                                                                 ****
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

Dim lngMenu                        As Long
Dim lngReturn                      As Long
Dim lngTaskBar                     As Long
Dim lngTitleBar                    As Long
Dim recRect                        As RECT
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    These are the statements that are used to get the various    ****
'   ****    object heights.  Some of these items will report the         ****
'   ****    particular Height even if it is not being used.  An          ****
'   ****    example of this is the Menu Height.  Therefore, they must    ****
'   ****    be checked to see if they are actually being used.           ****
'   ****                                                                 ****
'   ****    Please see the coding comments for the individual item       ****
'   ****    for more information.                                        ****
'   ****                                                                 ****
'   *************************************************************************
'
'   Initialize the appbardata structure
    ABD.cbSize = Len(ABD)
'   Get the taskbar's position
    SHAppBarMessage ABM_GETTASKBARPOS, ABD
'   Get the taskbar's state
    lngState = SHAppBarMessage(ABM_GETSTATE, ABD)
'   Get the Screen Height
    GetSystemMetrics SM_CYSCREEN
'   Get the TitleBar Height
    lngTitleBar = GetSystemMetrics(SM_CYCAPTION)
'   Get the MenuBar Height
    lngMenu = GetSystemMetrics(SM_CYMENU)
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Check to see if there is a Menu Rectangle.                   ****
'   ****                                                                 ****
'   ****    lngReturn will return a value of '0' (zero), if there is     ****
'   ****    no Menu or if the Menu is invisible.                         ****
'   ****                                                                 ****
'   *************************************************************************
'
    lngReturn = GetMenuItemRect(Me.HWnd, GetMenu(Me.HWnd), 0, recRect)

    If lngReturn = 0 Then
        lngMenu = 0
    End If

'   Check the Status of the TaskBar
'   If the AUTOHIDE feature is on, ignore the TaskBar size
    Select Case lngState
    Case ABS_AUTOHIDE
        lngTaskBar = 0

    Case ABS_ALWAYSONTOP
'       If ABD.rc.Top > 0 then the TaskBar is on the Bottom
        If ABD.rc.Top > 0 Then
            lngTaskBar = 0
            GoTo GetOut
        End If

'       If ABD.rc.Bottom < Screen.Height then the TaskBar is on the Top
        If ABD.rc.Bottom < Screen.Height \ Screen.TwipsPerPixelY Then
            lngTaskBar = ABD.rc.Bottom - ABD.rc.Top
            GoTo GetOut
        End If

    Case Else
        lngTaskBar = 0
    End Select               'lngState

GetOut:
'   Get the Top of the Usable Screen area
    lngTop = lngTitleBar + lngMenu + lngTaskBar
    tmrLastVBP.Enabled = True

End Sub

Private Sub ShowMenu()

Dim lngMainMenu                    As Long
Dim lngTPM                         As Long
Dim pntXorY                        As POINTAPI
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    This PopUp Menu is from the "PopUp Menu Tutorial" named      ****
'   ****    Cre8Popup, which can be found at:                            ****
'   ****                                                                 ****
'   ****    http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=68311&lngWId=1
'   ****                                                                 ****
'   ****    There is much more documentation there.                      ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
        lngMainMenu = CreatePopupMenu()
        GetCursorPos pntXorY

        AppendMenu lngMainMenu, MF_STRING, 1, "About"
        AppendMenu lngMainMenu, MF_STRING, 2, "Help"
        AppendMenu lngMainMenu, MF_SEPARATOR, 3, ByVal 0&
        AppendMenu lngMainMenu, MF_STRING, 4, "Delete the Text File!"
        AppendMenu lngMainMenu, MF_SEPARATOR, 5, ByVal 0&
        AppendMenu lngMainMenu, MF_STRING, 6, "Quit"

        lngTPM = TrackPopupMenuEx(lngMainMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, pntXorY.lngX, pntXorY.lngY, Me.HWnd, ByVal 0&)
        DestroyMenu lngMainMenu

        Select Case lngTPM
        Case 1
            RG_About.Show vbModal, Me
        Case 2
'           The name of the Help file.
            RG_Help.Help_File = App.Path & "\LastVBP.Html"
'           Display the Help page
            RG_Help.Show vbModal, Me
        Case 4
            Kill App.Path & "\Ra.ndy"
        Case 6
            Quit
        End Select           'lngTPM

        SetCursorPos pntXorY.lngX, pntXorY.lngY

End Sub

Private Sub Sort()

Dim blnSorted                      As Boolean
Dim lngActual                      As Long
Dim lngLoop2                       As Long
Dim strTemp                        As String
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Sort the File Names in Date order.                           ****
'   ****                                                                 ****
'   *************************************************************************
'
    ReDim Preserve arrFileInfo(lngFileCounter + 1)

    Do While Not blnSorted
        blnSorted = True
        For lngLoop2 = 1 To lngFileCounter
            If Left$(arrFileInfo(lngLoop2), 23) < Left$(arrFileInfo(lngLoop2 + 1), 23) Then
                strTemp = arrFileInfo(lngLoop2)
                arrFileInfo(lngLoop2) = arrFileInfo(lngLoop2 + 1)
                arrFileInfo(lngLoop2 + 1) = strTemp

                If GetInputState Then
                    DoEvents 'Yield for other processing.
                End If

                blnSorted = False
            End If
        Next lngLoop2
    Loop

    picLastVBP.Visible = False
    picDisplay.Visible = True

    Open App.Path & "\Ra.ndy" For Output As #1
    lngActual = lngFileCounter

    For lngLoop2 = 1 To lngFileCounter
        If Trim$(Mid$(arrFileInfo(lngLoop2), 25, 999)) = Trim$(Mid$(arrFileInfo(lngLoop2 + 1), 25, 999)) Then
            If Left$(arrFileInfo(lngLoop2), 23) = Left$(arrFileInfo(lngLoop2 + 1), 23) Then
                lngActual = lngActual - 1
                GoTo Nxt
            End If
        End If

        Print #1, " " & arrFileInfo(lngLoop2)
Nxt:
    Next lngLoop2

    Print #1,
    Print #1, " There are " & lngActual & " " & filLastVBP.Pattern & " files under the " & strFirstFolder & " folder."

    Close
    rtbDisplay.FileName = App.Path & "\Ra.ndy"

End Sub

Private Sub tmrLastVBP_Timer()

'   If the Right Mouse Button was pressed, go to (ShowMenu).
    If GetAsyncKeyState(VK_RBUTTON) <> 0 Then
        ShowMenu
    End If

End Sub

Private Sub txtFileType_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
'       The user has pressed the 'Enter' key.  Move the Focus back to the Start button.
        Focus
    End If

End Sub

':)Code Fixer V4.0.35 (Tuesday, 29 May 2007 11:49:00) 176 + 523 = 699 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 133302322223333232|033322232223332222222222222222222|1113222|2221222|333332222233|111111111111|1122222222222|333333|
