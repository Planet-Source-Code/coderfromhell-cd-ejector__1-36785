VERSION 5.00
Begin VB.Form frmTray 
   Caption         =   "CD Ejector"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.Image TrayI 
      Height          =   240
      Left            =   1080
      Picture         =   "frmTray.frx":0442
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image MenuPicContainer 
      Height          =   240
      Index           =   3
      Left            =   1440
      Picture         =   "frmTray.frx":058C
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image MenuPicContainer 
      Height          =   240
      Index           =   2
      Left            =   1200
      Picture         =   "frmTray.frx":068E
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image MenuPicContainer 
      Height          =   240
      Index           =   1
      Left            =   960
      Picture         =   "frmTray.frx":0790
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image MenuPicContainer 
      Height          =   240
      Index           =   0
      Left            =   720
      Picture         =   "frmTray.frx":0892
      Top             =   1080
      Width           =   240
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type IconeTray
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Dim IconeT As IconeTray

Private Type DrvInfo
    DriveLetter As String
    DriveName As String
    ID As Long
    CDOpen As Boolean
End Type

Private Const AJOUT = &H0
Private Const MODIF = &H1
Private Const SUPPRIME = &H2
Private Const MOUSEMOVE = &H200
Private Const MESSAGE = &H1
Private Const Icone = &H2
Private Const TIP = &H4

Private Const DOUBLE_CLICK_GAUCHE = &H203
Private Const BOUTON_GAUCHE_POUSSE = &H201
Private Const BOUTON_GAUCHE_LEVE = &H202
Private Const DOUBLE_CLICK_DROIT = &H206
Private Const BOUTON_DROIT_POUSSE = &H204
Private Const BOUTON_DROIT_LEVE = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As IconeTray) As Boolean
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Dim APICheck1 As Boolean
Dim CDDrives() As DrvInfo



Private Sub ShowMenu()
    Dim hPopupMenu1 As Long
    Dim mii1 As MENUITEMINFO
    Dim curpos As POINT_TYPE
    Dim menusel As Long
    Dim retval As Long
    Dim I As Integer, IndexCount As Integer
    
    hPopupMenu1 = CreatePopupMenu()
    
    With mii1
    .cbSize = Len(mii1)
    .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU
    End With
    
    IndexCount = 0
    
    For I = 1 To UBound(CDDrives)
        With CDDrives(I)
            .DriveName = StrConv(Dir((.DriveLetter & ":"), vbVolume), vbProperCase)
            If .DriveName = "" Then
                .DriveName = "Compact Disc (" & .DriveLetter & ":)"
            Else
                .DriveName = .DriveName & " (" & .DriveLetter & ":)"
            End If
        End With
    Next I
    
    For I = 1 To UBound(CDDrives)
        With CDDrives(I)
            Select Case .CDOpen
                Case False
                    mii1.fType = MFT_STRING
                    mii1.fState = MFS_ENABLED Or MFS_UNCHECKED
                    mii1.wID = .ID
                    mii1.dwTypeData = .DriveName
                    mii1.cch = Len(.DriveName)
                    mii1.hSubMenu = 0
                Case True
                    mii1.fType = MFT_STRING
                    mii1.fState = MFS_ENABLED Or MFS_CHECKED
                    mii1.wID = .ID
                    mii1.dwTypeData = .DriveName
                    mii1.cch = Len(.DriveName)
                    mii1.hSubMenu = 0
                Case Else
            End Select
        End With
        retval = InsertMenuItem(hPopupMenu1, IndexCount, 1, mii1)
        IndexCount = IndexCount + 1
    Next I
    
    With mii1
    .fType = MFT_SEPARATOR
    .fState = MFS_ENABLED
    .wID = 1001
    .dwTypeData = "/separator/"
    .cch = Len("/separator/")
    .hSubMenu = 0
    End With
    retval = InsertMenuItem(hPopupMenu1, IndexCount, 1, mii1)
    
    IndexCount = IndexCount + 1
    With mii1
    .fType = MFT_STRING
    .fState = MFS_ENABLED
    .wID = 1000
    .dwTypeData = "E&xit"
    .cch = Len("E&xit")
    .hSubMenu = 0
    End With
    retval = InsertMenuItem(hPopupMenu1, IndexCount, 1, mii1)

    'Add Bitmaps
    For I = 1 To UBound(CDDrives)
        retval = SetMenuItemBitmaps(hPopupMenu1, CDDrives(I).ID, 1, MenuPicContainer(0), MenuPicContainer(1))
    Next I
    retval = SetMenuItemBitmaps(hPopupMenu1, 1002, 1, MenuPicContainer(2), MenuPicContainer(2))
    retval = SetMenuItemBitmaps(hPopupMenu1, 1000, 1, MenuPicContainer(3), MenuPicContainer(3))
    '------------------------------------------------------------
    '------------------------------------------------------------
    
    retval = GetCursorPos(curpos)
    SetForegroundWindow hWnd
    menusel = TrackPopupMenu(hPopupMenu1, TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_LEFTALIGN Or TPM_LEFTBUTTON, curpos.X, curpos.Y, 0, hWnd, 0)
    retval = DestroyMenu(hPopupMenu1)
 
    For I = 1 To UBound(CDDrives)
        With CDDrives(I)
            If menusel = .ID Then
                .CDOpen = Not .CDOpen
                If .CDOpen = True Then
                    CDEjector.openCD .DriveLetter
                Else
                    CDEjector.closeCD (.DriveLetter)
                End If
                GoTo ExitMe
            End If
        End With
    Next I
    Select Case menusel
        Case 1000 '(Exit)
            Unload Me
        Case Else
    End Select
ExitMe:
End Sub

Private Sub Form_Load()
    Hide
    APICheck1 = False
    With IconeT
        .cbSize = Len(IconeT)
        .hWnd = Me.hWnd
        .uID = 1&
        .uFlags = Icone Or TIP Or MESSAGE
        .uCallbackMessage = MOUSEMOVE
        .hIcon = Me.TrayI.Picture
        .szTip = "Whizzo CD Ejector" & Chr$(0)
    End With
    Shell_NotifyIcon AJOUT, IconeT
    'First Get Drives
    Dim DriveNum As String, DriveType As Long, retval As Long
    DriveNum = 64
    On Error Resume Next
    retval = 0
    ReDim CDDrives(retval)
    Do
        DriveNum = DriveNum + 1
        DriveType = GetDriveType(Chr$(DriveNum) & ":\")
        If DriveNum > 90 Then Exit Do
        Select Case DriveType
'            Case 0: "unknown"
'            Case 2: "remove"
'            Case 3: "fixed"
'            Case 4: "remote"
            Case 5: 'cd
                retval = retval + 1
                ReDim Preserve CDDrives(retval)
                With CDDrives(retval)
                    .DriveLetter = Chr$(DriveNum)
                    .ID = 1003 + retval
                    .DriveName = StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase)
                    If .DriveName = "" Then .DriveName = "Compact Disc (" & .DriveLetter & ":)"
                    .CDOpen = False
                End With
'            Case 6: "ram"
        End Select
    Loop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static rec As Boolean, msg As Long
    
    msg = X
    
    If rec = False Then
        rec = True
        Select Case msg
            Case BOUTON_DROIT_LEVE: ShowMenu
        End Select
        rec = False
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    IconeT.cbSize = Len(IconeT)
    IconeT.hWnd = Me.hWnd
    IconeT.uID = 1&
    Shell_NotifyIcon SUPPRIME, IconeT
End Sub

