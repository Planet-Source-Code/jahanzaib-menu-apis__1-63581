VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Menus API"
   ClientHeight    =   5100
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8265
   Icon            =   "frmMenuAPI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   551
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "ABOUT   MENU API"
      Height          =   4935
      Left            =   7920
      TabIndex        =   61
      Top             =   120
      Width           =   255
   End
   Begin TabDlg.SSTab Tabs 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   8705
      _Version        =   393216
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "A - E Functions"
      TabPicture(0)   =   "frmMenuAPI.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "G - I Function"
      TabPicture(1)   =   "frmMenuAPI.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(2)=   "Frame8"
      Tab(1).Control(3)=   "Frame9"
      Tab(1).Control(4)=   "Frame10"
      Tab(1).Control(5)=   "Frame11"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "M - T Function"
      TabPicture(2)   =   "frmMenuAPI.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame12"
      Tab(2).Control(1)=   "Frame13"
      Tab(2).Control(2)=   "Frame14"
      Tab(2).Control(3)=   "Frame15"
      Tab(2).Control(4)=   "Frame16"
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame16 
         Caption         =   " TrackPopupMenu Function "
         Height          =   1215
         Left            =   -74880
         TabIndex        =   59
         Top             =   3480
         Width           =   5175
         Begin VB.CommandButton cmdTrackPopupMenu 
            Caption         =   "Track Popup Menu"
            Height          =   375
            Left            =   1200
            TabIndex        =   60
            Top             =   600
            Width           =   3135
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   " SetMenuItemBitmaps Function"
         Height          =   1455
         Left            =   -72240
         TabIndex        =   54
         Top             =   1920
         Width           =   4815
         Begin VB.PictureBox Picture4 
            AutoSize        =   -1  'True
            Height          =   300
            Left            =   3840
            Picture         =   "frmMenuAPI.frx":05DE
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   58
            Top             =   1080
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.PictureBox Picture3 
            AutoSize        =   -1  'True
            Height          =   300
            Left            =   240
            Picture         =   "frmMenuAPI.frx":0960
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   57
            Top             =   1080
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdSetMenuItemBitmaps 
            Caption         =   "Set Menu Item Bimaps"
            Height          =   375
            Left            =   1320
            TabIndex        =   56
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Sets bitmap for checkmark  for FILE --> SAVE item in FILE menu. Use CHECKMENUITEM function in A - E FUNCTIONS TAB to see result"
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   120
            TabIndex        =   55
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   " SetMenuDefaultItem Function "
         Height          =   1455
         Left            =   -74880
         TabIndex        =   51
         Top             =   1920
         Width           =   2535
         Begin VB.CommandButton cmdSetMenuDefaultItem 
            Caption         =   " Set Menu Default Item"
            Height          =   375
            Left            =   360
            TabIndex        =   53
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Sets FILE --> OPEN as a default item in FILE menu"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " RemoveMenu Function "
         Height          =   1335
         Left            =   -71280
         TabIndex        =   48
         Top             =   480
         Width           =   3855
         Begin VB.CommandButton cmdRemoveMenu 
            Caption         =   "Remove Menu"
            Height          =   375
            Left            =   1320
            TabIndex        =   50
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Removes RADIO 3 item in RADIO GROUP menu"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   " ModifyMenu Function"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   44
         Top             =   480
         Width           =   3495
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            Height          =   300
            Left            =   480
            Picture         =   "frmMenuAPI.frx":0CF1
            ScaleHeight     =   240
            ScaleWidth      =   270
            TabIndex        =   47
            Top             =   960
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.CommandButton cmdModifyMenu 
            Caption         =   "Modify Menu"
            Height          =   375
            Left            =   960
            TabIndex        =   46
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Replaces EDIT --> 1st Item string with a bitmap"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " InsertMenuItem Function "
         Height          =   1575
         Left            =   -72360
         TabIndex        =   40
         Top             =   3120
         Width           =   3975
         Begin VB.CommandButton cmdInsertMenuItem 
            Caption         =   "Insert Menu Item"
            Height          =   375
            Left            =   960
            TabIndex        =   43
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label14 
            Caption         =   "If menu does not appear, please put your cursor on Menu Bar"
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   3495
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Inserts HELP menu on the right side of Menu Bar"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   3735
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   " HiliteMenuItem Function "
         Height          =   1575
         Left            =   -74880
         TabIndex        =   37
         Top             =   3120
         Width           =   2415
         Begin VB.CommandButton cmdHiliteMenuItem 
            Caption         =   "Highlight"
            Height          =   435
            Left            =   360
            TabIndex        =   39
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Highlights RADIO 3 item in RADIO GROUPS menu"
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   " GetMenuItemRect Function"
         Height          =   1215
         Left            =   -72360
         TabIndex        =   34
         Top             =   1800
         Width           =   3975
         Begin VB.CommandButton cmdGetMenuItemRect 
            Caption         =   "Get Menu Item Rect"
            Height          =   375
            Left            =   960
            TabIndex        =   36
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Gets the bounding rectangle of EDIT --> 5th Item"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   315
            Width           =   3495
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   " GetMenuItemID Function "
         Height          =   1215
         Left            =   -74880
         TabIndex        =   31
         Top             =   1800
         Width           =   2415
         Begin VB.CommandButton cmdGetMenuItemId 
            Caption         =   "Get Menu Item ID"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   " Gets ID of EDIT --> 1st Item"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   " GetMenuItemCount Function"
         Height          =   1215
         Left            =   -71400
         TabIndex        =   28
         Top             =   480
         Width           =   3015
         Begin VB.CommandButton cmdGetMenuItemCount 
            Caption         =   "Get Menu Item Count"
            Height          =   375
            Left            =   480
            TabIndex        =   30
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Gets number of items in File Menu"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " GetMenuItemInfo Function"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   3375
         Begin VB.CommandButton cmdGetMnuItemInfo 
            Caption         =   "Get Menu Item Information"
            Height          =   375
            Left            =   600
            TabIndex        =   27
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Gets information about FILE --> NEW"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " CheckMenuItem Function "
         Height          =   1215
         Left            =   3240
         TabIndex        =   7
         Top             =   600
         Width           =   3855
         Begin VB.CommandButton cmdCheckMenu 
            Caption         =   "Check Menu Item"
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmMenuAPI.frx":0DA8
            Left            =   840
            List            =   "frmMenuAPI.frx":0DB2
            TabIndex        =   9
            Text            =   "File"
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frmMenuAPI.frx":0DC2
            Left            =   2880
            List            =   "frmMenuAPI.frx":0DD8
            TabIndex        =   11
            Text            =   "1"
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton cmbUncheckMenu 
            Caption         =   "Uncheck Menu item"
            Height          =   375
            Left            =   2040
            TabIndex        =   13
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Menu :"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   390
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Menu Item :"
            Height          =   195
            Left            =   1920
            TabIndex        =   10
            Top             =   390
            Width           =   840
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " CheckMenuRadioItem Function"
         Height          =   2175
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   3135
         Begin VB.CommandButton cmdCheckMenuRadio 
            Caption         =   "Check Menu Radio Items"
            Height          =   495
            Left            =   480
            TabIndex        =   16
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMenuAPI.frx":0DEE
            ForeColor       =   &H00FF0000&
            Height          =   975
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " DeleteMenu Functions "
         Height          =   1215
         Left            =   3480
         TabIndex        =   17
         Top             =   1920
         Width           =   3135
         Begin VB.CommandButton cmdDelMnu 
            Caption         =   "Delete Menu"
            Height          =   375
            Left            =   600
            TabIndex        =   19
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Delets first menu in Edit Menu"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " EnableMenuItem Function"
         Height          =   1455
         Left            =   3480
         TabIndex        =   20
         Top             =   3240
         Width           =   3135
         Begin VB.CommandButton cmdGrayedMenuItem 
            Caption         =   "Grayed"
            Height          =   375
            Left            =   2160
            TabIndex        =   24
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmdDisableMenuItem 
            Caption         =   "Disable"
            Height          =   375
            Left            =   1200
            TabIndex        =   23
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmdEnableMenuItem 
            Caption         =   "Enable"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Enables or disables 3rd and 4th items in EDIT menu"
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Append Menu Function "
         Height          =   1815
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   2895
         Begin VB.PictureBox Picture1 
            AutoSize        =   -1  'True
            Height          =   300
            Left            =   240
            Picture         =   "frmMenuAPI.frx":0EA3
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   6
            Top             =   1440
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.ComboBox cmbAppendMenuOptions 
            Height          =   315
            ItemData        =   "frmMenuAPI.frx":0F0D
            Left            =   720
            List            =   "frmMenuAPI.frx":0F17
            TabIndex        =   4
            Text            =   "Bitmap"
            Top             =   960
            Width           =   2055
         End
         Begin VB.CommandButton cmdAppendMenu 
            Caption         =   "Append Menu"
            Height          =   375
            Left            =   720
            TabIndex        =   5
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Appends a new menu in FILE menu"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Insert :"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   990
            Width           =   480
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuRadioGroup 
      Caption         =   "&Radio Groups"
      Begin VB.Menu mnuRadio1 
         Caption         =   "Radio &1"
      End
      Begin VB.Menu mnuRadio2 
         Caption         =   "Radio &2"
      End
      Begin VB.Menu mnuRadio3 
         Caption         =   "Radio &3"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim hMenu, hSubMenuFile, hSubMenuEdit, hSubMenuRadio
Dim funcResult, mainMenu, menuItem
Private Function refreshCombos()
If Combo1.Text = "Edit" Then
With Combo2
    .Clear
    .Text = "1"
    .AddItem "1"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
End With
Else
With Combo2
    .Clear
    .Text = "1"
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
End With
End If
End Function

Private Function determineCombosValue()
If Combo1.Text = "File" Then
mainMenu = hSubMenuFile
Else
mainMenu = hSubMenuEdit
End If
menuItem = CInt(Combo2.Text) - 1
End Function
Private Sub cmbUncheckMenu_Click()
determineCombosValue
funcResult = CheckMenuItem(mainMenu, menuItem, MF_BYPOSITION Or MF_UNCHECKED)
If funcResult = -1 Then
Print GetLastErrorString
End If
End Sub

Private Sub cmdAbout_Click()
frmAbout.Show 1
End Sub

Private Sub cmdAppendMenu_Click()
Dim cmbOpt, flag, newItem
cmbOpt = cmbAppendMenuOptions.Text
If cmbOpt = "Bitmap" Then
funcResult = AppendMenu(hSubMenuFile, MF_BITMAP, ID_NEWITEM, Picture1.Picture.Handle)
Else
funcResult = AppendMenu(hSubMenuFile, MF_STRING, ID_NEWITEM, "Custom Item")
End If
Call DrawMenuBar(Me.hwnd)
End Sub

Private Sub cmdCheckMenu_Click()
determineCombosValue
funcResult = CheckMenuItem(mainMenu, menuItem, MF_BYPOSITION Or MF_CHECKED)
End Sub

Private Sub cmdCheckMenuRadio_Click()
funcResult = CheckMenuRadioItem(hSubMenuRadio, 0, 2, 1, MF_BYPOSITION)
End Sub

Private Sub cmdDelMnu_Click()
funcResult = DeleteMenu(hSubMenuEdit, 0, MF_BYPOSITION)
End Sub

Private Sub cmdDisableMenuItem_Click()
funcResult = EnableMenuItem(hSubMenuEdit, 2, MF_BYPOSITION Or MF_DISABLED)
funcResult = EnableMenuItem(hSubMenuEdit, 3, MF_BYPOSITION Or MF_DISABLED)
End Sub

Private Sub cmdEnableMenuItem_Click()
funcResult = EnableMenuItem(hSubMenuEdit, 2, MF_BYPOSITION Or MF_ENABLED)
funcResult = EnableMenuItem(hSubMenuEdit, 3, MF_BYPOSITION Or MF_ENABLED)
End Sub

Private Sub cmdGetMenuItemCount_Click()
funcResult = GetMenuItemCount(hSubMenuFile)
MsgBox funcResult & " items found in File Menu"
End Sub

Private Sub cmdGetMenuItemId_Click()
funcResult = GetMenuItemID(hSubMenuEdit, 0)
MsgBox "ID of EDIT --> UNDO is " & funcResult
End Sub

Private Sub cmdGetMenuItemRect_Click()
funcResult = GetMenuItemRect(Me.hwnd, hSubMenuEdit, 4, rectInfo)
With rectInfo
MsgBox "Top             : " & .Top & vbCrLf & "Bottom        : " & .Bottom & vbCrLf & "Left              : " & .Left & vbCrLf & "Right            : " & .Right
End With
End Sub

Private Sub cmdGetMnuItemInfo_Click()
With mnuInfo
    .cbSize = Len(mnuInfo)
End With
funcResult = GetMenuItemInfo(hSubMenuFile, 0, 1, mnuInfo)
With mnuInfo
MsgBox "fMask                 : " & .fMask & vbCrLf & "fType                 : " & .fType & vbCrLf & "fState                 : " & .fState & vbCrLf & "wID                    : " & .wID & vbCrLf & "hSubMenu          : " & .hSubMenu & vbCrLf & "hbmpChecked    : " & .hbmpChecked & vbCrLf & "hbmpunchecked : " & .hbmpUnchecked & vbCrLf & "dwItemData       : " & .dwItemData & vbCrLf & "dwTypeData      : " & .dwTypeData & vbCrLf & "cch                     : " & .cch
End With
End Sub

Private Sub cmdGrayedMenuItem_Click()
funcResult = EnableMenuItem(hSubMenuEdit, 2, MF_BYPOSITION Or MF_GRAYED)
funcResult = EnableMenuItem(hSubMenuEdit, 3, MF_BYPOSITION Or MF_GRAYED)
End Sub

Private Sub cmdHiliteMenuItem_Click()
funcResult = HiliteMenuItem(Me.hwnd, hSubMenuRadio, 2, MF_BYPOSITION Or MF_HILITE)
End Sub

Private Sub cmdInsertMenuItem_Click()
With mnuInfo
    .cbSize = Len(mnuInfo)
    .fMask = MIIM_TYPE Or MIIM_ID
    .fType = MF_STRING Or MF_RIGHTJUSTIFY
    .wID = ID_HELPMENU
    .dwTypeData = "Help"
    .cch = Len("Help")
End With
funcResult = InsertMenuItem(hMenu, 3, True, mnuInfo)
End Sub

Private Sub cmdModifyMenu_Click()
funcResult = ModifyMenu(hSubMenuEdit, 0, MF_BYPOSITION Or MF_BITMAP, ID_EDITOPEN, Picture2.Picture.Handle)
End Sub

Private Sub cmdRemoveMenu_Click()
funcResult = RemoveMenu(hSubMenuRadio, 2, MF_BYPOSITION)
End Sub

Private Sub cmdSetMenuDefaultItem_Click()
funcResult = SetMenuDefaultItem(hSubMenuFile, 1, 1)
End Sub

Private Sub cmdSetMenuItemBitmaps_Click()
funcResult = SetMenuItemBitmaps(hSubMenuFile, 4, MF_BYPOSITION, Picture3.Picture.Handle, Picture4.Picture.Handle)
End Sub

Private Sub cmdTrackPopupMenu_Click()
Me.ScaleMode = vbPixels
pt.X = Me.ScaleLeft
pt.Y = Me.ScaleHeight / 2
ClientToScreen Me.hwnd, pt
funcResult = TrackPopupMenu(hSubMenuEdit, 0, pt.X, pt.Y, 0, Me.hwnd, ByVal 0&)
End Sub

Private Sub Combo1_Change()
refreshCombos
End Sub

Private Sub Combo1_Click()
refreshCombos
End Sub

Private Sub Form_Load()
hMenu = GetMenu(Me.hwnd)
hSubMenuFile = GetSubMenu(hMenu, 0)
hSubMenuEdit = GetSubMenu(hMenu, 1)
hSubMenuRadio = GetSubMenu(hMenu, 2)
refreshCombos
End Sub

Private Sub Form_Resize()
Dim frmWidth, frmHeight
frmWidth = Me.Width
frmHeight = Me.Height
Me.Width = frmWidth
Me.Height = frmHeight
End Sub

Private Sub lblPopupMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    With pt
        .X = Me.ScaleLeft
        .Y = Me.ScaleHeight / 2
    End With
    ClientToScreen Me.hwnd, pt
    funcResult = TrackPopupMenu(hSubMenuEdit, TPM_RIGHTALIGN Or TPM_RETURNCMD, X, Y, 0, Me.hwnd, rectInfo)
    lblPopupMenu = lblPopupMenu.Caption & vbCrLf & "User clicked " & funcResult
Else
    MsgBox "Problem"
End If
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub
