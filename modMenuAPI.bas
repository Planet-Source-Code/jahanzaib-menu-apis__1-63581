Attribute VB_Name = "modAPIDeclarations"

'API function to get the Handle fo menus
'required by almost all menu-related API functions
Public Declare Function GetMenu Lib "user32" ( _
    ByVal hwnd As Long) _
    As Long

'API function to get the handle fo submenu
'required by many menu-related API functions
Public Declare Function GetSubMenu Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal nPos As Long) _
    As Long

Public Declare Function DrawMenuBar Lib "user32" ( _
    ByVal hwnd As Long) _
    As Long

'Constant values
Public Const MF_BITMAP = &H4&
Public Const MF_CHECKED = &H8&
Public Const MF_DISABLED = &H2&
Public Const MF_ENABLED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_HILITE = &H80&
Public Const MF_UNHILITE = &H0&
Public Const MF_RIGHTJUSTIFY = &H4000&
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" ( _
    ByVal hMenu As Long, _
    ByVal wFlags As Long, _
    ByVal wIDNewItem As Long, _
    ByVal lpNewItem As Any) _
    As Long
    
Public Declare Function CheckMenuItem Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal wIDCheckItem As Long, _
    ByVal wCheck As Long) _
    As Long

Public Declare Function CheckMenuRadioItem Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal un1 As Long, _
    ByVal un2 As Long, _
    ByVal un3 As Long, _
    ByVal un4 As Long) _
    As Long
    
Public Declare Function DeleteMenu Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long) _
    As Long

Public Declare Function EnableMenuItem Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal wIDEnableItem As Long, _
    ByVal wEnable As Long) _
    As Long

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_CHECKMARKS = &H8
Public Const MIIM_TYPE = &H10
Public Const MIIM_DATA = &H20
Public Const MIIM_STRING = &H40
Public Const MIIM_BITMAP = &H80
Public Const MIIM_FTYPE = &H100

Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" ( _
    ByVal hMenu As Long, _
    ByVal un As Long, _
    ByVal b As Long, _
    lpMenuItemInfo As MENUITEMINFO) _
    As Long
    
Public mnuInfo As MENUITEMINFO

Public Declare Function GetMenuItemCount Lib "user32" ( _
    ByVal hMenu As Long) _
    As Long

Public Declare Function GetMenuItemID Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal nPos As Long) _
    As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function GetMenuItemRect Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hMenu As Long, _
    ByVal uItem As Long, _
    lprcItem As RECT) _
    As Long

Public rectInfo As RECT

Public Declare Function HiliteMenuItem Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hMenu As Long, _
    ByVal wIDHiliteItem As Long, _
    ByVal wHilite As Long) _
    As Long

Public Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" ( _
    ByVal hMenu As Long, _
    ByVal un As Long, _
    ByVal bool As Boolean, _
    ByRef lpcMenuItemInfo As MENUITEMINFO) _
    As Long

Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" ( _
    ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long, _
    ByVal wIDNewItem As Long, _
    ByVal lpString As Any) _
    As Long

Public Declare Function RemoveMenu Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long) _
    As Long

Public Declare Function SetMenuDefaultItem Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal uItem As Long, _
    ByVal fByPos As Long) _
    As Long

Public Declare Function SetMenuItemBitmaps Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long, _
    ByVal hBitmapUnchecked As Long, _
    ByVal hBitmapChecked As Long) _
    As Long

Public Const TPM_RIGHTALIGN = &H8&
Public Const TPM_RIGHTBUTTON = &H2&
Public Const TPM_LEFTBUTTON = &H0&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_RETURNCMD = &H100&
Public Const TPM_NONOTIFY = &H80&
Public Const TPM_VCENTERALIGN = &H10&
Public Const TPM_BOTTOMALIGN = &H20&
Public Const TPM_TOPALIGN = &H0&

Public Declare Function TrackPopupMenu Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal wFlags As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nReserved As Long, _
    ByVal hwnd As Long, _
    lprc As Any) _
    As Long

Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public pt As POINTAPI

