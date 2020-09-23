Attribute VB_Name = "Globali"

Option Explicit

Public Const LVM_FIRST = &H1000
Public Const HDS_BUTTONS = &H2
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const GWL_STYLE = (-16)
Public Const LVM_GETITEM = LVM_FIRST + 5
Public Const LVM_SETITEM = LVM_FIRST + 6
Public Const LVM_INSERTITEM = LVM_FIRST + 7
Public Const LVIF_INDENT = &H10
Public Const LVM_FINDITEM = LVM_FIRST + 13
Public Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Public Const LVNI_ALL = &H0
Public Const LVNI_FOCUSED = &H1
Public Const LVNI_SELECTED = &H2
Public Const LVNI_CUT = &H4
Public Const LVNI_DROPHILITED = &H8
Public Const LVNI_ABOVE = &H100
Public Const LVNI_BELOW = &H200
Public Const LVNI_TOLEFT = &H400
Public Const LVNI_TORIGHT = &H800
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVS_EX_GRIDLINES = &H1
Public Const LVS_EX_TRACKSELECT = &H8
Public Const LVS_EX_HEADERDRAGDROP = &H10
Public Const LVS_EX_CHECKBOXES = &H4
Public Const LVS_EX_SUBITEMIMAGES = &H2
Public Const LVM_GETITEMSTATE = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXT = (LVM_FIRST + 45)
Public Const LVM_SETCOLUMNWIDTH = LVM_FIRST + 30
Public Const LVIS_STATEIMAGEMASK = &HF000
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
Public Const ICC_LISTVIEW_CLASSES = &H1
Public Const CLR_NONE = &HFFFFFFFF
Public Const LVBKIF_SOURCE_NONE = &H0
Public Const LVBKIF_SOURCE_HBITMAP = &H1
Public Const LVBKIF_SOURCE_URL = &H2
Public Const LVBKIF_SOURCE_MASK = &H3
Public Const LVBKIF_STYLE_NORMAL = &H0
Public Const LVBKIF_STYLE_TILE = &H10
Public Const LVBKIF_STYLE_MASK = &H10
Public Const LVM_SETBKIMAGEA = (LVM_FIRST + 68)
Public Const LVM_SETBKIMAGEW = (LVM_FIRST + 138)
Public Const LVM_GETBKIMAGEA = (LVM_FIRST + 69)
Public Const LVM_GETBKIMAGEW = (LVM_FIRST + 139)
Public Const LVM_SETBKIMAGE = LVM_SETBKIMAGEA
Public Const LVM_GETBKIMAGE = LVM_GETBKIMAGEA
Public Const LVM_GETCOLUMN = (LVM_FIRST + 25)
Public Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
Public Const LVM_GETTEXTCOLOR = (LVM_FIRST + 35)
Public Const LVM_SETTEXTCOLOR = (LVM_FIRST + 36)
Public Const LVM_GETTEXTBKCOLOR = (LVM_FIRST + 37)
Public Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
Public Const LVM_REDRAWITEMS = (LVM_FIRST + 21)
Public Const LVCF_TEXT = &H4
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const LF_FACESIZE = 32
Public Const HDS_HOTTRACK = &H4
Public Const HDI_BITMAP = &H10
Public Const HDI_IMAGE = &H20
Public Const HDI_ORDER = &H80
Public Const HDI_FORMAT = &H4
Public Const HDI_TEXT = &H2
Public Const HDI_WIDTH = &H1
Public Const HDI_HEIGHT = HDI_WIDTH
Public Const HDF_LEFT = 0
Public Const HDF_RIGHT = 1
Public Const HDF_IMAGE = &H800
Public Const HDF_BITMAP_ON_RIGHT = &H1000
Public Const HDF_BITMAP = &H2000
Public Const HDF_STRING = &H4000
Public Const HDM_FIRST = &H1200
Public Const HDM_SETITEM = (HDM_FIRST + 4)
Public Const SWP_DRAWFRAME = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
  
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" _
    (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
    
Public Declare Function SendMessageAny Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
  
Public Declare Function SendMessageLong _
    Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Declare Function SendMessage _
    Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hWnd As Long, _
   ByVal nIndex As Long) As Long
   
Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hWnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
   
Public Declare Function SetWindowPos Lib "user32" _
   (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" _
   (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Public Declare Function SelectObject Lib "gdi32" _
   (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" _
   (ByVal hObject As Long) As Long

Public Declare Function CreateFontIndirect Lib "gdi32" _
    Alias "CreateFontIndirectA" _
    (lpLogFont As LOGFONT) As Long

Public Type HD_ITEM
   mask        As Long
   cxy         As Long
   pszText     As String
   hbm         As Long
   cchTextMax  As Long
   fmt         As Long
   lParam      As Long
   iImage      As Long
   iOrder      As Long
End Type

Public Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Public Type LVITEM
   mask         As Long
   iItem        As Long
   iSubItem     As Long
   state        As Long
   stateMask    As Long
   pszText      As String
   cchTextMax   As Long
   iImage       As Long
   lParam       As Long
   iIndent      As Long
End Type

Public Type LVBKIMAGE
    uFlags As Long
    hBmp As Long
    pszImage As String
    cchImageMax As Long
    xOffsetPercent As Long
    yOffsetPercent  As Long
End Type

Public Type tagINITCOMMONCONTROLSEX
    dwSize As Long
    dwICC As Long
End Type

Public Type LVCOLUMN
    mask As Long
    fmt As Long
    cx As Long
    pszText As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type

Public Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName(LF_FACESIZE) As Byte
End Type



