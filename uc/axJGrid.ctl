VERSION 5.00
Begin VB.UserControl axJGrid 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "axJGrid.ctx":0000
End
Attribute VB_Name = "axJGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' By J. Elihu @ FloresSystems.io
'----------------------------------------
' Version -> 2.3
' Dependencias:
'    - cScrollBar
'    - cSubClass
'----------------------------------------
' AxioUK - David Rojas A. 21-11-2020
'----------------------------------------
' Version -> 2.5.8
' Dependencias:
'    - cScrollBar
'    - cSubClass
'
'
Option Explicit

''GDI+ ------------------
Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal Token As Long)

Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal Brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long

Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private GdipToken As Long
Private nScale    As Single

Private Const LOGPIXELSX As Long = 88
Private Const WrapModeTileFlipXY = &H3
Private Const SmoothingModeAntiAlias As Long = 4

''----------------------
Private Type POINTAPI
    X   As Long
    Y   As Long
End Type

Private Type Rect
    l   As Long
    t   As Long
    r   As Long
    b   As Long
End Type

Private Type UcsRGBQuad
    r   As Byte
    g   As Byte
    b   As Byte
    A   As Byte
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

'/Header Style
Private Const HDS_HORZ = &H0
Private Const HDS_BUTTONS = &H2
Private Const HDS_HIDDEN = &H8
Private Const HDS_HOTTRACK = &H4
Private Const HDS_DRAGDROP = &H40
Private Const HDS_FULLDRAG = &H80

'/Header Item
Private Const HDI_WIDTH = &H1
Private Const HDI_HEIGHT = HDI_WIDTH
Private Const HDI_TEXT = &H2
Private Const HDI_FORMAT = &H4
Private Const HDI_LPARAM = &H8
Private Const HDI_BITMAP = &H10
Private Const HDI_IMAGE = &H20
Private Const HDI_DI_SETITEM = &H40
Private Const HDI_ORDER = &H80
Private Const HDI_FILTER = &H100

'/Header Messages
Private Const HDM_FIRST = &H1200
Private Const HDM_GETITEMCOUNT = (HDM_FIRST + 0)
Private Const HDM_INSERTITEM = (HDM_FIRST + 1)
Private Const HDM_HITTEST = (HDM_FIRST + 6)
Private Const HDM_GETITEMDROPDOWNRECT = (HDM_FIRST + 25)
Private Const HDM_DELETEITEM = (HDM_FIRST + 2)
Private Const HDM_GETITEM = (HDM_FIRST + 3)
Private Const HDM_SETITEM = (HDM_FIRST + 4)
Private Const HDM_LAYOUT = (HDM_FIRST + 5)
Private Const HDM_ORDERTOINDEX = (HDM_FIRST + 15)
Private Const HDM_GETITEMRECT = (HDM_FIRST + 7)
Private Const HDM_SETIMAGELIST = (HDM_FIRST + 8)
Private Const HDM_GETIMAGELIST = (HDM_FIRST + 9)

'/Header Messages
Private Const H_MAX As Long = &HFFFF + 1
Private Const HDN_FIRST = H_MAX - 300&                  '// header
Private Const HDN_LAST = H_MAX - 399&
Private Const HDN_ITEMCHANGING     As Long = (HDN_FIRST - 20) 'Unicode
Private Const HDN_ITEMCLICK        As Long = (HDN_FIRST - 22) 'Unicode
Private Const HDN_ITEMDBLCLICK     As Long = (HDN_FIRST - 23) 'Unicode
Private Const HDN_DIVIDERDBLCLICK  As Long = (HDN_FIRST - 25) 'unicode
Private Const HDN_BEGINTRACK    As Long = (HDN_FIRST - 26) 'Unicode
Private Const HDN_ENDTRACK      As Long = (HDN_FIRST - 27) 'Unicode
Private Const HDN_TRACK     As Long = (HDN_FIRST - 28) 'Unicode
Private Const HDN_DROPDOWN  As Long = (HDN_FIRST - 18)
Private Const HDN_FILTERBTNCLICK   As Long = (HDN_FIRST - 13)
Private Const HDN_FILTERCHANGE     As Long = (HDN_FIRST - 12)
Private Const HDN_ITEMCHECK        As Long = (HDN_FIRST - 16) 'The name is invented, not found his real name
Private Const HDN_BEGINDRAG = (HDN_FIRST - 10)
Private Const HDN_ENDDRAG = (HDN_FIRST - 11)

'/Header Flags
Private Const HDF_OWNERDRAW = &H8000
Private Const HDF_STRING = &H4000
Private Const HDF_BITMAP = &H2000
Private Const HDF_IMAGE = &H800

Private Type HDITEM
    mask        As Long
    cxy         As Long
    pszText     As String
    hbm         As Long
    cchTextMax  As Long
    fmt         As Long
    lParam      As Long
    iImage      As Long
    iOrder      As Long
    Type        As Long
    pvFilter    As Long
End Type
Private Type HDHITTESTINFO
    PT          As POINTAPI
    Flags       As Long
    iItem       As Long
End Type
Private Type NMHDR
    hwndFrom    As Long
    idfrom      As Long
    code        As Long
End Type
Private Type NMHEADER
    HDR         As NMHDR
    iItem       As Long
    iButton     As Long
    lPtrHDItem  As Long '    HDITEM  FAR* pItem
End Type

Private Type PAINTSTRUCT
    hDC                     As Long
    fErase                  As Long
    rcPaint                 As Rect
    fRestore                As Long
    fIncUpdate              As Long
    rgbReserved(1 To 32)    As Byte
End Type

'/Skin To Header
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'/Theme
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Any) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long

'/Header window
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetFocusEx Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'Mouse
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

'?Border
Private Declare Function GetWindowRect& Lib "user32" (ByVal hwnd As Long, lpRect As Rect)
Private Declare Function ExcludeClipRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long


'/WindowMessages
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'/ImageList
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
Private Declare Function ImageList_AddMasked Lib "Comctl32" (ByVal himl As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_Add Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "Comctl32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Private Declare Function ImageList_DrawEx Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long

'/Draw
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function OleTranslateColor2 Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As Rect, lpSourceRect As Rect) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'JGridEvents
Public Event ItemClick(ByVal Row As Long, ByVal Column As Long)
Public Event ItemDblClick(ByVal Row As Long, ByVal Column As Long)
Public Event ItemMouseUp(ByVal Row As Long, ByVal Column As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ItemMouseDown(ByVal Row As Long, ByVal Column As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event SelectionChanged(ByVal Row As Long, ByVal Column As Long)
Public Event RequestItemDrawingData(ByVal Row As Long, ByVal Column As Long, ByRef ForeColor1 As Long, ByRef ForeColor2 As Long, ByRef BackColor As Long, ByRef BorderColor As Long, ByRef Alpha As Long, ByRef ItemIdent As Long)
Public Event ItemDrawing(ByVal Row As Long, ByVal Column As Long, ByRef hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, ByRef CancelDraw As Boolean)
Public Event RequestEditControl(ByVal Row As Long, ByVal Column As Long, ByRef X As Long, ByRef Y As Long, ByRef W As Long, ByRef H As Long, ByRef Control As Object, ByRef CancelMove As Boolean)
Public Event RequestEdit(ByVal Row As Long, ByVal Column As Long, ByVal Text As String, ByRef Control As Object, ByRef Cancel As Boolean)
Public Event RequestCellUpdateT(ByVal Row As Long, ByVal Column As Long, ByRef NewText As String, ByRef Control As Object, ByRef Update As Boolean)
Public Event RequestCellUpdateS(ByVal Row As Long, ByVal Column As Long, ByRef NewSubText As String, ByRef Control As Object, ByRef Update As Boolean)
Public Event MouseEnter()
Public Event MouseExit()

'Header Events
Public Event ColumnClick(ByVal Column As Long)
Public Event ColumnRightClick(ByVal Column As Long)
Public Event ColumnDblClick(ByVal Column As Long)
Public Event ColumnSizeChangeStart(ByVal Column As Long, ByVal Width As Long, Cancel As Boolean)
Public Event ColumnSizeChanging(ByVal Column As Long, ByVal Width As Long, Cancel As Boolean)
Public Event ColumnSizeChanged(ByVal Column As Long, ByVal Width As Long)
Public Event ColumnDividerDblClick(ByVal Column As Long)

'Scroll
Public Event Scroll(eBar As EFSScrollBarConstants)
Public Event ScrollChange(eBar As EFSScrollBarConstants)
Public Event ScrollClick(eBar As EFSScrollBarConstants, eButton As MouseButtonConstants)

'Standart
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Enum eCoincidenceItem
    [ccWholeWord] = 0
    [ccPartial] = 1
End Enum

Public Enum eSortItemsOrder
    [AscendingOrder] = 0
    [DescendingOrder] = 1
End Enum

Private Type tHeader
    Text    As String
    Image   As Long
    Width   As Long
    id      As Long
    Fixed   As Boolean
    Aling   As Integer
    IAlign  As Integer
End Type

Private Type tCell
    Text    As String
    SubText As String
    Tag     As String
    Icon    As Long
End Type

Private Type tRow
    Cell()  As tCell
    Data    As Long
    RowTag  As String
End Type

Private Type tEventDrawing
    Fore1   As Long
    Fore2   As Long
    Back    As Long
    Border  As Long
    Ident   As Long
    Alpha   As Long
    Cancel  As Boolean
End Type

Private Type tMergedColumns
    Title   As String
    iStart  As Long
    iEnd    As Long
    Color   As Long
    eBold   As Boolean
End Type

Private m_CellH         As Long
Private m_HeaderH       As Long
Private m_GridColor As Long
Private m_GridStyle     As Integer
Private m_Striped       As Boolean
Private m_Header        As Boolean
Private m_FullRow       As Boolean
Private m_Editable      As Boolean
Private m_DrawEmpty     As Boolean
Private m_StripedColor  As Long
Private m_oSkin         As StdPicture
Private m_oFont         As StdFont
Private m_cFont         As StdFont
Private m_sFont         As StdFont

Private m_ForeColor     As OLE_COLOR
Private m_ForeColor2    As OLE_COLOR
Private m_SelColor      As OLE_COLOR
Private m_ForeSel       As OLE_COLOR
Private m_BorderColor   As OLE_COLOR
Private m_BackColor     As OLE_COLOR
Private m_Rounded       As Long
Private m_Alpha         As Long

Private m_RowH          As Long

Private WithEvents c_SubClass   As cSubClass
Attribute c_SubClass.VB_VarHelpID = -1
Private WithEvents c_Scroll     As cScrollBars
Attribute c_Scroll.VB_VarHelpID = -1


Private m_hWnd      As Long
Private m_Iml       As Long
Private pmTrack(3)  As Long
Private m_hSkin     As Long
Private m_bmdhFlag  As Boolean

Private m_GridH     As Long
Private m_GridW     As Long
Private m_ImgX      As Long
Private m_ImgY      As Long
Private m_imlFlag   As Boolean
Private m_bTrack    As Boolean
Private m_SelCol    As Long
Private m_SelRow    As Long
Private t_Col       As Long
Private t_Row       As Long
Private e_Row       As Long
Private e_Col       As Long
Private Th          As Integer

Private cHeader()   As tHeader
Private cRow()      As tRow
Private v_Merged()  As tMergedColumns

Private e_Ctrl      As Object
Private e_hWnd      As Long

Private b_DrawFlag  As Boolean
Private b_EditFlag  As Boolean
Private b_Prevent   As Boolean
Private b_ResizeFlag As Boolean
Private b_Merged    As Boolean
Private m_SelFirst  As Boolean

Private m_Gradient  As Boolean

Public Sub AddColumn(ByVal Text As String, Optional ByVal Width As Long = 100, Optional ByVal Alignment As AlignmentConstants, Optional Fixed As Boolean)
Dim l       As Long
Dim i       As Long
Dim tHI    As HDITEM


    l = ColumnCount
    ReDim Preserve cHeader(l)
    With cHeader(l)
        .Text = Text
        .Width = Width
        .Aling = Alignment
        .id = l
        .Fixed = Fixed
    End With
    
    'i = SendMessage(m_hWnd, HDM_GETITEMCOUNT, 0, ByVal 0)
    
    tHI.cxy = Width
    tHI.mask = HDI_TEXT Or HDI_WIDTH Or HDI_FORMAT Or HDI_LPARAM
    tHI.fmt = Alignment Or HDF_STRING
    tHI.lParam = l
    tHI.pszText = Text
    
    Call SendMessage(m_hWnd, HDM_INSERTITEM, l, tHI)
    m_GridW = m_GridW + Width
    If ItemCount And l < ColumnCount Then
        For i = 0 To ItemCount - 1
            ReDim Preserve cRow(i).Cell(l)
            cRow(i).Cell(l).Icon = -1
        Next
    End If
    UpdateScrollH
End Sub

Public Function AddItem(ByVal sText As String, Optional sSubText As String = vbNullString, Optional ByVal IconIndex As Long = -1, Optional ByVal ItemData As Long, Optional ByVal ItemTag As String = "") As Long
On Local Error Resume Next
Dim l   As Long
Dim i   As Long

    l = ItemCount
    ReDim Preserve cRow(l)
    
    With cRow(l)
        ReDim .Cell(ColumnCount - 1)
        .Cell(0).Text = sText
        .Cell(0).SubText = sSubText
        .Cell(0).Icon = IconIndex
        .Data = ItemData
        .RowTag = ItemTag
        For i = 1 To ColumnCount - 1
            .Cell(i).Icon = -1
        Next
    End With
    AddItem = l
    If b_Prevent Then Exit Function
    UpdateScrollV
    DrawGrid
End Function

Public Sub ClearColumns()
Dim l As Long

   Erase cRow
   Erase cHeader
   m_GridW = 0
   Do While SendMessage(m_hWnd, HDM_GETITEMCOUNT, 0, ByVal 0) <> 0
     Call SendMessageByLong(m_hWnd, HDM_DELETEITEM, 0, 0)
   Loop
   UpdateGrid
End Sub

Public Sub ClearItems()
    Erase cRow
    UpdateScrollV
    DrawGrid
End Sub

Public Sub CreateImageList(Optional Width As Integer = 16, Optional Height As Integer = 16, Optional hBitmap As Long, Optional MaskColor As Long = &HFFFFFFFF)
    If m_Iml And m_imlFlag Then ImageList_Destroy m_Iml
    m_Iml = ImageList_Create(Width, Height, &H20, 1, 1)
    m_imlFlag = m_Iml <> 0
    If m_Iml And hBitmap Then
        If (MaskColor <> &HFFFFFFFF) Then
            ImageList_AddMasked m_Iml, hBitmap, MaskColor
        Else
            ImageList_Add m_Iml, hBitmap, 0
        End If
    End If
    If m_Iml Then
        m_ImgX = Width
        m_ImgY = Height
    End If
    
    UpateValues1
End Sub

Public Sub EditEnd()
On Local Error Resume Next
Dim bvData1 As String, bvData2 As String
Dim Evt As Boolean

  If Not b_EditFlag Then Exit Sub
  If e_Ctrl Is Nothing Then Exit Sub
  Evt = False
  bvData1 = "" & e_Ctrl.Text
  bvData2 = "" & e_Ctrl.SubText
  
  If e_hWnd Then c_SubClass.UnSubclass e_hWnd
  e_Ctrl.Visible = False
  
  RaiseEvent RequestCellUpdateT(e_Row, e_Col, bvData1, e_Ctrl, Evt)
  RaiseEvent RequestCellUpdateS(e_Row, e_Col, bvData2, e_Ctrl, Evt)
    
  If bvData1 <> "" Then
    If Evt And cRow(e_Row).Cell(e_Col).Text <> bvData1 Then
        CellText(e_Row, e_Col) = bvData1
    End If
  End If
  If bvData2 <> "" Then
    If Evt And cRow(e_Row).Cell(e_Col).SubText <> bvData2 Then
        SubText(e_Row, e_Col) = bvData2
    End If
  End If
  
    Set e_Ctrl = Nothing
    b_EditFlag = False
    
End Sub

Public Sub EditStart(ByVal Item As Long, Subitem As Long, Optional ByVal bForceEdit As Boolean = False)
On Local Error Resume Next
Dim tHDI    As HDHITTESTINFO
Dim PT      As Rect
Dim Evt     As Boolean

    If Not m_Editable And Not bForceEdit Then Exit Sub
    
    If Item = -1 Or Subitem = -1 Then Exit Sub
    If Item > ItemCount - 1 Or Subitem > ColumnCount - 1 Then Exit Sub
    If b_EditFlag Then EditEnd
    
    e_Row = Item: e_Col = Subitem
    
    SendMessage m_hWnd, HDM_GETITEMRECT, e_Col, PT
    PT.l = PT.l + 1 - GetScroll(efsHorizontal)
    PT.t = ((e_Row * m_RowH) + lHeaderH) - GetScroll(efsVertical)
    PT.r = cHeader(e_Col).Width - IIf(m_GridStyle = 2 Or m_GridStyle = 3, 1, 0)
    PT.b = m_CellH
    
    RaiseEvent RequestEditControl(e_Row, e_Col, PT.l, PT.t, PT.r, PT.b, e_Ctrl, Evt)
    If Not Evt Then
        If Not e_Ctrl Is Nothing Then
            With e_Ctrl
                .Left = PT.l * 15: .Top = PT.t * 15: .Width = PT.r * 15: .Height = PT.b * 15
            End With
        End If
    End If
    Evt = False
    RaiseEvent RequestEdit(e_Row, e_Col, cRow(e_Row).Cell(e_Col).Text, e_Ctrl, Evt)
    If Evt Then Exit Sub
    
    e_Ctrl.Text = cRow(e_Row).Cell(e_Col).Text
    e_Ctrl.Alignment = cHeader(e_Col).Aling
    e_Ctrl.SelStart = 0
    e_Ctrl.SelLength = Len(cRow(e_Row).Cell(e_Col).Text)
    e_Ctrl.Visible = True
    
    e_hWnd = 0
    e_hWnd = e_Ctrl.hwnd
    e_Ctrl.SetFocus
    b_EditFlag = True
    
    If e_hWnd Then
         With c_SubClass
            If .Subclass(e_hWnd, , , Me) Then
                .AddMsg e_hWnd, WM_KILLFOCUS, MSG_AFTER
                .AddMsg e_hWnd, WM_CHAR, MSG_BEFORE
                .AddMsg e_hWnd, WM_KEYDOWN, MSG_BEFORE
            End If
        End With
    End If
    
zCancel:

End Sub

Public Function ItemFind(ByVal Text As String, Optional ByVal Coincidence As eCoincidenceItem = [ccWholeWord], Optional ByVal IgnoreCase As Boolean = True, Optional ByVal Column As Long) As Long
On Local Error GoTo zErr
Dim i As Long
Dim b   As Boolean
    
    If IgnoreCase Then Text = LCase(Text)
    For i = 0 To ItemCount - 1
        If Coincidence = ccWholeWord Then
            b = IIf(IgnoreCase, LCase(cRow(i).Cell(Column).Text) = Text, cRow(i).Cell(Column).Text = Text)
        Else
            b = InStr(1, IIf(IgnoreCase, LCase(cRow(i).Cell(Column).Text), cRow(i).Cell(Column).Text), Text) <> 0
        End If
        If b Then
            ItemFind = i
            Exit Function
        End If
    Next
zErr:
    ItemFind = -1
End Function

Public Function ItemFindData(Data As Long, Optional ByVal Subitem As Long) As Long
On Local Error GoTo zErr
Dim i As Long
    For i = 0 To ItemCount - 1
        If cRow(i).Data = Data Then
            ItemFindData = i
            Exit Function
        End If
    Next
zErr:
    ItemFindData = -1
End Function

Public Function ItemFindTag(ByVal Tag As String) As Long
On Local Error GoTo zErr
Dim i As Long
    For i = 0 To ItemCount - 1
        If cRow(i).RowTag = Tag Then
            ItemFindTag = i
            Exit Function
        End If
    Next
zErr:
    ItemFindTag = -1
End Function

Public Sub MergeColumn(ByVal Title As String, ByVal ColStart As Long, ByVal ColEnd As Long, Optional ByVal ForeColor As Long = -1, Optional FontBold As Boolean)
Dim l As Long
    
    If ColEnd > ColumnCount - 1 Then ColEnd = ColumnCount - 1
    
    l = MergedCount
    ReDim Preserve v_Merged(l)
    With v_Merged(l)
        .Title = Title
        .iStart = ColStart
        .iEnd = ColEnd
        .Color = ForeColor
        .eBold = FontBold
    End With
    DoEvents
    If Not b_Merged Then
        b_Merged = True
        Me.HeaderHeight = Me.HeaderHeight + 22
    End If
 
End Sub

Private Function DrawCell(ByVal hDC As Long, Rect As RECTL, ByVal Color As OLE_COLOR, Round As Long, ByVal Alpha As Long, ByVal Selected As Integer, Optional ByVal Angulo As Single = 0) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mPath As Long
    Dim hGraphics As Long
    Dim Color1 As Long, Color2 As Long
    Dim BorderCell As Long
    Dim InvColor As Long
    
   ' If Round <= 1 Then Round = 2
    
    Select Case Selected
        Case 0
            Color1 = ConvertColor(Color, 10)
        Case 1
            Color1 = ConvertColor(Color, 60)
            UserControl.ForeColor = vbWhite
        Case 2
            Color1 = ConvertColor(Color, Alpha)
    End Select
    
   ' InvColor = &HFFFFFF - Color1 '¿Useless?
    
    Color2 = ConvertColor(Color, Alpha)
    BorderCell = ConvertColor(m_BorderColor, 70)
    
    If m_Gradient Then Color2 = ConvertColor(Color1, Alpha + 15)
    
    GdipCreateFromHDC hDC, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    'Draw Border
    GdipCreatePen1 BorderCell, 1 * nScale, &H2, hPen
    GdipCreateLineBrushFromRectWithAngleI Rect, Color1, Color2, Angulo + 90, 0, WrapModeTileFlipXY, hBrush
    GdipCreatePath &H0, mPath
    
    With Rect
        If Round = 0 Then
            GdipDrawRectangleI hGraphics, hPen, .Left, .Top, .Width, .Height
            GdipAddPathLineI mPath, .Left, .Top, .Width + .Left, .Top            'Line-Top
            GdipAddPathLineI mPath, .Width + .Left, .Top, .Width + .Left, .Height + .Top   'Line-Left
            GdipAddPathLineI mPath, .Width + .Left, .Height + .Top, .Left, .Height + .Top   'Line-Bottom
            GdipAddPathLineI mPath, .Left, .Height + .Top, .Left, .Top         'Line-Right
        ElseIf Round = 1 Then
            GdipAddPathArcI mPath, (.Left + 1), (.Top + 1), 2, 2, 180, 90                       'Top-Left
            GdipAddPathArcI mPath, (.Left + .Width) - (2), .Top + 1, 1, 1, 270, 90      'Top-Right?
            GdipAddPathArcI mPath, (.Left + .Width) - (2), (.Top + .Height) - (2), 1, 1, 0, 90   'Bottom-Right
            GdipAddPathArcI mPath, .Left + 1, (.Top + .Height) - (2), 1, 1, 90, 90       'Bottom-Left
        Else
            GdipAddPathArcI mPath, (.Left + 1), (.Top + 1), Round + 1, Round + 1, 180, 90                       'Top-Left
            GdipAddPathArcI mPath, (.Left + .Width) - (Round + 1), .Top + 1, Round - 1, Round - 1, 270, 90      'Top-Right?
            GdipAddPathArcI mPath, (.Left + .Width) - (Round + 1), (.Top + .Height) - (Round + 1), Round - 1, Round - 1, 0, 90  'Bottom-Right
            GdipAddPathArcI mPath, .Left + 1, (.Top + .Height) - (Round + 1), Round - 1, Round - 1, 90, 90      'Bottom-Left
        End If
    End With
    
    GdipClosePathFigures mPath
    GdipFillPath hGraphics, hBrush, mPath
    GdipDrawPath hGraphics, hPen, mPath
    
    Call GdipDeletePath(mPath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)
    GdipDeleteGraphics hGraphics

    DrawCell = mPath
End Function

Private Function DrawCell2(ByVal hDC As Long, Rect As RECTL, tEDF As tEventDrawing, Round As Long, ByVal Selected As Integer, Optional ByVal Angulo As Single = 0) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mPath As Long
    Dim hGraphics As Long
    Dim Color1 As Long, Color2 As Long
    Dim BorderCell As Long
    Dim InvColor As Long
    
    'If Round <= 1 Then Round = 2
    
    Select Case Selected
        Case 0
            Color1 = ConvertColor(tEDF.Back, 10)
        Case 1
            Color1 = ConvertColor(tEDF.Back, 60)
            UserControl.ForeColor = vbWhite 'tEDF.Fore
        Case 2
            Color1 = ConvertColor(tEDF.Back, tEDF.Alpha)
    End Select
        
    Color2 = ConvertColor(tEDF.Back, tEDF.Alpha)
    BorderCell = ConvertColor(tEDF.Border, 70)
    
    If m_Gradient Then Color2 = ConvertColor(Color1, tEDF.Alpha + 15)
    
    GdipCreateFromHDC hDC, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    'Draw Border
    GdipCreatePen1 BorderCell, 1 * nScale, &H2, hPen
    GdipCreateLineBrushFromRectWithAngleI Rect, Color1, Color2, Angulo + 90, 0, WrapModeTileFlipXY, hBrush
    GdipCreatePath &H0, mPath
    
    With Rect
        If Round = 0 Then
            GdipDrawRectangleI hGraphics, hPen, .Left, .Top, .Width, .Height
            GdipAddPathLineI mPath, .Left + 1, .Top + 1, .Width, .Top + 1           'Line-Top
            GdipAddPathLineI mPath, .Width - 1, .Top + 1, .Width - 1, .Height - 1   'Line-Left
            GdipAddPathLineI mPath, .Width - 1, .Height - 1, .Left + 1, .Height - 1 'Line-Bottom
            GdipAddPathLineI mPath, .Left + 1, .Height - 1, .Left + 1, .Top + 1     'Line-Right
        ElseIf Round = 1 Then
            GdipAddPathArcI mPath, (.Left + 1), (.Top + 1), 2, 2, 180, 90                       'Top-Left
            GdipAddPathArcI mPath, (.Left + .Width) - (2), .Top + 1, 1, 1, 270, 90      'Top-Right?
            GdipAddPathArcI mPath, (.Left + .Width) - (2), (.Top + .Height) - (2), 1, 1, 0, 90   'Bottom-Right
            GdipAddPathArcI mPath, .Left + 1, (.Top + .Height) - (2), 1, 1, 90, 90       'Bottom-Left
        Else
            GdipAddPathArcI mPath, (.Left + 1), (.Top + 1), Round + 1, Round + 1, 180, 90                       'Top-Left
            GdipAddPathArcI mPath, (.Left + .Width) - (Round + 1), .Top + 1, Round - 1, Round - 1, 270, 90      'Top-Right?
            GdipAddPathArcI mPath, (.Left + .Width) - (Round + 1), (.Top + .Height) - (Round + 1), Round - 1, Round - 1, 0, 90  'Bottom-Right
            GdipAddPathArcI mPath, .Left + 1, (.Top + .Height) - (Round + 1), Round - 1, Round - 1, 90, 90      'Bottom-Left
        End If
    End With
    
    GdipClosePathFigures mPath
    GdipFillPath hGraphics, hBrush, mPath
    GdipDrawPath hGraphics, hPen, mPath
    
    Call GdipDeletePath(mPath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)
    GdipDeleteGraphics hGraphics

    DrawCell2 = mPath
End Function

'Inicia GDI+
Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub

'Termina GDI+
Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub

Public Sub RedrawGrid()
    DrawGrid True
End Sub

Public Sub RemoveItem(ByVal Index As Long)
On Local Error Resume Next
Dim j As Integer

    If ItemCount = 0 Or Index > ItemCount - 1 Or ItemCount < 0 Or Index < 0 Then Exit Sub
    
    If ItemCount > 1 Then
         For j = Index To UBound(cRow) - 1
            LSet cRow(j) = cRow(j + 1)
         Next
        ReDim Preserve cRow(UBound(cRow) - 1)
    Else
        Erase cRow
    End If
    
    UpdateScrollV
    If m_SelRow <> -1 Then
        If m_SelRow = Index Then m_SelRow = -1
        If m_SelRow > Index Then m_SelRow = m_SelRow - 1
    End If
    DrawGrid
    
End Sub

Public Function SelectedItemData() As Long
On Local Error GoTo zErr
    If m_SelRow = -1 Then Exit Function
    SelectedItemData = cRow(m_SelRow).Data
    Exit Function
zErr:
    SelectedItemData = -1
End Function

Public Function SetControlToGrid(ByRef Ctrl As Object) As Boolean
On Local Error Resume Next
Dim lp_hWnd As Long
    
    lp_hWnd = Ctrl.hwnd
    If lp_hWnd Then
        SetControlToGrid = SetParent(lp_hWnd, hwnd) <> 0
    End If
End Function

Public Sub SetItem(ByVal Item As Long, ByVal Col As Long, ByVal Text As String, Optional ByVal Icon As Long = -1)
On Error GoTo Err
     cRow(Item).Cell(Col).Text = Text
     cRow(Item).Cell(Col).Icon = Icon
Err:
End Sub

Public Sub SetRow(ByVal Item As Long, vRowData() As String)
On Error Resume Next
Dim i As Long
    For i = 0 To ColumnCount - 1
        cRow(Item).Cell(i).Text = vRowData(i)
    Next
    Call DrawGrid
End Sub

Public Sub SortItems(ByVal Column As Long, ByVal Order As eSortItemsOrder)
Dim out As tRow
Dim d1 As String
Dim d2 As String
Dim i As Long
Dim j As Long
Dim c As Long
Dim l As Long
Dim b As Boolean

    If b_EditFlag Then EditEnd
    
    UserControl.AutoRedraw = False
    
    c = ItemCount - 1
    l = m_SelRow
    For i = 0 To c - 1
         For j = 0 To (c - 1) - i
         
            d1 = cRow(j).Cell(Column).Text
            d2 = cRow(j + 1).Cell(Column).Text
            
            b = IIf(Order = 0, d1 > d2, d1 < d2)
            
            If b Then
            
                LSet out = cRow(j)
                LSet cRow(j) = cRow(j + 1)
                LSet cRow(j + 1) = out
                
                If l = j Then
                    l = j + 1
                ElseIf l = j + 1 Then
                    l = j
                End If
                
            End If
         Next
    Next
    
    Me.SelectedItem = l
    
    UserControl.AutoRedraw = True
End Sub

Public Sub UpdateGrid()
    UpdateScrollV
    UpdateScrollH
    RedrawGrid
    RedrawHeader
End Sub

Private Sub ChangeSelection(eRow As Long, eCol As Long)
    If eRow = m_SelRow And eCol = m_SelCol Then Exit Sub

    m_SelRow = eRow
    m_SelCol = eCol
    
    If m_SelRow = -1 Or m_SelCol = -1 Then
        DrawGrid
        GoTo Evt
    End If
    
    If Not IsCompleteVisibleItem(eRow, eCol) Then
        SetVisibleItem eRow, eCol
    Else
        DrawGrid
    End If
Evt:
    RaiseEvent SelectionChanged(eRow, eCol)
    
End Sub

Private Function CreateGrid()
Dim wStyle      As Long
Dim iFnt        As iFont

    wStyle = &H40000000 Or &H10000000 Or HDS_HORZ
    'wStyle = wStyle Or HDS_DRAGDROP
    wStyle = wStyle Or HDS_BUTTONS
        
    m_hWnd = CreateWindowEx(0, "SysHeader32", "", wStyle, 0, 0, UserControl.ScaleWidth, m_HeaderH, hwnd, 0, App.hInstance, 0)
    If m_hWnd Then
    
        Set iFnt = m_oFont
        
        ShowWindow m_hWnd, Abs(m_Header)
        SendMessage m_hWnd, &H30, iFnt.hFont, 0&
        SendMessage m_hWnd, &H2000 + 5, 1&, ByVal 0&
        
        With c_SubClass
            If .Subclass(m_hWnd, , , Me) Then
                .AddMsg m_hWnd, WM_PAINT, MSG_BEFORE
                .AddMsg m_hWnd, WM_LBUTTONUP, MSG_BEFORE
                .AddMsg m_hWnd, WM_LBUTTONDOWN, MSG_BEFORE
                .AddMsg m_hWnd, WM_SIZE, MSG_BEFORE
                .AddMsg m_hWnd, WM_ERASEBKGND, MSG_AFTER
            End If
        End With
        
    End If
End Function
Private Sub DestroyGrid()
    If m_hWnd Then
        c_SubClass.UnSubclass m_hWnd
        ShowWindow m_hWnd, 0
        DestroyWindow m_hWnd
        m_hWnd = 0
    End If
End Sub

'Private Sub DrawBack(lpDC As Long, Color As Long, Rct As Rect)
'Dim hBrush  As Long
'
'    hBrush = CreateSolidBrush(Color)
'    Call FillRect(lpDC, Rct, hBrush)
'    Call DeleteObject(hBrush)
'End Sub

Private Sub DrawBorder()
If UserControl.BorderStyle = 0 Then Exit Sub
Dim Rct     As Rect
Dim DC      As Long
Dim ix      As Long
Dim hPen    As Long
Dim OldPen  As Long
    
  DC = GetWindowDC(hwnd)
  GetWindowRect hwnd, Rct
          
  Rct.r = Rct.r - Rct.l
  Rct.b = Rct.b - Rct.t
  Rct.l = 0
  Rct.t = 0
  ix = GetSystemMetrics(6)
  ExcludeClipRect DC, ix + 1, ix + 1, Rct.r - (ix + 1), Rct.b - (ix + 1)
      
              
  hPen = CreatePen(0, 1, m_BorderColor)
  OldPen = SelectObject(DC, hPen)
  Rectangle DC, Rct.l, Rct.t, Rct.r, Rct.b
  Call SelectObject(DC, OldPen)
  DeleteObject hPen
             
  ReleaseDC hwnd, DC
End Sub

Private Sub DrawGrid(Optional ByVal bForce As Boolean)
On Local Error Resume Next
Dim lcol    As Long
Dim lRow    As Long
Dim ly      As Long
Dim lx      As Long
Dim lSx     As Long 'Start X
Dim lSCol   As Long 'Start Col
Dim lColW   As Long
Dim uDC     As Long
Dim iRct    As Rect
Dim tRct    As Rect
Dim sRct    As Rect
Dim cRect   As RECTL
Dim LPX     As Long
Dim lPx2    As Long
Dim tEvt    As tEventDrawing
Dim c       As Long
Dim cWidth  As Long
Dim cFore   As OLE_COLOR

    If b_Prevent Then Exit Sub
    
    b_DrawFlag = True
    With UserControl
      .AutoRedraw = True
      .Cls
      .BackColor = m_BackColor
    End With
    
    lcol = 0
    lRow = 0

    lx = -GetScroll(efsHorizontal)
    ly = -GetScroll(efsVertical)
    uDC = UserControl.hDC

    ly = ly + lHeaderH
    lSx = lx
    lSCol = -1
    
    Do While lRow <= ItemCount - 1 And ly < UserControl.ScaleHeight
        
        If ly + m_RowH > 0 Then '?Visible
            
            SetRect iRct, 0, ly, UserControl.ScaleWidth, ly + m_CellH
                                    
           Do While lcol < ColumnCount And lx < UserControl.ScaleWidth
           
              lColW = cHeader(lcol).Width
              
              If lx + lColW > 0 Then
              
                  If (lSCol = -1) Then
                      lSCol = lcol
                      lSx = lx
                  End If
''-> AxioUK
                ''Draw Stripped Rows
                  With cRect
                      .Left = lx + 1: .Top = ly: .Width = lColW: .Height = m_CellH
                  End With
                      
                  If m_Striped And lRow Mod 2 Then
                      DrawCell uDC, cRect, m_GridColor, m_Rounded, m_Alpha, 2
                  End If
                  
                  SetRect iRct, lx + 1, ly, lx + lColW + 1, ly + m_CellH
''-> AxioUK
                '''Start Draw Cells
                  DrawCell uDC, cRect, m_GridColor, m_Rounded, m_Alpha, 2
                '''End Draw Cells
                  
                  'Selection
                  If lRow = m_SelRow And lcol = m_SelCol And Not m_FullRow Then
                      DrawCell uDC, cRect, m_SelColor, m_Rounded, m_Alpha, 1
                  ElseIf lRow = t_Row And lcol = t_Col And Not m_FullRow Then
                      DrawCell uDC, cRect, m_SelColor, m_Rounded, m_Alpha, 0
                  ElseIf m_FullRow Then
                      If lRow = m_SelRow Then
                          DrawCell uDC, cRect, m_SelColor, m_Rounded, 50, 1
                      ElseIf lRow = t_Row Then
                          DrawCell uDC, cRect, m_SelColor, m_Rounded, 10, 0
                      End If
                  End If
                  
                  '?GridLines 0N,1H,2V,3B
                 ' If m_GridStyle = 2 Or m_GridStyle = 3 Then
                 '    DrawLine uDC, lx + lColW + 1, ly, lx + lColW + 1, ly + m_CellH, m_GridLineColor
                 ' End If
                                    
                  'RequestItemDrawing
                  tEvt = EventDrawingField(lRow, lcol)
                  If tEvt.Border <> -1 Then DrawCell2 uDC, cRect, tEvt, m_Rounded, 0
                  RaiseEvent ItemDrawing(lRow, lcol, uDC, iRct.l, iRct.t, iRct.r - iRct.l, iRct.b - iRct.t, tEvt.Cancel)
                  If tEvt.Cancel Then GoTo zDrawNext
                  
                  If m_Iml <> 0 And cRow(lRow).Cell(lcol).Icon <> -1 Then
                      
                      LPX = IIf(m_ImgX + 6 > lColW, lColW - 6, m_ImgX)
                      Select Case cHeader(lcol).IAlign
                          Case 0: SetRect tRct, 4, ((m_CellH - m_ImgY) \ 2), LPX, m_ImgY  'Left
                          Case 1: SetRect tRct, lColW - LPX - 1, ((m_CellH - m_ImgY) \ 2), LPX - 1, m_ImgY 'Right
                                  If tRct.r < 0 Then tRct.r = 0
                          Case 2: 'tRct.R = 0 'Center
                                  SetRect tRct, (lColW - m_ImgX) \ 2, ((m_CellH - m_ImgY) \ 2), LPX, m_ImgY
                                  If lColW - 6 < m_ImgX Then tRct.r = 0
                      End Select
                      If tRct.r Then ImageList_DrawEx m_Iml, cRow(lRow).Cell(lcol).Icon, uDC, lx + tRct.l, ly + tRct.t, tRct.r, 0, &HFFFFFFFF, &HFF000000, 0
                      LPX = m_ImgX + 2
                  Else
                      LPX = 0
                  End If
                  
                  If Trim(cRow(lRow).Cell(lcol).Text) <> vbNullString Then
                  
                      If Trim(cRow(lRow).Cell(lcol).SubText) <> vbNullString Then
                        'Text
                        SetRect tRct, lx + 6 + LPX, ly - (Th / 2), lx + lColW - 2, ly + m_CellH
                        'SubText
                        SetRect sRct, lx + 6 + LPX, ly + (Th / 2) + 1, lx + lColW - 2, ly + m_CellH
                      Else
                        'Text
                        SetRect tRct, lx + 6 + LPX, ly, lx + lColW - 2, ly + m_CellH
                      End If
                    
                      If LPX Then
                          Select Case cHeader(lcol).IAlign
                              Case 0 'Left
                              Case 1 'Right
                                  OffsetRect tRct, -LPX, 0
                              Case 2
                                  SetRect tRct, lx + 6, ly, lx + lColW - 2, ly + m_CellH
                          End Select
                      End If
                      If tRct.r < tRct.l Then tRct.r = tRct.l
                      
                                              
                      If tRct.r - tRct.l > 0 Then
                        'TEXT Normal
                        UserControl.ForeColor = IIf(tEvt.Fore1 <> -1, tEvt.Fore1, m_ForeColor)
                        'TEXT Selection
                        If lRow = m_SelRow And lcol = m_SelCol And Not m_FullRow Then
                            UserControl.ForeColor = InvColor(m_SelColor)
                        ElseIf m_FullRow Then
                            If lRow = m_SelRow Then
                                UserControl.ForeColor = InvColor(m_SelColor)
                            End If
                        End If
                        Set UserControl.Font = m_cFont
                        DrawText uDC, cRow(lRow).Cell(lcol).Text, Len(cRow(lRow).Cell(lcol).Text), tRct, GetTextFlag(lcol)
                        
                        'SUBTEXT Normal
                        UserControl.ForeColor = IIf(tEvt.Fore2 <> -1, tEvt.Fore2, m_ForeColor2)
                        'SUBTEXT Selection
                        If lRow = m_SelRow And lcol = m_SelCol And Not m_FullRow Then
                            UserControl.ForeColor = InvColor(m_SelColor)
                        ElseIf m_FullRow Then
                            If lRow = m_SelRow Then
                                UserControl.ForeColor = InvColor(m_SelColor)
                            End If
                        End If
                        Set UserControl.Font = m_sFont
                        DrawText uDC, cRow(lRow).Cell(lcol).SubText, Len(cRow(lRow).Cell(lcol).SubText), sRct, GetTextFlag(lcol)
                      End If
                  End If
zDrawNext:
              End If
              
              lx = lx + cHeader(lcol).Width
              lcol = lcol + 1
              
           Loop
           
          '?Reset to Scroll Position
          lcol = lSCol
          lx = lSx
        End If
        
        ly = ly + m_RowH
        lRow = lRow + 1
    Loop
    
    'Completar Rows
    If ly < UserControl.ScaleHeight And Not m_RowH = 0 And m_DrawEmpty And Ambient.UserMode Then
    
        LPX = ly
        If m_GridStyle = 1 Or m_GridStyle = 3 Then
            Do While ly < UserControl.ScaleHeight
            
                '?StripedGrid
                If lRow Mod 2 And m_Striped Then
                    SetRect iRct, 0, ly, UserControl.ScaleWidth, ly + m_CellH
                    'DrawBack uDC, SysColor(m_StripedColor), iRct
                End If
                    
                  '?GridLines 0N,1H,2V,3B
                  'If m_GridStyle = 1 Or m_GridStyle = 3 Then DrawLine uDC, 0, ly + m_CellH, UserControl.ScaleWidth, ly + m_CellH, m_GridLineColor
    
                  ly = ly + m_RowH
                  lRow = lRow + 1
            Loop
        End If
        
    End If

    UserControl.AutoRedraw = False
    
    b_DrawFlag = False
End Sub

Private Sub DrawHeader(lpDC As Long, Rct As Rect, bGradient As Boolean)
Dim hBmp    As Long
Dim DC      As Long
Dim hDCMem  As Long
Dim hPen    As Long
Dim hBrush  As Long
Dim Alpha1  As Long
Dim DivValue As Double
Dim i       As Long
Dim X As Long, Y As Long
Dim W As Long, H As Long
    
   X = Rct.l
   Y = Rct.t
   W = (Rct.l + Rct.r)
   H = (Rct.t + Rct.b) - 2
   
    DC = GetDC(0)
    hDCMem = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, 1, H)
    Call SelectObject(hDCMem, hBmp)
            
    hBrush = CreateSolidBrush(m_BackColor)
    Call SelectObject(lpDC, hBrush)
    Rectangle lpDC, Rct.l, Rct.t, Rct.r, Rct.b
    
    hPen = CreatePen(0, 3, vbWhite) 'pvAlphaBlend(m_GridColor, vbWhite, 1))
    Call SelectObject(lpDC, hPen)
    RoundRect lpDC, X, Y + 1, X + W - 1, Y + H - 1, m_Rounded, m_Rounded
    
    For i = 0 To H
      DivValue = 50
      If bGradient Then DivValue = ((i * 100) / H)            '((i * 255) / H)
      SetPixelV hDCMem, 0, i, pvAlphaBlend(m_GridColor, vbWhite, DivValue)
    Next
    StretchBlt lpDC, X + 1, Y + 1, W - 3, H - 2, hDCMem, 0, 1, 1, H - 1, vbSrcCopy
    
    
    DeleteObject hBrush
    DeleteObject hPen
    DeleteObject hBmp
    DeleteDC DC
    DeleteDC hDCMem
End Sub

Private Sub DrawLine(lpDC As Long, X As Long, Y As Long, x2 As Long, y2 As Long, Color As Long)
Dim PT      As POINTAPI
Dim hPen    As Long
Dim hPenOld As Long

    hPen = CreatePen(0, 1, Color)
    hPenOld = SelectObject(lpDC, hPen)
    Call MoveToEx(lpDC, X, Y, PT)
    Call LineTo(lpDC, x2, y2)
    Call SelectObject(lpDC, hPenOld)
    Call DeleteObject(hPen)
End Sub

Private Sub DrawSelection(lpDC As Long, X As Long, Y As Long, W As Long, H As Long, lIndex As Long)
Dim hBmp    As Long
Dim DC      As Long
Dim hDCMem  As Long
Dim DivValue    As Double
Dim hPen        As Long
Dim Alpha1  As Long
Dim lColor  As Long
Dim i       As Long

    Select Case lIndex
        Case 0: lColor = pvAlphaBlend(vbWhite, m_SelColor, 160)
        Case 1: lColor = pvAlphaBlend(vbWhite, m_SelColor, 60)
        Case 2: lColor = m_SelColor
    End Select

    DC = GetDC(0)
    hDCMem = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, 1, H)
    Call SelectObject(hDCMem, hBmp)
    
    Alpha1 = pvAlphaBlend(lColor, vbWhite, 80) ' pvAlphaBlend(oColorEnd, oColorStar, 25)
    For i = 0 To H
        DivValue = ((i * 100) / H)  '((i * 255) / H)
        SetPixelV hDCMem, 0, i, pvAlphaBlend(lColor, Alpha1, DivValue)
    Next
    StretchBlt lpDC, X + 1, Y + 1, W - 3, H - 2, hDCMem, 0, 1, 1, H - 1, vbSrcCopy
    
    hPen = CreatePen(0, 1, pvAlphaBlend(lColor, lColor, 255))
    Call SelectObject(lpDC, hPen)
    RoundRect lpDC, X, Y, X + W, Y + H, m_Rounded, m_Rounded
    DeleteObject hPen
    
    hPen = CreatePen(0, 1, pvAlphaBlend(vbWhite, lColor, 230))
    Call SelectObject(lpDC, hPen)
    RoundRect lpDC, X + 1, Y + 1, X + W - 1, Y + H - 1, m_Rounded, m_Rounded
    
    DeleteObject hPen
    DeleteObject hBmp
    DeleteDC DC
    DeleteDC hDCMem
End Sub

Private Sub DrawSkinHeader()
On Error Resume Next
Dim tHDHII      As HDHITTESTINFO
Dim PS          As PAINTSTRUCT
Dim iRct       As Rect
Dim tRct       As Rect
Dim m_bDown     As Boolean
Dim iIndex      As Long

Dim iFont       As iFont
Dim hDCMemory   As Long
Dim hBmp        As Long
Dim DC          As Long
Dim j           As Long

'/For Merged Columns
Dim m           As Long
Dim mFlag       As Boolean
Dim mDraw       As Boolean
Dim mRct        As Rect
Dim bFlag       As Boolean


    If m_oSkin Is Nothing Then Exit Sub
    
    If m_hSkin = 0 Then pSelectSkin m_oSkin.Handle
    Call BeginPaint(m_hWnd, PS)

    '/Crear DC en memoria
    DC = GetDC(0)
    hDCMemory = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, PS.rcPaint.r, PS.rcPaint.b)
    Call SelectObject(hDCMemory, hBmp)
    'hDCMemory = PS.Hdc
    
    pRenderSkin hDCMemory, 0, 0, PS.rcPaint.r, PS.rcPaint.b, m_hSkin, 0, 0, 4, 26, 0
    SetBkMode hDCMemory, 1
    
    Set iFont = m_oFont
    SelectObject hDCMemory, iFont.hFont

    Call GetCursorPos(tHDHII.PT)
    Call ScreenToClient(m_hWnd, tHDHII.PT)
    Call SendMessage(m_hWnd, HDM_HITTEST, 0, tHDHII)
    
    m_bDown = m_bmdhFlag
    iIndex = tHDHII.iItem
       
    For j = 0 To Me.ColumnCount - 1
    
        SendMessage m_hWnd, HDM_GETITEMRECT, j, iRct
        CopyRect tRct, iRct
        
        tRct.r = tRct.r - 10
        OffsetRect tRct, 5, 0
        
        '\Is Merged?
        If b_Merged Then
            If IsMerged(j, m, mDraw) Then
                tRct.t = 20
                If Not mFlag Then
                    CopyRect mRct, iRct
                    mRct.b = 20
                    mFlag = True
                Else
                    mRct.r = mRct.r + (iRct.r - iRct.l)
                End If
                iRct.t = 20
            End If
        End If
        
        If iIndex = j And m_bDown Then 'MouseDown Then
        
            OffsetRect tRct, 1, 1
            pRenderSkin hDCMemory, iRct.l, iRct.t, iRct.r - iRct.l, iRct.b - iRct.t, m_hSkin, 10, 0, 5, 26, 2
            SetTextColor hDCMemory, GetPixel(m_hSkin, 15, 2)
        
        ElseIf iIndex = j Then
            pRenderSkin hDCMemory, iRct.l, iRct.t, iRct.r - iRct.l, iRct.b - iRct.t, m_hSkin, 5, 0, 5, 26, 2
            SetTextColor hDCMemory, GetPixel(m_hSkin, 15, 1)
            
        Else
            pRenderSkin hDCMemory, iRct.l, iRct.t, iRct.r - iRct.l, iRct.b - iRct.t, m_hSkin, 0, 0, 5, 26, 2
            SetTextColor hDCMemory, GetPixel(m_hSkin, 15, 0)
            
        End If
        
        DrawText hDCMemory, cHeader(j).Text, Len(cHeader(j).Text), tRct, GetTextFlag(j)
        
        '\Draw Merged Columns
        If mDraw And b_Merged Then
        
            pRenderSkin hDCMemory, mRct.l, 0, mRct.r - mRct.l, mRct.b, m_hSkin, 0, 0, 5, 25, 2
            
            bFlag = m_oFont.Bold
            If Not bFlag And v_Merged(m).eBold Then
                m_oFont.Bold = True
                SelectObject hDCMemory, iFont.hFont
            End If
            
            SetTextColor hDCMemory, IIf(v_Merged(m).Color <> -1, v_Merged(m).Color, GetPixel(m_hSkin, 15, 0))
            DrawText hDCMemory, v_Merged(m).Title, Len(v_Merged(m).Title), mRct, &H4 Or &H20 Or &H40000 Or &H3
            SelectObject hDCMemory, iFont.hFont
            
            If Not bFlag And v_Merged(m).eBold Then
                m_oFont.Bold = False
                SelectObject hDCMemory, iFont.hFont
            End If
            
            mDraw = False
            mFlag = False

        End If
        
    Next
    
    StretchBlt PS.hDC, 0, 0, PS.rcPaint.r, PS.rcPaint.b, hDCMemory, 0, 0, PS.rcPaint.r, PS.rcPaint.b, vbSrcCopy
    
    Call EndPaint(m_hWnd, PS)
    DeleteObject hBmp
    DeleteDC DC
    DeleteDC hDCMemory

DrawH:

End Sub


'/ This is only for the merged columns because the cut has some flaws
Private Sub DrawThemeHeader()
On Error Resume Next
Dim uTheme      As Long
Dim tHDHII      As HDHITTESTINFO
Dim PS          As PAINTSTRUCT
Dim iRct       As Rect
Dim tRct       As Rect
Dim cRect      As RECTL
Dim hRect      As RECTL
Dim m_bDown     As Boolean
Dim iIndex      As Long

Dim iFont       As iFont
Dim hDCMemory   As Long
Dim hBmp        As Long
Dim DC          As Long
Dim j           As Long

'/For Merged Columns
Dim m           As Long
Dim mFlag       As Boolean
Dim mDraw       As Boolean
Dim mRct        As Rect
Dim bFlag       As Boolean

    uTheme = OpenThemeData(m_hWnd, StrPtr("Header"))
    If uTheme = 0 Then Exit Sub
    
    Call BeginPaint(m_hWnd, PS)

    '/Crear DC en memoria
    DC = GetDC(0)
    hDCMemory = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, PS.rcPaint.r, PS.rcPaint.b)
    Call SelectObject(hDCMemory, hBmp)
    
    SetBkMode hDCMemory, 1
    
    Set iFont = m_oFont
    SelectObject hDCMemory, iFont.hFont

    Call GetCursorPos(tHDHII.PT)
    Call ScreenToClient(m_hWnd, tHDHII.PT)
    Call SendMessage(m_hWnd, HDM_HITTEST, 0, tHDHII)
    
    m_bDown = m_bmdhFlag
    iIndex = tHDHII.iItem
    
    Call DrawThemeBackground(uTheme, hDCMemory, 0, 0&, PS.rcPaint, ByVal 0&)
    
    For j = 0 To ColumnCount - 1
    
        SendMessage m_hWnd, HDM_GETITEMRECT, j, iRct
        CopyRect tRct, iRct
        
'//-> tRct = Text Rect
        tRct.r = tRct.r - 10
        OffsetRect tRct, 5, 0
        
'//-> Round=0
        With cRect
          .Left = iRct.l + 1
          .Top = iRct.t
          .Width = iRct.r - iRct.l
          .Height = iRct.b - iRct.t
        End With
        
'\Is Merged?
        If IsMerged(j, m, mDraw) Then
            tRct.t = 20
            If Not mFlag Then
                CopyRect mRct, iRct
                mRct.b = 22
                mFlag = True
            Else
                mRct.r = mRct.r + (iRct.r - iRct.l)
            End If
            iRct.t = 20
            iRct.b = (iRct.b / 2) - 17
        End If
          
        If iIndex = j And m_bDown Then 'MouseDown Then
            OffsetRect tRct, 1, 1
            DrawCell hDCMemory, cRect, m_GridColor, m_Rounded, m_Alpha + 30, 2
        ElseIf iIndex = j Then
            DrawCell hDCMemory, cRect, m_GridColor, m_Rounded, m_Alpha, 2
        Else
            DrawCell hDCMemory, cRect, m_GridColor, m_Rounded, m_Alpha, 2
        End If
        
        SetTextColor hDCMemory, vbWindowText
        DrawText hDCMemory, cHeader(j).Text, Len(cHeader(j).Text), tRct, GetTextFlag(j)
        
        '\Draw Merged Columns
        If mDraw Then
        
            With hRect
              .Left = mRct.l
              .Top = mRct.t
              .Width = (mRct.r - mRct.l)
              .Height = (mRct.t + mRct.b) - 1
            End With
            
            DrawCell hDCMemory, hRect, m_GridColor, m_Rounded, m_Alpha, 2
            
            bFlag = m_oFont.Bold
            If Not bFlag And v_Merged(m).eBold Then
                m_oFont.Bold = True
                SelectObject hDCMemory, iFont.hFont
            End If
            
            SetTextColor hDCMemory, IIf(v_Merged(m).Color <> -1, v_Merged(m).Color, vbWindowText)
            DrawText hDCMemory, v_Merged(m).Title, Len(v_Merged(m).Title), mRct, &H4 Or &H20 Or &H40000 Or &H3
            SelectObject hDCMemory, iFont.hFont
            
            If Not bFlag And v_Merged(m).eBold Then
                m_oFont.Bold = False
                SelectObject hDCMemory, iFont.hFont
            End If
            
            mDraw = False
            mFlag = False

        End If
        
    Next
    
    StretchBlt PS.hDC, 0, 0, PS.rcPaint.r, PS.rcPaint.b, hDCMemory, 0, 0, PS.rcPaint.r, PS.rcPaint.b, vbSrcCopy
    
    Call EndPaint(m_hWnd, PS)
    Call CloseThemeData(uTheme)
    
    DeleteObject hBmp
    DeleteDC DC
    DeleteDC hDCMemory

End Sub

Private Function EventDrawingField(lRow As Long, lcol As Long) As tEventDrawing

    With EventDrawingField
        .Back = -1
        .Border = -1
        .Fore1 = -1
        .Fore2 = -1
        .Alpha = 0
        .Cancel = False
        .Ident = 0
         RaiseEvent RequestItemDrawingData(lRow, lcol, .Fore1, .Fore2, .Back, .Border, .Alpha, .Ident)
    End With
   
End Function

Private Function GetColFromX(ByVal X As Long) As Long
Dim iCol    As Long
Dim tHDI    As HDHITTESTINFO

    X = X + GetScroll(efsHorizontal) '+ 8
    tHDI.PT.X = X
    Call SendMessage(m_hWnd, HDM_HITTEST, 0, tHDI)
    
    GetColFromX = tHDI.iItem
    
    'iCol = Header.GetHeaderByX(X)
    If iCol <> -1 Then
        'If Header.GetHeaderX(iCol) + Header.GetHeaderWidth(iCol) < X Then iCol = iCol + 1
        'If iCol > Header.GetHeaderCount - 1 Then iCol = -1
    End If
    
    If iCol <> -1 Then
        'GetColFromX = 0 'Header.GetHeaderData(iCol)
    Else
        'GetColFromX = -1
    End If
    
End Function

Private Function GetRowFromY(ByVal Y As Long) As Long
    Y = Y + GetScroll(efsVertical) - lHeaderH
    GetRowFromY = Y \ m_RowH
    If GetRowFromY >= ItemCount Then GetRowFromY = -1
End Function

Private Function GetScroll(eBar As EFSScrollBarConstants) As Long
    GetScroll = IIf(c_Scroll.Visible(eBar), c_Scroll.Value(eBar), 0)
End Function

Private Function GetTextFlag(Col As Long) As Long
    'VerticalCenter-SingleLine-WordElipsis
    GetTextFlag = &H4 Or &H20 Or &H40000
    Select Case cHeader(Col).Aling
        Case 1: GetTextFlag = GetTextFlag Or &H2
        Case 2: GetTextFlag = GetTextFlag Or &H1
    End Select

End Function

Private Function IsCompleteVisibleItem(eRow As Long, eCol As Long) As Boolean
On Local Error Resume Next
Dim Y       As Long
Dim X       As Long
Dim bRow    As Boolean
Dim bCol    As Boolean
Dim tHI     As HDITEM
Dim Rct     As Rect
Dim lP      As Long

    
    SendMessage m_hWnd, HDM_GETITEMRECT, eCol, Rct
    Y = lGridH - ((lGridH + GetScroll(efsVertical)) - (eRow * m_RowH))
    X = Rct.l - (GetScroll(efsHorizontal))

    bRow = Y >= 0 And Y + m_RowH <= lGridH
    bCol = X >= 0 And X + cHeader(eCol).Width <= UserControl.ScaleWidth
    
    IsCompleteVisibleItem = bRow And bCol
End Function

Private Function IsMerged(Col As Long, MergeIndex As Long, uDrawFlag As Boolean) As Boolean
Dim i As Long
    For i = 0 To MergedCount - 1
        With v_Merged(i)
            If Col >= .iStart And Col <= .iEnd Then
                IsMerged = True
                MergeIndex = i
                If Col = .iEnd Then uDrawFlag = True Else uDrawFlag = False
                Exit Function
            End If
        End With
    Next
    MergeIndex = -1
End Function

Private Function IsVisibleItem(eRow As Long, ByVal eCol As Long) As Boolean
On Error Resume Next
Dim Y       As Long
Dim X       As Long
Dim bRow    As Boolean
Dim bCol    As Boolean
Dim tHI     As HDITEM
Dim Rct     As Rect

    SendMessage m_hWnd, HDM_GETITEMRECT, eCol, Rct
    Y = (eRow * m_RowH)
    X = Rct.l
    
    bRow = Y - m_CellH <= lGridH + GetScroll(efsVertical) And Y + lGridH > lGridH + GetScroll(efsVertical)
    bCol = X + cHeader(eCol).Width >= 0 And X <= UserControl.ScaleWidth + GetScroll(efsHorizontal)
    IsVisibleItem = bRow And bCol
End Function

Private Function IsVisibleRow(ByVal eRow As Long) As Boolean
On Error Resume Next
Dim Y As Long

    Y = (eRow * m_RowH)
    If c_Scroll.Visible(efsVertical) = False Then IsVisibleRow = True: Exit Function
    If Y - m_CellH <= lGridH + GetScroll(efsVertical) And Y + lGridH > lGridH + GetScroll(efsVertical) Then
        IsVisibleRow = True
    End If
    
End Function
Private Sub MoveHeader(Optional ByVal lLeft As Long = -1, Optional ByVal lWidth As Long = -1, Optional ByVal lHeight = -1)
    If lLeft = -1 Then lLeft = -GetScroll(efsHorizontal)
    If lWidth = -1 Then lWidth = m_GridW + 5
    If lHeight = -1 Then lHeight = m_HeaderH
    
    MoveWindow m_hWnd, lLeft, 0, lWidth, lHeight, 1
End Sub

Private Function pGetHeaderItemInfo(ByVal lcol As Long, tHI As HDITEM) As Boolean
      If Not (SendMessage(m_hWnd, HDM_GETITEM, lcol, tHI) = 0) Then
         pGetHeaderItemInfo = True
      End If
End Function

Private Function pRenderSkin(ByVal destDC As Long, ByVal destX As Long, ByVal destY As Long, ByVal DestW As Long, ByVal DestH As Long, ByVal SrcDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Size As Long, Optional MaskColor As Long = -1)
Dim Sx2 As Long
Sx2 = Size * 2

    If MaskColor <> -1 Then
        Dim mDC         As Long
        Dim mX          As Long
        Dim mY          As Long
        Dim DC          As Long
        Dim hBmp        As Long
        Dim hOldBmp     As Long
     
        mDC = destDC: DC = GetDC(0)
        destDC = CreateCompatibleDC(0)
        hBmp = CreateCompatibleBitmap(DC, DestW, DestH)
        hOldBmp = SelectObject(destDC, hBmp) ' save the original BMP for later reselection
        mX = destX: mY = destY
        destX = 0: destY = 0
    End If
 
        SetStretchBltMode destDC, vbPaletteModeNone
         
        BitBlt destDC, destX, destY, Size, Size, SrcDC, X, Y, vbSrcCopy  'TOP_LEFT
        StretchBlt destDC, destX + Size, destY, DestW - Sx2, Size, SrcDC, X + Size, Y, Width - Sx2, Size, vbSrcCopy 'TOP_CENTER
        BitBlt destDC, destX + DestW - Size, destY, Size, Size, SrcDC, X + Width - Size, Y, vbSrcCopy 'TOP_RIGHT
        StretchBlt destDC, destX, destY + Size, Size, DestH - Sx2, SrcDC, X, Y + Size, Size, Height - Sx2, vbSrcCopy 'MID_LEFT
        StretchBlt destDC, destX + Size, destY + Size, DestW - Sx2, DestH - Sx2, SrcDC, X + Size, Y + Size, Width - Sx2, Height - Sx2, vbSrcCopy 'MID_CENTER
        StretchBlt destDC, destX + DestW - Size, destY + Size, Size, DestH - Sx2, SrcDC, X + Width - Size, Y + Size, Size, Height - Sx2, vbSrcCopy 'MID_RIGHT
        BitBlt destDC, destX, destY + DestH - Size, Size, Size, SrcDC, X, Y + Height - Size, vbSrcCopy 'BOTTOM_LEFT
        StretchBlt destDC, destX + Size, destY + DestH - Size, DestW - Sx2, Size, SrcDC, X + Size, Y + Height - Size, Width - Sx2, Size, vbSrcCopy   'BOTTOM_CENTER
        BitBlt destDC, destX + DestW - Size, destY + DestH - Size, Size, Size, SrcDC, X + Width - Size, Y + Height - Size, vbSrcCopy 'BOTTOM_RIGHT

    If MaskColor <> -1 Then
        GdiTransparentBlt mDC, mX, mY, DestW, DestH, destDC, 0, 0, DestW, DestH, MaskColor
        SelectObject destDC, hOldBmp
        DeleteObject hBmp
        ReleaseDC 0&, DC
        DeleteDC destDC
    End If

End Function

Private Sub pSelectSkin(Optional lHandle As Long = 0)
    If m_hSkin Then Call DeleteDC(m_hSkin)     'Eliminar el DC
    If lHandle <> 0 Then
        m_hSkin = CreateCompatibleDC(0)        'Crearlo de nuevo
        Call SelectObject(m_hSkin, lHandle)    'Establecer la imagen
    End If

End Sub

Private Function pSetHeaderItemInfo(ByVal lcol As Long, tHI As HDITEM) As Boolean
      If Not (SendMessage(m_hWnd, HDM_SETITEM, lcol, tHI) = 0) Then
         pSetHeaderItemInfo = True
      End If
End Function

Private Sub RedrawHeader()
    RedrawWindow m_hWnd, ByVal 0&, ByVal 0&, &H1
End Sub

Private Sub SetHeaderWidth(eCol As Long, ByVal lWidth As Long)
Dim tHI As HDITEM
    

    tHI.mask = HDI_WIDTH
    Call pGetHeaderItemInfo(eCol, tHI)
    'If tHI.cxy <> lWidth Then
    tHI.cxy = lWidth
    If (pSetHeaderItemInfo(eCol, tHI)) Then
        'RaiseEvent ColumnSizeChanged(Index, Value)
    End If
    'End If
End Sub

Private Sub SetVisibleItem(eRow As Long, eCol As Long)
On Error GoTo zErr
Dim ly  As Integer
Dim lx  As Integer
Dim tHI     As HDITEM
Dim Rct     As Rect
    
    SendMessage m_hWnd, HDM_GETITEMRECT, eCol, Rct
    
    If eRow = -1 Or eCol = -1 Then Exit Sub
    '?Vertical
    ly = eRow * m_RowH
    If (ly + m_RowH) - lGridH > GetScroll(efsVertical) Then
        c_Scroll.Value(efsVertical) = ((ly + m_RowH) + 2) - lGridH
    ElseIf ly < GetScroll(efsVertical) Then
        c_Scroll.Value(efsVertical) = ly
    End If
    
    '?Horizantal
    lx = Rct.l
    If lx + cHeader(eCol).Width > UserControl.ScaleWidth + GetScroll(efsHorizontal) Then
        c_Scroll.Value(efsHorizontal) = ((lx + cHeader(eCol).Width) + 20) - UserControl.ScaleWidth
    ElseIf lx < UserControl.ScaleWidth + GetScroll(efsHorizontal) Then
        c_Scroll.Value(efsHorizontal) = lx - 5
    End If
zErr:
    DrawGrid
End Sub

Public Function ColorToHex(ByVal Color As Long) As String
    Dim bytOut(11) As Byte
    bytOut(0) = &H30& Or ((Color And &HF0&) \ &H10&)
    bytOut(2) = &H30& Or (Color And &HF&)
    bytOut(4) = &H30& Or ((Color And &HF000&) \ &H1000&)
    bytOut(6) = &H30& Or ((Color And &HF00&) \ &H100&)
    bytOut(8) = &H30& Or ((Color And &HF00000) \ &H100000)
    bytOut(10) = &H30& Or ((Color And &HF0000) \ &H10000)
    
    If bytOut(0) > &H39 Then bytOut(0) = bytOut(0) + 7
    If bytOut(2) > &H39 Then bytOut(2) = bytOut(2) + 7
    If bytOut(4) > &H39 Then bytOut(4) = bytOut(4) + 7
    If bytOut(6) > &H39 Then bytOut(6) = bytOut(6) + 7
    If bytOut(8) > &H39 Then bytOut(8) = bytOut(8) + 7
    If bytOut(10) > &H39 Then bytOut(10) = bytOut(10) + 7
    
    ColorToHex = bytOut
End Function

Private Function pvAlphaBlend(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
    Dim clrFore         As UcsRGBQuad
    Dim clrBack         As UcsRGBQuad
    
    OleTranslateColor clrFirst, 0, VarPtr(clrFore)
    OleTranslateColor clrSecond, 0, VarPtr(clrBack)
    With clrFore
        .r = (.r * lAlpha + clrBack.r * (255 - lAlpha)) / 255
        .g = (.g * lAlpha + clrBack.g * (255 - lAlpha)) / 255
        .b = (.b * lAlpha + clrBack.b * (255 - lAlpha)) / 255
    End With
    CopyMemory pvAlphaBlend, clrFore, 4
End Function

Private Function ConvertColor(ByVal Color As Long, ByVal Opacity As Long) As Long
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
  
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function

Private Function IsDarkColor(ByVal lColor As Long) As Boolean
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor lColor, 0, VarPtr(lColor)
    CopyMemory BGRA(0), lColor, 4&
  
    IsDarkColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) < 382
End Function

Private Function InvColor(ByVal Color As OLE_COLOR) As OLE_COLOR
  InvColor = &HFFFFFF - Color
End Function

Private Function SysColor(oColor As Long) As Long
    OleTranslateColor2 oColor, 0, SysColor
End Function

Private Sub UpateValues1()

    Th = UserControl.TextHeight("ÀJ") * 2
    If Th + 4 > m_CellH Then m_CellH = Th + 4
    If m_ImgY + 4 > m_CellH Then m_CellH = m_ImgY + 4
    m_RowH = m_CellH
    'If m_GridStyle = 1 Or m_GridStyle = 3 Then m_RowH = m_RowH + 1
    m_RowH = m_RowH + 1
End Sub

Private Sub UpdateScrollH()
On Local Error Resume Next
Dim lWidth      As Long
Dim lProportion As Long
Dim bFlag       As Boolean
    
    bFlag = c_Scroll.Visible(efsHorizontal)
    lWidth = m_GridW - (UserControl.ScaleWidth - 5)
    MoveHeader 0, UserControl.ScaleWidth + 5
    If (lWidth > 0) Then
        lProportion = lWidth \ (UserControl.ScaleWidth) + 1
        c_Scroll.LargeChange(efsHorizontal) = lWidth \ lProportion
        If c_Scroll.LargeChange(efsHorizontal) < 20 Then c_Scroll.LargeChange(efsHorizontal) = 20
        c_Scroll.Max(efsHorizontal) = lWidth
        c_Scroll.Visible(efsHorizontal) = True
        MoveHeader -GetScroll(efsHorizontal), m_GridW + 12
    Else
        c_Scroll.Visible(efsHorizontal) = False:
        MoveHeader 0, UserControl.ScaleWidth + 5
    End If
    If bFlag <> c_Scroll.Visible(efsHorizontal) Then UpdateScrollV
End Sub


Private Sub UpdateScrollHitTest()
Dim PT As POINTAPI
    GetCursorPos PT
    If WindowFromPoint(PT.X, PT.Y) = hwnd Then
        ScreenToClient hwnd, PT
        
        t_Row = GetRowFromY(PT.Y)
        t_Col = GetColFromX(PT.X)
        If t_Row <> -1 And t_Col = -1 And m_FullRow Then t_Row = -1
        
    Else
        t_Row = -1: t_Col = -1
    End If
End Sub

Private Sub UpdateScrollV()
On Local Error Resume Next
Dim lHeight     As Long
Dim lProportion As Long
Dim ly          As Long
Dim bFlag       As Boolean

    bFlag = c_Scroll.Visible(efsVertical)
    ly = lGridH
    lHeight = ((ItemCount * m_RowH) + 5) - ly
    
    If (lHeight > 0) Then
      lProportion = lHeight \ (ly + 1)
      c_Scroll.LargeChange(efsVertical) = lHeight \ lProportion
      c_Scroll.Max(efsVertical) = lHeight '+ 1
      c_Scroll.Visible(efsVertical) = True
    Else
      c_Scroll.Visible(efsVertical) = False
    End If
    If bFlag <> c_Scroll.Visible(efsVertical) Then UpdateScrollH
End Sub

Private Sub c_Scroll_Change(eBar As EFSScrollBarConstants)
    If b_EditFlag Then EditEnd
    UpdateScrollHitTest
    Call DrawGrid
    If eBar = efsHorizontal Then
        'Sleep 0
        MoveHeader -GetScroll(eBar)
    End If

    RaiseEvent ScrollChange(eBar)
End Sub

Private Sub c_Scroll_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
    DrawGrid
End Sub

Private Sub c_Scroll_Scroll(eBar As EFSScrollBarConstants)
    If b_EditFlag Then EditEnd
    UpdateScrollHitTest
    Call DrawGrid
    If eBar = efsHorizontal Then
        'Sleep 0
        MoveHeader -GetScroll(eBar)
    End If
    RaiseEvent Scroll(eBar)
End Sub

Private Sub c_Scroll_ScrollClick(eBar As EFSScrollBarConstants, eButton As MouseButtonConstants)
    If b_EditFlag Then EditEnd
    RaiseEvent ScrollClick(eBar, eButton)
End Sub

Private Sub UserControl_Click()
    If b_EditFlag Then EditEnd
    If t_Row <> -1 And t_Col <> -1 Then
        RaiseEvent ItemClick(t_Row, t_Col)
        If m_SelFirst And m_Editable Then EditStart t_Row, t_Col
    End If
    RaiseEvent Click
End Sub
Private Sub UserControl_DblClick()
    If t_Row <> -1 And t_Col <> -1 Then
        RaiseEvent ItemDblClick(t_Row, t_Col)
        If m_Editable Then EditStart t_Row, t_Col
    End If
    RaiseEvent DblClick
End Sub

Public Function GetWindowsDPI() As Double
    Dim hDC As Long, LPX  As Double
    hDC = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hDC, LOGPIXELSX))
    ReleaseDC 0, hDC

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Sub UserControl_Initialize()
    Set c_SubClass = New cSubClass
    Set c_Scroll = New cScrollBars
    
    InitGDI
    nScale = GetWindowsDPI
    
    t_Col = -1: t_Row = -1
    m_SelCol = -1: m_SelRow = -1
End Sub

Private Sub UserControl_InitProperties()
    m_HeaderH = 24
    m_GridColor = &HF0F0F0
    m_GridStyle = 3
    m_Striped = True
    m_StripedColor = &HFDFDFD
    m_SelColor = vbHighlight  '&HDDAC84
    m_BorderColor = &H908782    '&HB2ACA5
    m_FullRow = True
    m_Header = True
    m_DrawEmpty = True
    m_BackColor = UserControl.Parent.BackColor
    Set m_oFont = New StdFont
    m_oFont.Name = "Tahoma"
    m_Rounded = 3
    UserControl.BackColor = m_BackColor
    m_Alpha = 30
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim iRow As Long
Dim iCol As Long

    RaiseEvent KeyDown(KeyCode, Shift)
    If ItemCount = 0 Then Exit Sub
    If b_EditFlag Then EditEnd
    Select Case KeyCode
        Case vbKeyDown
            If m_SelRow < ItemCount - 1 Then ChangeSelection m_SelRow + 1, m_SelCol
        Case vbKeyUp
            If m_SelRow > 0 Then ChangeSelection m_SelRow - 1, m_SelCol
            
        Case vbKeyRight, vbKeyTab
            If m_SelCol < ColumnCount - 1 And Not m_FullRow Then
                ChangeSelection m_SelRow, m_SelCol + 1
            ElseIf m_SelCol = ColumnCount - 1 And Not m_FullRow Then
                If m_SelRow < ItemCount - 1 Then ChangeSelection m_SelRow + 1, 0
            End If
        Case vbKeyLeft
            If m_SelCol > 0 And Not m_FullRow Then
                ChangeSelection m_SelRow, m_SelCol - 1
            ElseIf m_SelCol = 0 And Not m_FullRow Then
                If Not m_SelRow = 0 Then ChangeSelection m_SelRow - 1, ColumnCount - 1
            End If
        Case vbKeyEnd, vbKeyHome
            If KeyCode = vbKeyEnd Then ChangeSelection ItemCount - 1, m_SelCol
            If KeyCode = vbKeyHome Then ChangeSelection 0, m_SelCol
        Case vbKeyPageDown, vbKeyPageUp
            If KeyCode = vbKeyPageDown Then c_Scroll.Value(efsVertical) = c_Scroll.Value(efsVertical) + c_Scroll.LargeChange(efsVertical)
            If KeyCode = vbKeyPageUp Then c_Scroll.Value(efsVertical) = c_Scroll.Value(efsVertical) - c_Scroll.LargeChange(efsVertical)
        Case Else
            
            On Error Resume Next
            Dim j           As Long
            Dim lStart      As Long
            Dim pChar       As String
            Dim iChar       As String
            Dim bFound      As Boolean
            Dim lcol        As Long
        
            lStart = m_SelRow + 1
            lcol = IIf(m_FullRow, 0, m_SelCol)
            If lStart > ItemCount - 1 Then lStart = 0
            pChar = Chr(KeyCode)
            If pChar = "" Then Exit Sub
            
            For j = lStart To ItemCount - 1
                iChar = UCase(Left(cRow(j).Cell(lcol).Text, 1))
                If iChar <> "" And pChar = iChar Then
                    ChangeSelection j, lcol
                    bFound = True
                    Exit For
                End If
            Next
            If Not bFound And lStart > 0 Then
                For j = 0 To lStart '- 1
                    iChar = UCase(Left(cRow(j).Cell(lcol).Text, 1))
                    If iChar <> "" And pChar = iChar Then
                        ChangeSelection j, lcol
                        Exit For
                    End If
                Next
            End If
            
    End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If m_Editable And KeyAscii = 13 And m_SelRow <> -1 And m_SelCol <> -1 Then EditStart m_SelRow, m_SelCol
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)
    
    If t_Row = -1 And t_Col = -1 Then
        t_Row = GetRowFromY(Y)
        t_Col = GetColFromX(X)
        If m_FullRow And t_Row <> -1 And t_Col = -1 Then t_Row = -1
    End If
    
    If t_Row <> -1 And t_Col <> -1 Then
        ChangeSelection t_Row, t_Col
        RaiseEvent ItemMouseDown(t_Row, t_Col, Button, Shift, X, Y)
    Else
        ChangeSelection -1, -1
    End If
    
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iCol    As Long
Dim iRow    As Long

    If Not m_bTrack Then
        TrackMouseEvent pmTrack(0)
        m_bTrack = True
        RaiseEvent MouseEnter
    End If
    
    iRow = GetRowFromY(Y)
    iCol = GetColFromX(X)
    If m_FullRow And iRow <> -1 And iCol = -1 Then iRow = -1
    If iRow <> t_Row Or iCol <> t_Col Then
        t_Col = iCol: t_Row = iRow
        DrawGrid
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If t_Row = m_SelRow And t_Col = m_SelCol Then
        RaiseEvent ItemMouseUp(m_SelRow, m_SelCol, Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_HeaderH = .ReadProperty("HeaderH", 24)
        m_GridColor = .ReadProperty("GridColor", &HF0F0F0)
        m_BackColor = .ReadProperty("BackColor", UserControl.Parent.BackColor)
        m_Striped = .ReadProperty("Striped", True)
        m_StripedColor = .ReadProperty("StripedColor", &HFDFDFD)
        m_SelColor = .ReadProperty("SelColor", &HDDAC84)
        m_CellH = .ReadProperty("CellH", 0)
        m_BorderColor = .ReadProperty("BorderColor", &HB2ACA5)
        m_Header = .ReadProperty("Header", True)
        m_FullRow = .ReadProperty("FullRow", True)
        m_ForeColor = .ReadProperty("ForeColor", 0)
        m_ForeColor2 = .ReadProperty("ForeColor2", 0)
        m_Editable = .ReadProperty("Editable", False)
        m_DrawEmpty = .ReadProperty("DrawEmpty", False)
        m_Gradient = .ReadProperty("Gradient", False)
        m_Rounded = .ReadProperty("RoundedCell", 1)
        m_Alpha = .ReadProperty("Alpha", 30)
        Set m_oSkin = .ReadProperty("HeaderSkin", Nothing)
        Set m_cFont = .ReadProperty("FontCellText", UserControl.Font)
        Set m_oFont = .ReadProperty("FontHeader", UserControl.Font)
        Set m_sFont = .ReadProperty("FontSubText", UserControl.Font)
        
    End With
    
    If Ambient.UserMode Then
        With c_Scroll
            .Create hwnd
            .SmallChange(efsHorizontal) = 20 '48
            .SmallChange(efsVertical) = 16
        End With
        With c_SubClass
            If .Subclass(hwnd, , , Me) Then
                .AddMsg hwnd, WM_NOTIFY, MSG_AFTER
                .AddMsg hwnd, WM_MOUSELEAVE, MSG_AFTER
                .AddMsg hwnd, WM_NCPAINT, MSG_AFTER
            End If
        End With
        
        pmTrack(0) = 16&
        pmTrack(1) = &H2
        pmTrack(2) = hwnd
    
        CreateGrid
    End If
    
    UpateValues1
    
End Sub

Private Sub UserControl_Resize()
On Local Error Resume Next
    If b_EditFlag Then EditEnd
    If b_ResizeFlag = True Then Exit Sub
    b_ResizeFlag = True
    Call UpdateGrid
    b_ResizeFlag = False
End Sub

Private Sub UserControl_Show()
    DrawGrid
End Sub

Private Sub UserControl_Terminate()
    DestroyGrid
    If m_Iml And m_imlFlag Then ImageList_Destroy m_Iml: m_Iml = 0
    pSelectSkin 0
    
    TerminateGDI

    Set c_SubClass = Nothing
    Set c_Scroll = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "HeaderH", m_HeaderH
        .WriteProperty "GridColor", m_GridColor
        .WriteProperty "BackColor", m_BackColor
        .WriteProperty "Striped", m_Striped
        .WriteProperty "StripedColor", m_StripedColor
        .WriteProperty "SelColor", m_SelColor
        .WriteProperty "CellH", m_CellH
        .WriteProperty "BorderColor", m_BorderColor
        .WriteProperty "Header", m_Header
        .WriteProperty "FullRow", m_FullRow
        .WriteProperty "ForeColor", m_ForeColor
        .WriteProperty "ForeColor2", m_ForeColor2
        .WriteProperty "Editable", m_Editable
        .WriteProperty "DrawEmpty", m_DrawEmpty
        .WriteProperty "Gradient", m_Gradient
        .WriteProperty "HeaderSkin", m_oSkin
        .WriteProperty "FontCellText", m_cFont
        .WriteProperty "FontSubText", m_sFont
        .WriteProperty "FontHeader", m_oFont
        .WriteProperty "RoundedCell", m_Rounded
        .WriteProperty "Alpha", m_Alpha
    End With
End Sub

Property Get AlignmentItemIcons(Column As Long) As AlignmentConstants
On Error Resume Next
    AlignmentItemIcons = cHeader(Column).IAlign
End Property
Property Let AlignmentItemIcons(Column As Long, Value As AlignmentConstants)
On Error Resume Next
    cHeader(Column).IAlign = Value
    If m_Iml Then DrawGrid
End Property

Property Get Alpha() As Long: Alpha = m_Alpha: End Property
Property Let Alpha(vAlpha As Long)
  m_Alpha = vAlpha
  RedrawGrid
  RedrawHeader
  PropertyChanged "Alpha"
End Property

Property Get BackColor() As OLE_COLOR: BackColor = m_BackColor: End Property
Property Let BackColor(ByVal Value As OLE_COLOR)
    m_BackColor = Value
    RedrawGrid
    RedrawHeader
    PropertyChanged "BackColor"
End Property

Property Get BorderColor() As OLE_COLOR: BorderColor = m_BorderColor: End Property
Property Let BorderColor(ByVal Value As OLE_COLOR)
    m_BorderColor = Value
    UserControl.Cls
    DrawBorder
    RedrawGrid
    RedrawHeader
    PropertyChanged "BorderColor"
End Property

Property Get ColumnCount() As Long
On Local Error Resume Next
    ColumnCount = UBound(cHeader) + 1
End Property

Property Get DrawEmptyGrid() As Boolean: DrawEmptyGrid = m_DrawEmpty: End Property
Property Let DrawEmptyGrid(Value As Boolean)
    m_DrawEmpty = Value
    DrawGrid
    PropertyChanged "DrawEmpty"
End Property

Public Property Get Editable() As Boolean: Editable = m_Editable: End Property
Public Property Let Editable(ByVal Value As Boolean)
    m_Editable = Value
    PropertyChanged "Editable"
End Property

Property Get FontCellText() As StdFont: Set FontCellText = m_cFont: End Property
Property Set FontCellText(ByVal Value As StdFont)
    Set m_cFont = Value
    PropertyChanged "FontCellText"
End Property

Property Get FontSubText() As StdFont: Set FontSubText = m_sFont: End Property
Property Set FontSubText(ByVal Value As StdFont)
    Set m_sFont = Value
    PropertyChanged "FontSubText"
End Property

Property Get FontHeader() As StdFont: Set FontHeader = m_oFont: End Property
Property Set FontHeader(ByVal Value As StdFont)
    Set m_oFont = Value
    PropertyChanged "HeaderFont"
End Property

Public Property Get ForeColor() As OLE_COLOR: ForeColor = m_ForeColor: End Property
Public Property Let ForeColor(ByVal Value As OLE_COLOR)
    m_ForeColor = Value
    RedrawGrid
    PropertyChanged "ForeColor"
End Property

Public Property Get ForeColor2() As OLE_COLOR: ForeColor2 = m_ForeColor2: End Property
Public Property Let ForeColor2(ByVal Value As OLE_COLOR)
    m_ForeColor2 = Value
    RedrawGrid
    PropertyChanged "ForeColor2"
End Property

Property Get FullRowSelection() As Boolean: FullRowSelection = m_FullRow: End Property
Property Let FullRowSelection(Value As Boolean)
    m_FullRow = Value
    PropertyChanged "FullRow"
    DrawGrid
End Property
'Gradient
Property Get Gradient() As Boolean: Gradient = m_Gradient: End Property
Property Let Gradient(Value As Boolean)
    m_Gradient = Value
    'DrawGrid True
    PropertyChanged "Gradient"
    RedrawGrid
    RedrawHeader
End Property

Property Get GridColor() As OLE_COLOR: GridColor = m_GridColor: End Property
Property Let GridColor(ByVal Value As OLE_COLOR)
    m_GridColor = Value
    PropertyChanged "GridColor"
    RedrawGrid
    RedrawHeader
End Property

Property Get Header() As Boolean: Header = m_Header: End Property

Property Get HeaderHeight() As Long: HeaderHeight = m_HeaderH: End Property
Property Let HeaderHeight(ByVal Value As Long)
    m_HeaderH = Value
    MoveHeader lHeight:=m_HeaderH
    UpdateScrollV
    RedrawHeader
    DrawGrid
    PropertyChanged "HeaderH"
End Property

Property Let Header(Value As Boolean)
    m_Header = Value
    If m_hWnd <> 0 Then ShowWindow m_hWnd, Abs(m_Header)
    UpdateScrollV
    DrawGrid
    PropertyChanged "Header"
End Property

Property Get HeaderSkin() As StdPicture: Set HeaderSkin = m_oSkin: End Property
Property Set HeaderSkin(oPic As StdPicture)

  Set m_oSkin = oPic
  
  pSelectSkin 0
  
  If Not m_oSkin Is Nothing Then
    pSelectSkin m_oSkin.Handle
  End If
  
  RedrawHeader
  PropertyChanged "HeaderSkin"
End Property

Property Get hwnd() As Long: hwnd = UserControl.hwnd: End Property
Property Get ItemCount() As Long
On Local Error Resume Next
    ItemCount = UBound(cRow) + 1
End Property

Property Get ItemData(ByVal Item As Long) As Long
On Local Error Resume Next
    ItemData = cRow(Item).Data
End Property
Property Let ItemData(ByVal Item As Long, Value As Long)
On Local Error Resume Next
     cRow(Item).Data = Value
End Property

Property Get ItemHeight() As Long: ItemHeight = m_CellH: End Property
Property Let ItemHeight(ByVal Value As Long)
    m_CellH = Value
    UpateValues1
    UserControl_Resize
    PropertyChanged "CellH"
End Property

Property Get ItemIcon(ByVal Item As Long, Optional ByVal Column As Long) As Long
On Local Error Resume Next
    ItemIcon = cRow(Item).Cell(Column).Icon
End Property
Property Let ItemIcon(ByVal Item As Long, Optional ByVal Column As Long, Value As Long)
On Local Error Resume Next
    If cRow(Item).Cell(Column).Icon = Value Then Exit Property
    cRow(Item).Cell(Column).Icon = Value
    If IsVisibleItem(Item, Column) Then DrawGrid
End Property

Property Get CellText(ByVal Row As Long, Optional ByVal Column As Long) As String
On Local Error Resume Next
    CellText = cRow(Row).Cell(Column).Text
End Property
Property Let CellText(ByVal Row As Long, Optional ByVal Column As Long, Value As String)
On Local Error Resume Next
    If cRow(Row).Cell(Column).Text = Value Then Exit Property
    cRow(Row).Cell(Column).Text = Value
    If IsVisibleItem(Row, Column) Then DrawGrid
End Property

Property Get SubText(ByVal Row As Long, Optional ByVal Column As Long) As String
On Local Error Resume Next
    SubText = cRow(Row).Cell(Column).SubText
End Property
Property Let SubText(ByVal Row As Long, Optional ByVal Column As Long, Value As String)
On Local Error Resume Next
    If cRow(Row).Cell(Column).SubText = Value Then Exit Property
    cRow(Row).Cell(Column).SubText = Value
    If IsVisibleItem(Row, Column) Then DrawGrid
End Property

Property Get MergedCount() As Long
On Local Error GoTo Err
    MergedCount = UBound(v_Merged) + 1
Err:
End Property

Property Let PreventGrid(Value As Boolean)
    b_Prevent = Value
    If Not Value Then UpdateGrid
End Property

Property Get RoundedCell() As Long: RoundedCell = m_Rounded: End Property
Property Let RoundedCell(Value As Long)
  m_Rounded = Value
  RedrawGrid
  RedrawHeader
  PropertyChanged "RoundedCell"
End Property


Property Get RowTag(ByVal Item As Long) As String
On Local Error Resume Next
    RowTag = cRow(Item).RowTag
End Property
Property Let RowTag(ByVal Item As Long, Value As String)
On Local Error Resume Next
     cRow(Item).RowTag = Value
End Property

Property Get SelectedColumn() As Long: SelectedColumn = m_SelCol: End Property
Property Let SelectedColumn(ByVal Value As Long)
    If Value < 0 Then Value = -1
    If Value > ColumnCount - 1 Then Value = -1
    If m_SelCol <> Value Then
        ChangeSelection m_SelRow, Value
    End If
End Property

Property Get SelectedItem() As Long: SelectedItem = m_SelRow: End Property
Property Let SelectedItem(ByVal Value As Long)
    If Value < 0 Then Value = -1
    If Value > ItemCount - 1 Then Value = -1
    If m_SelRow <> Value Then
        ChangeSelection Value, m_SelCol
    End If
End Property

Property Get SelectionColor() As OLE_COLOR: SelectionColor = m_SelColor: End Property
Property Let SelectionColor(ByVal Value As OLE_COLOR)
    m_SelColor = Value
    DrawGrid
    PropertyChanged "SelColor"
End Property

Property Get StripedGrid() As Boolean: StripedGrid = m_Striped: End Property
Property Let StripedGrid(ByVal Value As Boolean)
    m_Striped = Value
    DrawGrid
    PropertyChanged "Striped"
End Property

Private Property Get lGridH() As Long
    lGridH = UserControl.ScaleHeight - lHeaderH
End Property

Private Property Get lHeaderH() As Long
    lHeaderH = IIf(m_Header, m_HeaderH, 0)
End Property


'- ordinal #1
Private Sub WndProc(ByVal bBefore As Boolean, _
       ByRef bHandled As Boolean, _
       ByRef lReturn As Long, _
       ByVal hwnd As Long, _
       ByVal uMsg As Long, _
       ByVal wParam As Long, _
       ByVal lParam As Long, _
       ByRef lParamUser As Long)
'On Error Resume Next
Dim Evt As Boolean

    Select Case hwnd
        Case UserControl.hwnd
            Select Case uMsg
                Case WM_NOTIFY
                        
                        Dim tNMH    As NMHDR
                        Dim tHDN    As NMHEADER
                        Dim lHDI()  As Long

                        CopyMemory tHDN, ByVal lParam, Len(tHDN)
                        
                        ReDim lHDI(1)
                        Select Case tHDN.HDR.code
                            Case HDN_BEGINTRACK
                                If b_EditFlag Then EditEnd
                                
                                CopyMemory lHDI(0), ByVal tHDN.lPtrHDItem, 8
                                If cHeader(tHDN.iItem).Fixed Then
                                  lReturn = 1: bHandled = True
                                  
                                  Exit Sub
                                End If
                                m_bmdhFlag = False
                                RaiseEvent ColumnSizeChangeStart(tHDN.iItem, lHDI(1), Evt)
                                If Evt Then lReturn = 1: bHandled = True
                                
                            Case HDN_TRACK
                                CopyMemory lHDI(0), ByVal tHDN.lPtrHDItem, 8
                                RaiseEvent ColumnSizeChanging(tHDN.iItem, lHDI(1), Evt)
                                
                                If Evt Then
                                    lReturn = 1: bHandled = True
                                    lParam = lHDI(1)
                                    SendMessage m_hWnd, WM_PAINT, 0&, 0&
                                End If
                                
                            Case HDN_ENDTRACK
                                CopyMemory lHDI(0), ByVal tHDN.lPtrHDItem, 8
                                m_GridW = (m_GridW - cHeader(tHDN.iItem).Width) + lHDI(1)
                                cHeader(tHDN.iItem).Width = lHDI(1)
                                UpdateScrollH
                                DrawGrid
                            Case HDN_DIVIDERDBLCLICK
                                If b_EditFlag Then EditEnd
                                RaiseEvent ColumnDividerDblClick(tHDN.iItem)
                            Case HDN_ITEMCLICK
                                If b_EditFlag Then EditEnd
                                RaiseEvent ColumnClick(tHDN.iItem)
                            Case HDN_ITEMDBLCLICK
                                RaiseEvent ColumnDblClick(tHDN.iItem)
                            Case HDN_BEGINDRAG
                                If b_EditFlag Then EditEnd
                            
                            Case HDN_ENDDRAG
                                ReDim lHDI(8)
                                CopyMemory lHDI(0), ByVal tHDN.lPtrHDItem, 36
                                Debug.Print "Drag "; tHDN.iItem; vbTab; lHDI(8)
                        
                        End Select
                        
                Case WM_MOUSELEAVE
                    If t_Row <> -1 Or t_Col <> -1 Then
                        t_Col = -1: t_Row = -1
                        DrawGrid
                    End If
                    m_bTrack = False
                    RaiseEvent MouseExit

                Case WM_NCPAINT
                    If UserControl.BorderStyle = 0 Then Exit Sub
                    Dim Rct As Rect
                    Dim DC As Long
                    Dim ix As Long
                        
                    DC = GetWindowDC(hwnd)
                    GetWindowRect hwnd, Rct
                            
                    Rct.r = Rct.r - Rct.l
                    Rct.b = Rct.b - Rct.t
                    Rct.l = 0
                    Rct.t = 0
                    ix = GetSystemMetrics(6)
                    ExcludeClipRect DC, ix + 1, ix + 1, Rct.r - (ix + 1), Rct.b - (ix + 1)
                        
                    Dim hPen        As Long
                    Dim OldPen      As Long
                                
                    hPen = CreatePen(0, 1, m_BorderColor)
                    OldPen = SelectObject(DC, hPen)
                    Rectangle DC, Rct.l, Rct.t, Rct.r, Rct.b
                    Call SelectObject(DC, OldPen)
                    DeleteObject hPen
                               
                    ReleaseDC hwnd, DC
            End Select
            
        Case m_hWnd

            Select Case uMsg
                Case WM_PAINT
                    If (Not m_oSkin Is Nothing) Or b_Merged Then
                        If (Not m_oSkin Is Nothing) Then
                            DrawSkinHeader
                        Else
                            DrawThemeHeader
                        End If
                        lReturn = 1: bHandled = True
                    Else
                        DrawThemeHeader
                        lReturn = 1: bHandled = True
                    End If
                    
                Case WM_ERASEBKGND
                    lReturn = 1: bHandled = True
                    
                Case WM_SIZE
                     If Not b_ResizeFlag Then
                        GetWindowRect m_hWnd, Rct
                        Rct.r = Rct.r - Rct.l
                        If Rct.r < UserControl.ScaleWidth Then MoveHeader 0, UserControl.ScaleWidth + 5
                     End If
                     
                Case WM_LBUTTONDOWN, WM_LBUTTONUP
                    m_bmdhFlag = uMsg = WM_LBUTTONDOWN
                    If Not m_oSkin Is Nothing Then RedrawHeader
                    
            End Select
            
        Case e_hWnd
            Select Case uMsg
                Case WM_KILLFOCUS
                    If b_EditFlag Then EditEnd
                    
                Case WM_CHAR
                    If wParam = 27 Then
                    ElseIf wParam = 13 Then
                        If b_EditFlag Then EditEnd
                    ElseIf wParam = 9 Then
                    End If
                    
                Case WM_KEYDOWN
                    'Debug.Print wParam
            End Select
    End Select
 
End Sub




