VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{0FB79AC7-66DC-4902-9C66-F58D7CD3DFFF}#5.2#0"; "axJGrid2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8010
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   534
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
   StartUpPosition =   3  'Windows Default
   Begin AXJGDTL.axJGrid axJGrid 
      Height          =   4095
      Left            =   1485
      TabIndex        =   32
      Top             =   1065
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   7223
      HeaderH         =   24
      GridColor       =   15790320
      BackColor       =   -2147483633
      Striped         =   -1  'True
      StripedColor    =   16645629
      SelColor        =   -2147483635
      CellH           =   30
      BorderColor     =   9471874
      Header          =   -1  'True
      FullRow         =   -1  'True
      ForeColor       =   0
      ForeColor2      =   0
      Editable        =   0   'False
      DrawEmpty       =   -1  'True
      Gradient        =   0   'False
      BeginProperty FontCellText {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSubText {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCell     =   3
      Alpha           =   30
   End
   Begin VB.CommandButton btnReLoad 
      Caption         =   "Reload"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   31
      Top             =   570
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A"
      Height          =   315
      Left            =   10890
      TabIndex        =   30
      Top             =   60
      Width           =   300
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Border"
      Height          =   270
      Index           =   5
      Left            =   7785
      TabIndex        =   27
      Top             =   645
      Width           =   600
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Fore2"
      Height          =   270
      Index           =   4
      Left            =   7155
      TabIndex        =   26
      Top             =   645
      Width           =   600
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Back"
      Height          =   270
      Index           =   3
      Left            =   7785
      TabIndex        =   25
      Top             =   360
      Width           =   600
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "Use Request CellUpdate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   120
      TabIndex        =   24
      Top             =   3975
      Width           =   1215
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Fore1"
      Height          =   270
      Index           =   2
      Left            =   7155
      TabIndex        =   23
      Top             =   360
      Width           =   600
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   1410
      TabIndex        =   16
      Top             =   660
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Max             =   20
      SelStart        =   5
      TickStyle       =   3
      Value           =   5
      TextPosition    =   1
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Sel"
      Height          =   270
      Index           =   1
      Left            =   6540
      TabIndex        =   14
      Top             =   645
      Width           =   600
   End
   Begin VB.TextBox ctext 
      Height          =   285
      Left            =   4830
      TabIndex        =   13
      Text            =   "0:0 - ItemText"
      Top             =   30
      Width           =   4050
   End
   Begin VB.TextBox Tx 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   195
      TabIndex        =   12
      Text            =   "1000"
      Top             =   90
      Width           =   615
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "Gradient"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   5520
      TabIndex        =   11
      Top             =   555
      Width           =   855
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Grid"
      Height          =   270
      Index           =   0
      Left            =   6540
      TabIndex        =   10
      Top             =   345
      Width           =   600
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "Header SortItems"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   3270
      TabIndex        =   9
      Top             =   15
      Width           =   1425
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "Skin Header"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   2145
      TabIndex        =   8
      Top             =   15
      Width           =   1020
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "Show Header"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   7
      Left            =   8880
      TabIndex        =   7
      Top             =   9870
      Width           =   1140
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "StripedGrid"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   4470
      TabIndex        =   6
      Top             =   555
      Width           =   1035
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "Use RequestData"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   3285
      Width           =   1215
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "Use OwnerDraw"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2610
      Width           =   1215
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "FullRow selection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1935
      Width           =   1215
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1545
      Width           =   1215
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "Eliminar Seleccionado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   915
      Width           =   1215
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "Add Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   855
      TabIndex        =   0
      Top             =   60
      Width           =   900
   End
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   1560
      Top             =   6525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccionar Fuente"
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   660
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Max             =   100
      SelStart        =   30
      TickStyle       =   3
      Value           =   30
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   255
      Left            =   8745
      TabIndex        =   19
      Top             =   675
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Max             =   50
      SelStart        =   24
      TickStyle       =   3
      Value           =   24
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider Slider4 
      Height          =   255
      Left            =   10095
      TabIndex        =   21
      Top             =   675
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Min             =   17
      Max             =   50
      SelStart        =   17
      TickStyle       =   3
      Value           =   17
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider Slider5 
      Height          =   255
      Left            =   9975
      TabIndex        =   28
      Top             =   90
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   300
      SelStart        =   50
      TickStyle       =   3
      Value           =   50
      TextPosition    =   1
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ColWidth(1)"
      Height          =   195
      Left            =   9105
      TabIndex        =   29
      Top             =   90
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cell Height"
      Height          =   195
      Left            =   10155
      TabIndex        =   22
      Top             =   465
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Header Height"
      Height          =   195
      Left            =   8790
      TabIndex        =   20
      Top             =   465
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alpha Value :"
      Height          =   195
      Left            =   3000
      TabIndex        =   18
      Top             =   435
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rounded Cells :"
      Height          =   195
      Left            =   1470
      TabIndex        =   15
      Top             =   480
      Width           =   1125
   End
   Begin VB.Image skin2 
      Height          =   390
      Left            =   4860
      Picture         =   "Form1.frx":0000
      Top             =   5415
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image skin 
      Height          =   390
      Left            =   4440
      Picture         =   "Form1.frx":06C2
      Top             =   5400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image pb 
      Height          =   480
      Left            =   3480
      Picture         =   "Form1.frx":0BE4
      Top             =   5400
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image iml 
      Height          =   240
      Left            =   1920
      Picture         =   "Form1.frx":2A26
      Top             =   5520
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/Theme
Private Type Rect
    X1      As Long
    Y1      As Long
    x2      As Long
    y2      As Long
End Type

Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Any) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
'/-----

Private m_OwnerDraw As Boolean
Private m_OwnerData As Boolean
Private m_SortFlag  As Boolean

Private LastSortColumn As Long
Private lSortOrder As eSortItemsOrder


Private Sub btnReLoad_Click()
Form_Load
End Sub

Private Sub cmdColor_Click(Index As Integer)
With cmDlg
  .DialogTitle = "Seleccionar Color"
  .ShowColor
  
  Select Case Index
    Case Is = 0
      axJGrid.GridColor = .Color
    Case Is = 1
      axJGrid.SelectionColor = .Color
    Case Is = 2
      axJGrid.ForeColor = .Color
    Case Is = 3
      axJGrid.BackColor = .Color
    Case Is = 4
      axJGrid.ForeColor2 = .Color
    Case Is = 5
      axJGrid.BorderColor = .Color
  End Select
  
End With

End Sub

Private Sub Command1_Click()
Dim I As Long
With axJGrid
  For I = 0 To .ColumnCount - 1
    .AutoWidthCol I
  Next I
End With
End Sub

Private Sub Form_Load()


    With axJGrid
    
        .CreateImageList 16, 16, iml
        
        Slider5.Value = 150 'Set Width Column 1
        
        .AddColumn "Nombres", 150
        .AddColumn "Apellidos", 150
        .AddColumn "Ocupacion", 90
        .AddColumn "Estado", 60, vbCenter
        .AddColumn "Rango", 100, vbRightJustify
        .AddColumn "Sueldo", 80, vbRightJustify
        
        .AlignmentItemIcons(5) = vbRightJustify
        
        '.StripBackColor = RGB(250, 250, 250)
        .FullRowSelection = True
        
        .MergeColumn "Personas", 0, 1, vbBlue
        '.MergeColumn "Información", 1, 2, vbBlue
        
        .ForeColor = &HFF0000
        .ForeColor2 = &H400000
        .RoundedCell = Slider1.Value
    End With
    
    
    m_OwnerDraw = False
    m_OwnerData = False
    LastSortColumn = -1
    
    
    'btnMain_Click 0
    'btnMain_Click 8
        
End Sub
Private Sub Form_Resize()
On Error Resume Next
    axJGrid.Move 95, 65, Me.ScaleWidth - 106, Me.ScaleHeight - 100
End Sub
Private Sub btnMain_Click(Index As Integer)
Dim I As Long
Dim l As Long

    Select Case Index
    
        Case 0
            With axJGrid
            
                .Redraw = False
                For I = 1 To CInt(Tx.Text)
                    If RandomInt(0, 1) = 0 Then
                        'l = .AddItem(vbNullString, "", 1)
                        l = .AddItem(GetForename(ntMale) & " " & GetForename(ntMale), "", 1)
                        .SubText(l, 1) = "$ " & RandomInt(99000, 500000)
                    Else
                        l = .AddItem(GetForename(ntFemale) & " " & GetForename(ntFemale), "account with allocation amount...", 0)
                    End If
                    
                    .CellText(l, 1) = GetSurname() & " " & GetSurname()
                    '.CellText(l, 1) = vbNullString
                    .CellText(l, 3) = IIf(RandomInt(0, 1), "Si", "No")
                    .CellText(l, 4) = RandomInt(0, 100)
                    .CellText(l, 5) = RandomInt(0, 4000) & "$"  'Add text
                    
                    .SetItem l, 2, GetJobName(), 4
                    .SetItem l, 5, RandomInt(0, 4000) & "$", 2 'Add Icon
  
                Next
                
                .Redraw = True
            End With
        Case 1: axJGrid.RemoveItem axJGrid.SelectedItem
        Case 2: axJGrid.ClearGrid
        Case 3: axJGrid.FullRowSelection = Not axJGrid.FullRowSelection
        Case 4:
                m_OwnerDraw = Not m_OwnerDraw
                axJGrid.RefreshGrid
        Case 5:
                m_OwnerData = Not m_OwnerData
                axJGrid.RefreshGrid
        Case 6: axJGrid.StripedGrid = Not axJGrid.StripedGrid
        Case 7: axJGrid.Header = Not axJGrid.Header
        Case 8  'SKIN HEADER
                If axJGrid.HeaderSkin Is Nothing Then
                    Set axJGrid.HeaderSkin = skin.Picture
                Else
                    Set axJGrid.HeaderSkin = Nothing
                End If
        Case 9: m_SortFlag = Not m_SortFlag
        Case 10:
              
        Case 11: axJGrid.Gradient = Not axJGrid.Gradient
    End Select
End Sub

Private Sub axJGrid_ItemClick(ByVal Row As Long, ByVal Column As Long)
ctext.Text = Row & ":" & Column & " - " & axJGrid.CellText(Row, Column)
End Sub

Private Sub axJGrid_ItemDrawing(ByVal Item As Long, ByVal Column As Long, Hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, CancelDraw As Boolean)
    If Not m_OwnerDraw Then Exit Sub
    Select Case Column
        
        Case 3
            Dim hTheme As Long
            Dim Rct As Rect
            Dim x2 As Long
            Dim y2 As Long
            
            hTheme = OpenThemeData(0&, StrPtr("Button"))
            If hTheme Then
                x2 = (W - 14) \ 2
                y2 = (H - 14) \ 2
                
                SetRect Rct, X + x2, Y + y2, X + x2 + 14, Y + y2 + 14
                
                DrawThemeBackground hTheme, Hdc, 3, IIf(axJGrid.CellText(Item, 3) = "Si", 5, 0), Rct, ByVal 0&
                Call CloseThemeData(hTheme)
                CancelDraw = True
            End If
            
            
        Case 4
            Dim lP As Long
            lP = (Val(axJGrid.CellText(Item, 4)) * (W - (4))) / 100
            
             RenderStretchFromPicture Hdc, X + 2, Y + 2, W - 4, H - 4, pb, 0, 0, 60, 16, 4, vbMagenta
             RenderStretchFromPicture Hdc, X + 3, Y + 3, lP, H - 6, pb, 0, 16, 60, 16, 4, vbMagenta
             CancelDraw = True
    End Select
    
End Sub


'/ Ordering Items A-Z, Z-A
Private Sub axJGrid_ColumnClick(ByVal Column As Long)
    If Not m_SortFlag Then Exit Sub

    If LastSortColumn = Column Then lSortOrder = Not lSortOrder Else lSortOrder = False
    axJGrid.SortItems Column, Abs(lSortOrder)
    LastSortColumn = Column

End Sub

Private Sub axJGrid_RequestItemDrawingData(ByVal Row As Long, ByVal Column As Long, ForeColor1 As Long, ForeColor2 As Long, BackColor As Long, BorderColor As Long, Alpha As Long, ItemIdent As Long)
    If Not m_OwnerData Then Exit Sub
    If Column = 0 Then
        If axJGrid.SubText(Row, Column) <> vbNullString Then
            ForeColor1 = vbRed
            ForeColor2 = vbBlue
            BackColor = vbWhite
            BorderColor = vbRed
            Alpha = 100
        End If
    End If

End Sub

Private Sub Slider1_Click()
axJGrid.RoundedCell = Slider1.Value
End Sub

Private Sub Slider2_Click()
axJGrid.Alpha = Slider2.Value
End Sub

Private Sub Slider3_Click()
axJGrid.HeaderHeight = Slider3.Value
End Sub

Private Sub Slider4_Click()
axJGrid.ItemHeight = Slider4.Value
End Sub

Private Sub Slider5_Click()
axJGrid.ColumnWidth(1) = Slider5.Value
End Sub
