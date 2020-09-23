VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drive Space"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   Icon            =   "Drive Space.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   216
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2445
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   2760
      Width           =   1020
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1290
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2760
      Width           =   1020
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2760
      Width           =   1020
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   240
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2100
      Width           =   3300
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Used"
      Height          =   210
      Left            =   2445
      TabIndex        =   7
      Top             =   2520
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Free"
      Height          =   210
      Left            =   1290
      TabIndex        =   6
      Top             =   2520
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Total"
      Height          =   210
      Left            =   135
      TabIndex        =   5
      Top             =   2520
      Width           =   1020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_Free As Integer
Dim m_Total As Integer

Private Const GdiPlusVersion As Long = 1&
Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GdiplusStartupOutput
    NotificationHook As Long
    NotificationUnhook As Long
End Type

' The SmoothingMode enumeration specifies the type of
' smoothing (antialiasing) that is applied to lines and curves.
Private Enum SmoothingMode
   SmoothingModeInvalid
   ' Reserved.
   SmoothingModeDefault = 0
   ' Specifies that smoothing is not applied.
   SmoothingModeHighSpeed = 1
   ' Specifies that smoothing is not applied.
   SmoothingModeHighQuality = 2
   ' Specifies that smoothing is applied using an 8 X 4 box filter.
   SmoothingModeNone = 3
   ' Specifies that smoothing is not applied.
   SmoothingModeAntiAlias8x4 = 4
   ' Specifies that smoothing is applied using an 8 X 4 box filter.
   SmoothingModeAntiAlias
   ' Specifies that smoothing is applied using an 8 X 4 box filter.
   SmoothingModeAntiAlias8x8
   ' Specifies that smoothing is applied using an 8 X 8 box filter.
End Enum

Private Enum GpUnit
   UnitWorld = 0
   ' World coordinate (non-physical unit)
   UnitDisplay = 1
   ' Variable -- for PageTransform only
   UnitPixel = 2
   ' Each unit is one device pixel.
   UnitPoint = 3
   ' Each unit is a printer's point, or 1/72 inch.
   UnitInch = 4
   ' Each unit is 1 inch.
   UnitDocument = 5
   ' Each unit is 1/300 inch.
   UnitMillimeter = 6
   ' Each unit is 1 millimeter.
End Enum

' Common color constants
' NOTE: Oringinal enum was unnamed
Public Enum Colors
   AliceBlue = &HFFF0F8FF
   AntiqueWhite = &HFFFAEBD7
   Aqua = &HFF00FFFF
   Aquamarine = &HFF7FFFD4
   Azure = &HFFF0FFFF
   Beige = &HFFF5F5DC
   Bisque = &HFFFFE4C4
   Black = &HFF000000
   BlanchedAlmond = &HFFFFEBCD
   Blue = &HFF0000FF
   BlueViolet = &HFF8A2BE2
   Brown = &HFFA52A2A
   BurlyWood = &HFFDEB887
   CadetBlue = &HFF5F9EA0
   Chartreuse = &HFF7FFF00
   Chocolate = &HFFD2691E
   Coral = &HFFFF7F50
   CornflowerBlue = &HFF6495ED
   Cornsilk = &HFFFFF8DC
   Crimson = &HFFDC143C
   Cyan = &HFF00FFFF
   DarkBlue = &HFF00008B
   DarkCyan = &HFF008B8B
   DarkGoldenrod = &HFFB8860B
   DarkGray = &HFFA9A9A9
   DarkGreen = &HFF006400
   DarkKhaki = &HFFBDB76B
   DarkMagenta = &HFF8B008B
   DarkOliveGreen = &HFF556B2F
   DarkOrange = &HFFFF8C00
   DarkOrchid = &HFF9932CC
   DarkRed = &HFF8B0000
   DarkSalmon = &HFFE9967A
   DarkSeaGreen = &HFF8FBC8B
   DarkSlateBlue = &HFF483D8B
   DarkSlateGray = &HFF2F4F4F
   DarkTurquoise = &HFF00CED1
   DarkViolet = &HFF9400D3
   DeepPink = &HFFFF1493
   DeepSkyBlue = &HFF00BFFF
   DimGray = &HFF696969
   DodgerBlue = &HFF1E90FF
   Firebrick = &HFFB22222
   FloralWhite = &HFFFFFAF0
   ForestGreen = &HFF228B22
   Fuchsia = &HFFFF00FF
   Gainsboro = &HFFDCDCDC
   GhostWhite = &HFFF8F8FF
   Gold = &HFFFFD700
   Goldenrod = &HFFDAA520
   Gray = &HFF808080
   Green = &HFF008000
   GreenYellow = &HFFADFF2F
   Honeydew = &HFFF0FFF0
   HotPink = &HFFFF69B4
   IndianRed = &HFFCD5C5C
   Indigo = &HFF4B0082
   Ivory = &HFFFFFFF0
   Khaki = &HFFF0E68C
   Lavender = &HFFE6E6FA
   LavenderBlush = &HFFFFF0F5
   LawnGreen = &HFF7CFC00
   LemonChiffon = &HFFFFFACD
   LightBlue = &HFFADD8E6
   LightCoral = &HFFF08080
   LightCyan = &HFFE0FFFF
   LightGoldenrodYellow = &HFFFAFAD2
   LightGray = &HFFD3D3D3
   LightGreen = &HFF90EE90
   LightPink = &HFFFFB6C1
   LightSalmon = &HFFFFA07A
   LightSeaGreen = &HFF20B2AA
   LightSkyBlue = &HFF87CEFA
   LightSlateGray = &HFF778899
   LightSteelBlue = &HFFB0C4DE
   LightYellow = &HFFFFFFE0
   Lime = &HFF00FF00
   LimeGreen = &HFF32CD32
   Linen = &HFFFAF0E6
   Magenta = &HFFFF00FF
   Maroon = &HFF800000
   MediumAquamarine = &HFF66CDAA
   MediumBlue = &HFF0000CD
   MediumOrchid = &HFFBA55D3
   MediumPurple = &HFF9370DB
   MediumSeaGreen = &HFF3CB371
   MediumSlateBlue = &HFF7B68EE
   MediumSpringGreen = &HFF00FA9A
   MediumTurquoise = &HFF48D1CC
   MediumVioletRed = &HFFC71585
   MidnightBlue = &HFF191970
   MintCream = &HFFF5FFFA
   MistyRose = &HFFFFE4E1
   Moccasin = &HFFFFE4B5
   NavajoWhite = &HFFFFDEAD
   Navy = &HFF000080
   OldLace = &HFFFDF5E6
   Olive = &HFF808000
   OliveDrab = &HFF6B8E23
   Orange = &HFFFFA500
   OrangeRed = &HFFFF4500
   Orchid = &HFFDA70D6
   PaleGoldenrod = &HFFEEE8AA
   PaleGreen = &HFF98FB98
   PaleTurquoise = &HFFAFEEEE
   PaleVioletRed = &HFFDB7093
   PapayaWhip = &HFFFFEFD5
   PeachPuff = &HFFFFDAB9
   Peru = &HFFCD853F
   Pink = &HFFFFC0CB
   Plum = &HFFDDA0DD
   PowderBlue = &HFFB0E0E6
   Purple = &HFF800080
   Red = &HFFFF0000
   RosyBrown = &HFFBC8F8F
   RoyalBlue = &HFF4169E1
   SaddleBrown = &HFF8B4513
   Salmon = &HFFFA8072
   SandyBrown = &HFFF4A460
   SeaGreen = &HFF2E8B57
   SeaShell = &HFFFFF5EE
   Sienna = &HFFA0522D
   Silver = &HFFC0C0C0
   SkyBlue = &HFF87CEEB
   SlateBlue = &HFF6A5ACD
   SlateGray = &HFF708090
   Snow = &HFFFFFAFA
   SpringGreen = &HFF00FF7F
   SteelBlue = &HFF4682B4
   Tan = &HFFD2B48C
   Teal = &HFF008080
   Thistle = &HFFD8BFD8
   Tomato = &HFFFF6347
   Transparent = &HFFFFFF
   Turquoise = &HFF40E0D0
   Violet = &HFFEE82EE
   Wheat = &HFFF5DEB3
   White = &HFFFFFFFF
   WhiteSmoke = &HFFF5F5F5
   Yellow = &HFFFFFF00
   YellowGreen = &HFF9ACD32
End Enum

    ' Smototing & Rendering
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMode As SmoothingMode) As Long

Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, ByRef lpOutput As GdiplusStartupOutput) As Long

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef graphics As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal color As Colors, ByVal Width As Single, ByVal unit As GpUnit, ByRef pen As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal color As Colors, ByRef brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As Long

Private Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal path As Long) As Long

Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As Long

Private Declare Function GdipDrawPie Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipFillPie Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long

Private m_Token As Long
Private Function ShutdownGDIPlus() As Long
    ShutdownGDIPlus = GdiplusShutdown(m_Token)
End Function
Private Sub Combo1_Click()
On Error GoTo Error

    Dim fso, d, s
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = fso.GetDrive(fso.GetDriveName(fso.GetAbsolutePathName(Combo1.Text)))
    m_Total = Round(d.FreeSpace / Int(d.TotalSize / 360), 2)
    m_Free = 360 - Round(d.FreeSpace / Int(d.TotalSize / 360), 2)
    Text1.Text = FormatNumber(d.TotalSize / 1024, 0)
    Text2.Text = FormatNumber(d.FreeSpace / 1024, 0)
    Text3.Text = FormatNumber((d.TotalSize - d.FreeSpace) / 1024, 0)
    
    Dim m_Graphics As Long
    Dim m_Brush_01 As Long
    Dim m_Brush_02 As Long
    Dim m_Brush_03 As Long
    Dim m_Brush_04 As Long
    Dim m_Pen_01 As Long
    Dim m_Pen_02 As Long
    Dim m_Width As Single
    Dim m_Height As Single
    Dim m_Width_m As Single
    Dim m_Height_m As Single

    m_Width = 200
    m_Height = 90
    m_Width_m = Picture1.ScaleWidth / 2
    m_Height_m = Picture1.ScaleHeight / 2
    
    Picture1.Cls
    Call GdipCreateFromHDC(Picture1.hDC, m_Graphics)
    Call GdipSetSmoothingMode(m_Graphics, SmoothingModeAntiAlias8x4)
    
    Call GdipCreatePen1(Black, 2, UnitPixel, m_Pen_01)
    Call GdipDrawPie(m_Graphics, m_Pen_01, m_Width_m - m_Width / 2, m_Height_m - m_Height / 2 + 4, m_Width, m_Height, 0, m_Free)
    Call GdipCreateSolidFill(DarkBlue, m_Brush_03)
    Call GdipFillPie(m_Graphics, m_Brush_03, m_Width_m - m_Width / 2, m_Height_m - m_Height / 2 + 4, m_Width, m_Height, 0, m_Free)
    
    Call GdipCreatePen1(Black, 1, UnitPixel, m_Pen_02)
    Call GdipDrawPie(m_Graphics, m_Pen_02, m_Width_m - m_Width / 2, m_Height_m - m_Height / 2 + 4, m_Width, m_Height, m_Free, m_Total)
    Call GdipCreateSolidFill(DarkMagenta, m_Brush_04)
    Call GdipFillPie(m_Graphics, m_Brush_04, m_Width_m - m_Width / 2, m_Height_m - m_Height / 2 + 4, m_Width, m_Height, m_Free, m_Total)
       
    Call GdipCreatePen1(Black, 2, UnitPixel, m_Pen_01)
    Call GdipDrawPie(m_Graphics, m_Pen_01, m_Width_m - m_Width / 2, m_Height_m - m_Height / 2, m_Width, m_Height, 0, m_Free)
    Call GdipCreateSolidFill(Blue, m_Brush_01)
    Call GdipFillPie(m_Graphics, m_Brush_01, m_Width_m - m_Width / 2, m_Height_m - m_Height / 2, m_Width, m_Height, 0, m_Free)
    
    Call GdipCreatePen1(Black, 1, UnitPixel, m_Pen_02)
    Call GdipDrawPie(m_Graphics, m_Pen_02, m_Width_m - m_Width / 2, m_Height_m - m_Height / 2, m_Width, m_Height, m_Free, m_Total)
    Call GdipCreateSolidFill(Magenta, m_Brush_02)
    Call GdipFillPie(m_Graphics, m_Brush_02, m_Width_m - m_Width / 2, m_Height_m - m_Height / 2, m_Width, m_Height, m_Free, m_Total)
  
    Call GdipDeletePen(m_Pen_01)
    Call GdipDeletePen(m_Pen_02)
    Call GdipDeleteBrush(m_Brush_01)
    Call GdipDeleteBrush(m_Brush_02)
    Call GdipDeleteBrush(m_Brush_03)
    Call GdipDeleteGraphics(m_Graphics)

Error:
End Sub


Private Sub Form_Load()

    Dim fso, d, dc
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dc = fso.Drives
    For Each d In dc
        If d.DriveType = 2 Then
            Combo1.AddItem d
        End If
    Next
    Combo1.ListIndex = 0
    
    With Picture1
        .Left = 9
        .Top = 10
        .Height = 120
        .Width = 222
    End With
    
    With Combo1
        .Left = 9
        .Top = 140
        .Width = 222
    End With
    
    Call StartUpGDIPlus(GdiPlusVersion)
    Combo1_Click

    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call ShutdownGDIPlus
End Sub
Private Function StartUpGDIPlus(ByVal GdipVersion As Long) As Long
    Dim GdipStartupInput As GDIPlusStartupInput
    Dim GdipStartupOutput As GdiplusStartupOutput
    GdipStartupInput.GdiPlusVersion = GdipVersion
    StartUpGDIPlus = GdiplusStartup(m_Token, GdipStartupInput, GdipStartupOutput)
End Function

