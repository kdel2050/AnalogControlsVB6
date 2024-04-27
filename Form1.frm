VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   271
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar scrAxonesY 
      Height          =   1830
      LargeChange     =   2
      Left            =   4470
      Max             =   300
      TabIndex        =   7
      Top             =   1365
      Width           =   195
   End
   Begin VB.HScrollBar scrAxonesX 
      Height          =   210
      LargeChange     =   2
      Left            =   1875
      Max             =   300
      TabIndex        =   6
      Top             =   3465
      Width           =   2400
   End
   Begin VB.VScrollBar scrTimer 
      Height          =   3285
      LargeChange     =   100
      Left            =   5865
      Max             =   2000
      TabIndex        =   5
      Top             =   570
      Value           =   50
      Width           =   180
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      ClipControls    =   0   'False
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1800
      Left            =   1890
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   3
      Top             =   1380
      Width           =   2400
   End
   Begin VB.VScrollBar scrRadius 
      Height          =   1950
      LargeChange     =   10
      Left            =   450
      Max             =   500
      Min             =   1
      TabIndex        =   2
      Top             =   90
      Value           =   100
      Width           =   165
   End
   Begin VB.VScrollBar scrSmallRadius 
      Height          =   1440
      LargeChange     =   3
      Left            =   90
      Max             =   50
      Min             =   1
      TabIndex        =   1
      Top             =   75
      Value           =   1
      Width           =   180
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   180
      LargeChange     =   5
      Left            =   3030
      Max             =   180
      Min             =   1
      TabIndex        =   0
      Top             =   75
      Value           =   70
      Width           =   3120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5340
      Top             =   2460
   End
   Begin VB.Label lbDegree 
      Height          =   465
      Left            =   2070
      TabIndex        =   4
      Top             =   465
      Width           =   810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type COORD
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Private Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long, pcRect As RECT, ByVal un As Long) As Long
'Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
'Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Const ALTERNATE = 1 ' ALTERNATE and WINDING are
Const WINDING = 2 ' constants for FillMode.
Const BLACKBRUSH = 4 ' Constant for brush type.
Const DT_CENTER = &H1
Const DC_GRADIENT = &H20          'Only Win98/2000 !!
Dim Radius As Integer, smallRadius As Integer
Const PI = 3.14159265358979
Const degToRad = 1.74532925199432E-02
Dim degree As Integer
Dim rad As Double
Dim poly(1 To 3) As COORD, NumCoords As Long, wBrush As Long, yBrush As Long, rBrush As Long, gBrush As Long, hRgn As Long
Dim axones As COORD
Dim R As RECT
Dim a, b, c As Long
Dim CX, CY As Single
Dim txtW, txtH As Single
    

Private Sub Form_Load()
    degree = 0
    rad = 0
    axones.x = Pic1.ScaleWidth / 2
    axones.y = Pic1.ScaleHeight
    'axones.x = 300
    'axones.y = 300
    NumCoords = 3
    Radius = Pic1.ScaleHeight / 2 - 10
    smallRadius = 5
    ' Set scalemode to pixels to set up points of triangle.
    Pic1.ScaleMode = vbPixels
    'Pic1.ScaleLeft = 5000
    'Pic1.ScaleTop = 5000
    'Pic1.ScaleHeight = -100
    'hBrush = GetStockObject(BLACKBRUSH)
    gBrush = CreateSolidBrush(vbGreen)
    yBrush = CreateSolidBrush(vbYellow)
    rBrush = CreateSolidBrush(vbRed)
    wBrush = CreateSolidBrush(vbWhite)
    CX = Pic1.ScaleWidth     ' Set X position.
    CY = Pic1.ScaleHeight  ' Set Y position
    
    txtW = Pic1.TextWidth("180")
    txtH = Pic1.TextHeight("180")
    
    Call scrTimer_Change
End Sub

Private Sub Form_Paint()
    
    'MsgBox Me.Width & " " & Me.Height
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DeleteObject hRgn
    DeleteObject yBrush
    DeleteObject rBrush
    DeleteObject gBrush
    DeleteObject wBrush
    DeleteObject hBrush
    DeleteObject NumCoords
End Sub


Private Sub scrAxonesX_Change()
    axones.x = scrAxonesX.Value
    CX = axones.x * 2
End Sub

Private Sub scrAxonesX_Scroll()
    axones.x = scrAxonesX.Value
    CX = axones.x * 2
    'Giati otan ginetai Form_Load isxyoun:
    'axones.x = Pic1.ScaleWidth / 2
    'axones.y = Pic1.ScaleHeight
    'CX = Pic1.ScaleWidth
    'CY = Pic1.ScaleHeight
End Sub

Private Sub scrAxonesY_Change()
    axones.y = scrAxonesY.Value
    CY = axones.y
End Sub


Private Sub scrRadius_Change()
    Radius = scrRadius.Value
End Sub

Private Sub scrSmallRadius_Change()
    smallRadius = scrSmallRadius.Value
End Sub

Private Sub scrTimer_Change()
    Timer1.Interval = scrTimer.Value
End Sub

Private Sub scrTimer_Scroll()
    Timer1.Interval = scrTimer.Value
End Sub

Private Sub Timer1_Timer()
    'Radius = HScroll1.Value
    'degree = HScroll1.Value
    Pic1.Cls
    'Pic1
    ' Number of vertices in polygon.
    'SetRect R, axones.x - Radius - TextWidth("0"), axones.y - TextHeight("0"), axones.x - Radius, axones.y
    'MsgBox TextWidth("0") & " " & TextHeight("0")
    'SetRect R, axones.x - 20, axones.y - 20, axones.x + 20, axones.y - 20
    'SetRect R, CLng(axones.x - Radius - 20), CLng(axones.y - 20), CLng(axones.x - Radius), CLng(axones.y)
    'MsgBox CStr(axones.x - Radius - 20)
    'MsgBox CStr(axones.y + 20)
    'MsgBox CStr(axones.x - Radius)
    'MsgBox CStr(axones.y)
    'SetRect R, 0, 220 - TextWidth("0"), 300 - 12, 240, 300
    
    
    'SetRect R, 0, a, b, c
    'OffsetRect R, 0, 22
    'DrawFocusRect Pic1.hdc, R
    'DrawText Pic1.hdc, "0", Len("0"), R, DT_LEFT
    
    'OffsetRect R, CX / 2 - Radius * Cos(PI / 6), -Radius * Sin(PI / 6)
    'DrawFocusRect Pic1.hdc, R
    'DrawText Pic1.hdc, "30", Len("30"), R, DT_LEFT
    
     
    'SetRect R, 0, a, b, c
    'OffsetRect R, CX / 2 - Radius * Cos(PI / 3), -Radius * Sin(PI / 3)
    'DrawFocusRect Pic1.hdc, R
    'DrawText Pic1.hdc, "60", Len("60"), R, DT_LEFT
    
    'SetRect R, 0, a, b, c
    'OffsetRect R, CX / 2 - b, -Radius
    'DrawFocusRect Pic1.hdc, R
    'DrawText Pic1.hdc, "90", Len("90"), R, DT_CENTER
    
    ' Assign values to points.
    'poly(1).x = axones.x
    'poly(1).y = axones.y - smallRadius
    'poly(2).x = axones.x
    'poly(2).y = axones.y + smallRadius
    'poly(3).x = axones.x - Radius
    'poly(3).y = axones.y
    If degree = 180 Then degree = 0
    degree = degree + 1
    rad = degree * degToRad
    'rad = rad + 0.05
    poly(1).x = axones.x - smallRadius * Sin(rad)
    poly(1).y = axones.y + smallRadius * Cos(rad)
    poly(2).x = axones.x + smallRadius * Sin(rad)
    poly(2).y = axones.y - smallRadius * Cos(rad)
    poly(3).x = axones.x - Radius * Cos(rad)
    poly(3).y = axones.y - Radius * Sin(rad)
    
    '********************************************
    'poly(1).x = axones.x - smallRadius * Sin(rad)
    'poly(1).y = axones.y + smallRadius * Cos(rad)
    'poly(2).x = axones.x + smallRadius * Sin(rad)
    'poly(2).y = axones.y - smallRadius * Cos(rad)
    'poly(3).x = axones.x - Radius * Cos(rad)
    'poly(3).y = axones.y - Radius * Sin(rad)
    '********************************************
    
    ' Polygon function creates unfilled polygon on screen.
    ' Remark FillRgn statement to see results.
    Polygon Pic1.hdc, poly(1), NumCoords
    ' Gets stock black brush.
    
        'hBrush = GetStockObject(BLACKBRUSH)
    
    ' Creates region to fill with color.
    hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
    ' If the creation of the region was successful then color.
    Select Case degree
        Case 0 To 30
            'FillRgn Pic1.hdc, hRgn, rBrush
            FillRgn Pic1.hdc, hRgn, gBrush
            'Pic1.ForeColor = vbRed
        Case 30 To 60
            'FillRgn Pic1.hdc, hRgn, yBrush
            FillRgn Pic1.hdc, hRgn, gBrush
            'Pic1.ForeColor = vbYellow
        Case 60 To 120
            'FillRgn Pic1.hdc, hRgn, gBrush
            FillRgn Pic1.hdc, hRgn, gBrush
            'Pic1.ForeColor = vbGreen
        Case 120 To 150
            FillRgn Pic1.hdc, hRgn, gBrush
            'Pic1.ForeColor = vbYellow
        Case 150 To 180
            FillRgn Pic1.hdc, hRgn, gBrush
            'Pic1.ForeColor = vbRed
    End Select
        'Pic1.Circle (CX, CY), 2 * smallRadius, vbBlack
        'CX = 5
        'CY = CY - 5
        'Pic1.PSet (CX, CY), vbBlack
        Dim i As Integer, LineFactor As Single, smallLineFactor As Single
        LineFactor = Radius - 10
        smallLineFactor = Radius - 5
        Dim subdivision As Integer
        subdivision = 30
        For i = 0 To 180 Step 1
            rad = i * degToRad
            Select Case (i / subdivision)
                Case 0 '0-29
                    If i Mod subdivision = 0 Then
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - LineFactor * Cos(rad), axones.y - LineFactor * Sin(rad)), vbGreen
                        SetRect R, CX / 2 - Radius * Cos(rad) - txtW, CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad)
                        DrawText Pic1.hdc, CStr(i), Len(CStr(i)), R, DT_CENTER
                    Else
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - smallLineFactor * Cos(rad), axones.y - smallLineFactor * Sin(rad)), vbGreen
                    End If
                Case 1 '30-59
                    If i Mod subdivision = 0 Then
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - LineFactor * Cos(rad), axones.y - LineFactor * Sin(rad)), vbGreen
                        SetRect R, CX / 2 - Radius * Cos(rad) - txtW, CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad)
                        DrawText Pic1.hdc, CStr(i), Len(CStr(i)), R, DT_CENTER
                    Else
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - smallLineFactor * Cos(rad), axones.y - smallLineFactor * Sin(rad)), vbGreen
                    End If
                Case 2 '60-89
                    If i Mod subdivision = 0 Then
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - LineFactor * Cos(rad), axones.y - LineFactor * Sin(rad)), vbGreen
                        SetRect R, CX / 2 - Radius * Cos(rad) - txtW, CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad)
                        DrawText Pic1.hdc, CStr(i), Len(CStr(i)), R, DT_CENTER
                    Else
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - smallLineFactor * Cos(rad), axones.y - smallLineFactor * Sin(rad)), vbGreen
                        'If i <= 80 Then
                        '    SetRect R, CX / 2 - Radius * Cos(rad) - txtW, CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad)
                        'ElseIf i > 80 And i < 100 Then
                        '    SetRect R, CX / 2 - Radius * Cos(rad) - txtW / 2, CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad) + txtW / 2, CY - Radius * Sin(rad)
                        'Else
                        '    SetRect R, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad) + txtW, CY - Radius * Sin(rad)
                        'End If
                        'DrawText Pic1.hdc, CStr(i), Len(CStr(i)), R, DT_CENTER
                    End If
                Case 3 '90-119
                    If i Mod subdivision = 0 Then
                        'If i <= 80 Then
                        '    SetRect R, CX / 2 - Radius * Cos(rad) - txtW, CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad)
                        'ElseIf i > 80 And i < 100 Then
                            SetRect R, CX / 2 - Radius * Cos(rad) - txtW / 2, CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad) + txtW / 2, CY - Radius * Sin(rad)
                        'Else
                        '   SetRect R, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad) + txtW, CY - Radius * Sin(rad)
                        'End If
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - LineFactor * Cos(rad), axones.y - LineFactor * Sin(rad)), vbGreen
                        'SetRect R, CX / 2 - Radius * Cos(rad) - txtW, CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad)
                        DrawText Pic1.hdc, CStr(i), Len(CStr(i)), R, DT_CENTER
                    Else
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - smallLineFactor * Cos(rad), axones.y - smallLineFactor * Sin(rad)), vbGreen
                    End If
                Case 4 '120-149
                    If i Mod subdivision = 0 Then
                        SetRect R, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad) + txtW, CY - Radius * Sin(rad)
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - LineFactor * Cos(rad), axones.y - LineFactor * Sin(rad)), vbGreen
                        'SetRect R, CX / 2 - Radius * Cos(rad) - txtW, CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad)
                        DrawText Pic1.hdc, CStr(i), Len(CStr(i)), R, DT_CENTER
                    Else
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - smallLineFactor * Cos(rad), axones.y - smallLineFactor * Sin(rad)), vbGreen
                    End If
                Case 5 '150-179
                    If i Mod subdivision = 0 Then
                        SetRect R, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad) + txtW, CY - Radius * Sin(rad)
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - LineFactor * Cos(rad), axones.y - LineFactor * Sin(rad)), vbGreen
                        'SetRect R, CX / 2 - Radius * Cos(rad) - txtW, CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad)
                        DrawText Pic1.hdc, CStr(i), Len(CStr(i)), R, DT_CENTER
                    Else
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - smallLineFactor * Cos(rad), axones.y - smallLineFactor * Sin(rad)), vbGreen
                    End If
                Case 6 '180
                    If i Mod subdivision = 0 Then
                        SetRect R, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad) + txtW, CY - Radius * Sin(rad)
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - LineFactor * Cos(rad), axones.y - LineFactor * Sin(rad)), vbGreen
                        'SetRect R, CX / 2 - Radius * Cos(rad) - txtW, CY - Radius * Sin(rad) - txtH, CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad)
                        DrawText Pic1.hdc, CStr(i), Len(CStr(i)), R, DT_CENTER
                    Else
                        Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - smallLineFactor * Cos(rad), axones.y - smallLineFactor * Sin(rad)), vbGreen
                    End If
                Case Else
                    Exit Sub
            End Select
            'MsgBox i
            'rad = i * degToRad
            'Pic1.Line (CX / 2 - Radius * Cos(rad), CY - Radius * Sin(rad))-(axones.x - LineFactor * Cos(rad), axones.y - LineFactor * Sin(rad)), Pic1.ForeColor
            'Pic1.Line (CX / 2 - Radius * Cos(PI / i), CY - Radius * Sin(PI / i))-(axones.x, axones.y), vbBlack
        Next i
        'Pic1.FillColor = &H80FF&
        Pic1.Circle (CX / 2, CY), 2 * smallRadius, vbBlack
        lbDegree.Caption = degree

End Sub
