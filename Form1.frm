VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Terrain 3D - http://www.andreavb.com"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkWireFrame 
      Caption         =   "View Wire Frame"
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      Top             =   60
      Value           =   1  'Checked
      Width           =   2115
   End
   Begin VB.CheckBox chkRoll 
      Caption         =   "Enable Roll"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   60
      Width           =   1875
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "New Terrain"
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   60
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Const WATER_LEVEL = -5

Const MAX_X = 64
Const MAX_Y = 64
Private Type ScreenCoord
    X As Single
    Y As Single
End Type
Private Type Coord3D
    X As Single
    Y As Single
    Z As Single
End Type
Private Type HPoint
    H As Single
    bRecurring As Boolean
End Type

Dim pan As Coord3D
Dim altitude(0 To MAX_X - 1, 0 To MAX_Y - 1) As HPoint
Dim yaw As Single
Dim pitch As Single
Dim roll As Single
Dim FocalDistance As Single

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As PointApi, ByVal nCount As Long) As Long

Private Type PointApi   ' pt
    X As Long
    Y As Long
End Type

Private Sub CalculateMidPoints(x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    Dim xm As Long
    Dim ym As Long
    Dim d As Single
    Dim Y As Long
    
    xm = (x1 + x2 + 1) \ 2
    ym = (y1 + y2 + 1) \ 2
    If ym > y1 And ym < y2 And xm > x1 And xm < x2 And Not altitude(xm, ym).bRecurring Then
        d = Sqr((x2 - x1 + 1) ^ 2 + (y2 - y1 + 1) ^ 2) / 2
        altitude(xm, ym).H = (altitude(x1, y1).H + altitude(x1, y2).H + altitude(x2, y1).H + altitude(x2, y2).H) / 4 + d / 2 - Rnd * (d)
        altitude(xm, ym).bRecurring = True
        d = (x2 - x1 + 1) / 2
        If x1 <> 0 And y1 <> 0 And x2 <> MAX_Y - 1 And y2 <> MAX_Y - 1 Then
            altitude(x1, ym).H = (altitude(x1, y1).H + altitude(x1, y2).H) / 2 + d / 2 - Rnd * (d)
            altitude(x1, ym).bRecurring = True
            altitude(x2, ym).H = (altitude(x2, y1).H + altitude(x2, y2).H) / 2 + d / 2 - Rnd * (d)
            altitude(x2, ym).bRecurring = True
            altitude(xm, y1).H = (altitude(x1, y1).H + altitude(x2, y1).H) / 2 + d / 2 - Rnd * (d)
            altitude(xm, y1).bRecurring = True
            altitude(xm, y2).H = (altitude(x1, y2).H + altitude(x2, y2).H) / 2 + d / 2 - Rnd * (d)
            altitude(xm, y2).bRecurring = True
        End If
        CalculateMidPoints x1, y1, xm, ym
        CalculateMidPoints x1, ym, xm, y2
        CalculateMidPoints xm, y1, x2, ym
        CalculateMidPoints xm, ym, x2, y2
    End If
End Sub

Private Sub GenerateTerrain()
    Dim X As Integer
    Dim Y As Integer
    Dim i As Integer
    Dim d As Single

    'Reset array
    For X = 0 To MAX_X - 1
        For Y = 0 To MAX_Y - 1
            altitude(X, Y).H = 0
            altitude(X, Y).bRecurring = False
        Next
    Next
    CalculateMidPoints 0, 0, MAX_X - 1, MAX_Y - 1
    'smooth 3D surface
    For i = 1 To 2
        For X = 1 To MAX_X - 2
            For Y = 1 To MAX_Y - 2
                altitude(X, Y).H = (altitude(X - 1, Y - 1).H + altitude(X + 1, Y - 1).H + altitude(X - 1, Y + 1).H + altitude(X + 1, Y + 1).H + altitude(X, Y).H * 2) / 6 '+ (d / 2 - Rnd * (d / 2))
            Next
        Next
    Next
    'water level
    For X = 1 To MAX_X - 2
        For Y = 1 To MAX_Y - 2
            If altitude(X, Y).H < WATER_LEVEL Then altitude(X, Y).H = WATER_LEVEL
        Next
    Next
    DrawTerrain
End Sub

Private Sub chkRoll_Click()
    If chkRoll.Value = vbUnchecked Then
        roll = 0
        DrawTerrain
    End If
End Sub

Private Sub chkWireFrame_Click()
    DrawTerrain
End Sub

Private Sub cmdGenerate_Click()
    GenerateTerrain
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim Y As Integer
    
    Randomize Timer
    yaw = 0.88
    pitch = -0.6
    'set default pan
    pan.X = 0
    pan.Y = 0
    pan.Z = MAX_Y * 1.5
    'set focal distance
    FocalDistance = 8000
    'Me.AutoRedraw = True
    GenerateTerrain
End Sub

Private Sub DrawTerrain()
    Dim X As Long
    Dim Y As Long
    Dim p2(0 To MAX_X - 1, 0 To MAX_Y - 1) As ScreenCoord
    Dim mp As Coord3D
    Dim mc As Long
    Dim minh As Single
    Dim maxh As Single
    Dim point(0 To 5) As PointApi
    
    Cls
    'calculate screen coordinates of land points
    For X = 0 To MAX_X - 1
        For Y = 0 To MAX_Y - 1
            mp.X = X
            mp.Y = altitude(X, Y).H
            mp.Z = Y
            p2(X, Y) = PointToScreen(mp, pan, FocalDistance, yaw, pitch, roll)
            If altitude(X, Y).H < minh Then minh = altitude(X, Y).H
            If altitude(X, Y).H > maxh Then maxh = altitude(X, Y).H
        Next
    Next
    'Me.AutoRedraw = False
    'draw every polygon
    For X = 0 To MAX_X - 2
        For Y = MAX_Y - 2 To 0 Step -1
            If altitude(X, Y).H = WATER_LEVEL And altitude(X + 1, Y).H = WATER_LEVEL And altitude(X + 1, Y + 1).H = WATER_LEVEL And altitude(X, Y + 1).H = WATER_LEVEL Then
                'set water color to blue
                mc = RGB(80, 80, 210 + Rnd * 30)
                Me.ForeColor = mc
            Else
                'set land color
                mc = RGB(40, 50 + 205 * (altitude(X, Y).H - minh) / (maxh - minh), 10)
                If chkWireFrame.Value = vbChecked Then
                    Me.ForeColor = vbBlack
                Else
                    Me.ForeColor = mc
                End If
            End If
            Me.FillColor = mc
            Me.FillStyle = vbSolid
            point(0).X = Me.ScaleX(p2(X, Y).X, vbTwips, vbPixels)
            point(0).Y = Me.ScaleY(p2(X, Y).Y, vbTwips, vbPixels)
            point(1).X = Me.ScaleX(p2(X, Y + 1).X, vbTwips, vbPixels)
            point(1).Y = Me.ScaleY(p2(X, Y + 1).Y, vbTwips, vbPixels)
            point(2).X = Me.ScaleX(p2(X + 1, Y + 1).X, vbTwips, vbPixels)
            point(2).Y = Me.ScaleY(p2(X + 1, Y + 1).Y, vbTwips, vbPixels)
            point(3).X = Me.ScaleX(p2(X + 1, Y).X, vbTwips, vbPixels)
            point(3).Y = Me.ScaleY(p2(X + 1, Y).Y, vbTwips, vbPixels)
            point(4).X = point(0).X
            point(4).Y = point(0).Y
            Polygon Me.hdc, point(0), 4
        Next
    Next
    'Me.AutoRedraw = True
    Me.CurrentX = 0
    Me.CurrentY = 0
    Me.ForeColor = vbBlack
    Print "yaw=" & yaw
    Print "pitch=" & pitch
    Print "roll=" & roll
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static oldy As Single
    Static oldx As Single

    If oldy <> 0 And oldx <> 0 And Button = vbLeftButton Then
        pitch = pitch - (Y - oldy) / 1000
        yaw = yaw - (X - oldx) / 1000
        DrawTerrain
    End If
    If oldy <> 0 And oldx <> 0 And Button = vbRightButton Then
        If chkRoll.Value = vbChecked Then roll = roll + (X - oldx) / 1000
        FocalDistance = FocalDistance + (Y - oldy)
        DrawTerrain
    End If
    oldy = Y
    oldx = X
End Sub

Private Sub Form_Resize()
    Dim sx As Long
    
    DrawTerrain
    'move controls
    On Error Resume Next
    sx = Me.ScaleWidth / 4
    cmdGenerate.Move sx, 30, sx * 0.9, 315
    chkRoll.Move sx * 2, 30, sx * 0.9, 315
    chkWireFrame.Move sx * 3, 30, sx * 0.9, 315
End Sub

'convert 3d coordinates to screen coordinates
Private Function PointToScreen(p As Coord3D, pan As Coord3D, ByVal FocalDistance As Single, ByVal yaw As Single, ByVal pitch As Single, ByVal roll As Single) As ScreenCoord
    Dim np1 As Coord3D
    Dim np2 As Coord3D
    
    'apply pan to center 3d grid to view position
    np2.X = p.X - MAX_X / 2
    np2.Z = p.Z - MAX_Y / 2
    np2.Y = p.Y
   
    np1.X = np2.Z * Sin(yaw) + np2.X * Cos(yaw)
    np1.Y = np2.Y
    np1.Z = np2.Z * Cos(yaw) - np2.X * Sin(yaw)
    
    np2.X = np1.X
    np2.Y = np1.Y * Cos(pitch) - np1.Z * Sin(pitch)
    np2.Z = np1.Y * Sin(pitch) + np1.Z * Cos(pitch)
    
    np1.X = np2.Y * Sin(roll) + np2.X * Cos(roll)
    np1.Y = np2.Y * Cos(roll) - np2.X * Sin(roll)
    np1.Z = np2.Z
    
    np1.X = np1.X + pan.X
    np1.Y = np1.Y + pan.Y
    np1.Z = np1.Z + pan.Z
    
    If np1.Z <> 0 Then
        PointToScreen.X = np1.X * (FocalDistance) / np1.Z + Me.Width / 2
        PointToScreen.Y = -np1.Y * (FocalDistance) / np1.Z + Me.Height / 2
    End If
End Function


