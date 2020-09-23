VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OBJphysic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author : Creator Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'
Public Npoints
Private Point() As tPoint

Public Nlinks
Private Link() As tLink

Public NMuscles
Private MUSCLE() As tMuscle


Public DefaultStrength


Public Sub SetPointVelXY(p, vx, vy)
Point(p).vx = vx
Point(p).vy = vy

End Sub

Public Function ADDpoint(x, y)

Npoints = Npoints + 1
ReDim Preserve Point(Npoints)

Point(Npoints).x = x
Point(Npoints).y = y

End Function

Public Function ADDLink(P1, P2)
Nlinks = Nlinks + 1
ReDim Preserve Link(Nlinks)

Link(Nlinks).P1 = P1
Link(Nlinks).P2 = P2
Link(Nlinks).MainL = Distance(Point(P1), Point(P2))
Debug.Print Nlinks & "   " & Link(Nlinks).MainL
End Function

Public Function ADDMuscle(L1, L2, Strength)
NMuscles = NMuscles + 1
ReDim Preserve MUSCLE(NMuscles)

With MUSCLE(NMuscles)
    .L1 = L1
    .L2 = L2
    .f = Strength
    
    If Link(L1).P1 = Link(L2).P1 Then
        .P0 = Link(L1).P1
        .P1 = Link(L1).P2
        .P2 = Link(L2).P2
    End If
    
    If Link(L1).P1 = Link(L2).P2 Then
        .P0 = Link(L1).P1
        .P1 = Link(L1).P2
        .P2 = Link(L2).P1
    End If
    
    If Link(L1).P2 = Link(L2).P1 Then
        .P0 = Link(L1).P2
        .P1 = Link(L1).P1
        .P2 = Link(L2).P2
    End If
    
    If Link(L1).P2 = Link(L2).P2 Then
        .P0 = Link(L1).P2
        .P1 = Link(L1).P1
        .P2 = Link(L2).P1
    End If
    'Stop
    
    .MainA = Atan2(Point(.P1).x - Point(.P0).x, Point(.P1).y - Point(.P0).y) - _
            Atan2(Point(.P2).x - Point(.P0).x, Point(.P2).y - Point(.P0).y)
    
    
End With

End Function




Public Sub DRAW(ByRef p As PictureBox, Optional Extended = False)
Dim X1
Dim X2
Dim Y1
Dim Y2
'Stop

p.Circle (Point(MUSCLE(MUS).P0).x, (Point(MUSCLE(MUS).P0).y)), 5, vbRed
p.Line (Point(MUSCLE(MUS).P1).x, (Point(MUSCLE(MUS).P1).y))- _
        (Point(MUSCLE(MUS).P2).x, (Point(MUSCLE(MUS).P2).y)), vbRed


For i = 1 To Nlinks
    X1 = Point(Link(i).P1).x
    Y1 = Point(Link(i).P1).y
    X2 = Point(Link(i).P2).x
    Y2 = Point(Link(i).P2).y
    p.Line (X1, Y1)-(X2, Y2), vbGreen
    
    
    If Extended Then
        p.Circle (X1, Y1), 2, vbYellow
        p.CurrentX = X1
        p.CurrentY = Y1 - 10
        p.ForeColor = vbYellow
        p.Print Link(i).P1
        
        p.Circle (X2, Y2), 2, vbYellow
        p.CurrentX = X2
        p.CurrentY = Y2 - 10
        p.Print Link(i).P2
        
        X1 = (X1 + X2) / 2
        Y1 = (Y1 + Y2) / 2
        p.CurrentX = X1
        p.CurrentY = Y1 - 10
        p.ForeColor = vbGreen
        p.Print "-" & i & "-"
    End If
    
Next i

End Sub

Public Sub DRAW_DX()
For i = 1 To Nlinks
    
    
    With Sprite(i)
        
        .Ang = LINK_GetWorldAngle(i)
        .DrawCenter = .TexCenter
        .POS.x = (Point(Link(i).P1).x + Point(Link(i).P2).x) / 2
        .POS.y = (Point(Link(i).P1).y + Point(Link(i).P2).y) / 2
        .DrawPos = vSUB(.POS, .DrawCenter)
        
        dSprite.DRAW .Tex, ByVal 0, .DrawScala, .DrawCenter, -.Ang, .DrawPos, MyColor
        
    End With
    
    
Next

End Sub
Sub doPHYSICS()



For i = 1 To Npoints
    With Point(i)
        .OldX = .x
        .OldY = .y
    End With
Next


DoForces
DoMuscles


''''Update X Y   and Collision
For i = 1 To Npoints
    With Point(i)
        .vx = .vx * AIR
        .vy = .vy * AIR
        .vy = .vy + Gravity
        .x = .x + .vx
        .y = .y + .vy
        
        If .y > 437 Then
            'This is one of the most difficult point.
            'should Bounce, but doesnt.
            '(Do tests even with 0 muscles)
            .y = .OldY
            .vy = -.vy * 1.05 ' 0.9
            .vx = .vx * 0.5
        End If
        
        'If .y < 50 Then: .y = 50: .vx = 0: .vy = .vy = 0
        
        If .x < 3 Or .x > 421 Then
            .x = .OldX
            .vx = -.vx * 0.9
        End If
    End With
Next i

End Sub

Sub DoForces()
'VERLET
Dim Mx As Double
Dim My As Double

Dim P1x As Double
Dim P1y As Double
Dim P2x As Double
Dim P2y As Double
Dim A As Double


For L = 1 To Nlinks
    With Link(L)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Mx = (Point(.P1).x + Point(.P2).x) / 2
        My = (Point(.P1).y + Point(.P2).y) / 2
        
        'mx,my middle point of link L
        
        A = Atan2(Point(.P2).x - Point(.P1).x, Point(.P2).y - Point(.P1).y)
        
        'To preserve Link Lenght
        P1x = Mx - Cos(A) * .MainL / 2
        P1y = My - Sin(A) * .MainL / 2
        P2x = Mx + Cos(A) * .MainL / 2
        P2y = My + Sin(A) * .MainL / 2
        
        'Find Velocities
        Point(.P1).vx = Point(.P1).vx + (P1x - Point(.P1).x)
        Point(.P1).vy = Point(.P1).vy + (P1y - Point(.P1).y)
        Point(.P2).vx = Point(.P2).vx + (P2x - Point(.P2).x)
        Point(.P2).vy = Point(.P2).vy + (P2y - Point(.P2).y)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End With
Next L
End Sub

Sub DoMuscles()
Dim X1 As Double
Dim Y1 As Double
Dim X2 As Double
Dim Y2 As Double
Dim AM As Double
Dim Dx1 As Double
Dim Dy1 As Double
Dim Dx2 As Double
Dim Dy2 As Double
'For m = NMuscles To 1 Step -1
For M = 1 To NMuscles
    
    With MUSCLE(M)
        a1 = Atan2(Point(.P1).x - Point(.P0).x, Point(.P1).y - Point(.P0).y)
        a2 = Atan2(Point(.P2).x - Point(.P0).x, Point(.P2).y - Point(.P0).y)
        
        AM = (a1 - a2) - .MainA
        
        
        X1 = Point(.P0).x + Cos(a1 - AM) * Link(.L1).MainL
        Y1 = Point(.P0).y + Sin(a1 - AM) * Link(.L1).MainL
        X2 = Point(.P0).x + Cos(a2 + AM) * Link(.L2).MainL
        Y2 = Point(.P0).y + Sin(a2 + AM) * Link(.L2).MainL
        
        Dx1 = (Point(.P1).x - X1)
        Dy1 = (Point(.P1).y - Y1)
        Dx2 = (Point(.P2).x - X2)
        Dy2 = (Point(.P2).y - Y2)
        
        
        If frmPHYS.chLIMIT.Value = Checked Then
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Limit Muscles Speed
            'Not "Right", but works quite good
            
            Dx1 = IIf(Abs(Dx1) > kMuscleSpeedLimit, Sgn(Dx1) * kMuscleSpeedLimit, Dx1)
            Dy1 = IIf(Abs(Dy1) > kMuscleSpeedLimit, Sgn(Dy1) * kMuscleSpeedLimit, Dy1)
            Dx2 = IIf(Abs(Dx2) > kMuscleSpeedLimit, Sgn(Dx2) * kMuscleSpeedLimit, Dx2)
            Dy2 = IIf(Abs(Dy2) > kMuscleSpeedLimit, Sgn(Dy2) * kMuscleSpeedLimit, Dy2)
            '
            'D = Sqr(Dx1 * Dx1 + Dy1 * Dy1)
            'If D > kMuscleSpeedLimit Then
            'Dx1 = Sgn(Dx1) * kMuscleSpeedLimit
            'Dy1 = Sgn(Dy1) * kMuscleSpeedLimit
            'Dx2 = Sgn(Dx2) * kMuscleSpeedLimit
            'Dy2 = Sgn(Dy2) * kMuscleSpeedLimit
            'End If
            '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
        
        Point(.P1).vx = Point(.P1).vx - .f * Dx1
        Point(.P1).vy = Point(.P1).vy - .f * Dy1
        Point(.P2).vx = Point(.P2).vx - .f * Dx2
        Point(.P2).vy = Point(.P2).vy - .f * Dy2
        
        Point(.P0).vx = Point(.P0).vx + .f * (Dx1 + Dx2)
        Point(.P0).vy = Point(.P0).vy + .f * (Dy1 + Dy2)
        
    End With
    
Next

End Sub


Public Function MUSCLE_AddToMainANG(M, ValADD)
MUSCLE(M).MainA = MUSCLE(M).MainA + ValADD
End Function

Public Property Let MUSCLE_SetStrength(M, Strength)
MUSCLE(M).f = Strength
End Property

Public Property Get MUSCLE_MainANG(M)
MUSCLE_MainANG = MUSCLE(M).MainA
End Property

Public Property Let MUSCLE_MainANG(M, Value)
MUSCLE(M).MainA = Value
End Property

Public Property Get PointVX(p)
PointVX = Point(p).vx
End Property
Public Property Get PointVY(p)
PointVY = Point(p).vy
End Property
Public Property Get PointX(p)
PointX = Point(p).x
End Property
Public Property Get PointY(p)
PointY = Point(p).y
End Property

Public Property Let PointVX(p, V)
Point(p).vx = V
End Property
Public Property Let PointVY(p, V)
Point(p).vy = V
End Property
Public Property Let PointX(p, V)
Point(p).x = V
End Property
Public Property Let PointY(p, V)
Point(p).y = V
End Property

Public Function LINK_GetWorldAngle(L) As Double

Dim P1 As Integer
Dim P2 As Integer

P1 = Link(L).P1
P2 = Link(L).P2

LINK_GetWorldAngle = Atan2(Point(P2).x - Point(P1).x, Point(P2).y - Point(P1).y)


End Function

Public Sub Obj_SAVE(Optional Filename = "Obj.Txt")
Dim Cx As Double
Dim Cy As Double

Cx = 0
Cy = 0
For p = 1 To Npoints
    With Point(p)
        Cx = Cx + .x
        Cy = Cy + .y
    End With
Next p
Cx = Cx / Npoints
Cy = Cy / Npoints

Open App.Path & "\" & Filename For Output As 1

Print #1, Npoints
Print #1, Nlinks
Print #1, NMuscles

Print #1, "Points:"
For p = 1 To Npoints
    Print #1, Replace(Point(p).x - Cx, ",", ".")
    Print #1, Replace(Point(p).y - Cy, ",", ".")
Next

Print #1, "Links:"
For L = 1 To Nlinks
    Print #1, Link(L).P1
    Print #1, Link(L).P2
Next

Print #1, "Muscles:"
For M = 1 To NMuscles
    Print #1, MUSCLE(M).L1
    Print #1, MUSCLE(M).L2
    Print #1, Replace(MUSCLE(M).f, ",", ".")
Next

Close 1

End Sub

Public Sub Obj_LOADandPlace(Filename As String, CenterX, CenterY)
Dim S As String
Dim V1 As Double
Dim V2 As Double
Dim V3 As Double

Dim nP As Integer
Dim nL As Integer
Dim nM As Integer

Open App.Path & "\" & Filename For Input As 1

Input #1, nP 'Npoints
Input #1, nL 'Nlinks
Input #1, nM 'NMuscles

Npoints = 0
Nlinks = 0
NMuscles = 0


Input #1, S
For p = 1 To nP
    Input #1, V1 'Point(p).x
    Input #1, V2 'Point(p).y
    ADDpoint V1 + CenterX, V2 + CenterY
Next

Input #1, S
For L = 1 To nL
    Input #1, V1 'Link(l).P1
    Input #1, V2 'Link(l).P2
    ADDLink V1, V2
Next

Input #1, S
For M = 1 To nM
    Input #1, V1 'MUSCLE(M).L1
    Input #1, V2 'MUSCLE(M).L2
    Input #1, V3 'MUSCLE(M).F
    ADDMuscle V1, V2, V3
Next

Close 1

End Sub

Public Sub OBJ_SavePose(Filename As String)
Open App.Path & "\" & Filename For Output As 1
Print #1, NMuscles
For i = 1 To NMuscles
    Print #1, Replace(CStr(MUSCLE(i).MainA), ",", ".")
Next i

Close 1

End Sub
Public Sub OBJ_LoadPose(Filename As String)
Open App.Path & "\" & Filename For Input As 1
Input #1, nM
If nM <> NMuscles Then Close 1: Exit Sub
For i = 1 To NMuscles
    Input #1, MUSCLE(i).MainA
Next i

Close 1

End Sub
