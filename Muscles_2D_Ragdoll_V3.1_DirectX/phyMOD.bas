Attribute VB_Name = "phyMOD"
'Author : Creator Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'

Public Type tPoint
    x As Double
    y As Double
    OldX As Double
    OldY As Double
    vx As Double
    vy As Double
End Type

Public Type tLink
    P1 As Integer '     Point 1
    P2 As Integer '     Point 2
    MainL As Double '   Distance Between P1 and P2
End Type

Public Type tMuscle
    L1 As Integer '     Link1
    L2 As Integer '     Link2
    MainA As Double '   Angle that should be between L1 and L2
    P0 As Integer '     Common point of L1 and L2
    P1 As Integer '     Other point on L1
    P2 As Integer '     Other point on L2
    f As Double '       Muscle Force(strength)
End Type

Public Const PI = 3.14159265358979

Public AIR As Double
Public Gravity As Double
Public kMuscleSpeedLimit

Public MUS As Integer


Public Function Distance(P1 As tPoint, P2 As tPoint) As Double
Dim DX As Double
Dim dy As Double

DX = P1.x - P2.x
dy = P1.y - P2.y

Distance = Sqr(DX * DX + dy * dy)

End Function


Public Function Atan2(ByVal x As Double, ByVal y As Double) As Double
'This Should return Angle (X is deltaX,Y is DeltaY)


Dim theta As Double

If (Abs(x) < 0.0000001) Then
    If (Abs(y) < 0.0000001) Then
        theta = 0#
    ElseIf (y > 0#) Then
        'theta = 1.5707963267949
        theta = PI / 2
    Else
        'theta = -1.5707963267949
        theta = -PI / 2
    End If
Else
    theta = Atn(y / x)
    
    If (x < 0) Then
        If (y >= 0#) Then
            theta = PI + theta
        Else
            theta = theta - PI
        End If
    End If
End If

Atan2 = theta
End Function

