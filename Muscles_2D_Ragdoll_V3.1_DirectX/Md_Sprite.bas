Attribute VB_Name = "Md_Sprite"
Public Type tSprite
    Tex As Direct3DTexture8
    DrawCenter As D3DVECTOR2 '144'63
    DrawPos As D3DVECTOR2
    TexCenter As D3DVECTOR2
    POS As D3DVECTOR2
    POS1 As D3DVECTOR2
    
    Ang As Single
    Scala As D3DVECTOR2
    DrawScala As D3DVECTOR2
    
End Type

'Public CS As D3DVECTOR2
'Public wZOOM As D3DVECTOR2
'Public wPAN As D3DVECTOR2

Public MyColor As Long


Public Sprite() As tSprite


'Private Sub DrawSprite(ByRef SPR As tSprite)
'With SPR
'.DrawScala = vMUL(.Scala, wZOOM.x)
'.DrawCenter = vMUL(.TexCenter, .DrawScala.x) ')
'.DrawPos = vADD(vMUL(vSUB(.POS, .TexCenter), .DrawScala.x), CS)
'.DrawPos = vADD(.DrawPos, vMUL(wPAN, -wZOOM.x))
'dSprite.DRAW .Tex, ByVal 0, .DrawScala, .DrawCenter, -.Ang, .DrawPos, myColor
'End With
'End Sub


Function VEC(x As Single, y As Single) As D3DVECTOR2
VEC.x = x
VEC.y = y

End Function
Public Function vMUL(V2 As D3DVECTOR2, Val) As D3DVECTOR2
vMUL.x = V2.x * Val
vMUL.y = V2.y * Val
End Function

Public Function vADD(V1 As D3DVECTOR2, V2 As D3DVECTOR2) As D3DVECTOR2
vADD.x = V1.x + V2.x
vADD.y = V1.y + V2.y
End Function
Public Function vSUB(V1 As D3DVECTOR2, V2 As D3DVECTOR2) As D3DVECTOR2
vSUB.x = V1.x - V2.x
vSUB.y = V1.y - V2.y
End Function

