VERSION 5.00
Begin VB.Form frmPHYS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Muscles Driven 2D Ragdoll Physics"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chLIMIT 
      Caption         =   "Use Muscles Speed Limiter. (More Stable)"
      Height          =   375
      Left            =   6600
      TabIndex        =   16
      Top             =   3360
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   9360
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox PIC_S 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   14
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chGravity 
      Caption         =   "Gravity"
      Height          =   255
      Left            =   6600
      TabIndex        =   12
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton SavePos 
      Caption         =   "Save Pose"
      Height          =   375
      Left            =   8640
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   8520
      TabIndex        =   10
      ToolTipText     =   "Click to load Pose"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.HScrollBar GlobStrength 
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   3840
      Width           =   2655
   End
   Begin VB.HScrollBar MuscleANG 
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   7
      Top             =   4440
      Width           =   2655
   End
   Begin VB.CheckBox ApplyMuscle 
      Caption         =   "Use Muscles"
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Only Lines"
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      Top             =   1920
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton SLN 
      Caption         =   "Show Point Link Numbers"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   120
      ScaleHeight     =   439
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   423
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
   Begin VB.Timer TIMER1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8760
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RUN"
      Height          =   615
      Left            =   8880
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Interact with the figure using the mouse, picking it up at the points and tossing it around"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6840
      Width           =   6375
   End
   Begin VB.Label mDESC 
      Caption         =   "Label3"
      Height          =   255
      Index           =   0
      Left            =   9360
      TabIndex        =   8
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Muscles Driven 2D Ragdoll Engine based on Verlet physics.  "
      Height          =   1095
      Left            =   6600
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Muscles Strength"
      Height          =   495
      Left            =   9240
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
End
Attribute VB_Name = "frmPHYS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author : Creator Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'

Dim Doll As New OBJphysic

Dim doMouse As Boolean
Dim PtoMove As Integer
Dim mouseX As Single
Dim mouseY As Single



Private Sub ApplyMuscle_Click()
If ApplyMuscle.Value = Checked Then
    GlobStrength = GlobStrength.Max
Else
    GlobStrength = GlobStrength.Min
End If


End Sub

Private Sub chGravity_Click()
If chGravity.Value = Checked Then
    Gravity = 0.035 '0.035
Else
    Gravity = 0
    
End If

End Sub

Private Sub cmdExit_Click()
termina True

End Sub

Private Sub Command1_Click()
Command1.Visible = False
SavePos.Visible = True
cmdExit.Visible = True

MyColor = D3DColorMake(1, 1, 1, 1)

creaSchermo2 PIC.Width, PIC.Height, D3DFMT_A8R8G8B8, PIC.hwnd, True, 2, False



AIR = 0.99 '0.99
kMuscleSpeedLimit = 15 '10
chGravity_Click


S = 0.03 * 1.2 '* 0.9 '* 1.2 '* 0.9
Doll.DefaultStrength = S

''''''''''''''''''''''''''''''''''''''''
'RagDoll

'DOLL.ADDpoint 100, 200
Doll.ADDpoint 110, 200
Doll.ADDpoint 110, 170
Doll.ADDpoint 120, 140
Doll.ADDpoint 130, 170
Doll.ADDpoint 130, 200
'DOLL.ADDpoint 140, 200

Doll.ADDpoint 120, 110

Doll.ADDpoint 100, 110
Doll.ADDpoint 90, 130

Doll.ADDpoint 140, 110
Doll.ADDpoint 150, 130

Doll.ADDpoint 120, 90 + 5 - 5

'Links
Doll.ADDLink 1, 2
Doll.ADDLink 2, 3
Doll.ADDLink 5, 4
Doll.ADDLink 4, 3
Doll.ADDLink 6, 3

Doll.ADDLink 8, 7
Doll.ADDLink 7, 6

Doll.ADDLink 10, 9
Doll.ADDLink 9, 6

Doll.ADDLink 11, 6

'Muscles
'S = 0.03 * 1.2
Doll.ADDMuscle 1, 2, S
Doll.ADDMuscle 3, 4, S
Doll.ADDMuscle 2, 5, S
Doll.ADDMuscle 4, 5, S


Doll.ADDMuscle 6, 7, S
Doll.ADDMuscle 8, 9, S
Doll.ADDMuscle 7, 5, S
Doll.ADDMuscle 9, 5, S

Doll.ADDMuscle 10, 5, S * 0.9

Doll.Obj_SAVE
Doll.Obj_LOADandPlace "obj.txt", 200, 250


GlobStrength.Min = 0
GlobStrength.Max = Doll.DefaultStrength * 1000
GlobStrength.Value = GlobStrength.Max


For M = 2 To Doll.NMuscles
    Load MuscleANG(M - 1)
    MuscleANG(M - 1).Visible = True
    MuscleANG(M - 1).Top = MuscleANG(M - 2).Top + MuscleANG(M - 2).Height
    Load mDESC(M - 1)
    mDESC(M - 1).Top = MuscleANG(M - 1).Top
    mDESC(M - 1).Visible = True
Next

For M = 1 To Doll.NMuscles
    MuscleANG(M - 1).Min = -PI * 200
    MuscleANG(M - 1).Max = PI * 200
    MuscleANG(M - 1).Value = Doll.MUSCLE_MainANG(M) * 100
Next

'knee , hip, elbow, shoulder, head
mDESC(0) = "L - Knee"
mDESC(1) = "R - Knee"
mDESC(2) = "L - Hip"
mDESC(3) = "R - Hip"
mDESC(4) = "L - Elbow"
mDESC(5) = "R - Elbow"
mDESC(6) = "L - Shoulder"
mDESC(7) = "R - Shoulder"
mDESC(8) = "Head"


GetPoses
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReDim Sprite(Doll.Nlinks)

For i = 1 To Doll.Nlinks
    Set Sprite(i).Tex = creaTex(App.Path & "\Texture\T" & i & ".BMP", D3DColorMake(1, 0, 1, 1), True) 'transparent
    PIC_S = LoadPicture(App.Path & "\Texture\T" & i & ".BMP")
    PIC_S.Refresh
    Sprite(i).TexCenter.y = PIC_S.Height / 2
    Sprite(i).TexCenter.x = PIC_S.Width / 2
    Sprite(i).Scala.x = 1
    Sprite(i).Scala.y = 1
    Sprite(i).DrawScala.x = 1
    Sprite(i).DrawScala.y = 1
    
Next
'Don't know why but had to add this adjustment to do Right positioning in Doll.Draw_DX
Sprite(1).TexCenter.y = Sprite(1).TexCenter.y + 3
Sprite(3).TexCenter.y = Sprite(3).TexCenter.y + 3
Sprite(2).TexCenter.y = Sprite(2).TexCenter.y + 3
Sprite(4).TexCenter.y = Sprite(4).TexCenter.y + 3
Sprite(1).DrawScala.x = 1.1
Sprite(3).DrawScala.x = 1.1
Sprite(2).DrawScala.x = 1.1
Sprite(4).DrawScala.x = 1.1
Sprite(5).TexCenter.y = Sprite(5).TexCenter.y + 6
Sprite(5).DrawScala.y = 1.5
Sprite(6).TexCenter.x = Sprite(6).TexCenter.x + 4
Sprite(8).TexCenter.x = Sprite(8).TexCenter.x + 4
Sprite(6).DrawScala.x = 0.8
Sprite(8).DrawScala.x = 0.8
Sprite(7).DrawScala.x = 0.8
Sprite(9).DrawScala.x = 0.8
'Sprite(10).TexCenter.x = Sprite(10).TexCenter.x    'This is for SmileFace
'Sprite(10).TexCenter.y = Sprite(10).TexCenter.y - 2
'Sprite(10).DrawScala.x = 0.56
'Sprite(10).DrawScala.y = 0.56
Sprite(10).TexCenter.x = 17 'Sprite(10).TexCenter.x
Sprite(10).TexCenter.y = 12 'Sprite(10).TexCenter.y - 2
Sprite(10).DrawScala.x = 0.15 * 0.85
Sprite(10).DrawScala.y = 0.1 * 0.85



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
TIMER1.Enabled = True

End Sub



Private Sub Form_Load()
Me.Caption = Me.Caption & "  V" & App.Major & "." & App.Minor & "   [ DirectX ]"


End Sub

Private Sub GlobStrength_Change()

For M = 1 To Doll.NMuscles
    Doll.MUSCLE_SetStrength(M) = GlobStrength.Value / 1000
Next

End Sub

Private Sub GlobStrength_Scroll()
For M = 1 To Doll.NMuscles
    Doll.MUSCLE_SetStrength(M) = GlobStrength.Value / 1000
Next

End Sub

Private Sub List1_Click()
Doll.OBJ_LoadPose (List1)
For M = 1 To Doll.NMuscles
    MuscleANG(M - 1).Value = Doll.MUSCLE_MainANG(M) * 100
Next M

End Sub

Private Sub MuscleANG_Change(Index As Integer)
MUS = Index + 1
Doll.MUSCLE_MainANG(Index + 1) = CDbl(MuscleANG(Index).Value / 100)

End Sub

Private Sub MuscleANG_Scroll(Index As Integer)
MUS = Index + 1
Doll.MUSCLE_MainANG(Index + 1) = CDbl(MuscleANG(Index).Value / 100)
End Sub



Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim DMin As Single
Dim P1 As tPoint
Dim P2 As tPoint
DMin = 1E+19
P2.x = x
P2.y = y
For i = 1 To Doll.Npoints
    P1.x = Doll.PointX(i)
    P1.y = Doll.PointY(i)
    If Distance(P1, P2) < DMin Then DMin = Distance(P1, P2): PtoMove = i
Next
mouseX = x
mouseY = y
doMouse = True


End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    mouseX = x
    mouseY = y
End If

End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
doMouse = False
End Sub

Private Sub SavePos_Click()
Doll.OBJ_SavePose ("POS" & List1.ListCount & ".txt")
List1.AddItem "POS" & List1.ListCount & ".txt"

End Sub

Private Sub Timer1_Timer()

'PIC.Cls
'Doll.DRAW PIC, SLN


DollDRAW_DX
'Doll.DRAW PIC, SLN


Doll.doPHYSICS
If doMouse Then doMouseForces

End Sub

Sub GetPoses()
Dim D As String

D = Dir(App.Path & "\Pos" & "*.txt")
While D <> ""
    List1.AddItem D
    D = Dir
Wend

End Sub

Sub doMouseForces()
Doll.PointVX(PtoMove) = Doll.PointVX(PtoMove) - (Doll.PointX(PtoMove) - mouseX) * 0.015
Doll.PointVY(PtoMove) = Doll.PointVY(PtoMove) - (Doll.PointY(PtoMove) - mouseY) * 0.015
PIC.Line (Doll.PointX(PtoMove), Doll.PointY(PtoMove))-(mouseX, mouseY), vbGreen
End Sub

Sub DollDRAW_DX()
Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, D3DColorRGBA(0, 0, 0, 0), 1#, 0 'pulisce lo schermo
Device.BeginScene 'inizia il rendering
dSprite.Begin
'Stop

Doll.DRAW_DX


dSprite.End
'testo.DrawTextW txtSCR & "   " & CAR(BEST).dDISTtot, -1, r, DT_LEFT, D3DColorMake(1, 1, 1, 1)
Device.EndScene 'fa terminare il rendering
Device.Present ByVal 0, ByVal 0, 0, ByVal 0 'invia l'immagine al monitor



End Sub
