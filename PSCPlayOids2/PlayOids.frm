VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   165
   ClientTop       =   615
   ClientWidth     =   8505
   Icon            =   "PlayOids.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   502
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   567
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optAnim 
      Caption         =   "Rotate in plane"
      Height          =   570
      Index           =   2
      Left            =   6465
      TabIndex        =   31
      Top             =   195
      Width           =   900
   End
   Begin VB.OptionButton optAnim 
      Caption         =   "Stop rotation"
      Height          =   570
      Index           =   3
      Left            =   7425
      TabIndex        =   23
      Top             =   195
      Width           =   900
   End
   Begin VB.OptionButton optAnim 
      Caption         =   "Rotate about horizontal"
      Height          =   570
      Index           =   1
      Left            =   5325
      TabIndex        =   22
      Top             =   210
      Width           =   1065
   End
   Begin VB.OptionButton optAnim 
      Caption         =   "Rotate about poles"
      Height          =   540
      Index           =   0
      Left            =   4305
      TabIndex        =   21
      Top             =   225
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "SHAPES"
      ForeColor       =   &H00FF0000&
      Height          =   5190
      Left            =   225
      TabIndex        =   4
      Top             =   1140
      Width           =   1710
      Begin VB.OptionButton Option1 
         Caption         =   "Wheel"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   14
         Left            =   210
         TabIndex        =   20
         Top             =   4695
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Torus"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   13
         Left            =   210
         TabIndex        =   19
         Top             =   4425
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Semi-sphere"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   12
         Left            =   195
         TabIndex        =   17
         Top             =   4125
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Parabloid"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   11
         Left            =   195
         TabIndex        =   16
         Top             =   3825
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bendy blades"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   10
         Left            =   195
         TabIndex        =   15
         Top             =   3510
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sine wave"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   9
         Left            =   180
         TabIndex        =   14
         Top             =   3150
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cylinder"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   165
         TabIndex        =   13
         Top             =   600
         Width           =   1305
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2 Ellipsoids"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   165
         TabIndex        =   12
         Top             =   900
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Trophy"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   165
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cones"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   4
         Left            =   195
         TabIndex        =   10
         Top             =   1515
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Reflection"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   195
         TabIndex        =   9
         Top             =   1845
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Wine glass"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         Left            =   195
         TabIndex        =   8
         Top             =   2175
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Hyperboloid"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   7
         Left            =   195
         TabIndex        =   7
         Top             =   2505
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Flat surface"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   8
         Left            =   225
         TabIndex        =   6
         Top             =   2835
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sphere"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   165
         TabIndex        =   5
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Frame frmPerspec 
      Caption         =   "Perspective depth"
      Height          =   720
      Left            =   1500
      TabIndex        =   3
      Top             =   45
      Width           =   1515
      Begin VB.VScrollBar VScroll2 
         Height          =   360
         LargeChange     =   10
         Left            =   1080
         Max             =   100
         Min             =   350
         SmallChange     =   10
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   210
         Value           =   150
         Width           =   285
      End
      Begin VB.Label LabPerspec 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabPerspec"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   28
         Top             =   270
         Width           =   810
      End
   End
   Begin VB.CheckBox chkPerspective 
      Caption         =   "Perspective"
      Height          =   375
      Left            =   3060
      TabIndex        =   2
      Top             =   270
      Width           =   1170
   End
   Begin VB.Frame frmAspect 
      Caption         =   "Aspect"
      Height          =   720
      Left            =   225
      TabIndex        =   1
      Top             =   45
      Width           =   1230
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Left            =   855
         Max             =   0
         Min             =   16
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   225
         Value           =   10
         Width           =   270
      End
      Begin VB.Label LabAspect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabAspect"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   27
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Timer Timer2 
      Left            =   2595
      Top             =   6645
   End
   Begin VB.Timer Timer1 
      Left            =   2130
      Top             =   6645
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6030
      Left            =   2160
      ScaleHeight     =   398
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   396
      TabIndex        =   0
      Top             =   1230
      Width           =   6000
      Begin VB.Timer Timer3 
         Left            =   855
         Top             =   5385
      End
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      Height          =   195
      Index           =   1
      Left            =   8265
      TabIndex        =   30
      Top             =   7245
      Width           =   180
   End
   Begin VB.Label Label1 
      Caption         =   "Z"
      Height          =   195
      Index           =   0
      Left            =   1995
      TabIndex        =   29
      Top             =   885
      Width           =   180
   End
   Begin VB.Line Line2 
      X1              =   505
      X2              =   548
      Y1              =   490
      Y2              =   490
   End
   Begin VB.Line Line1 
      X1              =   137
      X2              =   137
      Y1              =   74
      Y2              =   117
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Height          =   90
      Left            =   4635
      TabIndex        =   24
      Top             =   7335
      Width           =   870
   End
   Begin VB.Label Label2 
      Caption         =   "Right click on image to save"
      Height          =   225
      Left            =   4065
      TabIndex        =   18
      Top             =   915
      Width           =   2115
   End
   Begin VB.Menu mF 
      Caption         =   "&FILE"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As bmp"
      End
      Begin VB.Menu zbrk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&XIT"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PLAYOIDS2  by Robert Rayment (May 2004) 30
'Three dimensional animated wire frame shapes

' Updates
' 1. 3rd timer for rotation in plane
' 2. Left click moves shape to the click point,
'    suggested by Roger Gilchrist

Option Explicit
Option Base 1

Private grxhi As Long, gryhi  As Long 'Number of grids

Private svrx() As Single          ' Start X,Y,Z values
Private svry() As Single
Private svrz() As Single
Private xs() As Single
Private zs() As Single            ' picDisplay points
Private xT As Single
Private yT As Single              ' Timer points
Private zT As Single              ' Timer points
Private xoff As Single
Private zoff As Single            ' Shape location
Private xangintv As Single
Private yangintv As Single        ' Timer angle intervals
Private zangintv As Single        ' Timer angle intervals
Private yeye As Single              ' eye y point
Private zaspect As Single         ' Spheroid aspect ratio
Private zrad As Single            ' Spheroid radius

Private phdc As Long
Private CheckPerspective As Long  'yeye, perspective depth

Private pp As POINTAPI         ' For MoveToEx bu not used

Private optShape As Long       ' Shape index

Private CommonDialog1 As OSDialog

Private FileSpec$, Pathspec$

Const LStep  As Long = 15      ' Need 360\LStep = integer

Const zpi# = 3.1415927
Const dtr# = zpi# / 180
'Const rtd# = 180 / zpi#


Private Sub Form_Load()
Dim k As Long
   'Centre form
   Form1.Top = (Screen.Height - Form1.Height) / 2
   Form1.Left = (Screen.Width - Form1.Width) / 2
   Form1.Height = 8400
   Show
   
   Caption = "PLAYOIDS2 by Robert Rayment"
   'Plotting intervals
   grxhi = (1 + 360 \ LStep): gryhi = (1 + 360 \ LStep)
   
   ReDim svrx(grxhi, gryhi), svry(grxhi, gryhi), svrz(grxhi, gryhi)
   ReDim xs(grxhi, gryhi), zs(grxhi, gryhi)
   
   Timer1.Interval = 1
   Timer1.Enabled = False
   Timer2.Interval = 1
   Timer2.Enabled = False
   Timer3.Interval = 1
   Timer3.Enabled = False
   'Timer angular steps
   xangintv = 2: yangintv = 2: zangintv = 2
   'Starting aspect & perspective depth
   LabAspect = "1"
   LabPerspec = "150"

   ' Basic radius
   zrad = 100
   picDisplay.Height = 400
   picDisplay.Width = 400
   
   ''For exact pic size without perspective
   'If za > 1 Then
   '   picDisplay.Height = za * 2 * zrad
   'Else
   '   picDisplay.Height = 2 * zrad
   'End If
   'picDisplay.Width = 2 * zrad
   'If za < 1 Then
   '   picDisplay.Width = 2 * zrad
   'Else
   '   picDisplay.Width = 2 * zrad
   'End If
      
   ' Line up Option buttons
   For k = 0 To 14
      Option1(k).Left = 150
      If k > 0 Then
         Option1(k).Top = Option1(0).Top + k * 300
      End If
   Next k
   
   xoff = picDisplay.Width / 2
   zoff = picDisplay.Height / 2
   
   'FILL ARRAYS WITH X,Y,Z VERTICES
   'USING SPHERICAL OR CYLINDRICAL COORDINATES
   optShape = 0
   Option1(0) = True
   zaspect = 1
   yeye = 150
   
   SHAPES
End Sub

Private Sub SHAPES()
Dim i As Long, j As Long
Dim ztheta As Single
Dim zphi As Single
Dim zz As Single
Dim zp As Single
Dim z As Single, x As Single, y As Single
Dim zzrad As Single

   j = 0
   For ztheta = 0 To 360 Step LStep
      i = 0
      j = j + 1
      zz = ztheta * dtr#
   For zphi = 0 To 360 Step LStep
      i = i + 1
      zp = zphi * dtr#
      
      Select Case optShape
      Case 0   ' ELLIPSOID
         z = zaspect * zrad * Cos(zz)
         x = zrad * Sin(zz) * Cos(zp)
         y = zrad * Sin(zz) * Sin(zp)
      Case 1   ' CYLINDER
         z = (ztheta - 180) / 2 'Make origin at zero
         x = zrad * Cos(zp)
         y = zaspect * zrad * Sin(zp)
      Case 2   ' TWO ELLIPSOIDS
         z = (ztheta - 180) / 2 'Make origin at zero
         x = zaspect * zrad * Sin(zz) * Cos(zp)
         y = zrad * Sin(zz) * Sin(zp)
      Case 3   ' TWO HALF & ONE WHOLE SPHERE,  Trophy
         z = (ztheta - 180) / 2 'Make origin at zero
         zzrad = zrad * Cos(zz)
         x = zzrad * Cos(zp)
         y = zaspect * zzrad * Sin(zp)
      Case 4   ' 2 LINEAR CONES
         z = (ztheta - 180) / 2 'Make origin at zero
         zzrad = zrad * z / 100
         x = zzrad * Cos(zp)
         y = zaspect * zzrad * Sin(zp)
      Case 5   ' WINE GLASS & REFLECTION
         z = (ztheta - 180) / 2 'Make origin at zero
         zzrad = zrad * Sqr(Abs(z)) / 20
         If zzrad = 0 Then
            zzrad = zrad
         ElseIf zzrad < 30 Then
            zzrad = 0
         End If
         x = zzrad * Cos(zp)
         y = zaspect * zzrad * Sin(zp)
      Case 6   ' WINE GLASS
         z = (ztheta - 180) / 2 'Make origin at zero
         If z >= 0 Then
            zzrad = zrad * Sqr(z) / 15
         ElseIf z > -80 Then
            zzrad = 2
         Else
            zzrad = zrad / 2
         End If
         x = zzrad * Cos(zp)
         y = zaspect * zzrad * Sin(zp)
      Case 7   ' HYPERBOLOID
         z = (ztheta - 180) / 2 'Make origin at zero
         zzrad = 25 + zrad * z * z / 10000
         x = zzrad * Cos(zp)
         y = zaspect * zzrad * Sin(zp)
      Case 8   ' FLAT SURFACE
         For j = 1 To gryhi
         For i = 1 To grxhi
            svrx(i, j) = 8 * (i - 10)
            svry(i, j) = 8 * (j - 10)
            svrz(i, j) = 8 * (i + j - 20)
         Next i
         Next j
         Exit For
      Case 9   ' SINE WAVE
         z = (ztheta - 180) / 2 'Make origin at zero
         x = zrad * Sin(zz) '* Cos(zp)
         y = zaspect * zrad * Sin(zz) '* Sin(zp)
      Case 10  ' BENDY BLADES
         z = zaspect * zrad * Cos(zz) ^ 2
         x = zrad * Sin(zz) * Cos(zp) / 2
         y = zrad * Sin(zz)
      Case 11   ' PARABLOID
         z = zaspect * zrad * Cos(zz) * Cos(zz)
         x = zrad * Sin(zz) * Cos(zp)
         y = zrad * Sin(zz) * Sin(zp)
      Case 12   ' SEMI-SPHERE
         z = zaspect * zrad * Abs(Cos(zz))
         x = zrad * Sin(zz) * Cos(zp)
         y = zrad * Sin(zz) * Sin(zp)
      Case 13  ' TORUS
         z = zaspect * zrad * Sin(zp) / 2
         x = zrad * Cos(zz) * (1 + Cos(zp) / 2)
         y = zrad * Sin(zz) * (1 + Cos(zp) / 2)
      Case 14  ' WHEEL
         z = zaspect * zrad * Cos(zz) * Sin(zz) ^ 3
         x = zrad * Sin(zz) * Cos(zp)
         y = zrad * Sin(zz) * Sin(zp)
      End Select
      
      svrx(i, j) = x
      svry(i, j) = y
      svrz(i, j) = z
   Next zphi
   If optShape = 8 Then Exit For
   Next ztheta
   
   xT = 0: yT = 0: zT = 0
   '------------------------------
   picDisplay_MouseMove 1, 0, xT, yT
   '------------------------------
End Sub

Private Sub mnuSave_Click()
Dim Title$, Filt$, InDir$
Dim Ext$
Dim FIndex As Long
Dim p As Long
Dim res As Long

Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
optAnim(3).Value = True

Set CommonDialog1 = New OSDialog
   Title$ = "Save As *.bmp"
   Filt$ = "Save Image|*.bmp"
   If FileSpec$ = "" Then
      InDir$ = Pathspec$
   Else
      p = InStrRev(FileSpec$, "\")
      InDir$ = Left$(FileSpec$, p)
      'InDir$ = FileSpec$
   End If
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
Set CommonDialog1 = Nothing
   
   If LenB(FileSpec$) > 0 Then
      FixExtension FileSpec$, ".bmp"
      SavePicture picDisplay.Image, FileSpec$
   End If
End Sub

Private Sub Option1_Click(Index As Integer)
   optShape = Index
   SHAPES
End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      xoff = x: zoff = y
   End If
End Sub

Private Sub picDisplay_Mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbRightButton Then
      mnuSave_Click
   End If
End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   xT = x
   If Timer3.Enabled = False Then
      yT = y
   Else
      zT = y
   End If
   '-----------------------------------
   CalculateScreenPoints x, y, zT
   '-----------------------------------
   Display
   '-----------------------------------
End Sub

Private Sub Display()
Dim i As Long, j As Long
Dim res As Long
   'Dim pp As POINTAPI Not used but API needs it
   'res = LineTo(Form1.hdc, xx&, yy&)
   'res = MoveToEx(Form1.hdc, xx&, yy&, pp)
   
   picDisplay.Cls 'Simplest & Fast enough
   phdc& = picDisplay.hdc  'picDisplay device context
   picDisplay.DrawWidth = 1  ' 2 is much slower
   
   'Draw X-latitude lines
   picDisplay.ForeColor = vbBlack
   For j = 1 To gryhi '/ 2
      res = MoveToEx(phdc&, xs(1, j), zs(1, j), pp)
   For i = 1 + 1 To grxhi
      res = LineTo(phdc&, xs(i, j), zs(i, j))
   Next i
   Next j
   
   'Draw Y-longitude lines
   picDisplay.ForeColor = vbBlack
   For i = 1 To grxhi '/ 2
      res = MoveToEx(phdc&, xs(i, 1), zs(i, 1), pp)
   For j = 1 + 1 To gryhi
      res = LineTo(phdc&, xs(i, j), zs(i, j))
   Next j
   Next i
   
   'Equator
   picDisplay.ForeColor = vbBlue
   picDisplay.DrawWidth = 2
   j = 7 'LStep / 2
      res = MoveToEx(phdc&, xs(1, j), zs(1, j), pp)
   For i = 1 + 1 To grxhi
      res = LineTo(phdc&, xs(i, j), zs(i, j))
   Next i
End Sub

Public Sub CalculateScreenPoints(x As Single, y As Single, zT As Single)
'r.. to z... Singles
Dim i As Long, j As Long
Dim xcen As Single, ycen As Single, zcen As Single
Dim zang As Single, xang As Single, yang As Single
Dim scfx As Single, scfz As Single
Dim xmin As Single, zmin As Single
'Dim yoff As Single, ymin As Single
Dim rxx As Single, ryy As Single, rzz As Single
Dim ryyy As Single, rzzz As Single
Dim xe As Single, ze As Single
Dim zd As Single ', za As Single

   'x,y cursor position from MouseDown
   
   'Get angles based on cursor position
   xcen = picDisplay.ScaleWidth / 2
   ycen = picDisplay.ScaleHeight / 2
   zcen = picDisplay.ScaleHeight / 2
   
   zang = -(zpi#) * ((x - xcen) / xcen)  'zang about z-axis
   xang = -(zpi#) * ((y - ycen) / ycen)  'xang about x-axis
   yang = -(zpi#) * ((zT - zcen) / zcen) 'yang about x-axis
   
   '-------  SCALE FACTORS  ----------
   scfx = 1
   'xoff = zrad   ''For exact pic size without perspective
   'xoff = picDisplay.Width / 2
   xmin = 0
   
   scfz = 1
   
   ''For exact pic size without perspective
   'If za > 1 Then
   '   zoff = zaspect * zrad
   'Else
   '   zoff = zrad
   'End If
   'zoff = picDisplay.Height / 2
   zmin = 0
   
   '--APPLY ROTATIONS --
   
   For j = 1 To gryhi
   For i = 1 To grxhi
'      If Timer3.Enabled = False Then
         'Apply rotation to original data about z-axis
         rxx = svrx(i, j) * Cos(zang) + svry(i, j) * Sin(zang)
         ryy = svry(i, j) * Cos(zang) - svrx(i, j) * Sin(zang)
         rzz = svrz(i, j)
         
         'Apply rotation about x-axis
         ryyy = ryy * Cos(xang) - rzz * Sin(xang)
         rzz = ryy * Sin(xang) + rzz * Cos(xang)
      If Timer3.Enabled = True Then
         'Apply rotation about y-axis
         rxx = svrx(i, j) * Cos(yang) + svrz(i, j) * Sin(yang)
         rzz = svrz(i, j) * Cos(yang) - svrx(i, j) * Sin(yang)
      End If
      
      If CheckPerspective = 1 Then
         'Find the intercept at plane y=0 (ie the screen plane) of the
         'line connecting the eye point (xe,yeye,ze) with each function
         'point in turn.
         'The display intercept points will be modified & unscaled
   
         'EYE POINT  yeye settable
         xe = 0: ze = 0
         zd = (yeye - ryy)
         If zd <> 0 Then
            rxx = -yeye * (xe - rxx) / zd
            rzz = -yeye * (ze - rzz) / zd
         End If
      End If
      
      '-----  GET PLOTTING POINTS  ----------------------
      
      xs(i, j) = scfx * (rxx - xmin) + xoff
      zs(i, j) = scfz * (rzz - zmin) + zoff
      
      If xs(i, j) > 1000 Then xs(i, j) = 1000
      If xs(i, j) < -1000 Then xs(i, j) = -1000
      If zs(i, j) > 1000 Then zs(i, j) = 1000
      If zs(i, j) < -1000 Then zs(i, j) = -1000
   
   Next i
   Next j
End Sub

Private Sub Timer1_Timer()
'Rotate about poles - Z
   xT = xT + xangintv
   If xT > 6000 Or xT < -6000 Then xangintv = -xangintv
   '--------------------------------
   picDisplay_MouseMove 1, 0, xT, yT
   '--------------------------------
End Sub

Private Sub Timer2_Timer()
'Rotate about horizontal - X
   yT = yT + yangintv
   If yT > 6000 Or yT < -6000 Then yangintv = -yangintv
   '--------------------------------
   picDisplay_MouseMove 1, 0, xT, yT
   '--------------------------------
End Sub

Private Sub Timer3_Timer()
'Rotate in plane - Y
   zT = zT + zangintv
   If zT > 6000 Or zT < -6000 Then zangintv = -zangintv
   '--------------------------------
   picDisplay_MouseMove 1, 0, xT, zT
   '--------------------------------
End Sub

Private Sub optAnim_Click(Index As Integer)
   If Index = 0 Then
      If Timer1.Enabled = True Then
         Timer1.Enabled = False
      Else
         Timer1.Enabled = True
      End If
      Timer2.Enabled = False
      Timer3.Enabled = False
   ElseIf Index = 1 Then
      If Timer2.Enabled = True Then
         Timer2.Enabled = False
      Else
         Timer2.Enabled = True
      End If
      Timer1.Enabled = False
      Timer3.Enabled = False
   ElseIf Index = 2 Then
      If Timer3.Enabled = True Then
         Timer3.Enabled = False
      Else
         Timer3.Enabled = True
      End If
      Timer1.Enabled = False
      Timer2.Enabled = False
   Else
      Timer1.Enabled = False
      Timer2.Enabled = False
      Timer3.Enabled = False
   End If
End Sub

Private Sub chkPerspective_Click()
   CheckPerspective = chkPerspective.Value
End Sub

Private Sub mnuExit_Click()
   Form_Unload 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Timer1.Enabled = False
   Timer2.Enabled = False
   Unload Me
   End
End Sub

Private Sub VScroll1_Change()
'Private zaspect As Single         ' Spheroid aspect ratio
'Aspect
   zaspect = VScroll1.Value / 10
   If zaspect < 1 And zaspect > 0 Then
      LabAspect = "0" & Str$(zaspect)
   Else
      LabAspect = Str$(zaspect)
   End If
   SHAPES
End Sub

Private Sub VScroll2_Change()
'Private yeye As Single              ' eye y point
'Perspective depth
   yeye = VScroll2.Value
   LabPerspec = Str$(yeye)
   SHAPES
End Sub
