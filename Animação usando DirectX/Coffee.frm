VERSION 5.00
Object = "{34F681D0-3640-11CF-9294-00AA00B8A733}#1.0#0"; "DANIM.DLL"
Begin VB.Form coffee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Coffee"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin DirectAnimationCtl.DAViewerControlWindowed DAViewerControlWindowed 
      Height          =   3850
      Left            =   120
      OleObjectBlob   =   "Coffee.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "coffee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  'Set the background
  Set imgBackGround = ImportImage(App.Path & "\Media\CLOUDS_COFFEE.GIF")

  'Create the final montage by layering the espresso machine, the cups and the steam
  'on top of each other.
  Set finalMtg = UnionMontage(UnionMontage(steamMontage(), montage()), machineMontage())

  'Create the final image.
  Set finalImage = Overlay(finalMtg.Render(), Overlay(beans(), imgBackGround))
  Set finalImage = Overlay(finalImage, SolidColorImage(White))
  
  DAViewerControlWindowed.UpdateInterval = 0.1
  
  'Display the final image.
  DAViewerControlWindowed.Image = finalImage
  
  'Set the sound.
  DAViewerControlWindowed.sound = Mix(sound().Pan(-1), sound().Pan(1))

  'Start the animation.
  DAViewerControlWindowed.Start
End Sub
Function sound()
  'This function creates the sound, which starts out as silence, and then
  'changes to steam.mp2 when the image is clicked.
  Set steamDurationConst = DANumber(7.25)
  Set steamSound = ImportSound(App.Path & "\Media\Radiotow.mp2").sound

  Set s0 = CreateObject("DirectAnimation.DASound")
  s0.Init DAStatics.Until(Silence, LeftButtonDown, DAStatics.Until(steamSound.Gain(0.85), TimerAnim(steamDurationConst), s0))

  Set sound = s0
End Function
Function montage()
  'This function creates a montage of five cups, which rotate around the espresso
  'machine.  The orbit is constructed by the orbitCup function.
  pi = 3.14159265359
  total = 5
  Set cupImageX = EmptyMontage

  For i = 0 To total
    Set cupImageX = UnionMontage(cupImageX, orbitCup(Add(Mul(DANumber(i), _
      Mul(DANumber(2), Div(DANumber(pi), DANumber(total)))), LocalTime)))
  Next
  Set montage = cupImageX
End Function
Function beans()
  'This function creates the beans you see in the background.  Two images,
  'bean1.gif and bean2.gif are imported, and then moved across the screen
  'while being rotated.
  Set delay = DANumber(0.5)
  Set Size = DANumber(0.5)
  Set initBean1 = ImportImage(App.Path & "\Media\Bean1.gif")
  Set initBean2 = ImportImage(App.Path & "\Media\Bean2.gif")

  Set image0 = CreateObject("DirectAnimation.DAImage")
  Set image1 = CreateObject("DirectAnimation.DAImage")
  image0.Init DAStatics.Until(initBean1, TimerAnim(delay), image1)
  image1.Init DAStatics.Until(initBean2, TimerAnim(delay), image0)

  Set bean1 = image0.Transform(Scale2UniformAnim(Size))
  Set bean2 = image1.Transform(Scale2UniformAnim(Size))

  Set beans = Overlay(bean1.Transform(Translate2(-0.01, -0.01)), bean2.Transform(Translate2(0.01, 0)))

  Set rain = beans.Tile()

  Set motion = Mul(LocalTime, Mul(DANumber(2), Div(DANumber(0.03), DANumber(4))))

  Set beans = rain.Transform(Translate2Anim(Neg(motion), Neg(motion)))
End Function
Function machineMontage()
  'This function displays the espresso machine.
  Set steamDurationConst = DANumber(7.25)
  Set espreso1 = ImportImage(App.Path & "\Media\espreso1.gif")
  Set espreso2 = ImportImage(App.Path & "\Media\espreso2.gif")

  Set image5 = CreateObject("DirectAnimation.DAImage")
  image5.Init DAStatics.Until(espreso1, LeftButtonDown, DAStatics.Until(espreso2, TimerAnim(steamDurationConst), image5))

  Set machineMontage = ImageMontage(image5, 0)
End Function
Function steamMontage()
  'This function displays the steam you see when you click on the image.
  Set steamDurationConst = DANumber(7.25)
  Dim steamImages(4)
  Set steamImages(0) = ImportImage(App.Path & "\Media\steam_1.gif")
  Set steamImages(1) = ImportImage(App.Path & "\Media\steam_2.gif")
  Set steamImages(2) = ImportImage(App.Path & "\Media\steam_3.gif")
  Set steamImages(3) = ImportImage(App.Path & "\Media\steam_4.gif")
  Set steamImages(4) = ImportImage(App.Path & "\Media\steam_5.gif")

  Set steamLen = DANumber(4)

  Set condition = GT(Add(Div(Mul(LocalTime, steamLen), steamDurationConst), DANumber(1)), steamLen)

  Set result2 = Add(Div(Mul(LocalTime, steamLen), steamDurationConst), DANumber(1))

  Set Index = Cond(condition, steamLen, result2)

  Set a = DAStatics.Array(steamImages)

  Set s0 = CreateObject("DirectAnimation.DAImage")
  s0.Init DAStatics.Until(EmptyImage, LeftButtonDown, DAStatics.Until(a.NthAnim(Index), TimerAnim(steamDurationConst), s0))

  Set image1 = s0.Transform(Translate2(-0.0085, 0.002))

  Set steamMontage = ImageMontage(image1, -0.0001)
  End Function
Function orbitCup(angle)
  pi = 3.14159265359
  Set pos = Point3(0, 0.05, 0)
  Set pos = pos.Transform(Compose3(Rotate3Anim(XVector3, Mul(DANumber(7), Div(DANumber(pi), DANumber(16)))), Rotate3Anim(ZVector3, angle)))

  Set cupAngle = LocalTime

  Set imageXX = cupImage(cupAngle).Transform(Compose2(Translate2Anim(pos.X, pos.Y), Scale2UniformAnim(DAStatics.Sub(DANumber(1), Mul(DAStatics.Abs(DAStatics.Cos(Div(angle, DANumber(2)))), DANumber(0.5))))))

  Set orbitCup = ImageMontageAnim(imageXX, Neg(pos.Z))
End Function
Function cupImage(cupAngle)
  pi = 3.14159265359
  Dim cupImages
  cupImages = Array(ImportImage(App.Path & "\Media\cup1.gif"), ImportImage(App.Path & "\Media\cup2.gif"), ImportImage(App.Path & "\Media\cup3.gif"), ImportImage(App.Path & "\Media\cup4.gif"), ImportImage(App.Path & "\Media\cup5.gif"), ImportImage(App.Path & "\Media\cup6.gif"), ImportImage(App.Path & "\Media\cup7.gif"), ImportImage(App.Path & "\Media\cup8.gif"))

  Set Number = DANumber(7)
  Set Index = Add(DAStatics.Mod(Mul(Number, Div(cupAngle, Mul(DANumber(2), DANumber(pi)))), Number), DANumber(1))

  Set a = DAStatics.Array(cupImages)

  Set cupImage = a.NthAnim(Index)
End Function

