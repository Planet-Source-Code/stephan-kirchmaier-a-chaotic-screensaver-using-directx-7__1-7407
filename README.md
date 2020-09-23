<div align="center">

## A chaotic ScreenSaver using DirectX 7


</div>

### Description

It's my first DirectX Project. Please say what you think of it!
 
### More Info
 
You must link "DirectX7 for Visual Basic Type Library" to your Project.

The ScreenSaver ends by clicking on the screen.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Stephan Kirchmaier](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stephan-kirchmaier.md)
**Level**          |Advanced
**User Rating**    |3.4 (24 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[DirectX](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/directx__1-44.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stephan-kirchmaier-a-chaotic-screensaver-using-directx-7__1-7407/archive/master.zip)





### Source Code

```
Option Explicit
Private DX7 As DirectX7
Private DXD As DirectDraw7
Private DXDS As DirectDrawSurface7
Private DXSD As DDSURFACEDESC2
Private Sub Form_Load()
 Dim i As Long, j As Long
 frmMain.Show
 'Create a DirectX7-Object and a DirectDraw-Object
 Set DX7 = New DirectX7
 Set DXD = DX7.DirectDrawCreate("")
 With DXSD
  .lFlags = DDSD_CAPS
  .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
 End With
 'Fullscreen and set the resolution to 640 X 480
 DXD.SetCooperativeLevel frmMain.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
 DXD.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT
 'Create the Surface using the Surfacedescription DXSD
 Set DXDS = DXD.CreateSurface(DXSD)
 i = 0
 Do Until DoEvents()
  For j = 0 To ScaleWidth Step 50
   'Set the Linecolor
   DXDS.SetForeColor i
   'Draw the line
   DXDS.DrawLine Rnd * Screen.Width, Rnd * Screen.Height, j, 0
   i = i + 1
   'Change the color
   If i = 65536 Then
    i = 0
   End If
  Next j
 Loop
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call endp
End Sub
Private Sub endp()
 'Clean up things
 DXD.RestoreDisplayMode
 Set DX7 = Nothing
 Set DXD = Nothing
 Set DXDS = Nothing
 End
End Sub
```

