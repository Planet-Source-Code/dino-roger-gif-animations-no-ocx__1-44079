<div align="center">

## Gif Animations No OCX


</div>

### Description

This code will display a gif animation (even if transparent). Uses no OCX files. I did not write this code, but found it on a site and thought PSC would make a good home for it. Enjoy.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dino Roger](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dino-roger.md)
**Level**          |Beginner
**User Rating**    |3.6 (18 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dino-roger-gif-animations-no-ocx__1-44079/archive/master.zip)





### Source Code

```
' Add a Timer called Timer1
' Add a image object called Image1
' Add a module (BAS) and paste the following
Option Explicit
Public RepeatTimes As Long 'This one calculates,
' but don't use in this sample. If You need, You
' can add simple checking at Timer1_Timer Procedure
Public TotalFrames As Long
Public Function LoadGif(sFile As String, aImg As Variant) As Boolean
 LoadGif = False
 If Dir$(sFile) = "" Or sFile = "" Then
  MsgBox "File " & sFile & " not found", vbCritical
  Exit Function
 End If
 On Error GoTo ErrHandler
 Dim fNum As Integer
 Dim imgHeader As String, fileHeader As String
 Dim buf$, picbuf$
 Dim imgCount As Integer
 Dim i&, j&, xOff&, yOff&, TimeWait&
 Dim GifEnd As String
 GifEnd = Chr(0) & Chr(33) & Chr(249)
 For i = 1 To aImg.Count - 1
  Unload aImg(i)
 Next i
 fNum = FreeFile
 Open sFile For Binary Access Read As fNum
  buf = String(LOF(fNum), Chr(0))
  Get #fNum, , buf 'Get GIF File into buffer
 Close fNum
 i = 1
 imgCount = 0
 j = InStr(1, buf, GifEnd) + 1
 fileHeader = Left(buf, j)
 If Left$(fileHeader, 3) <> "GIF" Then
  MsgBox "This file is not a *.gif file", vbCritical
  Exit Function
 End If
 LoadGif = True
 i = j + 2
 If Len(fileHeader) >= 127 Then
  RepeatTimes& = Asc(Mid(fileHeader, 126, 1)) + (Asc(Mid(fileHeader, 127, 1)) * 256&)
 Else
  RepeatTimes = 0
 End If
 Do ' Split GIF Files at separate pictures
  ' and load them into Image Array
  imgCount = imgCount + 1
  j = InStr(i, buf, GifEnd) + 3
  If j > Len(GifEnd) Then
   fNum = FreeFile
   Open "temp.gif" For Binary As fNum
    picbuf = String(Len(fileHeader) + j - i, Chr(0))
    picbuf = fileHeader & Mid(buf, i - 1, j - i)
    Put #fNum, 1, picbuf
    imgHeader = Left(Mid(buf, i - 1, j - i), 16)
   Close fNum
   TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256&)) * 10&
   If imgCount > 1 Then
    xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256&)
    yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * 256&)
    Load aImg(imgCount - 1)
    aImg(imgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
    aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
   End If
   ' Use .Tag Property to save TimeWait interval for separate Image
   aImg(imgCount - 1).Tag = TimeWait
   aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
   Kill ("temp.gif")
   i = j
  End If
  DoEvents
 Loop Until j = 3
' If there are one more Image - Load it
 If i < Len(buf) Then
  fNum = FreeFile
  Open "temp.gif" For Binary As fNum
   picbuf = String(Len(fileHeader) + Len(buf) - i, Chr(0))
   picbuf = fileHeader & Mid(buf, i - 1, Len(buf) - i)
   Put #fNum, 1, picbuf
   imgHeader = Left(Mid(buf, i - 1, Len(buf) - i), 16)
  Close fNum
  TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256)) * 10
  If imgCount > 1 Then
   xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256)
   yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * 256)
   Load aImg(imgCount - 1)
   aImg(imgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
   aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
  End If
  aImg(imgCount - 1).Tag = TimeWait
  aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
  Kill ("temp.gif")
 End If
 TotalFrames = aImg.Count - 1
 Exit Function
ErrHandler:
 MsgBox "Error No. " & Err.Number & " when reading file", vbCritical
 LoadGif = False
 On Error GoTo 0
End Function
'
'
'
'
'
'
'
'Paste the following in form code
'
Private Sub Form_Load()
 Timer1.Enabled = False
 If LoadGif("C:\Ball.gif", Image1) Then
' Change C:\Ball gif to your animation
  FrameCount = 0
  Timer1.Interval = CLng(Image1(0).Tag)
  Timer1.Enabled = True
 End If
End Sub
Private Sub Timer1_Timer()
 If FrameCount < TotalFrames Then
  Image1(FrameCount).Visible = False
  FrameCount = FrameCount + 1
  Image1(FrameCount).Visible = True
  Timer1.Interval = CLng(Image1(FrameCount).Tag)
 Else
  FrameCount = 0
  For i = 1 To Image1.Count - 1
   Image1(i).Visible = False
  Next i
  Image1(FrameCount).Visible = True
  Timer1.Interval = CLng(Image1(FrameCount).Tag)
 End If
End Sub
```

