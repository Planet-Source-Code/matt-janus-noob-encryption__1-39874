<div align="center">

## noob encryption


</div>

### Description

ive seen lot's of people's input on bad and insecure some encryption codes. this is one i put togeather. im wondering how (in)secure it is. please give me your opinion so i can further improve on it. you dont need to vote, i just want feedback.

sorry about the repost. i messed up the last one.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[matt\.janus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matt-janus.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matt-janus-noob-encryption__1-39874/archive/master.zip)





### Source Code

```
Function RandNum(min, max)
reRand:
DoEvents
Randomize
X = (Rnd * (max * 10)) / 10
If X > max Then GoTo reRand
If X < min Then GoTo reRand
X = Format(X, "0##")
RandNum = X
End Function
Function Encrypt(theString)
Dim theLen As Long, i As Long
Dim TempStr1 As String, TempStr2 As String, TempStr3 As String
Dim finalStr As String, tempChar As String
theLen = Len(theString)
For i = 1 To theLen Step 2
  TempStr1 = TempStr1 & Mid(theString, i, 1)
  TempStr2 = TempStr2 & Mid(theString, i + 1, 1)
Next
TempStr3 = LCase(TempStr2 & TempStr1)
For i = 1 To theLen
curChar = Mid(TempStr3, i, 1)
  Select Case curChar
    Case "a"
      tempChar = RandNum(969, 999)
    Case "b"
      tempChar = RandNum(50, 80)
    Case "c"
      tempChar = RandNum(939, 968)
    Case "d"
      tempChar = RandNum(126, 132)
    Case "e"
      tempChar = RandNum(901, 938)
    Case "f"
      tempChar = RandNum(888, 900)
    Case "g"
      tempChar = RandNum(81, 105)
    Case "h"
      tempChar = RandNum(106, 125)
    Case "i"
      tempChar = RandNum(133, 147)
    Case "j"
      tempChar = RandNum(601, 621)
    Case "k"
      tempChar = RandNum(148, 178)
    Case "l"
      tempChar = RandNum(179, 209)
    Case "m"
      tempChar = RandNum(209, 230)
    Case "n"
      tempChar = RandNum(231, 265)
    Case "o"
      tempChar = RandNum(622, 651)
    Case "p"
      tempChar = RandNum(266, 296)
    Case "q"
      tempChar = RandNum(297, 321)
    Case "r"
      tempChar = RandNum(652, 681)
    Case "s"
      tempChar = RandNum(322, 352)
    Case "t"
      tempChar = RandNum(682, 705)
    Case "u"
      tempChar = RandNum(353, 381)
    Case "v"
      tempChar = RandNum(382, 405)
    Case "w"
      tempChar = RandNum(406, 424)
    Case "x"
      tempChar = RandNum(424, 436)
    Case "y"
      tempChar = RandNum(437, 461)
    Case "z"
      tempChar = RandNum(462, 492)
    Case "."
      tempChar = RandNum(493, 505)
    Case ","
      tempChar = RandNum(506, 520)
    Case "?"
      tempChar = RandNum(521, 531)
    Case "!"
      tempChar = RandNum(532, 545)
    Case " "
      tempChar = RandNum(546, 600)
    Case Else
      tempChar = RandNum(493, 505)
  End Select
  DoEvents
finalStr = finalStr & tempChar
Next
finalStr = Replace(finalStr, "0", "A")
finalStr = Replace(finalStr, "1", "B")
finalStr = Replace(finalStr, "2", "C")
finalStr = Replace(finalStr, "3", "D")
finalStr = Replace(finalStr, "4", "E")
finalStr = Replace(finalStr, "5", "F")
finalStr = Replace(finalStr, "6", "G")
finalStr = Replace(finalStr, "7", "H")
finalStr = Replace(finalStr, "8", "I")
finalStr = Replace(finalStr, "9", "J")
Encrypt = finalStr
End Function
Function Decrypt(theStr)
theString = theStr
Dim theLen As Long
Dim curChar As String, tempChar As String, _
 temp1 As String, temp2 As String, finalStr As String
theString = Replace(theString, "A", "0")
theString = Replace(theString, "B", "1")
theString = Replace(theString, "C", "2")
theString = Replace(theString, "D", "3")
theString = Replace(theString, "E", "4")
theString = Replace(theString, "F", "5")
theString = Replace(theString, "G", "6")
theString = Replace(theString, "H", "7")
theString = Replace(theString, "I", "8")
theString = Replace(theString, "J", "9")
theLen = Len(theString) - 1
For i = 1 To theLen Step 3
curChar = Mid(theString, i, 3)
  Select Case Int(curChar)
    Case 969 To 999
      tempChar = "a"
    Case 66 To 80
      tempChar = "b"
    Case 939 To 968
      tempChar = "c"
    Case 126 To 132
      tempChar = "d"
    Case 901 To 938
      tempChar = "e"
    Case 888 To 900
      tempChar = "f"
    Case 81 To 105
      tempChar = "g"
    Case 106 To 125
      tempChar = "h"
    Case 133 To 147
      tempChar = "i"
    Case 601 To 621
      tempChar = "j"
    Case 148 To 178
      tempChar = "k"
    Case 179 To 209
      tempChar = "l"
    Case 210 To 230
      tempChar = "m"
    Case 231 To 265
      tempChar = "n"
    Case 622 To 651
      tempChar = "o"
    Case 266 To 296
      tempChar = "p"
    Case 297 To 321
      tempChar = "q"
    Case 652 To 681
      tempChar = "r"
    Case 322 To 352
      tempChar = "s"
    Case 682 To 705
      tempChar = "t"
    Case 353 To 381
      tempChar = "u"
    Case 382 To 405
      tempChar = "v"
    Case 406 To 424
      tempChar = "w"
    Case 425 To 436
      tempChar = "x"
    Case 437 To 461
      tempChar = "y"
    Case 462 To 492
      tempChar = "z"
    Case 493 To 505
      tempChar = "."
    Case 506 To 520
      tempChar = ","
    Case 521 To 531
      tempChar = "?"
    Case 532 To 545
      tempChar = "!"
    Case 546 To 600
      tempChar = " "
  End Select
  DoEvents
TempStr1 = TempStr1 & tempChar
Next
theLen = Len(TempStr1)
If theLen Mod 2 <> 0 Then
  temp1 = Mid(TempStr1, 1, (theLen / 2) - 0.5)
  temp2 = Mid(TempStr1, (theLen / 2) + 0.5, theLen)
  theLen = Len(temp2)
  Else
  temp1 = Mid(TempStr1, 1, (theLen / 2))
  temp2 = Mid(TempStr1, (theLen / 2) + 1, theLen)
  theLen = Len(temp2)
End If
For i = 1 To theLen
  finalStr = finalStr & Mid(temp2, i, 1) & Mid(temp1, i, 1)
Next
Decrypt = finalStr
End Function
```

