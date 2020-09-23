<div align="center">

## CopyMemory API, A Simple Tutorial


</div>

### Description

This gives a step by step explaination of the CopyMemory API Call.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-02-06 18:39:46
**By**             |[Sean Dittmar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sean-dittmar.md)
**Level**          |Intermediate
**User Rating**    |4.7 (33 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CopyMemory53519262002\.zip](https://github.com/Planet-Source-Code/sean-dittmar-copymemory-api-a-simple-tutorial__1-31555/archive/master.zip)

### API Declarations

```
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
```





