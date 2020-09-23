VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CopyMemory"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "The CopyMemory API call may look kinda intimidating, but it's actually really simple. Look at the code."
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'The CopyMemory function copies a block of memory from one
'location to another. Simple as that.

Private Sub Form_Load()

Dim MyStr As String

MyStr = "CopyMemory is easy!"

'Okay, so MyStr is now in memory, somewhere.....
'Q. How do we find the address(where in memory) of this variable?
'A. The function StrPtr

Dim MyStrAddr As Long
MyStrAddr = StrPtr(MyStr)

'We now have the location, in memory, of our MyStr variable.
'Of course, we need to reserve some space in memory for our copied
'string. Its important that we space out MyNewStr.

Dim MyNewStr As String
MyNewStr = Space(Len(MyStr))
'Now we find the address of our MyNewStr variable

Dim MyNewStrAddr As Long

MyNewStrAddr = StrPtr(MyNewStr)

'Now is the moment of truth. In this case, byval is very important because
'the CopyMemory call doesn't know what your intentions are, which is passing
'the address values by value and NOT by Reference. LenB is different
'from Len because it count the number of bytes, not the length of the string.

CopyMemory ByVal MyNewStrAddr, ByVal MyStrAddr, LenB(MyStr)

MsgBox MyNewStr






End Sub
