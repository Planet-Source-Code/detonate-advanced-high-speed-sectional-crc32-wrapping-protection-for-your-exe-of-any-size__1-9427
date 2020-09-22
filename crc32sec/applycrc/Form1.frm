VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Apply CRC32"
   ClientHeight    =   540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   540
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "YOU MUST EDIT THE CODE FOR THIS PROGRAM BEFORE RUNNING IT!"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sectional CRC32 Protection (High-Speed) for your Exe
'by Detonate (detonate@start.com.au), for PlanetSourceCode.com

'You are free to use this in your commercial/uncommercial/whatever
'programs, and I don't for any credit, but if you do publish this
'source code, please keep it intact and unmodified, that is all i ask.

'The idea for this stemmed because the original CRC32 wrapper
'i released at planetsourcecode.com/vb which used the entire file
'was too slow for practical use if the exe was over a few hundred
'kb in size. It's a relatively solid protection, as it's based on
'every byte in the file, meaning it's virtually impossible to change
'any bytes without changing the hash.

'Speed was the problem. If you don't understand why, try this:
'    For I = 1 to 5000000
'     Doevents
'    Next I
'That's looping through just 5 megabytes, and remember it's doing
'NO processing either!  yet it's still slow... so if we can't even do
'that, how can we get an effective CRC32 wrapper over all the bytes?
'... buggered if i know! :-)

'So i put my mind down a bit... and thought... we only really need to
'protect certain parts of the file... and if those parts of the file
'included the Exe header, the CRC32 instructions, the checksum , and
'various other fields, then the protection would be just about as good
'as one applied on the full file. The only catch is that you have to
'identify which parts of your exe are important... this is dead easy
'if you have a hex editor.

'Enter Sectional CRC32 Protection!
'This version only uses the sections of your .exe that you deem important
'For example, there is probably no need to protect a BMP image in your
'exe file, but you may want to protect where it says "UNLICENSED" :-)
'To do this, use a hex editor to find the byte entry point, and then
'figure out how many bytes you want to protect... whether its 1 or 100000
'or whatever. You add this string to the CRCTest string, which holds
'all of the "important" strings.
'After youve filled CRCTest with all of your important strings, you
'create a CRC32 hash on THAT, and then append that to the end of the
'file.

'This example uses two "protected fields" - the first 500 bytes of the
'file, ie. the exe header, and also a second imaginary field which could
'contain whatever, just to demonstrate how to select fields. Because
'its reading directly from the disk and not from memory, it doesnt matter
'if youre file is 1 meg or 100 megs.. the file contents are never read
'into memory

'---CODE STARTS---

'*************************************************
'*************************************************
Const ReadFile = "C:\crc32sec\program\myprog.exe"  '<- the exe to be protected
'*************************************************
'*************************************************

Private Sub Form_Load()
On Error Resume Next
Dim FileSize As Long
Dim NextString$
Dim CRCTest$
Dim I As Long

TotalValue = 0
FileSize = FileLen(ReadFile)
WholeFile$ = String$(FileSize, 0)
CRCTest$ = Trim(CStr(FileSize))  'reset

Open ReadFile For Binary Access Read As #1
'THIS IS THE ONLY BIT YOULL HAVE TO MODIFY
'THESE ARE THE PROTECTED FIELDS
'A hex editor is handy here :-)
 'Our first important string... the exe header, we'll grab the first 600 bytes
 NextString$ = String(600, 0) 'Reset NextString$ to a size of 600 bytes
 Get #1, 1, NextString$       'and fill the value from byte 1
 CRCTest$ = CRCTest$ & NextString$   'Add the string to our CRCTest string
  
 'The next "important-looking" field in the file starts at about byte 4096,
 'and only goes for about 800 bytes... we'll protect that too
 NextString$ = String(800, 0) 'Reset NextString$ to a size of 800 bytes
 Get #1, 4096, NextString$       'and fill the value from byte 4096
 CRCTest$ = CRCTest$ & NextString$   'Add the string to our CRCTest string

 'We can basically protect the ENTIRE registration section of the program
 'by reading in the last 12300 bytes of the file (73664 looks like a good
 'place to start - as seen in a hex editor, this is where the registration part lies.
 NextString$ = String(12300, 0) 'Reset NextString$ to a size of 12300 bytes
 Get #1, 73664, NextString$       'and fill the value from byte 73664
 CRCTest$ = CRCTest$ & NextString$   'Add the string to our CRCTest string
Close #1

'CRCTest now holds all of our important strings, bundled into the one string.
'We'll now make a CRC32 hash from that...
Dim CRC32 As String * 8
Dim OriginalCRC As String * 8
CRC32 = Compute(CRCTest)
OriginalCRC = CRC32

'mildly (but quickly) scramble the hash so it's not so plainly visible
For I = 1 To 8
 Mid(CRC32, I, 1) = Chr$(Asc(Mid(CRC32, I, 1)) Xor 30)
Next I

Open ReadFile For Binary Access Write As #1
 Put #1, FileLen(ReadFile) + 1, CRC32
Close #1
MsgBox "CRC32=" & OriginalCRC & vbCrLf & ReadFile & " has been protected!"
End
End Sub
