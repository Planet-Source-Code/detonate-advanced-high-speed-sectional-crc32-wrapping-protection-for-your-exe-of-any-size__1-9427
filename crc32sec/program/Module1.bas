Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text
'// Then declare this array variable Crc32Table
Private Crc32Table(255) As Long
'// Then all we have to do is writing public functions like these...

Public Function CheckCRC() As String
On Error Resume Next
Dim FileSize As Long
Dim NextString$
Dim CRCTest$
FileSize = FileLen(AppExe) - 8   'the size of the file, minus the crc32 hash
CRCTest$ = Trim(CStr(FileSize))  'reset the CRCTest string...
'we could have set that to "", but as an additional security measure,
'we'll make it the filesize - thus, if the size of the file
'changes, so will the checksum - regardless of what sections
'you have protected. :-)
Open AppExe For Binary Access Read As #1
'Now we read in the same fields ... exactly how
'we did it in the ApplyCRC program.
 
 NextString$ = String(600, 0) 'Reset NextString$ to a size of 600 bytes
 Get #1, 1, NextString$       'and fill the value from byte 1
 CRCTest$ = CRCTest$ & NextString$   'Add the string to our CRCTest string
  
 NextString$ = String(800, 0) 'Reset NextString$ to a size of 800 bytes
 Get #1, 4096, NextString$       'and fill the value from byte 4096
 CRCTest$ = CRCTest$ & NextString$   'Add the string to our CRCTest string

 NextString$ = String(12300, 0) 'Reset NextString$ to a size of 12300 bytes
 Get #1, 73664, NextString$       'and fill the value from byte 73664
 CRCTest$ = CRCTest$ & NextString$   'Add the string to our CRCTest string

Close #1

'CRCTest now holds all of our important strings, bundled into the one string.
'We'll now make a CRC32 hash from that...
Dim CRC32 As String
CheckCRC = Compute(CRCTest)
End Function

Public Function InitCrc32(Optional ByVal Seed As Long = &HEDB88320, Optional ByVal Precondition As Long = &HFFFFFFFF) As Long
    '// Declare counter variable iBytes, counter variable iBits, value variables lCrc32 and lTempCrc32
    Dim iBytes As Long, iBits As Integer, lCrc32 As Long, lTempCrc32 As Long
    '// Turn on error trapping
    On Error Resume Next
    '// Iterate 256 times

    For iBytes = 0 To 255
        '// Initiate lCrc32 to counter variable
        lCrc32 = iBytes
        '// Now iterate through each bit in counter byte


        For iBits = 0 To 7
            '// Right shift unsigned long 1 bit
            lTempCrc32 = lCrc32 And &HFFFFFFFE
            lTempCrc32 = lTempCrc32 \ &H2
            lTempCrc32 = lTempCrc32 And &H7FFFFFFF
            '// Now check if temporary is less than zero and then mix Crc32 checksum with Seed value


            If (lCrc32 And &H1) <> 0 Then
                lCrc32 = lTempCrc32 Xor Seed
            Else
                lCrc32 = lTempCrc32
            End If
        Next
        '// Put Crc32 checksum value in the holding array
        Crc32Table(iBytes) = lCrc32
    Next
    '// After this is done, set function value to the precondition value
    InitCrc32 = Precondition
End Function
'// The function above is the initializing function, now we have to write the computation function


Public Function AddCrc32(ByVal Item As String, ByVal CRC32 As Long) As Long
    '// Declare following variables
    Dim bCharValue As Byte, iCounter As Long, lIndex As Long
    Dim lAccValue As Long, lTableValue As Long
    '// Turn on error trapping
    On Error Resume Next
    '// Iterate through the string that is to be checksum-computed


    For iCounter = 1 To Len(Item)
        '// Get ASCII value for the current character
        bCharValue = Asc(Mid$(Item, iCounter, 1))
        '// Right shift an Unsigned Long 8 bits
        lAccValue = CRC32 And &HFFFFFF00
        lAccValue = lAccValue \ &H100
        lAccValue = lAccValue And &HFFFFFF
        '// Now select the right adding value from the holding table
        lIndex = CRC32 And &HFF
        lIndex = lIndex Xor bCharValue
        lTableValue = Crc32Table(lIndex)
        '// Then mix new Crc32 value with previous accumulated Crc32 value
        CRC32 = lAccValue Xor lTableValue
    Next
    '// Set function value the the new Crc32 checksum
    AddCrc32 = CRC32
End Function
'// At last, we have to write a function so that we can get the Crc32 checksum value at any time


Public Function GetCrc32(ByVal CRC32 As Long) As Long
    '// Turn on error trapping
    On Error Resume Next
    '// Set function to the current Crc32 value
    GetCrc32 = CRC32 Xor &HFFFFFFFF
End Function

'// To Test the Routines Above...
Public Function Compute(ToGet As String) As String
On Error Resume Next
    Dim lCrc32Value As Long
    On Error Resume Next
    lCrc32Value = InitCrc32()
    lCrc32Value = AddCrc32(ToGet, lCrc32Value)
    Compute = Hex$(GetCrc32(lCrc32Value))
End Function
Public Function AppExe() As String
On Error Resume Next
Dim AP As String
AP = App.Path
If Right(AP, 1) <> "\" Then AP = AP & "\"
AppExe = AP & App.EXEName & ".exe"
End Function


