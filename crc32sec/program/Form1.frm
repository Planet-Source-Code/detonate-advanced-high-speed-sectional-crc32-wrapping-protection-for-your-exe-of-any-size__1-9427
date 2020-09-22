VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "My CRC32 Sectionally-Protected Program"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1005
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   945
      ScaleWidth      =   5475
      TabIndex        =   4
      Top             =   0
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Program Registration"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enter your registration serial:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1980
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
'Boys and girls, please do NOT try to use this type of
'registration at home... it may prove dangerous to your
'application, and should be attempted only by fools! :-)
If Text1.Text = "theserial" Then
   MsgBox "Thankyou for registering my program, i'll be able to feed my kids this week"
Else
   MsgBox "Incorrect serial! Please pay for my program! :-)" & vbCrLf & "(or look at the source or something! ;-)"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim OurCRC As String
Dim FileCRC As String * 8
Dim FPos As Long
Dim I As Long
OurCRC = CheckCRC
FPos = FileLen(AppExe)
Open AppExe For Binary As #1
 Get #1, FPos - 7, FileCRC   'read the last 8 bytes of the file
Close #1

'Descramble the hash
For I = 1 To 8
 Mid(FileCRC, I, 1) = Chr$(Asc(Mid(FileCRC, I, 1)) Xor 30)
Next I

'Make sure the CRC hash is present
If OurCRC = "        " Or OurCRC = String(8, 0) Then
   MsgBox "This file hasn't been wrapped - dont forget to do that before you release it! :-)"
   End
End If

'Make sure the FileCRC is the same as the actual CRC
If FileCRC = OurCRC Then
   MsgBox "CRC check ok, file doesnt appear to have been tampered with"
Else
   MsgBox "CRC values have changed!" & vbCrLf & _
          " Its actually: " & OurCRC & vbCrLf & _
          " It should be: " & FileCRC & vbCrLf
   End
End If
End Sub

