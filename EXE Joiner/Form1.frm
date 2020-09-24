VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXE Joiner"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Join"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   3630
      TabIndex        =   8
      Top             =   960
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3630
      TabIndex        =   7
      Top             =   570
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3630
      TabIndex        =   6
      Top             =   180
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   900
      TabIndex        =   5
      Text            =   "Test.exe"
      Top             =   960
      Width           =   2600
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   570
      Width           =   2600
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   180
      Width           =   2600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Patch:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   990
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File2:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   330
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File1:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   330
      TabIndex        =   0
      Top             =   210
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright © 2002 By TSH
Option Explicit
Dim CommonDialog As New GCommonDialog
Dim File1 As String, File2 As String

Private Sub Command1_Click(Index As Integer)
On Error GoTo errhandle
Dim i As Integer

Select Case Index
    Case 0 'open File1
         If (CommonDialog.VBGetOpenFileName(File1, , , , , , "Program Files (*.EXE)|*.EXE", , , "Open File", "EXE", Me.hwnd)) Then
            Text1(0).Text = File1
            Text1(0).Text = Trim(Text1(0).Text)
         End If
    
    Case 1 'open File2
         If (CommonDialog.VBGetOpenFileName(File2, , , , , , "Program Files (*.EXE)|*.EXE", , , "Open File", "EXE", Me.hwnd)) Then
            Text1(1).Text = File2
            Text1(1).Text = Trim(Text1(1).Text)
         End If
    
    Case 2 'join files
    Dim a As Long
         For i = 0 To 2
             If Text1(i).Text = "" Then
                MsgBox "Please Select file and input patch file name.", vbExclamation, "EXE Joiner"
                Exit Sub
             ElseIf Right$(UCase(Text1(i).Text), 4) <> ".EXE" Then
                MsgBox "Please select program file or input execute file.", vbExclamation, "EXE Joiner"
                Exit Sub
             End If
         Next i
         
         If Mid$(Trim(Text1(2).Text), 2, 1) = ":" Then
            MsgBox "Please input patch file name only.", vbExclamation, "EXE Joiner"
            Exit Sub
         End If
         
         'if File1 is zero byte
         If FileLen(Text1(0).Text) = 0 Then
            MsgBox "File1 cannot is zero byte.", vbExclamation, "EXE Joiner"
            Exit Sub
         End If
         'if File2 is zero byte
         If FileLen(Text1(1).Text) = 0 Then
            MsgBox "File2 cannot is zero byte.", vbExclamation, "EXE Joiner"
            Exit Sub
         End If
         
         'if patch file exists, delete it
         If Dir(cPath & Trim(Text1(2).Text), vbNormal) <> "" Then
            Kill cPath & Trim(Text1(2).Text)
         End If
         'join the 2 EXE files
         Call Joiner(File1, File2)

End Select

Exit Sub
errhandle:
If Err.Number = 75 Then
   MsgBox "Patch file is running!", vbExclamation, "EXE Joiner"
End If
Exit Sub
MsgBox Err.Description, vbExclamation, "EXE Joiner"
End Sub

Private Sub Joiner(Join1 As String, Join2 As String)
On Error GoTo errhandle
Dim p1Array() As Byte, p2Array() As Byte
Dim MySize1 As Long, MySize2 As Long, pLength As Long
    
    Open Join1 For Binary Access Read As #1
         ReDim p1Array(LOF(1) - 1)
         MySize1 = LOF(1)
         Get #1, , p1Array
    Close #1
    
    Open Join2 For Binary Access Read As #1
         ReDim p2Array(LOF(1) - 1)
         MySize2 = LOF(1)
         Get #1, , p2Array
    Close #1
    
    'get the patch file
    Call LoadDataIntoFile(101, cPath & Trim(Text1(2).Text))
    'get the patch file size
    pLength = FileLen(cPath & Trim(Text1(2).Text))
    
    'Join 2 files together
    Open cPath & Trim(Text1(2).Text) For Binary Access Write As #1
         Put #1, pLength, p1Array 'put File1
         Put #1, pLength + MySize1, p2Array 'put File2
         'put 2 files size mark at last
         Put #1, pLength + MySize1 + MySize2, "[LEN]" & MySize1 & "," & MySize2
    Close #1

MsgBox "EXE joiner successful!", vbInformation, "EXE Joiner"

Exit Sub
errhandle:
MsgBox Err.Description, vbExclamation, "EXE Joiner"
End Sub

Private Function cPath() As String
If Right$(App.Path, 1) <> "\" Then
   cPath = App.Path & "\"
Else
   cPath = App.Path
End If
End Function

Private Sub LoadDataIntoFile(DataName As Integer, Filename As String)
Dim myArray() As Byte
Dim myFile As Long
    
    If Dir(Filename) = "" Then
       myArray = LoadResData(DataName, "CUSTOM")
       myFile = FreeFile
       Open Filename For Binary Access Write As #myFile
         Put #myFile, , myArray
       Close #myFile
    End If
    
End Sub

Private Sub Form_Load()
Me.Caption = "EXE Joiner v" & App.Major & "." & App.Minor & App.Revision & " By TSH"
End Sub
