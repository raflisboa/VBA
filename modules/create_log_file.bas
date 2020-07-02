Attribute VB_Name = "create_log_file"
Sub LogInformation(LogMessage As String)

Const LogFileName As String = "\\Placeholder\Placeholder\Placeholder.log"
Dim FileNum As Integer
    FileNum = FreeFile ' next file number
        Open LogFileName For Append As #FileNum ' creates the file if it doesn't exist
            Print #FileNum, LogMessage ' write information at the end of the text file
        Close #FileNum ' close the file
        
        ' backup process
        FileCopy "\\Placeholder\Placeholder\Placeholder.log", _
        "\\\Placeholder\Placeholder\Placeholder.csv"

End Sub

Public Sub DisplayLastLogInformation()

Const LogFileName As String = "\\Brpced0209\d#\Projetos\CHECKSUM\double_check_control.LOG"
Dim FileNum As Integer, tLine As String
    FileNum = FreeFile ' next file number
        Open LogFileName For Input Access Read Shared As #f ' open the file for reading
            Do While Not EOF(FileNum)
                Line Input #FileNum, tLine ' read a line from the text file
                    Loop ' until the last line is read
            Close #FileNum ' close the file

MsgBox tLine, vbInformation, "Last log information:"

End Sub
Sub DeleteLogFile(FullFileName As String)

On Error Resume Next ' ignore possible errors
    Kill FullFileName ' delete the file if it exists and it is possible
On Error GoTo 0 ' break on errors

End Sub
