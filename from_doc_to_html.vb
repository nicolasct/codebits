' This file is Git managed.
'
' Full comments on why this file and what it does: see at <a href="../../../../computing/lib/WordToHtml_VBA_script.doc#final_script"/> : please tell this href node if the file is moved or deleted.


Declare PtrSafe Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long 'à mettre en début de file, c'est un library call pour la création de multiples directories, genre "dir1/dir2/dir3/..."

Sub ChangeDocsToTxtOrRTFOrHTML()
'with export to PDF in Word 2007
    Dim fs As Object

    Dim strDocName As String
    Dim strDocRelativePath As String 'for ex: "Know_Yourself\ai\plymouth\iCubSim", while strDocName would be "iCubSim_notes" (le ".doc" sera enlevé et remplacé par l'extension voulu)
    Dim intPos As Integer
    Dim locFolder As String
    Dim buildFolder As String
    Dim fileType As String
    
    On Error Resume Next
    
    buildFolder = "C:\Users\pippo\cne"
    locFolder = InputBox("Enter the folder path to DOCs", "File Conversion", "C:\Users\pippo\Documents\Encyclopedia\Know_Yourself\ai\plymouth")
    Select Case Application.Version
        Case Is < 12
            Do
                fileType = UCase(InputBox("Change DOC to TXT, RTF, HTML", "File Conversion", "TXT"))
            Loop Until (fileType = "TXT" Or fileType = "RTF" Or fileType = "HTML")
        Case Is >= 12
            Do
                fileType = UCase(InputBox("Change DOC to TXT, RTF, HTML or PDF(2007+ only)", "File Conversion", "HTML"))
            Loop Until (fileType = "TXT" Or fileType = "RTF" Or fileType = "HTML" Or fileType = "PDF")
    End Select
    Application.ScreenUpdating = False
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    
    Dim files As New Collection
    GetFilesRecursive fs.GetFolder(locFolder), "doc", files, fs
    
    Dim oFile As Scripting.file
    For Each oFile In files
        Dim d As Document
        Set d = Application.Documents.Open(oFile.Path)
        
        ' only process if the file doesn't contain the string "DONTPUBLISH":
        With ActiveDocument.Content.Find
            .Text = "DONTPUBLISH"
            .Forward = True
            .Execute
            If .Found = False Then
            
            
                strDocName = ActiveDocument.Name
                intPos = InStrRev(strDocName, ".")
                strDocName = Left(strDocName, intPos - 1)
                strDocRelativePath = getMedianPath(oFile.parentFolder, "C:\Users\pippo\Documents\Encyclopedia", fs)
                'MsgBox "strDocRelativePath is " & strDocRelativePath
                Dim oBuildFolder_complete_String As String
                oBuildFolder_complete_String = buildFolder & "\" & strDocRelativePath
                MakeFullDir (oBuildFolder_complete_String)
                Set oBuildFolder_complete = fs.GetFolder(oBuildFolder_complete_String)
                ChangeFileOpenDirectory oBuildFolder_complete
                'MsgBox "oBuildFolder_complete is " & oBuildFolder_complete
            
                      
                    
                        Select Case fileType
                        Case Is = "TXT"
                            strDocName = strDocName & ".txt"
                            ActiveDocument.SaveAs FileName:=strDocName, FileFormat:=wdFormatText
                        Case Is = "RTF"
                            strDocName = strDocName & ".rtf"
                            ActiveDocument.SaveAs FileName:=strDocName, FileFormat:=wdFormatRTF
                            
                        Case Is = "HTML"
                            strDocName = strDocName & ".html"
                            ActiveDocument.SaveAs FileName:=strDocName, FileFormat:=wdFormatFilteredHTML
                            
                                'Loop through all hyperlinks and change .doc extension for .html
                                    Dim link_to_doc As String
                                    Dim link_to_html As String
                                    Set RegEx = CreateObject("vbscript.regexp")
                                For i = 1 To ActiveDocument.Hyperlinks.Count
                                    link_to_doc = ActiveDocument.Hyperlinks(i).Address
                                        With RegEx
                                            .IgnoreCase = True
                                            .Global = True
                                            .Pattern = "(.*).doc"
                                        End With
                                        link_to_html = RegEx.Replace(link_to_doc, "$1.html")
                                    ActiveDocument.Hyperlinks(i).Address = link_to_html
                                Next
                                
                                ' Change also any display names, if they contain a ".doc" extension:
                                With ActiveDocument.Content.Find
                                    .Text = ".doc"
                                    .Replacement.ClearFormatting
                                    .Replacement.Text = ".html"
                                    .Execute Replace:=wdReplaceAll, Forward:=True, _
                                    Wrap:=wdFindContinue
                                End With
                                
                                ActiveDocument.Save
                                
                
                
                
                        Case Is = "PDF"
                            strDocName = strDocName & ".pdf"
                            ' *** Word 2007 users - remove the apostrophe at the start of the next line ***
                            'ActiveDocument.ExportAsFixedFormat OutputFileName:=strDocName, ExportFormat:=wdExportFormatPDF
                            
                        End Select
                
            End If '(If .Found = False Then)
        End With '(With ActiveDocument.Content.Find : only process if the file doesn't contain the string "DONTPUBLISH":
       
    d.Close
    'ChangeFileOpenDirectory oFolder
    Next oFile
    
    Application.ScreenUpdating = True
    
End Sub



 
Public Sub MakeFullDir(strPath As String)
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\" 'Optional depending upon intent
    MakeSureDirectoryPathExists strPath
End Sub

Sub testMakeFullDir()
 MakeFullDir ("C:\Users\pippo\cne\d1\d2")
End Sub



' test for Function getMedianPath

Sub testGetMedianPath()
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    MsgBox "Returning " & getMedianPath("C:\Users\pippo\Documents\Encyclopedia\Know_Yourself\ai\plymouth\iCubSim", "C:\Users\pippo\Documents\Encyclopedia", fs)
End Sub

' get the relative path starting from a root path.
' For ex., with currentPath = "C:\Users\pippo\Documents\Encyclopedia\Know_Yourself\ai\plymouth\iCubSim"
' and given rootPath of "C:\Users\pippo\Documents\Encyclopedia", it will return "Know_Yourself\ai\plymouth\iCubSim"
' Use testGetMedianPath() for testing.

Function getMedianPath(currentPath As String, rootPath As String, fs As Scripting.FileSystemObject) As String

    Dim parentFolderS As String
    
    Dim medianPath As String
     
    Dim inputFolder As Object
    Set inputFolder = fs.GetFolder(currentPath)
    
    'MsgBox "inputFolder is" & inputFolder
    
    Set parent_Folder = inputFolder
    Do While parent_Folder <> rootPath
        Set parent_Folder = parent_Folder.parentFolder
        medianPath = Right(currentPath, Len(currentPath) - Len(parent_Folder) - 1)
    Loop
    'MsgBox "End of loop, parentFolder is " & parent_Folder
    'MsgBox "And medianPath is " & medianPath
    
    getMedianPath = medianPath
    
    
End Function



Sub GetFilesRecursive(f As Scripting.Folder, filter As String, c As Collection, fso As Scripting.FileSystemObject)
  Dim sf As Scripting.Folder
  Dim file As Scripting.file

  For Each file In f.files
    If InStr(1, fso.GetExtensionName(file.Name), filter, vbTextCompare) = 1 Then
      c.Add file, file.Path
    End If
  Next file

  For Each sf In f.SubFolders
    GetFilesRecursive sf, filter, c, fso
  Next sf
End Sub



