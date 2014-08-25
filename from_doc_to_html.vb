' This file is Git managed.
'
' Full comments on why this file and what it does: see at <a href="../../../../computing/lib/WordToHtml_VBA_script.doc#final_script"/> (path relative to my local file system) : please tell this href node if the file is moved or deleted.

Private Declare PtrSafe Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long 'à mettre en début de file, c'est un library call pour la création de multiples directories, genre "dir1/dir2/dir3/..."
 

Sub InsertTable1_1()
'
' InsertTable1_1 Macro
'
'
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:= _
        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "Table Grid" Then
            .Style = "Table Grid"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
End Sub


Sub ChangeDocsToTxtOrRTFOrHTML()
'with export to PDF in Word 2007
    Dim fs As Object

    Dim strDocName As String
    Dim strDocNameWithoutExtension As String 'keep the name without extension to use it for referencing the directory with with image files
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
    GetFilesRecursive fs.GetFolder(locFolder), "doc", files, fs, False
    
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
                strDocNameWithoutExtension = Left(strDocName, intPos - 1)
                strDocRelativePath = getMedianPath(oFile.parentFolder, "C:\Users\pippo\Documents\Encyclopedia", fs)
                'MsgBox "strDocRelativePath is " & strDocRelativePath
                Dim oBuildFolder_complete_String As String
                oBuildFolder_complete_String = buildFolder & "\" & strDocRelativePath
                MakeFullDir (oBuildFolder_complete_String)
                Set oBuildFolder_complete = fs.GetFolder(oBuildFolder_complete_String)
                'ChangeFileOpenDirectory oBuildFolder_complete // not used any more: we save locally now!
                ChangeFileOpenDirectory oFile.parentFolder
            
                      
                    
                        Select Case fileType
                        Case Is = "TXT"
                            strDocName = strDocNameWithoutExtension & ".txt"
                            ActiveDocument.SaveAs FileName:=strDocName, FileFormat:=wdFormatText
                        Case Is = "RTF"
                            strDocName = strDocNameWithoutExtension & ".rtf"
                            ActiveDocument.SaveAs FileName:=strDocName, FileFormat:=wdFormatRTF
                            
                      
                        Case Is = "HTML"
                            strDocName = strDocNameWithoutExtension & ".html"
                            'ActiveDocument.WebOptions.AlwaysSaveInDefaultEncoding = False (AlwaysSaveInDefaultEncoding ne s’applique pas à l’ActiveDocument level, seulement Application level)
                            ActiveDocument.WebOptions.Encoding = msoEncodingUTF8
                            ActiveDocument.SaveAs FileName:=strDocName, FileFormat:=wdFormatFilteredHTML, Encoding:=msoEncodingUTF8
                            
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
                                
                                ' Change also any link display names, if they contain a ".doc" extension (for coherency and user-friendlyness)
                                With ActiveDocument.Content.Find
                                    .Text = ".doc"
                                    .Replacement.ClearFormatting
                                    .Replacement.Text = ".html"
                                    .Execute Replace:=wdReplaceAll, Forward:=True, _
                                    Wrap:=wdFindContinue
                                End With
                                
                                ActiveDocument.Save
                                
                                d.Close
                                ' MsgBox "Saving ... " & oFile.parentFolder.Path & "\" & strDocName
                                fs.MoveFile oFile.parentFolder.Path & "\" & strDocName, oBuildFolder_complete_String & "\"
                               ' fs.MkDir oBuildFolder_complete_String & "\" & strDocNameWithoutExtension & "_files"
                                MakeFullDir (oBuildFolder_complete_String & "\" & strDocNameWithoutExtension & "_files")
                                fs.MoveFile oFile.parentFolder.Path & "\" & strDocNameWithoutExtension & "_files\*", oBuildFolder_complete_String & "\" & strDocNameWithoutExtension & "_files\"
                                fs.DeleteFolder oFile.parentFolder.Path & "\" & strDocNameWithoutExtension & "_files", True 'better to clean away this empty folder"
                
                
                        Case Is = "PDF"
                            strDocName = strDocNameWithoutExtension & ".pdf"
                            ' *** Word 2007 users - remove the apostrophe at the start of the next line ***
                            'ActiveDocument.ExportAsFixedFormat OutputFileName:=strDocName, ExportFormat:=wdExportFormatPDF
                            
                        End Select
                        

                
            End If '(If .Found = False Then)
        End With '(With ActiveDocument.Content.Find : only process if the file doesn't contain the string "DONTPUBLISH":
       
    
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



Sub GetFilesRecursive(f As Scripting.Folder, filter As String, c As Collection, fso As Scripting.FileSystemObject, recursive As Boolean)
  Dim sf As Scripting.Folder
  Dim file As Scripting.file

    For Each file In f.files
      If InStr(1, fso.GetExtensionName(file.Name), filter, vbTextCompare) = 1 Then
        c.Add file, file.Path
      End If
    Next file

    If recursive = True Then
      For Each sf In f.SubFolders
        GetFilesRecursive sf, filter, c, fso, recursive
      Next sf
    End If
End Sub


Sub Macro1()
'
' Macro1 Macro
'
'
End Sub
Sub clearFormating()
'
' clearFormating Macro
'
'
    Selection.ClearFormatting
End Sub
