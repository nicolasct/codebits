' This file is Git managed.
'
' Full comments on why this file and what it does: see at <a href="../../../../computing/lib/WordToHtml_VBA_script.doc#final_script"/> : please tell this href node if the file is moved or deleted.

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
                'ChangeFileOpenDirectory oBuildFolder_complete // not used any more: we save locally now!
                ChangeFileOpenDirectory oFile.parentFolder
            
                      
                    
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
                                Const OverwriteExisting = True
                                MsgBox "Saving ... " & oFile.parentFolder.Path & "\" & strDocName
                                fs.CopyFile oFile.parentFolder.Path & "\" & strDocName, oBuildFolder_complete_String & "\", OverwriteExisting
                
                
                        Case Is = "PDF"
                            strDocName = strDocName & ".pdf"
                            ' *** Word 2007 users - remove the apostrophe at the start of the next line ***
                            'ActiveDocument.ExportAsFixedFormat OutputFileName:=strDocName, ExportFormat:=wdExportFormatPDF
                            
                        End Select
                        

                
            End If '(If .Found = False Then)
        End With '(With ActiveDocument.Content.Find : only process if the file doesn't contain the string "DONTPUBLISH":
       
    
    'ChangeFileOpenDirectory oFolder
    Next oFile
    
    Application.ScreenUpdating = True
    
End Sub