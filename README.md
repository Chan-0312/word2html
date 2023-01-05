# doc/docx转 html的方法
- 作者：熊盛春
- 时间：2023.1.5

## 方法1: 利用VBA
- 利用windows环境下word中的VB语言进行批量转换
    - 优点：转换后的html格式完整，支持doc和docx
    - 缺点：依赖windows环境，并且需要文档开启了宏，未开启宏的无法进行转换。
- VB代码如下：
```vb
Sub ConvertDocuments()
    Application.ScreenUpdating = False
    Dim strFolder As String, strFiles As String, strDocNm As String, wdDoc As Document
    strDocNm = ActiveDocument.FullName
    strFolder = GetFolder
    Debug.Print strFolder
    If strFolder = "" Then Exit Sub
    strFiles = LoopThroughFiles(strFolder, ".docx", True)

    Dim iFiles() As String
    iFiles() = Split(strFiles, vbTab)

    Dim i As Long
    For i = LBound(iFiles) To UBound(iFiles)
        If iFiles(i) <> "" And iFiles(i) <> strDocNm Then
            On Error Resume Next
            Set wdDoc = Documents.Open(FileName:=iFiles(i), AddToRecentFiles:=False, Visible:=False)
            
            With wdDoc
            .SaveAs2 FileName:=Split(iFiles(i), ".doc")(0) & ".html", FileFormat:=wdFormatFilteredHTML, AddToRecentFiles:=False, Encoding:=msoEncodingUTF8
            .Close SaveChanges:=True
            End With
            
        End If
    Next i
    Set wdDoc = Nothing
    Application.ScreenUpdating = True
    MsgBox "Converted " & UBound(iFiles)
End Sub

Private Function LoopThroughFiles(inputDirectory As String, filenameCriteria As String, doTraverse As Boolean) As String
    Dim tmpOut As String
    Dim StrFile As String

    If doTraverse = True Then
        Dim allFolders As String
        Dim iFolders() As String
        allFolders = TraverseDir(inputDirectory & "\", 1, 100)
        iFolders() = Split(allFolders, vbTab)

        tmpOut = LoopThroughFiles(inputDirectory, filenameCriteria, False)
        Dim j As Long
        For j = LBound(iFolders) To UBound(iFolders)
            If iFolders(j) <> "" Then
                StrFile = LoopThroughFiles(iFolders(j), filenameCriteria, False)
                tmpOut = tmpOut & vbTab & StrFile
            End If
        Next j
        LoopThroughFiles = tmpOut
    Else
        'https://stackoverflow.com/a/45749626/4650297
        StrFile = Dir(inputDirectory & "\*" & filenameCriteria)
        Do While Len(StrFile) > 0
            tmpOut = tmpOut & vbTab & inputDirectory & "\" & StrFile
            StrFile = Dir()
        Loop
        LoopThroughFiles = tmpOut
    End If
End Function

Private Function TraverseDir(path As String, depth As Long, maxDepth As Long) As String
    'https://analystcave.com/vba-dir-function-how-to-traverse-directories/#Traversing_directories
    If depth > maxDepth Then
        TraverseDir = ""
        Exit Function
    End If
    Dim currentPath As String, directory As Variant
    Dim dirCollection As Collection
    Set dirCollection = New Collection
    Dim dirString As String

    currentPath = Dir(path, vbDirectory)

    'Explore current directory
    Do Until currentPath = vbNullString
        ' Debug.Print currentPath
        If Left(currentPath, 1) <> "." And (GetAttr(path & currentPath) And vbDirectory) = vbDirectory Then
            dirString = dirString & vbTab & path & currentPath
            dirCollection.Add currentPath
        End If
        currentPath = Dir()
    Loop

    TraverseDir = dirString
    'Explore subsequent directories
    For Each directory In dirCollection
        TraverseDir = TraverseDir & vbTab & TraverseDir(path & directory & "\", depth + 1, maxDepth)
    Next directory
End Function

Function GetFolder() As String
    Dim oFolder As Object
    GetFolder = ""
    Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0, "Please select the folder containing the Word documents to process", 0)
    If (Not oFolder Is Nothing) Then GetFolder = oFolder.Items.Item.path
    Set oFolder = Nothing
End Function
```

## 方法2：利用pywin32包
- 基于python语言，在windows环境下利用调用word的接口进行转换。本质上同方法1一样都是调用word自带的接口。
  - 优点：同方法1，并且利用脚本来进行批量转换，相对更方便。
  - 缺点：同方法1。
- 代码：[word2html_pywin32.py](word2html_pywin32.py)
- 
## 方法3：利用PyDocX包
- 基于python中的pyDocX包进行转换。
  - 优点：不依赖于windwos环境
  - 缺点：仅能对docx格式进行转换；转换后的html效果不如方法1和方法2(特别是多级列表的情况下)
- 代码：[word2html_pydocx.py](word2html_pydocx.py)

## 方法4：利用mammoth包
- 基于python中的mammoth包进行转换。
  - 优点：不依赖于windwos环境，转换效果优于pydocx包
  - 缺点：仅能对docx文件进行转换
- 代码：[word2html_mammoth.py](word2html_mammoth.py)

## 方法5：利用pandoc
- 基于pandoc进行转换。
  - 优点：可脚本批量处理
  - 缺点：转换后的效果较差，无法处理多级列表和标题
- 转换代码如下：`pandoc -s test.docx -o test.html`
- pandoc安装：`apt-get install pandoc`

## 方法6：利用libreoffice和mammoth
- 对于docx文件，直接用mammoth包进行转换；对于doc文件，需要先利用libreoffice转换成docx文件，并且libreoffice不依赖windows环境
  - 优点：libreoffice和mammoth的结合可以实现在linux环境下将doc/docx比较完美地转换为html格式
  - 缺点：libreoffice转换后的docx文件好像在word打不开(待确认？)，但是可以利用mammoth转换为html(无损坏)
- libreoffice安装：
  > apt-get install libreoffice-common libreoffice-writer
  > apt-get install default-jre libreoffice-java-common
- 利用libreoffice转换docx的命令：`libreoffice --invisible -convert-to docx:"MS Word 2007 XML" ./testdoc.doc`

