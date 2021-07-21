Attribute VB_Name = "ModEventBackup"
Option Explicit
'イベント用のプロシージャ
'中身をコピーして使う

Sub ワークブック保存時にフォルダに上書きバックアップ()
'ワークブック保存時にフォルダに上書きバックアップ
'20210721

    Dim FilePath$, FolderName$, BookName$
    FilePath = ThisWorkbook.Path
    FolderName = "Backup" 'バックアップするフォルダの名前←←←←←←←←←←←←←←←←←←←←←←←
    BookName = ThisWorkbook.Name
    
    If Dir(FilePath & "\" & FolderName, vbDirectory) = "" Then 'フォルダがない場合はフォルダを作成
        MkDir FilePath & "\" & FolderName
    End If
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CopyFile FilePath & "\" & BookName, FilePath & "\" & FolderName & "\" & BookName
    
End Sub
Sub ワークブック保存時にフォルダに日付をつけてバックアップ()
'ワークブック保存時にフォルダに日付をつけてバックアップ
'20210721

    Dim FilePath$, FolderName$, BookName$, extension$, BookName2$, StrTime$
    FilePath = ThisWorkbook.Path
    FolderName = "Backup" 'バックアップするフォルダの名前←←←←←←←←←←←←←←←←←←←←←←←
    BookName = ThisWorkbook.Name
    
    If Dir(FilePath & "\" & FolderName, vbDirectory) = "" Then 'フォルダがない場合はフォルダを作成
        MkDir FilePath & "\" & FolderName
    End If
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    extension = FSO.GetExtensionName(BookName)
    BookName2 = Replace(BookName, "." & extension, "")
    StrTime = Format(Now(), "YYYYMMDDhhmmss") '←←←←←←←←←←←←←←←←←←←←←←←
    FSO.CopyFile FilePath & "\" & BookName, FilePath & "\" & FolderName & "\" & BookName2 & " " & StrTime & "." & extension
    
End Sub
Sub ワークブック保存時に同じフォルダ上に上書きバックアップ()
'ワークブック保存時に同じフォルダ上に上書きバックアップ
'20210721

    Dim FilePath$, AddStr$, BookName$, extension$, BookName2$, StrTime$
    FilePath = ThisWorkbook.Path
    AddStr = "backup" 'バックアップファイルの語尾につく名前←←←←←←←←←←←←←←←←←←←←←←←
    BookName = ThisWorkbook.Name
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    extension = FSO.GetExtensionName(BookName)
    BookName2 = Replace(BookName, "." & extension, "")
    FSO.CopyFile FilePath & "\" & BookName, FilePath & "\" & BookName2 & "_" & AddStr & "." & extension
    
End Sub

Sub ワークブック保存時に同じフォルダ上に日付をつけて上書きバックアップ()
'ワークブック保存時に同じフォルダ上に日付をつけて上書きバックアップ
'20210721

    Dim FilePath$, AddStr$, BookName$, extension$, BookName2$, StrTime$
    FilePath = ThisWorkbook.Path
    AddStr = "backup" 'バックアップファイルの語尾につく名前←←←←←←←←←←←←←←←←←←←←←←←
    BookName = ThisWorkbook.Name
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    extension = FSO.GetExtensionName(BookName)
    BookName2 = Replace(BookName, "." & extension, "")
    StrTime = Format(Now(), "YYYYMMDDhhmmss") '←←←←←←←←←←←←←←←←←←←←←←←
    FSO.CopyFile FilePath & "\" & BookName, FilePath & "\" & BookName2 & "_" & AddStr & "_" & StrTime & "." & extension
    
End Sub
