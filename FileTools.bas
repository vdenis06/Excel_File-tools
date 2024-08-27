Attribute VB_Name = "FileTools"
'@IgnoreModule HostSpecificExpression
' File Tools (c) vdenis.net 2024
' read folder content to compare to another folder or to a saved previous version
'
Option Explicit

Public twb As Workbook
Public isInit As Boolean
Public ws_params As Worksheet
Public ws_liste As Worksheet
Public ws_save_list As Worksheet
Public ws_Diff As Worksheet
Public ws_list1 As Worksheet
Public ws_list2 As Worksheet




Const C_FilePath As Long = 1
Const C_Path As Long = 2
Const C_File As Long = 3
Const C_Size As Long = 4
Const C_DateCreation As Long = 5
Const C_DateUpdate As Long = 6
Const C_DateAccess As Long = 7
Const C_Type As Long = 8
Const C_Attribute As Long = 9

Public C_My_Path As Long
Public L_My_Path As Long
Public C_LastUpdate As Long
Public L_LastUpdate As Long
Public L_SaveLastUpdate As Long
Public My_Path As String
Public MyFilePath1 As String
Public MyFilePath2 As String

Public LL As Long
Public FL1 As Long
Public LL1 As Long
Public FL2 As Long
Public LL2 As Long
Public FLD As Long
Public LLD As Long
Public MyLine1 As Long
Public MyLine2 As Long
Public MyLineDiff As Long


Public sLog As String
Public Barre As String
Public ligne As Long
Const Debug_lvl As Long = 4

' ############################################
Public Sub A_Main()
    Const proc_name As String = "A_Main"
    Barre = "|"
    sLog = vbNullString
    If Debug_lvl > 0 Then sLog = sLog = sLog & Date & Time & " " & Barre & " " & proc_name & " start" & " ==|" & vbCrLf
    A_Init
    ws_liste.Activate
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    
    ws_liste.Columns("A:G").EntireColumn.AutoFit

End Sub

Public Sub A_Init()
    isInit = False
    Const proc_name As String = "init"
    Barre = Barre & "="
    sLog = vbNullString
    If Debug_lvl > 2 Then sLog = sLog & Date & Time & " " & Barre & " " & proc_name & "Start ==|" & vbCrLf
    Set twb = ThisWorkbook
    Set ws_params = twb.Sheets("Params")
    Set ws_Diff = twb.Sheets("Diff")
    
End Sub

Public Sub Init_List(ByVal My_List As String)

    Select Case My_List
    Case "Path1 List"
        C_My_Path = 3
        L_My_Path = 2
        L_LastUpdate = L_My_Path
        L_SaveLastUpdate = 4
        Set ws_liste = twb.Sheets("Path1 List")
        Set ws_save_list = twb.Sheets("Saved1 List")
    Case "Path2 List"
        C_My_Path = 3
        L_My_Path = 3
        L_LastUpdate = L_My_Path
        L_SaveLastUpdate = 5
        Set ws_liste = twb.Sheets("Path2 List")
        Set ws_save_list = twb.Sheets("Saved2 List")
    End Select
    
    C_LastUpdate = 4
    My_Path = ws_params.Cells(L_My_Path, C_My_Path)

End Sub

Public Sub Save_Path1()
    Save_Path ("Path1 List")
End Sub

Public Sub compare_P1_P2()
    compare_list "Path2 List", "Path1 List"
End Sub

Public Sub compare_list(ByRef My_List1 As String, ByRef My_list2 As String)
    Dim My_Find As Variant
    If Not isInit Then A_Init
    'Init_List My_List1
    Set ws_list1 = twb.Sheets(My_List1)
    'Init_List My_list2
    Set ws_list2 = twb.Sheets(My_list2)
    Clear_List ws_Diff
    FL1 = 2
    FL2 = 2
    FLD = 2
    LLD = 2
    MyLineDiff = 2
    LL1 = ws_list1.Cells.Find(What:="*", after:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LL2 = ws_list2.Cells.Find(What:="*", after:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    ' pour chaque ligne dans list1
    For MyLine1 = FL1 To LL1
        ' recherche file_path1 dans liste 2
        MyFilePath1 = ws_list1.Cells(MyLine1, C_FilePath)
        If MyFilePath1 <> vbNullString Then
            Set My_Find = ws_list2.Columns(C_FilePath).Find(What:=MyFilePath1, LookAt:=xlWhole, MatchCase:=True)
            If My_Find Is Nothing Then
                ' si pas trouvé => effacé
                ws_Diff.Cells(MyLineDiff, C_Attribute) = "File Not Found in" & My_list2
            Else
                ' si différent => update (détail)
                If ws_list1.Cells(MyLine1, C_Size) <> ws_list2.Cells(My_Find.Row, C_Size) Then
                    ws_Diff.Cells(MyLineDiff, C_Attribute) = "Size change to " & ws_list2.Cells(My_Find.Row, C_Size) & " in " & My_list2
                End If
                If ws_list1.Cells(MyLine1, C_DateUpdate) <> ws_list2.Cells(My_Find.Row, C_DateUpdate) Then
                    ws_Diff.Cells(MyLineDiff, C_Attribute) = ws_Diff.Cells(MyLineDiff, C_Attribute) & " - Date change to " & ws_list2.Cells(My_Find.Row, C_DateUpdate) & " in " & My_list2
                End If
            End If
            If ws_Diff.Cells(MyLineDiff, C_Attribute) <> vbNullString Then
                ws_Diff.Cells(MyLineDiff, C_FilePath) = MyFilePath1
                ws_Diff.Cells(MyLineDiff, C_Size) = ws_list1.Cells(MyLine1, C_Size)
                ws_Diff.Cells(MyLineDiff, C_Size).NumberFormat = "#,##0"
                ws_Diff.Cells(MyLineDiff, C_DateUpdate) = ws_list1.Cells(MyLine1, C_DateUpdate)
                ws_Diff.Cells(MyLineDiff, C_DateUpdate).NumberFormat = "dd/mm/yyyy hh:mm;@"
                MyLineDiff = MyLineDiff + 1
            End If
        End If
    Next MyLine1

    ' pour chaque ligne dans list2
    For MyLine2 = FL2 To LL2
        ' recherche file_path2 dans liste 1
        MyFilePath2 = ws_list2.Cells(MyLine2, C_FilePath)
        If MyFilePath2 <> vbNullString Then
            Set My_Find = ws_list1.Columns(C_FilePath).Find(What:=MyFilePath2, LookAt:=xlWhole, MatchCase:=True)
            If My_Find Is Nothing Then
                ' si pas trouvé => effacé
                ws_Diff.Cells(MyLineDiff, C_Attribute) = "File Added in " & My_List1
                ws_Diff.Cells(MyLineDiff, C_FilePath) = MyFilePath2
                MyLineDiff = MyLineDiff + 1
            End If
            ' si pas trouvé => ajout
        End If
    Next MyLine2
    LL = ws_Diff.Cells.Find(What:="*", after:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    ws_Diff.ListObjects.Add(xlSrcRange, Range("$A$1:$I$" & LL), , xlYes).Name = "Diff"

End Sub

Public Sub Save_Path(ByVal My_List As String)
    If Not isInit Then A_Init
    Init_List My_List
    Clear_List ws_save_list
    ws_liste.Activate
    ws_liste.Cells.Select
    ws_liste.Cells.Copy Destination:=ws_save_list.Range("A1")
    ws_save_list.Activate
End Sub

Public Sub Clear_List(ByVal ws_My_List As Worksheet)
    ws_My_List.Activate
    ws_My_List.Cells.Select
    Selection.ClearContents
    Selection.ClearFormats
    ws_My_List.Cells(1, 1).Select
    ws_My_List.Cells(1, C_FilePath) = "File Path"
    ws_My_List.Cells(1, C_Path) = "Path"
    ws_My_List.Cells(1, C_File) = "File Name"
    ws_My_List.Cells(1, C_Size) = "Size"
    ws_My_List.Cells(1, C_DateCreation) = "Date Creation"
    ws_My_List.Cells(1, C_DateUpdate) = "Date Update"
    ws_My_List.Cells(1, C_DateAccess) = "Date Access"
    ws_My_List.Cells(1, C_Type) = "Type"
    ws_My_List.Cells(1, C_Attribute) = "Attribute"
End Sub

Public Sub Lookup_Path1()
    Read_Tree "Path1 List"
End Sub

Public Sub Read_Tree(ByVal My_List As String)
    Dim fs As Variant
    Dim Root_Folder As Object
    Dim My_Level As Long
    If Not isInit Then A_Init
    Init_List My_List
    If My_Path = vbNullString Then
        Get_Path (My_List)
    End If
    ws_liste.Activate
    
    Application.ScreenUpdating = False
    
    Clear_List ws_liste
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Root_Folder = fs.GetFolder(My_Path)
    ligne = 2
    My_Level = 1
    Read_Folder Root_Folder, My_Level
    
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    LL = ws_liste.Cells.Find(What:="*", after:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    ws_liste.ListObjects.Add(xlSrcRange, Range("$A$1:$I$" & LL), , xlYes).Name = My_List
    ' Range("Path_List[#All]").Select

End Sub

Public Sub Read_Folder(ByVal My_Folder As Object, ByVal My_Level As Long)
    Dim My_File As Object
    Dim My_Sub_Folder As Object
    Dim Folder_Size As LongLong
    Dim Folder_Line As Long
    Dim C1 As Long
    Dim C2 As Long
    Dim C3 As Long
    Dim C4 As Long
    Dim File_Extension As String
    Const proc_name As String = "Read_Folder"
    Barre = Barre & "="
    Folder_Size = 0
    Folder_Line = ligne
    
    ws_liste.Cells(Folder_Line, C_FilePath).Select
    ws_liste.Cells(Folder_Line, C_FilePath) = My_Folder.Path
    ws_liste.Cells(Folder_Line, C_FilePath).Interior.ColorIndex = 36
    ws_liste.Cells(Folder_Line, C_Path) = My_Folder.Path
    ws_liste.Cells(Folder_Line, C_Path).Interior.ColorIndex = 36
    
    Application.StatusBar = Barre & " " & proc_name & " -> " & My_Folder & " [" & My_Level & "] ligne " & Folder_Line
    ligne = ligne + 1
    For Each My_File In My_Folder.Files
        Folder_Size = Folder_Size + My_File.Size
        If Not (vbHidden) Then
            File_Extension = Right$(My_File.Path, Len(My_File.Path) - InStrRev(My_File.Path, "."))
            '            File_Extension = Right(My_File.Path, Len(My_File.Path) - InStrRev(My_File.Path, "."))
            ws_liste.Cells(ligne, C_FilePath) = My_File.Path
            ws_liste.Cells(ligne, C_Path) = My_Folder.Path
            ws_liste.Cells(ligne, C_File) = My_File.Name
            'ws_liste.Cells(ligne, C_File).Interior.ColorIndex = xlNone
            ws_liste.Cells(ligne, C_Size) = My_File.Size
            ws_liste.Cells(ligne, C_Size).NumberFormat = "#,##0"
            ws_liste.Cells(ligne, C_DateCreation) = My_File.DateCreated
            ws_liste.Cells(ligne, C_DateCreation).NumberFormat = "dd/mm/yyyy hh:mm;@"
            ws_liste.Cells(ligne, C_DateUpdate) = My_File.DateLastModified
            ws_liste.Cells(ligne, C_DateUpdate).NumberFormat = "dd/mm/yyyy hh:mm;@"
            ws_liste.Cells(ligne, C_DateAccess) = My_File.DateLastAccessed
            ws_liste.Cells(ligne, C_DateAccess).NumberFormat = "dd/mm/yyyy hh:mm;@"
            ws_liste.Cells(ligne, C_Type) = File_Extension
            If My_File.Attributes And vbHidden Then ws_liste.Cells(ligne, C_Attribute) = "Hidden"
            C1 = InStrRev(ws_liste.Cells(ligne, C_Path), "\")
            If C1 > 1 Then
                C2 = InStrRev(ws_liste.Cells(ligne, C_Path), "\", C1 - 1)
                If C2 > C1 Then
                    C3 = InStrRev(ws_liste.Cells(ligne, C_Path), "\", C2 - 1)
                    If C3 > C2 Then
                        C4 = InStrRev(ws_liste.Cells(ligne, C_Path), "\", C3 - 1)
                    End If
                End If
            End If
            ligne = ligne + 1
        End If
    Next
    For Each My_Sub_Folder In My_Folder.SubFolders
        Read_Folder My_Sub_Folder, My_Level + 1
    Next
    ws_liste.Cells(Folder_Line, C_Size) = Folder_Size
    ws_liste.Cells(Folder_Line, C_Size).NumberFormat = "#,##0"
    Application.StatusBar = False
    Barre = Left$(Barre, Len(Barre) - 1)
End Sub

Public Sub Get_Path(ByVal My_List As String)
    Dim chemin As String
    Dim Repertoire As FileDialog
    
    If Not isInit Then A_Init
    Init_List My_List

    ws_params.Activate

    Set Repertoire = Application.FileDialog(msoFileDialogFolderPicker)
    Repertoire.Show
    
    If Repertoire.SelectedItems.Count > 0 Then
        chemin = Repertoire.SelectedItems(1)
        ws_params.Cells(L_My_Path, C_My_Path) = chemin
    End If

End Sub

Public Function GetThisWorkbookLocalPath1() As String
    
    If Not ThisWorkbook.Path Like "http*" Then
        GetThisWorkbookLocalPath1 = ThisWorkbook.Path
        Exit Function
    End If
    
    Static myLocalPathCache As String, lastUpdated As Date
    If myLocalPathCache <> "" And Now() - lastUpdated <= 30 / 86400 Then
        GetThisWorkbookLocalPath1 = myLocalPathCache
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim recentFolderPath As String
    recentFolderPath = Environ("USERPROFILE") & "\AppData\Roaming\Microsoft\Windows\Recent\"
    
    Dim baseName As String, lnkFilePath As String
    baseName = fso.GetBaseName(ThisWorkbook.Name)
    Select Case True
        Case fso.FileExists(recentFolderPath & ThisWorkbook.Name & ".LNK")
            lnkFilePath = recentFolderPath & ThisWorkbook.Name & ".LNK"
        Case fso.FileExists(recentFolderPath & baseName & ".LNK")
            lnkFilePath = recentFolderPath & baseName & ".LNK"
        Case Else
            ' No LNK file exists.
            Exit Function
    End Select
    
    Dim filePath As String
    filePath = CreateObject("WScript.Shell").CreateShortcut(lnkFilePath).TargetPath
    
    If Not fso.FileExists(filePath) Then Exit Function
    myLocalPathCache = fso.GetParentFolderName(filePath)
    lastUpdated = Now()
    GetThisWorkbookLocalPath1 = myLocalPathCache
        
End Function
