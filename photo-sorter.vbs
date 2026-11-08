' ============================================================
' Photo Sorter
' Sorts JPG photos into folders by the year they were taken and
' the camera / phone model, read from EXIF data:
'
'   <folder>\2021\Canon EOS 80D\IMG_0001.jpg
'   <folder>\2023\iPhone 13 Pro\IMG_4521.jpg
'   <folder>\_No EXIF\scan_01.jpg
'
' Usage : drag & drop a folder onto this script,
'         or double-click it and pick a folder.
' Needs : nothing. It uses the built-in Windows Image Acquisition (WIA).
' ============================================================
Option Explicit

' ---------- Settings --------------------------------------
Const BY_YEAR   = True      ' create a year folder level
Const BY_CAMERA = True      ' create a camera model folder level
Const NO_EXIF_FOLDER = "_No EXIF"
' ---------------------------------------------------------

Const APP = "Photo Sorter"
Dim fso, img, folderPath, file, paths, i, moved, skipped, noExif, errors
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count > 0 Then
    folderPath = WScript.Arguments(0)
Else
    folderPath = BrowseForFolder("Select the folder with your photos:")
End If
If folderPath = "" Then WScript.Quit 0
If Not fso.FolderExists(folderPath) Then
    MsgBox "Folder not found:" & vbCrLf & folderPath, vbCritical, APP
    WScript.Quit 1
End If
folderPath = fso.GetAbsolutePathName(folderPath)

If MsgBox("JPG photos in this folder will be moved into Year \ Camera subfolders:" & vbCrLf & vbCrLf & _
          folderPath & vbCrLf & vbCrLf & "Continue?", vbYesNo + vbQuestion, APP) <> vbYes Then WScript.Quit

Set img = CreateObject("WIA.ImageFile")

' Snapshot the file list first, since we move files as we go
paths = Array()
For Each file In fso.GetFolder(folderPath).Files
    Select Case LCase(fso.GetExtensionName(file.Name))
        Case "jpg", "jpeg"
            ReDim Preserve paths(UBound(paths) + 1)
            paths(UBound(paths)) = file.Path
    End Select
Next

moved = 0 : skipped = 0 : noExif = 0 : errors = ""
For i = 0 To UBound(paths)
    SortPhoto paths(i)
Next

Dim msg
msg = "Done!" & vbCrLf & vbCrLf & _
      "Photos sorted   : " & moved & vbCrLf & _
      "Without EXIF    : " & noExif & " (moved to " & NO_EXIF_FOLDER & ")" & vbCrLf & _
      "Skipped / errors: " & skipped
If errors <> "" Then msg = msg & vbCrLf & vbCrLf & Left(errors, 800)
MsgBox msg, vbInformation, APP

' ================= FUNCTIONS =================

Sub SortPhoto(path)
    Dim year, model, target

    On Error Resume Next
    img.LoadFile path
    If Err.Number <> 0 Then
        Err.Clear
        skipped = skipped + 1
        errors = errors & fso.GetFileName(path) & " (cannot read)" & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    year = ExifYear()
    model = ExifModel()

    If year = "" And model = "" Then
        target = fso.BuildPath(folderPath, NO_EXIF_FOLDER)
        noExif = noExif + 1
    Else
        If year = "" Then year = "Unknown year"
        If model = "" Then model = "Unknown camera"
        target = folderPath
        If BY_YEAR Then target = fso.BuildPath(target, year)
        If BY_CAMERA Then target = fso.BuildPath(target, model)
    End If

    If LCase(target) = LCase(folderPath) Then Exit Sub
    CreateFolderTree target
    If MoveUnique(path, target) Then
        If year <> "" Or model <> "" Then moved = moved + 1
    Else
        skipped = skipped + 1
    End If
End Sub

' EXIF date looks like "2021:05:16 15:34:41". Prefer DateTimeOriginal.
Function ExifYear()
    Dim v
    ExifYear = ""
    v = Prop("ExifDTOrig")
    If v = "" Then v = Prop("DateTime")
    If Len(v) >= 4 Then
        If IsNumeric(Left(v, 4)) Then
            If CInt(Left(v, 4)) > 1900 Then ExifYear = Left(v, 4)
        End If
    End If
End Function

' Camera model, e.g. "Canon EOS 80D". The maker name is added when it is
' missing from the model ("SM-G991B" -> "samsung SM-G991B").
Function ExifModel()
    Dim make, model
    make = CleanName(Prop("EquipMake"))
    model = CleanName(Prop("EquipModel"))
    If model <> "" And make <> "" Then
        If InStr(1, model, Split(make, " ")(0), vbTextCompare) = 0 Then model = make & " " & model
    ElseIf model = "" Then
        model = make
    End If
    ExifModel = model
End Function

Function Prop(name)
    Prop = ""
    On Error Resume Next
    If img.Properties.Exists(name) Then Prop = CStr(img.Properties(name).Value)
    Err.Clear
    On Error GoTo 0
End Function

' Strip characters that are not allowed in Windows folder names
Function CleanName(s)
    Dim bad, c
    s = Replace(s, Chr(0), "")
    For Each c In Array("<", ">", ":", """", "/", "\", "|", "?", "*")
        s = Replace(s, c, " ")
    Next
    Do While InStr(s, "  ") > 0 : s = Replace(s, "  ", " ") : Loop
    s = Trim(s)
    Do While Right(s, 1) = "." : s = Trim(Left(s, Len(s) - 1)) : Loop
    CleanName = s
End Function

Sub CreateFolderTree(path)
    Dim parent
    parent = fso.GetParentFolderName(path)
    If parent <> "" And Not fso.FolderExists(parent) Then CreateFolderTree parent
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

' Moves a file without overwriting: photo.jpg -> photo_1.jpg ...
Function MoveUnique(src, destFolder)
    Dim dest, n
    dest = fso.BuildPath(destFolder, fso.GetFileName(src))
    n = 1
    Do While fso.FileExists(dest)
        dest = fso.BuildPath(destFolder, fso.GetBaseName(src) & "_" & n & "." & fso.GetExtensionName(src))
        n = n + 1
    Loop
    On Error Resume Next
    fso.MoveFile src, dest
    MoveUnique = (Err.Number = 0)
    If Err.Number <> 0 Then errors = errors & fso.GetFileName(src) & " (" & Err.Description & ")" & vbCrLf
    Err.Clear
    On Error GoTo 0
End Function

Function BrowseForFolder(prompt)
    Dim f
    Set f = CreateObject("Shell.Application").BrowseForFolder(0, prompt, &H10 Or &H200, 0)
    If f Is Nothing Then BrowseForFolder = "" Else BrowseForFolder = f.Self.Path
End Function
