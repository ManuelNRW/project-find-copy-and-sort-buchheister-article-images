
call main


sub Main()

    'load pathes
    dim target_folder, source_pathes

    target_folder = LoadTargetFolder
    source_pathes = LoadSourcePathes

    'loop source pathes
    for i = 0 to ubound(source_pathes)

        msgbox source_pathes(i)

    next

end sub


function LoadTargetFolder()

    filePath = ".\target_path.txt"

    Dim fs, f

    Set fs = CreateObject("Scripting.FileSystemObject")

    Set f = fs.OpenTextFile(filePath)

    LoadTargetFolder = f.ReadLine()

    f.Close

end function


function LoadSourcePathes()

    filePath = ".\source_path.txt"

    Dim fs, f, lines, i

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(filePath)

    ReDim lines(-1)

    Do Until f.AtEndOfStream
        ReDim Preserve lines(UBound(lines) + 1)
        lines(UBound(lines)) = f.ReadLine()
    Loop

    f.Close

    LoadSourcePathes = lines

end function

