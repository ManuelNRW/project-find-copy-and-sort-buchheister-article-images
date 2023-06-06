
call main


sub Main()

    dim destination_path, source_pathes, input_data

    'load and check destination path
    destination_path = LoadDestinationPath()
    if IsFolderAvailable(destination_path, "Der Zielpfad ist nicht vorhanden!") = false then exit sub

    'load and check source pathes
    source_pathes = LoadSourcePathes()
    for i = 0 to ubound(source_pathes)
        if IsFolderAvailable(source_pathes(i), "Einer der Quellordner ist nicht vorhanden!") = false then exit sub
    next

    'load input data / user phrases
    input_data = LoadUserPhrases

    'check destination folder
    call create_destination_folder(destination_path & "\" & input_data(0))

    call SearchAlgorithm(destination_path, source_pathes, input_data)

    msgbox "Fertig!"
end sub

sub SearchAlgorithm(destination_path, source_pathes, input_data)

    dim file_name

    'target folder to objekt
    Set fso_destination = CreateObject("Scripting.FileSystemObject")
    Set obj_destination_path = fso_destination.GetFolder(destination_path)

    'source folder to objekt
    Set fso_source = CreateObject("Scripting.FileSystemObject")
    dim obj_source_pathes

    'loop source pathes
    for i = 0 to ubound(source_pathes)

        Set obj_source_pathes = fso_destination.GetFolder(source_pathes(i))
        
        'loop files in source pathes
        For Each objFile In obj_source_pathes.Files
            file_name = cstr(objfile.name)
            'ignore files if not .jpg or .jpeg
            if right(file_name, 4) = ".jpg" or right(file_name, 5) = ".jpeg" then

                'loop mathing phrases
                for j = 1 to ubound(input_data)
                    
                    if instr(1, file_name, input_data(j), 1) > 0 then
                        fso_source.copyFile objFile.Path, destination_path & "\" & input_data(0) & "\"
                    end if

                next

            end if

        next

    next

end sub

function LoadUserPhrases()

    filePath = ".\matching_phrases.txt"

    Dim fs, f, lines, i

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(filePath)

    ReDim lines(-1)

    Do Until f.AtEndOfStream
        ReDim Preserve lines(UBound(lines) + 1)
        lines(UBound(lines)) = f.ReadLine()
    Loop

    f.Close

    LoadUserPhrases = lines

end function


function LoadDestinationPath()

    filePath = ".\destination_path.txt"

    Dim fs, f

    Set fs = CreateObject("Scripting.FileSystemObject")

    Set f = fs.OpenTextFile(filePath)

    LoadDestinationPath = f.ReadLine()

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

sub create_destination_folder(folderPath)

    if IsFolderAvailable(folderPath, "") = false then

        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder(folderPath)

    end if


end sub

function IsFolderAvailable(folderPath, msg)

    Set fsAvailable = CreateObject("Scripting.FileSystemObject")

    If fsAvailable.FolderExists(folderPath) Then
        IsFolderAvailable = true
    Else
        if msg <> "" then msgbox msg
        IsFolderAvailable = false
    End If

end function