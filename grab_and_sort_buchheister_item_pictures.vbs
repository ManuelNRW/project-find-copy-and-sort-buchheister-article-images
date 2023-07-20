
call main

sub Main()
    msgbox "Die Suche startet im Hintergrund sobald Sie auf 'OK' geklickt haben! Warte, bis die Meldung 'Fertig' erscheint."

    dim destination_path, source_pathes, input_data, user_name

    'get user name
    user_name = GetUserName()

    'load and check destination path
    destination_path = LoadDestinationPath()
    if IsFolderAvailable(destination_path, "Der Zielpfad ist nicht vorhanden!") = false then exit sub

    'load and check source pathes
    source_pathes = LoadSourcePathes()
    for i = 0 to ubound(source_pathes)
        if IsFolderAvailable(source_pathes(i), "Einer der Quellordner ist nicht vorhanden!") = false then 
        msgbox source_pathes(i)
        exit sub
        end if
    next

    'load input data / user phrases
    input_data = LoadUserPhrases

    'check destination folder
    call create_destination_folder(destination_path & "\" & input_data(0))

    call SearchAlgorithm(destination_path, source_pathes, input_data, user_name)

    call copy_rename_file(destination_path & "\" & input_data(0))

    msgbox "Fertig!"
end sub

sub SearchAlgorithm(destination_path, source_pathes, input_data, user_name)

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
                    
                    if is_every_split_in_array(input_data(J), file_name) = true then
                        fso_source.copyFile objFile.Path, destination_path & "\" & input_data(0) & "\"
                        call LogMatch(user_name, destination_path, objFile.Path, input_data(0))
                        exit for
                    end if

                next

            end if

        next

    next

end sub

function is_every_split_in_array(array_to_split, file_name)

    'dim split_array()
    dim split_match

    split_array = split(array_to_split, "*")

    split_match = true
    for k = 0 to ubound(split_array)
        if instr(1, file_name, split_array(k), 1) = 0 then
            split_match = false
            exit for
        end if
    next

    is_every_split_in_array = split_match

end function

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

    filePath = ".\database\destination_path.txt"

    Dim fs, f

    Set fs = CreateObject("Scripting.FileSystemObject")

    Set f = fs.OpenTextFile(filePath)

    LoadDestinationPath = f.ReadLine()

    f.Close

end function


function LoadSourcePathes()

    filePath = ".\database\source_path.txt"

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

sub LogMatch(user_name, destination_path, file_name, destination_name)

    ' Erzeuge ein FileSystemObject
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Öffne die Textdatei im Schreibmodus
    Dim logfile
    Set logfile = fso.OpenTextFile(destination_path & "\" & "log_matches.txt" , 8, True)

    ' Füge eine neue Zeile hinzu
    logfile.WriteLine now() & "|" & user_name  & "|" & right(file_name, len(file_name)-2) & "|" & destination_name

    ' Schließe die Textdatei
    logfile.Close

end sub

function GetUserName()
    Set objNetwork = CreateObject("WScript.Network")
    GetUserName = objNetwork.UserName
end function


sub copy_rename_file(destination_path)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile ".\scripts\rename_and_copy.vbs", destination_path & "\" & "rename_and_copy.vbs"
end sub