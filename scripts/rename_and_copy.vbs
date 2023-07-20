call rename_and_copy

sub rename_and_copy

    dim file_Name
    dim img_type
    dim img_attr
    dim img_number
    dim loop_result_boolean
    dim loop_result

    loop_result_boolean = true

    do While loop_result_boolean = true
    
        call create_destination_folder

        file_Name = get_random_img_filename()

        if file_Name <> "currentfile.jpg" then call to_current_filename(file_Name)

        call open_picture()

        img_type = get_img_type()
        if img_type = "null" then
            msgbox "Es wurde kein g" & chr(252) & "ltiger Bildtyp angegeben!"
            exit sub
        end if

        img_attr = get_img_attr(file_Name)
        if img_attr = "" then
            msgbox "Es wurde kein g" & chr(252) & "ltiges Bildattribut gesetzt!"
            exit sub
        end if

        img_number = get_img_number(img_type, img_attr)

        call rename_and_move(img_type & img_number & "_" & img_attr & ".jpg")

        loop_result = MsgBox("M" & chr(246) & "chten Sie weitermachen?", vbOKCancel)

        If loop_result = vbCancel Then
            loop_result_boolean = false
        End If

        call unload_photo_explorer()
        call unload_irfan()

    loop

    end sub

sub create_destination_folder()

    if IsFolderAvailable(".\benannt", "") = false then

        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder(".\benannt")

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

function IsFileAvailable(folderPath, msg)

    Set fsAvailable = CreateObject("Scripting.FileSystemObject")

    If fsAvailable.fileExists(folderPath) Then
        IsFileAvailable = true
    Else
        if msg <> "" then msgbox msg
        IsFileAvailable = false
    End If

end function

sub open_picture()

    Set objShell = CreateObject("WScript.Shell")
    objShell.Run ".\currentfile.jpg", 1, False

end sub

function get_random_img_filename()

    if IsFileAvailable("currentfile.jpg","") = true then

        get_random_img_filename = "currentfile.jpg"

    else

        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFolder = objFSO.GetFolder(".\")

        For Each objFile In objFolder.Files
            If LCase(Right(objFile.Name, 4)) = ".jpg" Then
                get_random_img_filename = objFile.Name
                exit for
            End If
        Next

    end if
end function

function get_img_attr(file_Name)

    dim input
    dim attr
    attr = ""

    input = inputbox("Anzahl L" & chr(228) & "ufe? Pfad:"  & chr(10) & file_Name,"Anzahl L" & chr(228) & "ufe")
    if isnumeric(input) = true then
    input = input & "lauf"
    end if
    if input ="" then
    elseif attr = "" then
    attr = input
    else
    attr = attr & "_" & input
    end if

    input = inputbox("Farbe? Pfad:"  & chr(10) & file_Name,"Farbe")
    if input ="" then
    elseif attr = "" then
    attr = input
    else
    attr = attr & "_" & input
    end if

    input = inputbox("Bitte geben Sie sonstige Artikelattribute mit Unterstrich separiert ein. Pfad:"  & chr(10) & file_Name,"Sonstige Artikelattribute")
    if input = "" then
    elseif attr = "" then
    attr = input
    else
    attr = attr & "_" & input
    end if

    get_img_attr = attr

end function

function get_img_type()

    img_type = inputbox("Bitte geben Sie den Bildtyp an:" & chr(10) & "1 = Freisteller" & chr(10) & "2 = Detailbild" & chr(10) & "3 = Bema" & chr(223) &  "ung" & chr(10) & "4 = Mileu" & chr(10) & "5 = Zubeh" & chr(246) & "r" & chr(10) & "6 = Montagebild" & chr(10) & "7 = Sonstiges","Bildtyp")


    if isnumeric(img_type) = false then
    get_img_type = "null"
    elseif cint(img_type) = 1 then
    get_img_type = "clipping_"
    elseif cint(img_type) = 2 then
    get_img_type = "detail_"
    elseif cint(img_type) = 3 then
    get_img_type = "dimensions_"
    elseif cint(img_type) = 4 then
    get_img_type = "mileu_"
    elseif cint(img_type) = 5 then
    get_img_type = "accessories_"
    elseif cint(img_type) = 6 then
    get_img_type = "allembly_"
    elseif cint(img_type) = 7 then
    get_img_type = "_"
    else
    get_img_type = "null"
    end if

end function

function get_img_number(img_type, img_attr)

    dim i 

    i = 1

    do While IsFileAvailable(".\benannt\" & img_type & i & "_" & img_attr & ".jpg", "") = True
        
        i = i + 1
    loop

    get_img_number = i

end function

sub rename_and_move(file_Name_New)

    dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFile ".\currentfile.jpg",  ".\benannt\" & file_Name_New

end sub


sub to_current_filename(file_Name)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFile ".\" & file_Name, ".\" & "currentfile.jpg"
end sub

sub unload_photo_explorer()

    Dim objWMIService, colProcesses, objProcess

    ' Windows-Fotoanzeige-Prozess beenden
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'PhotosApp.exe'")

    For Each objProcess In colProcesses
        objProcess.Terminate()
    Next

    Set objProcess = Nothing
    Set colProcesses = Nothing
    Set objWMIService = Nothing
end sub

sub unload_irfan()

    Dim objWMIService, colProcesses, objProcess

    ' Windows-Fotoanzeige-Prozess beenden
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'i_view32.exe'")

    For Each objProcess In colProcesses
        objProcess.Terminate()
    Next

    Set objProcess = Nothing
    Set colProcesses = Nothing
    Set objWMIService = Nothing
end sub

