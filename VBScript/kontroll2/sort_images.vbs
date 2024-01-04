option explicit
dim fso : set fso = CreateObject("Scripting.FileSystemObject")
dim sourceFolder, destinationFolder
sourceFolder = WScript.Arguments(0) & "\"
destinationFolder = WScript.Arguments(1) & "\"
dim jpegExt, jpgExt
jpegExt = "jpeg"
jpgExt = "jpg"
dim fileCount, folderCount
fileCount = 0
folderCount = 0

if Not fso.FolderExists(sourceFolder) Or Not fso.FolderExists(destinationFolder) then
    WScript.Echo "Source or destination folder does not exist."
    WScript.Quit
end if

function PadDigits(number, totalDigits)
    PadDigits = Right(String(totalDigits, "0") & number, totalDigits)
end function

function ProcessFolder(folderPath)
    dim folder, subFolder, file
    set folder = fso.GetFolder(folderPath)

    for each file In folder.Files
        ProcessFile file
    next

    for each subFolder In folder.SubFolders
        ProcessFolder subFolder.Path
    next
end function

function ProcessFile(file)
    if LCase(fso.GetExtensionName(file.Path)) = jpegExt Or LCase(fso.GetExtensionName(file.Path)) = jpgExt then
        dim fileYear, formattedDate, yearPath, datePath
        fileYear = Year(file.DateLastModified)
        formattedDate = fileYear & "-" & PadDigits(Month(file.DateLastModified), 2) & "-" & PadDigits(Day(file.DateLastModified), 2)
        yearPath = destinationFolder & fileYear & "\"
        datePath = yearPath & formattedDate & "\"

        if Not fso.FolderExists(yearPath) then
            fso.CreateFolder yearPath
            folderCount = folderCount + 1
        end if

        if Not fso.FolderExists(datePath) then
            fso.CreateFolder datePath
            folderCount = folderCount + 1
        end if

        file.Copy datePath
        fileCount = fileCount + 1

    end if
end function

function OutputLog(destinationFolderPath)
    dim folder, subFolder, file
    dim fileList, fileCount
    set folder = fso.GetFolder(destinationFolderPath)

    for each subFolder In folder.SubFolders
        set fileList = CreateObject("System.Collections.ArrayList")
        fileCount = 0

        for each file In subFolder.Files
            fileList.Add(file.Name)
            fileCount = fileCount + 1
        next

        if fileCount > 0 then
            WScript.Echo "--------"
            if fileCount = 1 then
                WScript.Echo fileCount & " file"
            else
                WScript.Echo fileCount & " files"
            end if
            WScript.Echo Join(fileList.ToArray(), ", ")
            WScript.Echo "were moved to folder"
            WScript.Echo subFolder.Path
        end if

        OutputLog subFolder.Path
    next
end function

ProcessFolder sourceFolder
OutputLog(destinationFolder)

if fileCount = 1 and folderCount = 1 then
    WScript.Echo fileCount & " picture was sorted into " & folderCount & " folder."
elseif fileCount = 1 then
    WScript.Echo fileCount & " picture was sorted into " & folderCount & " folders."
elseif folderCount = 1 then
    WScript.Echo fileCount & " pictures were sorted into " & folderCount & " folder."
else 
    WScript.Echo fileCount & " pictures were sorted into " & folderCount & " folders."
end if
