Set objFso = CreateObject(“Scripting.FileSystemObject”)


Set Folder = objFSO.GetFolder(“ENTER\PATH\HERE”)

For Each File In Folder.Files

    sNewFile = File.Name

    sNewFile = Replace(sNewFile,”ORIGINAL”,”REPLACEMENT”)

    if (sNewFile<>File.Name) then

        File.Move(File.ParentFolder+”\”+sNewFile)

    end if

Next
