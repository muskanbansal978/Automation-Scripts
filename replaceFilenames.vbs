Set objFso = CreateObject("Scripting.FileSystemObject")


Set Folder = objFSO.GetFolder("C:\Users\musbansa\Desktop\CEI\CEIApplicationAnalysis\PS_CEI2_APP_SUBS_CITY_2_1_DAY")

For Each File In Folder.Files

    sNewFile = File.Name

    sNewFile = Replace(sNewFile,"-5GSA","")

    if (sNewFile<>File.Name) then

        File.Move(File.ParentFolder+"\"+sNewFile)

    end if

Next