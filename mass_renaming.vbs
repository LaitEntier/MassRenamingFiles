Set objFso = CreateObject("Scripting.FileSystemObject")

Set Folder = objFSO.GetFolder(".\")

	Dim sInputOld
	Dim sInputNew
	
	sInputOld = InputBox("Enter current file name : ")
	
	sInputNew = InputBox("Enter new file name : ")
	
For Each File In Folder.Files

    sNewFile = File.Name

    sNewFile = Replace(sNewFile,sInputOld,sInputNew)

    if (sNewFile<>File.Name) then

        File.Move(File.ParentFolder+"\"+sNewFile)

    end if

Next