

'https://gallery.technet.microsoft.com/scriptcenter/fe809794-5e4c-498d-83a6-4832f3d9160b
'http://techshalvin.blogspot.jp/2014/04/zip-filesfolders-using-7-zip-and-vbs.html

Dim ArrayDateDiff()


'�萔���w�肷��
Const C7Za = "C:\Program Files\7-Zip\7z.exe"
Const C7Zexe = "7z.exe"
Const TargetFol = "C:\work\test"

Set objFSO = CreateObject("Scripting.FileSystemObject") 		' Create FileSystemObject
Set objFolder = objFSO.GetFolder(TargetFol)						' Get Folder Object
Set objFiles = objFolder.Files									' Get Files Collection
Set objDictionary = CreateObject("Scripting.Dictionary")		' Get Dictionary Collection


'������擾
Call Month(objFiles,objDictionary,arrItems)


For Each strItem in arrItems

	'������(True:�����ȊO�AFalse:����)
	If Monthtest(strItem) = True Then
	
		'�����EZIP�ȊO�������Ƀt�@�C���擾
		Call MonthFile(objFolder,objFiles,strItem,CreateD,Files)
		
		For cnt = 0 to UBound(Files)-1
		
			'�����Ńt�H���_���쐬
			Call CreateFolderStatus(objFolder.Path,CreateD)
			
			'�Y���t�@�C���̈ړ�
			Call MoveAfile(objFolder.Path & "\" & Files(cnt),objFolder.Path & "\" & CreateD,Files(cnt))
			
		Next
		
		'�Y���t�H���_�̈��k����
		Call CompresFLD(C7Za,objFolder.Path & "\" & CreateD & ".Zip",objFolder.Path & "\" & CreateD)
		
		'���k��̃t�H���_�̍폜����
		Call DeleteAFolder(objFolder.Path & "\" & CreateD)
		
	End If
	
Next
	


'��������l�̏d���r��
Function Month(ByVal objFiles,ByVal objDictionary,ByRef arrItems)

	ReDim ArrayDateDiff(objFiles.count-1)

	cnt = 0
	
	For Each objFile in objFiles
	
		ArrayDateDiff(cnt) = FCntDateDiff(objFile.DateLastModified)
		cnt = cnt + 1
		
	Next

	For cnt = 0 to UBound(ArrayDateDiff)
	
		strItem = ArrayDateDiff(cnt)
		If Not objDictionary.Exists(strItem) Then
			objDictionary.Add strItem,strItem
		End If
		
	Next

	intItems = objDictionary.Count - 1

	ReDim arrItems(intItems)

	i = 0

	For Each strKey in objDictionary.Keys
	
		arrItems(i) = strKey
		i = i + 1
		
	Next
	
End Function

'������(True:�����ȊO�AFalse:����)
Function Monthtest(MonDiff)

	If MonDiff = 0 Then
	
		Monthtest = False
		
	Else
	
		Monthtest = True
		
	End If

End Function

'���k�Ώۃt�@�C���̍X�V���t�𓖌��Ɣ�r
Function FCntDateDiff(DateLast)

	FCntDateDiff = DateDiff("m",FormatDateTime(date,vbGeneralDate),DateLast)

End Function

'�����ȑO�̃t�@�C���܂��AZIP�ȊO�̃t�@�C���������ɃO���[�v��
Sub MonthFile(ByVal objFolder,ByVal objFiles,ByVal key,ByRef CreateD,ByRef Files)

	'�����̃t�@�C�����J�E���g
	Fcnt = 0
	
	For Each objFile in objFiles 
		'�����ȊO
		IF DateDiff("m",FormatDateTime(date,vbGeneralDate),objFile.DateLastModified) = key Then
			'ZIP�ȊO
			If FileExtension(objFile.Name) <> "zip" Then
			
				Fcnt = Fcnt + 1
				
			End If
			
		End if
		
	Next

	ReDim Files(Fcnt)
	
	'�����̃t�@�C�������i�[
	Fcnt = 0
	
	CreateD = mid(DateAdd("m",key,date),1,4) & mid(DateAdd("m",key,date),6,2)
	
	For Each objFile in objFiles 
	
		IF DateDiff("m",FormatDateTime(date,vbGeneralDate),objFile.DateLastModified) = key Then
			'ZIP�ȊO
			If FileExtension(objFile.Name) <> "zip" Then

				Files(Fcnt) = objFile.Name
				Fcnt = Fcnt + 1
				
			End if
			
		End if
		
	Next 
	
End Sub

'���t�H���_�̑��݂𔻕ʂ��쐬����B
Sub CreateFolderStatus(ByVal Path,ByVal CreateD)

   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")

	fldr = Path & "\" & CreateD

	If ReportFolderStatus(fldr) = False Then
	
		fso.CreateFolder fldr
		
	End If
	
End Sub

'���t�@�C���̑��݂𔻕ʂ��ړ�����
Sub MoveAfile(ByVal Source,ByVal Destination,ByVal File)

   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   
	If ReportFileStatus(Destination & "\" & File) = False Then
	
		fso.MoveFile Source,Destination & "\" & File
		
	End If
	
End Sub

'�t�H���_�����݂��邩���ʂ���
Function ReportFolderStatus(fldr)

   Dim fso, msg
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   If (fso.FolderExists(fldr)) Then
   
      msg = True
      
   Else
   
      msg = False
      
   End If
   
   ReportFolderStatus = msg
   
End Function

'�t�@�C�������݂��邩���ʂ���
Function ReportFileStatus(filespec)
   Dim fso, msg
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   If (fso.FileExists(filespec)) Then
   
      msg = True
      
   Else
   
      msg = False
      
   End If
   
   ReportFileStatus = msg
   
End Function

'�t�@�C���̊g���q���擾����
Function FileExtension(filespec)

	Dim fso, msg
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	Extention = fso.GetExtensionName(filespec)
	
	FileExtension = LCase(Extention)

End Function

'�Y���t�H���_�̈��k����
Sub CompresFLD(ByVal InstPath,Byval Zipfilename,Byval CompSource)

	Dim objShell
	
	'ex "C:\Program Files\7-Zip\7z.exe" a -tzip -mx=9 "C:\work\test\201502.Zip" "C:\work\test\201502"
	Set objShell = CreateObject("WScript.Shell")
	
	If ReportFolderStatus(CompSource) = True Then
	
		Command = Chr(34) & InstPath & Chr(34) & " a -tzip -mx=9 " & Chr(34) & Zipfilename & Chr(34) & Space(1) & Chr(34) & CompSource & Chr(34)
		objShell.Run Command,0,True
		
	End If
	
End Sub

'�t�H���_���폜����
Sub DeleteAFolder(folderspec)

	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	If ReportFolderStatus(folderspec) = True Then
	
		fso.DeleteFolder(folderspec)
		
	End If

End Sub

'������ҋ@����
Function WaitEvent(EventID)
	
	WaitEvent = False
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
	errResult = objWMIService.Create("7z.exe", null, null, EventID) 
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
	Set colProcesses = objWMIService.ExecNotificationQuery("Select * From __InstanceDeletionEvent Within 1 Where TargetInstance ISA 'Win32_Process'") 

	Do Until i = 999 
	
	    Set objProcess = colProcesses.NextEvent 
	    
	    If objProcess.TargetInstance.ProcessID = EventID Then 
	    
	        Exit Do 
	        
	    End If 
	    
	Loop
	
	WaitEvent = True
	
End Function