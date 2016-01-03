

'https://gallery.technet.microsoft.com/scriptcenter/fe809794-5e4c-498d-83a6-4832f3d9160b
'http://techshalvin.blogspot.jp/2014/04/zip-filesfolders-using-7-zip-and-vbs.html

Dim ArrayDateDiff()


'定数を指定する
Const C7Za = "C:\Program Files\7-Zip\7z.exe"
Const C7Zexe = "7z.exe"
Const TargetFol = "C:\work\test"

Set objFSO = CreateObject("Scripting.FileSystemObject") 		' Create FileSystemObject
Set objFolder = objFSO.GetFolder(TargetFol)						' Get Folder Object
Set objFiles = objFolder.Files									' Get Files Collection
Set objDictionary = CreateObject("Scripting.Dictionary")		' Get Dictionary Collection


'月判定取得
Call Month(objFiles,objDictionary,arrItems)


For Each strItem in arrItems

	'月判定(True:当月以外、False:当月)
	If Monthtest(strItem) = True Then
	
		'当月・ZIP以外を月毎にファイル取得
		Call MonthFile(objFolder,objFiles,strItem,CreateD,Files)
		
		For cnt = 0 to UBound(Files)-1
		
			'月名でフォルダを作成
			Call CreateFolderStatus(objFolder.Path,CreateD)
			
			'該当ファイルの移動
			Call MoveAfile(objFolder.Path & "\" & Files(cnt),objFolder.Path & "\" & CreateD,Files(cnt))
			
		Next
		
		'該当フォルダの圧縮処理
		Call CompresFLD(C7Za,objFolder.Path & "\" & CreateD & ".Zip",objFolder.Path & "\" & CreateD)
		
		'圧縮後のフォルダの削除処理
		Call DeleteAFolder(objFolder.Path & "\" & CreateD)
		
	End If
	
Next
	


'月毎判定値の重複排除
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

'月判定(True:当月以外、False:当月)
Function Monthtest(MonDiff)

	If MonDiff = 0 Then
	
		Monthtest = False
		
	Else
	
		Monthtest = True
		
	End If

End Function

'圧縮対象ファイルの更新日付を当月と比較
Function FCntDateDiff(DateLast)

	FCntDateDiff = DateDiff("m",FormatDateTime(date,vbGeneralDate),DateLast)

End Function

'当月以前のファイルまた、ZIP以外のファイルを月毎にグループ化
Sub MonthFile(ByVal objFolder,ByVal objFiles,ByVal key,ByRef CreateD,ByRef Files)

	'月毎のファイル数カウント
	Fcnt = 0
	
	For Each objFile in objFiles 
		'当月以外
		IF DateDiff("m",FormatDateTime(date,vbGeneralDate),objFile.DateLastModified) = key Then
			'ZIP以外
			If FileExtension(objFile.Name) <> "zip" Then
			
				Fcnt = Fcnt + 1
				
			End If
			
		End if
		
	Next

	ReDim Files(Fcnt)
	
	'月毎のファイル名を格納
	Fcnt = 0
	
	CreateD = mid(DateAdd("m",key,date),1,4) & mid(DateAdd("m",key,date),6,2)
	
	For Each objFile in objFiles 
	
		IF DateDiff("m",FormatDateTime(date,vbGeneralDate),objFile.DateLastModified) = key Then
			'ZIP以外
			If FileExtension(objFile.Name) <> "zip" Then

				Files(Fcnt) = objFile.Name
				Fcnt = Fcnt + 1
				
			End if
			
		End if
		
	Next 
	
End Sub

'同フォルダの存在を判別し作成する。
Sub CreateFolderStatus(ByVal Path,ByVal CreateD)

   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")

	fldr = Path & "\" & CreateD

	If ReportFolderStatus(fldr) = False Then
	
		fso.CreateFolder fldr
		
	End If
	
End Sub

'同ファイルの存在を判別し移動する
Sub MoveAfile(ByVal Source,ByVal Destination,ByVal File)

   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   
	If ReportFileStatus(Destination & "\" & File) = False Then
	
		fso.MoveFile Source,Destination & "\" & File
		
	End If
	
End Sub

'フォルダが存在するか判別する
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

'ファイルが存在するか判別する
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

'ファイルの拡張子を取得する
Function FileExtension(filespec)

	Dim fso, msg
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	Extention = fso.GetExtensionName(filespec)
	
	FileExtension = LCase(Extention)

End Function

'該当フォルダの圧縮処理
Sub CompresFLD(ByVal InstPath,Byval Zipfilename,Byval CompSource)

	Dim objShell
	
	'ex "C:\Program Files\7-Zip\7z.exe" a -tzip -mx=9 "C:\work\test\201502.Zip" "C:\work\test\201502"
	Set objShell = CreateObject("WScript.Shell")
	
	If ReportFolderStatus(CompSource) = True Then
	
		Command = Chr(34) & InstPath & Chr(34) & " a -tzip -mx=9 " & Chr(34) & Zipfilename & Chr(34) & Space(1) & Chr(34) & CompSource & Chr(34)
		objShell.Run Command,0,True
		
	End If
	
End Sub

'フォルダを削除する
Sub DeleteAFolder(folderspec)

	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	If ReportFolderStatus(folderspec) = True Then
	
		fso.DeleteFolder(folderspec)
		
	End If

End Sub

'処理を待機する
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