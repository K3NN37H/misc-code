option explicit

' Quick Checksum
' (c) 2020 Ken

' This is a script for Directory Opus.
' See https://www.gpsoft.com.au/DScripts/redirect.asp?page=scripts for development information.



' Called by Directory Opus to initialize the script
Function OnInit(initData)
	initData.name = "Quick Checksum"
	initData.version = "1.0"
	initData.copyright = "(c) 2020 Ken"
	initData.url = "https://github.com/K3NN37H/misc-code/directory-opus-scripts"
	initData.desc = "Adds a few commands to calculate the checksum of the selected file using MD5, SHA1, SHA256, SHA512"
	initData.default_enable = true
	' FSUtil added SHA256 and SHA512 in 12.7
	' FSUtil added CRC32 in 12.13
	initData.min_version = "12.13"
	
	Dim sha1Cmd, sha256Cmd, sha512Cmd, md5Cmd, crc32Cmd
	Set sha1Cmd = initData.AddCommand()
	sha1Cmd.name = "HashFileSha1"
	sha1Cmd.method = "HashFileSha1"
	sha1Cmd.icon = "filecommands"
	sha1Cmd.desc = "Hash the selected file using SHA1"
	sha1Cmd.label = "SHA1"
	
	Set sha256Cmd = initData.AddCommand()
	sha256Cmd.name = "HashFileSha256"
	sha256Cmd.method = "HashFileSha256"
	sha256Cmd.icon = "filecommands"
	sha256Cmd.desc = "Hash the selected file using SHA256"
	sha256Cmd.label = "SHA256"
	
	Set sha512Cmd = initData.AddCommand()
	sha512Cmd.name = "HashFileSha512"
	sha512Cmd.method = "HashFileSha512"
	sha512Cmd.icon = "filecommands"
	sha512Cmd.desc = "Hash the selected file using SHA512"
	sha512Cmd.label = "SHA512"
	
	Set md5Cmd = initData.AddCommand()
	md5Cmd.name = "HashFileMd5"
	md5Cmd.method = "HashFileMd5"
	md5Cmd.icon = "filecommands"
	md5Cmd.desc = "Hash the selected file using MD5"
	md5Cmd.label = "MD5"
	
	Set crc32Cmd = initData.AddCommand()
	crc32Cmd.name = "HashFileCrc32"
	crc32Cmd.method = "HashFileCrc32"
	crc32Cmd.icon = "filecommands"
	crc32Cmd.desc = "Hash the selected file using CRC32"
	crc32Cmd.label = "CRC32"
End Function

Function HashFileSha1(ByRef cmdData)
	Call OnHashFile(cmdData, "sha1")
End Function

Function HashFileSha256(ByRef cmdData)
	Call OnHashFile(cmdData, "sha256")
End Function

Function HashFileSha512(ByRef cmdData)
	Call OnHashFile(cmdData, "sha512")
End Function

Function HashFileMd5(ByRef cmdData)
	Call OnHashFile(cmdData, "md5")
End Function

Function HashFileCrc32(ByRef cmdData)
	Call OnHashFile(cmdData, "crc32")
End Function

Dim checksumLength
Set checksumLength = CreateObject("Scripting.Dictionary")
checksumLength.Add "sha1", 40
checksumLength.Add "sha256", 64
checksumLength.Add "sha512", 128
checksumLength.Add "md5", 32
checksumLength.Add "crc32", 8

Function OnHashFile(ByRef clickData, hashType)
	DOpus.ClearOutput
	Dim cmd, selItem, clipHash
	' ---------------------------------------------------------
	Set cmd = clickData.func.command
	cmd.deselect = false ' Prevent automatic deselection
	cmd.RunCommand "Set UTILITY=otherlog"
	Dim selectedItems, hash, dialog
	Set selectedItems = clickData.func.sourcetab.selected
	If selectedItems.count = 1 Then
		For Each selItem In selectedItems
			If Not selItem.is_dir Then
				DOpus.Output "Hashing, please wait..."
				hash = DOpus.FSUtil.Hash(selItem.RealPath, hashType)
				DOpus.Output UCase(hashType) & ": " & hash
				clipHash = DOpus.GetClip("text")
				If checksumLength.Exists(hashType) AND Len(clipHash) = checksumLength.Item(hashType) Then
					If LCase(clipHash) = hash Then
						DOpus.Output "Clipboard hash matches!"
					Else
						DOpus.Output "Clipboard hash DOES NOT MATCH (shown below)"
						DOpus.Output UCase(hashType) & ": " & clipHash
					End If
				End If
			End If
		Next
	End If

End Function
