isAddVSCode=msgbox("右键菜单添加或移除VsCode，是（添加），否(移除)",vbYesNoCancel,"欢迎")
Function SelectFile()
    '使用javascript选择文件并获取文件路径，写入用户变量MsgResp，传递到此脚本，并清理该临时变量
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tempFolder : Set tempFolder = fso.GetSpecialFolder(2)
    Dim tempName : tempName = fso.GetTempName()
    Dim tempFile : Set tempFile = tempFolder.CreateTextFile(tempName & ".hta")
    tempFile.Write _
    "<input type='file' id='f' />" & _
    "<script type='text/javascript'>" & _
    "var f = document.getElementById('f');" & _
    "f.click();" & _
    "var shell = new ActiveXObject('WScript.Shell');" & _
    "shell.RegWrite('HKEY_CURRENT_USER\\Volatile Environment\\MsgResp', f.value);" & _
    "window.close();" & _
    "</script>"
    tempFile.Close
    wso.Run tempFolder & "\" & tempName & ".hta", 0, True
    SelectFile = wso.RegRead("HKEY_CURRENT_USER\Volatile Environment\MsgResp")
    wso.RegDelete "HKEY_CURRENT_USER\Volatile Environment\MsgResp"
End Function

sub checkVscode(vspath)'检查VsCode是否存在子过程'
    if vspath="" then
        wscript.quit
    end if
    if fso.fileExists(vspath) then'检查Vscode是否存在'
        if right(vspath,len(vspath)-instrrev(vspath,"\"))<>"Code.exe" then'检查可执行文件名字
            msgbox "文件不是VsCode的标准名字，请重新选择"
            checkVscode SelectFile
            wscript.quit
        end if
        vsroot=left(vspath,instrrev(vspath,"\"))
        '文件右键'
        wso.RegWrite "HKEY_CLASSES_ROOT\*\shell\VSCode\","Open with Code","REG_SZ"
        wso.RegWrite "HKCR\*\shell\VSCode\Icon",vspath,"REG_SZ"
        wso.RegWrite "HKCR\*\shell\VSCode\command\",""""+vspath+""" ""%1""","REG_SZ"
        '目录右键'
        wso.RegWrite "HKCR\Directory\shell\VSCode\","Open with Code","REG_SZ"
        wso.RegWrite "HKCR\Directory\shell\VSCode\Icon",vspath,"REG_SZ"
        wso.RegWrite "HKCR\Directory\shell\VSCode\command\",""""+vspath+""" ""%v""","REG_SZ"
        '桌面右键'
        wso.RegWrite "HKCR\Directory\Background\shell\VSCode\","Open with Code","REG_SZ"
        wso.RegWrite "HKCR\Directory\Background\shell\VSCode\Icon",vspath,"REG_SZ"
        wso.RegWrite "HKCR\Directory\Background\shell\VSCode\command\",""""+vspath+""" ""%v""","REG_SZ"
        msgbox "添加完成"
    else
        msgbox "找不到VsCode，需要手动指定路径"
        checkVscode SelectFile
    end if
end sub
Dim fso,wso
set wso=CreateObject("WScript.Shell")
set fso=CreateObject("Scripting.FileSystemObject") 
if isAddVSCode=vbYes then
    '搜索默认安装路径
    vspath=wso.ExpandenVironmentStrings("%LOCALAPPDATA%\Programs\Microsoft VS Code\Code.exe")
    checkVscode(vspath)
elseif isAddVSCode=vbNo then
    regList=array("HKCR\*\shell\VSCode\command\","HKCR\*\shell\VSCode\","HKCR\Directory\shell\VSCode\command\","HKCR\Directory\shell\VSCode\","HKCR\Directory\Background\shell\VSCode\command\","HKCR\Directory\Background\shell\VSCode\") 
    for i=0 to UBound(regList)
        on error resume next
        wso.RegDelete regList(i)
    Next
    msgbox "移除完成"
end if