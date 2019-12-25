isAddVSCode=msgbox("右键菜单添加或移除VsCode，是（添加），否(移除)",vbYesNoCancel,"欢迎")

sub checkVscode(vspath)'检查VsCode是否存在子过程'
    if vspath="" then
        wscript.quit
    end if
    if fso.fileExists(vspath) then'检查Vscode是否存在'
        if right(vspath,len(vspath)-instrrev(vspath,"\"))<>"Code.exe" then'检查可执行文件名字
            checkVscode inputbox("可执行文件路径","文件不是VsCode的标准名字，请重新选择",vspath)
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
        checkVscode inputbox("可执行文件路径","找不到VsCode，需要手动指定路径",vspath)
    end if
end sub
Dim fso,wso
set wso=CreateObject("WScript.Shell")
set fso=CreateObject("Scripting.FileSystemObject") 
if isAddVSCode=vbYes then
    '搜索默认安装路径
    vspath=wso.ExpandenVironmentStrings("%LOCALAPPDATA%\Programs\Microsoft VS Code\Code.exe1")
    checkVscode(vspath)
elseif isAddVSCode=vbNo then
    regList=array("HKCR\*\shell\VSCode\command\","HKCR\*\shell\VSCode\","HKCR\Directory\shell\VSCode\command\","HKCR\Directory\shell\VSCode\","HKCR\Directory\Background\shell\VSCode\command\","HKCR\Directory\Background\shell\VSCode\") 
    for i=0 to UBound(regList)
        on error resume next
        wso.RegDelete regList(i)
    Next
    msgbox "移除完成"
end if