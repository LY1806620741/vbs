isAddVSCode=msgbox("�Ҽ��˵���ӻ��Ƴ�VsCode���ǣ���ӣ�����(�Ƴ�)",vbYesNoCancel,"��ӭ")

sub checkVscode(vspath)'���VsCode�Ƿ�����ӹ���'
    if vspath="" then
        wscript.quit
    end if
    if fso.fileExists(vspath) then'���Vscode�Ƿ����'
        if right(vspath,len(vspath)-instrrev(vspath,"\"))<>"Code.exe" then'����ִ���ļ�����
            checkVscode inputbox("��ִ���ļ�·��","�ļ�����VsCode�ı�׼���֣�������ѡ��",vspath)
            wscript.quit
        end if
        vsroot=left(vspath,instrrev(vspath,"\"))
        '�ļ��Ҽ�'
        wso.RegWrite "HKEY_CLASSES_ROOT\*\shell\VSCode\","Open with Code","REG_SZ"
        wso.RegWrite "HKCR\*\shell\VSCode\Icon",vspath,"REG_SZ"
        wso.RegWrite "HKCR\*\shell\VSCode\command\",""""+vspath+""" ""%1""","REG_SZ"
        'Ŀ¼�Ҽ�'
        wso.RegWrite "HKCR\Directory\shell\VSCode\","Open with Code","REG_SZ"
        wso.RegWrite "HKCR\Directory\shell\VSCode\Icon",vspath,"REG_SZ"
        wso.RegWrite "HKCR\Directory\shell\VSCode\command\",""""+vspath+""" ""%v""","REG_SZ"
        '�����Ҽ�'
        wso.RegWrite "HKCR\Directory\Background\shell\VSCode\","Open with Code","REG_SZ"
        wso.RegWrite "HKCR\Directory\Background\shell\VSCode\Icon",vspath,"REG_SZ"
        wso.RegWrite "HKCR\Directory\Background\shell\VSCode\command\",""""+vspath+""" ""%v""","REG_SZ"
        msgbox "������"
    else
        checkVscode inputbox("��ִ���ļ�·��","�Ҳ���VsCode����Ҫ�ֶ�ָ��·��",vspath)
    end if
end sub
Dim fso,wso
set wso=CreateObject("WScript.Shell")
set fso=CreateObject("Scripting.FileSystemObject") 
if isAddVSCode=vbYes then
    '����Ĭ�ϰ�װ·��
    vspath=wso.ExpandenVironmentStrings("%LOCALAPPDATA%\Programs\Microsoft VS Code\Code.exe1")
    checkVscode(vspath)
elseif isAddVSCode=vbNo then
    regList=array("HKCR\*\shell\VSCode\command\","HKCR\*\shell\VSCode\","HKCR\Directory\shell\VSCode\command\","HKCR\Directory\shell\VSCode\","HKCR\Directory\Background\shell\VSCode\command\","HKCR\Directory\Background\shell\VSCode\") 
    for i=0 to UBound(regList)
        on error resume next
        wso.RegDelete regList(i)
    Next
    msgbox "�Ƴ����"
end if