Option Explicit

Dim a
a = 2

' if��
if a = 0 Then
    WScript.Echo "0�ł�"
elseif a = 2 Then
    WScript.Echo "2�ł�"
else
    WScript.Echo "0�ȊO�ł�"
end if

' for��
for a = 1 To 5
    WScript.Echo a
Next

' while��
Do while a < 100 
    a = a + 1
Loop

WScript.Echo a

' Function�Ăяo���i���l���Q�Ɠn������Ă���悤�ł��j
WScript.Echo addNum( a )
' ������+1����Ă�
WScript.Echo a

' Function(sub�͂����)
Function addNum( num )
    num = num + 1
    addNum = num
End Function



' �������� init�t�@�C���擾�B�ʓ|����������I�H
' �����̃}���R�s�� https://nkb.hatenablog.com/entry/20210123/1611371545
Function get_ini(sectionName,keyName,iniFile)
    Dim iniDic
    Set iniDic = read_ini(iniFile)
    get_ini = iniDic.Item(sectionName).Item(keyName)
End Function

Function read_ini(iniFile)

    Dim fso         'file system object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(iniFile) Then

        Dim dic         'Dictionary
        Set dic = CreateObject("Scripting.Dictionary")
        Dim sectionDic  'section Dictionary
        Set sectionDic = CreateObject("Scripting.Dictionary")
        Dim line
        Dim sectionName : sectionName = ""
        Dim a           'array

        Dim f           'file
        Set f = fso.OpenTextFile(iniFile)

        Do Until f.AtEndOfStream
            line = Trim(f.ReadLine)

            'Section
            If Left(line,1) = "[" And Right(line,1) = "]" Then
                sectionName = Mid(line, 2, Len(line) - 2)

                If Not dic.Exists(sectionName) Then
                  dic.Add sectionName, sectionDic
                End If

            'Parameter
            ElseIf Instr(line,"=") > 1 And sectionName <> "" Then

                'Key & Value
                a = Split(line,"=")
                dic(sectionName).Add Trim(a(0)), Trim(a(1))

            'comment'
            ElseIf Left(line,1) = ";" Then
            End If
        Loop

        Set read_ini = dic

        f.Close
        Set f = Nothing
        Set sectionDic = Nothing
        Set dic = Nothing
    End If

    Set fso = Nothing
End Function

WScript.Echo get_ini("test1","data2","setting.ini")
' �����܂� init�t�@�C���擾�B�ʓ|����������I�H

' �t�H���_����Excel�t�@�C����A1��\��
Function get_file_from_folder(folder)

    Dim fso         'file system object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folder) Then
        Dim objFolder
        Set objFolder = fso.GetFolder(folder)
        Dim objFiles
        Set objFiles = objFolder.Files
        Dim objFile, objExcel, excel, sheet ' for���Ȃ��Ŏg���ϐ���
        For Each objFile In objFiles
            WScript.Echo objFile.name
            Set objExcel = WScript.CreateObject("Excel.Application")
            Set excel = objExcel.WorkBooks.Open(objFile)
            Set sheet = excel.WorkSheets.Item(1)
            WScript.Echo sheet.Cells(1, 1)
            objExcel.Quit()
        Next
    else
        WScript.Echo folder+"�Ƃ����t�H���_�͂Ȃ��ł�"
    End If

    Set fso = Nothing
End Function
get_file_from_folder("../sampleBook")
