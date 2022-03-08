Option Explicit

' setting.ini��ǂݍ���Ŋe��ϐ��Ɋi�[����
' �i�[�����e��ϐ�����ʂɕ\������i�m�F�p�A���������邩�Ȃ��B�B0�Ȃ�Ԃ������肷��Ƃ������ȁj
' �����̌���̉ғ����v�Z����(����̃Z�������邾�����A�����܂ł̑����Z�����邩�B����̃Z���̏ꍇ�A��񂶂ē��͂��ł��Ȃ�)
' �c��̉ғ���(�c�蕽�� - �j��)����z�莞�Ԃ��Z�o����(�c��̉ғ��� * 8h * kadouritsu)
' �ȉ���\������
'   - �����܂ł̉ғ����ԁFxxx h
'   - �c��̑z��ғ����ԁFyyy h(�Z��)
'   - �z�荡���̉ғ����ԁFzzz h(�c��ғ������ׂĒ莞�Ƃ���)


' �t�H���_���̓���̃t�@�C������Excel�t�@�C����A1��\��
Function get_file_from_folder(folder)

    Dim fso         'file system object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folder) Then
        Dim objFolder
        Set objFolder = fso.GetFolder(folder)
        Dim objFiles
        Set objFiles = objFolder.Files
        Dim objRegExp
        Set objRegExp = CreateObject("VBScript.RegExp")
        objRegExp.Pattern = "^"+get_month_prefix()+"*"
        Dim objFile, objExcel, excel, sheet ' for�����Ŏg���ϐ���
        For Each objFile In objFiles
            WScript.Echo objFile.name
            If objRegExp.Test(objFile.name) Then
                Set objExcel = WScript.CreateObject("Excel.Application")
                Set excel = objExcel.WorkBooks.Open(objFile)
                Set sheet = excel.WorkSheets.Item(1)
                WScript.Echo sheet.Cells(1, 1)
            Else
                WScript.Echo "�����t�@�C���F"+objFile.name
            End If
            objExcel.Quit()
        Next
    else
        WScript.Echo folder+"�Ƃ����t�H���_�͂Ȃ��ł�"
    End If

    Set fso = Nothing
End Function

' ������prefix��������擾
Function get_month_prefix()
    get_month_prefix = Replace(Left(Now(),7), "/", "")
End Function

' init�t�@�C���̓ǂݍ��݊֐�
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