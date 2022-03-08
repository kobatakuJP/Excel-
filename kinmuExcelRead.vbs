Option Explicit

' setting.iniを読み込んで各種変数に格納する
' 格納した各種変数を画面に表示する（確認用、しかし見るかなぁ。。0なら赤くしたりするといいかな）
' 今月の現状の稼働を計算する(特定のセルから取るだけか、今日までの足し算をするか。特定のセルの場合、先んじて入力ができない)
' 残りの稼働日(残り平日 - 祝日)から想定時間を算出する(残りの稼働日 * 8h * kadouritsu)
' 以下を表示する
'   - 今日までの稼働時間：xxx h
'   - 残りの想定稼働時間：yyy h(〇日)
'   - 想定今月の稼働時間：zzz h(残り稼働日すべて定時として)


' フォルダ内の特定のファイル名のExcelファイルのA1を表示
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
        Dim objFile, objExcel, excel, sheet ' for文内で使う変数ら
        For Each objFile In objFiles
            WScript.Echo objFile.name
            If objRegExp.Test(objFile.name) Then
                Set objExcel = WScript.CreateObject("Excel.Application")
                Set excel = objExcel.WorkBooks.Open(objFile)
                Set sheet = excel.WorkSheets.Item(1)
                WScript.Echo sheet.Cells(1, 1)
            Else
                WScript.Echo "無視ファイル："+objFile.name
            End If
            objExcel.Quit()
        Next
    else
        WScript.Echo folder+"というフォルダはないです"
    End If

    Set fso = Nothing
End Function

' 今月のprefix文字列を取得
Function get_month_prefix()
    get_month_prefix = Replace(Left(Now(),7), "/", "")
End Function

' initファイルの読み込み関数
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