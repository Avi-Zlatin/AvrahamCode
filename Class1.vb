Imports System.IO
Imports System.Net
Imports Microsoft.Win32
Imports System.Management
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports bVirtualKeyLib.wkcheck
''Imports System.Windows.Forms



'' 12-3-18 enumerating folder content
''  openXLS2  -- SCRIPT FOR MASON
''Avi: solving issue with the ownership report
''Avi : mrd10Diff     7-11-17
''Avi: getRecentVersion 4-21-7
''Avi: new chk for update 12-30-16
'Avi: comm regchk 11-29-2016
'Avi: last change date 5-13-2016
'Avi: isSiteOK check http  6-8-17

'E x e c P l a n

Public Class Class1
    ''fingerprint

    Public lnrStr, PwdChk As String
    Public strPath As String

    Public folderName As String = "ExecPlan"  'Avi: 5-12-16
    Public isExecPlan, isExecPlanSt As Boolean

    'Avi: set Folder  path
    '' 12-6-17
    Public Function SetFolderXLS(fldr As String) As String

        strPath = Environment.CurrentDirectory
        Dim position As Short = fldr.LastIndexOf("\"c)

        If fldr.Trim = "" Or position = -1 Then
            strPath = Environment.CurrentDirectory
            Return strPath
        End If


        If position = -1 Then
            strPath = Environment.CurrentDirectory
        Else
            strPath = fldr.Substring(0, position)
        End If

        Try

        Catch ex As Exception
            Return strPath
        End Try

        '' MsgBox("strPath=" & strPath)

        Return strPath

        '  strPath = fldr
    End Function

    ''Avi: solving issue with the ownership report 7-5-18
    Public Function openXLS2(fldr As String, outputF As String) As Short
        Dim dataPathxls, strB, strA, dataPathxls1, dataPathxlsNew As String
        Dim xls As Excel.Application
        Dim workbook1, workbook2 As Excel.Workbook
        Dim ind As Short
        Dim Str As String
        Dim oSheet As Excel.Worksheet

        xls = New Excel.Application

        xls.Visible = False

        Dim sData As String()
        sData = fldr.Split(":")

        'Dim folderBrowse As FolderBrowserDialog
        'Dim result As DialogResult = folderBrowse.ShowDialog()

        'If result = DialogResult.OK Then
        '    MsgBox(folderBrowse.SelectedPath)
        'Else
        '    MsgBox(" no folder name")
        'End If


        ''  MsgBox("IN BVIRTUAL ")
        Try
            ''strPath = Environment.CurrentDirectory
            '   MsgBox("outputF=" & outputF)

            strA = sData(0) & ".xls"
            '   dataPathxls1 = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlanExpress\" & strA
            dataPathxls1 = strPath & "\" & strA
            workbook1 = xls.Workbooks.Open(dataPathxls1)
            workbook1.Sheets(1).Name = sData(0).Trim

            Dim excelWorkSheet As Excel.Worksheet

            For ind = 1 To sData.GetUpperBound(0) - 1
                strB = sData(ind) & ".xls"
                '    MsgBox("ind=" & ind & "  strB=" & strB)
                '  dataPathxls = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlanExpress\" & strB
                dataPathxls = strPath & "\" & strB
                workbook2 = xls.Workbooks.Open(dataPathxls)

                Str = workbook2.Name
                Str = workbook1.Sheets(1).Name
                workbook2.Worksheets.Copy(workbook1.Sheets(1))
                workbook1.Sheets(1).Name = sData(ind).Trim
                workbook2.Close()
            Next


            If (outputF.Trim = "") Then
                'dataPathxlsNew = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlanExpress\CombinedReport.xls"
                dataPathxlsNew = strPath & "\CombinedReport.xls"
                '  workbook1.SaveAs(dataPathxlsNew)
            Else
                '  dataPathxlsNew = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlanExpress\" & outputF & ".xls"
                dataPathxlsNew = strPath & "\" & outputF & ".xls"
            End If
            '   MsgBox("dataPathxlsNew11=" & dataPathxlsNew & "upperbound=" & sData.GetUpperBound(0))

            Dim i As Short = sData.GetUpperBound(0)
            ind = 0

            For Each sheet As Worksheet In workbook1.Worksheets
                '   MsgBox("i=" & i)
                Str = sheet.Name
                'excelWorkSheet = workbook1.Sheets(i)
                'excelWorkSheet.Name = sData(ind)
                ''  MsgBox("ind=" & ind & "   sData=" & sData(ind))
               

                'excelWorkSheet.Cells(4, 4) = "gg11g"
                ''excelWorkSheet.Cells(4, 4)
                'excelWorkSheet.Range("A3").Value = excelWorkSheet.Rows.Count
                ''  excelWorkSheet.Range("A3:A3").NumberFormat = "@"
                'excelWorkSheet.Range("A3:A3").Style.Font.Name = "Comic Sans MS"

                'oSheet = workbook1.ActiveSheet

                'Str = oSheet.Cells(1, 1).Value.ToString()

                'Str = oSheet.Name  ''/Get the name of worksheet.

                Dim ind1, bln As Short
                bln = 0
                ind1 = 1
                Str = sheet.Name
                '    If (sData(ind) = "Ownership" Or sData(ind) = "ownership") Then
                If (Str = "Ownership" Or Str = "ownership") Then

                    Do
                        If sheet.Cells(ind1, 1).Value <> vbNullString Then

                            Str = sheet.Cells(ind1, 1).Value.ToString()

                            If (Str.Trim() = "PERCENT OF TOTAL") Then

                                sheet.Cells(ind1, 3).NumberFormat = "#0.00%"
                                sheet.Cells(ind1, 4).NumberFormat = "#0.00%"
                                sheet.Cells(ind1, 5).NumberFormat = "#0.00%"
                                sheet.Cells(ind1, 6).NumberFormat = "#0.00%"
                                bln = 1
                            End If
                        End If
                        ind1 += 1
                    Loop Until ind1 > 300 Or bln = 1

                End If

                'If (Str = "Ownership" Or Str = "ownership") Then

                '    Do
                '        If oSheet.Cells(ind1, 1).Value <> vbNullString Then

                '            Str = oSheet.Cells(ind1, 1).Value.ToString()

                '            If (Str.Trim() = "PERCENT OF TOTAL") Then

                '                '      excelWorkSheet.Cells(i, 1).Style.Font.Name = "Comic Sans MS"
                '                '     excelWorkSheet.Cells(i, 3).Style = excelWorkSheet.Cells(i, 2).Style
                '                '    excelWorkSheet.Cells(i, 1).NumberFormat = "#0.00%"
                '                excelWorkSheet.Cells(ind1, 3).NumberFormat = "#%"
                '                excelWorkSheet.Cells(ind1, 4).NumberFormat = "#%"
                '                excelWorkSheet.Cells(ind1, 5).NumberFormat = "#%"
                '                excelWorkSheet.Cells(ind1, 6).NumberFormat = "#%"
                '                bln = 1
                '            End If
                '        End If
                '        ind1 += 1
                '    Loop Until ind1 > 300 Or bln = 1

                'End If

                ind += 1
                i -= 1

            Next

            'MsgBox("dataPathxlsNew12=" & dataPathxlsNew)
            workbook1.SaveAs(dataPathxlsNew)

            workbook1.Close()

            xls.Quit()
            Return 1


        Catch ex As Exception
            '   "Can't write file - code 1010" + ex.Message
            xls.Quit()
            MsgBox(ex.Message)
            Return -1
        End Try

    End Function



    'Public Function openXLS2(fldr As String) As Short
    '    Dim dataPathxls, strB, strA, dataPathxls1, dataPathxlsNew As String
    '    Dim xls As Excel.Application
    '    Dim workbook1, workbook2 As Excel.Workbook
    '    Dim ind As Short
    '    xls = New Excel.Application

    '    xls.Visible = True

    '    Dim sData As String()
    '    sData = fldr.Split(":")

    '    Try

    '        strA = sData(ind) & ".xls"
    '        dataPathxls1 = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlanExpress\" & strA

    '        For ind = 2 To sData.GetUpperBound(0)

    '            strB = sData(ind) & ".xls"
    '            dataPathxls = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlanExpress\" & strB
    '        Next



    '        strA = "book1.xls"
    '        dataPathxls1 = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlanExpress\" & strA

    '        strB = "book2.xls"
    '        dataPathxls = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlanExpress\" & strB
    '        dataPathxlsNew = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlanExpress\CombinedReport.xls"


    '        Dim i As Short
    '        i = 1
    '        workbook1 = xls.Workbooks.Open(dataPathxls1)

    '        'Dim excelWorkSheet As Excel.Worksheet
    '        'excelWorkSheet = workbook1.Sheets(1)
    '        'excelWorkSheet.Name = "incommmmmmmmm"

    '        workbook2 = xls.Workbooks.Open(dataPathxls)
    '        'excelWorkSheet = workbook2.Sheets(1)
    '        'excelWorkSheet.Name = "balance"

    '        workbook2.Worksheets.Copy(workbook1.Sheets(1))

    '        workbook1.SaveAs(dataPathxlsNew)
    '        workbook1.Close()
    '        workbook2.Close()

    '        Return 1


    '    Catch ex As Exception
    '        '   "Can't write file - code 1010" + ex.Message
    '        MsgBox(ex.Message)
    '        Return -1
    '    End Try

    'End Function

    Public Function generateHWprofile() As Boolean
        Dim moReturn As Management.ManagementObjectCollection
        Dim moSearch As Management.ManagementObjectSearcher
        Dim mo As Management.ManagementObject
        Dim FilePath, lnr, SerialNumber, DriveType As String

        lnr = ""
        Try
            '  FilePath = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlan\epxhw.hsd" 'Personal data foldder - password 
            FilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName & "\epxhw.hsd" 'Personal data foldder - password 

            Dim objWriter As New System.IO.StreamWriter(FilePath)

            moSearch = New Management.ManagementObjectSearcher("Select * from Win32_LogicalDisk")
            moReturn = moSearch.Get
            For Each mo In moReturn

                If Not IsNothing(mo("Volumeserialnumber")) Then
                    SerialNumber = mo("Volumeserialnumber").ToString
                End If

                If Not IsNothing(mo("DriveType")) Then
                    DriveType = mo("DriveType").ToString

                    If CDbl(DriveType) = 3 Then
                        lnr = lnr + SerialNumber
                    End If
                End If
            Next

            moSearch = New Management.ManagementObjectSearcher("Select * from Win32_Processor")
            moReturn = moSearch.Get
            For Each mo In moReturn
                Dim ProcessorId As String = mo("ProcessorId").ToString
                lnr = lnr + ProcessorId
            Next

            lnrStr = lnr
            lnr = strEncrypt(lnr)
            '  MsgBox("lnr=" & lnr)
            objWriter.WriteLine(lnr)
            objWriter.Close()
            '  MsgBox("true")
            Return True
        Catch ex As Exception
            '   "Can't write file - code 1010" + ex.Message
            MsgBox(ex.Message)
            Return False
        End Try

    End Function


    Private Function strEncrypt(ByVal pWd As String) As String
        Dim i As Integer
        Dim ResultStr, EncryptionKey As String

        ResultStr = ""
        EncryptionKey = "A"

        Dim KeyChar As Integer
        KeyChar = Asc(EncryptionKey)

        For i = 1 To Len(pWd)
            ResultStr &= _
               Chr(KeyChar Xor _
               Asc(Mid(pWd, i, 1)))
        Next
        Return ResultStr
    End Function

    Public Function chkValidityHW(ByVal lnr As String) As Boolean
        Dim moReturn As Management.ManagementObjectCollection
        Dim moSearch As Management.ManagementObjectSearcher
        Dim mo As Management.ManagementObject
        Dim FilePath, SerialNumber, FilelogPath As String
        Dim counter As Integer = 0
        FilePath = ""

        'Try
        ' FilelogPath = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlan\chHW.log" 'Personal data foldder - password 
        FilelogPath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName & "\chHW.log" 'Personal data foldder - password 


        Dim objtmpWriter As New System.IO.StreamWriter(FilelogPath)
        objtmpWriter.WriteLine("in chkValidityHW. LNR=" & lnr)
        'Catch

        'End Try

        '' will check against the online lnr
        'Try
        '    FilePath = (My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlan\epxhw.hsd")
        'Catch fileException As Exception
        '    '  MsgBox("epxhw.hsd not found in folder" & FilePath)
        '    objtmpWriter.WriteLine("in chkValidityHW")
        '    objtmpWriter.Close()
        '    Return False
        '    '    Throw fileException
        'End Try

        'If File.Exists(FilePath) = True Then
        '    Dim objReader As New System.IO.StreamReader(FilePath)
        '    lnr = objReader.ReadLine
        '    objReader.Close()
        '    lnr = strEncrypt(lnr)
        '    '   MsgBox("hw file=" & lnr)
        'Else
        '    objtmpWriter.WriteLine("no hd file ")
        '    objtmpWriter.Close()
        '    Return False
        'End If

        moSearch = New Management.ManagementObjectSearcher("Select * from Win32_LogicalDisk")
        moReturn = moSearch.Get
        For Each mo In moReturn
            If Not IsNothing(mo("Volumeserialnumber")) Then
                SerialNumber = mo("Volumeserialnumber").ToString
            End If

            If Not IsNothing(mo("DriveType")) Then
                If CDbl(mo("DriveType")) = 3 Then
                    '  MsgBox("mo serial=" & SerialNumber)
                    If lnr.IndexOf(SerialNumber) <> -1 Then
                        counter += 1
                    End If
                End If
            End If
        Next

        moSearch = New Management.ManagementObjectSearcher("Select * from Win32_Processor")
        moReturn = moSearch.Get
        For Each mo In moReturn
            Dim ProcessorId As String = mo("ProcessorId").ToString
            If lnr.IndexOf(ProcessorId) <> -1 Then
                counter += 1
            End If
        Next

        objtmpWriter.WriteLine("counter" & counter)
        objtmpWriter.Close()
        If counter > 0 Then Return True Else Return False
    End Function

    'check registry and password and also if ther is a match to psd file
    Public Function regchk(yourkey As String) As Integer
        Dim regKey As RegistryKey
        Dim ThisCopy, FilePath As String
        Dim i2, result As Integer
        Dim ver, ReverseA As String ' version of the program

        'create ExecPlan folder if not there
        If (Not System.IO.Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName)) Then
            System.IO.Directory.CreateDirectory(System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName)
        End If

        FilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName & "\regch.log" 'Personal data foldder - password 
        Dim objtmpWriter As New System.IO.StreamWriter(FilePath)
        objtmpWriter.WriteLine("in regchk")

        Try

            If (isExecPlanSt = True) Then
                regKey = Registry.LocalMachine.OpenSubKey("Software\EpxSt", True)
                ThisCopy = CStr(regKey.GetValue("ExecPlanSt"))
            Else
                regKey = Registry.LocalMachine.OpenSubKey("Software\EpxS", True)
                ThisCopy = CStr(regKey.GetValue("excplan"))
            End If


            ver = Str(regKey.GetValue("Version"))
            ReverseA = CStr(regKey.GetValue("Access"))
            PwdChk = CStr(regKey.GetValue("pds" & yourkey))

            '  MsgBox("in chk pswd=" & PwdChk)
            objtmpWriter.WriteLine("PwdChk:" & PwdChk)
            objtmpWriter.WriteLine("yourkey:" & yourkey)
            PwdChk = strEncrypt(PwdChk)
            ' MsgBox("in chk pswd decrypt=" & strEncrypt(PwdChk))

            'If (yourkey <> strEncrypt(PwdChk)) Then
            '    objtmpWriter.WriteLine("not same pswd " & yourkey & ", PwdChk:" & strEncrypt(PwdChk))
            '    objtmpWriter.Close()
            '    Return False    'the virtual key that you using is not the same as the resitered one
            'Else

            '  MsgBox("in regcheck= key=" & yourkey)
            result = chkValidityPwd(yourkey)
            If (ThisCopy = "isProduction" And result = 1) Then
                objtmpWriter.WriteLine("is production")
                objtmpWriter.Close()
                Return 1
            Else
                objtmpWriter.WriteLine("not production")
                objtmpWriter.Close()
                Return result
            End If
            '   End If

        Catch ex As Exception
            objtmpWriter.WriteLine("error-no entry")
            objtmpWriter.Close()
            ''  MsgBox("Error reading reg entry. no entry =" & ex.Message)
            Return 3
        End Try

    End Function


    'register online - key contains either the key alon or concateneted with the activation code 
    Public Function registrationProc(newkey As String) As Integer
        Dim Str, FilePath, pswCode As String
        Dim ws As Service1 = New Service1

        ' ws.unRegKeySt("aa", "aa")
        'create ExecPlan folder if not there
        If (Not System.IO.Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName)) Then
            System.IO.Directory.CreateDirectory(System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName)
        End If


        FilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName & "\eeonlnreg.log" 'Personal data foldder - password 
        Dim objtmpWriter As New System.IO.StreamWriter(FilePath)

        If (generateHWprofile() = False) Then
            objtmpWriter.WriteLine("couldn't create HW file")
            objtmpWriter.Close()
            ws.Abort()
            Return 2
        End If

        objtmpWriter.WriteLine("registrationProc newkey=" & newkey)


        If (isExecPlanSt = True) Then
            ' pswCode = newkey.Substring(5)
            '  Str = ws.pwdRegSt(newkey.Substring(0, 5), pswCode, lnrStr)
            Str = ws.pwdRegSt(newkey, lnrStr)
        Else
            pswCode = newkey.Substring(5)
            Str = ws.pwdReg(newkey.Substring(0, 5), pswCode, lnrStr)
        End If


        'Str = ws.pwdReg(newkey.Substring(0, 5), pswCode, lnrStr)
        ' MsgBox("str=" & Str)
        'end of testing


        objtmpWriter.WriteLine("in registrationProc. pswCode=" & pswCode)
        objtmpWriter.WriteLine("str=" & Str)

        If (Str <> "Activated-") Then
            ws.Abort()
            objtmpWriter.WriteLine("couldn't activate:" & Str)
            objtmpWriter.Close()

            If (Str = "No password match-0") Then
                Return 31
            End If

            If (Str = "AlreadyActive-") Then
                Return 32
            End If

            If (Str = "NotFound") Then
                Return 30

            End If

            Return 40
        Else
            If (registerKey(newkey) = False) Then
                If (isExecPlanSt = True) Then
                    ws.unRegisterPwdSt(newkey)
                Else
                    ws.unRegisterPwd(newkey)
                End If

                ws.Abort()
                objtmpWriter.WriteLine("couldn't local register")
                objtmpWriter.Close()
                Return 4
            End If

            ws.Abort()

            objtmpWriter.WriteLine("str=" & RandomString(10))

            objtmpWriter.WriteLine("all well")
            objtmpWriter.Close()
            Return 1
        End If

    End Function


    'unregister your key
    Public Function UNregistrationProc(newkey As String) As Integer
        Dim Str, FilePath As String

        Dim ws As Service1 = New Service1

        Str = ws.unRegKey(newkey.Substring(0, 5), newkey.Substring(5))
        '   MsgBox("newkey=" & newkey.Substring(0, 5) & "pswCode=" & newkey.Substring(5) & "STR=" & Str)

        FilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName & "\eeonlnreg.log" 'Personal data foldder - password 
        Dim objtmpWriter As New System.IO.StreamWriter(FilePath)
        objtmpWriter.WriteLine("in unregister")
        objtmpWriter.WriteLine("str=" & Str)
        ws.Abort()

        If (Str = "Unregistered-" Or Str = "AlreadyNotRegistered-") Then

            Try
                Dim regKey As RegistryKey

                If (isExecPlanSt = True) Then
                    regKey = Registry.LocalMachine.OpenSubKey("Software\epxSt", True)

                Else
                    regKey = Registry.CurrentUser.OpenSubKey("Software\epxS", True)
                End If

                regKey.DeleteValue("pds" & newkey.Substring(0, 5))
                regKey.Close()
                objtmpWriter.WriteLine("need to delete reg=" & newkey.Substring(0, 5))
                '   MsgBox("in unreg 1")
            Catch ex As Exception
                objtmpWriter.WriteLine("couldn't unreg from reg=" & newkey.Substring(0, 5))

            End Try
            ' FilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName & "\dwp" & newkey.Substring(0, 5) & ".psd" 'Personal data foldder - password 
            FilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\dwp" & newkey.Substring(0, 5) & ".psd" 'Personal data foldder - password 

            File.Delete(FilePath)

            objtmpWriter.WriteLine("deleted f")
            objtmpWriter.Close()
            Return 1
        Else
            '  MsgBox("in unreg 2")
            objtmpWriter.WriteLine("couldn't onlin unreg")
            objtmpWriter.Close()
            Return 10
        End If

    End Function

    'register a new key localy
    Public Function registerKey(newkey As String) As Boolean
        Dim regKey As RegistryKey
        Dim i2 As Integer
        'Dim ver As Decimal
        Dim ThisCopy, chkResult, strE, strEncrypted, FilePath As String

        'create ExecPlan folder if not there
        If (Not System.IO.Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName)) Then
            System.IO.Directory.CreateDirectory(System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName)
        End If

        FilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName & "\reg5.log" 'Personal data foldder - password 
        Dim objtmpWriter As New System.IO.StreamWriter(FilePath)
        objtmpWriter.WriteLine("in registerKey - no write reg ver.")

        '  MsgBox("in register key proc")
        'Try


        'Catch ex As Exception
        '    objtmpWriter.WriteLine("can't register key in registry:" + ex.Message)
        '    objtmpWriter.Close()
        '    Return False
        'End Try

        Try
            If (isExecPlanSt = True) Then


                'regKey = Registry.LocalMachine.OpenSubKey("Software\epxSt", True)
                'ThisCopy = CStr(regKey.GetValue("ExecPlanSt"))
                'regKey.SetValue("ExecPlanSt", "isProduction")
                'objtmpWriter.WriteLine("ExecPlanSt val")

                'regKey.SetValue("Version", "1.0")
                ''   regKey.SetValue("Access", ReverseA)
                strE = strEncrypt(newkey)
                ' regKey.SetValue("pds" & newkey.Substring(0, 5), strE)
                objtmpWriter.WriteLine("St  entries created. newkey: " & newkey.Substring(0, 5))

            Else
                ''  AVI: Remove REG operations   12-19-16

                'regKey = Registry.LocalMachine.OpenSubKey("Software\epxS", True)
                'ThisCopy = CStr(regKey.GetValue("execplan"))
                'regKey.SetValue("execplan", "isProduction")
                'objtmpWriter.WriteLine("ExecPlan val")
                'regKey.SetValue("Version", "1.0")
                ''   regKey.SetValue("Access", ReverseA)
                'strE = strEncrypt(newkey)
                'regKey.SetValue("pds" & newkey.Substring(0, 5), strE)
                'objtmpWriter.WriteLine("reg entries created " & newkey)

            End If


            Try
                strEncrypted = strEncrypt(newkey) 'create the file
                '   FilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName & "\dwp" & newkey.Substring(0, 5) & ".psd" 'Personal data foldder - password 
                FilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\dwp" & newkey.Substring(0, 5) & ".psd" 'Personal data foldder - password 

                Dim objWriter As New System.IO.StreamWriter(FilePath)
                objWriter.WriteLine(strEncrypted)
                objWriter.Close()

                objtmpWriter.WriteLine("psd created")
            Catch
                MsgBox("can't create the pswd file")
                objtmpWriter.Close()
                Return False
            End Try
        Catch ex As Exception
            '   MsgBox("can't create the reg")
            MsgBox("err:" & ex.Message)
            objtmpWriter.Close()
            Return False
        End Try

        objtmpWriter.Close()
        Return True
    End Function


    'chek if the same key as in the psd file
    Public Function chkValidityPwd(ByVal myPwd As String) As Integer
        Dim validityResult As Integer = 1
        Dim pwd, strDate, FilePathlog, FilePath, Str As String

        'create ExecPlan folder if not there
        If (Not System.IO.Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName)) Then
            System.IO.Directory.CreateDirectory(System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName)
        End If

        FilePathlog = (System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments)) & "\" & folderName & "\eePwd.log" 'Personal data foldder - password 
        Dim objtmpWriter As New System.IO.StreamWriter(FilePathlog)
        objtmpWriter.WriteLine("in chkValidityPwd=" & myPwd)

        Try

            '     FilePath = (System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName & "\dwp" & myPwd & ".psd")
            FilePath = (System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\dwp" & myPwd & ".psd")

        Catch fileException As Exception
            validityResult = 20          
            objtmpWriter.WriteLine("No such path, can't validate dwp")
            objtmpWriter.Close()
            Return 20
        End Try


        If File.Exists(FilePath) = True Then

            Try
                Dim objReader As New System.IO.StreamReader(FilePath)
                pwd = objReader.ReadLine
                objReader.Close()
                objtmpWriter.WriteLine("1pasw comp is valid pwd=" & pwd & " PwdChk=" & PwdChk)
                '' pwd value from the file , PwdChk from registry
                If String.Compare(pwd, strEncrypt(PwdChk)) = 0 Then
                    validityResult = 1
                    objtmpWriter.WriteLine("2pasw comp is valid pwd=" & pwd & " PwdChk=" & PwdChk)
                    '   MsgBox("2pasw comp is valid pwd=" & pwd & " mypwd=" & PwdChk)
                Else
                    ' MsgBox("No match comp isnt valid pwd=" & pwd & " mypwd=" & PwdChk)
                    objtmpWriter.WriteLine("3pasw comp No Match=" & pwd & " PwdChk=" & PwdChk)
                End If
            Catch
                objtmpWriter.WriteLine("can't open the psd file")
                MsgBox("can't open the psd file")
                objtmpWriter.Close()
                validityResult = 21

            End Try

        Else
            MsgBox("psd file does not exist")
            objtmpWriter.WriteLine("psd file does not exist " & FilePath)
            objtmpWriter.Close()
            validityResult = 22
        End If

        'Dim dtDate As Date
        'Dim dateDiff As TimeSpan
        'FilePath = (My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\TaxMode\ord.psd")
        'If File.Exists(FilePath) = True Then
        '    Dim objReader As New System.IO.StreamReader(FilePath)
        '    strDate = objReader.ReadLine
        '    strDate = strEncrypt(strDate)
        '    dtDate = Convert.ToDateTime(strDate)
        '    If dtDate.Date < Now.Date Then
        '        dateDiff = Now.Subtract(dtDate)
        '        If dateDiff.TotalDays > 365 Then           'expired
        '            validityResult = False
        '        Else
        '            validityResult = True
        '        End If
        '    End If
        '    objReader.Close()
        'Else
        '    validityResult = False                 'No date file
        'End If
        objtmpWriter.Close()
        '  MsgBox("validy is=" & validityResult)
        Return validityResult
    End Function

    'checks existence, expiration and machine id
    'This version only checks the key
    Public Function onlinekeychk(ByVal keyCode As String) As Short
        Dim Str, StrDate, FilePath, lnrStr As String
        Dim DateTime As DateTime

        FilePath = (System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments)) & "\" & folderName & "\eeonlnchk.log" 'Personal data foldder - password 
        Dim objtmpWriter As New System.IO.StreamWriter(FilePath)

        objtmpWriter.WriteLine("in onlinekeychk.  ")

        If (isConnected() = False) Then
            objtmpWriter.WriteLine("no internet connection ")
            objtmpWriter.Close()
            Return 2 'no internet connection
        End If


        Dim ws As Service1 = New Service1
        Str = ws.pwdChk(keyCode)
        lnrStr = ws.getlnr(keyCode)
        ws.Abort()

        objtmpWriter.WriteLine("str=" & Str)

        If (Str <> "Found") Then

            objtmpWriter.WriteLine("not found:" & keyCode)
            objtmpWriter.Close()
            Return 3
        Else
            '  MsgBox("key is: " & keyCode)
            StrDate = ws.getExpirationDate(keyCode).Trim()
            ' MsgBox("1date is: " & StrDate)
            Try
                DateTime = Date.ParseExact(StrDate, "d", Globalization.CultureInfo.InvariantCulture)

                If (Now > DateTime) Then

                    objtmpWriter.WriteLine("expired")
                    objtmpWriter.Close()
                    Return 4
                Else
                    If (chkValidityHW(lnrStr) = False) Then

                        objtmpWriter.WriteLine("failed the same machine match")
                        objtmpWriter.Close()
                        Return 5
                    End If

                    objtmpWriter.Close()
                    Return 1 'OK
                End If

            Catch
                '   MsgBox("date conversion failed: " & DateTime)
                objtmpWriter.WriteLine("date conversion failed:" & StrDate)
                objtmpWriter.WriteLine("error:" & Err.Description)
                objtmpWriter.Close()
                Return 6
            End Try

        End If

    End Function

    'Avi: chng 5-11-2016
    'ExecPlan proc
    'checks existence, expiration and machine id
    'This version is for the keyless version hence we check also the validity of the passcode
    'KeyCode is just the key and we add the passcode from the registry before we check the database online
    Public Function onlinekeychkCombo(ByVal keyCode As String) As Integer
        Dim Str, StrDate, FilePath, lnrStr As String
        Dim DateTime As DateTime
        Dim result, numD As Integer

        setProgType("ExecPlan")        

        '  MsgBox(" in onlinekeychkCombo. key=" & keyCode)

        FilePath = (System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments)) & "\" & folderName & "\eeonlnchk.log" 'Personal data foldder - password 
        Dim objtmpWriter As New System.IO.StreamWriter(FilePath)
        objtmpWriter.WriteLine("in onlinekeychk combo. key:" & keyCode)
        objtmpWriter.Flush()


        If (isConnected() = False And keyCode.Trim <> "") Then

            objtmpWriter.WriteLine("no internet connection or server down ")
            objtmpWriter.Flush()
            numD = chkdays()
            '  MsgBox("numD " & numD)
            If (numD >= 0) Then
                objtmpWriter.WriteLine("numday with no int ok: ")
                objtmpWriter.Close()
                Return 1 'OK
            End If

            objtmpWriter.Close()          
            Return 2 'no internet connection
        End If


        ''Avi: remove regchk due users not able to read/write reg  12-19-16
        'result = regchk(keyCode)
        'If (result <> 1) Then
        '    objtmpWriter.WriteLine("no reg. match=" & keyCode & ". result=" & result)
        '    objtmpWriter.Close()
        '    Return result
        'End If


        Dim ws As Service1 = New Service1
        '  Str = ws.pwdChk(PwdChk)   'when key was integrated with pasw
        '  lnrStr = ws.getlnr(PwdChk) ''when key was integrated with pasw

        '   MsgBox("keytocheck=" & PwdChk)
        'Str = ws.pwdChk(PwdChk.Substring(0, 5))
        'lnrStr = ws.getlnr(PwdChk.Substring(0, 5))

        Str = ws.pwdChk(keyCode)
        lnrStr = ws.getlnr(keyCode)

        '    MsgBox(" in onlinekeychkCombo. Str=" & Str)

        ws.Abort()

        objtmpWriter.WriteLine("str=" & Str & "  lnrstr=" & lnrStr)
        objtmpWriter.Flush()

        If (Str <> "Found") Then
            objtmpWriter.WriteLine("not found:" & keyCode)
            objtmpWriter.Close()
            Return 3
        Else

            '            StrDate = ws.getExpirationDate(PwdChk.Substring(0, 5)).Trim()
            StrDate = ws.getExpirationDate(keyCode.Trim())

            Try
                DateTime = Date.ParseExact(StrDate, "d", Globalization.CultureInfo.InvariantCulture)

                If (Now > DateTime) Then
                    objtmpWriter.WriteLine("expired")
                    objtmpWriter.Close()
                    Return 4
                Else
                    If (chkValidityHW(lnrStr) = False) Then
                        objtmpWriter.WriteLine("failed the same machine match..")
                        objtmpWriter.Close()
                        Return 5
                    End If
                    objtmpWriter.Close()

                    updDays()
                    Return 1 'OK
                End If

            Catch
                '   MsgBox("date conversion failed: " & DateTime)
                objtmpWriter.WriteLine("date conversion failed:" & StrDate)
                objtmpWriter.WriteLine("error:" & Err.Description)
                objtmpWriter.Close()
                Return 6
            End Try

        End If

    End Function

    ''init no internet days
    Public Function updDays() As Short
        Dim days2try As String
        Dim pathString As String


        '' MsgBox(System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments))
        pathString = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\ExecPlan\dint.chk"

        ''  MsgBox("IN updDays:pathString" & pathString)

        days2try = "execplan0"
        days2try = strEncrypt(days2try)

        Try

            Dim objtmp As New System.IO.StreamWriter(pathString)

            objtmp.WriteLine(days2try)
            objtmp.Close()
        Catch ex As Exception

            Return -1
        End Try

    End Function


  


    ''getting the recent app version
    Public Function getRecentVersion(ByVal prgType As String) As String
        Dim Str, FilePath, graphName As String
        Dim arrInp3(5) As String 'Double
        Dim ind, pointsArray(5) As Short


        'Dim chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart
        ''For ind = 0 To 100
        ''    arrInp3(ind) = ind * 100
        ''Next


        'arrInp3 = {"Cat", "Dog", "Bird", "Monkey"}
        'pointsArray = {2, 1, 7, 5}

        'graphName = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlanExpress\networth.jpeg"
        'MsgBox("in getRecentVersion 1")
        'MsgBox("grapname=" & graphName)

        'chart1.ChartAreas.Add("mychart")
        'MsgBox("in getRecentVersion 1A")

        'chart1.ChartAreas(0) = New System.Windows.Forms.DataVisualization.Charting.ChartArea()

        'MsgBox("in getRecentVersion 2")


        'chart1.Titles.Add("Animals")

        'MsgBox("in getRecentVersion 2A")
        'For ind = 0 To 3
        '    chart1.Series.Add(arrInp3(ind))
        '    chart1.Series(ind).Points.Add(pointsArray(ind))
        'Next


        'MsgBox("in getRecentVersion 2B")
        '' chart1.Series.Add("s1")

        ''  chart1.Series(0).Points.DataBindY(arrInp3)

        ''   chart1.Series(0).ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
        'MsgBox("in getRecentVersion 3")
        'chart1.Width = 500
        'chart1.Height = 500

        'MsgBox("in getRecentVersion 4")
        'Try
        '    chart1.SaveImage(graphName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Jpeg)
        'Catch ex As Exception
        '    MsgBox(ex.Message)

        'End Try


        'MsgBox("in getRecentVersion 5")
        '' = New System.Windows.Forms.DataVisualization.Charting.ChartArea()

        ''     System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
        ''     System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();





        Str = ""
        Try
            Str = ""
            'FilePath = (System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments)) & "\" & folderName & "\eeonlnver.log" 'Personal data foldder - password 
            'Dim objtmpWriter As New System.IO.StreamWriter(FilePath)
            'objtmpWriter.WriteLine("in getRecentVersion")

            If (isConnected() = False) Then
                'objtmpWriter.WriteLine("no internet connection ")
                'objtmpWriter.Close()
                Return "NOINTERNET"  'no internet connection
            End If

            Dim ws As Service1 = New Service1


            If (prgType = "ExecPlan") Then

                Str = ws.getVerEP()   ''Avi: new chk for update 12-30-16
                '    MsgBox("in getRecentVersion 4" & Str)
            End If

            If (prgType = "ExecPlanStandard") Then
                Str = ws.getVersionSt()
            End If

            ws.Abort()
            '  objtmpWriter.Close()
        Catch ex As Exception
            Str = ""
        End Try


        Return Str

    End Function

    ''getting the recent app version
    Public Function chkWordPath(ByVal prgType As String) As String
        Dim Str As String
        Dim ind As Short

        Str = ""
        Try
            Str = IsOfficeInstalled()
            If Str <> "" Then
                ind = Str.IndexOf("Offic 10")
            End If

        Catch ex As Exception
            Str = ""
        End Try


        Return Str

    End Function

    Public Function IsOfficeInstalled() As String
        Dim PStr As String
        Dim key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe")

        PStr = ""
        If key IsNot Nothing Then
            PStr = key.GetValue("Path")
            key.Close()
        End If
        Return key IsNot Nothing
    End Function




    ''getting the recent app version
    Public Function chkExpiry(ByVal code As String) As String
        Dim Str, FilePath As String

        Str = ""
        FilePath = (System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments)) & "\" & folderName & "\eeonlnver.log" 'Personal data foldder - password 
        Dim objtmpWriter As New System.IO.StreamWriter(FilePath)
        objtmpWriter.WriteLine("in chkExpiry.")

        If (isConnected() = False) Then
            objtmpWriter.WriteLine("no internet connection. ")
            objtmpWriter.Close()
            Return "NOINTERNET"  'no internet connection
        End If


        Dim ws As Service1 = New Service1


        Str = ws.chkExpirySt(code)

        ws.Abort()
        objtmpWriter.Close()

        Return Str

    End Function

    'check the internet availability 
    Private Function isConnected() As Boolean
        Dim bln As Boolean

        Try
            '   My.Computer.Network.Ping("www.microsoft.com")
            My.Computer.Network.Ping("www.execplanexpress.com")

        Catch ex As Exception
            Return False
            Exit Function
        End Try

        bln = isSiteok()
        If bln = False Then Return False

        '' MsgBox("in isConnected True")
        Return True
    End Function

    ''6-8-17
    Private Function isSiteok() As Boolean
        Try
            Dim request As WebRequest = WebRequest.Create("http://www.execplanexpress.com/")
            Dim response As WebResponse = request.GetResponse()

            response.Close()

        Catch ex As Exception

            Return False
            Exit Function
        End Try

        Return True
    End Function

    Public Function RandomString(ByVal length As Integer) As String
        Dim random As New Random()
        Dim charOutput As Char() = New Char(length - 1) {}
        For i As Integer = 0 To length - 1
            Dim selector As Integer = random.[Next](65, 101)
            If selector > 90 Then
                selector -= 43
            End If
            charOutput(i) = Convert.ToChar(selector)
        Next
        Return New String(charOutput)
    End Function


    ' ''set  days to use with no internet/server connection
    ' ''Avi 5-11-16
    'Public Function setNumDays() As Short
    '    Dim days2try As Short

    '    MsgBox("bvirDaysTrySettings1=" & My.Settings.bvirDaysTry)

    '    days2try = My.Settings.bvirDaysTry
    '    If (days2try > 7) Then   ''exceeded the number of days try with no internet
    '        Return -1
    '    Else
    '        days2try += 1
    '        My.Settings.bvirDaysTry = days2try
    '        MsgBox("bvirDaysTrySettings2=" & My.Settings.bvirDaysTry)
    '        Return days2try
    '    End If
    'End Function


    ''set  days to use with no internet/server connection
    ''Avi 5-11-16
    Public Function chkdays() As Short
        Dim logPath, pathString, strtrys, stDays As String
        Dim days2try As Short

        pathString = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\ExecPlan\dint.chk"

        If File.Exists(pathString) = True Then
            Dim objReader As New System.IO.StreamReader(pathString)
            stDays = objReader.ReadLine
            stDays = strEncrypt(stDays)
            ' MsgBox("chkdays: stDays1=" & stDays)
            stDays = stDays.Substring(8, 1)
            '  MsgBox("chkdays: stDays2=" & stDays)
            days2try = Convert.ToInt32(stDays)
            objReader.Close()
        Else
            Return -1
        End If


        If (days2try > 7) Then   ''exceeded the number of days try with no internet
            Return -1
        Else
            days2try += 1
            stDays = "execplan" & days2try.ToString
            stDays = strEncrypt(stDays)
            Try
                Dim objtmp As New System.IO.StreamWriter(pathString)
                objtmp.WriteLine(stDays)
                objtmp.Close()

                Return days2try
            Catch ex As Exception
                Return -1
            End Try
        End If

    End Function


    ''set  if regular ExecPlan Express or full blown ExecPlan
    Public Function setProgType(ByVal prgType As String) As Boolean

        If (prgType = "ExecPlan") Then
            folderName = "ExecPlan"
            isExecPlan = True
            isExecPlanSt = False
        End If

        If (prgType = "ExecPlanStandard") Then
            folderName = "ExecPlanExpress"
            isExecPlan = False
            isExecPlanSt = True
        End If

        Return (True)
    End Function




    ''''==============================================================================================='''''
    '''' Improving different aspects of EP
    '''' Avi 7-11-17
    ''''
    ''''==============================================================================================''''''

    ''' check cl age more then 10 years apart
    '''
    ''' check cl age more then 10 years apart
    Public Function mrd10Diff(AgeCl As Short, AgeSp As Short) As String
        Dim line, dataPath As String
        Dim RetV As String
        Dim i, limitUp As Short
        Dim strArg() As String

        RetV = ""
        If AgeCl < 20 Or AgeCl > 133 Then Return "ERR"
        If AgeSp < 20 Or AgeSp > 133 Then Return "ERR"

        '   limitUp=

        dataPath = AppDomain.CurrentDomain.BaseDirectory & "IRADistrTable2.csv" 'Personal data foldder - password 

        'MsgBox(dataPath)
        '  dataPath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName & "\IRADistrTable2.csv" 'Personal data foldder - password 


        Try
            Dim file As New System.IO.StreamReader(dataPath)


            For i = 1 To 116
                line = file.ReadLine()
                '  MsgBox("line=" & line)

                strArg = line.Split(",")
                If strArg(0) = AgeCl.ToString Then  'got to the correct line 
                    '   If strArg(AgeSp - 20 + 1) <> "" Then
                    RetV = line
                    'End If
                    '  MsgBox("RetV1=" & RetV)

                    Exit For

                End If

            Next

        Catch ex As Exception
            MsgBox("ERROR --mrd10Diff=" & ex.Message)
        End Try

        'MsgBox("RetV=" & RetV)
        Return RetV

    End Function

    'Public Function mrd10Diff(AgeCl As Short, AgeSp As Short) As Double
    '    Dim line, dataPath As String
    '    Dim RetV As Double
    '    Dim i, limitUp As Short
    '    Dim strArg() As String


    '    RetV = -1
    '    If AgeCl < 20 Or AgeCl > 115 Then Return -1
    '    If AgeSp < 20 Or AgeSp > 115 Then Return -1

    '    '   limitUp=


    '    dataPath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86) & "\" & folderName & "\epxhw.hsd" 'Personal data foldder - password 

    '    MsgBox(dataPath)
    '    dataPath = AppDomain.CurrentDomain.BaseDirectory & "\" & folderName & "\epxhw.hsd" 'Personal data foldder - password 

    '    MsgBox(dataPath)
    '    dataPath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\" & folderName & "\IRADistrTable2.csv" 'Personal data foldder - password 


    '    Try
    '        Dim file As New System.IO.StreamReader(dataPath)


    '        For i = 1 To 98
    '            line = file.ReadLine()

    '            strArg = line.Split(",")
    '            If strArg(0) = AgeCl.ToString Then  'got to the correct line 
    '                If strArg(AgeSp - 20 + 1) <> "" Then
    '                    RetV = strArg(AgeSp - 20 + 1)
    '                End If

    '                Exit For

    '            End If

    '        Next

    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try

    '    Return RetV

    'End Function


    ''7-5-2017
    ''Snap report 1 changes.
    Public Function snapshot1(ByVal prgType As String) As String
        'Dim Str, FilePath, graphName As String
        'Dim arrInp3(5) As String 'Double
        'Dim ind, pointsArray(5) As Short

        ''reading the data

        Dim sData As String()
        '    sData = strWebVer.Split(":")


        'Dim chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart
        ''For ind = 0 To 100
        ''    arrInp3(ind) = ind * 100
        ''Next


        'arrInp3 = {"Cat", "Dog", "Bird", "Monkey"}
        'pointsArray = {2, 1, 7, 5}

        'graphName = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlanExpress\networth.jpeg"
        'MsgBox("in getRecentVersion 1")
        'MsgBox("grapname=" & graphName)

        'chart1.ChartAreas.Add("mychart")
        'MsgBox("in getRecentVersion 1A")

        'chart1.ChartAreas(0) = New System.Windows.Forms.DataVisualization.Charting.ChartArea()

        'MsgBox("in getRecentVersion 2")


        'chart1.Titles.Add("Animals")

        'MsgBox("in getRecentVersion 2A")
        'For ind = 0 To 3
        '    chart1.Series.Add(arrInp3(ind))
        '    chart1.Series(ind).Points.Add(pointsArray(ind))
        'Next


        'MsgBox("in getRecentVersion 2B")
        '' chart1.Series.Add("s1")

        ''  chart1.Series(0).Points.DataBindY(arrInp3)

        ''   chart1.Series(0).ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
        'MsgBox("in getRecentVersion 3")
        'chart1.Width = 500
        'chart1.Height = 500

        'MsgBox("in getRecentVersion 4")
        'Try
        '    chart1.SaveImage(graphName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Jpeg)
        'Catch ex As Exception
        '    MsgBox(ex.Message)

        'End Try


        'MsgBox("in getRecentVersion 5")
        '' = New System.Windows.Forms.DataVisualization.Charting.ChartArea()

        ''     System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
        ''     System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
    End Function

    '' Net Worth
    'Private Sub drawNetWorth(Chart1 As Chart, title As String, chartType As SeriesChartType)
    '    seriesA = "Assets"
    '    seriesB = "Net Worth"
    '    seriesC = "Liabilities"

    '    Chart1.Series.Add("s1")
    '    Chart1.Series.Add("s2")
    '    Chart1.Series.Add("s3")

    '    Chart1.Series(0).Points.DataBindY(arrInp3)
    '    Chart1.Series(1).Points.DataBindY(arrInp2)
    '    Chart1.Series(2).Points.DataBindY(arrInp1)


    '    Chart1.Series(0).Name = seriesA
    '    Chart1.Series(0).Color = Color.Blue

    '    Chart1.Series(1).Name = seriesB
    '    Chart1.Series(1).Color = Color.Green

    '    Chart1.Series(2).Name = seriesC
    '    Chart1.Series(2).Color = Color.Red

    '    Chart1.Series(0).ChartType = chartType
    '    Chart1.Series(1).ChartType = chartType
    '    Chart1.Series(2).ChartType = chartType

    '    Call chartSettings(Chart1, title)


    '    graphName = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlanExpress\networth.jpeg"
    '    Try
    '        If (textReportMode = True) Then Chart1.SaveImage(graphName, ChartImageFormat.Jpeg)
    '    Catch
    '    End Try
    'End Sub


    '' 12-3-18 enumerating folder 
    Public Function folderProp()
        'Dim diTop As DirectoryInfo = New DirectoryInfo("C:\Program Files (x86)\ExecPlan5")
        Dim diTop As DirectoryInfo = New DirectoryInfo(System.IO.Directory.GetCurrentDirectory())
        Dim FilePath As String

        ' MsgBox("in folderProp ")
        Try
            '  FilePath = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ExecPlan\epxhw.hsd" 'Personal data foldder - password 
            FilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\ExecPlan" & "\folderInfo.txt" 'Personal data foldder - password 

            Dim objWriter As New System.IO.StreamWriter(FilePath)


            objWriter.WriteLine(System.IO.Directory.GetCurrentDirectory())
            objWriter.WriteLine("File Name,    Last Write Date,      Size")
            objWriter.WriteLine("==========================================")

            For Each fi In diTop.EnumerateFiles()

                Try

                    '     MsgBox("in =  " & fi.FullName)
                    '   Console.WriteLine("{0}" & vbTab & vbTab & "{1}", fi.FullName, fi.Length.ToString("N0"))
                    objWriter.WriteLine(fi.Name & "  " & fi.LastWriteTime & "  " & fi.Length.ToString("N0"))


                Catch UnAuthTop As UnauthorizedAccessException
                    objWriter.WriteLine(fi.FullName & "  " & UnAuthTop.Message)
                End Try
            Next


            'Catch DirNotFound As DirectoryNotFoundException
            '    Console.WriteLine("{0}", DirNotFound.Message)
            'Catch UnAuthDir As UnauthorizedAccessException
            '    Console.WriteLine("UnAuthDir: {0}", UnAuthDir.Message)
            'Catch LongPath As PathTooLongException
            '    Console.WriteLine("{0}", LongPath.Message)

            objWriter.Close()

        Catch ex As Exception
            MsgBox("ERROR --folderProp=" & ex.Message)
        End Try



    End Function


End Class


