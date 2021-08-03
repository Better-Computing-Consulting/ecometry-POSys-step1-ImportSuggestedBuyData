Imports System.IO
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Text.RegularExpressions
Module Module1
   Public ITEMDBSTR As String = ""
   Public ITEMMASTDB As String = ""
    'Public ITEMDBSTR As String = "Data Source=ECOMDB3;Initial Catalog=RECBUYS;UID=sss;PWD=sssss"
    'Public ITEMMASTDB As String = "Data Source=ecom-db1;Initial Catalog=ECOMLIVE;UID=sss;PWD=sssss"
    Dim RunProduction As Boolean = False
   Sub Main()
      Dim body As String = "Job Started: " & Now.ToString & Environment.NewLine & Environment.NewLine
      Dim ReportFiles() As String = {"RECMDLTW", "IAA1W", "IAC3W"}
      Dim FilesProcessed As New List(Of String)
      Dim FilesWithError As New List(Of String)
      For Each StreamFile As String In ReportFiles
         Dim args As String = ""
         If RunProduction Then
            args = "E:\ECOMETRY\LIVECODE\UserStr2\" & StreamFile & " JOBS 31000 E:\ECOMETRY\LISTENER\ECOMLIVE.ini"
         Else
            args = "E:\ECOMETRY\VERCODE\UserStr2\" & StreamFile & " JOBS 31040 E:\ECOMETRY\LISTENER\ECOMVER.ini"
                ITEMDBSTR = "Data Source=ECOMDB3;Initial Catalog=RECBYSDEV;UID=sss;PWD=sssss"
                ITEMMASTDB = "Data Source=ecom-db2;Initial Catalog=ECOMVER;UID=sss;PWD=sssss"
            End If
         Dim p As New Process
         With p.StartInfo
            .RedirectStandardOutput = True
            .FileName = "E:\ECOMETRY\VERCODE\PRO\cronsched.exe"
            .CreateNoWindow = False
            .UseShellExecute = False
            .Arguments = args
         End With
         p.Start()
         p.WaitForExit()
         Dim ReportFile As String = ""
         Do While p.StandardOutput.Peek >= 0
            Dim s As String = p.StandardOutput.ReadLine
            If s.Contains("StdList") Then
               Dim path As String = s.Substring(s.IndexOf("\\"), s.LastIndexOf("\") - s.IndexOf("\\") + 1)
               Dim jobNo As String = s.Substring(s.LastIndexOf("."))
               Dim SpoolDir As New DirectoryInfo(path)
               For Each f As FileInfo In SpoolDir.GetFiles(StreamFile & "_*" & jobNo)
                  ReportFile = f.FullName
                  Select Case StreamFile
                     Case "IAA1W"
                        body &= "A1 Items" & Environment.NewLine
                     Case "IAC3W"
                        body &= "C3 Items" & Environment.NewLine
                     Case "RECMDLTW"
                        body &= "R1 Items" & Environment.NewLine
                  End Select
                  body &= f.FullName.Substring(2) & Environment.NewLine
               Next
            End If
         Loop
         If ReportFile.Length > 10 Then
            body &= ProcessOneFile(ReportFile)
            FilesProcessed.Add(ReportFile)
         Else
            FilesWithError.Add(StreamFile)
         End If
      Next
      Dim anEmail As New MailMessage
      With anEmail
            .CC.Add("federico@ecommerce.com")
            If RunProduction Then
                '.To.Add("tony@ecommerce.com")
                '.To.Add("paul@ecommerce.com")
            End If
            .From = New MailAddress(My.Computer.Name & "@ecommerce.com")
            Dim subject As String = "Import of R1, A1 and C3 items into Recommended Buy PO System executed"
         For Each s As String In FilesProcessed
            .Attachments.Add(New Attachment(s))
         Next
         If FilesWithError.Count > 0 Then
            .Priority = MailPriority.High
            For Each s As String In FilesWithError
               Select Case s
                  Case "IAA1W"
                     body &= "Error processing the A1 items stream file IAA1W" & Environment.NewLine
                  Case "IAC3W"
                     body &= "Error processing the C3 items stream file IAC3W" & Environment.NewLine
                  Case "RECMDLTW"
                     body &= "Error processing the R1 items stream file RECMDLTW" & Environment.NewLine
               End Select
            Next
            subject &= " with Error!"
         End If
         .Body = body & Environment.NewLine & "Job Ended: " & Now.ToString
         .Subject = subject
      End With
      Dim aSMTPClient As New SmtpClient("172.16.92.10")
      Try
         aSMTPClient.Send(anEmail)
      Catch ex As Exception
         Console.WriteLine(ex.Message)
         LogThis("EMAILERROR", ex.Message)
      End Try
      Console.WriteLine(vbCrLf & "done")
   End Sub
   Function ProcessOneFile(ByVal FilePath As String) As String
      Dim tmpResult As String = ""
      Dim CurrentVendor As String = ""
      Dim rgx As New Regex("/[0-9\s][0-9\s]/", RegexOptions.IgnoreCase)
      Using sr As New StreamReader(FilePath)
         Dim counter As ULong = 0
         Do While sr.Peek >= 0
            Dim tmpline As String = sr.ReadLine
            If tmpline.Contains("Selection: Vendor") Then
               CurrentVendor = tmpline.Substring(19, 10).Trim
            End If
            Dim matches As MatchCollection = rgx.Matches(tmpline)
            If matches.Count = 1 Then
               If New FileInfo(FilePath).Name.ToUpper.StartsWith("RECM") Then
                  If tmpline.Trim.Length = 131 Then
                     Dim itemdesc As String = sr.ReadLine
                     Dim vendorline As String = sr.ReadLine
                     Dim vendordesc As String = sr.ReadLine
                     If CurrentVendor <> "PRID" Then
                        Dim tmpItem As New bikeItem(tmpline, itemdesc, vendorline, vendordesc, CurrentVendor, True)
                        tmpResult &= tmpItem.oItemNo & Environment.NewLine
                        Console.WriteLine(tmpItem.oEDPNO & vbTab & tmpItem.oItemNo & vbTab)
                        counter += 1
                     End If
                  End If
               Else
                  If tmpline.Trim.Length = 132 Then
                     Dim itemdesc As String = sr.ReadLine
                     Dim vendorline As String = sr.ReadLine
                     Dim vendordesc As String = sr.ReadLine
                     If CurrentVendor <> "PRID" Then
                        Dim tmpItem As New bikeItem(tmpline, itemdesc, vendorline, vendordesc, CurrentVendor, False)
                        Console.WriteLine(tmpItem.oEDPNO & vbTab & tmpItem.oItemNo & vbTab)
                        tmpResult &= tmpItem.oItemNo & Environment.NewLine
                        counter += 1
                     End If
                  End If
               End If
            End If
         Loop
         tmpResult &= "Added " & counter & " items" & Environment.NewLine & Environment.NewLine
         Console.WriteLine(vbCrLf & counter)
      End Using
      Return tmpResult
   End Function
   Public Sub LogThis(ByVal ShortText As String, ByVal LongText As String)
      Dim entryDateTime As DateTime = Now
      Dim logfile As String = My.Application.Info.DirectoryPath & "\" & My.Application.Info.ProductName & ".log"
      My.Computer.FileSystem.WriteAllText(logfile, "[" & entryDateTime.ToString("yyyy-MM-dd HH:mm:ss") & "]" & " " & String.Format("{0,-15}{1}" & Environment.NewLine, ShortText, LongText), True)
   End Sub
End Module
Class bikeItem
#Region "Class Private Variables"
    'Private ITEMDBSTR As String = "Data Source=ECOMDB3;Initial Catalog=RECBYSDEV;UID=sss;PWD=sssss"
    Public oItemNo As String = ""
   Public oEDPNO As String = ""
   Public oItemDesc As String = ""
   Public oVendor As String = ""
   Public oVendorNo As String = ""
   Public oVendorDesc As String = ""
   Public oStatus As String = ""
   Public oPrice As String = ""
   Public oMarginPercent As String = ""
   Public oAvg52Wk As String = ""
   Public oAvg26Wk As String = ""
   Public oAvg13Wk As String = ""
   Public oAvg8Wk As String = ""
   Public oAvg4Wk As String = ""
   Public oLastWk As String = ""
   Public oPONum1 As String = ""
   Public oPOExpDate1 As String = ""
   Public oPOQty1 As String = ""
   Public oPONum2 As String = ""
   Public oPOExpDate2 As String = ""
   Public oPOQty2 As String = ""
   Public oPONum3 As String = ""
   Public oPOExpDate3 As String = ""
   Public oPOQty3 As String = ""
   Public oPONum4 As String = ""
   Public oPOExpDate4 As String = ""
   Public oPOQty4 As String = ""
   Public oTotalDue As String = ""
   Public oOnHand As String = ""
   Public oWeeksOfStock As String = ""
   Public oReorderLevel As String = ""
   Public oVendorPrice As String = ""
   Public oLandedCost As String = ""
   Public oMinQty As String = ""
   Public oBOQty As String = ""
   Public oRecmdBuy As String = ""
#End Region
   Sub New(ByVal ItemLine As String, ByVal ItemDescriptionLine As String, ByVal VendorLine As String, ByVal VendorDescriptionLine As String, ByVal CurrentVendor As String, ByVal IsR1 As Boolean)
      If IsR1 Then
         oItemDesc = ItemDescriptionLine.Substring(3, 50).Trim
         oAvg52Wk = ItemLine.Substring(15, 6).Trim
         oAvg26Wk = ItemLine.Substring(21, 6).Trim
         oAvg13Wk = ItemLine.Substring(28, 6).Trim
         oAvg8Wk = ItemLine.Substring(35, 6).Trim
         oAvg4Wk = ItemLine.Substring(42, 6).Trim
         oLastWk = ItemLine.Substring(49, 4).Trim
         oPONum1 = ItemLine.Substring(54, 9).Trim
         oPOExpDate1 = ItemLine.Substring(64, 8).Trim
         oPOQty1 = ItemLine.Substring(73, 5).Trim
         oPONum2 = ItemDescriptionLine.Substring(54, 9).Trim
         oPOExpDate2 = ItemDescriptionLine.Substring(64, 8).Trim
         oPOQty2 = ItemDescriptionLine.Substring(73, 5).Trim
         oPONum3 = VendorLine.Substring(54, 9).Trim
         oPOExpDate3 = VendorLine.Substring(64, 8).Trim
         oPOQty3 = VendorLine.Substring(73, 5).Trim
         oPONum4 = VendorDescriptionLine.Substring(54, 9).Trim
         oPOExpDate4 = VendorDescriptionLine.Substring(64, 8).Trim
         oPOQty4 = VendorDescriptionLine.Substring(73, 5).Trim
         oTotalDue = ItemLine.Substring(79, 5).Trim
         oOnHand = ItemLine.Substring(85, 6).Trim
         oWeeksOfStock = ItemLine.Substring(92, 6).Trim
         oReorderLevel = ItemLine.Substring(99, 6).Trim
         oVendorPrice = ItemLine.Substring(106, 7).Trim
         oMinQty = ItemLine.Substring(114, 4).Trim
         oBOQty = ItemLine.Substring(119, 5).Trim
         oRecmdBuy = ItemLine.Substring(127, 5).Trim
         oVendor = CurrentVendor
         oVendorNo = VendorLine.Substring(1, 23).Trim
         oStatus = VendorLine.Substring(25, 2)
         oPrice = VendorLine.Substring(30, 8).Trim
         oMarginPercent = VendorLine.Substring(40, 6).Trim
         oVendorDesc = VendorDescriptionLine.Substring(3, 49).Trim
         SetItemNoAndEDPNO(ItemLine.Substring(0, 14).Trim)
      Else
         oItemDesc = ItemDescriptionLine.Substring(3, 50).Trim
         oAvg52Wk = ItemLine.Substring(17, 6).Trim
         oAvg26Wk = ItemLine.Substring(23, 6).Trim
         oAvg13Wk = ItemLine.Substring(30, 6).Trim
         oAvg8Wk = ItemLine.Substring(37, 6).Trim
         oAvg4Wk = ItemLine.Substring(44, 6).Trim
         oLastWk = ItemLine.Substring(51, 4).Trim
         oPONum1 = ItemLine.Substring(56, 9).Trim
         oPOExpDate1 = ItemLine.Substring(66, 8).Trim
         oPOQty1 = ItemLine.Substring(75, 5).Trim
         oPONum2 = ItemDescriptionLine.Substring(56, 9).Trim
         oPOExpDate2 = ItemDescriptionLine.Substring(66, 8).Trim
         oPOQty2 = ItemDescriptionLine.Substring(75, 5).Trim
         oPONum3 = VendorLine.Substring(56, 9).Trim
         oPOExpDate3 = VendorLine.Substring(66, 8).Trim
         oPOQty3 = VendorLine.Substring(75, 5).Trim
         oPONum4 = VendorDescriptionLine.Substring(56, 9).Trim
         oPOExpDate4 = VendorDescriptionLine.Substring(66, 8).Trim
         oPOQty4 = VendorDescriptionLine.Substring(75, 5).Trim
         oTotalDue = ItemLine.Substring(82, 5).Trim
         oOnHand = ItemLine.Substring(87, 6).Trim
         oWeeksOfStock = ItemLine.Substring(94, 6).Trim
         oReorderLevel = ItemLine.Substring(101, 6).Trim
         oVendorPrice = ItemLine.Substring(109, 9).Trim
         oLandedCost = ItemLine.Substring(120, 9).Trim
         oMinQty = ItemLine.Substring(129, 4).Trim
         oBOQty = ItemDescriptionLine.Substring(129, 4).Trim
         oVendor = CurrentVendor
         oVendorNo = VendorLine.Substring(1, 23).Trim
         oStatus = VendorLine.Substring(25, 2)
         oPrice = VendorLine.Substring(30, 8).Trim
         oMarginPercent = VendorLine.Substring(40, 6).Trim
         oVendorDesc = VendorDescriptionLine.Substring(3, 49).Trim
         SetItemNoAndEDPNO(ItemLine.Substring(0, 16).Trim)
      End If
      EnterIntoDB()
   End Sub
   Private Sub SetItemNoAndEDPNO(ByVal tmpItemNumber As String)
      Dim queryString As String = "SELECT EDPNO,ITEMNO FROM ITEMMAST WHERE DESCRIPTION = '" & oItemDesc.Replace("'", "''") & "' AND ITEMNO LIKE '" & tmpItemNumber & "%'"
      Using conn As New SqlConnection(ITEMMASTDB)
         Dim cmd As New SqlCommand(queryString, conn)
         Try
            conn.Open()
            Dim r As SqlDataReader = cmd.ExecuteReader
            If r.HasRows Then
               r.Read()
               Dim tmp As String = r.Item(0)
               oEDPNO = tmp.Trim
               tmp = r.Item(1)
               oItemNo = tmp.Trim
            Else
               Console.WriteLine(tmpItemNumber & " not found")
            End If
         Catch ex As Exception
            Console.WriteLine(ex.Message)
         End Try
      End Using
   End Sub
   Private Sub EnterIntoDB()
      Dim queryString As String = "INSERT INTO RECMDBUYITEMS" & _
           "(RBI_NUMBER,RBI_EDPNO_NUMBER,RBI_DESCRIPTION,RBI_STATUS,RBI_PRICE,RBI_MARGIN,RBI_VENDOR,RBI_VENDOR_ITM_NUMBER,RBI_VENDOR_DESCRIPTION,RBI_VENDOR_PRICE,RBI_ORIGINAL_VENDOR_PRICE,RBI_LANDED_COST,RBI_AVG52WK" & _
           ",RBI_AVG26WK,RBI_AVG13WK,RBI_AVG8WK,RBI_AVG4WK,RBI_LASTWK,RBI_PO_NUMBER1,RBI_PO_EXPDATE1,RBI_PO_QTY1,RBI_PO_NUMBER2,RBI_PO_EXPDATE2,RBI_PO_QTY2,RBI_PO_NUMBER3" & _
           ",RBI_PO_EXPDATE3,RBI_PO_QTY3,RBI_PO_NUMBER4,RBI_PO_EXPDATE4,RBI_PO_QTY4,RBI_TOTAL_DUE,RBI_ONHAND,RBI_WEEKSOFSTOCK,RBI_REORDERLEVEL,RBI_MINQTY,RBI_BOQTY" & _
           ",RBI_RECMDBUY,RBI_ADDED_DATE) VALUES(" & _
            "'" & oItemNo & "'," & _
            "" & oEDPNO & "," & _
            "'" & oItemDesc.Replace("'", "''") & "'," & _
            "'" & oStatus & "'," & _
            "'" & oPrice & "'," & _
            "'" & oMarginPercent & "'," & _
            "'" & oVendor & "'," & _
            "'" & oVendorNo & "'," & _
            "'" & oVendorDesc.Replace("'", "''") & "'," & _
            "'" & oVendorPrice & "'," & _
            "'" & oVendorPrice & "'," & _
            "'" & oLandedCost & "'," & _
            "'" & oAvg52Wk & "'," & _
            "'" & oAvg26Wk & "'," & _
            "'" & oAvg13Wk & "'," & _
            "'" & oAvg8Wk & "'," & _
            "'" & oAvg4Wk & "'," & _
            "'" & oLastWk & "'," & _
            "'" & oPONum1 & "'," & _
            "'" & oPOExpDate1 & "'," & _
            "'" & oPOQty1 & "'," & _
            "'" & oPONum2 & "'," & _
            "'" & oPOExpDate2 & "'," & _
            "'" & oPOQty2 & "'," & _
            "'" & oPONum3 & "'," & _
            "'" & oPOExpDate3 & "'," & _
            "'" & oPOQty3 & "'," & _
            "'" & oPONum4 & "'," & _
            "'" & oPOExpDate4 & "'," & _
            "'" & oPOQty4 & "'," & _
            "'" & oTotalDue & "'," & _
            "'" & oOnHand & "'," & _
            "'" & oWeeksOfStock & "'," & _
            "'" & oReorderLevel & "'," & _
            "'" & oMinQty & "'," & _
            "'" & oBOQty & "'," & _
            "'" & oRecmdBuy & "'," & _
            "'" & Now.ToShortDateString & "')"
      Using conn As New SqlConnection(ITEMDBSTR)
         Dim cmd As New SqlCommand(queryString, conn)
         Try
            conn.Open()
            cmd.ExecuteNonQuery()
         Catch ex As Exception
            Console.WriteLine(queryString)
            Console.WriteLine(ex.Message)
         End Try
      End Using
   End Sub
End Class
