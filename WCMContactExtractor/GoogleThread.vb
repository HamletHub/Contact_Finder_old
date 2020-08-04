Imports System.ComponentModel
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Xml
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Public Class GoogleThread
    Private WithEvents BWorker As BackgroundWorker 'Declaring Background worker
    Public ThreadWorking As Boolean = False
    Public ThreadInfo As String
    Public driver2 As IWebDriver
    ReadOnly sql As New SQLiteControl()

    Dim CityarrayNew As New ArrayList

    Public Sub InitializeGoogleThread()
        ThreadWorking = True
        ThreadInfo = "Initializing thread..."
        Dim driverService2 = ChromeDriverService.CreateDefaultService()
        driverService2.HideCommandPromptWindow = True
        Dim optionOn As New ChromeOptions
        optionOn.AddUserProfilePreference("profile.default_content_setting_values.geolocation", 2)
        optionOn.AddArgument("start-maximized")
        optionOn.AddArgument("--disable-infobars")
        optionOn.AddArgument("--lang=en-GB")
        optionOn.AddArguments("headless")
        optionOn.AddUserProfilePreference("profile.default_content_setting_values.images", 2) 'Disable or enable images
        optionOn.AddArgument("--blink-settings=imagesEnabled=false") 'Disable or enable images

        driver2 = New ChromeDriver(driverService2, optionOn)

        Dim element As IWebElement

        driver2.Navigate.GoToUrl("https://google.com")

        Try
            element = driver2.FindElement(By.LinkText("English"))
            element.Click()
        Catch ex As Exception
        End Try

        BWorker = New BackgroundWorker 'Initializing a new instance of Background worker
        BWorker.RunWorkerAsync() 'Run previously initialized background worker

    End Sub

    Private Sub BwRunSelenium_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles BWorker.DoWork
        Dim FinishWithScraping As Boolean = False
        Do
            Dim SelectedKeyword As String = ""

            ThreadInfo = "Getting keywords..."

            Try
                sql.ExecQuery("SELECT BKeywords FROM PendingKeywords;")
                SelectedKeyword = sql.DBDT.Rows(0).Item(0)
            Catch ex As Exception
                MsgBox("NO KEYWORDS LEFT IN THE DATABASE! PLEASE FILL IT WITH NEW KEYWORDS TO CONTINUE WORKING")
                Exit Sub
            End Try

            Dim TempListOfKeywords As New RichTextBox
            Dim SuggestedKeywords As New ArrayList

            Dim ListOfGoogleSearchLinks As New ArrayList
            '    'Getting keyword sugggestions from google

            TempListOfKeywords.Text = EnterKeywordOnGoogle(SelectedKeyword)
            TempListOfKeywords.AppendText(Environment.NewLine & SelectedKeyword)

            Dim ExceptKeyword As String = SelectedKeyword

            For Each line As String In TempListOfKeywords.Lines
                If Not line = "" Then If IsCleanTownOnTheList(line) = True Then If CheckDoneKeyword(line) = False Then If CheckPendingKeyword(line, ExceptKeyword) = False Then SuggestedKeywords.Add(line)
            Next

            ThreadInfo = "Getting google search result..."

            Thread.Sleep(1000)

            For Each SugKeyword As String In SuggestedKeywords
                ScrapeGoogle(SugKeyword)

                sql.AddParam("@item", SelectedKeyword) : sql.ExecQuery("DELETE FROM PendingKeywords WHERE BKeywords LIKE @item")
                sql.AddParam("@kword", SelectedKeyword) : sql.ExecQuery("INSERT INTO GoogleDoneKeywords(keywords) VALUES(@kword);")

                For Each item As String In SuggestedKeywords
                    sql.AddParam("@kword", item) : sql.ExecQuery("INSERT INTO GoogleDoneKeywords(keywords) VALUES(@kword);")
                Next
            Next


            sql.ExecQuery("SELECT bplaceid FROM bcontacts;")
            If sql.HasExpetion(True) Then FinishWithScraping = True
            Dim ii As Integer = 0
            For Each r As DataRow In sql.DBDT.Rows : ii += 1 : Next
            If ii > 150 Then FinishWithScraping = True

        Loop Until FinishWithScraping = True
    End Sub

    Private Sub BWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles BWorker.RunWorkerCompleted
        QuitGoogleThread()
    End Sub

    Private Sub QuitGoogleThread()
        driver2.Dispose()
        driver2.Quit()
        sql.DBCon.Close()
        ThreadWorking = False
        ThreadInfo = ""
    End Sub


    Private Sub ScrapeGoogle(ByVal link As String)
        driver2.Navigate.GoToUrl("https://www.google.com/maps/search/recreation+and+sports+in+newtown+ct/@41.4066774,-73.3320721,14z")
        Thread.Sleep(3000)
        Dim element As IWebElement
        element = driver2.FindElement(By.Id("searchboxinput"))
        element.Clear()
        Thread.Sleep(1500)
        element.SendKeys(link)
        Thread.Sleep(1500)
        element.SendKeys(Keys.Enter)

        Dim NoMorePages As Boolean = False

        Dim CurrentPageInteger As Integer = 1

        Do

            Do : Loop Until WaitForClassElement("section-result") = True

            CurrentPageInteger += 1

            Dim elementTexts As List(Of String) = New List(Of String)(driver2.FindElements(By.ClassName("section-result-content")).[Select](Function(iw) iw.GetAttribute("outerHTML")))
            Dim i As Integer = 1
            For Each ContactEntry As String In elementTexts
                If FWorkSpace.QuitGoogleThread = True Then QuitGoogleThread() : Exit Sub

                ThreadInfo = "Google: Page - " & CurrentPageInteger & " - item: " & i.ToString
                Thread.Sleep(3000)
                i += 2
                Dim BusinessName As String = ""
                Dim FullBusinessName As String = ""
                Dim BusinessAddress As String = ""
                Dim BusinessPhone As String = ""
                Dim TempWebsite As String = ""
                Dim BusinessWebsite As String = ""

                Dim ShouldSaveEntry As Boolean = False

                Dim TempBusinessName As String = ExtractData(ContactEntry, "class=""section-result-title""><span ", "</span>")
                BusinessName = GetStringBeforeOrAfter(TempBusinessName, """>", False, True)
                TempWebsite = ReturnMatchedURLS(ContactEntry, False)

                If CheckBName(BusinessName) = False Then

                    Try
                        element = driver2.FindElement(By.XPath("//*[@id='pane']/div/div[1]/div/div/div[4]/div[1]/div[" & i.ToString & "]/div[1]"))

                        Thread.Sleep(1000)

                        element.Click()
                    Catch ex As Exception

                    End Try


                    If BusinessName = "Stamford IT Consultants" Then
                        BusinessName = "Stamford IT Consultants"
                    End If

                    If TempWebsite = "False" Then TempWebsite = ""

                    If Not TempWebsite = "" Then

                        BusinessWebsite = FormatWebsite(TempWebsite)

                        If Not CheckDuplicateWebsite(BusinessWebsite) = True Then
                            If WaitForXPathElement("//*[@id='pane']/div/div[1]/div/div/button/span") = True Then
                                element = driver2.FindElement(By.Id("pane"))
                                Dim BusRTB As New RichTextBox
                                BusRTB.Text = element.GetAttribute("outerHTML")
                                Try

                                    Dim RemoveLeft As New RichTextBox
                                    RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Address: ", False, True)

                                    Dim RemoveRight As New RichTextBox
                                    RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False)
                                    BusinessAddress = RemoveRight.Text


                                    RemoveLeft.Text = ""
                                    RemoveRight.Text = ""
                                    RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Phone: ", False, True)

                                    RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False)
                                    BusinessPhone = RemoveRight.Text

                                    RemoveLeft.Text = ""
                                    RemoveRight.Text = ""
                                    RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Website: ", False, True)

                                    'RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False)
                                    'BusinessWebsite = RemoveRight.Text

                                    ShouldSaveEntry = True
                                Catch ex As Exception
                                End Try

                            End If
                        End If
                    End If

                    If BusinessWebsite = "" Then
                        If WaitForXPathElement("//*[@id='pane']/div/div[1]/div/div/button/span") = True Then
                            element = driver2.FindElement(By.Id("pane"))
                            Dim BusRTB As New RichTextBox
                            BusRTB.Text = element.GetAttribute("outerHTML")
                            Try
                                Dim RemoveLeft As New RichTextBox
                                Dim RemoveRight As New RichTextBox
                                RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Website: ", False, True)
                                RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False)
                                TempWebsite = RemoveRight.Text

                                If TempWebsite = "False" Then
                                    TempWebsite = ""

                                    RemoveLeft.Text = ""
                                    RemoveRight.Text = ""

                                    RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Address: ", False, True)


                                    RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False)
                                    BusinessAddress = RemoveRight.Text


                                    RemoveLeft.Text = ""
                                    RemoveRight.Text = ""
                                    RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Phone: ", False, True)

                                    RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False)
                                    BusinessPhone = RemoveRight.Text

                                    ShouldSaveEntry = True
                                Else
                                    BusinessWebsite = FormatWebsite(TempWebsite)

                                    If Not CheckDuplicateWebsite(BusinessWebsite) = True Then
                                        RemoveLeft.Text = ""
                                        RemoveRight.Text = ""

                                        RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Address: ", False, True)


                                        RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False)
                                        BusinessAddress = RemoveRight.Text


                                        RemoveLeft.Text = ""
                                        RemoveRight.Text = ""
                                        RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Phone: ", False, True)

                                        RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False)
                                        BusinessPhone = RemoveRight.Text

                                        ShouldSaveEntry = True
                                    End If
                                End If

                            Catch ex As Exception
                            End Try

                        End If
                    End If

                    Thread.Sleep(3500)

                    Try
                        element = driver2.FindElement(By.XPath("//*[@id='pane']/div/div[1]/div/div/button/span"))
                        element.Click()
                    Catch ex As Exception
                        Exit For
                    End Try

                    BusinessName = BusinessName.Replace("&amp;", "&")
                    If BusinessAddress = "False" Then BusinessAddress = ""
                    If BusinessPhone = "False" Then BusinessPhone = ""

                    If ShouldSaveEntry = True Then
                        If IsTownOnTheList(BusinessAddress) = True Then SaveIntoDatabase(BusinessName, BusinessAddress, BusinessPhone, BusinessWebsite, "GMB") Else SaveIntoOTDatabase(BusinessName, BusinessAddress, BusinessPhone, BusinessWebsite, "GMB")
                    Else
                        BlackListEntry(BusinessName, BusinessWebsite)
                    End If
                End If
            Next

            Thread.Sleep(1000)

            Try
                element = driver2.FindElement(By.XPath("/html/body/jsl/div[3]/div[9]/div[8]/div/div[1]/div/div/div[4]/div[2]/div/div[1]/div/button[2]/span")) : element.Click()

                Thread.Sleep(5000)
            Catch ex As Exception
                NoMorePages = True
            End Try
        Loop Until NoMorePages = True
    End Sub

    Public Function FormatWebsite(ByVal BusinessWebsite As String)
        BusinessWebsite = BusinessWebsite.Replace("https://www.", "").Replace("http://www.", "").Replace("https://", "").Replace("http://", "").Replace("www.", "")
        Dim TempBusinessWebsite As String = BusinessWebsite
        If TempBusinessWebsite.Contains("/") Then BusinessWebsite = TempBusinessWebsite.Substring(0, TempBusinessWebsite.IndexOf("/"))
        BusinessWebsite = BusinessWebsite.Replace("/", "")

        Return BusinessWebsite
    End Function

    Private Sub BlackListEntry(ByVal BName As String, ByVal BWebsite As String)
        Dim BWebsiteExists As Boolean

        Dim BNameExists As Boolean = CheckBName(BName)
        If Not BWebsite = "" Then BWebsiteExists = CheckDuplicateWebsite(BWebsite)

        If BNameExists = False Then
            sql.AddParam("@EN", BName)
            sql.ExecQuery("INSERT INTO bDuplicateNames(bnames) VALUES(@EN);")
        End If

        If BWebsiteExists = False Then
            sql.AddParam("@EN", BWebsite)
            sql.ExecQuery("INSERT INTO bDuplicateWebsites(bwebsites) VALUES(@EN);")
        End If

    End Sub

    Private Sub SaveIntoDatabase(ByVal BName, ByVal BAddress, ByVal BPhone, ByVal BWebsite, ByVal EleIndustry)
        Dim PiD As String = "NONE"

        sql.AddParam("@piD", PiD)
        sql.AddParam("@EN", BName)
        sql.AddParam("@EA", BAddress)
        sql.AddParam("@EP", BPhone)
        sql.AddParam("@EW", BWebsite)
        sql.AddParam("@EI", EleIndustry)

        'Execute query - Insert results into the database
        sql.ExecQuery("INSERT INTO bcontacts(bplaceid, bname, bphone, baddress, bwebsite, btype) VALUES(@piD, @EN, @EP, @EA, @EW, @EI);")
    End Sub
    Private Sub SaveIntoOTDatabase(ByVal BName, ByVal BAddress, ByVal BPhone, ByVal BWebsite, ByVal EleIndustry)
        Dim PiD As String = "NONE"

        sql.AddParam("@piD", PiD)
        sql.AddParam("@EN", BName)
        sql.AddParam("@EA", BAddress)
        sql.AddParam("@EP", BPhone)
        sql.AddParam("@EW", BWebsite)
        sql.AddParam("@EI", EleIndustry)

        'Execute query - Insert results into the database
        sql.ExecQuery("INSERT INTO bcontactsOT(bplaceid, bname, bphone, baddress, bwebsite, btype) VALUES(@piD, @EN, @EP, @EA, @EW, @EI);")
    End Sub


    Private Function GetStringBeforeOrAfter(ByVal TString As String, ByVal TSeparator As String, ByVal TLeft As Boolean, ByVal TRight As Boolean)
        Dim StringOnTheRight As String
        Dim StringOnTheLeft As String

        Try
            Dim original As String = TString : Dim cut_at As String = TSeparator : Dim stringSeparators() As String = {cut_at} : Dim split = original.Split(stringSeparators, 2, StringSplitOptions.RemoveEmptyEntries)
            StringOnTheRight = split(1)
            StringOnTheLeft = split(0)
        Catch ex As Exception
            Return False
        End Try

        If TLeft = True Then Return StringOnTheLeft

        If TRight = True Then Return StringOnTheRight

        If TLeft = False And TRight = False Then Return False

        Return False
    End Function

    Private Function ReturnMatchedURLS(ByVal outerHTML As String, ByVal RequestMoreDetails As Boolean)
        Dim oHTML As String = outerHTML
        Dim MappedUrlList As Boolean = RequestMoreDetails

        Dim LinkURLList As New ArrayList

        Dim LinksRTB As New RichTextBox

        Dim links As MatchCollection = Regex.Matches(oHTML, "<a.*?href=""(.*?)"".*?>(.*?)</a>")
        For Each match As Match In links
            ' Dim dr As DataRow = dt.NewRow
            Dim matchUrl As String = match.Groups(1).Value
            'Ignore all anchor links
            If matchUrl.StartsWith("#") Then Continue For
            'Ignore all javascript calls
            If matchUrl.ToLower.StartsWith("javascript:") Then Continue For
            'Ignore all email links
            If matchUrl.ToLower.StartsWith("mailto:") Then Continue For

            'For internal links, build the url mapped to the base address
            If Not matchUrl.StartsWith("http://") And Not matchUrl.StartsWith("https://") Then matchUrl = MapUrl(driver2.Url.ToString, matchUrl).Replace("&amp;", "&")

            LinkURLList.Add(matchUrl)

            LinksRTB.AppendText(matchUrl & " - " & match.Groups(2).Value.Replace("&amp;", "&") & Environment.NewLine)
        Next

        If MappedUrlList = False Then If LinkURLList.Count > 0 Then Return LinkURLList.Item(0) Else Return Nothing Else If Not LinksRTB.Text = "" Then Return LinksRTB.Text Else Return Nothing
    End Function

    Private Function IsTownOnTheList(ByVal s)
        CityarrayNew.Add("Danbury,")
        CityarrayNew.Add("Darien,")
        CityarrayNew.Add("Milford,")
        CityarrayNew.Add("New Canaan,")
        CityarrayNew.Add("Newtown,")
        CityarrayNew.Add("Norwalk,")
        CityarrayNew.Add("Ridgefield,")
        CityarrayNew.Add("Stamford,")
        CityarrayNew.Add("Westport,")

        For Each value As String In CityarrayNew
            If s.Contains(value) Then Return True
        Next
        Return False
    End Function
    Public Function IsCleanTownOnTheList(ByVal s As String, Optional ReturnCityName As Boolean = False)
        Dim CityarrayNew2 As New ArrayList From {
            "danbury",
            "darien",
            "milford",
            "new canaan",
            "newtown",
            "norwalk",
            "ridgefield",
            "stamford",
            "westport"
        }

        For Each value As String In CityarrayNew2
            If s.ToLower.Contains(value.ToLower) Then
                If ReturnCityName = True Then Return value Else Return True
            End If
        Next
        Return False
    End Function

    Private Function CheckDoneKeyword(ByVal line As Object)
        sql.AddParam("@bkeyword", line)
        sql.ExecQuery("SELECT * FROM GoogleDoneKeywords WHERE keywords LIKE @bkeyword;")

        If sql.RecordCount = 0 Then Return False Else Return True
    End Function
    Private Function CheckPendingKeyword(ByVal line As String, ByVal ExceptKeyword As String)
        If Not line.ToLower = ExceptKeyword.ToLower Then
            sql.AddParam("@bkeyword", line)
            sql.ExecQuery("SELECT * FROM PendingKeywords WHERE BKeywords LIKE @bkeyword;")

            If sql.RecordCount = 0 Then Return False Else Return True
        Else
            Return False
        End If
    End Function
    Private Function CheckBName(ByVal BusinessName)
        Dim BName As String = BusinessName

        Dim TempBName As String = BName
        If ReturnIfNameIsPresent(TempBName) = True Then Return True

        TempBName = TempBName.Replace(" Co Inc", "").Replace(" Co. Inc.", "").Replace(", LLC", "").Replace(", Inc.", "").Replace(" Inc.", "").Replace(" LLC", "").Replace(" INC", "").Replace(" llc", "").Replace(" Inc", "").Replace(" LLC.", "").Replace(",Inc.", "").Replace(" Inc", "").Replace(" Ltd", "")
        If ReturnIfNameIsPresent(TempBName) = True Then Return True


        If TempBName.Contains("&") Then
            TempBName = TempBName.Replace("&", "and")
            If ReturnIfNameIsPresent(TempBName) = True Then Return True
        ElseIf TempBName.Contains("and") Then
            TempBName = TempBName.Replace("and", "&")
            If ReturnIfNameIsPresent(TempBName) = True Then Return True
        End If


        If TempBName.Contains("’") Then
            TempBName = TempBName.Replace("’", "'")
            If ReturnIfNameIsPresent(TempBName) = True Then Return True
        ElseIf TempBName.Contains("'") Then
            TempBName = TempBName.Replace("'", "’")
            If ReturnIfNameIsPresent(TempBName) = True Then Return True
        End If


        If TempBName.Contains("’s") Then
            TempBName = TempBName.Replace("’s", "s")
            If ReturnIfNameIsPresent(TempBName) = True Then Return True
        ElseIf TempBName.Contains("'s") Then
            TempBName = TempBName.Replace("'s", "s")
            If ReturnIfNameIsPresent(TempBName) = True Then Return True
        End If

        Return False

    End Function

    Private Function ReturnIfNameIsPresent(ByVal BusinessName)
        Dim IsDuplicate As Boolean
        sql.AddParam("@bname", BusinessName)
        sql.ExecQuery("SELECT * FROM bcontacts WHERE bname LIKE @bname;")
        If sql.RecordCount = 0 Then IsDuplicate = False Else IsDuplicate = True

        If IsDuplicate = True Then Return True

        sql.AddParam("@bname", BusinessName)
        sql.ExecQuery("SELECT * FROM bcontactsOT WHERE bname LIKE @bname;")
        If sql.RecordCount = 0 Then IsDuplicate = False Else IsDuplicate = True

        If IsDuplicate = True Then Return True

        sql.AddParam("@bname", BusinessName)
        sql.ExecQuery("SELECT * FROM bDuplicateNames WHERE bnames LIKE @bname;")
        If sql.RecordCount = 0 Then IsDuplicate = False Else IsDuplicate = True

        If IsDuplicate = True Then Return True Else Return False
    End Function

    Private Function CheckDuplicateWebsite(ByVal BusinessWebsite As String)
        If BusinessWebsite = "" Then Return True

        Dim IsDuplicate As Boolean = False

        sql.AddParam("@bweb", BusinessWebsite.ToLower)
        sql.ExecQuery("SELECT * FROM bDuplicateWebsites WHERE bwebsites LIKE @bweb;")
        If sql.RecordCount = 0 Then IsDuplicate = False Else IsDuplicate = True

        If IsDuplicate = True Then Return True

        sql.AddParam("@bweb", BusinessWebsite.ToLower)
        sql.ExecQuery("SELECT * FROM bcontactsOT WHERE bwebsite LIKE @bweb;")
        If sql.RecordCount = 0 Then IsDuplicate = False Else IsDuplicate = True

        If IsDuplicate = True Then Return True

        sql.AddParam("@bweb", BusinessWebsite.ToLower)
        sql.ExecQuery("SELECT * FROM bcontacts WHERE bwebsite LIKE @bweb;")
        If sql.RecordCount = 0 Then IsDuplicate = False Else IsDuplicate = True

        If IsDuplicate = True Then Return True Else Return False

    End Function


    Private Function WaitForXPathElement(ByVal DElement As Object)
        Dim element As IWebElement
        Dim ThreadSleepCount As Integer = 0

        Dim ElementFound As Boolean = False
        Try
            Do
                Thread.Sleep(1000)
                ThreadSleepCount += 1

                If ThreadSleepCount = 10 Then Return False

                element = driver2.FindElement(By.XPath(DElement))
                ElementFound = True

            Loop Until ElementFound = True
        Catch ex As Exception
        End Try

        If ElementFound = True Then Return True Else Return False

    End Function
    Private Function WaitForClassElement(ByVal DElement As Object)
        Dim element As IWebElement
        Dim ThreadSleepCount As Integer = 0

        Dim ElementFound As Boolean = False
        Try
            Do
                Thread.Sleep(1000)
                ThreadSleepCount += 1

                If ThreadSleepCount = 10 Then Return False

                element = driver2.FindElement(By.ClassName(DElement))
                ElementFound = True

            Loop Until ElementFound = True
        Catch ex As Exception
        End Try

        If ElementFound = True Then Return True Else Return False

    End Function
    Private Function WaitForIDElement(ByVal DElement As Object)
        Dim element As IWebElement
        Dim ThreadSleepCount As Integer = 0

        Dim ElementFound As Boolean = False
        Try
            Do
                Thread.Sleep(1000)
                ThreadSleepCount += 1

                If ThreadSleepCount = 10 Then Return False

                element = driver2.FindElement(By.Id(DElement))
                ElementFound = True

            Loop Until ElementFound = True
        Catch ex As Exception
        End Try

        If ElementFound = True Then Return True Else Return False

    End Function

    Public Function ExtractData(ByRef pHTML As String, ByRef pSearchStart As String, ByRef pSearchEnd As String, Optional ByRef pSearchSub As String = "") As String
        Try
            Dim lonPhrasePos1 As Integer
            Dim lonPhrasePos2 As Integer

            ExtractData = ""
            ' Look for the search phrase. If not found, exit this function.
            lonPhrasePos1 = pHTML.ToUpper.IndexOf(pSearchStart.ToUpper)
            If lonPhrasePos1 = 0 Then Exit Function
            ' If the optional pSearchSub parameter has been provided, find the string AFTER the
            ' first search string's location.
            If pSearchSub <> "" Then
                lonPhrasePos1 = pHTML.IndexOf(pSearchSub.ToUpper, lonPhrasePos1 + 1)
                If lonPhrasePos1 = 0 Then Exit Function
            End If
            ' Now look for the ending search phrase. Everything in-between gets returned as the value
            ' of this function. If the ending search phrase is not found, exit this function.
            lonPhrasePos2 = pHTML.ToUpper.IndexOf(pSearchEnd.ToUpper, lonPhrasePos1)
            If lonPhrasePos2 = 0 Or lonPhrasePos1 = lonPhrasePos2 Then Exit Function

            ' Extract the data between the two given search strings.
            If pSearchSub <> "" Then
                ExtractData = pHTML.Substring(lonPhrasePos1 + Len(pSearchSub), lonPhrasePos2 - (lonPhrasePos1 + Len(pSearchSub)))
            Else
                ExtractData = pHTML.Substring(lonPhrasePos1 + Len(pSearchStart), lonPhrasePos2 - (lonPhrasePos1 + Len(pSearchStart)))
            End If
        Catch ex As Exception

        End Try

    End Function

    Private Sub SendEnterOnGoogle()
        Dim element As IWebElement
        Try
            element = driver2.FindElement(By.XPath("//*[@id='tsf']/div[2]/div/div[1]/div/div[1]/input"))
            element.SendKeys(Keys.Enter)
        Catch ex As Exception
            element = driver2.FindElement(By.XPath("//*[@id='tsf']/div[2]/div[1]/div[1]/div/div[2]/input"))
            element.SendKeys(Keys.Enter)
        End Try
    End Sub

    Private Function EnterKeywordOnGoogle(ByVal Keyword As String)
        Dim element As IWebElement
        driver2.Navigate.GoToUrl("https://google.com")
        Try
            element = driver2.FindElement(By.XPath("//*[@id='tsf']/div[2]/div/div[1]/div/div[1]/input"))
            element.Clear()
            element.SendKeys(Keyword.ToLower)

            Thread.Sleep(3000)

            element = driver2.FindElement(By.XPath("//*[@id='tsf']/div[2]/div[1]/div[2]/div[2]/ul"))
            Return element.Text

        Catch ex As Exception
            element = driver2.FindElement(By.XPath("//*[@id='tsf']/div[2]/div[1]/div[1]/div/div[2]/input"))
            element.Clear()
            element.SendKeys(Keyword.ToLower)

            Thread.Sleep(3000)

            element = driver2.FindElement(By.XPath("//*[@id='tsf']/div[2]/div[1]/div[2]/div[2]/ul"))
            Return element.Text
        End Try

        Return "NONE"
    End Function


    Private Function ContainBannedWords(ByVal matchUrl)
        Dim BWords As New ArrayList From {
            "webcache.googleusercontent.com",
            "=related:",
            "FAQ_Answers",
            "LocationPhotoDirectLink",
            "ShowUserReviews"
        }

        Dim MUrl As String = matchUrl
        For Each item As String In BWords
            If MUrl.Contains(item) Then Return True
        Next
        Return False
    End Function
    Private Function IsGMBLink(ByVal link)
        If link.Contains("https://www.google.com/search?q=") Then If Not link.Contains("=related:") Then If link.Contains("rldoc") Then Return True Else Return False Else Return False Else Return False
    End Function

    Public Function MapUrl(ByVal baseAddress As String, ByVal relativePath As String) As String
        Dim u As New Uri(baseAddress)

        If relativePath = "./" Then relativePath = "/"

        If relativePath.StartsWith("/") Then
            Return u.Scheme + Uri.SchemeDelimiter + u.Authority + relativePath
        Else
            Dim pathAndQuery As String = u.AbsolutePath

            pathAndQuery = pathAndQuery.Split("?")(0).TrimEnd("/")
            If pathAndQuery.Split("/")(pathAndQuery.Split("/").Count - 1).Contains(".") Then pathAndQuery = pathAndQuery.Substring(0, pathAndQuery.LastIndexOf("/"))
            baseAddress = u.Scheme + Uri.SchemeDelimiter + u.Authority + pathAndQuery

            While relativePath.StartsWith("../")
                relativePath = relativePath.Substring(3)
                If baseAddress.LastIndexOf("/") > baseAddress.IndexOf("//" + 2) Then baseAddress = baseAddress.Substring(0, baseAddress.LastIndexOf("/")).TrimEnd("/")
            End While

            Return baseAddress + "/" + relativePath
        End If
    End Function
End Class
