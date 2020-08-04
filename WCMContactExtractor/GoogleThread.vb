Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Public Class GoogleThread
    Private WithEvents BWorker As BackgroundWorker 'Declaring Background worker
    Public ThreadWorking As Boolean = False 'Declaring Thread boolean to know when thread is working
    Public driver2 As IWebDriver 'Chrome WebDriver
    ReadOnly sql As New SQLiteControl() 'Declare SQL variable

    Dim CityarrayNew As New ArrayList 'City array list

    Dim CurrentLogPath As String

    Public Sub InitializeGoogleThread()
        Dim theContDir As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) 'Specify MyDocuments folder for WCM Logs
        Dim newDir As String = "WCM Logs" 'Folder name
        Dim thePath As String = IO.Path.Combine(theContDir, newDir) 'Combine MyDocuments with folder name
        If Not Directory.Exists(thePath) Then Directory.CreateDirectory(thePath) 'Create directory if it does not exists

        Dim localDate = DateTime.Now
        CurrentLogPath = thePath & "Log:" & localDate & ".txt"

        Dim fs As FileStream = File.Create(CurrentLogPath) ' Create or overwrite the file.

        Dim info As Byte() = New UTF8Encoding(True).GetBytes("Google Thread session - " & localDate & Environment.NewLine) ' Add text to the file.
        fs.Write(info, 0, info.Length)
        fs.Close()

        ThreadWorking = True 'Set this to true to inform the app this thread is working

        SaveLog("Setup - Initializing Chrome instance")

        Dim driverService2 = ChromeDriverService.CreateDefaultService()
        driverService2.HideCommandPromptWindow = True  'Hide command prompt from user
        Dim optionOn As New ChromeOptions
        optionOn.AddUserProfilePreference("profile.default_content_setting_values.geolocation", 2) 'Disable location tracking for this Chrome instance
        optionOn.AddArgument("start-maximized") 'Instruct Chrome to run in maximized mode
        optionOn.AddArgument("--disable-infobars") 'Disable infobars
        optionOn.AddArgument("--lang=en-GB") 'Set default language to English
        optionOn.AddArguments("headless") 'Hide Chrome browser from the user. Headless mode means Chrome will run in the background
        optionOn.AddUserProfilePreference("profile.default_content_setting_values.images", 2) 'Disable or enable images
        optionOn.AddArgument("--blink-settings=imagesEnabled=false") 'Disable or enable images

        driver2 = New ChromeDriver(driverService2, optionOn) 'Initialize ChromeDriver with added arguments

        SaveLog("Setup - Successfully Initialized Chrome instance")

        Dim element As IWebElement ' Declare Chrome iWebElement

        driver2.Navigate.GoToUrl("https://google.com") 'Navigate browser to Google homepage

        Try
            element = driver2.FindElement(By.LinkText("English")) 'Find "English" button and click on it (to set default Google language)
            element.Click() 'Click on that button
            SaveLog("Setup - Google changed to English language")
        Catch ex As Exception
        End Try

        SaveLog("Setup - Initializing Background worker")
        BWorker = New BackgroundWorker 'Initializing a new instance of Background worker
        BWorker.RunWorkerAsync() 'Run previously initialized background worker
    End Sub
    Private Sub BwRunSelenium_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles BWorker.DoWork
        Dim FinishWithScraping As Boolean = False 'Declare variable to indicate when scraping is completed
        SaveLog("Background worker - Successfully started")
        Do
            Dim SelectedKeyword As String 'Declare variable to store selected keyword
            SaveLog("Background worker - Getting keyword from SQL database")
            Try
                sql.ExecQuery("SELECT BKeywords FROM PendingKeywords;") 'Select table
                SelectedKeyword = sql.DBDT.Rows(0).Item(0) 'Get first row / keywor5d
                SaveLog("Background worker - Selected SQL Keyword from Pending keywords: " & SelectedKeyword.ToString)
            Catch ex As Exception
                MsgBox("NO KEYWORDS LEFT IN THE DATABASE! PLEASE FILL IT WITH NEW KEYWORDS TO CONTINUE WORKING") 'In case of an error, that means there are no keywords in the database
                SaveLog("Background worker - NO KEYWORDS IN THE DATABASE - ABORT")
                Exit Sub
            End Try

            Dim TempListOfKeywords As New RichTextBox, SuggestedKeywords As New ArrayList, ListOfGoogleSearchLinks As New ArrayList

            TempListOfKeywords.Text = EnterKeywordOnGoogle(SelectedKeyword) 'Enter selected keyword to google.com
            TempListOfKeywords.AppendText(Environment.NewLine & SelectedKeyword) 'Append additional keywords to the RTB

            Dim ExceptKeyword As String = SelectedKeyword

            SaveLog("Background worker - Getting keywords from google.com")

            'Check if this keyword already exists, if not, add found keywords to the list
            For Each line As String In TempListOfKeywords.Lines
                If Not line = "" Then If IsCleanTownOnTheList(line) = True Then If CheckDoneKeyword(line) = False Then If CheckPendingKeyword(line, ExceptKeyword) = False Then SuggestedKeywords.Add(line)
            Next

            Thread.Sleep(1000) 'Sleep for 1000ms (1 second)

            SuggestedKeywords.Add(SelectedKeyword)

            For Each SugKeyword As String In SuggestedKeywords 'For each keyword on the list
                SaveLog("Background worker - Run ScrapeGoogle() sub - Keyword: " & SugKeyword.ToString)
                ScrapeGoogle(SugKeyword) 'Scrape google
                'After scraping is done, delete this keyword from "PendingKeywords" table, and add it to "DoneKeywords" table
                SaveLog("Background worker - Deleting " & SugKeyword & " from PendingKeywords table")
                sql.AddParam("@item", SugKeyword) : sql.ExecQuery("DELETE FROM PendingKeywords WHERE BKeywords LIKE @item")
                SaveLog("Background worker - Inserting " & SugKeyword & " to the GoogleDoneKeywords table")
                sql.AddParam("@kword", SugKeyword) : sql.ExecQuery("INSERT INTO GoogleDoneKeywords(keywords) VALUES(@kword);")
            Next

            sql.ExecQuery("SELECT bplaceid FROM bcontacts;") 'Select bcontacts table
            If sql.HasExpetion(True) Then
                FinishWithScraping = True 'If SQL has exception, finish with scraping and try again later
                SaveLog("Background worker - Google thread finished - SQL Exception: " & sql.Exception.ToString)
            End If

            Dim rowsCount As Integer = 0
            For Each r As DataRow In sql.DBDT.Rows : rowsCount += 1 : Next 'Count rows
            If rowsCount > 150 Then
                FinishWithScraping = True 'If there are more than 150 contacts in the table, stop with scraping and continue later
                SaveLog("Background worker - Google thread finished - rowsCount > 150")
            End If
        Loop Until FinishWithScraping = True
        SaveLog("Background worker - task completed")
    End Sub
    Private Sub ScrapeGoogle(ByVal link As String)
        SaveLog("Google thread - Navigating to google.com maps")
        driver2.Navigate.GoToUrl("https://www.google.com/maps/search/recreation+And+sports+in+newtown+ct/@41.4066774,-73.3320721,14z") 'Navigate to Google maps
        Dim element As IWebElement 'Declare Chrome iWebElement
        Do : Loop Until WaitForElement("searchboxinput", "ID", 15) = True 'Wait for the page to load
        element = driver2.FindElement(By.Id("searchboxinput")) 'Find search textbox on Google Maps
        element.Clear() 'Clear it
        Thread.Sleep(500) 'Wait for a moment
        element.SendKeys(link) 'Send keyword
        Thread.Sleep(500) 'Wait for a moment
        element.SendKeys(Keys.Enter) 'Press enter
        SaveLog("Google thread - keyword entered")

        Dim NoMorePages As Boolean = False 'Variable to determine when there are no more pages

        Do 'Do all this until app can't find "Next" button
            SaveLog("Google thread - Waiting for results to load")
            Do : Loop Until WaitForElement("section-result", "Class", 15) = True 'Wait for the page to load
            SaveLog("Google thread - Results loaded")
            Dim elementTexts As List(Of String) = New List(Of String)(driver2.FindElements(By.ClassName("section-result-content")).[Select](Function(iw) iw.GetAttribute("outerHTML"))) 'Get all elements with contacts data
            Dim ContactID As Integer = 1 'Variable to determine contact ID
            For Each ContactEntry As String In elementTexts 'For each contact on the page...
                If FWorkSpace.QuitGoogleThread = True Then
                    QuitGoogleThread() 'Checks if user requested to stop scraping
                    SaveLog("Google thread - User requested to abort scraping")
                    Exit Sub
                End If

                ContactID += 2 'First contact (starts at #3)
                SaveLog("Google thread - Processing contact ID: " & ContactID)

                'Declare contact details variables
                Dim BusinessName As String = "", FullBusinessName As String = "", BusinessAddress As String = "", BusinessPhone As String = "", TempWebsite As String = "", BusinessWebsite As String = ""

                Dim ShouldSaveEntry As Boolean = False 'This variable indicates if we should save this contact or not

                Dim TempBusinessName As String = ExtractData(ContactEntry, "class=""section-result-title""><span ", "</span>") 'Extract contact name string
                BusinessName = GetStringBeforeOrAfter(TempBusinessName, """>", False, True) 'Format string
                TempWebsite = ReturnMatchedURLS(ContactEntry, False) 'Extract website

                SaveLog("Google thread - Processing contact name: " & BusinessName)

                If CheckBName(BusinessName) = False Then 'If contact name does not exist in the database, we can open it and proceed with scraping
                    Try
                        SaveLog("Google thread - Opening contact window")
                        element = driver2.FindElement(By.XPath("//*[@id='pane']/div/div[1]/div/div/div[4]/div[1]/div[" & ContactID.ToString & "]/div[1]")) 'Find contact element
                        Thread.Sleep(1000) 'Wait for a second
                        element.Click() 'Click on the contact
                    Catch ex As Exception
                        SaveLog("Google thread - Failed to open contact window: " & ex.Message)
                    End Try

                    If TempWebsite = "False" Then TempWebsite = "" 'If website is false, set it to nothing. This sometimes can happen when function does not find website, and then it returns "False" instead of nothing

                    If Not TempWebsite = "" Then 'If website is not nothing...
                        SaveLog("Google thread - " & BusinessName & " website found: " & TempWebsite)
                        BusinessWebsite = FormatWebsite(TempWebsite) 'Format website
                        SaveLog("Google thread - Formatted website: " & BusinessWebsite)

                        If Not CheckDuplicateWebsite(BusinessWebsite) = True Then 'If this website is not in the database, proceed with scraping
                            If WaitForElement("//*[@id='pane']/div/div[1]/div/div/button/span", "xPath", 15) = True Then 'If contact details are loaded in the browser
                                element = driver2.FindElement(By.Id("pane")) 'Find main ID element
                                Dim BusRTB As New RichTextBox 'Variable to store outerHTML of the element
                                BusRTB.Text = element.GetAttribute("outerHTML") 'Set element outerHTML to the RTB variable
                                SaveLog("Google thread - Getting " & BusinessWebsite & " outerHTML")
                                Try
                                    Dim RemoveLeft As New RichTextBox, RemoveRight As New RichTextBox 'RTBs to store text from the left and from the right

                                    RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Address: ", False, True) 'Get address string from the left
                                    RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False) 'Get address string from the right
                                    BusinessAddress = RemoveRight.Text 'String between is our contact address

                                    SaveLog("Google thread - Contact address: " & BusinessAddress)

                                    RemoveLeft.Text = "" 'Clear variable
                                    RemoveRight.Text = "" 'Clear variable

                                    RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Phone: ", False, True) 'Get phone string from the left
                                    RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False) 'Get phone string from the right
                                    BusinessPhone = RemoveRight.Text 'String between is our contact phone

                                    SaveLog("Google thread - Contact phone: " & BusinessPhone)

                                    ShouldSaveEntry = True 'All good - we can save this contact
                                Catch ex As Exception
                                End Try
                            Else
                                SaveLog("Google thread - element xPath: //*[@id='pane']/div/div[1]/div/div/button/span not found")
                            End If
                        Else
                            SaveLog("Google thread - " & BusinessWebsite & " already exists in the database")
                        End If
                    End If

                    If BusinessWebsite = "" Then 'If business website is nothing...
                        SaveLog("Google thread -" & BusinessName & " website not found")
                        If WaitForElement("//*[@id='pane']/div/div[1]/div/div/button/span", "xPath", 15) = True Then 'If contact details are loaded in the browser
                            element = driver2.FindElement(By.Id("pane")) 'Find main ID element
                            Dim BusRTB As New RichTextBox 'Variable to store outerHTML of the element
                            BusRTB.Text = element.GetAttribute("outerHTML") 'Set element outerHTML to the RTB variable
                            SaveLog("Google thread - Getting " & BusinessWebsite & " outerHTML")
                            Try
                                Dim RemoveLeft As New RichTextBox, RemoveRight As New RichTextBox 'RTBs to store text from the left and from the right

                                SaveLog("Google thread - Still trying to find website")

                                RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Website: ", False, True) 'Get website string from the left
                                RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False) 'Get website string from the right
                                TempWebsite = RemoveRight.Text 'String between is our contact website

                                If TempWebsite = "False" Then 'If TempWebsite is false, that means google did not provide website for this contact
                                    SaveLog("Google thread - Website not found")

                                    TempWebsite = "" 'Clear variable

                                    RemoveLeft.Text = "" 'Clear variable
                                    RemoveRight.Text = "" 'Clear variable

                                    RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Address: ", False, True) 'Get address string from the left
                                    RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False) 'Get address string from the right
                                    BusinessAddress = RemoveRight.Text 'String between is our contact address
                                    SaveLog("Google thread - Contact address: " & BusinessAddress)

                                    RemoveLeft.Text = "" 'Clear variable
                                    RemoveRight.Text = "" 'Clear variable

                                    RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Phone: ", False, True) 'Get phone string from the left
                                    RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False) 'Get phone string from the right
                                    BusinessPhone = RemoveRight.Text 'String between is our contact phone
                                    SaveLog("Google thread - Contact phone: " & BusinessPhone)

                                    ShouldSaveEntry = True 'All good - we can save this contact
                                Else 'If TempWebsite is not False, that means google provided website for this contact
                                    SaveLog("Google thread - " & BusinessName & " website found: " & TempWebsite)

                                    BusinessWebsite = FormatWebsite(TempWebsite) 'Format the website
                                    SaveLog("Google thread - Formatted website: " & BusinessWebsite)

                                    If Not CheckDuplicateWebsite(BusinessWebsite) = True Then 'If this website is not in the database, proceed with scraping
                                        RemoveLeft.Text = "" 'Clear variable
                                        RemoveRight.Text = "" 'Clear variable

                                        RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Address: ", False, True) 'Get address string from the left
                                        RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False) 'Get address string from the right
                                        BusinessAddress = RemoveRight.Text 'String between is our contact address
                                        SaveLog("Google thread - Contact address: " & BusinessAddress)

                                        RemoveLeft.Text = "" 'Clear variable
                                        RemoveRight.Text = "" 'Clear variable

                                        RemoveLeft.Text = GetStringBeforeOrAfter(BusRTB.Text, "Phone: ", False, True) 'Get phone string from the left
                                        RemoveRight.Text = GetStringBeforeOrAfter(RemoveLeft.Text, """", True, False) 'Get phone string from the right
                                        BusinessPhone = RemoveRight.Text 'String between is our contact phone
                                        SaveLog("Google thread - Contact phone: " & BusinessPhone)

                                        ShouldSaveEntry = True 'All good - we can save this contact
                                    Else
                                        SaveLog("Google thread - " & BusinessWebsite & " already exists in the database")
                                    End If
                                End If
                            Catch ex As Exception
                            End Try
                        Else
                            SaveLog("Google thread - element xPath: //*[@id='pane']/div/div[1]/div/div/button/span not found")
                        End If
                    End If

                    Thread.Sleep(1000) 'Wait for 3+ seconds
                    SaveLog("Google thread - Waiting for 'Back' button")
                    If WaitForElement("//*[@id='pane']/div/div[1]/div/div/button/span", "xPath", 10) = True Then 'If "Back" button is present...
                        Try
                            element = driver2.FindElement(By.XPath("//*[@id='pane']/div/div[1]/div/div/button/span")) 'Find "Back" button element
                            element.Click() 'Click on it
                            SaveLog("Google thread - Clicked on the back button")
                        Catch ex As Exception
                            SaveLog("Google thread - Exception occured while attempting to click on the 'back' button: " & ex.Message & " - Trying again...")
                            Thread.Sleep(5000) 'Wait for 5 seconds and try again
                            element = driver2.FindElement(By.XPath("//*[@id='pane']/div/div[1]/div/div/button/span")) 'Find "Back" button element
                            element.Click() 'Click on it
                            SaveLog("Google thread - Clicked on the back button")
                        End Try

                        SaveLog("Google thread - Formatting data...")
                        BusinessName = BusinessName.Replace("&amp;", "&") 'Format contact name
                        If BusinessAddress = "False" Then BusinessAddress = "" 'Format address
                        If BusinessPhone = "False" Then BusinessPhone = "" 'Format phone

                        If ShouldSaveEntry = True Then 'If we can save this contact...
                            If IsTownOnTheList(BusinessAddress) = True Then SaveIntoDatabase(BusinessName, BusinessAddress, BusinessPhone, BusinessWebsite, "GMB") Else SaveIntoOTDatabase(BusinessName, BusinessAddress, BusinessPhone, BusinessWebsite, "GMB") 'Save into adequate database (this splits contacts from towns we're working on from others into different SQL tables)
                        Else
                            SaveLog("Google thread - Blacklisting: " & BusinessName & " and " & BusinessWebsite)
                            BlackListEntry(BusinessName, BusinessWebsite) 'If we can't save this contact for any reason, add this contact to duplicate tables
                        End If
                    End If
                Else
                    SaveLog("Google thread - Duplicate contact: " & BusinessName)
                End If
            Next

            SaveLog("Google thread - Finished with this page. Trying to click on the 'Next' button")

            Thread.Sleep(1000) 'Wait for a second...

            Try
                element = driver2.FindElement(By.XPath("/html/body/jsl/div[3]/div[9]/div[8]/div/div[1]/div/div/div[4]/div[2]/div/div[1]/div/button[2]/span")) : element.Click() 'Try to find a "next" button on the google page, and if it exists, click on it
                SaveLog("Google thread - Clicked on the next button")
                Thread.Sleep(5000) 'Wait for 5 seconds
            Catch ex As Exception
                SaveLog("Google thread - No 'Next' button is present")
                NoMorePages = True 'In case of an error, set this to true.
            End Try
        Loop Until NoMorePages = True 'Means there are no more pages to navigate to, so scraping job is done! Moving on to the next keyword.
        SaveLog("Google thread - task completed")
    End Sub


#Region "Procedures"
    Public Sub SaveLog(ByVal LogLine As String)
        Dim theContDir As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) 'Specify MyDocuments folder for WCM Logs
        Dim newDir As String = "WCM Logs" 'Log Folder name
        Dim thePath As String = Path.Combine(theContDir, newDir) 'Combine MyDocuments with folder name

        Dim sb As StringBuilder = New StringBuilder()
        Dim localDate = DateTime.Now 'Get current time
        sb.AppendLine(localDate & ": " & LogLine & Environment.NewLine)
        File.AppendAllText(CurrentLogPath, sb.ToString())
    End Sub
    Private Sub SendEnterOnGoogle()
        Dim element As IWebElement 'Declare Chrome iWebElement
        Try
            element = driver2.FindElement(By.XPath("//*[@id='tsf']/div[2]/div/div[1]/div/div[1]/input")) 'Try to find element
            element.SendKeys(Keys.Enter) 'If element is present, send enter key
        Catch ex As Exception
            element = driver2.FindElement(By.XPath("//*[@id='tsf']/div[2]/div[1]/div[1]/div/div[2]/input")) 'Try to find element again
            element.SendKeys(Keys.Enter) 'If element is present, send enter key
        End Try
    End Sub
    Private Sub QuitGoogleThread()
        SaveLog("Google Thread - Quit Chrome driver and SQL")
        driver2.Dispose() 'Dispose Chrome instance
        driver2.Quit() 'Quit Chrome instance
        sql.DBCon.Close() 'Close SQL connection
        ThreadWorking = False 'Thread not working
        SaveLog("Google Thread - Success")
    End Sub
    Private Sub BlackListEntry(ByVal BName As String, ByVal BWebsite As String)
        'This sub adds contact to the database, in case we want to blacklist it (we don't need it)
        Dim BWebsiteExists As Boolean 'Specify if website exists

        Dim BNameExists As Boolean = CheckBName(BName) 'Check if business name exists in the database
        If Not BWebsite = "" Then BWebsiteExists = CheckDuplicateWebsite(BWebsite) 'Check if website exists in the database

        If BNameExists = False Then 'If business name does not exist...
            sql.AddParam("@EN", BName) 'Add SQL parameter
            sql.ExecQuery("INSERT INTO bDuplicateNames(bnames) VALUES(@EN);") 'Insert into duplicate database
        End If

        If BWebsiteExists = False Then 'If business website does not exist...
            sql.AddParam("@EN", BWebsite) 'Add SQL parameter
            sql.ExecQuery("INSERT INTO bDuplicateWebsites(bwebsites) VALUES(@EN);") 'Insert into duplicate database
        End If
    End Sub
    Private Sub SaveIntoDatabase(ByVal BName, ByVal BAddress, ByVal BPhone, ByVal BWebsite, ByVal EleIndustry)
        'This sub adds data to the bcontacts table
        Dim PiD As String = "NONE"
        'Add multiple parameters
        sql.AddParam("@piD", PiD)
        sql.AddParam("@EN", BName)
        sql.AddParam("@EA", BAddress)
        sql.AddParam("@EP", BPhone)
        sql.AddParam("@EW", BWebsite)
        sql.AddParam("@EI", EleIndustry)

        SaveLog("Saving following results into the database: " & BName & " - " & BAddress & " - " & BPhone & " - " & BWebsite & " - " & EleIndustry)

        'Execute query - Insert results into the database
        sql.ExecQuery("INSERT INTO bcontacts(bplaceid, bname, bphone, baddress, bwebsite, btype) VALUES(@piD, @EN, @EP, @EA, @EW, @EI);")
        If sql.HasExpetion Then SaveLog("FATAL SQL EXCEPTION: " & sql.Exception)
    End Sub
    Private Sub SaveIntoOTDatabase(ByVal BName, ByVal BAddress, ByVal BPhone, ByVal BWebsite, ByVal EleIndustry)
        'This sub adds data to the bcontactsOT table, which are all contacts from different towns & areas. We might use them later.
        Dim PiD As String = "NONE"
        'Add multiple parameters
        sql.AddParam("@piD", PiD)
        sql.AddParam("@EN", BName)
        sql.AddParam("@EA", BAddress)
        sql.AddParam("@EP", BPhone)
        sql.AddParam("@EW", BWebsite)
        sql.AddParam("@EI", EleIndustry)

        SaveLog("Saving following results into OT the database: " & BName & " - " & BAddress & " - " & BPhone & " - " & BWebsite & " - " & EleIndustry)

        'Execute query - Insert results into the database
        sql.ExecQuery("INSERT INTO bcontactsOT(bplaceid, bname, bphone, baddress, bwebsite, btype) VALUES(@piD, @EN, @EP, @EA, @EW, @EI);")
        If sql.HasExpetion Then SaveLog("FATAL SQL EXCEPTION: " & sql.Exception)
    End Sub
    Private Sub BWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles BWorker.RunWorkerCompleted
        QuitGoogleThread() 'Thread finished with scraping >> quit
    End Sub
#End Region

#Region "Functions"
    Private Function EnterKeywordOnGoogle(ByVal Keyword As String)
        Dim element As IWebElement 'Declare Chrome iWebElement
        driver2.Navigate.GoToUrl("https://google.com") 'Navigate to google.com
        Try
            element = driver2.FindElement(By.XPath("//*[@id='tsf']/div[2]/div/div[1]/div/div[1]/input")) 'Find element
            element.Clear() 'Clear it
            element.SendKeys(Keyword.ToLower) 'Enter keyword

            Thread.Sleep(3000) 'Sleep for 3 seconds

            element = driver2.FindElement(By.XPath("//*[@id='tsf']/div[2]/div[1]/div[2]/div[2]/ul")) 'Find element dropdown list
            Return element.Text 'Return list
        Catch ex As Exception
            element = driver2.FindElement(By.XPath("//*[@id='tsf']/div[2]/div[1]/div[1]/div/div[2]/input")) 'On error, try to find element again
            element.Clear() 'If element is present, clear it...
            element.SendKeys(Keyword.ToLower) 'And enter keyword

            Thread.Sleep(3000) 'Sleep for 3 seconds

            element = driver2.FindElement(By.XPath("//*[@id='tsf']/div[2]/div[1]/div[2]/div[2]/ul")) 'Find element dropdown list
            Return element.Text 'Return list
        End Try
        Return "" 'If no dropdown is present, return nothing
    End Function
    Private Function ContainBannedWords(ByVal matchUrl)
        'If match URL contains some of these words, discard it
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
    Public Function MapUrl(ByVal baseAddress As String, ByVal relativePath As String) As String
        Dim u As New Uri(baseAddress)
        'Map URL according to website host - some websites shows only pages without hostname. This function returns pages combined with hostname address.
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
    Public Function ExtractData(ByRef pHTML As String, ByRef pSearchStart As String, ByRef pSearchEnd As String, Optional ByRef pSearchSub As String = "") As String
        Try
            Dim lonPhrasePos1 As Integer, lonPhrasePos2 As Integer
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
    Private Function IsTownOnTheList(ByVal s)
        'Add town list
        CityarrayNew.Add("Danbury,")
        CityarrayNew.Add("Darien,")
        CityarrayNew.Add("Milford,")
        CityarrayNew.Add("New Canaan,")
        CityarrayNew.Add("Newtown,")
        CityarrayNew.Add("Norwalk,")
        CityarrayNew.Add("Ridgefield,")
        CityarrayNew.Add("Stamford,")
        CityarrayNew.Add("Westport,")

        'Check if this address contains town from the list above
        For Each value As String In CityarrayNew
            If s.Contains(value) Then Return True
        Next
        Return False
    End Function
    Public Function IsCleanTownOnTheList(ByVal s As String, Optional ReturnCityName As Boolean = False)
        'Check if town is on the list
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
        'Check if keyword already exists in the database
        If sql.RecordCount = 0 Then Return False Else Return True
    End Function
    Private Function CheckPendingKeyword(ByVal line As String, ByVal ExceptKeyword As String)
        'Check if keyword already exists in the database
        If Not line.ToLower = ExceptKeyword.ToLower Then
            sql.AddParam("@bkeyword", line)
            sql.ExecQuery("SELECT * FROM PendingKeywords WHERE BKeywords LIKE @bkeyword;")

            If sql.RecordCount = 0 Then Return False Else Return True
        Else
            Return False
        End If
    End Function
    Private Function CheckBName(ByVal BusinessName)
        Dim BName As String = BusinessName 'Declare variable to get business name

        Dim TempBName As String = BName 'Temporary business name
        If ReturnIfNameIsPresent(TempBName) = True Then Return True 'If business name exists in the database, return true

        TempBName = TempBName.Replace(" Co Inc", "").Replace(" Co. Inc.", "").Replace(", LLC", "").Replace(", Inc.", "").Replace(" Inc.", "").Replace(" LLC", "").Replace(" INC", "").Replace(" llc", "").Replace(" Inc", "").Replace(" LLC.", "").Replace(",Inc.", "").Replace(" Inc", "").Replace(" Ltd", "") 'If not, format business name and try again
        If ReturnIfNameIsPresent(TempBName) = True Then Return True 'If business name exists in the database, return true

        If TempBName.Contains("&") Then 'More formatting
            TempBName = TempBName.Replace("&", "and")
            If ReturnIfNameIsPresent(TempBName) = True Then Return True 'If business name exists in the database, return true
        ElseIf TempBName.Contains("and") Then
            TempBName = TempBName.Replace("and", "&")
            If ReturnIfNameIsPresent(TempBName) = True Then Return True 'If business name exists in the database, return true
        End If

        If TempBName.Contains("’") Then 'More formatting
            TempBName = TempBName.Replace("’", "'")
            If ReturnIfNameIsPresent(TempBName) = True Then Return True 'If business name exists in the database, return true
        ElseIf TempBName.Contains("'") Then
            TempBName = TempBName.Replace("'", "’")
            If ReturnIfNameIsPresent(TempBName) = True Then Return True 'If business name exists in the database, return true
        End If

        If TempBName.Contains("’s") Then 'More formatting
            TempBName = TempBName.Replace("’s", "s")
            If ReturnIfNameIsPresent(TempBName) = True Then Return True 'If business name exists in the database, return true
        ElseIf TempBName.Contains("'s") Then
            TempBName = TempBName.Replace("'s", "s")
            If ReturnIfNameIsPresent(TempBName) = True Then Return True 'If business name exists in the database, return true
        End If

        Return False 'Business name is not found in the database, so return false
    End Function
    Private Function ReturnIfNameIsPresent(ByVal BusinessName)
        Dim IsDuplicate As Boolean 'Variable to determine if business exists in the database
        sql.AddParam("@bname", BusinessName) 'Add parameter
        sql.ExecQuery("SELECT * FROM bcontacts WHERE bname LIKE @bname;") 'Execute SQL query
        If sql.RecordCount = 0 Then IsDuplicate = False Else IsDuplicate = True 'If business exists in the database, set variable to true

        If IsDuplicate = True Then Return True 'If business is found in the database, return true

        sql.AddParam("@bname", BusinessName) 'Add parameter
        sql.ExecQuery("SELECT * FROM bcontactsOT WHERE bname LIKE @bname;") 'Execute SQL query
        If sql.RecordCount = 0 Then IsDuplicate = False Else IsDuplicate = True 'If business exists in the database, set variable to true

        If IsDuplicate = True Then Return True 'If business is found in the database, return true

        sql.AddParam("@bname", BusinessName) 'Add parameter
        sql.ExecQuery("SELECT * FROM bDuplicateNames WHERE bnames LIKE @bname;") 'Execute SQL query
        If sql.RecordCount = 0 Then IsDuplicate = False Else IsDuplicate = True 'If business exists in the database, set variable to true

        If IsDuplicate = True Then Return True Else Return False 'If business is found in the database, return true
    End Function
    Private Function CheckDuplicateWebsite(ByVal BusinessWebsite As String)
        If BusinessWebsite = "" Then Return True 'If business website is empty, return true

        Dim IsDuplicate As Boolean 'Variable to determine if website already exists in the database

        sql.AddParam("@bweb", BusinessWebsite.ToLower) 'Add parameter
        sql.ExecQuery("SELECT * FROM bDuplicateWebsites WHERE bwebsites LIKE @bweb;") 'Execute SQL query
        If sql.RecordCount = 0 Then IsDuplicate = False Else IsDuplicate = True 'If website exists in the database, set boolean to true

        If IsDuplicate = True Then Return True 'If website exists in the database, return true. If not, continue

        sql.AddParam("@bweb", BusinessWebsite.ToLower) 'Add parameter
        sql.ExecQuery("SELECT * FROM bcontactsOT WHERE bwebsite LIKE @bweb;") 'Execute SQL query
        If sql.RecordCount = 0 Then IsDuplicate = False Else IsDuplicate = True 'If website exists in the database, set boolean to true

        If IsDuplicate = True Then Return True 'If website exists in the database, return true. If not, continue

        sql.AddParam("@bweb", BusinessWebsite.ToLower) 'Add parameter
        sql.ExecQuery("SELECT * FROM bcontacts WHERE bwebsite LIKE @bweb;") 'Execute SQL query
        If sql.RecordCount = 0 Then IsDuplicate = False Else IsDuplicate = True 'If website exists in the database, set boolean to true

        If IsDuplicate = True Then Return True Else Return False 'If website exists in the database, return true
    End Function
    Private Function WaitForElement(ByVal element As String, ByVal elementMechanism As String, ByVal SecondsToWait As Integer)
        Dim TimeOut As Integer = 0 'Set timeout integer
        Do
            If TimeOut > SecondsToWait Then Return False 'If TimeOut is > than 15, element is not loaded

            If elementMechanism = "ID" Then If Not driver2.FindElements(By.Id(element)).Count = 0 Then Return True 'If there are more than 1 element of type "ID", return True
            If elementMechanism = "Class" Then If Not driver2.FindElements(By.ClassName(element)).Count = 0 Then Return True 'If there are more than 1 element of type "Class", return True
            If elementMechanism = "xPath" Then If Not driver2.FindElements(By.XPath(element)).Count = 0 Then Return True 'If there are more than 1 element of type "xPath", return True

            Thread.Sleep(1000) : TimeOut += 1 'Sleep for 1000ms
        Loop
    End Function
    Private Function GetStringBeforeOrAfter(ByVal TString As String, ByVal TSeparator As String, ByVal TLeft As Boolean, ByVal TRight As Boolean)
        Dim StringOnTheRight As String, StringOnTheLeft As String 'Specify string on the right and string on the left
        Try
            Dim original As String = TString : Dim cut_at As String = TSeparator : Dim stringSeparators() As String = {cut_at} : Dim split = original.Split(stringSeparators, 2, StringSplitOptions.RemoveEmptyEntries)
            StringOnTheRight = split(1)
            StringOnTheLeft = split(0)
        Catch ex As Exception
            Return False
        End Try

        If TLeft = True Then Return StringOnTheLeft 'Return string from the left

        If TRight = True Then Return StringOnTheRight 'Return string from the right

        If TLeft = False And TRight = False Then Return False 'Return false in case user did not specify right or left

        Return False
    End Function
    Private Function ReturnMatchedURLS(ByVal outerHTML As String, ByVal RequestMoreDetails As Boolean)
        Dim oHTML As String = outerHTML 'Get HTML of the page
        Dim MappedUrlList As Boolean = RequestMoreDetails

        Dim LinkURLList As New ArrayList, LinksRTB As New RichTextBox

        Dim links As MatchCollection = Regex.Matches(oHTML, "<a.*?href=""(.*?)"".*?>(.*?)</a>")
        For Each match As Match In links
            Dim matchUrl As String = match.Groups(1).Value
            If matchUrl.StartsWith("#") Then Continue For 'Ignore all anchor links
            If matchUrl.ToLower.StartsWith("javascript:") Then Continue For 'Ignore all javascript calls
            If matchUrl.ToLower.StartsWith("mailto:") Then Continue For 'Ignore all email links
            If Not matchUrl.StartsWith("http://") And Not matchUrl.StartsWith("https://") Then matchUrl = MapUrl(driver2.Url.ToString, matchUrl).Replace("&amp;", "&") 'For internal links, build the url mapped to the base address
            LinkURLList.Add(matchUrl) 'Add URL to the list
            LinksRTB.AppendText(matchUrl & " - " & match.Groups(2).Value.Replace("&amp;", "&") & Environment.NewLine) 'Add URL along with link text of the URL
        Next

        If MappedUrlList = False Then If LinkURLList.Count > 0 Then Return LinkURLList.Item(0) Else Return Nothing Else If Not LinksRTB.Text = "" Then Return LinksRTB.Text Else Return Nothing
    End Function
    Public Function FormatWebsite(ByVal BusinessWebsite As String)
        BusinessWebsite = BusinessWebsite.Replace("https://www.", "").Replace("http://www.", "").Replace("https://", "").Replace("http://", "").Replace("www.", "") 'Replace unnecesary values from string
        Dim TempBusinessWebsite As String = BusinessWebsite
        If TempBusinessWebsite.Contains("/") Then BusinessWebsite = TempBusinessWebsite.Substring(0, TempBusinessWebsite.IndexOf("/")) 'Format string
        BusinessWebsite = BusinessWebsite.Replace("/", "")

        Return BusinessWebsite 'Return formated string
    End Function


#End Region

End Class
