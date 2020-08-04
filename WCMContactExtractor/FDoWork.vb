Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Threading
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium.Remote
Imports OpenQA.Selenium.Support.UI
Public Class FDoWork
    Dim driver As IWebDriver
    Dim driver2 As IWebDriver

    Dim BusinessEmail As String
    Dim NewBusinessEmail As String
    Dim BwReportProgres As String = ""
    Dim NewBwReportProgres As String = ""

    ReadOnly WCMNamesList As New ArrayList
    Dim WCMNamesIndex As Integer = -1
    Dim WCMTown As String

    ReadOnly WCMAddressList As New ArrayList

    ReadOnly FilterKeywords As New ListBox

    Dim Email1Entered As Boolean = False
    Dim Email2Entered As Boolean = False
    Dim Email1TEXT As String = ""
    Dim Email2TEXT As String = ""

    Dim ContactName As String
    Dim ContactNameForEmail As String
    Dim ContactPhone As String
    Dim ContactWebsite As String
    Dim ContactTwitter As String
    Dim ContactInstagram As String
    Dim ContactFacebook As String
    Dim ContactLinkedIn As String

    Dim EmailThread As Thread

    Dim LastKnownElementString As String = ""
    Dim NumberOfOccurences As Integer = 0

    Dim ClipboardContainer As String
    Dim EmailFound As Boolean = False

    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

    Private Sub FDoWork_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False

        InitializeChromeOptions()
        AddEmailFilterKeywords()

        For Each item As String In Form1.RichTextBox1.Lines
            If Not item = "" Then
                If item.Contains("%") Then
                    Dim original As String = item
                    Dim cut_at As String = "%"

                    Dim stringSeparators() As String = {cut_at}
                    Dim split = original.Split(stringSeparators, 2, StringSplitOptions.RemoveEmptyEntries)

                    Dim string_before As String = split(0)

                    Dim FinalValue As String = string_before.Replace(vbTab, "").ToLower
                    Dim FinalString As String = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(FinalValue)

                    WCMNamesList.Add(FinalString)
                Else
                    Dim FinalValue As String = item.Replace(vbTab, "")
                    Dim FinalString As String = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(FinalValue.ToLower)
                    WCMNamesList.Add(FinalString)
                End If

            End If

        Next

        WCMTown = Form1.ComboBox1.SelectedItem.ToString

        For Each item As String In Form1.RichTextBox2.Lines
            WCMAddressList.Add(item)
        Next

        driver.Navigate.GoToUrl("https://crm.zoho.com/crm/org19427096/tab/Contacts/create")

        If CheckBox2.Checked = True Then TTabHandler.Start()

        If CheckBox3.Checked = True Then TDrawHandler.Start()
    End Sub

    Private Sub AddEmailFilterKeywords()
        FilterKeywords.Items.Add("johndoe")
        FilterKeywords.Items.Add("domain")
        FilterKeywords.Items.Add("name")
        FilterKeywords.Items.Add("jpg")
        FilterKeywords.Items.Add("png")
        FilterKeywords.Items.Add("jpeg")
        FilterKeywords.Items.Add("img")
        FilterKeywords.Items.Add("ico")
        FilterKeywords.Items.Add("png")
        FilterKeywords.Items.Add(".js")
        FilterKeywords.Items.Add("example")
        FilterKeywords.Items.Add("godaddy")
        FilterKeywords.Items.Add("email.com")
        FilterKeywords.Items.Add("placeholder")
        FilterKeywords.Items.Add("someone.com")
        FilterKeywords.Items.Add("wix.com")
        FilterKeywords.Items.Add("wix")
        FilterKeywords.Items.Add("email.com")
        FilterKeywords.Items.Add("abuse")
        FilterKeywords.Items.Add("spam")
        FilterKeywords.Items.Add("..")
        FilterKeywords.Items.Add("broofa.com")
    End Sub

    Private Sub InitializeChromeOptions()
        Dim optionOn As New ChromeOptions

        Dim driverService = ChromeDriverService.CreateDefaultService()

        driverService.HideCommandPromptWindow = True

        optionOn.AddUserProfilePreference("profile.default_content_setting_values.geolocation", 2)


        optionOn.AddArgument("start-maximized")
        optionOn.AddArgument("--disable-infobars")



        Dim theContDir As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim newDir As String = "WCMChromeCookies"

        Dim thePath As String = IO.Path.Combine(theContDir, newDir)

        If Not IO.Directory.Exists(thePath) Then
            Try
                IO.Directory.CreateDirectory(thePath)
            Catch UAex As UnauthorizedAccessException
                ' exception
            End Try
        End If
        Dim tdir As String = thePath

        optionOn.AddArgument("user-data-dir=" + tdir + "/chrome-session")
        optionOn.AddArgument("--profile-directory=Default")
        optionOn.AddArgument("--disable-application-cache")

        Try
            driver = New ChromeDriver(driverService, optionOn)
        Catch ex As Exception
            Application.Exit()
        End Try


        'Initiate another Chrome web driver for email hunting
        Dim driverService2 = ChromeDriverService.CreateDefaultService()
        driverService2.HideCommandPromptWindow = True
        Dim optionOn2 As New ChromeOptions
        optionOn2.AddUserProfilePreference("profile.default_content_setting_values.images", 2) 'Disable or enable images
        optionOn2.AddArgument("--blink-settings=imagesEnabled=false") 'Disable or enable images
        optionOn2.AddArguments("headless")
        driver2 = New ChromeDriver(driverService2, optionOn2)

    End Sub

    Private Sub FDoWork_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            driver.Dispose()
            driver2.Dispose()
            driver.Quit()
            driver2.Quit()
        Catch ex As Exception

        End Try

        Form1.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button1.Enabled = False
        WCMNamesIndex += 1

        '  TTabHandler.Stop()

        Label14.Text = WCMNamesIndex & "/" & WCMNamesList.Count
        Try
            Label4.Text = WCMAddressList.Item(WCMNamesIndex)

            ListBox1.Items.Clear()
            ListBox2.Items.Clear()
            ListBox1.BackColor = Color.White
            ListBox2.BackColor = Color.White

            ContactName = ""
            ContactNameForEmail = ""
            ContactPhone = ""
            ContactWebsite = ""
            ContactTwitter = ""
            ContactInstagram = ""
            ContactFacebook = ""
            ContactLinkedIn = ""

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""

            BWLoadContact.RunWorkerAsync()
        Catch ex As Exception
            MsgBox("Done!")
            Application.Exit()
        End Try
    End Sub

    Private Sub BWLoadContact_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWLoadContact.DoWork
        Email1Entered = False : Email2Entered = False : Email1TEXT = "" : Email2TEXT = ""

        Dim element As IWebElement
        Dim js As IJavaScriptExecutor = driver

        Thread.Sleep(500)

        Try
            driver.SwitchTo().Window(driver.WindowHandles(1))
            driver.Navigate.GoToUrl("https://google.com")
        Catch ex As Exception
            driver.SwitchTo().Window(driver.WindowHandles.Last)
            js.ExecuteScript("window.open();")
            driver.SwitchTo().Window(driver.WindowHandles(1))

            driver.Navigate.GoToUrl("https://google.com")
        End Try

        Try
            element = driver.FindElement(By.LinkText("English"))
            element.Click()
        Catch ex As Exception
        End Try

        Try
            element = driver.FindElement(By.XPath("//*[@id='tsf']/div[2]/div/div[1]/div/div[1]/input"))
            element.SendKeys(WCMNamesList.Item(WCMNamesIndex) & " " & WCMTown & " CT")
            Thread.Sleep(500)
            element.SendKeys(Keys.Enter)
        Catch ex As Exception
            element = driver.FindElement(By.XPath("//*[@id='tsf']/div[2]/div[1]/div[1]/div/div[2]/input"))
            element.SendKeys(WCMNamesList.Item(WCMNamesIndex) & " " & WCMTown & " CT")
            Thread.Sleep(500)
            element.SendKeys(Keys.Enter)
        End Try


        ContactName = WCMNamesList.Item(WCMNamesIndex)
        TextBox1.Text = ContactName
        TextBox4.Text = WCMTown

        Dim SideBarExists As Boolean = False
        Try
            element = driver.FindElement(By.XPath("//*[@id='rhs']"))
            SideBarExists = True
        Catch ex As Exception

        End Try

        If SideBarExists = True Then
            Dim SidebarOuterHTML As String
            SidebarOuterHTML = element.GetAttribute("outerHTML")

            Dim TempContactName As String = ExtractData(SidebarOuterHTML, "data-attrid=""title""", "</div>")

            ContactNameForEmail = StringBetween(TempContactName, "<span>", "</span>")
            TextBox8.Text = ContactNameForEmail

            ContactPhone = ExtractData(SidebarOuterHTML, "<span>+1 ", "</span>")
            TextBox2.Text = ContactPhone

            Dim SidebarLinks As MatchCollection = Regex.Matches(SidebarOuterHTML, "<a.*?href=""(.*?)"".*?>(.*?)</a>")

            For Each match As Match In SidebarLinks
                Dim matchUrl As String = match.Groups(1).Value
                'Ignore all anchor links
                If matchUrl.StartsWith("#") Then
                    Continue For
                End If
                'Ignore all javascript calls
                If matchUrl.ToLower.StartsWith("javascript:") Then
                    Continue For
                End If
                'Ignore all email links
                If matchUrl.ToLower.StartsWith("mailto:") Then
                    Continue For
                End If

                If match.Groups(2).Value = "Website" Then ContactWebsite = matchUrl : TextBox3.Text = ContactWebsite
                If matchUrl.Contains("facebook.com") Then ContactFacebook = matchUrl : TextBox5.Text = ContactFacebook
                If matchUrl.Contains("twitter.com") Then ContactTwitter = matchUrl : TextBox6.Text = ContactTwitter
                If matchUrl.Contains("instagram.com") Then ContactInstagram = matchUrl : TextBox9.Text = ContactInstagram
                If matchUrl.Contains("linkedin.com") Then ContactLinkedIn = matchUrl : TextBox7.Text = ContactLinkedIn
            Next

            Dim MainBodyOuterHtml As String
            element = driver.FindElement(By.XPath("//*[@id='rso']"))
            MainBodyOuterHtml = element.GetAttribute("outerHTML")
            Dim MainBodyLinks As MatchCollection = Regex.Matches(MainBodyOuterHtml, "<a.*?href=""(.*?)"".*?>(.*?)</a>")
            For Each match As Match In MainBodyLinks
                Dim matchUrl As String = match.Groups(1).Value
                'Ignore all anchor links
                If matchUrl.StartsWith("#") Then
                    Continue For
                End If
                'Ignore all javascript calls
                If matchUrl.ToLower.StartsWith("javascript:") Then
                    Continue For
                End If
                'Ignore all email links
                If matchUrl.ToLower.StartsWith("mailto:") Then
                    Continue For
                End If

                If matchUrl.Contains("facebook.com") Then If ContactFacebook = "" Then ContactFacebook = matchUrl : TextBox5.Text = ContactFacebook
                If matchUrl.Contains("twitter.com") Then If ContactTwitter = "" Then ContactTwitter = matchUrl : TextBox6.Text = ContactTwitter
                If matchUrl.Contains("instagram.com") Then If ContactInstagram = "" Then ContactInstagram = matchUrl : TextBox9.Text = ContactInstagram
                If matchUrl.Contains("linkedin.com") Then If ContactLinkedIn = "" Then ContactLinkedIn = matchUrl : TextBox7.Text = ContactLinkedIn
            Next



            If Not TextBox3.Text = "" Then
                ContactWebsite = TextBox3.Text
                LEmailThread.Text = "Searching for emails..."
                LEmailThread.ForeColor = Color.Red

                EmailThread = New Thread(AddressOf CheckForEmails)
                EmailThread.Start()

                Try
                    driver.SwitchTo().Window(driver.WindowHandles(2))
                Catch ex As Exception
                    driver.SwitchTo().Window(driver.WindowHandles.Last)
                    js.ExecuteScript("window.open();")
                    driver.SwitchTo().Window(driver.WindowHandles(2))
                End Try

                driver.Navigate.GoToUrl(ContactWebsite)

                Dim AdditionalLinks As MatchCollection = Regex.Matches(driver.PageSource, "<a.*?href=""(.*?)"".*?>(.*?)</a>")
                For Each match As Match In AdditionalLinks
                    Dim matchUrl As String = match.Groups(1).Value
                    'Ignore all anchor links
                    If matchUrl.StartsWith("#") Then
                        Continue For
                    End If
                    'Ignore all javascript calls
                    If matchUrl.ToLower.StartsWith("javascript:") Then
                        Continue For
                    End If
                    'Ignore all email links
                    If matchUrl.ToLower.StartsWith("mailto:") Then
                        Continue For
                    End If

                    If matchUrl.Contains("facebook.com") Then
                        If ContactFacebook = "" Then
                            If FacebookBannedWords(matchUrl) = False Then ContactFacebook = matchUrl : TextBox5.Text = ContactFacebook
                        End If
                    End If
                    If matchUrl.Contains("twitter.com") Then
                        If ContactTwitter = "" Then
                            If TwitterBannedWords(matchUrl) = False Then ContactTwitter = matchUrl : TextBox6.Text = ContactTwitter
                        End If
                    End If
                    If matchUrl.Contains("instagram.com") Then
                        If ContactInstagram = "" Then
                            If InstagramBannedWords(matchUrl) = False Then ContactInstagram = matchUrl : TextBox9.Text = ContactInstagram
                        End If
                    End If
                    If matchUrl.Contains("linkedin.com") Then
                        If ContactLinkedIn = "" Then
                            If LinkedInBannedWords(matchUrl) = False Then ContactLinkedIn = matchUrl : TextBox7.Text = ContactLinkedIn
                        End If
                    End If
                Next

            Else
                Try
                    driver.SwitchTo().Window(driver.WindowHandles(2))
                Catch ex As Exception
                    driver.SwitchTo().Window(driver.WindowHandles.Last)
                    js.ExecuteScript("window.open();")
                    driver.SwitchTo().Window(driver.WindowHandles(2))
                End Try

                If Not driver.Url.Contains("bing.com") Then driver.Navigate.GoToUrl("https://bing.com")

            End If


            'CTDATA SEARCH
            If CheckBox1.Checked = True Then

                Try
                    Try
                        driver.SwitchTo().Window(driver.WindowHandles(3))
                    Catch ex As Exception
                        driver.SwitchTo().Window(driver.WindowHandles.Last)
                        js.ExecuteScript("window.open();")
                        driver.SwitchTo().Window(driver.WindowHandles(3))
                        driver.Navigate.GoToUrl("http://searchctbusiness.ctdata.org/search_results?query=aaaaaaaa&query_limit=&index_field=business_name&start_date=1900-01-01&end_date=2019-02-01&active=&sort_by=nm_name&sort_order=asc&page=1")
                    End Try

                    If Not ContactWebsite = "" Then
                        Try
                            element = driver.FindElement(By.XPath("//*[@id='query']"))
                            element.SendKeys(Keys.Control + Keys.Backspace)
                        Catch ex As Exception
                            driver.Navigate.GoToUrl("http://searchctbusiness.ctdata.org/search_results?query=aaaaaaaa&query_limit=&index_field=business_name&start_date=1900-01-01&end_date=2019-02-01&active=&sort_by=nm_name&sort_order=asc&page=1")
                        End Try


                        Try
                            element = driver.FindElement(By.XPath("//*[@id='query']")) : element.Clear()

                            Thread.Sleep(500)
                            element.SendKeys(TextBox1.Text)

                            element = driver.FindElement(By.XPath("//*[@id='sidebar']/form/div/div/div[6]/span/button/i"))
                            Thread.Sleep(500)
                            element.Click()
                        Catch ex As Exception

                        End Try



                        Try
                            element = driver.FindElement(By.XPath("//*[@id='main']/div/div[2]/table/tbody/tr[1]/td[1]/a"))
                        Catch ex As Exception
                            Dim bb As String = TextBox1.Text

                            element = driver.FindElement(By.XPath("//*[@id='query']"))
                            element.SendKeys(Keys.Control + Keys.Backspace)

                            element = driver.FindElement(By.XPath("//*[@id='sidebar']/form/div/div/div[6]/span/button/i"))
                            Thread.Sleep(500)
                            element.Click()
                        End Try
                    End If

                Catch ex As Exception
                End Try
            End If
            'CTDATA SEARCH





            driver.SwitchTo().Window(driver.WindowHandles(0))

            'CLEARING DATA
            If WaitForElement("//*[@id='Crm_Contacts_ACCOUNTID']", "xPath") = True Then
                Try
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ACCOUNTID']")) : element.Clear() 'Account name
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF13']")) : element.Clear()  'Account name for email
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF122']")) : element.Clear() 'Website
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF48']")) : element.Clear() 'Facebook
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF155']")) : element.Clear() 'Instagram
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF156']")) : element.Clear() 'Twitter
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF30']")) : element.Clear() 'LinkedIn
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_PHONE']")) : element.Clear() 'Phone
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_FIRSTNAME']")) : element.Clear() 'First name
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : element.Clear() 'Dear field
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : element.Clear() 'Last name
                Catch ex As Exception
                    driver.Navigate.GoToUrl("https://crm.zoho.com/crm/org19427096/tab/Contacts/create")

                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ACCOUNTID']")) : element.Clear() 'Account name
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF13']")) : element.Clear()  'Account name for email
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF122']")) : element.Clear() 'Website
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF48']")) : element.Clear() 'Facebook
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF155']")) : element.Clear() 'Instagram
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF156']")) : element.Clear() 'Twitter
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF30']")) : element.Clear() 'LinkedIn
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_PHONE']")) : element.Clear() 'Phone
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_FIRSTNAME']")) : element.Clear() 'First name
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : element.Clear() 'Dear field
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : element.Clear() 'Last name
                End Try
            End If




            'element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF151']")) : element.Clear() 'Social tags
            'CLEARING DATA



            'LOAD DATA
            ''''''''''''remove this later
            'element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_DESCRIPTION']")) : element.Clear()
            'element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_DESCRIPTION']")) : element.SendKeys("WCM NEWTOWN BUSINESS")
            ''' 


            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, "Hello")
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, "Hello")

            ' element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : element.SendKeys("Hello")
            'element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : element.SendKeys("Hello")

            TextBox8.Text = TextBox8.Text.Replace(" Co Inc", "").Replace(" Co. Inc.", "").Replace(", LLC", "").Replace(", Inc.", "").Replace(" Inc.", "").Replace(" LLC", "").Replace(" INC", "").Replace(" llc", "").Replace(" Inc", "").Replace(" LLC.", "").Replace(",Inc.", "").Replace(" Inc", "").Replace(" L.L.C.", "")

            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ACCOUNTID']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox1.Text) 'Account name
            'element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ACCOUNTID']")) : element.SendKeys(TextBox1.Text) 'Account name

            If TextBox8.Text = "" Then TextBox8.Text = TextBox1.Text
            ' element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF13']")) : element.SendKeys(TextBox8.Text) 'Account name for email
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF13']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox8.Text) 'Account name for email

            If Not TextBox3.Text = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF122']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox3.Text) 'Contact website
            'If Not TextBox3.Text = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF122']")) : element.SendKeys(TextBox3.Text) 'Contact website

            'If Not ContactFacebook = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF48']")) : element.SendKeys(TextBox5.Text)
            'If Not ContactInstagram = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF155']")) : element.SendKeys(TextBox9.Text)
            'If Not ContactTwitter = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF156']")) : element.SendKeys(TextBox6.Text)
            'If Not ContactLinkedIn = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF30']")) : element.SendKeys(TextBox7.Text)
            If Not ContactFacebook = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF48']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox5.Text)
            If Not ContactInstagram = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF155']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox9.Text)
            If Not ContactTwitter = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF156']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox6.Text)
            If Not ContactLinkedIn = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF30']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox7.Text)


            'element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_PHONE']")) : element.SendKeys(TextBox2.Text) 'Phone
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_PHONE']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox2.Text) 'Phone


            element = driver.FindElement(By.XPath("//*[@id='Contacts_fldRow_CONTACTCF157']/div[2]/div/span/span[1]/span/ul/li/input")) 'WCM Town
            element.SendKeys(Keys.Backspace)

            Thread.Sleep(500)

            element = driver.FindElement(By.XPath("//*[@id='Contacts_fldRow_CONTACTCF157']/div[2]/div/span/span[1]/span/ul/li/input"))
            element.SendKeys(WCMTown + Keys.Enter)


            Thread.Sleep(500)
            element = driver.FindElement(By.ClassName("crmBodyWin"))
            element.Click()
            element.SendKeys(Keys.PageUp)
            element.SendKeys(Keys.PageUp)

            driver.SwitchTo().Window(driver.WindowHandles(2))
        Else
            ' Button2.Enabled = True
        End If

        TTabHandler.Start()
        TBackgroundCheck.Start()
        Button1.Enabled = True
    End Sub

    Private Function WaitForElement(ByVal element As String, ByVal elementMechanism As String)
        Dim TimeOut As Integer = 0
        Do
            If TimeOut > 15 Then Return False

            If elementMechanism = "ID" Then If Not driver.FindElements(By.Id(element)).Count = 0 Then Return True
            If elementMechanism = "Class" Then If Not driver.FindElements(By.ClassName(element)).Count = 0 Then Return True
            If elementMechanism = "xPath" Then If Not driver.FindElements(By.XPath(element)).Count = 0 Then Return True

            Thread.Sleep(1000) : TimeOut += 1
        Loop
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '  Button2.Enabled = False
        '  Button1.Enabled = False
        ' BWAnalyzeReport.RunWorkerAsync()
    End Sub

    Private Sub BWAnalyzeReport_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWAnalyzeReport.DoWork
        ContactName = WCMNamesList.Item(WCMNamesIndex)
        ContactNameForEmail = WCMNamesList.Item(WCMNamesIndex)

        Dim element As IWebElement
        Dim js As IJavaScriptExecutor = driver

        If Not TextBox3.Text = "" Then

            ContactWebsite = TextBox3.Text

            LEmailThread.Text = "Searching for emails..."
            LEmailThread.ForeColor = Color.Red

            EmailThread = New Thread(AddressOf CheckForEmails)
            EmailThread.Start()

            Try
                driver.SwitchTo().Window(driver.WindowHandles(2))
            Catch ex As Exception
                driver.SwitchTo().Window(driver.WindowHandles.Last)
                js.ExecuteScript("window.open();")
                driver.SwitchTo().Window(driver.WindowHandles(2))
            End Try

            driver.Navigate.GoToUrl(ContactWebsite)

            Dim Additionallinks As MatchCollection = Regex.Matches(driver.PageSource, "<a.*?href=""(.*?)"".*?>(.*?)</a>")
            For Each match As Match In Additionallinks
                Dim matchUrl As String = match.Groups(1).Value
                'Ignore all anchor links
                If matchUrl.StartsWith("#") Then
                    Continue For
                End If
                'Ignore all javascript calls
                If matchUrl.ToLower.StartsWith("javascript:") Then
                    Continue For
                End If
                'Ignore all email links
                If matchUrl.ToLower.StartsWith("mailto:") Then
                    Continue For
                End If

                If matchUrl.Contains("facebook.com") Then
                    If ContactFacebook = "" Then
                        If FacebookBannedWords(matchUrl) = False Then ContactFacebook = matchUrl : TextBox5.Text = ContactFacebook
                    End If
                End If
                If matchUrl.Contains("twitter.com") Then
                    If ContactTwitter = "" Then
                        If TwitterBannedWords(matchUrl) = False Then ContactTwitter = matchUrl : TextBox6.Text = ContactTwitter
                    End If
                End If
                If matchUrl.Contains("instagram.com") Then
                    If ContactInstagram = "" Then
                        If InstagramBannedWords(matchUrl) = False Then ContactInstagram = matchUrl : TextBox9.Text = ContactInstagram
                    End If
                End If
                If matchUrl.Contains("linkedin.com") Then
                    If ContactInstagram = "" Then
                        If LinkedInBannedWords(matchUrl) = False Then ContactLinkedIn = matchUrl : TextBox7.Text = ContactLinkedIn
                    End If
                End If
            Next

        Else
            driver.SwitchTo().Window(driver.WindowHandles(2))

            If Not driver.Url.Contains("bing.com") Then
                driver.Navigate.GoToUrl("https://bing.com")
            End If
        End If



        'CTDATA SEARCH
        If CheckBox1.Checked = True Then

            Try
                Try
                    driver.SwitchTo().Window(driver.WindowHandles(3))
                Catch ex As Exception
                    driver.SwitchTo().Window(driver.WindowHandles.Last)
                    js.ExecuteScript("window.open();")
                    driver.SwitchTo().Window(driver.WindowHandles(3))
                    driver.Navigate.GoToUrl("http://searchctbusiness.ctdata.org/search_results?query=aaaaaaaa&query_limit=&index_field=business_name&start_date=1900-01-01&end_date=2019-02-01&active=&sort_by=nm_name&sort_order=asc&page=1")
                End Try

                If Not ContactWebsite = "" Then
                    Try
                        element = driver.FindElement(By.XPath("//*[@id='query']"))
                        element.SendKeys(Keys.Control + Keys.Backspace)
                    Catch ex As Exception
                        driver.Navigate.GoToUrl("http://searchctbusiness.ctdata.org/search_results?query=aaaaaaaa&query_limit=&index_field=business_name&start_date=1900-01-01&end_date=2019-02-01&active=&sort_by=nm_name&sort_order=asc&page=1")
                    End Try


                    Try
                        element = driver.FindElement(By.XPath("//*[@id='query']"))
                        element.Clear()

                        Thread.Sleep(500)
                        element.SendKeys(TextBox1.Text)

                        element = driver.FindElement(By.XPath("//*[@id='sidebar']/form/div/div/div[6]/span/button/i"))
                        Thread.Sleep(500)
                        element.Click()
                    Catch ex As Exception

                    End Try



                    Try
                        element = driver.FindElement(By.XPath("//*[@id='main']/div/div[2]/table/tbody/tr[1]/td[1]/a"))
                    Catch ex As Exception
                        Dim bb As String = TextBox1.Text

                        element = driver.FindElement(By.XPath("//*[@id='query']"))
                        element.SendKeys(Keys.Control + Keys.Backspace)

                        element = driver.FindElement(By.XPath("//*[@id='sidebar']/form/div/div/div[6]/span/button/i"))
                        Thread.Sleep(500)
                        element.Click()
                    End Try
                End If

            Catch ex As Exception
            End Try
        End If
        'CTDATA SEARCH



        driver.SwitchTo().Window(driver.WindowHandles(0))

        'CLEARING DATA
        Try
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ACCOUNTID']")) : element.Clear() 'Account name
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF13']")) : element.Clear()  'Account name for email
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF122']")) : element.Clear() 'Website
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF48']")) : element.Clear() 'Facebook
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF155']")) : element.Clear() 'Instagram
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF156']")) : element.Clear() 'Twitter
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF30']")) : element.Clear() 'LinkedIn
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_PHONE']")) : element.Clear() 'Phone
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_FIRSTNAME']")) : element.Clear() 'First name
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : element.Clear() 'Dear field
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : element.Clear() 'Last name
        Catch ex As Exception
            driver.Navigate.GoToUrl("https://crm.zoho.com/crm/org19427096/tab/Contacts/create")

            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ACCOUNTID']")) : element.Clear() 'Account name
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF13']")) : element.Clear()  'Account name for email
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF122']")) : element.Clear() 'Website
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF48']")) : element.Clear() 'Facebook
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF155']")) : element.Clear() 'Instagram
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF156']")) : element.Clear() 'Twitter
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF30']")) : element.Clear() 'LinkedIn
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_PHONE']")) : element.Clear() 'Phone
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_FIRSTNAME']")) : element.Clear() 'First name
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : element.Clear() 'Dear field
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : element.Clear() 'Last name
        End Try

        element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_EMAIL']")) : element.Clear() 'Email
        element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ADDN_EMAIL']")) : element.Clear() 'Secondary email
        'CLEARING DATA


        'LOAD DATA
        js.ExecuteScript("arguments[0].value = arguments[1]", element, "Hello")


        element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, "Hello")
        element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, "Hello")


        ' element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : element.SendKeys("Hello")
        ' element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : element.SendKeys("Hello")

        TextBox8.Text = TextBox8.Text.Replace(" Co Inc", "").Replace(" Co. Inc.", "").Replace(", LLC", "").Replace(", Inc.", "").Replace(" Inc.", "").Replace(" LLC", "").Replace(" INC", "").Replace(" llc", "").Replace(" Inc", "").Replace(" LLC.", "").Replace(",Inc.", "").Replace(" Inc", "").Replace(" L.L.C.", "")

        element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ACCOUNTID']")) : element.SendKeys(TextBox1.Text) 'Account name

        If TextBox8.Text = "" Then TextBox8.Text = TextBox1.Text
        element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF13']")) : element.SendKeys(TextBox8.Text) 'Account name for email

        If Not TextBox3.Text = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF122']")) : element.SendKeys(TextBox3.Text) 'Contact website

        If Not ContactFacebook = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF48']")) : element.SendKeys(TextBox5.Text)
        If Not ContactInstagram = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF155']")) : element.SendKeys(TextBox9.Text)
        If Not ContactTwitter = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF156']")) : element.SendKeys(TextBox6.Text)
        If Not ContactLinkedIn = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF30']")) : element.SendKeys(TextBox7.Text)

        element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_PHONE']")) : element.SendKeys(TextBox2.Text) 'Phone

        element = driver.FindElement(By.XPath("//*[@id='Contacts_fldRow_CONTACTCF157']/div[2]/div/span/span[1]/span/ul/li/input")) : element.SendKeys(Keys.Backspace) 'WCM Town

        Thread.Sleep(500)

        element = driver.FindElement(By.XPath("//*[@id='Contacts_fldRow_CONTACTCF157']/div[2]/div/span/span[1]/span/ul/li/input")) : element.SendKeys(WCMTown + Keys.Enter)

        Thread.Sleep(500)
        element = driver.FindElement(By.ClassName("crmBodyWin"))
        element.Click()
        element.SendKeys(Keys.PageUp) : element.SendKeys(Keys.PageUp)

        driver.SwitchTo().Window(driver.WindowHandles(1))

        ' Button2.Enabled = False
        Button1.Enabled = True
    End Sub

#Region "Email SCRAPING"
    Private Sub CheckForEmails()
        ListBox1.Items.Clear() : ListBox2.Items.Clear()

        Try
            driver2.Navigate.GoToUrl(ContactWebsite)
        Catch ex As Exception
        End Try

        Try
            If Not ContactWebsite.ToString = "" Then
                Dim tempTB As String = ContactWebsite.ToString
                tempTB = tempTB.Replace("https://www.", "").Replace("http://www.", "").Replace("https://", "").Replace("http://", "").Replace("www.", "")
            Else
                Exit Sub
            End If

        Catch ex As Exception
            Exit Sub
        End Try




        Dim element As IWebElement

        Dim PageText As New RichTextBox

        Try
            element = driver2.FindElement(By.XPath("html"))
            PageText.Text = element.Text
        Catch ex As Exception
            PageText.Text = ""
        End Try


        PageText.AppendText(Environment.NewLine)
        PageText.AppendText(driver2.PageSource)

        Dim emails As New List(Of String)

        Dim Newemails As New List(Of String)

        Dim adrRx As Regex = New Regex("([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})")
        For Each ii As Match In adrRx.Matches(PageText.Text.ToLower)
            Dim CheckForInvalidDomains As Boolean = FilterResults(ii.Value)
            If Not CheckForInvalidDomains = True Then emails.Add(ii.Value)
        Next


        Dim SecondTryFound As Boolean = False

        Dim emailsString As String

        Dim TextLinks As New ListBox
        TextLinks.Items.Add("Contact")
        TextLinks.Items.Add("CONTACT")
        TextLinks.Items.Add("Feedback")
        TextLinks.Items.Add("FEEDBACK")
        TextLinks.Items.Add("REQUEST")
        TextLinks.Items.Add("About")
        TextLinks.Items.Add("ABOUT")
        TextLinks.Items.Add("Get in touch")
        TextLinks.Items.Add("GET IN TOUCH")
        TextLinks.Items.Add("Staff")
        TextLinks.Items.Add("STAFF")
        TextLinks.Items.Add("Leadership")
        TextLinks.Items.Add("LEADERSHIP")

        For Each item As String In TextLinks.Items
            Try
                element = driver2.FindElement(By.PartialLinkText(item.ToString))

                driver2.Navigate.GoToUrl(element.GetAttribute("href"))

                element = driver2.FindElement(By.XPath("html"))
                PageText.Text = element.Text

                PageText.AppendText(Environment.NewLine)
                PageText.AppendText(driver2.PageSource)


                For Each ii As Match In adrRx.Matches(PageText.Text.ToLower)
                    Dim CheckForInvalidDomains As Boolean = FilterResults(ii.Value)
                    If Not CheckForInvalidDomains = True Then emails.Add(ii.Value)
                Next

                If Not emails.Count = 0 Then
                    emailsString = Join(emails.Distinct.ToArray, "; ")
                    BusinessEmail = emailsString
                    BwReportProgres = BusinessEmail

                    SecondTryFound = True
                End If
            Catch ex As Exception
            End Try

        Next

        Try

            driver2.Navigate.GoToUrl(ContactFacebook)

            Dim CurFBUrl As String = driver2.Url

            CurFBUrl = CurFBUrl.Replace("https://www.facebook.com/", "")
            CurFBUrl = CurFBUrl.Replace("/", "")

            driver2.Navigate.GoToUrl("https://www.facebook.com/pg/" & CurFBUrl & "/about/?ref=page_internal")

            element = driver2.FindElement(By.XPath("html"))

            PageText.AppendText(Environment.NewLine & element.Text)

            For Each ii As Match In adrRx.Matches(PageText.Text.ToLower)
                Dim CheckForInvalidDomains As Boolean = FilterResults(ii.Value)

                If Not CheckForInvalidDomains = True Then
                    If Not emails.Contains(ii.Value) Then
                        emails.Add(ii.Value)
                    End If
                End If
            Next
        Catch ex As Exception

        End Try


        Try
            emailsString = Join(emails.Distinct.ToArray, "; ")
            BusinessEmail = emailsString
            BwReportProgres = BusinessEmail
        Catch ex As Exception

        End Try


        If SecondTryFound = False Then
            Dim CurrentDomain As String = ContactWebsite
            CurrentDomain = CurrentDomain.Replace("https://www.", "")
            CurrentDomain = CurrentDomain.Replace("http://www.", "")
            CurrentDomain = CurrentDomain.Replace("https://", "")
            CurrentDomain = CurrentDomain.Replace("http://", "")
            CurrentDomain = CurrentDomain.Replace("www.", "")

            Dim string_before As String = ""

            Try
                Dim original As String = CurrentDomain
                Dim cut_at As String = "/"

                Dim stringSeparators() As String = {cut_at}
                Dim split = original.Split(stringSeparators, 2, StringSplitOptions.RemoveEmptyEntries)

                string_before = split(0)
            Catch ex As Exception
            End Try


            Try
                driver2.Navigate.GoToUrl("https://www.bing.com/search?q=%22*%40" & string_before & "%22")
                element = driver2.FindElement(By.XPath("html"))
                PageText.Text = element.Text


                driver2.Navigate.GoToUrl("https://www.bing.com/search?q=%22*%40" & string_before & "&first=11&FORM=PERE")
                element = driver2.FindElement(By.XPath("html"))
                PageText.Text &= element.Text
            Catch ex As Exception

            End Try


            Try
                For Each ii As Match In adrRx.Matches(PageText.Text.ToLower)
                    Dim CheckForInvalidDomains As Boolean = FilterResults(ii.Value)

                    If Not CheckForInvalidDomains = True Then
                        If ii.Value.Contains(string_before) Then
                            If Not emails.Contains(ii.Value) Then
                                Newemails.Add(ii.Value)
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
            End Try

        End If

        If emails.Count = 0 Then
            BwReportProgres = BusinessEmail
        Else
            emailsString = Join(emails.Distinct.ToArray, "; ")
            BusinessEmail = emailsString
            BwReportProgres = BusinessEmail
        End If



        If Newemails.Count = 0 Then
            NewBwReportProgres = NewBusinessEmail
        Else
            Dim NewemailsString As String = Join(Newemails.Distinct.ToArray, "; ")
            NewBusinessEmail = NewemailsString
            NewBwReportProgres = NewBusinessEmail
        End If



        Dim TempTB1 As New TextBox
        Dim TempTB2 As New TextBox


        TempTB1.Text = BwReportProgres
        TempTB2.Text = NewBwReportProgres

        TempTB1.Text = TempTB1.Text.Replace("; ", Environment.NewLine)
        TempTB2.Text = TempTB2.Text.Replace("; ", Environment.NewLine)

        For Each line As String In TempTB1.Lines
            If Not line = "" Then
                ListBox1.Items.Add(line)
            End If
        Next
        For Each line As String In TempTB2.Lines
            If Not line = "" Then
                ListBox2.Items.Add(line)
            End If
        Next

        BusinessEmail = ""
        NewBusinessEmail = ""

        driver2.Navigate.GoToUrl("https://www.bing.com")
        LEmailThread.Text = "Thread idle..."
        LEmailThread.ForeColor = Color.Green
    End Sub

    Private Function FacebookBannedWords(ByVal link As String)
        Dim BannedWords As New ListBox
        Dim BannedLinks As New ListBox

        BannedWords.Items.Add("sharer")
        BannedWords.Items.Add("share")
        BannedWords.Items.Add("wix")

        BannedLinks.Items.Add("https://www.facebook.com/")
        BannedLinks.Items.Add("https://www.facebook.com")
        BannedLinks.Items.Add("https//www.facebook.com/")
        BannedLinks.Items.Add("www.facebook.com")
        BannedLinks.Items.Add("www.facebook.com/")

        For Each item As String In BannedWords.Items
            If link.Contains(item) Then Return True
        Next

        For Each item As String In BannedLinks.Items
            If link = item Then Return True
        Next

        Return False
    End Function
    Private Function TwitterBannedWords(ByVal link As String)
        Dim BannedWords As New ListBox
        Dim BannedLinks As New ListBox

        BannedWords.Items.Add("sharer")
        BannedWords.Items.Add("share")
        BannedWords.Items.Add("wix")

        BannedLinks.Items.Add("https://www.twitter.com/")
        BannedLinks.Items.Add("https://www.twitter.com")
        BannedLinks.Items.Add("https//www.twitter.com/")
        BannedLinks.Items.Add("www.twitter.com")
        BannedLinks.Items.Add("www.twitter.com/")

        For Each item As String In BannedWords.Items
            If link.Contains(item) Then Return True
        Next

        For Each item As String In BannedLinks.Items
            If link = item Then Return True
        Next

        Return False
    End Function
    Private Function InstagramBannedWords(ByVal link As String)
        Dim BannedWords As New ListBox
        Dim BannedLinks As New ListBox

        BannedWords.Items.Add("sharer")
        BannedWords.Items.Add("share")
        BannedWords.Items.Add("wix")

        BannedLinks.Items.Add("https://www.instagram.com/")
        BannedLinks.Items.Add("https://www.instagram.com")
        BannedLinks.Items.Add("https//www.instagram.com/")
        BannedLinks.Items.Add("www.instagram.com")
        BannedLinks.Items.Add("www.instagram.com/")

        For Each item As String In BannedWords.Items
            If link.Contains(item) Then Return True
        Next

        For Each item As String In BannedLinks.Items
            If link = item Then Return True
        Next

        Return False
    End Function
    Private Function LinkedInBannedWords(ByVal link As String)
        Dim BannedWords As New ListBox
        Dim BannedLinks As New ListBox

        BannedWords.Items.Add("sharer")
        BannedWords.Items.Add("share")
        BannedWords.Items.Add("wix")

        BannedLinks.Items.Add("https://www.linkedin.com/")
        BannedLinks.Items.Add("https://www.linkedin.com")
        BannedLinks.Items.Add("https//www.linkedin.com/")
        BannedLinks.Items.Add("www.linkedin.com")
        BannedLinks.Items.Add("www.linkedin.com/")

        For Each item As String In BannedWords.Items
            If link.Contains(item) Then Return True
        Next

        For Each item As String In BannedLinks.Items
            If link = item Then Return True
        Next

        Return False
    End Function

    Private Function FilterResults(ByVal ii)
        Dim SubjEmail As String = ii

        For Each item As String In FilterKeywords.Items
            If SubjEmail.Contains(item) Then
                Return True
            End If
        Next

        Return False
    End Function
#End Region
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
            Return ""
        End Try

    End Function
    Function StringBetween(value As String, a As String,
                     b As String) As String
        ' Get positions for both string arguments.
        Dim posA As Integer = value.IndexOf(a)
        Dim posB As Integer = value.LastIndexOf(b)
        If posA = -1 Then
            Return ""
        End If
        If posB = -1 Then
            Return ""
        End If

        Dim adjustedPosA As Integer = posA + a.Length
        If adjustedPosA >= posB Then
            Return ""
        End If

        ' Get the substring between the two positions.
        Return value.Substring(adjustedPosA, posB - adjustedPosA)
    End Function

    Public Function MapUrl(ByVal baseAddress As String, ByVal relativePath As String) As String

        Dim u As New System.Uri(baseAddress)

        If relativePath = "./" Then
            relativePath = "/"
        End If

        If relativePath.StartsWith("/") Then
            Return u.Scheme + Uri.SchemeDelimiter + u.Authority + relativePath
        Else
            Dim pathAndQuery As String = u.AbsolutePath
            ' If the baseAddress contains a file name, like ..../Something.aspx
            ' Trim off the file name
            pathAndQuery = pathAndQuery.Split("?")(0).TrimEnd("/")
            If pathAndQuery.Split("/")(pathAndQuery.Split("/").Count - 1).Contains(".") Then
                pathAndQuery = pathAndQuery.Substring(0, pathAndQuery.LastIndexOf("/"))
            End If
            baseAddress = u.Scheme + Uri.SchemeDelimiter + u.Authority + pathAndQuery

            'If the relativePath contains ../ then
            ' adjust the baseAddress accordingly

            While relativePath.StartsWith("../")
                relativePath = relativePath.Substring(3)
                If baseAddress.LastIndexOf("/") > baseAddress.IndexOf("//" + 2) Then
                    baseAddress = baseAddress.Substring(0, baseAddress.LastIndexOf("/")).TrimEnd("/")
                End If
            End While

            Return baseAddress + "/" + relativePath
        End If

    End Function

    Private Sub BWLoadWebsite_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWLoadContact.RunWorkerCompleted

        TextBox1.Text = ContactName

        TextBox8.Text = ContactNameForEmail

        TextBox2.Text = ContactPhone

        TextBox4.Text = WCMTown
    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        Dim SV As String = ListBox1.SelectedItem
        Dim element As IWebElement

        driver.SwitchTo().Window(driver.WindowHandles(0))

        'EMAIL HANDLING ON DOUBLE CLICK LB1 AND LB2
        Try
            If Email1Entered = False Then
                If Not Email1TEXT = SV Then
                    If Not Email2TEXT = SV Then
                        element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_EMAIL']"))
                        element.SendKeys(SV)
                        Email1TEXT = SV
                        Email1Entered = True
                    End If
                End If
            End If
        Catch ex As Exception
        End Try

        Try
            If Email2Entered = False Then
                If Not Email1TEXT = SV Then
                    If Not Email2TEXT = SV Then
                        element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ADDN_EMAIL']"))
                        element.SendKeys(SV)
                        Email2TEXT = SV
                        Email2Entered = True
                    End If
                End If
            End If
        Catch ex As Exception
        End Try

        Try
            If Email1Entered = True Then
                If Email2Entered = True Then My.Computer.Clipboard.SetText(SV)
            End If
        Catch ex As Exception
        End Try

        'EMAIL HANDLING ON DOUBLE CLICK LB1 AND LB2
    End Sub

    Private Sub ListBox2_DoubleClick(sender As Object, e As EventArgs) Handles ListBox2.DoubleClick
        Dim SV As String = ListBox2.SelectedItem
        Dim element As IWebElement

        driver.SwitchTo().Window(driver.WindowHandles(0))


        'EMAIL HANDLING ON DOUBLE CLICK LB1 AND LB2
        Try
            If Email1Entered = False Then
                If Not Email1TEXT = SV Then
                    If Not Email2TEXT = SV Then
                        element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_EMAIL']"))
                        element.SendKeys(SV)
                        Email1TEXT = SV
                        Email1Entered = True
                    End If
                End If
            End If
        Catch ex As Exception
        End Try

        Try
            If Email2Entered = False Then
                If Not Email1TEXT = SV Then
                    If Not Email2TEXT = SV Then
                        element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ADDN_EMAIL']"))
                        element.SendKeys(SV)
                        Email2TEXT = SV
                        Email2Entered = True
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
        Try
            If Email1Entered = True Then
                If Email2Entered = True Then My.Computer.Clipboard.SetText(SV)
            End If
        Catch ex As Exception
        End Try

        'EMAIL HANDLING ON DOUBLE CLICK LB1 AND LB2
    End Sub

    Private Sub EnterOwnerName(ByVal OwnerName As String)
        Try
            Dim s As String = OwnerName.ToLower

            s = s.Replace(" secretary", "").Replace(" vice president", "").Replace(" president", "").Replace(" manager", "").Replace(" member", "").Replace(" owner", "").Replace(" president and member", "").Replace(" treasurer", "").Replace(" director", "").Replace(" and ", "").Replace(" sole proprietor", "").Replace(" managing", "").Replace("chairman", "")

            Dim result As String = Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s)

            Dim rtb9 As String, rtb10 As String

            rtb9 = result.Substring(0, s.IndexOf(" "))

            result = result.Replace(OwnerName & " ", "").Replace(0, "").Replace(vbTab, "").Replace(rtb9, "")

            rtb10 = result

            TextBox10.Text = ""

            driver.SwitchTo().Window(driver.WindowHandles(0))

            Dim element As IWebElement



            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_FIRSTNAME']")) : element.Clear()
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : element.Clear()
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : element.Clear()


            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_FIRSTNAME']")) : element.SendKeys(rtb9)
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : element.SendKeys(rtb9)
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : element.SendKeys(rtb10)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        EnterOwnerName(TextBox10.Text)
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then TTabHandler.Start() Else TTabHandler.Stop()
    End Sub

    Private Sub TTabHandler_Tick(sender As Object, e As EventArgs) Handles TTabHandler.Tick
        TTabHandler.Stop()

        'Current mouse position
        Dim xOse As Integer = Cursor.Position.X, yOse As Integer = Cursor.Position.Y

        Try
            If yOse < 31 Then
                'Y range
                Dim TabsYRangeStart As Integer = -1, TabsYRangeEnd As Integer = 30


                'DETECT FIRST TAB
                Dim FirstTabRangeXStart As Integer = 0, FirstTabRangeXEnd As Integer = 230
                If xOse > FirstTabRangeXStart And xOse < FirstTabRangeXEnd Then If yOse > TabsYRangeStart And yOse < TabsYRangeEnd Then driver.SwitchTo().Window(driver.WindowHandles(0))

                Try
                    'DETECT SECOND TAB
                    Dim SecondTabRangeXStart As Integer = 260, SecondTabRangeXEnd As Integer = 470
                    If xOse > SecondTabRangeXStart And xOse < SecondTabRangeXEnd Then If yOse > TabsYRangeStart And yOse < TabsYRangeEnd Then driver.SwitchTo().Window(driver.WindowHandles(1))
                Catch ex As Exception

                End Try


                Try
                    'DETECT THIRD TAB
                    Dim ThirdTabRangeXStart As Integer = 500, ThirdTabRangeXEnd As Integer = 710
                    If xOse > ThirdTabRangeXStart And xOse < ThirdTabRangeXEnd Then If yOse > TabsYRangeStart And yOse < TabsYRangeEnd Then driver.SwitchTo().Window(driver.WindowHandles(2))
                Catch ex As Exception

                End Try


                Try
                    'DETECT FOURTH TAB
                    Dim FourthTabRangeXStart As Integer = 740, FourthTabRangeXEnd As Integer = 950
                    If xOse > FourthTabRangeXStart And xOse < FourthTabRangeXEnd Then If yOse > TabsYRangeStart And yOse < TabsYRangeEnd Then driver.SwitchTo().Window(driver.WindowHandles(3))
                Catch ex As Exception

                End Try


                Try
                    'DETECT FIFTH TAB
                    Dim FifthTabRangeXStart As Integer = 980, FifthTabRangeXEnd As Integer = 1160
                    If xOse > FifthTabRangeXStart And xOse < FifthTabRangeXEnd Then If yOse > TabsYRangeStart And yOse < TabsYRangeEnd Then driver.SwitchTo().Window(driver.WindowHandles(4))
                Catch ex As Exception

                End Try

            End If
        Catch ex As Exception

        End Try

        If CheckBox2.Checked = True Then TTabHandler.Start()
    End Sub

    Private Sub TDrawHandler_Tick(sender As Object, e As EventArgs) Handles TDrawHandler.Tick
        TDrawHandler.Stop()
        Try
            If driver.Url.Contains("searchctbusiness.ctdata.org") Then

                Dim element As IWebElement
                Dim js As IJavaScriptExecutor = driver

                element = CType(CType(driver, IJavaScriptExecutor).ExecuteScript("return document.elementFromPoint(arguments[0], arguments[1])", Cursor.Position.X, Cursor.Position.Y - 80), IWebElement)

                Dim jsDriver = CType(driver, IJavaScriptExecutor)
                Dim highlightJavascript As String = "arguments[0].style.cssText = ""border-width: 1px; border-style: solid; border-color: yellow"";"
                jsDriver.ExecuteScript(highlightJavascript, New Object() {element})

                If element.Text = LastKnownElementString Then
                    NumberOfOccurences += 1

                    If GetAsyncKeyState(1) <> 0 Then

                        If NumberOfOccurences = 20 Then

                            EnterOwnerName(LastKnownElementString)
                            ' TextBox9.Text = LastKnownElementString
                            NumberOfOccurences = 0
                            LastKnownElementString = ""

                            Dim ThighlightJavascript As String = "arguments[0].style.cssText = ""border-width: 3px; border-style: solid; border-color: red"";"
                            jsDriver.ExecuteScript(ThighlightJavascript, New Object() {element})
                        End If

                    End If
                Else
                    NumberOfOccurences = 0
                    LastKnownElementString = element.Text
                End If
            End If

        Catch ex As Exception
        End Try

        If CheckBox3.Checked = True Then TDrawHandler.Start()
    End Sub

    Private Sub TextBox5_DoubleClick(sender As Object, e As EventArgs) Handles TextBox5.DoubleClick
        driver.SwitchTo().Window(driver.WindowHandles.Last)
        Dim js As IJavaScriptExecutor = driver

        Try
            js.ExecuteScript("window.open();")
            driver.SwitchTo().Window(driver.WindowHandles.Last)
            driver.Navigate.GoToUrl(TextBox5.Text)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TextBox6_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles TextBox6.MouseDoubleClick
        driver.SwitchTo().Window(driver.WindowHandles.Last)
        Dim js As IJavaScriptExecutor = driver

        Try
            js.ExecuteScript("window.open();")
            driver.SwitchTo().Window(driver.WindowHandles.Last)
            driver.Navigate.GoToUrl(TextBox6.Text)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TextBox9_DoubleClick(sender As Object, e As EventArgs) Handles TextBox9.DoubleClick
        driver.SwitchTo().Window(driver.WindowHandles.Last)
        Dim js As IJavaScriptExecutor = driver

        Try
            js.ExecuteScript("window.open();")
            driver.SwitchTo().Window(driver.WindowHandles.Last)
            driver.Navigate.GoToUrl(TextBox9.Text)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TextBox7_DoubleClick(sender As Object, e As EventArgs) Handles TextBox7.DoubleClick
        driver.SwitchTo().Window(driver.WindowHandles.Last)
        Dim js As IJavaScriptExecutor = driver

        Try
            js.ExecuteScript("window.open();")
            driver.SwitchTo().Window(driver.WindowHandles.Last)
            driver.Navigate.GoToUrl(TextBox7.Text)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        Button1.Enabled = False
        Label4.Text = WCMAddressList.Item(WCMNamesIndex)
        Try
            Dim TestString As String = WCMNamesList.Item(WCMNamesIndex)

            ListBox1.Items.Clear()
            ListBox2.Items.Clear()
            ListBox1.BackColor = Color.White
            ListBox2.BackColor = Color.White

            BWLoadContact.RunWorkerAsync()
        Catch ex As Exception
            MsgBox("Done!")
            Application.Exit()
        End Try
    End Sub

    Private Sub TBackgroundCheck_Tick(sender As Object, e As EventArgs) Handles TBackgroundCheck.Tick
        TBackgroundCheck.Stop()

        ClipboardContainer = My.Computer.Clipboard.GetText

        BWBackgroundHandler.RunWorkerAsync()
    End Sub

    Private Sub BWBackgroundHandler_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWBackgroundHandler.DoWork
        Dim element As IWebElement

        Try
            If ClipboardContainer.Contains("@") Then
                Try
                    Dim CBC As New RichTextBox
                    CBC.Text = ClipboardContainer
                    For Each line As String In CBC.Lines
                        If Not line = "" Then
                            Dim LbContainsEmail As Boolean = False
                            For Each item As String In ListBox1.Items
                                If item = line Then
                                    LbContainsEmail = True
                                End If
                            Next

                            If LbContainsEmail = False Then
                                ListBox1.Items.Add(line)
                                ListBox1.BackColor = Color.PaleGreen
                            End If
                        End If
                    Next

                    My.Computer.Clipboard.Clear()
                Catch ex As Exception

                End Try

            End If

        Catch ex As Exception

        End Try

        Try

            If driver.WindowHandles(1) Then
                If EmailFound = False Then
                    Dim PageText As New RichTextBox
                    Try
                        element = driver.FindElement(By.XPath("html"))
                        PageText.Text = element.Text
                    Catch ex As Exception
                        PageText.Text = ""
                    End Try
                    PageText.AppendText(Environment.NewLine)
                    PageText.AppendText(driver.PageSource)
                    Dim emails As New List(Of String)
                    Dim adrRx As Regex = New Regex("([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})")
                    For Each ii As Match In adrRx.Matches(PageText.Text.ToLower)
                        Dim CheckForInvalidDomains As Boolean = FilterResults(ii.Value)
                        If Not CheckForInvalidDomains = True Then
                            emails.Add(ii.Value)
                        End If
                    Next
                    If Not emails.Count = 0 Then
                        Dim emailsString As String = Join(emails.Distinct.ToArray, "; ")

                        Dim TempTB As New TextBox
                        TempTB.AppendText(emailsString)
                        TempTB.Text = TempTB.Text.Replace("; ", Environment.NewLine)

                        For Each line As String In TempTB.Lines
                            If Not line = "" Then
                                Dim LbContainsEmail As Boolean = False
                                For Each item As String In ListBox1.Items
                                    If item = line Then LbContainsEmail = True
                                Next

                                If LbContainsEmail = False Then
                                    ListBox1.Items.Add(line)
                                    ListBox1.BackColor = Color.PaleGreen
                                End If
                            End If
                        Next

                        EmailFound = True
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress

        If Asc(e.KeyChar) = 13 Then
            Button1.Enabled = False
            Label4.Text = WCMAddressList.Item(WCMNamesIndex)
            Try
                e.Handled = True
                Dim TestString As String = WCMNamesList.Item(WCMNamesIndex)

                ListBox1.Items.Clear()
                ListBox2.Items.Clear()
                ListBox1.BackColor = Color.White
                ListBox2.BackColor = Color.White

                BWLoadContact.RunWorkerAsync()
            Catch ex As Exception
                MsgBox("Done!")
                Application.Exit()
            End Try
        End If

    End Sub
End Class