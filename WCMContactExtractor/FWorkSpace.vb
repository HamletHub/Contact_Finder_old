Imports System.Text.RegularExpressions
Imports System.Threading
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium.Remote
Imports OpenQA.Selenium.Support.UI

Public Class FWorkSpace
    Public driver As IWebDriver 'Chrome webdriver initialization
    Public driver2 As IWebDriver 'Chrome webdriver initialization for finding emails in the background
    Dim FilterKeywords As New ListBox 'List of email strings to ignore
    ReadOnly sql As New SQLiteControl() 'SQL initialization
    Public GoogleThread As New GoogleThread 'Initialization of the Google Thread which finds new contacts
    Public QuitGoogleThread As Boolean = False 'Boolean to quit Google thread (ex. when user quits the application)

    Dim Email1Entered As Boolean = False, Email2Entered As Boolean = False
    Dim Email1TEXT As String = "", Email2TEXT As String = ""

    Dim EmailThread As Thread 'Initialization of background thread that looks for emails

    Dim BusinessEmail As String, NewBusinessEmail As String
    Dim BwReportProgres As String = "", NewBwReportProgres As String = ""

    Dim CurrentTID As Integer 'Current SQL row ID

    Dim ContactName As String, ContactNameForEmail As String, ContactPhone As String, ContactAddress As String, ContactWebsite As String, ContactTwitter As String, ContactInstagram As String, ContactFacebook As String, ContactLinkedIn As String 'Contact details

    Private Sub FWorkSpace_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeChromeOptions() 'Initialize all Chrome instances

        AddEmailFilterKeywords() 'Add email keywords to the filter

        CheckForIllegalCrossThreadCalls = False 'Makes sure that the app won't freeze in case of heavy load, so that the app is always responsive

        LoadZoho() 'Load Zoho website

        'Counts all rows in bcontacts table SQL database
        Dim rowsCount As Integer = 0
        sql.ExecQuery("SELECT bplaceid FROM bcontacts;")
        For Each r As DataRow In sql.DBDT.Rows : rowsCount += 1 : Next
        LRecordCount.Text = "Total pending records: " & rowsCount.ToString 'Display number of rows

        My.Computer.Clipboard.Clear() 'Clear the computer clipboard

        'If there are less than 100 rows in the database, run Google thread to get more contacts
        If rowsCount < 100 Then If GoogleThread.ThreadWorking = False Then GoogleThread.InitializeGoogleThread() 'Check for more contacts from google maps

        'If there are less than 10 contacts in the database, disable "NEXT" button, inform user and wait until there are at least 11 contacts in the database.
        If rowsCount < 10 Then
            Button1.Enabled = False
            MsgBox("You're running out of contacts in the database, and you have less than 10. Please hang on a little bit while the app gets contacts from the internet! Next button will became available once you have more than 10 contacts in the database.")
            TCheckContactsCount.Start() 'Start the timer to check how many contacts are in the database
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim rowCount As Integer = 0 'Declare rowCount variable

        Button1.Enabled = False 'Disable button

        TCheckContactsCount.Stop() 'Stop the timer

        ListBox1.Items.Clear() : ListBox2.Items.Clear() : ListBox1.BackColor = Color.White : ListBox2.BackColor = Color.White 'Clear listboxes

        ContactName = "" : ContactNameForEmail = "" : ContactPhone = "" : ContactAddress = "" : ContactWebsite = "" : ContactTwitter = "" : ContactInstagram = "" : ContactFacebook = "" : ContactLinkedIn = "" 'Clear contact details from variables

        TextBox1.Text = "" : TextBox2.Text = "" : TextBox3.Text = "" : TextBox4.Text = "" : TextBox5.Text = "" : TextBox6.Text = "" : TextBox7.Text = "" : TextBox8.Text = "" : TextBox9.Text = "" 'Clear textboxes

        Email1Entered = False : Email2Entered = False : Email1TEXT = "" : Email2TEXT = "" 'Clear email variables

        CurrentTID = 0 'Set TID to = 0 - CurrentTID is used to remove currently processed contact from the database

        driver2.Navigate.GoToUrl("https://www.bing.com") 'Navigate to the bing.com

        My.Computer.Clipboard.Clear() 'Clear computer clipboard

        sql.ExecQuery("SELECT bplaceid FROM bcontacts;")

        For Each r As DataRow In sql.DBDT.Rows : rowCount += 1 : Next 'Count contact rows in the database
        LRecordCount.Text = "Total pending records on the list: " & rowCount.ToString 'Display number of contacts

        If rowCount < 150 Then If GoogleThread.ThreadWorking = False Then GoogleThread.InitializeGoogleThread() 'If row count is less than 150, run google thread to get more contacts

        sql.ExecQuery("SELECT * FROM bcontacts;") 'Select bcontacts table in SQL db

        Dim rn As New Random, nb As Integer
        nb = rn.Next(0, sql.DBDT.Rows.Count) 'Generate random number

        Try
            CurrentTID = sql.DBDT.Rows(nb).Item(0) 'Set ID from the bcontacts table based on the random number
        Catch ex As Exception
            MsgBox("Following error occured: " & ex.Message & " - You're most likely out of contacts in the database. Please hang on a little bit while the app gets contacts from the internet!") 'If unable to set ID, that means the table is empty
            Button1.Enabled = False 'Enable button
            TCheckContactsCount.Start() 'Start the timer to check how many contacts are in the database
            Exit Sub
        End Try


        'Assign contact name to the variable and textbox
        ContactName = sql.DBDT.Rows(nb).Item(2)
        TextBox1.Text = sql.DBDT.Rows(nb).Item(2)
        'Assign contact phone to the variable and textbox
        TextBox2.Text = sql.DBDT.Rows(nb).Item(3)
        ContactPhone = sql.DBDT.Rows(nb).Item(3)
        'Assign contact address to the variable and textbox
        TextBox4.Text = sql.DBDT.Rows(nb).Item(4)
        ContactAddress = sql.DBDT.Rows(nb).Item(4)
        'Assign contact website to the variable and textbox
        TextBox3.Text = sql.DBDT.Rows(nb).Item(5)
        ContactWebsite = sql.DBDT.Rows(nb).Item(5)

        BLoadWebsite.RunWorkerAsync() 'Run background worker to load the website
    End Sub
    Private Sub BLoadWebsite_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BLoadWebsite.DoWork
        Dim element As IWebElement 'Declare iWebElement
        Dim js As IJavaScriptExecutor = driver 'Declare IJavaScriptExecutor to be able to use execute JavaScript

        Thread.Sleep(500) 'Wait for 500ms

        Try
            driver.SwitchTo().Window(driver.WindowHandles(1)) 'Switch to the second tab
            driver.Navigate.GoToUrl("https://google.com") 'Navigate to Google
        Catch ex As Exception
            driver.SwitchTo().Window(driver.WindowHandles.Last) 'Second tab does not exist, open new window and then switch to the second tab
            js.ExecuteScript("window.open();") 'Open new tab
            driver.SwitchTo().Window(driver.WindowHandles(1)) 'Finally switch to the second tab

            driver.Navigate.GoToUrl("https://google.com") 'Navigate to Google
        End Try

        Try
            element = driver.FindElement(By.LinkText("English")) 'Find "English" button to change Google to English language
            element.Click() 'Click on that button
        Catch ex As Exception
        End Try

        Dim TownForGoogleSearch As String = "ct" 'If no town is provided, "ct" is default
        If Not ContactAddress = "" Then TownForGoogleSearch = GetTownFromAddress(ContactAddress) 'If contact address is not nothing, extract town from this full address to be able to search google with proper town

        Try
            element = driver.FindElement(By.XPath("//*[@id='tsf']/div[2]/div/div[1]/div/div[1]/input")) 'Find google input element
            element.SendKeys(TextBox1.Text & " " & TownForGoogleSearch & " CT") 'Send search query to the input element
            Thread.Sleep(500) 'Wait for 500ms
            element.SendKeys(Keys.Enter) 'Send enter key to the input element
        Catch ex As Exception
            element = driver.FindElement(By.XPath("//*[@id='tsf']/div[2]/div[1]/div[1]/div/div[2]/input")) 'In case if an error, try another xPath to input search query
            element.SendKeys(TextBox1.Text & " " & TownForGoogleSearch & " CT") 'Send search query
            Thread.Sleep(500) 'Sleep for 500ms
            element.SendKeys(Keys.Enter) 'Send enter
        End Try

        Dim SideBarExists As Boolean = False 'Declare variable to find out if Google sidebar exists. We use that sidebar to scrape data
        Try
            element = driver.FindElement(By.XPath("//*[@id='rhs']")) 'If this xPath exists, that means sidebar also exists
            SideBarExists = True 'Set this to true if sidebar exists
        Catch ex As Exception
            'If sidebar does not exist, SideBarExists will remain false
        End Try

        If SideBarExists = True Then 'If sidebar exists...
            Dim SidebarOuterHTML As String 'Declare string to store outerHTML of the sidebar
            SidebarOuterHTML = element.GetAttribute("outerHTML") 'Get sidebar outerHTML

            Dim TempContactName As String = ExtractData(SidebarOuterHTML, "data-attrid=""title""", "</div>") 'Extract temporary Contact name

            ContactNameForEmail = StringBetween(TempContactName, "<span>", "</span>") 'Extract contact name for email
            ContactNameForEmail = ContactNameForEmail.Replace("&amp;", "&")
            ContactName = ContactName.Replace("&amp;", "&")

            TextBox8.Text = ContactNameForEmail 'Set textbox value

            If ContactPhone = "" Then ContactPhone = ExtractData(SidebarOuterHTML, "<span>+1 ", "</span>") 'Extract contact phone
            TextBox2.Text = ContactPhone 'Set textbox value

            If ContactAddress = "ct" Then ContactAddress = ExtractData(SidebarOuterHTML, "data-address=""", """ ") 'Extract contact address
            TextBox4.Text = ContactAddress 'Set textbox value

            Dim SidebarLinks As MatchCollection = Regex.Matches(SidebarOuterHTML, "<a.*?href=""(.*?)"".*?>(.*?)</a>") 'Get all links from the sidebar outerHTML

            For Each match As Match In SidebarLinks 'For each match (link) in the sidebar outerHTML
                Dim matchUrl As String = match.Groups(1).Value 'Get URL match
                If matchUrl.StartsWith("#") Then Continue For 'Ignore all anchor links
                If matchUrl.ToLower.StartsWith("javascript:") Then Continue For 'Ignore all javascript calls
                If matchUrl.ToLower.StartsWith("mailto:") Then Continue For 'Ignore all email links

                If TextBox3.Text = "" Then If match.Groups(2).Value = "Website" Then TextBox3.Text = matchUrl 'If textbox with website is empty, get website from the sidebar as well

                If matchUrl.Contains("facebook.com") Then ContactFacebook = matchUrl : TextBox5.Text = ContactFacebook 'Get Facebook link from the sidebar
                If matchUrl.Contains("twitter.com") Then ContactTwitter = matchUrl : TextBox6.Text = ContactTwitter 'Get Twitter link from the sidebar
                If matchUrl.Contains("instagram.com") Then ContactInstagram = matchUrl : TextBox9.Text = ContactInstagram 'Get Instagram link from the sidebar
                If matchUrl.Contains("linkedin.com") Then ContactLinkedIn = matchUrl : TextBox7.Text = ContactLinkedIn 'Get LinkedIn link from the sidebar
            Next

            'Dim MainBodyOuterHtml As String 'Declare string to store outerHTML of the main body. We use it to get additional social media links
            'element = driver.FindElement(By.XPath("//*[@id='rso']")) 'Find main body xPath
            'MainBodyOuterHtml = element.GetAttribute("outerHTML") 'Get outerHTML of the main body
            'Dim MainBodyLinks As MatchCollection = Regex.Matches(MainBodyOuterHtml, "<a.*?href=""(.*?)"".*?>(.*?)</a>") 'Get all links from the main body outerHTML
            'For Each match As Match In MainBodyLinks 'For each match (link) in the main body outerHTML
            '    Dim matchUrl As String = match.Groups(1).Value 'Get URL match
            '    If matchUrl.StartsWith("#") Then Continue For 'Ignore all anchor links
            '    If matchUrl.ToLower.StartsWith("javascript:") Then Continue For 'Ignore all javascript calls
            '    If matchUrl.ToLower.StartsWith("mailto:") Then Continue For 'Ignore all email links

            '    If matchUrl.Contains("facebook.com") Then ContactFacebook = matchUrl : TextBox5.Text = ContactFacebook 'Get Facebook link from the sidebar
            '    If matchUrl.Contains("twitter.com") Then ContactTwitter = matchUrl : TextBox6.Text = ContactTwitter 'Get Twitter link from the sidebar
            '    If matchUrl.Contains("instagram.com") Then ContactInstagram = matchUrl : TextBox9.Text = ContactInstagram 'Get Instagram link from the sidebar
            '    If matchUrl.Contains("linkedin.com") Then ContactLinkedIn = matchUrl : TextBox7.Text = ContactLinkedIn 'Get LinkedIn link from the sidebar
            'Next

            If Not TextBox3.Text = "" Then 'If textbox with website is not empty...
                If Not TextBox3.Text.Contains("http") Then 'If website does not contain http, we need to add that to the website
                    Dim TempWB As String = TextBox3.Text
                    ContactWebsite = "http://" & TempWB 'Add https:// to the website
                Else
                    ContactWebsite = TextBox3.Text 'If website contain http, no worries! :)
                End If

                LEmailThread.Text = "Searching for emails..." 'Set label to indicate that email thread is working
                LEmailThread.ForeColor = Color.Red 'Set label color

                EmailThread = New Thread(AddressOf CheckForEmails) 'Initialize thread
                EmailThread.Start() 'Start thread

                Try
                    driver.SwitchTo().Window(driver.WindowHandles(2)) 'Switch to the third tab
                Catch ex As Exception
                    driver.SwitchTo().Window(driver.WindowHandles.Last) 'If third tab is not opened...
                    js.ExecuteScript("window.open();") 'Open it...
                    driver.SwitchTo().Window(driver.WindowHandles(2)) 'And then switch to it
                End Try

                driver.Navigate.GoToUrl(ContactWebsite) 'Navigate to the website

                Dim AdditionalLinks As MatchCollection = Regex.Matches(driver.PageSource, "<a.*?href=""(.*?)"".*?>(.*?)</a>") 'Get additional links directly from the website
                For Each match As Match In AdditionalLinks 'For each link on the website...
                    Dim matchUrl As String = match.Groups(1).Value 'Get URL match
                    If matchUrl.StartsWith("#") Then Continue For 'Ignore all anchor links
                    If matchUrl.ToLower.StartsWith("javascript:") Then Continue For 'Ignore all javascript calls
                    If matchUrl.ToLower.StartsWith("mailto:") Then Continue For 'Ignore all email links

                    'Find Facebook link
                    If matchUrl.Contains("facebook.com") Then
                        If ContactFacebook = "" Then If FacebookBannedWords(matchUrl) = False Then ContactFacebook = matchUrl : TextBox5.Text = ContactFacebook
                    End If

                    'Find Twitter link
                    If matchUrl.Contains("twitter.com") Then
                        If ContactTwitter = "" Then If TwitterBannedWords(matchUrl) = False Then ContactTwitter = matchUrl : TextBox6.Text = ContactTwitter
                    End If

                    'Find Instagram link
                    If matchUrl.Contains("instagram.com") Then
                        If ContactInstagram = "" Then If InstagramBannedWords(matchUrl) = False Then ContactInstagram = matchUrl : TextBox9.Text = ContactInstagram
                    End If

                    'Find LinkedIn link
                    If matchUrl.Contains("linkedin.com") Then
                        If ContactLinkedIn = "" Then If LinkedInBannedWords(matchUrl) = False Then ContactLinkedIn = matchUrl : TextBox7.Text = ContactLinkedIn
                    End If
                Next

            Else 'Website does not exists
                Try
                    driver.SwitchTo().Window(driver.WindowHandles(2)) 'Switch to the third tab
                Catch ex As Exception
                    driver.SwitchTo().Window(driver.WindowHandles.Last) 'If third tab is not opened...
                    js.ExecuteScript("window.open();") 'Open it...
                    driver.SwitchTo().Window(driver.WindowHandles(2)) 'And then switch to it
                End Try

                If Not driver.Url.Contains("bing.com") Then driver.Navigate.GoToUrl("https://bing.com") 'Navigate to the bing.com
            End If

            If CheckBox1.Checked = True Then 'If this is checked, a user wants to search searchctbusiness.ctdata.org for owner names. Checked by default
                Try
                    Try
                        driver.SwitchTo().Window(driver.WindowHandles(3)) 'Switch to the fourth tab
                    Catch ex As Exception
                        driver.SwitchTo().Window(driver.WindowHandles.Last) 'If fourth tab is not opened...
                        js.ExecuteScript("window.open();") 'Open it...
                        driver.SwitchTo().Window(driver.WindowHandles(3)) 'Switch to it and navigate to the website
                        driver.Navigate.GoToUrl("http://searchctbusiness.ctdata.org/search_results?query=aaaaaaaa&query_limit=&index_field=business_name&start_date=1900-01-01&end_date=2019-02-01&active=&sort_by=nm_name&sort_order=asc&page=1")
                    End Try

                    Try
                        element = driver.FindElement(By.XPath("//*[@id='query']")) 'Find textbox to enter business name
                        element.SendKeys(Keys.Control + Keys.Backspace) 'Clear search field
                    Catch ex As Exception
                        driver.Navigate.GoToUrl("http://searchctbusiness.ctdata.org/search_results?query=aaaaaaaa&query_limit=&index_field=business_name&start_date=1900-01-01&end_date=2019-02-01&active=&sort_by=nm_name&sort_order=asc&page=1")
                    End Try

                    Try
                        element = driver.FindElement(By.XPath("//*[@id='query']")) : element.Clear() 'Clear the search field
                        Thread.Sleep(500) 'Wait for 500ms
                        element.SendKeys(TextBox1.Text) 'Send contact name to the search field

                        element = driver.FindElement(By.XPath("//*[@id='sidebar']/form/div/div/div[6]/span/button/i")) 'Find search button
                        Thread.Sleep(500) 'Wait for 500ms
                        element.Click() 'Click on the button
                    Catch ex As Exception
                    End Try

                    Try
                        element = driver.FindElement(By.XPath("//*[@id='main']/div/div[2]/table/tbody/tr[1]/td[1]/a")) 'Find table with search results
                    Catch ex As Exception
                        Dim bb As String = TextBox1.Text 'If there is no table, that means a search result returned no matches, so try to remove one word from the contact name, and try again

                        element = driver.FindElement(By.XPath("//*[@id='query']")) 'Find search field
                        element.SendKeys(Keys.Control + Keys.Backspace) 'Remove one word
                        element = driver.FindElement(By.XPath("//*[@id='sidebar']/form/div/div/div[6]/span/button/i")) 'Find search button
                        Thread.Sleep(500) 'Wait 500ms
                        element.Click() 'Click on the button
                    End Try

                Catch ex As Exception
                End Try
            End If

            driver.SwitchTo().Window(driver.WindowHandles(0)) 'Switch to the first tab

            If WaitForElement("//*[@id='Crm_Contacts_ACCOUNTID']", "xPath") = True Then 'Wait for the Zoho to load properly
                Try
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ACCOUNTID']")) : element.Clear() 'Clear Account name
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF13']")) : element.Clear()  'Clear Account name for email
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_EMAIL']")) : element.Clear() 'Clear Email
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ADDN_EMAIL']")) : element.Clear() 'Clear Secondary Email
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF122']")) : element.Clear() 'Clear Website
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF48']")) : element.Clear() 'Clear Facebook
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF155']")) : element.Clear() 'Clear Instagram
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF156']")) : element.Clear() 'Clear Twitter
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF30']")) : element.Clear() 'Clear LinkedIn
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_PHONE']")) : element.Clear() 'Clear Phone
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_FIRSTNAME']")) : element.Clear() 'Clear First name
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : element.Clear() 'Clear Dear field
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : element.Clear() 'Clear Last name
                Catch ex As Exception
                    driver.Navigate.GoToUrl("https://crm.zoho.com/crm/org19427096/tab/Contacts/create") 'If first tab loaded another page instead of "Create contact" page, navigate to the "Create contact" Zoho URL and clear input elements again

                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ACCOUNTID']")) : element.Clear() 'Clear Account name
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF13']")) : element.Clear()  'Clear Account name for email
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_EMAIL']")) : element.Clear() 'Clear Email
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ADDN_EMAIL']")) : element.Clear() 'Clear Secondary Email
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF122']")) : element.Clear() 'Clear Website
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF48']")) : element.Clear() 'Clear Facebook
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF155']")) : element.Clear() 'Clear Instagram
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF156']")) : element.Clear() 'Clear Twitter
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF30']")) : element.Clear() 'LClear inkedIn
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_PHONE']")) : element.Clear() 'Clear Phone
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_FIRSTNAME']")) : element.Clear() 'Clear First name
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : element.Clear() 'Clear Dear field
                    element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : element.Clear() 'Clear Last name
                End Try
            End If

            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, "Hello") 'Enter last name
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, "Hello") 'Enter first name

            TextBox8.Text = TextBox8.Text.Replace(" Co Inc", "").Replace(" Co. Inc.", "").Replace(", LLC", "").Replace(", Inc.", "").Replace(" Inc.", "").Replace(" LLC", "").Replace(" INC", "").Replace(" llc", "").Replace(" Inc", "").Replace(" LLC.", "").Replace(",Inc.", "").Replace(" Inc", "").Replace(" L.L.C.", "") 'Format contact name

            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_ACCOUNTID']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox1.Text) 'Enter account name

            If TextBox8.Text = "" Then TextBox8.Text = TextBox1.Text 'If acc. name for email is empty, assing account name to it
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF13']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox8.Text) 'Enter account name for email

            If Not TextBox3.Text = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF122']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox3.Text) 'Enter contact website
            If Not ContactFacebook = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF48']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox5.Text) 'Enter contact Facebook
            If Not ContactInstagram = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF155']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox9.Text) 'Enter contact Instagram
            If Not ContactTwitter = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF156']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox6.Text) 'Enter contact Twitter
            If Not ContactLinkedIn = "" Then element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF30']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox7.Text) 'Enter contact LinkedIn


            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_PHONE']")) : js.ExecuteScript("arguments[0].value = arguments[1]", element, TextBox2.Text) 'Enter contact Phone


            element = driver.FindElement(By.XPath("//*[@id='Contacts_fldRow_CONTACTCF157']/div[2]/div/span/span[1]/span/ul/li/input")) 'Find WCM Town element
            element.SendKeys(Keys.Backspace) 'Clear it

            Thread.Sleep(500) 'Sleep for 500ms

            Dim WCMTown As String = GetTownFromAddress(TextBox4.Text)
            element = driver.FindElement(By.XPath("//*[@id='Contacts_fldRow_CONTACTCF157']/div[2]/div/span/span[1]/span/ul/li/input"))  'Find WCM Town element once again
            element.SendKeys(WCMTown + Keys.Enter) 'Get town from address and enter it into Zoho

            Thread.Sleep(500) 'Wait 500ms
            element = driver.FindElement(By.ClassName("crmBodyWin")) 'Find Zoho main body element
            element.Click() 'Click on it
            element.SendKeys(Keys.PageUp) 'And move UP
            element.SendKeys(Keys.PageUp) 'Move UP again!

            driver.SwitchTo().Window(driver.WindowHandles(2)) 'Switch to the third tab
        End If

        sql.AddParam("@id", CurrentTID) : sql.ExecQuery("DELETE FROM bcontacts WHERE ROWID=@id") 'Select current contact ID and delete it from bcontacts table

        Try
            If Not TextBox1.Text = "" Then 'If textbox with contact name is not empty....
                sql.AddParam("@name", TextBox1.Text) : sql.ExecQuery("INSERT INTO bDuplicateNames (bnames) VALUES (@name);") 'Add contact to the duplicate database
                If sql.HasExpetion(True) Then MsgBox("SEND A SCREENSHOT OF THIS ERROR TO DALIBOR! - " & sql.Exception) 'In case of error, inform user what is going on
            End If
        Catch ex As Exception
            MsgBox("SEND A SCREENSHOT OF THIS ERROR TO DALIBOR ASAP! - " & ex.Message)
        End Try

        Try
            If Not TextBox8.Text = "" Then 'If textbox with contact name is not empty....
                If Not TextBox1.Text = TextBox8.Text Then 'If contact name and contact name for email textboxes are not empty...
                    sql.AddParam("@name", TextBox8.Text) : sql.ExecQuery("INSERT INTO bDuplicateNames (bnames) VALUES (@name);") 'Add contact to the duplicate database
                    If sql.HasExpetion(True) Then MsgBox("SEND A SCREENSHOT OF THIS ERROR TO DALIBOR! - " & sql.Exception) 'In case of error, inform user what is going on
                End If
            End If
        Catch ex As Exception
            MsgBox("SEND A SCREENSHOT OF THIS ERROR TO DALIBOR ASAP! - " & ex.Message)
        End Try


        Try
            If Not TextBox3.Text = "" Then 'If website is not nothing...
                If Not TextBox3.Text = "NONE" Then 'And if website is not = NONE
                    Dim FormattedWebsite As String = GoogleThread.FormatWebsite(TextBox3.Text) 'Format website for the database
                    sql.AddParam("@website", FormattedWebsite) : sql.ExecQuery("INSERT INTO bDuplicateWebsites (bwebsites)  VALUES (@website);") 'Add contact to the website duplicate database
                    If sql.HasExpetion(True) Then MsgBox("SEND A SCREENSHOT OF THIS ERROR TO DALIBOR! - " & sql.Exception) 'In case of error, inform user what is going on
                End If
            End If
        Catch ex As Exception
            MsgBox("SEND A SCREENSHOT OF THIS ERROR TO DALIBOR ASAP! - " & ex.Message)
        End Try

        Button1.Enabled = True 'Enable NEXT button
    End Sub


#Region "Functions"
    Private Function GetTownFromAddress(ByVal Address As String)
        'Get town from the full address
        If Address.ToLower.Contains("danbury, ct") Then Return "Danbury"
        If Address.ToLower.Contains("fairfield, ct") Then Return "Fairfield"
        If Address.ToLower.Contains("milford, ct") Then Return "Milford"
        If Address.ToLower.Contains("new canaan, ct") Then Return "New canaan"
        If Address.ToLower.Contains("newtown, ct") Then Return "Newtown"
        If Address.ToLower.Contains("norwalk, ct") Then Return "Norwalk"
        If Address.ToLower.Contains("ridgefield, ct") Then Return "Ridgefield"
        If Address.ToLower.Contains("stamford, ct") Then Return "Stamford"
        If Address.ToLower.Contains("westport, ct") Then Return "Westport"

        Return ""
    End Function
    Public Function ExtractData(ByRef pHTML As String, ByRef pSearchStart As String, ByRef pSearchEnd As String, Optional ByRef pSearchSub As String = "") As String
        Try
            Dim lonPhrasePos1 As Integer, lonPhrasePos2 As Integer
            ExtractData = ""

            lonPhrasePos1 = pHTML.ToUpper.IndexOf(pSearchStart.ToUpper) ' Look for the search phrase. If not found, exit this function.
            If lonPhrasePos1 = 0 Then Exit Function
            ' If the optional pSearchSub parameter has been provided, find the string AFTER the first search string's location.
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
    Function StringBetween(value As String, a As String, b As String) As String
        ' Get positions for both string arguments.
        Dim posA As Integer = value.IndexOf(a)
        Dim posB As Integer = value.LastIndexOf(b)
        If posA = -1 Then Return ""
        If posB = -1 Then Return ""

        Dim adjustedPosA As Integer = posA + a.Length
        If adjustedPosA >= posB Then Return ""

        ' Get the substring between the two positions.
        Return value.Substring(adjustedPosA, posB - adjustedPosA)
    End Function
    Private Function WaitForElement(ByVal element As String, ByVal elementMechanism As String)
        Dim TimeOut As Integer = 0 'Set timeout integer
        Do
            If TimeOut > 15 Then Return False 'If TimeOut is > than 15, element is not loaded

            If elementMechanism = "ID" Then If Not driver.FindElements(By.Id(element)).Count = 0 Then Return True 'If there are more than 1 element of type "ID", return True
            If elementMechanism = "Class" Then If Not driver.FindElements(By.ClassName(element)).Count = 0 Then Return True 'If there are more than 1 element of type "Class", return True
            If elementMechanism = "xPath" Then If Not driver.FindElements(By.XPath(element)).Count = 0 Then Return True 'If there are more than 1 element of type "xPath", return True

            Thread.Sleep(1000) : TimeOut += 1 'Sleep for 1000ms
        Loop
    End Function
    Private Function FacebookBannedWords(ByVal link As String)
        Dim BannedWords As New ListBox, BannedLinks As New ListBox 'Declare listbox variables

        'Add banned words for facebook
        BannedWords.Items.Add("sharer")
        BannedWords.Items.Add("share")
        BannedWords.Items.Add("wix")

        'Add banned links for facebook
        BannedLinks.Items.Add("https://www.facebook.com/")
        BannedLinks.Items.Add("https://www.facebook.com")
        BannedLinks.Items.Add("https//www.facebook.com/")
        BannedLinks.Items.Add("www.facebook.com")
        BannedLinks.Items.Add("www.facebook.com/")

        For Each item As String In BannedWords.Items
            If link.Contains(item) Then Return True 'If link contains banned word, return true
        Next

        For Each item As String In BannedLinks.Items
            If link = item Then Return True 'If link is the same as one of banned links, return true
        Next
        Return False
    End Function
    Private Function TwitterBannedWords(ByVal link As String)
        Dim BannedWords As New ListBox, BannedLinks As New ListBox

        'Add banned words for Twitter
        BannedWords.Items.Add("sharer")
        BannedWords.Items.Add("share")
        BannedWords.Items.Add("wix")

        'Add banned links for twitter
        BannedLinks.Items.Add("https://www.twitter.com/")
        BannedLinks.Items.Add("https://www.twitter.com")
        BannedLinks.Items.Add("https//www.twitter.com/")
        BannedLinks.Items.Add("www.twitter.com")
        BannedLinks.Items.Add("www.twitter.com/")

        For Each item As String In BannedWords.Items
            If link.Contains(item) Then Return True 'If link contains banned word, return true
        Next

        For Each item As String In BannedLinks.Items
            If link = item Then Return True 'If link is the same as one of banned links, return true
        Next
        Return False
    End Function
    Private Function InstagramBannedWords(ByVal link As String)
        Dim BannedWords As New ListBox, BannedLinks As New ListBox

        'Add banned words for Instagram
        BannedWords.Items.Add("sharer")
        BannedWords.Items.Add("share")
        BannedWords.Items.Add("wix")

        'Add banned links for Instagram
        BannedLinks.Items.Add("https://www.instagram.com/")
        BannedLinks.Items.Add("https://www.instagram.com")
        BannedLinks.Items.Add("https//www.instagram.com/")
        BannedLinks.Items.Add("www.instagram.com")
        BannedLinks.Items.Add("www.instagram.com/")

        For Each item As String In BannedWords.Items
            If link.Contains(item) Then Return True 'If link contains banned word, return true
        Next

        For Each item As String In BannedLinks.Items
            If link = item Then Return True 'If link is the same as one of banned links, return true
        Next
        Return False
    End Function
    Private Function LinkedInBannedWords(ByVal link As String)
        Dim BannedWords As New ListBox, BannedLinks As New ListBox

        'Add banned words for LinkedIn
        BannedWords.Items.Add("sharer")
        BannedWords.Items.Add("share")
        BannedWords.Items.Add("wix")

        'Add banned links for LinkedIn
        BannedLinks.Items.Add("https://www.linkedin.com/")
        BannedLinks.Items.Add("https://www.linkedin.com")
        BannedLinks.Items.Add("https//www.linkedin.com/")
        BannedLinks.Items.Add("www.linkedin.com")
        BannedLinks.Items.Add("www.linkedin.com/")

        For Each item As String In BannedWords.Items
            If link.Contains(item) Then Return True 'If link contains banned word, return true
        Next

        For Each item As String In BannedLinks.Items
            If link = item Then Return True 'If link is the same as one of banned links, return true
        Next

        Return False
    End Function
    Private Function FilterResults(ByVal ii)
        Dim SubjEmail As String = ii
        'If email contains any of items in FilterKeywords list, that means we don't want that email > return true
        For Each item As String In FilterKeywords.Items
            If SubjEmail.Contains(item) Then Return True
        Next
        Return False
    End Function
#End Region

#Region "Email SCRAPING"
    Private Sub CheckForEmails()
        'This Chrome instance works in the background
        ListBox1.Items.Clear() : ListBox2.Items.Clear() 'Clear listboxes

        driver2.Navigate.GoToUrl(ContactWebsite) 'Navigate to the website

        Dim element As IWebElement 'Declare iWebEleemnt

        Dim PageText As New RichTextBox 'Temporary Richtextbox

        Try
            element = driver2.FindElement(By.XPath("html")) 'Find HTML of the page
            PageText.Text = element.Text 'Get all the text from the HTML page
        Catch ex As Exception
            PageText.Text = "" 'On error, set richtextbox text to nothing
        End Try

        PageText.AppendText(Environment.NewLine) 'Append new line to the richtextbox
        PageText.AppendText(driver2.PageSource) 'Get page source of the page

        Dim emails As New List(Of String), Newemails As New List(Of String) 'Declare a variable to store list of emails and a variable to store list of new emails

        Dim adrRx As Regex = New Regex("([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})") 'Regex pattern to get all email variants
        For Each ii As Match In adrRx.Matches(PageText.Text.ToLower)
            Dim CheckForInvalidDomains As Boolean = FilterResults(ii.Value)
            If Not CheckForInvalidDomains = True Then emails.Add(ii.Value) 'If email is not invalid, add it to the list
        Next

        Dim SecondTryFound As Boolean = False 'Declare second try boolean

        Dim emailsString As String

        'Add potential page names to the text links listbox. We will use this list to navigate through the pages on the website
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

        For Each item As String In TextLinks.Items 'For each item in text links listbox...
            Try
                element = driver2.FindElement(By.PartialLinkText(item.ToString)) 'Find that element...

                driver2.Navigate.GoToUrl(element.GetAttribute("href")) 'And try to navigate!

                element = driver2.FindElement(By.XPath("html")) 'If navigate was successfull, get page HTML
                PageText.Text = element.Text
                'Append page HTML to the rtb
                PageText.AppendText(Environment.NewLine)
                PageText.AppendText(driver2.PageSource)

                For Each ii As Match In adrRx.Matches(PageText.Text.ToLower)
                    Dim CheckForInvalidDomains As Boolean = FilterResults(ii.Value)
                    If Not CheckForInvalidDomains = True Then emails.Add(ii.Value) 'If email is not invalid, add it to the list
                Next

                If Not emails.Count = 0 Then 'If thread found more than 0 emails...
                    emailsString = Join(emails.Distinct.ToArray, "; ") 'Join them together, separated by ";"
                    BusinessEmail = emailsString
                    BwReportProgres = BusinessEmail
                    SecondTryFound = True
                End If
            Catch ex As Exception
            End Try 'Try - Catch ensures that all goes smoothly. If it throws here, that means that particular link does not exists
        Next

        Try
            driver2.Navigate.GoToUrl(ContactFacebook) 'Navigate to facebook (if exists)

            Dim CurFBUrl As String = driver2.Url 'GET Facebook URL
            'Replace a few things...
            CurFBUrl = CurFBUrl.Replace("https://www.facebook.com/", "")
            CurFBUrl = CurFBUrl.Replace("/", "")

            driver2.Navigate.GoToUrl("https://www.facebook.com/pg/" & CurFBUrl & "/about/?ref=page_internal") 'And create about page from user ID

            element = driver2.FindElement(By.XPath("html")) 'Get page HTML

            PageText.AppendText(Environment.NewLine & element.Text) 'Append text to the  rtb

            For Each ii As Match In adrRx.Matches(PageText.Text.ToLower) 'For each match in this RTB...
                Dim CheckForInvalidDomains As Boolean = FilterResults(ii.Value)
                If Not CheckForInvalidDomains = True Then If Not emails.Contains(ii.Value) Then emails.Add(ii.Value) 'If email is not invalid or already found, add it to the list
            Next
        Catch ex As Exception
        End Try

        Try
            emailsString = Join(emails.Distinct.ToArray, "; ") 'Join them together, separated by ";"
            BusinessEmail = emailsString
            BwReportProgres = BusinessEmail
        Catch ex As Exception
        End Try

        If SecondTryFound = False Then 'In case we need more emails, we continue looking for emails on bing.com
            Dim CurrentDomain As String = ContactWebsite
            CurrentDomain = CurrentDomain.Replace("https://www.", "").Replace("http://www.", "").Replace("https://", "").Replace("http://", "").Replace("www.", "") 'Format website string
            Dim string_before As String = ""
            Try
                'Get formatted website required to search bing.com
                Dim original As String = CurrentDomain, cut_at As String = "/", stringSeparators() As String = {cut_at}, split = original.Split(stringSeparators, 2, StringSplitOptions.RemoveEmptyEntries)
                string_before = split(0)
            Catch ex As Exception
            End Try

            Try
                driver2.Navigate.GoToUrl("https://www.bing.com/search?q=%22*%40" & string_before & "%22") 'Search bing
                element = driver2.FindElement(By.XPath("html")) 'Find HTML of the page
                PageText.Text = element.Text 'Get page text

                driver2.Navigate.GoToUrl("https://www.bing.com/search?q=%22*%40" & string_before & "&first=11&FORM=PERE") 'Navigate to the second page
                element = driver2.FindElement(By.XPath("html")) 'Find HTML of the page
                PageText.Text &= element.Text 'Get page text of the second page
            Catch ex As Exception
            End Try

            Try
                For Each ii As Match In adrRx.Matches(PageText.Text.ToLower) 'For each match in this RTB...
                    Dim CheckForInvalidDomains As Boolean = FilterResults(ii.Value)
                    If Not CheckForInvalidDomains = True Then If ii.Value.Contains(string_before) Then If Not emails.Contains(ii.Value) Then Newemails.Add(ii.Value) 'If email is not invalid or already found, add it to the list
                Next
            Catch ex As Exception
            End Try
        End If

        If emails.Count = 0 Then 'If we still did not found any emails...
            BwReportProgres = BusinessEmail 'Report 0
        Else
            emailsString = Join(emails.Distinct.ToArray, "; ") ' Else, join emails
            BusinessEmail = emailsString 'Report
            BwReportProgres = BusinessEmail 'Report
        End If



        If Newemails.Count = 0 Then 'If we still did not found any emails...
            NewBwReportProgres = NewBusinessEmail ' Report 0
        Else
            Dim NewemailsString As String = Join(Newemails.Distinct.ToArray, "; ") 'Else join emails
            NewBusinessEmail = NewemailsString 'Report
            NewBwReportProgres = NewBusinessEmail 'Report
        End If

        Dim TempTB1 As New TextBox, TempTB2 As New TextBox 'Declare temporary textboxes

        TempTB1.Text = BwReportProgres 'Add values to the first one
        TempTB2.Text = NewBwReportProgres 'Add values to the second one

        TempTB1.Text = TempTB1.Text.Replace("; ", Environment.NewLine) 'Replace ";"
        TempTB2.Text = TempTB2.Text.Replace("; ", Environment.NewLine) 'Replace ";"

        For Each line As String In TempTB1.Lines
            If Not line = "" Then ListBox1.Items.Add(line) 'Add emails to the listbox
        Next
        For Each line As String In TempTB2.Lines
            If Not line = "" Then ListBox2.Items.Add(line) 'Add emails to the listbox
        Next

        BusinessEmail = "" 'Clear variable
        NewBusinessEmail = "" 'Clear variable

        driver2.Navigate.GoToUrl("https://www.bing.com") 'Navigate to bing and stand by
        LEmailThread.Text = "Thread idle..." 'Inform user that thread is idle
        LEmailThread.ForeColor = Color.Green 'Set color
    End Sub
#End Region

#Region "Object events"
    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        SendEmailsToZoho(ListBox1.SelectedItem) 'Get selected item value
    End Sub
    Private Sub ListBox2_DoubleClick(sender As Object, e As EventArgs) Handles ListBox2.DoubleClick
        SendEmailsToZoho(ListBox2.SelectedItem) 'Get selected item value
    End Sub
    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        EnterOwnerName(TextBox10.Text) 'Enter owner name
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then TTabHandler.Start() Else TTabHandler.Stop() 'Set tab handler on or off
    End Sub
    Private Sub TextBox5_DoubleClick(sender As Object, e As EventArgs) Handles TextBox5.DoubleClick
        driver.SwitchTo().Window(driver.WindowHandles.Last) 'Switch to the last tab
        Dim js As IJavaScriptExecutor = driver 'Declare JS

        Try
            js.ExecuteScript("window.open();")
            driver.SwitchTo().Window(driver.WindowHandles.Last)
            driver.Navigate.GoToUrl(TextBox5.Text) 'Navigate to the website
        Catch ex As Exception
        End Try
    End Sub
    Private Sub TextBox6_DoubleClick(sender As Object, e As EventArgs) Handles TextBox6.DoubleClick
        driver.SwitchTo().Window(driver.WindowHandles.Last) 'Switch to the last tab
        Dim js As IJavaScriptExecutor = driver 'Declare JS

        Try
            js.ExecuteScript("window.open();")
            driver.SwitchTo().Window(driver.WindowHandles.Last)
            driver.Navigate.GoToUrl(TextBox6.Text) 'Navigate to the website
        Catch ex As Exception
        End Try
    End Sub
    Private Sub TextBox9_DoubleClick(sender As Object, e As EventArgs) Handles TextBox9.DoubleClick
        driver.SwitchTo().Window(driver.WindowHandles.Last) 'Switch to the last tab
        Dim js As IJavaScriptExecutor = driver 'Declare JS

        Try
            js.ExecuteScript("window.open();")
            driver.SwitchTo().Window(driver.WindowHandles.Last)
            driver.Navigate.GoToUrl(TextBox9.Text) 'Navigate to the website
        Catch ex As Exception
        End Try
    End Sub
    Private Sub TextBox7_DoubleClick(sender As Object, e As EventArgs) Handles TextBox7.DoubleClick
        driver.SwitchTo().Window(driver.WindowHandles.Last) 'Switch to the last tab
        Dim js As IJavaScriptExecutor = driver 'Declare JS

        Try
            js.ExecuteScript("window.open();")
            driver.SwitchTo().Window(driver.WindowHandles.Last)
            driver.Navigate.GoToUrl(TextBox7.Text) 'Navigate to the website
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        Button1.Enabled = False 'Disable NEXT button while refreshing the website
        ListBox1.Items.Clear() : ListBox2.Items.Clear() : ListBox1.BackColor = Color.White : ListBox2.BackColor = Color.White 'Clear Listboxes
        BLoadWebsite.RunWorkerAsync() 'Run background worker again
    End Sub
    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) = 13 Then 'On enter keypress...
            Button1.Enabled = False 'Disable NEXT button while refreshing the website
            Try
                e.Handled = True 'Set e.Handled to true to make sure the enter key is handled properly
                ListBox1.Items.Clear() : ListBox2.Items.Clear() : ListBox1.BackColor = Color.White : ListBox2.BackColor = Color.White 'Clear Listboxes
                BLoadWebsite.RunWorkerAsync() 'Run background worker again
            Catch ex As Exception
            End Try
        End If
    End Sub
#End Region

#Region "Procedures and timers"
    Private Sub TCheckContactsCount_Tick(sender As Object, e As EventArgs) Handles TCheckContactsCount.Tick
        TCheckContactsCount.Stop() 'Stop the timer
        Dim rowsCount As Integer = 0

        sql.ExecQuery("SELECT bplaceid FROM bcontacts;") 'Select bcontacts table

        For Each r As DataRow In sql.DBDT.Rows : rowsCount += 1 : Next 'Count rows
        LRecordCount.Text = "Total pending records: " & rowsCount.ToString 'Set label to show rows count

        If rowsCount > 10 Then Button1.Enabled = True Else TCheckContactsCount.Start() 'If rows count is > than 10, enable NEXT button, else start this timer again
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
    Private Sub InitializeChromeOptions()
        Dim optionOn As New ChromeOptions

        Dim driverService = ChromeDriverService.CreateDefaultService()

        driverService.HideCommandPromptWindow = True 'Hide command prompt from user

        optionOn.AddUserProfilePreference("profile.default_content_setting_values.geolocation", 2) 'Disable location tracking for this Chrome instance

        optionOn.AddArgument("start-maximized") 'Instruct Chrome to run in maximized mode
        optionOn.AddArgument("--disable-infobars") 'Disable infobars

        Dim theContDir As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) 'Specify MyDocuments folder for Chrome session
        Dim newDir As String = "ChromeCookies" 'Folder name

        Dim thePath As String = IO.Path.Combine(theContDir, newDir) 'Combine MyDocuments with folder name

        If Not IO.Directory.Exists(thePath) Then IO.Directory.CreateDirectory(thePath) 'Create directory if it does not exists

        optionOn.AddArgument("user-data-dir=" + thePath + "/chrome-session") 'Specify Chrome user data directory
        optionOn.AddArgument("--profile-directory=Default") 'Set as default
        optionOn.AddArgument("--disable-application-cache") 'Disable caching

        driver = New ChromeDriver(driverService, optionOn) 'Initialize ChromeDriver with added arguments

        'Initiate another Chrome web driver that will search for emails
        Dim driverService2 = ChromeDriverService.CreateDefaultService()
        driverService2.HideCommandPromptWindow = True 'Hide command prompt
        Dim optionOn2 As New ChromeOptions
        optionOn2.AddUserProfilePreference("profile.default_content_setting_values.images", 2) 'Disable or enable images
        optionOn2.AddArgument("--blink-settings=imagesEnabled=false") 'Disable or enable images
        optionOn2.AddArguments("headless") 'Hide Chrome browser from the user. Headless mode means Chrome will run in the background
        driver2 = New ChromeDriver(driverService2, optionOn2) 'Initialize ChromeDriver with added arguments
    End Sub
    Private Sub AddEmailFilterKeywords()
        'Adding specific strings to filter common emails we don't need
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
    Private Sub LoadZoho()
        Dim Element As IWebElement 'Declaring iWebElement to manupulate Chrome elements
        Try
            driver.Navigate.GoToUrl("https://accounts.zoho.com/signin?servicename=ZohoCRM&signupurl=https://www.zoho.com/crm/signup.html?plan=enterprise") 'Navigating to Zoho

            Element = driver.FindElement(By.XPath("//*[@id='lid']")) 'Finding username element
            Element.SendKeys("danb@hamlethub.com") 'Send string

            Element = driver.FindElement(By.XPath("//*[@id='pwd']")) 'Finding password element
            Element.SendKeys("Dadada92") 'Send string

            Element = driver.FindElement(By.XPath("//*[@id='signin_submit']")) 'Finding submit element button
            Element.Click() 'Click on the element

        Catch ex As Exception
        End Try

        driver.Navigate.GoToUrl("https://crm.zoho.com/crm/org19427096/tab/Contacts/create") 'Navigate to the Zoho "Create" page 
    End Sub
    Private Sub FWorkSpace_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            driver.Dispose() 'Dispose Chrome instance
            driver2.Dispose() 'Dispose Chrome instance that search for emails
            driver.Quit() 'Quit Chrome instance
            driver2.Quit() 'Quit Chrome instance that search for emails

            GoogleThread.driver2.Dispose() 'Dispose Chrome / Google thread instance
            GoogleThread.driver2.Quit() 'Quit Chrome / Google thread instance

            QuitGoogleThread = True
        Catch ex As Exception

        End Try

        Form1.Show() 'Show main form
    End Sub
    Private Sub EnterOwnerName(ByVal OwnerName As String)
        Try
            Dim s As String = OwnerName.ToLower 'Declare owner name

            s = s.Replace(" secretary", "").Replace(" vice president", "").Replace(" president", "").Replace(" manager", "").Replace(" member", "").Replace(" owner", "").Replace(" president and member", "").Replace(" treasurer", "").Replace(" director", "").Replace(" and ", "").Replace(" sole proprietor", "").Replace(" managing", "").Replace("chairman", "") 'Format owner name

            Dim result As String = Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s) 'To title case

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

            'Enter formatted owner name
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_FIRSTNAME']")) : element.SendKeys(rtb9)
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_CONTACTCF9']")) : element.SendKeys(rtb9)
            element = driver.FindElement(By.XPath("//*[@id='Crm_Contacts_LASTNAME']")) : element.SendKeys(rtb10)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub SendEmailsToZoho(ByVal ListboxValue As String)
        Dim SV As String = ListboxValue 'Get selected item value

        Dim element As IWebElement 'Declare iWebElement

        driver.SwitchTo().Window(driver.WindowHandles(0)) 'Switch to the first tab

        'HANDLING DOUBLE CLICK LB1 AND LB2
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
            If Email1Entered = True Then If Email2Entered = True Then My.Computer.Clipboard.SetText(SV)
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class