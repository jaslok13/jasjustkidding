Imports Microsoft.VisualBasic
Module YouTube
    Public str_Error As String = String.Empty
    Public str_XML_Parameters As String = String.Empty
    Public str_URL_Chrome As String = String.Empty
    Public str_URL_Edge As String = String.Empty
    Public str_URL_Firefox As String = String.Empty
    Public str_URL_Opera As String = String.Empty
    Public str_File_Path_Nord_VPN As String = String.Empty

    Public bool_Open_Chrome As Boolean = True
    Public bool_Open_Edge As Boolean = True
    Public bool_Open_Firefox As Boolean = True
    Public bool_Open_Opera As Boolean = True
    Public bool_Connect_VPN As Boolean = True
    Public dt_Country_List As New System.Data.DataTable
    Public int_Wait_After_Opening_URL As Integer = 15
    Sub Main()
        Try
            Dim File_Path_Parameters_XML = "C:\Users\" & Environment.UserName() & "\Do Not Delete - Banarasi Singaporiya.xml"
            If System.IO.File.Exists(File_Path_Parameters_XML) Then
                str_XML_Parameters = System.IO.File.ReadAllText(File_Path_Parameters_XML, System.Text.Encoding.UTF8)
                str_URL_Chrome = Get_Parameter_Value("URL_Chrome")
                str_URL_Edge = Get_Parameter_Value("URL_Edge")
                str_URL_Firefox = Get_Parameter_Value("URL_Firefox")
                str_URL_Opera = Get_Parameter_Value("URL_Opera")
                str_File_Path_Nord_VPN = Get_Parameter_Value("File_Path_Nord_VPN")

                bool_Open_Chrome = IIf(Get_Parameter_Value("Open_Chrome").ToUpper.Equals("TRUE"), True, False)
                bool_Open_Edge = IIf(Get_Parameter_Value("Open_Edge").ToUpper.Equals("TRUE"), True, False)
                bool_Open_Firefox = IIf(Get_Parameter_Value("Open_Firefox").ToUpper.Equals("TRUE"), True, False)
                bool_Open_Opera = IIf(Get_Parameter_Value("Open_Opera").ToUpper.Equals("TRUE"), True, False)
                bool_Connect_VPN = IIf(Get_Parameter_Value("Connect_VPN").ToUpper.Equals("TRUE"), True, False)

                Dim str_Wait = Get_Parameter_Value("Wait_After_Opening_URL")
                If IsNumeric(str_Wait) Then
                    int_Wait_After_Opening_URL = Convert.ToInt32(str_Wait)
                End If
            Else
                Write_To_CSV("File doesn't exist: " & File_Path_Parameters_XML)
            End If

            If bool_Connect_VPN Then
                Connect_VPN()
            End If
            Close_Existing_Browers()
            Open_YouTube_Playlist()
        Catch ex As Exception
            Write_To_CSV(ex.Message)
        End Try

    End Sub
    Private Sub Get_Random_Country()
        Dim collection_Country_List As New System.Data.DataTable
        collection_Country_List.Columns.Add("Country")
        collection_Country_List.Rows.Add("Albania")
        collection_Country_List.Rows.Add("Argentina")
        collection_Country_List.Rows.Add("Austria")
        collection_Country_List.Rows.Add("Belgium")
        collection_Country_List.Rows.Add("Bosnia and Herzegovina")
        Dim text_Country As String = ""
        Try
            ' Initialize the random-number generator.
            Randomize()
            ' Generate random value between 1 and collection length.
            Dim int_Row_Index As Integer = CInt(Int(((collection_Country_List.Rows.Count - 1) * Rnd()) + 1))
            If int_Row_Index >= 0 And int_Row_Index < collection_Country_List.Rows.Count Then
                text_Country = collection_Country_List.Rows(int_Row_Index).Item(0).ToString
            Else
                text_Country = "United States"
            End If
        Catch ex As Exception
            Write_To_CSV(ex.Message)
        End Try
    End Sub

    Private Sub Close_Existing_Browers()
        Try
            Dim counter As Integer = 0
            Dim list_Browser As List(Of String) = New List(Of String)({"msedge", "chrome", "firefox", "opera"})
            For Each browser As String In list_Browser
                counter = 0

                If (browser = "msedge" And bool_Open_Edge) Or (browser = "chrome" And bool_Open_Chrome) Or (browser = "firefox" And bool_Open_Firefox) Or (browser = "opera" And bool_Open_Opera) Then
                    While System.Diagnostics.Process.GetProcessesByName(browser).Count > 0
                        counter = counter + 1
                        If counter <= 3 Then
                            For Each process_Browser As System.Diagnostics.Process In System.Diagnostics.Process.GetProcessesByName(browser)
                                Try
                                    'Try to close process Window
                                    process_Browser.CloseMainWindow()
                                    process_Browser.Close()
                                Catch ex As Exception
                                End Try
                            Next
                            System.Threading.Thread.Sleep(1000)
                        Else
                            Exit While
                        End If
                    End While

                    'Forcekill to make sure nothing is present
                    Microsoft.VisualBasic.Interaction.Shell("taskkill /f /im " & browser & ".exe", Microsoft.VisualBasic.AppWinStyle.Hide)
                    System.Threading.Thread.Sleep(1000)
                End If
            Next
            'Wait 10 Seconds, this required to clean up memory
            System.Threading.Thread.Sleep(10000)
        Catch ex As Exception
            Write_To_CSV(ex.Message)
        End Try
    End Sub
    Private Sub Open_YouTube_Playlist()
        'Open URL in Browser, neep to wait 10 seconds so that video autoplay
        'Firt we are just opening the browser and then opening the YouTube URL in a new tab
        'This is required as if we don't do this youtube videos doesn't get played automatically sometimes
        Try
            If bool_Open_Edge Then
                System.Diagnostics.Process.Start("msedge.exe")
                System.Threading.Thread.Sleep(int_Wait_After_Opening_URL * 1000)
                System.Diagnostics.Process.Start("msedge.exe", str_URL_Edge)
                System.Threading.Thread.Sleep(1000)
                Write_To_CSV("YouTube playlist opened in Edge browser: " & str_URL_Edge)
                'Microsoft.VisualBasic.Interaction.Shell("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe """ & str_URL_Edge & """", Microsoft.VisualBasic.AppWinStyle.Hide)
                'System.Threading.Thread.Sleep(number_Wait_After_Opening_URL * 1000)
            Else
                Write_To_CSV("YouTube playlist not opened in Edge browser as parameter Open_Edge is set to False")
            End If
        Catch ex As Exception
            Write_To_CSV(ex.Message)
        End Try

        Try
            If bool_Open_Chrome Then
                System.Diagnostics.Process.Start("chrome.exe")
                System.Threading.Thread.Sleep(int_Wait_After_Opening_URL * 1000)
                System.Diagnostics.Process.Start("chrome.exe", str_URL_Chrome)
                System.Threading.Thread.Sleep(1000)
                Write_To_CSV("YouTube playlist opened in Chrome browser: " & str_URL_Chrome)
                'Microsoft.VisualBasic.Interaction.Shell("C:\Program Files\Google\Chrome\Application\chrome.exe """ & str_URL_Chrome & """", Microsoft.VisualBasic.AppWinStyle.Hide)
                'System.Threading.Thread.Sleep(number_Wait_After_Opening_URL * 1000)
            Else
                Write_To_CSV("YouTube playlist not opened in Chrome browser as parameter Open_Chrome is set to False")
            End If
        Catch ex As Exception
            Write_To_CSV(ex.Message)
        End Try

        Try
            If bool_Open_Firefox Then
                System.Diagnostics.Process.Start("firefox.exe")
                System.Threading.Thread.Sleep(int_Wait_After_Opening_URL * 1000)
                System.Diagnostics.Process.Start("firefox.exe", str_URL_Firefox)
                System.Threading.Thread.Sleep(1000)
                Write_To_CSV("YouTube playlist opened in Firefox browser: " & str_URL_Firefox)
                'Microsoft.VisualBasic.Interaction.Shell("C:\Program Files\Mozilla Firefox\firefox.exe """ & str_URL_Firefox & """", Microsoft.VisualBasic.AppWinStyle.Hide)
                'System.Threading.Thread.Sleep(number_Wait_After_Opening_URL * 1000	
            Else
                Write_To_CSV("YouTube playlist not opened in Firefox browser as parameter Open_Firefox is set to False")
            End If
        Catch ex As Exception
            Write_To_CSV(ex.Message)
        End Try

        Try
            If bool_Open_Opera Then
                System.Diagnostics.Process.Start("opera.exe")
                System.Threading.Thread.Sleep(int_Wait_After_Opening_URL * 1000)
                System.Diagnostics.Process.Start("opera.exe", str_URL_Opera)
                System.Threading.Thread.Sleep(1000)
                Write_To_CSV("YouTube playlist opened in Opera browser: " & str_URL_Opera)
                'Microsoft.VisualBasic.Interaction.Shell("C:\Users\JAS\AppData\Local\Programs\Opera\opera.exe """ & str_URL_Opera & """", Microsoft.VisualBasic.AppWinStyle.Hide)
                'System.Threading.Thread.Sleep(number_Wait_After_Opening_URL * 1000)
            Else
                Write_To_CSV("YouTube playlist not opened in Opera browser as parameter Open_Opera is set to False")
            End If
        Catch ex As Exception
            Write_To_CSV(ex.Message)
        End Try
    End Sub
    Private Sub Write_To_CSV(text_To_Write)
        Try
            Dim text_File_Path_YouTube_Run_Logs As String = "C:\Users\" & Environment.UserName() & "\Documents\YouTube Run Logs " & DateAndTime.Now().ToString("yyyyMMdd") & ".csv"
            Dim sw As System.IO.StreamWriter
            If System.IO.File.Exists(text_File_Path_YouTube_Run_Logs) Then
                sw = System.IO.File.AppendText(text_File_Path_YouTube_Run_Logs)
            Else
                sw = System.IO.File.CreateText(text_File_Path_YouTube_Run_Logs)
                sw.WriteLine("DateTime,Remarks")
            End If
            sw.WriteLine("""" & DateAndTime.Now.ToString("dd MMM yyyy HH:mm : ss") & """, """ & text_To_Write.Replace("""", """""") & """")
            sw.Flush()
            sw.Close()
        Catch ex As Exception
            str_Error = ex.ToString()
        End Try
    End Sub
    Private Function Get_Parameter_Value(str_Parameter_Name As String) As String
        Dim str_Parameter_Value As String = String.Empty
        Try
            If Not String.IsNullOrEmpty(str_Parameter_Name) Then
                Dim xdoc_Parameters As System.Xml.Linq.XDocument = System.Xml.Linq.XDocument.Parse(str_XML_Parameters.Replace("&", "&amp;"))
                If xdoc_Parameters.Descendants.Any(Function(e) e.Name.LocalName.Trim.ToUpper.Equals(str_Parameter_Name.Trim.ToUpper)) Then
                    str_Parameter_Value = xdoc_Parameters.Descendants.First(Function(e) e.Name.LocalName.Trim.ToUpper.Equals(str_Parameter_Name.Trim.ToUpper)).Value
                Else
                    Write_To_CSV("Parameter not found in XML: " & str_Parameter_Name)
                End If

            Else
                Write_To_CSV("Blank parameter name passed")
            End If
        Catch ex As Exception
            Write_To_CSV(ex.Message)
        End Try
        Return str_Parameter_Value
    End Function
    Private Sub Get_Country_List()
        Dim str_Parameter_Value As String = String.Empty
        Try
            Dim xdoc_Parameters As System.Xml.Linq.XDocument = System.Xml.Linq.XDocument.Parse(str_XML_Parameters.Replace("&", "&amp;"))
            If xdoc_Parameters.Descendants.Any(Function(e) e.Name.LocalName.Trim.ToUpper.Equals("NRI_Country_List_VPN".ToUpper)) Then
                dt_Country_List.Columns.Add("Country")
                For Each xelement_Country As System.Xml.Linq.XElement In xdoc_Parameters.Descendants.First(Function(e) e.Name.LocalName.Trim.ToUpper.Equals("NRI_Country_List_VPN".ToUpper)).Descendants
                    dt_Country_List.Rows.Add(xelement_Country.Value)
                Next
            Else
                Write_To_CSV("Country List not found in XML")
            End If
        Catch ex As Exception
            Write_To_CSV(ex.Message)
        End Try
    End Sub
    Private Sub Connect_VPN()
        Try
            If System.IO.File.Exists(str_File_Path_Nord_VPN) Then
                Dim str_Country As String = String.Empty
                Get_Country_List()
                If dt_Country_List.Rows.Count > 0 Then
                    ' Initialize the random-number generator.
                    Randomize()
                    ' Generate random value between 1 and collection length.
                    Dim int_Row_Index As Integer = CInt(Int(((dt_Country_List.Rows.Count - 1) * Rnd()) + 1))
                    If int_Row_Index >= 0 And int_Row_Index < dt_Country_List.Rows.Count Then
                        str_Country = dt_Country_List.Rows(int_Row_Index).Item(0).ToString
                    Else
                        str_Country = "United States"
                    End If
                    Microsoft.VisualBasic.Interaction.Shell("""" & str_File_Path_Nord_VPN & """ -c -g """ & str_Country & """", Microsoft.VisualBasic.AppWinStyle.Hide)
                    Write_To_CSV("VPN Connected: " & str_Country)
                Else
                    Write_To_CSV("Country list have no rows")
                End If
            Else
                Write_To_CSV("NordVPN exe not found at path: " & str_File_Path_Nord_VPN)
            End If
        Catch ex As Exception
            Write_To_CSV(ex.Message)
        End Try
    End Sub
End Module
