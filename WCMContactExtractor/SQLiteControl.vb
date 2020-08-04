Imports System.Data.SQLite

Public Class SQLiteControl
    Private ReadOnly location As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
    Private ReadOnly fileName As String = "WCMDB.db"
    Private ReadOnly fullPath As String = System.IO.Path.Combine(location, fileName)
    Private ReadOnly connectionString As String = String.Format("Data Source = {0}", fullPath)

    Public DBCon As New SQLiteConnection(connectionString)

    Private DBCmd As SQLiteCommand

    ' DB DATA
    Public DBDA As SQLiteDataAdapter
    Public DBDT As DataTable


    ' QUERY PARAMETERS
    Public Params As New List(Of SQLiteParameter)


    'QUERY STATISTICS
    Public RecordCount As Integer
    Public Exception As String

    Public Sub New()
    End Sub


    'ALLOW CONNECTION STRING OVERRIDE
    Public Sub New(ConnectionString As String)
        DBCon = New SQLiteConnection(ConnectionString)
    End Sub


    'EXECUTE QUERY SUB  
    Public Sub ExecQuery(Query As String, Optional ReturnIdentity As Boolean = False)
        'RESET QUERY STATS
        RecordCount = 0
        Exception = ""

        Try
            DBCon.Open()

            ' CREATE DB COMMAND
            DBCmd = New SQLiteCommand(Query, DBCon)

            ' LOAD PARAMS INTO DB COMMAND
            Params.ForEach(Sub(p) DBCmd.Parameters.Add(p))

            'clear param list
            Params.Clear()


            'EXECUTE COMMAND AND FILL DATASET
            DBDT = New DataTable
            DBDA = New SQLiteDataAdapter(DBCmd)
            RecordCount = DBDA.Fill(DBDT)

            If ReturnIdentity = True Then
                Dim ReturnQuery As String = "SELECT @@IDENTITY As LastID;"
                DBCmd = New SQLiteCommand(ReturnQuery, DBCon)
                DBDT = New DataTable
                DBDA = New SQLiteDataAdapter(DBCmd)
                RecordCount = DBDA.Fill(DBDT)
            End If
        Catch ex As Exception
            Exception = "ExecQuery Error: " & vbNewLine & ex.Message
        Finally
            'CLOSE CONNECTION
            If DBCon.State = ConnectionState.Open Then DBCon.Close()

        End Try
    End Sub


    'ADD PARAMS

    Public Sub AddParam(Name As String, Value As Object)
        Dim NewParam As New SQLiteParameter(Name, Value)
        Params.Add(NewParam)
    End Sub

    'ERROR CHECKING

    Public Function HasExpetion(Optional Report As Boolean = False) As Boolean
        If String.IsNullOrEmpty(Exception) Then Return False

        If Report = True Then MsgBox(Exception, MsgBoxStyle.Critical, "EXCEPTION")

        Return True

    End Function
End Class
