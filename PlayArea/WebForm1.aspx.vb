Imports Common
Imports System.Data.SqlClient
Public Class WebForm1
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim dc As New SQL

        Dim str As String = "Data Source=localhost;Initial Catalog=ERP;Integrated Security=True"
        SqlDataSource2.ConnectionString = str
        'dc.ConnectionString = str
        'Dim con As New SqlConnection
        'con = dc.ConnectDB()
        ' ASPxGridView1.KeyFieldName = "Address_PK"
        Dim SqlData2 As String = "SELECT * FROM ADDRESS"
        'Dim tbl As DataTable = dc.ExecuteSQLQuery(con, strString, "tbl")
        ''dc.CloseConnection(con)
        ''ASPxGridView1.DataSource = tbl
        'SQLDB.ConnectionString = str
        'ASPxGridView1.DataSource = tbl

        SqlDataSource2.SelectCommand = SqlData2



    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ASPxGridView2.Columns.Clear()
        ASPxGridView2.DataSource = Nothing
        ASPxGridView2.DataSourceID = Nothing
        Dim SqlData2 As String = "SELECT * FROM PERSON"

        SqlDataSource2.SelectCommand = SqlData2
        'Dim tbl As DataTable = dc.ExecuteSQLQuery(con, strString, "tbl")
        ''dc.CloseConnection(con)
        ''ASPxGridView1.DataSource = tbl
        'SQLDB.ConnectionString = str
        'ASPxGridView1.DataSource = tbl
        ASPxGridView2.AutoGenerateColumns = True
        ASPxGridView2.DataSource = SqlDataSource2

    End Sub
End Class