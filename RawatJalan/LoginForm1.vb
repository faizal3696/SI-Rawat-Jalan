Imports System.Data.OleDb
Public Class LoginForm1
    Dim Conn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim LokasiDB As String
    Dim id As New TextBox
    Dim jk As New TextBox

    Public dt As DataTable
    Public ds As DataSet

    Sub Koneksi()
        LokasiDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=sirjdb.accdb"
        Conn = New OleDbConnection(LokasiDB)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub
    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Koneksi()
        If UsernameTextBox.Text = "" Or PasswordTextBox.Text = "" Then
            MsgBox("Harap Isi Username Dan Password !", MsgBoxStyle.Critical, "Error Message")
        Else
            Dim CMD As OleDbCommand
            Dim RD As OleDbDataReader
            Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(PasswordTextBox.Text)
            CMD = New OleDbCommand("Select * From pemakai where nama_pemakai = '" & UsernameTextBox.Text & "' and password = '" & Convert.ToBase64String(byt) & "'", Conn)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("Username Tidak Ditemukan")
                UsernameTextBox.Text = ""
                PasswordTextBox.Text = ""
            Else
                Me.Hide()
                MainMenu.Label3.Text = "Selamat Datang " & RD.GetString(1)
                MainMenu.Label4.Text = RD.GetString(0)
                MainMenu.Show()
            End If

        End If
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

End Class
