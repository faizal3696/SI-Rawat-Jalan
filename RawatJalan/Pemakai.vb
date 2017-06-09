Imports System.Data.OleDb
Public Class pemakai
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
    Sub loaddata()
        Koneksi()

        da = New OleDbDataAdapter("Select * from pemakai", Conn)
        dt = New DataTable
        da.Fill(dt)
        DataGridView1.DataSource = dt
        'DataGridView1.DataSource = (ds.Tables("pasien"))
        DataGridView1.Columns(0).HeaderText = "Kode"
        DataGridView1.Columns(1).HeaderText = "Nama"

        DataGridView1.Columns(2).Visible = False
        DataGridView1.Columns(3).HeaderText = "Status"
        DataGridView1.AutoResizeColumns()
        DataGridView1.Columns(1).Width = 180
        DataGridView1.Columns(3).Width = 150
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Gray

    End Sub
    Sub refreshdata()
        Koneksi()
        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader
        Dim urutan As String
        Dim hitung As Long
        CMD = New OleDbCommand("Select top 1 kode_pemakai From pemakai where kode_pemakai like '%" & ComboBox1.SelectedItem & "%' order by kode_pemakai desc", Conn)
        'MsgBox("Select kode_obat From obat where kode_obat = '" & ComboBox1.SelectedItem & "' order by kode_obat desc limit 1")
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            urutan = ComboBox1.SelectedItem & "1"
            TextBox1.Text = urutan

        Else
            Dim substrdata As String
            substrdata = ComboBox1.SelectedItem
            'MsgBox(RD.GetString(0))
            'MsgBox(substrdata.Length - 1)
            hitung = Microsoft.VisualBasic.Right(RD.GetString(0), 1) + 1
            'MsgBox(Microsoft.VisualBasic.Right(RD.GetString(0), 5))
            'MsgBox("000" & hitung)
            urutan = ComboBox1.SelectedItem & "" & hitung & ""
            'Dim substrdata As String
            'substrdata = ComboBox1.SelectedItem
            'Console.WriteLine(substrdata)
            'Data = substrdata.Substring(0, substrdata.Length)
            'Console.WriteLine(Data)
            'MsgBox(RD.HasRows)
            TextBox1.Text = urutan
        End If
        TextBox1.Enabled = False
        Return
    End Sub

    Private Sub pemakai_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Koneksi()
        loaddata()
        Button4.Enabled = False
        Button5.Enabled = False
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        ComboBox1.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
       

        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        Me.id.Text = DataGridView1.Item(0, i).Value
        Me.ComboBox1.SelectedItem = DataGridView1.Item(3, i).Value
        Me.TextBox1.Text = DataGridView1.Item(0, i).Value
        Me.TextBox2.Text = DataGridView1.Item(1, i).Value
        Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(DataGridView1.Item(2, i).Value)
        Dim s As String = DataGridView1.Item(2, i).Value
        'Me.TextBox3.Text = Convert.ToString(DataGridView1.Item(2, i).Value)

        Button2.Enabled = True
        Button4.Enabled = False
        Button5.Enabled = True
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        refreshdata()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = True
        Button5.Enabled = True
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""

        TextBox1.Enabled = False
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        ComboBox1.Enabled = True

        ComboBox1.Focus()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.SelectedItem = ""

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        ComboBox1.Enabled = False
        Button5.Enabled = False

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click


        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.SelectedItem = "" Then
            MsgBox("Silahkan Isi Semua Form", MsgBoxStyle.Critical, "Error Message")
        Else
            If MessageBox.Show("Apakah Yakin Akan Menyimpan...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Try


                    Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(TextBox3.Text)


                    Dim CMD As OleDbCommand
                    Dim simpan As String = "insert into pemakai values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & Convert.ToBase64String(byt) & "','" & ComboBox1.SelectedItem & "')"

                    'MsgBox(simpan)
                    'Clipboard.SetText(CStr(simpan))
                    Koneksi()
                    CMD = New OleDbCommand(simpan, Conn)
                    'Clipboard.SetText(CStr(CMD))
                    CMD.ExecuteNonQuery()
                    MsgBox("Input data berhasil")
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    ComboBox1.SelectedItem = ""

                    TextBox1.Enabled = False
                    TextBox2.Enabled = False
                    TextBox3.Enabled = False
                    ComboBox1.Enabled = False
                    TextBox1.Text = ""
                    loaddata()
                Catch ex As Exception
                    MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
                End Try

            End If

        End If
        Button5.Enabled = False
        Button2.Enabled = True
        Button2.Enabled = True
        Button4.Enabled = False
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Koneksi()
            Dim CMD As OleDbCommand
            Dim edit As String
            Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(TextBox3.Text)

            'edit = "update pemakai set password = '" & Convert.ToBase64String(byt) & "',status = '" & ComboBox1.SelectedItem & "' where kode_pemakai = '" & id.Text & "'"
            edit = "UPDATE    pemakai SET [password] = '" & Convert.ToBase64String(byt) & "', status = '" & ComboBox1.SelectedItem & "' WHERE     (kode_pemakai = '" & id.Text & "')"
            'MsgBox(edit)
            'Clipboard.SetText(CStr(edit))
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.SelectedItem = "" Then
                MsgBox("Data Gagal Diupdate", MsgBoxStyle.Critical, "Error Message")
            Else
                CMD = New OleDbCommand(edit, Conn)
                CMD.ExecuteNonQuery()
                MsgBox("Data Berhasil diUpdate")
                loaddata()

            End If
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            ComboBox1.SelectedItem = ""


            TextBox1.Enabled = False
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            ComboBox1.Enabled = False
            Button5.Enabled = False

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If id.Text = "" Then
            MsgBox("Silahkan Pilih Data yang akan di hapus")
        Else
            If MessageBox.Show("Yakin akan dihapus..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Koneksi()
                Dim CMD As OleDbCommand
                Dim hapus As String = "delete From pemakai  where kode_pemakai ='" & id.Text & "'"
                CMD = New OleDbCommand(hapus, Conn)
                CMD.ExecuteNonQuery()
                ComboBox1.SelectedItem = ""
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""

                ComboBox1.Enabled = False
                TextBox1.Enabled = False
                TextBox2.Enabled = False
                TextBox3.Enabled = False
                
                MsgBox("Data Berhasil Dihapus", MsgBoxStyle.Information)
                loaddata()

            End If
        End If
        Button5.Enabled = False
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub
End Class