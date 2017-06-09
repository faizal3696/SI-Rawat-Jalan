Imports System.Data.OleDb
Public Class Obat
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

        da = New OleDbDataAdapter("Select * from obat", Conn)
        dt = New DataTable
        da.Fill(dt)
        DataGridView1.DataSource = dt
        'DataGridView1.DataSource = (ds.Tables("pasien"))
        DataGridView1.Columns(0).HeaderText = "Kode Obat"
        DataGridView1.Columns(1).HeaderText = "Nama Obat"
        DataGridView1.Columns(2).HeaderText = "Kategori"
        DataGridView1.Columns(3).HeaderText = "Jenis"
        DataGridView1.Columns(4).HeaderText = "Harga"
        DataGridView1.Columns(4).DefaultCellStyle.Format = "##,0"
        DataGridView1.Columns(5).HeaderText = "Jumlah"
        DataGridView1.AutoResizeColumns()
        DataGridView1.Columns(5).Width = 83
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Gray

    End Sub
    Sub refreshdata()
        Koneksi()
        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader
        Dim urutan As String
        Dim hitung As Int32
        CMD = New OleDbCommand("Select top 1 kode_obat From obat where kode_obat like '%" & ComboBox1.SelectedItem & "%' order by kode_obat desc", Conn)
        'MsgBox("Select kode_obat From obat where kode_obat = '" & ComboBox1.SelectedItem & "' order by kode_obat desc limit 1")
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            urutan = ComboBox1.SelectedItem & "01"
            TextBox1.Text = urutan

        Else
            Dim substrdata As String
            substrdata = ComboBox1.SelectedItem
            'MsgBox(RD.GetString(0))
            'MsgBox(substrdata.Length - 1)
            hitung = Microsoft.VisualBasic.Right(RD.GetString(0), 1) + 1
            'MsgBox(Microsoft.VisualBasic.Right(RD.GetString(0), 5))
            'MsgBox("000" & hitung)
            If hitung > 9 Then
                urutan = ComboBox1.SelectedItem & "" & hitung & ""
                'Dim substrdata As String
                'substrdata = ComboBox1.SelectedItem
                'Console.WriteLine(substrdata)
                'Data = substrdata.Substring(0, substrdata.Length)
                'Console.WriteLine(Data)
                'MsgBox(RD.HasRows)
                TextBox1.Text = urutan
            Else
                'MsgBox(hitung.ToString("D2"))

                urutan = ComboBox1.SelectedItem + hitung.ToString("D2")
                TextBox1.Text = urutan
            End If

        End If
            TextBox1.Enabled = False
            Return
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = True
        Button5.Enabled = True
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True

        ComboBox1.Focus()
    End Sub

    Private Sub Obat_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Koneksi()
        loaddata()
        Button4.Enabled = False
        Button5.Enabled = False
    

        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader
        'CMD = New OleDbCommand("Select * From poli", Conn)
        'RD = CMD.ExecuteReader

        'dt.Load(RD)

        'ComboBox1.ValueMember = "nama_poli"
        'ComboBox1.DisplayMember = "nama_poli"
        'ComboBox1.DataSource = dt

        CMD = New OleDbCommand("select * FROM poli", Conn)
        RD = CMD.ExecuteReader
        ComboBox1.Items.Add("")
        Do While RD.Read
            ComboBox1.Items.Add(RD.Item(1))
        Loop
        TextBox1.Text = ""
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And TextBox3.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
        If e.KeyChar = Chr(13) Then
            TextBox3.Text = FormatNumber(TextBox3.Text, 0) ' ini format angka
        End If
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

        If IsNumeric(TextBox3.Text) Then

            Dim temp As Integer = TextBox3.Text

            TextBox3.Text = Format(temp, "N")

            TextBox3.SelectionStart = TextBox3.TextLength - 3


        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.SelectedItem = ""
        ComboBox2.SelectedItem = ""


        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        Button5.Enabled = False

        TextBox1.Text = ""
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Button5.Enabled = False

        If TextBox1.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.SelectedItem = "" Or ComboBox2.SelectedItem = "" Then
            MsgBox("Silahkan Isi Semua Form", MsgBoxStyle.Critical, "Error Message")
        Else
            If MessageBox.Show("Apakah Yakin Akan Menyimpan...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Try
                    Koneksi()



                    Dim CMD As OleDbCommand
                    Koneksi()
                    Dim simpan As String = "insert into obat (kode_obat, nama_obat,jenis,kategori_obat,harga,jumlah) values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox2.SelectedItem & "','" & ComboBox1.SelectedItem & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"

                    CMD = New OleDbCommand(simpan, Conn)
                    CMD.ExecuteNonQuery()
                    MsgBox("Input data berhasil")
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    ComboBox1.SelectedItem = ""
                    ComboBox2.SelectedItem = ""
                    TextBox1.Text = ""

                    TextBox1.Enabled = False
                    TextBox2.Enabled = False
                    TextBox3.Enabled = False
                    TextBox4.Enabled = False
                    ComboBox1.Enabled = False
                    ComboBox2.Enabled = False

                    loaddata()
                Catch ex As Exception
                    MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
                End Try

            End If
            Button2.Enabled = True
            Button3.Enabled = True
            Button4.Enabled = False

        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If id.Text = "" Then
            MsgBox("Silahkan Pilih Data yang akan di hapus")
        Else
            If MessageBox.Show("Yakin akan dihapus..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Koneksi()
                Dim CMD As OleDbCommand
                'MsgBox(id.Text)
                Dim hapus As String = "delete From obat where kode_obat='" & id.Text & "'"
                CMD = New OleDbCommand(hapus, Conn)
                CMD.ExecuteNonQuery()

                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                ComboBox1.SelectedItem = ""
                ComboBox2.SelectedItem = ""
                TextBox1.Text = ""

                TextBox1.Enabled = False
                TextBox2.Enabled = False
                TextBox3.Enabled = False
                TextBox4.Enabled = False
                ComboBox1.Enabled = False
                ComboBox2.Enabled = False

                MsgBox("Data Berhasil Dihapus", MsgBoxStyle.Information)
                loaddata()

            End If
        End If
        Button5.Enabled = False

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        refreshdata()
        'TextBox1.Text = ""
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True


        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        Me.id.Text = DataGridView1.Item(0, i).Value
        Me.ComboBox2.SelectedItem = DataGridView1.Item(3, i).Value
        'MsgBox("data2 " & DataGridView1.Item(2, i).Value)
        Me.ComboBox1.SelectedItem = DataGridView1.Item(2, i).Value
        Me.TextBox1.Text = DataGridView1.Item(0, i).Value
        Me.TextBox2.Text = DataGridView1.Item(1, i).Value
        Me.TextBox3.Text = DataGridView1.Item(4, i).Value
        Me.TextBox4.Text = DataGridView1.Item(5, i).Value
        Button2.Enabled = True
        Button4.Enabled = False
        Button5.Enabled = True
    End Sub


  

    Private Sub TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And TextBox4.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
        If e.KeyChar = Chr(13) Then
            TextBox4.Text = FormatNumber(TextBox4.Text, 0) ' ini format angka
        End If
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Koneksi()
            Dim CMD As OleDbCommand
            Dim edit As String
            'nama_dokter,spesialis,alamat,telpon,kode_poli,tarif
            edit = "update obat set nama_obat = '" & TextBox2.Text & "', jenis = '" & ComboBox2.SelectedItem & "', kategori_obat = '" & ComboBox1.SelectedItem & "', harga = '" & TextBox3.Text & "', jumlah = '" & TextBox4.Text & "'  where kode_obat = '" & id.Text & "'"

            'MsgBox(edit)

            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.SelectedItem = "" Or ComboBox2.SelectedItem = "" Then
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
            TextBox4.Text = ""
            ComboBox1.SelectedItem = ""
            ComboBox2.SelectedItem = ""
            TextBox1.Text = ""

            TextBox1.Enabled = False
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            ComboBox1.Enabled = False
            ComboBox2.Enabled = False

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
        End Try
        Button5.Enabled = False

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If TextBox5.Text = "" Then
            loaddata()

        Else
            Koneksi()
            Dim CMD As OleDbCommand
            Dim RD As OleDbDataReader
            CMD = New OleDbCommand("Select * From obat where nama_obat like '%" & TextBox5.Text & "%'", Conn)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("Data Pasien Tidak Ditemukan")
                TextBox5.Focus()
            Else
                da = New OleDbDataAdapter("Select * From obat where nama_obat like '%" & TextBox5.Text & "%'", Conn)
                dt = New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
                'DataGridView1.DataSource = (ds.Tables("pasien"))
                DataGridView1.Columns(0).HeaderText = "Kode Obat"
                DataGridView1.Columns(1).HeaderText = "Nama Obat"
                DataGridView1.Columns(2).HeaderText = "Jenis"
                DataGridView1.Columns(3).HeaderText = "Kategori"
                DataGridView1.Columns(4).HeaderText = "Harga"
                DataGridView1.Columns(4).DefaultCellStyle.Format = "##,0"
                DataGridView1.Columns(5).HeaderText = "Jumlah"
                DataGridView1.AutoResizeColumns()
                DataGridView1.Columns(5).Width = 83
                DataGridView1.AutoResizeColumns()
                DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Gray
            End If
        End If
        Button5.Enabled = True
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        If TextBox5.Text = "" Then
            loaddata()
        End If
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class