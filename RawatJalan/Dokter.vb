Imports System.Data.OleDb
Public Class Dokter
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

    Sub kodedokter()
        Koneksi()
        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader
        Dim urutan As Int32
        Dim formatnya As Int32
        Dim newstring As String
        'Dim hitung As Long
        Dim a As String

        CMD = New OleDbCommand("select kode_dokter from dokter Where kode_dokter In(Select Max(kode_dokter)From dokter)Order By kode_dokter Desc", Conn)
        'MsgBox("Select kode_obat From obat where kode_obat = '" & ComboBox1.SelectedItem & "' order by kode_obat desc limit 1")
        RD = CMD.ExecuteReader
        RD.Read()

        'MsgBox(a + 1)
        If Not RD.HasRows Then
            formatnya = Format(Now, "MM")

            If formatnya > 9 Then
                urutan = formatnya.ToString() + "001"
                'MsgBox(urutan)
                TextBox6.Text = urutan

            Else
                'urutan = formatnya.ToString("D2") + "001"
                'urutan = formatnya + numbering.ToString("D3")
                'MsgBox(formatnya.ToString("D2"))
                newstring = Format(Now, "MM") + "001"
                'MsgBox("data string " + newstring)
                TextBox6.Text = newstring
            End If

            'MsgBox(urutan.ToString("D2"))


        Else

            a = RD.GetString(0)
            Dim x As Int32
            x = Microsoft.VisualBasic.Right(a, 3) + 1
            'MsgBox(x.ToString("D3"))
            If x > 9 Then
                urutan = formatnya.ToString() + x.ToString("D2")
                TextBox6.Text = urutan
            ElseIf x > 99 Then
                urutan = formatnya.ToString() + x
                TextBox6.Text = urutan
            Else
                newstring = Format(Now, "MM") + x.ToString("D3")
                'MsgBox(newstring)
                TextBox6.Text = newstring
            End If



        End If
        TextBox6.Enabled = False

        Return
    End Sub
    Sub loaddata()
        da = New OleDbDataAdapter("Select * from dokter order by kode_dokter asc", Conn)
        dt = New DataTable
        da.Fill(dt)
        DataGridView1.DataSource = dt
        'DataGridView1.DataSource = (ds.Tables("pasien"))
        DataGridView1.Columns(0).HeaderText = "Kode Dokter"
        DataGridView1.Columns(1).HeaderText = "Nama Dokter"
        DataGridView1.Columns(2).HeaderText = "Spesialis"
        DataGridView1.Columns(3).HeaderText = "Alamat"
        DataGridView1.Columns(4).HeaderText = "Telpon"
        DataGridView1.Columns(5).HeaderText = "Kode Poli"
        DataGridView1.Columns(6).HeaderText = "Tarif"
        DataGridView1.Columns(6).DefaultCellStyle.Format = "##,0"
        DataGridView1.AutoResizeColumns()
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Gray

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        kodedokter()

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

        ComboBox1.Focus()

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Button5.Enabled = False
        
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.SelectedItem = "" Then
            MsgBox("Silahkan Isi Semua Form", MsgBoxStyle.Critical, "Error Message")
        Else
            If MessageBox.Show("Apakah Yakin Akan Menyimpan...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Try
                    Koneksi()
                    Dim convert As Int32
                    Dim result As String
                    convert = ComboBox1.SelectedIndex
                    If convert > 9 Then
                        result = ComboBox1.SelectedIndex
                    Else
                        result = convert.ToString("D2")
                    End If

                    'MsgBox(ComboBox1.SelectedItem)
                    'MsgBox(convert.ToString("D2"))
                    Dim CMD As OleDbCommand
                    Dim simpan As String = "insert into dokter (kode_dokter, nama_dokter,spesialis,alamat,telpon,kode_poli,tarif) values ('" & TextBox6.Text & "','" & TextBox1.Text & "','" & ComboBox1.SelectedItem & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & result & "','" & TextBox4.Text & "')"
                    ' Clipboard.SetText(CStr(simpan))
                     CMD = New OleDbCommand(simpan, Conn)
                    CMD.ExecuteNonQuery()
                    MsgBox("Input data berhasil")
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    TextBox6.Text = ""
                    ComboBox1.SelectedItem = ""


                    TextBox1.Enabled = False
                    TextBox2.Enabled = False
                    TextBox3.Enabled = False
                    TextBox4.Enabled = False
                    TextBox6.Enabled = False
                    ComboBox1.Enabled = False

                    loaddata()
                Catch ex As Exception
                    MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
                End Try

            End If
            Button4.Enabled = False
            Button2.Enabled = True
            Button3.Enabled = True

        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox6.Text = ""
        ComboBox1.SelectedItem = ""


        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox6.Enabled = False
        ComboBox1.Enabled = False
        Button5.Enabled = False

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If id.Text = "" Then
            MsgBox("Silahkan Pilih Data yang akan di hapus")
        Else
            If MessageBox.Show("Yakin akan dihapus..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Koneksi()
                Dim CMD As OleDbCommand
                Dim hapus As String = "delete From dokter  where kode_dokter='" & id.Text & "'"
                CMD = New OleDbCommand(hapus, Conn)
                CMD.ExecuteNonQuery()
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox6.Text = ""
                ComboBox1.SelectedItem = ""


                TextBox1.Enabled = False
                TextBox2.Enabled = False
                TextBox3.Enabled = False
                TextBox4.Enabled = False
                TextBox6.Enabled = False
                ComboBox1.Enabled = False

                MsgBox("Data Berhasil Dihapus", MsgBoxStyle.Information)
                loaddata()

            End If
        End If
        Button5.Enabled = False
    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        ComboBox1.Enabled = True


        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        Me.id.Text = DataGridView1.Item(0, i).Value
        Me.ComboBox1.SelectedItem = DataGridView1.Item(2, i).Value
        Me.TextBox1.Text = DataGridView1.Item(1, i).Value
        Me.TextBox2.Text = DataGridView1.Item(3, i).Value
        Me.TextBox3.Text = DataGridView1.Item(4, i).Value
        Me.TextBox4.Text = DataGridView1.Item(6, i).Value
        Me.TextBox6.Text = DataGridView1.Item(0, i).Value
        Button2.Enabled = True
        Button4.Enabled = False
        Button5.Enabled = True

    End Sub



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Koneksi()
            Dim CMD As OleDbCommand
            Dim edit As String
            'nama_dokter,spesialis,alamat,telpon,kode_poli,tarif
            Dim convert As Int32
            convert = ComboBox1.SelectedIndex
            Dim result As String
            If convert > 9 Then
                result = ComboBox1.SelectedIndex
            Else
                result = convert.ToString("D2")
            End If
            edit = "update dokter set nama_dokter = '" & TextBox1.Text & "', spesialis = '" & ComboBox1.SelectedItem & "', alamat = '" & TextBox2.Text & "', telpon = '" & TextBox3.Text & "', kode_poli = '" & result & "', tarif = '" & TextBox4.Text & "'  where kode_dokter = '" & id.Text & "'"
            'Clipboard.SetText(CStr(edit))
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.SelectedItem = "" Then
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
            TextBox6.Text = ""
            ComboBox1.SelectedItem = ""


            TextBox1.Enabled = False
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox6.Enabled = False
            ComboBox1.Enabled = False

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
        End Try


    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        If TextBox5.Text = "" Then
            loaddata()
        End If

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If TextBox5.Text = "" Then
            loaddata()

        Else
            Koneksi()
            Dim CMD As OleDbCommand
            Dim RD As OleDbDataReader
            CMD = New OleDbCommand("Select * From dokter where nama_dokter like '%" & TextBox5.Text & "%'", Conn)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("Data Pasien Tidak Ditemukan")
                TextBox5.Focus()
            Else
                da = New OleDbDataAdapter("Select * From dokter where nama_dokter like '%" & TextBox5.Text & "%'", Conn)
                dt = New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
                'DataGridView1.DataSource = (ds.Tables("pasien"))
                DataGridView1.Columns(0).HeaderText = "Kode Dokter"
                DataGridView1.Columns(1).HeaderText = "Nama Dokter"
                DataGridView1.Columns(2).HeaderText = "Spesialis"
                DataGridView1.Columns(3).HeaderText = "Alamat"
                DataGridView1.Columns(4).HeaderText = "Telpon"
                DataGridView1.Columns(5).HeaderText = "Kode Poli"
                DataGridView1.Columns(6).HeaderText = "Tarif"
                DataGridView1.Columns(6).DefaultCellStyle.Format = "##,0"
                DataGridView1.AutoResizeColumns()
                DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Gray
            End If
        End If
        Button5.Enabled = True
    End Sub

    Private Sub Dokter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Koneksi()
        loaddata()
        Button4.Enabled = False
        Button5.Enabled = False
        da = New OleDbDataAdapter("Select * From poli", Conn)
        dt = New DataTable

        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader
        'CMD = New OleDbCommand("Select * From poli", Conn)
        'RD = CMD.ExecuteReader

        'dt.Load(RD)

        'ComboBox1.ValueMember = "nama_poli"
        'ComboBox1.DisplayMember = "nama_poli"
        'ComboBox1.DataSource = dt

        CMD = New OleDbCommand("select * FROM poli order by kode_poli asc", Conn)
        RD = CMD.ExecuteReader
        ComboBox1.Items.Add("")
        Do While RD.Read
            ComboBox1.Items.Add(RD.Item(1))
        Loop
    End Sub


    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

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


        If IsNumeric(TextBox4.Text) Then

            Dim temp As Integer = TextBox4.Text

            TextBox4.Text = Format(temp, "N")

            TextBox4.SelectionStart = TextBox4.TextLength - 3


        End If
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub
End Class
