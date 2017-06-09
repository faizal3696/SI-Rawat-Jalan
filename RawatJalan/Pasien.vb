Imports System.Data.OleDb

Public Class Pasien
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

        da = New OleDbDataAdapter("Select * from pasien", Conn)
        dt = New DataTable
        da.Fill(dt)
        DataGridView1.DataSource = dt
        'DataGridView1.DataSource = (ds.Tables("pasien"))
        DataGridView1.Columns(0).HeaderText = "Kode Pasien"
        DataGridView1.Columns(1).HeaderText = "Nama Pasien"
        DataGridView1.Columns(2).HeaderText = "Alamat"
        DataGridView1.Columns(3).HeaderText = "Jenis Kelamin"
        DataGridView1.Columns(4).HeaderText = "Umur"
        DataGridView1.Columns(5).HeaderText = "Telepon"

        DataGridView1.AutoResizeColumns()
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Gray

    End Sub
    Sub kodepasien()
        Koneksi()
        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader
        'Dim urutan As Int32
        Dim formatnya As String
        Dim format1 As Int32
        Dim z As Int32
        Dim format2 As Int32
        Dim z2 As Int32
        Dim format3 As Int32
        Dim z3 As Int32
        'Dim newstring As String
        'Dim hitung As Long
        Dim a As Int32

        'CMD = New OleDbCommand("select kode_pasien from pasien Where kode_pasien In(Select Max(kode_pasien)From pasien) Order By kode_pasien Desc", Conn)
        'MsgBox("Select kode_obat From obat where kode_obat = '" & ComboBox1.SelectedItem & "' order by kode_obat desc limit 1")
        CMD = New OleDbCommand("select top 1 kode_pasien from pasien where kode_pasien like '%" & Format(Now, "ddMMyy") & "%' Order By kode_pasien Desc", Conn)
        'Clipboard.SetText("select top 1 kode_pasien from pasien where kode_pasien like '%" & Format(Now, "ddMMyy") & "%' Order By kode_pasien Desc")
        RD = CMD.ExecuteReader
        RD.Read()

        'MsgBox(a + 1)
        If Not RD.HasRows Then
            format1 = Format(Now, "dd")
            format2 = Format(Now, "MM")
            format3 = Format(Now, "yy")
            'MsgBox("format 1 is " & format1)
            'formatnya = Format(Now, "ddMMyy")
            If format1 > 9 And format2 > 9 And format3 > 9 Then
                z = format1
                z2 = format2
                z3 = format3
                formatnya = z & "" & z2 & "" & z3
                TextBox6.Text = formatnya + "01"
            Else
                z = format1.ToString("D2")
                'MsgBox("Z is " & z.ToString("D2"))

                z2 = format2.ToString("D2")
                'MsgBox("Z2 is " & z2.ToString("D2"))
                z3 = format3.ToString("D2")
                'MsgBox("Z3 is " & z3.ToString("D2"))
                formatnya = z.ToString("D2") & "" & z2.ToString("D2") & "" & z3.ToString("D2")
                'MsgBox("formatnya " & formatnya)
                TextBox6.Text = formatnya & "01"
            End If


        Else
            a = RD.GetString(0)
          
            'Dim x As Int32
            'x = Microsoft.VisualBasic.Right(a, 3) + 1
                'MsgBox(x.ToString("D3"))
            'If x > 9 Then
            'urutan = formatnya.ToString() + x.ToString("D2")
            'TextBox6.Text = urutan
            'ElseIf x > 99 Then
            'urutan = formatnya.ToString() + x
            'TextBox6.Text = urutan
            '  Else
            'newstring = Format(Now, "MM") + x.ToString("D3")
            'MsgBox(newstring)
            format1 = Format(Now, "dd")
            If format1 > 9 Then
                TextBox6.Text = a + 1
            Else
                TextBox6.Text = "0" & a + 1
            End If

            ' End If



        End If
        TextBox6.Enabled = False

        Return
    End Sub


    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        kodepasien()
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = True
        Button5.Enabled = True
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True

        TextBox1.Focus()

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        If (RadioButton1.Checked) Then
            jk.Text = "Pria"
        ElseIf (RadioButton2.Checked) Then
            jk.Text = "Wanita"
        Else
            jk.Text = ""
        End If
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or jk.Text = "" Then
            MsgBox("Silahkan Isi Semua Form", MsgBoxStyle.Critical, "Error Message")
        Else
            If MessageBox.Show("Apakah Yakin Akan Menyimpan...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Try
                    Koneksi()



                    Dim CMD As OleDbCommand
                    Koneksi()
                    Dim simpan As String = "insert into pasien (kode_pasien, nama_pasien,alamat,jenis_kelamin,umur,telepon) values ('" & TextBox6.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & jk.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
                    CMD = New OleDbCommand(simpan, Conn)
                    CMD.ExecuteNonQuery()
                    MsgBox("Input data berhasil")
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    TextBox6.Text = ""
                    RadioButton1.Checked = False
                    RadioButton2.Checked = False
                    TextBox6.Text = ""

                    TextBox1.Enabled = False
                    TextBox2.Enabled = False
                    TextBox3.Enabled = False
                    TextBox4.Enabled = False
                    TextBox6.Enabled = False
                    RadioButton1.Enabled = False
                    RadioButton2.Enabled = False
                    loaddata()
                Catch ex As Exception
                    MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
                End Try

            End If
            Button4.Enabled = False
            Button2.Enabled = True
            Button3.Enabled = True
            Button5.Enabled = False
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox6.Text = ""
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        TextBox6.Text = ""

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox6.Enabled = False

        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        Button5.Enabled = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If id.Text = "" Then
            MsgBox("Silahkan Pilih Data yang akan di hapus")
        Else
            If MessageBox.Show("Yakin akan dihapus..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Koneksi()
                Dim CMD As OleDbCommand
                Dim hapus As String = "delete From pasien  where kode_pasien='" & id.Text & "'"
                CMD = New OleDbCommand(hapus, Conn)
                CMD.ExecuteNonQuery()
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox6.Text = ""
                RadioButton1.Checked = False
                RadioButton2.Checked = False
                TextBox1.Enabled = False
                TextBox2.Enabled = False
                TextBox3.Enabled = False
                TextBox4.Enabled = False
                TextBox6.Enabled = False
                RadioButton1.Enabled = False
                RadioButton2.Enabled = False
                MsgBox("Data Berhasil Dihapus", MsgBoxStyle.Information)
                loaddata()
                Button5.Enabled = False
            End If
        End If
    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True

        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        Me.id.Text = DataGridView1.Item(0, i).Value
        Me.TextBox1.Text = DataGridView1.Item(1, i).Value
        Me.TextBox2.Text = DataGridView1.Item(2, i).Value
        If DataGridView1.Item(3, i).Value = "Pria" Then
            Me.RadioButton1.Checked = True
        ElseIf DataGridView1.Item(3, i).Value = "Wanita" Then
            Me.RadioButton2.Checked = True
        End If

        Me.TextBox3.Text = DataGridView1.Item(4, i).Value
        Me.TextBox4.Text = DataGridView1.Item(5, i).Value
        Me.TextBox6.Text = DataGridView1.Item(0, i).Value
        TextBox6.Enabled = False
        Button2.Enabled = True
        Button4.Enabled = False
        Button5.Enabled = True

    End Sub



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If (RadioButton1.Checked) Then
                jk.Text = "Pria"
            ElseIf (RadioButton2.Checked) Then
                jk.Text = "Wanita"
            Else
                jk.Text = ""
            End If
            Koneksi()
            Dim CMD As OleDbCommand
            Dim edit As String
            edit = "update pasien set nama_pasien = '" & TextBox1.Text & "', alamat = '" & TextBox2.Text & "', jenis_kelamin = '" & jk.Text & "', umur = '" & TextBox3.Text & "', telepon = '" & TextBox4.Text & "'  where kode_pasien = '" & id.Text & "'"
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or jk.Text = "" Then
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
            RadioButton1.Checked = False
            RadioButton2.Checked = False

            TextBox1.Enabled = False
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox6.Enabled = False
            RadioButton1.Enabled = False
            RadioButton2.Enabled = False
            Button5.Enabled = False
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
        End Try


    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress


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
            CMD = New OleDbCommand("Select * From pasien where nama_pasien like '%" & TextBox5.Text & "%'", Conn)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("Data Pasien Tidak Ditemukan")
                TextBox5.Focus()
            Else
                da = New OleDbDataAdapter("Select * From pasien where nama_pasien like '%" & TextBox5.Text & "%'", Conn)
                dt = New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
                'DataGridView1.DataSource = (ds.Tables("pasien"))
                DataGridView1.Columns(0).HeaderText = "Kode Pasien"
                DataGridView1.Columns(1).HeaderText = "Nama Pasien"
                DataGridView1.Columns(2).HeaderText = "Alamat"
                DataGridView1.Columns(3).HeaderText = "Jenis Kelamin"
                DataGridView1.Columns(4).HeaderText = "Umur"
                DataGridView1.Columns(5).HeaderText = "Telepon"

                DataGridView1.AutoResizeColumns()
                DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Gray
            End If
        End If
        Button5.Enabled = True
    End Sub


    Private Sub Pasien_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loaddata()
        Button4.Enabled = False
        Button5.Enabled = False
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub
End Class