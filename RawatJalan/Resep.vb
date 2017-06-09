Imports System.Data.OleDb
Public Class Resep
    Dim Conn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim LokasiDB As String
    Public dt As DataTable
    Public ds As DataSet
    Public subtotal As Double = 0
    Sub Koneksi()
        LokasiDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=sirjdb.accdb"
        Conn = New OleDbConnection(LokasiDB)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub
    Sub loadpendaftaran()
        Koneksi()

        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader

        CMD = New OleDbCommand("Select pendaftaran.nomor_daftar, pasien.nama_pasien From pendaftaran left join pasien on pasien.kode_pasien = pendaftaran.kode_pasien where nomor_daftar like '%" & Format(Now, "yyMMdd") & "%'", Conn)
        'Clipboard.SetText("Select * From pendaftaran left join pasien on pasien.kode_pasien = pendaftaran.kode_pasien where nomor_daftar like '%" & Format(Now, "yyMMdd") & "%'")
        RD = CMD.ExecuteReader
        ComboBox1.Items.Add("")
        Do While RD.Read
            ComboBox1.Items.Add(RD.Item(0) & "-" & RD.Item(1))
        Loop
    End Sub
    Private Sub Resep_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadpendaftaran()
        TextBox6.Text = Format(Now, "dd-MM-yyyy")
        DataGridView1.Enabled = False
        DataGridView1.Columns(2).DefaultCellStyle.Format = "N2"
        DataGridView1.Columns(4).DefaultCellStyle.Format = "N2"
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem = "" Then

        Else
            Button4.Enabled = True

            TextBox9.Text = ""
            TextBox11.Text = ""
            TextBox12.Text = ""
            DataGridView1.Rows.Clear()

            Dim klm As String = ComboBox1.SelectedItem
            Dim arr() As String = klm.Split("-")
            Dim kodedkt As New TextBox
            Dim kodepasien As New TextBox
            Dim kodepoli As New TextBox
            Dim kategoriobat As New TextBox


            Dim CMD As OleDbCommand
            Dim RD As OleDbDataReader

            CMD = New OleDbCommand("Select top 1 kode_dokter, kode_pasien, kode_poli From pendaftaran where nomor_daftar = '" & arr.First() & "'", Conn)
            Clipboard.SetText("Select * From pendaftaran left join pasien on pasien.kode_pasien = pendaftaran.kode_pasien where nomor_daftar like '%" & Format(Now, "yyMMdd") & "%'")
            RD = CMD.ExecuteReader
            Do While RD.Read
                kodedkt.Text = RD.Item(0)
                kodepasien.Text = RD.Item(1)
                kodepoli.Text = RD.Item(2)
            Loop

            Dim CMD1 As OleDbCommand
            Dim RD1 As OleDbDataReader

            CMD1 = New OleDbCommand("Select top 1 kode_dokter, nama_dokter From dokter where kode_dokter = '" & kodedkt.Text & "'", Conn)
            'Clipboard.SetText("Select * From pendaftaran left join pasien on pasien.kode_pasien = pendaftaran.kode_pasien where nomor_daftar like '%" & Format(Now, "yyMMdd") & "%'")
            RD1 = CMD1.ExecuteReader
            Do While RD1.Read
                TextBox1.Text = RD1.Item(0)
                TextBox8.Text = RD1.Item(1)
            Loop

            Dim CMD2 As OleDbCommand
            Dim RD2 As OleDbDataReader

            CMD2 = New OleDbCommand("Select top 1 kode_pasien, nama_pasien From pasien where kode_pasien = '" & kodepasien.Text & "'", Conn)
            Clipboard.SetText("Select top 1 kode_pasien, nama_pasien From pasien where kode_pasien = '" & kodepasien.Text & "'")
            RD2 = CMD2.ExecuteReader
            Do While RD2.Read
                TextBox2.Text = RD2.Item(0)
                TextBox7.Text = RD2.Item(1)
            Loop


            Dim CMD3 As OleDbCommand
            Dim RD3 As OleDbDataReader

            CMD3 = New OleDbCommand("Select top 1 kode_poli, nama_poli From poli where kode_poli = '" & kodepoli.Text & "'", Conn)
            'Clipboard.SetText("Select * From pendaftaran left join pasien on pasien.kode_pasien = pendaftaran.kode_pasien where nomor_daftar like '%" & Format(Now, "yyMMdd") & "%'")
            RD3 = CMD3.ExecuteReader
            Do While RD3.Read
                TextBox3.Text = RD3.Item(0)
                TextBox4.Text = RD3.Item(1)
                kategoriobat.Text = RD3.Item(1)
            Loop


            Dim CMD4 As OleDbCommand
            Dim RD4 As OleDbDataReader

            CMD4 = New OleDbCommand("Select kode_obat, nama_obat From obat where kategori_obat = '" & kategoriobat.Text & "'", Conn)
            Clipboard.SetText("Select kode_obat, nama_obat From obat where kategori_obat = '" & kategoriobat.Text & "'")
            RD4 = CMD4.ExecuteReader
            ListBox1.Items.Clear()
            Do While RD4.Read
                ListBox1.Items.Add(RD4.Item(1))
            Loop
            ListBox1.Refresh()
        End If
    End Sub

    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick
        DataGridView1.Enabled = True
        
        Dim a As String
        a = ListBox1.SelectedItem
        Dim i As Integer = 0

        Dim CMD4 As OleDbCommand
        Dim RD4 As OleDbDataReader

        CMD4 = New OleDbCommand("Select kode_obat, nama_obat, harga From obat where nama_obat = '" & a & "'", Conn)
        'Clipboard.SetText("Select kode_obat, nama_obat From obat where nama_obat = '" & a & "'")
        RD4 = CMD4.ExecuteReader

        Do While RD4.Read
            DataGridView1.Rows.Add(New String() {RD4.Item(0), RD4.Item(1), RD4.Item(2), "", RD4.Item(2)})
            DataGridView1.Columns(2).DefaultCellStyle.Format = "##,0"
            DataGridView1.Columns(4).DefaultCellStyle.Format = "##,0"

        Loop

        DataGridView1.Columns(0).ReadOnly = True
        DataGridView1.Columns(1).ReadOnly = True
        DataGridView1.Columns(2).ReadOnly = True
        DataGridView1.Columns(3).ReadOnly = False
        DataGridView1.Columns(4).ReadOnly = True


    End Sub

    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Try
            Dim i = DataGridView1.CurrentRow.Index
            Dim CMD9 As OleDbCommand
            Dim RD9 As OleDbDataReader
            CMD9 = New OleDbCommand("Select jumlah From obat where kode_obat = '" & DataGridView1.Item(0, i).Value & "'", Conn)
            'Clipboard.SetText("Select kode_obat, nama_obat From obat where nama_obat = '" & a & "'")
            RD9 = CMD9.ExecuteReader
            Do While RD9.Read
                If CDbl(DataGridView1.Item(3, i).Value) > RD9.Item(0) Then
                    MsgBox("Stok Obat Kurang ", MsgBoxStyle.Critical, "Error Message")
                    Button4.Enabled = False
                Else
                    DataGridView1.Item(4, i).Value = CDbl(DataGridView1.Item(2, i).Value) * CDbl(DataGridView1.Item(3, i).Value)
                End If

            Loop

            
            For Each row As DataGridViewRow In Me.DataGridView1.Rows
                If (Not row.IsNewRow) Then
                    subtotal = subtotal + row.Cells(4).Value.ToString
                End If
            Next

            TextBox12.Text = subtotal
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub DataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint

        Dim strRowNumber As String = (e.RowIndex + 1).ToString

        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)

        If DataGridView1.RowHeadersWidth < CInt((size.Width + 20)) Then
            DataGridView1.RowHeadersWidth = CInt((size.Width + 20))
        End If

        Dim b As Brush = SystemBrushes.ControlText


        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, _
                               e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) _
                                                         / 2))

    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged
        If IsNumeric(TextBox12.Text) Then

            Dim temp As Integer = TextBox12.Text

            TextBox12.Text = Format(temp, "N")

            TextBox9.SelectionStart = TextBox12.TextLength - 3


        End If
    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged
        If IsNumeric(TextBox9.Text) Then

            Dim temp As Integer = TextBox9.Text

            TextBox9.Text = Format(temp, "N")

            TextBox9.SelectionStart = TextBox9.TextLength - 3


        End If
    End Sub

    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox11.TextChanged
        If IsNumeric(TextBox11.Text) Then

            Dim temp As Integer = TextBox11.Text

            TextBox11.Text = Format(temp, "N")

            TextBox11.SelectionStart = TextBox11.TextLength - 3

            TextBox9.Text = TextBox11.Text - TextBox12.Text
        End If
    End Sub

    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        'TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        ComboBox1.SelectedItem = ""
        ListBox1.Items.Clear()
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        bersih()

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If ComboBox1.SelectedItem = "" Or TextBox9.Text = "" Or TextBox11.Text = "" Or TextBox12.Text = "" Then
            MsgBox("Silahkan Isi Semua Form", MsgBoxStyle.Critical, "Error Message")
        Else
            If MessageBox.Show("Apakah Yakin Akan Menyimpan...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Try
                    Koneksi()
                    Dim data As String = ComboBox1.SelectedItem
                    Dim arr() As String = data.Split("-")
                    Dim CMD As OleDbCommand
                    Dim simpan As String = "insert into resep (nomor_resep, tanggal_resep, kode_dokter, kode_pasien, kode_poli, kode_pemakai, total_harga, dibayar, kembali) values ('" & arr.First & "','" & TextBox6.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & MainMenu.Label4.Text & "','" & TextBox12.Text & "','" & TextBox11.Text & "','" & TextBox9.Text & "')"
                    'Clipboard.SetText(CStr(simpan))
                    CMD = New OleDbCommand(simpan, Conn)
                    CMD.ExecuteNonQuery()


                    Dim x As OleDbCommand
                    Dim update As String = "update pendaftaran set ket='1' where nomor_daftar='" & arr.First & "'"
                    x = New OleDbCommand(update, Conn)
                    x.ExecuteNonQuery()


                    Dim i = DataGridView1.CurrentRow.Index

                    'DataGridView1.Item(4, i).Value = CDbl(DataGridView1.Item(2, i).Value) * CDbl(DataGridView1.Item(3, i).Value)

                    For Each row As DataGridViewRow In Me.DataGridView1.Rows
                        If (Not row.IsNewRow) Then
                            'subtotal = subtotal + row.Cells(4).Value.ToString
                            Dim CMD2 As OleDbCommand
                            Dim simpan2 As String = "insert into detail (nomor_resep, kode_obat, harga, dosis, subtotal) values ('" & arr.First & "','" & row.Cells(0).Value.ToString & "','" & row.Cells(2).Value.ToString & "','" & row.Cells(3).Value.ToString & "','" & row.Cells(4).Value.ToString & "')"

                            CMD2 = New OleDbCommand(simpan2, Conn)
                            CMD2.ExecuteNonQuery()

                            Dim CMD4 As OleDbCommand
                            Dim RD4 As OleDbDataReader
                            CMD4 = New OleDbCommand("Select jumlah From obat where kode_obat = '" & row.Cells(0).Value.ToString & "'", Conn)
                            RD4 = CMD4.ExecuteReader
                            Do While RD4.Read
                                Dim angka As Integer
                                angka = row.Cells(3).Value.ToString
                                Dim xx As OleDbCommand
                                Dim updatee As String = "update obat set jumlah='" & RD4.Item(0) - angka & "' where kode_obat='" & row.Cells(0).Value.ToString & "'"
                                xx = New OleDbCommand(updatee, Conn)
                                'Clipboard.SetText(CStr(updatee))
                                xx.ExecuteNonQuery()
                            Loop

                        End If
                    Next

                    Dim CMD3 As OleDbCommand
                    Dim simpan3 As String = "insert into pembayaran (nomor_byr, kode_pasien, tanggal_bayar, jumlah_bayar) values ('" & arr.First & "','" & TextBox2.Text & "','" & TextBox6.Text & "','" & TextBox12.Text & "')"
                    Clipboard.SetText(CStr(simpan3))
                    CMD3 = New OleDbCommand(simpan3, Conn)
                    CMD3.ExecuteNonQuery()

                    MsgBox("Input data berhasil")

                    bersih()
                    
                Catch ex As Exception
                    MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
                End Try

            End If
        End If

    End Sub

   
 
End Class