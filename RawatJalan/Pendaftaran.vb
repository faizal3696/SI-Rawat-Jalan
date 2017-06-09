Imports System.Data.OleDb

Public Class Pendaftaran
    Dim Conn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim LokasiDB As String
    Dim id As New System.Windows.Forms.TextBox
    Dim save As New System.Windows.Forms.TextBox
    Dim jk As New System.Windows.Forms.TextBox
    Dim getdata As New TextBox
    Dim getdata2 As New TextBox
    Dim getdata3 As New TextBox
    Public dt As DataTable
    Public ds As DataSet

    Sub Koneksi()
        LokasiDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=sirjdb.accdb"
        Conn = New OleDbConnection(LokasiDB)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub
    Sub loaddata()
        Koneksi()

        da = New OleDbDataAdapter("Select * from pendaftaran where (tanggal_daftar = #" & Format(Now, "M/dd/yyyy") & "#) ", Conn)
        dt = New DataTable
        da.Fill(dt)
        DataGridView1.DataSource = dt
        'DataGridView1.DataSource = (ds.Tables("pasien"))
        DataGridView1.Columns(0).HeaderText = "Nomor Daftar"
        DataGridView1.Columns(1).HeaderText = "Tanggal Daftar"
        DataGridView1.Columns(2).HeaderText = "Kode Dokter"
        DataGridView1.Columns(3).HeaderText = "Kode Pasien"
        DataGridView1.Columns(4).HeaderText = "Kode Poli"
        DataGridView1.Columns(5).HeaderText = "Kode Pemakai"
        DataGridView1.Columns(6).HeaderText = "Biaya"
        DataGridView1.Columns(6).DefaultCellStyle.Format = "##,0"
        DataGridView1.Columns(7).HeaderText = "Keterangan"
        DataGridView1.Columns(8).HeaderText = "Nomor Antrian"
        DataGridView1.Columns(8).DisplayIndex = 0
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Gray

    End Sub
    Sub nomorpendaftaran()
        Koneksi()
        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader
        Dim urutan As String
        'Dim hitung As Long
        Dim a As Long

        'CMD = New OleDbCommand("Select top 1 nomor_daftar From pendaftaran order by nomor_daftar desc", Conn)
        CMD = New OleDbCommand("select nomor_daftar from pendaftaran Where nomor_daftar In(Select Max(nomor_daftar)From pendaftaran)Order By nomor_daftar Desc", Conn)
        'MsgBox("Select kode_obat From obat where kode_obat = '" & ComboBox1.SelectedItem & "' order by kode_obat desc limit 1")
        RD = CMD.ExecuteReader
        RD.Read()

        'MsgBox(a + 1)
        If Not RD.HasRows Then
            urutan = Format(Now, "yyMMdd") & "0001"
            TextBox1.Text = urutan

        Else
            a = RD.GetString(0)
            urutan = a + 1
           
            TextBox1.Text = urutan
        End If
        TextBox1.Enabled = False

        Return
    End Sub
    Sub nomorantrian()
        Koneksi()
        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader
        Dim urutan As String
        'Dim hitung As Long
        Dim a As Long
        Dim query As String

        'CMD = New OleDbCommand("Select top 1 nomor_daftar From pendaftaran order by nomor_daftar desc", Conn)
        'query = "select nomor_antri from pendaftaran Where tanggal_daftar = #" & Format(Now, "MM/dd/yyyy") & "# order by tanggal_daftar desc"
        query = "SELECT     * FROM         pendaftaran WHERE     (tanggal_daftar = #" & Format(Now, "M/dd/yyyy") & "#) ORDER BY tanggal_daftar DESC"
        CMD = New OleDbCommand(query, Conn)
        'Clipboard.SetText(CStr(query))
        RD = CMD.ExecuteReader
        RD.Read()

        'MsgBox(a + 1)
        If Not RD.HasRows Then
            urutan = "1"
            TextBox3.Text = urutan

        Else
            'MsgBox(RD.GetString(0))
            'MsgBox(substrdata.Length - 1)
            'hitung = a + 1
            'MsgBox(Microsoft.VisualBasic.Right(RD.GetString(0), 5))
            'MsgBox("000" & hitung)
            MsgBox(RD.GetString(0))
            'MsgBox(RD.GetString(1))
            MsgBox(RD.GetString(2))
            MsgBox(RD.GetString(3))
            MsgBox(RD.GetString(4))
            MsgBox(RD.GetString(5))
            'MsgBox(RD.GetString(6))
            MsgBox(RD.GetString(7))
            MsgBox(RD.GetValue(8))
            a = RD.GetValue(8)
            urutan = a + 1
            'Dim substrdata As String
            'substrdata = ComboBox1.SelectedItem
            'Console.WriteLine(substrdata)
            'Data = substrdata.Substring(0, substrdata.Length)
            'Console.WriteLine(Data)
            'MsgBox(RD.HasRows)
            TextBox3.Text = urutan
        End If
        'TextBox1.Enabled = False

        Return
    End Sub

    Sub loaddatapasien()
        Koneksi()

        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader
        Dim format1 As Int32
        Dim z As Int32
        Dim format2 As Int32
        Dim z2 As Int32
        Dim format3 As Int32
        Dim z3 As Int32
        Dim formatnya As String

        format1 = Format(Now, "dd")
        format2 = Format(Now, "MM")
        format3 = Format(Now, "yy")
        formatnya = z & "" & z2 & "" & z3
        CMD = New OleDbCommand("select * FROM pasien where kode_pasien like '%" & Format(Now, "ddMMyy") & "%'order by kode_pasien desc", Conn)
        'Clipboard.SetText("select * FROM pasien where kode_pasien like '%" & Format(Now, "ddMMyy") & "%'order by kode_pasien desc")
        RD = CMD.ExecuteReader
        ComboBox2.Items.Clear()
        ComboBox2.Items.Add("")
        Do While RD.Read
            ComboBox2.Items.Add(RD.Item(1))
        Loop
        ComboBox2.Refresh()
        'MsgBox("hello")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        nomorpendaftaran()
        nomorantrian()
        TextBox2.Text = Format(Now, "dd-MM-yyy")
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ListBox1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
    End Sub

    Private Sub Pendaftaran_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Koneksi()

        loaddata()
        loaddatapasien()

        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader
        
        CMD = New OleDbCommand("select * FROM poli order by kode_poli asc", Conn)
        RD = CMD.ExecuteReader
        ComboBox1.Items.Add("")
        Do While RD.Read
            ComboBox1.Items.Add(RD.Item(1))
        Loop
        Button2.Enabled = False
        'Button4.Enabled = False
        Button3.Enabled = False


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Koneksi()
        

        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader

        Dim convert As Int32
        Dim result As String
        convert = ComboBox1.SelectedIndex
        If convert > 9 Then
            result = ComboBox1.SelectedIndex
        Else
            result = convert.ToString("D2")
        End If
        'MsgBox(result)

        CMD = New OleDbCommand("select * FROM dokter where kode_poli ='" & result & "'", Conn)
        RD = CMD.ExecuteReader
        ListBox1.Items.Clear()
        Do While RD.Read
            ListBox1.Items.Add(RD.Item(1))
        Loop
        ListBox1.Refresh()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ComboBox1.SelectedItem = "" Then
            MsgBox("Isikan kode poli !", MsgBoxStyle.Critical, "Error Message")
            ComboBox1.Focus()
        ElseIf TextBox9.Text = "" Then
            MsgBox("Silahkan pilih dokter!", MsgBoxStyle.Critical, "Error Message")
            ListBox1.Focus()
        ElseIf ComboBox2.SelectedItem = "" Then
            If MessageBox.Show("Kode pasien kosong tambah baru pasien...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                TextBox4.Enabled = True
                TextBox5.Enabled = True
                TextBox7.Enabled = True
                TextBox8.Enabled = True
                RadioButton1.Enabled = True
                RadioButton2.Enabled = True
                TextBox5.Focus()
            Else
                MsgBox("silahkan pilih data pasien")
                TextBox4.Enabled = False
                TextBox5.Enabled = False
                TextBox7.Enabled = False
                TextBox8.Enabled = False
                RadioButton1.Enabled = False
                RadioButton2.Enabled = False
            End If

            'MsgBox("Isikan kode pasien !", MsgBoxStyle.Critical, "Error Message")
            'ComboBox2.Focus()

        Else
            If MessageBox.Show("Apakah Yakin Akan Menyimpan...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Try
                    Koneksi()

                    Dim CMD1 As OleDbCommand
                    Dim RD1 As OleDbDataReader
                    CMD1 = New OleDbCommand("select * FROM dokter where nama_dokter ='" & ListBox1.SelectedItem & "'", Conn)
                    RD1 = CMD1.ExecuteReader
                    Do While RD1.Read
                        getdata.Text = RD1.Item(0)
                    Loop

                    Dim CMD2 As OleDbCommand
                    Dim RD2 As OleDbDataReader
                    CMD2 = New OleDbCommand("select * FROM pasien where nama_pasien ='" & ComboBox2.SelectedItem & "'", Conn)
                    RD2 = CMD2.ExecuteReader
                    Do While RD2.Read
                        getdata2.Text = RD2.Item(0)
                    Loop


                    Dim CMD3 As OleDbCommand
                    Dim RD3 As OleDbDataReader
                    CMD3 = New OleDbCommand("select * FROM poli where nama_poli ='" & ComboBox1.SelectedItem & "'", Conn)
                    RD3 = CMD3.ExecuteReader
                    Do While RD3.Read
                        getdata3.Text = RD3.Item(0)
                    Loop


                    Dim CMD As OleDbCommand
                    Dim simpan As String
                    Dim converts As Long

                    converts = TextBox9.Text

                    simpan = "insert into pendaftaran(nomor_daftar,tanggal_daftar,kode_dokter,kode_pasien,kode_poli,kode_pemakai,biaya,ket,nomor_antri) values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & getdata.Text & "','" & getdata2.Text & "','" & getdata3.Text & "','" & MainMenu.Label4.Text & "'," & converts & ",'0','" & TextBox3.Text & "')"
                    Clipboard.SetText(CStr(simpan))

                    'MsgBox(simpan)
                    Clipboard.SetText(CStr(simpan))

                    CMD = New OleDbCommand(simpan, Conn)
                    'Clipboard.SetText(CStr(CMD))
                    CMD.ExecuteNonQuery()
                    MsgBox("Input data berhasil")
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    TextBox5.Text = ""

                    TextBox7.Text = ""
                    TextBox8.Text = ""
                    TextBox9.Text = ""
                    ComboBox1.SelectedItem = ""
                    ComboBox2.SelectedItem = ""
                    ListBox1.Items.Clear()
                    TextBox4.Text = ""

                    TextBox1.Enabled = False
                    TextBox2.Enabled = False
                    TextBox3.Enabled = False
                    ComboBox1.Enabled = False
                    ComboBox2.Enabled = False

                    Button2.Enabled = False
                    Button3.Enabled = False
                    Button4.Enabled = False

                    loaddata()
                Catch ex As Exception
                    MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
                End Try

            End If

        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            Koneksi()
            Dim CMD As OleDbCommand
            Dim RD As OleDbDataReader
            Dim formatnya As String
            Dim format1 As Int32
            Dim z As Int32
            Dim format2 As Int32
            Dim z2 As Int32
            Dim format3 As Int32
            Dim z3 As Int32
            Dim a As Int32
            Dim CMDPasienn As OleDbCommand
            Dim simpanPasien As String

            CMD = New OleDbCommand("select top 1 kode_pasien from pasien where kode_pasien like '%" & Format(Now, "ddMMyy") & "%' Order By kode_pasien Desc", Conn)
            RD = CMD.ExecuteReader
            RD.Read()

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
                    save.Text = formatnya + "01"
                Else
                    z = format1.ToString("D2")
                    z2 = format2.ToString("D2")
                    z3 = format3.ToString("D2")
                    formatnya = z.ToString("D2") & "" & z2.ToString("D2") & "" & z3.ToString("D2")
                    save.Text = formatnya & "01"
                End If

            Else
                a = RD.GetString(0)

                format1 = Format(Now, "dd")
                If format1 > 9 Then
                    save.Text = a + 1
                Else
                    save.Text = "0" & a + 1
                End If

            End If


            If (RadioButton1.Checked) Then
                jk.Text = "Pria"
            ElseIf (RadioButton2.Checked) Then
                jk.Text = "Wanita"
            Else
                jk.Text = ""
            End If
            If MessageBox.Show("Apakah Yakin Akan Menyimpan...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Try
                    simpanPasien = "insert into pasien (kode_pasien,nama_pasien,alamat,jenis_kelamin,umur,telepon) values ('" & save.Text & "','" & TextBox5.Text & "','" & TextBox4.Text & "','" & jk.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "')"
                    CMDPasienn = New OleDbCommand(simpanPasien, Conn)
                    CMDPasienn.ExecuteNonQuery()
                    loaddatapasien()
                Catch ex As Exception
                    MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
                End Try

            End If

            
            TextBox4.Text = ""
            TextBox5.Text = ""
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            TextBox7.Text = ""
            TextBox8.Text = ""

            TextBox4.Enabled = False
            TextBox5.Enabled = False
            TextBox7.Enabled = False
            TextBox8.Enabled = False
            RadioButton1.Enabled = False
            RadioButton2.Enabled = False



        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
        End Try

    End Sub

    Private Sub TextBox9_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox9.KeyPress
        
    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged
        If IsNumeric(TextBox9.Text) Then

            Dim temp As Integer = TextBox9.Text

            TextBox9.Text = Format(temp, "N")

            TextBox9.SelectionStart = TextBox9.TextLength - 3


        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""

        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        ComboBox1.SelectedItem = ""
        ComboBox2.SelectedItem = ""
        ListBox1.Items.Clear()
        TextBox4.Text = ""

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False

        Button2.Enabled = False
        Button3.Enabled = False

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Koneksi()
        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader
        CMD = New OleDbCommand("select * FROM dokter where nama_dokter ='" & ListBox1.SelectedItem & "'", Conn)
        RD = CMD.ExecuteReader
        Do While RD.Read

            TextBox9.Text = RD.Item(6)
        Loop

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Koneksi()
        If ComboBox2.SelectedItem = "" Then
            Button4.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox7.Enabled = True
            TextBox8.Enabled = True
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
        Else
            Button4.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            TextBox7.Enabled = False
            TextBox8.Enabled = False
            RadioButton1.Enabled = False
            RadioButton2.Enabled = False

            Dim CMD As OleDbCommand
            Dim RD As OleDbDataReader

            CMD = New OleDbCommand("select * FROM pasien where nama_pasien = '" & ComboBox2.SelectedItem & "'", Conn)
            RD = CMD.ExecuteReader
            Do While RD.Read
                ComboBox1.Items.Add(RD.Item(1))
                TextBox5.Text = RD.Item(1)
                TextBox4.Text = RD.Item(2)
                TextBox7.Text = RD.Item(4)
                TextBox8.Text = RD.Item(5)
                If RD.Item(3) = "Pria" Then
                    RadioButton1.Checked = True
                Else
                    RadioButton2.Checked = True
                End If
            Loop

        End If
        
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub
End Class