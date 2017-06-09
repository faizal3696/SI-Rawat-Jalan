Imports System.Data.OleDb
Public Class Poli

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
    Sub kodepoli()
        Koneksi()
        Dim CMD As OleDbCommand
        Dim RD As OleDbDataReader
        Dim urutan As Int32
        'Dim hitung As Long
        Dim a As String


        CMD = New OleDbCommand("select kode_poli from poli Where kode_poli In(Select Max(kode_poli)From poli)Order By kode_poli Desc", Conn)
        'MsgBox("Select kode_obat From obat where kode_obat = '" & ComboBox1.SelectedItem & "' order by kode_obat desc limit 1")
        RD = CMD.ExecuteReader
        RD.Read()

        'MsgBox(a + 1)
        If Not RD.HasRows Then
            urutan = "1"
            TextBox1.Text = urutan.ToString("D2")

        Else
            a = RD.GetString(0)
            'MsgBox(a)
            urutan = a + 1
            'MsgBox(urutan)
            TextBox1.Text = urutan.ToString("D2")
        End If
        TextBox1.Enabled = False


    End Sub


    Sub loaddata()
        da = New OleDbDataAdapter("Select * from poli order by kode_poli asc", Conn)
        dt = New DataTable
        da.Fill(dt)
        DataGridView1.DataSource = dt
        'DataGridView1.DataSource = (ds.Tables("pasien"))
        DataGridView1.Columns(0).HeaderText = "Kode Poli"
        DataGridView1.Columns(1).HeaderText = "Nama Poli"
        DataGridView1.AutoResizeColumns()
        DataGridView1.Columns(1).Width = 293
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Gray

    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = True
        Button5.Enabled = True
        TextBox1.Text = ""
        TextBox1.Enabled = True
        TextBox2.Text = ""
        TextBox2.Enabled = True
        TextBox1.Focus()
        kodepoli()

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Button5.Enabled = False
        If TextBox1.Text = "" Then
            MsgBox("Silahkan Isi Semua Form", MsgBoxStyle.Critical, "Error Message")
        Else
            If MessageBox.Show("Apakah Yakin Akan Menyimpan...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Try
                    Koneksi()

                    Dim CMD As OleDbCommand
                    Koneksi()
                    Dim simpan As String = "insert into poli (kode_poli,nama_poli) values ('" & TextBox1.Text & "','" & TextBox2.Text & "')"
                    CMD = New OleDbCommand(simpan, Conn)
                    CMD.ExecuteNonQuery()
                    MsgBox("Input data berhasil")
                    TextBox1.Text = ""
                    TextBox2.Text = ""

                    TextBox1.Enabled = False
                    TextBox2.Enabled = False
                    loaddata()
                Catch ex As Exception
                    MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
                End Try

            End If

        End If
        Button4.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox1.Text = ""
        TextBox2.Text = ""

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = False
        Button5.Enabled = False

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If id.Text = "" Then
            MsgBox("Silahkan Pilih Data yang akan di hapus")
        Else
            If MessageBox.Show("Yakin akan dihapus..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Koneksi()
                Dim CMD As OleDbCommand
                Dim hapus As String = "delete From poli  where kode_poli='" & id.Text & "'"
                CMD = New OleDbCommand(hapus, Conn)
                CMD.ExecuteNonQuery()
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox1.Enabled = False
                TextBox2.Enabled = False
                MsgBox("Data Berhasil Dihapus", MsgBoxStyle.Information)
                loaddata()

            End If
        End If
        Button5.Enabled = False
    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox1.Enabled = False
        TextBox2.Enabled = True

        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        Me.id.Text = DataGridView1.Item(0, i).Value
        Me.TextBox1.Text = DataGridView1.Item(0, i).Value
        Me.TextBox2.Text = DataGridView1.Item(1, i).Value
        Button2.Enabled = True
        Button4.Enabled = False
        Button5.Enabled = True

    End Sub



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try

            Koneksi()
            Dim CMD As OleDbCommand
            Dim edit As String
            edit = "update poli set nama_poli = '" & TextBox1.Text & "' where kode_poli = " & id.Text & ""
            If TextBox1.Text = "" Then
                MsgBox("Data Gagal Diupdate", MsgBoxStyle.Critical, "Error Message")
            Else
                CMD = New OleDbCommand(edit, Conn)
                CMD.ExecuteNonQuery()
                MsgBox("Data Berhasil diUpdate")
                loaddata()

            End If
            TextBox1.Text = ""
            TextBox1.Enabled = False
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
            CMD = New OleDbCommand("Select * From poli where nama_poli like '%" & TextBox5.Text & "%'", Conn)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("Data Pasien Tidak Ditemukan")
                TextBox5.Focus()
            Else
                da = New OleDbDataAdapter("Select * From poli where nama_poli like '%" & TextBox5.Text & "%'", Conn)
                dt = New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
                'DataGridView1.DataSource = (ds.Tables("pasien"))
                DataGridView1.Columns(0).HeaderText = "Kode Poli"
                DataGridView1.Columns(1).HeaderText = "Nama Poli"

                DataGridView1.AutoResizeColumns()
                DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Gray
            End If
        End If
        Button5.Enabled = True
    End Sub

    Private Sub Poli_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Koneksi()
        loaddata()
        Button4.Enabled = False
        Button5.Enabled = False
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub
End Class