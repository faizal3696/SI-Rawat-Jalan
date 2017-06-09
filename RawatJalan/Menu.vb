Public Class MainMenu
    Dim iduser As New System.Windows.Forms.TextBox
    Private Sub DataPemakaiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Poli.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dokter.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Obat.Show()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Pasien.Show()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        pemakai.Show()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Pendaftaran.Show()
    End Sub

    Private Sub DataPoliToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataPoliToolStripMenuItem.Click
        Dim a As New PoliReport
        Dim b As New PoliDesign
        b.CrystalReportViewer1.ReportSource = a
        b.ShowDialog()
    End Sub

    Private Sub QuitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuitToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub DataDokterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataDokterToolStripMenuItem.Click
        Dim a As New DokterReport
        Dim b As New DokterDesign
        b.CrystalReportViewer1.ReportSource = a
        b.ShowDialog()
    End Sub

    Private Sub DataObatToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataObatToolStripMenuItem.Click
        Dim a As New ObatReport
        Dim b As New ObatDesign
        b.CrystalReportViewer1.ReportSource = a
        b.ShowDialog()
    End Sub

    Private Sub DataPasienToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataPasienToolStripMenuItem.Click
        Dim a As New PasienReport
        Dim b As New PasienDesign
        b.CrystalReportViewer1.ReportSource = a
        b.ShowDialog()
    End Sub

    Private Sub DataPemakaiToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataPemakaiToolStripMenuItem.Click
        Dim a As New PemakaiReport
        Dim b As New PemakaiDesign
        b.CrystalReportViewer1.ReportSource = a
        b.ShowDialog()
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dokter.Show()

    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Pasien.Show()
    End Sub

    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dokter.Show()

    End Sub

    Private Sub Button6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Pendaftaran.Show()

    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Obat.Show()

    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        pemakai.Show()

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Poli.Show()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Dim TD As Date
        TD = DateTime.Now
        Label2.Text = TD
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Application.Exit()
    End Sub

    Private Sub MainMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label4.Hide()

    End Sub

    Private Sub DataPendaftaranToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataPendaftaranToolStripMenuItem.Click
        PendaftaranDesign.Show()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Resep.Show()

    End Sub
End Class