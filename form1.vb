Imports System.Data.Odbc
Public Class Form1
    Sub show_db_mahasiswa()
        Call connection()
        cmd = New OdbcCommand("SELECT * FROM `db_mahasiswa`", conn)
        rd = cmd.ExecuteReader
        DataGridView1.Rows.Clear()
        Do While rd.Read = True
            DataGridView1.Rows.Add(rd(0), rd(1), rd(2), rd(3))
        Loop
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 200
        DataGridView1.Columns(2).Width = 150
        DataGridView1.Columns(3).Width = 120
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        ComboBox1.Text = DataGridView1.Item(2, i).Value
        ComboBox2.Text = DataGridView1.Item(3, i).Value
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KondisiAwal()
    End Sub

    Sub KondisiAwal()
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        TextBox1.MaxLength = 11
        Button1.Text = "Input Data Baru"
        Button2.Text = "Ubah Data"
        Button3.Text = "Hapus Data"
        Button4.Text = "Keluar"
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Call connection()
        Call show_db_mahasiswa()
    End Sub

    Sub RowAktif()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        TextBox1.Focus()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then e.Handled = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Button1.Text = "Input Data Baru" Then
            Button1.Text = "Simpan"
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Text = "Batal"
            Call RowAktif()
            TextBox1.Text = ""
            TextBox2.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Or ComboBox2.Text = "" Then
                MsgBox("Pastikan Semua Data Terisi !", MsgBoxStyle.Exclamation, "")
            Else
                Call connection()
                Dim InputData As String = "INSERT INTO `db_mahasiswa` VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "')"
                cmd = New OdbcCommand(InputData, conn)
                cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil Di Simpan !", MsgBoxStyle.Information, "")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Button2.Text = "Ubah Data" Then
            Button2.Text = "Ubah"
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Text = "Batal"
            TextBox2.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            TextBox2.Focus()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Or ComboBox2.Text = "" Then
                MsgBox("Pilih Data Yang Akan Di Ubah Pada Tabel !", MsgBoxStyle.Exclamation, "")
            Else
                Call connection()
                Dim UbahData As String = "UPDATE `db_mahasiswa` SET `NPM`='" & TextBox1.Text & "',`Nama`='" & TextBox2.Text & "',`Jurusan`='" & ComboBox1.Text & "',`Jenis Kelamin`='" & ComboBox2.Text & "' WHERE `NPM`='" & TextBox1.Text & "'"
                cmd = New OdbcCommand(UbahData, conn)
                cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil Di Ubah !", MsgBoxStyle.Information, "")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Button3.Text = "Hapus Data" Then
            Button3.Text = "Hapus"
            TextBox1.Enabled = False
            Button1.Enabled = False
            Button2.Enabled = False
            Button4.Text = "Batal"
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Or ComboBox2.Text = "" Then
                MsgBox("Pilih Data Yang Akan Di Hapus Pada Tabel !", MsgBoxStyle.Exclamation, "")
            Else
                Call connection()
                Dim HapusData As String = "DELETE FROM `db_mahasiswa` WHERE `NPM`='" & TextBox1.Text & "'"
                cmd = New OdbcCommand(HapusData, conn)
                cmd.ExecuteNonQuery()
                MsgBox("Data berhasil di hapus !", MsgBoxStyle.Information, "")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Button4.Text = "Keluar" Then
            If MessageBox.Show("Lanjutkan Menutup Aplikasi ?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Close()
            End If
        Else
            Call KondisiAwal()
        End If
    End Sub
End Class
