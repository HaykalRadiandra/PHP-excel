<?php 
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

$koneksi = mysqli_connect('localhost:3308', 'root', '', 'dummy');
if (!$koneksi) {
    die("Koneksi Gagal: " . mysqli_connect_error());
}

$result = mysqli_query($koneksi, "SELECT * FROM siswa");

if (isset($_POST['import'])) {
    $targetPath = $_FILES['fileexcel']['tmp_name'];
    $spreadsheet = IOFactory::load($targetPath);
    $data = $spreadsheet->getActiveSheet()->toArray();

    $headerRowIndex = null;
    $columns = [];

    foreach ($data as $i => $row) {
        $lowerRow = array_map(function($cell) {
            return is_string($cell) ? strtolower(trim($cell)) : '';
        }, $row);

        echo "<pre>Baris ke-$i:\n";
        print_r($lowerRow);
        echo "</pre>";

        if (
            in_array("nama", $lowerRow) &&
            in_array("nis", $lowerRow) &&
            in_array("nisn", $lowerRow) &&
            in_array("masa sekolah", $lowerRow) &&
            in_array("kelas", $lowerRow) &&
            in_array("jurusan", $lowerRow) &&
            in_array("indeks", $lowerRow) &&
            in_array("alamat", $lowerRow) &&
            (in_array("no telepon", $lowerRow) || in_array("notelp", $lowerRow)) &&
            in_array("tanggal lahir", $lowerRow)
        ) {
            echo "<p style='color:green'>Header ditemukan di baris: $i</p>";
            $headerRowIndex = $i;

            foreach ($lowerRow as $idx => $colName) {
                if ($colName === 'nama') $columns['nama'] = $idx;
                if ($colName === 'nis') $columns['nis'] = $idx;
                if ($colName === 'nisn') $columns['nisn'] = $idx;
                if ($colName === 'masa sekolah') $columns['masasekolah'] = $idx;
                if ($colName === 'kelas') $columns['kelas'] = $idx;
                if ($colName === 'jurusan') $columns['jurusan'] = $idx;
                if ($colName === 'indeks') $columns['indeks'] = $idx;
                if ($colName === 'alamat') $columns['alamat'] = $idx;
                if ($colName === 'no telepon' || $colName === 'notelp') $columns['notelp'] = $idx;
                if ($colName === 'tanggal lahir') $columns['tgllahir'] = $idx;
            }

            break;
        } else {
            echo "<p style='color:red'>Header belum ditemukan di baris: $i</p>";
        }
    }

    if ($headerRowIndex !== null) {
        echo "<pre>Total baris data: " . count($data) . "</pre>";

        for ($i = $headerRowIndex + 1; $i < count($data); $i++) {
            $row = $data[$i];

            $nama        = $row[$columns['nama']] ?? '';
            $nis         = $row[$columns['nis']] ?? '';
            $nisn        = $row[$columns['nisn']] ?? '';
            $masasekolah = $row[$columns['masasekolah']] ?? '';
            $kelas       = $row[$columns['kelas']] ?? '';
            $jurusan     = $row[$columns['jurusan']] ?? '';
            $indeks      = $row[$columns['indeks']] ?? '';
            $alamat      = $row[$columns['alamat']] ?? '';
            $notelp      = $row[$columns['notelp']] ?? '';
            $tgllahir    = $row[$columns['tgllahir']] ?? '';
            $userid = 'admin';

            echo "<pre>Baris ke-$i:\n";
            echo "Nama: $nama\n";
            echo "NIS: $nis\n";
            echo "NISN: $nisn\n";
            echo "Masa Sekolah: $masasekolah\n";
            echo "Kelas: $kelas\n";
            echo "Jurusan: $jurusan\n";
            echo "Indeks: $indeks\n";
            echo "Alamat: $alamat\n";
            echo "No Telp: $notelp\n";
            echo "Tgl Lahir: $tgllahir\n";
            echo "</pre>";

            if ($nama && $nis && $nisn && $masasekolah && $kelas && $jurusan && $indeks && $alamat && $notelp && $tgllahir) {
                $cek = mysqli_query($koneksi, "SELECT MAX(RIGHT(kodesiswa,4)) AS nourut FROM siswa");
                $rec = mysqli_fetch_assoc($cek);
                $nourut = ($rec['nourut'] ?? 0) + 1;
                if ($nourut > 9999) $nourut = 1;

                $kodesiswa = "SISW" . sprintf("%04s", $nourut);

                mysqli_query($koneksi, "INSERT INTO siswa (kodesiswa, nama, nis, nisn, masasekolah, kelas, kodejurusan, indeks, alamat, notelp, tgllahir, tglentry, userid, tglupdate, userupdate, onview) VALUES (
                    '$kodesiswa', '$nama', '$nis', '$nisn', '$masasekolah', '$kelas', '$jurusan', '$indeks', '$alamat', '$notelp', '$tgllahir', NOW(), '$userid', NOW(), '$userid', 1)");

                echo "<p style='color:green'>✔ Data berhasil di-insert: $kodesiswa</p>";
            } else {
                echo "<p style='color:red'>✘ Data tidak lengkap, tidak di-insert!</p>";
            }
        }
    } else {
        echo "<p style='color:red'>Header tidak ditemukan!</p>";
    }
}
?>


<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Import Excel</title>
</head>
<body>
    <br><br>
    <form method="post" enctype="multipart/form-data">
        <input type="file" name="fileexcel" accept=".xls, .xlsx" required>
        <button type="submit" name="import">Import Excel</button>
    </form>
    <hr>
    <br>
    <table border="1">
        <tr>
            <td>No</td>
            <td>Kodesiswa</td>
            <td>Nama</td>
            <td>NIS</td>
            <td>NISN</td>
            <td>Masa Sekolah</td>
            <td>Kelas</td>
            <td>Jurusan</td>
            <td>Indeks</td>
            <td>Alamat</td>
            <td>No Telepon</td>
            <td>Tanggal Lahir</td>
        </tr>
        <?php 
        $no = 1;
        $rows = [];
        while ($row = mysqli_fetch_assoc($result)) {
            $rows[] = $row;
        }
        foreach ($rows as $row):
        ?>
        <tr>
            <td><?= $no++ ?></td>
            <td><?= $row['kodesiswa']; ?></td>
            <td><?= $row['nama']; ?></td>
            <td><?= $row['nis']; ?></td>
            <td><?= $row['nisn']; ?></td>
            <td><?= $row['masasekolah']; ?></td>
            <td><?= $row['kelas']; ?></td>
            <td><?= $row['kodejurusan']; ?></td>
            <td><?= $row['indeks']; ?></td>
            <td><?= $row['alamat']; ?></td>
            <td><?= $row['notelp']; ?></td>
            <td><?= $row['tgllahir']; ?></td>
        </tr>
        <?php endforeach; ?>
    </table>

</body>
</html>
