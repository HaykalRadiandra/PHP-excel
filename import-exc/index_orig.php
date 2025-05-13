<?php 
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

$conn = new mysqli('localhost:3308', 'root', '', 'dummy');

if ($conn->connect_error) {
    die("Koneksi Gagal: " . $conn->connect_error);
}

$result = $conn->query("SELECT * FROM users");
$rows = $result->fetch_all(MYSQLI_ASSOC);

if (isset($_POST['import'])) {
    $filename = $_FILES['fileexcel'];
    $fileExtension = explode('.', $filename['name']);
    $fileExtension = strtolower(end($fileExtension));

    $newFileName = date("Y.m.d") . "-" . date("h.i.sa") . "." . $fileExtension;
    $targetPath = "uploads/" . $newFileName;
    move_uploaded_file($_FILES['fileexcel']['tmp_name'], $targetPath);

    error_reporting(0);
    ini_set('display_errors', 0);

    $spreadsheet = IOFactory::load($targetPath);
    $data = $spreadsheet->getActiveSheet()->toArray();

    foreach ($data as $index => $row) {
        // ** Lewati Baris Header A **
        if ($index == 0) continue;

        // ** Ambil Data dari Kolom B (1), C (2), D (3) **
        $name = $row[1];
        $age = $row[2];
        $country = $row[3];

        if (!empty($name)) {
            $stmt = $conn->prepare("INSERT INTO users (nama, age, country) VALUES (?, ?, ?)");
            $stmt->bind_param("sis", $name, $age, $country);
            $stmt->execute();
            echo
                "
                <script>
                alert('Succesfully Imported');
                document.location.href = '';
                </script>
                ";
        } else {
            echo
                "
                <script>
                alert('Failed Imported');
                document.location.href = '';
                </script>
                ";
        }
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
            <td>Name</td>
            <td>Age</td>
            <td>Country</td>
        </tr>
        <?php 
        $no = 1;
        foreach ($rows as $row):
        ?>
        <tr>
            <td><?= $no++ ?></td>
            <td><?= $row['nama']; ?></td>
            <td><?= $row['age']; ?></td>
            <td><?= $row['country']; ?></td>
        </tr>
        <?php endforeach ; ?>
    </table>
</body>
</html>