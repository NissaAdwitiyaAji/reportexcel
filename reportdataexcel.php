<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
    $sheet->setCellValue('A1', 'No');
    $sheet->setCellValue('B1', 'Jenis Pendaftaran');
    $sheet->setCellValue('C1', 'Tgl Masuk Sekolah');
    $sheet->setCellValue('D1', 'NIS');
    $sheet->setCellValue('E1', 'No. Peserta Ujian');
    $sheet->setCellValue('F1', 'Pernah Paud');
    $sheet->setCellValue('G1', 'Pernah Tk');
    $sheet->setCellValue('H1', 'No.Skhun');
    $sheet->setCellValue('I1', 'No.Ijazah');
    $sheet->setCellValue('J1', 'Hobi');
    $sheet->setCellValue('K1', 'Cita-Cita');
    $sheet->setCellValue('L1', 'Nama Lengkap');
    $sheet->setCellValue('M1', 'Jk');
    $sheet->setCellValue('N1', 'NISN');
    $sheet->setCellValue('O1', 'NIK');
    $sheet->setCellValue('P1', 'Tempat Lahir');
    $sheet->setCellValue('Q1', 'Tanggal Lahir');
    $sheet->setCellValue('R1', 'Agama');
    $sheet->setCellValue('S1', 'Berkebutuhan Khusus');
    $sheet->setCellValue('T1', 'Alamat');
    $sheet->setCellValue('U1', 'RT');
    $sheet->setCellValue('V1', 'RW');
    $sheet->setCellValue('W1', 'Dusun');
    $sheet->setCellValue('X1', 'Kelurahan');
    $sheet->setCellValue('Y1', 'Kecamatan');
    $sheet->setCellValue('Z1', 'Kode Pos');
    $sheet->setCellValue('AA1', 'Tempat Tinggal');
    $sheet->setCellValue('AB1', 'Transportasi');
    $sheet->setCellValue('AC1', 'No.Hp');
    $sheet->setCellValue('AD1', 'No.Tlp');
    $sheet->setCellValue('AE1', 'Email');
    $sheet->setCellValue('AF1', 'Penerima Kps');
    $sheet->setCellValue('AG1', 'No.Kps');
    $sheet->setCellValue('AH1', 'Kewarganegaraan');
    $sheet->setCellValue('AI1', 'Negara');
    $sheet->setCellValue('AJ1', 'Nama Ayah Kandung');
    $sheet->setCellValue('AK1', 'Tahun Lahir');
    $sheet->setCellValue('AL1', 'Pendidikan');
    $sheet->setCellValue('AM1', 'Pekerjaan');
    $sheet->setCellValue('AN1', 'Penghasilan Bulanan');
    $sheet->setCellValue('AO1', 'Berkebutuhan Khusus');
    $sheet->setCellValue('AP1', 'Nama Ibu Kandung');
    $sheet->setCellValue('AQ1', 'Tahun Lahir');
    $sheet->setCellValue('AR1', 'Pendidikan');
    $sheet->setCellValue('AS1', 'Pekerjaan');
    $sheet->setCellValue('AT1', 'Penghasilan');
    $sheet->setCellValue('AU1', 'Berkebutuhan Khusus');

$koneksi = mysqli_connect("localhost", "root", "", "form_peserta");
$sql= mysqli_query ($koneksi, "SELECT * FROM data_pribadi, pribadi, ayah, ibu");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A1' . $i, $no++);
    $sheet->setCellValue('B1' . $i, $row['jenis_pendaftaran']);
    $sheet->setCellValue('C1' . $i, $row['tanggal_masuk_sekolah']);
    $sheet->setCellValue('D1' . $i, $row['nis']);
    $sheet->setCellValue('E1' . $i, $row['no_peserta_ujian']);
    $sheet->setCellValue('F1' . $i, $row['apakah_pernah_paud']);
    $sheet->setCellValue('G1' . $i, $row['apakah_pernah_tk']);
    $sheet->setCellValue('H1' . $i, $row['no_seri_skhun_sebelumnya']);
    $sheet->setCellValue('I1' . $i, $row['no_seri_ijazah_sebelumnya']);
    $sheet->setCellValue('J1' . $i, $row['hobi']);
    $sheet->setCellValue('K1' . $i, $row['citacita']);
    $sheet->setCellValue('L1' . $i, $row['nama_lengkap']);
    $sheet->setCellValue('M1' . $i, $row['jenis_kelamin']);
    $sheet->setCellValue('N1' . $i, $row['nisn']);
    $sheet->setCellValue('O1' . $i, $row['nik']);
    $sheet->setCellValue('P1' . $i, $row['tempat_lahir']);
    $sheet->setCellValue('Q1' . $i, $row['tempat_lahir']);
    $sheet->setCellValue('R1' . $i, $row['agama']);
    $sheet->setCellValue('S1' . $i, $row['berkebutuhan_khusus']);
    $sheet->setCellValue('T1' . $i, $row['alamat']);
    $sheet->setCellValue('U1' . $i, $row['rt']);
    $sheet->setCellValue('V1' . $i, $row['rw']);
    $sheet->setCellValue('W1' . $i, $row['nama_dusun']);
    $sheet->setCellValue('X1' . $i, $row['nama_kel']);
    $sheet->setCellValue('Y1' . $i, $row['kecamatan']);
    $sheet->setCellValue('Z1' . $i, $row['kode_pos']);
    $sheet->setCellValue('AA1' . $i, $row['tempat_tinggal']);
    $sheet->setCellValue('AB1' . $i, $row['transportasi']);
    $sheet->setCellValue('AC1' . $i, $row['no_hp']);
    $sheet->setCellValue('AD1' . $i, $row['no_telp']);
    $sheet->setCellValue('AE1' . $i, $row['email']);
    $sheet->setCellValue('AF1' . $i, $row['kpspkh']);
    $sheet->setCellValue('AG1' . $i, $row['nokpspkh']);
    $sheet->setCellValue('AH1' . $i, $row['kewarnegaraan']);
    $sheet->setCellValue('AI1' . $i, $row['nama_negara']);
    $sheet->setCellValue('AJ1' . $i, $row['nama_ayah']);
    $sheet->setCellValue('AK1' . $i, $row['tahun_lahir']);
    $sheet->setCellValue('AL1' . $i, $row['pendidikan']);
    $sheet->setCellValue('AM1' . $i, $row['pekerjaan']);
    $sheet->setCellValue('AN1' . $i, $row['penghasilan_bulanan']);
    $sheet->setCellValue('AO1' . $i, $row['kebutuhan_khusus']);
    $sheet->setCellValue('AP1' . $i, $row['nama_ibu']);
    $sheet->setCellValue('AQ1' . $i, $row['tahun_lahir_ibu']);
    $sheet->setCellValue('AR1' . $i, $row['pendidikan_ibu']);
    $sheet->setCellValue('AS1' . $i, $row['pekerjaan_ibu']);
    $sheet->setCellValue('AT1' . $i, $row['penghasilan_bulanan_ibu']);
    $sheet->setCellValue('AU1' . $i, $row['kebutuhan_khusus_ibu']);
   

    $i++;
}

$styleArray = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];
$sheet->getStyle('A1:AU1'.$i)->applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
$writer->save('Report Peserta.xlsx');
?>