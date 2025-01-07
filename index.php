<?php
if (isset($_POST['submit'])) {
    $file = $_FILES['excelFile'];
    
    if ($file['type'] == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
        require 'vendor/autoload.php';
        
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load($file['tmp_name']);
        
        $studentsSheet = $spreadsheet->getSheet(0)->toArray(null, true, true, true);
        $centersSheet = $spreadsheet->getSheet(1)->toArray(null, true, true, true);
        
        $students = array_slice($studentsSheet, 1);
        $centers = array_slice($centersSheet, 1);
        
        $schedule = [];
        $centerCapacity = [];
        
        foreach ($centers as $center) {
            if($center['A']=="Grand Total" || empty($center['B'])){
                continue;
            }
            $centerCapacity[$center['A']] = $center['C'];
        }
        //print_r($students);die;
        usort($students, function($a, $b) {
            if ($a['E'] != $b['E']) return $a['E'] == 'YES' ? -1 : 1; 
            if ($a['C'] != $b['C']) return $a['C'] == 'FEMALE' ? -1 : 1;
            return 0;
        });
        
        foreach ($students as $student) {
            $assigned = false;
            foreach ($centers as $center) {
                if (strpos($student['B'], $center['B']) !== false && $centerCapacity[$center['A']] > 0) {
                    $schedule[] = [
                        'Name' => $student['A'],
                        'Trade' => $student['F'],
                        'Day' => $student['G'],
                        'Center' => $center['A']
                    ];
                    $centerCapacity[$center['A']]--;
                    $assigned = true;
                    break;
                }
            }
            if (!$assigned) {

                foreach ($centers as $center) {
                    if ($centerCapacity[$center['A']] > 0) {
                        $schedule[] = [
                            'Name' => $student['A'],
                            'Trade' => $student['F'],
                            'Day' => $student['G'],
                            'Center' => $center['A']
                        ];
                        $centerCapacity[$center['A']]--;
                        $assigned = true;
                        break;
                    }
                }

                if (!$assigned) {
                    $schedule[] = [
                        'Name' => $student['A'],
                        'Trade' => $student['F'],
                        'Day' => $student['G'],
                        'Center' => 'Not Assigned'
                    ];
                }
            }
        }
        
        $output = fopen('php://output', 'w');
        header('Content-Type: text/csv');
        header('Content-Disposition: attachment;filename="schedule.csv"');
        fputcsv($output, ['Name', 'Trade', 'Day', 'Center']);
        foreach ($schedule as $row) {
            fputcsv($output, $row);
        }
        fclose($output);
        exit;
    } else {
        echo 'Invalid file format. Please upload an Excel file.';
    }
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Training Scheduler</title>
</head>
<body>
    <h2>Upload Training Data</h2>
    <form action="" method="post" enctype="multipart/form-data">
        <input type="file" name="excelFile" required>
        <button type="submit" name="submit">Upload and Process</button>
    </form>
</body>
</html>
