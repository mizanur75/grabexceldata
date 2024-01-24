<?php
session_start(); 
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Csv;

  $servername = "localhost";
  $username = "root";
  $password = "";

  try {
    $conn = new PDO("mysql:host=$servername; dbname=grab", $username, $password);
    // set the PDO error mode to exception
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    // echo "Connected successfully <br>";
  } catch(PDOException $e) {
    echo "Connection failed: " . $e->getMessage();
  }


  
// Check if a file is uploaded
if (isset($_FILES['excel_file'])) {
    $file = $_FILES['excel_file'];

    // Check for errors in file upload
    if ($file['error'] === UPLOAD_ERR_OK) {
        // Specify the path to store the uploaded file
        $uploadPath = __DIR__ . '/uploads/';
        $uploadedFile = $uploadPath . basename($file['name']);
        if(!file_exists($uploadPath)){
            mkdir('uploads', 777, true);
        }

        // Move the uploaded file to the specified directory
        try {
            if (move_uploaded_file($file['tmp_name'], $uploadedFile)) {

                // Load the Excel file
                $spreadsheet = IOFactory::load($uploadedFile);

                //Get all sheets name
                // $sheets = $spreadsheet->getSheetNames();


                // Get the desired sheet by name
                $entity_profile = $spreadsheet->getSheetByName('Entity Profile');
                $business_summary = $spreadsheet->getSheetByName('Business Summary');
                $annual_summary = $spreadsheet->getSheetByName('Annual Summary - Comparison');

                $wholesale_business = $entity_profile->getCell('B35')->getValue();
                $retail_business = $entity_profile->getCell('B36')->getValue();
                $others = $entity_profile->getCell('B37')->getValue();

                $gross_turnover24 = $business_summary->getCell('C6')->getValue();
                $gross_turnover23 = $business_summary->getCell('D6')->getValue();

                $net_revenue24 = $annual_summary->getCell('B4')->getValue();
                $net_revenue23 = $annual_summary->getCell('C4')->getValue();

                $existingExcelFile = __DIR__ . '/uploads/extracted_data.xlsx';

                if (file_exists($existingExcelFile)) {
                    // Load the existing spreadsheet
                    $existingSpreadsheet = IOFactory::load($existingExcelFile);

                    // Access the active sheet in the existing spreadsheet
                    $existingSheet = $existingSpreadsheet->getActiveSheet();

                    // Assuming $newData is an array representing the new row of data
                    $newData = [$wholesale_business, $retail_business, $others, $gross_turnover24, $gross_turnover23, $net_revenue24, $net_revenue23];

                    // Find the last row with data in the existing sheet
                    $lastRow = $existingSheet->getHighestRow() + 1;

                    // Loop through the new data and set the values in the new row
                    foreach ($newData as $columnIndex => $value) {
                        $existingSheet->setCellValueByColumnAndRow($columnIndex + 1, $lastRow, $value);
                    }

                    // Save the modified spreadsheet back to the file
                    $writer = new Xlsx($existingSpreadsheet);
                    $writer->save($existingExcelFile);

                }else{
                    // Create a new spreadsheet
                    $newSpreadsheet = new Spreadsheet();
                    $newSheet = $newSpreadsheet->getActiveSheet();

                    //Set the column name first
                    $newSheet->setCellValue('A1', $entity_profile->getCell('A35')->getValue());
                    $newSheet->setCellValue('B1', $entity_profile->getCell('A36')->getValue());
                    $newSheet->setCellValue('C1', $entity_profile->getCell('A37')->getValue());
                    $newSheet->setCellValue('D1', $business_summary->getCell('B6')->getValue().'24');
                    $newSheet->setCellValue('E1', $business_summary->getCell('B6')->getValue().'23');
                    $newSheet->setCellValue('F1', $annual_summary->getCell('A4')->getValue().'24');
                    $newSheet->setCellValue('G1', $annual_summary->getCell('A4')->getValue().'23');


                    // Write the extracted data to the new spreadsheet
                    $newSheet->setCellValue('A2', $wholesale_business);
                    $newSheet->setCellValue('B2', $retail_business);
                    $newSheet->setCellValue('C2', $others);
                    $newSheet->setCellValue('D2', $gross_turnover24);
                    $newSheet->setCellValue('E2', $gross_turnover23);
                    $newSheet->setCellValue('F2', $net_revenue24);
                    $newSheet->setCellValue('G2', $net_revenue23);

                    // Save the new spreadsheet as an Excel file (Xlsx format)
                    $newExcelFile = __DIR__ . '/uploads/extracted_data.xlsx';
                    $writer = new Xlsx($newSpreadsheet);
                    $writer->save($newExcelFile);
                }
                

                // Alternatively, you can save the data as a CSV file
                $existingCsvFile = __DIR__ . '/uploads/extracted_data.csv';

                if (file_exists($existingCsvFile)) {
                    // Load the existing spreadsheet
                    $existingCsv = IOFactory::load($existingCsvFile);

                    // Access the active sheet in the existing spreadsheet
                    $existingCsvF = $existingCsv->getActiveSheet();

                    // Assuming $newData is an array representing the new row of data
                    $newData = [$wholesale_business, $retail_business, $others, $gross_turnover24, $gross_turnover23, $net_revenue24, $net_revenue23];

                    // Find the last row with data in the existing sheet
                    $lastRow = $existingCsvF->getHighestRow() + 1;

                    // Loop through the new data and set the values in the new row
                    foreach ($newData as $columnIndex => $value) {
                        $existingCsvF->setCellValueByColumnAndRow($columnIndex + 1, $lastRow, $value);
                    }

                    // Save the modified spreadsheet back to the file
                    $writer = new Xlsx($existingCsv);
                    $writer->save($existingCsvFile);
                }else{
                    $newCsvFile = __DIR__ . '/uploads/extracted_data.csv';
                    $writer = new Csv($newSpreadsheet);
                    $writer->setDelimiter(',');
                    $writer->save($newCsvFile);
                }
                

                // Output a link to the new file
                $_SESSION['button'] = '<a class="btn btn-success" href="uploads/extracted_data.xlsx">Download Extracted Data</a>';

                $insert_query = "INSERT INTO `excel_datas` (`wholesale_business`, `retail_business`, `others`, `gross_turnover24`, `gross_turnover23`, `net_revenue24`, `net_revenue23`) VALUES(?,?,?,?,?,?,?)";
                $res = $conn->prepare($insert_query);
                $result = $res->execute([$wholesale_business, $retail_business, $others, $gross_turnover24, $gross_turnover23, $net_revenue24, $net_revenue23]);


                if($result){
                    // Close and delete the uploaded file
                    unlink($uploadedFile);
                    $_SESSION['success'] = "Data inserted successfully!";
                }else{
                    $_SESSION['error'] = "Opps! Something went wrong.";             
                }

                header('Location: index.php');
                
            } else {
                $_SESSION['error'] = "Error moving the uploaded file.";
                header('Location: index.php');
            }
        } catch (Exception $e) {
            return $e;
        }
    } else {

        $_SESSION['error'] = "File upload error: " . $file['error'];
        header('Location: index.php');
    }
}