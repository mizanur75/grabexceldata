<?php
session_start(); 
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

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
        // var_dump($uploadedFile); exit;

        // Move the uploaded file to the specified directory
        if (move_uploaded_file($file['tmp_name'], $uploadedFile)) {
            // Specify the sheet name you want to extract data from
            // $sheetNames = ['Entity Profile', 'Business Summary', 'Annual Summary - Comparison']; // Replace with your sheet name

            // Load the Excel file
            $spreadsheet = IOFactory::load($uploadedFile);
            // $sheets = $spreadsheet->getSheetNames();

            	

            // Get the desired sheet by name
            $entity_profile = $spreadsheet->getSheetByName('Entity Profile');
            $business_summary = $spreadsheet->getSheetByName('Business Summary');
            $annual_summary = $spreadsheet->getSheetByName('Annual Summary - Comparison');

            $wholesale_business = $entity_profile->getCell('B35')->getValue() ?? 0;
            $retail_business = $entity_profile->getCell('B36')->getValue() ?? 0;
            $others = $entity_profile->getCell('B37')->getValue() ?? 0;

            $gross_turnover24 = $business_summary->getCell('C6')->getValue() ?? 0;
            $gross_turnover23 = $business_summary->getCell('D6')->getValue() ?? 0;

            $net_revenue24 = $annual_summary->getCell('B4')->getValue() ?? 0;
            $net_revenue23 = $annual_summary->getCell('C4')->getValue() ?? 0;

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
            echo "Error moving the uploaded file.";
        }
    } else {
        echo "File upload error: " . $file['error'];
    }
}