<?php

/********************************************************
*Ethan Eldridge 1/4/2013
* 
* This file creates an excel sheet that pulls data from the
* api.census.gov website. This file requires the PHPExcel 
* library to function. 
*
*/

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

require_once dirname(__FILE__) . '/../Classes/PHPExcel.php';

//Create the excel object to use
$excel = new PHPExcel();

//Set Document Properties
$excel->getProperties()->setCreator("Ethan Eldridge")
					    ->setLastModifiedBy("Ethan Eldridge")
					    ->setTitle("Census Information")
						->setSubject("Census Information")
						->setDescription("Spreadsheet using information from api.census.gov")
						->setKeywords("Census,API,api.census.gov")
						->setCategory("census");
$excel->getActiveSheet()->setTitle("API Data Filled");


$excel->setActiveSheetIndex(0);

//Style the headers with color
$excel->getActiveSheet()->getStyle('A1:E1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$excel->getActiveSheet()->getStyle('A1:E1')->getFill()->getStartColor()->setARGB('FFFFFF99');
$excel->getActiveSheet()->getStyle('B1')->getAlignment()->setTextRotation(90);

//Set the widths of the header columns
foreach (array('A','C','D','E') as $letter) {
	$excel->getActiveSheet()->getColumnDimension($letter)->setAutoSize(true);	
}
$excel->getActiveSheet()->getColumnDimension("B")->setWidth(4);

//Set the height of the headers to be kinda big
$excel->getActiveSheet()->getRowDimension(1)->setRowHeight(80);



//Set the headers
$excel->setActiveSheetIndex(0)
      ->setCellValue('A1',"Text from Online Profile")
      ->setCellValue('B1',"FieldName")
      ->setCellValue('C1',"Table Number")
      ->setCellValue('D1',"Computed Value")
      ->setCellValue('E1',"Query to get Data");



//save
$objWriter = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME));


?>