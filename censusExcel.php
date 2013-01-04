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
$excel->getActiveSheet()->getColumnDimension("B")->setWidth(6);

//Set the height of the headers to be kinda big (fit vertical text)
$excel->getActiveSheet()->getRowDimension(1)->setRowHeight(80);

//Set the headers
$excel->setActiveSheetIndex(0)
      ->setCellValue('A1',"Text from Online Profile")
      ->setCellValue('B1',"FieldName")
      ->setCellValue('C1',"Table Number")
      ->setCellValue('D1',"Computed Value")
      ->setCellValue('E1',"Query to get Data");

//Less typing please
$sheet = $excel->getActiveSheet();


//Set the 'row header' for the first section of information:
$sheet->setCellValue('A2',"Housing Demand");
$sheet->getStyle('A2:E2')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A2:E2')->getFill()->getStartColor()->setARGB('FFF79646');
$sheet->getStyle("A2:E2")->applyFromArray(array("font" => array( "bold" => true)));

/*
The A column is a bunch of pre definied text, I'm sorely tempted 
to load it from a file, same with the field names, and tables.
I could create an upload tool that would send the file to play with
over to this file, but... I see issues with this.
1. The file sent over would have to be opened and read.
   a. This would slow the script down, and php has a timeout
   b. This file is already intensive enough so I dont want to add
      more memory to it.
   c. The excel sheet coming from this file is a one off file every
      time the census is updated (not often). Maintaining another
      file besides this one just adds more overhead that's not neccesary
2. Assuming the other file would be excel, the phpexcel is lazy load but
   it'd still be intensive for memory and would slow down and once again
   the timeout on php is something like 30 seconds and we needs to query
   the api with that time as well, so that might not be a good idea.
*/

//Create the api interface
require_once('APIInterface.php');
$API = new APIInterface();

//Set up the initial Query default
$API->setSurvey('acs5')
	->setTable('B25039_001E')
	->setState('46') //46 is vermonts code
	->setYear('2010')
	->constructQuery();


//We place this into the session so the helper functon can see it
$_SESSION['sheet'] = $sheet;
//Define a helper function to assist us:
function fillCell($cellindex,$val)
{
	$_SESSION['sheet']->setCellValue($cellindex, $val);
}

//This is going to get long very fast
fillCell('A3',"Total Population");
fillCell('B3',"popn");

//When we figure out what data to use for rows 3-17 I'll come back


$result = $API->runQuery();

print_r($result);


//save
$objWriter = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME));


?>