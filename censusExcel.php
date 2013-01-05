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

//Create a helper function to fill a whole row
function fillRow($rowNum,$vals){
	$i = 0;
	$len = count($vals);
	foreach(array('A','B','C','D','E') as $col){
		if($i < $len){
			fillCell($col . $rowNum, $vals[$i]);	
		}
		$i++; 	
	}
}
//This is going to get long very fast
fillCell('A4',"Total Population");
fillCell('B4',"popn");

//Row5
fillRow(5,array('... in occupied housing units','pocc'));

//Row6
fillRow(6,array('    ... owner occupied','pown'));

//Row7
fillRow(7,array('    ... renter occupied','pren'));

//Row8
fillRow(8,array('Total group quarters population','grou'));

//Row9
fillRow(9,array('Number of households','hhld'));

//Row10
fillRow(10,array('... owning home','thuo'));

//Row11
fillRow(11,array('... renting home','thur'));

//Row12
fillRow(12,array('... number of families','fams'));

//Row13
fillRow(13,array('Average household size','ahhs'));

//Row14
fillRow(14,array('... in owner occupied housing units','hhso'));

//Row15
fillRow(15,array('... in renter occupied housing units','hhsr'));

//Row16
fillRow(16,array('Average family size','afas'));

//Row17
fillRow(17,array("Owner-Occupied Units","croo","B25014"));

//Row18
$API->setTable('B25014I_003E,B25014A_002E,B25014A_003E,B25014H_003E,B25014C_003E,B25014D_003E,B25014E_003E,B25014F_003E,B25014G_003E');
$API->constructQuery();
$result = $API->runQuery();
$computedValue = 0;
for($i=0; $i <  count($result[1])-1; $i++) {
	$computedValue += $result[1][$i];
}
fillRow(18,array("..with 1.01 or more people per room","crom",'B25014',$computedValue,$API->getQuery()));

//Row 19
fillRow(19,array('Renter-Occupied Units','crro','B25014'));

//Row 20
fillRow(20,array("..with 1.01 or more people per room","crom",'B25014'));

//Row 21
fillRow(21,array('Year Householder Moved Into Unit'));

//Row22
fillRow(22,array('...For Owner-Occupied Units'));

//Row23
fillRow(23,array('   ...2005 or later','ymo5','B25038'));

//Row24
fillRow(24,array('   ...2000 to 2004','ymo4','B25038'));

//Row25
fillRow(25,array('   ...1990 to March 2000','ymo9','B25038'));

//Row26
fillRow(26,array('   ...1980 to 1989','ymo8','B25038'));

//Row27
fillRow(27,array('   ...1970 to 1980','ymo7','B25038'));

//Row28
fillRow(28,array('   ...1969 or Earlier','ymo6','B25038'));

//Row29
fillRow(29,array('...For Renter-Occupied Units:'));

//Row30
fillRow(30,array('   ...2005 or later','ymr5','B25038'));

//Row31
fillRow(31,array('   ...2000 to 2004','ymr4','B25038'));

//Row32
fillRow(32,array('   ...1990 to March 2000','ymr9',	'B25038'));

//Row33
fillRow(33,array('   ...1980 to 1989','ymr8','B25038'));

//Row34
fillRow(34,array( '   ...1970 to 1979','ymr7','B25038'));

//Row35
fillRow(35,array('   ...1969 or Earlier','ymr6','B25038'));

//Row36
fillRow(36,array('Median Year Householder Moved Into Unit'));

//Row37
fillRow(37,array('...for all Occupied Units','myma','B25039'));

//Row38
fillRow(38,array('   ...Owner-Occupied','mymo','B25039'));

//Row39
fillRow(39,array('   ...Renter-Occupied','mymr','B25039'));

//Row40
fillRow(40,array('Total workers 16 years of age and over','woto'));

//Row41
fillRow(41,array('... working outside town or city of residence ','wotn','B08009'));

//Row42
fillRow(42,array('... working outside county of residence','woco','B08007'));

//Row43
fillRow(43,array('... working outside Vermont','wost','B08007'));

//Row44 Header for ability to afford
$sheet->setCellValue('A44',"Ability to Afford");
$sheet->getStyle('A44:E44')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A44:E44')->getFill()->getStartColor()->setARGB('FFF79646');
$sheet->getStyle("A44:E44")->applyFromArray(array("font" => array( "bold" => true)));

//Row45
fillCell('A45','Housing Wage');

//Row46
fillRow(46,array('... 0 bedroom unit','hwa0'));

//Row47
fillRow(47,array('... 1 bedroom unit','hwa1'));

//Row48
fillRow(48,array('... 2 bedroom unit','hwa2'));

//Row49
fillRow(49,array('... 3 bedroom unit','hwa3'));

//Row50
fillRow(50,array('... 4 bedroom unit','hwa4'));

//Row51
fillRow(51,array('housing wage as a percentage of the state minimum wage'));

//Row52
fillRow(52,array('... 0 bedroom unit','hwp0'));

//Row53
fillRow(53,array('... 1 bedroom unit','hwp1'));

//Row54
fillRow(54,array('... 2 bedroom unit','hwp2'));

//Row55
fillRow(55,array('... 3 bedroom unit','hwp3'));

//Row56
fillRow(56,array('... 4 bedroom unit','hwp4'));

//Row57
fillRow(57,array('Income needed to afford an apartmnet at HUD\'s FMR'));

//Row58
fillRow(58,array('... 0 bedroom unit','ina0'));

//Row59
fillRow(59,array('... 1 bedroom unit','ina1'));

//Row60
fillRow(60,array('... 2 bedroom unit','ina2'));

//Row61
fillRow(61,array('... 3 bedroom unit','ina3'));

//Row62
fillRow(62,array('... 4 bedroom unit','ina4'));

//Heading for ability to afford 2
$sheet->setCellValue('A63',"Ability to Afford");
$sheet->getStyle('A63:E63')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A63:E63')->getFill()->getStartColor()->setARGB('FFF79646');
$sheet->getStyle("A63:E63")->applyFromArray(array("font" => array( "bold" => true)));



//save
$objWriter = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME));


?>