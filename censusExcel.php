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

//Row64
fillRow(64,array('Median household income (Census)','hhai','B19013;'));

//Row65
fillRow(65,array('...homeowner households','hhio','B25119'));

//Row66
fillRow(66,array('...renter households','hhir','B25119' ));

//Row67
fillRow(67,array('Median family income (Census)','fmai','B19019'));

//Row68
fillRow(68,array('...for 2-person families','fmi2','B19019'));

//Row69
fillRow(69,array('...for 3-person families','fmi3','B19019'));

//Row70
fillRow(70,array('...for 4-person families','fmi4','B19019'));

//Row71
fillRow(71,array('...for 5-person families','fmi5','B19019'));

//Row72
fillRow(72,array('...for 6-person families','fmi6','B19019'));

//Row73
fillRow(73,array('...for 7-person families','fmi7','B19019'));

//Row74
fillRow(74,array('Median household income for family of four(HUD)','mf4i','B19019'));

//Row75
fillRow(75,array('Median family adjusted gross income','agif'));

//Row76
fillRow(76,array('Annual average wage(VT DET)'));

//Row77
fillRow(77,array('Per capita income (Census)','prci','B19301'));

//Row78
fillRow(78,array('Labor Force (VT DET)','lafo'));

//Row79
fillRow(79,array('...employed','aemp'));

//Row80
fillRow(80,array('...unemployed','uemp'));

//Row81
fillRow(81,array('...unemployment rate','uemr' ));

//Row83 is a header row
$sheet->setCellValue('A83',"Housing Stock");
$sheet->getStyle('A83:E83')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A83:E83')->getFill()->getStartColor()->setARGB('FFF79646');
$sheet->getStyle("A83:E83")->applyFromArray(array("font" => array( "bold" => true)));

//Row84
fillRow(84,'Total housing units','tohu');

//Row85
fillRow(85,array('...owner occupied'));

//Row86
fillRow(86,array('...renter occupied','thur'));

//Row87
fillRow(87,array('...vacant housing units','vahu'));

//Row88
fillRow(88,array('   ...for seasonal, recreational, occasional use','vhse'));

//Row89
fillRow(89,array('   ...for rent','vhre'));

//Row90
fillRow(90,array('   ...for sale only','vhsa'));

//Row91
fillRow(91,array('Housing by Units in Structure'));

//Row92
fillRow(92,array('...All Housing Units','ustt','B25024'));

//Row93
fillRow(93,array('   ...in Buildings with 1 Unit','ust1','B25024'));

//Row94
fillRow(94,array('   ...in Buildings with 2 Units','ust2','B25024'));

//Row95
fillRow(95,array('   ...in Buildings with 3+ Units','ust3','B25024'));

//Row96
fillRow(96,array('   ...that Are Mobile Homes','ustm','B25024'));

//Row97
fillRow(97,array('   ...that Are Boats, RVs, Vans or Other','usto','B25024'));

//Row98
fillRow(98,array('...Owner-Occupied Housing Units','usot','B25032'));

//Row99
fillRow(99,array('   ...in Buildings with 1 Unit','uso1','B25032'));

//Row100
fillRow(100,array('   ...in Buildings with 2 Units','uso2','B25032'));

//Row101
fillRow(101,array('   ...in Buildings with 3+ Units','uso3','B25032'));

//Row102
fillRow(102,array('   ...that Are Mobile Homes','usom','B25032'));

//Row103
fillRow(103,array('   ...that Are Boats, RVs, Vans or Other','usoo','B25032'));

//Row104
fillRow(104,array('...Renter-Occupied Housing Units','usrt','B25032'));

//Row105
fillRow(105,array('   ...in Buildings with 1 Unit','usr1','B25032'));

//Row106
fillRow(106,array('   ...in Buildings with 2 Units','usr2','B25032'));

//Row107
fillRow(107,array('   ...in Buildings with 3+ Units','usr3','B25032'));

//Row108
fillRow(108,array('   ...that are Mobile Homes','usrm','B25032'));

//Row109
fillRow(109,array('   ...that are Boats, RVs, Vans or Other','usro','B25032'));

//Row110
fillRow(110,array('Building permits estimated (total units)','bupe'));

//Row111
fillRow(111,array('...single family unit permits reported','bpsf'));

//Row112
fillRow(112,array('...multifamily unit permits reported','bpmf'));

//Row113
fillRow(113,array('Home heating fuel','','B25040'));

//Row114
fillRow(114,array('...All Occupied Housing Units','ftot','B25040'));

//Row115
fillRow(115,array('   ...gas','futg','B25040'));

//Row116
fillRow(116,array('   ...electric','fute','B25040'));

//Row117
fillRow(117,array('   ...fuel oil, kerosene','futo','B25040'));

//Row118
fillRow(118,array('   ...wood','fuow','B25040'));

//Row119
fillRow(119,array('   ...all other fuels','futt','B25040'));

//Row120
fillRow(120,array('   ...no fuel used','futn','B25040'));

//Row121
fillRow(121,array('...Owner-Occupied Housing Units','','B25117'));

//Row122
fillRow(122,array('   ...gas','fuog','B25117'));

//Row123
fillRow(123,array('   ...electric','fuoe','B25117'));

//Row124
fillRow(124,array('   ...fuel, oil kerosene','fuoo','BB25117'));

//Row125
fillRow(125,array('   ...wood','fuow','B25117'));

//RoW126
fillRow(126,array('   ...all other fuels','fuot','B25117'));

//Row127
fillRow(127,array('   ...no fuel used','fuon','B25117'));

//Row128
fillRow(128,array('...Renter-Occupied Housing Units'));

//Row129
fillRow(129,array('   ...gas','furg','B25117'));

//Row130
fillRow(130,array('   ...electric','fure','B25117'));

//Row131
fillRow(131,array('   ...fuel, oil kerosene','furo','BB25117'));

//Row132
fillRow(132,array('   ...wood','furw','B25117'));

//RoW133
fillRow(133,array('   ...all other fuels','furt','B25117'));

//Row134
fillRow(134,array('   ...no fuel used','furn','B25117'));

//Row135 
fillCell('A135','	Homeownership costs (hoCosts1 (pink), hoCosts2 (orange), hoCosts3 (yellow), hoCosts4 (green))');
$sheet->getStyle('A83:E83')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A83:E83')->getFill()->getStartColor()->setARGB('FFFFC000');

//Apply some styling to the cells for home ownership color coded area
$sheet->getStyle('A136:E167')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A136:E138')->getFill()->getStartColor()->setARGB('FFFF99CC');
$sheet->getStyle('A140:E142')->getFill()->getStartColor()->setARGB('FFFF99CC');
$sheet->getStyle('A144:E146')->getFill()->getStartColor()->setARGB('FFFF99CC');
$sheet->getStyle('A139:E139')->getFill()->getStartColor()->setARGB('FFFFCC99');
$sheet->getStyle('A143:E143')->getFill()->getStartColor()->setARGB('FFFFCC99');
$sheet->getStyle('A147:E147')->getFill()->getStartColor()->setARGB('FFFFCC99');
$sheet->getStyle('A148:E150')->getFill()->getStartColor()->setARGB('FFFFFF99');
$sheet->getStyle('A151:E154')->getFill()->getStartColor()->setARGB('FFFFFF99');
$sheet->getStyle('A156:E158')->getFill()->getStartColor()->setARGB('FFFFFF99');
$sheet->getStyle('A151:E151')->getFill()->getStartColor()->setARGB('FFCCFFCC');
$sheet->getStyle('A155:E155')->getFill()->getStartColor()->setARGB('FFCCFFCC');
$sheet->getStyle('A159:E159')->getFill()->getStartColor()->setARGB('FFCCFFCC');
$sheet->getStyle('A161:E163')->getFill()->getStartColor()->setARGB('FFFFCC99');
$sheet->getStyle('A165:E167')->getFill()->getStartColor()->setARGB('FFCCFFCC');


//Row136
fillRow(136,array('Number of primary residences sold','nphs'));

//Row137
fillRow(137,array('... single family homes','nsfs'));

//Row138
fillRow(138,array('... condominiums','ncon'));

//Row139
fillRow(139,array('... mobile homes with land','nmhl'));

//Row140
fillRow(140,array('Average Price of residences sold','aphs'));

//Row141
fillRow(141,array('... single family homes','asfs'));

//Row142
fillRow(142,array('... condominiums','acon'));

//Row143
fillRow(143,array('... mobile homes with land','amhl'));

//Row144
fillRow(144,array('Median price of primary residences sold','mphs'));

//Row145
fillRow(145,array('... single family homes','msfs'));

//Row146
fillRow(146,array('... condominiums','mcon'));

//Row147
fillRow(147,array('... mobile homes with land','mmhl'));

//Row148
fillRow(148,array('Number of vacation residences sold','nvac'));

//Row149
fillRow(149,array('... single family vacation homes','nvsf'));

//Row150
fillRow(150,array('... vacation condominiums','nvcn'));

//Row151
fillRow(151,array('... mobile homes with land','nvml'));

//Row152
fillRow(152,array('Average price of vacation residences sol','avac'));

//Row153
fillRow(153,array('... single family vacation homes','avsf'));

//Row154
fillRow(154,array('... vacation condominiums','avcn'));

//Row155
fillRow(155,array('... mobile homes with land','avml'));

//Row156
fillRow(156,array('Median price of vacation residences sold','mvac'));

//Row157
fillRow(157,array('... single family vacation homes','mvsf'));

//Row158
fillRow(158,array('... vacation condominiums','mvcn'));

//Row159
fillRow(159,array('... mobile homes with land','mvml'));

//Row160
fillRow(160,array('Primary residence mobile homes sold without land'));

//Row161
fillRow(161,array('   ... number of sales','nmho'));

//Row162
fillRow(162,array('   ... average price','amho'));

//Row163
fillRow(163,array('   ... median price','mmho'));

//Row164
fillRow(164,array('Vacation residence mobile homes sold without land'));

//Row165
fillRow(165,array('   ... number of sales','nvmo'));

//Row166
fillRow(166,array('   ... average price','avmo'));

//Row167
fillRow(167,array('   ... median price','mvmo'));

//Header for Ho Costs 5
$sheet->setCellValue('A168',"HO Costs 5");
$sheet->getStyle('A168:E168')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A168:E168')->getFill()->getStartColor()->setARGB('FFF79646');
$sheet->getStyle("A168:E168")->applyFromArray(array("font" => array( "bold" => true)));

//Row169
fillRow(168,array('Median monthly owner costs','octo','B25088'));

//Row170
fillRow(170,array('... with mortage','ocwi','B25088'));

//Row171
fillRow(171,array('... without mortage','ocwo','B25088'));

//Row172
fillRow(172,array('... as percentage of household income','ochi','B25092'));

//Row173
fillRow(173,array('Owner-Occupied housing units','sphu','B25003'));

//Row174
fillRow(174,array('... at or above 30% of household income','oca3','B25091'));

//Row175
fillRow(175,array('... at or above 50% of household income','oca5','B25091'));

//Row176
fillRow(176,array('... with a mortage','Ocmt','B25091'));

//Row177
fillRow(177,array('   ... with owner costs at or above 30% of household income','Ocm3','B25091'));

//Row178
fillRow(178,array('   ... with owner costs at or above 50% of household income','Ocm5','B25091'));

//Row179
fillRow(179,array('... without a mortage','Ocnt','B25091'));

//Row180
fillRow(180,array('   ... with owner costs at or above 30% of household income','Ocn3','B25091'));

//Row181
fillRow(181,array('   ... with owner costs at or above 50% of household income','Ocn5','B25091'));

//Row182
fillRow(182,array('Median value of owner occupied housing units','mvho'));

//Row183
fillRow(183,array('Municipal Tax Rate','mutr'));

//Row184
fillRow(184,array('Educational Tax Rate for Homesteads','eths'));

//Row185
fillRow(185,array('Educational Tax Rate for Non-Residential','etnr'));

//Row186
fillRow(186,array('Common Level of Appraisal Ratio','clar'));

//Row187
fillRow(187,array('Number of primary residences sold (YTD)','ytdn'));

//Row188
fillRow(188,array('Median price of primary residences sold (YTD)','ytdm'));

//Row189
fillRow(189,array('Average price of primary residences sold (YTD)','tyda'));

//Header for Rental housing costs
$sheet->setCellValue('A191',"Rental Housing Costs");
$sheet->getStyle('A191:E191')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A191:E191')->getFill()->getStartColor()->setARGB('FFF79646');
$sheet->getStyle("A191:E191")->applyFromArray(array("font" => array( "bold" => true)));

//Row192
fillRow(192,array('Median gross rent (all units)','megr','B25064'));

//Row193
fillRow(193,array('   ...as a percentage of household income','mgrp','B25071'));

//Row194
fillRow(194,array('Specified housing units with gross rent (total)','hugr','B25603'));

//Row195
fillRow(195,array('... at or above 30% of household income','mgra','B25070'));

//Row196
fillRow(196,array('... at or above 50% of household income','mgrc','B25070'));

//Row197
fillRow(197,array('Fair Market Rent (HUD)'));

//Row198
fillRow(198,array('...0 bedroom unit (40%)','fmr0'));

//Row199
fillRow(199,array('...1 bedroom unit (40%)','fmr1'));

///Row200
fillRow(200,array('...2 bedroom unit (40%)','fmr2'));

//Row201
fillRow(201,array('...3 bedroom unit (40%)','fmr3'));

//Row202
fillRow(202,array('...4 bedroom unit (40%)','fmr4'));

//Row203
fillRow(203,array('Median rents (HUD)'));

//Row204
fillRow(204,array('...0 bedroom unit (50%) - Median Rent','mer0'));

//Row205
fillRow(205,array('...1 bedroom unit (50%) - Median Rent','mer1'));

//Row206
fillRow(206,array('...2 bedroom unit (50%) - Median Rent','mer2'));

//Row207
fillRow(207,array('...3 bedroom unit (50%) - Median Rent','mer3'));

//Row208
fillRow(208,array('...4 bedroom unit (50%) - Median Rent','mer4'));

//Header for special needs
$sheet->setCellValue('A209',"Special Needs");
$sheet->getStyle('A209:E209')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A209:E209')->getFill()->getStartColor()->setARGB('FFF79646');
$sheet->getStyle("A209:E209")->applyFromArray(array("font" => array( "bold" => true)));

//Row210
fillRow(210,array('Supplemental security income recips','ssit'));

//Row211
fillRow(211,array('...younger than age 18','ssa1'));

//Row212
fillRow(212,array('...aged 16-64','ssa2'));

//Row213
fillRow(213,array('...aged 65 and older','ssa3'));

//Row214
fillRow(214,array('Monthly SSI payments in Vermont','sspt'));

//Row215
fillRow(215,array('...amount available for housing','ssph'));

//Row216
fillRow(216,array('... % income needed for efficiency','sspe'));

//Row217
fillRow(217,array('... % income needed for 1-bedroom','ssp1'));


//save
$objWriter = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME));


?>