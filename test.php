<?php
?>
<script type="text/javascript">
	function httpGet(url){
		var xmlHttp = null;
		xmlHttp = new XMLHttpRequest();
		//Method, url to query, async or not
		xmlHttp.open("GET",url,false);
		xmlHttp.send(null);
		alert(xmlHttp.responseText);
		return xmlHttp.responseText;
	}
</script>

<input type="button" value="Click Here" onClick=httpGet('http://api.census.gov/data/2010/sf1?key=51c2d5ad9882a926f52fca3d47d749a963fda0a8&get=PCT012A015,PCT012A119&for=state:01');>
<?php
?>

<html>
<head>
<title></title>

</head>
<body>
<?php
date_default_timezone_set('America/Los_Angeles');
//Ethan Eldridge ( ejayeldridge@gmail.com )
//October 24 2012, file to offer a simple form based querying for the census API
//This file will more than likely be called with an include, so  we won't make it fully modular
//We'll use a class structure to store the data though

		
//http://nadeausoftware.com/articles/2007/07/php_tip_how_get_web_page_using_fopen_wrappers
function get_web_page( $url )
{
    $options = array( 'http' => array(
        'user_agent'    => 'spider',    // who am i
        'max_redirects' => 10,          // stop after 10 redirects
        'timeout'       => 120,         // timeout on response
    ) );
    $context = stream_context_create( $options );
    $page    = @file_get_contents( $url, false, $context );
 
    $result  = array( );
    if ( $page != false )
        $result['content'] = $page;
    else if ( !isset( $http_response_header ) )
        return null;    // Bad url, timeout

    // Save the header
    $result['header'] = $http_response_header;

    // Get the *last* HTTP status code
    $nLines = count( $http_response_header );
    for ( $i = $nLines-1; $i >= 0; $i-- )
    {
        $line = $http_response_header[$i];
        if ( strncasecmp( "HTTP", $line, 4 ) == 0 )
        {
            $response = explode( ' ', $line );
            $result['http_code'] = $response[1];
            break;
        }
    }
 
    return $result;
}

class APIInterface{
	// //API Key for VHFA
	private $key = '51c2d5ad9882a926f52fca3d47d749a963fda0a8';
	//State to query for, default as VT
	private $stateNum = '46';
	//County to ask for, default all
	private $county = '*';
	//Town to ask for, default all
	private $town = '*';
	//Year to get data for, default last year
	private $year = '2010';
	//Survey to query against
	private $survey = 'acs5';
	//Table in survey to retrieve data from
	private $table = 'B25039_001E';
	//Base for making our queries
	private $qBase = 'http://api.census.gov/data/';
	//Actual Query
	private $query = "";
	
	//Constants for tables:
	// http://www.census.gov/developers/data/2010acs5_variables.xml
	
	
	public function __construct(){
		//Construct API query from defaults
		$this->constructQuery();
	}
	
	public function setYear($year = 2010){
		$this->year = $year;
	}
	
	public function setSurvey($survey = 'asc5'){
		$this->survey = $survey;
	}
	
	public function setTable($table = 'B25039_001E'){
		$this->table = $table;
	}
	
	public function setState($state = '46'){
		$this->state = $state;
	}
	
	//Runs the Query and returns the JSON
	public function runQuery(){
		
		echo 'Running Query: ' . $this->query . '<br />';
		$result = get_web_page($this->query);
		
		//Did we get something back?
		if($result['http_code'] == '200'){
			print_r (json_decode($result["content"]));
		}
		
	}
	
	public function constructQuery(){
		//Use the variables we've defined to perform out query.
		$this->query = $this->qBase . $this->year .'/'. $this->survey .'?key='.$this->key.'&get='.$this->table.'&for=state:'.$this->stateNum.'&for=county:'.$this->county.'&for=town:'.$this->town;
	}
	
	public function constructTownlessQuery(){
		//Create query without town
		$this->query = $this->qBase . $this->year .'/'. $this->survey .'?key='.$this->key.'&get='.$this->table.'&for=state:'.$this->stateNum.'&for=county:'.$this->county;
	}
	
	public function constructCountylessQuery(){
		//Create query without a county
		$this->query = $this->qBase . $this->year .'/'. $this->survey .'?key='.$this->key.'&get='.$this->table.'&for=state:'.$this->stateNum;
	}
	
	
}

//Now that I've defined my interface lets make one.
$API = new APIInterface();

?>

<form name="APIForm" action="" method= "POST">
Run query on 
<select name = "survey" >
	<option value = "acs5">acs5</option>
	<option value = "sf1">sf1</option>
</select>
survey, getting data from

<select name= "table" >
	<option value = "B25039_001E">Median Year Householder moved in (Totals)</option>
	<option value = "B25039_002E">Median Year Householder moved in (Owners)</option>
	<option value = "B25039_003E">Median Year Householder moved in (Renter)</option>
	<option value = "B25038_001E">Tenure by Year Householder Moved into Unit (Total)</option>
</select>

for the state of
<select name = "state">
	<option value="46">Vermont</option>
</select>
during the year 

<select name = "year">
	<option value="2010">2010</option>
</select>
 
 

<input type="submit" value="Run Query">

</form>

<?php

//If we submitted then run the query:
if(isset($_POST['survey'])){
	$API->setSurvey($_POST['survey']);
	$API->setTable($_POST['table']);
	$API->setState($_POST['state']);
	$API->setYear($_POST['year']);
	$API->constructQuery();
	$API->runQuery();
}




?>
<h1>API RUN</h1>
</body>
</html>