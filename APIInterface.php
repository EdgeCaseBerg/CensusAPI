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
		return $this;
	}
	
	public function setSurvey($survey = 'asc5'){
		$this->survey = $survey;
		return $this;
	}
	
	public function setTable($table = 'B25039_001E'){
		$this->table = $table;
		return $this;
	}
	
	public function setState($state = '46'){
		$this->state = $state;
		return $this;
	}
	
	//Runs the Query and returns the JSON
	public function runQuery(){
		
		$result = get_web_page($this->query);
		
		//Did we get something back?
		if($result['http_code'] == '200'){
			return (json_decode($result["content"]));
		}

		return null;
		
	}
	
	public function constructQuery(){
		//Use the variables we've defined to perform out query.
		$this->query = $this->qBase . $this->year .'/'. $this->survey .'?key='.$this->key.'&get='.$this->table.'&for=state:'.$this->stateNum.'&for=county:'.$this->county.'&for=town:'.$this->town;
		return $this;
	}
	
	public function constructTownlessQuery(){
		//Create query without town
		$this->query = $this->qBase . $this->year .'/'. $this->survey .'?key='.$this->key.'&get='.$this->table.'&for=state:'.$this->stateNum.'&for=county:'.$this->county;
	}
	
	public function constructCountylessQuery(){
		//Create query without a county
		$this->query = $this->qBase . $this->year .'/'. $this->survey .'?key='.$this->key.'&get='.$this->table.'&for=state:'.$this->stateNum;
	}

	public function getQuery(){
		return $this->query;
	}
	
	
}

?>