<?php
//header('Content-Type: application/json');

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use OntraportAPI\Ontraport;
$inputFileName = './DatosOportunidad.xlsx';

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);
$worksheet = $spreadsheet->getActiveSheet();
$i=0;
$servername = "127.0.0.1";
$username = "root";
$password = "root";
$dbname = "bmw";
$conn = new PDO("mysql:host=$servername;dbname=$dbname", $username, $password);
$conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$existentes = array();
$client = new OntraportAPI\Ontraport("2_165242_leJiCRh3x","OUnU0EmLXURF0kY");

foreach ($worksheet->getRowIterator() as $row) {
	if($i>5){
		$cellIterator = $row->getCellIterator();
	    $cellIterator->setIterateOnlyExistingCells(FALSE); // This loops through all cells,
	    $j=0;
	    $text = '';
	    foreach ($cellIterator as $cell) {
	    	$text .='|'.$cell->getValue();
	    }
	    $campos = explode("|",$text);
	   //insert


	    try {
	    	
	    	$sql = "INSERT INTO leads (id_oportunidad,codigo_ciclo_venta,ubicacion,id_usuario,codigo_marca,estado_oportunidad,fecha_creacion,fecha_cierre,nombre,email,phone_no,mobile_phone_no) VALUES (:id_oportunidad,:codigo_ciclo_venta,:ubicacion,:id_usuario,:codigo_marca,:estado_oportunidad,:fecha_creacion,:fecha_cierre,:nombre,:email,:phone_no,:mobile_phone_no)";
	    	$stmt = $conn->prepare($sql);
			    // use exec() because no results are returned
	    	$stmt->execute(
	    		array(
	    			':id_oportunidad' => $campos[1],
	    			':codigo_ciclo_venta' => $campos[2],
	    			':ubicacion' => $campos[3],
	    			':id_usuario' => $campos[4],
	    			':codigo_marca' => $campos[5],
	    			':estado_oportunidad' => $campos[6],
	    			':fecha_creacion' => $campos[7],
	    			':fecha_cierre' => $campos[8],
	    			':nombre' => $campos[0],
	    			':email' => $campos[10],
	    			':phone_no' => $campos[11],
	    			':mobile_phone_no' => $campos[12]
	    		)
	    	);
			 //   echo "New record created successfully";
	    }
	    catch(PDOException $e)
	    {
	    	//echo $sql . "<br>" . $e->getMessage();
	    }
	}
	$i++;
}
$emails = array();
$sql = "SELECT distinct(email) FROM leads";
$stmt = $conn->prepare($sql);
$stmt->execute();
$result = $stmt->fetchAll(PDO::FETCH_BOTH);
foreach ($result as $r) {
	# code...
	array_push($emails, $r['email']);
}
$emails = array_chunk($emails, 100);
foreach ($emails as $em) {
	$json= '';
	$count = count($em)-1;
	$i=0;
	foreach ($em as $e) {
		if($i<$count){
			$json .= '{"value":"'.$e.'"},';
		}else{
			$json .= '{"value":"'.$e.'"}';
		}
		$i++;
	}
	rtrim($json,',');
	$queryParams = array(
		"condition"     => '[{
			"field":{"field":"email"},
			"op":"IN",
			"value":{"list":[
				'.$json.'
				]}
			}]',
		"listFields" => "id,email"
	);
	try {
	    $response = $client->contact()->retrieveMultiple($queryParams);
	    $response = json_decode($response);
	    $response = $response->data;
	    foreach ($response as $r) {
	    	$nuevo = array(
	    		'id' => $r->id,
	    		'email' => $r->email
		    );
		    array_push($existentes, $nuevo);
	    }
	} catch (Exception $e) {
		print_r($e);
	}

}
foreach ($existentes as $ex) {
	$sql = "UPDATE leads set id_ontraport = :id_ontraport WHERE email = :email";
	$stmt = $conn->prepare($sql);
	try {
	if($ex['email']!=''){
		$stmt->execute(
			array(
				':id_ontraport' => $ex['id'],
				':email' => $ex['email']
			)
		);
		print_r($ex);
	}
	} catch (Exception $e) {
		print_r($e);
	}
}

$sql = "SELECT * FROM leads WHERE id_ontraport <> 'NULL' AND id_ontraport <> ''";
$stmt = $conn->prepare($sql);
$stmt->execute();
$result = $stmt->fetchAll(PDO::FETCH_BOTH);
echo "<pre>";
print_r($result);

foreach ($result as $ult) {
	try {
    $requestParams = array(
        "firstname" => "John",
        "lastname"  => "Doe",
        "email"     => "user@ontraport.com"
    );
    $response = $client->contact()->saveOrUpdate($requestParams);
	} catch (Exception $e) {
		
	}

	# code...
}

/* 

[1] => OP-0005891
    [2] => RT-LEAD
    [3] => CAMPOSE
    [4] => 10760
    [5] => BMW-MOT
    [6] => En Curso
    [7] => 43032
    [8] => 0
    [9] => MARCOS ACHAR CONTRERAS
    [10] => acharmarcos@gmail.com
    [11] => 5510486486
    [12] => 5510486486
*/