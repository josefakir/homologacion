<?php
header('Content-Type: application/json');

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use OntraportAPI\Ontraport;
$inputFileName = './DatosOportunidad.xlsx';

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);
$worksheet = $spreadsheet->getActiveSheet();
$i = 0;
$k = 0;
$emails = array();
$existentes = array();
foreach ($worksheet->getRowIterator() as $row) {
    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(FALSE); // This loops through all cells,
    if($i>=6){
	    $j=0;
		foreach ($cellIterator as $cell) {
			switch ($j) {
				case 0:
					$id_oportunidad = $cell->getValue();
				break;
				case 1:
					$codigo_ciclo_venta = $cell->getValue();
				break;
				case 2:
					$ubicacion = $cell->getValue();
				break;
				case 3:
					$id_usuario = $cell->getValue();
				break;
				case 4:
					$codigo_marca = $cell->getValue();
				break;
				case 5:
					$estado_oportunidad = $cell->getValue();
				break;
				case 6:
					$fecha_creacion = $cell->getValue();
				break;
				case 7:
					$fecha_cierre = $cell->getValue();
				break;
				case 8:
					$name = $cell->getValue();
				break;
				case 9:
					$email = $cell->getValue();
				break;
				case 10:
					$phone = $cell->getValue();
				break;
				case 11:
					$mobile = $cell->getValue();
				break;
			}
			$j++;
			if($j==12){
				$j=0;
			}
		}
		if(!empty($email)){
			array_push($emails, $email);
			///crear, la parte de homologación en ontraport;
		}
	}
	$i++;
}
$emails = array_unique($emails);
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
	$client = new OntraportAPI\Ontraport("2_165242_leJiCRh3x","OUnU0EmLXURF0kY");
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
print_r($existentes);


