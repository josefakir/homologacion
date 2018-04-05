<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$inputFileName = './DatosOportunidad.xlsx';

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);
$worksheet = $spreadsheet->getActiveSheet();
$i = 0;
$k = 0;
echo "<table border='1'>";
foreach ($worksheet->getRowIterator() as $row) {
    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(FALSE); // This loops through all cells,
    if($i>=6){
    echo "<tr>";
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

		}
		echo "</tr>";
	}
	$i++;
}
echo "</table>";
