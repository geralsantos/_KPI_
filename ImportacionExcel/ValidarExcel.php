<?php
error_reporting(E_ALL);
ini_set('display_errors', 'On');
//if (($fichero = fopen($tmp_archivo, "r")) !== FALSE)
//{
	//$chk_ext = explode(".",$nombre_archivo);
	//if(strtolower($extension) == "xlsx")
    //{
	require 'phpExcel/PHPExcel/IOFactory.php';
function GetCuentas(){
	$upload_folder  = dirname(__FILE__).'/excels/';
	$nombre_archivo = $_FILES['archivo']['name'];
	$tipo_archivo   = $_FILES['archivo']['type'];
	$tamano_archivo = $_FILES['archivo']['size'];
	$tmp_archivo    = $_FILES['archivo']['tmp_name'];
	$extension		= pathinfo($nombre_archivo, PATHINFO_EXTENSION);
	$result=[];
	$fichero_subido = $upload_folder . basename($nombre_archivo);
	if (strtolower($extension) == "xlsx" || strtolower($extension) == "xls")
	{
		if (move_uploaded_file($tmp_archivo, $fichero_subido))
		{
			$objPHPExcel = PHPExcel_IOFactory::load($fichero_subido);
			$objPHPExcel = $objPHPExcel->setActiveSheetIndex(0);
			if (!empty(ValidarCampos($objPHPExcel)['Error']) )
			{
				$result = ValidarCampos($objPHPExcel)['Error'];
			}
			else{
				//----------Filas----------//
				$CuentasVer=[];
				$columna ='A';
				$dia = date('d');
			  	$mes = date('m');
			  	$anio = date('Y');
			  	$today = $anio."-".$mes."-".$dia ;
				$ultimaFila = $objPHPExcel->getHighestRow();
				for ($row=1; $row <=$ultimaFila ; $row++)
				{

					$lastColumn =$objPHPExcel->getHighestColumn();
					$lastColumn++;
					//cuenta	debe	haber	tipo doc	serie	correlativo	documento/dni	fecha emision	fecha vencimiento
					//----------Columnas----------//
					$i=0;
					$Cuentas=[];
					for ($column = 'A'; $column != $lastColumn; $column++)
					{
						if ($row==1) {
							/*$response[$i] = $Cabeceras[$i];
							$i++;*/
						}else{
							$value = $objPHPExcel->getCell($column.$row)->getCalculatedValue();
							$value = ($column =="E" || $column =="F" || $column =="G" || $column =="H") ? \PHPExcel_Style_NumberFormat::toFormattedString($value, 'DD/MM/YYYY') : $value;
							/*$value = $column =="B" || $column =="C" ? ($value !="" ? $value : 0) : $value;*/
							$response2[] = [($column.$row),$value];
							$i++;
						}
					}
					if ($row>1)
					{
						array_push($Cuentas, $response2 );

						array_push($CuentasVer,($Cuentas));
					}
					//----------Columnas----------//
				}
				//----------Filas----------//
				$result=array('Error'=>'','Success'=>$CuentasVer);
			}
			unlink($fichero_subido);
		}else{
			$result= array('Error' => ["El archivo no ha podido ser leido correctamente, por favor verifique que el archivo excel cumpla con todos los estandares"] );
		}

	}else{
		$result=array('Error'=>["Solo se permiten archivos de formato excel (.xlsx, .xls)"]);
	}
	print_r($result);
	return (json_encode($result));
}
function ValidarCampos ($objPHPExcel){
	$ultimaFila = $objPHPExcel->getHighestRow();
	$vacios =[];
	$Permitecolumnasvacias = array();
	$fechas=[];
	$numeros=[];
	$columns = ["A","B","C","D","F","G","H","I","J","K"];
	for ($row=2; $row <=$ultimaFila ; $row++)
	{
		/*$lastColumn =$objPHPExcel->getHighestColumn();
		$lastColumn++;*/
		$i=0;
		for ($column = 0; $column < count($columns); $column++)
		{
			$value =trim($objPHPExcel->getCell($columns[$column].$row)->getCalculatedValue());
			//valores vacios
			if (!in_array($columns[$column],$Permitecolumnasvacias)) {
				$d = $value == "" ? array_push($vacios,($columns[$column].$row)) : "";
			}
			//fechas
			//el array almacena si hay error en el formato de fecha
			$d = ($columns[$column] =="E" || $columns[$column] =="F" || $columns[$column] =="G" || $columns[$column] =="H") ? (validateDate( \PHPExcel_Style_NumberFormat::toFormattedString($value, 'DD/MM/YYYY') ) ? $value : array_push($fechas,($columns[$column].$row) ) ) : $value ;
			//solo numeros
		/*	$d = ($column =="B" || $column =="C") && $value!="" ? (is_numeric(str_replace(",","",$value)) ? $value : array_push($numeros,($column.$row) ) ) : $value;
*/
		}
	}
	$dia = date('d');
  	$mes = date('m');
  	$anio = date('Y');
  	$today = $dia."/".$mes."/".$anio ;
	$contenido 	=[$vacios,$fechas,$numeros];
	$msg		=["No se permiten valores nulos en las siguientes celdas del archivo excel: ","Se han encontrado fechas que no cumplen con el formato (ex: ".$today.") verificar las siguientes celdas en el archivo excel: ","No se permiten valores de tipo texto, por favor verificar las siguientes celdas en el archivo excel: "];
	$zeno		=[];
	for ($i=0; $i <=count($contenido) ; $i++)
	{
		if (!empty($contenido[$i]))
		{
			array_push($zeno,$msg[$i].implode(", ", $contenido[$i]) );
		}
	}
	$result =  array('Error' =>  $zeno);
	return $result;
}
function validateDate($date, $format = 'd/m/Y')
{
    $d = DateTime::createFromFormat($format, $date);
    return $d && $d->format($format) == $date;
}

    //}
//}
GetCuentas();
?>
