<?php

//if (($fichero = fopen($tmp_archivo, "r")) !== FALSE)
//{
	//$chk_ext = explode(".",$nombre_archivo);
	//if(strtolower($extension) == "xlsx")
    //{
	require 'phpExcel/PHPExcel/IOFactory.php';
function GetCuentas(){
	$upload_folder  = dirname(__DIR__.PHP_EOL).'/Controles/DetDoc/Excel/';
	$nombre_archivo = $_FILES['archivo']['name'];
	$tipo_archivo   = $_FILES['archivo']['type'];
	$tamano_archivo = $_FILES['archivo']['size'];
	$tmp_archivo    = $_FILES['archivo']['tmp_name'];
	$extension		= pathinfo($nombre_archivo, PATHINFO_EXTENSION);
	$result=[];
		$fichero_subido = $upload_folder . basename($nombre_archivo);
	if (strtolower($extension) == "xlsx" || strtolower($extension) == "xls" || strtolower($extension) == "csv")
	{
		if (move_uploaded_file($tmp_archivo, $fichero_subido))
		{
			$objPHPExcel = PHPExcel_IOFactory::load($fichero_subido);
			$objPHPExcel = $objPHPExcel->setActiveSheetIndex(0);
			$Cabeceras =["Cuenta","Debe","Haber","TipoDoc","Serie","Correlativo","NroDocumento","FEmision","FVencimiento"]; // cabeceras del excel
			if (!empty(ValidarCampos($objPHPExcel,$value,$Cabeceras)['Error']) )
			{
				$result = ValidarCampos($objPHPExcel,$value,$Cabeceras);
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
							$response[$i] = $Cabeceras[$i];
							$i++;
						}else{
							$value = $objPHPExcel->getCell($column.$row)->getCalculatedValue();
							$value = ($column =="H" || $column =="I") ? ($column=="I" && $value == ''?$today:\PHPExcel_Style_NumberFormat::toFormattedString($value, 'YYYY-MM-DD'))   : $value;
							$value = $column =="B" || $column =="C" ? ($value !="" ? $value : 0) : $value;
							$response2[$response[$i]] = [($column.$row),$value];
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

	return (json_encode($result));
}

    //}
//}
function ValidarCampos ($objPHPExcel,$value,$cabeceras){
	$columna ='A';
	$ultimaFila = $objPHPExcel->getHighestRow();
	$vacios =[];
	$fechas=[];
	$numeros=[];
	for ($row=2; $row <=$ultimaFila ; $row++)
	{
		$lastColumn =$objPHPExcel->getHighestColumn();
		$lastColumn++;
		$i=0;
		for ($column = 'A'; $column < $lastColumn-1; $column++)
		{
			$value =trim($objPHPExcel->getCell($column.$row)->getCalculatedValue());
			//valores vacios
			$d = $value == "" && ($column !="B" && $column !="C" && $column !="I") ? array_push($vacios,($column.$row)) : "";
			//fechas
			//el array almacena si hay error en el formato de fecha
			$d = ($column =="H" || $column =="I" && $value !="" ) ? (validateDate( \PHPExcel_Style_NumberFormat::toFormattedString($value, 'YYYY/MM/DD') ) ? $value : array_push($fechas,($column.$row) ) ) : $value ;
			//solo numeros
			$d = ($column =="B" || $column =="C") && $value!="" ? (is_numeric(str_replace(",","",$value)) ? $value : array_push($numeros,($column.$row) ) ) : $value;

		}
	}
	$dia = date('d');
  	$mes = date('m');
  	$anio = date('Y');
  	$today = $anio."/".$mes."/".$dia ;

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
function validateDate($date, $format = 'Y/m/d')
{
    $d = DateTime::createFromFormat($format, $date);
    return $d && $d->format($format) == $date;
}

function GetDataFromExcel(){

	$upload_folder  = dirname(__DIR__.PHP_EOL).'/Controles/DetDoc/Excel/';
	$nombre_archivo = $_FILES['archivo']['name'];
	$tipo_archivo   = $_FILES['archivo']['type'];
	$tamano_archivo = $_FILES['archivo']['size'];
	$tmp_archivo    = $_FILES['archivo']['tmp_name'];
	$extension		= pathinfo($nombre_archivo, PATHINFO_EXTENSION);
	$result=[];
	$fichero_subido = $upload_folder . basename($nombre_archivo);
	$control 		= $_POST["control"];
	if (strtolower($extension) == "xlsx")
	{
		if (move_uploaded_file($tmp_archivo, $fichero_subido))
		{
			$objPHPExcel = PHPExcel_IOFactory::load($fichero_subido);
			$objPHPExcel = $objPHPExcel->setActiveSheetIndex(0);
			switch ($_REQUEST["Accion"]) {
				case 'SocioNegocio':
					$Cabeceras =["numero_doc","tipo_doc","cliente","proveedor","trabajador","telefono","correo","direcciones","contactos"];
					$result_ = ValidarSocioNegocio($objPHPExcel,$value,$Cabeceras,$control);

					break;
				case 'ListaPrecios':
					$Cabeceras =["codigo_producto","prodserv","moneda_venta","valor_compra"];
					$lastColumn =$objPHPExcel->getHighestColumn();
					$lastColumn++;
					$i=0;
					for ($column = 'A'; $column != $lastColumn; $column++)
					{
						$Cabeceras[$i] = ($column!="A" && $column!="B" && $column!="C" && $column!="D") ? trim($objPHPExcel->getCell($column.'1')->getCalculatedValue()) : $Cabeceras[$i];
						$i++;
					}
					$result_ = ValidarListaPrecios($objPHPExcel,$value,$Cabeceras,$control);
					break;
				default:
					$Cabeceras =["Codigo","Descripcion","Serie"];
					$result_ = ValidarFormato($objPHPExcel,$value,$Cabeceras,$control);
					break;
			}
			if (!empty($result_["Error"]))
			{
				$result = $result_;
			}
			else{
				//----------Filas----------//
				$getData=[];
				$columna ='A';
				$ultimaFila = $objPHPExcel->getHighestRow();
				for ($row=1; $row <=$ultimaFila ; $row++)
				{
					$lastColumn =$objPHPExcel->getHighestColumn();
					$lastColumn++;
					//----------Columnas----------//
					$i=0;
					$Series=[];
					if ($_REQUEST["Accion"]=="ListaPrecios") {
						$response2=[];
					}
					for ($column = 'A'; $column != $lastColumn; $column++)
					{
						if ($row==1) { // titulos del excel
							$response[$i] = $Cabeceras[$i];
							$i++;
						}else{
							$value = $objPHPExcel->getCell($column.$row)->getCalculatedValue();
							if ($_REQUEST["Accion"]=="ListaPrecios") {
								if (($response[$i]!="codigo_producto" && $response[$i]!="prodserv" && $response[$i]!="moneda_venta" && $response[$i]!="valor_compra")) {
									$response2["ListaPrecios"][] = [($column.$row),$value,$response[$i],($column.'1')]; //['A1','Valor']
								}else{
									$response2[$response[$i]] = [($column.$row),$value,$response[$i]]; //['A1','Valor']
								}
							}else{
								$response2[$response[$i]] = [($column.$row),$value]; //['A1','Valor']
							}
							$i++;
						}
					}
					if ($row>1)//pasa de la fila de los titulos
					{
						array_push($getData,$response2);
					}
					//----------Columnas----------//
				}
				//----------Filas----------//
				$result=array('Error'=>'','Success'=>$getData);
			}
			unlink($fichero_subido);
		}else{
			$result= array('Error' => ["El archivo no ha podido ser leido correctamente, por favor verifique que el archivo excel cumpla con todos los estandares"] );
		}

	}else{
		$result=array('Error'=>["Solo se permiten archivos de formato excel (.xlsx)"]);
	}

	return (json_encode($result));
}
function ValidarSocioNegocio ($objPHPExcel,$value,$cabeceras,$control){
	$columna ='A';
	$ultimaFila = $objPHPExcel->getHighestRow();
	$vacios =[];$rucdigitos=[];$cliprovtrab=[];$repiteexcel=[];$dataexcel=[];
	for ($row=2; $row <=$ultimaFila ; $row++)
	{
		$lastColumn =$objPHPExcel->getHighestColumn();
		$lastColumn++;
		$i=0;

		for ($column = 'A'; $column < $lastColumn; $column++)
		{

			$value =trim($objPHPExcel->getCell($column.$row)->getCalculatedValue());
			//valores vacios
			$d =  $value == "" && ($column!="F" && $column!="G" && $column!="H" && $column!="I") ? array_push($vacios,($column.$row)) : "";
			$d = trim($objPHPExcel->getCell("B".$row)->getCalculatedValue()) == "43" && $value != "" && strlen($value) != 11 && $column=="A" ? array_push($rucdigitos,($column.$row)) : "";
			$d =  $value != "" && ($value != "1" && $value != "0") && ($column=="C" || $column=="D" || $column=="E") ? array_push($cliprovtrab,($column.$row)) : "";
			if (trim($objPHPExcel->getCell("B".$row)->getCalculatedValue()) == "43" && $value != "" && $column=="A") {
				if (in_array($value, $dataexcel)) {
					array_push($repiteexcel,($column.$row));
				}
					array_push($dataexcel,$value);
			}
		}
	}
	$contenido 	=[$vacios,$rucdigitos,$cliprovtrab,$repiteexcel];
	$msg		=["No se permiten valores nulos en las siguientes celdas del archivo excel: ","El RUC no tiene la cantidad de 11 digitos en las siguientes celdas del archivo excel: ","Las columnas 'C','P','T' solo permite valores '1' y '0' verificar en las siguientes celdas del archivo excel: ","No se puede repetir el número de RUC, favor de verificar las siguientes celdas del excel: "];
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
function ValidarListaPrecios ($objPHPExcel,$value,$cabeceras,$control){
	$columna ='A';
	$ultimaFila = $objPHPExcel->getHighestRow();
	$vacios =[];$numeros=[];$repiteexcel=[];$dataexcel=[];
	for ($row=2; $row <=$ultimaFila ; $row++)
	{
		$lastColumn =$objPHPExcel->getHighestColumn();
		$lastColumn++;
		$i=0;

		for ($column = 'A'; $column < $lastColumn; $column++)
		{

			$value =trim($objPHPExcel->getCell($column.$row)->getCalculatedValue());
			//valores vacios
			$d =  $value == "" ? array_push($vacios,($column.$row)) : "";
			$d = ($column !="A" && $column !="B") && $value!="" ? ($value!="*" ? (is_numeric(str_replace(",","",$value)) ? $value : array_push($numeros,($column.$row) ) ):$value) : $value;
			if ($value != "" && $column=="A") {
				$newvalue = ($value."||".trim(strtolower($objPHPExcel->getCell("B".$row)->getCalculatedValue()) )."||".trim($objPHPExcel->getCell("C".$row)->getCalculatedValue() ) );
				if (in_array($newvalue, $dataexcel)) {
					array_push($repiteexcel,($row));
				}
				array_push($dataexcel,$newvalue );
			}
		}
	}
	$contenido 	=[$vacios,$numeros,$repiteexcel];
	$msg		=["No se permiten valores nulos en las siguientes celdas del archivo excel: ","Solo se permiten valores numericos en las siguientes celdas del archivo excel: ","Se está repitiendo el codigo de producto con la misma moneda, favor de verificar las siguientes filas del excel: "];
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
function ValidarFormato ($objPHPExcel){
	$columna ='A';
	$ultimaFila = $objPHPExcel->getHighestRow();
	$vacios =[];
	// $row=2, la fila del excel donde comienza la data que se quiere guardar
	for ($row=2; $row <=$ultimaFila ; $row++) //fila
	{
		$lastColumn =$objPHPExcel->getHighestColumn();
		$lastColumn++;
		$i=0;
		for ($column = 'A'; $column < $lastColumn-1; $column++) //columna
		{
			$value =trim($objPHPExcel->getCell($column.$row)->getCalculatedValue()); //valor de la celda
			//valores vacios
			$d = $value == "" && ($column !="B") ? array_push($vacios,($column.$row)) : "";//guarda las coordenadas de la celda vacia ex: A1
		}
	}
	$contenido 	=[$vacios];
	$msg		=["No se permiten valores nulos en las siguientes celdas del archivo excel: "];
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

switch ($_REQUEST["Accion"]) {
	case 'getCuentas':
	case 'getcuentas':
		 print_r(GetCuentas());
		break;
	case 'GetDataFromExcel':
		 print_r(GetDataFromExcel());
		break;
	case 'SocioNegocio':
		print_r(GetDataFromExcel());
	break;
	case 'ListaPrecios':
		print_r(GetDataFromExcel());
	break;
	default:
		# code...
		break;
}
?>
