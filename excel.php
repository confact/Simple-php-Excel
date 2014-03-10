<?php
class Excel
{
	
	function __construct($file, $sheet = 0) {
		date_default_timezone_set('Europe/Stockholm');
		require_once 'PHPExcel/Classes/PHPExcel/IOFactory.php';
		require_once 'PHPExcel/Classes/PHPExcel.php';
		$this->filename = $file;
		try {
		    $inputFileType = PHPExcel_IOFactory::identify($file);
		    $this->filetype = $inputFileType;
		    $objReader = PHPExcel_IOFactory::createReader($inputFileType);
		    $this->excel = $objReader->load($file);
		    $this->worksheet = $this->excel->setActiveSheetIndex($sheet) ;
		} catch(Exception $e) {
		    die('Error loading file "'.pathinfo($file,PATHINFO_BASENAME).'": '.$e->getMessage());
		}
		$this->highestRow = $this->worksheet->getHighestRow(); 
		$this->highestColumn = $this->worksheet->getHighestColumn();
	}
	
	function getRowsForColumn($column, $skipheadcolumn = true) {
		$array = array();
		$lastRow = $this->highestRow;
		if($skipheadcolumn) {
			for ($row = 2; $row <= $lastRow; $row++) {
			    $cell = $this->worksheet->getCell($column.$row);
			    $array[$row] = $cell->getCalculatedValue();
			}
		}
		else {
			for ($row = 1; $row <= $lastRow; $row++) {
			    $cell = $this->worksheet->getCell($column.$row);
			    $array[$row] = $cell->getCalculatedValue();
			}
		}
		return $array;
	}
	
	function set($column, $row, $value) {
		echo "Setting Columns in excel \n";
		$this->worksheet->setCellValueExplicit(
    $column.$row, 
    $value, 
    PHPExcel_Cell_DataType::TYPE_STRING
);
	}
	
	function getspecificrowcolumn($column, $row) {
		$cell = $this->worksheet->getCell($column.$row);
		return $cell->getCalculatedValue();
	}
	
	function getRows($skipheadcolumn = true) {		
		$rowIterator = $this->excel->getActiveSheet()->getRowIterator();
		$array_data = array();
		foreach($rowIterator as $row){
		    $cellIterator = $row->getCellIterator();
		    $cellIterator->setIterateOnlyExistingCells(true); // Loop only existing cells
		    if($skipheadcolumn && 1 == $row->getRowIndex ()) continue;//skip first row
		    $rowIndex = $row->getRowIndex ();
		    $array_data[$rowIndex] = array();
		     
		    foreach ($cellIterator as $cell) {
				$array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
		    }
		}
		return $array_data;
	}
	
	function save() {
		$objWriter = PHPExcel_IOFactory::createWriter($this->excel, $this->filetype);
		$objWriter->save($this->filename);
	}
	
}
?>