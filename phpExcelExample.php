<?php
/**
 * Created by PhpStorm.
 * User: hannah
 * Date: 2018/5/10
 * Time: 下午3:59
 */
require_once 'vendor/autoload.php';

/**
 * @note Read Create and Write PHPExcel OpenXML in php
 * composer require phpoffice/phpexcel
 * This package is abandoned and no longer maintained
 * @param $file_name
 * @return mixed
**/
function readEntireExcel($file_name){
    try{
        /** Create a reader by explicitly setting the file type.
        // $inputFileType = 'Excel5';
        // $inputFileType = 'Excel2007';
        // $inputFileType = 'Excel2003XML';
        // $inputFileType = 'OOCalc';
        // $inputFileType = 'SYLK';
        // $inputFileType = 'Gnumeric';
        // $inputFileType = 'CSV';
        $excelReader = PHPExcel_IOFactory::createReader($inputFileType);
         */

        //automatically detect the correct reader to load for this file
        $excelReader = PHPExcel_IOFactory::createReaderForFile($file_name);
        if (!$excelReader->canRead($file_name)) throw new \Exception('Unable read file',1);

        $excelObj = $excelReader->load($file_name);
        //only display the information on the currently active sheet, which is the last loaded one
        //default entire excel file loaded at once,it will be much faster to split it in chunks and work with each individual chunk at a time
        $worksheetNames = $excelObj->getSheetNames();
        $result=array();
        foreach ($worksheetNames as $key =>$sheetName){
            $excelObj->setActiveSheetIndexByName($sheetName);
            $result[$sheetName] = $excelObj->getActiveSheet()->toArray(null,true,true,true);
        }
        return $result;
    }catch (\Exception $e){
        echo $e->getMessage();
    }
}


function readChunkExcel($file_name){
    try{
        $excelReader = PHPExcel_IOFactory::createReaderForFile($file_name);

        $result = [];
        $chunkSize = 2048;
        $chunkFilter = new ChunkReadFilter();
        /** Tell the Reader that we want to use the Read Filter **/
        $excelReader->setReadFilter($chunkFilter);
        /** Loop to read our worksheet in "chunk size" blocks **/
        for ($startRow = 2; $startRow <= 65536; $startRow += $chunkSize){
            $chunkFilter->setRows($startRow,$chunkSize);
            /** Load only the rows that match our filter **/
            $excelObj = $excelReader->load($file_name);
            $result[] = $excelObj->getActiveSheet()->toArray(null, true,true,true);
            // Do some processing here - the $data variable will contain an array which is always limited to 2048 elements regardless of the size of the entire sheet
        }
        return $result;
    }catch (\Exception $e){
        echo $e->getMessage();
    }
}


//Loading data in chunks
class ChunkReadFilter implements PHPExcel_Reader_IReadFilter{
    private $_startRow =0;
    private $_endRow = 0;

    /**
     * Set the list of rows that we want to read
     * @param $startRow
     * @param $chunkSize
     **/
    public function setRows($startRow,$chunkSize){
        $this->_startRow = $startRow;
        $this->_endRow = $startRow + $chunkSize;
    }

    public function readCell($column, $row, $worksheetName = '')
    {
        // Only read the heading row, and the configured rows
        if ($row ==1 || ($row >= $this->_startRow && $row < $this->_endRow))
            return true;
        else return false;
    }
}


function writeExcel(){
    $objExcel = new PHPExcel();
    //// Set document properties
//$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
//    ->setLastModifiedBy("Maarten Balliauw")
//    ->setTitle("Office 2007 XLSX Test Document")
//    ->setSubject("Office 2007 XLSX Test Document")
//    ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
//    ->setKeywords("office 2007 openxml php")
//    ->setCategory("Test result file");

    //Add some data
    $objExcel->setActiveSheetIndex(0)
        ->setCellValue('A1', 'Hello')
        ->setCellValue('B2', 'world!')
        ->setCellValue('C1', 'Hello')
        ->setCellValue('D2', 'world!');
    //Rename worksheet
    $objExcel->getActiveSheet()->setTitle('Simple');
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objExcel->setActiveSheetIndex(0);

    //Redirect output a client's web browser
    //header('Content-Type: application/octet-stream'); // unknown mime-type
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="01simple.xlsx"');
    header('Cache-Control: max-age=0');

    //If you're serving to IE9, then the following may be needed
    header('Cache-Control: max-age=1');
    // If you're serving to IE over SSL, then the following may be needed
    header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
    header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
    header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
    $objWriter = PHPExcel_IOFactory::createWriter($objExcel,'Excel2007');
    $objWriter->save('php://output');
}
