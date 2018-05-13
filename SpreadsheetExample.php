<?php
/**
 * Created by PhpStorm.
 * User: hannah
 * Date: 2018/5/10
 * Time: 下午3:51
 */
require_once 'vendor/phpoffice/phpspreadsheet';

/**
* @note PhpSpreadsheet is the next version of PHPExcel.
 * composer require phpoffice/phpspreadsheet
 * @param $file_name
 * @return mixed
 **/
function readExcel($file_name){
    try{
        //Create a reader by explicitly setting the file type
//        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
//        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($inputFileType);
        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($file_name);

        if ($reader->canRead($file_name)) throw new \Exception('Unable create Reader',1);
        $spreadsheet = $reader->load($file_name);
        $worksheetNames = $spreadsheet->getSheetNames();
        $result=array();
        foreach ($worksheetNames as $key =>$sheetName){
            $spreadsheet->setActiveSheetIndexByName($sheetName);
            $result[$sheetName] = $spreadsheet->getActiveSheet()->toArray(null,true,true,true);
        }
        return $result;
    }catch (\Exception $exception){
        echo $exception->getMessage();
    }
}

//readChunkExcel just the same as phpExcelExample


function createExcel(){
    $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
    $spreadsheet->setActiveSheetIndex(0)
        ->setCellValue('A1', 'Hello')
        ->setCellValue('B2', 'world!')
        ->setCellValue('C1', 'Hello')
        ->setCellValue('D2', 'world!');

    //Rename worksheet
    $spreadsheet->getActiveSheet()->setTitle('Simple');
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $spreadsheet->setActiveSheetIndex(0);

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
    $objWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet,'Excel2007');
    $objWriter->save('php://output');
}