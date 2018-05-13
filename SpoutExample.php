<?php
/**
 * Created by PhpStorm.
 * User: hannah
 * Date: 2018/5/8
 * Time: 下午3:49
 */
require_once 'vendor/autoload.php';
use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Common\Type;

/**
 * @note Spout is a PHP library to read and write CSV and XLSX files, in a fast and scalable way
 * @param $file_name
 * @return mixed
**/

function readExcel($file_name){
    try{
        $reader = ReaderFactory::create(Type::XLSX);
        //如果注释掉，单元格内的日期类型将会是DateTime，不注释的话Spout自动帮你将日期转化成string
        //$reader->setShouldFormatDates(true);
        $reader->open($file_name);

        $result = [];
        //getData from sheet1 rewind() current() next() key()
        foreach ($reader->getSheetIterator() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                $result[] =$row;
            }
        }
//        while ($reader->hasNextSheet()) {
//            $reader->nextSheet();
//            while ($reader->hasNextRow()) {
//                $result[] = $reader->nextRow();
//            }
//        }
        $reader->close();
        return $result;
    }catch (\Exception $e){
        echo $e->getMessage();
    }
}


function writeExcel(){
    $writer = WriterFactory::create(Type::XLSX);

    $writer->setTempFolder('')
        ->setCurrentSheet($writer->getCurrentSheet())
        ->setShouldUseInlineStrings(true)  // default (and recommended) value
        ->setShouldCreateNewSheetsAutomatically(true); // default value

    $fileName='test.xlsx';
    $data=[[1,2,3],[4,5,6]];
    $writer->openToFile($fileName);
//    $writer->openToBrowser($fileName);
    $writer->addRows($data);
    $writer->close();
}
