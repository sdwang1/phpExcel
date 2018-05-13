<?php
/**
 * Created by PhpStorm.
 * User: hannah
 * Date: 2018/5/13
 * Time: 下午3:19
 */
require_once  'vendor/autoload.php';

/**
 * @note A lightly Excel XLSx files reader 读取兼容效果欠佳【尤其是维度值为空】
 * composer require jakeroid/simplexlsx
 * @param $file_name
 * @return mixed
**/
function readExcel($file_name){
    try{
        $xlsx = new \jakeroid\tools\SimpleXLSX($file_name);

//        $sheetNum = $xlsx->sheetsCount();
        // output worsheet 1
        $sheetIdx = 1;
        list($num_cols, $num_rows) = $xlsx->dimension($sheetIdx);

        echo '<table cellpadding="10">
                <tr><td valign="top">';
        echo '<h1>Sheet 1</h1>';
        echo '<table>';
        foreach( $xlsx->rows($sheetIdx) as $r ) {
            echo '<tr>';
            for( $i=0; $i < $num_cols; $i++ )
                echo '<td>'.( (!empty($r[$i])) ? $r[$i] : '&nbsp;' ).'</td>';
            echo '</tr>';
        }
        echo '</table>';
        echo '</td><td valign="top">';
    }catch (\Exception $e){
        echo $e->getMessage();
    }
}




