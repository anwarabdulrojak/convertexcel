<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Helper\Sample;

$helper = new Sample();

if(isset($_POST['Submit'])){

    $mimes = ['application/vnd.ms-excel','text/xls','text/xlsx','application/vnd.oasis.opendocument.spreadsheet','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];

    if(in_array($_FILES["file"]["type"],$mimes)){

        $uploadFilePath = 'uploads/'.basename($_FILES['file']['name']);
        move_uploaded_file($_FILES['file']['tmp_name'], $uploadFilePath);

        //GET FILENAME FROM UPLOAD
        $filename = pathinfo($uploadFilePath, PATHINFO_FILENAME);
        
        $spreadsheet = IOFactory::load($uploadFilePath);
        $sheetData = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);

        //GRID VIEW
        // $helper->displayGrid($sheetData);

        $new_sheet = [];
        $new_sheet_not_allow = [];

        $new_sheet_not_allow[] = $sheetData[1];

        //REMOVE HEADER/FIRST ROW
        array_shift($sheetData);

        //SET HEADER        
        $sheet = ['sku','name','url_key','quantity_dus','weight','product_type','categories','price','jade_price','diamond_price','gold_price','silver_price','ruby_price','qty','attribute_set_code','visibility','additional_attributes','configurable_variation_labels','configurable_variations','base_image','small_image','thumbnail_image','additional_images','product_websites','product_online','website_id','page_layout'];
        $new_sheet[] = $sheet;

        $oldsku = '';
        $oldname = '';
        $oldproduct = [];
        $configurable_variations_list = [];
        
        $defineFirstSKU = ['JH-','JHL-','HL-','LC-','BS-','TSB-','JHX-','FG-','JM-','HM-','JT JHW ','JHW ','JH ','HJ-','GQ-','LC ','JT ','HL ','BS '];
        $defineAllowedCat = ['BAG','WATCH','WALLET','PACKAGING','HEELS','WOMAN','HOMEWARES','POWERBANK','TUMBLER'];

        foreach ($sheetData as $key => $value) {
            if (($value['A'] == 'NORAH HL-852 (1Q :30/26pc)') || ($value['A'] == 'CALLY BAG')) {
                continue;
            }
            if ($value['D'] == '') {
                $value['D'] = 'NO-VARIANT';
            }
            if ($value['A'] == 'EMMA JHX-202 GOLD-DIAMOND (1Q:12BOX)'){
                if ($value['I'] == '' || $value['I'] == null) {
                    $value['I'] = 1656000;
                }
            }
            if ($value['A'] == 'AUDREY JHX-102 GOLD-DIAMOND (1Q:12BOX)'){
                if ($value['I'] == '' || $value['I'] == null) {
                    $value['I'] = 1680000;
                }
            }
            if ($value['A'] == 'MAUDY JHX-205 JADE-SILVER (1Q:24 BOX)'){
                if ($value['K'] == '' || $value['K'] == null) {
                    $value['K'] = 3792000;
                }
            }
            if ($value['A'] == 'JHW 32 (1Q:60)'){
                if ($value['K'] == '' || $value['K'] == null) {
                    $value['K'] = 135000;
                }
            }
            if ($value['F'] == '' || $value['F'] == null) {
                $value['F'] = $value['G'];
            }
            if ($value['F'] == '' || $value['G'] == '' || $value['H'] == '' || $value['I'] == '' || $value['K'] == '') {
                $new_sheet_not_allow[] = $value;
                continue;
            }
            
            if (array_search($value['B'], $defineAllowedCat) === false) {
                $new_sheet_not_allow[] = $value;
                continue;
            }

            $sku = $value['A'];
            if ($value['B'] != 'PACKAGING' && $value['B'] != 'HOMEWARES' && $value['B'] != 'POWERBANK' && $value['B'] != 'TUMBLER') {
                foreach ($defineFirstSKU as $val) {
                    if (strpos($value['A'], $val) !== FALSE) {
                        $sku = substr($value['A'], strpos($value['A'], $val));
                        if($val == 'JHW ' || $val == 'JT ' || $val == 'JT JHW '){
                            $sku;
                        }
                        else if(($val == 'JH ' || $val == 'LC ' || $val == 'HL ' || $val == 'HL-' || $val == 'BS ') && str_contains($sku,'(1')) 
                        {
                            $sku = substr($sku, 0, strpos($sku, '(1'));
                        }
                        else
                        {
                            if(!empty(strpos($sku, ' '))) $sku = substr($sku, 0, strpos($sku, ' '));
                        }
                        break;
                    }
                }
            } 

            $name = str_replace([$sku], '', $value['A']);
            if (str_contains($sku,'(1') || str_contains($sku,'( 1')) $sku = substr($sku, 0, strpos($sku, '('));

            if ($name == '') {
                $name = $value['A'];
                if ($value['B'] != 'PACKAGING' && $value['B'] != 'HOMEWARES' && $value['B'] != 'POWERBANK' && $value['B'] != 'TUMBLER') {
                    if (!str_contains($sku,'JH') && !str_contains($sku,'JT')) {
                        if (preg_match('/\\d/', $sku)){
                            $sku = intval(preg_replace('/[^0-9]+/', '', $sku), 10);
                            if ($name != $sku) $name = str_replace([$sku], '', $name);
                        }
                    }
                }
            }
            if ($sku == 'ADELLE BAG 1071') {
                $sku = '1071';
                $name = str_replace([$sku], '', $name);
            }

            $name = preg_replace('/\s+/', ' ', $name);
            if (preg_match('/^\s+|\s+$/u',$name)) $name = preg_replace('/^\s+|\s+$/u', '', $name);
            if (preg_match('/^\s+|\s+$/u',$sku)) $sku = preg_replace('/^\s+|\s+$/u', '', $sku);

            if (str_contains($name,'JADE-GOLD')) {
                $sku = $sku.'-JADE-GOLD';
            }
            if (str_contains($name,'SILVER-RUBY')) {
                $sku = $sku.'-SILVER-RUBY';
            }
            if (str_contains($name,'GOLD-DIAMOND')) {
                $sku = $sku.'-GOLD-DIAMOND';
            }
            if (str_contains($name,'RUBY-SILVER')) {
                $sku = $sku.'-RUBY-SILVER';
            }
            if (str_contains($name,'BRONZE-SILVER')) {
                $sku = $sku.'-BRONZE-SILVER';
            }
            if ($name == 'SPARKLE BAG TH(1Q:32)') {
                $sku = '9TH';
                $name = 'SPARKLE BAG (1Q:32)';
            }
            if ($name == 'MARU (1Q : 48pc) TANPA TALI') {
                $sku = $sku.'-TANPA TALI';
            }

            if (str_contains($value['D'],"'")) $value['D'] = str_replace(["'"], '', $value['D']);
            $value['D'] = strtoupper($value['D']);

            $skucolor = $sku.' '.$value['D'];
            $image = str_replace(['-', ' ','.','(',')'], '',$skucolor).'.jpg';

            $duplicateVariant = array_search($skucolor, array_column($new_sheet, '0'));
            if($duplicateVariant !== false) {
                //replace price dan stock
                $new_sheet[$duplicateVariant][7] = $value['K'];
                $new_sheet[$duplicateVariant][8] = $value['F'];
                $new_sheet[$duplicateVariant][9] = $value['G'];
                $new_sheet[$duplicateVariant][10] = $value['H'];
                $new_sheet[$duplicateVariant][11] = $value['I'];
                $new_sheet[$duplicateVariant][12] = $value['K'];
                $new_sheet[$duplicateVariant][13] = $value['N'];

                continue;
            }

            $quantity_dus = '';
            if (str_contains($name,'1Q') || str_contains($name,'1 dus') || str_contains($name,'1dus') || str_contains($name,'1DUS') || str_contains($name,'1 Q') || str_contains($name,'1 DUS')) {
                $quantity_dus = str_replace([' ','pc','PC','Pc','box','BOX','roll'], '', $name);
                if(str_contains($quantity_dus,'1dus')) $quantity_dus = str_replace('dus', 'Q', $quantity_dus);
                if(str_contains($quantity_dus,'1DUS')) $quantity_dus = str_replace('DUS', 'Q', $quantity_dus);
                if(str_contains($quantity_dus,'Q)')) $quantity_dus = str_replace('Q)', ')', $quantity_dus);
                if(strpos($quantity_dus, '1Q:') === false) $quantity_dus = str_replace('1Q', '1Q:', $quantity_dus);

                $subtring_start = strpos($quantity_dus, '(1Q:');
                $subtring_start += strlen('(1Q:');
                $size = strpos($quantity_dus, ')', $subtring_start) - $subtring_start;
                $quantity_dus = substr($quantity_dus, $subtring_start, $size);
            }

            if ($sku == 'JH-518') {
                $quantity_dus = str_replace([' ','pc','PC','Pc','box','BOX','roll'], '', $name);
                $subtring_start = strpos($quantity_dus, '(');
                $subtring_start += strlen('(');
                $size = strpos($quantity_dus, ')', $subtring_start) - $subtring_start;
                $quantity_dus = substr($quantity_dus, $subtring_start, $size);
            }

            if ($sku == 'CARA BAG') $quantity_dus = '36';
            if ($sku == 'JH-813') $quantity_dus = '40';

            $category = '';
            if ($value['B'] == 'BAG') {
                $category = 'Default Category,Default Category/SHOP WOMEN,Default Category/SHOP WOMEN/BAG';
                if ($quantity_dus == 0 || $quantity_dus == null) $quantity_dus = 35;
            }
            else if ($value['B'] == 'WALLET') 
            {
                $category = 'Default Category,Default Category/SHOP WOMEN,Default Category/SHOP WOMEN/WALLET';
                if ($quantity_dus == 0 || $quantity_dus == null) $quantity_dus = 100;
            }
            else if ($value['B'] == 'HEELS')
            {
                $category = 'Default Category,Default Category/SHOP WOMEN,Default Category/SHOP WOMEN/HEELS';
                if ($quantity_dus == 0 || $quantity_dus == null) $quantity_dus = 12;
            }
            else if ($value['B'] == 'WOMEN')
            {
                $category = 'Default Category,Default Category/SHOP WOMEN,Default Category/SHOP WOMEN/BAG';
                if ($quantity_dus == 0 || $quantity_dus == null) $quantity_dus = 35;
            }
            else if ($value['B'] == 'WATCH')
            {
                $category = 'Default Category,Default Category/WATCHES';
                if ($quantity_dus == 0 || $quantity_dus == null) $quantity_dus = 100;
            }
            else
            {
                if ($quantity_dus == 0 || $quantity_dus == null) $quantity_dus = 100;
                $category = 'Default Category,Default Category/'.$value['B'];
            }

            if ($value['D'] == 'HITAM') $value['D'] = 'BLACK';
            if ($value['D'] == 'BIRUMUDA') $value['D'] = 'SKYBLUE';
            if ($value['D'] == 'BIRUTUA') $value['D'] = 'BLUE';
            if ($value['D'] == 'MERAH') $value['D'] = 'RED';
            if ($value['D'] == 'ABU') $value['D'] = 'GREY';
            if ($value['D'] == 'MERAHMUDA') $value['D'] = 'PINK';
            if ($value['D'] == 'PUTIH') $value['D'] = 'WHITE';
            if ($value['D'] == 'COKLATGELAP') $value['D'] = 'DARKBROWN';
            if ($value['D'] == 'COKLAT') $value['D'] = 'BROWN';
            if ($value['D'] == 'KARAMEL') $value['D'] = 'CARAMEL';
            if ($value['D'] == 'GREEN-ARMY') $value['D'] = 'GREENARMY';
            if ($value['D'] == 'DARK-GREEN') $value['D'] = 'DARKGREEN';
            if ($value['D'] == 'WHITE-APRCT') $value['D'] = 'APRICOT-WHITE';
            if ($value['D'] == 'LO2') $value['D'] = 'L02';
            if ($value['D'] == 'LO3') $value['D'] = 'L03';

            //CONFIGURABLE
            if (($oldsku != '' && $oldsku != $sku)) {
                $oldlink = $oldproduct[2];
                $newlink = $oldproduct[1].'-'.$oldsku;
                $oldlink = preg_replace('/\s+/', '', $oldlink);
                
                if ($oldlink == $newlink) {
                    $newlink = $oldproduct[1].'-'.$oldsku.'-CB';
                }
                $sheet = [];
                $oldproduct[0] = $oldsku;
                $oldproduct[2] = $newlink;
                $oldproduct[5] = 'configurable';
                $oldproduct[13] = '';
                $oldproduct[15] = 'Catalog, Search';
                if (($oldproduct[6] == 'Default Category,Default Category/HOMEWARES' || $oldproduct[6] == 'Default Category,Default Category/PACKAGING')  && $oldproduct[1] != 'STANDING A4 ACRILIC DISPLAY (DENGAN LED)') {
                    $oldproduct[17] = 'size=Size';
                } 
                else 
                {
                    $oldproduct[17] = 'color=Color';
                    if (str_contains($oldproduct[16],'COLOR') && str_contains($oldproduct[16],'SIZE')) $oldproduct[17] = 'color=Color,size=Size';
                }
                $oldproduct[16] = '';
                $oldproduct[18] = implode('|',$configurable_variations_list);
                $oldproduct[26] = 'Product -- Full Width';
                $new_sheet[] = $oldproduct;

                //RESET DATA
                $oldsku = '';
                $oldproduct = [];
                $configurable_variations_list = [];
            }

            $color = '';
            $size = '';
            if ($value['B'] == 'HEELS' && str_contains($value['D'],'-')) {

                $colorsize = $value['D'];
                $subtring_start = strpos($colorsize, '-');
                $subtring_start += strlen('-');

                $size = substr($colorsize, $subtring_start, 2);
                $color = substr($colorsize, 0, strpos($colorsize, '-'));
            }
            
            $sheet = [];
            $sheet[] = $skucolor;
            $sheet[] = $name;
            $sheet[] = $name.'-'.$skucolor;
            $sheet[] = $quantity_dus;
            $sheet[] = '500';
            $sheet[] = 'simple';
            $sheet[] = $category;
            $sheet[] = $value['K'];
            $sheet[] = $value['F'];
            $sheet[] = $value['G'];
            $sheet[] = $value['H'];
            $sheet[] = $value['I'];
            $sheet[] = $value['K'];
            $sheet[] = $value['N'];
            $sheet[] = 'Default';
            $sheet[] = 'Not Visible Individually';
            if (($value['B'] == 'PACKAGING' || $value['B'] == 'HOMEWARES') && $name != 'STANDING A4 ACRILIC DISPLAY (DENGAN LED)'){
                $sheet[] = 'SIZE='.$value['D'];
            }
            else 
            {
                if ($value['B'] == 'HEELS' && str_contains($value['D'],'-')) {
                    $sheet[] = 'COLOR='.$color.',SIZE='.$size;
                }
                else 
                {
                    $sheet[] = 'COLOR='.$value['D'];
                }
            }
            $sheet[] = '';
            $sheet[] = '';
            $sheet[] = $image;
            $sheet[] = $image;
            $sheet[] = $image;
            $sheet[] = $image;
            $sheet[] = 'base';
            $sheet[] = '1';
            $sheet[] = '1';
            $sheet[] = '';
            $new_sheet[] = $sheet;

            //FOR CONFIGURABLE
            $oldname = $name;
            $oldsku = $sku;
            $oldproduct = $sheet;
            $configurable_variations = '';
            if (($value['B'] == 'PACKAGING' || $value['B'] == 'HOMEWARES') && $name != 'STANDING A4 ACRILIC DISPLAY (DENGAN LED)'){
                $configurable_variations = 'SKU='.$skucolor.',SIZE='.$value['D'];
            }
            else 
            {
                if ($value['B'] == 'HEELS' && str_contains($value['D'],'-')) {
                    $configurable_variations = 'SKU='.$skucolor.',COLOR='.$color.',SIZE='.$size;
                }
                else 
                {
                    $configurable_variations = 'SKU='.$skucolor.','.'COLOR='.$value['D'];
                }
            }
            $configurable_variations_list[] = $configurable_variations;
        }

        //LAST CONFIGURABLE
        if (!empty($oldproduct)) {
            $oldlink = $oldproduct[2];
            $newlink = $oldproduct[1].'-'.$oldsku;
            if ($oldlink == $newlink) {
                $newlink = $oldproduct[1].'-'.$oldsku.'-CB';
            }
            $sheet = [];
            $oldproduct[0] = $oldsku;
            $oldproduct[2] = $newlink;
            $oldproduct[5] = 'configurable';
            $oldproduct[13] = '';
            $oldproduct[15] = 'Catalog, Search';
            if (($oldproduct[6] == 'Default Category,Default Category/HOMEWARES' || $oldproduct[6] == 'Default Category,Default Category/PACKAGING')  && $oldproduct[1] != 'STANDING A4 ACRILIC DISPLAY (DENGAN LED)') {
                $oldproduct[17] = 'size=Size';
            }
            else 
            {
                $oldproduct[17] = 'color=Color';
                if (str_contains($oldproduct[16],'COLOR') && str_contains($oldproduct[16],'SIZE')) $oldproduct[17] = 'color=Color,size=Size';
            }
            $oldproduct[16] = '';
            $oldproduct[18] = implode('|',$configurable_variations_list);
            $oldproduct[26] = 'Product -- Full Width';
            $new_sheet[] = $oldproduct;
        }

        $new_spreadsheet = new Spreadsheet();

        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $new_spreadsheet->setActiveSheetIndex(0);

        $new_spreadsheet->getActiveSheet()->fromArray($new_sheet, null, 'A1');

        // //GRID VIEW
        // $newsheetData = $new_spreadsheet->getActiveSheet()->toArray(null, true, true, true);
        // $helper->displayGrid($newsheetData);

        //AUTO DOWNLOAD FILE EXCEL
        // Redirect output to a client’s web browser (CSV)
        header('Content-Type: text/csv');
        header('Content-Disposition: attachment;filename="'.$filename.'-convert.csv"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');
        
        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $writer = new Csv($new_spreadsheet);
        // $writer = IOFactory::createWriter($new_spreadsheet, 'CSV');
        $writer->save('php://output');

        // echo "<h1>Data Yang tidak masuk ke dalam excel</h1>";
        // //SET TO GRID VIEW NOT ALLOWED DATA
        // $new_spreadsheet_not_allow = new Spreadsheet();
        // // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        // $new_spreadsheet_not_allow->setActiveSheetIndex(0);
        // $new_spreadsheet_not_allow->getActiveSheet()->fromArray($new_sheet_not_allow, null, 'A1');
        // //GRID VIEW
        // $newsheetDataNotAllow = $new_spreadsheet_not_allow->getActiveSheet()->toArray(null, true, true, true);
        // $helper->displayGrid($newsheetDataNotAllow);

    }
    else 
    { 
        die("<br/>Sorry, File type is not allowed. Only Excel file."); 
    }
}
?>