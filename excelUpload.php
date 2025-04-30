<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Helper\Sample;

$helper = new Sample();
$defineLocationSKU = [
    ['sku' => 'JH-336','location' => '2.C8'],
    ['sku' => 'JH-357','location' => '4.B14'],
    ['sku' => 'JH-632','location' => '4.C11'],
    ['sku' => 'HL-863','location' => 'M.WIP'],
    ['sku' => 'JH-815','location' => '4.C3'],
    ['sku' => 'JH-368','location' => 'Depan Jendela Lt.4'],
    ['sku' => 'LC-7110','location' => '2.B8'],
    ['sku' => 'JH-625','location' => '4.C7'],
    ['sku' => 'JH-907','location' => '4.B12'],
    ['sku' => 'LC-7091','location' => '2.A1'],
    ['sku' => 'LC-7115','location' => 'Depan Lift Lt.2'],
    ['sku' => 'JH-715','location' => 'Depan Toilet Lt.1'],
    ['sku' => 'JH-1108','location' => '4.B16'],
    ['sku' => 'JHX-103','location' => '4.B5'],
    ['sku' => 'JH-517','location' => 'Depan Lift Lt.4'],
    ['sku' => 'JH-722','location' => '4.B3'],
    ['sku' => 'BS-062','location' => 'Depan Lift Lt.2'],
    ['sku' => 'JH-335','location' => '4.B6'],
    ['sku' => 'JH-333','location' => '4.B6'],
    ['sku' => 'JH-1015','location' => '4.B3'],
    ['sku' => 'JH-1027','location' => '3.B18'],
    ['sku' => 'LC-7101','location' => '2.A6'],
    ['sku' => 'JH-918','location' => 'M.C11'],
    ['sku' => 'JH-2009','location' => 'M.C4'],
    ['sku' => 'JH-523','location' => 'M.C3'],
    ['sku' => 'JH-818','location' => '4.B'],
    ['sku' => 'BOX JAM TANGAN EKSKLUSIF 8125 l05','location' => '1.B13'],
    ['sku' => 'JH-922','location' => 'Depan Jendela Lt.3'],
    ['sku' => 'BUBBLE MAILER','location' => '4.B20 - B22'],
    ['sku' => 'LC-7118','location' => 'Jendela Lt.2'],
    ['sku' => 'LC-7119','location' => 'Jendela Lt.2'],
    ['sku' => 'JH-622','location' => '4.B9'],
    ['sku' => 'JH-350','location' => '4.B12'],
    ['sku' => 'JHX-201-GOLD-DIAMOND','location' => '4.B10'],
    ['sku' => 'JHX-201-RUBY-SILVER','location' => '4.B10'],
    ['sku' => 'JH-1028','location' => '4.B15'],
    ['sku' => 'JH-1016','location' => '4.A1 -A2'],
    ['sku' => 'JH-326','location' => '4.B13'],
    ['sku' => 'JH-520','location' => 'M.B7'],
    ['sku' => 'JHX-101-JADE-GOLD','location' => 'M.C2'],
    ['sku' => 'JHX-101-SILVER-RUBY','location' => 'M.C2'],
    ['sku' => 'JH-361','location' => '2.C17 - C18'],
    ['sku' => 'JH-711','location' => '4.C14'],
    ['sku' => 'JH-525','location' => '4.B14'],
    ['sku' => 'HL-862','location' => 'M.C3'],
    ['sku' => 'JH-1017','location' => '4.B1'],
    ['sku' => 'JH-706','location' => '4.B11'],
    ['sku' => 'JH-811','location' => '4.B4'],
    ['sku' => 'JH-817','location' => 'M.WIP'],
    ['sku' => 'JH-1023','location' => '2.C13'],
    ['sku' => 'JH-1021','location' => '4.C4'],
    ['sku' => 'ELLIS CARDHOLDER','location' => '1.B15'],
    ['sku' => 'BS-066','location' => '2.C2'],
    ['sku' => 'LC-7086','location' => '2.B21'],
    ['sku' => 'JHX-202-BRONZE-SILVER','location' => '4.B6'],
    ['sku' => 'JHX-202-GOLD-DIAMOND','location' => '4.B5'],
    ['sku' => 'LC-7087','location' => 'M.C14 & B8'],
    ['sku' => 'JH-203','location' => '4.C8 - C9'],
    ['sku' => 'LC-7030','location' => '1.B15'],
    ['sku' => 'JH-359','location' => '4.A5'],
    ['sku' => 'GOODIE BAG HOLOGRAM MOE','location' => '2'],
    ['sku' => 'JH-913','location' => '4.B2'],
    ['sku' => 'JH-369','location' => 'Depan Lift Lt.4'],
    ['sku' => 'JH-512','location' => '1.B12'],
    ['sku' => 'JH-2013','location' => '1.B8'],
    ['sku' => 'LC-7098','location' => 'Jendela Lt.2'],
    ['sku' => 'JH-329','location' => '4.C15'],
    ['sku' => 'JH-701','location' => '4.B9'],
    ['sku' => 'LC-7105','location' => '2.C3 - C4'],
    ['sku' => 'JH-360','location' => '4.C6'],
    ['sku' => 'JH-516','location' => 'M.C7'],
    ['sku' => 'JHW 18','location' => '1.B11'],
    ['sku' => 'JHW 23','location' => '1.B15'],
    ['sku' => 'JHW 26','location' => '1.B11'],
    ['sku' => 'JHW 27','location' => '1.B14'],
    ['sku' => 'JHW 30','location' => '1.B15'],
    ['sku' => 'JHW 31','location' => '1.B13'],
    ['sku' => 'JHW 32','location' => '1.B15'],
    ['sku' => 'JHW 33','location' => '1.B10'],
    ['sku' => 'JHW 38','location' => '1.B11'],
    ['sku' => 'JHW 39','location' => '1.B11'],
    ['sku' => 'JHW 50','location' => '1.B14'],
    ['sku' => 'JHW 52','location' => 'Tangga Office Lt.2'],
    ['sku' => 'JHW 53','location' => 'Office Lt 2'],
    ['sku' => 'JH-1109','location' => 'Depan Jendela Lt.4'],
    ['sku' => 'JH-712','location' => 'M.WIP'],
    ['sku' => 'LC-7095','location' => '2.B19'],
    ['sku' => 'TSB-003','location' => 'M.C20'],
    ['sku' => 'JT 2139','location' => '1.B15'],
    ['sku' => 'JT 8011','location' => '1.B13'],
    ['sku' => 'JT 8027','location' => '1.B11'],
    ['sku' => 'JT 8062','location' => '1.B12'],
    ['sku' => 'JT 8086','location' => '1.B13'],
    ['sku' => 'JT 8123','location' => '1.B10'],
    ['sku' => 'JT 8125 PLUS','location' => '1.B10'],
    ['sku' => 'JT 8138','location' => '1.B11'],
    ['sku' => 'JT 8151','location' => '1.B11'],
    ['sku' => 'JH-809','location' => '4.B14'],
    ['sku' => 'LC-7107','location' => '2.B13 & B12'],
    ['sku' => 'LC-7111','location' => 'Jendela Lt.2'],
    ['sku' => 'LC-7071','location' => '2.B9'],
    ['sku' => 'JHX-209','location' => '4.B17'],
    ['sku' => 'JH-613','location' => '4.B13'],
    ['sku' => 'JH-703','location' => '4.B14'],
    ['sku' => 'JH-917','location' => '4.B3'],
    ['sku' => 'JH-920','location' => 'Depan Jendela Lt.4'],
    ['sku' => 'JH-713','location' => '4.B14'],
    ['sku' => 'JH-921','location' => 'Depan Jendela Lt.3'],
    ['sku' => 'LAKBAN BENING','location' => 'Lift Lt.2'],
    ['sku' => 'LC-7106','location' => '2.B4'],
    ['sku' => 'LC-7083','location' => '2.B3'],
    ['sku' => 'JH-1025','location' => '4.A5 - A6'],
    ['sku' => 'JH-2012','location' => 'M.C17'],
    ['sku' => 'JH-2011','location' => '1.B8'],
    ['sku' => 'JH-337','location' => '4.B18'],
    ['sku' => 'JH-1022','location' => '2.C13'],
    ['sku' => 'JH-2008','location' => 'M.C5'],
    ['sku' => 'JH-355','location' => '3.B20'],
    ['sku' => 'LC-7113','location' => '2.B15'],
    ['sku' => 'LC-7075','location' => '2.B15'],
    ['sku' => 'JH-1018','location' => 'M.B1 - B3'],
    ['sku' => 'JH-519','location' => 'M.C4'],
    ['sku' => 'JH-338','location' => '2.B22'],
    ['sku' => 'LC-7088','location' => '2.B5'],
    ['sku' => 'BS-069','location' => '2.B17'],
    ['sku' => 'JH-912','location' => '3.C16'],
    ['sku' => 'JH-631','location' => 'Depan Lift Lt.4'],
    ['sku' => 'JH-721','location' => 'M.B7 - B8'],
    ['sku' => 'JH-363','location' => '4.B18'],
    ['sku' => 'JH-1107','location' => '1.B15'],
    ['sku' => 'TSB-006','location' => 'M.B12 & B13'],
    ['sku' => 'JH-630','location' => '4.B2'],
    ['sku' => 'LC-7109','location' => '2.B6'],
    ['sku' => 'LC-7092','location' => '2. depan lift'],
    ['sku' => 'JH-339','location' => '4.B16'],
    ['sku' => 'JH-926','location' => '4.B4'],
    ['sku' => 'JH-925','location' => 'M.B5 - B6'],
    ['sku' => 'JH-358','location' => '4. samping lift'],
    ['sku' => 'JH-513','location' => 'Depan Jendela Lt.4'],
    ['sku' => 'JH-628','location' => '4.C17'],
    ['sku' => 'TSB-007','location' => 'M.C15 - C16'],
    ['sku' => 'JH-910','location' => '1.WIP'],
    ['sku' => 'JH-915','location' => '4.A3 & B1'],
    ['sku' => 'PLASTIK JH PINK','location' => 'Tangga Office Lt.2'],
    ['sku' => 'PLASTIK JIMS HONEY','location' => 'Tangga Office Lt.2'],
    ['sku' => 'PLASTIK POLYMAILER BESAR','location' => '2.B14'],
    ['sku' => 'PLASTIK POLYMAILER KECIL','location' => '2.B14'],
    ['sku' => 'LC-7077','location' => '1.B10'],
    ['sku' => 'JH-2007','location' => 'M.C6'],
    ['sku' => 'TSB-005','location' => 'M.WIP'],
    ['sku' => 'JH-708','location' => 'Depan Jendela Lt.3'],
    ['sku' => 'JH-332','location' => '2.C5 - C7'],
    ['sku' => 'JH-366','location' => 'M.B14'],
    ['sku' => 'JH-716','location' => '1.Wip'],
    ['sku' => 'JHX-208-JADE-GOLD','location' => '4.A4'],
    ['sku' => 'JHX-208-SILVER-RUBY','location' => '4.A4'],
    ['sku' => 'TSB-002','location' => 'M.B14'],
    ['sku' => 'LC-7090','location' => '2.B2 - B.3'],
    ['sku' => 'JH-813','location' => '2.C15 - C16'],
    ['sku' => 'STANDING A4 ACRILIC DISPLAY (DENGAN LED)','location' => '2.B1'],
    ['sku' => 'JH-356','location' => '1.Wip'],
    ['sku' => 'JH-1012','location' => 'Depan Jendela Lt.4'],
    ['sku' => 'JH-720','location' => '4.B8'],
    ['sku' => 'JH-1102','location' => 'M.B3 - B4'],
    ['sku' => 'JH-521','location' => 'M.C11'],
    ['sku' => 'JH-330','location' => '4.B2'],
    ['sku' => 'LC-7121','location' => '2.B11'],
    ['sku' => 'JH-370','location' => '2.C 16'],
    ['sku' => 'JH-919','location' => '4.C11'],
    ['sku' => 'LC-7103','location' => '2.B1'],
    ['sku' => 'LC-7058','location' => '2.B1'],
    ['sku' => 'LC-7099','location' => '2.C3'],
    ['sku' => 'JH-1105','location' => '4.C13'],
    ['sku' => 'JH-1007','location' => '4.C6'],
    ['sku' => 'JH-1026','location' => '4.B4'],
    ['sku' => 'JH-808','location' => '4.B3'],
    ['sku' => 'JHX-207-JADE-GOLD','location' => 'Depan Jendela Lt.4'],
    ['sku' => 'JHX-207-SILVER-RUBY','location' => 'Depan Jendela Lt.4'],
    ['sku' => 'JH-511','location' => 'M.B5'],
    ['sku' => 'JH-801','location' => 'M.B3']
];

// $keyLoc = array_search('JH-366', array_column($defineLocationSKU, 'sku'));
// print_r(($keyLoc) ? $defineLocationSKU[$keyLoc]['location'] : '');
// die;

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
        $sheet = ['sku','name','url_key','quantity_dus','weight','product_type','categories','price','jade_price','diamond_price','gold_price','silver_price','ruby_price','qty','attribute_set_code','visibility','additional_attributes','configurable_variation_labels','configurable_variations','base_image','small_image','thumbnail_image','additional_images','product_websites','website_id','page_layout','location'];
        $new_sheet[] = $sheet;

        $oldsku = '';
        $oldname = '';
        $oldproduct = [];
        $configurable_variations_list = [];
        
        $defineFirstSKU = ['JH-','JHL-','HL-','LC-','BS-','TSB-','JHX-','FG-','JM-','HM-','JT JHW ','JHW ','JH ','HJ-','GQ-','LC ','JT ','HL ','BS '];
        $defineAllowedCat = ['BAG','WATCH','WALLET','PACKAGING','HEELS','WOMAN','HOMEWARES','POWERBANK','TUMBLER'];
        $defineHaveConfigurableImg = ['JH-357','JH-1108','JH-502','JHX-102-BRONZE-SILVER','JH-717','JH-918','JH-818','JH-922','JHX-201-RUBY-SILVER','JHX-201-GOLD-DIAMOND','JH-1028','LC-7112','JH-520','JHX-101-JADE-GOLD','JHX-101-SILVER-RUBY','HL-858','JH-361','JHX-203-BRONZE-SILVER','JHX-203-GOLD-DIAMOND','JH-1023','JH-307','ELLIS CARDHOLDER','JHX-202-BRONZE-SILVER','JHX-202-GOLD-DIAMOND','JH-359','JH-369','JH-2013','JH-508','JH-360','JH-306','JH-516','JH-510','BS-032','JHW 05','JHW 16','JHW 35','JHW 37','JHW 39','JHW 50','JH-1109','TSB-003','JT 2139','JT 8125 BIASA','LC-7107','LC-7111','600617','JH-1025','','JH-2012','JH-1022','LC-7052','JH-355','HL-857','LC-7089','JH-605','JH-519','MICHELLE BAG','JH-631','JH-721','JH-363','JH-1107','LC-7067','JH-503','TSB-006','','LC-7109','JH-1002','JH-358','LC-7102','TSB-007','JH-910','LC-7074','JH-915','JH-2005','BS-052','QUEEN BAG','TSB-005','JHX-208-JADE-GOLD','JHX-208-SILVER-RUBY','TSB-002','9TH','SUNNY WALLET','JH-718','JH-356','JH-720','JH-1102','JH-521','LC-7121','JH-1103','TUMBLER Ver4','JH-919','LC-7103','JH-1105','LC-7057','JH-1026','LC-7073','JH-309','JH-627','JH-518'];

        foreach ($sheetData as $key => $value) {
            //DUPLICATE
            if ($value['A'] == 'BEATRICE PLUS LC-7054 (1 dus : 60)') continue;
            if ($value['A'] == 'SCARLETT HL-836 (1Q :36pc)') continue;
            if ($value['A'] == 'CLARA BAG') continue;
            if ($value['A'] == 'CLARA HL-858 (1 DUS : 16pc)') continue;
            if ($value['A'] == 'TYA PLUS BAG JH-903 (1Q:24)') continue;
            if ($value['A'] == 'VANIA BAG') continue;
            if ($value['A'] == 'SUNNY WALLET') continue;
            if ($value['A'] == 'ROSIE BAG') continue;
            if ($value['A'] == 'NIKI WALLET') continue;
            if ($value['A'] == 'MERLIN BAG') continue;
            if ($value['A'] == 'AMOUR BAG') continue;
            if ($value['A'] == 'ADELINE BAG') continue;
            if ($value['A'] == 'jHW 53 (1Q:100)') $value['A'] = 'JHW 53 (1Q:100)';
            
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
            if ($value['D'] == 'IVORYWHTE') $value['D'] = 'IVORYWHITE';

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
                $oldproduct[25] = 'Product -- Full Width';
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
            $sheet[] = '';
            $keyLoc = array_search($sku, array_column($defineLocationSKU, 'sku'));
            $sheet[] = ($keyLoc != '') ? $defineLocationSKU[$keyLoc]['location'] : '';
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
            $oldproduct[25] = 'Product -- Full Width';
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