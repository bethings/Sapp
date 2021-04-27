<?php

$pack = new Pack();

$pack->run();

Class Pack
{
    /* Настройки */

    private $main_file = 'inv_nov.xlsx';// Эталонный файл где хранится информация о сопоставлении: инвентарные номера - штрихкоды
    private $main_inv_column = 2;// Порядковый номер столбца с инвентарным номером в основном файле (вместо A,B - 1,2)
    private $main_barcode_column = 3;// Порядковый номер столбца со штрихкодом в Эталонном файле (вместо A,B - 1,2)

    private $folder = 'files';// Папка с файлами для записи штрихкодов (будут обработаны все файлы, которые находятся в папке)
    private $file_inv_column = 2;// Порядковый номер столбца с инвентарным номером в целевых файлах
    private $file_barcode_column = 3;// Порядковый номер столбца для записи штрих кода в целевых файлах

    /* ----- */

    private $barcodes = array();

    public function __construct() {
        require_once 'library/PHPExcel.php';
        require_once 'library/PHPExcel/IOFactory.php';

        $this->main_inv_column--;
        $this->main_barcode_column--;

        $this->file_inv_column--;
        $this->file_barcode_column--;
    }

    public function run() {
        $this->importAllBarcodes();
        echo "import barcodes: success";
    
        if(!is_dir($this->folder)) {
            exit('Folder with files is not correct');
        }

        $files = scandir($this->folder);
        
        unset($files[0]);
        unset($files[1]);

        if(empty($files)) {
            exit('Folder with files is empty');
        }
        
        foreach($files as $file) {
            if(strpos($file, '~') === 0) {
                continue;
            }
            echo "<br>" . $file;
            $this->fillFile($file);
        }
    }

    private function importAllBarcodes() {
        if(!file_exists($this->main_file)) {
            exit('Cannot read all barcodes file');
        }

        $xls = PHPExcel_IOFactory::load($this->main_file);
        $xls->setActiveSheetIndex(0);
        $sheet = $xls->getActiveSheet();

        for($i = 1; $i <= $sheet->getHighestRow(); $i++) {	 
            $inv = $sheet->getCellByColumnAndRow($this->main_inv_column, $i)->getValue();
            $barcode = $sheet->getCellByColumnAndRow($this->main_barcode_column, $i)->getValue();
            
            // error_log($inv . " --- " . $barcode);

            if(!empty($inv)) {
                $this->barcodes[$inv] = $barcode;
            }
        }

        $xls = null;
        $sheet = null;
        unset($xls);
        unset($sheet);
    }

    public function fillFile($file) {
        $work_file = $this->folder . '/' . $file;

        if(empty($file) || !file_exists($work_file)) {
            echo "<br>[" . $file . "] is not exists";
            return false;
        }

        // $xls = PHPExcel_IOFactory::load($work_file);
        // $xls->setActiveSheetIndex(0);
        // $sheet = $xls->getActiveSheet();

        $file_type = 'Excel2007';
        
        $objReader = PHPExcel_IOFactory::createReader($file_type);
        $objPHPExcel = $objReader->load($work_file);

        $objPHPExcel->setActiveSheetIndex(0);
        $sheet = $objPHPExcel->getActiveSheet();
        
        for($i = 1; $i <= $sheet->getHighestRow(); $i++) {	 
            $inv = $sheet->getCellByColumnAndRow($this->file_inv_column, $i)->getValue();

            if(empty($inv)) {
                continue;
            }

            if(isset($this->barcodes[$inv])) {
                $sheet
                ->getCellByColumnAndRow($this->file_barcode_column, $i)
                ->setValueExplicit($value, PHPExcel_Cell_DataType::TYPE_NUMERIC);

                $sheet
                ->getStyleByColumnAndRow($this->file_barcode_column, $i)
                ->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
                // $sheet->setCellValueExplicit('A1', $val,PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->setCellValueByColumnAndRow($this->file_barcode_column, $i, (string)"" . $this->barcodes[$inv]);
            }
            else {
                error_log('not isset');
            }
        }

        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $file_type);
        $objWriter->save($work_file);

		// $objWriter = new PHPExcel_Writer_Excel5($xls);
		// $objWriter->save($work_file);
    }

    // public function import_excel() {

        
            //     if (file_exists($file_route)) {
    
            //         $json = array() ;
    
            //         require_once DIR_SYSTEM.'/library/PHPExcel.php';
    
            //         require_once DIR_SYSTEM.'/library/PHPExcel/IOFactory.php';
    
            //         $xls = PHPExcel_IOFactory::load($file_route);
    
            //         $xls->setActiveSheetIndex(0);
            //         // Получаем активный лист
            //         $sheet = $xls->getActiveSheet();
    
            //         $xls->setActiveSheetIndex(0);
            //         // Получаем активный лист
            //         $sheet = $xls->getActiveSheet();
                    
            //         for ($i = 2; $i <= $sheet->getHighestRow(); $i++) {	 
            //             $json[$i] = array();
    
            //             $json[$i]['product_id']		    = $sheet->getCellByColumnAndRow(0, $i)->getValue();
            //             $json[$i]['offer_product_id']   = $sheet->getCellByColumnAndRow(1, $i)->getValue();
            //             $json[$i]['language'][2]['title']   = $sheet->getCellByColumnAndRow(2, $i)->getValue();
            //             $json[$i]['language'][3]['title']   = $sheet->getCellByColumnAndRow(3, $i)->getValue();
            //             $json[$i]['language'][2]['link']   = $sheet->getCellByColumnAndRow(4, $i)->getValue();
            //             $json[$i]['language'][3]['link']   = $sheet->getCellByColumnAndRow(5, $i)->getValue();
            //             $json[$i]['image']              = $sheet->getCellByColumnAndRow(6, $i)->getValue();
            //             $json[$i]['title_pa']           = htmlspecialchars($sheet->getCellByColumnAndRow(7, $i)->getValue());
            //             $json[$i]['description_pa']     = htmlspecialchars($sheet->getCellByColumnAndRow(8, $i)->getValue());
            //             $json[$i]['priority']           = $sheet->getCellByColumnAndRow(9, $i)->getValue();
            //             $json[$i]['sort']               = $sheet->getCellByColumnAndRow(10, $i)->getValue();
            //             $json[$i]['status']             = $sheet->getCellByColumnAndRow(11, $i)->getValue();
            //         }
    
            //         $json = array_reverse($json);
            //     }
            // } else {
            //     $json['error'] = 'Імпортувати можна тільки файли EXCEL';
            // }				
        // }
                  
        // $this->response->setOutput(json_encode($json));						   
    // }
}
