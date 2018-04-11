<?php
/**
 * 数组数据转换
 * @author wangjl
 */
class ArrayToExcel {
    
    public $objPHPExcel;
    public $objSheet;//当前sheet
    public $position;//位置

    private static $_instance;
    
    private $sheetList = 'ABCDEFJHIJKLMNOPQRSTUVWXYZ';
    
    public function __construct()
    {
        require_once dirname(__FILE__).'./phpoffice/phpexcel/Classes/PHPExcel.php';
//         require dirname(__FILE__).'/../vendor/autoload.php';
        // Create new PHPExcel object
        $this->objPHPExcel = new PHPExcel();
        // Set document properties
        $this->objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
        ->setLastModifiedBy("Maarten Balliauw")
        ->setTitle("Office 2007 XLSX Test Document")
        ->setSubject("Office 2007 XLSX Test Document")
        ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
        ->setKeywords("office 2007 openxml php")
        ->setCategory("Test result file");
        
        $this->objSheet = $this->objPHPExcel->getActiveSheet();
        
        $this->objSheet->getDefaultStyle()->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
        ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        
        //设置水平垂直居中
        $this->objSheet->getDefaultStyle()->getFont()->setName("微软雅黑")->setSize(12); //设置默认字体大小
        
        $this->objPHPExcel->setActiveSheetIndex(0);
        
        $this->position = 1;
    }
    
    public static function getInstance()
    {
        if(empty(self::$_instance))
        {
            self::$_instance = new self();
        }
        return self::$_instance;
    }
    
    /**
     * 渲染单个项目
     * @param array $pj
     */
    public function render(array $pj)
    {
        if(!empty($pj['fields']))
        {
            $fieldNums = count($pj['fields']);
        }
        $this->objSheet->getDefaultStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_BLACK);
        
        //标题
        if(!empty($pj['title']))
        {
            $titleBox = "A".$this->position.":".$this->toAZ($fieldNums-1).$this->position;
            $this->objSheet->mergeCells($titleBox); //合并单元格
            $this->objSheet->setCellValue("A".$this->position,$pj['title']);
            $this->objSheet->getStyle("A".$this->position)->getFont()->setSize(14)->setBold(true); //标题字体
            $this->objSheet->getStyle("A".$this->position)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('3399ff'); //设置标题背景颜色
            $this->position++;            
        }

        //描述
        if(!empty($pj['desc']))
        {
            $this->objSheet->setCellValue("A".$this->position,$pj['desc']);
            $titleBox = "A".$this->position.":".$this->toAZ($fieldNums-1).$this->position;
            $this->objSheet->mergeCells($titleBox); //合并单元格            
            $this->objSheet->getStyle('A'.$this->position)->getAlignment()->setWrapText(true); //设置换行
            $this->objSheet->getStyle("A".$this->position)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('3399ff'); //设置标题背景颜色
            $this->objSheet->getDefaultStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
            $this->position++;
        }
        
        $this->objSheet->getDefaultStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_BLACK);
        
        //遍历字段
        foreach($pj['fields'] as $f=>$field)
        {
            $this->objSheet->setCellValue($this->toAZ($f).$this->position,$field);
            $this->objSheet->getStyle($this->toAZ($f).$this->position)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('3399ff'); //设置标题背景颜色
            $this->objSheet->getStyle($this->toAZ($f).$this->position)->applyFromArray(getBorderStyle("#66FF99")); //设置标题边框
            
        }
        $this->position++;
        
        //便利数据
        foreach($pj['list'] as $data)
        {
            foreach($data as $k=>$d)
            {
                $this->renderLine($k,$d);
            }
            $this->position++;
        }
    }
    
    /**
     * 渲染单行
     * @param 列数 $k
     * @param 数据值 $d
     */
    public function renderLine($k,$d)
    {
        if(is_array($d))
        {
               $d = json_encode($d);
        }
        $this->objSheet->setCellValue($this->toAZ($k).$this->position,$d);
    }
    
    
    /**
     * 渲染多个数组
     * @param array $projects
     */
    public function multiRender(array $projects)
    {
        foreach($projects as $pl)
        {
            $this->render($pl);
            $this->position++;
        }
    }
    
    private $headerHtml;
    private $footerHtml;
    private $styleHtml;
    
    public function setStyleHtml($styleHtml)
    {
        $this->styleHtml = $styleHtml;
    }
    
    public function setHeaderHtml($html)
    {
        $this->headerHtml = $html;
    }
    
    public function setfooterHtml($html)
    {
        $this->footerHtml = $html;
    }
    
    
    /**
     * 获取html标签数据
     */
    public function getHtml()
    {
        $objWriteHTML = new \PHPExcel_Writer_HTML($this->objPHPExcel);  //读取excel文件，并将它实例化为PHPExcel_Writer_HTML对象
        $style='<style> .sheet0 td{border:0px solid #fff }</style>';
//         $style='<link rel="stylesheet" href="//cdnjs.cloudflare.com/ajax/libs/bootstrap-table/1.12.1/bootstrap-table.min.css">';
        $html = $objWriteHTML->generateHTMLHeader(true).$style.$objWriteHTML->generateSheetData().$objWriteHTML->generateHTMLFooter();
        return $html;
    }
    
    /**
     * 保存xlsx文件
     * @param string $path
     */
    public function saveExcel($path='')
    {
        $objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel2007');        
        $objWriter->save( 'export.xlsx');
    }
    
    
    /**
     * 导出下载excel
     * @param string $fileName
     */
    public function exportExcel($fileName = 'temp')
    {
        //生成xlsx文件
        
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $fileName . '.xlsx"');
        header('Cache-Control: max-age=0');
        $objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel2007');
        
        //生成xls文件
        /*
         header('Content-Type: application/vnd.ms-excel');
         header('Content-Disposition: attachment;filename="'.$filename.'.xls"');
         header('Cache-Control: max-age=0');
         $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        */
        $objWriter->save('php://output');
    }
    
    private function toAZ($n)
    {
        $str = (string) $this->sheetList;
        return $str{$n};
    }
}


/*
 *获得不同颜色的边框
 */
function getBorderStyle($color){
    $styleArray = array(
        'borders' => array(
            'outline' => array(
                'style' => PHPExcel_Style_Border::BORDER_THICK,
                'color' => array('rgb' => $color),
            ),
        ),
    );
    return $styleArray;
}