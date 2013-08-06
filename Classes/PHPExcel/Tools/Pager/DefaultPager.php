<?php
/**
 * Пейджер заглушка - делегирует разбивку на старницы далее
 * @category  
 * @package   
 * @subpackage 
 * @author: u.lebedev
 * @date: 05.08.13
 * @version    $Id: $
 */

class PHPExcel_Tools_Pager_DefaultPager implements  IExcelPager{
    protected $excelSheet;

    /**
     *
     * @param PHPExcel_Worksheet $excelSheet
     * @param null               $lastDataRow
     */

    public function __construct (
        PHPExcel_Worksheet $excelSheet,
        $lastDataRow=null
    ){
        $this->setExcelSheet($excelSheet);
    }

    /**
     * Получить карту разбивки по страницам
     * @return array
     */
    public function getSmoothedPageMap(){
        $highestRow = $this->getExcelSheet()->getHighestRow();
        return array (1=>new PHPExcel_Tools_Document_PageModel(1, $highestRow));
    }

    /**
     * @param PHPExcel_Worksheet $excelSheet
     * @return $this
     */
    protected function setExcelSheet( PHPExcel_Worksheet $excelSheet)
    {
        $this->excelSheet = $excelSheet;
        return $this;
    }

    /**
     * @return PHPExcel_Worksheet
     */
    protected function getExcelSheet()
    {
        return $this->excelSheet;
    }

}