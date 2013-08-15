<?php
require_once ("IExcelPager.php");
/**
 * Пейджер для XLS шаблонов
 * @category  
 * @package   
 * @subpackage 
 * @author: u.lebedev
 * @date: 12.08.13
 * @version    $Id: $
 */

class PHPExcel_Tools_Pager_XlsPager extends PHPExcel_Tools_Pager_Pager implements IExcelPager{
    public function getSmoothedPageMap(){
        $result = parent::getSmoothedPageMap();
        $sheet = $this->getExcelSheet();
        $sheet->setBreaks(array());
        foreach ($result as $page){
            $sheet->setBreak('A'.$page->getFinish(),PHPExcel_Worksheet::BREAK_ROW);
        }
        return $result;
    }

}