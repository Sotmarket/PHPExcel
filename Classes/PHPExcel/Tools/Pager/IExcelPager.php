<?php
/**
 * 
 * @category  
 * @package   
 * @subpackage 
 * @author: u.lebedev
 * @date: 05.08.13
 * @version    $Id: $
 */
interface IExcelPager{
    public function __construct(
        PHPExcel_Worksheet $excelSheet,
        $lastDataRow=null
    );
    public function getSmoothedPageMap();
}