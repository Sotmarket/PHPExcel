<?php
/**
 * Модель странички, с какого ряда начинать, каким заканчивать
 * @category  
 * @package   
 * @subpackage 
 * @author: u.lebedev
 * @date: 02.08.13
 * @version    $Id: $
 */
class PHPExcel_Tools_Document_PageModel{
    public $start;
    public $finish;

    /**
     * @param integer $start
     * @param integer $finish
     */
    public function __construct ($start, $finish){
        $this->start =  $start;
        $this->finish =  $finish;
    }

    public function getFinish()
    {
        return $this->finish;
    }

    public function getStart()
    {
        return $this->start;
    }

}