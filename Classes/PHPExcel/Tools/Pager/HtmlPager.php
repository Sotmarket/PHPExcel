<?php
/**
 * ������� ��� html.
 * @category  
 * @package   
 * @subpackage 
 * @author: u.lebedev
 * @date: 06.08.13
 * @version    $Id: $
 */
class PHPExcel_Tools_Pager_HtmlPager  extends PHPExcel_Tools_Pager_Pager implements IExcelPager{

    const PRECISION_ROW=5;

    public function getPageHeight()
    {
        // ��������� ���, �.� �������� �� ������� �������������� ������ �����, ����� ������ �������
        return parent::getPageHeight()-120;
    }

}