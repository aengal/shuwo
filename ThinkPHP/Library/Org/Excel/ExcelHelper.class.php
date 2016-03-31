<?php
namespace Org\Excel;
class ExcelHelper {  
	public function __construct() {  
		Vendor("Excel.PHPExcel");//引入phpexcel
	} 

	/**
	 * read
	 *
	 * 读取excel文件
	 *
	 * @param 	string 	$fileDir 	文件路径
	 * @param	int 	$sheet 		表编号，默认0为第一张表
	 * @return 	array 	$array      数组元素 bool result, string msg, array data
	 */
	public function read($fileDir, $sheet = 0){
		//判断文件后缀
		$ext = pathinfo($fileDir,PATHINFO_EXTENSION);
		if (strtolower($ext) !== 'xls' && strtolower($ext) !== 'xlsx') {
			return array('result'=>false,'msg'=>'文件格式错误');
		}
		try {
			//初始化
			new \PHPExcel;
			//载入文件
			$PHPExcel = \PHPExcel_IOFactory::load($fileDir);
			//获取工作表
			$currentSheet = $PHPExcel->getSheet($sheet);
			//获取总列数
			$allColumn = $currentSheet->getHighestColumn();
			//获取总行数
			$allRow = $currentSheet->getHighestRow();
			//循环获取表中的数据，$currentRow表示当前行，从哪行开始读取数据，索引值从0开始
			$sheetData = array();
			for ($currentRow=2; $currentRow<=$allRow; $currentRow++){
				//从哪列开始，A表示第一列
				for($currentColumn='A'; $currentColumn<=$allColumn; $currentColumn++){
					//数据坐标
					$address = $currentColumn.$currentRow;
					//读取到的数据，保存到数组$sheetData中
					// $sheetData[$currentRow][$currentColumn] = $currentSheet->getCell($address)->getValue();
					$array[$currentColumn] = $currentSheet->getCell($address)->getValue();
				}
				array_push($sheetData, $array);
			}
		} catch (\Exception $e) {
			return array('result'=>false,'msg'=>$e->getMessage());
		}
		return array('result'=>true,'data'=>$sheetData);
	}

	public function write($fileName,$headArr,$data){
		//对数据进行检验
	    if(empty($data) || !is_array($data)){
	        die("data must be a array");
	    }
	    //检查文件名
	    if(empty($fileName)){
	        exit;
	    }

	    $date = date("Y_m_d",time());
	    $fileName .= "_{$date}.xls";

		//创建PHPExcel对象，注意，不能少了\
	    $objPHPExcel = new \PHPExcel();
	    $objProps = $objPHPExcel->getProperties();
		
	    //设置表头
	    $key = ord("A");
	    foreach($headArr as $v){
	        $colum = chr($key);
	        $objPHPExcel->setActiveSheetIndex(0) ->setCellValue($colum.'1', $v);
	        $key += 1;
	    }
	    
	    $column = 2;
	    $objActSheet = $objPHPExcel->getActiveSheet();
	    foreach($data as $key => $rows){ //行写入
	        $span = ord("A");
	        foreach($rows as $keyName=>$value){// 列写入
	            $j = chr($span);
	            $objActSheet->setCellValueExplicit($j.$column, $value, \PHPExcel_Cell_DataType::TYPE_STRING);//指定字符串写入
	            $span++;
	        }
	        $column++;
    	}

	    $fileName = iconv("utf-8", "gb2312", $fileName);
	    //重命名表
	   	// $objPHPExcel->getActiveSheet()->setTitle('test');
	    //设置活动单指数到第一个表,所以Excel打开这是第一个表
	    $objPHPExcel->setActiveSheetIndex(0);
	    header('Content-Type: application/vnd.ms-excel');
		header("Content-Disposition: attachment;filename=\"$fileName\"");
		header('Cache-Control: max-age=0');

	  	$objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
	    $objWriter->save('php://output'); //文件通过浏览器下载
	    exit;
	}
}
?>