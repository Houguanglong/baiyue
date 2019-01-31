<?php
/**
 * Created by PhpStorm.
 * User: 侯光龙
 * Date: 2019/1/6
 * Time: 18:17
 */
namespace app\web\model\Excel;

vendor('PHPExcel.PHPExcel');

class Excel
{

    //传入的数组数据
    protected $data;

    //Excel类实例
    protected $excel_obj;

    //单元格索引
    protected $letter = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];

    //输出错误内容
    protected $error;

    //文件类型
    protected $writer_type = 'Excel5';

    //保存的文件名
    protected $save_file_name = 'Excel.xls';


    //表头实例
    protected $excel_sheet;

    //数据数量
    protected $data_count;

    //数据表头数据数组
    protected $data_head_field = [];

    //数据表头字段数组
    protected $data_head_fieldname=[];


    /**
     * 初始化方法
     * @param array $data 需要导出的数据
     * @param array $head_list 需要导出的数据的表头数组
     * @param string $save_name 导出的文件名
     * @param string $writer_type 导出Excel类型 默认Excel5
     */
    public function __construct(array $data,array $head_list,string $save_name,string $writer_type)
    {
        if(!is_array($data) || empty($data[0])){
            $this->error = '请传入要导出的数据!';
            return $this->show_error();
        }
        if(!is_array($head_list)){
            $this->error = '请传入生存表头数据!';
            return $this->show_error();
        }
        if(!empty($save_name)){
            $this->save_file_name = $save_name;
        }
        if(!empty($writer_type)){
            $this->writer_type = $writer_type;
        }
        $this->data_head_field = $head_list;
        $this->set_data($data);
        $this->browser_export($this->writer_type,$this->save_file_name);
    }


    /**
     * 设置导出的数据
     * @param array $data
     */
    public function set_data(array $data)
    {
        $this->data = $data;
        $this->data_count = count($this->data);
    }

    /**
     * 填充表格表头
     */
    public function set_cell_head()
    {
        $head_field_value = [];
        foreach ($this->data_head_field as $key=>$value1){
            array_push($this->data_head_fieldname,$key);
            array_push($head_field_value,$value1);
        }
        $this->data_head_field = $head_field_value;
        //设置表头
        foreach ($this->data_head_field as $key=>$value){
            $this->excel_sheet->setCellValue($this->letter[$key].'1',$value);
        }
        return $this;
    }


    /**
     * 填充表格内容
     */
    public function set_cell_value()
    {
        $j = 2;
        foreach ($this->data as $key=>$value){
            for ($i=0;$i<$this->data_count;$i++){
                $this->excel_sheet->setCellValue($this->letter[$i].($key+2),$this->data[$key][$this->data_head_fieldname[$i]]);
                $j++;
            }
        }
        return $this;
    }

    /**
     * 生成excel文件
     */
    public function save()
    {
        $objWriter=\PHPExcel_IOFactory::createWriter($this->excel_obj,$this->writer_type);
        $objWriter->save("php://output");
    }


    /**
     * 设置http头部信息
     * @param string $type 导出execl类型
     * @param string $filename 导出文件名
     */
    function browser_export($type,$filename){
        if($type=="Excel5"){
            header('Content-Type: application/vnd.ms-excel');//告诉浏览器将要输出excel03文件
        }else{
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');//告诉浏览器数据excel07文件
        }
        header('Content-Disposition: attachment;filename="'.$filename.'"');//告诉浏览器将输出文件的名称
        header('Cache-Control: max-age=0');//禁止缓存
    }

    /**
     * 导出文件
     */
    public function export()
    {
        $this->excel_obj = new \PHPExcel();
        $this->excel_sheet = $this->excel_obj->getActiveSheet();//获取当前活动sheet操作对象
        $this->excel_sheet->setTitle('内容表');
        $this->set_cell_head()->set_cell_value()->save();
    }


    /**
     * 输出错误信息
     */
    public function show_error()
    {
        echo $this->error;die;
    }


}