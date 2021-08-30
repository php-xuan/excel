<?php
/**
 * Created by
 * User: GuoZhaoXuan
 * Date: 2021/8/24
 * Time: 9:46
 */

namespace App\Utility\Excel;


use EasySwoole\Component\Singleton;
use App\Utility\Common;

/**
 * php-excel 封装
 *    composer: phpoffice/phpexcel
 *     支持:1.断点续导 2.一次分批导出 3.多sheet/无限极头
 *
 * @author vartruexuan <guozhaoxuanx@163.com>
 * @package App\Utility\Excel
 */
class PhpExcelHelper
{
    use Singleton;

    /**
     * 导出excel
     *
     * @param array  $sheets   页码参数,可多页
     * @param string $filePath 导出文件地址（未提供，则导入到临时文件）
     * @param int    $pageSize 分页导出时，一页导出数量
     * @param bool   $isAdd    是否追加
     *
     * @return bool
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     * @throws \PHPExcel_Writer_Exception
     */
    public function exportExcel(array $sheets, $filePath = "", $pageSize = 5000, $isAdd = false)
    {
        // 页码sheets设置备注
        /*        [
                    [
                        'sheet_name' => '页码名称',
                        'sheet_header' => '页头信息(支持多级)',// array:参考 setHeader()
                        // 全局操作
                        'default_format' => function (\PHPExcel_Worksheet $sheet) {
                        },
                        // $params => ['limit'=>1,'offset'=>1]  设置 data_count 时提供
                        'data' => function ($params) { //  数据： callback | array
                        },
                        "data_count" => function () {  //    数据总数量: integer| callback
                            return 2;
                        },
                        "is_calculation_colspan" => true, // 是否智能计算colspan

                    ],
                ];*/

        // 是否追加数据
        $excelObj = $this->getExcel($isAdd, $filePath);
        $pageSize = $pageSize ? $pageSize : 5000;
        foreach ($sheets as $k => $sheet) {
            $sheetName = $sheet['sheet_name'];
            $sheetHeader = $sheet['sheet_header'];
            $dataFun = $sheet['data'] ?? [];
            $dataCount = $sheet['data_count'] ?? false;
            $defaultFormat = $sheet['default_format'] ?? null; // 全局默认单元格样式
            $isCalculationColspan = $sheet['is_calculation_colspan'] ?? true;// 是否只能计算colspan
            $isFreezePane=$sheet['is_freeze_pane']??false;// 是否冻结头部
            $totalCount = $dataCount;
            if (is_callable($dataCount)) {
                $totalCount = $dataCount();
            }
            if (!$excelObj->sheetNameExists($sheetName)) {
                $sheet = $excelObj->createSheet($k);
                $sheet->setTitle($sheetName);
            }
            $sheet = $excelObj->setActiveSheetIndex($k);
            // 设置头
            if ($isCalculationColspan) {
                $sheetHeader = $this->calculationColspan($sheetHeader);
            }
            $endColIndex = -1;
            $this->addHeader($sheetHeader, $sheet, $maxRow, $dataHeaders, 1, $endColIndex,$isAdd,$isFreezePane);
            if ($isAdd) {
                // 追加时起始位置
                $maxRow = $sheet->getHighestRow();
            }
            // 设置全局样式
            if ($defaultFormat) {
                $defaultFormat($sheet);
            }
            // 获取列的位置
            $keysIndex = array_flip(array_column($dataHeaders, 'key'));

            // 设置数据
            $pageCount = 1;
            if (is_callable($dataFun) && $totalCount) {
                $pageCount = ceil($totalCount / $pageSize);
            }
            // 分批次导入
            for ($page = 1; $page <= $pageCount; $page++) {
                $params['limit'] = $pageSize;
                $params['offset'] = ($page - 1) * $pageSize;
                $data = $dataFun;
                if (is_callable($dataFun)) {
                    $data = call_user_func($dataFun, $params);
                }
                // 数据格式化
                foreach ($data ?? [] as $k => &$v) {
                    $newVal = [];
                    $rowIndex=$maxRow + $params['offset'] + $k + 1;
                    foreach ($dataHeaders as $colIndex => $head) {
                        // 执行单元格回调
                        if (is_callable($head['cellFormat'])) {
                            $newVal[$head['key']] = call_user_func_array($head['cellFormat'], [
                                'key' => $head['key'],
                                'row' => $v, // 行数据
                                'rowIndex' =>$rowIndex, // 行索引
                                'colIndex' => $keysIndex[$head['key']], // 列索引
                            ]);
                        } else {
                            $newVal[$head['key']] = $v[$head['key']] ?? '';
                        }
                        // 样式回调
                        if (is_callable($head['style'])) {
                            $head['style']($sheet, $rowIndex, $keysIndex[$head['key']],false,$v);
                        }
                    }
                    $dataType = array_column($dataHeaders, 'type');
                    // 插入数据
                    $this->writerRow($sheet, $newVal, $rowIndex, $dataType);

                }
            }
        }

        $writer = \PHPExcel_IOFactory::createWriter($excelObj, 'Excel2007');
        $writer->save($filePath);

        return true;
    }

    /**
     * 获取excel对象
     *
     * @param false $isAdd 是否追加
     * @param null  $filePath
     *
     * @return \PHPExcel
     * @throws \PHPExcel_Reader_Exception
     */
    protected function getExcel($isAdd = false, $filePath = null)
    {
        if ($isAdd) {
            $type = \PHPExcel_IOFactory::identify($filePath);
            $objReader = \PHPExcel_IOFactory::createReader($type);
            $excelObj = $objReader->load($filePath);
        } else {
            $excelObj = new \PHPExcel();
        }

        return $excelObj;
    }


    /**
     * 设置头部(支持多级)
     *
     * @param  array  $headers
     * @param  \PHPExcel_Worksheet  $sheet
     * @param  int  $maxRow
     * @param  array  $dataHeaders
     * @param  int  $rowIndex
     * @param  int  $endColIndex
     * @param  false  $isAdd
     * @param  false  $isFreezePane
     *
     * @throws \PHPExcel_Exception
     */
    protected function addHeader(array $headers, \PHPExcel_Worksheet $sheet, &$maxRow = 1, &$dataHeaders = [], $rowIndex = 1, &$endColIndex = -1, $isAdd = false,$isFreezePane=false)
    {
        $this->setHeader($headers, $sheet, $maxRow, $dataHeaders, 1, $endColIndex,$isAdd);
        if($isFreezePane){
            // 冻结头部
            for($j=$rowIndex;$j<=$maxRow+1;$j++){
                $sheet->freezePaneByColumnAndRow(0,$j);
            }
        }
    }

    /**
     *
     * 设置头部(支持多级)
     *
     * @param array               $headers     头参数
     * @param \PHPExcel_Worksheet $sheet       操作对象
     * @param int                 $maxRow      头部占用行数
     * @param array               $dataHeaders 数据字段信息（包含 key ）
     * @param int                 $rowIndex    当前行
     * @param int                 $endColIndex 当前结束列
     * @param bool                $isAdd       是否追加数据
     * @param array               $style       样式（待）
     * @param int                 $level       当前层级
     *
     * @return bool
     * @throws \PHPExcel_Exception
     */
    protected function setHeader(array $headers, \PHPExcel_Worksheet $sheet, &$maxRow = 1, &$dataHeaders = [], $rowIndex = 1, &$endColIndex = -1, $isAdd = false, $style = [], $level = 1)
    {

        // headers 配置：可无限极
        /*        [
                    [
                        "title" => "列名称",
                        "type" => \PHPExcel_Cell_DataType::TYPE_STRING,// 数据类型
                        "key" => "name", // 数据key
                        "style" => function (\PHPExcel_Worksheet $sheet,$rowIndex, $colIndex,$isHeader,$row) { // 当前head 单元格样式
                                    // $sheet 页操作对象
                                    // $rowIndex 当前行索引
                                    // $colIndex 当前列索引
                                    // $isHeader 当前是否是header头部
                                    // $row 当前行数据（$isHeader=false）
                        },
                        // 列数据格式化 (提供key时)
                        "cellFormat" => function ($key, $row, $rowIndex, $colIndex, $colIndexStr) {
                            // $key 键值  $row 当前行 $rowIndex 当前行下标   $colIndex 当前列下标
                            return (float)$row[$key];
                        },

                        "rowspan" => 2, // 跨行数
                        "colspan" => 2,  // 跨列数
                        // 子级
                        "children" => [
                            [
                                "title" => '',
                                ...
                            ]

                        ],
                    ]
                ];*/

        foreach ($headers as $headerIndex => $head) {
            // 设置默认参数
            $head = array_merge([
                "title" => "",
                "type" => \PHPExcel_Cell_DataType::TYPE_STRING,
                "key" => "",
                "style" => function (\PHPExcel_Worksheet $sheet, $rowIndex, $startColIndex) {
                },
                "width" => 0,// 宽度
                "children" => [],
                "colspan" => 1,
                "rowspan" => 1,
            ], $head);


            // 设置数据 key/类型 （排序位置）
            if ($head['key']) {
                $dataHeaders[] = [
                    'key' => $head['key'],
                    'type' => $head['type'] ?? \PHPExcel_Cell_DataType::TYPE_STRING,
                    'cellFormat' => $head['cellFormat'] ?? null,
                    'style'=>$head['style']??null,
                ];
            }

            // 列位置：起始-结束
            $startColIndex = $endColIndex + 1;
            $endColIndex = $startColIndex + $head['colspan'] - 1;

            // 行位置：起始-结束
            $startRow = $rowIndex;
            $endRow = $startRow + $head['rowspan'] - 1;
            if ($endRow > $maxRow) {
                $maxRow = $endRow;
            }

            // 非追加数据时：不插入头信息
            if (!$isAdd) {
                // 设置列宽
                $head['width'] = $head['width'] > 0 ? $head['width'] : strlen($head['title']) + 5;
                // 合并行 A1:B3
                $sheet->mergeCellsByColumnAndRow($startColIndex, $startRow, $endColIndex, $endRow);

                $this->writerCell($sheet, $startColIndex, $startRow, $head['title'], \PHPExcel_Cell_DataType::TYPE_STRING, function (\PHPExcel_Cell $cell) use ($level, $sheet, $rowIndex, $startColIndex, $head, $endRow, $endColIndex, $headerIndex) {
                    // 设置默认样式
                    $cell->getStyle()->getFont()->setBold(true);// 加粗
                    $cell->getStyle()->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);// 文字居中
                    $column = $sheet->getColumnDimension(self::stringFromColumnIndex($startColIndex));
                    if ($column->getWidth() < $head['width']) {
                        $column->setWidth($head['width']);// 自动宽度
                    }
                    // 相同父子级加上边框
                    if ($level == 1) {
                        $styleThinBlackBorderOutline = [
                            'borders' => [
                                'allborders' => [ //设置全部边框
                                                  'style' => \PHPExcel_Style_Border::BORDER_THIN, //粗的是thick
                                                  'color' => [
                                                      'argb' => \PHPExcel_Style_Color::COLOR_BLUE
                                                  ],
                                ],
                            ],
                        ];
                        $cellRange = self::stringFromColumnIndex($startColIndex) . $rowIndex . ':' . self::stringFromColumnIndex($endColIndex) . $endRow;
                        $sheet->getStyle($cellRange)->applyFromArray($styleThinBlackBorderOutline);
                        // 设置个背景色
                        $color = $headerIndex % 2 == 0 ? \PHPExcel_Style_Color::COLOR_DARKBLUE : \PHPExcel_Style_Color::COLOR_DARKGREEN;
                        $sheet->getStyle($cellRange)->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB($color);
                        $sheet->getStyle($cellRange)->getFont()->getColor()->setARGB(\PHPExcel_Style_Color::COLOR_WHITE);
                    }
                    // 样式回调
                    if (isset($head['style']) && is_callable($head['style'])) {
                        $head['style']($sheet, $rowIndex, $startColIndex,true);
                    }
                });
            }

            // 子级
            if (isset($head['children']) && $head['children']) {
                $endColIndex = $startColIndex - 1;
                $this->setHeader($head['children'], $sheet, $maxRow, $dataHeaders, $rowIndex + $head['rowspan'], $endColIndex, $isAdd, [], $level + 1);
            }
        }

        return true;
    }

    /**
     * 计算colspan(多级)
     *
     * @param     $header
     * @param int $level
     *
     * @return mixed
     */
    protected function calculationColspan($header, $level = 1)
    {
        // 子集colspan之和
        foreach ($header as &$head) {
            $children = $head['children'] ?? [];
            if ($children) {
                $head['children'] = $this->calculationColspan($children, $level + 1);
                $head['colspan'] = array_sum(array_column($head['children'], 'colspan'));
            } else {
                $head['colspan'] = 1;
            }
        }
        return $header;
    }

    /**
     * 写入行数据
     *
     * @param \PHPExcel_Worksheet $sheet        当前sheet对象
     * @param array               $row          当前行数据
     * @param int                 $rowIndex     行下标
     * @param array               $colDataTypes 列数据类型 []
     *
     * @author:郭昭璇
     */
    protected function writerRow(\PHPExcel_Worksheet $sheet, $row, $rowIndex = 1, $colDataTypes = [])
    {
        $row=array_values($row);
        foreach ($row as $columnIndex => $value) {
            $pDataType = \PHPExcel_Cell_DataType::TYPE_STRING;// 默认字符串
            $pDataType = $colDataTypes[$columnIndex] ?? $pDataType;
            $this->writerCell($sheet, $columnIndex, $rowIndex, $value, $pDataType);
        }
    }

    /**
     *
     * 写入单元格数据
     *
     * @param \PHPExcel_Worksheet $sheet
     * @param int                 $columnIndex 列下标
     * @param int                 $rowIndex    行下标
     * @param mixed               $value       列值
     * @param string              $pDataType   当前列数据类型PHPExcel_Cell_DataType
     * @param callable|null       $cellFun     单元格处理回调:function(\PHPExcel_Cell $cell){ .... }
     */
    protected function writerCell(\PHPExcel_Worksheet $sheet, $columnIndex, $rowIndex, $value, $pDataType = \PHPExcel_Cell_DataType::TYPE_STRING, callable $cellFun = null)
    {
        $cell = $sheet->setCellValueExplicitByColumnAndRow($columnIndex, $rowIndex, $value, $pDataType, true);
        if (is_callable($cellFun)) {
            $cellFun($cell);
        }
        return $cell;
    }

    /**
     * 列标字符转化（下标转字符）
     *
     * @param $index
     *
     * @return string|mixed
     */
    public static function stringFromColumnIndex($index)
    {
        return \PHPExcel_Cell::stringFromColumnIndex($index);
    }

    /**
     * 列标字符转化（字符转下标）
     *
     * @param $strIndex
     *
     * @return integer|mixed
     */
    public static function columnIndexFromString($strIndex)
    {
        return \PHPExcel_Cell::columnIndexFromString($strIndex);
    }

    /**
     * 获取header主题，配色方案(待)
     */
    protected function getHeaderTheme()
    {

        return [
            'default'=>[],
        ];
    }

}
