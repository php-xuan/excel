<?php
/**
 * Created by PhpStorm.
 * User: 小昭昭
 * Date: 2018/6/13
 * Time: 13:56
 */

namespace app\common\components\excel;

use yii\helpers\FileHelper;

class ExcelHelper
{
    #####说明####
    # 1. 基于插件 mk-j/php_xlsxwrite composer安装
    # 使用:
    /*      导出:
                $rs = ExcelHelper::exportExcel("", $sheets);
                if (is_int($rs)) {
                     return $this->jsonReturnError("导出文件失败({$rs})");
                }
                # 返回下载地址
                $result['url'] = ExcelHelper::getUrl("学员统计", $rs['token']);
    */

    # 2. self::exportCsv 导出cvs 无需插件
    ######## 单元格可设置样式表 ########
    # 样式	        允许值
    # font	        Arial, Times New Roman, Courier New, Comic Sans MS
    # font-size	    8,9,10,11,12 ...
    # font-style	bold, italic, underline, strikethrough or multiple ie: 'bold,italic'
    # border	    left, right, top, bottom, or multiple ie: 'top,left'
    # border-style	thin, medium, thick, dashDot, dashDotDot, dashed, dotted, double, hair, mediumDashDot, mediumDashDotDot, mediumDashed, slantDashDot
    # border-color	#RRGGBB, ie: #ff99cc or #f9c
    # color     	#RRGGBB, ie: #ff99cc or #f9c
    # fill	        #RRGGBB, ie: #eeffee or #efe
    # halign	    general, left, right, justify, center
    # valign	    bottom, center, distributed
    /**
     * 下载对应文件地址
     *
     * @var string
     */
    public static $excelDownUrl = "/common/excel-down";
    /**
     * excel文件临时存放目录(相对于项目根目录)
     *
     * @var string
     */
    public static $temDir = "/uploads/excel_tmp";
    /**
     * 错误码
     *
     * @var array
     */
    public static $errors = [
        1001 => "获取临时文件名失败",
        1002 => "文件不存在",
        1003 => "sheet页参数为空",
        1004 => "读取文件失败",
    ];
    /**
     * 默认文件失效时间1天
     *
     * @var int
     */
    public static $expireTime = 3600 * 24;

    private static $hooks = [];
    /**
     * 写入行事件钩子
     * params:
     *  row:当前行数据
     *  row_index:当前行索引
     *  write:当前写入对象
     */
    public const HOOK_WRITE_ROW_BEFORE = "writeRowBefore";
    public const HOOK_WRITE_ROW_AFTER = "writeRowAfter";

    /**
     * 保存成excel(保存到文件或直接下载) 多页sheets
     *
     * @param array  $sheets sheet页集合,参数
     *         [
     *               [
     *                     # 页名1 (string)
     *                     "sheet_name"=>"sheet1",
     *                     # 列标题 (array)
     *                     "sheet_header"=>[
     *                            [
     *                                "title"=>"名称",//excel显示的标题
     *                                "type"=>"string",//字段类型
     *                                "key"=>"name",//字段名
     *                                "style"=>[],//当前列样式,例如:["font-size"=>18]
     *                            ],
     *                            [
     *                                "title"=>"名称2",//excel显示的标题
     *                                "type"=>"string",//字段类型
     *                                "key"=>"sex",//字段名
     *                            ]
     *                     ],
     *                     # 数据回调 (callback)
     *                     "data_fun"=>function($params){
     *                            return table::getData($params['limit'],$params['offset']);
     *                     },
     *                     # 数据总数量(int|callback|false 只导出一次适用于不查询的一次导出)
     *                     "data_count"=>1000 | function(){return total;}
     *                     # 字段值格式化 (头字段使用数据中没有的字段可通过此方式赋值)
     *                     "data_filter"=>[
     *                           # return fasle:销毁该字段, 否则便是重新赋值给该字段
     *                           "字段1"=>function($row){ return "";},
     *                           字段2"=>function($row){return "";},
     *                     ],
     *                     # 行样式
     *                     # ["字段key"=>["font-size"=>18]] 可为当前行某个单元格设置样式
     *                     # ["font-size"=>18] 为行样式应用当前行所有单元格
     *                     # $row:当前行数据 $row_index:当前行索引 $keys 当前要导出的键  (可用来一些隔行换色操作)
     *                     "row_style"=>function($row,$row_index,$keys){
     *                                 if($row['sex']='男'){
     *                                        return ["font-size"=>18]
     *                                 }
     *                                 return ["sex"=>["font-size"=>18]];
     *                     },
     *                     # 合并行|单元格回调
     *                     "merge_fun" => function ($row, $row_index,$keys) {
     *                              return [
     *                                        [
     *                                            "start_row" => $row_index,//合并起始行索引
     *                                            "end_row" => $row_index-1,//合并结束行索引
     *                                            "start_col" => 1,//合并列起始索引
     *                                            "end_col" => 2,//合并列结束索引
     *                                         ]
     *                              ];
     *                     },
     *
     *              ],
     *
     *        ]
     * @param string $filePath 指定路径保存| 未指定路径将保存到临时路径中(可返回token | url)
     * @param int    $pageSize 每次读取输出数量(默认5000)
     *
     * @return bool|int|array  (array('path'=>'文件绝对地址','token'=>'提供给down_excel() 下载文件时使用') 不下载文件地址 int 错误码)
     * @throws \yii\base\Exception
     */
    public static function exportExcel(Array $sheets, $filePath = "", $pageSize = 5000)
    {
        if (!$sheets) {
            return 1003;
        }
        @ob_end_clean();
        set_time_limit(0);
        # 查询每页的最大数量
        $pageSize = $pageSize ? $pageSize : 5000;
        $writer = new \XLSXWriter();

        # 写入头部列名
        foreach ($sheets as $sheetIndex => $sheet) {
            # 分别获取头部信息
            $titles = array_column($sheet['sheet_header'], 'title'); // 展示字段头
            $types = array_column($sheet['sheet_header'], 'type');   // 字段展示类型
            $keys = array_column($sheet['sheet_header'], 'key');     // 对应数据中的字段名
            $row_style = $sheet['row_style'] ?? []; //行样式
            $merge_fun = $sheet['merge_fun'] ?? null;
            $column_style = array_column($sheet['sheet_header'], 'style', 'key');//当前列默认样式
            # 设置列字段
            $sheetHeader = array_combine($titles, $types);
            # 绑定事件(合并行|单元格)
            self::bindHook(self::HOOK_WRITE_ROW_AFTER,
                function ($row, $row_index, $writer) use ($merge_fun, $sheet, $keys) {
                    if (is_callable($merge_fun)) {
                        $merge_params = call_user_func_array($merge_fun, [
                            $row,
                            $row_index,
                            $keys,
                        ]);
                        $merge_params = $merge_params ?? [];
                        foreach ($merge_params as $mp) {
                            # 合并单元格
                            $writer->markMergedCell($sheet['sheet_name'], $mp['start_row'], $mp['start_col'],
                                $mp['end_row'], $mp['end_col']);
                        }
                    }

                });
            $writer->writeSheetHeader($sheet['sheet_name'], $sheetHeader);//optional
            # 进行分页分批导出（>0 :限制总导出数量 / null(默认) :则不限制直到导出至无数据 /false: 只导出一次）
            if (isset($sheet['data_count'])) {
                $total_count = (int)is_callable($sheet['data_count']) ? $sheet['data_count']() : $sheet['data_count'];
                # 只导出一次
                if ($total_count === false) {
                    $page_count = 1;
                }
                # 限制总导出数量
                if ($total_count > 0) {
                    $page_count = ceil($total_count / $pageSize);
                }
            }
            for ($page = 1; (isset($page_count) ? ($page <= $page_count) : true); $page++) {
                $params['limit'] = $pageSize;
                $params['offset'] = ($page - 1) * $pageSize;
                # 最后一页
                if (isset($page_count) && $total_count !== false && $page_count == $page) {
                    $params['limit'] = $total_count % $pageSize;
                }
                # 查询对应数据（数据为空则代表最后）
                $data = $sheet['data_fun']($params);
                if (empty($data)) {
                    break;
                }
                # 写入数据
                foreach ($data as $k => $v) {
                    $row_index = $params['offset'] + $k + 1;
                    # 设置排序 和过滤字段
                    $v2 = [];
                    $style = is_callable($row_style) ? $row_style($v, $row_index, $keys) : $row_style;//行样式
                    foreach ($keys as $key) {
                        $v2[ $key ] = (isset($v[ $key ]) && $v[ $key ] !== null) ? $v[ $key ] : '';
                        # 格式化数据
                        $filter_fun = $sheet['data_filter'][ $key ] ?? null;
                        if ($filter_fun && is_callable($filter_fun)) {
                            $v2[ $key ] = $filter_fun($v);
                        }
                        # 处理样式
                        $merge_style[ $key ] = $column_style[ $key ] ?? [];//默认列样式
                        if ($style && is_array($style)) {
                            # 取行样式中对应列样式
                            $merge_style[ $key ] = (count($style) == count($style,
                                    COUNT_RECURSIVE)) ? $style : ($style[ $key ] ?? $merge_style[ $key ]);
                        }
                    }
                    # 合并行处理
                    $values = array_values($v2);
                    # 写入数据前事件触发
                    self::triggerHook(self::HOOK_WRITE_ROW_BEFORE, [
                        "writer"    => $writer,
                        "row"       => $v,
                        "row_index" => $row_index,
                    ]);

                    $writer->writeSheetRow($sheet['sheet_name'], $values, array_values($merge_style));
                    # 写入数据后事件触发
                    self::triggerHook(self::HOOK_WRITE_ROW_AFTER, [
                        "row"       => $v,//当前行数据
                        "row_index" => $row_index,//当前行索引
                        "writer"    => $writer,//当前写入对象
                    ]);
                }
            }
        }

        $is_tmp = false;
        # 如果文件地址为空则写入临时文件中
        if (!$filePath) {
            $is_tmp = true;
            $filePath = self::getTmpfile();
        }
        if (!$filePath) {
            return 1001;
        }
        $token = '';
        if ($is_tmp) {
            $token = substr($filePath, strlen(self::getTmpDir()) + 1);
        } else {
            # 非临时文件将不返回token
            if (!is_file($filePath)) {
                # 创建目录
                $dirname = dirname($filePath);
                if (!is_dir($dirname)) {
                    FileHelper::createDirectory($dirname, 0775, true);
                }
                if (is_dir($dirname)) {
                    # 创建一个空文件
                    if (!$handle = fopen($filePath, 'w')) {
                        return 1002;
                    } else {
                        fclose($handle);
                    }
                } else {
                    return 1002;
                }
            }
            // $token = $fileName;
        }
        $writer->writeToFile($filePath);
        $excelInfo = [
            "is_tmp"    => $is_tmp,
            'token'     => $token,
            'path'      => $filePath,
            "file_name" => basename($filePath),
        ];

        return $excelInfo;
    }


    /**
     * 导出csv
     *
     * @param array  $sheets 页参数（只可写一个成员）
     *        [
     *              [
     *                     # 列标题 (array)
     *                     "sheet_header"=>[
     *                             [
     *                                  "title"=>"名称",//excel显示的标题
     *                                   "type"=>"string",//字段类型
     *                                   "key"=>"name",//字段名
     *                             ]
     *                      ],
     *                      # 数据回调 (callback)
     *                      "data_fun"=>function($params){
     *                                return table::getData($params['limit'],$params['offset']);
     *                       },
     *                      # 数据总数量(int|callback|false 只导出一次适用于不查询的一次导出)
     *                      "data_count"=>1000,//1000 | function(){return total;}
     *                      # 字段值格式化 (头字段使用数据中没有的字段可通过此方式赋值)
     *                      data_filter"=>[
     *                                "字段1"=>function($row){ return "";},
     *                                "字段2"=>function($row){return "";},
     *                      ],
     *              ]
     *        ]
     *
     * @param string $filePath
     *                                         1. 全路径
     *                                         2. null 写入临时文件返回token
     *                                         3.$isPutPhp=true 时 此处写下载文件名
     * @param int    $pageSize 每次写入数量
     * @param bool   $isPutPhp true: 输入输出流中php://output
     *
     * @return int
     */
    public static function exportCsv(Array $sheets, $filePath = "", $pageSize = 5000, $isPutPhp = false)
    {
        $path = '';
        $sheet = $sheets[0];
        ob_clean();
        setlocale(LC_ALL, 'zh_CN');
        set_time_limit(0);
        # 输入到输出流中
        if ($isPutPhp) {
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="' . ($filePath ? $filePath : 'file') . '.csv"');
            $path = 'php://output';
        }
        # 设置文件地址
        $is_tmp = false;
        if (!$path) {
            $is_tmp = true;
            $path = self::getTmpfile(null, 'csv');
        }
        if (!$path) {
            return 1001;
        }
        $token = '';
        if ($is_tmp) {
            $token = substr($path, strlen(self::getTmpDir()) + 1);
        }
        if (!$fp = fopen($path, 'a')) {
            return 1002;
        }
        fwrite($fp, chr(0xEF) . chr(0xBB) . chr(0xBF));
        # 设置列名
        $titles = array_column($sheet['sheet_header'], 'title');
        if ($titles && is_array($titles)) {
            fputcsv($fp, $titles);
        }
        # 获取数据总数量
        if (isset($sheet['data_count'])) {
            $total_count = (int)is_callable($sheet['data_count']) ? $sheet['data_count']() : $sheet['data_count'];
            $page_count = ceil($total_count / $pageSize);
        }
        # 开始插入数据
        for ($page = 1; (isset($page_count) ? ($page <= $page_count) : true); $page++) {
            $params['limit'] = $pageSize;
            $params['offset'] = ($page - 1) * $pageSize;
            # 最后一页
            if (isset($page_count) && $total_count !== false && $page_count == $page) {
                $params['limit'] = $total_count % $pageSize;
            }
            # 查询对应数据（数据为空则代表最后）
            $data = $sheet['data_fun']($params);
            if (empty($data)) {
                break;
            }
            foreach ($data as $k => $v) {
                # 设置排序 和过滤字段
                $v2 = [];
                foreach ($sheet['sheet_header'] as $header) {
                    $key = $header['key'];
                    $v2[ $key ] = (isset($v[ $key ]) && $v[ $key ] !== null) ? $v[ $key ] : '';
                    # 数据格式化
                    $filter_fun = $sheet['data_filter'][ $key ] ?? null;
                    if ($filter_fun && is_callable($filter_fun)) {
                        $v2[ $key ] = $filter_fun($v);
                    }
                    # 格式按字符串走
                    if ($header['type'] == 'string') {
                        $v2[ $key ] = "\t" . $v2[ $header['key'] ];
                    }
                }
                # 写入数据
                fputcsv($fp, $v2);
            }
            # 刷新缓冲区
            if ($isPutPhp) {
                ob_flush();
                flush();
            }
        }
        fclose($fp);
        !$isPutPhp || die;
        $excelInfo = [
            "is_tmp"    => $is_tmp,
            'token'     => $token,
            'path'      => $path,
            "file_name" => basename($path),
        ];
        ob_clean();

        return $excelInfo;
    }

    /**
     * 绑定事件回调
     *
     * @param          $hookName
     * @param callable $fun
     */
    public static function bindHook($hookName, callable $fun)
    {
        self::$hooks[ $hookName ][] = $fun;
    }

    /**
     * 触发事件回调
     *
     * @param       $hookName
     * @param array $param
     */
    public static function triggerHook($hookName, $param = [])
    {
        $funs = self::$hooks[ $hookName ] ?? [];
        foreach ($funs as $fun) {
            call_user_func_array($fun, $param);
        }
    }

    /**
     * 删除已经失效的生成的excel文件临时文件
     *
     * @return array 返回删除成功的文件名集合
     */
    public static function delFileByExpire()
    {
        $cTime = time();
        $tempdir = self::getTmpDir();
        # 删除临时目录下生成的失效excel
        $files = scandir($tempdir);
        $files = array_filter(array_map(function ($file) use ($tempdir, $cTime) {
            $full_path = $tempdir . "/" . $file;
            if (!in_array($file, [".", ".."]) && !is_dir($full_path)) {
                $fileCtime = filectime($full_path);
                if ($fileCtime && ($cTime - $fileCtime >= self::$expireTime)) {
                    if (@unlink($full_path)) {
                        return $file;
                    }
                }
            }
        }, $files));

        return $files;
    }

    /**
     * 根据token获取extension
     *
     * @param $token
     */
    public static function getExtensionByToken($token)
    {
        list($extension, $filename) = explode('_', $token);

        return $extension;
    }

    /**
     * @param      $showFilename  显示文件名
     * @param      $tmpFilename   对应文件路径（全路径/token）
     * @param bool $isDel 是否下载完删除对应文件
     * @param null $extension 设置下载拓展名(token时自动根据名字获取,全路径可指定(xlsx/csv))
     *
     * @return int
     */
    public static function downExcel($showFilename, $tmpFilename, $isDel = true, $extension = null)
    {
        # 删除已经失效的文件
        self::delFileByExpire();
        ob_end_clean();

        if (!is_file($tmpFilename)) {
            $extension = self::getExtensionByToken($tmpFilename);
            $path = self::getPathByToken($tmpFilename);
        } else {
            $path = $tmpFilename;
        }
        if (!file_exists($path)) {
            return 1002;
        }

        header("Pragma: public");
        header("Expires: 0");
        header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
        header("Content-Type:application/force-download");
        header("Content-Type:application/vnd.ms-excel");
        header("Content-Type:application/octet-stream");
        header("Content-Type:application/download");
        header('Content-Disposition:attachment;filename=' . $showFilename . '.' . ($extension ?? 'xlsx'));
        header("Content-Transfer-Encoding:binary");

        /*        $fp = fopen($path, 'rb');
                if ($fp) {
                    if (!fpassthru($fp)) {
                        return 1004;
                    }
                }
                @unlink($path);*/
        # 读取文件
        if (!readfile($path)) {
            if ($isDel) {
                @unlink($path);
            }

            return 1004;
        }
        # 删除文件
        if ($isDel) {
            @unlink($path);
        }
        exit();
    }

    /**
     * 根据参数 token 获取文件地址
     *
     * @param $fileName
     * @param $token
     *
     * @return string
     * @throws \Exception
     */
    public static function getUrl($fileName, $token)
    {
        if (!$token) {
            throw new \Exception("非临时目录时将不返回url,请自行根据自定义路径拼接");
        }

        return self::$excelDownUrl . "?filename={$fileName}&token={$token}";
    }

    /**
     * 在指定目录申请一个临时文件
     *
     * @param null   $tmpDir 指定目录
     * @param string $extension 拓展名标识(xlsx/csv)
     *
     * @return bool|string
     */
    public static function getTmpfile($tmpDir = null, $extension = 'xlsx')
    {
        if (!$tmpDir) {
            $tmpDir = self::getTmpDir();
        }
        $tmpPath = tempnam($tmpDir, '');
        $baseName = basename($tmpPath);
        $filePath = $tmpDir . '/' . $extension . '_' . $baseName;
        rename($tmpPath, $filePath);

        return $filePath;
    }

    /**
     * 获取系统临时目录
     *
     * @return string
     */
    public static function getTmpDir()
    {
        # 根目录申请一个临时目录
        $dir = \Yii::getAlias("@webroot") . self::$temDir;
        if (!is_dir($dir)) {
            mkdir($dir, 0755, true);
        }

        return $dir;

        //  return sys_get_temp_dir();
    }

    /**
     * 根据token 获取完整地址
     *
     * @param $token
     *
     * @return string
     */
    public static function getPathByToken($token)
    {
        return self::getTmpDir() . '/' . $token;
    }
}
