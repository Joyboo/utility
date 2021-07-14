<?php

use \Vtiful\Kernel\Excel;

/**
 * Create by Joyboo 2021-06-29
 *
 * XlsWriter官方文档：https://xlswriter-docs.viest.me/zh-cn/kuai-su-shang-shou/reader
 * 来自easyswoole官方的推荐 http://www.easyswoole.com/OpenSource/xlsWriter.html
 *
 * 支持游标模式的超大文件读取，内存消耗不到1MB，需安装XlsWriter.so扩展
 *
 * Class XlsWriter
 */
class XlsWriter
{
    const TYPE_INT = Excel::TYPE_INT;
    const TYPE_STRING = Excel::TYPE_STRING;
    const TYPE_DOUBLE = Excel::TYPE_DOUBLE;
    const TYPE_TIMESTAMP = Excel::TYPE_TIMESTAMP;

    protected $excel = null;

    protected $offset = 0;

    protected $setType = [];

    public function __construct($path = '')
    {
        if (!is_dir($path))
        {
//            mkdir($path, 0777, true);
            throw new \Exception('没有这个目录：' . $path);
        }

        $config = ['path' => $path];
        $this->excel = new Excel($config);
    }

    /**
     * 设置读取参数
     * @param int $offset 偏移量，传1会丢弃第一行，传2会丢弃第一行和第二行 ...
     * @param array $setType 列单元格数据类型，从0开始 [2 => \XlsWriter::TYPE_TIMESTAMP]表示第三列的单元格是时间类型
     * @return $this
     */
    public function setReader(int $offset = 0,array $setType = [])
    {
        $this->offset = $offset;
        $this->setType = $setType;
        return $this;
    }

    /**
     * 游标模式
     * @param $file
     * @param callable $callback function(int $row, int $cell, $data)
     */
    public function cursorFile($file, callable $callback)
    {
        $sheetList = $this->excel->openFile($file)->sheetList();

        foreach ($sheetList as $sheetName)
        {
            $sheet = $this->excel->openSheet($sheetName);
            if ($this->offset > 0)
            {
                $sheet->setSkipRows($this->offset);
            }
            if ($this->setType)
            {
                $sheet->setType($this->setType);
            }
            $sheet->nextCellCallback($callback);
        }
    }

    /**
     * 全量模式
     * @param $file
     */
    public function readFile($file)
    {
        $sheetList = $this->excel->openFile($file)->sheetList();

        $result = [];
        foreach ($sheetList as $sheetName)
        {
            $sheet = $this->excel->openSheet($sheetName);
            if ($this->offset > 0)
            {
                $sheet->setSkipRows($this->offset);
            }
            if ($this->setType)
            {
                $sheet->setType($this->setType);
            }
            $sheetData = $sheet->getSheetData();
            $result = array_merge($result, $sheetData);
            unset($sheetData, $sheet);
        }

        return $result;
    }

    public function ouputFile($file, $data = [], $header = [])
    {
        $object = $this->excel->fileName($file);
        if ($header)
        {
            $object->header($header);
        }
        $object->data($data)->output();
    }
}
