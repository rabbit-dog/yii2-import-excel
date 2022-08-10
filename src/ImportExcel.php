<?php
/**
 * 使用说明
 * User: xing.chen
 * Date: 2018/10/6
 * Time: 21:10
 *
 *
 *
ImportExcel::init(''文件路径, $rowsSet, 1)
->valueMap($valueMap)
->valueMapDefault($valueMapDefault)
->formatFields(['birthday' => 'date', 'activistApplyDate' => 'date'])
->run(function($data) use ($type) {

// 增加/修改
if ($type == 'update') $m = PartyMember::findCard($data['cardNumber']) ?: new PartyMember();
else $m = new PartyMember();

$m->load($data, '');
if (!$m->save()) throw new ModelException($m);
});
 */

namespace xing\yiiImportExcel;

use xing\yiiImportExcel\cache\CacheFactory;
use Yii;
use PhpOffice\PhpSpreadsheet\Reader\Xls;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

/**
 * Class ImportExcel
 * @property string $file 文件路径
 * @property array $rowsSet  列对应的字段名
 * @property int $startRow 从第几行开始处理
 * @property array $valueMap 选项值配置
 * @property array $valueMapDefault 选项值不存在时的默认值
 * @property array $formatFields 格式设置，如日期需要设置，否则读取到值 会有问题
 * @property callable $transactionRollBack 事务回滚匿名函数
 * @property callable $checkFunction 事务回滚匿名函数
 * @property array $uniqueFields 唯一字段，需要检查表格中是否存在重复
 * @property bool $transaction 是否开启事务
 * @package xing\helper\yii\xml
 */
class ImportExcel
{

    protected $file;

    protected $rowsSet;

    protected $valueMap;

    protected $valueMapDefault;

    protected $formatFields = [];

    // 当前行
    public static $currentRow = 0;
    public static $currentCol = 0;

    protected $transactionRollBack;

    protected $uniqueFields = [];

    private $checkUniqueText = [];
    private $transaction = false;
    public static $defaultDate = '1000-01-01';


    /**
     * 初始化
     * @param string $file 文件所在路径
     * @param array $rowsSet 列对应字段
     * @param int $startRow 从第几行开始处理
     * @return ImportExcel
     */
    public static function init($file, $rowsSet, $startRow = 0)
    {
        $class = new self;
        $class->file = $file;
        $class->rowsSet = $rowsSet;
        $class->startRow = $startRow;
        return $class;
    }

    public function valueMap($valueMap)
    {
        $this->valueMap = $valueMap;
        return $this;
    }

    public function valueMapDefault($valueMapDefault)
    {
        $this->valueMapDefault = $valueMapDefault;
        return $this;
    }

    public function formatFields($formatFields)
    {
        $this->formatFields = $formatFields;
        return $this;
    }

    public function setTransactionRollBack($transactionRollBack)
    {
        $this->transactionRollBack = $transactionRollBack;
        return $this;
    }

    public function setUnique($fields)
    {
        $this->uniqueFields = $fields;
        return $this;
    }

    public function setTransaction($bool)
    {
        $this->transaction = $bool;
        return $this;
    }

    public function setCheck($function)
    {
        $this->checkFunction = $function;
        return $this;
    }

    /**
     * 执行导入
     * @param callable $saveFunction 回调:保存处理程序
     * @return int
     * @throws \Exception
     */
    public function run(callable $saveFunction)
    {

        set_time_limit(0);

        /*$objReader = \PHPExcel_IOFactory::createReader('Excel2007');
        if(!$objReader->canRead($this->file)){
            $reader = new Xls();
        } else {
            $reader = new Xlsx();
        }*/

        // 设置缓存,以处理大型表格
        $cache = CacheFactory::getInstance('FileCache');
        \PhpOffice\PhpSpreadsheet\Settings::setCache($cache);

//        $reader->setReadDataOnly(true);
//        $spreadsheet = $reader->load($this->file);
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($this->file);

        $worksheet = $spreadsheet->getActiveSheet();

        // 表格序列对应的字段名
        $rowsSet = $this->rowsSet;

        // 键值设置
        $valueMap = $this->valueMap;

        // 键值默认设置，如果没有设置的话，将会抛出错误
        $valueMapDefault = $this->valueMapDefault;


        if ($this->transaction) $transaction = Yii::$app->db->beginTransaction();
        $nnn = 0;
        try {
            // 检查数据（考虑内存，不另存数据）
            foreach ($worksheet->getRowIterator() as $k => $row) {
                static::$currentRow = $row->getRowIndex();
                // 跳过
                if ($this->startRow > $k ) continue;

                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(FALSE);
                $data = [];
                // 检查当前一行
                foreach ($cellIterator as $col => $cell) {

                    static::$currentCol = $cell->getColumn();
                    $fieldName = $rowsSet[$col] ?? null;
                    if (empty($fieldName)) continue;

                    $value = $cell->getCalculatedValue() ?: $cell->getValue();

                    // 如果是选项
                    // 如果值为空，并且没有默认值的设置
                    if (is_null($value) && !isset($valueMapDefault[$fieldName])) {
                        throw new \Exception("第{$k}行{$col}列格式错误，值不能为空。", 10000);
                    }
                    // 重复检查
                    if (!is_null($value) && in_array($fieldName, $this->uniqueFields)) {
                        if (!isset($this->checkUniqueText[$fieldName])) $this->checkUniqueText[$fieldName] = '|';
                        if (stripos($this->checkUniqueText[$fieldName], '|' . $value . '|') !== false) {
                            throw new \Exception("第{$k}行{$col}的值重复");
                        }
                        $this->checkUniqueText[$fieldName] .= $value . '|';
                    }
                    if (!empty($this->checkFunction)) {
                        if (isset($this->formatFields[$fieldName])) {
                            $value = $this->format($fieldName, $value);
                        }
                        $data[$fieldName] = trim((string) $value);
                    }
                }
                if (!empty($this->checkFunction)) ($this->checkFunction)($data);
            }

            // 删除变量，释放内存
            unset($this->checkUniqueText);
            $this->checkUniqueText = null;

            foreach ($worksheet->getRowIterator() as $k => $row) {
                static::$currentRow = $row->getRowIndex();

                // 跳过
                if ($this->startRow > $k ) continue;

                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(FALSE);

                $data = [];
                $nullNumber = 0;

                // 空行检查
                foreach ($cellIterator as $col => $cell) {

                    $value = $cell->getCalculatedValue() ?: $cell->getValue();
                    // 空值统计
                    if (is_null($value)) $nullNumber ++;
                }

                // 循环赋值表格一行的数据
                if ($nullNumber < count($this->rowsSet)) foreach ($cellIterator as $col => $cell) {

                    static::$currentCol = $cell->getColumn();
                    $fieldName = $rowsSet[$col] ?? null;
                    if (empty($fieldName)) continue;
                    $value = $cell->getCalculatedValue() ?: $cell->getValue();


                    // 如果值为空，并且有默认值的设置，则设置为默认值，否则如果值存在，则值换为键名
                    if (is_null($value) && isset($valueMapDefault[$fieldName])) {
                        $value = $valueMapDefault[$fieldName];
                    } else if ((!is_null($value)) && isset($valueMap[$fieldName]) && is_array($valueMap[$fieldName])) {
                        $value = array_keys($valueMap[$fieldName], $value)[0] ?? '';
                    }
                    // 格式处理
                    if (isset($this->formatFields[$fieldName])) $value = $this->format($fieldName, $value);

                    $data[$fieldName] = trim((string) $value);
                }


                // 执行保存程序
                $saveFunction($data);
                ++ $nnn;
            }

            if ($this->transaction) $transaction->commit();
            return $nnn;
        } catch (\Exception $e) {
            if ($this->transaction) {
                $transaction->rollBack();
                if (!empty($this->transactionRollBack)) ($this->transactionRollBack)($e);
            }
//            throw $e;
            $msg = '操作失败，'. ($this->transaction ? '本次操作全部取消，' : '') .'请修正后重新上传';
            throw new \Exception('<h3>' . $msg . '</h3>
<p>错误信息：'.$e->getMessage().'</p>
<p>表格处理至第 '. static::$currentRow .' 行 ' . static::$currentCol . ' 列</p>');
        }
    }

    /**
     * 格式处理
     * @param string $fieldName
     * @param $value
     * @return mixed
     */
    protected function format(string $fieldName, $value)
    {
        $format = $this->formatFields[$fieldName] ?? null;
            // 格式处理
        switch ($format) {
            case 'date':
                $value = $value ? date('Y-m-d', \PhpOffice\PhpSpreadsheet\Shared\Date::excelToTimestamp($value)) : static::$defaultDate;
                break;
            case 'date:int':
                $value = $value ? \PhpOffice\PhpSpreadsheet\Shared\Date::excelToTimestamp($value) : 0;
                break;
        }
        return $value;
    }
}
