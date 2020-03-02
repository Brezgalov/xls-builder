<?php

namespace Brezgalov\XlsBuilder;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;

abstract class BaseTableTemplateBuilder extends BaseTemplateBuilder
{
    /**
     * @var array
     */
    public $tableData = [];

    /**
     * Автоматом проставлять границу таблицы
     * @var bool
     */
    public $autoBorder = true;

    /**
     * Масcив связывающий данные с таблицей вида ['my_column' => ['A']]
     * @return array
     */
    public abstract function getTablePropertiesMap();

    /**
     * Откуда начинаются данные таблицы
     * @return int
     */
    public abstract function getTableDataStartRow();

    /**
     * Левый край таблицы, по нему ставится граница
     * @return string
     */
    public abstract function getTableLeftColumnName();

    /**
     * Правый край таблицы, по нему ставится граница
     * @return string
     */
    public abstract function getTableRightColumnName();

    /**
     * @return string
     */
    public function getTableBorderWidth()
    {
        return Border::BORDER_THIN;
    }

    /**
     * @param array $data
     */
    public function setTableData(array $data)
    {
        $this->tableData = $data;
    }

    /**
     * Puts table to spreadsheet
     * @param Spreadsheet $spreadsheet
     * @return int
     */
    protected function putTable(Spreadsheet $spreadsheet)
    {
        $page = $spreadsheet->getActiveSheet();
        $nextRow = $this->getTableDataStartRow();
        $tableFieldsMap = $this->getTablePropertiesMap();

        $startColumn = $this->getTableLeftColumnName();
        $endColumn = $this->getTableRightColumnName();

        foreach ($this->tableData as $tableRow) {
            foreach ($tableFieldsMap as $field => $columns) {
                foreach ($columns as $column) {
                    if (array_key_exists($field, $tableRow)) {
                        $page->setCellValue("{$column}{$nextRow}", $tableRow[$field]);
                    }
                }
            }

            if ($this->autoBorder) {
                $page->getStyle("{$startColumn}{$nextRow}:{$endColumn}{$nextRow}")->applyFromArray([
                    'borders' => [
                        'allBorders' => [
                            'borderStyle'   => $this->getTableBorderWidth(),
                            'color'         => ['argb' => '00000000'],
                        ],
                    ],
                ]);
            }

            $nextRow += 1;
        }

        return $nextRow;
    }

    /**
     * @return bool|\PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public function buildFile()
    {
        $spreadsheet = parent::buildFile();

        if (!empty($this->tableData)) {
            $this->putTable($spreadsheet);
        }

        return $spreadsheet;
    }
}