<?php

namespace Alastor141\Bitrix\Google\Sheets\Model;

use \Alastor141\Google\Model\Row as BaseRow;
use Google\Service\Sheets\CellData;
use Google\Service\Sheets\CellFormat;
use Google\Service\Sheets\ExtendedValue;

class Row extends BaseRow
{
    protected function getCellData($value, $key)
    {
        $fields = $this->getFields();
        $format = ['wrapStrategy' => 'WRAP'];
        if ($this->formatCells[$key]) {
            $this->setFields($fields.',userEnteredFormat.backgroundColor');
            $format = array_merge($format, $this->formatCells[$key]);
        }

        if ($this->textFormatRuns[$key]) {
            $this->setFields($fields.',textFormatRuns');
        }

        $cellDataParams = [
            'userEnteredValue' => new ExtendedValue([
                'stringValue' => $value
            ]),
            'userEnteredFormat' => new CellFormat($format),
            'textFormatRuns' => $this->textFormatRuns[$key] ? $this->textFormatRuns[$key] : []
        ];

        return new CellData($cellDataParams);
    }

    public function getMetadataKey(): string
    {
        return 'object';
    }
}