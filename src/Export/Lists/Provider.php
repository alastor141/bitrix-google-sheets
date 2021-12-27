<?php


namespace Alastor141\Bitrix\Google\Sheets\Export\Lists;

use Alastor141\Bitrix\Google\Sheets\Model\Row;
use Alastor141\Google\Collection\Rows;
use Alastor141\Service\Sheets\Provider as BaseProvider;
use CCurrencyLang;
use Google\Service\Sheets\AddSheetRequest;
use Google\Service\Sheets\Color;
use Google\Service\Sheets\CreateDeveloperMetadataRequest as CreateDeveloperMetadataRequestAlias;
use Google\Service\Sheets\DeveloperMetadata;
use Google\Service\Sheets\DeveloperMetadataLocation;
use Google\Service\Sheets\GridProperties;
use Google\Service\Sheets\GridRange;
use Google\Service\Sheets\MergeCellsRequest;
use Google\Service\Sheets\Request;
use Google\Service\Sheets\SheetProperties;
use Google\Service\Sheets\TextFormat;
use Google\Service\Sheets\TextFormatRun;
use Google\Service\Sheets\UpdateSheetPropertiesRequest;


/**
 * Class ProviderObject
 * @package App\Service\Bitrix\Lists
 */
class Provider extends BaseProvider
{
    const STATUSES = [
        971 => [
            'name' => 'sale',
            'rgba' => [
                "red" => 0.91,
                "green" => 0.96,
                "blue" => 0.83,
                "alpha" => 1
            ]
        ],
        972 => [
            'name' => 'after_reserved',
            'rgba' => [
                "red" => 0.90,
                "green" => 0.67,
                "blue" => 0.64,
                "alpha" => 1
            ]
        ],
        973 => [
            'name' => 'registration',
            'rgba' => [
                "red" => 0.91,
                "green" => 0.96,
                "blue" => 0.83,
                "alpha" => 1
            ]
        ],
        974 => [
            'name' => 'reserved',
            'rgba' => [
                "red" => 1,
                "green" => 0.93,
                "blue" => 0.84,
                "alpha" => 1
            ]
        ],
    ];
    
    public function execution() : Rows
    {
		$filter = $this->config['filter'];

        $select = [
            'ID',
            'PROPERTY_HOUSE_ID',
            'PROPERTY_BUILDING_ID',
            'PROPERTY_SECTION_ID',
            'PROPERTY_HOUSE_ID.NAME',
            'PROPERTY_BUILDING_ID.NAME',
            'PROPERTY_SECTION_ID.NAME',
            'PROPERTY_PRICE_PER_SQM_VALUE',
            'PROPERTY_APARTMENT_NUMBER',
            'PROPERTY_AREA_VALUE',
            'PROPERTY_ROOMS',
            'PROPERTY_FLOOR',
            'PROPERTY_PRICE_TOTAL_VALUE',
            'PROPERTY_STATUS_ID'
        ];

        $dbResult = \CIBlockElement::GetList(['PROPERTY_FLOOR' => 'DESC', 'PROPERTY_APARTMENT_NUMBER' => 'ASC'], $filter, false, false, $select);


        $headers = [];
        $houseMetadataIds = [];
        $buildingMetadataIds = [];
        $sectionMetadataIds = [];
        $buildings = [];
        while ($row = $dbResult->Fetch()) {
            $buildingId = (int)$row['PROPERTY_BUILDING_ID_VALUE'];
            $houseId = (int)$row['PROPERTY_HOUSE_ID_VALUE'];
            $sectionId = (int)$row['PROPERTY_SECTION_ID_VALUE'];
            $floorId = (int)$row['PROPERTY_FLOOR_VALUE'];

            $houseMetadataIds[$houseId] = $houseId;

            $buildings[$buildingId]['name'] = $row['PROPERTY_BUILDING_ID_NAME'];
            $buildings[$buildingId]['houses'][$houseId]['name'] = $row['PROPERTY_HOUSE_ID_NAME'];
            $buildings[$buildingId]['houses'][$houseId]['sections'][$sectionId]['name'] = $row['PROPERTY_SECTION_ID_NAME'];
            $buildings[$buildingId]['houses'][$houseId]['sections'][$sectionId]['floors'][$floorId][] = $row;
        }

        $sections = [];
        $sheetTitle = [];
        $endColumnIndex = 0;
        foreach ($buildings as $buildingId => $building) {
            foreach ($building['houses'] as $houseId => $house) {
                $sheetTitle[$houseId] = $building['name'].' '.$house['name'];
                $buildingKey = ($buildingId+$houseId)*2;
                $buildingMetadataIds[$buildingKey] = $buildingKey;
				//Убрана строка названия ЖК
				//$headers[$houseId]['values'][$buildingKey] = [$building['name']];
                $headers[$houseId]['values'][$houseId] = [$house['name']];
                foreach ($house['sections'] as $sectionId => $section) {
                    $sectionMetadataIds[] = $sectionId;
                    $sections[$houseId]['values'][$sectionId]['values'][$sectionId] = [
                         $section['name']
                    ];

                    foreach ($section['floors'] as $floorId => $floor) {
                        $cells = [];
                        $room = [];
                        $formatCells = [];
                        $textFormatRuns = [];
                        $cells[$sectionId.abs($floorId)] = (string)$floorId;
                        $formatCells[$sectionId.abs($floorId)] = [
                            'backgroundColor' => new Color([
                                "red" => 0.82,
                                "green" => 0.88,
                                "blue" => 0.89,
                                "alpha" => 1
                            ])
                        ];
                        $room[] = 'Этаж';
                        foreach ($floor as $item) {
                            $statusCode = false;
                            if ($item['PROPERTY_STATUS_ID_VALUE']) {
                                $status = self::STATUSES[$item['PROPERTY_STATUS_ID_VALUE']];
                                if ($status && $status['rgba']) {
                                    $statusCode = $status['name'];
                                    $formatCells[$item['ID']]['backgroundColor'] = new Color($status['rgba']);
                                }
                            }

                            $room[] = $item['PROPERTY_ROOMS_VALUE']." комн. кв.";
                            $price = $statusCode != 'sale' && $statusCode != 'registration' ? CCurrencyLang::CurrencyFormat($item['PROPERTY_PRICE_TOTAL_VALUE_VALUE'], 'RUB') : '';
                            $rooms = $item['PROPERTY_ROOMS_VALUE']." комн. кв.\n";
                            $startIndexApartment =  mb_strlen($rooms) - 1;
                            $text = $rooms."№".$item['PROPERTY_APARTMENT_NUMBER_VALUE']."\n";
                            $startIndexArea = mb_strlen($text) - 1;
                            $area = $item['PROPERTY_AREA_VALUE_VALUE'].'м²'."\n";
                            $text = $text.$area;
                            $cells[$item['ID']] = $text.$price;
                            $startIndexPrice = mb_strlen($text) - 1;

                            $textFormatRuns[$item['ID']][] = new TextFormatRun([
                                'startIndex' => $startIndexApartment,
                                'format' => new TextFormat([
                                    'italic' => true
                                ])
                            ]);

                            $textFormatRuns[$item['ID']][] = new TextFormatRun([
                                'startIndex' => $startIndexArea,
                                'format' => new TextFormat([
                                    'bold' => true
                                ])
                            ]);

                            $textFormatRuns[$item['ID']][] = new TextFormatRun([
                                'startIndex' => $startIndexPrice,
                                'format' => new TextFormat([
                                    'bold' => true,
                                    'italic' => true
                                ])
                            ]);

                        }

                        $count = count($cells);
                        if ($endColumnIndex < $count) {
                            $endColumnIndex = $count;
                        }

                        $roomId = ($sectionId+1)*4;
						//Убрана строка количества комнат
						//$sections[$houseId]['values'][$sectionId]['values'][$roomId] = $room;
                        $sections[$houseId]['values'][$sectionId]['values'][$sectionId.abs($floorId)] = $cells;
                        $sections[$houseId]['values'][$sectionId]['format'][$sectionId.abs($floorId)] = $formatCells;
                        $sections[$houseId]['values'][$sectionId]['textFormatRuns'][$sectionId.abs($floorId)] = $textFormatRuns;
                        $sections[$houseId]['columnCount'] = $endColumnIndex;
                    }
                }
            }
        }

        $collection = new Rows();

        $sheetIndex = 0;
        $sheets = $this->spreadsheets->getSheets();

        foreach ($sections as $houseId => $section) {
            $uid = ($houseId+1)*6;
            $sheet = false;
            foreach ($sheets as $s) {
                $developerMetadata = $s->getDeveloperMetadata();
                foreach ($developerMetadata as $item) {
                    if ($item->getMetadataId() == $uid) {
                        $sheet = $s;
                    }
                }
            }  

            if ($sheet) {
                $sheetId = $sheet->getProperties()->getSheetId();
                $updateSheetProperties = new UpdateSheetPropertiesRequest([
                    'fields' => 'title',
                    'properties' => new SheetProperties([
                        'sheetId' => $sheetId,
						'title' => $sheetTitle[$houseId]
                    ])
                ]);

                $request = new Request();
                $request->setUpdateSheetProperties($updateSheetProperties);
				$this->beforeAdditionalRequest[] = $request;

            } else {
                $sheetId = $uid;
                $addSheetRequest = new AddSheetRequest([
                    'properties' => new SheetProperties([
                        'sheetId' => $sheetId,
                        'title' => $sheetTitle[$houseId],
                        'gridProperties' => new GridProperties([
                            'columnCount' => $section['columnCount']
                        ])
                    ])
                ]);

                $request = new Request();
                $request->setAddSheet($addSheetRequest);
                $this->afterAdditionalRequest[] = $request;
            }

            $sheetDeveloperMetadata = $this->developerMetadata[$uid];
            if ($sheetDeveloperMetadata) {
                $sheetId = $sheetDeveloperMetadata->getLocation()->getSheetId();
            } else {
                $developerMetadataParams = [
                    'developerMetadata' => new DeveloperMetadata([
                        'metadataId' => $uid,
                        'metadataKey' => $this->getMetadataKey(),
                        'location' => new DeveloperMetadataLocation([
                            'sheetId' => $sheetId
                        ]),
                        'visibility' => 'DOCUMENT'
                    ])
                ];

                $request = new Request();
                $request->setCreateDeveloperMetadata(new CreateDeveloperMetadataRequestAlias($developerMetadataParams));
                $this->beforeAdditionalRequest[] = $request;
            }

            $startRowIndex = 0;

            foreach ($section['values'] as $values) {

                if ($startRowIndex === 0) {
					//Убрана строка литера
					//$values['values'] = $headers[$houseId]['values'] + $values['values'];
                }

                foreach ($values['values'] as $uid => $value) {
					//Убрана строка подъезда
					if (in_array($uid, $sectionMetadataIds)) {

						continue;
					}

                    $developerMetadata = $this->developerMetadata[$uid];
                    if ($developerMetadata) {
                        $startRowIndex = $developerMetadata->getLocation()->getDimensionRange()->getStartIndex();
                    }

                    if (in_array($uid, array_merge($houseMetadataIds,$buildingMetadataIds,$sectionMetadataIds))) {
                        $mergeRequest = new Request();
                        $mergeRequest->setMergeCells(new MergeCellsRequest([
                            'mergeType' => 'MERGE_ROWS',
                            'range' => new GridRange([
                                'sheetId' => $sheetId,
                                'startColumnIndex' => 0,
                                'endColumnIndex' => $endColumnIndex,
                                'startRowIndex' => $startRowIndex,
                                'endRowIndex' => $startRowIndex + 1,
                            ])
                        ]));

						$this->beforeAdditionalRequest[] = $mergeRequest;
                    }

                    $params = [
                        'sheetId' => $sheetId,
                        'startRowIndex' => $startRowIndex,
                        'endRowIndex' => ++$startRowIndex,
                        'startColumnIndex' => 0,
                        'metadataKey' => $this->getMetadataKey()
                    ];

                    $rowData = new Row($uid, $value, $values['format'][$uid], $values['textFormatRuns'][$uid]);
                    $rowData->setDeveloperMetadata($developerMetadata);
                    $rowData->setParams($params);
                    $collection[] = $rowData;
                }

                $startRowIndex++;
            }
            $sheetIndex++;
        }

        return $collection;
    }

    public function getMetadataKey(): string
    {
        return 'object';
    }
}
