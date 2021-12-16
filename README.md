# Google Sheets

Пример конфигурационного файла

```yaml
parameters:
  table: '1ayf2NMi5A8VUaWrgDiLYs48lanJ-q2Mk6k6STJIk8sI'
  google.client.config:
    access_type: offline
    credentials: '%kernel.project_dir%/google.key.json'
    scopes:
      - !php/const Google\Service\Sheets::SPREADSHEETS
  bitrix.export.config:
      filter:
        IBLOCK_ID: 42
        ACTIVE: 'Y'

services:
  google.client:
    class: Google\Client
    arguments: ['%google.client.config%']
  google.sheets:
    class: Google\Service\Sheets
    arguments: ['@google.client']
  service.google.sheets:
    class: Alastor141\Google\Api\Sheets
    arguments: ['@google.sheets']
  bitrix.export.object.provider:
    class: Alastor141\Bitrix\Google\Sheets\Export\Lists\ProviderObject
    arguments: [ '%bitrix.export.config%' ]
  bitrix.export.object:
    class: Alastor141\Bitrix\Google\Sheets\Export\Lists\Export
    arguments: ['@service.google.sheets', '@bitrix.export.object.provider', '%table%']
    public: true
```

 Реализовав класс провайдера достаточно зарегистрировать сервис передав в него аргумент в виде массива конфигурации для самого провайдера по которому он например сделает выборку данных

```yaml
bitrix.export.object.provider:
    class: Alastor141\Bitrix\Google\Sheets\Export\Lists\ProviderObject
    arguments: [ '%bitrix.export.config%' ]
```

После этого в PHP 
```php
use Alastor141\Bitrix\Google\Sheets\Kernel;

try {
    $kernel = new Kernel();
    $sheets = $kernel->getContainer()->get('bitrix.export.object');
    $sheets->export();

	$sheets = $kernel->getContainer()->get('bitrix.export.gray.object');
    $sheets->export();
} catch (\Exception $error) {
    dump($error);
}
```
