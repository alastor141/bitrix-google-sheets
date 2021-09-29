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
  bitrix.export.object:
    class: Alastor141\Bitrix\Google\Sheets\Export\Lists\Export
    arguments: ['@service.google.sheets', '@bitrix.export.object.provider', '%table%']
    public: true
```