---
name: Microsoft Graph
x-slug: microsoft-graph
description: 'Microsoft Graph exposes multiple APIs from Office 365 and other Microsoft
  cloud services through a single endpoint: https://graph.microsoft.com. Microsoft
  Graph simplifies queries that would otherwise be more complex. You can use Microsoft
  Graph to: Access data from multiple Microsoft cloud services, including Azure Active
  Directory, Exchange Online as part of Office 365, SharePoint, OneDrive, OneNote,
  and Planner. Navigate between entities and relationships. Access intelligence and
  insights from the Microsoft cloud (for commercial users).'
image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
x-kinRank: "10"
x-alexaRank: "0"
tags: Worksheets
created: "2018-08-28"
modified: "2018-08-28"
url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/apis.md
specificationVersion: "0.14"
apis:
- name: Microsoft Graph API - List Worksheets
  x-api-slug: workbookworksheets-get
  description: List worksheets Retrieve a list of worksheet objects.
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheets-get-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheets-get-openapi.md
- name: Microsoft Graph API - Worksheet Cell
  x-api-slug: workbookworksheetsltidnamegtcellrowltrowgtcolumnltcolumngt-get
  description: 'Worksheet: Cell Gets the range object containing the single cell based
    on row and column numbers. The cell can be outside the bounds of its parent range,
    so long as it''s stays within the worksheet grid.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtcellrowltrowgtcolumnltcolumngt-get-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtcellrowltrowgtcolumnltcolumngt-get-openapi.md
- name: Microsoft Graph API - Worksheet Delete
  x-api-slug: workbookworksheetsltidnamegtdelete-post
  description: 'Worksheet: delete Deletes the worksheet from the workbook.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtdelete-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtdelete-post-openapi.md
- name: Microsoft Graph API - Get Worksheet
  x-api-slug: workbookworksheetsltidnamegt-get
  description: Get Worksheet Retrieve the properties and relationships of worksheet
    object.
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegt-get-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegt-get-openapi.md
- name: Microsoft Graph API - Worksheet Range
  x-api-slug: workbookworksheetsltidnamegtrange-post
  description: 'Worksheet: Range Gets the range object specified by the address or
    name.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtrange-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtrange-post-openapi.md
- name: Microsoft Graph API - Update Worksheet
  x-api-slug: workbookworksheetsltidnamegt-patch
  description: Update worksheet Update the properties of worksheet object.
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegt-patch-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegt-patch-openapi.md
- name: Microsoft Graph API - Worksheet Used Range
  x-api-slug: workbookworksheetsltidnamegtusedrange-post
  description: 'Worksheet: UsedRange The used range is the smallest range that encompasses
    any cells that have a value or formatting assigned to them. If the worksheet is
    blank, this function will return the top left cell.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtusedrange-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtusedrange-post-openapi.md
- name: Microsoft Graph API - Worksheet Collection Add
  x-api-slug: workbookworksheets-post
  description: 'WorksheetCollection: add Adds a new worksheet to the workbook. The
    worksheet will be added at the end of existing worksheets. If you wish to activate
    the newly added worksheet, call ".activate() on it.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheets-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheets-post-openapi.md
- name: Microsoft Graph API - Get Worksheet Protection
  x-api-slug: workbookworksheetsltidnamegtprotection-get
  description: Get WorksheetProtection Retrieve the properties and relationships of
    worksheetprotection object.
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotection-get-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotection-get-openapi.md
- name: Microsoft Graph API - Worksheet Protection Protect
  x-api-slug: workbookworksheetsltidnamegtprotectionprotect-post
  description: 'WorksheetProtection: protect Protect a worksheet. It throws if the
    worksheet has been protected.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotectionprotect-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotectionprotect-post-openapi.md
- name: Microsoft Graph API - Worksheet Protection Unprotect
  x-api-slug: workbookworksheetsltidnamegtprotectionunprotect-post
  description: 'WorksheetProtection: unprotect Unprotect a worksheet'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotectionunprotect-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotectionunprotect-post-openapi.md
- name: Microsoft Graph API - Worksheet Cell
  x-api-slug: workbookworksheetsltidnamegtcellrowltrowgtcolumnltcolumngt-get
  description: 'Worksheet: Cell Gets the range object containing the single cell based
    on row and column numbers. The cell can be outside the bounds of its parent range,
    so long as it''s stays within the worksheet grid.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtcellrowltrowgtcolumnltcolumngt-get-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtcellrowltrowgtcolumnltcolumngt-get-openapi.md
- name: Microsoft Graph API - Worksheet Delete
  x-api-slug: workbookworksheetsltidnamegtdelete-post
  description: 'Worksheet: delete Deletes the worksheet from the workbook.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtdelete-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtdelete-post-openapi.md
- name: Microsoft Graph API - Get Worksheet
  x-api-slug: workbookworksheetsltidnamegt-get
  description: Get Worksheet Retrieve the properties and relationships of worksheet
    object.
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegt-get-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegt-get-openapi.md
- name: Microsoft Graph API - Worksheet Range
  x-api-slug: workbookworksheetsltidnamegtrange-post
  description: 'Worksheet: Range Gets the range object specified by the address or
    name.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtrange-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtrange-post-openapi.md
- name: Microsoft Graph API - Update Worksheet
  x-api-slug: workbookworksheetsltidnamegt-patch
  description: Update worksheet Update the properties of worksheet object.
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegt-patch-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegt-patch-openapi.md
- name: Microsoft Graph API - Worksheet Used Range
  x-api-slug: workbookworksheetsltidnamegtusedrange-post
  description: 'Worksheet: UsedRange The used range is the smallest range that encompasses
    any cells that have a value or formatting assigned to them. If the worksheet is
    blank, this function will return the top left cell.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtusedrange-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtusedrange-post-openapi.md
- name: Microsoft Graph API - Worksheet Collection Add
  x-api-slug: workbookworksheets-post
  description: 'WorksheetCollection: add Adds a new worksheet to the workbook. The
    worksheet will be added at the end of existing worksheets. If you wish to activate
    the newly added worksheet, call ".activate() on it.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheets-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheets-post-openapi.md
- name: Microsoft Graph API - Get Worksheet Protection
  x-api-slug: workbookworksheetsltidnamegtprotection-get
  description: Get WorksheetProtection Retrieve the properties and relationships of
    worksheetprotection object.
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotection-get-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotection-get-openapi.md
- name: Microsoft Graph API - Worksheet Protection Protect
  x-api-slug: workbookworksheetsltidnamegtprotectionprotect-post
  description: 'WorksheetProtection: protect Protect a worksheet. It throws if the
    worksheet has been protected.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotectionprotect-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotectionprotect-post-openapi.md
- name: Microsoft Graph API - Worksheet Protection Unprotect
  x-api-slug: workbookworksheetsltidnamegtprotectionunprotect-post
  description: 'WorksheetProtection: unprotect Unprotect a worksheet'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotectionunprotect-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotectionunprotect-post-openapi.md
- name: Microsoft Graph API - Worksheet Protection Unprotect
  x-api-slug: workbookworksheetsltidnamegtprotectionunprotect-post
  description: 'WorksheetProtection: unprotect Unprotect a worksheet'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotectionunprotect-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotectionunprotect-post-openapi.md
- name: Microsoft Graph API - Worksheet Protection Protect
  x-api-slug: workbookworksheetsltidnamegtprotectionprotect-post
  description: 'WorksheetProtection: protect Protect a worksheet. It throws if the
    worksheet has been protected.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotectionprotect-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotectionprotect-post-openapi.md
- name: Microsoft Graph API - Get Worksheet Protection
  x-api-slug: workbookworksheetsltidnamegtprotection-get
  description: Get WorksheetProtection Retrieve the properties and relationships of
    worksheetprotection object.
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotection-get-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtprotection-get-openapi.md
- name: Microsoft Graph API - Worksheet Collection Add
  x-api-slug: workbookworksheets-post
  description: 'WorksheetCollection: add Adds a new worksheet to the workbook. The
    worksheet will be added at the end of existing worksheets. If you wish to activate
    the newly added worksheet, call ".activate() on it.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheets-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheets-post-openapi.md
- name: Microsoft Graph API - Worksheet Used Range
  x-api-slug: workbookworksheetsltidnamegtusedrange-post
  description: 'Worksheet: UsedRange The used range is the smallest range that encompasses
    any cells that have a value or formatting assigned to them. If the worksheet is
    blank, this function will return the top left cell.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtusedrange-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtusedrange-post-openapi.md
- name: Microsoft Graph API - Update Worksheet
  x-api-slug: workbookworksheetsltidnamegt-patch
  description: Update worksheet Update the properties of worksheet object.
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegt-patch-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegt-patch-openapi.md
- name: Microsoft Graph API - Worksheet Range
  x-api-slug: workbookworksheetsltidnamegtrange-post
  description: 'Worksheet: Range Gets the range object specified by the address or
    name.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtrange-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtrange-post-openapi.md
- name: Microsoft Graph API - Get Worksheet
  x-api-slug: workbookworksheetsltidnamegt-get
  description: Get Worksheet Retrieve the properties and relationships of worksheet
    object.
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegt-get-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegt-get-openapi.md
- name: Microsoft Graph API - Worksheet Delete
  x-api-slug: workbookworksheetsltidnamegtdelete-post
  description: 'Worksheet: delete Deletes the worksheet from the workbook.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtdelete-post-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtdelete-post-openapi.md
- name: Microsoft Graph API - Worksheet Cell
  x-api-slug: workbookworksheetsltidnamegtcellrowltrowgtcolumnltcolumngt-get
  description: 'Worksheet: Cell Gets the range object containing the single cell based
    on row and column numbers. The cell can be outside the bounds of its parent range,
    so long as it''s stays within the worksheet grid.'
  image: http://kinlane-productions.s3.amazonaws.com/api-evangelist-site/company/logos/microsoft-graph.png
  humanURL: https://developer.microsoft.com/en-us/graph/
  baseURL: https://graph.microsoft.com//
  tags: Microsoft, Files, Notes, Tasks, Stack Network, API Provider, Contacts, Emails,
    Profiles, Service API, Relative Data
  properties:
  - type: x-postman-collection
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtcellrowltrowgtcolumnltcolumngt-get-postman.md
  - type: x-openapi-spec
    url: https://raw.githubusercontent.com/streamdata-gallery-topics/worksheets/master/_listings/microsoft-graph/workbookworksheetsltidnamegtcellrowltrowgtcolumnltcolumngt-get-openapi.md
x-common:
- type: x-api-gallery
  url: http://messente.api.gallery.streamdata.io
- type: x-api-stack
  url: http://microsoft.graph.stack.network
- type: x-change-loge
  url: https://developer.microsoft.com/en-us/graph/docs/overview/changelog
- type: x-documentation
  url: https://developer.microsoft.com/en-us/graph/docs
- type: x-explorer
  url: https://developer.microsoft.com/en-us/graph/graph-explorer
- type: x-getting-started
  url: https://developer.microsoft.com/en-us/graph/docs/get-started/rest
- type: x-github
  url: https://github.com/microsoftgraph
- type: x-sdk
  url: https://developer.microsoft.com/en-us/graph/code-samples-and-sdks
- type: x-website
  url: https://developer.microsoft.com/en-us/graph/
include: []
maintainers:
- FN: Kin Lane
  x-twitter: apievangelist
  email: info@apievangelist.com
---