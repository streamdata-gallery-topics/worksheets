---
swagger: "2.0"
x-collection-name: Microsoft Graph
x-complete: 0
info:
  title: Microsoft Graph Worksheet Range
  description: 'Worksheet: Range Gets the range object specified by the address or
    name.'
  version: 1.0.0
host: graph.microsoft.com
basePath: /
schemes:
- http
produces:
- application/json
consumes:
- application/json
paths:
  /workbook/worksheets:
    get:
      summary: List Worksheets
      description: List worksheets Retrieve a list of worksheet objects.
      operationId: ListWorksheets
      x-api-path-slug: workbookworksheets-get
      parameters:
      - in: header
        name: Authorization
        description: Bearer
      responses:
        200:
          description: OK
      tags:
      - List
      - Worksheets
  /workbook/worksheets(&lt;id|name&gt;)/Cell(row=&lt;row&gt;,column=&lt;column&gt;):
    get:
      summary: Worksheet Cell
      description: 'Worksheet: Cell Gets the range object containing the single cell
        based on row and column numbers. The cell can be outside the bounds of its
        parent range, so long as it''s stays within the worksheet grid.'
      operationId: Worksheet:Cell
      x-api-path-slug: workbookworksheetsltidnamegtcellrowltrowgtcolumnltcolumngt-get
      parameters:
      - in: header
        name: Authorization
        description: Bearer
      responses:
        200:
          description: OK
      tags:
      - Worksheet
      - Cell
  /workbook/worksheets(&lt;id|name&gt;)/delete:
    post:
      summary: Worksheet Delete
      description: 'Worksheet: delete Deletes the worksheet from the workbook.'
      operationId: Worksheet:Delete
      x-api-path-slug: workbookworksheetsltidnamegtdelete-post
      parameters:
      - in: header
        name: Authorization
        description: Bearer
      responses:
        200:
          description: OK
      tags:
      - Worksheet
  /workbook/worksheets(&lt;id|name&gt;):
    get:
      summary: Get Worksheet
      description: Get Worksheet Retrieve the properties and relationships of worksheet
        object.
      operationId: GetWorksheet
      x-api-path-slug: workbookworksheetsltidnamegt-get
      parameters:
      - in: header
        name: Authorization
        description: Bearer
      responses:
        200:
          description: OK
      tags:
      - Worksheet
    patch:
      summary: Update Worksheet
      description: Update worksheet Update the properties of worksheet object.
      operationId: UpdateWorksheet
      x-api-path-slug: workbookworksheetsltidnamegt-patch
      parameters:
      - in: header
        name: Authorization
        description: Bearer
      responses:
        200:
          description: OK
      tags:
      - Worksheet
  /workbook/worksheets(&lt;id|name&gt;)/Range:
    post:
      summary: Worksheet Range
      description: 'Worksheet: Range Gets the range object specified by the address
        or name.'
      operationId: Worksheet:Range
      x-api-path-slug: workbookworksheetsltidnamegtrange-post
      parameters:
      - in: header
        name: Authorization
        description: Bearer
      responses:
        200:
          description: OK
      tags:
      - Worksheet
      - Range
  /workbook/worksheets(&lt;id|name&gt;)/UsedRange:
    post:
      summary: Worksheet Used Range
      description: 'Worksheet: UsedRange The used range is the smallest range that
        encompasses any cells that have a value or formatting assigned to them. If
        the worksheet is blank, this function will return the top left cell.'
      operationId: Worksheet:UsedRange
      x-api-path-slug: workbookworksheetsltidnamegtusedrange-post
      parameters:
      - in: header
        name: Authorization
        description: Bearer
      responses:
        200:
          description: OK
      tags:
      - Worksheet
      - Used
      - Range
  /workbook/worksheets/:
    post:
      summary: Worksheet Collection Add
      description: 'WorksheetCollection: add Adds a new worksheet to the workbook.
        The worksheet will be added at the end of existing worksheets. If you wish
        to activate the newly added worksheet, call ".activate() on it.'
      operationId: WorksheetCollection:Add
      x-api-path-slug: workbookworksheets-post
      parameters:
      - in: header
        name: Authorization
        description: Bearer
      responses:
        200:
          description: OK
      tags:
      - Worksheet
      - Collection
  /workbook/worksheets(&lt;id|name&gt;)/protection:
    get:
      summary: Get Worksheet Protection
      description: Get WorksheetProtection Retrieve the properties and relationships
        of worksheetprotection object.
      operationId: GetWorksheetProtection
      x-api-path-slug: workbookworksheetsltidnamegtprotection-get
      parameters:
      - in: header
        name: Authorization
        description: Bearer
      responses:
        200:
          description: OK
      tags:
      - Worksheet
      - Protection
  /workbook/worksheets(&lt;id|name&gt;)/protection/protect:
    post:
      summary: Worksheet Protection Protect
      description: 'WorksheetProtection: protect Protect a worksheet. It throws if
        the worksheet has been protected.'
      operationId: WorksheetProtection:Protect
      x-api-path-slug: workbookworksheetsltidnamegtprotectionprotect-post
      parameters:
      - in: header
        name: Authorization
        description: Bearer
      responses:
        200:
          description: OK
      tags:
      - Worksheet
      - Protection
      - Protect
  /workbook/worksheets(&lt;id|name&gt;)/protection/unprotect:
    post:
      summary: Worksheet Protection Unprotect
      description: 'WorksheetProtection: unprotect Unprotect a worksheet'
      operationId: WorksheetProtection:Unprotect
      x-api-path-slug: workbookworksheetsltidnamegtprotectionunprotect-post
      parameters:
      - in: header
        name: Authorization
        description: Bearer
      responses:
        200:
          description: OK
      tags:
      - Worksheet
      - Protection
      - Unprotect
x-streamrank:
  polling_total_time_average: 0
  polling_size_download_average: 0
  streaming_total_time_average: 0
  streaming_size_download_average: 0
  change_yes: 0
  change_no: 0
  time_percentage: 0
  size_percentage: 0
  change_percentage: 0
  last_run: ""
  days_run: 0
  minute_run: 0
---