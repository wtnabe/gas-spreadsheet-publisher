# SpreadsheetPublisher for Google Apps Script

SpreadsheetPublisher Library for Google Apps Script Project ( **Sheets Container-bound** script )

## Prerequisite

Sheets Container-bound script Project.

## How to Use

### 0. Prepare Sheet

This library is for Spreadsheet Container-bound Script as it adds menus to Spreadsheet.

### 1. Add Library code

Choose one of them please.

 1. add Project ID for your project as Library `1reFNMbPJLeqNcUr2Agyq6VWsQyXcJWW-7nkQs152SYPGYmZhJZ4YxLZ7`
 2. Copy and Paste this code

I would recommend #2 for speed of execution, but #1 is also a good option for administrative costs.

### 2. Write setup code

open Script Editor

```javascript
function onOpen () {
  SpreadsheetPubliser.register(<target Spreadsheet ID>, [<src sheet>, <src sheet>, ..])
}
```

If no `<src sheet>`s specified, all sheets will be copied.

### 3. Reload Spreadsheet

### 4. Publish from Spreadsheet UI

open menu `[ Publisher ]` -> `[ Publish ]`
