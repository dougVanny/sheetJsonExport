# JSON Export

This document is a massive WIP and has nothing but the bare minimum to let someone start using this script

## How to setup

Add the attached script to your spreadsheet's Apps Script and, in another file, add the following code

    var ui = SpreadsheetApp.getUi();
    
    function onOpen()
    {
      var menu = ui.createMenu("Tools");
      menu = JSONExport.addToMenu(menu,"JSONExport");
      menu.addToUi();
    }
    
    function exportJson()
    {
      JSONExport.exportJson();
    }


## Spreadsheet Format

The following table will generate the follow JSON

|Character|Character|Character|Character.Skills|Character.Skills|
|--|--|--|--|--|
||**int**|**bool**||**float**|
|**name**|**hp**|**isPlayable**|**skillName**|**damage**|
|_Frozen Rows Division_|
|marcio|100|TRUE|jump|10|
||||fireball|37.5|
|ruiji|120|FALSE|jump|12|
||||vacuum|26.12|
||||escape||

```json
{"Character":[
  {"name": "marcio",
   "hp": 100,
   "isPlayable": true,
   "Skills": [
	 {"skillName": "jump",
	  "damage": 10},
	 {"skillName": "fireball",
	  "damage": 37.5}]
  },
  {"name": "ruiji",
   "hp": 120,
   "isPlayable": false,
   "Skills": [
	 {"skillName": "jump",
	  "damage": 12},
	 {"skillName": "vacuum",
	  "damage": 26.12},
	 {"skillName": "escape"}]
  }
]}
```
