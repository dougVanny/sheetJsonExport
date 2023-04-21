/*
 *     Copyright 2023 Douglas Vanny @ https://github.com/dougVanny/
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *          http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

function addToMenu(menu, module)
{
  var ui = SpreadsheetApp.getUi();
  menu = menu.addSubMenu(ui.createMenu("JSON")
        .addItem("Export as JSON", (module?module+".exportJson":"exportJson"))
        .addSeparator()
        .addItem("JSON Export - v2.0",(module?module+".same":"same")));
  return menu;
}

function exportJson()
{
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var parseResult = exportJsonFromSpreadsheet(spreadsheet);
  ui.alert("Log",JSON.stringify(parseResult),ui.ButtonSet.OK);
}

function exportJsonFromSpreadsheet(spreadsheet)
{
  var sheets = spreadsheet.getSheets();
  var parseResult = {};

  var globalObjects = {}

  for(var s=0; s<sheets.length; s++)
  {
    if(sheets[s].getName()[0]=="+")
    {
      var objects = parseSheet(sheets[s], globalObjects);

      for(var key in objects)
      {
        if(!parseResult[key]) parseResult[key] = [];

        parseResult[key] = parseResult[key].concat(objects[key])
      }
    }
  }

  return parseResult;
}

function parseSheet(sheet, globalObjects)
{
  // We only look at the bottom 2 frozen rows for header. Everything below it is content
  var headerRow = sheet.getFrozenRows()-2;

  var values = sheet.getRange(headerRow,1,sheet.getLastRow()-headerRow+1,sheet.getLastColumn()).getDisplayValues();
  var paths = values[0];
  var types = values[1];
  var fields = values[2];
  values.splice(0,3)

  // Root Objects are objects that don't live inside another. Any non root object lives inside a root one, directly or indirectly
  var rootObjects = {}

  // List of currently known objects. One object of each Path is known. If an object is 4 rows high, it will be on this list until reaching row 5, where it will be replaced by another
  var currentObjects = {}

  for(var row=0; row<values.length; row++)
  {
    // For each row, we track which paths have already been seen. This is done becuase, when we first find a field of Path X in a certain, we consider that a new object, if we find another field
    // in the same row, we ae able to reuse the previously created object instead of creting a new one
    var visitedPaths = []

    var recordedPath = null;
    var recordedField = null;
    var recordedType = null;

    for(var col=0; col<values[row].length; col++)
    {
      // Path is a terrible name for what is actually a struct. But this is more like a path to a struct, like "ClassA.ClassB", so here we are
      // We remove whitespaces as they can be used in the spreadsheet for better readibility
      var path = paths[col].replace(/\s/g,"").trim()
      // Name of the field where we will store the value
      var field = fields[col].trim()
      var type = types[col].trim()
      // This is the value of each individual cell
      var value = values[row][col].trim()

      if(path.length==0)
      {
        if(recordedPath==null) continue;

        path = recordedPath;
      }
      else
      {
        recordedPath=null;
      }

      if(field.length==0)
      {
        if(recordedField!=null) field = recordedField;
        if(recordedType!=null) type = recordedType;
      }
      else
      {
        recordedField=null;
        recordedType=null;
      }

      if(path.indexOf("...") == path.length-3)
      {
        path = path.substring(0,path.length-3);
        recordedPath = path;
      }

      if(field.indexOf("...") == field.length-3)
      {
        field = field.substring(0,field.length-3);
        recordedField = field;
        recordedType = type;
      }

      // Due to spreadsheet constains, it is not trivial to differentiate lists from non-lists. We therefore use '{}' as a way to tell if a path is supposed to be stored
      // as a single object instead of a list in it's parent path
      var singleObjectPath = path.includes("{}");
      path = path.replace("{}","")

      // '*' is used to link between objects across different sheets. This is further explained down below, along with the '&' symbol
      if(field.length!=0 && value.length!=0 && field.includes("*"))
      {
        field = field.replace("*","")

        currentObjects[path] = globalObjects[path][field][value]

        visitedPaths.push(path)

        continue;
      }

      // We need to know if the current column if the first ocurrence of a path. If the path contains a '?', this check is bypassed
      var firstColPath;
      
      if(path.includes("?"))
      {
        firstColPath = false;
        path = path.replace("?","")
      }
      else if(value.length == 0)
      {
        if(field.length!=0)
        {
          firstColPath = false;
        }
        else
        {
          if(path.includes("."))
          {
            var parent = currentObjects[path.substring(0,path.lastIndexOf("."))];

            if(visitedPaths.indexOf(parent)>=0) firstColPath=false;
          }
        }
      }
      else
      {
        firstColPath = !visitedPaths.includes(path);
      }
      
      if(firstColPath)
      {
        // If firstColPath is true, we clean whatever object previously existed under the current path with a new one, and start building from it
        currentObjects[path] = {}
        visitedPaths.push(path)
        
        // If the current path has multiple parts, we look at the parent path, find it's current object, and attach the new object as a child
        if(path.includes("."))
        {
          var parent = currentObjects[path.substring(0,path.lastIndexOf("."))]
          var pathLeaf = path.substring(path.lastIndexOf(".")+1);

          if(singleObjectPath)
          {
            parent[pathLeaf] = currentObjects[path]
          }
          else
          {
            if(!parent[pathLeaf]) parent[pathLeaf] = [];
            parent[pathLeaf].push(currentObjects[path])
          }
        }
        else
        {
          // If path contians only one part, we consider that a root object, not linking it to any nonexistant parent, but instead adding it to the rootObjects dict that is later returned
          if(!rootObjects[path]) rootObjects[path]=[]

          rootObjects[path].push(currentObjects[path])
        }
      }

      if(value.length==0 || field.length==0) continue;

      // Once we guarantee the object is created and well stored, we start to fill its fields
      var obj = currentObjects[path];

      // If the field contains a '[]', instead of directly assigning it to the object, we push them inside a list. This is used for both horizontal and vertical lists within the spreadsheet
      if(field.includes("[]"))
      {
        // We do not trust spreadsheet types, so we run this through a parser first, always using the display value of the cell
        var parsedValue = PARSERS[type](value)

        field = field.replace("[]","")

        if(!obj[field]) obj[field] = []

        obj[field].push(parsedValue)
      }
      // If the field contains a '[X]', we treat the value contents as a single list, splitting it by 'X' and storing within the object
      else if(field.includes("[") && field.includes("]") && field.indexOf("[") < field.indexOf("]"))
      {
        var separator = field.substring(field.indexOf("[")+1,field.indexOf("]"))

        field = field.replace("["+separator+"]","")

        var valueList = value.split(separator)

        if(!obj[field]) obj[field] = []
        for(var i=0; i<valueList.length; i++)
        {
          obj[field].push(PARSERS[type](valueList[i].trim()))
        }
      }
      // If no list indicator is found, store that value directly into the object
      else
      {
        // If the field name contains a '&', we also store the object as a global object. This is used to cross reference objects across multiple sheets
        // If, at any other sheet, a '*' is found in the field, instead of creating a new object, we pull the one stored as a global object
        if(field.includes("&"))
        {
          field = field.replace("&","")

          if(!globalObjects[path]) globalObjects[path] = {};
          if(!globalObjects[path][field]) globalObjects[path][field] = {};
          globalObjects[path][field][value] = obj;
        }

        var parsedValue = PARSERS[type](value)

        obj[field] = parsedValue
      }
    }
  }

  return rootObjects;
}

function same(_){return _}
function parseBool(value)
{
  return value.toLowerCase()!="false";
}
function _parseFloat(value)
{
  if(value[value.length-1]=="%")
  {
    return parseFloat(value.substring(0,value.length-1))/100.0;
  }

  return parseFloat(value);
}

PARSERS = {
  "": same,
  "str": same,
  string: same,
  int: parseInt,
  float: _parseFloat,
  bool: parseBool,
}
