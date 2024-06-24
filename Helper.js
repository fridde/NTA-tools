class SpreadsheetHelper {
  
  /**
  * @param {Spreadsheet} aSS - The active Spreadsheet instance
  */
  constructor(aSS)
  {
    this.aSS = aSS;            
    this.Logger = BetterLog.useSpreadsheet();
  }

  /**
  * @param {String} title - the main title of  
  * @param {Array<Array>} items - a 2-dimensional array with each item consisting of a title and a function name
  */
  createMenu(title, items)
  {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu(title);

    items.forEach((item) => menu.addItem(item[0], item[1]));

    menu.addToUi();
  }

  /**
  * @param {String|Sheet} sheet - A string or a sheet object
  *
  * @return {Sheet} a Sheet object
  */
  getSheetAsObject(sheet)
  {
    if(typeof sheet === 'string'){
      sheet = this.aSS.getSheetByName(sheet);
    }

    return sheet;
  }

  /**
  * @param {String|Sheet} sheet - A string or a sheet object
  * @param {Array<Number>} location - The location given as 2 coordinates in the form [row, col] (beginning at 1)
  * @param {Array<Array> | Array<Object>} values
  * @param {Boolean} [clearRightAndBelow=true] -  Should everything "beyond" be cleared?
  */
  insertValuesAt(sheet, location, values, clearRightAndBelow = true)
  {
    sheet = this.getSheetAsObject(sheet);
    values = this.ensureRowsAreArrays(values);

    const row = location[0];
    const col = location[1];

    let range =  sheet.getRange(row, col, values.length, values[0].length);
    if(clearRightAndBelow){
      sheet.getRange(row, col, sheet.getLastRow(), sheet.getLastColumn()).clear();
    }
    range.setValues(values);
}

  /**
  * Converts the array to an object using the key-array as keys
  * (["a", "b", "c"], ["x", "y", "z"]) becomes {"a": "x", "b": "y", "c": "z"}
  * 
  * @param {Array<string>} keys - The keys to use for the values
  * @param {Array} array - the array containing the values
  * @return {Object} the object
  */
  createObjectWithKeys(keys, array)
  {
    return Object.fromEntries(keys.map((k, i) => [k, array[i]]));
  }

  /**
   * @example 
   * arr1 = [{a: "x", b:12}, {a: "y", b: 13}, {a: "z", b: 14, c: "hello"}]
   * this.redefineKeysByProperty(arr1, "a")
   * // becomes {x: {a: "x", b:12}, y: {a: "y", b: 13},  z: {a: "z", b: 14, c: "hello"}}
   * 
   * @param {Array<Object>} array -  The array of objects where each object contains the keyString as key
   * @param {string} keyString - The key of what should be used as the key for each object
   * 
   * @return {Object<Object>} - a 2-dimensional object that uses the given values as keys for each object 
   */
  redefineKeysByProperty(array, keyString)
  {
    const returnObj = {};
    array.forEach(element => returnObj[element[keyString]] = element);

    return returnObj;
  }

  /**
   * @param {Array<Array> | Array<Object>} array - The array consisting either of arrays or objects
   * 
   * @return {Array<Array>}
   */
  ensureRowsAreArrays(array)
  {
    if(typeof array[0] !== 'array'){
      array = array.map(row => Object.values(row));
    }
    return array;
  }

  firstRowAsHeader(values)
  {
    const header = values.shift();
  
    return values.map(row => this.createObjectWithKeys(header, row));  
  }

  /**
  * @param {String} rangeName - The name of the range
  * @return {Range} the Range object
  */
  getNamedRange(rangeName)
  {
    return this.aSS.getRangeByName(rangeName);
  }
  
  /**
  * @param {String} rangeName - The name of the range

  * @return {Array<Array>} an 2-dimensional array with a row of values as each element
  */
  getNamedValues(rangeName)
  {
    return this.getNamedRange(rangeName).getValues();
  }

  /**
  * @param {String} rangeName - The name of the range

  * @return {String|Number} the value of the single named cell or the left-upper-most cell of the named range
  */
  getSingleNamedValue(rangeName)
  {
    return this.getNamedRange(rangeName).getValue();
  }

  /**
   * @param {String|Sheet} sheet - A string or a sheet object
   * 
   * @return {Array<Array>} an 2-dimensional array with a row of values as each element
   */
  getValuesFromSheet(sheet)
  {
    return this.getSheetAsObject(sheet).getDataRange().getValues();
  }


  /**
   * @param {Array<Object>} array - The array of objects to clean
   * @param {String} [keyToCheck='*'] - Either the * symbol if all columns have to be empty or the key of which element to check for emptyness 
   * 
   * @return {Array<Object>} -  the cleaned array of objects
   */
  
  removeIfEmpty(array, keyToCheck = '*')
  {
    return array.filter(function(row){
      if(keyToCheck === '*'){
        return Object.values(row).some(cell => cell !== '');
      }
      return row[keyToCheck] !== '';  
    });
  }

  /**
   * Creates a unique array using the JSON string representation for comparison
   * 
   * @param {Array} array
   * 
   * @return {Array}
   */
  uniqueArray(array)
  {    
      const set = [...new Set(array.map(JSON.stringify))];
   
      return Array.from(set, JSON.parse);
  }

  /** 
   * @param {Object} obj
   *
   * @return {Boolean} - Tells if the object is empty or not
   */
  objectIsEmpty(obj)
  {
    return Object.keys(obj).length === 0;
  }

  /**
   * Adds or substracts days from a date object
   *
   * @param {Date} date the Date given as a Date object
   * @param {number} nrDays the number of days, given either as positive or negative integer
   * @return {Date} the altered date
   */
  addDays(date, nrDays)
  {  
    let clonedDate = new Date(date.getTime());
    clonedDate.setDate(clonedDate.getDate() + nrDays);
    
    return clonedDate;
  }

  /**
   * Split array into chunks of a given size
   * 
   * @example
   * chunk([1, 2, 3, 4, 5], 2); 
   * // [[1, 2], [3, 4], [5]]
   * 
   * @param {Array} array - The long array
   * @param {number} size - The maximum size each chunk should have.
   * 
   * @return {Array<Array>} - an array chunked up into pieces each the size of "size", the last elements dangling in an array of their own
   */
  chunk(array, size)
  { 
    
    return Array.from({ length: Math.ceil(array.length / size) }, (v, i) =>
      array.slice(i * size, i * size + size));
  }


  /**
  * @param {Array|String} text - Either the text to be shown or an array of strings, where each element creates a new line
  */
  alert(text)
  {
    if(Array.isArray(text)){
      text = text.join("\n");
    }
  
    SpreadsheetApp.getUi().alert(text);
  }


  /**
  * @param {Array<Object>} arrayOfObjects
  *
  * @return {Array<Array>}
  */
  convertToGridWithHeaders(arrayOfObjects)
  {
    const headers = Object.keys(arrayOfObjects[0]);

    const arrayOfArrays = arrayOfObjects.map(obj => Object.values(obj));

    arrayOfArrays.unshift(headers);

    return arrayOfArrays;
  }


  /**
   * @param {Array<Array>} arrayOfArrays
   * 
   * @return {String}
   */
  convertToHTMLTable(arrayOfArrays)
  {
    const [headings, ...rows] = arrayOfArrays;
    
    const headingsText = headings.map(th => '<th>'+th+'</th>').join('');
    const rowsText = rows.map(row => {
      return '<tr>' + row.map(td => '<td>'+td+'</td>').join('') + '</tr>';
    }).join('');
    
    return `<table> 
      <thead>
        ${headingsText}
      </thead>
      <tbody>
        ${rowsText}
      </tbody>
    </table>
    <style>
    table, th, td {
      border: 0.5px solid black;
    }
    </style>
      `;
  }  
}