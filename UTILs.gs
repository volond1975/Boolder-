/**
 * @description Изменяет запрос по принципу Utils.format из обьекта Опций
 * @param  {string} query
 * @param  {object} opsQuery
 * @returns {string} query
 */
function queryCreate(query,opsQuery){
  //var opsQuery={
  //  lisn:SpreadsheetApp.getActive().getRangeByName('ХідРозробкиЛісництво').getValue(),
  //  dil:SpreadsheetApp.getActive().getRangeByName('ХідРозробкиДілянка').getValue()
  //}
  //var query = "SELECT DISTINCT " + "[0],[1]" + "  FROM ? WHERE [0] LIKE '%"+"{lisn}"+"%'   ANd [1] LIKE '%"+"{dil}"+"%'";//Ділянки по ЛК 
  for(var key in opsQuery){
    var regex2 = new RegExp("\{"+key+"\}",'g');

    query=query.replace(regex2,opsQuery[key]);
     }
  return query;
  
} 


/*
  
  * input:
  
 
     arr2 =       [['a', 'cat',  1],
                   ['b', 'dog',  1],
                   ['c', 'bird', 1]];
    keyColumns =  [0, 2];
    
    // Note! The code assumes vakue-key pairs are unique:
    //   a1, b1, c1 
     
     
    
  * result
  
   object = { 
              a1: ['cat'],
              b1: ['dog'],
              c1: ['bird']
            };
*/

function convertArrayToObject(arr, keyColumns) {
  var result = {};
  
  var row = [];
  var keyArr = [];
  var valueArr = [];
  var key = '';
  
  for(var i = 0; i < arr.length; i++) {
  
    row = arr[i];
    keyArr = [];
    valueArr = [];

    for (var j = 0; j < row.length; j++) {
      if (keyColumns.indexOf(j) > -1) {
        keyArr.push(row[j]);
      }
      else
      {
        valueArr.push(row[j]);
      }    
    }
      
      key = keyArr.join('');      
      result[key] = valueArr;
    
  }

  return result;

}

function RangeToJSON(range,RangeKey){
Logger.log(typeof RangeKey)
var jsn=JSON.stringify(convertArrayToObject([range],[RangeKey]));
Logger.log(jsn)
Logger.log(JSON.stringify(RangeKey))
return jsn
}

function ARangeToJSON(range, RangeKey){
var r=  _ . fromPairs (range);
// => {'a': 1, 'b': 2}
   
//    var jsn = JSON.stringify(convertArrayToObject(r = range, rk = RangeKey));  
//    var ran = [];
  
    return JSON.stringify(r) ;
}

function createBanding(query,opsQuery){
  
  
  return queryCreate(query,JSON.parse(opsQuery))

}



/**
 * Takes a colour string and returns it to a hex string. If a non-matching string is
 * passed, it will return the argument as is - for this situation it means that a
 * hex string can be passed to it and be returned as is. This is not for production.
 * 
 * @param  {string} color    Must be either a colour name or hex string of color eg ("#ffffff")
 * 
 * @return {object|string}          hex string of color eg ("#ffffff") or the argument given.
 */
/**
 * Sums cell values in a range if they have the given background color
 * 
 * @param  {String} color    Hex string of color eg ("#ffffff")
 * @param  {Array.Array} range    Values of the desired range
 * @param  {int} startcol The column of the range
 * @param  {int} startrow The first row of the range
 * 
 * @return {int}          Sum of all cell values matching the condition
 */
function sumIfBgColor(color, range, startcol, startrow){
  // convert from int to ALPHANUMERIC - thanks to 
  // Daniel at http://stackoverflow.com/a/3145054/2828136
  var col_id = String.fromCharCode(64 + startcol);
  var endrow = startrow + range.length - 1
  // build the range string, then get the background colours
  var range_string = col_id + startrow + ":" + col_id + endrow
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var getColors = ss.getRange(range_string).getBackgrounds();

  var x = 0;
  for(var i = 0; i < range.length; i++) {
    for(var j = 0; j < range[0].length; j++) {
      // Sometimes the cell background is eg 'white' rather than '#ffffff'.
      // I don't know why - I think it's a bug.
      // so we remove that inconsistency with colourNameToHex
      // courtesy of Greg at http://stackoverflow.com/a/1573141/2828136
      if(colourNameToHex(getColors[i][j].toString()) == color) {
        x += range[i][j];
      }
    }
  }
  return x;
}

/**
 * Takes a colour string and returns it to a hex string. If a non-matching string is
 * passed, it will return the argument as is - for this situation it means that a
 * hex string can be passed to it and be returned as is. This is not for production.
 * 
 * @param  {string} color    Must be either a colour name or hex string of color eg ("#ffffff")
 * 
 * @return {object|string}          hex string of color eg ("#ffffff") or the argument given.
 */
function colourNameToHex(colour)
{
    var colours = {"aliceblue":"#f0f8ff","antiquewhite":"#faebd7","aqua":"#00ffff","aquamarine":"#7fffd4","azure":"#f0ffff",
    "beige":"#f5f5dc","bisque":"#ffe4c4","black":"#000000","blanchedalmond":"#ffebcd","blue":"#0000ff","blueviolet":"#8a2be2","brown":"#a52a2a","burlywood":"#deb887",
    "cadetblue":"#5f9ea0","chartreuse":"#7fff00","chocolate":"#d2691e","coral":"#ff7f50","cornflowerblue":"#6495ed","cornsilk":"#fff8dc","crimson":"#dc143c","cyan":"#00ffff",
    "darkblue":"#00008b","darkcyan":"#008b8b","darkgoldenrod":"#b8860b","darkgray":"#a9a9a9","darkgreen":"#006400","darkkhaki":"#bdb76b","darkmagenta":"#8b008b","darkolivegreen":"#556b2f",
    "darkorange":"#ff8c00","darkorchid":"#9932cc","darkred":"#8b0000","darksalmon":"#e9967a","darkseagreen":"#8fbc8f","darkslateblue":"#483d8b","darkslategray":"#2f4f4f","darkturquoise":"#00ced1",
    "darkviolet":"#9400d3","deeppink":"#ff1493","deepskyblue":"#00bfff","dimgray":"#696969","dodgerblue":"#1e90ff",
    "firebrick":"#b22222","floralwhite":"#fffaf0","forestgreen":"#228b22","fuchsia":"#ff00ff",
    "gainsboro":"#dcdcdc","ghostwhite":"#f8f8ff","gold":"#ffd700","goldenrod":"#daa520","gray":"#808080","green":"#008000","greenyellow":"#adff2f",
    "honeydew":"#f0fff0","hotpink":"#ff69b4",
    "indianred ":"#cd5c5c","indigo ":"#4b0082","ivory":"#fffff0","khaki":"#f0e68c",
    "lavender":"#e6e6fa","lavenderblush":"#fff0f5","lawngreen":"#7cfc00","lemonchiffon":"#fffacd","lightblue":"#add8e6","lightcoral":"#f08080","lightcyan":"#e0ffff","lightgoldenrodyellow":"#fafad2",
    "lightgrey":"#d3d3d3","lightgreen":"#90ee90","lightpink":"#ffb6c1","lightsalmon":"#ffa07a","lightseagreen":"#20b2aa","lightskyblue":"#87cefa","lightslategray":"#778899","lightsteelblue":"#b0c4de",
    "lightyellow":"#ffffe0","lime":"#00ff00","limegreen":"#32cd32","linen":"#faf0e6",
    "magenta":"#ff00ff","maroon":"#800000","mediumaquamarine":"#66cdaa","mediumblue":"#0000cd","mediumorchid":"#ba55d3","mediumpurple":"#9370d8","mediumseagreen":"#3cb371","mediumslateblue":"#7b68ee",
    "mediumspringgreen":"#00fa9a","mediumturquoise":"#48d1cc","mediumvioletred":"#c71585","midnightblue":"#191970","mintcream":"#f5fffa","mistyrose":"#ffe4e1","moccasin":"#ffe4b5",
    "navajowhite":"#ffdead","navy":"#000080",
    "oldlace":"#fdf5e6","olive":"#808000","olivedrab":"#6b8e23","orange":"#ffa500","orangered":"#ff4500","orchid":"#da70d6",
    "palegoldenrod":"#eee8aa","palegreen":"#98fb98","paleturquoise":"#afeeee","palevioletred":"#d87093","papayawhip":"#ffefd5","peachpuff":"#ffdab9","peru":"#cd853f","pink":"#ffc0cb","plum":"#dda0dd","powderblue":"#b0e0e6","purple":"#800080",
    "red":"#ff0000","rosybrown":"#bc8f8f","royalblue":"#4169e1",
    "saddlebrown":"#8b4513","salmon":"#fa8072","sandybrown":"#f4a460","seagreen":"#2e8b57","seashell":"#fff5ee","sienna":"#a0522d","silver":"#c0c0c0","skyblue":"#87ceeb","slateblue":"#6a5acd","slategray":"#708090","snow":"#fffafa","springgreen":"#00ff7f","steelblue":"#4682b4",
    "tan":"#d2b48c","teal":"#008080","thistle":"#d8bfd8","tomato":"#ff6347","turquoise":"#40e0d0",
    "violet":"#ee82ee",
    "wheat":"#f5deb3","white":"#ffffff","whitesmoke":"#f5f5f5",
    "yellow":"#ffff00","yellowgreen":"#9acd32"};

    if (typeof colours[colour.toLowerCase()] != 'undefined')
        return colours[colour.toLowerCase()];

    return colour;
}


/**
 * Takes a colour string and returns it to a hex string. If a non-matching string is
 * passed, it will return the argument as is - for this situation it means that a
 * hex string can be passed to it and be returned as is. This is not for production.
 * 
 * @param  {string} color    Must be either a colour name or hex string of color eg ("#ffffff")
 * 
 * @return {object|string}          hex string of color eg ("#ffffff") or the argument given.
 */
function HexToColourName(hex)
{
  var hexes = {"#f0f8ff":"aliceblue","#faebd7":"antiquewhite","#00ffff":"aqua","#7fffd4":"aquamarine","#f0ffff":"azure","#f5f5dc":"beige","#ffe4c4":"bisque",
      "#000000":"black","#ffebcd":"blanchedalmond","#0000ff":"blue","#8a2be2":"blueviolet","#a52a2a":"brown","#deb887":"burlywood","#5f9ea0":"cadetblue",
      "#7fff00":"chartreuse","#d2691e":"chocolate","#ff7f50":"coral","#6495ed":"cornflowerblue","#fff8dc":"cornsilk","#dc143c":"crimson","#00ffff":"cyan",
      "#00008b":"darkblue","#008b8b":"darkcyan","#b8860b":"darkgoldenrod","#a9a9a9":"darkgray","#006400":"darkgreen","#bdb76b":"darkkhaki",
      "#8b008b":"darkmagenta","#556b2f":"darkolivegreen","#ff8c00":"darkorange","#9932cc":"darkorchid","#8b0000":"darkred","#e9967a":"darksalmon",
      "#8fbc8f":"darkseagreen","#483d8b":"darkslateblue","#2f4f4f":"darkslategray","#00ced1":"darkturquoise","#9400d3":"darkviolet","#ff1493":"deeppink",
      "#00bfff":"deepskyblue","#696969":"dimgray","#1e90ff":"dodgerblue","#b22222":"firebrick","#fffaf0":"floralwhite","#228b22":"forestgreen",
      "#ff00ff":"fuchsia","#dcdcdc":"gainsboro","#f8f8ff":"ghostwhite","#ffd700":"gold","#daa520":"goldenrod","#808080":"gray","#008000":"green",
      "#adff2f":"greenyellow","#f0fff0":"honeydew","#ff69b4":"hotpink","#cd5c5c":"indianred ","#4b0082":"indigo ","#fffff0":"ivory","#f0e68c":"khaki",
      "#e6e6fa":"lavender","#fff0f5":"lavenderblush","#7cfc00":"lawngreen","#fffacd":"lemonchiffon","#add8e6":"lightblue","#f08080":"lightcoral",
      "#e0ffff":"lightcyan","#fafad2":"lightgoldenrodyellow","#d3d3d3":"lightgrey","#90ee90":"lightgreen","#ffb6c1":"lightpink","#ffa07a":"lightsalmon",
      "#20b2aa":"lightseagreen","#87cefa":"lightskyblue","#778899":"lightslategray","#b0c4de":"lightsteelblue","#ffffe0":"lightyellow","#00ff00":"lime",
      "#32cd32":"limegreen","#faf0e6":"linen","#ff00ff":"magenta","#800000":"maroon","#66cdaa":"mediumaquamarine","#0000cd":"mediumblue",
      "#ba55d3":"mediumorchid","#9370d8":"mediumpurple","#3cb371":"mediumseagreen","#7b68ee":"mediumslateblue","#00fa9a":"mediumspringgreen",
      "#48d1cc":"mediumturquoise","#c71585":"mediumvioletred","#191970":"midnightblue","#f5fffa":"mintcream","#ffe4e1":"mistyrose","#ffe4b5":"moccasin",
      "#ffdead":"navajowhite","#000080":"navy","#fdf5e6":"oldlace","#808000":"olive","#6b8e23":"olivedrab","#ffa500":"orange","#ff4500":"orangered",
      "#da70d6":"orchid","#eee8aa":"palegoldenrod","#98fb98":"palegreen","#afeeee":"paleturquoise","#d87093":"palevioletred","#ffefd5":"papayawhip",
      "#ffdab9":"peachpuff","#cd853f":"peru","#ffc0cb":"pink","#dda0dd":"plum","#b0e0e6":"powderblue","#800080":"purple","#ff0000":"red","#bc8f8f":"rosybrown",
      "#4169e1":"royalblue","#8b4513":"saddlebrown","#fa8072":"salmon","#f4a460":"sandybrown","#2e8b57":"seagreen","#fff5ee":"seashell","#a0522d":"sienna",
      "#c0c0c0":"silver","#87ceeb":"skyblue","#6a5acd":"slateblue","#708090":"slategray","#fffafa":"snow","#00ff7f":"springgreen","#4682b4":"steelblue",
      "#d2b48c":"tan","#008080":"teal","#d8bfd8":"thistle","#ff6347":"tomato","#40e0d0":"turquoise","#ee82ee":"violet","#f5deb3":"wheat","#ffffff":"white",
      "#f5f5f5":"whitesmoke","#ffff00":"yellow","#9acd32":"yellowgreen"} 
  
 if (typeof hexes[hex.toLowerCase()] != 'undefined')
        return hexes[hex.toLowerCase()];

    return hex;
}

function noBlankColumnsArr2D(arr){
var csvData =arr ,
    columns = csvData.reduce(
       function (r, a)  {return(a.forEach(function (v, i)  { return r[i] = r[i] || v}), r)},

        csvData[0].map(function(_)  {return false})
    );

csvData = csvData.map(function(a){return a.filter(function(_, i) {
    return columns[i]
})}

);
return csvData
}


//https://docs.google.com/spreadsheets/d/1SdTm3PCJ0FySN01bPgXKF2H_0HVJSOjL7XT21URV-5k/edit#gid=0
/**
 * Unpivot a pivot table of any size.
 *
 * @param {A1:D30} data The pivot table.
 * @param {1} fixColumns Number of columns, after which pivoted values begin. Default 1.
 * @param {1} fixRows Number of rows (1 or 2), after which pivoted values begin. Default 1.
 * @param {"city"} titlePivot The title of horizontal pivot values. Default "column".
 * @param {"distance"[,...]} titleValue The title of pivot table values. Default "value".
 * @return The unpivoted table
 * @customfunction
 */
function unpivot(data,fixColumns,fixRows,titlePivot,titleValue) {  
  var fixColumns = fixColumns || 1; // how many columns are fixed
  var fixRows = fixRows || 1; // how many rows are fixed
  var titlePivot = titlePivot || 'column';
  var titleValue = titleValue || 'value';
  var ret=[],i,j,row,uniqueCols=1;
  
  // we handle only 2 dimension arrays
  if (!Array.isArray(data) || data.length < fixRows || !Array.isArray(data[0]) || data[0].length < fixColumns)
    throw new Error('no data');
  // we handle max 2 fixed rows
  if (fixRows > 2)
    throw new Error('max 2 fixed rows are allowed');
  
  // fill empty cells in the first row with value set last in previous columns (for 2 fixed rows)
  var tmp = '';
  for (j=0;j<data[0].length;j++)
    if (data[0][j] != '') 
      tmp = data[0][j];
    else
      data[0][j] = tmp;
  
  // for 2 fixed rows calculate unique column number
  if (fixRows == 2)
  {
    uniqueCols = 0;
    tmp = {};
    for (j=fixColumns;j<data[1].length;j++)
      if (typeof tmp[ data[1][j] ] == 'undefined')
      {
        tmp[ data[1][j] ] = 1;
        uniqueCols++;
      }
  }
  
  // return first row: fix column titles + pivoted values column title + values column title(s)
  row = [];
    for (j=0;j<fixColumns;j++) row.push(fixRows == 2 ? data[0][j]||data[1][j] : data[0][j]); // for 2 fixed rows we try to find the title in row 1 and row 2
    for (j=3;j<arguments.length;j++) row.push(arguments[j]);
  ret.push(row);
    
  // processing rows (skipping the fixed columns, then dedicating a new row for each pivoted value)
  for (i=fixRows;i<data.length && data[i].length > 0 && data[i][0];i++)
  {
    row = [];
    for (j=0;j<fixColumns && j<data[i].length;j++)
      row.push(data[i][j]);
    for (j=fixColumns;j<data[i].length;j+=uniqueCols)
      ret.push( 
        row.concat([data[0][j]]) // the first row title value
        .concat(data[i].slice(j,j+uniqueCols)) // pivoted values
      );
  }

  return ret;
}
