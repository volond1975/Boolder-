var _ = LodashGS.load()


function onEdit(e){
//  var range = e.range;
  //SpreadsheetApp.getActiveSpreadsheet().toast(range.getSheet().getName())
  TESTgetDataRegion(e.range.getSheet().getName())
 // Logger.log(RangeToJSON(range.getSheet().getRange("G25:H28"),range.getSheet().getRange("G24")))
  //SpreadsheetApp.getActiveSpreadsheet().toast( ubdate())
}

function GETBACKGROUND(ref) {
    return SpreadsheetApp.getActiveRange().getBackground();
}
function ubdate(){
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheet = ss.getActiveSheet(),
    range = sheet.getDataRange(),
    formulas = range.getFormulas();
    range.setFormulas(formulas)
    Logger.log(formulas)
    return true
}
function splitByRegExp(range){
  
  var text="Виділ 10-5"//range.getValue()
var VP=text.replace("Виділ","").trim().split(/[-_.,]/)
 objHeader['Виділ']=VP[0]
 objHeader['ПідВиділ']=VP[1]||"0"

}


function createTextFinder(){

  var urlGet='https://docs.google.com/spreadsheets/d/1Mut0h6RBi35y9Y7o9pmNhSrea7sP09IlrWKqme9h3YU/edit?usp=sharing'
  var spreadsheetGet = SpreadsheetApp.openByUrl(urlGet)
  var sortimentna={}
  var findAll={}
  var arr=[]
var FinderText= ['виходу сортиментів з  лісосіки головного користування','Лісгосп','Лісництво','Квартал','Виділ','Площа','Порода','кбм']
FinderText.forEach(function(item,i){
findAll[item]=spreadsheetGet.createTextFinder(item).findAll()
}) 
  
var sortimentna={}
var funSpl=FuncConstr("obj,key,i","return obj[key][i].getValue().split(' ')[1]")
 findAll[FinderText[0]].forEach(function(range,i){
   var objHeader={}
   
    if(range.getValue()===FinderText[0]+' 2019 року.'){
  //  Logger.log(range.getSheet().getName()+'!-'+range.getA1Notation())
 objHeader['Лісгосп']=findAll['Лісгосп'][i].getValue().match(/(\")(.*?)(\")/)[2].replace("ЛГ","").trim() //Віделяем из кавічек
 objHeader['Лісництво']=funSpl(findAll,'Лісництво',i)//findAll['Лісництво'][i].getValue().split(" ")[1]
 var objDil={}
   objDil["Квартал"]=funSpl(findAll,'Квартал',i)//findAll['Квартал'][i].getValue().split(" ")[1]
   objDil["VP"]=findAll['Виділ'][i].getValue().replace("Виділ","").trim().split(/[-_.,]/)
   objDil["Виділ"]=objDil['VP'][0]
   objDil["ПідВиділ"]=objDil["VP"][1]||"0"
   objDil["Площа"]=funSpl(findAll,'Площа',i);//findAll['Площа'][i].getValue().split(" ")[1]
   objHeader['Ділянка']=queryCreate('{Квартал} кв ({Виділ} вид) {ПідВиділ} діл.',objDil)//'{Квартал} кв ({Виділ} вид) {ПідВиділ} діл.'

   objHeader['Данні']=[]
// Logger.log(findAllEndColumn[i].getMergedRanges())
   // Получение таблицы 
var Data= getValuesTableRange(findAll,'Порода','кбм',i) 


//Обработка таблицы
objHeader['Данні']=unpivotTableRange(Data.arr)['Данні']


var Headers=objHeader['Данні'].splice(0,1)
var keyObj=[objHeader['Лісництво'],objHeader['Ділянка'],objHeader['Площа']].join("_")
//objHeader['Заголовки Данні']= Headers  
   
   
   
   
   
   
   
   
   var keyObj=[objHeader['Лісництво'],objHeader['Ділянка'],objDil['Площа']].join("_")
sortimentna[keyObj]=objHeader
    }
})
// Logger.log(JSON.stringify(sortimentna))
var NewArr=TESTImportJSONFromSheet(sortimentna)
Logger.log(NewArr)
//Добавим на лист Сортиментна
var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Сортиментна')
sheet.clear()
var range = sheet.getRange(2,1,NewArr.length, NewArr[0].length); // check edges or uses a1 notation
range.setValues(NewArr)

                         }

//funct

function TESTImportJSONFromSheet(object){
  
  var tmplArr=["Лісгосп","Лісництво","Ділянка"]//["Лісгосп","Лісництво","Ділянка"]
  var sheetName="Parametrs"
  var arrNew=[]
// var object = getDataFromNamedSheet_(sheetName);
//Logger.log(JSON.stringify(object))//JSON.stringify(object[key]) 
  for (fuldil in object){
    var  arr=[] 
    for (key in object[fuldil]){
     if (_.isArray(object[fuldil][key])){ 
        var result=[]
      result =  object[fuldil][key].map(function(item) {
  return arr.concat(item);
  //  var  arr=[] 
})
  //Logger.log(result)
      
        arrNew.push(result)
      
     }
      else
        
      {
      arr.splice(tmplArr.indexOf(key), 0, object[fuldil][key]);
   //   Logger.log(JSON.stringify(arr))  
      }
    }
  }
 var arrNew=arrNew.reduce(function(r,cur){return r.concat(cur)})
Logger.log(arrNew)
//Logger.log(f.length)
//Logger.log(f[0].length)
// var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Сортиментна')
//sheet.clear()
//Logger.log(sheet.getLastRow())
//arrNew.forEach(function( item,i){
//var range = sheet.getRange(sheet.getLastRow()+1,1,item.length, item[0].length)
//Logger.log(range.getA1Notation())
 
  //range.setValues(item)
  //})
  return arrNew
}




function getValuesTableRange(findAll,keyStart,KeyEnd,i){
  
var Data={}
 Data.rowData=findAll[keyStart][i].getRow()+findAll[keyStart][i].getMergedRanges()[0].getNumRows(),//Строка где начинаються данні
 Data.columnData=findAll[keyStart][i].getColumn(),
 Data.numColumnsData=findAll[KeyEnd][i].getColumn()-findAll[keyStart][i].getColumn()+1,//Количество столбцов
 Data.numRowsData=findAll[keyStart][i].getDataRegion().getNumRows()-findAll[keyStart][i].getMergedRanges()[0].getNumRows(),//Количество столбцов -findAllEndColumn[i].getRows()
 Data.range=findAll[keyStart][i].getSheet().getRange(Data.rowData, Data.columnData, Data.numRowsData, Data.numColumnsData),
 Data.arr=Data.range.getValues()

return Data
}

function unpivotTableRange(arr){
  var obj={}
_.remove(arr, function(item,i,arr) {
  return item[0]==='Всього' || +item[1]===0;
});
  
obj['Pivot']=[['Порода','Всього','Ділова','A','B','C','D','ПП','НП']].concat(noBlankColumnsArr2D(arr))     
obj['Данні']=unpivot(obj['Pivot'],1,1,"Клас якості","Обь'ем" ) 
_.remove(obj['Данні'], function(item,i,arr) {
  return item[1]==='Всього' || item[1]==='Ділова' || item[2]==0;
});
return obj
}








function getRangeTextFinder(findTextYear,ops){
  var url='https://docs.google.com/spreadsheets/d/1Mut0h6RBi35y9Y7o9pmNhSrea7sP09IlrWKqme9h3YU/edit?usp=sharing'
    var spreadsheet = SpreadsheetApp.openByUrl(url)||SpreadsheetApp.getActiveSpreadsheet();
//var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();  
  var objHeader={} 
  //FIXME Учесть любовь юзверей к задвоеным пробелам
//findTextYear=findTextYear||'виходу сортиментів з  лісосіки головного користування 2019 року.'

var f=[]
findTextYear=findTextYear||'виходу сортиментів з  лісосіки головного користування'
 var firstOccurrence=spreadsheet.createTextFinder(findTextYear).findAll()
 var findAllLisgosp=spreadsheet.createTextFinder('Лісгосп').findAll()
 var findAllLisn=spreadsheet.createTextFinder('Лісництво').findAll()
 var findAllKv=spreadsheet.createTextFinder('Квартал').findAll()
 var findAllVid=spreadsheet.createTextFinder('Виділ').findAll() 
 var findAllPl=spreadsheet.createTextFinder('Площа').findAll()
 var findAllPoroda=spreadsheet.createTextFinder('Порода').findAll()
 var findAllEndColumn=spreadsheet.createTextFinder('кбм').findAll()
 var sortimentna={}
 
 firstOccurrence.forEach(function(range,i){
   if(range.getValue()===findTextYear+' 2019 року.'){
 objHeader['Лісгосп']=findAllLisgosp[i].getValue().match(/(\")(.*?)(\")/)[2].replace("ЛГ","").trim() //Віделяем из кавічек

 objHeader['Лісництво']=findAllLisn[i].getValue().split(" ")[1]
//  f.push(findAllLisn[i].getValue().split(" ")[1])
 objHeader['Квартал']=findAllKv[i].getValue().split(" ")[1]
//  f.push(findAllKv[i].getValue().split(" ")[1])
 var VP=findAllVid[i].getValue().replace("Виділ","").trim().split(/[-_.,]/)
 
 objHeader['Виділ']=VP[0]
//  f.push(VP.split(/[-_.,]/)[0])
 objHeader['ПідВиділ']=VP[1]||"0"
//  f.push(VP.split(/[-_.,]/)[1]||"0")
 objHeader['Площа']=findAllPl[i].getValue().split(" ")[1]
 objHeader['Ділянка']=queryCreate('{Квартал} кв ({Виділ} вид) {ПідВиділ} діл.',objHeader)//'{Квартал} кв ({Виділ} вид) {ПідВиділ} діл.'
//   f.push(VP.split(queryCreate('{Квартал} кв ({Виділ} вид) {ПідВиділ} діл.',objHeader)//'{Квартал} кв ({Виділ} вид) {ПідВиділ} діл.')
 objHeader['Данні']=[]
// Logger.log(findAllEndColumn[i].getMergedRanges())
 var Data={}
 Data.rowData=findAllPoroda[i].getRow()+findAllPoroda[i].getMergedRanges()[0].getNumRows(),//Строка где начинаються данні
 Data.columnData=findAllPoroda[i].getColumn(),
 Data.numColumnsData=findAllEndColumn[i].getColumn()-findAllPoroda[i].getColumn()+1,//Количество столбцов
 Data.numRowsData=findAllPoroda[i].getDataRegion().getNumRows()-findAllPoroda[i].getMergedRanges()[0].getNumRows(),//Количество столбцов -findAllEndColumn[i].getRows()
 Data.range=findAllPoroda[i].getSheet().getRange(Data.rowData, Data.columnData, Data.numRowsData, Data.numColumnsData)
 Data.arr=findAllPoroda[i].getSheet().getRange(Data.rowData, Data.columnData, Data.numRowsData, Data.numColumnsData).getValues()

_.remove(Data.arr, function(item,i,arr) {
  return item[0]==='Всього' || +item[1]===0;
});
  
objHeader['Pivot']=[['Порода','Всього','Ділова','A','B','C','D','ПП','НП']].concat(noBlankColumnsArr2D(Data.arr))     
objHeader['Данні']=unpivot(objHeader['Pivot'],1,1,"Клас якості","Обь'ем" )  
//Удалим отсутствующие породы
 _.remove(objHeader['Данні'], function(n) {
  return n[2]== 0;
}); 

objHeader['Заголовки Данні']=objHeader['Данні'].splice(0,1)
var keyObj=[objHeader['Лісництво'],objHeader['Ділянка'],objHeader['Площа']].join("_")
sortimentna[keyObj]=objHeader
//objHeader=undefined
//unpivot(data,fixColumns,fixRows,titlePivot,titleValue)


//Logger.log([['Порода','Всього','Ділова','A','B','C','D','ПП','НП']].concat(noBlankColumnsArr2D(Data.arr))) 




//DataArr=[['Порода','','','A','B','C','D','ПП','НП']].concat(findAllPoroda[i].getSheet().getRange(Data.rowData, Data.columnData, Data.numRowsData, Data.numColumnsData).getValues()))//getRange(rowData, columnData, numRowsData, numColumnsData).getValues()
 
   }
  })
Logger.log(JSON.stringify(sortimentna))


}





function f(){
    // Assume the active spreadsheet is blank.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[1];
    sheet.getRange("C2").setValue(100);
    sheet.getRange("B3").setValue(100);
    sheet.getRange("D3").setValue(100);
    sheet.getRange("C4").setValue(100);
    // Logs "C2:C4"
    Logger.log(sheet.getRange("C3").getDataRegion(SpreadsheetApp.Dimension.ROWS).getA1Notation());
    // Logs "B3:D3"
    Logger.log(sheet.getRange("C3").getDataRegion(SpreadsheetApp.Dimension.COLUMNS).getA1Notation());
}


function getMyDataRegion(opsDataRegion){
var sheetName_=opsDataRegion.sheetName||"Border";
var ss_ = SpreadsheetApp.getActiveSpreadsheet();
var sheet_ =  ss_.getSheetByName(sheetName_);
var A1Notation_=opsDataRegion.A1Notation || ss_.getRangeByName("A1Notation").getValue();
var range_=sheet_.getRange(A1Notation_).getDataRegion()
return {
  
  
  ss:ss_,
  sheet:sheet_,
  range:range_
 }
}
function TESTgetDataRegion(sheetName,A1Notation){
   try{
/*
       // Assume the active spreadsheet is blank.
       var ss = SpreadsheetApp.getActiveSpreadsheet();
    
       var sheet =  ss.getSheetByName(sheetName)//ss.getSheets()[0];
       A1Notation = A1Notation || ss.getRangeByName("A1Notation").getValue();
       // Logs "B2:D4"
       var range=sheet.getRange(A1Notation).getDataRegion()						
       // Logger.log(nf)
						
       //var df = txtFnRange.getValue().split("~")					
       //var nf = FuncConstr(df[0], df[1])	
       */
     //sheetName
       var ops=new getMyDataRegion(sheetName,A1Notation)
       Logger.log(ops.range.getA1Notation());
       var nf = createFunctionValueNameRange(ops.ss,"txtFnRange")
       Logger.log(nf)
       nf(ops)												
												
       return ops.range.getA1Notation();
   }
   catch(err){
       return ops.range.getA1Notation(); 
   }

}

/**
 * @param  {} args
 * @param  {} body
 * @returns {Function} f
 */
function FuncConstr(args, body) {
    var f = new Function(args, body);
    return f;
}

function createFunctionValueNameRange(ss,rangeName){
    ss = ss || SpreadsheetApp.getActiveSpreadsheet(); 
    var txtFnRange = ss.getRangeByName(rangeName)
    var df = txtFnRange.getValue().split("~")					
    var nf = FuncConstr(df[0], df[1])	
    return nf
}


