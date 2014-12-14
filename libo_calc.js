importClass(Packages.com.sun.star.uno.UnoRuntime);
importClass(Packages.com.sun.star.frame.XComponentLoader);
importClass(Packages.com.sun.star.frame.XStorable);
importClass(Packages.com.sun.star.sheet.XSpreadsheet);
importClass(Packages.com.sun.star.sheet.XSpreadsheetDocument);
importClass(Packages.com.sun.star.util.XCloseable);
importClass(Packages.com.sun.star.connection.NoConnectException);
importClass(Packages.com.sun.star.container.XIndexAccess);
importClass(Packages.com.sun.star.comp.helper.Bootstrap);
importClass(Packages.com.sun.star.comp.helper.BootstrapException);
importClass(Packages.com.sun.star.beans.PropertyValue);
importClass(Packages.com.sun.star.comp.servicemanager.ServiceManager);


////////////////////////////////


var Calc = (function(){

  function Calc(){}

  Calc.open = function(path, fn){
    var /* XMultiComponentFactory */ mcf,
    /* XComponent */ component,
    /* XComponentContext */ context,
    /* XSpreadsheetDocument */ doc,
    /* XComponentLoader */ componentLoader,
    /* com.sun.star.frame.Desktop */ desktop,
    /* com.sun.star.beans.PropertyValue[] */ args,
    fileUrl = "file:///" + path;

    function arg(name, value){
      var _arg = new PropertyValue();
      _arg.Name = name;
      _arg.Value = value;
      return _arg;
    }

    context = Bootstrap.bootstrap();

    mcf = context.getServiceManager();

    desktop = mcf.createInstanceWithContext(
      "com.sun.star.frame.Desktop", context);

    componentLoader = UnoRuntime.queryInterface(XComponentLoader, desktop);

    // GUI表示なし
    args = [ arg("Hidden", true) ];
    component = componentLoader.loadComponentFromURL(
      fileUrl, "_blank", 0, args);

    try{
      doc = new CalcDocument(component);
      fn(doc);
    }finally{
      doc.close();
    }
  };

  return Calc;
})();


////////////////////////////////


/**
 * Excel の Book に相当。
 * 
 * _component: com.sun.star.lang.XComponent
 * _doc: com.sun.star.sheet.XSpreadsheetDocument
 */
var CalcDocument = (function(){

  /**
   * @param {com.sun.star.lang.XComponent} component
   */
  function CalcDocument(component){
    this._component = component;

    this._doc = UnoRuntime.queryInterface(
      XSpreadsheetDocument, this._component);
  }

  var __ = CalcDocument.prototype;

  __.save = function(){
    var /* XStorable */ storable = UnoRuntime.queryInterface(
      XStorable, this._component);
    storable.store();
  };

  __.close = function(){
    var /* XCloseable */ closable = UnoRuntime.queryInterface(
      XCloseable, this._component);
    closable.close(true);
  };

  __.getSheets = function(){
    var sheets = [],
    sheetNames = this._doc.getSheets().getElementNames(),
    i, sheetName;

    for(i=0,len=sheetNames.length; i<len; i++){
      sheetName = sheetNames[i];
      var /* XSpreadsheet */ sheet = this.getSheetByIndex(i);
      sheets.push(new Sheet(sheet, sheetName));
    }

    return sheets;
  };

  __.each = function(fn){
    var sheetNames = this._doc.getSheets().getElementNames(),
    i;

    for(i=0,len=sheetNames.length; i<len; i++ ){
      fn(this.getSheetByIndex(i), i);
    }
  };

  /**
   * @return {XSpreadsheet}
   */
  __.getSheetByIndex = function(i){
    var /* XSpreadsheets */ sheets,
    /* XIndexAccess */ indexAccess;

    sheets = this._doc.getSheets();

    indexAccess = UnoRuntime.queryInterface(XIndexAccess, sheets);

    return UnoRuntime.queryInterface(
      XSpreadsheet, indexAccess.getByIndex(i));
  };

  /**
   * @return {Sheet}
   */
  __.getSheetByName = function(name){
    var sheets = this._doc.getSheets(),
    /* XSpreadsheet */ sheet,
    sheetNames = sheets.getElementNames(),
    sheetName,
    i, targetIndex;

    for(i=0,len=sheetNames.length; i<len; i++ ){
      sheetName = sheetNames[i];
      if(sheetName == name){
        targetIndex = i;
        break;
      }
    }

    sheet = this.getSheetByIndex(targetIndex);
    return new Sheet(sheet, sheetName);
  };

  return CalcDocument;
})();


////////////////////////////////


/**
 * _sheet: com.sun.star.sheet.XSpreadsheet
 */
var Sheet = (function(){

  /**
   * @param {com.sun.star.sheet.XSpreadsheet} sheet
   * @param {string} name
   */
  function Sheet(sheet, name){
    this._sheet = sheet;
    this.name = name;
  }

  var __ = Sheet.prototype;

  /**
   * @return {string} JavaScript の Stringプリミティブ値
   *     （nullは返さない）
   */
  __.get = function(row, col){
    var /* XCell */ cell;

    cell = this._sheet.getCellByPosition(col, row);
    return "" + cell.getFormula();
  };

  __.set = function(row, col, val){
    var /* XCell */ cell;

    cell = this._sheet.getCellByPosition(col, row);
    cell.setFormula(val);
  };

  __.getInt = function(row, col){
    var cell = this._sheet.getCellByPosition(col, row);
    return parseInt(cell.getFormula(), 10);
  };

  __.getFloat = function(row, col){
    var cell = this._sheet.getCellByPosition(col, row);
    return parseFloat(cell.getFormula());
  };

  return Sheet;
})();
