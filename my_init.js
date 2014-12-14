"use strict";

function puts(){
  for(var i=0,len=arguments.length; i<len; i++){
    print(arguments[i]);
    print("\n");
  }
}

function putsp(k, v){
  puts("" + k + " (" + v + ")");
}

function dump(obj){
  for(var k in obj){
    putsp(k, obj[k]);
  }
}

function sleep(sec){
  java.lang.Thread.sleep(sec * 1000);
}


////////////////////////////////


var _ma;

(function(){
  function MyArray(arg){

    if(arg.constructor === MyArray){
      return arg;
    }

    if(typeof arg === "undefined"){
      throw new Error("illegal argument");
    }
    this._list = arg;
  }

  _ma = function(list){
    return new MyArray(list);
  };

  MyArray.prototype.size = function(){
    return this._list.length;
  };

  MyArray.prototype.get = function(i){
    if(arguments.length === 0){
      return this._list;
    }
    return this._list[i];
  };

  MyArray.prototype.push = function(el){
    return this._list.push(el);
  };

  MyArray.prototype.each = function(fn){
    var ret;
    for(var i=0,len=this._list.length; i<len; i++){
      ret = fn(this._list[i], i);
      if(ret === false){ break; }
    }
  };

  MyArray.prototype.map = function(fn){
    var list = _ma([]);
    for(var i=0,len=this._list.length; i<len; i++){
      list.push( fn(this._list[i], i) );
    }
    return list;
  };

  MyArray.prototype.filter = function(fn){
    var list = _ma([]);
    for(var i=0,len=this._list.length; i<len; i++){
      if( fn(this._list[i], i) ){
        list.push( this._list[i] );
      }
    }
    return list;
  };

  MyArray.prototype.join = function(sep){
    var s = "";
    for(var i=0,len=this.size(); i<len; i++){
      if(i >= 1){
        s += sep;
      }
      s += this.get(i);
    }
    return s;
  };
})();

function range(from, to){
  var xs = [];
  for(var x=from; x<=to; x++){
    xs.push(x);
  }
  return _ma(xs);
}

////////////////////////////////

var _File = (function(){
  var _File = {};

  _File.exists = function(path){
    return new File(path).exists();
  };

  function createArray(type, size){
    var javaType = type;
    if(type === "Byte"){
      javaType = java.lang.Byte.TYPE;
    }else{
      throw new Error("not yet implemented");
    }
    return java.lang.reflect.Array.newInstance(
      javaType, size);
  };

  function withFileInputStream(file, fn){
    var fis;
    try{
      fis = new FileInputStream(file);
      fn(fis);
    }finally{
      if(fis){
        fis.close();
      }
    }
  }

  function withBufferedInputStream(is, fn){
    var bis;
    try{
      bis = new BufferedInputStream(is);
      fn(bis);
    } finally {
      if(bis){
        bis.close();
      }
    }
  }

  function withFileWriter(file, fn){
    var fw = new FileWriter(file);
    try{
      fn(fw);
    } finally {
      if(fw){ fw.close(); }
    }
  }

  function withBufferedWriter(writer, fn){
    var bw = new BufferedWriter(writer);
    try {
      fn(writer);
    } finally {
      if(bw){ bw.close(); }
    }
  }

  function withFileOutputStream(path, fn){
    var fos;
    try{
      fos = new FileOutputStream(path);
      fn(fos);
    }finally{
      if(fos){ fos.close(); }
    }
  }

  function withOutputStreamWriter(fos, encoding, fn){
    var osw;
    try{
      osw = new OutputStreamWriter(fos, encoding);
      fn(osw);
    }finally{
      if(osw){ osw.close(); }
    }
  }

  // ----------------

  _File.read = function(path, opts){
    if(opts && opts.binary){

      var file = new File(path);
      var result = createArray("Byte", file.length());

      withFileInputStream(file, function(is){
        withBufferedInputStream(is, function(bis){
          bis.read(result);
        });
      });

      return result;
    }else{
      var byteArray = _File.read(path, { binary: true });
      var enc = (opts && opts.encoding) || "UTF-8";
      return "" + (new java.lang.String(byteArray, enc));
    }
  };

  _File.write = function(path, text, opts){
    opts = opts || {};
    opts.encoding = opts.encoding || "UTF-8";

    withFileOutputStream(path, function(fos){
      withOutputStreamWriter(fos, opts.encoding, function(osw){
        osw.write(text);
      });
    });
  };

  _File.fixPath = function(path){
    return path.replace(/\\/g, "/");
  };

  _File.isAbsolutePath = function(path){
    return /^[A-Z]:\//i.test(_File.fixPath(path));
  };

  return _File;
})();

////////////////////////////////

var FileUtils = (function(){

  var FileUtils = {};

  function createByteArray(size){
    return java.lang.reflect.Array.newInstance(
      java.lang.Byte.TYPE, size);
  }

  FileUtils.cp = function(src, dest){
    if(new File(dest).exists()){
      throw new Error("dest file exists");
    }
    var fis, fos;
    try{
      fis = new FileInputStream(src);
      fos = new FileOutputStream(dest);
      var buf = createByteArray(1024);
      var len;
      while(true){
        len = fis.read(buf);
        if(len === -1){
          break;
        }
        fos.write(buf, 0, len);
      }
    }finally{
      if(fis){ fis.close(); }
      if(fos){ fos.close(); }
    }
  };

  FileUtils.rm = function(path){
    new File(path)["delete"]();
  };

  return FileUtils;
})();

////////////////////////////////

var Dir = (function(){

  var Dir = {};

  Dir.pwd = function(){
    return _File.fixPath("" + new File(".").getCanonicalPath());
  };

  return Dir;
})();

////////////////////////////////

var Tsv = (function(){

  var Tsv = {};

  Tsv.escape = function(s){
    return s
        .replace(/\\/g, "\\\\")
        .replace(/\r/g, "\\r")
        .replace(/\n/g, "\\n")
        .replace(/\t/g, "\\t")
    ;
  };

  Tsv.toTsv = function(list){
    return _ma(list).map(function(it){
      return Tsv.escape("" + it);
    }).join("\t");
  };

  return Tsv;
})();

////////////////////////////////

var Ltsv = (function(){

  var Ltsv = {};

  Ltsv.parseLine = function(line){
    var pairs = line.split("\t");
    var obj = {};
    _ma(pairs).each(function(pair){
      pair.match( /^(.+?):(.*)$/ );
      obj[ RegExp.$1 ] = RegExp.$2;
    });
    return obj;
  };

  return Ltsv;
})();

////////////////////////////////

function loadUtf8(str) {
  var stream = inStream(str);
  var bstream = new BufferedInputStream(stream);
  var reader = new BufferedReader(new InputStreamReader(bstream, "UTF-8"));
  var oldFilename = engine.get(engine.FILENAME);
  engine.put(engine.FILENAME, str);
  try {
    engine.eval(reader);
  }
  finally {
    engine.put(engine.FILENAME, oldFilename);
    streamClose(stream);
  }
}
