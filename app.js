
/**
 * Module dependencies.
 */

var express = require('express'),
  routes = require('./routes'),
  http = require('http'),
  path = require('path');
  excelbuilder = require('msexcel-builder');
  filehandler = require('./addModules');
  open = require('open');
  filemonitor = require('fs');
  excelParser = require('excel-parser');



var app = express();
//var curFilename = 'data/sample.xlsx';
var curFilename = 'data/Testfile3.xlsx';
var curFilenameOut = 'data/Testfile4.xlsx';

app.configure(function(){
  app.set('port', process.env.PORT || 3000);
  app.set('views', __dirname + '/views');
  app.set('view engine', 'jade');
  app.use(express.favicon());
  app.use(express.logger('dev'));
  app.use(express.bodyParser());
  app.use(express.methodOverride());
  app.use(app.router);
  app.use(require('stylus').middleware(__dirname + '/public'));
  app.use(express.static(path.join(__dirname, 'public')));
  app.engine('html', require('ejs').renderFile);
  
});

app.configure('development', function(){
  app.use(express.errorHandler());
});



 
http.createServer(app).listen(app.get('port'), function(){
  console.log("Express server listening on port " + app.get('port'));
});


app.get('/', function(req, res){
  res.render('index.html');
});

app.post('/', function(request, response){
  var surveyArray = [request.body.S1_Q1_1, request.body.S1_Q2_1, request.body.S1_Q2_2, request.body.S1_Q3_1, request.body.S1_Q3_2];
  for (var i = 0; i < surveyArray.length; i++)
    console.log("Submitted argument: " + surveyArray[i]);
  var reciept = filehandler.writeToFile(surveyArray, curFilenameOut, 'sheet1');
});



filemonitor.watchFile(curFilename, function (curr, prev) {
  console.log('the current mtime is: ' + curr.mtime);
  console.log('the previous mtime was: ' + prev.mtime);
  if (String(curr.mtime).localeCompare(String(prev.mtime)) != 0)
  {
  excelParser.parse({ inFile: curFilename, worksheet: 1, skipEmpty: true, searchFor: { term: [''], type: 'loose'} },
                  function(err, records){
  if(err) console.error(err);
  var ad2 = filehandler.openBrowser('http://localhost:3000', records)
});  
      } 
});



