
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



//
//var app = require('express')()
//  , server = require('http').createServer(app)
//  , io = require('socket.io').listen(server);
//
//server.listen(8090);
//
//app.get('/', function (req, res) {
//  res.sendfile('index2.html');
//  //res.render('index.html');
//});
//
//io.sockets.on('connection', function (socket) {
//  socket.emit('news', { hello: 'world' });
//  socket.on('my other event', function (data) {
//    console.log(data);
//  });
//});





//app.post('/', function(request, response){
//  var surveyArray = [request.body.S1_Q1_1, request.body.S1_Q2_1, request.body.S1_Q2_2, request.body.S1_Q3_1, request.body.S1_Q3_2];
//  for (var i = 0; i < surveyArray.length; i++)
//    console.log("Submitted argument: " + surveyArray[i]);
//  var reciept = filehandler.writeToFile(surveyArray, curFilenameOut, 'sheet1');
//});
//
//
//
//filemonitor.watchFile(curFilename, function (curr, prev) {
//  console.log('the current mtime is: ' + curr.mtime);
//  console.log('the previous mtime was: ' + prev.mtime);
//  if (String(curr.mtime).localeCompare(String(prev.mtime)) != 0)
//  {
//  excelParser.parse({ inFile: curFilename, worksheet: 1, skipEmpty: true, searchFor: { term: [''], type: 'loose'} },
//                  function(err, records){
//  if(err) console.error(err);
//  var ad2 = filehandler.openBrowser('http://localhost:3000', records)
//});  
//      } 
//});


////***********************
//var http = require('http'),  
//    io = require('socket.io'), // for npm, otherwise use require('./path/to/socket.io') 
//
//
//server = http.createServer(function(req, res){ 
//    // your normal server code 
//    res.writeHead(200, {'Content-Type': 'text/html'}); 
//    res.end('<h1>Hello world</h1>'); 
//});
//
//server.listen(8090);
//
//// socket.io 
//var socket = io.listen(server); 
//socket.on('connection', function(client){ 
//    // new client is here! 
//    client.on('message', function(){ 
//        console.log('message arrive');
//        client.send('some message');
//    });
//
//    client.on('disconnect', function(){
//        console.log('connection closed');
//    });
//
//});






