var express = require('express');
var path = require('path');
var favicon = require('serve-favicon');
var logger = require('morgan');
var cookieParser = require('cookie-parser');
var bodyParser = require('body-parser');
var xlsx = require('xlsx');
var cheerio = require('cheerio');
var request = require('sync-request');

var routes = require('./routes/index');
var users = require('./routes/users');
var Workbook = require('workbook');

var app = express();

// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'jade');

// uncomment after placing your favicon in /public
//app.use(favicon(__dirname + '/public/favicon.ico'));
app.use(logger('dev'));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

app.use('/', routes);
app.use('/users', users);

// catch 404 and forward to error handler
app.use(function(req, res, next) {
  var err = new Error('Not Found');
  err.status = 404;
  next(err);
});

//initial file to read linkedin links from
/////////////CREATE A FILE WITH YOUR URLS AND PUT THE PATH HERE////////////
var fileName = "source.xlsx";

function parseAndWrite() {
  var readWorkbook = xlsx.readFile(fileName);
  var firstSheet = readWorkbook.SheetNames[0];

  var sheet = readWorkbook.Sheets[firstSheet];

  var data = [];

  //collect arrays of info on each person, arrays of arrays make up the workbook
  for (cell in sheet){
    if (sheet[cell]) {
      var person = [];
        if (sheet[cell].v){
          person = getLinkedinInfo(sheet[cell].v);
        }

        //data is so mean, pushing people and stuff
        data.push(person);
    }
  }

   var workbook = new Workbook().addRowsToSheet("Info", data).finalize();
   xlsx.writeFile(workbook, 'li-positions.xlsx'); 

   //c'est fini
   console.log("DONNNNNEE");
}

//grab JSON deets of current position(s)
function getLinkedinInfo(url, cell) {   
 var res = request('GET', url);    
 var $ = cheerio.load(res.body); 
 var currentArray = [];

 console.log($('.full-name').text());

 //push each persons info into one array
 currentArray.push($('.full-name').html());

 //account for multiple current positions
 $('div.current-position > div > header').each(function() {

     //dig out title and company from DOM
     var header = $(this);
     var headerKids = header.children();
     var title = $(headerKids).eq(0).text();
     var second = $(headerKids).eq(1).text();

     var company = "";

     //sometimes company and title are in different places
     if (title == ""){
      title = second;
      company = $(headerKids).eq(2).text();
     }
     else{
      company = second;
     }
    
     currentArray.push(title);
     currentArray.push(company);

     var experienceDiv = header.parent();
     var experienceHtml = experienceDiv.html();

     //extract end time from raw html by finding encoded dash
     var n = experienceHtml.indexOf("&#x2013;");
     experienceHtml = experienceHtml.slice(n, experienceHtml.length);
     var m = experienceHtml.indexOf("<");
     var end = experienceHtml.slice(0, m);
     end = end.replace("&#x2013;", " -");

     var time = experienceDiv.children('.experience-date-locale').children();
     var start = $(time).eq(0).text();

     currentArray.push(start + end);
  });      

 return currentArray;   
}

// error handlers

// development error handler
// will print stacktrace
if (app.get('env') === 'development') {
  app.use(function(err, req, res, next) {
    res.status(err.status || 500);
    res.render('error', {
      message: err.message,
      error: err
    });
  });
}

// production error handler
// no stacktraces leaked to user
app.use(function(err, req, res, next) {
  res.status(err.status || 500);
  res.render('error', {
    message: err.message,
    error: {}
  });
});

parseAndWrite();
module.exports = app;
