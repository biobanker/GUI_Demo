
exports.writeToFile = function (argArray, filename, sheetIndex) {
  var workbook = excelbuilder.createWorkbook('./', filename)
  var sheet1 = workbook.createSheet('sheet1', 10, 12);
  
  sheet1.set(1, 1, "REPLY");
  for (var i = 0; i < argArray.length; i++)
    sheet1.set(1, i+2, argArray[i]);

  workbook.save(function(ok){
    if (!ok) 
      workbook.cancel();
    else
      console.log('congratulations, your workbook created');
  });
    return;
}


exports.openBrowser = function (browserUrl, userFormArray)
  {
    open(browserUrl);    
  return;
  } 




