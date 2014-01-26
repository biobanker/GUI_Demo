
exports.writeToFile = function (argArray, filename, sheetIndex) {
  // Create a new workbook file in current working-path
  var workbook = excelbuilder.createWorkbook('./', filename)

  // Create a new worksheet with 10 columns and 12 rows
  var sheet1 = workbook.createSheet('sheet1', 10, 12);
    
  for (var i = 0; i < argArray.length; i++)
    sheet1.set(1, i+1, argArray[i]);

  // Save it
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






//
//exports.writeToFile = function (arg1, arg2) {
//  // Create a new workbook file in current working-path
//  var workbook = excelbuilder.createWorkbook('./', 'sample.xlsx')
//
//  // Create a new worksheet with 10 columns and 12 rows
//  var sheet1 = workbook.createSheet('sheet1', 10, 12);
//
//  // Fill some data
//  sheet1.set(1, 1, 'This is what you wrote in the GUI:');
//  sheet1.set(2, 1, arg1);
//  sheet1.set(3, 1, arg2);  
//  //for (var i = 2; i < 5; i++)
//  //  sheet1.set(i, 1, 'test'+i);
//
//  // Save it
//  workbook.save(function(ok){
//    if (!ok) 
//      workbook.cancel();
//    else
//      console.log('congratulations, your workbook created');
//  });
//    return;
//}
//
//exports.openBrowser = function (browserUrl, userFormArray)
//  {
//    open(browserUrl);    
//  return;
//  } 
//
//
//
//
