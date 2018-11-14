const BotDriver = require('botium-core').BotDriver
const Capabilities = require('botium-core/index').Capabilities
const Source = require('botium-core/index').Source
const fs = require('fs')
const XLSX = require('xlsx')
const config = require('./config');

var baseFolder = config.DIR.basedir
var resultsFolder = config.DIR.resultsdir

// building Excel begin
function datenum(v, date1904) {
  if (date1904) v += 1462;
  var epoch = Date.parse(v);
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}
function sheet_from_array_of_arrays(data, opts) {
  var ws = {};
  var range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
  for (var R = 0; R != data.length; ++R) {
    for (var C = 0; C != data[R].length; ++C) {
      if (range.s.r > R) range.s.r = R;
      if (range.s.c > C) range.s.c = C;
      if (range.e.r < R) range.e.r = R;
      if (range.e.c < C) range.e.c = C;
      var cell = { v: data[R][C] };
      if (cell.v == null) continue;
      var cell_ref = XLSX.utils.encode_cell({ c: C, r: R });

      if (typeof cell.v === 'number') cell.t = 'n';
      else if (typeof cell.v === 'boolean') cell.t = 'b';
      else if (cell.v instanceof Date) {
        cell.t = 'n'; cell.z = XLSX.SSF._table[14];
        cell.v = datenum(cell.v);
      }
      else cell.t = 's';

      ws[cell_ref] = cell;
    }
  }
  if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
  return ws;
}
// building Excel end
var data = []
var index = 0

function assert(actualList, expectedList, usersays, isList) {
  var testCase = []

  if (typeof actualList == 'undefined' || typeof expectedList == 'undefined') {
    console.log("--------------------------")

    console.log(`ERROR: Expected <${expectedList != 'undefined' ? expectedList : "UNDEFINED INPUT"}>, got <${typeof actualList != 'undefined' ? actualList : 'UNIDENTIFIED INTENT'}>`)
    testCase.push(usersays)
    testCase.push(typeof actualList != 'undefined' ? actualList : 'UNIDENTIFIED INTENT')
    testCase.push(typeof expectedList != 'undefined' ? expectedList : 'UNDEFINED INPUT')
    testCase.push("FAIL")
    data.push(testCase)

    console.log("--------------------------")

    return false
  }

  if (!isList) {
    assertb(actualList, expectedList, usersays);
  } else {
    //assert for adaptive cards
    var exist = false
    for (var key in expectedList) {
      try {
        exist = actualList[key].includes(expectedList[key])
      } catch (err) {
        exist = false
        console.log("Mismatch response found between actual and expected.......!")
      }
      if (!exist) {
        console.log("--------------------------")

        console.log(`ERROR: Expected <${expectedList}>, got <${actualList}>`)
        testCase.push(usersays)
        testCase.push(actualList)
        testCase.push(expectedList)
        testCase.push("FAIL")
        data.push(testCase)

        console.log("--------------------------")

        return false
      }
    }
    console.log("--------------------------")

    console.log(`SUCCESS: Got Expected <${expectedList}>`)
    testCase.push(usersays)
    testCase.push(actualList)
    testCase.push(expectedList)
    testCase.push("PASS")
    data.push(testCase)

    console.log("--------------------------")
    return true
  }
}

function assertb(actual, expected, usersays) {

  var testCase = []
  if (!actual || actual.indexOf(expected) < 0) {
    console.log("--------------------------")

    console.log(`ERROR: Expected <${expected}>, got <${actual}>`)
    testCase.push(usersays)
    testCase.push(actual)
    testCase.push(expected)
    testCase.push("FAIL")
    data.push(testCase)

    console.log("--------------------------")

    return false
  } else {
    console.log("--------------------------")

    console.log(`SUCCESS: Got Expected <${expected}>`)
    testCase.push(usersays)
    testCase.push(actual)
    testCase.push(expected)
    testCase.push("PASS")
    data.push(testCase)

    console.log("--------------------------")
    return true
  }
}

function fail(err) {
  console.log(`ERROR: <${err}>`)

  var ws_name = config.Excel_Properties.sheetName;

  var testResults = resultsFolder + "/" + "Results-" + file

  function Workbook() {
    if (!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
  }
  var wb = new Workbook(), ws = sheet_from_array_of_arrays(data);
  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = ws;
  XLSX.writeFile(wb, testResults);
  index++
  if (fileList.length > index) {
    readNextFile(index);
  }

}

//reading a folder 

var fileList = []

//Reading Excel
fs.readdirSync(baseFolder).forEach(file => {

  fileList.push(file)
})



var readNextFile = function (index) {
  file = fileList[index]
  //console.log("file>>>"+file)

  if (fileList.length > 0) {

    data = [];

    var ws_name = config.Excel_Properties.sheetName;
    var header = config.Excel_Properties.Headers;
    data.push(header)

    function Workbook() {
      if (!(this instanceof Workbook)) return new Workbook();
      this.SheetNames = [];
      this.Sheets = {};
    }

    console.log("Testing file:" + file)
    var testData = baseFolder + "/" + file
    var testResults = resultsFolder + "/" + "Results-" + file
    const script = fs.readFileSync(testData)

    const driver = new BotDriver()
      .setCapability(Capabilities.SCRIPTING_XLSX_SHEETNAMES, 'Dialogs')
      .setCapability(Capabilities.SCRIPTING_XLSX_SHEETNAMES_UTTERANCES, 'Utterances')

    driver.BuildFluent()
      .UserSaysText()
      .Compile(script, 'SCRIPTING_FORMAT_XSLX')
      .Compile(script, 'SCRIPTING_FORMAT_XSLX', 'SCRIPTING_TYPE_UTTERANCES')
      .RunScripts(assert, fail)
      .Exec()
      .then(() => {

        var wb = new Workbook(), ws = sheet_from_array_of_arrays(data);
        wb.SheetNames.push(ws_name);
        wb.Sheets[ws_name] = ws;
        XLSX.writeFile(wb, testResults);
        index++
        if (fileList.length > index) {
          readNextFile(index);
        }else{
          console.log("End of Testing Press Ctrl+c")
        }
      }).catch((err) => {
        console.log('ERROR: ', err)

        index++
        if (fileList.length > index) {
          readNextFile(index);
        }


      })
  } else {
    console.log("No TestData Found");
  }

}


readNextFile(index)
