var OfficeExtension = require('@microsoft/office-api/office.runtime');
var FakeExcelApi = require('@microsoft/office-api/tests/fakeexcelapi');

function checkLibraryLoaded() {
    var ctx = new FakeExcelApi.ExcelClientRequestContext();
    var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
    var result1 = range.replaceValue("Hello");
    var result2 = range.replaceValue("HelloWorld");
    ctx.load(range, "Value, RowIndex");
    range.activate();
}


function testSimpleRequest() {
    var ctx = new FakeExcelApi.ExcelClientRequestContext();
    var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
    var result1 = range.replaceValue("Hello");
    var result2 = range.replaceValue("HelloWorld");
    ctx.load(range, "Value, RowIndex");
    range.activate();
    return ctx.sync().then(function () {
        console.log(result1.value);
        console.log("Done successfully!");
    });
}

function testObjectNewAndObjectAsParameter() {
    var ctx = new FakeExcelApi.ExcelClientRequestContext();
    var testCase = FakeExcelApi.TestCaseObject.newObject(ctx);
    var workbook = ctx.application.activeWorkbook;
    var range = workbook.activeWorksheet.range("A1");
    var result = testCase.calculateAddressAndSaveToRange("One Microsoft Way", "Redmond", range);
    return ctx.sync().then(
        function () {
            console.log(result.value);
            console.log("Done");
        });
}

function testArrayValue() {
    var ctx = new FakeExcelApi.ExcelClientRequestContext();
    var sheet = ctx.application.activeWorkbook.activeWorksheet;
    var range = sheet.range("A1");
    range.value = ['Hello', 123, true];
    ctx.load(range);
    return ctx.sync().then(
        function () {
            console.log("Succeeded");
            console.log("Range.Value=" + JSON.stringify(range.value));
            console.log("Done");
        });
}

function testValue2DArray() {
    var ctx = new FakeExcelApi.ExcelClientRequestContext();
    var sheet = ctx.application.activeWorkbook.activeWorksheet;
    var range = sheet.range("A1");
    range.valueArray = [['Hello', 123, true], ['World', 456, false]];
    ctx.load(range);
    return ctx.sync().then(
        function () {
            console.log("Succeeded");
            console.log("Range.ValueArray=" + JSON.stringify(range.valueArray));
        });
}

function testRest() {
    return OfficeExtension.HttpUtility.sendLocalDocumentRequest({
        url: "activeWorkbook/sheets",
        method: "GET"
    })
    .then((response) => {
        console.log(JSON.stringify(response));
    })
}

function doTests(){
    var tests = [
        testSimpleRequest, 
        testObjectNewAndObjectAsParameter, 
        testArrayValue, 
        testValue2DArray,
        testRest];
    var p = OfficeExtension.Utility._createPromiseFromResult(null);
    for (var i = 0; i < tests.length; i++){
        p = p.then(createOneTestFunc(tests[i]));
    }
    p = p.then(function(){
        console.log("--Finished all tests--");
    });
}

function createOneTestFunc(func){
    return function(){
        console.log("Running");
        console.log(func.toString());
        return func();
    };
}

OfficeExtension.HostBridgeTest.setTestFunc(doTests);

checkLibraryLoaded();

