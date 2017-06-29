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
    ctx.sync().then(function () {
        console.log(result1.value);
        console.log("Done successfully!");
    });
}

OfficeExtension.NativeBridgeTest.setTestFunc(testSimpleRequest);

checkLibraryLoaded();

