(function(filename) {
    var i,
        xs = [filename],
        book,
        sheet;

    ExcelApp.Visible = true;
    ExcelApp.DisplayAlerts = false;

    book = ExcelApp.Workbooks.Open(filename);
    for(i = 0; i < book.Worksheets.Count; i++) {
      sheet = book.Worksheets(i + 1);
      xs.push(sheet.Name);
    }
    WScript.Echo(xs.join(getResource("SEPARATOR") || '\t'));
    book.Close();
    ExcelApp.Quit();
})(WScript.Arguments.Named("filename"));
