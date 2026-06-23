const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('issue 2265 - Empty <t> with <r> rich text in shared strings', () => {
    let wb;
    before(() => {
      wb = new ExcelJS.Workbook();
      return wb.xlsx.readFile(
        `${__dirname}/../data/bug-repro-empty-shared-string.xlsx`
      );
    });

    it('should load without TypeError', () => {
      expect(wb.worksheets.length).to.equal(1);
    });
  });
});
