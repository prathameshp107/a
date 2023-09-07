import { Component, ViewChild } from '@angular/core';
import * as GC from '@grapecity/spread-sheets';
import * as Excel from '@grapecity/spread-excelio';
import { saveAs } from 'file-saver';
// LICENSE INFORMATION
var SpreadJSKey = "";
GC.Spread.Sheets.LicenseKey = SpreadJSKey;
// NEED TO SET SpreadJS Key to EXCELIO also
(<any>Excel).LicenseKey = SpreadJSKey;
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  spreadBackColor = 'aliceblue';
  hostStyle = {
    width: '95vw',
    height: '80vh'
  };
  private spread!: GC.Spread.Sheets.Workbook;;
  private excelIO;
  constructor() {
    this.excelIO = new Excel.IO();
  }
  workbookInit(args: { spread: GC.Spread.Sheets.Workbook; }) {
    const self = this;
    self.spread = args.spread;
    const sheet = self.spread.getActiveSheet();
    sheet.getCell(0, 0).text('Test ExcelIO').foreColor('blue');
  }
  onFileChange(args: any) {
    const self = this, file = args.srcElement && args.srcElement.files && args.srcElement.files[0];
    if (self.spread && file) {
      self.excelIO.open(file, (json: Object) => {
        self.spread.fromJSON(json, {});
        setTimeout(() => {
          alert('Excel loaded successfully');
        }, 0);
      }, (error: any) => {
        alert('load fail');
      });
    }
  }
  onClickMe(args: any) {
    const self = this;
    const filename = 'ExportedExcel.xlsx';
    const json = JSON.stringify(self.spread.toJSON());
    self.excelIO.save(json, function (blob: any) {
      saveAs(blob, filename);
    }, function (e: any) {
      console.log(e);
    });
  }
}