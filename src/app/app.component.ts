import { Component, VERSION } from "@angular/core";
import * as XLSX from "xlsx";
import * as FileSaver from "file-saver";

@Component({
  selector: "my-app",
  templateUrl: "./app.component.html",
  styleUrls: ["./app.component.css"]
})
export class AppComponent {
  form = [
    {
      sex: "girl",
      name: "AA"
    }
  ];
  userName = "";
  userSex = "girl";

  public addRow() {
    this.form.push({
      sex: this.userSex,
      name: this.userName
    });
  }

  private getToday() {
    const day = new Date();
    const year = day.getFullYear().toString();
    const month = (day.getMonth() + 1).toString();
    const date = day.getDate().toString();
    return year + month + date;
  }

  public downloadFile() {
    this.exportAsExcel(this.form, this.getToday());
    // this.exportAsText(JSON.stringify(this.form), this.getToday());
  }

  private exportAsExcel(json, fileName) {
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
    const workbook: XLSX.WorkBook = {
      Sheets: { data: worksheet },
      SheetNames: ["data"]
    };
    const excelBuffer = XLSX.write(workbook, {
      bookType: "xls",
      type: "array"
    });
    const data: Blob = new Blob([excelBuffer], {
      type:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8"
    });
    FileSaver.saveAs(data, fileName + ".xls");
  }

  // download .txt
  private exportAsText(string, fileName) {
    const data: Blob = new Blob([string], { type: "text/plain;charset=utf-8" });
    FileSaver.saveAs(data, fileName + ".txt");
  }
}
