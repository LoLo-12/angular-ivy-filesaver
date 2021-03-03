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

  public downloadFile() {
    const day = new Date();
    const year = day.getFullYear().toString();
    const month = (day.getMonth() + 1).toString();
    const date = day.getDate().toString();
    this.exportAsExcel(this.form, year + month + date);
    // this.exportAsText(JSON.stringify(this.form), year + month + date);
  }

  private exportAsExcel(json, fileName) {
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
    const workbook: XLSX.WorkBook = {
      Sheets: { data: worksheet },
      SheetNames: ["data"]
    };
    const excelBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array"
    });
    const data: Blob = new Blob([excelBuffer], {
      type:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8"
    });
    FileSaver.saveAs(data, fileName + ".xlsx");
  }

  // download .txt
  private exportAsText(string, fileName) {
    const data: Blob = new Blob([string], { type: "text/plain;charset=utf-8" });
    FileSaver.saveAs(data, fileName + ".txt");
  }
}
