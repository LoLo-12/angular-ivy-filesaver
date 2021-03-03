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
      id: 0,
      sex: "girl",
      name: "AA"
    }
  ];
  user_name = "";
  user_sex = "girl";
  user_id = 1;

  public addRow() {
    this.form.push({
      id: this.user_id,
      sex: this.user_sex,
      name: this.user_name
    });
    this.user_id += 1;
  }

  public downloadFile() {
    this.exportAsExcel(JSON.stringify(this.form), "2021/3/3");
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
}
