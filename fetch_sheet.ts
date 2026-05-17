import fs from "fs";
const url = "https://docs.google.com/spreadsheets/d/1R_O4llA1K43Y97GAgkK97WMvWbqg-tftz_FXpcUSZPU/gviz/tq?tqx=out:csv&sheet=Examiner%20Information";
fetch(url).then(r => r.text()).then(t => {
  fs.writeFileSync("output.html", t);
  console.log("Done", t.length);
}).catch(e => console.error(e));
