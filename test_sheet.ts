const url = "https://docs.google.com/spreadsheets/d/1R_O4llA1K43Y97GAgkK97WMvWbqg-tftz_FXpcUSZPU/gviz/tq?tqx=out:json&tq&sheet=Examiner%20Information";
fetch(url).then(r => r.text()).then(t => console.log(t.slice(0, 500))).catch(e => console.error(e));
