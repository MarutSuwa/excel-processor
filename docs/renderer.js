document
  .getElementById("fileInput")
  .addEventListener("change", function (event) {
    var file = event.target.files[0];
    var reader = new FileReader();
    reader.onload = function (e) {
      var data = new Uint8Array(e.target.result);
      var workbook = XLSX.read(data, { type: "array" });

      // Assuming the first sheet
      var firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      var jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

      // Process the data here
      console.log(jsonData);
      displayData(jsonData);
    };
    reader.readAsArrayBuffer(file);
  });

function displayData(data) {
  var output = document.getElementById("output");
  output.innerHTML = "<pre>" + JSON.stringify(data, null, 2) + "</pre>";
}
