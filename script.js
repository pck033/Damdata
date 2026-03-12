function processData(){

let data = document.getElementById("excelData").value.trim();

let rows = data.split("\n");

let tableBody = document.querySelector("#resultTable tbody");

tableBody.innerHTML="";

let totalInflow = 0;
let totalRelease = 0;

rows.forEach(row => {

let cols = row.split("\t");

if(cols.length < 8) return;

let month = cols[1];

let inflow = parseFloat(cols[6]);
let release = parseFloat(cols[7]);

if(isNaN(inflow) || isNaN(release)) return;

totalInflow += inflow;
totalRelease += release;

let tr = document.createElement("tr");

tr.innerHTML = `
<td>${month}</td>
<td>${inflow.toFixed(3)}</td>
<td>${release.toFixed(3)}</td>
`;

tableBody.appendChild(tr);

});

document.getElementById("totalInflow").innerText = totalInflow.toFixed(3);
document.getElementById("totalRelease").innerText = totalRelease.toFixed(3);

}
