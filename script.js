// ── GLOBALS ──────────────────────────────────────────────────────
let inflowTableData = [], outflowTableData = []
let inflowAvg = [], outflowAvg = []
let reservoirName = "", provinceName = ""
let inflowChartObj = null, outflowChartObj = null

const monthMap = {
  "มกราคม":"Jan","กุมภาพันธ์":"Feb","มีนาคม":"Mar","เมษายน":"Apr",
  "พฤษภาคม":"May","มิถุนายน":"Jun","กรกฎาคม":"Jul","สิงหาคม":"Aug",
  "กันยายน":"Sep","ตุลาคม":"Oct","พฤศจิกายน":"Nov","ธันวาคม":"Dec"
}
// ปีน้ำ: เม.ย.–มี.ค.
const months = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]
const earlyMonths = new Set(["Jan","Feb","Mar"])

function toWaterYear(yearBE, month) {
  return earlyMonths.has(month) ? yearBE - 1 : yearBE
}

// ── FORMAT DETECTION ─────────────────────────────────────────────
function trimCols(line) {
  const c = line.split("\t").map(s => s.trim())
  while (c.length && c[c.length - 1] === "") c.pop()
  return c
}

function detectFormat(raw) {
  for (const line of raw.split(/\r?\n/)) {
    const c = trimCols(line)
    if (c.length >= 10) {
      const p = c[1]?.split(" ")
      if (p && monthMap[p[0]] && parseInt(p[1]) > 2400 &&
          /^(เหนือ|กลาง|ตะวันออกเฉียงเหนือ|ใต้|อีสาน)$/.test(c[2])) return 1
    }
    if (c.length === 9) {
      const p = c[1]?.split(" ")
      if (p && monthMap[p[0]] && parseInt(p[1]) > 2400 && !isNaN(parseFloat(c[2]))) return 3
    }
  }
  return 2
}

// ── FORMAT 1 PARSER ──────────────────────────────────────────────
// [0]ลำดับ [1]เดือน+ปีพ.ศ. [2]ภาค [3]ความจุ [4]น้ำใช้การ
// [5]ระดับน้ำ [6]น้ำในอ่าง [7]%รนก. [8]ไหลลงอ่าง [9]ระบาย
function parseFormat1(raw) {
  const nm = raw.match(/^([^\n\r\t]+?)\s+ภาค(?:เหนือ|กลาง|ตะวันออกเฉียงเหนือ|ใต้)/m)
  reservoirName = nm ? nm[1].trim() : "Reservoir"
  provinceName  = ""
  const data = []
  raw.split("\n").forEach(r => {
    const c = r.split("\t")
    if (c.length < 10) return
    const p = c[1]?.trim().split(" ")
    const monthThai = p?.[0], yearBE = parseInt(p?.[1])
    const month = monthMap[monthThai]
    if (!month || isNaN(yearBE)) return
    const inflow  = parseFloat(c[8]?.replace(/,/g, "")) || 0
    const outflow = parseFloat(c[9]?.replace(/,/g, "")) || 0
    data.push({ year: toWaterYear(yearBE, month), month, inflow, outflow })
  })
  return data
}

// ── FORMAT 3 PARSER ──────────────────────────────────────────────
// [0]ลำดับ [1]เดือน+ปีพ.ศ. [2]ความจุ [3]รนก.ต่ำสุด [4]น้ำในอ่าง
// [5]%รนก. [6]ไหลลงอ่าง [7]ระบาย [8]ปริมาณน้ำที่ใช้การได้
function parseFormat3(raw) {
  reservoirName = "Reservoir"; provinceName = ""
  for (const ln of raw.split(/\r?\n/).slice(0, 10)) {
    const j = ln.indexOf("\u0E08.")  // "จ."
    if (j > 0) {
      let before = ln.slice(0, j).trim()
      const pre = "\u0E2D\u0E48\u0E32\u0E07\u0E40\u0E01\u0E47\u0E1A\u0E19\u0E49\u0E33"
      if (before.startsWith(pre)) before = before.slice(pre.length).trim()
      reservoirName = before
      provinceName  = ln.slice(j + 2).trim().split(/[\s,]/)[0]
      break
    }
  }
  const data = []
  for (const line of raw.split(/\r?\n/)) {
    const c = trimCols(line)
    if (c.length !== 9) continue
    const p = c[1].split(" ")
    const month = monthMap[p[0]], yearBE = parseInt(p[1])
    if (!month || !(yearBE > 2400)) continue
    const inflow  = parseFloat(c[6].replace(/,/g, "")) || 0  // ไหลลงอ่าง
    const outflow = parseFloat(c[7].replace(/,/g, "")) || 0  // ระบาย
    data.push({ year: toWaterYear(yearBE, month), month, inflow, outflow })
  }
  return data
}

// ── FORMAT 2 PARSER ──────────────────────────────────────────────
// Compact (ไม่มี tab คั่น) — regex จับตัวเลขต่อเนื่อง
function parseFormat2(raw) {
  const nm = raw.match(/อ่างเก็บน้ำ\s*([^\s]+)\s+จ\.([^\s\r\n,]+)/)
  if (nm) { reservoirName = nm[1]; provinceName = nm[2] }
  else    { reservoirName = "Reservoir"; provinceName = "" }

  const thaiMonths = Object.keys(monthMap)
  const mp = thaiMonths.join("|")
  const rx = new RegExp(`\\d+\\s*(${mp})\\s+(\\d{4})([\\d.]+)([\\d.]+)([\\d.]+)([\\d.]+)([\\d.]+)([\\d.]+)([\\d.]+)([\\d.]+)`, "g")
  const data = []
  let m
  while ((m = rx.exec(raw)) !== null) {
    const month = monthMap[m[1]], yearBE = parseInt(m[2])
    if (!month || isNaN(yearBE)) continue
    const inflow  = parseFloat(m[9])  || 0
    const outflow = parseFloat(m[10]) || 0
    data.push({ year: toWaterYear(yearBE, month), month, inflow, outflow })
  }

  // fallback: line-by-line
  if (data.length === 0) {
    for (const line of raw.split(/\r?\n/)) {
      let foundThai = null
      for (const th of thaiMonths) { if (line.includes(th)) { foundThai = th; break } }
      if (!foundThai) continue
      const ym = line.match(new RegExp(foundThai + "\\s*(\\d{4})"))
      if (!ym) continue
      const yearBE = parseInt(ym[1]), month = monthMap[foundThai]
      if (!month || isNaN(yearBE)) continue
      const after = line.slice(line.indexOf(ym[0]) + ym[0].length)
      const nums  = (after.match(/[\d.]+/g) || [])
      if (nums.length < 8) continue
      const inflow  = parseFloat(nums[6]) || 0
      const outflow = parseFloat(nums[7]) || 0
      data.push({ year: toWaterYear(yearBE, month), month, inflow, outflow })
    }
  }
  return data
}

// ── LIVE DETECT ──────────────────────────────────────────────────
document.getElementById("excelData").addEventListener("input", function () {
  const v = this.value.trim()
  setFormatBadge(v ? detectFormat(v) : 0)
})

function setFormatBadge(fmt) {
  const badge = document.getElementById("formatBadge")
  const label = document.getElementById("formatLabel")
  badge.className = "format-detect"
  if (fmt === 1) { badge.classList.add("fmt1"); label.textContent = "รูปแบบที่ 1 · เขื่อน / อ่างเก็บน้ำขนาดใหญ่" }
  else if (fmt === 3) { badge.classList.add("fmt2"); label.textContent = "รูปแบบที่ 2 · อ่างเก็บน้ำขนาดกลาง (Tab)" }
  else if (fmt === 2) { badge.classList.add("fmt2"); label.textContent = "รูปแบบที่ 2 · อ่างเก็บน้ำขนาดกลาง (Compact)" }
  else label.textContent = "ยังไม่ได้วางข้อมูล"
}

// ── MAIN ─────────────────────────────────────────────────────────
function generateTables() {
  const raw = document.getElementById("excelData").value.trim()
  if (!raw) { alert("กรุณาวางข้อมูลก่อน"); return }
  const fmt  = detectFormat(raw)
  setFormatBadge(fmt)
  const data = fmt === 1 ? parseFormat1(raw) : fmt === 3 ? parseFormat3(raw) : parseFormat2(raw)
  if (data.length === 0) { alert("ไม่พบข้อมูล กรุณาตรวจสอบรูปแบบ"); return }

  document.getElementById("output").style.display = "block"
  document.getElementById("exportBtn").disabled = false

  const title = reservoirName + (provinceName ? " จ." + provinceName : "")
  document.getElementById("inflowTitle").textContent  = " " + title
  document.getElementById("outflowTitle").textContent = " " + title
  document.getElementById("inflowSectionTitle").textContent  = "อ่างเก็บน้ำ" + title + " · ปริมาณน้ำไหลลงอ่าง (ล้าน ลบ.ม.)"
  document.getElementById("outflowSectionTitle").textContent = "อ่างเก็บน้ำ" + title + " · ปริมาณน้ำระบาย (ล้าน ลบ.ม.)"

  const years = [...new Set(data.map(d => d.year))].sort()
  inflowTableData  = buildTableData(data, "inflow")
  outflowTableData = buildTableData(data, "outflow")

  const avgIn  = inflowTableData[inflowTableData.length - 1][13]
  const avgOut = outflowTableData[outflowTableData.length - 1][13]

  document.getElementById("metaBar").innerHTML = `
    <div class="meta-item"><div class="meta-label">อ่างเก็บน้ำ</div><div class="meta-value">${title}</div></div>
    <div class="meta-item"><div class="meta-label">ช่วงปีน้ำ (พ.ศ.)</div><div class="meta-value">${years[0]}–${years[years.length - 1]}</div></div>
    <div class="meta-item"><div class="meta-label">จำนวนปีน้ำ</div><div class="meta-value">${years.length} ปี</div></div>
    <div class="meta-item"><div class="meta-label">เฉลี่ย Inflow / ปี</div><div class="meta-value">${avgIn  != null ? avgIn.toFixed(3)  : "-"}</div></div>
    <div class="meta-item"><div class="meta-label">เฉลี่ย Outflow / ปี</div><div class="meta-value">${avgOut != null ? avgOut.toFixed(3) : "-"}</div></div>
  `
  renderTable("inflowTableDiv",  inflowTableData)
  renderTable("outflowTableDiv", outflowTableData)
  inflowAvg  = inflowTableData[inflowTableData.length - 1].slice(1, 13)
  outflowAvg = outflowTableData[outflowTableData.length - 1].slice(1, 13)
  drawChart("inflowChart",  "Inflow",  inflowTableData,  inflowAvg,  true)
  drawChart("outflowChart", "Outflow", outflowTableData, outflowAvg, false)
}

// ── BUILD TABLE DATA ──────────────────────────────────────────────
function buildTableData(data, type) {
  const years = [...new Set(data.map(d => d.year))].sort()
  const mSum = {}, mCnt = {}
  months.forEach(m => { mSum[m] = 0; mCnt[m] = 0 })
  const sheet = [["ปีน้ำ (พ.ศ.)", ...months, "รวม (ล้าน ลบ.ม.)"]]
  years.forEach(y => {
    let total = 0
    const row = [y]
    months.forEach(m => {
      const f = data.find(d => d.year === y && d.month === m)
      if (f) {
        const v = type === "inflow" ? f.inflow : f.outflow
        total += v; mSum[m] += v; mCnt[m]++; row.push(v)
      } else row.push(null)
    })
    row.push(total); sheet.push(row)
  })
  const avgRow = ["Average"]
  let avgTotal = 0
  months.forEach(m => {
    const v = mCnt[m] ? mSum[m] / mCnt[m] : null
    avgRow.push(v)
    if (v !== null) avgTotal += v
  })
  avgRow.push(avgTotal)
  sheet.push(avgRow)
  return sheet
}

// ── RENDER TABLE ─────────────────────────────────────────────────
function renderTable(id, tbl) {
  let h = "<table><thead><tr>"
  tbl[0].forEach(c => { h += `<th>${c}</th>` })
  h += "</tr></thead><tbody>"
  for (let i = 1; i < tbl.length - 1; i++) {
    const row = tbl[i]; h += "<tr>"
    row.forEach((cell, ci) => {
      if (ci === 0) h += `<td>${row[0]}</td>`
      else if (ci === row.length - 1) h += `<td><b>${cell !== null ? parseFloat(cell).toFixed(3) : "–"}</b></td>`
      else h += cell !== null ? `<td>${parseFloat(cell).toFixed(3)}</td>` : `<td class="dash">–</td>`
    })
    h += "</tr>"
  }
  h += "</tbody><tfoot><tr>"
  tbl[tbl.length - 1].forEach((cell, ci) => {
    if (ci === 0) h += "<td>Average</td>"
    else h += cell !== null ? `<td>${parseFloat(cell).toFixed(3)}</td>` : `<td class="dash">–</td>`
  })
  h += "</tr></tfoot></table>"
  document.getElementById(id).innerHTML = h
}

// ── CHART ─────────────────────────────────────────────────────────
const PAL = [
  "#58a6ff","#3fb950","#f78166","#d2a8ff","#ffa657","#79c0ff",
  "#56d364","#ff7b72","#bc8cff","#ffb77c","#a5d6ff","#85e89d",
  "#e3b341","#f0883e","#7ee787","#cae8ff"
]

function drawChart(canvasId, label, tbl, avgData, isInflow) {
  if (isInflow  && inflowChartObj)  { inflowChartObj.destroy();  inflowChartObj  = null }
  if (!isInflow && outflowChartObj) { outflowChartObj.destroy(); outflowChartObj = null }
  const datasets = []
  for (let i = 1; i < tbl.length - 1; i++) {
    const row = tbl[i]
    datasets.push({
      label: String(row[0]),
      data: row.slice(1, 13).map(v => v !== null ? parseFloat(v) : null),
      borderColor: PAL[(i - 1) % PAL.length], backgroundColor: "transparent",
      borderWidth: 1.5, tension: .35, pointRadius: 2, spanGaps: true
    })
  }
  datasets.push({
    label: "Average",
    data: avgData.map(v => v !== null ? parseFloat(v) : null),
    borderColor: "#ffffff", backgroundColor: "rgba(255,255,255,.05)",
    borderWidth: 3, borderDash: [8, 5], tension: .35,
    pointRadius: 4, pointBackgroundColor: "#ffffff", spanGaps: true
  })
  const obj = new Chart(document.getElementById(canvasId), {
    type: "line",
    data: { labels: months, datasets },
    options: {
      responsive: true,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { labels: { color: "#8b949e", font: { size: 10, family: "IBM Plex Mono" }, boxWidth: 20 } },
        title: {
          display: true,
          text: `${label} — ${reservoirName}${provinceName ? " จ." + provinceName : ""} (ล้าน ลบ.ม.)`,
          color: "#e6edf3", font: { size: 13, family: "IBM Plex Sans Thai" }
        },
        tooltip: {
          backgroundColor: "#161b22", borderColor: "#30363d", borderWidth: 1,
          titleColor: "#e6edf3", bodyColor: "#8b949e"
        }
      },
      scales: {
        x: { ticks: { color: "#8b949e", font: { size: 11 } }, grid: { color: "rgba(48,54,61,.5)" } },
        y: {
          ticks: { color: "#8b949e", font: { size: 11 } }, grid: { color: "rgba(48,54,61,.5)" },
          title: { display: true, text: "ล้าน ลบ.ม.", color: "#8b949e", font: { size: 11 } }
        }
      }
    }
  })
  if (isInflow) inflowChartObj = obj; else outflowChartObj = obj
}

// ── CLEAR ─────────────────────────────────────────────────────────
function clearAll() {
  document.getElementById("excelData").value = ""
  document.getElementById("output").style.display = "none"
  document.getElementById("exportBtn").disabled = true
  setFormatBadge(0)
  if (inflowChartObj)  { inflowChartObj.destroy();  inflowChartObj  = null }
  if (outflowChartObj) { outflowChartObj.destroy(); outflowChartObj = null }
}

// ── EXPORT EXCEL ─────────────────────────────────────────────────
async function exportExcel() {
  const wb = new ExcelJS.Workbook()
  async function addSheet(name, tbl, canvasId) {
    const ws = wb.addWorksheet(name)
    const labelMap = { "Inflow": "ปริมาณน้ำไหลลงอ่าง", "Outflow": "ปริมาณน้ำระบาย" }
    const titleRow = ws.addRow([reservoirName + (provinceName ? " จ." + provinceName : "") + " · " + labelMap[name] + " (ล้าน ลบ.ม.)"])
    titleRow.getCell(1).font = { bold: true, size: 13, color: { argb: "FF58A6FF" } }
    ws.addRow([])
    tbl.forEach((r, ri) => {
      const row = ws.addRow(r)
      if (ri === 0) row.eachCell(c => {
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF21262D" } }
        c.font = { bold: true, color: { argb: "FF8B949E" } }
        c.alignment = { horizontal: "center" }
      })
      if (ri === tbl.length - 1) row.eachCell(c => { c.font = { bold: true, color: { argb: "FF3FB950" } } })
    })
    ws.columns.forEach((col, i) => { col.width = i === 0 ? 18 : 12 })
    const imgData = document.getElementById(canvasId).toDataURL("image/png")
    const imgId = wb.addImage({ base64: imgData, extension: "png" })
    ws.addImage(imgId, { tl: { col: 0, row: tbl.length + 4 }, ext: { width: 900, height: 380 } })
  }
  await addSheet("Inflow",  inflowTableData,  "inflowChart")
  await addSheet("Outflow", outflowTableData, "outflowChart")
  const buf = await wb.xlsx.writeBuffer()
  saveAs(new Blob([buf]), reservoirName + (provinceName ? "_จ." + provinceName : "") + ".xlsx")
}
