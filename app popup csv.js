const $ = (id) => document.getElementById(id);

const fileInput = $("file");
const btn = $("btn");
const statusEl = $("status");

fileInput.addEventListener("change", () => {
  btn.disabled = !fileInput.files.length;
});

const SEP = ";";
const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;

const FAT_FDT_HEADERS = [
  "Pole ID (New)",
  "Coordinate (Lat) NEW",
  "Coordinate (Long) NEW",
  "Pole Provider (New)",
  "Pole Type",
  "FAT ID/NETWORK ID",
];

const HOME_HEADERS = [
  "HOMEPASS_ID",
  "CLUSTER_NAME",
  "PREFIX_ADDRESS",
  "STREET_NAME",
  "HOUSE_NUMBER",
  "BLOCK",
  "FLOOR",
  "RT",
  "RW",
  "DISTRICT",
  "SUB_DISTRICT",
  "FDT_CODE",
  "FAT_CODE",
  "BUILDING_LATITUDE",
  "BUILDING_LONGITUDE",
  "Category BizPass",
  "POST CODE",
  "ADDRESS POLE / FAT",
  "OV_UG",
  "HOUSE_COMMENT_",
  "BUILDING_NAME",
  "TOWER",
  "APTN",
  "FIBER_NODE__HFC_",
  "ID_Area",
  "Clamp_Hook_ID",
  "DEPLOYMENT_TYPE",
  "NEED_SURVEY",
];

const POLE_HEADERS = [
  "Pole ID (New)",
  "Coordinate (Lat) NEW",
  "Coordinate (Long) NEW",
  "Pole Provider (New)",
  "Pole Type",
  "LINE",
];

function toCSV(headers, rows) {
  const out = [headers.map(esc).join(SEP)];
  rows.forEach((r) => out.push(headers.map((h) => esc(r[h] || "")).join(SEP)));
  return out.join("\n");
}

function updatePreview(counts) {
  $("cHome").textContent = counts.HOME;
  $("cHomeBiz").textContent = counts.HOME_BIZ;
  $("cFat").textContent = counts.FAT;
  $("cFdt").textContent = counts.FDT;
  $("cPole").textContent = counts.POLE;
}

function splitFatNetworkId(value) {
  return String(value ?? "")
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean);
}

btn.addEventListener("click", async () => {
  const file = fileInput.files[0];
  if (!file) return;

  statusEl.textContent = "Memproses Master Excel...";
  btn.disabled = true;

  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheetName = wb.SheetNames.find((s) => s === "Master Data") || wb.SheetNames[0];
    const worksheet = wb.Sheets[sheetName];

    const rowsAsArray = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    const headerRow = rowsAsArray[0];
    if (!headerRow || !headerRow.length) throw new Error("Header Excel tidak terbaca");

    const LAST_COL_NAME = headerRow[headerRow.length - 1];
    const master = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    if (!master.length) throw new Error("Master kosong");

    const area = $("areaName").value || master[0]["ID_Area"] || file.name.replace(/\.(xlsx|xls)$/i, "");

    const HOME = [];
    const HOME_BIZ = [];
    master.forEach((r) => {
      const lastVal = String(r[LAST_COL_NAME] || "").trim().toUpperCase().replace(/\s+/g, "");
      const isBiz = lastVal.includes("BIZ") || lastVal.includes("BIS") || lastVal.includes("BUSINESS");
      const row = Object.fromEntries(HOME_HEADERS.map((h) => [h, r[h] || ""]));
      row["Category BizPass"] = r["Category BizPass"] || "";
      if (isBiz) HOME_BIZ.push(row);
      else HOME.push(row);
    });

    const FAT = [];
    master.forEach((r) => {
      const poleId = String(r["Pole ID (New)"] || "").trim().toUpperCase();
      if (poleId === "1A") return;

      const fatIds = splitFatNetworkId(r["FAT ID/NETWORK ID"]);
      const repeatValues = fatIds.length ? fatIds : [r["FAT ID/NETWORK ID"] || ""];

      repeatValues.forEach((fatId) => {
        const row = Object.fromEntries(FAT_FDT_HEADERS.map((h) => [h, r[h] || ""]));
        row["FAT ID/NETWORK ID"] = fatId;
        FAT.push(row);
      });
    });

    const fdtSource = master.find((r) => String(r["Pole ID (New)"] || "").trim().toUpperCase() === "1A");
    if (!fdtSource) throw new Error('Data POLE "1A" tidak ditemukan di Master');
    const FDT = [Object.fromEntries(FAT_FDT_HEADERS.map((h) => [h, fdtSource[h] || ""]))];

    // POLE tetap kondisi awal: 1 row master = 1 row POLE.
    // Split multi FAT hanya berlaku untuk FAT.csv, bukan POLE.csv.
    const POLE = master.map((r) => Object.fromEntries(POLE_HEADERS.map((h) => [h, r[h] || ""])));

    const zip = new JSZip();
    const folder = zip.folder(area);
    folder.file("HOME.csv", toCSV(HOME_HEADERS, HOME));
    folder.file("HOME-BIZ.csv", toCSV(HOME_HEADERS, HOME_BIZ));
    folder.file("FAT.csv", toCSV(FAT_FDT_HEADERS, FAT));
    folder.file("FDT.csv", toCSV(FAT_FDT_HEADERS, FDT));
    folder.file("POLE.csv", toCSV(POLE_HEADERS, POLE));

    const blob = await zip.generateAsync({ type: "blob" });
    saveAs(blob, `${area}_POPUP.zip`);

    updatePreview({ HOME: HOME.length, HOME_BIZ: HOME_BIZ.length, FAT: FAT.length, FDT: FDT.length, POLE: POLE.length });
    statusEl.textContent = `SELESAI: HOME=${HOME.length} | HOME-BIZ=${HOME_BIZ.length} | FAT=${FAT.length} | FDT=${FDT.length} | POLE=${POLE.length}`;
  } catch (e) {
    console.error(e);
    statusEl.textContent = "ERROR: " + e.message;
  } finally {
    btn.disabled = false;
  }
});
