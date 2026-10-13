// Peminjaman Mobil Kantor - Google Apps Script (terikat ke Google Spreadsheet)
// Peran: ADMIN (General Affair, PIN di sheet "pengaturan") dan MEMBER (staf, mendaftar lalu disetujui admin).
const COLS = {
  cars: ["id", "name", "plate", "photo"],
  bookings: ["id", "carId", "name", "purpose", "dest", "sd", "ed", "status", "start", "end", "before", "after", "note", "video", "mid", "tm"],
  members: ["id", "name", "phone", "dept", "salt", "hash", "status", "created"]
};
const DATA = ["cars", "bookings"];
const PHOTO_COLS = ["photo", "before", "after", "video"];
const NUM_COLS = ["start", "end"];
const PUBLIC_BOOKING = ["id", "carId", "name", "sd", "ed", "status", "start", "end"]; // yang terlihat oleh member lain

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Peminjaman Mobil Kantor")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// Jalankan SEKALI dari editor (dan lagi setelah update): membuat sheet, PIN admin awal, folder foto, memberi izin.
function setup() {
  ["cars", "bookings", "members"].forEach(n => {
    const s = sheet_(n);
    s.getRange(1, 1, 1, COLS[n].length).setValues([COLS[n]]).setFontWeight("bold");
    if (s.getMaxRows() > 1) s.getRange(2, 1, s.getMaxRows() - 1, COLS[n].length).setNumberFormat("@");
  });
  settings_(); folder_();
}

function sheet_(n) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(n);
  if (!s) {
    s = ss.insertSheet(n);
    s.getRange(1, 1, s.getMaxRows(), COLS[n].length).setNumberFormat("@");
    s.getRange(1, 1, 1, COLS[n].length).setValues([COLS[n]]).setFontWeight("bold");
    s.setFrozenRows(1);
  }
  return s;
}

function settings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName("pengaturan");
  if (!s) {
    s = ss.insertSheet("pengaturan");
    s.getRange("A1:B2").setNumberFormat("@");
    s.getRange("A1:B1").setValues([["PIN admin (boleh diganti)", String(Math.floor(100000 + Math.random() * 900000))]]);
    s.getRange("A2").setValue("Jangan bagikan akses sheet ini ke staf.");
  }
  return s;
}

// ---------- Autentikasi ----------
function isAdmin_(auth) {
  const pin = String(settings_().getRange("B1").getValue()).trim();
  return !!(auth && auth.apin && pin && String(auth.apin).trim() === pin);
}
function checkAdmin(pin) { return isAdmin_({ apin: pin }); }

const hash_ = (pin, salt) => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ":" + pin));
const phone_ = p => String(p || "").replace(/\D/g, "").replace(/^62/, "0");
const members_ = () => sheet_("members").getDataRange().getValues().slice(1).filter(r => r[0]);

function member_(auth) {
  if (!auth || !auth.mid || !auth.pin) return null;
  const r = members_().find(r => String(r[0]) === String(auth.mid));
  if (!r || String(r[6]) !== "aktif") return null;
  return hash_(String(auth.pin).trim(), String(r[4])) === String(r[5]) ? { id: String(r[0]), name: String(r[1]) } : null;
}
function who_(auth) {
  if (isAdmin_(auth)) return { admin: true, id: "admin", name: String((auth && auth.name) || "Admin GA") };
  const m = member_(auth);
  return m ? { admin: false, id: m.id, name: m.name } : null;
}

function register(d) {
  const name = String(d.name || "").trim(), phone = phone_(d.phone), dept = String(d.dept || "").trim(), pin = String(d.pin || "").trim();
  if (!name || phone.length < 9 || !/^\d{4,6}$/.test(pin)) throw new Error("Lengkapi nama, nomor HP, dan PIN 4-6 angka.");
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (members_().some(r => String(r[2]) === phone)) throw new Error("Nomor HP ini sudah terdaftar. Silakan masuk.");
    const salt = Utilities.getUuid(), s = sheet_("members");
    s.getRange(s.getLastRow() + 1, 1, 1, COLS.members.length)
      .setValues([["m" + Date.now(), name, phone, dept, salt, hash_(pin, salt), "menunggu", String(Date.now())]]);
    return true;
  } finally { lock.releaseLock(); }
}

function login(phone, pin) {
  const r = members_().find(r => String(r[2]) === phone_(phone));
  if (!r || hash_(String(pin || "").trim(), String(r[4])) !== String(r[5])) throw new Error("Nomor HP atau PIN salah.");
  if (r[6] === "menunggu") throw new Error("Pendaftaran Anda masih menunggu persetujuan admin (General Affair).");
  if (r[6] !== "aktif") throw new Error("Akun Anda tidak aktif. Hubungi admin.");
  return { id: String(r[0]), name: String(r[1]) };
}

// ---------- Admin: kelola anggota ----------
function listMembers(auth) {
  if (!isAdmin_(auth)) throw new Error("Hanya admin.");
  return members_().map(r => ({ id: String(r[0]), name: String(r[1]), phone: String(r[2]), dept: String(r[3]), status: String(r[6]) }));
}
function memberRow_(s, id) {
  const vals = s.getDataRange().getValues();
  const i = vals.findIndex((r, k) => k > 0 && String(r[0]) === String(id));
  if (i < 1) throw new Error("Anggota tidak ditemukan.");
  return i + 1;
}
function setMemberStatus(id, status, auth) {
  if (!isAdmin_(auth)) throw new Error("Hanya admin.");
  if (["aktif", "nonaktif"].indexOf(status) < 0) throw new Error("Status tidak valid.");
  const s = sheet_("members");
  s.getRange(memberRow_(s, id), 7).setValue(status);
}
function setMemberPin(id, pin, auth) {
  if (!isAdmin_(auth)) throw new Error("Hanya admin.");
  if (!/^\d{4,6}$/.test(String(pin))) throw new Error("PIN harus 4-6 angka.");
  const s = sheet_("members"), row = memberRow_(s, id), salt = Utilities.getUuid();
  s.getRange(row, 5, 1, 2).setValues([[salt, hash_(String(pin), salt)]]);
}

// ---------- Foto & video ----------
function folder_() {
  const nm = "Foto Peminjaman Mobil", it = DriveApp.getFoldersByName(nm);
  return it.hasNext() ? it.next() : DriveApp.createFolder(nm);
}
function savePhoto_(dataUrl, name) {
  const m = dataUrl.match(/^data:(.+?);base64,(.*)$/);
  const mime = m[1], video = mime.indexOf("video/") === 0;
  const ext = video ? mime.split("/")[1].split(";")[0].replace("quicktime", "mov") : "jpg";
  const f = folder_().createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), mime, name + "." + ext));
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { f.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); }
  return video ? "https://drive.google.com/file/d/" + f.getId() + "/view" : "https://drive.google.com/thumbnail?id=" + f.getId() + "&sz=w1000";
}

// ---------- Data peminjaman ----------
// Admin melihat semua. Member melihat detail lengkap hanya untuk peminjamannya sendiri;
// peminjaman orang lain hanya tampil nama, mobil, dan tanggal (untuk kalender).
function loadAll(auth) {
  const w = who_(auth), out = { cars: {}, bookings: [] };
  DATA.forEach(n => {
    sheet_(n).getDataRange().getValues().slice(1).forEach(r => {
      if (!r[0]) return;
      const d = {};
      COLS[n].forEach((c, i) => {
        if (r[i] === "" || r[i] == null) return;
        d[c] = NUM_COLS.includes(c) ? Number(r[i]) : String(r[i]);
      });
      if (n === "cars") { out.cars[d.id] = d; return; }
      if (w && (w.admin || d.mid === w.id)) out.bookings.push(d);
      else { const m = {}; PUBLIC_BOOKING.forEach(c => { if (d[c] !== undefined) m[c] = d[c]; }); out.bookings.push(m); }
    });
  });
  return out;
}

function save(n, id, data, auth) {
  if (DATA.indexOf(n) < 0) throw new Error("Data tidak dikenal.");
  const w = who_(auth);
  if (!w) throw new Error("Sesi tidak valid. Silakan masuk lagi.");
  if (n === "cars" && !w.admin) throw new Error("Hanya admin yang boleh mengubah data mobil.");
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = sheet_(n), cols = COLS[n], vals = s.getDataRange().getValues();
    const idx = vals.findIndex((r, i) => i > 0 && String(r[0]) === String(id));
    if (n === "bookings" && !w.admin) {
      if (idx > 0 && String(vals[idx][cols.indexOf("mid")]) !== w.id)
        throw new Error("Hanya admin atau peminjam yang bersangkutan yang boleh mengubah data ini.");
      data = Object.assign({}, data, { name: w.name, mid: w.id });
    }
    const row = idx > 0 ? vals[idx].slice(0, cols.length) : cols.map(() => "");
    row[0] = id;
    cols.forEach((c, i) => {
      if (i === 0 || !(c in data)) return;
      let v = data[c];
      if (PHOTO_COLS.includes(c) && typeof v === "string" && v.indexOf("data:") === 0) v = savePhoto_(v, n + "-" + id + "-" + c);
      row[i] = v == null ? "" : String(v);
    });
    s.getRange(idx > 0 ? idx + 1 : s.getLastRow() + 1, 1, 1, cols.length).setValues([row]);
    return true;
  } finally { lock.releaseLock(); }
}

function remove(n, id, auth) {
  if (DATA.indexOf(n) < 0) throw new Error("Data tidak dikenal.");
  const w = who_(auth);
  if (!w) throw new Error("Sesi tidak valid. Silakan masuk lagi.");
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = sheet_(n), vals = s.getDataRange().getValues();
    const idx = vals.findIndex((r, i) => i > 0 && String(r[0]) === String(id));
    if (idx < 1) return;
    if (!w.admin) {
      const c = COLS.bookings, r = vals[idx];
      if (n !== "bookings" || String(r[c.indexOf("mid")]) !== w.id || String(r[c.indexOf("status")]) !== "booked")
        throw new Error("Hanya admin yang boleh menghapus data ini.");
    }
    s.deleteRow(idx + 1);
  } finally { lock.releaseLock(); }
}
